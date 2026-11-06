const express = require('express');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;

const SHEET_ID = process.env.SHEET_ID;
const HOJA = process.env.HOJA || 'Hoja 2';

// --- Configuración del mail (todo por variables de entorno en Render) ---
const RESEND_API_KEY = process.env.RESEND_API_KEY;      // clave de Resend
const MAIL_FROM = process.env.MAIL_FROM || 'onboarding@resend.dev';
const MAIL_TO = process.env.MAIL_TO;                    // uno o varios, separados por coma
const CLAVE_ENVIO = process.env.CLAVE_ENVIO;            // clave secreta para /enviar
const ENVIAR_CADA_MIN = Number(process.env.ENVIAR_CADA_MIN || 0); // 0 = desactivado

// Genera el Excel con solo la pestaña elegida. Devuelve un Buffer.
async function generarExcel() {
  if (!SHEET_ID) throw new Error('Falta configurar SHEET_ID');

  const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=xlsx`;
  const respuesta = await fetch(url);
  if (!respuesta.ok) {
    throw new Error('No se pudo descargar la hoja. ¿Está compartida con "cualquier persona con el enlace"?');
  }
  const buffer = Buffer.from(await respuesta.arrayBuffer());

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(buffer);

  const ws = wb.getWorksheet(HOJA);
  if (!ws) {
    const nombres = wb.worksheets.map((w) => w.name).join(', ');
    throw new Error('No encontré la pestaña "' + HOJA + '". Pestañas disponibles: ' + nombres);
  }

  // Fórmulas -> valores
  ws.eachRow((row) => {
    row.eachCell((cell) => {
      const v = cell.value;
      if (v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v)) {
        cell.value = v.result === undefined ? '' : v.result;
      }
    });
  });

  wb.worksheets
    .filter((w) => w.id !== ws.id)
    .forEach((w) => wb.removeWorksheet(w.id));

  return Buffer.from(await wb.xlsx.writeBuffer());
}

// Envía el Excel por mail usando la API de Resend (HTTPS, no SMTP)
async function enviarMail() {
  if (!RESEND_API_KEY) throw new Error('Falta RESEND_API_KEY');
  if (!MAIL_TO) throw new Error('Falta MAIL_TO');

  const excel = await generarExcel();
  const ahora = new Date().toLocaleString('es-AR', { timeZone: 'America/Argentina/Buenos_Aires' });

  const r = await fetch('https://api.resend.com/emails', {
    method: 'POST',
    headers: {
      Authorization: 'Bearer ' + RESEND_API_KEY,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      from: MAIL_FROM,
      to: MAIL_TO.split(',').map((s) => s.trim()),
      subject: 'Informe ' + ahora,
      text: 'Adjunto el informe actualizado al ' + ahora + '.',
      attachments: [{ filename: 'informe.xlsx', content: excel.toString('base64') }],
    }),
  });

  if (!r.ok) throw new Error('Error de Resend: ' + (await r.text()));
  console.log('Mail enviado', ahora);
}

app.get('/', (req, res) => {
  res.send(`<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="utf-8">
  <title>Descargar informe</title>
</head>
<body style="font-family: sans-serif; margin: 40px;">
  <h1>Informe</h1>
  <a href="/descargar"><button style="font-size: 18px; padding: 10px 20px;">Descargar Excel</button></a>
</body>
</html>`);
});

app.get('/descargar', async (req, res) => {
  try {
    const salida = await generarExcel();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="informe.xlsx"');
    res.send(salida);
  } catch (e) {
    res.status(500).send('Error: ' + e.message);
  }
});

// Para disparar el envío desde afuera (ej: cron-job.org cada 5 minutos)
// Uso: https://TU-APP.onrender.com/enviar?clave=TU_CLAVE
app.get('/enviar', async (req, res) => {
  if (!CLAVE_ENVIO || req.query.clave !== CLAVE_ENVIO) {
    return res.status(401).send('No autorizado');
  }
  try {
    await enviarMail();
    res.send('Mail enviado');
  } catch (e) {
    console.error(e);
    res.status(500).send('Error: ' + e.message);
  }
});

app.listen(PORT, () => {
  console.log('Servicio corriendo en el puerto ' + PORT);

  // Opción interna: si definís ENVIAR_CADA_MIN=5, se envía solo cada 5 minutos
  if (ENVIAR_CADA_MIN > 0) {
    setInterval(() => {
      enviarMail().catch((e) => console.error('Falló el envío:', e.message));
    }, ENVIAR_CADA_MIN * 60 * 1000);
    console.log('Envío automático cada ' + ENVIAR_CADA_MIN + ' min');
  }
});
