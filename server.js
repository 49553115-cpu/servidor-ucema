const express = require('express');
const ExcelJS = require('exceljs');

const app = express();
const PORT = process.env.PORT || 3000;

const SHEET_ID = process.env.SHEET_ID;
// Nombre de la pestaña que quieres descargar (por defecto "Hoja 2")
const HOJA = process.env.HOJA || 'Hoja 2';

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
    if (!SHEET_ID) return res.status(500).send('Falta configurar SHEET_ID');

    // 1) Baja el libro completo
    const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=xlsx`;
    const respuesta = await fetch(url);
    if (!respuesta.ok) {
      return res.status(502).send('No se pudo descargar la hoja. ¿Está compartida con "cualquier persona con el enlace"?');
    }
    const buffer = Buffer.from(await respuesta.arrayBuffer());

    // 2) Lo abre y se queda solo con la pestaña elegida
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(buffer);

    const ws = wb.getWorksheet(HOJA);
    if (!ws) {
      const nombres = wb.worksheets.map((w) => w.name).join(', ');
      return res.status(404).send('No encontré la pestaña "' + HOJA + '". Pestañas disponibles: ' + nombres);
    }

    // Convierte fórmulas en valores, para que no se rompan al sacar las otras pestañas
    ws.eachRow((row) => {
      row.eachCell((cell) => {
        const v = cell.value;
        if (v && typeof v === 'object' && ('formula' in v || 'sharedFormula' in v)) {
          cell.value = v.result === undefined ? '' : v.result;
        }
      });
    });

    // Borra las demás pestañas
    wb.worksheets
      .filter((w) => w.id !== ws.id)
      .forEach((w) => wb.removeWorksheet(w.id));

    const salida = await wb.xlsx.writeBuffer();

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="informe.xlsx"');
    res.send(Buffer.from(salida));
  } catch (e) {
    res.status(500).send('Error: ' + e.message);
  }
});

app.listen(PORT, () => console.log('Servicio corriendo en el puerto ' + PORT));