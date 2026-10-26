const express = require('express');

const app = express();
const PORT = process.env.PORT || 3000;

// El ID de la Google Sheet se pone en Render como variable de entorno (ver abajo)
const SHEET_ID = process.env.SHEET_ID;

// Página con el botón
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

// Descarga la Google Sheet como .xlsx y se la pasa al usuario
app.get('/descargar', async (req, res) => {
  try {
    if (!SHEET_ID) return res.status(500).send('Falta configurar SHEET_ID');

    const url = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=xlsx`;
    const respuesta = await fetch(url);

    if (!respuesta.ok) {
      return res.status(502).send('No se pudo descargar la hoja. ¿Está compartida con "cualquier persona con el enlace"?');
    }

    const buffer = Buffer.from(await respuesta.arrayBuffer());
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', 'attachment; filename="informe.xlsx"');
    res.send(buffer);
  } catch (e) {
    res.status(500).send('Error: ' + e.message);
  }
});

app.listen(PORT, () => console.log('Servicio corriendo en el puerto ' + PORT));