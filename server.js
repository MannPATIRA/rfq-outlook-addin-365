/**
 * Minimal Outlook Add-in server
 * Outlook on the web requires HTTPS. Use cert.pem and key.pem for local HTTPS.
 */

const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const https = require('https');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors({
  origin: function (origin, callback) {
    if (!origin) return callback(null, true);
    const allowed = [
      'https://outlook.office.com',
      'https://outlook.office365.com',
      'https://outlook.live.com',
      'https://localhost:3000',
      'http://localhost:3000'
    ];
    const ok = allowed.some(a => origin.startsWith(a)) ||
      origin.includes('outlook.office') ||
      origin.includes('outlook.live') ||
      origin.includes('localhost');
    callback(null, ok);
  },
  credentials: true,
  methods: ['GET', 'POST', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization', 'X-Requested-With']
}));

app.use('/assets', express.static(path.join(__dirname, 'assets')));
app.use('/assets/templates', express.static(path.join(__dirname, 'src/assets/templates')));

app.get('/taskpane.html', (req, res) => {
  res.sendFile(path.join(__dirname, 'src/taskpane/taskpane.html'));
});

app.get('/taskpane.js', (req, res) => {
  res.sendFile(path.join(__dirname, 'src/taskpane/taskpane.js'));
});

app.get('/services/auth.js', (req, res) => {
  res.type('application/javascript');
  res.sendFile(path.join(__dirname, 'src/services/auth.js'));
});

app.get('/data/rfq-data.js', (req, res) => {
  res.type('application/javascript');
  res.sendFile(path.join(__dirname, 'src/data/rfq-data.js'));
});

app.get('/data/email-templates.js', (req, res) => {
  res.type('application/javascript');
  res.sendFile(path.join(__dirname, 'src/data/email-templates.js'));
});

app.get('/data/attachments.js', (req, res) => {
  res.type('application/javascript');
  res.sendFile(path.join(__dirname, 'src/data/attachments.js'));
});

const templatesDir = path.join(__dirname, 'src/assets/templates');
app.get('/api/attachments/pdf', (req, res) => {
  const filePath = path.join(templatesDir, '412600xx.pdf');
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: '412600xx.pdf not found' });
  const buf = fs.readFileSync(filePath);
  res.json({ name: '412600xx.pdf', contentType: 'application/pdf', contentBytes: buf.toString('base64') });
});
app.get('/api/attachments/xlsx', (req, res) => {
  const filePath = path.join(templatesDir, '412600xx.xlsx');
  if (!fs.existsSync(filePath)) return res.status(404).json({ error: '412600xx.xlsx not found' });
  const buf = fs.readFileSync(filePath);
  res.json({ name: '412600xx.xlsx', contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', contentBytes: buf.toString('base64') });
});

app.get('/styles.css', (req, res) => {
  res.type('text/css');
  res.sendFile(path.join(__dirname, 'src/taskpane/styles.css'));
});

app.get('/commands.html', (req, res) => {
  res.send(`<!DOCTYPE html>
<html>
<head>
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  <script>Office.onReady(function() {});</script>
</head>
<body></body>
</html>`);
});

app.get('/manifest.xml', (req, res) => {
  res.type('application/xml');
  res.sendFile(path.join(__dirname, 'manifest.xml'));
});

const certPath = path.join(__dirname, 'cert.pem');
const keyPath = path.join(__dirname, 'key.pem');
const useHttps = fs.existsSync(certPath) && fs.existsSync(keyPath);

if (useHttps) {
  const options = {
    cert: fs.readFileSync(certPath),
    key: fs.readFileSync(keyPath)
  };
  https.createServer(options, app).listen(PORT, () => {
    console.log('Outlook add-in server at https://localhost:' + PORT);
    console.log('Taskpane: https://localhost:' + PORT + '/taskpane.html');
    console.log('Manifest: https://localhost:' + PORT + '/manifest.xml');
    console.log('(Using self-signed cert; you may need to accept it in the browser.)');
  });
} else {
  app.listen(PORT, () => {
    console.log('Outlook add-in server at http://localhost:' + PORT);
    console.log('Taskpane: http://localhost:' + PORT + '/taskpane.html');
    console.log('Manifest: http://localhost:' + PORT + '/manifest.xml');
    console.log('');
    console.log('For Outlook on the web, HTTPS is required. Generate certs:');
    console.log('  openssl req -x509 -newkey rsa:2048 -keyout key.pem -out cert.pem -days 365 -nodes -subj "/CN=localhost"');
    console.log('Then restart the server.');
  });
}
