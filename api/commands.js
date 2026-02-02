// Serverless function for commands.html
export default function handler(req, res) {
  res.setHeader('Content-Type', 'text/html');
  res.status(200).send(`<!DOCTYPE html>
<html>
<head>
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  <script>Office.onReady(function() {});</script>
</head>
<body></body>
</html>`);
}
