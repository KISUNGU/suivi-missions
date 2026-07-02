const express = require('express');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Serve static files from the project root
app.use(express.static(path.join(__dirname), {
  index: 'index.html'
}));

// Explicit routes for each page
app.get('/', (_req, res) => res.sendFile(path.join(__dirname, 'index.html')));
app.get('/admin', (_req, res) => res.sendFile(path.join(__dirname, 'admin.html')));
app.get('/superadmin', (_req, res) => res.sendFile(path.join(__dirname, 'superadmin.html')));

app.listen(PORT, () => {
  console.log(`PNDA Missions — serveur démarré sur http://localhost:${PORT}`);
});
