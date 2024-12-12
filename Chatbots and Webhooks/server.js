const express = require('express');
const bodyParser = require('body-parser');
const { verifyWebhook, handleWebhook } = require('./webhookHandler');

const app = express();
const PORT = 4000;

// Middleware
app.use(bodyParser.json());

// Rutas
app.get('/webhook', verifyWebhook);
app.post('/webhook', handleWebhook);

// Iniciar el servidor
app.listen(PORT, () => {
    console.log(`Servidor corriendo en http://localhost:${PORT}`);
});
