const { handleIncomingMessage } = require('./messageHandler');
const { VERIFY_TOKEN } = require('./config');

// Verificar webhook
function verifyWebhook(req, res) {
    const mode = req.query['hub.mode'];
    const token = req.query['hub.verify_token'];
    const challenge = req.query['hub.challenge'];

    if (mode && token === VERIFY_TOKEN) {
        if (mode === "subscribe") {
            console.log("Webhook verificado correctamente.");
            res.status(200).send(challenge);
        }
    } else {
        console.log("Error: Token de verificación no válido.");
        res.status(403).send("Forbidden");
    }
}

// Manejar eventos del webhook
function handleWebhook(req, res) {
    if (req.body.object) {
        req.body.entry.forEach(entry => {
            entry.changes.forEach(change => {
                if (change.field === "messages") {
                    const message = change.value.messages[0];
                    console.log("Mensaje recibido:", message);

                    // Procesar mensaje
                    const response = handleIncomingMessage(message);
                    console.log("Respuesta generada:", response);
                }
            });
        });
        res.status(200).send("EVENT_RECEIVED");
    } else {
        console.log("Evento vacío recibido.");
        res.status(404).send();
    }
}

module.exports = { verifyWebhook, handleWebhook };
