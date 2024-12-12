const { saveAppointment, getAppointment, cancelAppointment } = require('./databaseHandler');

function handleIncomingMessage(message) {
    const text = message.text.body.toLowerCase();

    if (text.includes("hola")) {
        return "¡Hola! Bienvenido al Laboratorio de Salud Visual. ¿Cómo puedo ayudarte?\n1. Agendar cita\n2. Consultar cita\n3. Cancelar cita\n4. Información del laboratorio.";
    } else if (text === "1") {
        return "Por favor, envíame tu nombre, fecha y hora para agendar tu cita.";
    } else if (text === "2") {
        return "Por favor, envíame tu folio de cita para consultarla.";
    } else if (text === "3") {
        return "Envíame tu folio de cita para cancelarla.";
    } else if (text.includes("información")) {
        return "Estamos ubicados en la ENES, Instituto Nacional de Neurobiología. Nuestro horario es de lunes a viernes de 9:00 AM a 5:00 PM.";
    } else {
        return "Lo siento, no entiendo tu mensaje. Por favor intenta con una opción válida.";
    }
}

module.exports = { handleIncomingMessage };
