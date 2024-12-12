# Proyecto: Chatbot para Gestión de Citas con WhatsApp

## Descripción
Este proyecto implementa un chatbot para la gestión de citas utilizando la **API de WhatsApp Cloud** y un servidor local desarrollado en **Node.js**. El chatbot permite a los usuarios:

- **Agendar citas** basándose en disponibilidad de horarios y personal.
- **Consultar citas** previamente agendadas mediante un folio único.
- **Reagendar o cancelar citas** con condiciones específicas.
- **Acceder a preguntas frecuentes** sobre información del servicio (ubicación, precios, etc.).
---

### Herramientas necesarias
- **Node.js**: Servidor principal.
- **Express**: Manejo de rutas y endpoints del webhook.
- **SQLite**: Base de datos ligera para almacenar información de citas.
- **WhatsApp Cloud API**: Comunicación con los usuarios a través de WhatsApp.
- **ngrok**: Herramienta para exponer el servidor local a Internet durante el desarrollo.
- **No-IP (opcional)**: Para asignar un subdominio gratuito a tu servidor si decides configurarlo como permanente desde casa.
 ### Para pruebas del chatbot
- **WhatsApp (App de prueba)**:Usa un número registrado en WhatsApp Cloud API para probar los mensajes.
- **Postman (opcional)**: Para simular llamadas a la API y probar las rutas del servidor.
---

## Características Principales

### Funcionalidades del Chatbot
1. **Agendar citas**:
   - Permite al usuario elegir un horario disponible y asignar personal automáticamente según disponibilidad.
   - Genera un folio único como identificador de la cita.

2. **Consultar citas**:
   - Proporciona los detalles de la cita al usuario con base en el folio único.

3. **Reagendar o cancelar citas**:
   - Los usuarios pueden modificar o cancelar sus citas con un día de anticipación.

4. **Preguntas frecuentes**:
   - Respuestas predefinidas a preguntas comunes como "¿Dónde estamos ubicados?" o "¿Cuáles son los precios del servicio?".

