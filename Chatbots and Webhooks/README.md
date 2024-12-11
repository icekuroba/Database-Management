# Proyecto: Chatbot para Gestión de Citas con WhatsApp

## Descripción
Este proyecto implementa un chatbot para la gestión de citas utilizando la **API de WhatsApp Cloud** y un servidor local desarrollado en **Node.js**. El chatbot permite a los usuarios:

- **Agendar citas** basándose en disponibilidad de horarios y personal.
- **Consultar citas** previamente agendadas mediante un folio único.
- **Reagendar o cancelar citas** con condiciones específicas.
- **Acceder a preguntas frecuentes** sobre información del servicio (ubicación, precios, etc.).


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

