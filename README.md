# Proyecto: Chatbot para Gestión de Citas con WhatsApp
- Revisa tambien en este documento de drive: https://docs.google.com/document/d/1YGCP4cNDfUpnH5E07NcwRjhLl78VQ4Q_/edit

## Descripción
Este proyecto implementa un chatbot para la gestión de citas utilizando la **API de WhatsApp Cloud** y un servidor local desarrollado en **Node.js**. 

El chatbot permite a los usuarios:
- **Agendar citas** basándose en disponibilidad de horarios y personal.
- **Consultar citas** previamente agendadas mediante un folio único.
- **Reagendar o cancelar citas** con condiciones específicas.
- **Acceder a preguntas frecuentes** sobre información del servicio (ubicación, precios, etc.).

---

## Herramientas necesarias

### Para el desarrollo
- **Node.js**: Servidor principal.
- **Express**: Manejo de rutas y endpoints del webhook.
- **SQLite**: Base de datos ligera para almacenar información de citas.
- **WhatsApp Cloud API**: Comunicación con los usuarios a través de WhatsApp.

### Para exponer el servidor
- **Railway**: Herramienta para exponer el servidor local a Internet durante el desarrollo.
- **No-IP (opcional)**: Asigna un subdominio gratuito a tu servidor si decides configurarlo como permanente desde casa.

### Para pruebas del chatbot
- **WhatsApp (App de prueba)**: Usa un número registrado en WhatsApp Cloud API para probar los mensajes.
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
   - Respuestas predefinidas a preguntas comunes como:
     - "¿Dónde estamos ubicados?"
     - "¿Cuáles son los precios del servicio?"

---
## Estructura del Proyecto

```
├── server.js          # Archivo principal del servidor
├── citas.db           # Base de datos SQLite
├── package.json       # Dependencias del proyecto
├── README.md          # Documentación del proyecto
└── routes/            # (Opcional) Directorio para modularizar rutas del servidor
```

---

## Requisitos Previos
1. **Node.js** y **npm** instalados.
2. **SQLite** configurado en el entorno local.
3. Una cuenta de **Meta for Developers** para acceder a la API de WhatsApp Cloud.
4. **ngrok** instalado para exponer el servidor durante el desarrollo.
---

### Notas adicionales
- Este proyecto está diseñado para ser **completamente gratuito** y utiliza herramientas accesibles para todos los desarrolladores.
- Se recomienda implementar buenas prácticas de programación y manejo de datos para garantizar la escalabilidad del chatbot.
