### Tecnologías Utilizadas
- **Node.js**: Servidor principal.
- **Express**: Manejo de rutas y endpoints del webhook.
- **SQLite**: Base de datos ligera para almacenar información de citas.
- **WhatsApp Cloud API**: Comunicación con los usuarios a través de WhatsApp.
- **ngrok**: Herramienta para exponer el servidor local a Internet durante el desarrollo.

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
