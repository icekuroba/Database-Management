# Project: Appointment Management Chatbot with WhatsApp
- Also see this reference document: [Google Drive Link](https://docs.google.com/document/d/1YGCP4cNDfUpnH5E07NcwRjhLl78VQ4Q_/edit)

## Description
This project implements a chatbot for appointment management using the **WhatsApp Cloud API** and a local server developed with **Node.js**.

The chatbot allows users to:
- **Schedule appointments** based on staff and time slot availability.
- **Check previously scheduled appointments** using a unique tracking ID.
- **Reschedule or cancel appointments** under specific conditions.
- **Access frequently asked questions (FAQs)** regarding the service (location, pricing, etc.).

---

## Required Tools

### For development
- **Node.js** – Main server environment.
- **Express** – Routing and webhook endpoint management.
- **SQLite** – Lightweight database for storing appointment information.
- **WhatsApp Cloud API** – Enables communication with users via WhatsApp.

### For exposing the server
- **Railway** – Tool to expose the local server to the Internet during development.
- **No-IP** (optional) – Assigns a free subdomain to your server if you plan to run it permanently from home.

### For chatbot testing
- **WhatsApp (Test App)** – Use a phone number registered in the WhatsApp Cloud API for testing messages.
- **Postman** (optional) – For simulating API calls and testing server routes.

---

## Key Features

### Chatbot Functionalities
1. **Schedule Appointments**:
   - Allows the user to select an available time slot and automatically assigns staff based on availability.
   - Generates a unique appointment ID.

2. **Check Appointments**:
   - Provides appointment details to the user using the unique appointment ID.

3. **Reschedule or Cancel Appointments**:
   - Users can modify or cancel their appointments at least one day in advance.

4. **Frequently Asked Questions (FAQ)**:
   - Predefined responses to common questions such as:
     - "Where are you located?"
     - "What are your service prices?"

---

## Project Structure

```
├── server.js          # Main server file
├── citas.db           # SQLite database
├── package.json       # Project dependencies
├── README.md          # Project documentation
└── routes/            # (Optional) Directory for modular server routes
```

---

## Prerequisites
1. **Node.js** and **npm** installed.
2. **SQLite** configured in the local environment.
3. A **Meta for Developers** account to access the WhatsApp Cloud API.
4. **ngrok** installed to expose the server during development.

---

### Additional Notes
- This project is designed to be **completely free** and uses tools accessible to all developers.
- It is recommended to implement good programming practices and proper data handling to ensure chatbot scalability.


