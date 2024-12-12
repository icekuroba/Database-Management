const sqlite3 = require('sqlite3').verbose();
const db = new sqlite3.Database('./citas.db');

// Crear tablas
db.serialize(() => {
    db.run("CREATE TABLE IF NOT EXISTS citas (id INTEGER PRIMARY KEY, folio TEXT, nombre TEXT, fecha TEXT, hora TEXT, estatus TEXT)");
});

// Guardar cita
function saveAppointment(nombre, fecha, hora, callback) {
    const folio = `CITA-${Date.now()}`;
    const sql = "INSERT INTO citas (folio, nombre, fecha, hora, estatus) VALUES (?, ?, ?, ?, ?)";
    db.run(sql, [folio, nombre, fecha, hora, "activa"], function (err) {
        callback(err, folio);
    });
}

// Consultar cita
function getAppointment(folio, callback) {
    const sql = "SELECT * FROM citas WHERE folio = ?";
    db.get(sql, [folio], (err, row) => {
        callback(err, row);
    });
}

// Cancelar cita
function cancelAppointment(folio, callback) {
    const sql = "UPDATE citas SET estatus = 'cancelada' WHERE folio = ?";
    db.run(sql, [folio], function (err) {
        callback(err, this.changes);
    });
}

module.exports = { saveAppointment, getAppointment, cancelAppointment };
