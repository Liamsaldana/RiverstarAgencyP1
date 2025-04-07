# -----------------------------
# Parte 1: Aplicación de Escritorio (PyQt5)
# Archivo: desktop_app/main.py
# -----------------------------

import sys
import requests
import sqlite3
import datetime
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel, QLineEdit, QPushButton, QMessageBox, QHBoxLayout
from PyQt5.QtCore import QTimer

SERVER_URL = "http://localhost:5000/registro"  # Cambia esto si usas otra IP pública

# Base de datos local para respaldo
conn = sqlite3.connect("respaldo.db")
cursor = conn.cursor()
cursor.execute("""
    CREATE TABLE IF NOT EXISTS registros (
        matricula TEXT,
        estado TEXT,
        fecha TEXT
    )
""")
conn.commit()

# Seguimiento del estado de la matrícula
estado_matriculas = {}

class App(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Control de Gimnasio")
        self.setGeometry(100, 100, 300, 180)
        self.initUI()
        self.iniciar_sincronizador()

    def initUI(self):
        layout = QVBoxLayout()

        self.label = QLabel("Ingresa la matrícula del alumno:")
        layout.addWidget(self.label)

        self.input_matricula = QLineEdit()
        layout.addWidget(self.input_matricula)

        boton_layout = QHBoxLayout()

        self.boton = QPushButton("Registrar")
        self.boton.clicked.connect(self.registrar_asistencia)
        boton_layout.addWidget(self.boton)

        self.boton_sync = QPushButton("Sincronizar")
        self.boton_sync.clicked.connect(self.sincronizar_datos_locales)
        boton_layout.addWidget(self.boton_sync)

        layout.addLayout(boton_layout)
        self.setLayout(layout)

    def registrar_asistencia(self):
        matricula = self.input_matricula.text().strip()
        if not matricula:
            QMessageBox.warning(self, "Error", "Por favor ingresa una matrícula.")
            return

        # Determinar estado (entrada o salida)
        estado_actual = estado_matriculas.get(matricula, "entrada")
        nuevo_estado = "salida" if estado_actual == "entrada" else "entrada"
        estado_matriculas[matricula] = nuevo_estado

        data = {
            "matricula": matricula,
            "estado": nuevo_estado
        }

        if not self.enviar_a_web(data):
            # Guardar localmente si no hay conexión
            fecha = datetime.datetime.now().isoformat()
            cursor.execute("INSERT INTO registros VALUES (?, ?, ?)", (matricula, nuevo_estado, fecha))
            conn.commit()
            QMessageBox.information(self, "Sin conexión", f"{nuevo_estado.title()} guardada localmente.")
        else:
            QMessageBox.information(self, "Éxito", f"{nuevo_estado.title()} registrada correctamente.")

        self.input_matricula.clear()

    def enviar_a_web(self, data):
        try:
            response = requests.post(SERVER_URL, json=data, timeout=5)
            return response.status_code == 200
        except:
            return False

    def sincronizar_datos_locales(self):
        cursor.execute("SELECT * FROM registros")
        pendientes = cursor.fetchall()
        exitos = 0
        for matricula, estado, fecha in pendientes:
            data = {"matricula": matricula, "estado": estado}
            if self.enviar_a_web(data):
                exitos += 1
                cursor.execute("DELETE FROM registros WHERE matricula = ? AND estado = ? AND fecha = ?", (matricula, estado, fecha))
                conn.commit()
        if exitos:
            QMessageBox.information(self, "Sincronización", f"{exitos} registros sincronizados con éxito.")

    def iniciar_sincronizador(self):
        timer = QTimer(self)
        timer.timeout.connect(self.sincronizar_datos_locales)
        timer.start(60000)  # Cada 60 segundos

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = App()
    window.show()
    sys.exit(app.exec_())

# -----------------------------
# Parte 2: Página Web (Flask)
# Archivo: web_server/app.py
# -----------------------------

from flask import Flask, request, render_template
import sqlite3
import datetime

app = Flask(__name__)

# Crear base de datos web
conn = sqlite3.connect("database.db", check_same_thread=False)
cursor = conn.cursor()
cursor.execute("""
    CREATE TABLE IF NOT EXISTS registros (
        matricula TEXT,
        estado TEXT,
        fecha TEXT
    )
""")
conn.commit()

@app.route("/registro", methods=["POST"])
def recibir_datos():
    data = request.get_json()
    matricula = data.get("matricula")
    estado = data.get("estado")
    fecha = datetime.datetime.now().isoformat()

    if matricula and estado:
        cursor.execute("INSERT INTO registros VALUES (?, ?, ?)", (matricula, estado, fecha))
        conn.commit()
        return {"message": "Registro exitoso"}, 200
    return {"message": "Datos incompletos"}, 400

@app.route("/")
def index():
    cursor.execute("SELECT * FROM registros ORDER BY fecha DESC")
    datos = cursor.fetchall()
    return render_template("index.html", datos=datos)

if __name__ == "__main__":
    app.run(debug=True)

# -----------------------------
# HTML para mostrar tabla
# Archivo: web_server/templates/index.html
# -----------------------------

# <!DOCTYPE html>
# <html>
# <head>
#     <title>Registros de Asistencia</title>
#     <style>
#         body { font-family: Arial, sans-serif; margin: 40px; }
#         table { width: 100%; border-collapse: collapse; }
#         th, td { padding: 10px; border: 1px solid #ccc; text-align: center; }
#     </style>
# </head>
# <body>
#     <h2>Registros Recientes</h2>
#     <table>
#         <tr>
#             <th>Matrícula</th>
#             <th>Estado</th>
#             <th>Fecha</th>
#         </tr>
#         {% for matricula, estado, fecha in datos %}
#         <tr>
#             <td>{{ matricula }}</td>
#             <td>{{ estado }}</td>
#             <td>{{ fecha }}</td>
#         </tr>
#         {% endfor %}
#     </table>
# </body>
# </html>
