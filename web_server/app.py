from flask import Flask, request, render_template
import sqlite3
import datetime

app = Flask(__name__)

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