import sys
import datetime
import pandas as pd
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLineEdit, QPushButton, QLabel, QTableWidget, QTableWidgetItem,
    QMessageBox, QDialog, QGroupBox, QStatusBar
)
from PyQt5.QtCore import Qt

# -----------------------------
# Clase Student
# -----------------------------
class Student:
    def __init__(self, matricula, nombre, categoria):
        self.matricula = matricula  # Ej: "AL001", "TO002"
        self.nombre = nombre
        self.categoria = categoria  # "universidad", "preparatoria" o "colaborador"

# -----------------------------
# Lógica del gimnasio con integración a Excel
# -----------------------------
class GymManager:
    MAX_CAPACITY = 25

    def __init__(self, excel_file):
        self.student_db = self.load_database(excel_file)
        self.current_inside = {}      # {matricula: hora de entrada}
        self.waiting_list = []        # Lista de matriculas en espera
        self.day_logs = []            # Registros completos del día
        self.daily_summary = {"universidad": 0, "preparatoria": 0, "colaborador": 0}

    def load_database(self, excel_file):
        """Carga la base de datos desde un archivo Excel."""
        try:
            df = pd.read_excel(excel_file)
            df.columns = df.columns.str.strip()  # Elimina espacios extra en los encabezados
        except Exception as e:
            print(f"Error al cargar el archivo Excel: {e}")
            return {}
        student_db = {}
        # Se esperan las columnas: 'matricula', 'nombre' y 'categoria'
        for index, row in df.iterrows():
            matricula = str(row['matricula']).strip()
            student_db[matricula] = Student(matricula, row['nombre'], row['categoria'])
        return student_db

    def get_student_by_input(self, input_matricula):
        """
        Busca el estudiante en la base de datos.
        Si se ingresa el código completo (con prefijo) se usa directamente;
        si se ingresa solo la parte numérica, se busca comparando la parte numérica.
        Retorna una tupla (clave, Student) o (None, None) si no se encuentra.
        """
        input_matricula = input_matricula.strip()
        # Búsqueda exacta
        if input_matricula in self.student_db:
            return input_matricula, self.student_db[input_matricula]
        # Si el input es solo dígitos, comparar la parte numérica de cada matrícula
        if input_matricula.isdigit():
            for key, student in self.student_db.items():
                numeric_part = ''.join(filter(str.isdigit, key))
                if numeric_part == input_matricula:
                    return key, student
        # Intentar extraer la parte numérica de un input mixto
        numeric_input = ''.join(filter(str.isdigit, input_matricula))
        if numeric_input:
            for key, student in self.student_db.items():
                numeric_part = ''.join(filter(str.isdigit, key))
                if numeric_part == numeric_input:
                    return key, student
        return None, None

    def register_entry(self, input_matricula):
        key, student = self.get_student_by_input(input_matricula)
        if student is None:
            return {"error": f"La matrícula '{input_matricula}' no se encuentra en la base de datos."}

        if key in self.current_inside:
            return {"error": f"{student.nombre} ya se encuentra registrado dentro del gimnasio."}
        if key in self.waiting_list:
            return {"error": f"{student.nombre} ya está en la lista de espera."}

        entry_time = datetime.datetime.now()
        if len(self.current_inside) < GymManager.MAX_CAPACITY:
            self.current_inside[key] = entry_time
            self.daily_summary[student.categoria] += 1
            self.day_logs.append({
                "matricula": key,
                "nombre": student.nombre,
                "categoria": student.categoria,
                "entrada": entry_time,
                "salida": None
            })
            return {"success": f"Entrada registrada: {student.nombre} a las {entry_time.strftime('%H:%M:%S')}.",
                    "student": student, "entry_time": entry_time}
        else:
            # Si no hay espacio, agregar a lista de espera
            self.waiting_list.append(key)
            return {"warning": f"Capacidad máxima alcanzada. {student.nombre} ha sido agregado a la lista de espera.",
                    "student": student}

    def register_exit(self, input_matricula):
        key, student = self.get_student_by_input(input_matricula)
        if student is None:
            return {"error": f"La matrícula '{input_matricula}' no se encuentra en la base de datos."}
        if key not in self.current_inside:
            return {"error": f"No se encontró la matrícula '{key}' entre los estudiantes adentro."}

        entry_time = self.current_inside.pop(key)
        exit_time = datetime.datetime.now()
        for record in self.day_logs:
            if record["matricula"] == key and record["salida"] is None:
                record["salida"] = exit_time
                break
        return {"success": f"Salida registrada: {student.nombre} a las {exit_time.strftime('%H:%M:%S')}.",
                "student": student, "exit_time": exit_time}

    def admit_from_waiting(self, key):
        """Admite manualmente a un estudiante desde la lista de espera, si hay espacio."""
        if key in self.waiting_list and len(self.current_inside) < GymManager.MAX_CAPACITY:
            self.waiting_list.remove(key)
            entry_time = datetime.datetime.now()
            self.current_inside[key] = entry_time
            student = self.student_db[key]
            self.daily_summary[student.categoria] += 1
            self.day_logs.append({
                "matricula": key,
                "nombre": student.nombre,
                "categoria": student.categoria,
                "entrada": entry_time,
                "salida": None
            })
            return {"success": f"El estudiante {student.nombre} fue admitido desde la lista de espera."}
        else:
            return {"error": "No se pudo admitir al estudiante (verifique la capacidad o la lista de espera)."}

    def cancel_waiting(self, key):
        """Elimina manualmente un estudiante de la lista de espera."""
        if key in self.waiting_list:
            self.waiting_list.remove(key)
            student = self.student_db[key]
            return {"success": f"{student.nombre} ha sido removido de la lista de espera."}
        else:
            return {"error": "El estudiante no se encontró en la lista de espera."}

# -----------------------------
# Ventana de Registro Completo (logs del día)
# -----------------------------
class RegistroCompletoWindow(QDialog):
    def __init__(self, gym_manager):
        super().__init__()
        self.gym_manager = gym_manager
        self.setWindowTitle("Registro Completo del Día")
        self.resize(600, 400)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        self.table = QTableWidget()
        self.table.setColumnCount(5)
        self.table.setHorizontalHeaderLabels(["Matrícula", "Nombre", "Categoría", "Entrada", "Salida"])
        layout.addWidget(self.table)
        self.setLayout(layout)
        self.refresh_table()

    def refresh_table(self):
        logs = self.gym_manager.day_logs
        self.table.setRowCount(len(logs))
        for row, record in enumerate(logs):
            self.table.setItem(row, 0, QTableWidgetItem(record["matricula"]))
            self.table.setItem(row, 1, QTableWidgetItem(record["nombre"]))
            self.table.setItem(row, 2, QTableWidgetItem(record["categoria"]))
            entrada = record["entrada"].strftime('%H:%M:%S')
            self.table.setItem(row, 3, QTableWidgetItem(entrada))
            salida = record["salida"].strftime('%H:%M:%S') if record["salida"] is not None else ""
            self.table.setItem(row, 4, QTableWidgetItem(salida))
        self.table.resizeColumnsToContents()

# -----------------------------
# Ventana de Estadísticas (diarias, semanales y mensuales)
# -----------------------------
class EstadisticasWindow(QDialog):
    def __init__(self, gym_manager):
        super().__init__()
        self.gym_manager = gym_manager
        self.setWindowTitle("Estadísticas")
        self.resize(300, 200)
        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()
        daily = self.gym_manager.daily_summary
        daily_label = QLabel(
            f"Diario:\nUniversidad: {daily['universidad']}\nPreparatoria: {daily['preparatoria']}\nColaborador: {daily['colaborador']}"
        )
        daily_label.setAlignment(Qt.AlignCenter)
        weekly_label = QLabel("Semanal: Datos no disponibles")
        weekly_label.setAlignment(Qt.AlignCenter)
        monthly_label = QLabel("Mensual: Datos no disponibles")
        monthly_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(daily_label)
        layout.addWidget(weekly_label)
        layout.addWidget(monthly_label)
        self.setLayout(layout)

# -----------------------------
# Ventana principal
# -----------------------------
class MainWindow(QMainWindow):
    def __init__(self, gym_manager):
        super().__init__()
        self.gym_manager = gym_manager
        self.setWindowTitle("Sistema de Asistencia del Gimnasio")
        self.resize(900, 600)
        self.init_ui()

    def init_ui(self):
        central_widget = QWidget()
        main_layout = QVBoxLayout()

        # Indicador de espacios restantes
        self.spaces_label = QLabel()
        self.spaces_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(self.spaces_label)

        # Área de ingreso de matrícula y botones
        ingreso_group = QGroupBox("Ingreso de Matrícula")
        ingreso_layout = QHBoxLayout()
        self.matricula_input = QLineEdit()
        self.matricula_input.setPlaceholderText("Ingrese la matrícula (Ej: AL001, TO001 o 001)")
        btn_entrada = QPushButton("Registrar Entrada")
        btn_salida = QPushButton("Registrar Salida")
        ingreso_layout.addWidget(self.matricula_input)
        ingreso_layout.addWidget(btn_entrada)
        ingreso_layout.addWidget(btn_salida)
        ingreso_group.setLayout(ingreso_layout)
        main_layout.addWidget(ingreso_group)

        # Área de tablas: Activos y Lista de Espera
        tablas_group = QGroupBox("Estado Actual")
        tablas_layout = QHBoxLayout()
        # Tabla de estudiantes activos
        self.active_table = QTableWidget()
        self.active_table.setColumnCount(3)
        self.active_table.setHorizontalHeaderLabels(["Matrícula", "Nombre", "Entrada"])
        # Tabla de lista de espera con columna de selección
        self.waiting_table = QTableWidget()
        self.waiting_table.setColumnCount(4)
        self.waiting_table.setHorizontalHeaderLabels(["Seleccionar", "Orden", "Matrícula", "Nombre"])
        tablas_layout.addWidget(self.active_table)
        tablas_layout.addWidget(self.waiting_table)
        tablas_group.setLayout(tablas_layout)
        main_layout.addWidget(tablas_group)

        # Botones para procesar la lista de espera
        waiting_buttons_layout = QHBoxLayout()
        self.btn_admitir = QPushButton("Admitir Seleccionados")
        self.btn_cancelar = QPushButton("Cancelar Seleccionados")
        waiting_buttons_layout.addWidget(self.btn_admitir)
        waiting_buttons_layout.addWidget(self.btn_cancelar)
        main_layout.addLayout(waiting_buttons_layout)

        # Botones para abrir ventanas adicionales
        botones_layout = QHBoxLayout()
        btn_registro = QPushButton("Ver Registro Completo")
        btn_estadisticas = QPushButton("Ver Estadísticas")
        botones_layout.addWidget(btn_registro)
        botones_layout.addWidget(btn_estadisticas)
        main_layout.addLayout(botones_layout)

        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)

        # Conectar señales
        btn_entrada.clicked.connect(self.handle_entry)
        btn_salida.clicked.connect(self.handle_exit)
        btn_registro.clicked.connect(self.show_registro_completo)
        btn_estadisticas.clicked.connect(self.show_estadisticas)
        self.btn_admitir.clicked.connect(self.admit_selected)
        self.btn_cancelar.clicked.connect(self.cancel_selected)

        self.update_tables()

    def handle_entry(self):
        matricula = self.matricula_input.text().strip()
        if not matricula:
            QMessageBox.warning(self, "Aviso", "Debe ingresar una matrícula.")
            return
        result = self.gym_manager.register_entry(matricula)
        if "error" in result:
            QMessageBox.critical(self, "Error", result["error"])
        elif "warning" in result:
            QMessageBox.warning(self, "Lista de Espera", result["warning"])
        elif "success" in result:
            QMessageBox.information(self, "Entrada", result["success"])
        self.matricula_input.clear()
        self.update_tables()

    def handle_exit(self):
        matricula = self.matricula_input.text().strip()
        if not matricula:
            QMessageBox.warning(self, "Aviso", "Debe ingresar la matrícula para registrar la salida.")
            return
        result = self.gym_manager.register_exit(matricula)
        if "error" in result:
            QMessageBox.critical(self, "Error", result["error"])
        else:
            QMessageBox.information(self, "Salida", result["success"])
        self.matricula_input.clear()
        self.update_tables()

    def update_tables(self):
        # Actualizar indicador de espacios restantes
        available = GymManager.MAX_CAPACITY - len(self.gym_manager.current_inside)
        self.spaces_label.setText(f"Espacios restantes: {available}")

        # Actualizar tabla de estudiantes activos
        active = self.gym_manager.current_inside
        self.active_table.setRowCount(len(active))
        for row, (matricula, entry_time) in enumerate(active.items()):
            student = self.gym_manager.student_db[matricula]
            self.active_table.setItem(row, 0, QTableWidgetItem(matricula))
            self.active_table.setItem(row, 1, QTableWidgetItem(student.nombre))
            self.active_table.setItem(row, 2, QTableWidgetItem(entry_time.strftime('%H:%M:%S')))
        self.active_table.resizeColumnsToContents()

        # Actualizar tabla de lista de espera
        waiting = self.gym_manager.waiting_list
        self.waiting_table.setRowCount(len(waiting))
        for row, matricula in enumerate(waiting, start=1):
            student = self.gym_manager.student_db[matricula]
            # Columna 0: Checkbox (selección manual)
            chk_item = QTableWidgetItem()
            chk_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled)
            chk_item.setCheckState(Qt.Unchecked)
            self.waiting_table.setItem(row-1, 0, chk_item)
            # Columna 1: Orden
            self.waiting_table.setItem(row-1, 1, QTableWidgetItem(str(row)))
            # Columna 2: Matrícula
            self.waiting_table.setItem(row-1, 2, QTableWidgetItem(matricula))
            # Columna 3: Nombre
            self.waiting_table.setItem(row-1, 3, QTableWidgetItem(student.nombre))
        self.waiting_table.resizeColumnsToContents()

    def admit_selected(self):
        # Itera sobre las filas de la tabla de lista de espera y procesa las seleccionadas
        rows_to_process = []
        for row in range(self.waiting_table.rowCount()):
            item = self.waiting_table.item(row, 0)
            if item is not None and item.checkState() == Qt.Checked:
                matricula = self.waiting_table.item(row, 2).text()
                rows_to_process.append(matricula)
        if not rows_to_process:
            QMessageBox.information(self, "Admitir", "No se seleccionó ningún estudiante para admitir.")
            return
        # Procesa cada selección, respetando la capacidad disponible
        admitted = []
        for matricula in rows_to_process:
            available = GymManager.MAX_CAPACITY - len(self.gym_manager.current_inside)
            if available <= 0:
                break
            result = self.gym_manager.admit_from_waiting(matricula)
            if "success" in result:
                admitted.append(matricula)
        if admitted:
            QMessageBox.information(self, "Admitir", f"Admitidos: {', '.join(admitted)}")
        else:
            QMessageBox.warning(self, "Admitir", "No se pudo admitir a ningún estudiante.")
        self.update_tables()

    def cancel_selected(self):
        # Procesa la cancelación de las filas seleccionadas en la lista de espera
        rows_to_cancel = []
        for row in range(self.waiting_table.rowCount()):
            item = self.waiting_table.item(row, 0)
            if item is not None and item.checkState() == Qt.Checked:
                matricula = self.waiting_table.item(row, 2).text()
                rows_to_cancel.append(matricula)
        if not rows_to_cancel:
            QMessageBox.information(self, "Cancelar", "No se seleccionó ningún estudiante para cancelar.")
            return
        cancelled = []
        for matricula in rows_to_cancel:
            result = self.gym_manager.cancel_waiting(matricula)
            if "success" in result:
                cancelled.append(matricula)
        if cancelled:
            QMessageBox.information(self, "Cancelar", f"Cancelados: {', '.join(cancelled)}")
        self.update_tables()

    def show_registro_completo(self):
        self.registro_window = RegistroCompletoWindow(self.gym_manager)
        self.registro_window.exec_()

    def show_estadisticas(self):
        self.estadisticas_window = EstadisticasWindow(self.gym_manager)
        self.estadisticas_window.exec_()

# -----------------------------
# Función principal
# -----------------------------
def main():
    # Ruta absoluta al archivo Excel (ajusta según tu sistema)
    excel_file = "DBGym.xlsx"  # Reemplaza 'tu_usuario' según corresponda
    gym_manager = GymManager(excel_file)
    app = QApplication(sys.argv)
    window = MainWindow(gym_manager)
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
