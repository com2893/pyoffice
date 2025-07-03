import sys
import csv
import re
import math

# Importaciones de PyQt5 para la interfaz gráfica
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTableWidget, QTableWidgetItem,
    QAction, QFileDialog, QMessageBox, QInputDialog, QDialog, QVBoxLayout, QPushButton, QLabel,
    QColorDialog, QFontDialog, QMenu, QStatusBar
)
from PyQt5.QtGui import QColor, QFont
from PyQt5.QtCore import Qt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

class ChartDialog(QDialog):
    """
    Diálogo para mostrar gráficos de barras o pastel usando matplotlib.
    """
    def __init__(self, data, chart_type, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Gráfico")
        layout = QVBoxLayout(self)
        fig = Figure(figsize=(5, 3))
        self.canvas = FigureCanvas(fig)
        ax = fig.add_subplot(111)
        # Selección del tipo de gráfico
        if chart_type == 'Barras':
            ax.bar(range(len(data)), data)
        elif chart_type == 'Pastel':
            ax.pie(data, labels=[str(i + 1) for i in range(len(data))], autopct='%1.1f%%')
        layout.addWidget(self.canvas)
        btn = QPushButton("Cerrar")
        btn.clicked.connect(self.accept)
        layout.addWidget(btn)

class ExcelClone(QMainWindow):
    """
    Ventana principal que simula una hoja de cálculo básica.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("pycalc")
        self.resize(1200, 800)
        # Inicializa la tabla con 100 filas y 26 columnas (A-Z)
        self.table = QTableWidget(100, 26)
        self.setCentralWidget(self.table)
        self.create_menu()
        self.create_statusbar()
        # Conexiones de señales y slots
        self.table.cellChanged.connect(self.evaluate_formula)
        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_context_menu)
        self.table.itemSelectionChanged.connect(self.update_statusbar)
        # Variables auxiliares
        self.clipboard = None
        self.undo_stack = []
        self.redo_stack = []

    def create_statusbar(self):
        """
        Crea y actualiza la barra de estado inferior.
        """
        self.statusbar = QStatusBar()
        self.setStatusBar(self.statusbar)
        self.update_statusbar()

    def update_statusbar(self):
        """
        Muestra información de la celda seleccionada en la barra de estado.
        """
        items = self.table.selectedItems()
        if items:
            row = items[0].row() + 1
            col = chr(items[0].column() + 65)
            self.statusbar.showMessage(f"Celda: {col}{row} | Valor: {items[0].text()}")
        else:
            self.statusbar.showMessage("")

    def show_context_menu(self, pos):
        """
        Muestra el menú contextual al hacer clic derecho en la tabla.
        """
        menu = QMenu()
        menu.addAction("Cortar", self.cut_cells)
        menu.addAction("Copiar", self.copy_cells)
        menu.addAction("Pegar", self.paste_cells)
        menu.addAction("Borrar contenido", self.clear_cells)
        menu.addSeparator()
        menu.addAction("Insertar fila", self.insert_row)
        menu.addAction("Eliminar fila", self.delete_row)
        menu.addAction("Insertar columna", self.insert_col)
        menu.addAction("Eliminar columna", self.delete_col)
        menu.addSeparator()
        menu.addAction("Color de fondo", self.set_bg_color)
        menu.addAction("Color de texto", self.set_fg_color)
        menu.addAction("Fuente...", self.set_font)
        menu.addSeparator()
        menu.addAction("Negrita", self.set_bold)
        menu.addAction("Cursiva", self.set_italic)
        menu.addAction("Subrayado", self.set_underline)
        menu.addSeparator()
        menu.addAction("Alinear izquierda", lambda: self.set_alignment(Qt.AlignLeft))
        menu.addAction("Alinear centro", lambda: self.set_alignment(Qt.AlignCenter))
        menu.addAction("Alinear derecha", lambda: self.set_alignment(Qt.AlignRight))
        menu.addSeparator()
        menu.addAction("Insertar fecha/hora", self.insert_datetime)
        menu.addAction("Limpiar formato", self.clear_format)
        menu.exec_(self.table.viewport().mapToGlobal(pos))

    # --- Funciones de edición de celdas ---
    def cut_cells(self):
        """
        Corta las celdas seleccionadas (copia y borra).
        """
        self.copy_cells()
        self.clear_cells()

    def copy_cells(self):
        """
        Copia las celdas seleccionadas al portapapeles interno.
        """
        items = self.table.selectedItems()
        if items:
            self.clipboard = [(i.row(), i.column(), i.text()) for i in items]

    def paste_cells(self):
        """
        Pega el contenido del portapapeles en las celdas correspondientes.
        """
        if self.clipboard:
            for r, c, val in self.clipboard:
                self.table.setItem(r, c, QTableWidgetItem(val))

    def clear_cells(self):
        """
        Borra el contenido de las celdas seleccionadas.
        """
        for i in self.table.selectedItems():
            i.setText("")

    # --- Funciones de filas y columnas ---
    def insert_row(self):
        """
        Inserta una fila en la posición actual.
        """
        row = self.table.currentRow()
        self.table.insertRow(row if row >= 0 else 0)

    def delete_row(self):
        """
        Elimina la fila actual.
        """
        row = self.table.currentRow()
        if row >= 0:
            self.table.removeRow(row)

    def insert_col(self):
        """
        Inserta una columna en la posición actual.
        """
        col = self.table.currentColumn()
        self.table.insertColumn(col if col >= 0 else 0)

    def delete_col(self):
        """
        Elimina la columna actual.
        """
        col = self.table.currentColumn()
        if col >= 0:
            self.table.removeColumn(col)

    # --- Funciones de formato ---
    def set_bg_color(self):
        """
        Cambia el color de fondo de las celdas seleccionadas.
        """
        color = QColorDialog.getColor()
        if color.isValid():
            for i in self.table.selectedItems():
                i.setBackground(color)

    def set_fg_color(self):
        """
        Cambia el color del texto de las celdas seleccionadas.
        """
        color = QColorDialog.getColor()
        if color.isValid():
            for i in self.table.selectedItems():
                i.setForeground(color)

    def set_font(self):
        """
        Cambia la fuente de las celdas seleccionadas.
        """
        font, ok = QFontDialog.getFont()
        if ok:
            for i in self.table.selectedItems():
                i.setFont(font)

    def set_bold(self):
        """
        Aplica negrita al texto de las celdas seleccionadas.
        """
        for i in self.table.selectedItems():
            f = i.font()
            f.setBold(True)
            i.setFont(f)

    def set_italic(self):
        """
        Aplica cursiva al texto de las celdas seleccionadas.
        """
        for i in self.table.selectedItems():
            f = i.font()
            f.setItalic(True)
            i.setFont(f)

    def set_underline(self):
        """
        Aplica subrayado al texto de las celdas seleccionadas.
        """
        for i in self.table.selectedItems():
            f = i.font()
            f.setUnderline(True)
            i.setFont(f)

    def set_alignment(self, align):
        """
        Cambia la alineación del texto en las celdas seleccionadas.
        """
        for i in self.table.selectedItems():
            i.setTextAlignment(align)

    def insert_datetime(self):
        """
        Inserta la fecha y hora actual en las celdas seleccionadas.
        """
        import datetime
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for i in self.table.selectedItems():
            i.setText(now)

    def clear_format(self):
        """
        Restaura el formato predeterminado de las celdas seleccionadas.
        """
        for i in self.table.selectedItems():
            i.setBackground(QColor(Qt.white))
            i.setForeground(QColor(Qt.black))
            f = i.font()
            f.setBold(False)
            f.setItalic(False)
            f.setUnderline(False)
            i.setFont(f)
            i.setTextAlignment(Qt.AlignLeft)

    # --- Menú principal ---
    def create_menu(self):
        """
        Crea la barra de menús principal con todas las acciones.
        """
        menubar = self.menuBar()
        # Menú Archivo
        file_menu = menubar.addMenu("Archivo")
        open_action = QAction("Abrir", self)
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)
        save_action = QAction("Guardar", self)
        save_action.triggered.connect(self.save_file)
        file_menu.addAction(save_action)
        exit_action = QAction("Salir", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)
        # Menú Inicio
        home_menu = menubar.addMenu("Inicio")
        home_menu.addAction("Copiar", self.copy_cells)
        home_menu.addAction("Cortar", self.cut_cells)
        home_menu.addAction("Pegar", self.paste_cells)
        home_menu.addAction("Borrar contenido", self.clear_cells)
        home_menu.addSeparator()
        home_menu.addAction("Negrita", self.set_bold)
        home_menu.addAction("Cursiva", self.set_italic)
        home_menu.addAction("Subrayado", self.set_underline)
        home_menu.addSeparator()
        home_menu.addAction("Color de fondo", self.set_bg_color)
        home_menu.addAction("Color de texto", self.set_fg_color)
        home_menu.addAction("Fuente...", self.set_font)
        home_menu.addSeparator()
        home_menu.addAction("Alinear izquierda", lambda: self.set_alignment(Qt.AlignLeft))
        home_menu.addAction("Alinear centro", lambda: self.set_alignment(Qt.AlignCenter))
        home_menu.addAction("Alinear derecha", lambda: self.set_alignment(Qt.AlignRight))
        # Menú Insertar
        insert_menu = menubar.addMenu("Insertar")
        insert_menu.addAction("Gráfico...", self.insert_chart)
        insert_menu.addAction("Insertar fila", self.insert_row)
        insert_menu.addAction("Insertar columna", self.insert_col)
        insert_menu.addAction("Insertar fecha/hora", self.insert_datetime)
        # Menú Datos
        data_menu = menubar.addMenu("Datos")
        data_menu.addAction("Buscar/Reemplazar", self.find_replace)
        data_menu.addAction("Ordenar ascendente", self.sort_asc)
        data_menu.addAction("Ordenar descendente", self.sort_desc)
        # Menú Ver
        view_menu = menubar.addMenu("Ver")
        view_menu.addAction("Seleccionar todo", self.select_all)
        view_menu.addAction("Ajustar ancho de columna", self.auto_resize_columns)
        view_menu.addAction("Ajustar alto de fila", self.auto_resize_rows)
        # Menú Ayuda
        help_menu = menubar.addMenu("Ayuda")
        help_menu.addAction("Acerca de", self.show_about)

    # --- Funciones de archivo ---
    def open_file(self):
        """
        Abre un archivo CSV y carga su contenido en la tabla.
        """
        path, _ = QFileDialog.getOpenFileName(self, "Abrir archivo CSV", "", "CSV Files (*.csv)")
        if path:
            try:
                with open(path, newline='', encoding='utf-8') as f:
                    reader = csv.reader(f)
                    data = list(reader)
                    self.table.setRowCount(len(data))
                    self.table.setColumnCount(max(len(row) for row in data))
                    for i, row in enumerate(data):
                        for j, cell in enumerate(row):
                            self.table.setItem(i, j, QTableWidgetItem(cell))
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo abrir el archivo:\n{e}")

    def save_file(self):
        """
        Guarda el contenido de la tabla en un archivo CSV.
        """
        path, _ = QFileDialog.getSaveFileName(self, "Guardar archivo CSV", "", "CSV Files (*.csv)")
        if path:
            try:
                with open(path, 'w', newline='', encoding='utf-8') as f:
                    writer = csv.writer(f)
                    for row in range(self.table.rowCount()):
                        row_data = []
                        for col in range(self.table.columnCount()):
                            item = self.table.item(row, col)
                            row_data.append(item.text() if item else "")
                        writer.writerow(row_data)
            except Exception as e:
                QMessageBox.critical(self, "Error", f"No se pudo guardar el archivo:\n{e}")

    # --- Utilidades para fórmulas ---
    def cell_to_pos(self, ref):
        """
        Convierte una referencia de celda tipo 'A1' en coordenadas (fila, columna).
        """
        match = re.match(r"([A-Za-z]+)([0-9]+)", ref)
        if not match:
            return None
        # Calcula la columna considerando letras múltiples (AA, AB, etc.)
        col = sum([(ord(c.upper()) - 65 + 1) * 26 ** i for i, c in enumerate(match.group(1)[::-1])]) - 1
        row = int(match.group(2)) - 1
        return row, col

    def get_cell_value(self, row, col):
        """
        Devuelve el valor de una celda, evaluando la fórmula si es necesario.
        """
        item = self.table.item(row, col)
        if item:
            val = item.text()
            if val.startswith("="):
                return self.evaluate_formula_direct(val[1:])
            return val
        return ""

    def evaluate_formula(self, row, col):
        """
        Evalúa la fórmula de una celda si comienza con '='.
        """
        item = self.table.item(row, col)
        if not item:
            return
        text = item.text()
        if text.startswith("="):
            try:
                result = self.evaluate_formula_direct(text[1:])
                item.setText(str(result))
            except Exception as e:
                item.setText(f"#ERROR: {e}")

    def evaluate_formula_direct(self, formula):
        """
        Evalúa una fórmula simple (SUMA, PROMEDIO, RAIZ, operaciones básicas y referencias).
        """
        formula = formula.upper().replace(" ", "")
        if formula.startswith("SUMA(") and formula.endswith(")"):
            rng = formula[5:-1]
            vals = self.get_range_values(rng)
            return sum(vals)
        elif formula.startswith("PROMEDIO(") and formula.endswith(")"):
            rng = formula[9:-1]
            vals = self.get_range_values(rng)
            return sum(vals) / len(vals) if vals else 0
        elif formula.startswith("RAIZ(") and formula.endswith(")"):
            ref = formula[5:-1]
            val = self.get_single_value(ref)
            return math.sqrt(val)
        else:
            # Soporta operaciones simples como =A1+B2
            tokens = re.split(r'([+\-*/])', formula)
            result = None
            op = None
            for t in tokens:
                t = t.strip()
                if not t:
                    continue
                if t in '+-*/':
                    op = t
                else:
                    v = None
                    if re.match(r'^[A-Z]+[0-9]+$', t):
                        v = self.get_single_value(t)
                    else:
                        try:
                            v = float(t)
                        except Exception:
                            v = 0
                    if result is None:
                        result = v
                    else:
                        if op == '+':
                            result += v
                        elif op == '-':
                            result -= v
                        elif op == '*':
                            result *= v
                        elif op == '/':
                            result /= v if v != 0 else 0
            return result if result is not None else ''

    def get_range_values(self, rng):
        """
        Devuelve una lista de valores numéricos de un rango (ejemplo: 'A1:A5').
        """
        if ':' in rng:
            start, end = rng.split(':')
            r1, c1 = self.cell_to_pos(start)
            r2, c2 = self.cell_to_pos(end)
            vals = []
            for row in range(min(r1, r2), max(r1, r2) + 1):
                for col in range(min(c1, c2), max(c1, c2) + 1):
                    v = self.get_cell_value(row, col)
                    try:
                        vals.append(float(v))
                    except Exception:
                        pass
            return vals
        else:
            r, c = self.cell_to_pos(rng)
            v = self.get_cell_value(r, c)
            try:
                return [float(v)]
            except Exception:
                return [0]

    def get_single_value(self, ref):
        """
        Devuelve el valor numérico de una celda referenciada.
        """
        pos = self.cell_to_pos(ref)
        if pos:
            v = self.get_cell_value(*pos)
            try:
                return float(v)
            except Exception:
                return 0
        return 0

    # --- Funciones de gráficos ---
    def insert_chart(self):
        """
        Solicita un rango y tipo de gráfico, y muestra el gráfico correspondiente.
        """
        rng, ok = QInputDialog.getText(self, "Rango de datos", "Introduce el rango (ej: A1:A5):")
        if not ok or not rng:
            return
        data = self.get_range_values(rng.upper())
        if not data:
            QMessageBox.warning(self, "Datos", "No se encontraron datos numéricos en el rango.")
            return
        chart_type, ok = QInputDialog.getItem(self, "Tipo de gráfico", "Selecciona tipo:", ["Barras", "Pastel"], 0, False)
        if not ok:
            return
        dlg = ChartDialog(data, chart_type, self)
        dlg.exec_()

    # --- Buscar y reemplazar ---
    def find_replace(self):
        """
        Busca y reemplaza texto en todas las celdas de la tabla.
        """
        find, ok = QInputDialog.getText(self, "Buscar", "Texto a buscar:")
        if not ok or not find:
            return
        replace, ok = QInputDialog.getText(self, "Reemplazar", f"Reemplazar '{find}' por:")
        if not ok:
            return
        for r in range(self.table.rowCount()):
            for c in range(self.table.columnCount()):
                item = self.table.item(r, c)
                if item and find in item.text():
                    item.setText(item.text().replace(find, replace))

    # --- Ordenar y selección ---
    def sort_asc(self):
        """
        Ordena la columna actual de forma ascendente.
        """
        col = self.table.currentColumn()
        self.table.sortItems(col, Qt.AscendingOrder)

    def sort_desc(self):
        """
        Ordena la columna actual de forma descendente.
        """
        col = self.table.currentColumn()
        self.table.sortItems(col, Qt.DescendingOrder)

    def select_all(self):
        """
        Selecciona todas las celdas de la tabla.
        """
        self.table.selectAll()

    def auto_resize_columns(self):
        """
        Ajusta automáticamente el ancho de las columnas.
        """
        self.table.resizeColumnsToContents()

    def auto_resize_rows(self):
        """
        Ajusta automáticamente el alto de las filas.
        """
        self.table.resizeRowsToContents()

    def show_about(self):
        """
        Muestra información sobre la aplicación.
        """
        QMessageBox.information(self, "Acerca de", "pycalc\nFunciones: edición, formato, gráficos, búsqueda, ordenación, etc.")

if __name__ == "__main__":
    # Punto de entrada principal de la aplicación
    app = QApplication(sys.argv)
    window = ExcelClone()
    window.show()
    sys.exit(app.exec_())