import sys
from PyQt5.QtGui import QIcon, QTextCharFormat, QColor, QFont
from PyQt5.QtCore import QFileInfo, Qt, QSize
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QTextEdit, QAction, QTabWidget, QWidget, QVBoxLayout, QFileDialog, QMessageBox, QToolBar, QFontComboBox, QComboBox, QColorDialog, QLabel, QHBoxLayout
)

class DocumentTab(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.editor = QTextEdit()
        self.editor.setPlaceholderText("Empieza a escribir tu documento...")
        layout = QVBoxLayout()
        layout.addWidget(self.editor)
        self.setLayout(layout)
        self.file_path = None

class WordApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PyWrite")
        self.setGeometry(100, 100, 1000, 700)
        self.tabs = QTabWidget()
        self.tabs.setDocumentMode(True)
        self.setCentralWidget(self.tabs)
        self.create_menus()
        self.create_toolbar()
        self.create_statusbar()
        self.apply_modern_style()
        self.new_tab()

    def create_menus(self):
        menubar = self.menuBar()
        menubar.setNativeMenuBar(False)
        # Archivo
        file_menu = menubar.addMenu("Archivo")
        new_action = QAction(QIcon.fromTheme("document-new"), "Nuevo", self)
        new_action.triggered.connect(self.new_tab)
        open_action = QAction(QIcon.fromTheme("document-open"), "Abrir...", self)
        open_action.triggered.connect(self.open_file)
        save_action = QAction(QIcon.fromTheme("document-save"), "Guardar", self)
        save_action.triggered.connect(self.save_file)
        saveas_action = QAction(QIcon.fromTheme("document-save-as"), "Guardar como...", self)
        saveas_action.triggered.connect(self.save_file_as)
        close_action = QAction(QIcon.fromTheme("window-close"), "Cerrar pestaña", self)
        close_action.triggered.connect(self.close_tab)
        exit_action = QAction(QIcon.fromTheme("application-exit"), "Salir", self)
        exit_action.triggered.connect(self.close)
        file_menu.addActions([new_action, open_action, save_action, saveas_action, close_action])
        file_menu.addSeparator()
        file_menu.addAction(exit_action)
        # Edición
        edit_menu = menubar.addMenu("Edición")
        undo_action = QAction(QIcon.fromTheme("edit-undo"), "Deshacer", self)
        undo_action.triggered.connect(lambda: self.current_editor().undo())
        redo_action = QAction(QIcon.fromTheme("edit-redo"), "Rehacer", self)
        redo_action.triggered.connect(lambda: self.current_editor().redo())
        cut_action = QAction(QIcon.fromTheme("edit-cut"), "Cortar", self)
        cut_action.triggered.connect(lambda: self.current_editor().cut())
        copy_action = QAction(QIcon.fromTheme("edit-copy"), "Copiar", self)
        copy_action.triggered.connect(lambda: self.current_editor().copy())
        paste_action = QAction(QIcon.fromTheme("edit-paste"), "Pegar", self)
        paste_action.triggered.connect(lambda: self.current_editor().paste())
        edit_menu.addActions([undo_action, redo_action])
        edit_menu.addSeparator()
        edit_menu.addActions([cut_action, copy_action, paste_action])
        # Ver
        view_menu = menubar.addMenu("Ver")
        zoom_in_action = QAction("Aumentar UI (Ctrl + +)", self)
        zoom_in_action.setShortcut("Ctrl++")
        zoom_in_action.triggered.connect(self.zoom_in_ui)
        view_menu.addAction(zoom_in_action)
        zoom_out_action = QAction("Reducir UI (Ctrl + -)", self)
        zoom_out_action.setShortcut("Ctrl+-")
        zoom_out_action.triggered.connect(self.zoom_out_ui)
        view_menu.addAction(zoom_out_action)
        reset_zoom_action = QAction("Restablecer UI", self)
        reset_zoom_action.triggered.connect(self.reset_zoom_ui)
        view_menu.addAction(reset_zoom_action)
        # Ayuda
        help_menu = menubar.addMenu("Ayuda")
        about_action = QAction(QIcon.fromTheme("help-about"), "Acerca de", self)
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def create_toolbar(self):
        toolbar = QToolBar("Formato de texto")
        toolbar.setMovable(False)
        toolbar.setIconSize(QSize(22, 22))
        self.addToolBar(Qt.TopToolBarArea, toolbar)
        # Negrita
        bold_action = QAction(QIcon.fromTheme("format-text-bold"), "Negrita", self)
        bold_action.setCheckable(True)
        bold_action.setShortcut("Ctrl+B")
        bold_action.setToolTip("Negrita")
        bold_action.triggered.connect(self.toggle_bold)
        toolbar.addAction(bold_action)
        # Cursiva
        italic_action = QAction(QIcon.fromTheme("format-text-italic"), "Cursiva", self)
        italic_action.setCheckable(True)
        italic_action.setShortcut("Ctrl+I")
        italic_action.setToolTip("Cursiva")
        italic_action.triggered.connect(self.toggle_italic)
        toolbar.addAction(italic_action)
        # Subrayado
        underline_action = QAction(QIcon.fromTheme("format-text-underline"), "Subrayado", self)
        underline_action.setCheckable(True)
        underline_action.setShortcut("Ctrl+U")
        underline_action.setToolTip("Subrayado")
        underline_action.triggered.connect(self.toggle_underline)
        toolbar.addAction(underline_action)
        toolbar.addSeparator()
        # Fuente
        self.font_box = QFontComboBox(self)
        self.font_box.setToolTip("Fuente")
        self.font_box.setCurrentFont(QFont('Arial', 14))
        self.font_box.currentFontChanged.connect(self.set_font_family)
        toolbar.addWidget(self.font_box)
        # Tamaño
        self.size_box = QComboBox(self)
        self.size_box.setEditable(True)
        self.size_box.setToolTip("Tamaño de fuente")
        for size in range(8, 36, 2):
            self.size_box.addItem(str(size))
        self.size_box.setCurrentText("14")
        self.size_box.currentTextChanged.connect(self.set_font_size)
        toolbar.addWidget(self.size_box)
        # Color
        color_action = QAction(QIcon.fromTheme("format-color-text"), "Color", self)
        color_action.setToolTip("Color de texto")
        color_action.triggered.connect(self.set_font_color)
        toolbar.addAction(color_action)
        toolbar.addSeparator()
        # Alineación
        align_left = QAction(QIcon.fromTheme("format-justify-left"), "Izquierda", self)
        align_left.setToolTip("Alinear a la izquierda")
        align_left.triggered.connect(lambda: self.set_alignment(Qt.AlignLeft))
        toolbar.addAction(align_left)
        align_center = QAction(QIcon.fromTheme("format-justify-center"), "Centrar", self)
        align_center.setToolTip("Centrar")
        align_center.triggered.connect(lambda: self.set_alignment(Qt.AlignCenter))
        toolbar.addAction(align_center)
        align_right = QAction(QIcon.fromTheme("format-justify-right"), "Derecha", self)
        align_right.setToolTip("Alinear a la derecha")
        align_right.triggered.connect(lambda: self.set_alignment(Qt.AlignRight))
        toolbar.addAction(align_right)
        align_justify = QAction(QIcon.fromTheme("format-justify-fill"), "Justificar", self)
        align_justify.setToolTip("Justificar")
        align_justify.triggered.connect(lambda: self.set_alignment(Qt.AlignJustify))
        toolbar.addAction(align_justify)

    def create_statusbar(self):
        self.status = self.statusBar()
        self.status.setStyleSheet("background: #e0e7ef; color: #3949ab; font-size: 14px;")
        self.status.showMessage("Listo para escribir")

    def toggle_bold(self):
        editor = self.current_editor()
        if editor:
            fmt = QTextCharFormat()
            weight = QFont.Bold if not editor.fontWeight() == QFont.Bold else QFont.Normal
            fmt.setFontWeight(weight)
            editor.mergeCurrentCharFormat(fmt)

    def toggle_italic(self):
        editor = self.current_editor()
        if editor:
            fmt = QTextCharFormat()
            fmt.setFontItalic(not editor.fontItalic())
            editor.mergeCurrentCharFormat(fmt)

    def toggle_underline(self):
        editor = self.current_editor()
        if editor:
            fmt = QTextCharFormat()
            fmt.setFontUnderline(not editor.fontUnderline())
            editor.mergeCurrentCharFormat(fmt)

    def set_font_family(self, font):
        editor = self.current_editor()
        if editor:
            fmt = QTextCharFormat()
            fmt.setFontFamily(font.family())
            editor.mergeCurrentCharFormat(fmt)

    def set_font_size(self, size):
        editor = self.current_editor()
        if editor:
            try:
                size = int(size)
                fmt = QTextCharFormat()
                fmt.setFontPointSize(size)
                editor.mergeCurrentCharFormat(fmt)
            except ValueError:
                pass

    def set_font_color(self):
        editor = self.current_editor()
        if editor:
            color = QColorDialog.getColor()
            if color.isValid():
                fmt = QTextCharFormat()
                fmt.setForeground(color)
                editor.mergeCurrentCharFormat(fmt)

    def set_alignment(self, alignment):
        editor = self.current_editor()
        if editor:
            editor.setAlignment(alignment)

    def apply_modern_style(self):
        self.setStyleSheet('''
            QMainWindow {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #f4f6fa, stop:1 #e3e9f7);
            }
            QTabWidget::pane {
                border: 2px solid #b0b0b0;
                background: #fff;
                border-radius: 12px;
                margin: 10px;
            }
            QTabBar::tab {
                background: #e0e7ef;
                border: 2px solid #b0b0b0;
                border-bottom: none;
                padding: 12px 32px;
                margin-right: 6px;
                font-size: 16px;
                border-top-left-radius: 10px;
                border-top-right-radius: 10px;
                color: #3a3a3a;
            }
            QTabBar::tab:selected {
                background: #fff;
                font-weight: bold;
                color: #1a237e;
            }
            QTextEdit {
                background: #fafdff;
                border: 2px solid #d0d0d0;
                font-size: 17px;
                font-family: 'Segoe UI', 'Arial', sans-serif;
                padding: 18px;
                border-radius: 10px;
                color: #222;
            }
            QToolBar {
                background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #e0e7ef, stop:1 #cfd8dc);
                border-bottom: 2.5px solid #b0b0b0;
                spacing: 12px;
                padding: 8px;
            }
            QToolButton {
                background: #f5f7fa;
                border: 2px solid #b0b0b0;
                border-radius: 8px;
                padding: 8px 12px;
                margin: 3px;
                font-size: 18px;
            }
            QToolButton:hover {
                background: #e3e9f7;
                border: 2px solid #1a237e;
                color: #1a237e;
            }
            QToolButton:checked {
                background: #c5cae9;
                border: 2px solid #3949ab;
                color: #3949ab;
            }
            QFontComboBox, QComboBox {
                background: #fafdff;
                border: 2px solid #b0b0b0;
                border-radius: 8px;
                padding: 6px 10px;
                font-size: 16px;
                min-width: 90px;
            }
            QMenuBar {
                background: #e0e7ef;
                border-bottom: 2px solid #b0b0b0;
                font-size: 16px;
            }
            QStatusBar {
                background: #e0e7ef;
                border-top: 2px solid #b0b0b0;
            }
        ''')

    def new_tab(self):
        tab = DocumentTab()
        idx = self.tabs.addTab(tab, "Sin título")
        tab.editor.setFont(QFont('Arial', 14))
        self.tabs.setCurrentIndex(idx)
        if hasattr(self, 'status') and self.status:
            self.status.showMessage("Nuevo documento creado")

    def open_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "Abrir archivo", "", "Documentos de texto (*.txt);;Todos los archivos (*)")
        if path:
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read()
            tab = DocumentTab()
            tab.editor.setText(content)
            tab.file_path = path
            tab.editor.selectAll()
            fmt = QTextCharFormat()
            fmt.setFontFamily('Arial')
            fmt.setFontPointSize(14)
            tab.editor.mergeCurrentCharFormat(fmt)
            tab.editor.moveCursor(tab.editor.textCursor().End)
            idx = self.tabs.addTab(tab, QFileInfo(path).fileName())
            self.tabs.setCurrentIndex(idx)
            if hasattr(self, 'status') and self.status:
                self.status.showMessage(f"Archivo abierto: {QFileInfo(path).fileName()}")

    def save_file(self):
        tab = self.current_tab()
        if tab and hasattr(tab, 'file_path') and tab.file_path:
            with open(tab.file_path, 'w', encoding='utf-8') as f:
                f.write(tab.editor.toPlainText())
            if hasattr(self, 'status') and self.status:
                self.status.showMessage(f"Archivo guardado: {QFileInfo(tab.file_path).fileName()}")
        else:
            self.save_file_as()

    def save_file_as(self):
        tab = self.current_tab()
        if tab:
            path, _ = QFileDialog.getSaveFileName(self, "Guardar como", "", "Documentos de texto (*.txt);;Todos los archivos (*)")
            if path:
                with open(path, 'w', encoding='utf-8') as f:
                    f.write(tab.editor.toPlainText())
                tab.file_path = path
                self.tabs.setTabText(self.tabs.currentIndex(), QFileInfo(path).fileName())
                if hasattr(self, 'status') and self.status:
                    self.status.showMessage(f"Archivo guardado como: {QFileInfo(path).fileName()}")

    def close_tab(self):
        idx = self.tabs.currentIndex()
        if self.tabs.count() > 1:
            self.tabs.removeTab(idx)
            if hasattr(self, 'status') and self.status:
                self.status.showMessage("Pestaña cerrada")
        else:
            self.close()

    def show_about(self):
        QMessageBox.about(self, "Acerca de PyWrite", "<b>PyWrite</b> - Editor de texto con pestañas<br>Inspirado en Word<br>Desarrollado con PyQt5<br><br><i>2025</i>")

    def current_tab(self):
        return self.tabs.currentWidget() if hasattr(self, 'tabs') else None

    def current_editor(self):
        tab = self.current_tab()
        return tab.editor if tab and hasattr(tab, 'editor') else None

    def zoom_in_ui(self):
        self.change_ui_scale(1.1)

    def zoom_out_ui(self):
        self.change_ui_scale(0.9)

    def reset_zoom_ui(self):
        self.set_ui_scale(1.0)

    def change_ui_scale(self, factor):
        if not hasattr(self, '_ui_scale'):
            self._ui_scale = 1.0
        self._ui_scale *= factor
        self.set_ui_scale(self._ui_scale)

    def set_ui_scale(self, scale):
        # Cambia el tamaño de fuente global y de widgets principales
        font = QFont('Arial', int(14 * scale))
        self.setFont(font)
        for i in range(self.tabs.count()):
            tab = self.tabs.widget(i)
            if hasattr(tab, 'editor'):
                tab.editor.setFont(QFont('Arial', int(14 * scale)))
        self.font_box.setCurrentFont(QFont('Arial', int(14 * scale)))
        self.size_box.setCurrentText(str(int(14 * scale)))
        # Cambia el tamaño de la barra de herramientas
        self.findChild(QToolBar).setIconSize(QSize(int(22 * scale), int(22 * scale)))

if __name__ == "__main__":
    import traceback
    try:
        app = QApplication(sys.argv)
        window = WordApp()
        window.show()
        sys.exit(app.exec_())
    except Exception as e:
        print("Error al iniciar PyWrite:", e)
        traceback.print_exc()