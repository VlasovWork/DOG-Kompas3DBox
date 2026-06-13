import sys
from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, 
    QFileDialog, QLabel, QTableWidget, QTableWidgetItem, 
    QLineEdit
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon

sys.path.append(r"C:\Users\vlasovpi\Desktop\DOG-Kompas3DBox\DOG-macros")

from Macros.razvertka_dxf import convert_to_dxf

class FileTableWidget(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        
        # Конфигурация колонок таблицы
        self.setColumnCount(3)
        self.setHorizontalHeaderLabels(["Путь к файлу", "Кол-во", "Гибка"])
        
        # Растягиваем колонку с путем, фиксируем остальные
        self.horizontalHeader().setSectionResizeMode(0, self.horizontalHeader().ResizeMode.Stretch)
        self.horizontalHeader().setSectionResizeMode(1, self.horizontalHeader().ResizeMode.Interactive)
        self.horizontalHeader().setSectionResizeMode(2, self.horizontalHeader().ResizeMode.Interactive)
        self.setColumnWidth(1, 80)
        self.setColumnWidth(2, 70)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event):
        event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if file_path.lower().endswith(".m3d"):
                self.add_file_row(file_path)
        event.acceptProposedAction()

    def add_file_row(self, file_path):
        # Защита от дубликатов файлов в таблице
        for row in range(self.rowCount()):
            if self.item(row, 0).text() == file_path:
                return

        row_position = self.rowCount()
        self.insertRow(row_position)

        # 1. Колонка: Путь к файлу (выключаем редактирование текста)
        path_item = QTableWidgetItem(file_path)
        path_item.setFlags(path_item.flags() ^ Qt.ItemIsEditable)
        self.setItem(row_position, 0, path_item)

        # 2. Колонка: Поле ввода количества (валидация на числа)
        count_input = QLineEdit()
        count_input.setText("1")
        count_input.setAlignment(Qt.AlignCenter)
        count_input.setInputMask("999999") 
        self.setCellWidget(row_position, 1, count_input)

        # 3. Колонка: Интерактивный чекбокс (True/False)
        bend_item = QTableWidgetItem()
        bend_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
        bend_item.setCheckState(Qt.Unchecked)
        self.setItem(row_position, 2, bend_item)


class MainWindow(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowIcon(QIcon(r"C:\Users\vlasovpi\Desktop\DOG-Kompas3DBox\DOG-macros\GUI\icons\icon.ico"))
        self.setWindowTitle("Kompas DXF Converter")
        self.setMinimumSize(700, 450)

        layout = QVBoxLayout()

        self.label = QLabel("Список файлов для обработки:")
        layout.addWidget(self.label)

        self.table_widget = FileTableWidget()
        layout.addWidget(self.table_widget)

        btn_add = QPushButton("Добавить файлы")
        btn_add.clicked.connect(self.add_files)
        layout.addWidget(btn_add)

        self.btn_convert = QPushButton("Конвертировать в DXF")
        self.btn_convert.clicked.connect(self.convert)
        layout.addWidget(self.btn_convert)

        self.setLayout(layout)

    def add_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Выбери .m3d", "", "3D Models (*.m3d)"
        )
        for f in files:
            self.table_widget.add_file_row(f)

    def convert(self):
        row_count = self.table_widget.rowCount()
        if row_count == 0:
            print("Список файлов пуст.")
            return

        # Наполняем чистый список данных для отправки в ваш скрипт
        prepared_data = []
        for row in range(row_count):
            file_path = self.table_widget.item(row, 0).text()
            if not file_path or not os.path.exists(file_path):
                continue
            
            count_widget = self.table_widget.cellWidget(row, 1)
            file_count = int(count_widget.text()) if count_widget.text() else 1
            
            bend_item = self.table_widget.item(row, 2)
            has_bending = bend_item.checkState() == Qt.Checked
            
            # Структура: (строка_пути, число_количество, булева_гибка)
            prepared_data.append((file_path, file_count, has_bending))

        print("Запуск конвертации...")
        
        # Передаем подготовленный список данных в вашу функцию
        convert_to_dxf(prepared_data) 
        
        print("Данные отправлены:", prepared_data)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())