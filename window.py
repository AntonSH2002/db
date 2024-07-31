from PyQt5 import QtWidgets
from PyQt5.Qt import *
from PyQt5.QtWidgets import QTableWidgetItem, QMessageBox
from interface import Ui_MainWindow
import pyodbc
import openpyxl


class Window(QtWidgets.QMainWindow):
    def __init__(self):
        super(Window, self).__init__()
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.action_2.triggered.connect(self.select_patients)
        self.ui.action_3.triggered.connect(self.select_doctors)
        self.ui.action_4.triggered.connect(self.select_analysis_types)
        self.ui.action_5.triggered.connect(self.select_analysis_directions)
        self.ui.action_6.triggered.connect(self.save_to_excel)
        self.connection = pyodbc.connect(r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};'
                                         r'DBQ=C:\\Users\\Anton\\PycharmProjects\\task1\\database.accdb')
        self.cursor = self.connection.cursor()

    def tab(self, rows, columns, data):
        self.ui.tableWidget.setRowCount(rows)
        self.ui.tableWidget.setColumnCount(columns)

        for i in range(rows):
            for j in range(columns):
                item = QTableWidgetItem("{}".format(data[i][j]))
                item.setTextAlignment(Qt.AlignHCenter)
                self.ui.tableWidget.setItem(i, j, item)

        self.ui.tableWidget.resizeColumnsToContents()
        self.ui.tableWidget.resizeRowsToContents()

    def fetch_data(self, query):
        try:
            self.cursor.execute(query)
            data = self.cursor.fetchall()
            if not data:
                message_box = QMessageBox()
                message_box.setText("Данные не извлекаются из базы данных")
                message_box.exec_()
                return None
            return data
        except pyodbc.Error as e:
            message_box = QMessageBox()
            message_box.setText(f"Ошибка при извлечении данных: {e}")
            message_box.exec_()
            return None

    def update_table(self, data, column_count, headers):
        self.ui.tableWidget.setRowCount(0)
        self.ui.tableWidget.setColumnCount(column_count)
        self.ui.tableWidget.setHorizontalHeaderLabels(headers)
        if data:
            rows = len(data)
            columns = len(data[0])
            self.tab(rows, columns, data)

    def select_patients(self):
        query = 'SELECT * FROM patients'
        data = self.fetch_data(query)
        if data:
            headers = ("Идентификатор", "Фамилия", "Имя", "Отчество", "Дата рождения", "Пол")
            self.update_table(data, 6, headers)

    def select_doctors(self):
        query = 'SELECT * FROM doctors'
        data = self.fetch_data(query)
        if data:
            headers = ("Идентификатор", "Фамилия", "Имя", "Отчество")
            self.update_table(data, 4, headers)

    def select_analysis_types(self):
        query = 'SELECT * FROM analysis_types'
        data = self.fetch_data(query)
        if data:
            headers = ("Идентификатор", "Название", "Цена")
            self.update_table(data, 3, headers)

    def select_analysis_directions(self):
        query = 'SELECT * FROM analysis_directions'
        data = self.fetch_data(query)
        if data:
            headers = ("Идентификатор", "Идентификатор вида исследования",
                       "Идентификатор пациента", "Идентификатор врача",
                       "Дата назначения", "Дата результата", "Оплата")
            self.update_table(data, 7, headers)

    def save_to_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Сохранить файл", "", "Excel Files (*.xlsx)")
        if not path:
            return

        workbook = openpyxl.Workbook()
        sheet = workbook.active

        headers = [self.ui.tableWidget.horizontalHeaderItem(col).text() for col in
                   range(self.ui.tableWidget.columnCount())]
        sheet.append(headers)

        for row in range(self.ui.tableWidget.rowCount()):
            row_data = [self.ui.tableWidget.item(row, col).text() if self.ui.tableWidget.item(row, col) else '' for col
                        in range(self.ui.tableWidget.columnCount())]
            sheet.append(row_data)

        try:
            workbook.save(path)
            message_box = QMessageBox()
            message_box.setText(f"Файл успешно сохранен в {path}")
            message_box.exec_()
        except Exception as e:
            message_box = QMessageBox()
            message_box.setText(f"Ошибка при сохранении файла: {e}")
            message_box.exec_()
            