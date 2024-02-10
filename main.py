import os
import sys
import datetime
import numpy as np
import plotly.offline as offline
import pandas as pd
import plotly.express as px
import matplotlib as mp
import sympy as sp
from PyQt6.QtCore import *
from PyQt6.QtSvgWidgets import QSvgWidget
from PyQt6.QtWebEngineWidgets import QWebEngineView
from PyQt6.QtWidgets import *
from PyQt6.QtGui import *
from matplotlib_inline.backend_inline import FigureCanvas
from openpyxl.reader.excel import load_workbook

from forecast import main_import_file
from matplotlib.figure import Figure
import utilities as util


class FileSelectionApp(QMainWindow):
    """Main window of this program"""

    def __init__(self):
        super(FileSelectionApp, self).__init__(None)
        self.selected_file = None  # Объявляем selected_file как атрибут экземпляра класса
        self.selected_sheet = None  # Объявляем selected_sheet
        self.setWindowTitle('ADL Model')
        self.resize(1200, 800)
        # Создаем вертикальный лейаут для размещения виджетов
        layout = QVBoxLayout()

        # Создаем QLabel для отображения выбранного файла
        self.file_label = QLabel('Выберите файл:')
        layout.addWidget(self.file_label)

        # Создаем кнопку для выбора листа
        self.select_file_button = QPushButton('Выбрать файл')
        self.select_file_button.clicked.connect(self.open_file_and_choose_sheet)
        layout.addWidget(self.select_file_button)

        # Создаем кнопку для выбора листа
        self.select_sheet_button = QPushButton('Выбрать лист')
        self.select_sheet_button.setEnabled(False)
        self.select_sheet_button.clicked.connect(self.choose_excel_sheet)
        layout.addWidget(self.select_sheet_button)

        # Создаем кнопку для создания прогноза
        self.create_prediction_on_chosen_file = QPushButton('Создать прогноз')
        self.create_prediction_on_chosen_file.setEnabled(False)
        # Используем lambda-функцию, чтобы передать параметр
        self.create_prediction_on_chosen_file.clicked.connect(self.create_short_term_prediction)
        layout.addWidget(self.create_prediction_on_chosen_file)

        # Создаем кнопку для создания графика
        self.create_prediction_graph = QPushButton('Создать график')
        self.create_prediction_graph.setEnabled(False)
        self.create_prediction_graph.clicked.connect(self.choose_column_name_for_plot)
        layout.addWidget(self.create_prediction_graph)

        # Создаем QLabel для отображения выбранного листа
        self.sheet_label = QLabel('Выбранный лист:')
        layout.addWidget(self.sheet_label)

        # Создаем QTableWidget для отображения данных
        self.table_widget = QTableWidget()
        layout.addWidget(self.table_widget)

        # Создаем виджет и устанавливаем наш лейаут
        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        # Добавляем меню в верхнюю панель
        menu_bar = self.menuBar()

        # Создаем QAction "File" с выпадающим меню
        file_menu_action = QAction("File", self)
        file_menu = QMenu(self)

        # Добавляем в меню кнопку "Выбрать файл"
        select_file_action = QAction("Выбрать файл", self)
        select_file_action.triggered.connect(self.open_file_and_choose_sheet)
        file_menu.addAction(select_file_action)

        # Создаем QWebEngineView для отображения графика
        self.webview = QWebEngineView()
        layout.addWidget(self.webview)

        file_menu_action.setMenu(file_menu)
        menu_bar.addAction(file_menu_action)

    def create_prediction(self):
        if self.selected_file and self.selected_sheet:
            main_import_file(self.selected_file, self.selected_sheet)

    def open_file_and_choose_sheet(self):
        try:
            # Открываем диалог выбора файла и фильтра для Excel
            file_dialog = QFileDialog(self)
            file_dialog.setNameFilter("Excel Files (*.xlsx)")
            selected_file, _ = file_dialog.getOpenFileName(self, 'Выберите файл', '', 'Excel Files (*.xlsx)')

            # Обновляем QLabel с выбранным файлом
            self.file_label.setText(f'Выбранный файл: {selected_file}')

            # Выбираем лист Excel, если это файл Excel
            if selected_file.endswith('.xlsx'):
                self.selected_file = selected_file
                self.select_sheet_button.setEnabled(True)
                self.create_prediction_on_chosen_file.setEnabled(True)
                self.create_prediction_graph.setEnabled(True)
                self.choose_excel_sheet()

        except Exception as e:
            print(f"Ошибка при открытии файла: {e}")

    def choose_excel_sheet(self):
        try:
            workbook = load_workbook(filename=self.selected_file, read_only=True)
            sheet_names = workbook.sheetnames

            # Используем QInputDialog для выбора листа
            input_dialog = QInputDialog(self)
            input_dialog.setWindowTitle('Выбор листа')
            input_dialog.setComboBoxItems(sheet_names)
            input_dialog.setComboBoxEditable(False)

            # Устанавливаем фиксированный размер
            input_dialog.setFixedSize(400, 200)

            # Отображаем диалоговое окно
            ok_pressed = input_dialog.exec()

            if ok_pressed:
                # Очищаем таблицу перед отображением новых данных
                self.table_widget.clear()
                self.selected_sheet = input_dialog.textValue()
                # Загружаем данные из выбранного листа
                data = pd.read_excel(self.selected_file, self.selected_sheet, engine='openpyxl')
                self.display_data_in_table(data)

                # Обновляем QLabel с выбранным листом
                self.sheet_label.setText(f'Выбранный лист: {self.selected_sheet}')

        except Exception as e:
            print(f"Ошибка при загрузке данных: {e}")

    def display_data_in_table(self, data):
        try:
            # Устанавливаем количество строк и столбцов
            self.table_widget.setRowCount(data.shape[0])
            self.table_widget.setColumnCount(data.shape[1])

            # Устанавливаем заголовки столбцов
            self.table_widget.setHorizontalHeaderLabels(data.columns)

            # Заполняем ячейки данными
            for i in range(data.shape[0]):
                for j in range(data.shape[1]):
                    item = QTableWidgetItem(str(data.iloc[i, j]))
                    self.table_widget.setItem(i, j, item)

        except Exception as e:
            print(f"Ошибка при отображении данных в таблице: {e}")

    def choose_column_name_for_plot(self):
        try:
            if self.selected_file and self.selected_sheet:
                column_names = [self.table_widget.horizontalHeaderItem(j).text() for j in
                                range(self.table_widget.columnCount())]

                # Создаем новый экземпляр DataSelectionDialog
                data_selection_dialog = DataSelectionDialogPlot(column_names=column_names,
                                                                selected_file=self.selected_file,
                                                                selected_sheet=self.selected_sheet)
                data_selection_dialog.exec()
        except Exception as e:
            print(f"Ошибка при запуске диалогового окна выбора осей графика: {e}")

    def create_short_term_prediction(self):
        try:
            if self.selected_file and self.selected_sheet:
                column_names = [self.table_widget.horizontalHeaderItem(j).text() for j in
                                range(self.table_widget.columnCount())]

                # Создаем новый экземпляр DataSelectionDialog
                data_selection_dialog = ForecastWindow(column_names=column_names, selected_file=self.selected_file,
                                                       selected_sheet=self.selected_sheet)
                data_selection_dialog.exec()
        except Exception as e:
            print(f"Ошибка при запуске диалогового окна выбора настройки создания прогноза: {e}")


class ForecastWindow(QDialog):
    """In this window created predicts"""

    def __init__(self, column_names, selected_file, selected_sheet):
        super().__init__()
        self.count_factors = None
        self.column_names = column_names
        self.selected_file = selected_file
        self.selected_sheet = selected_sheet
        self.setWindowTitle('Выбор данных для создания прогноза')
        self.adjustSize()

        layout = QVBoxLayout()

        # Добавляем QLabel и QComboBox для выбора данных по оси X
        self.label_1 = QLabel('Выберите столбец, для которого необходимо составить прогнозную модель:')
        self.menu_data_prediction_name_column = QComboBox()
        self.menu_data_prediction_name_column.addItems(column_names)

        # Добавляем QLabel и QComboBox для выбора данных по оси X
        self.label_2 = QLabel('Выберите столбец данных с временными метками:')
        self.menu_time_label_name_column = QComboBox()
        self.menu_time_label_name_column.addItems(column_names)
        self.label_3 = QLabel('Выберите до 3-ёх факторов прогнозной модели:')
        self.label_4 = QLabel("Выберите размер тестовой выборки:")

        self.slider_box = QSlider()
        self.slider_box.setMinimum(0)
        self.slider_box.setMaximum(100)
        self.slider_box.setOrientation(Qt.Orientation.Horizontal)
        self.slider_box.setValue(66)
        self.percent_label = QLabel("Обучающая выборка: 66% Тестовая выборка: 34%")
        self.slider_box.valueChanged.connect(self.slider_value_changed)

        # Добавляем QTableWidget для выбора факторов
        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(1)
        self.table_widget.setRowCount(len(column_names))
        self.table_widget.setHorizontalHeaderLabels(['Выбрать столбец'])
        self.table_widget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table_widget.itemChanged.connect(self.handle_item_changed)
        self.table_widget.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

        for i, column_name in enumerate(column_names):
            item = QTableWidgetItem(column_name)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Unchecked)
            self.table_widget.setItem(i, 0, item)

        # Кнопки "OK" и "Отмена"
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout.addWidget(self.label_1)
        layout.addWidget(self.menu_data_prediction_name_column)
        layout.addWidget(self.label_2)
        layout.addWidget(self.menu_time_label_name_column)
        layout.addWidget(self.label_3)
        layout.addWidget(self.table_widget)
        layout.addWidget(self.label_4)
        layout.addWidget(self.slider_box)
        layout.addWidget(self.percent_label)
        layout.addWidget(buttons)

        self.setLayout(layout)

    def slider_value_changed(self, value):
        self.percent_label.setText(f"Обучающая выборка: {value}% Тестовая выборка: {100-value}%")

    def handle_item_changed(self):
        checked_count = (1 for row in range(self.table_widget.rowCount()) if
                         self.table_widget.item(row, 0).checkState() == Qt.CheckState.Checked)

        try:
            checked_count_sum = sum(checked_count)
            if checked_count_sum > 3:
                for row in range(self.table_widget.rowCount()):
                    item = self.table_widget.item(row, 0)
                    item.setCheckState(Qt.CheckState.Unchecked)
        except Exception as e:
            print(f"Exception occurred: {e}")

    def accept(self):
        chosen_column_for_predict = self.menu_data_prediction_name_column.currentText()
        chosen_column_with_time_label = self.menu_time_label_name_column.currentText()
        selected_columns = []
        for row in range(self.table_widget.rowCount()):
            item = self.table_widget.item(row, 0)
            if item.checkState() == Qt.CheckState.Checked:
                selected_columns.append(item.text())
        prepared_data = util.data_preparation(self.selected_file, self.selected_sheet,
                                              chosen_column_with_time_label, selected_columns)
        data = util.create_lags(prepared_data, selected_columns, 1)
        data_learn, data_test = util.separation_data(data, self.slider_box.value())
        print(f"{len(data_learn)} - Длинна обучающей выборки ")
        print(f"{len(data_test)} - Длинна тестовой выборки ")
        model = util.create_model(data_learn, chosen_column_for_predict, selected_columns, 1)
        result_data = util.predict_on_params(data, model.params, self.slider_box.value(), chosen_column_for_predict)
        print(util.calculate_mape(result_data, chosen_column_for_predict))
        print(util.write_to_excel(result_data, f"Prediction"))
        super().accept()


class DataSelectionDialogPlot(QDialog):
    """In this window is happening choosing data for creating predict model"""
    def __init__(self, column_names, selected_file, selected_sheet):
        super().__init__()
        self.selected_file = selected_file
        self.selected_sheet = selected_sheet
        self.setWindowTitle('Выбор данных для графика')
        self.column_names = column_names

        layout = QVBoxLayout()

        # Добавляем QLabel и QLineEdit для выбора данных по оси X
        self.label_x = QLabel('Выберите данные для оси X:')
        self.menu_x = QComboBox()
        self.menu_x.addItems(column_names)
        self.menu_x.currentIndexChanged.connect(self.update_table_widget)
        layout.addWidget(self.label_x)
        layout.addWidget(self.menu_x)

        # Добавляем QLabel и QTableWidget для выбора данных по оси Y
        self.label_y = QLabel('Выберите данные для оси Y:')
        self.table_widget = QTableWidget()
        self.table_widget.setColumnCount(1)
        self.table_widget.setRowCount(len(column_names)-1)
        self.table_widget.setHorizontalHeaderLabels(['Выбрать столбец'])
        self.table_widget.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table_widget.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)

        list_columns_name = list(column_names)
        list_columns_name.remove(self.menu_x.currentText())

        for i, column_name in enumerate(list_columns_name):
            item = QTableWidgetItem(column_name)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Unchecked)
            self.table_widget.setItem(i, 0, item)

        layout.addWidget(self.label_y)
        layout.addWidget(self.table_widget)

        # Кнопки "OK" и "Отмена"
        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

        self.setLayout(layout)

    def get_selected_data(self):
        # Возвращаем выбранные данные по осям X и Y
        selected_x = self.menu_x.currentText()
        selected_y = [self.table_widget.item(i, 0).text() for i in range(self.table_widget.rowCount())
                      if self.table_widget.item(i, 0).checkState() == Qt.CheckState.Checked]
        return selected_x, selected_y

    def update_table_widget(self):
        selected_x = self.menu_x.currentText()
        list_columns_name = list(self.column_names)
        list_columns_name.remove(selected_x)
        # Очистка содержимого table_widget
        self.table_widget.clearContents()

        # Заполнение table_widget данными, исключая выбранный в menu_x
        for i, column_name in enumerate(list_columns_name):
            if column_name != selected_x:
                item = QTableWidgetItem(column_name)
                item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Unchecked)
                self.table_widget.setItem(i, 0, item)

    def get_selected_data(self):
        # Возвращаем выбранные данные по осям X и Y
        selected_x = self.menu_x.currentText()
        selected_y = [self.table_widget.item(i, 0).text() for i in range(self.table_widget.rowCount())
                      if self.table_widget.item(i, 0).checkState() == Qt.CheckState.Checked]
        return selected_x, selected_y

    def accept(self):
        selected_x, selected_y = self.get_selected_data()
        print(selected_y)

        if selected_x and selected_y:
            # Создаем новый экземпляр PlotWindow с выбранными данными
            data = pd.read_excel(self.selected_file, sheet_name=self.selected_sheet)
            plot_window = PlotWindow(data=data, x_column=selected_x, y_columns=selected_y)
            super().accept()
            plot_window.exec()


class PlotWindow(QDialog):
    """In this window is happening choosing data for creating plot data"""

    def __init__(self, data, x_column, y_columns):
        super().__init__()
        self.setWindowTitle('График')
        self.resize(1200, 800)
        try:
            # Построение графика с несколькими y
            fig = px.line(data, x=x_column, y=y_columns, title='Название графика')

            # Создаем html-код фигуры
            html = '<html><body>'
            html += offline.plot(fig, output_type='div', include_plotlyjs='cdn')
            html += '</body></html>'

            # Создаем экземпляр QWebEngineView и устанавливаем html-код
            plot_widget = QWebEngineView()
            plot_widget.setHtml(html)

            # Размещаем QWebEngineView в макете
            layout = QVBoxLayout()
            layout.addWidget(plot_widget)
            self.setLayout(layout)

        except Exception as e:
            print(f"Ошибка при отображении графика: {e}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = FileSelectionApp()
    window.show()
    sys.exit(app.exec())
