from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QLabel

app = QApplication([])

# Основное окно
main_window = QWidget()

# Главный вертикальный макет
main_layout = QVBoxLayout(main_window)

# Вертикальные макеты для столбцов
column_layout1 = QVBoxLayout()
column_layout2 = QVBoxLayout()

# Добавление элементов в столбцы
column_layout1.addWidget(QLabel("Элемент 1"))
column_layout1.addWidget(QLabel("Элемент 2"))

column_layout2.addWidget(QLabel("Элемент 3"))
column_layout2.addWidget(QLabel("Элемент 4"))

# Добавление столбцов в главный макет
main_layout.addLayout(column_layout1)
main_layout.addLayout(column_layout2)

main_window.show()
app.exec_()