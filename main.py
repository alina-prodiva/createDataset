import sys
import pandas as pd
import xlsxwriter
from PyQt6.QtWidgets import (
    QWidget, QHBoxLayout, QVBoxLayout, QLabel, QApplication, QTableWidget, QMessageBox,
    QTableWidgetItem, QPushButton, QFileDialog, QTextEdit, QComboBox, QLineEdit)
from PyQt6.QtCore import Qt


class MainWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setGeometry(100, 50, 1400, 700)
        self.setWindowTitle('Обезличивание персональных данных')

        # Основной горизонтальный макет (слева — таблица, справа — настройки)
        self.layout = QHBoxLayout()
        self.setLayout(self.layout)

        # 1. Левая часть — таблица
        self.table_layout = QVBoxLayout()
        self.table_label = QLabel("Табличные данные")
        self.table_label.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.table_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.table_layout.addWidget(self.table_label)
        self.table = QTableWidget()
        self.table_layout.addWidget(self.table)
        self.layout.addLayout(self.table_layout, 2)

        # 3. Правая часть настройки
        self.settings_layout = QVBoxLayout()
        self.settings_label = QLabel("Настройки")
        self.settings_label.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.settings_label.setStyleSheet("font-weight: bold; font-size: 14px;")
        self.settings_layout.addWidget(self.settings_label)

        self.button_select = QPushButton("Выберите файл")
        self.button_select.clicked.connect(self.on_button_select)
        self.settings_layout.addWidget(self.button_select)

        self.filePath = QLineEdit()
        self.settings_layout.addWidget(self.filePath)

        self.ids = ['Идентификатор', 'Сущность', 'Объект', 'Документ', 'Адрес', 'Описание объекта', 'Информация']

        la_id_spr = QLabel('Справочник вариантов идентификаторов')
        self.id_spr = QComboBox()
        self.id_spr.addItems([f'{i + 1} - {item}' for i, item in enumerate(self.ids)])
        self.settings_layout.addWidget(la_id_spr)
        self.settings_layout.addWidget(self.id_spr)

        la_cols = QLabel('Укажите номер столбца - номер варианта идентификатора, перечисление через запятую')
        self.cols = QTextEdit()
        self.settings_layout.addWidget(la_cols)
        self.settings_layout.addWidget(self.cols)

        self.button_renew = QPushButton("Обезличить данные")
        self.button_renew.clicked.connect(self.on_button_renew)
        self.settings_layout.addWidget(self.button_renew)

        self.button_export = QPushButton("Экспорт данных")
        self.button_export.clicked.connect(self.on_button_export)
        self.settings_layout.addWidget(self.button_export)

        self.export_text = QTextEdit()
        self.settings_layout.addWidget(self.export_text)

        self.settings_layout.addSpacing(400)
        self.layout.addLayout(self.settings_layout)

    def on_button_select(self):
        self.file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите файл",
            "",  # начальная директория (пустая = домашняя)
            "Табличные данные (*.xlsx *.xls *.csv);;Все файлы (*.*)"
        )
        if self.file_path:
            self.filePath.setText(self.file_path)
            if self.file_path.endswith('.csv'):
                self.df = pd.read_csv(self.file_path)
                self.data = self.df.values.tolist()
            elif self.file_path.endswith(('.xlsx', '.xls')):
                self.df = pd.read_excel(self.file_path)
                self.data = self.df.values.tolist()
            else:
                raise ValueError("Поддерживаются только файлы CSV и Excel (.csv, .xlsx, .xls)")
        else:
            return

        if self.data:
            # Заполняем таблицу
            self.table.clear()
            self.table.setRowCount(self.df.shape[0])
            self.table.setColumnCount(self.df.shape[1])
            self.table.setHorizontalHeaderLabels([f'{i}, {str(x)}' for i, x in enumerate(list(self.df.columns.values), 1)])

            for i, row in enumerate(self.data):
                for j, elem in enumerate(row):
                    self.table.setItem(i, j, QTableWidgetItem(str(elem)))
        self.update()

    def on_button_renew(self):
        self.new_data = self.data[:]
        try:
            cols = [x.strip() for x in self.cols.toPlainText().split(',')]
            for line in cols:
                ii, jj = line.split('-')
                for i, row in enumerate(self.data):
                    for j, elem in enumerate(row):
                        if j == int(ii) - 1:
                            self.table.setItem(i, j, QTableWidgetItem(f'{self.ids[int(jj) - 1]} {i + 1}'))
                            self.new_data[i][j] = f'{self.ids[int(jj) - 1]} {i + 1}'
            dl = QMessageBox()
            dl.setWindowTitle('Сообщение')
            dl.setText('Процедура обезличивания окончена!')
            dl.exec()
        except Exception:
            dl = QMessageBox()
            dl.setWindowTitle('Внимание!')
            dl.setText('Задан неверный формат обезличивания!')
            dl.exec()
        self.update()

    def on_button_export(self):
        file1 = 'dataset_' + (self.filePath.text().split('/')[-1]).split('.')[0] + '.xlsx'
        wb1 = xlsxwriter.Workbook(file1)
        w1 = wb1.add_worksheet()

        for i, x in enumerate(list(self.df.columns.values)):
            w1.write(0, i, str(x))
        for i, row in enumerate(self.new_data, 1):
            for j, elem in enumerate(row):
                w1.write(i, j, str(elem))

        wb1.close()
        self.export_text.setText(f'Результат автоматизированного обезличивания:\nОбезличенная таблица - {file1}')
        dl = QMessageBox()
        dl.setWindowTitle('Сообщение')
        dl.setText(f'Создан файл {file1} с обезличенными данными!')
        dl.exec()
        self.update()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    win = MainWidget()
    win.show()
    sys.exit(app.exec())