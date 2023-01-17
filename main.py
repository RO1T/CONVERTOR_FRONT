import sys
from datetime import datetime
from PyQt5 import uic
from PyQt5.QtWidgets import (QApplication,
                             QDialog,
                             QStackedWidget,
                             QFileDialog,
                             QMessageBox,
                             QDesktopWidget)
from PyQt5.QtGui import QFont, QFontDatabase, QIcon
from PyQt5.QtCore import QAbstractTableModel, Qt
from convertor import Convertor
from ehandler import ExcpetionHandler


class TableModel(QAbstractTableModel):
    def __init__(self, data):
        super(TableModel, self).__init__()
        self._data = data

    def load_data(self, data):
        self.beginResetModel()
        self._data = data
        self.endResetModel()

    def data(self, index, role):
        if role == Qt.DisplayRole:
            value = self._data.iloc[index.row(), index.column()]
            if isinstance(value, datetime):
                # Render time to YYY-MM-DD.
                return value.strftime("%Y-%m-%d")
            if isinstance(value, int):
                # Render time to YYY-MM-DD.
                return str(value)

            if isinstance(value, float):
                # Render float to 2 dp
                return "%.2f" % value

            if isinstance(value, str):
                # Render strings with quotes
                return '%s' % value
            return str(value)

    def rowCount(self, index):
        return self._data.shape[0]

    def columnCount(self, index):
        return self._data.shape[1]

    def headerData(self, section, orientation, role):
        # section is the index of the column/row.
        if role == Qt.DisplayRole:
            if orientation == Qt.Horizontal:
                return str(self._data.columns[section])

            if orientation == Qt.Vertical:
                return str(self._data.index[section])


class WorkWindow(QDialog):
    def __init__(self, excel_path, name_chose, json_path):
        super(WorkWindow, self).__init__()
        self.name_chose = name_chose
        self.json_path = json_path
        self.ehandler = ExcpetionHandler()
        uic.loadUi('dialog_3.ui', self)
        QFontDatabase.addApplicationFont("font/Gilroy-Regular.ttf")
        # данные с екселя
        self.excel_path = excel_path
        self.convertor = Convertor(self.excel_path, self.json_path)
        df_input = self.convertor.original
        df_result = self.convertor.result
        # кнопки
        self.original_clear.clicked.connect(self.clear_orig)
        self.result_clear.clicked.connect(self.clear_res)
        self.apply_btn.clicked.connect(self.apply_changes)
        self.change_file_btn.clicked.connect(self.change_file_func)
        self.go_to_download_btn.clicked.connect(self.changer)
        self.msg = QMessageBox()
        # таблица исходная
        self.model_original = TableModel(df_input)
        self.table_before.setModel(self.model_original)
        self.table_before.horizontalHeader().sectionClicked.connect(self.click_handler_original)
        # таблица итоговая
        self.model_result = TableModel(df_result)
        self.table_after.setModel(self.model_result)
        self.table_after.horizontalHeader().sectionClicked.connect(self.click_handler_result)

    def download_fun(self):
        if self.name_chose == 'exel':
            self.path = QFileDialog.getSaveFileName(self, f"Куда сохранить файл?", "",
                                                    "Excel (*.xlsx *.xls)")
            self.file_name = self.path[0].split('/')[-1]
            self.file_path_abs = self.path[0]
            if self.file_path_abs == '':
                button = self.call_error()
                if button != QMessageBox.No:
                    self.download_fun()
            else:
                self.convertor.to_exel(self.path[0])
        elif self.name_chose == 'json':
            self.path = QFileDialog.getSaveFileName(self, f"Куда сохранить файл?", "",
                                                    "Json (*.json)")
            self.file_name = self.path[0].split('/')[-1]
            self.file_path_abs = self.path[0]
            if self.file_path_abs == '':
                button = self.call_error()
                if button != QMessageBox.No:
                    self.download_fun()
            else:
                self.convertor.to_json(self.path[0])

    def call_error(self):
        return self.ehandler.warning_choice_msg('Ошибка', 'Вы должны выбрать куда загружать файл!')

    def not_implemented_alert(self, message):
        return self.ehandler.critical_msg('Ошибка', message)

    def empty_column_warning_no_yes(self, message):
        button = self.ehandler.warning_choice_msg('Ошибка', message)
        if button == QMessageBox.Yes:
            self.download_fun()

    def change_filled_warning_no_yes(self, message):
        button = self.ehandler.warning_choice_msg('Ошибка', message)
        if button == QMessageBox.Yes:
            self.apply()

    def clear_orig(self):
        self.original.setText('')

    def clear_res(self):
        self.result.setText('')

    def click_handler_original(self, e):
        column_text = self.model_original.headerData(e, Qt.Horizontal, Qt.DisplayRole)
        self.original.setText(self.original.text() + column_text + ', ')

    def click_handler_result(self, e):
        column_text = self.model_result.headerData(e, Qt.Horizontal, Qt.DisplayRole)
        self.result.setText(self.result.text() + column_text + ', ')

    def apply_changes(self):
        command = self.get_command()
        no_split = command[0] == 'SPLIT' and (len(command[1]) > 1 or len(command[2]) == 1)
        no_rename = command[0] == 'RENAME' and (len(command[1]) > 1 or len(command[2]) > 1)
        no_zip = command[0] == 'ZIP' and (len(command[1]) == 1 or len(command[2]) > 1)
        change_filled = False
        for column in command[2]:
            if column in self.convertor.between.columns:
                change_filled = True
        if command[0] == '':
            self.not_implemented_alert('Вы ничего не сделали!')
        elif no_rename:
            self.not_implemented_alert(
                'Для выполнения команды RENAME в каждой таблице выберите только по одному столбцу!')
        elif no_split:
            self.not_implemented_alert('Для выполнения команды SPLIT в  исходной таблице выберите только один стобец!')
        elif no_zip:
            self.not_implemented_alert('Для выполнения команды ZIP в  итоговой таблице выберите только один стобец!')
        elif change_filled:
            self.change_filled_warning_no_yes(
                'Вы пытаетесь изменить уже заполненную колонку, уверены, что хотите продолжить?')
        else:
            self.apply()

    def apply(self):
        try:
            command = self.get_command()
            self.convertor.execute(command)
            self.model_result.load_data(self.convertor.result)
        except:
            self.not_implemented_alert('Данную функцию нельзя выполнить')
        finally:
            self.clear_res()
            self.clear_orig()

    def get_command(self):
        return self.command.currentText(), self.original.text()[:-2].split(', '), self.result.text()[:-2].split(', ')

    def change_file_func(self):
        widgets.setCurrentIndex(widgets.currentIndex() - 1)
        widgets.removeWidget(self)

    def changer(self):
        if self.convertor.have_empty_columns():
            self.empty_column_warning_no_yes('У вас остались незаполненные колонки, уверены, что хотите продолжить?')
        else:
            self.download_fun()


class InputWindow(QDialog):
    def __init__(self, name_chose):
        super(InputWindow, self).__init__()
        self.name_chose = name_chose
        # если джсон не нужен, но он нужен, крч надо тут
        self.json_path = None
        self.ehandler = ExcpetionHandler()
        QFontDatabase.addApplicationFont("font/Gilroy-Regular.ttf")
        font = QFont('Gilroy')

        if self.name_chose == 'exel':
            uic.loadUi('dialog_2.ui', self)
            self.change_btn.setFont(font)
            self.input_btn.setFont(font)
            font_underline = self.change_btn.font()
            font_underline.setUnderline(True)
            self.change_btn.setFont(font_underline)
            self.change_btn.setStyleSheet("background-color: white")
            self.input_btn.clicked.connect(self.input_func)
            self.change_btn.clicked.connect(self.change_func)
        elif self.name_chose == 'json':
            uic.loadUi('dialog_2.1.ui', self)
            self.change_btn.setFont(font)
            self.input_btn_ex.setFont(font)
            self.input_btn_js.setFont(font)
            self.label.setFont(font)
            font_underline = self.change_btn.font()
            font_underline.setUnderline(True)
            self.change_btn.setFont(font_underline)
            self.change_btn.setStyleSheet("background-color: white")

            self.input_btn_ex.clicked.connect(self.input_func)
            self.input_btn_js.clicked.connect(self.input_json_func)

            self.change_btn.clicked.connect(self.change_func)

            self.next_btn.clicked.connect(self.next_wind)

    def next_wind(self):
        try:
            self.work_w = WorkWindow(self.excel_path, self.name_chose, self.json_path)
            widgets.addWidget(self.work_w)
            widgets.setCurrentIndex(widgets.currentIndex() + 1)
        except (ValueError, FileNotFoundError):
            self.ehandler.critical_msg('Ошибка', 'Не правильный исходный файл!')

    def input_json_func(self):
        self.json_path = QFileDialog.getOpenFileName(self, f"Выберите файл {self.name_chose}", "",
                                                     "Json (*.json)")[0]
        if self.json_path == '':
            button = self.ehandler.warning_choice_msg('Ошибка',
                                                      'Вы должны выбрать файл!\nЕсли вы передумали, нажмите No.')
            if button != QMessageBox.No:
                self.input_json_func()

    def input_func(self):
        self.excel_path = QFileDialog.getOpenFileName(self, f"Выберите файл {self.name_chose}", "",
                                                      "Excel (*.xlsx *.xls)")[0]
        if self.excel_path == '':
            button = self.ehandler.warning_choice_msg('Ошибка',
                                                      'Вы должны выбрать файл!\nЕсли вы передумали, нажмите No.')
            if button != QMessageBox.No:
                self.input_func()
        else:
            if self.name_chose == 'exel':
                self.next_wind()

    def change_func(self):
        widgets.setCurrentIndex(widgets.currentIndex() - 1)
        widgets.removeWidget(self)


class MainWindow(QDialog):
    def __init__(self):
        super().__init__()
        uic.loadUi('dialog_1.ui', self)
        QFontDatabase.addApplicationFont("font/Gilroy-Regular.ttf")
        font = QFont('Gilroy')
        self.ehandler = ExcpetionHandler()
        labels = [self.label_up, self.label_down]
        for label in labels:
            label.setFont(font)
        buttons = [self.exel_exel_btn, self.exel_word_btn, self.exel_json_btn]
        for btn in buttons:
            btn.setFont(font)
        self.exel_exel_btn.clicked.connect(self.exel_exel_btn_fun)
        self.exel_word_btn.clicked.connect(self.exel_word_btn_fun)
        self.exel_json_btn.clicked.connect(self.exel_json_btn_fun)

    def exel_exel_btn_fun(self):
        self.input_w = InputWindow('exel')
        widgets.addWidget(self.input_w)
        widgets.setCurrentIndex(widgets.currentIndex() + 1)

    def exel_word_btn_fun(self):
        self.not_implemented_alert()

    def exel_json_btn_fun(self):
        self.input_w = InputWindow('json')
        widgets.addWidget(self.input_w)
        widgets.setCurrentIndex(widgets.currentIndex() + 1)

    def not_implemented_alert(self):
        self.ehandler.critical_msg('Ошибка', 'В процессе разработки!')


if __name__ == "__main__":
    app = QApplication(sys.argv)
    widgets = QStackedWidget()

    main_w = MainWindow()
    widgets.addWidget(main_w)

    widgets.setGeometry(main_w.geometry())

    widgets.setWindowTitle('Конвертор')
    widgets.setWindowIcon(QIcon('logo.png'))
    # widgets.setMaximumSize(1160, 591)
    qtRectangle = widgets.frameGeometry()
    centerPoint = QDesktopWidget().availableGeometry().center()
    qtRectangle.moveCenter(centerPoint)
    widgets.move(qtRectangle.topLeft())

    widgets.show()

    sys.exit(app.exec_())
