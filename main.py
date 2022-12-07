import os.path
import time
from docxcompose.composer import Composer
import docxcompose

from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QGridLayout, QMessageBox, QFontDialog
from PyQt6.QtWidgets import (
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDateTimeEdit,
    QDial,
    QDoubleSpinBox,
    QFontComboBox,
    QLabel,
    QLCDNumber,
    QLineEdit,
    QMainWindow,
    QProgressBar,
    QPushButton,
    QRadioButton,
    QSlider,
    QSpinBox,
    QTimeEdit,
    QVBoxLayout,
    QWidget,
)
from PyQt6.QtCore import QSize, Qt
from PyQt6.QtCore import *
from PyQt6.QtGui import *
from PyQt6 import QtGui
from pathlib import Path

import sys
from random import choice

import Excel_redactor

def get_file_extension(fullpath):
    if fullpath == "":
        print("No file selected")
        return 1
    else:
        if fullpath.find('.xls') == -1:
            print("Wrong file selected")
            return 2
        else:
            print("Input file - ok.")
            return 0

# class PopUp(QWidget):
#     def __init__(self):
#         QWidget.__init__(self)
#         self.event()
#
#     def Event(self):
#         msg = QMessageBox()
#         msg.setWindowTitle("Ошибка")
#         msg.setText("Вы выбрали файл неверного расширения. Допустимые файлы *.doc и *.docx")
#         msg.setIcon(QMessageBox.Critical)

class MainWindow(QMainWindow):
    def __init__(self):

        super().__init__()
        self.setupUi()
        self.file_name  = ""
        self.file_name2 = ""
        self.times_clk_btn = 0



    def setupUi(self):

        self.resize(800, 150)

        #print(QFontDialog.getFont())

        self.text_bar1 = QLabel()
        # self.text_bar1.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        #self.text_bar1.adjustSize()
        self.text_bar1.setHidden(True)

        self.text_bar2 = QLabel()
        # self.text_bar2.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
        #self.text_bar2.adjustSize()
        self.text_bar2.setHidden(True)

        # self.cancel_btn_1 = QPushButton("Отмена")
        # self.cancel_btn_1.adjustSize()

        self.setWindowIcon(QIcon("Graphicloads-Filetype-Excel-xls.ico"))
        self.setWindowTitle("Выгрузка отчета риска")

        self.button = QPushButton("Загрузить файл типа 'Scens'")
        self.button.clicked.connect(self.the_button_was_clicked)
        self.button.adjustSize()

        self.button1 = QPushButton("Загрузить файл типа 'Зоны'")
        self.button1.clicked.connect(self.the_button_was_clicked_second)
        self.button1.adjustSize()

        self.button2 = QPushButton("Переформатировать и выгрузить")
        self.button2.clicked.connect(self.btn2_was_clicked)
        self.button2.adjustSize()

        self.button3 = QPushButton("Переформатировать и выгрузить")
        self.button3.clicked.connect(self.btn3_was_clicked)
        self.button3.adjustSize()

        self.progressbar = QProgressBar()
        self.progressbar.setValue(0)
        self.progressbar.adjustSize()

        self.progressbar2 = QProgressBar()
        self.progressbar2.setValue(0)
        self.progressbar2.adjustSize()

        self.input1 = QLineEdit("")
        self.input1.setPlaceholderText("Введите путь до файла: C:\\user\\file.xlsx")
        self.input2 = QLineEdit("")
        self.input2.setPlaceholderText("Введите путь до файла: C:\\user\\file.xlsx")

        layout = QGridLayout()

        layout.addWidget(self.button, 0, 0)
        layout.addWidget(self.input1, 0, 1)
        layout.addWidget(self.button2, 0, 2)
        layout.addWidget(self.text_bar1, 2, 0)
        # layout.addWidget(self.cancel_btn_1, 2, 2)
        # layout.addWidget(self.label1, 0, 2)
        # layout.addWidget(self.label1, 0, 3)

        layout.addWidget(self.button1, 1, 0)
        layout.addWidget(self.input2, 1, 1)
        layout.addWidget(self.button3, 1, 2)
        layout.addWidget(self.text_bar2, 3, 0)
        #layout.addWidget(self.label1, 1, 0)

        #layout.addWidget(self.label1, 0, 0)
        # layout.addWidget(self.label_jif, 0, 3)
        # layout.addWidget(self.label_jif2, 1, 3)

        layout.addWidget(self.progressbar, 2, 1, 1, 2)
        self.progressbar.setHidden(True)
        layout.addWidget(self.progressbar2, 3, 1, 1, 2)
        self.progressbar2.setHidden(True)


        layout.setRowStretch(4, 1)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)


    def the_button_was_clicked_second(self):
        print("Clicked. second")
        fname = QFileDialog.getOpenFileName(self, 'Выбор Excel файла', 'f:\\', "Excel файлы (*.xls *.xlsx)")
        check = get_file_extension(fname[0])
        if check == 1:
            self.input2.clear()
            print()
            # nothing
        if check == 2:
            self.input2.clear()
            QMessageBox.critical(self, 'Ошибка', 'Вы выбрали файл неверного расширения. Допустимые файлы *.xls и *.xlsx')
            #pop-up wrong file
        if check == 0:
            self.input2.setText(fname[0])
        return fname

    def the_button_was_clicked(self):
        print("Clicked. first")
        fname = QFileDialog.getOpenFileName(self, 'Выбор Excel файла', 'f:\\', "Excel файлы (*.xls *.xlsx)")
        check = get_file_extension(fname[0])
        if check == 1:
            self.input1.clear()
            print()
            # nothing
        if check == 2:
            self.input1.clear()
            msg = QMessageBox()
            msg.setWindowTitle("Ошибка")
            msg.setText("Вы выбрали файл неверного расширения. Допустимые файлы *.xls и *.xlsx")
            msg.setIcon(QMessageBox.Icon.Critical)
            a = msg.exec()
            #QMessageBox.critical(self, 'Ошибка', 'Вы выбрали файл неверного расширения. Допустимые файлы *.doc и *.docx')
            #pop-up wrong file
        if check == 0:
            self.input1.setText(fname[0])
        return fname

        # saveFile = QFileDialog.getSaveFileName(self, 'Save File', "Новый отчет.docx)")
        # print("save = " + str(saveFile[0]))
        # file = open(saveFile[0], 'w')
        # file.write("1111")
        # file.close()
    def btn2_was_clicked(self):
        self.times_clk_btn += 1
        if self.times_clk_btn == 1:
            msg = QMessageBox()
            msg.setWindowTitle("Внимание!")
            msg.setText("Для корректной работы программы, необходимо,"
                        " чтобы выбранный файл для word отчета был закрыт, если вы хотите его перезаписать.")
            msg.setIcon(QMessageBox.Icon.Information)
            a = msg.exec()
        if get_file_extension(self.input1.text()) == 0:
            if Path(self.input1.text()).exists():
                saveFile = QFileDialog.getSaveFileName(self, 'Save File', "Новый отчет.docx", "Word файл (*.docx)")
                file_name = os.path.basename(Path(saveFile[0]))
                print(file_name)
                self.file_name = os.path.basename(Path(saveFile[0]))
                btn_text = "Запись в файл " + self.file_name + ": "
                if self.file_name == self.file_name2:
                    msg = QMessageBox()
                    msg.setWindowTitle("Ошибка")
                    msg.setText("Вы выбрали имя файла, которое уже находится в обработке")
                    msg.setIcon(QMessageBox.Icon.Critical)
                    a = msg.exec()
                else:
                    if saveFile[0] != "":
                        self.text_bar1.setHidden(False)
                        self.text_bar1.setText(btn_text)
                        self.text_bar1.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                        # self.text_bar1.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)
                        # self.text_bar1.setFont(QFont('Segoe UI', 10))
                        #self.text_bar1.setFixedSize(18)
                        self.button.setEnabled(False)
                        self.button2.setEnabled(False)
                        self.progressbar.setHidden(False)
                        # self.progressbar.setAlignment(Qt.AlignmentFlag.AlignVCenter)
                        input_data = self.input1.text()

                        self.thread = QThread()

                        self.worker = Excel_redactor.Worker_Excel(input_data, saveFile[0])

                        self.worker.moveToThread(self.thread)

                        self.thread.started.connect(self.worker.run)
                        #self.cancel_btn_1.clicked.connect(self.worker.terminate)

                        self.worker._signal.connect(self.signal_accept)

                        self.worker.error.connect(self.error_signal)

                        self.worker.finished.connect(self.thread_complete)

                        self.worker.finished.connect(self.thread.quit)
                        self.worker.finished.connect(self.worker.deleteLater)
                        self.thread.finished.connect(self.thread.deleteLater)

                        print("STARTING THREAD")
                        self.thread.start()
                    else:
                        print("no file selected")
                        msg = QMessageBox()
                        msg.setWindowTitle("Выберите файл!")
                        msg.setText("Вы не выбрали название и расположение word-файла")
                        msg.setIcon(QMessageBox.Icon.Information)
                        a = msg.exec()
            else:
                msg = QMessageBox()
                msg.setWindowTitle("Ошибка")
                msg.setText("Вы выбрали несуществующий файл")
                msg.setIcon(QMessageBox.Icon.Critical)
                a = msg.exec()
        else:
            msg = QMessageBox()
            msg.setWindowTitle("Ошибка")
            msg.setText("Вы выбрали файл неверного расширения. Допустимые файлы *.xls и *.xlsx")
            msg.setIcon(QMessageBox.Icon.Critical)
            a = msg.exec()
            self.input1.clear()


    def btn3_was_clicked(self):
        self.times_clk_btn += 1
        if self.times_clk_btn == 1:
            msg = QMessageBox()
            msg.setWindowTitle("Внимание!")
            msg.setText("Для корректной работы программы, необходимо,"
                        " чтобы выбранный файл для word отчета был закрыт, если вы хотите его перезаписать.")
            msg.setIcon(QMessageBox.Icon.Information)
            a = msg.exec()
        if get_file_extension(self.input2.text()) == 0:
            if Path(self.input2.text()).exists():
                saveFile = QFileDialog.getSaveFileName(self, 'Save File', "Новый отчет.docx", "Word файл (*.docx)")
                file_name = os.path.basename(Path(saveFile[0]))
                print(file_name)
                self.file_name2 =os.path.basename(Path(saveFile[0]))
                btn_text = "Запись в файл " + self.file_name2  + ": "
                if self.file_name2 == self.file_name:
                    msg = QMessageBox()
                    msg.setWindowTitle("Ошибка")
                    msg.setText("Вы выбрали имя файла, которое уже находится в обработке")
                    msg.setIcon(QMessageBox.Icon.Critical)
                    a = msg.exec()
                else:
                    if saveFile[0] != "":
                        self.text_bar2.setHidden(False)
                        self.text_bar2.setText(btn_text)
                        self.text_bar2.setAlignment(Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter)

                        self.button1.setEnabled(False)
                        self.button3.setEnabled(False)
                        self.progressbar2.setHidden(False)
                        input_data = self.input2.text()

                        self.thread2 = QThread()

                        self.worker2 = Excel_redactor.Worker_Excel_RPR(input_data, saveFile[0]) # change func -> new class .

                        self.worker2.moveToThread(self.thread2)

                        self.thread2.started.connect(self.worker2.run)

                        self.worker2._signal.connect(self.signal_accept_zone)

                        self.worker2.error.connect(self.error_signal_zone)

                        self.worker2.finished.connect(self.thread_complete_zone)

                        self.worker2.finished.connect(self.thread2.quit)
                        self.worker2.finished.connect(self.worker2.deleteLater)
                        self.thread2.finished.connect(self.thread2.deleteLater)

                        print("STARTING THREAD")
                        self.thread2.start()
                    else:
                        print("no file selected")
                        msg = QMessageBox()
                        msg.setWindowTitle("Выберите файл!")
                        msg.setText("Вы не выбрали название и расположение word-файла")
                        msg.setIcon(QMessageBox.Icon.Information)
                        a = msg.exec()
            else:
                msg = QMessageBox()
                msg.setWindowTitle("Ошибка")
                msg.setText("Вы выбрали несуществующий файл")
                msg.setIcon(QMessageBox.Icon.Critical)
                a = msg.exec()
        else:
            msg = QMessageBox()
            msg.setWindowTitle("Ошибка")
            msg.setText("Вы выбрали файл неверного расширения. Допустимые файлы *.xls и *.xlsx")
            msg.setIcon(QMessageBox.Icon.Critical)
            a = msg.exec()
            self.input2.clear()

        # msg = QMessageBox()
        # msg.setWindowTitle("Ой!")
        # msg.setText("С выгрузкой файлов типа 'Зоны' все еще есть проблемы. Данная кнопка временно отключена.")
        # msg.setIcon(QMessageBox.Icon.Information)
        # a = msg.exec()
        # self.input2.clear()


    def signal_accept(self, msg):
        #print(self.progressbar.value())
        cur_value = self.progressbar.value()
        cur_value += int(msg)
        if cur_value > 100:
            self.progressbar.setValue(100)
        else:
            self.progressbar.setValue(cur_value)


    def signal_accept_zone(self, msg):

        cur_value = self.progressbar2.value()
        print("CUR = " + str(cur_value))
        print(msg)
        cur_value += int(msg)
        print("CUR + MSG = " + str(cur_value))
        if cur_value > 100:
            self.progressbar2.setValue(100)
        else:
            self.progressbar2.setValue(cur_value)


    def thread_complete(self, x):
        if x == 0:
            msg = QMessageBox()
            msg.setWindowTitle("Готово!")
            text = "Word файл " + self.file_name +  " успешно создан"
            msg.setText(text)
            msg.setIcon(QMessageBox.Icon.Information)
            a = msg.exec()
            self.progressbar.setValue(0)
            # self.button2.setEnabled(True)
        else:
            self.progressbar.setValue(0)

        self.text_bar1.clear()
        self.text_bar1.setHidden(True)
        self.input1.clear()
        self.progressbar.setHidden(True)

        self.button.setEnabled(True)
        self.button2.setEnabled(True)
        self.file_name = ""


    def thread_complete_zone(self, x):
        if x == 0:
            msg = QMessageBox()
            msg.setWindowTitle("Готово!")
            text = "Word файл " + self.file_name2 + " успешно создан"
            msg.setText(text)
            msg.setIcon(QMessageBox.Icon.Information)
            a = msg.exec()
            self.progressbar2.setValue(0)
            # self.button2.setEnabled(True)
        else:
            self.progressbar2.setValue(0)

        self.input2.clear()
        self.progressbar2.setHidden(True)
        self.text_bar2.clear()
        self.text_bar2.setHidden(True)

        self.button1.setEnabled(True)
        self.button3.setEnabled(True)
        self.file_name2 = ""

    def error_signal(self, y):
        if y == 1:
            msg = QMessageBox()
            msg.setWindowTitle("Ошибка")
            msg.setText("Структура выбраного файла не соответсвует структуре 'Scens'.")
            msg.setIcon(QMessageBox.Icon.Critical)
            a = msg.exec()
            self.button.setEnabled(True)
            self.button2.setEnabled(True)
            self.file_name = ""
    def error_signal_zone(self, y):
        if y == 2:
            msg = QMessageBox()
            msg.setWindowTitle("Ошибка")
            msg.setText("Структура выбраного файла не соответсвует структуре 'Зоны'.")
            msg.setIcon(QMessageBox.Icon.Critical)
            a = msg.exec()
            self.button1.setEnabled(True)
            self.button3.setEnabled(True)
            self.file_name2 = ""







app = QApplication(sys.argv)
print("i am here111")
window = MainWindow()
window.show()
#window.init()
# worker = Excel_redactor.Excel()
# worker._signal.connect(window.signal_accept)
# worker.start()



# Excel = Excel_redactor.Excel()
# Excel.edit_excel('C:/Users/fmoro/Desktop/Scens_1_метео 14.06.2022.xlsx')
sys.exit(app.exec())
