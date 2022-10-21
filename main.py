import time

from PyQt6.QtWidgets import QApplication, QMainWindow, QPushButton, QFileDialog, QGridLayout, QMessageBox
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

window_titles = [
    'My App',
    'My App',
    'Still My App',
    'Still My App',
    'What on earth',
    'What on earth',
    'This is surprising',
    'This is surprising',
    'Something went wrong'
]

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

        self.resize(800, 150)
        self.n_times_clicked = 0

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

        # self.progress = QProgressBar()
        # self.progress.setRange(0, 100)




        self.input1 = QLineEdit("")
        self.input1.setPlaceholderText("Введите путь до файла: C:\\user\\file.docx")
        self.input2 = QLineEdit("")
        self.input2.setPlaceholderText("Введите путь до файла: C:\\user\\file.docx")

        self.label1 = QLabel("")
        self.label_jif = QLabel()
        self.movie = QMovie("Spinner-1s-25px.gif")
        self.label_jif.setMovie(self.movie)
        self.movie.start()
        self.label_jif.setHidden(True)

        self.label_jif2 = QLabel()
        self.movie2 = QMovie("Spinner-1s-25px.gif")
        self.label_jif2.setMovie(self.movie2)
        self.movie2.start()
        self.label_jif2.setHidden(True)
        #self.label_jif.setFixedSize(100, 100)




        layout = QGridLayout()

        layout.addWidget(self.button, 0, 0)
        layout.addWidget(self.input1, 0, 1)
        layout.addWidget(self.button2, 0, 2)
        # layout.addWidget(self.label1, 0, 2)
        # layout.addWidget(self.label1, 0, 3)

        layout.addWidget(self.button1, 1, 0)
        layout.addWidget(self.input2, 1, 1)
        layout.addWidget(self.button3, 1, 2)

        #layout.addWidget(self.label1, 1, 0)

        #layout.addWidget(self.label1, 0, 0)
        layout.addWidget(self.label_jif, 0, 3)
        layout.addWidget(self.label_jif2, 1, 3)

        layout.setRowStretch(2, 1)


        # layout.addWidget(self.progress, 2, 0, 2, 3)

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

        if get_file_extension(self.input1.text()) == 0:
            if Path(self.input1.text()).exists():
                self.label_jif.setHidden(False)
                saveFile = QFileDialog.getSaveFileName(self, 'Save File', "Новый отчет.docx")
                #self.label_jif.setHidden(False)
                if saveFile[0] != "":
                    # self.progress


                    #print("save = " + str(saveFile[0]))
                    input_data = self.input1.text()
                    #self.label_jif.setHidden(False)
                    time.sleep(10)
                    #self.movie.start()

                    Excel = Excel_redactor.Excel()
                    check = 0
                    check = Excel.edit_excel(input_data, saveFile[0])
                    self.input1.clear()
                    self.label_jif.setHidden(True)
                    if check == 1:
                        msg = QMessageBox()
                        msg.setWindowTitle("Ошибка")
                        msg.setText("Структура выбраного файла не соответсвует структуре 'Scens'.")
                        msg.setIcon(QMessageBox.Icon.Critical)
                        a = msg.exec()
                    #self.movie.stop()
                    else:
                        msg = QMessageBox()
                        msg.setWindowTitle("Готово!")
                        msg.setText("Word файл успешно создан")
                        msg.setIcon(QMessageBox.Icon.Information)
                        a = msg.exec()


                else:
                    print("no file selected")
                    self.label_jif.setHidden(True)
                #do smf with file
                # print("INPUT")
                # print(self.input1)
                # print(self.input1.text())
                # input_data = self.input1.text()
                #
                # Excel = Excel_redactor.Excel()
                # Excel.edit_excel(input_data)
                # self.input1.clear()
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
        if get_file_extension(self.input2.text()) == 0:
            if Path(self.input2.text()).exists():
                self.label_jif2.setHidden(False)
                saveFile = QFileDialog.getSaveFileName(self, 'Save File', "Новый отчет.docx")
                if saveFile[0] != "":
                    start_time = time.time()
                    while int(time.time() - start_time) <= 600:
                        print(time.time() - start_time)
                    msg = QMessageBox()
                    msg.setWindowTitle("Ошибка")
                    msg.setText("Что-то пошло не так")
                    msg.setIcon(QMessageBox.Icon.Critical)
                    a = msg.exec()

                    self.label_jif2.setHidden(True)
                    file = open(saveFile[0], 'w')
                    # file.write("1111")
                    file.close()
                    self.input2.clear()

                else:
                    self.label_jif2.setHidden(True)
                    print("no file selected")
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

        # print()
        # saveFile = QFileDialog.getSaveFileName(self, 'Save File', "Новый отчет.docx")
        # print("save = " + str(saveFile[0]))
        # file = open(saveFile[0], 'w')
        # # file.write("1111")
        # file.close()
        # self.input2.clear()




app = QApplication(sys.argv)
print("i am here111")
window = MainWindow()
#window.init()
window.show()
# Excel = Excel_redactor.Excel()
# Excel.edit_excel('C:/Users/fmoro/Desktop/Scens_1_метео 14.06.2022.xlsx')
app.exec()
