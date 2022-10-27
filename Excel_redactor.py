import collections
import datetime
import math
import os
import shutil
import time
from copy import copy
import binascii


# import locale
# os.environ["PYTHONIOENCODING"] = "utf-8";
# myLocale=locale.setlocale(category=locale.LC_ALL, locale="en_GB.UTF-8")

import docx
import openpyxl
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
from docx.shared import Pt
from docx.table import _Cell
from openpyxl.styles import Border, Side, Alignment, Font
from openpyxl import Workbook
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from decimal import Decimal
from PyQt6.QtCore import QThread, pyqtSignal, QObject, QRunnable, pyqtSlot

import traceback, sys

def change_gas(cur_list):
    cur_str = str(cur_list)
    pos = cur_str.find(" (")
    if pos != -1:
        new_str = cur_str[:pos]
    else:
        new_str = cur_list
    return new_str

def change_state(r_otv, n_gas):

    r_otv = r_otv.strip()

    if r_otv is not None and r_otv != "None": # if otverstie est' -
        if r_otv != "Полное разрушение": # if ne polnoe razryshenie
            r_otv = float(r_otv)
            d_otv = math.sqrt((r_otv * 4) / math.pi) * 1000
            round_num = round(d_otv * 2) / 2
            if str(round_num)[-1] == "0":
                res = str(round_num)
                pos = res.find(".0")
                res = res[:pos]
                res += " мм"
                # print(res)
            else:
                res = str(round_num)
                res = res.replace(".", ",")
                res += " мм"
        if r_otv == "Полное разрушение":
            res = "ПР"
    else:
        res = ''

    if n_gas is not None and n_gas != "None":
        new_state = n_gas + ' ' + res
    else:
        new_state = res

    return (new_state)


def get_super(x):
    normal = "0123456789+-=()"
    super_s = "⁰¹²³⁴⁵⁶⁷⁸⁹⁺⁻⁼⁽⁾"
    res = x.maketrans("".join(normal), "".join(super_s))
    return x.translate(res)


def another_e(chislo):
    power = int(str(chislo)[-2] + str(chislo)[-1])
    value_new = float(str(chislo * pow(10, power))[:-2])
    answer = str(round(value_new, 2))
    p = get_super(str(10))
    answer = answer + ' × 10' + p
    return answer

def filtery(sort_dict, obor, ov, pdo, max_1,max_2):
    #Структура: [0]Оборудование, [1]Состояние, [2]Опасное вещество,
    # [3]Площадь отверстия, [4]Масса, [5]Дрейф, [6]1,[7]10,[8]25,[9]50,[10]90,[11]99
    rez = []
    #1.Выбираем ооборудование
    #2.Выбираем опасное вещество
    for o in range(len(obor)):
        for v in range(len(ov)):
            for p in range(len(pdo)):
                a = []
                max = 0
                max_dr = 0
                b = []
                max_str = []
                for i in range(len(sort_dict)):
                    if sort_dict[i][0] == obor[o] and sort_dict[i][2] == ov[v] and sort_dict[i][3] == pdo[p]:
                        a.append(sort_dict[i])
                    if i == len(sort_dict)-1:
                        for j in range(len(a)):
                            if a[j][max_1] is None: a[j][max_1] = 0
                            if a[j][max_1] > max:
                                b.clear()
                                b.append(a[j])
                                max = a[j][max_1]
                            elif a[j][max_1] == max:
                                b.append(a[j])
                        for j in range(len(b)):
                            if b[j][max_2] is None: b[j][max_2] = 0
                            if b[j][max_2] >= max_dr:
                                max_str = (b[j])
                                max_dr = b[j][max_2]
                        if max_str is not []:
                            rez.append(max_str)
    rez = [x for x in rez if x != []]
    for i in range (len(rez)):
        rez[i].pop(3)
        rez[i].pop(2)
    return rez

def filtery_for_3(sort_dict, obor, ov, pdo, max_1,max_2,max_3):
    #Структура: [0]Оборудование, [1]Состояние, [2]Опасное вещество,
    # [3]Площадь отверстия, [4]Масса, [5]Дрейф, [6]1,[7]10,[8]25,[9]50,[10]90,[11]99
    rez = []
    for o in range(len(obor)):
        for v in range(len(ov)):
            for p in range(len(pdo)):
                a = []
                max = 0
                max_dr = 0
                max_dr_dr = 0
                b = []
                c = []
                max_str = []
                for i in range(len(sort_dict)):
                    if sort_dict[i][0] == obor[o] and sort_dict[i][2] == ov[v] and sort_dict[i][3] == pdo[p]:
                        a.append(sort_dict[i])
                    if i == len(sort_dict)-1:
                        for j in range(len(a)):
                            if a[j][max_1] is None: a[j][max_1] = 0
                            if a[j][max_1] > max:
                                b.clear()
                                b.append(a[j])
                                max = a[j][max_1]
                            elif a[j][max_1] == max:
                                b.append(a[j])
                        for j in range(len(b)):
                            if b[j][max_2] is None: b[j][max_2] = 0
                            if b[j][max_2] > max_dr:
                                c.clear()
                                c.append(b[j])
                                max = b[j][max_2]
                            elif b[j][max_2] == max:
                                c.append(b[j])
                        for j in range(len(c)):
                            if c[j][max_3] is None: c[j][max_3] = 0
                            if c[j][max_3] > max_dr_dr:
                                max_str = (c[j])
                                max_dr = c[j][max_2]
                        if max_str is not []:
                            rez.append(max_str)
    rez = [x for x in rez if x != []]
    for i in range (len(rez)):
        rez[i].pop(3)
        rez[i].pop(2)
    return rez

def filtery_for_1(sort_dict, obor, ov, pdo, max_1):
    #Структура: [0]Оборудование, [1]Состояние, [2]Опасное вещество,
    # [3]Площадь отверстия, [4]Масса, [5]Дрейф, [6]1,[7]10,[8]25,[9]50,[10]90,[11]99
    rez = []
    for o in range(len(obor)):
        for v in range(len(ov)):
            for p in range(len(pdo)):
                a = []
                max = 0
                max_str = []
                for i in range(len(sort_dict)):
                    if sort_dict[i][0] == obor[o] and sort_dict[i][2] == ov[v] and sort_dict[i][3] == pdo[p]:
                        a.append(sort_dict[i])
                    if i == len(sort_dict)-1:
                        for j in range(len(a)):
                            if a[j][max_1] is None: a[j][max_1] = 0
                            if a[j][max_1] >= max:
                                max_str = (a[j])
                                max = a[j][max_1]
                        if max_str is not []:
                            rez.append(max_str)
    rez = [x for x in rez if x != []]
    for i in range (len(rez)):
        rez[i].pop(3)
        rez[i].pop(2)
    return rez

def chistka_3_ogn(dict):
    for i in range (len(dict)):
        dict[i].pop(3)
        dict[i].pop(2)
    return dict




def sort_dict(dict):
    dictionary_keys = list(dict.keys())
    sorted_dict = {dictionary_keys[x]: sorted(
        dict.values())[x] for x in range(len(dictionary_keys))}
    return sorted_dict

class Worker_Excel(QObject):

    _signal = pyqtSignal(int)
    finished = pyqtSignal(int)
    error = pyqtSignal(int)
    # terminate = pyqtSignal()

    def __init__(self, excel_path, word_path):
        super(Worker_Excel, self).__init__()
        self.excel_path = excel_path
        self.word_path = word_path
    #
    # def __del__(self):
    #     self.wait()
    def run(self):
        print("IN THE RUN")
        start_time = time.time()

        wb = openpyxl.load_workbook(self.excel_path)
        sheet_name = (wb.get_sheet_names())
        sheet = wb.get_sheet_by_name(sheet_name[0])
        print(sheet_name[0])
        if sheet_name[0] != 'Scens':
            wb.close()
            self.error.emit(1)
            self.finished.emit(-1)
            return 1

            #return 1


        # B - Наименование оборудования
        # D - опасное вещество
        # Е - исход
        # P - диаметр !!!!!!!!!!!!!!!!!!!!
        # F - частота сценария
        # H - масса Гф
        # I - масса Жф

        dict = []
        stroka = []
        all_rows = sheet.max_row
        j = 0
        for i in range(3, all_rows):
            stroka.append(sheet['B' + str(i)].value)
            stroka.append(sheet['D' + str(i)].value)
            stroka.append(sheet['E' + str(i)].value)

            q = (sheet['G' + str(i)].value)
            if q == '-':
                stroka.append(0)
            else:
                if q is None:
                    stroka.append(0)
                else:
                    q = (math.sqrt(float(q) * 4 / float(math.pi)) * 1000)
                    if int(q) == 12 or int(q) == 13:
                        stroka.append(12.5)
                    else:
                        stroka.append(int(q))

            with_10_ = another_e(sheet['F' + str(i)].value)
            stroka.append(with_10_)

            stroka.append(sheet['H' + str(i)].value)
            stroka.append(sheet['I' + str(i)].value)

            dict.append(stroka)
            stroka = []

        wb.close()
        print("start DOCX")
        (dict.sort(key = lambda row: (row[1],row[2],row[3],row[0])))

        doc = Document('Scene.docx')
        doc_table = doc.tables

        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(10)
        paragraph_format = doc.styles['Normal'].paragraph_format
        paragraph_format.line_spacing = Pt(12)  ###################
        paragraph_format.space_after = (0)  ################

        step = all_rows / 100
        percent = 1

        for k in range(2, all_rows-3):
            if 1.0 > k % step >= 0.0:
                percent += 1

                self._signal.emit(percent)


            list = dict[k]
            row = doc_table[0].add_row().cells  # олучаем все ячейки ряда
            for col in range(7):
                cell = row[col]
                paragraph = cell.paragraphs[0]
                paragraph.style = doc.styles['Normal']
                if col == 3:
                    if (list[col]) == 0:
                        list[col] = 'Полное разрущение'
                    else:
                        list[col] = str(list[col]).replace('.', ',')
                if col == 4: list[col] = str(list[col]).replace('.', ',')
                if col == 0:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                else:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = paragraph.add_run(str(list[col]))
        print("SAVE------------------------")
        doc.save(self.word_path)
        print("--- %s seconds ---" % (time.time() - start_time))
        self.finished.emit(0)
        #return 0




        #
        # j = 0
        # step = all_rows / 100
        # percent = 1
        # #self._signal.emit(percent)
        # for row in range(2, all_rows-2):
        #     # print(row)
        #     # print("%%%%%%%%%%%% -- " + str(percent))
        #     # print(complete_percent_value)
        #     if row % step < 1.0 and row % step >= 0.0:
        #         #percent+=percent
        #         percent += 1
        #         self._signal.emit(percent)
        #     list = dict[j]
        #     j = j + 1
        #     row = doc_table[0].add_row().cells  # олучаем все ячейки ряда
        #     for col in range(7):
        #         cell = row[col]
        #         paragraph = cell.paragraphs[0]
        #         paragraph.style = doc.styles['Normal']
        #         if col == 3:
        #             if (list[col]) == 0:
        #                 list[col] = 'Полное разрущение'
        #             else:
        #                 list[col] = str(list[col]).replace('.', ',')
        #         if col == 4: list[col] = str(list[col]).replace('.', ',')
        #         if col == 0:
        #             paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        #         else:
        #             paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        #         run = paragraph.add_run(str(list[col]))
        #
        #         cell = doc_table[0].cell(row, col)
        #         cell.text = str(list[col])
        #
        #
        # pos = excel_path.find('.xls')
        # print("I AM HERE")
        # print(pos)
        # copy_path = excel_path
        # copy_path = excel_path[:pos] + " -copy" + excel_path[pos:]
        # print(copy_path)
        # print(excel_path)
        #
        # shutil.copy(excel_path, copy_path)
        # book = openpyxl.load_workbook(copy_path)
        # cur_sheet = book.active
        # value = 0
        #
        #
        # for row in cur_sheet[3:cur_sheet.max_row]:
        #     # print('in loop')
        #     cur_cell = row[6]
        #     if cur_cell.value == '-':
        #         # print('Complete ')
        #         continue
        #     else:
        #         if math.sqrt(float(cur_cell.value) * 4 / float(math.pi))*1000 < 0 or None:
        #             print('BO4ik potik')
        #         else:
        #             value = math.sqrt(float(cur_cell.value) * 4 / float(math.pi))*1000
        #
        #         # #print(value)
        #         # print('in loop 2')
        #
        # value = str(value)
        #
        # value = value.replace('.', ',')
        # print("VALUE , = " + value)
        # index = value.find(',')
        # value = value[:index+3]
        # print("VALUE = " + value)
        #     #cell.value = '=КОРЕНЬ(F' + str(myiter) + '*4/ПИ())*1000'
        #
        # book.save(copy_path)
        # book = openpyxl.load_workbook(copy_path, data_only=True)
        #
        # cur_sheet = book.active
        # cur_cell = cur_sheet['F3']
        #
        #
        #
        # print(power)
        # print(value_new)
        # #print(value * pow(10, power))
        #
        # #print(str(value)[-1])
        # os.remove(copy_path)
        #
        # book = openpyxl.load_workbook(excel_path)

class Worker_Excel_RPR(QObject):

    _signal = pyqtSignal(int)
    finished = pyqtSignal(int)
    error = pyqtSignal(int)

        # terminate = pyqtSignal()

    def __init__(self, excel_path, word_path):
        super(Worker_Excel_RPR, self).__init__()
        self.excel_path = excel_path
        self.word_path = word_path

        #
        # def __del__(self):
        #     self.wait()
    def run(self):
        start_time = time.time()
        wb = openpyxl.load_workbook(self.excel_path)
        sheet_name = (wb.get_sheet_names())
        print(len(sheet_name))
        if len(sheet_name) != 8 and  len(sheet_name) != 9:
            wb.close()
            print("WRONG STRUCT")
            self.error.emit(2)
            self.finished.emit(-1)
            return 1

        if sheet_name[4] == '3_Взрыв ТВС Избыточное давл':
            var = 'OPO'
            doc = Document('Zona_OB_OPO.docx')
        if sheet_name[4] == '3_Огненный шар Вероятностно':
            var = 'RPR'
            doc = Document('Zona_RPR.docx')
        if sheet_name[4] != '3_Взрыв ТВС Избыточное давл' and sheet_name[4] != '3_Огненный шар Вероятностно':
            wb.close()
            print("WRONG STRUCT")
            self.error.emit(2)
            self.finished.emit(-1)
            return 1


        doc_table = doc.tables
        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(10)
        paragraph_format = doc.styles['Normal'].paragraph_format
        paragraph_format.line_spacing = Pt(12)
        paragraph_format.space_after = (0)

        percent = 0
        for i in range(len(sheet_name)):
            if var == 'RPR':
                print("IN RPR")
                if i == 5 or i == 6 or i == 1 or i == 0 or i == 4 or i == 3:  # 2ой лист
                    self._signal.emit(percent)
                    percent += 16
                    dict = []
                    l = 0
                    print(sheet_name[i])
                    print(i)
                    sheet = wb.get_sheet_by_name(sheet_name[i])
                    all_rows = sheet.max_row
                    oborudovanie = []
                    ov = []
                    pdo = []
                    if i == 0:
                        for j in range(4, all_rows):
                            stroka = []
                            obor = str(sheet['C' + str(j)].value).strip()  # Оборудование
                            stroka.append(obor)
                            if not (obor in oborudovanie): oborudovanie.append(obor)

                            ve = str(sheet['E' + str(j)].value).strip()  # Вещество
                            ve = change_gas(ve)  # Вещество
                            s_def_otv = str(sheet['H' + str(j)].value)  # Площадь деф отв

                            state = change_state(s_def_otv, ve)  # Состояние
                            stroka.append(state)

                            stroka.append(ve)
                            if not (ve in ov): ov.append(ve)

                            stroka.append(s_def_otv)  # Площадь деф отв
                            if not (s_def_otv in pdo): pdo.append(s_def_otv)

                            stroka.append(sheet['I' + str(j)].value)  # Масса
                            stroka.append(sheet['K' + str(j)].value)  # Радиус НКПВ
                            dict.append(stroka)
                    if i == 3:
                        for j in range(4, all_rows):
                            stroka = []
                            obor = str(sheet['C' + str(j)].value).strip()  # Оборудование
                            stroka.append(obor)
                            ve = str(sheet['E' + str(j)].value).strip()  # Вещество
                            ve = change_gas(ve)  # Вещество
                            s_def_otv = str(sheet['F' + str(j)].value)  # Площадь деф отв

                            state = change_state(s_def_otv, ve)  # Состояние
                            stroka.append(state)

                            stroka.append(ve)

                            stroka.append(s_def_otv)  # Площадь деф отв

                            stroka.append(sheet['G' + str(j)].value)  # Масса
                            stroka.append(sheet['I' + str(j)].value)  # Радуис
                            dict.append(stroka)
                    if i == 5 or i == 6 or i == 1 or i == 4:
                        for j in range(4, all_rows):
                            stroka = []

                            obor = str(sheet['C' + str(j)].value).strip()  # Оборудование
                            stroka.append(obor)
                            if not (obor in oborudovanie): oborudovanie.append(obor)

                            ve = str(sheet['E' + str(j)].value).strip()  # Вещество
                            ve = change_gas(ve)
                            if i == 5 or i == 6:
                                s_def_otv = str(sheet['H' + str(j)].value)  # Площадь деф отв)
                            elif i == 1 or i == 4:
                                s_def_otv = str(sheet['F' + str(j)].value)  # Площадь деф отв)

                            state = change_state(s_def_otv, ve)  # Состояние
                            stroka.append(state)

                            stroka.append(ve)
                            if not (ve in ov): ov.append(ve)

                            if i == 5 or i == 6:
                                # s = sheet['H' + str(j)].value # Площадь деф отв
                                stroka.append(s_def_otv)
                                if not (s_def_otv in pdo): pdo.append(s_def_otv)
                                stroka.append(sheet['I' + str(j)].value)  # Масса
                                if i == 5:
                                    stroka.append(sheet['J' + str(j)].value)  # Дрейф
                                stroka.append(sheet['K' + str(j)].value)  # Радиус 1
                                stroka.append(sheet['L' + str(j)].value)  # Радиус 10
                                stroka.append(sheet['M' + str(j)].value)  # Радиус 25
                                stroka.append(sheet['N' + str(j)].value)  # Радиус 50
                                stroka.append(sheet['O' + str(j)].value)  # Радиус 90
                                stroka.append(sheet['P' + str(j)].value)  # Радиус 99
                                if i == 6:
                                    stroka.append(sheet['Q' + str(j)].value)  # Радиус 100
                            elif i == 1 or i == 4:
                                # s = sheet['F' + str(j)].value # Площадь деф отв
                                stroka.append(s_def_otv)
                                if not (s_def_otv in pdo): pdo.append(s_def_otv)
                                stroka.append(sheet['G' + str(j)].value)  # Масса
                                stroka.append(sheet['I' + str(j)].value)  # Радиус 1
                                stroka.append(sheet['J' + str(j)].value)  # Радиус 10
                                stroka.append(sheet['K' + str(j)].value)  # Радиус 25
                                stroka.append(sheet['L' + str(j)].value)  # Радиус 50
                                stroka.append(sheet['M' + str(j)].value)  # Радиус 90
                                stroka.append(sheet['N' + str(j)].value)  # Радиус 99
                                stroka.append(sheet['O' + str(j)].value)  # Радиус 100

                            dict.append(stroka)
                            # l = l + 1

                    #print(dict)

                    if i == 5 or i == 6:
                        dict.sort(key=lambda row: (row[2] or 0, row[3] or 0, row[0] or 0))
                        #print(dict)
                        rez = filtery(dict, oborudovanie, ov, pdo, 6, 5)
                    elif i == 0:
                        dict.sort(key=lambda row: (row[2] or 0, row[3] or 0, row[0] or 0))
                        #print(dict)
                        rez = filtery_for_1(dict, oborudovanie, ov, pdo, 5)
                    elif i == 4:
                        dict.sort(key=lambda row: (row[0] or 0))
                        #print(dict)
                        rez = chistka_3_ogn(dict)
                    elif i == 3:
                        dict.sort(key=lambda row: (row[2] or 0, row[3] or 0, row[0] or 0))
                        #print(dict)
                        rez = chistka_3_ogn(dict)
                    elif i == 1:
                        dict.sort(key=lambda row: (row[2] or 0, row[3] or 0, row[0] or 0))
                        #print(dict)
                        rez = filtery_for_1(dict, oborudovanie, ov, pdo, 5)
                    #print(rez)

                    if i == 0 or i == 3:
                        count_stolb = 4
                    else:
                        count_stolb = 10

                    for k in range(len(rez)):
                        list = rez[k]
                        if i == 0: row = doc_table[2].add_row().cells  # олучаем все ячейки ряда
                        if i == 1: row = doc_table[5].add_row().cells  # олучаем все ячейки ряда
                        if i == 3: row = doc_table[4].add_row().cells  # олучаем все ячейки ряда
                        if i == 4: row = doc_table[3].add_row().cells  # олучаем все ячейки ряда
                        if i == 5: row = doc_table[0].add_row().cells  # олучаем все ячейки ряда
                        if i == 6: row = doc_table[1].add_row().cells  # олучаем все ячейки ряда
                        for col in range(count_stolb):
                            cell = row[col]
                            paragraph = cell.paragraphs[0]
                            paragraph.style = doc.styles['Normal']
                            if col == 0:
                                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                            else:
                                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            if str(list[col]) is None or str(list[col]) == 'None': list[col] = ''
                            run = paragraph.add_run(str(list[col]))
                    doc.save(self.word_path)

                    print("--- %s seconds ---" % (time.time() - start_time))
            if var == 'OPO':
                if i == 4 or i == 6 or i == 7 or i == 3 or i == 1 or i == 0:
                    self._signal.emit(percent)
                    percent += 16
                    dict = []
                    l = 0
                    print(sheet_name[i])
                    print(i)
                    sheet = wb.get_sheet_by_name(sheet_name[i])
                    all_rows = sheet.max_row
                    oborudovanie = []
                    ov = []
                    pdo = []
                    if i == 4 or i == 6:
                        for j in range(4, all_rows):
                            stroka = []

                            obor = str(sheet['C' + str(j)].value).strip()  # Оборудование
                            stroka.append(obor)
                            if not (obor in oborudovanie): oborudovanie.append(obor)

                            ve = str(sheet['E' + str(j)].value).strip()  # Вещество
                            ve = change_gas(ve)
                            if i == 4 or i == 6:
                                s_def_otv = str(sheet['H' + str(j)].value)  # Площадь деф отв

                            state = change_state(s_def_otv, ve)  # Состояние
                            stroka.append(state)

                            stroka.append(ve)
                            if not (ve in ov): ov.append(ve)

                            if i == 4 or i == 6:
                                # s = sheet['H' + str(j)].value # Площадь деф отв
                                stroka.append(s_def_otv)
                                if not (s_def_otv in pdo): pdo.append(s_def_otv)

                                stroka.append(sheet['I' + str(j)].value)  # Масса
                                stroka.append(sheet['J' + str(j)].value)  # Дрейф
                                stroka.append(sheet['K' + str(j)].value)  # Радиус 1
                                stroka.append(sheet['L' + str(j)].value)  # Радиус 2
                                stroka.append(sheet['M' + str(j)].value)  # Радиус 3
                                stroka.append(sheet['N' + str(j)].value)  # Радиус 5
                                stroka.append(sheet['O' + str(j)].value)  # Радиус 6
                                stroka.append(sheet['P' + str(j)].value)  # Радиус 7
                                stroka.append(sheet['Q' + str(j)].value)  # Радиус 12
                            dict.append(stroka)
                    if i == 7:
                        for j in range(4, all_rows):
                            stroka = []

                            obor = str(sheet['C' + str(j)].value).strip()  # Оборудование
                            stroka.append(obor)
                            if not (obor in oborudovanie): oborudovanie.append(obor)

                            ve = str(sheet['E' + str(j)].value).strip()  # Вещество
                            ve = change_gas(ve)

                            s_def_otv = str(sheet['H' + str(j)].value)  # Площадь деф отв

                            state = change_state(s_def_otv, ve)  # Состояние
                            stroka.append(state)

                            stroka.append(ve)
                            if not (ve in ov): ov.append(ve)

                            if i == 7:
                                # s = sheet['H' + str(j)].value # Площадь деф отв
                                stroka.append(s_def_otv)
                                if not (s_def_otv in pdo): pdo.append(s_def_otv)

                                stroka.append(round(sheet['I' + str(j)].value, 2))  # Масса
                                stroka.append(round(sheet['J' + str(j)].value, 2))  # Дрейф
                                stroka.append(sheet['L' + str(j)].value)  # По ветру
                                stroka.append(sheet['M' + str(j)].value)  # Против ветра
                                stroka.append(sheet['N' + str(j)].value)  # Полуширина
                            dict.append(stroka)
                    if i == 3:
                        for j in range(4, all_rows):
                            stroka = []

                            obor = str(sheet['C' + str(j)].value).strip()  # Оборудование
                            stroka.append(obor)
                            if not (obor in oborudovanie): oborudovanie.append(obor)

                            ve = str(sheet['E' + str(j)].value).strip()  # Вещество
                            ve = change_gas(ve)

                            s_def_otv = str(sheet['F' + str(j)].value)  # Площадь деф отв

                            state = change_state(s_def_otv, ve)  # Состояние
                            stroka.append(state)

                            stroka.append(ve)
                            if not (ve in ov): ov.append(ve)
                            if i == 3:
                                stroka.append(s_def_otv)
                                if not (s_def_otv in pdo): pdo.append(s_def_otv)
                                stroka.append(round(sheet['G' + str(j)].value, 2))  # Масса
                                stroka.append(sheet['I' + str(j)].value)  # Радиус 1
                                stroka.append(sheet['J' + str(j)].value)  # Радиус 10
                                stroka.append(sheet['K' + str(j)].value)  # Радиус 25
                                stroka.append(sheet['L' + str(j)].value)  # Радиус 50
                                stroka.append(sheet['M' + str(j)].value)  # Радиус 90
                                stroka.append(sheet['N' + str(j)].value)  # Радиус 99
                                stroka.append(sheet['O' + str(j)].value)  # Радиус 100
                            dict.append(stroka)
                    if i == 1:
                        for j in range(4, all_rows):
                            stroka = []

                            obor = str(sheet['C' + str(j)].value).strip()  # Оборудование
                            stroka.append(obor)
                            if not (obor in oborudovanie): oborudovanie.append(obor)

                            ve = str(sheet['E' + str(j)].value).strip()  # Вещество
                            ve = change_gas(ve)

                            s_def_otv = str(sheet['F' + str(j)].value)  # Площадь деф отв

                            state = change_state(s_def_otv, ve)  # Состояние
                            stroka.append(state)

                            stroka.append(ve)
                            if not (ve in ov): ov.append(ve)
                            if i == 1:
                                stroka.append(s_def_otv)
                                if not (s_def_otv in pdo): pdo.append(s_def_otv)
                                stroka.append(round(sheet['G' + str(j)].value, 2))  # Масса
                                stroka.append(sheet['I' + str(j)].value)  # Радиус 10
                                stroka.append(sheet['J' + str(j)].value)  # Радиус 100
                            dict.append(stroka)
                    if i == 0:
                        for j in range(4, all_rows):
                            stroka = []

                            obor = str(sheet['C' + str(j)].value).strip()  # Оборудование
                            stroka.append(obor)
                            if not (obor in oborudovanie): oborudovanie.append(obor)

                            ve = str(sheet['E' + str(j)].value).strip()  # Вещество
                            ve = change_gas(ve)
                            s_def_otv = str(sheet['F' + str(j)].value)  # Площадь деф отв

                            state = change_state(s_def_otv, ve)  # Состояние
                            stroka.append(state)

                            stroka.append(ve)
                            if not (ve in ov): ov.append(ve)

                            stroka.append(s_def_otv)
                            if not (s_def_otv in pdo): pdo.append(s_def_otv)

                            stroka.append(sheet['G' + str(j)].value)  # Масса
                            stroka.append(sheet['I' + str(j)].value)  # Радиус 1
                            stroka.append(sheet['J' + str(j)].value)  # Радиус 10
                            stroka.append(sheet['K' + str(j)].value)  # Радиус 25
                            stroka.append(sheet['L' + str(j)].value)  # Радиус 50
                            stroka.append(sheet['M' + str(j)].value)  # Радиус 90
                            stroka.append(sheet['N' + str(j)].value)  # Радиус 99
                            stroka.append(sheet['O' + str(j)].value)  # Радиус 100
                            dict.append(stroka)
                    #print(dict)

                    if i == 4:
                        dict.sort(key=lambda row: ((row[2]) or 0, row[3] or 0, row[0] or 0))
                        #print(dict)
                        rez = filtery(dict, oborudovanie, ov, pdo, 6, 5)
                    elif i == 6:
                        dict.sort(key=lambda row: ((row[2]) or 0, row[3] or 0, row[0] or 0))
                        #print(dict)
                        rez = filtery(dict, oborudovanie, ov, pdo, 6, 5)
                        for o in range(len(rez)):
                            rez[o].pop(3)
                    elif i == 7:
                        dict.sort(key=lambda row: ((row[0]) or 0, row[2] or 0, row[3] or 0))
                        #print(dict)
                        rez = filtery_for_3(dict, oborudovanie, ov, pdo, 4, 8, 5)
                    elif i == 3:
                        dict.sort(key=lambda row: ((row[0]) or 0))
                        #print(dict)
                        rez = chistka_3_ogn(dict)
                    elif i == 1:
                        dict.sort(key=lambda row: ((row[0]) or 0, row[2] or 0, row[3] or 0))
                        #print(dict)
                        rez = chistka_3_ogn(dict)
                    elif i == 0:
                        dict.sort(key=lambda row: ((row[2]) or 0, row[3] or 0, row[0] or 0))
                        #print(dict)
                        rez = filtery_for_1(dict, oborudovanie, ov, pdo, 5)

                    #print(rez)

                    if i == 4:
                        count_stolb = 11
                    elif i == 6 or i == 3 or i == 0:
                        count_stolb = 10
                    elif i == 7:
                        count_stolb = 7
                    elif i == 1:
                        count_stolb = 5

                    for k in range(len(rez)):
                        list = rez[k]
                        if i == 4: row = doc_table[0].add_row().cells  # олучаем все ячейки ряда
                        if i == 6: row = doc_table[1].add_row().cells  # олучаем все ячейки ряда
                        if i == 7: row = doc_table[2].add_row().cells  # олучаем все ячейки ряда
                        if i == 3: row = doc_table[3].add_row().cells  # олучаем все ячейки ряда
                        if i == 1: row = doc_table[4].add_row().cells  # олучаем все ячейки ряда
                        if i == 0: row = doc_table[5].add_row().cells  # олучаем все ячейки ряда
                        for col in range(count_stolb):
                            cell = row[col]
                            paragraph = cell.paragraphs[0]
                            paragraph.style = doc.styles['Normal']
                            if col == 0:
                                paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                            else:
                                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            if str(list[col]) is None or str(list[col]) == 'None': list[col] = ''
                            run = paragraph.add_run(str(list[col]))
                        doc.save(self.word_path)
                    print("--- %s seconds ---" % (time.time() - start_time))

        self._signal.emit(100)
        time.sleep(2)
        self.finished.emit(0)

    # def terminate(self):
    #     print("terminate func enter")
    #     self.finished.emit(2)
    #     return 123


