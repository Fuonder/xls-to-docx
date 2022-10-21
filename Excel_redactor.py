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


class Excel:

    def another_e(self, chislo):
        power = int(str(chislo)[-2] + str(chislo)[-1])
        value_new = float(str(chislo * pow(10, power))[:-2])
        return value_new

    def sort_dict(self, dict):
        dictionary_keys = list(dict.keys())
        sorted_dict = {dictionary_keys[x]: sorted(
            dict.values())[x] for x in range(len(dictionary_keys))}
        return sorted_dict

    def edit_excel(self, excel_path, word_path):

        start_time = time.time()

        wb = openpyxl.load_workbook(excel_path)
        sheet_name = (wb.get_sheet_names())
        sheet = wb.get_sheet_by_name(sheet_name[0])
        print(sheet_name[0])
        if sheet_name[0] != 'Scens':
            return 1


        # B - Наименование оборудования
        # D - опасное вещество
        # Е - исход
        # P - диаметр !!!!!!!!!!!!!!!!!!!!
        # F - частота сценария
        # H - масса Гф
        # I - масса Жф

        dict = {}
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
                if math.sqrt(float(q) * 4 / float(math.pi)) * 1000 < 0 or None:
                    stroka.append(0)
                else:
                    q = (math.sqrt(float(q) * 4 / float(math.pi)) * 1000)
                    q = str(q)
                    index = q.find('.')
                    q = q[:index + 3]
                    q = float(q)
                    stroka.append(q)

            with_10_ = self.another_e(sheet['F' + str(i)].value)
            with_10_ = str(with_10_)
            index = with_10_.find('.')
            with_10_ = with_10_[:index + 3]
            with_10_ = float(with_10_)
            stroka.append(with_10_)

            stroka.append(sheet['H' + str(i)].value)
            stroka.append(sheet['I' + str(i)].value)
            dict[j] = stroka
            j = j + 1
            stroka = []

        wb.close()

        dict = self.sort_dict(dict)

        doc = Document('Scene.docx')
        doc_table = doc.tables

        style = doc.styles['Normal']
        style.font.name = 'Times New Roman'
        style.font.size = Pt(10)

        j = 0
        for row in range(2, all_rows-2):
            list = dict[j]
            j = j + 1
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

                # cell = doc_table[0].cell(row, col)
                # cell.text = str(list[col])

        doc.save(word_path)
        print("--- %s seconds ---" % (time.time() - start_time))
        return 0


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
        # print(power)
        # print(value_new)
        # #print(value * pow(10, power))
        #
        # #print(str(value)[-1])
        # os.remove(copy_path)

        # book = openpyxl.load_workbook(excel_path)


