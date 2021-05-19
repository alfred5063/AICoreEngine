
# Python bytecode 3.7 (3394)
# Embedded file name: C:\Users\junshou.leow\source\repos\RoboticProcessAutomation\RPA.DTI\TESTING\junshou\mediclinic\transformation\generate_dc
# .py
# Size of source mod 2**32: 5439 bytes
# Decompiled by https://python-decompiler.com
import pandas as pd
from xlutils.copy import copy
from xlrd import *
from xlwt import *
import pandas as pd
from PIL import Image
from io import BytesIO
from utils.audit_trail import audit_log

def build_rb(file, formatting_info=False):
    if formatting_info == True:
        rb = open_workbook(file, formatting_info=True)
    else:
        rb = open_workbook(file)
    return rb


def build_wb(file, formatting_info=False):
    if formatting_info == True:
        rb = open_workbook(file, formatting_info=True)
    else:
        rb = open_workbook(file)
    wb = copy(rb)
    return wb


def savefile(wb, path):
    wb.save(path)


def generate_dc(bord_number, ammount, bord_date, running_no, name, bdx_type, dc_template_path, path,medi_base):
    file_name = running_no + ' ' + bdx_type.upper()
    path = path + '\\' + file_name + '.xls'
    name = '(' + name + ')'
    template = dc_template_path
    if False:
      print()
    else:
        print(template)
        try:
          wb = build_wb(template, formatting_info=True)
        except:
          audit_log("Template for Disbursement Claim Listing is not FOUND", "Completed...", medi_base)
          audit_log("No Disbursement Claim Listing is generated", "Completed...", medi_base)
          return ""
        w_sheet = wb.get_sheet(0)
        style10 = XFStyle()
        borders10 = Borders()
        borders10.left = Borders.THIN
        style10.borders = borders10
        style = XFStyle()
        font = Font()
        borders = Borders()
        borders.left = Borders.THIN
        borders.right = Borders.THIN
        style.borders = borders
        font.name = 'Tahoma'
        font.height = 200
        style.num_format_str = '"RM" #,##0.00'
        font.bold = True
        style.font = font
        Alignment99 = Alignment()
        Alignment99.horz = Alignment.HORZ_CENTER
        Alignment99.vert = Alignment.VERT_BOTTOM
        style.alignment.wrap = 5
        style.alignment = Alignment99
        style5 = XFStyle()
        font5 = Font()
        borders5 = Borders()
        borders5.left = Borders.THIN
        borders5.top = Borders.THIN
        borders5.bottom = Borders.THIN
        borders5.right = Borders.THIN
        style5.borders = borders5
        font5.name = 'Tahoma'
        font5.height = 200
        style5.num_format_str = '"RM" #,##0.00'
        font5.bold = True
        style5.font = font5
        Alignment6 = Alignment()
        Alignment6.horz = Alignment.HORZ_CENTER
        Alignment6.vert = Alignment.VERT_BOTTOM
        style5.alignment.wrap = 5
        style5.alignment = Alignment6
        style4 = XFStyle()
        font4 = Font()
        borders4 = Borders()
        borders4.left = Borders.THIN
        borders4.right = Borders.THIN
        borders4.top = Borders.THIN
        borders4.bottom = Borders.THIN
        style4.borders = borders4
        Alignment4 = Alignment()
        Alignment4.horz = Alignment.HORZ_CENTER
        Alignment4.vert = Alignment.VERT_BOTTOM
        style4.alignment.wrap = 5
        style4.alignment = Alignment4
        font4.name = 'Arial'
        font4.height = 200
        style4.font = font4
        style3 = XFStyle()
        font3 = Font()
        Alignment3 = Alignment()
        Alignment3.horz = Alignment.HORZ_LEFT
        Alignment3.vert = Alignment.VERT_BOTTOM
        style3.alignment.wrap = 5
        style3.alignment = Alignment4
        font3.name = 'Arial'
        font3.height = 160
        style3.font = font3
        style2 = XFStyle()
        font2 = Font()
        Alignment2 = Alignment()
        Alignment2.horz = Alignment.HORZ_CENTER
        Alignment2.vert = Alignment.VERT_BOTTOM
        style2.alignment.wrap = 5
        style2.alignment = Alignment2
        font2.name = 'Tahoma'
        font2.height = 200
        style2.font = font2
        rb = open_workbook(template, formatting_info=True)
        readerSheet = rb.sheet_by_index(0)
        row_running_no = 0
        row_bord_number = 0
        row_name = 0
        row_ammount = 0
        row_total_ammount = 0

        style8 = XFStyle()
        font8 = Font()
        Alignment8 = Alignment()
        Alignment8.horz = Alignment.HORZ_LEFT
        Alignment8.vert = Alignment.VERT_BOTTOM
        style8.alignment.wrap = 5
        style8.alignment = Alignment8
        style8.name = 'Arial'
        style8.height = 160
        style8.font = font3
        style8.num_format_str = '0.00'
        
        for i in range(100):
            if readerSheet.cell(i, 3).value == 'Disbursement Claims No':
                row_running_no = i + 1
            if readerSheet.cell(i, 3).value == 'Bordereaux No':
                row_bord_number = i + 1
            if readerSheet.cell(i, 5).value == 'ASIA ASSISTANCE NETWORK (M) SDN BHD' or readerSheet.cell(i, 6).value == 'ASIA ASSISTANCE NETWORK (M) SDN BHD' or readerSheet.cell(i, 7).value == 'ASIA ASSISTANCE NETWORK (M) SDN BHD':
                row_name = i + 4
                break
            if readerSheet.cell(i, 0).value == 'Total claim incurred' or readerSheet.cell(i, 0).value == 'Total claim incurred ':
                row_ammount = i
            if readerSheet.cell(i, 7).value == 'Grand Total:':
                row_total_ammount = i

        formula = 'I' + str(row_ammount + 1)
        for i in range(5, row_running_no - 2):
            w_sheet.write(i, 6, '', style10)

        w_sheet.write(row_running_no, 3, running_no, style4)
        w_sheet.write(row_running_no, 8, bord_date, style4)
        w_sheet.write(row_bord_number, 3, bord_number, style3)

        print('Amount to be list: {0}'.format(ammount))
        try:
          w_sheet.write(row_bord_number, 7, float("{0:.2f}".format(float(ammount))), style8)
        except Exception as error:
          w_sheet.write(row_bord_number, 7, float(ammount), style8)
          print(error)
          
        w_sheet.write(row_name, 6, name, style2)
        w_sheet.write(row_ammount, 8, Formula('SUM(H27:H30)'), style)
        w_sheet.write(row_total_ammount, 8, Formula('SUM({})'.format(formula)), style5)
        savefile(wb, path)
