# coding=utf-8
import tkinter
from tkinter import filedialog, messagebox

import xlrd
import xlwt
import importlib,sys
from xlutils.copy import copy as xl_copy
import os

importlib.reload(sys)

class Person(object):

    def __init__(self):
        self._primary_key = ''  # 主键
        self._house_number = ''  # 房号
        self._arrears_period = 0  # 欠费周期
        self._no_take_principal = 0  # 未收本金
        self._take_principal = 0  # 已收本金
        self._should_take_principal = 0  # 应收本金
        self._no_take_invalid_principal = 0  # 未收违约金
        self._take_invalid_principal = 0  # 已收违约金
        self._should_take_invalid_principal = 0  # 应收违约金
        self._all_arrears_number = 0  # 欠费总金额
        self._butler_name = ''  # 管家

    @property
    def primary_key(self):
        return self._primary_key

    @primary_key.setter
    def primary_key(self, primary_key):
        if primary_key is not None and isinstance(primary_key, str) and len(primary_key) > 0:
            self._primary_key = primary_key
        # else:
        #     raise Exception('primary_key can not be none')
        # pass

    @property
    def house_number(self):
        return self._house_number

    @house_number.setter
    def house_number(self, house_number):
        if house_number is not None:
            self._house_number = house_number
        else:
            raise Exception('house_number can not be none')
        pass

    @property
    def arrears_period(self):
        return self._arrears_period

    @arrears_period.setter
    def arrears_period(self, arrears_period):
        if arrears_period >= 0:
            self._arrears_period = arrears_period
        else:
            raise Exception('arrears_period should >= 0 or type is not int')
        pass

    @property
    def no_take_principal(self):
        return self._no_take_principal

    @no_take_principal.setter
    def no_take_principal(self, no_take_principal):
        if no_take_principal >= 0:
            self._no_take_principal = no_take_principal
        else:
            raise Exception('no_take_principal should >= 0')
        pass

    @property
    def take_principal(self):
        return self._take_principal

    @take_principal.setter
    def take_principal(self, take_principal):
        if take_principal >= 0:
            self._take_principal = take_principal
        else:
            raise Exception('take_principal should >= 0')
        pass

    @property
    def no_take_invalid_principal(self):
        return self._no_take_invalid_principal

    @no_take_invalid_principal.setter
    def no_take_invalid_principal(self, no_take_invalid_principal):
        if no_take_invalid_principal >= 0:
            self._no_take_invalid_principal = no_take_invalid_principal
        else:
            raise Exception('no_take_invalid_principal should >= 0')
        pass

    @property
    def take_invalid_principal(self):
        return self._take_invalid_principal

    @take_invalid_principal.setter
    def take_invalid_principal(self, take_invalid_principal):
        if take_invalid_principal >= 0:
            self._take_invalid_principal = take_invalid_principal
        else:
            raise Exception('no_take_invalid_principal should >= 0')
        pass

    @property
    def should_take_invalid_principal(self):
        return self._should_take_invalid_principal

    @should_take_invalid_principal.setter
    def should_take_invalid_principal(self, should_take_invalid_principal):
        if should_take_invalid_principal >= 0:
            self._should_take_invalid_principal = should_take_invalid_principal
        else:
            raise Exception('should_take_invalid_principal should >= 0')
        pass

    @property
    def all_arrears_number(self):
        return self._all_arrears_number

    @all_arrears_number.setter
    def all_arrears_number(self, all_arrears_number):
        if all_arrears_number >= 0:
            self._all_arrears_number = all_arrears_number
        else:
            raise Exception('all_arrears_number should >= 0')
        pass

    @property
    def should_take_principal(self):
        return self._should_take_principal

    @should_take_principal.setter
    def should_take_principal(self, should_take_principal):
        if should_take_principal >= 0:
            self._should_take_principal = should_take_principal
        else:
            raise Exception('all_arrears_number should >= 0')
        pass

    @property
    def butler_name(self):
        return self._butler_name

    @butler_name.setter
    def butler_name(self, butler_name):
        if butler_name is not None and isinstance(butler_name, str) and len(butler_name) > 0:
            self._butler_name = butler_name
        else:
            raise Exception('butler_name can not be none')
        pass

    @staticmethod
    def get_house_index():
        return 0

    @staticmethod
    def get_arrears_period_index():
        return Person.get_house_index() + 1

    @staticmethod
    def get_no_take_principal_index():
        return Person.get_arrears_period_index() + 1

    @staticmethod
    def get_no_take_invalid_principal_index():
        return Person.get_no_take_principal_index() + 1

    @staticmethod
    def get_all_arrears_number_index():
        return Person.get_no_take_invalid_principal_index() + 1

    @staticmethod
    def get_butler_name_index():
        return Person.get_all_arrears_number_index() + 1


excel_file = None
excel_src_path = ''


def create_sheet_style():
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = 'Arial'
    font.bold = False
    font.colour_index = 4
    font.height = 10
    alignment = xlwt.Alignment()
    alignment.horz = 0x02
    alignment.vert = 0x01
    style.alignment = alignment
    return style
    pass


def write_to_new_sheet(all_persons):
    style = create_sheet_style()
    work_book = xl_copy(excel_file)
    new_sheet = work_book.add_sheet(u'二号地欠费汇总1', False)
    row0 = [u'房号', u'欠费周期', u'未收本金', u'未收违约金', u'欠费总金额', u'管家']
    for i in range(len(row0)):
        new_sheet.write(0, i, row0[i], style)
    persons = list(all_persons.values())
    sorted(persons, key=lambda person: person.house_number[0:1])
    for i in range(len(persons)):
        p = persons[i]
        new_sheet.write(i + 1, Person.get_house_index(), p.house_number, style)
        new_sheet.write(i + 1, Person.get_arrears_period_index(), p.arrears_period, style)
        new_sheet.write(i + 1, Person.get_no_take_principal_index(), p.no_take_principal, style)
        new_sheet.write(i + 1, Person.get_no_take_invalid_principal_index(), p.no_take_invalid_principal,
                        style)
        new_sheet.write(i + 1, Person.get_all_arrears_number_index(), p.no_take_principal + p.no_take_invalid_principal,
                        style)
        new_sheet.write(i + 1, Person.get_butler_name_index(), u'', style)
    pass
    work_book.save(excel_src_path)


def read_excel():
    all_persons = {}
    property_map = {}
    sheet_0 = excel_file.sheet_by_index(0)

    for i in range(sheet_0.ncols):
        column_name = sheet_0.cell(4, i).value
        if column_name == '房产主键':
            property_map['primary_key'] = i
        elif column_name == '房产':
            property_map['house_number'] = i
        elif column_name == '欠费周期':
            property_map['arrears_period'] = i
        elif column_name == '应收本金':
            property_map['should_take_principal'] = i
        elif column_name == '应收违约金':
            property_map['should_take_invalid_principal'] = i
        elif column_name == '已收本金':
            property_map['take_principal'] = i
        elif column_name == '已收违约金':
            property_map['take_invalid_principal'] = i
        elif column_name == '未收本金':
            property_map['no_take_principal'] = i
        elif column_name == '未收违约金':
            property_map['no_take_invalid_principal'] = i
    row_count = sheet_0.nrows
    for i in range(row_count - 2):
        if i < 5:
            continue
        house_number = sheet_0.cell(i, property_map['house_number']).value
        if house_number is None:
            continue
        if all_persons.__contains__(house_number):
            p = all_persons[house_number]
        else:
            p = Person()
            all_persons[house_number] = p
            p.house_number = house_number
        p.arrears_period += 1
        p.take_principal = p.take_principal + float(sheet_0.cell(i, property_map['take_principal']).value)
        p.should_take_principal = p.should_take_principal + float(sheet_0.cell(i, property_map[
            'should_take_principal']).value)
        p.no_take_principal = p.no_take_principal + float(sheet_0.cell(i, property_map['no_take_principal']).value)
        p.take_invalid_principal = p.take_invalid_principal + float(sheet_0.cell(i, property_map[
            'take_invalid_principal']).value)
        p.should_take_invalid_principal = p.should_take_invalid_principal + float(sheet_0.cell(i, property_map[
            'should_take_invalid_principal']).value)
        p.no_take_invalid_principal = p.no_take_invalid_principal + float(sheet_0.cell(i, property_map[
            'no_take_invalid_principal']).value)
        pass
    pass
    write_to_new_sheet(all_persons)


def execute(excel_path):
    if not os.path.exists(excel_path):
        return False
    global excel_src_path
    excel_src_path = excel_path
    global excel_file
    excel_file = xlrd.open_workbook(excel_path)
    try:
        read_excel()
        return True
    except Exception:
        return False
    pass

class Application(tkinter.Frame):
    def __init__(self, master=None):
        tkinter.Frame.__init__(self, master)
        self.pack()
        self.createWidgets()

    def createWidgets(self):
        self.tip_label = tkinter.Label(self, text='选择要处理的Excel文件地址：')
        self.tip_label.pack()
        self.execute_button = tkinter.Button(self, text='选择文件', command=self.select_file)
        self.execute_button.pack()
        self.path_input = tkinter.Entry(self)
        self.path_input.pack()
        self.execute_button = tkinter.Button(self, text='执行', command=self.handle_excel)
        self.execute_button.pack()
        pass

    def select_file(self):
        file_name = filedialog.askopenfilename()
        self.path_input.delete(0, tkinter.END)
        if file_name != '':
            self.path_input.insert(0, file_name)
        pass

    def handle_excel(self):
        excel_path = self.path_input.get()
        result = execute(excel_path)
        if result:
            messagebox.showinfo('执行成功')
            self.path_input.delete(0, tkinter.END)
        else:
            messagebox.showerror('执行失败')
        pass


app = Application()
app.master.title('珂珂珂')
app.mainloop()
