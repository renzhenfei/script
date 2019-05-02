# coding=utf-8
import tkMessageBox
from Tkinter import *
import tkFileDialog
import parse_script as kk


class Application(Frame):
    def __init__(self, master=None):
        Frame.__init__(self, master)
        self.pack()
        self.createWidgets()

    def createWidgets(self):
        self.tip_label = Label(self, text='选择要处理的Excel文件地址：')
        self.tip_label.pack()
        self.execute_button = Button(self, text='选择文件', command=self.select_file)
        self.execute_button.pack()
        self.path_input = Entry(self)
        self.path_input.pack()
        self.execute_button = Button(self, text='执行', command=self.handle_excel)
        self.execute_button.pack()
        pass

    def select_file(self):
        file_name = tkFileDialog.askopenfilename()
        self.path_input.delete(0, END)
        if file_name != '':
            self.path_input.insert(0, file_name)
        pass

    def handle_excel(self):
        excel_path = self.path_input.get()
        result = kk.execute(excel_path)
        if result:
            tkMessageBox.showinfo('执行成功')
            self.path_input.delete(0, END)
        else:
            tkMessageBox.showerror('执行失败')
        pass


app = Application()
app.master.title('珂珂珂')
app.mainloop()
