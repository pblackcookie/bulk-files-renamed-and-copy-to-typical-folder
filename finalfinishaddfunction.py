'''
@author : 溺水的饼干
@create_time : 2022/7/29 - 17:00
@function : 批量挂接文件的实现
@version : 2.1_将所有按钮封装在类里面，实现可视化GUI窗口弹窗顺序顺序正常化
 version : 2.2_更改功能
 原功能：将原本文件夹内的文件改名，并且将原文件存放至指定路径文件夹
 现功能：将未改名的文件保留至原路径不动，并将改名后的文件挪至指定文件夹
 version : 2.3_更改excel读取功能 2022/8/2 - 11:00
 逻辑：使用功能读取excel整表，并且删除列表中空行，使程序适应不同的表头
 1.读取excel的表头
 2.使用dataframe删除所有空白列
 3.按行依次读取新生成的xlsx文件并修改（此文件仅存在于程序进程中，不会改变原本xlsx表的数据格式）

'''

# 实现复制文件，读取execl表格文件所用的库
import os
import xlrd
import pandas as pd
from shutil import copyfile
# 实现窗口可视化所引用的库
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

class Appliation(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.master = master
        self.filepath = tk.StringVar()
        self.pack()
        self.main_config() # 对界面的一些基本配置
        self.create_widgets()
        self.create_widgets1()
        self.create_widgets2()
# 窗口的一些基本设置
    def main_config(self):
        #设置窗口大小不可变
        root.resizable(0,0)
        #设置主标题
        self.getFile_bt = tk.Label(self)
        self.getFile_bt['width'] = 50
        self.getFile_bt['height'] = 3
        self.getFile_bt['font'] = ('Consolas', 18)
        self.getFile_bt['fg'] = 'black'
        self.getFile_bt['text'] = "欢迎来到批量更改文件名称系统，请根据需求选择操作"
        self.getFile_bt.pack(side="top")

# 按钮1所实现的功能
    def create_widgets(self):
        #获取文件
        self.getFile_bt = tk.Button(self)
        self.getFile_bt['width'] = 40
        self.getFile_bt['height'] = 3
        self.getFile_bt['font'] = ('Consolas', 18)
        self.getFile_bt['background'] = 'black'
        self.getFile_bt['fg'] = 'white'
        self.getFile_bt['text'] = "请点击此处选择读取数据的excel文件"
        self.getFile_bt['command'] = self._getFile
        self.getFile_bt.pack(side="top")

        #显示文件路径
        self.filePath_en = tk.Entry(self, width = 30)
        self.filePath_en.pack(side="top")

        self.filePath_en.delete(0, "end")
        self.filePath_en.insert(0, "请选择文件")
# 打开文件并且显示路径
    def _getFile(self):
        default_dir = r"文件路径"
        self.filePath = tk.filedialog.askopenfilename(title=u'选择文件',initialdir=(os.path.expanduser(default_dir)))
        print(self.filePath)
        self.filePath_en.delete(0,"end")
        self.filePath_en.insert(0,self.filePath)

# 按钮2所实现的功能
    def create_widgets1(self):
        self.getFile_bt = tk.Button(self)
        self.getFile_bt['width'] = 40
        self.getFile_bt['height'] = 3
        self.getFile_bt['font'] = ('Consolas', 18)
        self.getFile_bt['background'] = 'black'
        self.getFile_bt['fg'] = 'white'
        self.getFile_bt['text'] = '选择要将改名后文件挂载到的目标文件夹'
        self.getFile_bt['command'] = self._getFloder
        self.getFile_bt.pack(side="top")

        # 显示文件夹路径
        self.filePath_en1 = tk.Entry(self, width=30)
        self.filePath_en1.pack(side="top")

        self.filePath_en1.delete(0, "end")
        self.filePath_en1.insert(0, "请选择文件夹")
# 打开文件夹并且显示路径
    def _getFloder(self):
        default_dir =r"文件夹路径"
        self.filePath = tk.filedialog.askdirectory(title=u'选择文件夹',initialdir=(os.path.expanduser(default_dir)))
        print(self.filePath)
        self.filePath_en1.delete(0,"end")
        self.filePath_en1.insert(0,self.filePath)

    def create_widgets2(self):
        self.getFile_bt = tk.Button(self)
        self.getFile_bt['width'] = 40
        self.getFile_bt['height'] = 3
        self.getFile_bt['font'] = ('Consolas', 18)
        self.getFile_bt['background'] = 'black'
        self.getFile_bt['fg'] = 'white'
        self.getFile_bt['text'] = '点击此处开始批量挂载文件'
        self.getFile_bt['command'] = self._transfer
        self.getFile_bt.pack(expand="yes")

        # 显示转化挂载提示
        self.filePath_en2 = tk.Entry(self, width=30)
        self.filePath_en2.pack(side="top")

        self.filePath_en2.delete(0, "end")
        self.filePath_en2.insert(0, "点击上述按钮开始转换文件名称与挂载")
#转化与挂载文件
    def _transfer(self):
        # 获取目标文件数据
        # 使用dataframe删除excel文件中空白行
        df = pd.read_excel(self.filePath_en.get(), index_col=0)
        final_df = df[df.columns.drop(list(df.filter(regex='Unnamed')))]
        # final_df.to_excel(self.filePath_en.get())
        # 在创造出的新document文件中进行数据更改，不破坏原本文档的数据格式
        final_df.to_excel('document.xlsx')
        # aimdata = xlrd.open_workbook(self.filePath_en.get(), "rb")
        aimdata = xlrd.open_workbook('document.xlsx', "rb")
        table = aimdata.sheets()[0]
        # print(type(table))
        aimfolder = self.filePath_en1.get()

        ### 创建一个字典贮存execl表里面的文字
        # def import_execl(execl):
        count = 0
        # 在循环之前创建一个空列表，存储execl的数据
        tables = []
        #获取execl表格的文件行数
        line_data = pd.read_excel(self.filePath_en.get())
        end_line = len(line_data) + 1
        # define array and the data is from the start row in the execl sheet
        # default desire get data form is as the list below
        # array = {'原本文件名': '', '目标文件名': '', '文件存放路径': '', '文件格式': ''}
        # get all data as a list in the row[0]
        # if there is a NaN column, then delete this column
        array = table.row_values(0)
        # print(array)
        for rown in range(1, end_line):
            array[0] = table.cell_value(rown, 0) # initial file name
            # print('原文件名:',array[0])
            array[1] = table.cell_value(rown, 1) # final file name
            # print('目标文件名:',array[1])
            array[2] = table.cell_value(rown, 2) # file store directory
            # print('目标挂载路径：',array[2])
            array[3] = table.cell_value(rown, 3) # file format
            # print('文件类型：',array[3])
            count += 1
            # 复制一份文件至指定文件夹
            # 若要修改指定复制的目标文件夹，更改aimfolder路径
            cpy_path = aimfolder + '\\' + array[1]
            filename = array[2]
            ori_path = filename + '\\' + array[0]
            copyfile(ori_path, cpy_path)
            # 修改复制后文件的名称
            # tar_path 是目标路径，如需换路径，更改filename至指定路径即可
            # tar_path = filename + '\\' + array['原本文件名']
            # os.rename(ori_path, tar_path)
            os.rename(cpy_path, cpy_path + "." + array[3])
            # 打印操作数据一览
            # print("已将" + array[0] + "更改为：" + array[1] + "." + array[3])
            tables.append(array)


        # print("\n\t如上所示，共修改" + str(count) + "条数据\n")
        # print("________________________________________________________________")
        os.remove('document.xlsx')
        self.filePath_en2.delete(0, "end")
        self.filePath_en2.insert(0, "操作成功,恭喜")



root = tk.Tk()
root.title("批量文件处理系统")
root.geometry("640x480+600+300")

app = Appliation(master=root)
#关闭窗口提示框
def on_close():
    if tk.messagebox.askokcancel("提示窗口", "您确定要退出吗？"):
        root.destroy()
root.protocol('WM_DELETE_WINDOW', on_close)
app.mainloop()
