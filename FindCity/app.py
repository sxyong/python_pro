import tkinter as tk
from tkinter import filedialog
from FindCity import *

class Application(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack()
        self.createWidgets()
        InitCitys()
        file_path = ''
        

    def createWidgets(self):
        self.lb = tk.Label(self, text='')
        self.lb.pack()
        self.bt = tk.Button(self, text='Open', command=self.Open)
        self.bt.pack()
        self.bt2 = tk.Button(self, text='Run', command=self.Run)
        self.bt2.pack()
        

    def Run(self):
        self.lb.config(text='执行中，请等待')
        hExcel = xlrd.open_workbook(self.file_path)
        #选择sheet
        hSheet = hExcel.sheet_by_index(0)
        nRows = hSheet.nrows

        hExcelNew = copy(hExcel)
        hSheetNew = hExcelNew.get_sheet(0)

        for i in range(nRows):
            #查找第一行地址，如果找到，写入第二行
            a = FindPlace(hSheet.row_values(i)[0])
            if(len(a) > 0):
                hSheetNew.write(i, 1, a[0])
                hSheetNew.write(i, 2, a[1])

            hExcelNew.save(self.file_path)

        self.lb.config(text='finish!')

    def Open(self):
        self.file_path = filedialog.askopenfilename()
        self.lb.config(text=self.file_path)

root = tk.Tk()
root.geometry('300x80')    
app = Application()
app.master.title('sfy')
app.mainloop()