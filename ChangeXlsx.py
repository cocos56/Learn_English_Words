from openpyxl import load_workbook
import os
import xlrd
import Spyder
from openpyxl.styles import PatternFill, Border, Side, Alignment, Protection, Font
import FN
import ModifyData
import wget

class ChangeXlsx():
    def __init__(self, recreate = False):
        self.d = os.getcwd() + '\\Datas'
        self.fn = FN.FN(d=self.d + '\\ModifyXlsx')
        self.fn.analyzeExtensions()
        self.font = Font(name='Arial',
                 size=11,
                 bold=False,
                 italic=False,
                 vertAlign=None,
                 underline='none',
                 strike=False,
                 color='FF000000')
        self.space = [' ', ' ']
        if(recreate):
            md = ModifyData.MD()
            md.test()
        self.spy = Spyder.Spy()
        if(not os.path.exists(self.d + '\\Mp3')):
            os.mkdir(self.d + '\\Mp3')
        pass

    def addPhoneticSymbol(self):
        for file_dir in self.fn.filesName:
            data = xlrd.open_workbook(file_dir)
            table = data.sheets()[0]

            wb = load_workbook(filename= file_dir)
            ws = wb[wb.sheetnames[0]]
            for i in range(2, table.nrows+1):
                if not ws.cell(row=i, column=1).value == None:
                    ws.cell(row=i, column=1).value = ws.cell(row=i, column=1).value.replace('‘', '\'')
                    ws.cell(row=i, column=1).value = ws.cell(row=i, column=1).value.replace('’', '\'')
                if(self.__isWord(ws.cell(row=i, column=1).value)):
                    ws.cell(row=i, column=2).font = self.font
                    print(ws.cell(row=i, column=1).value)
                    ps = self.spy.getData(ws.cell(row=i, column=1).value)
                    print(ps)
                    if not ps == '':
                        ws.cell(row=i, column=2).value = ps[0][0]
                        self.__getMp3(ws.cell(row=i, column=1).value, ps[0][1])
                pass
            wb.save(file_dir)

    def addMeaning(self):
        for file_dir in self.fn.filesName:
            data = xlrd.open_workbook(file_dir)
            table = data.sheets()[0]

            wb = load_workbook(filename= file_dir)
            ws = wb[wb.sheetnames[0]]
            for i in range(2, table.nrows+1):
                if(self.__isWord(ws.cell(row=i, column=1).value)):
                    print(ws.cell(row=i, column=1).value)
                    ps = self.spy.getMeaning(ws.cell(row=i, column=1).value)
                    str1 = ''
                    for i1 in ps:
                        str2 = i1[0] + '  '
                        str3 = ''
                        for i3 in i1[1]:
                            str3 += i3
                        str2 += str3
                        str1 += str2
                    print(str1)
                    ws.cell(row=i, column=4).value = str1
                pass
            wb.save(file_dir)

    def __getMp3(self, word, url):
        if(not os.path.exists(self.d + '\\Mp3\\' + word + '.mp3')):
            wget.download(url, out=self.d + '\\Mp3\\' + word + '.mp3')
        pass


    def __isWord(self, str):
        if str == None:
            return False
        if not str == '':
            return  str.find(' ') == -1

if __name__ == '__main__':
    pass