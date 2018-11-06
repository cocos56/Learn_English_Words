import os
import xlrd
import xlsxwriter
import FN

class MD:
    def __init__(self):
        self.d = os.getcwd() + '\\Datas'
        self.fn = FN.FN(d=self.d)
        self.fn.analyzeExtensions()
        self.filesName = []
        for i in self.fn.filesName:
            self.filesName.append(self.fn.analyzeFN(i))
            pass

    def test(self):
        for i in self.filesName:
            if(not i[1] == '0000' and i[2] == '.xlsx'):
                self.__openXlsx(i)
        pass

    def __openXlsx(self, fileName):
        data_file = fileName[0] + fileName[1] + fileName[2]
        workbook = xlrd.open_workbook(data_file)
        print(workbook.sheet_names())
        sheet1_name = workbook.sheet_names()[0]
        print(sheet1_name)
        sheet1 = workbook.sheet_by_index(0)
        print(sheet1.name, sheet1.nrows, sheet1.ncols)
        a = sheet1.cell(0,0).value
        b = sheet1.cell(1,0).value
        print(sheet1.cell(0,0).value)
        print(sheet1.cell(1,0).value)

        data_file = fileName[0] + '\\Modify\\' + fileName[1] + fileName[2]
        workbook = xlsxwriter.Workbook(data_file)
        worksheet = workbook.add_worksheet()

        title = ['单词', '音标', '用户释义','词典释义', '词汇特性', '学习次数', '检测次数', '错误次数', '正确率']
        title_len = len(title)

        for i in range(title_len):
            worksheet.write(0, i ,title[i])

        k = 1
        j = 1
        for i in range(sheet1.nrows):
            data = self.__remove_end_space(sheet1.cell(i, 0).value)
            if(not self.__check_contain_chinese(data)):
                worksheet.write(k, 0, data)
                k += 1
                f = 1
            else:
                worksheet.write(j, 2, data)
                j += 1
                f += 1
                if(f>2):
                    k += 1

        workbook.close()

    def __others(self):
        pass
    
    def __check_contain_chinese(self, check_str):
        for ch in check_str:
            if u'\u4e00' <= ch <= u'\u9fff':
                return True
        return False
    
    def __remove_end_space(self, s):
        length = len(s)
        space_ords = [32, 160] #空格有32的空格和160的空格，32的空格为常见空格
        while(length and ord(s[length-1]) in space_ords):
            length -= 1
        
        s = s[:length]
        return s