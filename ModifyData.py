import os
import xlrd
import xlsxwriter
import FN

class MD:
    def __init__(self):
        self.dict = {
            3 : [1, 2, 4, 5, 8]
            }
        self.d = os.getcwd() + '\\Datas'
        self.fn = FN.FN(d=self.d + '\\RawXlsx')
        self.fn.analyzeExtensions()
        self.filesName = []
        for i in self.fn.filesName:
            self.filesName.append(self.fn.analyzeFN(i))
            pass

    def test(self):
        for i in self.filesName:
            self.__openXlsx(i)
        pass

    def createXlsx(self):
        for d in self.dict:
            for i in self.dict[d]:
                for t in ['A', 'B']:
                    fileDir = self.d + '\\RawXlsx\\' + 'B' + str(d) + 'Unit' + str(i) + t + '.xlsx'
                    print(fileDir)
                    if not os.path.exists(fileDir):
                        workbook = xlsxwriter.Workbook(fileDir)
                        worksheet = workbook.add_worksheet()
                        workbook.close()

    def __openXlsx(self, fileName):
        data_file = fileName[0] + fileName[1] + fileName[2]
        print(data_file)
        workbook = xlrd.open_workbook(data_file)
        print(workbook.sheet_names())
        sheet1_name = workbook.sheet_names()[0]
        print(sheet1_name)
        sheet1 = workbook.sheet_by_index(0)
        print(sheet1.name, sheet1.nrows, sheet1.ncols)
        if (not (sheet1.nrows and sheet1.ncols)):
            return 
        a = sheet1.cell(0,0).value
        b = sheet1.cell(1,0).value
        print(a, type(a))
        print(b, type(b))

        data_file = self.d + '\\ModifyXlsx\\'  + fileName[1] + fileName[2]
        workbook = xlsxwriter.Workbook(data_file)
        worksheet = workbook.add_worksheet()

        title = ['单词', '音标', '用户释义','词典释义', '词汇特性', '只记词汇', '学习次数', '检测次数', '错误次数', '正确率']
        title_len = len(title)

        for i in range(title_len):
            worksheet.write(0, i ,title[i])

        k = 1
        j = 1
        begin = 0
        while(sheet1.cell(begin,0).value == '??'):
            begin += 1
        while(sheet1.cell(begin,0).value == chr(65279)):
            begin += 1
        for i in range(begin,sheet1.nrows):
            data = sheet1.cell(i, 0).value
            if not data=='':
                data = self.__remove_end_space(data)
                if(not (self.__check_contain_chinese(data) or not data.find('(')== -1)):
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
    
    def __check_contain_chinese(self, check_str):
        for ch in check_str:
            if u'\u4e00' <= ch <= u'\u9fff':
                return True
        return False
    
    def __modifySpaceAndQM(self,s):
        '''
        空格有32的空格和160的空格，32的空格为常见空格，
        这里将160的空格统一转成32的
        '''
        return s.replace(chr(160), chr(32))

    def __remove_end_space(self, s):
        s = self.__modifySpace(s)
        length = len(s)
        while(length and s[length-1] == ' '):
            length -= 1
        return s[:length]

if __name__ == '__main__':
    md = MD()
    #md.createXlsx()
    md.test()