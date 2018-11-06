import os

class FN:
    def __init__(self, d = os.path.dirname(os.getcwd())):
        self.dir = d
        self.filesName = []
        self.extensions = {}
        self.filesNum = []

    def analyzeExtensions(self):
            self.filesName = []
            self.__getFilesName()
            for i in self.filesName:
                e = self.analyzeFN(i)[2]
                if e not in self.extensions:
                    self.extensions.update({e:1})
                else:
                    self.extensions.update({e:self.extensions[e]+1})
            self.__getFilesNum()

    def analyzeFN(self, s):
        '''
        s = [
            '01_考研形势',
            '01_考研形势.mp4',
            'D:\\01_考研形势',
            'D:\\01_考研形势.mp4'
        ]
        '''
        
        t = ['', '', '']    #t[0] 文件路径, t[1] 文件名（不含后缀）, t[2] 文件后缀（含'.'）
        length = len(s)
        
        if(length == 0):
            return t

        while(s[length-1] != '\\'):
            if(s[length-1] == '.'):
                break
            elif(length == 1):
                t[1] = s    #传入的字符串为纯文件名
                return t
            length -= 1

        if(s[length-1] == '.'): #传入的字符串含后缀名
            t[2] = s[length-1:]
            s =  s[:length-1]
            length -= 1
        elif(s[length-1] == '\\'):  #传入的字符串不含后缀名
            t[1] = s[length:]
            t[0] = s[:length]
            return t
        
        if(len(s) == 0):
            return t
        while(s[length - 1] != '\\'):
            if(length == 1):
                t[1] = s
                return t
            length -= 1
        t[1] = s[length:]
        t[0] = s[:length]
        return t

    def __getFilesNum(self):
        sum = 0
        for k in self.extensions:
            st = str(k) + ':' + str(self.extensions[k])
            self.filesNum.append(st)
            sum += self.extensions[k]
        st = 'sum=' + str(sum)
        self.filesNum.append(st)

    def __getFilesName(self):
        for root, dirs, files in os.walk(self.dir):
            #print(root)    # 当前目录路径
            #print(dirs)    # 当前路径下的所有子目录
            #print(files)   # 当前目录下的所有非目录子文件
            if(not root == self.dir):
                break
            for i in files:
                temp = root + '\\' + i
                if(i != ''):
                    self.filesName.append(temp)

if __name__ == '__main__':
    d = os.getcwd() + '\\Datas'
    fn = FN(d=d)
    fn.analyzeExtensions()
    for i in fn.filesName:
        print(i)
    print(fn.filesNum)