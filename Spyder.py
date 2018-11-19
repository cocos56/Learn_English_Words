import requests
import os
import re

class Spy:
    '''
    本类用于爬取信息
    '''
    def __init__(self):
        self.IndexUrl = r'http://www.iciba.com/'
        self.headers = {
        'User-Agent':'Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/63.0.3239.26 Safari/537.36 Core/1.63.6776.400 QQBrowser/10.3.2601.400'
        }
        self.cwd = os.getcwd()
        self.dir = self.cwd + r'\Datas\Htmls'
        if not os.path.exists(self.dir):
            os.mkdir(self.dir)
        pass

    def getData(self, s):
        text = self.__getHtml(s)
        pattern = '<span>英 \[(.+?)\]</span>.+?<i class="new-speak-step" ms-on-mouseover="sound\(\'(.+?)\'\)\"></i>'
        reg = re.compile(pattern, re.S)
        results = reg.findall(text)
        if results == []:
            pattern = '<span>\[(.+?)\]</span>.+?<i class="new-speak-step" ms-on-mouseover="sound\(\'(.+?)\'\)\"></i>'
            reg = re.compile(pattern, re.S)
            results = reg.findall(text)
        try:
            results[0][0]
        except IndexError:
            return ''
        return results

    def getMeaning(self, s):
        text = self.__getHtml(s)
        pattern = ' <li class="clearfix">(.+?)</li>'
        reg = re.compile(pattern, re.S)
        results = reg.findall(text)
        results2 =  []
        pattern1 = '<span class="prop">(.+?)</span>'
        reg1 = re.compile(pattern1, re.S)
        pattern2 = '<span>(.+?)</span>'
        reg2 = re.compile(pattern2, re.S)
        for i in results:
            temp = reg1.findall(i)
            t = reg2.findall(i)
            temp.append(t)
            results2.append(temp)

        #print(results2)
        return results2
        

    def __getHtml(self, s):
        fileDir = self.dir + '\\' + s + '.html'
        if not os.path.exists(fileDir):
            url = self.IndexUrl + s
            r = requests.get(url, headers=self.headers)
            with open(fileDir,'w',encoding = 'utf-8') as f:
                    f.write(r.text)
                    f.close()
        with open(fileDir, 'r', encoding = 'utf-8') as f:
            text = f.read()
            f.close()
        return text


if __name__ == '__main__':
    spy = Spy()
    print(spy.getMeaning('abort'))