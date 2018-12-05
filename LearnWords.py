import xlrd
from playsound import playsound
import os
import time
import pygame
import threading
import sys
import subprocess
import json
import datetime
import re

class LW():
        def __init__(self, cwd, myWordsData, workbookDir):
                self.cwd = cwd
                self.myWordsData = myWordsData
                self.workbook = xlrd.open_workbook(r'D:\0COCO\本科\大二上学期\英语三\Datas\ModifyXlsx\B3.xlsx')
                self.worksheet = ['8A']
                self.tasks = 0
                self.count = 0
                self.words=[]
                mp3dir = self.cwd + '\\Datas\\Mp3\\'
                os.chdir(mp3dir)

        def __del__(self):
                os.chdir(self.cwd)
                pass
        
        def __updatReadingTimes(self, word):
                t = []
                if word in self.myWordsData:
                        times = self.myWordsData[word][0] + 1
                        t.append(times) #浏览次数
                        t.append(self.myWordsData[word][1])     #跟敲次数
                        t.append(self.myWordsData[word][2])     #跟敲正确次数
                        t.append(self.myWordsData[word][3])     #默写次数
                        t.append(self.myWordsData[word][4])     #默写正确次数
                else:
                        times = 1
                        t.append(times) #浏览次数
                        t.append(0)     #跟敲次数
                        t.append(0)     #跟敲正确次数
                        t.append(0)     #默写次数
                        t.append(0)     #默写正确次数
                self.myWordsData.update({word:t})

        def __getWords(self):
                for name in self.worksheet:
                        sheet = self.workbook.sheet_by_name(name)
                        self.tasks += (sheet.nrows-1)
                        for i in range(1, sheet.nrows):
                                word = []
                                for j in range(sheet.ncols):
                                        word.append(sheet.cell(i,j).value)
                                self.words.append(word)

        def learnWithTime(self, Time=5):
                Time = Time*60
                beginT = time.time()
                self.__getWords()
                num = 0
                while(self.count < self.tasks):
                        rT = 1
                        while(rT<=3):
                                self.__readingWord(rT)
                                r = self.__checkWithAll()
                                while(r == 0):
                                        r = self.__checkWithAll()
                                if(r == 2):
                                        break
                                rT += 1
                        if(r == 2):
                                break
                        num += 1
                        if (time.time-beginT<Time and not self.count==self.tasks-1):
                                self.count+=1
                        else:
                                print('\n\n您已进入检测模式，请根据中文释义拼写出英文词汇')
                                Number = num
                                while(not num==0):
                                        n = Number-num
                                        print(n+1, '//', Number)
                                        r = self.__checkWithMeaning(n)
                                        while(r == 0):
                                                r = self.__checkWithMeaning(n)
                                        if(r == 2):
                                                break
                                        num -= 1
                                if(r == 2):
                                        break
                return self.count + 1

        def learnWithNumber(self, Number=6):
                self.__getWords()
                num = 0
                while(True):
                        rT = 1
                        while(rT<=3):
                                self.__readingWord(rT)
                                r = self.__checkWithAll()
                                while(r == 0):
                                        r = self.__checkWithAll()
                                if(r == 2):
                                        break
                                rT += 1
                        if(r == 2):
                                break
                        num += 1
                        if (num<Number and not self.count==self.tasks-1):
                                self.count+=1
                        else:
                                print('\n\n您已进入检测模式，请根据中文释义拼写出英文词汇')
                                while(not num==0):
                                        print(Number-num+1, '/', Number)
                                        r = self.__checkWithMeaning(num-1)
                                        while(r == 0):
                                                rT = 1
                                                while(rT<=3):
                                                        self.__readingWord(rT, num-1)
                                                        r = self.__checkWithAll(num-1)
                                                        while(r == 0):
                                                                r = self.__checkWithAll(num-1)
                                                        if(r == 2):
                                                                break
                                                        rT += 1
                                                if(r == 2):
                                                        break
                                                self.words.append(self.words[self.count-num+1])
                                        if(r == 2):
                                                break
                                        num -= 1
                                if(r == 2):
                                        break
                                self.count += 1
                        if(self.count == self.tasks-1):
                                if(len(self.words)==self.tasks):
                                        break
                                else:
                                        self.words = self.words[self.tasks:]
                                        self.tasks = len(self.words)
                                        print('\n请重新学习刚才拼错的单词')
                                        self.count = 0
                return self.count + 1

        def __checkWithMeaning(self, n):
                word = self.words[self.count-n]
                print(word[2])
                s = input()
                flag = 1
                times = self.myWordsData[word[0]][3]
                times2 = self.myWordsData[word[0]][4]
                if s == word[0]:
                        times += 1
                        times2 += 1
                        print('恭喜，拼写正确')
                        # self.__play2(r'D:\0COCO\本科\大二上学期\英语三\Datas\right.mp3', 0.2)
                        # time.sleep(1)
                        mp3file = word[0] + '.mp3'
                        self.__PlayMp3(mp3file)
                elif s == '##quit':
                        flag = 2
                else:
                        times += 1
                        print('抱歉，拼写错误，您刚才输入的值为：', s)
                        flag = 0
                        self.__play2(r'D:\0COCO\本科\大二上学期\英语三\Datas\wrong.mp3', 0.6)
                        time.sleep(1)
                text = '已拼写{}次，拼写正确{}次'.format(times,times2)
                t = []
                t.append(self.myWordsData[word[0]][0]) #浏览次数
                t.append(self.myWordsData[word[0]][1])     #跟敲次数
                t.append(self.myWordsData[word[0]][2])     #跟敲正确次数
                t.append(times)     #默写次数
                t.append(times2)    #默写正确次数
                self.myWordsData.update({word[0]:t})
                print(text)
                return flag

        def __readingWord(self, rT, offset = 0):
                word = self.words[self.count-offset]
                self.__updatReadingTimes(word[0])
                mp3file = word[0] + '.mp3'
                text = '\n{}/{}：({}/{})'.format(self.count+1-offset, self.tasks, rT, 3)
                print(text)
                print(word[0], '\n[', word[1], '] ', word[2], sep = '')
                self.__getSpecialProperty(word)
                text = '\"{}\"已浏览{}次'.format(word[0], self.myWordsData[word[0]][0])
                print(text)
                if(os.path.exists(mp3file)):
                        t = threading.Thread(target=self.__PlayMp3, args= (mp3file,))
                        t.start()

        def __getSpecialProperty(self, word, showS = True, showS2 = True):
                s = '词汇特性：'
                s2 = '特性释义：'
                if(not word[4]==''):
                        s += '考研'
                        s2 += '考研：'
                        s2 += word[4]
                if(not word[5]==''):
                        if(not s == '词汇特性：'):
                                s += '、'
                                s2 += '    '
                        s += '六级'
                        s2 += '六级：'
                        s2 += word[5]
                if(not word[6]==''):
                        if(not s == '词汇特性：'):
                                s += '、'
                                s2 += '    '
                        s += '四级'
                        s2 += '四级：'
                        s2 += word[6]
                if(not word[7]==''):
                        if(not s == '词汇特性：'):
                                s += '、'
                                s2 += '    '
                        s += '高考'
                        s2 += '高考：'
                        s2 += word[7]
                if(not s == '词汇特性：'):
                        s += '词汇'
                else:
                        s += '无'
                if(not s == '特性释义：'):
                        pass
                else:
                        s2 += '无'
                if(showS):
                        print(s)
                if(showS2):
                        print(s2)

        def __checkWithAll(self, offset = 0):
                word = self.words[self.count-offset]
                s = input()
                flag = 1
                times = self.myWordsData[word[0]][1]
                times2 = self.myWordsData[word[0]][2]
                if not s == '':
                        if s == word[0]:
                                times += 1
                                times2 += 1
                                print('恭喜，跟敲正确')
                                # self.__play2(r'D:\0COCO\本科\大二上学期\英语三\Datas\right.mp3', 0.2)
                                # time.sleep(1)
                                mp3file = word[0] + '.mp3'
                                if(os.path.exists(mp3file)):
                                        t = threading.Thread(target=self.__PlayMp3, args= (mp3file,True))
                                        t.start()
                                        t.join()
                        elif s == '##quit':
                                flag = 2
                        else:
                                times += 1
                                print('抱歉，跟敲错误，您刚才输入的值为：', s)
                                flag = 0
                                self.__play2(r'D:\0COCO\本科\大二上学期\英语三\Datas\wrong.mp3', 0.6)
                        text = '已跟敲{}次，跟敲正确{}次'.format(times,times2)
                        t = []
                        t.append(self.myWordsData[word[0]][0]) #浏览次数
                        t.append(times)     #跟敲次数
                        t.append(times2)     #跟敲正确次数
                        t.append(self.myWordsData[word[0]][3])     #默写次数
                        t.append(self.myWordsData[word[0]][4])     #默写正确次数
                        self.myWordsData.update({word[0]:t})
                        print(text)
                return flag

        def __play2(self, filename, volume = 1):
                pygame.mixer.init()
                pygame.mixer.music.load(filename)
                pygame.mixer.music.set_volume(volume)
                pygame.mixer.music.play()

        def __play(self, filename):
                playsound(filename, block = True)

        def __PlayMp3(self, filename, wait = False):
                command = ["ffprobe.exe", "-i", filename]
                result = subprocess.Popen(command,shell=True,stdout = subprocess.PIPE, stderr = subprocess.STDOUT)
                out = result.stdout.read()
                temp = str(out.decode('utf-8'))
                #print(temp)
                if(not temp.find('Estimating duration from bitrate, this may be inaccurate') == -1):
                        self.__play2(filename, volume = 1)
                else:
                       self.__play(filename)
                if(wait):
                        temp = re.findall(r'Duration: (\d+):(\d+):(\d+).(\d+),', temp)[0]
                        seconds = float('0.' + temp[3])
                        seconds += float(temp[2])
                        seconds += float(temp[1]*60)
                        seconds += float(temp[0]*3600)
                        time.sleep(seconds+1)