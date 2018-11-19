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

def check(word, myWordsData):
        s = input()
        flag = 1
        times = myWordsData[word][1]
        times2 = myWordsData[word][2]
        if not s == '':
                if s == word:
                        times += 1
                        times2 += 1
                        print('恭喜，跟敲正确')
                        play2(r'D:\0COCO\本科\大二上学期\英语三\Datas\right.mp3')
                        time.sleep(1)
                elif s == '##quit':
                        flag = 2
                else:
                        times += 1
                        print('抱歉，跟敲错误，您刚才输入的值为：', s)
                        flag = 0
                        play2(r'D:\0COCO\本科\大二上学期\英语三\Datas\wrong.mp3')
                text = '已跟敲{}次，跟敲正确{}次'.format(times,times2)
                t = []
                t.append(myWordsData[word][0]) #浏览次数
                t.append(times)     #跟敲次数
                t.append(times2)     #跟敲正确次数
                t.append(myWordsData[word][3])     #默写次数
                t.append(myWordsData[word][4])     #默写正确次数
                myWordsData.update({word:t})
                print(text)
        return flag

def play2(file):
        pygame.mixer.init()
        track = pygame.mixer.music.load(file)
        pygame.mixer.music.play()
def play(file):
        playsound(file, block = True)

def learn(myWordsData):
        workbook = xlrd.open_workbook(r'D:\0COCO\本科\大二上学期\英语三\Datas\ModifyXlsx\B3Unit5A.xlsx')
        sheet1 = workbook.sheet_by_index(0)
        print(sheet1.nrows,sheet1.ncols)

        total = 0
        for i in range(1, sheet1.nrows):
                if not sheet1.cell(i,0).value == '':
                        total += 1
        c = 0
        mp3dir = cwd + '\\Datas\\Mp3\\'
        os.chdir(mp3dir)
        for i in range(1, sheet1.nrows):
                if not sheet1.cell(i,0).value == '':
                        word = sheet1.cell(i,0).value
                        mp3file = word + '.mp3'
                        if word in myWordsDataFile:
                                times = myWordsDataFile[word][0] + 1
                        else:
                                t = []
                                times = 1
                                t.append(times) #浏览次数
                                t.append(0)     #跟敲次数
                                t.append(0)     #跟敲正确次数
                                t.append(0)     #默写次数
                                t.append(0)     #默写正确次数
                        myWordsData.update({word:t})
                        c += 1
                        text = '\n{}/{}：'.format(c, total)
                        print(text)
                        print(word, '\n[', sheet1.cell(i,1).value, ']', sheet1.cell(i,2).value)
                        text = '\"{}\"已浏览{}次'.format(word, times)
                        print(text)
                        if(os.path.exists(mp3file)):
                                t = threading.Thread(target=play2, args= (mp3file,))
                                t.start()
                        r = check(word, myWordsData)
                        while(r == 0):
                                r = check(word, myWordsData)
                        if(r == 2):
                                break
        return c
if __name__=='__main__':
        cwd = os.getcwd()
        myWordsDataFile = cwd+'\\Datas\\myWordsData'
        os.chdir(cwd)
        if(os.path.exists(myWordsDataFile)):
                with open(myWordsDataFile) as f:
                        myWordsData = json.load(f)
        else:
                myWordsData = {}
        c = 0
        while(1):
                now_time = datetime.datetime.now()
                total = learn(myWordsData)
                end_time = datetime.datetime.now()
                text = '恭喜您完成第{}轮学习，本轮共遍历{}词汇，总用时{}，平均每个词汇用时{}'.format(c, total, end_time-now_time, (end_time-now_time)/total)
                c += 1
                print(text)
                s = input()
                if(s == '##quit'):
                        break
        with open(myWordsDataFile, 'w') as f:
                json.dump(myWordsData, f)
