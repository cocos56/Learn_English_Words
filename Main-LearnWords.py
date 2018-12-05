from LearnWords import LW
import json
import os
import datetime

if __name__=='__main__':
        cwd = os.getcwd()
        myWordsDataFileDir = cwd+'\\Datas\\myWordsDat.json'
        workbookDir = r'D:\0COCO\本科\大二上学期\英语三\Datas\ModifyXlsx\B3Unit5B.xlsx'
        if(os.path.exists(myWordsDataFileDir)):
                with open(myWordsDataFileDir) as f:
                        myWordsData = json.load(f)
        else:
                myWordsData = {}
        c = 0
        lw = LW(cwd, myWordsData, workbookDir)
        while(1):
                now_time = datetime.datetime.now()
                total = lw.learnWithNumber()
                end_time = datetime.datetime.now()
                text = '恭喜您完成第{}轮学习，本轮共遍历{}词汇，总用时{}，平均每个词汇用时{}'.format(c, total, end_time-now_time, (end_time-now_time)/total)
                c += 1
                print(text)
                s = input()
                if(s == '##quit'):
                        break
        del lw
        with open(myWordsDataFileDir, 'w') as f:
                json.dump(myWordsData, f)
