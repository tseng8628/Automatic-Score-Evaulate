#!/usr/bin/python
# -*- coding: utf-8 -*-

import pandas as pd
from itertools import zip_longest
import openpyxl

class scoreEvaulate():

    def __init__(self, classroom, path, input_excel, output_excel, basic_score, last_input_row, write_row):
        self.classroom = classroom
        self.path = path
        self.input_excel = input_excel
        self.output_excel = output_excel
        self.basic_score = basic_score
        self.last_input_row = last_input_row
        self.write_row = write_row

    def readFile(self):
        df = pd.read_excel(self.path + self.classroom + '/' + self.input_excel)
        name = df['請輸入您的姓名'].tolist()
        number = df['請選擇您的座號'].tolist()

        return name, number, df

    def writeScore(self, saveScore):

        df1 = pd.read_excel(self.path + self.output_excel, self.classroom)
        number1 = list(map(int, df1['座號'].tolist()[0:30]))
        name1 = df1['姓名'].tolist()[0:30]
        hw_score = df1[df1.columns[self.write_row]].tolist()[0:30]

        for i in range(len(saveScore)):
            if saveScore[i][1] in name1:
                index = name1.index(saveScore[i][1])
                hw_score[index] = saveScore[i][2]
            else:
                pass

        wb = openpyxl.load_workbook(self.path + self.output_excel)
        ws = wb[self.classroom]

        for i in range(1, len(number1) + 1):
            ws.cell(i + 1, self.write_row).value = hw_score[i - 1]

        wb.save(self.path + self.output_excel)


    def main(self):
        print('檔案讀取中...')
        name, number, df = self.readFile()

        print('成績計算中...')
        score = []
        score_dic = {'順利完成': 2, '仍有小臭蟲(bug)': 1, '尚未撰寫': 0.5}

        for i in range(len(df)):
            temp = 0
            aaa = df.iloc[i].tolist()[5:self.last_input_row]
            for x in aaa:
                temp += score_dic[x]
            score.append(round(temp) + self.basic_score)

        saveScore = list(zip_longest(number, name, score))

        print('成績計算完畢！')

        print('成績寫入中...')
        self.writeScore(saveScore)

        print('成績寫入完畢！')

if __name__ == '__main__':

    # input參數
    # 注意還有人數要先更改
    classroom = '803'
    path = '/Volumes/GoogleDrive/我的雲端硬碟/02 110-2/04 各班成績/'
    input_excel = '【803回饋表】課堂練習2-打磚塊(Scratch)  (回覆).xlsx'
    output_excel = '110-2各班成績.xlsx'

    # hw1 = {'basic_score': 65, 'last_input_row': 20, 'write_row': 3}
    # hw2 = {'basic_score': 65, 'last_input_row': 19, 'write_row': 4}
    # hw3 = {'basic_score': 70, 'last_input_row': 16, 'write_row': 5}
    # hw4 = {'basic_score': 70, 'last_input_row': 15, 'write_row': 7}

    basic_score = 65
    last_input_row = 19
    write_row = 4

    a = scoreEvaulate(classroom, path, input_excel, output_excel, basic_score, last_input_row, write_row)
    a.main()