# -*- coding: utf-8 -*-
import json

import openpyxl
import os
import shutil
from glob import glob
import os  # os是用来切换路径和创建文件夹的。
import shutil
import time

from pydub import AudioSegment

FromRoot = r"D:\vocal\The Witcher 3\ogg"  #硬盘路径 来源路径
ToRoot = r"wavss"  #硬盘路径   目标路径 别人使用的时候 修改这俩个就好
#此方法用于更新
def importDict():
    xlsxPath = 'Witcher 3 Dialog.xlsx'#这是你的文件
    # 第一步打开工作簿
    wb = openpyxl.load_workbook(xlsxPath)
    # 第二步选取表单
    sheet = wb.active
    # 按行获取数据转换成列表
    rows_data = list(sheet.rows)
    # 获取表单的表头信息(第一行)，也就是列表的第一个元素
    titles = [title.value for title in rows_data.pop(0)]
    print(titles)


    all_row_dict = []

    # 遍历出除了第一行的其他行
    for a_row in rows_data:
        the_row_data = [cell.value for cell in a_row]
        # 将表头和该条数据内容，打包成一个字典
        row_dict = dict(zip(titles, the_row_data))
        #print(row_dict)
        all_row_dict.append(row_dict)
        a=1
    for i in all_row_dict:
        key=i.get('IDs')#表格第一列列名
        #print(key)
        if a>350:
            break
        if key!=None and 'Geralt:' in key:
            try:
                asfds=str(key).split('  ')
                for s in asfds:
                    if '0x' in s:
                        wan=s
                        print('s--->'+s)
                        want=s.replace('0x','0')
                sa=str(asfds[3]).split(': ')
                if len(sa[1]) < 16:
                    continue
                if a<310:
                    igo='wavs/'+want+'.wav|'+sa[1]+'\n'
                    print('wavs/'+want+'.wav|'+sa[1])
                    file = open('list.txt', 'a')
                    file.write(igo)
                    file.close()
                else:

                    igo = 'wavs/' + want + '.wav|' + sa[1] + '\n'
                    print('wavs/' + want + '.wav|' + sa[1])
                    file = open('lista.txt', 'a')
                    file.write(igo)
                    file.close()
                a+=1

                wansa=wan+'.wav.ogg'
                wanss=want+'.wav'
                file_path1 = FromRoot+ '/' + str(wansa)  # 第一个文件的来源
                file_to1 = ToRoot + '/' + str(wanss)
                sound = AudioSegment.from_file(file=file_path1, format='ogg')
                sound.export(out_f=file_to1, format='wav')  # 用于保存剪辑之后的音频文件


                #shutil.copyfile(file_path1, file_to1)


            except:
                print('error line')
            print('--------------')


def check():
    file = open('list.txt', 'r',encoding='utf-8')
    for i in file.readlines():
        ass=i
        ass=ass.split('/')
        ass=ass[1].split('.wav')
        ass=ass[0]+'.wav'
        filenames = os.listdir(r'wavss')
        if ass not in filenames:
            print(ass)
    file.close()

    print('---------------------')
    file = open('lista.txt', 'r', encoding='utf-8')
    for i in file.readlines():
        ass = i
        ass = ass.split('/')
        ass = ass[1].split('.wav')
        ass = ass[0] + '.wav'
        filenames = os.listdir(r'wavss')
        if ass not in filenames:
            print(ass)
    file.close()





if __name__ == '__main__':
    importDict()
    #check()

