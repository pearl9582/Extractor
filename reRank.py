# -*- coding: utf-8 -*-
# 读对topN 个P进行重排序
# 2018.7.1 Pearl

import json,re,os
import xlrd
from xlutils.copy import copy


def getNumOfP(res_Path, pio_Path):
    '''

    :param res_Path:  待计算的topN表格路径
    :param pio_Path:  json文件路径
    :return:返回这一篇文献的标准P中含有的数字
    '''
    p_num = []
    with open(pio_Path, 'r') as load_f:
        pio_json = json.load(load_f)
    for pio in pio_json['content']:
        if res_Path.split('_')[0] == pio['Title']:
            participants = pio['Participants'].split(' ')
            for p in participants:  #得到P中所有的数字
                t = re.sub("\D", "", p)
                if t:
                    p_num.append(t)
                # print(p_num)
    return p_num

def getNumOfRef(filename,sheetNum):
    '''

    :param filename:  topN所在文件
    :return:  返回topN这些句子中的num
    '''
    resPNum = []
    pNum = []
    rd = xlrd.open_workbook(filename)
    sheet = rd.sheet_by_index(sheetNum)  # top N所在表格
    nrows = sheet.nrows
    # wb = copy(rd)
    for i in range(0,nrows-1):
        participants = bytes.decode(sheet.cell(i, 0).value.encode('utf-8'))
        participants = re.sub("[[](.*?)[]]", "", participants)
        participants = participants.split(' ')
        for p in participants:
            t = re.sub("\D",'',p)
            if t:
                pNum.append(t)
        resPNum.append(pNum)
        pNum = []
    return resPNum

def numSimilarity(resPNum, p_Num):
    '''

    :param resPNum:
    :param p_Num:
    :return: 返回每个句子与标准P的相同数字个数
    '''
    result = []
    number = 0
    for rp in resPNum:
        for n in rp:
            for pn in p_Num:
                if n == pn:
                    number = number +1
        result.append(number)
        number = 0
    return result

def reRankP(res_content,res_Path, pio_Path,sheetNum ,topN):
    '''
    对文件res_Path
    :param res_Path:
    :param pio_Path:
    :param sheetNum:
    :return:
    '''
    filename = res_content + '/' + res_Path
    rd = xlrd.open_workbook(filename)
    sheet = rd.sheet_by_index(sheetNum)  # top N所在表格
    nrows = sheet.nrows
    wb = copy(rd)
    ws = wb.get_sheet(sheetNum)
    pNum = getNumOfP(res_Path, pio_Path)
    # print(pNum)

    resPNum = getNumOfRef(filename, sheetNum)
    # print(resPNum)
    ns = numSimilarity(resPNum, pNum)
    # print(ns)
    for i in range(0, topN):
        maxn = max(ns)
        index = ns.index(maxn)
        ws.write(i + nrows + 5, 0, sheet.cell(index, 0).value)
        ns[index] = -1
    wb.save(filename)





if __name__ == '__main__':
    res_Path = 'Jafari 2012_textExcludeStopWord.xls'
    res_content = 'F:/TestPaper/ref_Music/textExcludeStopWord'
    pio_Path = 'F:/TestPaper/PIO/Tabula-Music for stress and anxiety reduction in coronary heartdisease patients.csv.json'
    for i in range(2,6):
        reRankP(res_content,res_Path, pio_Path,i ,8)