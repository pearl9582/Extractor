# -*- coding: utf-8 -*-
# 计算实验结果与标准文本间的ROUGE矩阵
# 2018.6.5  Pearl

from rouge import Rouge
import rouge
import xlrd
import os,json
from xlutils.copy import copy
import string
from nltk.corpus import stopwords
from nltk.stem.lancaster import LancasterStemmer

# rougeMatrix = ['P-rouge1-f', 'P-rouge1-p', 'P-rouge1-r', 'P-rouge2-f', 'P-rouge2-p', 'P-rouge2-r', 'P-rougeL-f', 'P-rougeL-p', 'P-rougeL-r',
#                'I-rouge1-f', 'I-rouge1-p', 'I-rouge1-r',  'I-rouge2-f', 'I-rouge2-p', 'I-rouge2-r','I-rougeL-f', 'I-rougeL-p', 'I-rougeL-r',
#                'O-rouge1-f', 'O-rouge1-p', 'O-rouge1-r', 'O-rouge2-f', 'O-rouge2-p', 'O-rouge2-r', 'O-rougeL-f', 'O-rougeL-p', 'O-rougeL-r' ]

lancaster_stemmer = LancasterStemmer()#进行词干提取的对象
rougeMatrix = ['P-rouge2-f', 'P-rouge2-r', 'P-rougeL-f', 'P-rougeL-r',
               'I-rouge2-f', 'I-rouge2-r','I-rougeL-f', 'I-rougeL-r',
               'O-rouge2-f', 'O-rouge2-r', 'O-rougeL-f', 'O-rougeL-r']
rougeMatrixIndex = [4,6,7,9,13,15,16,18,22,24,25,27]

textOriginal = []
textStem = [] #提取词干后的文本
textExcludeStopWord = [] #去掉stop words的文本

def rouge_filescores(hyp_path,ref_path):
    '''
    :param hyp_path: 实验结果文本
    :param ref_path: 标准
    :return: 实验结果文本相对于标准文本的分数
    '''
    files_rouge = rouge.FilesRouge(hyp_path, ref_path)
    scores = files_rouge.get_scores()
    return scores



def rouge_filescores(hyp_path,ref_path):
    '''
    :param hyp_path: 实验结果文本
    :param ref_path: 标准
    :return: 实验结果文本相对于标准文本的分数
    '''
    files_rouge = rouge.FilesRouge(hyp_path, ref_path)
    scores = files_rouge.get_scores()
    return scores
def writeRouge(row,col,value,sheet):
    '''
    在原文所在文件的第二个sheet写入相似度数据，每调用一次写入一个句子与一个PIO信息的ROUGE矩阵P{rouge-1[f,p,r]}{rouge-2[f,p,r]}{rouge-l[f,p,r]}
    :param row: 写在第几行
    :param col: 数据开始的列
    :param value: ROUGE矩阵
    :param sheet: 写入的表格
    :return:
    '''
    sheet.write(row, col+0, value[0]['rouge-1']['f'])
    sheet.write(row, col+1, value[0]['rouge-1']['p'])
    sheet.write(row, col+2, value[0]['rouge-1']['r'])
    sheet.write(row, col+3, value[0]['rouge-2']['f'])
    sheet.write(row, col+4, value[0]['rouge-2']['p'])
    sheet.write(row, col+5, value[0]['rouge-2']['r'])
    sheet.write(row, col+6, value[0]['rouge-l']['f'])
    sheet.write(row, col+7, value[0]['rouge-l']['p'])
    sheet.write(row, col+8, value[0]['rouge-l']['r'])

def excludeStopWords(str):
    '''
    返回字符串str去除stop words 的结果
    :param str:
    :return:
    '''
    str = " ".join([word for word in str.translate(str.maketrans('', '', string.punctuation)).split()
                        if word not in stopwords.words('english') ])
    if not str.strip():
        str = str + '*'
    return str

def rougeScoreExcludeStopWords(hyp_str,ref_str):
    '''
    返回两个句子去除stop word 后的
    :param hyp_str: 原文对应的句子
    :param ref_str: SR中的PIO句子
    :return:
    '''
    rouge = Rouge()
    hyp_str = " ".join([word for word in hyp_str.translate(str.maketrans('', '', string.punctuation)).split()
                        if word not in stopwords.words('english') ])
    # print(hyp_str)
    if not hyp_str.strip():
        hyp_str = hyp_str + '*'
    ref_str = " ".join([word for word in ref_str.translate(str.maketrans('', '', string.punctuation)).split()
                        if word not in stopwords.words('english') ])
    # print(ref_str)
    if not ref_str.strip():
        ref_str = ref_str + '*'
    return rouge.get_scores(hyp_str,ref_str)

def rougeSorceWithStemming(hyp_str,ref_str):
    '''
    返回两个句子进行词干提取的相似度
    :param hyp_str:
    :param ref_str:
    :return:
    '''

def calculateRouge(hyp_Path, ref_Path,option):
    '''
    计算PIO信息与原文中每个句子的相似度
    :param hyp_Path:原文所在文件夹
    :param ref_Path:SR中PIO存储的json文件
    :option {'textOriginal','textStem','textExcludeStopWord'}
    :return: 结果存储在原文所在文件的sheet2中
    '''
    with open(ref_Path, 'r') as load_f:
        pio_json = json.load(load_f)
    for pio in pio_json['content']:
        title = pio['Title']
        #PIO信息分别为， pio['Participants']  pio['Interventions'] pio['Outcomes']
        exist = False #标记pio参考文献在原文文件夹中是否存在
        year = 2001
        if len(title.split(' ')) <2 :
            year = 2001
        if len(title.split(' ')) >2:
            year = 2001
        if len(title.split(' ')) == 2:
            year = int(title.split(' ')[1][0:4])
        if year >= 2000:  #去掉2000年前的论文
            path_list = os.listdir(hyp_Path)
            for filename in path_list:  #在文件夹中查找与该参考文献对应的原文标题
                str = filename.split('_')
                if str[0] == pio['Title'] and filename.endswith('.xls'):
                    exist = True
                    break
            if exist == True:  #标记pio参考文献在原文文件夹中存在
                rd = xlrd.open_workbook(hyp_Path+'/'+filename)
                sheet = rd.sheet_by_index(0) #原文所在表格
                nrows = sheet.nrows
                # ncols = sheet.ncols
                wb = copy(rd)
                try:
                    sheet1 = wb.get_sheet(1)
                    # sheet1.write(range(0,nrows+1),range(0,27),'')
                except Exception as err:
                    sheet1 = wb.add_sheet('ROUGE Matrix', cell_overwrite_ok=True)  # 增加一个工作表，记录ROUGE矩阵
                sheet1.write_merge(0, 0, 1, 9, 'P{rouge-1[f,p,r]}{rouge-2[f,p,r]}{rouge-l[f,p,r]}')
                sheet1.write_merge(0, 0, 10, 18, 'I{rouge-1[f,p,r]}{rouge-2[f,p,r]}{rouge-l[f,p,r]}')
                sheet1.write_merge(0, 0, 19, 27, 'O{rouge-1[f,p,r]}{rouge-2[f,p,r]}{rouge-l[f,p,r]}')
                rouge = Rouge()
                for i in range(0,nrows):
                    sheet1.write(i+1, 0, i+1)
                    tempStr = bytes.decode(sheet.cell(i, 0).value.encode('utf-8'))
                    # textOriginal.append(tempStr) #存储原始文本
                    # textExcludeStopWord.append(excludeStopWords(tempStr)) #原始文本去除stop words
                    # textStem.append(lancaster_stemmer.stem(tempStr)) #原始文本进行词干提取

                    textOriginal = tempStr
                    textExcludeStopWord = excludeStopWords(tempStr)
                    textStem = lancaster_stemmer.stem(tempStr)

                    if option == 'textOriginal':
                        #原文本与PIO相似度
                        score_p = rouge.get_scores(textOriginal, pio['Participants'])
                        score_i = rouge.get_scores(textOriginal, pio['Interventions'])
                        score_o = rouge.get_scores(textOriginal, pio['Outcomes'])
                    if option == 'textStem':
                        #提取词干后 文本与PIO相似度
                        score_p = rouge.get_scores(textStem, lancaster_stemmer.stem(pio['Participants']))
                        score_i = rouge.get_scores(textStem, lancaster_stemmer.stem(pio['Interventions']))
                        score_o = rouge.get_scores(textStem, lancaster_stemmer.stem(pio['Outcomes']))
                    if option == 'textExcludeStopWord':
                        # 去除stop words后 文本与PIO相似度
                        score_p = rouge.get_scores(textExcludeStopWord, excludeStopWords(pio['Participants']))
                        score_i = rouge.get_scores(textExcludeStopWord, excludeStopWords(pio['Interventions']))
                        score_o = rouge.get_scores(textExcludeStopWord, excludeStopWords(pio['Outcomes']))

                    writeRouge(i + 1, 1, score_p, sheet1)
                    writeRouge(i + 1, 10, score_i, sheet1)
                    writeRouge(i + 1, 19, score_o, sheet1)
                if not os.path.exists(hyp_Path+'/'+option):
                    os.makedirs(hyp_Path+'/'+option)
                wb.save(hyp_Path+'/'+option+'/'+filename.split('_')[0]+'_'+option+'.xls')
                print(filename + ' ROUGE Matrix has generated')
                # print(score_p,score_i,score_o)

def selectSentence(filePath):
    rd = xlrd.open_workbook(filePath)
    sheet1 = rd.get_sheet(0) #原文所在工作表
    sheet2 = rd.get_sheet(1) #ROUGE矩阵所在工作表
    nrows = sheet1.nrows #共有多少个句子
    cols = sheet2.col_values(1)
    maxcols = max(cols)
    print(maxcols)

def rougeRankBy(Path,topN,by):
    '''
    按照by的ROUGE进行排序，取topN个句子
    :param Path:
    :param topN:
    :param by:字符串{'P-rouge-1-p','P-rouge-1-r','P-rouge-1-f'}，以此类推，共3*3*3种
    :return:
    '''
    # Path = 'F:/TestPaper/reference'
    path_list = os.listdir(Path)
    for filePath in path_list:
        if filePath.endswith('.xls'):
            filePath= Path+'/'+filePath
            rd = xlrd.open_workbook(filePath)
            sheet1 = rd.sheet_by_index(0) #原文所在工作表
            sheet2 = rd.sheet_by_index(1) #ROUGE矩阵所在工作表
            nrows = sheet1.nrows #共有多少个句子
            cols = sheet2.col_values(rougeMatrixIndex[by]) #获取一列数据###################
            cols.pop(0)#去掉表头
            if not cols:
                os.remove(filePath)
                return
            wb = copy(rd)
            try:
                sheet3 = wb.get_sheet(by+2)
            except Exception as err:
                sheet3 = wb.add_sheet('原文相似度排序by'+rougeMatrix[by], cell_overwrite_ok=True)  # 增加一个工作表，记录ROUGE矩阵

            for i in range(0,topN): #取相似度最高的top n
                maxcols = max(cols)
                index = cols.index(maxcols)
                # print(index)
                # print(sheet1.cell(index,0).value)
                sheet3.write(i,0,sheet1.cell(index,0).value)
                cols[index] = 0
            sheet3.write(i+1,0,'###### top '+str(topN)+' end')
            sheet3.write(i+2,0,'')
            wb.save(filePath)
            print(filePath+'原文相似度排序top'+str(topN)+'by'+rougeMatrix[by] + ' has generated')


def rougeRank(Path,topN):
    '''
    按ROUGE的某一项进行排序，得到topN个句子
    :param Path:
    :param topN:
    :return:
    '''
    for r in rougeMatrix:
        rougeRankBy(Path,topN,rougeMatrix.index(r))


if __name__ == '__main__':

    hyp_Path = 'C:/Users/Pearl/Desktop/hox1/ref_ziehm1996'  # 原文所在文件夹
    ref_Path = 'C:/Users/Pearl/Desktop/hox1/PIO/tabula-84ziehm1996.csv.json'  # SR中提取出的PIO信息 所在文件夹
    #####三种计算相似度的方式
    calculateRouge(hyp_Path, ref_Path,'textStem') #计算相似度矩阵
    rougeRank(hyp_Path+'/textStem',15) #按相似度对句子进行排序，取topN的句子
    calculateRouge(hyp_Path, ref_Path, 'textOriginal')  # 计算相似度矩阵
    rougeRank(hyp_Path + '/textOriginal', 15)  # 按相似度对句子进行排序，取topN的句子
    calculateRouge(hyp_Path, ref_Path, 'textExcludeStopWord')  # 计算相似度矩阵
    rougeRank(hyp_Path + '/textExcludeStopWord', 15)  # 按相似度对句子进行排序，取topN的句子