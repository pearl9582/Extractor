# -*- coding: utf-8 -*-
#  构建训练集要执行的步骤
# 2018.6.24 Pearl

from rightCode import csvToJson,extractFullText,getRougeScore
import os
from tkinter import *
from tkinter import messagebox
from tkinter.filedialog import askdirectory,askopenfilename

def step(filePath,reference_Path):
    #输入从tabula 网站得到的.csv文件，转换成JSON
    #输出文本保存在filePath下
    # filePath = 'F:/TestPaper/PIO/Tabula-Inhaled nitric oxide for the postoperative management of pulmonary hypertension in infants and children with congenital heart disease.csv'   #地址可修改
    jsonPath = csvToJson.csvToJson(filePath)

    #输入参考文献所在文件夹，一次读取其中的pdf文件，转化成xls文件
    #输出文本保存在path下
    # reference_Path = "F:/TestPaper/ref_Inhaled"  # 待读取的文件夹                                                                      #地址可修改
    path_list = os.listdir(reference_Path)
    for filename in path_list:
        if filename.endswith('.xls'):
            filePath = os.path.join(reference_Path, filename)
            os.remove(filePath)
            print('Delete '+filename)
    for filename in path_list:
        if filename.endswith('.pdf'):
            filePath = os.path.join(reference_Path, filename)
            extractFullText.readPDFtoCSV(filePath)

    #计算P/I/O与参考文献中句子的ROUGE矩阵，并对相似度进行排序
    #输出结果在hyp_Path下创建相应文件夹{textOriginal , textStem , textExcludeStopWord}分别是：原文，原文进行词干提取，原文去除stop words
    hyp_Path = reference_Path
    ref_Path = jsonPath
    #####三种计算相似度的方式

    getRougeScore.calculateRouge(hyp_Path, ref_Path, 'textOriginal')  # 计算相似度矩阵
    getRougeScore.rougeRank(hyp_Path + '/textOriginal', 15)  # 按相似度对句子进行排序，取topN的句子

    getRougeScore.calculateRouge(hyp_Path, ref_Path,'textStem') #计算相似度矩阵
    getRougeScore.rougeRank(hyp_Path+'/textStem',15) #按相似度对句子进行排序，取topN的句子

    getRougeScore.calculateRouge(hyp_Path, ref_Path, 'textExcludeStopWord')  # 计算相似度矩阵
    getRougeScore.rougeRank(hyp_Path + '/textExcludeStopWord', 15)  # 按相似度对句子进行排序，取topN的句子


def selectPath():
    path_ = askdirectory()
    path2.set(path_)

def selectFile():
    path_ = askopenfilename()
    path1.set(path_)

def excuteStep():
    if path1.get() and path2.get():
        step(path1.get(),path2.get())
        r = messagebox.showwarning('消息框', '数据处理完成！')
        print('showwarning:', r)
    else:
        r = messagebox.showwarning('消息框', '路径存在问题')
        print('showwarning:', r)



if __name__ == '__main__':
    root = Tk()
    path1 = StringVar()
    path2 = StringVar()

    Label(root,text = "PIO文件地址:").grid(row = 0, column = 0)
    Entry(root, textvariable = path1).grid(row = 0, column = 1)
    Button(root, text = "选择", command = selectFile).grid(row = 0, column = 4)


    Label(root,text = "参考文献目录:").grid(row = 1, column = 0)
    Entry(root, textvariable = path2).grid(row = 1, column = 1)
    Button(root, text = "选择", command = selectPath).grid(row = 1, column = 4)


    Button(root, text = "计算", command = excuteStep).grid(row = 2, column = 2)


    root.mainloop()


