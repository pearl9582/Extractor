# -*- coding: utf-8 -*-
# 读取PDF文档 存储到TXT或者xls文件中
# 2018.6.5 Pearl
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfpage import PDFTextExtractionNotAllowed
from pdfminer.pdfinterp import PDFResourceManager
from pdfminer.pdfinterp import PDFPageInterpreter
from pdfminer.layout import *
from pdfminer.converter import PDFPageAggregator
from textblob import TextBlob

import xlwt
import os

def readPDFtoCSV(filePath):
    '''
    :param filePath:  需要读取的pdf路径
    :return:
    '''
    #excel写入
    wb = xlwt.Workbook(encoding = 'ascii')
    ws = wb.add_sheet('原文')
    #创建一个pdf文档分析器
    fileNames = os.path.splitext(filePath)
    fp = open(filePath, 'rb')
    try:
        parser = PDFParser(fp)
        #创建一个PDF文档对象存储文档结构
        document = PDFDocument(parser)
        # 检查文件是否允许文本提取

        if not document.is_extractable:
            raise PDFTextExtractionNotAllowed
        else:
            # 创建一个PDF资源管理器对象来存储共赏资源
            rsrcmgr=PDFResourceManager()
            # 设定参数进行分析
            laparams=LAParams()
            # 创建一个PDF设备对象
            device=PDFPageAggregator(rsrcmgr,laparams=laparams)
            # 创建一个PDF解释器对象
            interpreter=PDFPageInterpreter(rsrcmgr,device)
            # 处理每一页
            i=0 #标记Excel的行
            endPDF = False
            for page in PDFPage.create_pages(document):
                interpreter.process_page(page)
                layout=device.get_result()
                for x in layout:
                    if hasattr(x, "get_text"):
                        results = x.get_text()
                        if results.lower().startswith('references\n') or results.lower().startswith('references \n') :
                            endPDF = True
                            break
                        blob = TextBlob(results)
                        res = blob.sentences
                        # print(type(res[0]))
                        # results = results.replace('\n',' ') #换行问题
                        # results = results.replace('- ','') #去掉换行后的‘- ’
                        # res = results.split('. ')
                        pre = ''
                        for r in res:
                            r = str(r)
                            r = pre + ' '+ r #如果pre不为空，说明上一句以 ‘i,e,’结尾，两句话合并
                            if r.strip().endswith('i.e.') or r.strip().endswith('Fig.'):
                                pre = r.strip() #该句以 ‘i.e.’结尾，设置下一句的开头是本句，本句设为0
                                r = ''
                            else:
                                pre = ''
                            if r.strip() and len(r)>4:
                                wstr = r.replace('\n',' ').replace('- ','').replace('(cid:129)',' ')
                                ws.write(i,0,wstr)
                                i = i + 1

                    if endPDF == True:
                        break
                if endPDF == True:
                    break
                wb.save(fileNames[0] + '.xls')#生成TXT文件
                # if not os.path.getsize(fileNames[0] + '.xls'):
                #     os.remove(fileNames[0] + '.xls')
    except :
        print('some error with '+fileNames[0])
        return
    print(fileNames[0]+'.xls has generated'+'\n')




if __name__ == '__main__':
    path = "F:/TestPaper/test"  # 待读取的文件夹
    path_list = os.listdir(path)
    for filename in path_list:
        if filename.endswith('.xls'):
            filePath = os.path.join(path, filename)
            os.remove(filePath)
            print('Delete '+filename)
    for filename in path_list:
        if filename.endswith('.pdf'):
            filePath = os.path.join(path, filename)
            readPDFtoCSV(filePath)