# Extractor
1.csvToJson
    从tabula得到一篇SR中所有PIO信息的.csv文件，这些文本格式有问题，需要处理一下。转换为JSON格式

2.extractFullText
    读取参考文献原文的PDF文档 存储到xls文件中.一个句子存储一行

3.getRougeScore
    计算实验结果与标准文本间的ROUGE矩阵,根据相似度对原文中句子进行排序，取topN
