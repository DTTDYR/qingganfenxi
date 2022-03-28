import xlrd
import json
from snownlp import SnowNLP
from snownlp import sentiment

def nlp_text(text='提醒大家注意无症状感染者，是好事呀。 原文在这。许多“较真的网友”不关心患者的命运，而是拿着显微镜找披露事实者叙述的毛病。'):
    resp = SnowNLP(text)  # 不只是情感分析，训练出的snownlp模型还有提取关键字，摘要等功能
    # print(resp.sentiments)
    # print(resp.keywords(3))
    # print(resp.summary(3))
    return resp.sentiments
    

FILE_INFO = '数据3.xlsx'


def parse_xlsx():
    npl_result = []
    xl = xlrd.open_workbook(FILE_INFO)
    
    sheet_ref = xl.sheet_by_index(0)
    for r_num in range(1, sheet_ref.nrows):
        # 微博名
        name = sheet_ref.cell_value(r_num, 0)
        # # 博文内容
        content = sheet_ref.cell_value(r_num, 1)
        # #地址
        address = sheet_ref.cell_value(r_num, 2)

        # result.append((name,content,address))
        # 情感分析
        npl_result.append([
            name, 
            nlp_text(content), 
            address
        ])
        print('npl_result')
        print(npl_result)

    resp = {
        "npl_result": npl_result
    }
    return resp


if __name__ == '__main__':
    parse_xlsx()
