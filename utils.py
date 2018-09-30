from typing import Dict, Any

import xlrd
import xlwt
import sys
from xlutils.copy import copy
from keys import NO_CATEGORY
import pickle


def load_category(filename='c:/excel/汇总.xlsm', sheet_index=2):
    """
    读取excel中主题分类的类别
    :param filename: excel文件路径
    :param sheet_index: 所需解析的类别在excel中的sheet序列
    :return:
    """

    # 读取excel
    excel_file = xlrd.open_workbook(filename)
    sheet = excel_file.sheet_by_index(sheet_index)

    # 一级标题列表
    category_list1 = []
    # 二级标题列表
    category_list2 = []
    # 上一行的一级标题ID名，防止重复存储
    old_category = ''
    for i in range(1, sheet.nrows):
        # 一级标题
        if old_category != sheet.cell(i, 1).value:
            category1 = {'id': sheet.cell(i, 1).value,
                         'content': sheet.cell(i, 0).value}
            category_list1.append(category1)
        old_category = sheet.cell(i, 1).value
        # 二级标题
        category2 = {'id': sheet.cell(i, 3).value,
                     'content': sheet.cell(i, 2).value}
        category_list2.append(category2)
    return category_list1, category_list2


def id2title(id):
    """
    通过类目的ID查找类目标题
    :param id: 所需找到的类目id
    :return: 类目标题
    """
    for category in category_list:
        if id == category['id']:
            return category['content']
    return 'None'


def load_id2word(filename='c:/excel/汇总.xlsm'):
    """
    通过ID补全汇总.xlsm中的主题分类title
    :return:
    """

    # 读取excel
    excel_file = xlrd.open_workbook(filename)
    sheet = excel_file.sheet_by_index(0)

    # 加载类目
    category_list1, category_list2 = load_category()
    # 合并两个类目列表
    global category_list
    category_list = category_list2 + category_list1

    # 遍历每行，补全ID对应的标题
    new_categorys = []
    for i in range(0, sheet.nrows):
        zt_id = sheet.cell(i, 2).value
        new_categorys.append(str(zt_id) + id2title(zt_id))

    return new_categorys


def write_category(file='c:/excel/汇总.xlsm', new_categorys=None, sheet_index=0, cols=2):
    """
        向表格中写入ID对应的
        :param file:excel文件
        :param zt_list:新的数据列表
        :param sheet_index:
        :param cols:
        :return:
        """
    ExcelFile = xlrd.open_workbook(file)
    nrows = ExcelFile.sheet_by_index(sheet_index).nrows
    new_excel = copy(ExcelFile)
    sheet = new_excel.get_sheet(sheet_index)
    for i in range(0, nrows):
        try:
            sheet.write(i, cols, new_categorys[i])
        except Exception as e:
            s = sys.exc_info()
            print("Error '%s' happened on line %d" % (s[1], s[2].tb_lineno))
    new_excel.save(file)


def write_result(file, zt_list, sheet_index=0, cols=3):
    """
    向表格中写入新的数据
    :param file:excel文件
    :param zt_list:新的数据列表
    :param sheet_index:
    :param cols:
    :return:
    """
    ExcelFile = xlrd.open_workbook(file)
    nrows = ExcelFile.sheet_by_index(sheet_index).nrows
    new_excel = copy(ExcelFile)
    sheet = new_excel.get_sheet(sheet_index)
    for i in range(0, nrows):
        try:
            sheet.write(i, cols, str(zt_list[i]['id']) + zt_list[i]['title'])
            word_text = ""
            for word in zt_list[i]['words']:
                word_text += word['word'] + ":" + str(word['num']) + ";"
            sheet.write(i, 6, word_text)
        except Exception as e:
            s = sys.exc_info()
            print("Error '%s' happened on line %d" % (s[1], s[2].tb_lineno))
    new_excel.save(file)


def classify_subject(article_list, keys_list):
    """
    通过关键词对文章进行分类
    :param article_list: 文章列表
    :param keys_list: 分类列表
    :return:
    """

    # 存放分类结果（id:类目的二级id；title：类目的名称；keys：类目包含的所有关键词；keys_sort:筛选出靠前的关键词供填入表格中）
    result_list = []
    print(len(article_list))
    # 遍历文章列表
    for article in article_list:
        # 最大匹配数标志
        max_num = 0
        # 当前文章最佳分类的结果
        result = {}
        # 最佳分类的关键词及匹l'l
        max_words = []
        for keys in keys_list:
            num = 0
            words = []
            for key in keys['keys']:
                temp = {}
                temp_num = 0
                temp_num += article['title'].count(key)
                temp_num += article['content'].count(key)
                temp['num'] = temp_num
                temp['word'] = key
                num += temp_num
                words.append(temp)
            if max_num < num:
                max_words = words

                # print('--------------')
                max_num = num
                result = keys.copy()

        new_max_words = sort_keys(max_words)
        if len(new_max_words) is 0:
            print('空')
        result['words'] = new_max_words
        # print(result)
        # print('出现频率：', max_num)

        result_list.append(result)

    return result_list


def filter_no_category(key_arr, category):
    """
    过滤非主题分类和已存在的的关键词
    :param key_arr:
    :return:
    """
    new_key_arr = []
    for key in key_arr:
        try:
            if NO_CATEGORY.count(key) <= 0 and category['std_key'].count(key) <= 0 and key != '':
                new_key_arr.append(key)
        except KeyError:
            if key != '':
                new_key_arr.append(key)
    return new_key_arr


def load_standard_key(file='c:/excel/汇总 - 副本.xlsm', sheet_index=0):
    """
    从汇总.xls中获取官方的key列表
    :param file:
    :return:
    """

    # 获取主题分类的类目列表
    category_list1, category_list2 = load_category(filename='c:/excel - 副本/汇总.xlsm')

    # 加载excel文件句柄
    excel_file = xlrd.open_workbook(file)
    sheet = excel_file.sheet_by_index(sheet_index)

    # 合并两个类目列表
    global category_list
    category_list = category_list2 + category_list1

    # 遍历每一行数据,
    for i in range(0, sheet.nrows):
        try:
            # 识别每行对应的类目
            id = sheet.cell(i, 2).value
            for category in category_list:
                # 将对应id的基准key值存入
                if id == category['id']:
                    # 先对key分词'安全生产;规定;通知'
                    key_str = sheet.cell(i, 4).value
                    key_arr = key_str.split(';')
                    # 去除非主题分类的关键词和已存入的关键词
                    key_arr = filter_no_category(key_arr,category)
                    try:
                        category['std_key'] += key_arr
                    except KeyError:
                        category['std_key'] = key_arr
                    finally:
                        break
        except Exception as e:
            s = sys.exc_info()
            print("Error '%s' happened on line %d" % (s[1], s[2].tb_lineno))

    data_file = 'category_list.pkl'
    f = open(data_file, 'wb')
    pickle.dump(category_list, f)


def sort_keys(words):
    """
    对匹配的关键词详情进行排序，筛选出匹配数大于1的关键词
    :param words: 关键词详情，包含每个关键词的匹配数
    :return:
    """

    # 排序后的关键词详情
    new_words = []

    for word in words:
        new_word = {}
        # 如果该关键词匹配数大于1，则放入匹配结果关键词列表中
        if word['num'] >= 1:
            new_word['num'] = word['num']
            new_word['word'] = word['word']
            new_words.append(new_word)

    return new_words


def load_excel(filename):
    ExcelFile = xlrd.open_workbook(filename)
    sheet = ExcelFile.sheet_by_index(0)
    article_list = []
    for i in range(0, sheet.nrows):
        article = {}
        article['title'] = sheet.cell(i, 0).value
        article['content'] = str(sheet.cell(i, 5).value)
        article_list.append(article)
    return article_list


def load_std_keys():
    """
    加载基准关键词
    :return:
    """
    data_file = 'category_list.pkl'
    f = open(data_file, 'rb')
    std_keys = pickle.load(f)
    return std_keys


if __name__ == '__main__':
    # category_list = load_id2word()
    # write_category(new_categorys=category_list)
    load_standard_key()
