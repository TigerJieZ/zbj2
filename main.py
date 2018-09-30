import xlrd
from keys import keys_list_new
import xlwt
import sys
from xlutils.copy import copy
import os
import utils
import jieba
import jieba.posseg
import numpy as np


def check(s1, s2):
    return sum(map(lambda ch: s1.count(ch), s2))


def load_key(filename):
    ExcelFile = xlrd.open_workbook(filename)
    sheet = ExcelFile.sheet_by_index(1)
    keys_list = []
    for i in range(1, sheet.nrows - 2):
        keys = {}
        title_cell = sheet.cell(i, 2)
        id_cols = sheet.cell(i, 3)
        keys['id'] = int(id_cols.value)
        key_arr = title_cell.value.split('、')
        content = []
        for key in key_arr:
            # print(key, ' ', end="")
            content.append(key)
        # keys['content']=content
        keys['content'] = title_cell.value
        keys_list.append(keys)
    print(keys_list)
    return keys_list


def load_excel(filename):
    ExcelFile = xlrd.open_workbook(filename)
    sheet = ExcelFile.sheet_by_index(0)
    article_list = []
    for i in range(0, sheet.nrows):
        article = {}
        article['title'] = sheet.cell(i, 0).value
        article['content'] = str(sheet.cell(i, 7).value)
        article_list.append(article)
    return article_list


def in_words(key, words):
    """
    判断key是否包含在words中的某个word中
    :param key:
    :param words:
    :return:
    """
    for word in words:
        try:
            key.index(word['word'])
            return True
        except ValueError:
            pass
    return False


def is_joint(word1, word2):
    """
    判断两个相邻word是否可lianjie
    :param word1:
    :param word2:
    :return:
    """

    try:
        word1.flag.index('n')
        word2.flag.index('v')
        return True
    except ValueError:
        pass

    try:
        word1.flag.index('a')
        word2.flag.index('v')
        return True
    except ValueError:
        pass

    try:
        word1.flag.index('n')
        word2.flag.index('n')
        return True
    except ValueError:
        pass

    return False


def rebuild(words):
    """
    重构分词列表
    规则：名词+动词/形容词+动词/名词+名词
    :param words:
    :return:
    """
    new_words = []
    is_skip = False
    for i in range(len(words) - 1):
        if is_skip:
            is_skip = False
            continue
        if is_joint(words[i], words[i + 1]):
            new_words.append(words[i].word + words[i + 1].word)
            is_skip = True
        else:
            new_words.append(words[i].word)
    return new_words


def filter_keys(new_max_words, title):
    """
    根据题目筛选出合适的关键词
    :param new_max_words:关键词列表
    :param title:标题名
    :return:
    """

    # result
    result = []

    # 对题目进行分词
    words = list(jieba.posseg.cut(title))

    # 对题目的分词结果重构，将符合连接规则的相邻词连接成新的词
    words = rebuild(words)

    # 筛选关键词，判断每个词是否在关键词列表中或者包含在某个关键词中，若是则判断该词为关键词
    for word in words:
        if new_max_words.count(word) > 0 or in_words(word, new_max_words):
            result.append(word)

    return result


def classify(article_list, keys_list, index='keys'):
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
            try:
                for key in keys[index]:
                    temp = {}
                    temp_num = 0
                    temp_num += article['title'].count(key)
                    # temp_num += article['content'].count(key)
                    temp['num'] = temp_num
                    temp['word'] = key
                    num += temp_num
                    words.append(temp)
                if max_num < num:
                    max_words = words

                    # print('--------------')
                    max_num = num
                    result = keys.copy()
            except KeyError:
                # print('no std key')
                pass

        new_max_words = sort_keys(max_words)
        # 从关键词中筛选出合适的
        new_max_words = filter_keys(new_max_words, article['title'])
        if len(new_max_words) is 0:
            print('空')
        result['words'] = new_max_words
        # print(result)
        # print('出现频率：', max_num)

        result_list.append(result)

    return result_list


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
        # 如果该关键词匹配数大于0，则放入匹配结果关键词列表中
        if word['num'] >= 1:
            new_word['num'] = word['num']
            new_word['word'] = word['word']
            new_words.append(new_word)

    return new_words


def write_result(file, zt_list):
    ExcelFile = xlrd.open_workbook(file)
    nrows = ExcelFile.sheet_by_index(0).nrows
    new_excel = copy(ExcelFile)
    sheet = new_excel.get_sheet(0)
    for i in range(0, nrows):
        try:
            sheet.write(i, 3, str(zt_list[i]['id']) + zt_list[i]['content'])
            word_text = ""
            for word in zt_list[i]['words']:
                word_text += word + ';'
            sheet.write(i, 6, word_text)
            i += 1
        except Exception as e:
            s = sys.exc_info()
            print("Error '%s' happened on line %d" % (s[1], s[2].tb_lineno))
    new_excel.save(file)


if __name__ == '__main__':
    EXTERSIONS = ['xls', 'xlsm', 'xlsx']
    std_keys = utils.load_std_keys()
    for path in os.listdir('C:/excel/'):
        if EXTERSIONS.count(path.split('.')[-1]) > 0:
            article_list = load_excel('C:/excel/' + path)
            # keys_list_ = load_key('C:/excel/数据.xls')
            result_list = classify(article_list, std_keys, index='std_key')
            write_result('C:/excel/' + path, result_list)

    # article_list = load_excel('C:/excel/信息公开目录.xls')
    # # keys_list_ = load_key('C:/excel/数据.xls')
    # result_list = classify(article_list, keys_list_new)
    # write_result('C:/excel/信息公开目录.xls', result_list)
