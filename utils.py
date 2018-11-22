import xlrd
import sys
from xlutils.copy import copy
from keys import NO_CATEGORY
import pickle
import jieba
import jieba.posseg
import numpy as np

CATEGORY_TYPE_zt = 1
CATEGORY_TYPE_tc = 2
CATEGORY_TYPE_gw = 3


def load_category(filename='c:/excel/汇总.xls', sheet_index=2):
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
        if sheet.cell(i, 0).value != '':
            if old_category != sheet.cell(i, 1).value:
                category1 = {'id': sheet.cell(i, 1).value,
                             'content': sheet.cell(i, 0).value}
                category_list1.append(category1)
        old_category = sheet.cell(i, 1).value
        # 二级标题
        if sheet.cell(i, 2).value != '':
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


def load_id2word(filename='c:/excel/汇总.xls', sheet_index=0, colx=2, category_sheet_index=2):
    """
    通过ID补全汇总.xlsm中的主题分类title
    :return:
    """

    # 读取excel
    excel_file = xlrd.open_workbook(filename)
    sheet = excel_file.sheet_by_index(sheet_index)

    # 加载类目
    category_list1, category_list2 = load_category(filename, category_sheet_index)
    # 合并两个类目列表
    global category_list
    category_list = category_list2 + category_list1

    # 遍历每行，补全ID对应的标题
    new_categorys = []
    for i in range(0, sheet.nrows):
        zt_id = sheet.cell(i, colx=colx).value
        new_categorys.append(str(zt_id) + id2title(zt_id))

    return new_categorys


def write_category(file='c:/excel/汇总.xls', new_categorys=None, sheet_index=0, cols=2):
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


def classify_subject(article_list, keys_list, index='keys'):
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
                for key in keys['std_key']:
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


def load_standard_key(file='c:/excel - 副本/汇总 - 副本.xlsm', sheet_index=0, category_type=1, is_save=True, key_col=4):
    """

    :param file: excel文件路径
    :param sheet_index: 存放文章数据的sheet索引
    :param category_type: 分类的类别（主题分类、体裁分类、公文种类）
    :param is_save: 是否将解析出的关键词保存至pkl文件中
    :param key_col: 关键词的所在列索引
    :return:
    """

    if category_type is CATEGORY_TYPE_zt:
        # 获取主题分类的类目列表
        category_list1, category_list2 = load_category(filename='c:/excel - 副本/汇总.xlsm')
        # 合并两个类目列表
        category_list_ = category_list2 + category_list1
        id_col = 2
    elif category_type is CATEGORY_TYPE_tc:
        category_list_ = parse_style()
        id_col = 1
    else:
        return

    # 加载excel文件句柄
    excel_file = xlrd.open_workbook(file)
    sheet = excel_file.sheet_by_index(sheet_index)

    # 遍历每一行数据,
    for i in range(0, sheet.nrows):
        try:
            # 识别每行对应的类目
            id_ = sheet.cell(i, id_col).value
            for category in category_list_:
                # 将对应id的基准key值存入
                if id_ == category['id']:
                    # 先对key分词'安全生产;规定;通知'
                    key_str = sheet.cell(i, key_col).value
                    if category_type is CATEGORY_TYPE_tc:
                        key_arr = [key_str.split(';')[-1]]
                    elif category_type is CATEGORY_TYPE_zt:
                        key_arr = key_str.split(';')
                        # 去除非主题分类的关键词和已存入的关键词
                        key_arr = filter_no_category(key_arr, category)
                    else:
                        return
                    try:
                        category['std_key'] += key_arr
                    except KeyError:
                        category['std_key'] = key_arr
                    finally:
                        break
        except Exception as e:
            print(e)
            s = sys.exc_info()
            print("Error '%s' happened on line %d" % (s[1], s[2].tb_lineno))

    if category_type is CATEGORY_TYPE_tc:
        category_list_ = get_larger_keys(category_list_, threshold=0.0)
        pass

    if is_save:
        data_file = 'category_list.pkl'
        f = open(data_file, 'wb')
        pickle.dump(category_list_, f)
    else:
        return category_list_


def get_keys_num(keys_list):
    """
    计算keys的出现次数
    :param keys_list:
    :return:
    """
    keys_num = {}
    for key in keys_list:
        if keys_num.get(key) is None:
            keys_num[key] = 1
        else:
            keys_num[key] += 1

    return keys_num


def filter_larger_key(keys_num, threshold):
    """
    过滤出关键词中占比率超过阈值的部分
    :param keys_num:
    :param threshold:
    :return:
    """
    new_keys = []
    sum_num = np.array([x for x in keys_num.values()]).sum()
    for key in keys_num:
        if keys_num[key] / sum_num > threshold:
            new_keys.append(key)

    return new_keys


def get_larger_keys(category_list_, threshold=0.2):
    """
    过滤出关键词中占比率超过阈值的部分
    :param category_list_:
    :param threshold:
    :return:
    """
    for category in category_list_:
        try:
            keys_num = get_keys_num(category['std_key'])
            category['std_key'] = filter_larger_key(keys_num, threshold)
        except KeyError:
            print(str(category['id']) + ' ' + category['content'] + '没有标准关键词')

    return category_list_


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


def filter_no_key():
    """
    删除无意义的key
    :return:
    """
    std_keys = load_std_keys()
    for key in std_keys:
        print(key)


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
        try:
            if is_joint_3(words[i], words[i + 1], words[i + 2]):
                new_words.append(words[i].word + words[i + 1].word)
                is_skip = True
            else:
                new_words.append(words[i].word)
        except IndexError:
            if is_joint_2(words[i], words[i + 1]):
                new_words.append(words[i].word + words[i + 1].word)
                is_skip = True
            else:
                new_words.append(words[i].word)
    return new_words


def is_joint_3(word1, word2, word3):
    """
    在旧函数的基础上优化了判断规则，并且设置组词的优先级，当当前组词优先级低于下一个组词的优先级则跳过
    :param word1:
    :param word2:
    :param word3:
    :return:
    """
    try:
        word1.flag.index('a')
        word2.flag.index('v')
        if get_priority(word1, word2) >= get_priority(word2, word3):
            return True
    except ValueError:
        pass

    try:
        word1.flag.index('a')
        word2.flag.index('n')
        if get_priority(word1, word2) >= get_priority(word2, word3):
            return True
    except ValueError:
        pass

    try:
        word1.flag.index('v')
        word2.flag.index('n')
        if get_priority(word1, word2) >= get_priority(word2, word3):
            return True
    except ValueError:
        pass

    try:
        word1.flag.index('n')
        word2.flag.index('v')
        if get_priority(word1, word2) >= get_priority(word2, word3):
            return True
    except ValueError:
        pass

    try:
        word1.flag.index('n')
        word2.flag.index('n')
        try:
            # 且不 名词+形容词
            word1.flag.index('n')
            word2.flag.index('a')
            return False
        except ValueError:
            pass

        # 且不 地名+非地名
        if word1.flag == 'ns' and word2.flag == 'ns':
            if get_priority(word1, word2) >= get_priority(word2, word3):
                return True
        if word1.flag != 'ns' and word2.flag != 'ns':
            if get_priority(word1, word2) >= get_priority(word2, word3):
                return True

    except ValueError:
        pass

    return False


def is_joint_2(word1, word2):
    """
    在旧函数的基础上优化了判断规则
    :param word1:
    :param word2:
    :param word3:
    :return:
    """
    try:
        word1.flag.index('a')
        word2.flag.index('v')
        return True
    except ValueError:
        pass

    try:
        word1.flag.index('a')
        word2.flag.index('n')
        return True
    except ValueError:
        pass

    try:
        word1.flag.index('v')
        word2.flag.index('n')
        return True
    except ValueError:
        pass

    try:
        word1.flag.index('n')
        word2.flag.index('v')
        return True
    except ValueError:
        pass

    try:
        word1.flag.index('n')
        word2.flag.index('n')
        try:
            # 且不 名词+形容词
            word1.flag.index('n')
            word2.flag.index('a')
            return False
        except ValueError:
            pass

        # 且不 地名+非地名
        if word1.flag == 'ns' and word2.flag == 'ns':
            return True
        if word1.flag != 'ns' and word2.flag != 'ns':
            return True

    except ValueError:
        pass

    return False


def get_priority(word1, word2):
    """
    获取组词的优先级
    :param word1:
    :param word2:
    :return:
    """
    try:
        word1.flag.index('a')
        word2.flag.index('v')
        return 5
    except ValueError:
        pass

    try:
        word1.flag.index('a')
        word2.flag.index('n')
        return 4
    except ValueError:
        pass

    try:
        word1.flag.index('v')
        word2.flag.index('n')
        return 3
    except ValueError:
        pass

    try:
        word1.flag.index('n')
        word2.flag.index('v')
        return 2
    except ValueError:
        pass

    try:
        word1.flag.index('n')
        word2.flag.index('n')
        if word1 == 'n' and word2 == 'n':
            return 0
        return 1
    except ValueError:
        pass

    return 0


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
        word1.flag.index('v')
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
                # 关键词不能以内蒙古自治区开头
                try:
                    word.index('内蒙古自治区')
                except:
                    word_text += word + ';'
            sheet.write(i, 6, word_text)
            i += 1
        except Exception as e:
            s = sys.exc_info()
            print("Error '%s' happened on line %d" % (s[1], s[2].tb_lineno))
    new_excel.save(file)


def parse_style(filename='I:/Tencent Files/1700117425/FileRecv/数据.xls'):
    """
    解析文件中的体裁分类类目
    :param filename:
    :return:
    """
    (style_list1, style_list2) = load_category(filename, sheet_index=2)
    style_list = style_list1 + style_list2
    del style_list1, style_list2
    return style_list


if __name__ == '__main__':
    # 通过id补全主题分类
    # category_list = load_id2word()
    # write_category(new_categorys=category_list)

    # 通过id补全体裁分类

    # category_list = load_id2word(sheet_index=1, colx=1, category_sheet_index=1)
    # write_category(new_categorys=category_list, sheet_index=0, cols=1)

    # load_standard_key()
    # filter_no_key()
    # print(parse_style())
    print(load_standard_key(file='I:/Tencent Files/1700117425/FileRecv/数据样例.xlsx', category_type=CATEGORY_TYPE_tc, is_save=False,key_col=7))
