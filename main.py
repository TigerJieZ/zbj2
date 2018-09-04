import xlrd
from keys import keys_list

def check(s1, s2):
    return sum(map(lambda ch: s1.count(ch), s2))


def load_key(filename):
    ExcelFile = xlrd.open_workbook(filename)
    sheet = ExcelFile.sheet_by_index(1)
    keys_list = []
    for i in range(1, sheet.nrows-2):
        keys = {}
        title_cell = sheet.cell(i, 2)
        id_cols = sheet.cell(i, 3)
        keys['id'] = int(id_cols.value)
        key_arr = title_cell.value.split('、')
        content=[]
        for key in key_arr:
            # print(key, ' ', end="")
            content.append(key)
        keys['content']=content
        keys_list.append(keys)
    print(keys_list[:-2])
    return keys_list[:-2]


def load_excel(filename):
    ExcelFile = xlrd.open_workbook(filename)
    sheet = ExcelFile.sheet_by_index(0)
    article_list = []
    for i in range(0, sheet.nrows):
        article = {}
        article['title'] = sheet.cell(i, 0).value
        article['content'] = sheet.cell(i, 7).value
        article_list.append(article)
    return article_list


def classify(article_list, keys_list):
    for article in article_list:
        max_num = 0
        result = []
        print(article['title'])
        for keys in keys_list:
            num = 0
            for key in keys['content']:
                num += article['title'].count(key)
                num += article['content'].count(key)
            if max_num < num:
                max_num = num
                result = keys
        print(result)
        print('出现频率：', max_num)
        print('--------------')


if __name__ == '__main__':
    article_list = load_excel('C:/excel/数据.xls')
    # keys_list = load_key('C:/excel/数据.xls')
    classify(article_list,keys_list)
