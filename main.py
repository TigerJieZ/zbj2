import xlrd
import os
import utils


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


def load_excel(filename: str, sheet_index: int, title_col: int):
    excel_file = xlrd.open_workbook(filename)
    sheet = excel_file.sheet_by_index(sheet_index)
    article_list_ = []
    for i in range(1, sheet.nrows):
        article = {'title': sheet.cell(i, title_col).value}
        article_list_.append(article)
    return article_list_


if __name__ == '__main__':
    EXTERSIONS = ['xls', 'xlsm', 'xlsx']
    std_keys = utils.load_std_keys()
    dir_ = 'F:/pycharm/zbj2/test_data/'
    for path in os.listdir(dir_):
        if EXTERSIONS.count(path.split('.')[-1]) > 0:
            article_list = load_excel(dir_ + path, sheet_index=1, title_col=0)
            # keys_list_ = load_key('C:/excel/数据.xls')
            result_list = utils.classify_subject(article_list, std_keys, index='std_key')
            utils.write_result(dir_ + path, result_list, w_key_col=5, sheet_index=1, w_id_col=4)

    # article_list = load_excel('C:/excel/信息公开目录.xls')
    # # keys_list_ = load_key('C:/excel/数据.xls')
    # result_list = classify(article_list, keys_list_new)
    # write_result('C:/excel/信息公开目录.xls', result_list)
