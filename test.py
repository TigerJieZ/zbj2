import jieba.posseg
import pytest


@pytest.fixture
def texts():
    return ['呼伦贝尔：扎兰屯市农机局开展“安全生产月”质量监督检查活动', '呼伦贝尔市：莫旗农机质监站扎实开展农机维修市场整顿工作',
            '   呼伦贝尔市：鄂伦春旗农牧业局召开农机深松整地技术推广座谈会', '呼伦贝尔市：鄂伦春旗农牧业局高效调解农机质量投诉解民忧',
            '包头市：2016年农机质量跟踪调查机型已经确定', '乌兰察布市：组织党员干部参观战役纪念馆实践活动',
            '内蒙古公布肉及肉制品抽检结果：不合格6批次', '内蒙古自治区政府法律顾问委员会办公室招聘工作人员的公告',
            '关于公布自治区人民政府继续有效规范性文件目录的通告', '2016年度内蒙古职称外语等级考试12月10日起报名',
            '内蒙古自治区成品油价格调整公告']


def test_jieba(texts):
    """
    测试jieba的词性，通过词性将相邻词连接成短语
    :param text:
    :return:
    """
    for line in texts:
        words_ = jieba.posseg.cut(line)
        for word in words_:
            print(word)