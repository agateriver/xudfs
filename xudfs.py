# —*- coding: utf8 -*-

#### NOTE: Don't use VBA keywords for functions' arguments

from random import choice
import string
import xlwings as xw
import pandas as pd  # noqa: F401
import numpy as np
from faker import Faker
import re
import math
import pypinyin as py
import chinese_stroke_sorting as css
import pyodbc

__password_chars__ = list(
    set(string.ascii_letters  + string.digits).difference(
        set('01lIoO\'"[]:;{}()`@%~.,<')))

__password_chars_2__ = list(
    set(string.ascii_letters + string.punctuation + string.digits).difference(
        set('01lIoO\'"[]:;{}()`@%~.,<')))

def get_rand_password(digits=8,include_punctuation=False):
    if include_punctuation:
        return "".join(choice(__password_chars_2__) for x in range(0, digits))
    return "".join(choice(__password_chars__) for x in range(0, digits))

@xw.func
@xw.arg("digits", doc=": 密码位数，默认为8")
@xw.arg("include_punctuation", doc=": 是否包含标点符号，默认为False")
def xxRandPassword(digits =8, include_punctuation=False):
    """返回随机密码"""
    return get_rand_password(int(digits),include_punctuation)

@xw.func
@xw.arg("number_", doc=": 待转换的值")
@xw.arg("is_int", doc=": 是否为整数，默认为True")
def xxToText(number_, is_int=True):
    """返回去除首尾指定字符的字符串，默认去除首位全角空格、半角空格及换行"""
    if not isinstance(number_, str):
        if isinstance(is_int, bool) and is_int:
            return str(int(number_))
        else:
            return str(number_)
    else:
        return number_


@xw.func
@xw.arg("text", doc=": 待转换的文本")
@xw.arg("to_int", doc=": 是否转换为整数，默认为True")
def xxToNumber(text, to_int=True):
    """返回去除首尾指定字符的字符串，默认去除首位全角空格、半角空格及换行"""
    if isinstance(to_int, bool) and to_int:
        return int(text)
    else:
        return float(text)


@xw.func
@xw.arg("text", doc=": 待修剪的文本")
@xw.arg(
    "pattern",
    doc=": 首尾要去除部分的正则表达式，默认为'　  \\r\\n' (0x3000,0x0020,0x00A0,\\r,\\n)")
def xxStrip(text, pattern="　  \r\n"):  # 三种空格(0x3000,0x0020,0x00A0)、换行
    """修剪掉字符串首尾匹配指定模式的字符,默认去除首位全角和半角空格及换行"""
    if isinstance(text, str) and text:
        return str.strip(text, pattern)
    else:
        return text


@xw.func
@xw.arg("text", doc=": 待截取的文本")
@xw.arg("start_", doc=": 开始位置, 默认=''")
@xw.arg("end_", doc=": 结束位置, 默认=''")
def xxSlice(text, start_="", end_=""):
    """返回起始范围内的子字串"""
    if isinstance(text, str) and text:
        if start_ == '' and end_ == '':
            return text
        elif end_ == '':
            return text[int(start_):]
        elif start_ == '':
            return text[:int(end_)]
        else:
            return text[int(start_):int(end_)]
    else:
        return text


@xw.func
@xw.arg("text", doc=": 待替换的文本")
@xw.arg("pattern", doc=": 待替换部分模式的正则表达式")
@xw.arg("repl", doc=": 替换字符串")
def xxRegexSub(text, pattern, repl):
    """替换某字符串匹配模式的部分为指定字符串"""
    if text and isinstance(text, str):
        if pattern and isinstance(pattern, str):
            if repl is None:
                return re.sub(pattern, '', text,0,re.MULTILINE | re.DOTALL)
            if repl and isinstance(repl, str):
                return re.sub(pattern, repl, text,0,re.MULTILINE | re.DOTALL)
            return text
        else:
            return text
    else:
        return text


@xw.func
@xw.arg("text", doc=": 待分割的文本")
@xw.arg("sep_pattern", doc=": 分隔符的正则表达式")
def xxRegexSplitH(text, sep_pattern):
    """用正则表达式分割字符串，结果横向显示"""
    return re.split(sep_pattern, text)


@xw.func
@xw.arg("text", doc=": 待分割的文本")
@xw.arg("sep_pattern", doc=": 分隔符的正则表达式")
def xxRegexSplitV(text, sep_pattern):
    """用正则表达式分割字符串，结果纵向显示"""
    return [[s] for s in xxRegexSplitH(text, sep_pattern)]


@xw.func
@xw.arg("range_", ndim=2, doc=": 选定的范围(Range)")
@xw.arg("sep", doc=": 分隔符，默认为','")
def xxJoin(range_, sep=","):
    """将选定范围内的文本用指定的分隔符连接起来"""
    cells = [cell for row in range_ for cell in row]
    return sep.join(cells)


@xw.func
@xw.arg("ranges", ndim=2, doc=": 选定的范围(Ranges)")
def xxSetUnionH(*ranges):
    """返回所选ranges内所有唯一值的并集，结果横向显示"""
    ss = set()
    for range in [rng for rng in ranges if rng is not None]:
        for row in range:
            for cell in row:
                ss.add(cell)
    return sorted([s for s in ss])


@xw.func
@xw.arg("ranges", expand="table", ndim=2,doc=": 选定的范围(Ranges)")
def xxSetUnionV(*ranges):
    """返回所选ranges内所有唯一值的并集，结果纵向显示"""
    return [[s] for s in xxSetUnionH(*ranges)]


@xw.func
@xw.arg("range1", np.array, ndim=2, doc=": 代表集合1的范围(Range)")
@xw.arg("range2", np.array, ndim=2, doc=": 代表集合2的范围(Range)")
def xxSetDiffH(range1, range2):
    """返回两个所选范围的差集，结果横向显示"""
    ss1 = set()
    for row in range1:
        for cell in row:
            ss1.add(cell)
    ss2 = set()
    for row in range2:
        for cell in row:
            ss2.add(cell)
    set_diff = ss1.difference(ss2)
    if set_diff:
        return sorted([s for s in set_diff])
    else:
        return None


@xw.func
@xw.arg("range1", np.array, ndim=2, doc=": 代表集合1的范围(Range)")
@xw.arg("range2", np.array, ndim=2, doc=": 代表集合2的范围(Range)")
def xxSetDiffV(range1, range2):
    """返回两个所选范围的差集，结果纵向显示"""
    diff = xxSetDiffH(range1, range2)
    if diff:
        return [[s] for s in diff]
    else:
        return None


@xw.func
@xw.arg("range1", np.array, ndim=2, doc=": 代表集合1的范围(Range)")
@xw.arg("range2", np.array, ndim=2, doc=": 代表集合2的范围(Range)")
def xxSetSymDiffH(range1, range2):
    """返回两个集合的对称差集，结果横向显示"""
    ss1 = set()
    for row in range1:
        for cell in row:
            ss1.add(cell)
    ss2 = set()
    for row in range2:
        for cell in row:
            ss2.add(cell)
    set_diff = ss1.symmetric_difference(ss2)
    if set_diff:
        return sorted([s for s in set_diff])
    else:
        return None


@xw.func
@xw.arg("range1", np.array, ndim=2, doc=": 代表集合1的范围(Range)")
@xw.arg("range2", np.array, ndim=2, doc=": 代表集合2的范围(Range)")
def xxSetSymDiffV(range1, range2):
    """返回两个集合的对称差集，结果纵向显示"""
    diff = xxSetSymDiffH(range1, range2)
    if diff:
        return [[s] for s in diff]
    else:
        return None


@xw.func
@xw.arg("ranges", ndim=2, doc=": 选定的范围(Ranges)")
def xxSetIntersectH(*ranges):
    """返回所选集合的交集，结果横向显示"""
    ss = set()
    for idx, range in enumerate([rng for rng in ranges if rng is not None]):
        if idx==0:
            ss=set([cell for row in range for cell in row])
        else:
            ss = ss.intersection(set([cell for row in range for cell in row]))
    return sorted([s for s in ss])


@xw.func
@xw.arg("ranges", ndim=2, doc=": 选定的范围(Ranges)")
def xxSetIntersectV(*ranges):
    """返回两个集合的交集，结果纵向显示"""
    return [[s] for s in xxSetIntersectH(*ranges)]


@xw.func
@xw.arg("range1", np.array, ndim=2, doc=": 代表集合1的范围(Range)")
@xw.arg("range2", np.array, ndim=2, doc=": 代表集合2的范围(Range)")
def xxSetIsSubset(range1, range2):
    """报告第一个集合是否是第二个集合的子集"""
    ss1 = set()
    for row in range1:
        for cell in row:
            ss1.add(cell)
    ss2 = set()
    for row in range2:
        for cell in row:
            ss2.add(cell)
    return ss1.issubset(ss2)


@xw.func
@xw.arg("range1", np.array, ndim=2, doc=": 代表集合1的范围(Range)")
@xw.arg("range2", np.array, ndim=2, doc=": 代表集合2的范围(Range)")
def xxSetIsSuperSet(range1, range2):
    """报告第一个集合是否是第二个集合的超集"""
    ss1 = set()
    for row in range1:
        for cell in row:
            ss1.add(cell)
    ss2 = set()
    for row in range2:
        for cell in row:
            ss2.add(cell)
    return ss1.issuperset(ss2)


@xw.func
@xw.arg("range1", np.array, ndim=2, doc=": 代表集合1的范围(Range)")
@xw.arg("range2", np.array, ndim=2, doc=": 代表集合2的范围(Range)")
def xxSetIsDisjoint(range1, range2):
    """报告两个集合是否没有交集"""
    ss1 = set()
    for row in range1:
        for cell in row:
            ss1.add(cell)
    ss2 = set()
    for row in range2:
        for cell in row:
            ss2.add(cell)
    return ss1.isdisjoint(ss2)

@xw.func
@xw.arg("ranges", ndim=2, doc=": 选定的范围(Ranges)")
def xxVStack(*ranges):
    """新Excel函数VStack模拟"""
    return np.vstack(ranges)

@xw.func
@xw.arg("ranges", ndim=2, doc=": 选定的范围(Ranges)")
def xxHStack(*ranges):
    """新Excel函数HStack模拟"""
    return np.hstack(ranges)

@xw.func
@xw.arg("n", doc=": 生成的假人名数")
@xw.arg("locale", default ="zh_CN",doc=": locale,默认zh_CN")
def xxFakePersonName(n,locale ="zh_CN"):
    """Fake Name"""
    fake = Faker(locale)
    return [[fake.name()] for i in range(int(n))]

@xw.func
@xw.arg("n",doc=": 生成的假身份证数")
@xw.arg("locale", default ="zh_CN",doc=": locale,默认zh_CN")
def xxFakeSSN(n,locale ="zh_CN"):
    """Fake Name"""
    fake = Faker(locale)
    return [[fake.ssn()] for i in range(int(n))]

@xw.func
@xw.arg("n",doc=": 生成的假邮编数")
@xw.arg("locale", default ="zh_CN",doc=": locale,默认zh_CN")
def xxFakePostcode(n,locale ="zh_CN"):
    """Fake Postcode"""
    fake = Faker(locale)
    return [[fake.postcode()] for i in range(int(n))]

@xw.func
@xw.arg("n",doc=": 生成的假公司名数")
@xw.arg("locale", default ="zh_CN",doc=": locale,默认zh_CN")
def xxFakeCommany(n,locale ="zh_CN"):
    """Fake Company"""
    fake = Faker(locale)
    return [[fake.company()] for i in range(int(n))]

@xw.func
@xw.arg("n",doc=": 生成的假地址数")
@xw.arg("locale", default ="zh_CN",doc=": locale,默认zh_CN")
def xxFakeAddress(n,locale ="zh_CN"):
    """Fake Address"""
    fake = Faker(locale)
    return [[fake.address()] for i in range(int(n))]

@xw.func
@xw.arg("n",doc=": 生成的电话号码数")
@xw.arg("locale", default ="zh_CN",doc=": localee,默认zh_CN")
def xxFakePhoneNumber(n,locale ="zh_CN"):
    """Fake Phone Number"""
    fake = Faker(locale)
    return [[fake.phone_number()] for i in range(int(n))]


@xw.func
@xw.arg("names",doc=": 表示人名的列数据")
@xw.arg("cellsPerRow",default = 5, doc=": 转换后每行的单元格数")
@xw.arg("wrapByRow",default = True, doc=": 转换后的数据是按行折返还是按列折返,TRUE or FALSE,默认TRUE按行折返")
@xw.arg("fillBlank",default = True, doc=": 转换后对两字名是否填充空格补全为三字宽度,TRUE or FALSE,默认TRUE填充")
@xw.arg("ordyBy",default = "pinyin", doc=": 转换后的数据是按pinyin或是stroke排序,默认按pinyin排序")
def xxWrapNames(names, cellsPerRow=5, wrapByRow=True,fillBlank=True,ordyBy="pinyin"):
    """将一行/列中文人名转换为按拼音或笔画排序的矩阵"""
    len_names = len(names)
    if ordyBy == "pinyin":
        names.sort(key=lambda x: py.lazy_pinyin(x, style=py.Style.FIRST_LETTER))
    elif ordyBy == "stroke":
        names= css.sort_by_stroke(names)
    else:
        raise ValueError("ordyBy must be pinyin or stroke")
    if fillBlank:
        for i in range(len_names):
            if len(names[i]) == 2 and names[i][1] not in ["　"," "]:
                names[i] = names[i][0] + "　"+names[i][1]
    cellsPerRow=int(cellsPerRow)
    rows = math.ceil(len_names/cellsPerRow )
    result=[]
    if wrapByRow:
        for  i in range(rows):
            if i< rows-1:
                result.append([names[i*cellsPerRow+j] for j in range(cellsPerRow)])
            else:
                result.append([names[i*cellsPerRow+j] for j in range(len_names-i*cellsPerRow)])  # noqa: E501
    else:
        cellsInLastCol = len_names - (cellsPerRow-1)* rows
        for  i in range(rows):
            result.append([])
            
        if cellsInLastCol == cellsPerRow:
            for c in range(cellsPerRow):
                for r in range(rows):
                    result[r].append(names[c*rows+r])
        else:
            for c in range(cellsPerRow-1):
                for r in range(rows):
                    result[r].append(names[c*rows+r])
            for r in range(cellsInLastCol): # 最后一列的数据
                result[r].append(names[(cellsPerRow-1)*rows+r])

    for row in result:
        if len(row)<cellsPerRow:
            row.extend([None]*(cellsPerRow-len(row)))
    return result

@xw.func
@xw.arg("names",doc=": 表示人名的列或列数据")
@xw.arg("ordyBy",default = "pinyin", doc=": 转换后的数据是按pinyin或是stroke排序,默认按pinyin排序")
@xw.arg("sqlConStr",default = "Driver={SQL Server};Server=.;Database=msdb;Trusted_Connection=yes", doc=": SQLServer 连接字符串，默认为本机信任连接")
def xxSortCNamesViaSQLServerH(names,ordyBy = "pinyin",sqlConStr="Driver={SQL Server};Server=.;Database=msdb;Trusted_Connection=yes"):
    """通过SQL Server的排序规则将一行/列中文人名转换为按拼音或笔画排序,可指定排序规则实现其它排序""" 
    conn = pyodbc.connect(sqlConStr)
    cursor = conn.cursor()
    s="""'),('""".join(names)
    collate = ordyBy
    if ordyBy == "pinyin":
        collate = "Chinese_Simplified_Pinyin_100_CI_AS_KS_WS"
    if ordyBy in ["bihua","stroke"] :
        collate =  "Chinese_Simplified_Stroke_Order_100_CS_AS_KS_WS"
    query = f"""SELECT C FROM (VALUES ('{s}')) as T(C) order by C collate {collate}"""  # noqa: E501
    cursor.execute(query)
    result=[]
    for row in cursor:
        result.append(row[0])
    return result

@xw.func
@xw.arg("names",doc=": 表示人名的列或列数据")
@xw.arg("ordyBy",default = "pinyin", doc=": 转换后的数据是按pinyin或是stroke排序,默认按pinyin排序")
@xw.arg("sqlConStr",default = "Driver={SQL Server};Server=.;Database=msdb;Trusted_Connection=yes", doc=": SQLServer 连接字符串，默认为本机信任连接")
def xxSortCNamesViaSQLServerV(names,ordyBy = "pinyin",sqlConStr="Driver={SQL Server};Server=.;Database=msdb;Trusted_Connection=yes"):
    """通过SQL Server的排序规则将一行/列中文人名按拼音或笔画排序,可指定其它排序规则实现更多排序""" 
    conn = pyodbc.connect(sqlConStr)
    cursor = conn.cursor()
    s="""'),('""".join(names)
    collate = ordyBy
    if ordyBy.lower() == "pinyin":
        collate = "Chinese_Simplified_Pinyin_100_CI_AS_KS_WS"
    if ordyBy.lower() in ["bihua","stroke"] :
        collate =  "Chinese_Simplified_Stroke_Order_100_CS_AS_KS_WS"
    query = f"""SELECT C FROM (VALUES ('{s}')) as T(C) order by C collate {collate}"""  # noqa: E501
    cursor.execute(query)
    result=[]
    for row in cursor:
        result.append([row[0],])  
    return result


@xw.func
@xw.arg("data",doc=": 待随机分组的行或列数据")
@xw.arg("n", doc=": 分成多少组")
def xxRandomGroup(data, n):
    """将数据均分成n组,每组一列"""
    def chunks(data, n):
        import random
        random.shuffle(data)  # 随机洗牌
        int_part, rem_part = divmod(len(data), int(n))  
        i = 0
        while i < len(data):
            if rem_part > 0: #如若总样本不是是分组数的整数倍，则前几组多分一个
                yield data[i:i + int_part+1]
                rem_part -= 1
                i += int_part+1
            else: # 如若总样本正好是分组数的整数倍则均分
                yield data[i:i + int_part]
                i += int_part     
    result = list(chunks(data, int(n)))
    int_part, rem_part = divmod(len(data), int(n))  
    if rem_part>0: # 补齐为矩阵
        for i in range(rem_part,int(n)):
            result[i].append(None)
    transposed = list(map(list, zip(*result))) # 转置为每列一组
    return transposed
    
@xw.func
@xw.arg("data",ndim=2,doc=":样本总体，行、列或矩阵")
@xw.arg("n", doc=": 抽样数")
def xxRandomSampleH(data, n):
    """从总体中抽n个样本"""
    import random
    result = []
    xdata=[j for i in data for j in i]
    for i in random.sample(xdata, int(n)):
        result.append(i)
    return result

@xw.func
@xw.arg("data",ndim=2,doc=":样本总体,行、列或矩阵")
@xw.arg("n", doc=": 抽样数")
def xxRandomSampleV(data, n):
    """从总体中抽n个样本"""
    import random
    result = []
    xdata=[j for i in data for j in i]
    for i in random.sample(xdata, int(n)):
        result.append([i,])
    return result
    
@xw.func
@xw.arg("lookup_value",doc=":  查找值")
@xw.arg("lookup_array", ndim=2,doc=": 在哪一列查找")
@xw.arg("return_array", ndim=2,doc=": 返回值所在列")
def xxLookupMultiple(lookup_value, lookup_array,return_array):
    """多值查找"""
    result = []
    flatten_lookup_array = [j for i in lookup_array for j in i]
    flatten_return_array = [j for i in return_array for j in i]
    for idx,value in enumerate(flatten_lookup_array):
        if lookup_value == value:
            result.append(flatten_return_array[idx])
    return result

@xw.func
@xw.arg("data",convert=pd.DataFrame, index=0, ndim=2,doc=": 待查询的数据区，第一行为列名")
@xw.arg("expr",doc=": 查询表达式，写法参见pandas文档。如：'A > 0 and `B 1` < 0' and C.str.startswith('a') and D in [1,2,3]'")
@xw.ret(index=False)
def xxPandasQuery(data, expr):
    """pandas.DataFrame.query()的封装。"""
    qry = data.query(expr, inplace=False)
    return qry

@xw.func
@xw.arg("col_index",doc=": 以字母表示的列索引")
def xxColumnIndexToNumber(col_index):
    """将以字母表示的列索引转换为数字表示"""
    num = 0
    for c in col_index:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord('A')) + 1
    return num

@xw.func
@xw.arg("col_index",doc=": 以数字表示的列索引")
def xxColumnIndexToLetter(col_index):
    """将以数字表示的列索引转换为字母表示"""
    letter = ""
    _col_index = int(col_index)
    while _col_index > 0:
        remainder = _col_index % 26
        if remainder == 0:
            remainder = 26
        letter += chr(remainder + 64)
        _col_index = (_col_index - remainder) // 26
    return letter[::-1]

# for debug
if __name__ == "__main__":
    xw.serve()
