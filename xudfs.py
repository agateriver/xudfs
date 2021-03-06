# —*- coding: utf8 -*-

#### NOTE: Don't use VBA keywords for functions' arguments

import xlwings as xw
import pandas as pd
import numpy as np
import re


# 转换为文本
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


# 转换为数值
@xw.func
@xw.arg("text", doc=": 待转换的文本")
@xw.arg("to_int", doc=": 是否转换为整数，默认为True")
def xxToNumber(text, to_int=True):
    """返回去除首尾指定字符的字符串，默认去除首位全角空格、半角空格及换行"""
    if isinstance(to_int, bool) and to_int:
        return int(text)
    else:
        return float(text)


# 去除字符串首尾指定的字符,默认去除首位全角和半角空格及换行
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


# 返回起始范围内的子字串
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


# 替换某字符串匹配模式的部分为指定字符串
@xw.func
@xw.arg("text", doc=": 待替换的文本")
@xw.arg("pattern", doc=": 待替换部分模式的正则表达式")
@xw.arg("repl", doc=": 替换字符串")
def xxRegexSub(text, pattern, repl):
    """替换某字符串匹配模式的部分为指定字符串"""
    if text and isinstance(text, str):
        return re.sub(pattern, repl, text, re.MULTILINE | re.DOTALL)
    else:
        return text


# 用正则表达式分割字符串，结果横向显示
@xw.func
@xw.arg("text", doc=": 待分割的文本")
@xw.arg("pattern", doc=": 分隔符的正则表达式")
def xxRegexSplitH(text, pattern):
    """用正则表达式分割字符串，结果横向显示"""
    return re.split(pattern, text)


# 用正则表达式分割字符串，结果纵向显示
@xw.func
@xw.arg("text", doc=": 待分割的文本")
@xw.arg("pattern", doc=": 分隔符的正则表达式")
def xxRegexSplitV(text, pattern):
    """用正则表达式分割字符串，结果纵向显示"""
    return [[s] for s in xxRegexSplitH(text, pattern)]


# 将选定范围内的文本用指定的分隔符连接起来
@xw.func
@xw.arg("range_", ndim=2, doc=": 选定的范围(Range)")
@xw.arg("sep", doc=": 分隔符，默认为','")
def xxJoin(range_, sep=","):
    """将选定范围内的文本用指定的分隔符连接起来"""
    cells = [cell for row in range_ for cell in row]
    return sep.join(cells)


@xw.func
@xw.arg("ranges", expand="table", ndim=2, doc=": 选定的范围(Ranges)")
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
    return sorted([s for s in set_diff])


@xw.func
@xw.arg("range1", np.array, ndim=2, doc=": 代表集合1的范围(Range)")
@xw.arg("range2", np.array, ndim=2, doc=": 代表集合2的范围(Range)")
def xxSetDiffV(range1, range2):
    """返回两个所选范围的差集，结果纵向显示"""
    return [[s] for s in xxSetDiffH(range1, range2)]


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
    return sorted([s for s in set_diff])


@xw.func
@xw.arg("range1", np.array, ndim=2, doc=": 代表集合1的范围(Range)")
@xw.arg("range2", np.array, ndim=2, doc=": 代表集合2的范围(Range)")
def xxSetSymDiffV(range1, range2):
    """返回两个集合的对称差集，结果纵向显示"""
    return [[s] for s in xxSetSymDiffH(range1, range2)]


@xw.func
@xw.arg("range1", np.array, ndim=2, doc=": 代表集合1的范围(Range)")
@xw.arg("range2", np.array, ndim=2, doc=": 代表集合2的范围(Range)")
def xxSetIntersectH(range1, range2):
    """返回两个集合的交集，结果横向显示"""
    ss1 = set()
    for row in range1:
        for cell in row:
            ss1.add(cell)
    ss2 = set()
    for row in range2:
        for cell in row:
            ss2.add(cell)
    set_intersect = ss1.intersection(ss2)
    return sorted([s for s in set_intersect])


@xw.func
@xw.arg("range1", np.array, ndim=2, doc=": 代表集合1的范围(Range)")
@xw.arg("range2", np.array, ndim=2, doc=": 代表集合2的范围(Range)")
def xxSetIntersectV(range1, range2):
    """返回两个集合的交集，结果纵向显示"""
    return [[s] for s in xxSetIntersectH(range1, range2)]


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
@xw.arg("ranges", expand="table", ndim=2,doc=": 选定的范围(Ranges)")
def xxVStack(*ranges):
    """新Excel函数VStack模拟"""
    return np.vstack(ranges)

@xw.func
@xw.arg("ranges", expand="table", ndim=2,doc=": 选定的范围(Ranges)")
def xxHStack(*ranges):
    """新Excel函数HStack模拟"""
    return np.hstack(ranges)

# for debug
if __name__ == "__main__":
    xw.serve()
