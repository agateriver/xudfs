# —*- coding: utf8 -*-

#### NOTE: Don't use VBA keywords for functions' arguments

# from typing import Set
# from ordered_set import OrderedSet as set
import xlwings as xw
import pandas as pd
import numpy as np
import re


# 转换为文本
@xw.func
@xw.arg("number_", doc="To be converted value")
@xw.arg("is_int", doc="value is integer")
def xxToText(number_, is_int=True):
    """返回去除首尾指定字符的字符串，默认去除首位全角空格、半角空格及换行"""
    if not isinstance(number_, str):
        if is_int:
            return str(int(number_))
        else:
            return str(number_)
    else:
        return number_


# 转换为数值
@xw.func
@xw.arg("text", doc="To be converted string")
@xw.arg("to_int", doc="convert to integer?")
def xxToNumber(text, to_int=True):
    """返回去除首尾指定字符的字符串，默认去除首位全角空格、半角空格及换行"""
    if to_int:
        return int(text)
    else:
        return float(text)


# 默认去除首位全角和半角空格及换行
@xw.func
@xw.arg("text", doc="To be stripped string")
@xw.arg(
    "pattern",
    doc="Regex for pattern, default is '　  \r\n' (0x3000,0x0020,0x00A0,\r,\n)")
def xxStrip(text, pattern="　  \r\n"):  # 三种空格(0x3000,0x0020,0x00A0)、换行
    """返回去除首尾指定字符的字符串，默认去除首位全角空格、半角空格及换行"""
    if isinstance(text, str) and text:
        return str.strip(text, pattern)
    else:
        return text


# 返回起始范围内的字串
@xw.func
@xw.arg("text", doc="To be sliced string")
@xw.arg("start_", doc="index for starting, default=''")
@xw.arg("end_", doc="index for endding, default=''")
def xxSlice(text, start_="", end_=""):
    """返回起始范围内的字串"""
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


# 用正则表达式替换字符串
@xw.func
@xw.arg("text", doc="To be replaced string")
@xw.arg("pattern", doc="Regex for pattern")
@xw.arg("repl", doc="Replacement string")
def xxRegexSub(text, pattern, repl):
    """用正则表达式替换字符串"""
    if text and isinstance(text, str):
        return re.sub(pattern, repl, text, re.MULTILINE | re.DOTALL)
    else:
        return text


# 用正则表达式分割字符串，结果横向显示
@xw.func
@xw.arg("text", doc="To be replaced string")
@xw.arg("pattern", doc="Regex for pattern")
def xxRegexSplitH(text, pattern):
    """用正则表达式分割字符串，结果横向显示"""
    return re.split(pattern, text)


# 用正则表达式分割字符串，结果纵向显示
@xw.func
@xw.arg("text", doc="To be replaced string")
@xw.arg("pattern", doc="Regex for pattern")
def xxRegexSplitV(text, pattern):
    """用正则表达式分割字符串，结果纵向显示"""
    return [[s] for s in xxRegexSplitH(text, pattern)]


# 将选定range内的字符串用sep连接起来
@xw.func
@xw.arg("range_", ndim=2, doc="Selected Range")
@xw.arg("sep", doc="sep")
def xxJoin(range_, sep=","):
    """将选定range内的字符串用sep连接起来"""
    cells = [cell for row in range_ for cell in row]
    return sep.join(cells)


@xw.func
@xw.arg("ranges", expand="table", ndim=2)
def xxSetUnionH(*ranges):
    """以列的形式返回所选ranges内所有唯一值的并集"""
    ss = set()
    for range in [rng for rng in ranges if rng is not None]:
        for row in range:
            for cell in row:
                ss.add(cell)
    return sorted([s for s in ss])


@xw.func
@xw.arg("ranges", expand="table", ndim=2)
def xxSetUnionV(*ranges):
    """以列的形式返回所选ranges内所有唯一值的并集"""
    return [[s] for s in xxSetUnionH(*ranges)]


@xw.func
@xw.arg("range1", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2", np.array, ndim=2, doc="Range for Set 2")
def xxSetDiffH(range1, range2):
    """以列的形式返回两个所选范围的差集"""
    ss1 = set()
    for row in range1:
        for cell in row:
            # print(cell)
            ss1.add(cell)
    ss2 = set()
    for row in range2:
        for cell in row:
            ss2.add(cell)
    set_diff = ss1.difference(ss2)
    return sorted([s for s in set_diff])


@xw.func
@xw.arg("range1", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2", np.array, ndim=2, doc="Range for Set 2")
def xxSetDiffV(range1, range2):
    """以列的形式返回两个所选范围的差集"""
    return [[s] for s in xxSetDiffH(range1, range2)]


@xw.func
@xw.arg("range1", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2", np.array, ndim=2, doc="Range for Set 2")
def xxSetSymDiffH(range1, range2):
    """以列的形式返回两个所选范围的对称差集"""
    ss1 = set()
    for row in range1:
        for cell in row:
            # print(cell)
            ss1.add(cell)
    ss2 = set()
    for row in range2:
        for cell in row:
            ss2.add(cell)
    set_diff = ss1.symmetric_difference(ss2)
    return sorted([s for s in set_diff])


@xw.func
@xw.arg("range1", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2", np.array, ndim=2, doc="Range for Set 2")
def xxSetSymDiffV(range1, range2):
    """以行的形式返回两个所选范围的对称差集"""
    return [[s] for s in xxSetSymDiffH(range1, range2)]


@xw.func
@xw.arg("range1", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2", np.array, ndim=2, doc="Range for Set 2")
def xxSetIntersectH(range1, range2):
    """以列的形式返回两个所选范围的交集"""
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
@xw.arg("range1", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2", np.array, ndim=2, doc="Range for Set 2")
def xxSetIntersectV(range1, range2):
    """以行的形式返回两个所选范围的交集"""
    return [[s] for s in xxSetIntersectH(range1, range2)]


@xw.func
@xw.arg("range1", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2", np.array, ndim=2, doc="Range for Set 2")
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
@xw.arg("range1", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2", np.array, ndim=2, doc="Range for Set 2")
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
@xw.arg("range1", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2", np.array, ndim=2, doc="Range for Set 2")
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


# for debug
if __name__ == "__main__":
    xw.serve()
