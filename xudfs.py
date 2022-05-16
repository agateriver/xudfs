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
@xw.arg("string_", doc="To be converted string")
@xw.arg("to_int", doc="convert to integer")
def xxToNumber(string_, to_int=True):
    """返回去除首尾指定字符的字符串，默认去除首位全角空格、半角空格及换行"""
    if to_int:
        return int(string_)
    else:
        return float(string_)


# 默认去除首位全角和半角空格及换行
@xw.func
@xw.arg("string_", doc="To be stripped string")
@xw.arg(
    "pattern",
    doc="Regex for pattern, default is '　  \r\n' (0x3000,0x0020,0x00A0,\r,\n)")
def xxStrip(string_, pattern="　  \r\n"):  # 三种空格(0x3000,0x0020,0x00A0)、换行
    """返回去除首尾指定字符的字符串，默认去除首位全角空格、半角空格及换行"""
    if isinstance(string_, str) and string_:
        return str.strip(string_, pattern)
    else:
        return string_


# 返回起始范围内的字串
@xw.func
@xw.arg("string_", doc="To be sliced string")
@xw.arg("start_", doc="index for starting, default=''")
@xw.arg("end_", doc="index for endding, default=''")
def xxSlice(string_, start_="", end_=""):
    """返回起始范围内的字串"""
    if isinstance(string_, str) and string_:
        if start_ == '' and end_ == '':
            return string_
        elif end_ == '':
            return string_[int(start_):]
        elif start_ == '':
            return string_[:int(end_)]
        else:
            return string_[int(start_):int(end_)]
    else:
        return string_


# 用正则表达式替换字符串
@xw.func
@xw.arg("string_", doc="To be replaced string")
@xw.arg("pattern", doc="Regex for pattern")
@xw.arg("repl", doc="Replacement string")
def xxRegexSub(string_, pattern, repl):
    """用正则表达式替换字符串"""
    if string_ and isinstance(string_, str):
        return re.sub(pattern, repl, string_, re.MULTILINE | re.DOTALL)
    else:
        return string_


# 用正则表达式分割字符串，结果横向显示
@xw.func
@xw.arg("string_", doc="To be replaced string")
@xw.arg("pattern", doc="Regex for pattern")
@xw.ret(expand='table')
def xxRegexSplitH(string_, pattern):
    """用正则表达式分割字符串，结果横向显示"""
    if string_ and isinstance(string_, str):
        return re.split(pattern, string_)
    else:
        return string_


# 用正则表达式分割字符串，结果纵向显示
@xw.func
@xw.arg("string_", doc="To be replaced string")
@xw.arg("pattern", doc="Regex for pattern")
@xw.ret(expand='table')
def xxRegexSplitV(string_, pattern):
    """用正则表达式分割字符串，结果纵向显示"""
    if string_ and isinstance(string_, str):
        return [[s] for s in re.split(pattern, string_)]
    else:
        return string_


# 将选定range内的字符串用sep连接起来
@xw.func
@xw.arg("range_", ndim=2, doc="Selected Range")
@xw.arg("sep", doc="sep")
def xxJoin(range_, sep=","):
    """将选定range内的字符串用sep连接起来"""
    cells = [cell for row in range_ for cell in row]
    return sep.join(cells)


# 提取选定区域的唯一字符串集合
@xw.func
@xw.arg("range1", np.array, ndim=2, doc="Range  To be disctincted")
@xw.ret(expand='table')
def xxDistinct(range1):
    """以列的形式返回单个所选ranges内所有唯一的值"""
    ss = set()
    for range in [
            range1,
    ]:
        for row in range:
            for cell in row:
                ss.add(cell)
    if ss:
        return sorted(list([[s] for s in ss]))
    else:
        return None


# 提取选定区域的唯一字符串集合
@xw.func
@xw.arg("range1", np.array, ndim=2, doc="Range  To be disctincted")
@xw.ret(expand='table')
def xxDistinctH(range1):
    """以行的形式返回单个所选ranges内所有唯一的值"""
    ss = set()
    for range in [
            range1,
    ]:
        for row in range:
            for cell in row:
                ss.add(cell)
    if ss:
        return sorted([s for s in ss])
    else:
        return None


@xw.func
@xw.arg("range1", np.array, ndim=2, doc="Range 1 To be disctincted")
@xw.arg("range2", np.array, ndim=2, doc="Range 2 To be disctincted")
@xw.arg("range3", np.array, ndim=2, doc="Range 3 To be disctincted")
@xw.arg("range4", np.array, ndim=2, doc="Range 4 To be disctincted")
@xw.arg("range5", np.array, ndim=2, doc="Range 5 To be disctincted")
@xw.arg("range6", np.array, ndim=2, doc="Range 6 To be disctincted")
@xw.arg("range7", np.array, ndim=2, doc="Range 7 To be disctincted")
@xw.arg("range8", np.array, ndim=2, doc="Range 8 To be disctincted")
@xw.arg("range9", np.array, ndim=2, doc="Range 9 To be disctincted")
@xw.arg("range10", np.array, ndim=2, doc="Range 10 To be disctincted")
@xw.arg("range11", np.array, ndim=2, doc="Range 11 To be disctincted")
@xw.arg("range12", np.array, ndim=2, doc="Range 12 To be disctincted")
@xw.arg("range13", np.array, ndim=2, doc="Range 13 To be disctincted")
@xw.arg("range14", np.array, ndim=2, doc="Range 14 To be disctincted")
@xw.arg("range15", np.array, ndim=2, doc="Range 15 To be disctincted")
@xw.arg("range16", np.array, ndim=2, doc="Range 16 To be disctincted")
@xw.arg("range17", np.array, ndim=2, doc="Range 17 To be disctincted")
@xw.arg("range18", np.array, ndim=2, doc="Range 18 To be disctincted")
@xw.arg("range19", np.array, ndim=2, doc="Range 19 To be disctincted")
@xw.arg("range20", np.array, ndim=2, doc="Range 20 To be disctincted")
@xw.ret(expand='table')
def xxSetUnion(range1,
               range2=None,
               range3=None,
               range4=None,
               range5=None,
               range6=None,
               range7=None,
               range8=None,
               range9=None,
               range10=None,
               range11=None,
               range12=None,
               range13=None,
               range14=None,
               range15=None,
               range16=None,
               range17=None,
               range18=None,
               range19=None,
               range20=None):
    """以列的形式返回最多20个所选ranges内所有唯一值的并集"""
    ranges = [
        rng for rng in [
            range1, range2, range3, range4, range5, range6, range7, range8,
            range9, range10, range11, range12, range13, range14, range15,
            range16, range17, range18, range19, range20
        ] if rng is not None
    ]
    ss = set()
    for range in ranges:
        for row in range:
            for cell in row:
                ss.add(cell)
    if ss:
        return sorted(list([[s] for s in ss]))
    else:
        return None


# Return the difference of two sets as a new set.
@xw.func
@xw.arg("range1_", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2_", np.array, ndim=2, doc="Range for Set 2")
@xw.ret(expand='table')
def xxSetDiff(range1_, range2_):
    """以列的形式返回两个所选范围的差集"""
    ss1 = set()
    for row in range1_:
        for cell in row:
            # print(cell)
            ss1.add(cell)
    ss2 = set()
    for row in range2_:
        for cell in row:
            ss2.add(cell)
    set_diff = ss1.difference(ss2)
    if set_diff:
        # return sorted(list([[s] for s in set_diff]))
        return list([[s] for s in set_diff])
    else:
        return None


@xw.func
@xw.arg("range1", np.array, ndim=2, doc="Range 1 To be disctincted")
@xw.arg("range2", np.array, ndim=2, doc="Range 2 To be disctincted")
@xw.arg("range3", np.array, ndim=2, doc="Range 3 To be disctincted")
@xw.arg("range4", np.array, ndim=2, doc="Range 4 To be disctincted")
@xw.arg("range5", np.array, ndim=2, doc="Range 5 To be disctincted")
@xw.arg("range6", np.array, ndim=2, doc="Range 6 To be disctincted")
@xw.arg("range7", np.array, ndim=2, doc="Range 7 To be disctincted")
@xw.arg("range8", np.array, ndim=2, doc="Range 8 To be disctincted")
@xw.arg("range9", np.array, ndim=2, doc="Range 9 To be disctincted")
@xw.arg("range10", np.array, ndim=2, doc="Range 10 To be disctincted")
@xw.arg("range11", np.array, ndim=2, doc="Range 11 To be disctincted")
@xw.arg("range12", np.array, ndim=2, doc="Range 12 To be disctincted")
@xw.arg("range13", np.array, ndim=2, doc="Range 13 To be disctincted")
@xw.arg("range14", np.array, ndim=2, doc="Range 14 To be disctincted")
@xw.arg("range15", np.array, ndim=2, doc="Range 15 To be disctincted")
@xw.arg("range16", np.array, ndim=2, doc="Range 16 To be disctincted")
@xw.arg("range17", np.array, ndim=2, doc="Range 17 To be disctincted")
@xw.arg("range18", np.array, ndim=2, doc="Range 18 To be disctincted")
@xw.arg("range19", np.array, ndim=2, doc="Range 19 To be disctincted")
@xw.arg("range20", np.array, ndim=2, doc="Range 20 To be disctincted")
@xw.ret(
    expand='table',)
def xxSetUnionH(range1,
                range2=None,
                range3=None,
                range4=None,
                range5=None,
                range6=None,
                range7=None,
                range8=None,
                range9=None,
                range10=None,
                range11=None,
                range12=None,
                range13=None,
                range14=None,
                range15=None,
                range16=None,
                range17=None,
                range18=None,
                range19=None,
                range20=None):
    """以行的形式返回最多20个所选ranges内所有唯一值的并集"""
    ranges = [
        rng for rng in [
            range1, range2, range3, range4, range5, range6, range7, range8,
            range9, range10, range11, range12, range13, range14, range15,
            range16, range17, range18, range19, range20
        ] if rng is not None
    ]
    ss = set()
    for range in ranges:
        for row in range:
            for cell in row:
                ss.add(cell)
    if ss:
        return sorted([s for s in ss])
    else:
        return None


# Return the difference of two sets as a new set.
@xw.func
@xw.arg("range1_", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2_", np.array, ndim=2, doc="Range for Set 2")
@xw.ret(expand='table')
def xxSetDiff(range1_, range2_):
    """以列的形式返回两个所选范围的差集"""
    ss1 = set()
    for row in range1_:
        for cell in row:
            # print(cell)
            ss1.add(cell)
    ss2 = set()
    for row in range2_:
        for cell in row:
            ss2.add(cell)
    set_diff = ss1.difference(ss2)
    if set_diff:
        # return sorted(list([[s] for s in set_diff]))
        return list([[s] for s in set_diff])
    else:
        return None


# Return the difference of two sets as a new set.
@xw.func
@xw.arg("range1_", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2_", np.array, ndim=2, doc="Range for Set 2")
@xw.ret(expand='table')
def xxSetDiffH(range1_, range2_):
    """以行的形式返回两个所选范围的差集"""
    ss1 = set()
    for row in range1_:
        for cell in row:
            # print(cell)
            ss1.add(cell)
    ss2 = set()
    for row in range2_:
        for cell in row:
            ss2.add(cell)
    set_diff = ss1.difference(ss2)
    if set_diff:
        # return sorted(list([[s] for s in set_diff]))
        return list([[s] for s in set_diff])
    else:
        return None


# Return the symmetric difference of two sets as a new set.
@xw.func
@xw.arg("range1_", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2_", np.array, ndim=2, doc="Range for Set 2")
@xw.ret(expand='table')
def xxSetSymDiff(range1_, range2_):
    """以列的形式返回两个所选范围的对称差集"""
    ss1 = set()
    for row in range1_:
        for cell in row:
            # print(cell)
            ss1.add(cell)
    ss2 = set()
    for row in range2_:
        for cell in row:
            ss2.add(cell)
    set_diff = ss1.symmetric_difference(ss2)
    if set_diff:
        return sorted(list([[s] for s in set_diff]))
    else:
        return None


# Return the symmetric difference of two sets as a new set.
@xw.func
@xw.arg("range1_", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2_", np.array, ndim=2, doc="Range for Set 2")
@xw.ret(expand='table')
def xxSetSymDiffH(range1_, range2_):
    """以行的形式返回两个所选范围的对称差集"""
    ss1 = set()
    for row in range1_:
        for cell in row:
            # print(cell)
            ss1.add(cell)
    ss2 = set()
    for row in range2_:
        for cell in row:
            ss2.add(cell)
    set_diff = ss1.symmetric_difference(ss2)
    if set_diff:
        return sorted(list([[s] for s in set_diff]))
    else:
        return None


# Return the intersection of two sets as a new set.
@xw.func
@xw.arg("range1_", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2_", np.array, ndim=2, doc="Range for Set 2")
@xw.ret(expand='table')
def xxSetIntersect(range1_, range2_):
    """以列的形式返回两个所选范围的交集"""
    ss1 = set()
    for row in range1_:
        for cell in row:
            ss1.add(cell)
    ss2 = set()
    for row in range2_:
        for cell in row:
            ss2.add(cell)
    set_intersect = ss1.intersection(ss2)
    if set_intersect:
        return sorted(list([[s] for s in set_intersect]))
    else:
        return None


# Return the intersection of two sets as a new set.
@xw.func
@xw.arg("range1_", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2_", np.array, ndim=2, doc="Range for Set 2")
@xw.ret(expand='table')
def xxSetIntersectH(range1_, range2_):
    """以行的形式返回两个所选范围的交集"""
    ss1 = set()
    for row in range1_:
        for cell in row:
            ss1.add(cell)
    ss2 = set()
    for row in range2_:
        for cell in row:
            ss2.add(cell)
    set_intersect = ss1.intersection(ss2)
    if set_intersect:
        return sorted(list([[s] for s in set_intersect]))
    else:
        return None


# Report whether another set contains this set.
@xw.func
@xw.arg("range1_", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2_", np.array, ndim=2, doc="Range for Set 2")
def xxSetIsSubset(range1_, range2_):
    """报告第一个集合是否是第二个集合的子集"""
    ss1 = set()
    for row in range1_:
        for cell in row:
            ss1.add(cell)
    ss2 = set()
    for row in range2_:
        for cell in row:
            ss2.add(cell)
    return ss1.issubset(ss2)


# Report whether this set contains another set.
@xw.func
@xw.arg("range1_", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2_", np.array, ndim=2, doc="Range for Set 2")
def xxSetIsSuperSet(range1_, range2_):
    """报告第一个集合是否是第二个集合的超集"""
    ss1 = set()
    for row in range1_:
        for cell in row:
            ss1.add(cell)
    ss2 = set()
    for row in range2_:
        for cell in row:
            ss2.add(cell)
    return ss1.issuperset(ss2)


# Return True if two sets have a null intersection.
@xw.func
@xw.arg("range1_", np.array, ndim=2, doc="Range for Set 1")
@xw.arg("range2_", np.array, ndim=2, doc="Range for Set 2")
def xxSetIsDisjoint(range1_, range2_):
    """报告两个集合是否没有交集"""
    ss1 = set()
    for row in range1_:
        for cell in row:
            ss1.add(cell)
    ss2 = set()
    for row in range2_:
        for cell in row:
            ss2.add(cell)
    return ss1.isdisjoint(ss2)


# for debug
if __name__ == "__main__":
    xw.serve()
