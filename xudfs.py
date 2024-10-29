# —*- coding: utf8 -*-

#### NOTE: Don't use VBA keywords for functions' arguments

import string
import xlwings as xw
import pandas as pd
import numpy as np
import re
import math
import pyodbc
from random import choice
from dbfread import DBF
from faker import Faker
import pypinyin as py
import chinese_stroke_sorting as css

pd.options.future.infer_string = True  # for pnadas>2.1

__password_chars__ = list(
    set(string.ascii_letters + string.digits).difference(
        set("01lIoO'\"[]:;{}()`@%~.,<")
    )
)

__password_chars_2__ = list(
    set(string.ascii_letters + string.punctuation + string.digits).difference(
        set("01lIoO'\"[]:;{}()`@%~.,<")
    )
)

__kilos__ = "天地玄黄，宇宙洪荒。日月盈昃，辰宿列张。寒来暑往，秋收冬藏。闰余成岁，律吕调阳。云腾致雨，露结为霜。金生丽水，玉出昆冈。剑号巨阙，珠称夜光。果珍李柰，菜重芥姜。海咸河淡，鳞潜羽翔。龙师火帝，鸟官人皇。始制文字，乃服衣裳。推位让国，有虞陶唐。吊民伐罪，周发殷汤。坐朝问道，垂拱平章。爱育黎首，臣伏戎羌。遐迩一体，率宾归王。鸣凤在竹，白驹食场。化被草木，赖及万方。盖此身发，四大五常。恭惟鞠养，岂敢毁伤。女慕贞洁，男效才良。知过必改，得能莫忘。罔谈彼短，靡恃己长。信使可覆，器欲难量。墨悲丝染，诗赞羔羊。景行维贤，克念作圣。德建名立，形端表正。空谷传声，虚堂习听。祸因恶积，福缘善庆。尺璧非宝，寸阴是竞。资父事君，曰严与敬。孝当竭力，忠则尽命。临深履薄，夙兴温凊。似兰斯馨，如松之盛。川流不息，渊澄取映。容止若思，言辞安定。笃初诚美，慎终宜令。荣业所基，籍甚无竟。学优登仕，摄职从政。存以甘棠，去而益咏。乐殊贵贱，礼别尊卑。上和下睦，夫唱妇随。外受傅训，入奉母仪。诸姑伯叔，犹子比儿。孔怀兄弟，同气连枝。交友投分，切磨箴规。仁慈隐恻，造次弗离。节义廉退，颠沛匪亏。性静情逸，心动神疲。守真志满，逐物意移。坚持雅操，好爵自縻。都邑华夏，东西二京。背邙面洛，浮渭据泾。宫殿盘郁，楼观飞惊。图写禽兽，画彩仙灵。丙舍旁启，甲帐对楹。肆筵设席，鼓瑟吹笙。升阶纳陛，弁转疑星。右通广内，左达承明。既集坟典，亦聚群英。杜稿钟隶，漆书壁经。府罗将相，路侠槐卿。户封八县，家给千兵。高冠陪辇，驱毂振缨。世禄侈富，车驾肥轻。策功茂实，勒碑刻铭。磻溪伊尹，佐时阿衡。奄宅曲阜，微旦孰营。桓公匡合，济弱扶倾。绮回汉惠，说感武丁。俊乂密勿，多士寔宁。晋楚更霸，赵魏困横。假途灭虢，践土会盟。何遵约法，韩弊烦刑。起翦颇牧，用军最精。宣威沙漠，驰誉丹青。九州禹迹，百郡秦并。岳宗泰岱，禅主云亭。雁门紫塞，鸡田赤城。昆池碣石，钜野洞庭。旷远绵邈，岩岫杳冥。治本于农，务兹稼穑。俶载南亩，我艺黍稷。税熟贡新，劝赏黜陟。孟轲敦素，史鱼秉直。庶几中庸，劳谦谨敕。聆音察理，鉴貌辨色。贻厥嘉猷，勉其祗植。省躬讥诫，宠增抗极。殆辱近耻，林皋幸即。两疏见机，解组谁逼。索居闲处，沉默寂寥。求古寻论，散虑逍遥。欣奏累遣，戚谢欢招。渠荷的历，园莽抽条。枇杷晚翠，梧桐早凋。陈根委翳，落叶飘摇。游鹍独运，凌摩绛霄。耽读玩市，寓目囊箱。易輶攸畏，属耳垣墙。具膳餐饭，适口充肠。饱饫烹宰，饥厌糟糠。亲戚故旧，老少异粮。妾御绩纺，侍巾帷房。纨扇圆絜，银烛炜煌。昼眠夕寐，蓝笋象床。弦歌酒宴，接杯举觞。矫手顿足，悦豫且康。嫡后嗣续，祭祀烝尝。稽颡再拜，悚惧恐惶。笺牒简要，顾答审详。骸垢想浴，执热愿凉。驴骡犊特，骇跃超骧。诛斩贼盗，捕获叛亡。布射僚丸，嵇琴阮啸。恬笔伦纸，钧巧任钓。释纷利俗，并皆佳妙。毛施淑姿，工颦妍笑。年矢每催，曦晖朗曜。璇玑悬斡，晦魄环照。指薪修祜，永绥吉劭。矩步引领，俯仰廊庙。束带矜庄，徘徊瞻眺。孤陋寡闻，愚蒙等诮。谓语助者，焉哉乎也。"

__DLTB_Fields__ = {  # 地类图斑字段名
    "BSM": "标识码",
    "YSDM": "要素代码",
    "TBYBH": "图斑预编号",
    "TBBH": "图斑编号",
    "DLBM": "地类编码",
    "DLMC": "地类名称",
    "QSXZ": "权属性质",
    "QSDWDM": "权属单位代码",
    "QSDWMC": "权属单位名称",
    "ZLDWDM": "座落单位代码",
    "ZLDWMC": "座落单位名称",
    "TBMJ": "图斑面积",
    "KCDLBM": "扣除地类编码",
    "KCXS": "扣除地类系数",
    "KCMJ": "扣除地类面积",
    "TBDLMJ": "图斑地类面积",
    "GDLX": "耕地类型",
    "GDPDJB": "耕地坡度级别",
    "XZDWKD": "线状地物宽度",
    "XXTBKD": "线性图斑宽度",
    "TBXHDM": "图斑细化代码",
    "TBXHMC": "图斑细化名称",
    "GDZZSXDM": "耕地种植属性代码",
    "GDZZSXMC": "耕地种植属性名称",
    "GDDB": "耕地等别",
    "FRDBS": "飞入地标识",
    "CZCSXM": "城镇村属性码",
    "SJNF": "数据年份",
    "BZ": "备注",
    "XZQMC": "所属乡镇名称",
    "XZQDM": "所属乡镇代码",
}

__YZ_Fields__ = {  # 研招字段
    "BKDSM": "报考点说明",
    "BKDWDM": "报考单位代码",
    "BKDWMC": "报考单位名称",
    "BKLBM": "报考类别码",
    "BKXXFS": "报考学习方式",
    "BKYXSM": "报考院系所码",
    "BKYXSMC": "报考院系所名称",
    "BKZYDM": "报考专业代码",
    "BKZYMC": "报考专业名称",
    "BMH": "报名号",
    "BMSJ": "报名时间",
    "BYDW": "毕业单位",
    "BYDWM": "毕业单位码",
    "BYDWSSM": "毕业单位省市代码",
    "BYNY": "毕业年月",
    "BYXX": "备用信息",
    "BYXX1": "备用信息1",
    "BYXX2": "备用信息2",
    "BYXX3": "备用信息3",
    "BYZYDM": "毕业专业代码",
    "BYZYMC": "毕业专业名称",
    "CSDM": "出生地码",
    "CSJF": "初试加分",
    "CSRQ": "出生日期",
    "DASZDM": "档案所在地码",
    "DASZDW": "档案所在单位",
    "DASZDWDZ": "档案所在单位地址",
    "DASZDWYZBM": "档案所在单位邮政编码",
    "DXWPDW": "定向委培单位",
    "DXWPDWSZDM": "定向委培单位所在地码",
    "DZXX": "电子信箱",
    "FSTZHFSJ": "复试通知回复时间",
    "HFM": "婚姻状况码",
    "HKSZDM": "户口所在地码",
    "HKSZDXXDZ": "户口所在地详细地址",
    "JFBZ": "交费标志",
    "JGSZDM": "籍贯所在地码",
    "JLCF": "奖励惩罚",
    "JTCY": "家庭成员",
    "KSBH": "考生编号",
    "KSFSM": "考试方式码",
    "KSLYM": "考生来源码",
    "KSZBQK": "考生作弊情况",
    "KZYXX": "跨专业信息",
    "MZM": "民族码",
    "RWNY": "入伍年月",
    "RWPZDW": "批准入伍机关名称",
    "RWPZSBH": "入伍批准书编号",
    "RWQGXDM": "入伍前所在高校代码",
    "RWQGXMC": "入伍前所在高校名称",
    "RWQGXSSM": "入伍前所在高校省市码",
    "RWQRXNY": "入伍前所在高校入学年月",
    "RWQXLZSBH": "入伍前所在高校毕业证书编号",
    "RWQXXXS": "入伍前所在高校学习形式",
    "RWQZYDM": "入伍前所在高校学习专业代码",
    "RWQZYMC": "入伍前所在高校学习专业名称",
    "SSMZZG": "少数民族照顾",
    "TCXYZBH": "退出现役证编号",
    "TJ_BZ": "调剂_备注",
    "TJ_DWDM": "调剂_单位代码",
    "TJ_DWMC": "调剂_单位名称",
    "TJ_LXDH": "调剂_联系电话",
    "TJ_LXDH2": "调剂_紧急联系电话",
    "TJ_SFYZY": "调剂_是否一志愿",
    "TJ_SSMZZG": "调剂_少数民族政策",
    "TJ_XXFS": "调剂_学习方式",
    "TJ_YJFXDM": "调剂_研究方向代码",
    "TJ_YJFXMC": "调剂_研究方向名称",
    "TJ_YXSDM": "调剂_院系所代码",
    "TJ_YXSMC": "调剂_院系所名称",
    "TJ_ZXJH": "调剂_专项计划",
    "TJ_ZYDM": "调剂_专业代码",
    "TJ_ZYMC": "调剂_专业名称",
    "TXDZ": "通讯地址",
    "TYNY": "批准退役年月",
    "TYPZDW": "批准退役机关名称",
    "TYZX": "退役专项",
    "WGY": "外国语",
    "WGYM": "外国语码",
    "WGYMC": "外国语名称",
    "XBM": "性别码",
    "XLM": "学历码",
    "XLZSBH": "学历证书编号",
    "XM": "姓名",
    "XMPY": "姓名拼音",
    "XWM": "最后学位码",
    "XWZSBH": "学位证书编号",
    "XXGZDW": "学习工作单位",
    "XXGZJL": "学习工作经历",
    "XXXS": "最后学历学习形式",
    "XYJRM": "现役军人码",
    "YDDH": "移动电话",
    "YJFXM": "研究方向码",
    "YJFXMC": "研究方向名称",
    "YWK1": "业务课1",
    "YWK1M": "业务课1码",
    "YWK1MC": "业务课1名称",
    "YWK2": "业务课2",
    "YWK2M": "业务课2码",
    "YWK2MC": "业务课2名称",
    "ZCXH": "注册学号",
    "ZF": "总分",
    "ZJHM": "证件号码",
    "ZJLX": "证件类型",
    "ZSDWSM": "招生单位说明",
    "ZXBZ": "照相标志",
    "ZXJH": "专项计划",
    "ZZLL": "政治理论成绩",
    "ZZLLM": "政治理论码",
    "ZZLLMC": "政治理论名称",
    "ZZMMM": "政治面貌码",
}


def get_rand_password(digits=8, include_punctuation=False):
    if include_punctuation:
        return "".join(choice(__password_chars_2__) for x in range(0, digits))
    return "".join(choice(__password_chars__) for x in range(0, digits))


@xw.func
@xw.arg("digits", doc=": 密码位数，缺省为8", numbers=int)
@xw.arg("include_punctuation", doc=": 是否包含标点符号，缺省为False")
def xxRandPassword(digits=8, include_punctuation=False):
    """返回随机密码"""
    return get_rand_password(digits, include_punctuation)


@xw.func
@xw.arg("number_", doc=": 待转换的值")
@xw.arg("is_int", doc=": 是否为整数，缺省为True")
def xxToText(number_, is_int=True):
    """数值转换为文本"""
    if not isinstance(number_, str):
        if isinstance(is_int, bool) and is_int:
            return str(int(number_))
        else:
            return str(number_)
    else:
        return number_


@xw.func
@xw.arg("text", doc=": 待转换的文本")
@xw.arg("to_int", doc=": 是否转换为整数，缺省为True")
def xxToNumber(text, to_int=True):
    """将文本转换位数值"""
    if isinstance(to_int, bool) and to_int:
        return int(text)
    else:
        return float(text)


@xw.func
@xw.arg("text", doc=": 待修剪的文本")
@xw.arg(
    "pattern",
    doc=": 首尾要去除部分的正则表达式，缺省为'　  \\r\\n' (0x3000,0x0020,0x00A0,\\r,\\n)",
)
def xxStringStrip(text, pattern="　  \r\n"):  # 三种空格(0x3000,0x0020,0x00A0)、换行
    """修剪掉字符串首尾匹配指定模式的字符,缺省去除首位全角和半角空格及换行"""
    if isinstance(text, str) and text:
        return str.strip(text, pattern)
    else:
        return text


@xw.func
@xw.arg("text", doc=": 待截取的文本")
@xw.arg("start_", doc=": 开始位置, 1-based，缺省 = 1", numbers=int)
@xw.arg("end_", doc=": 结束位置,1-based, 缺省 = len(text)", numbers=int)
def xxStringSlice(text: str, start_: int = 1, end_: int = None):
    """返回起始范围内的子字串"""
    if start_ < 1:
        raise "起始位置指定错误"
    if end_ is not None:
        if end_ < start_:
            raise "结束位置指定错误"
        return text[(start_ - 1) : (end_)]
    else:
        return text[(start_ - 1) :]


@xw.func
@xw.arg("text", doc=": 待替换的文本")
@xw.arg("pattern", doc=": 待替换部分模式的正则表达式")
@xw.arg("repl", doc=": 替换字符串")
def xxRegexSub(text, pattern, repl):
    """替换某字符串匹配模式的部分为指定字符串"""
    if text and isinstance(text, str):
        if pattern and isinstance(pattern, str):
            if repl is None:
                return re.sub(pattern, "", text, 0, re.MULTILINE | re.DOTALL)
            if repl and isinstance(repl, str):
                return re.sub(pattern, repl, text, 0, re.MULTILINE | re.DOTALL)
            return text
        else:
            return text
    else:
        return text


@xw.func
@xw.arg("text", doc=": 待分割的文本")
@xw.arg("sep_pattern", doc=": 分隔符的正则表达式")
@xw.arg(
    "item",
    doc=": 返回数组的第几项(1-based)。缺省为0则返回所有项",
    default=0,
    numbers=int,
)
def xxRegexSplitH(text, sep_pattern, item=0):
    """用正则表达式分割字符串，结果横向显示"""
    result = re.split(sep_pattern, text)
    if item == 0:
        return result
    else:
        return result[item - 1]


@xw.func
@xw.arg("text", doc=": 待分割的文本")
@xw.arg("pattern", doc=": 分隔符的正则表达式")
@xw.arg(
    "group",
    doc=": 返回第几个匹配组。缺省为1。如果用命名组，也可输入组名。",
    default=1,
    numbers=int,
)
def xxRegexExtract(text, pattern, group=1):
    """用正则表达式分割字符串，结果横向显示"""
    reobj = re.compile(pattern, re.IGNORECASE | re.DOTALL | re.MULTILINE)
    match = reobj.search(text)
    if match:
        return match.group(group)
    else:
        return ""


@xw.func
@xw.arg("text", doc=": 待分割的文本")
@xw.arg("sep_pattern", doc=": 分隔符的正则表达式")
@xw.arg(
    "item",
    doc=": 返回数组的第几项(1-based)。缺省为0则返回所有项",
    default=0,
    numbers=int,
)
def xxRegexSplitV(text, sep_pattern, item=0):
    """用正则表达式分割字符串，结果纵向显示"""
    result = [[s] for s in xxRegexSplitH(text, sep_pattern)]
    if item == 0:
        return result
    else:
        return result[item - 1][0]


@xw.func
@xw.arg("range_", ndim=2, doc=": 选定的范围(Range)")
@xw.arg("sep", doc=": 分隔符，缺省为','")
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
@xw.arg("ranges", expand="table", ndim=2, doc=": 选定的范围(Ranges)")
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
        if idx == 0:
            ss = set([cell for row in range for cell in row])
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
@xw.arg("n", doc=": 生成的假人名数", numbers=int)
@xw.arg("locale", default="zh_CN", doc=": locale,缺省zh_CN")
def xxFakePersonName(n, locale="zh_CN"):
    """Fake Name"""
    fake = Faker(locale)
    return [[fake.name()] for i in range(n)]


@xw.func
@xw.arg("n", doc=": 生成的假身份证数", numbers=int)
@xw.arg("locale", default="zh_CN", doc=": locale,缺省zh_CN")
def xxFakeSSN(n, locale="zh_CN"):
    """Fake Name"""
    fake = Faker(locale)
    return [[fake.ssn()] for i in range(n)]


@xw.func
@xw.arg("n", doc=": 生成的假邮编数", numbers=int)
@xw.arg("locale", default="zh_CN", doc=": locale,缺省zh_CN")
def xxFakePostcode(n, locale="zh_CN"):
    """Fake Postcode"""
    fake = Faker(locale)
    return [[fake.postcode()] for i in range(n)]


@xw.func
@xw.arg("n", doc=": 生成的假公司名数", numbers=int)
@xw.arg("locale", default="zh_CN", doc=": locale,缺省zh_CN")
def xxFakeCommany(n, locale="zh_CN"):
    """Fake Company"""
    fake = Faker(locale)
    return [[fake.company()] for i in range(n)]


@xw.func
@xw.arg("n", doc=": 生成的假地址数", numbers=int)
@xw.arg("locale", default="zh_CN", doc=": locale,缺省zh_CN")
def xxFakeAddress(n, locale="zh_CN"):
    """Fake Address"""
    fake = Faker(locale)
    return [[fake.address()] for i in range(n)]


@xw.func
@xw.arg("n", doc=": 生成的电话号码数", numbers=int)
@xw.arg("locale", default="zh_CN", doc=": localee,缺省zh_CN")
def xxFakePhoneNumber(n, locale="zh_CN"):
    """Fake Phone Number"""
    fake = Faker(locale)
    return [[fake.phone_number()] for i in range(n)]


@xw.func
@xw.arg("names", doc=": 表示人名的列数据")
@xw.arg("cellsPerRow", default=5, doc=": 转换后每行的单元格数", numbers=int)
@xw.arg(
    "wrapByRow",
    default=True,
    doc=": 转换后的数据是按行折返还是按列折返,TRUE or FALSE,缺省TRUE按行折返",
)
@xw.arg(
    "fillBlank",
    default=True,
    doc=": 转换后对两字名是否填充空格补全为三字宽度,TRUE or FALSE,缺省TRUE填充",
)
@xw.arg(
    "ordyBy",
    default="pinyin",
    doc=": 转换后的数据是按pinyin或是stroke排序,缺省按pinyin排序",
)
def xxWrapNames(names, cellsPerRow=5, wrapByRow=True, fillBlank=True, ordyBy="pinyin"):
    """将一行/列中文人名转换为按拼音或笔画排序的矩阵"""
    len_names = len(names)
    if ordyBy in ["py", "pinyin"]:
        names.sort(key=lambda x: py.lazy_pinyin(x, style=py.Style.FIRST_LETTER))
    elif ordyBy in ["stroke", "bihua"]:
        names = css.sort_by_stroke(names)
    else:
        raise ValueError("ordyBy must be pinyin/py or bihua/stroke")
    if fillBlank:
        for i in range(len_names):
            if len(names[i]) == 2 and names[i][1] not in ["　", " "]:
                names[i] = names[i][0] + "　" + names[i][1]
    result = []
    if wrapByRow:
        (rows, cellsInLastCol) = divmod(len_names, cellsPerRow)
        for i in range(rows):
            result.append([names[i * cellsPerRow + j] for j in range(cellsPerRow)])
        if cellsInLastCol > 0:
            result.append(names[-cellsInLastCol:])  # noqa: E501
        if cellsInLastCol > 0:
            result[rows].extend([None] * (cellsPerRow - cellsInLastCol))
    else:
        (rows, mod) = divmod(len_names, cellsPerRow)
        if mod == 0:
            for i in range(rows):
                result.append([])
            for c in range(cellsPerRow):
                for r in range(rows):
                    result[r].append(names[c * rows + r])
        else:
            rows = math.ceil(len_names / cellsPerRow)
            while (cellsPerRow - 1) * rows > len_names:
                cellsPerRow -= 1
                rows = math.ceil(len_names / cellsPerRow)
            for i in range(rows):
                result.append([])
            for c in range(cellsPerRow - 1):
                for r in range(rows):
                    result[r].append(names[c * rows + r])
            mod = len_names - (cellsPerRow - 1) * rows
            for r in range(mod):  # 最后一列的数据
                result[r].append(names[(cellsPerRow - 1) * rows + r])
            for r in range(mod, rows):
                result[r].extend([None])
    return result


@xw.func(call_in_wizard=False)
@xw.arg("names", doc=": 表示人名的列或列数据")
@xw.arg(
    "ordyBy",
    default="pinyin",
    doc=": 转换后的数据是按pinyin或是stroke排序,缺省按pinyin排序",
)
@xw.arg(
    "sqlConStr",
    default="Driver={SQL Server};Server=.;Trusted_Connection=yes",
    doc=": SQLServer 连接字符串，缺省为本机信任连接",
)
def xxSortCNamesViaSQLServerH(
    names,
    ordyBy="pinyin",
    sqlConStr="Driver={SQL Server};Server=.;Trusted_Connection=yes",
):
    """通过SQL Server的排序规则将一行/列中文人名转换为按拼音或笔画排序,可指定排序规则实现其它排序"""
    conn = pyodbc.connect(sqlConStr)
    cursor = conn.cursor()
    s = """'),('""".join(names)
    collate = ordyBy
    if ordyBy == "pinyin":
        collate = "Chinese_Simplified_Pinyin_100_CI_AS_KS_WS_SC"
    if ordyBy in ["bihua", "stroke"]:
        collate = "Chinese_Simplified_Stroke_Order_100_CS_AS_KS_WS_SC"
    query = f"""SELECT C FROM (VALUES ('{s}')) as T(C) order by C collate {collate}"""  # noqa: E501
    cursor.execute(query)
    result = []
    for row in cursor:
        result.append(row[0])
    return result


@xw.func(call_in_wizard=False)
@xw.arg("names", doc=": 表示人名的列或列数据")
@xw.arg(
    "ordyBy",
    default="pinyin",
    doc=": 转换后的数据是按pinyin或是stroke排序,缺省按pinyin排序",
)
@xw.arg(
    "sqlConStr",
    default="Driver={SQL Server};Server=.;Trusted_Connection=yes",
    doc=": SQLServer 连接字符串，缺省为本机信任连接",
)
def xxSortCNamesViaSQLServerV(
    names,
    ordyBy="pinyin",
    sqlConStr="Driver={SQL Server};Server=.;Trusted_Connection=yes",
):
    """通过SQL Server的排序规则将一行/列中文人名按拼音或笔画排序,可指定其它排序规则实现更多排序"""
    conn = pyodbc.connect(sqlConStr, unicode_results=True, timeout=5)
    cursor = conn.cursor()
    s = """'),('""".join(names)
    collate = ordyBy
    if ordyBy.lower() == "pinyin":
        collate = "Chinese_Simplified_Pinyin_100_CI_AS_KS_WS"
    if ordyBy.lower() in ["bihua", "stroke"]:
        collate = "Chinese_Simplified_Stroke_Order_100_CS_AS_KS_WS"
    query = f"""SELECT C FROM (VALUES ('{s}')) as T(C) order by C collate {collate}"""  # noqa: E501
    cursor.execute(query)
    result = []
    for row in cursor:
        result.append(
            [
                row[0],
            ]
        )
    return result


@xw.func
@xw.arg("names", doc=": 要排序的人名区域")
@xw.arg("sa_pwd", doc=": SQLServer sa用户密码")
@xw.arg(
    "server",
    default=".",
    doc=": SqlServer 服务器实例地址（如 127.0.0.1\\mssql1,61433），缺省为本机",
)
def xxSortCNamesByBihuaV(
    names,
    sa_pwd,
    server="127.0.0.1\\mssql1,61433",
):
    """按笔画排序"""
    return xxSortCNamesViaSQLServerV(
        names,
        ordyBy="bihua",
        sqlConStr=f"Driver={{SQL Server}};Server={server};UID=sa;PWD={sa_pwd}",
    )


@xw.func
@xw.arg("names", doc=": 要排序的人名区域")
@xw.arg("sa_pwd", doc=": SQLServer sa用户密码")
@xw.arg(
    "server",
    default=".",
    doc=": SqlServer 服务器实例地址（如 127.0.0.1\\mssql1,61433），缺省为本机",
)
def xxSortCNamesByPinyinV(
    names,
    sa_pwd,
    server="127.0.0.1\\mssql1,61433",
):
    """按笔画排序"""
    return xxSortCNamesViaSQLServerV(
        names,
        ordyBy="pinyin",
        sqlConStr=f"Driver={{SQL Server}};Server={server};UID=sa;PWD={sa_pwd}",
    )


@xw.func(call_in_wizard=False)
@xw.arg("data", doc=": 待随机分组的行或列数据")
@xw.arg("n", doc=": 分成多少组", numbers=int)
def xxRandomGroup(data, n):
    """将数据均分成n组,每组一列"""

    def chunks(data, n):
        import random

        random.shuffle(data)  # 随机洗牌
        int_part, rem_part = divmod(len(data), n)
        i = 0
        while i < len(data):
            if rem_part > 0:  # 如若总样本不是是分组数的整数倍，则前几组多分一个
                yield data[i : i + int_part + 1]
                rem_part -= 1
                i += int_part + 1
            else:  # 如若总样本正好是分组数的整数倍则均分
                yield data[i : i + int_part]
                i += int_part

    result = list(chunks(data, n))
    int_part, rem_part = divmod(len(data), n)
    if rem_part > 0:  # 补齐为矩阵
        for i in range(rem_part, n):
            result[i].append(None)
    transposed = list(map(list, zip(*result)))  # 转置为每列一组
    return transposed


@xw.func(call_in_wizard=False)
@xw.arg("data", ndim=2, doc=":样本总体，行、列或矩阵")
@xw.arg("n", doc=": 抽样数", numbers=int)
def xxRandomSampleH(data, n):
    """从总体中抽n个样本"""
    import random

    result = []
    xdata = [j for i in data for j in i]
    for i in random.sample(xdata, n):
        result.append(i)
    return result


@xw.func(call_in_wizard=False)
@xw.arg("data", ndim=2, doc=":样本总体,行、列或矩阵")
@xw.arg("n", doc=": 抽样数", numbers=int)
def xxRandomSampleV(data, n):
    """从总体中抽n个样本"""
    import random

    result = []
    xdata = [j for i in data for j in i]
    for i in random.sample(xdata, n):
        result.append(
            [
                i,
            ]
        )
    return result


@xw.func
@xw.arg("lookup_value", doc=":  查找值")
@xw.arg("lookup_array", ndim=2, doc=": 在哪一列查找")
@xw.arg("return_array", ndim=2, doc=": 返回值所在列")
def xxLookupMultiple(lookup_value, lookup_array, return_array):
    """多值查找"""
    result = []
    flatten_lookup_array = [j for i in lookup_array for j in i]
    flatten_return_array = [j for i in return_array for j in i]
    for idx, value in enumerate(flatten_lookup_array):
        if lookup_value == value:
            result.append(flatten_return_array[idx])
    return result


@xw.func(call_in_wizard=True)
@xw.arg(
    "data", convert=pd.DataFrame, index=0, ndim=2, doc=": 待查询的数据区，第一行为列名"
)
@xw.arg(
    "expr",
    doc=": 查询表达式，写法参见pandas文档。如：'A > 0 and `B 1` < 0' and C.str.startswith('a') and D in [1,2,3]'",
)
@xw.arg("cols", doc=": 返回各列的列名，多个列名用逗号分隔,缺省为空返回全部列")
@xw.arg("sorted_by", doc=": 按某列排序，缺省为空")
@xw.arg("ascending", doc=": 是否升序，缺省为True")
@xw.arg("headers", doc=": 是否返回列名，缺省为TRUE")
@xw.ret(index=False)
def xxPandasQuery(data, expr, cols=None, sorted_by=None, ascending=True, headers=True):
    """pandas.DataFrame.query()的封装。"""
    qry = data.query(expr, inplace=False)
    if sorted_by:
        qry = qry.sort_values(by=sorted_by, ascending=ascending)
    if cols:
        qry = qry[re.split(r"""[,，]\s*""", cols)]
    if headers:
        return qry
    else:
        return qry.values


@xw.func(call_in_wizard=True)
@xw.arg(
    "data", convert=pd.DataFrame, index=0, ndim=2, doc=": 待查询的数据区，第一行为列名"
)
@xw.arg(
    "expr",
    doc=": 查询表达式，写法参见pandas文档。如：'A > 0 and `B 1` < 0' and C.str.startswith('a') and D in [1,2,3]'",
)
@xw.ret(index=False)
def xxCountPandasQuery(data, expr, cols=None):
    """返回 pandas.DataFrame.query()的结果行数"""
    qry = data.query(expr, inplace=False)
    return qry.shape[0]


@xw.func
@xw.arg("col_index", doc=": 以字母表示的列索引")
def xxColumnIndexToNumber(col_index):
    """将以字母表示的列索引转换为数字表示"""
    num = 0
    for c in col_index:
        if c in string.ascii_letters:
            num = num * 26 + (ord(c.upper()) - ord("A")) + 1
    return num


@xw.func
@xw.arg("col_index", doc=": 以数字表示的列索引", numbers=int)
def xxColumnIndexToLetter(col_index):
    """将以数字表示的列索引转换为字母表示"""
    letter = ""
    while col_index > 0:
        remainder = col_index % 26
        if remainder == 0:
            remainder = 26
        letter += chr(remainder + 64)
        col_index = (col_index - remainder) // 26
    return letter[::-1]


@xw.func
@xw.arg("begin", doc=": 千字文起始句数。", numbers=int)
@xw.arg("total", doc=": 总句数。", numbers=int)
def xxQianZiWen(begin: int = 1, total: int = 125) -> str:
    """千字文字符串生成"""
    if begin == 1 and total == 125:
        return __kilos__
    s = begin if begin < 125 else 125
    div, mod = divmod(s + total - 1, 125)
    if div == 0:
        return __kilos__[(s - 1) * 10 : ((s - 1) + total) * 10]
    elif div == 1 and mod == 0:
        return __kilos__[(s - 1) * 10 :]
    else:
        interm = div * __kilos__
        s1 = __kilos__[(s - 1) * 10 :]
        s2 = __kilos__[: (mod) * 10]
        return s1 + interm + s2


@xw.func
@xw.arg("dms", doc=": 度分秒字符串。")
def xxDMS2DEC(dms: str) -> float:
    """将度分秒转换为十进制度数"""
    if not dms:
        return 0.0
    if not isinstance(dms, str):
        return 0.0
    if re.match(r"(-?\d+)°(\d+)'(\d+\.?\d*)\"", dms):
        x = re.split(r"""[°'"]""", dms)
        y = int(x[0])
        z = int(x[1])
        m = float(x[2])
        return y + z / 60 + m / 3600
    else:
        return 0.0


@xw.func
@xw.arg("field", doc=": 原字段名。")
def xxDLTBFieldsRename(field: str) -> str:
    """将国土调查数据库中的地类图斑图层的字段名转换为中文"""
    if not field:
        return ""
    if not isinstance(field, str):
        return ""
    fo = list(filter(lambda f: field == f, __DLTB_Fields__.keys()))
    if fo:
        return __DLTB_Fields__[fo[0]]
    else:
        return field

@xw.func
@xw.arg("field", doc=": 原字段名。")
def xxYZFieldsRename(field: str) -> str:
    """将研招数据库中的字母缩写形式的字段名转换为可读的中文名称"""
    if not field:
        return ""
    if not isinstance(field, str):
        return ""
    fo = list(filter(lambda f: field == f, __YZ_Fields__.keys()))
    if fo:
        return __YZ_Fields__[fo[0]]
    else:
        return field

@xw.func(call_in_wizard=False)
@xw.arg("path", doc=": DBF文件路径。")
@xw.arg("encoding", doc=": DBF文件编码，缺省为UTF8。")
def xxReadDBF(path: str, encoding: str = "UTF8") -> str:
    """读DBF文件"""
    table = DBF(path, encoding=encoding)
    return pd.DataFrame(iter(table))


@xw.func(call_in_wizard=False)
@xw.arg("table", pd.DataFrame, index=False, header=True)
@xw.arg("columns", doc=": 要相加的列名，多个列名用逗号分隔")
@xw.arg("condition_for_row", doc=": 用于选取某唯一行的条件")
@xw.ret(index=False, header=False)
def xxSumTableColumns(table, columns: str, condition_for_row: str = ""):
    """按指定的条件获取某行指定列的值之和"""
    _columns = re.split(r"""[,，]\s*""", columns)
    return table.query(condition_for_row)[_columns].sum(axis=1).iloc[0]


@xw.func(call_in_wizard=False)
@xw.arg("table", pd.DataFrame, index=False, header=True)
@xw.arg("columns", doc=": 要相加的列名，多个列名用逗号分隔")
@xw.arg("condition_for_row", doc=": 用于选取某唯一行的条件")
@xw.ret(index=False, header=False)
def xxSumTableColumnsAsMu(table, columns: str, condition_for_row: str = ""):
    """按指定的条件获取某行指定列的值之和"""
    _columns = re.split(r"""[,，]\s*""", columns)
    return table.query(condition_for_row)[_columns].sum(axis=1).iloc[0] * 3 / 2000


@xw.func
@xw.arg("tables", pd.DataFrame, index=False, header=True, doc=": 指定要合并的表(Table)")
@xw.ret(index=False, header=True, expand="table")
def xxConcatTables(*tables):
    """纵向合并多个表(Table)，保留所有表的列"""
    return pd.concat(tables, ignore_index=True, axis=0)


@xw.func
@xw.arg("sq_meters", doc=": 平方米数")
def xxSqMetersToMu(sq_meters: float) -> float:
    """将平方米转换为亩"""
    return sq_meters * 3.0 / 2000


@xw.func
@xw.arg("rng", ndim=2, doc=": 选定的范围(Range)")
@xw.ret(expand="table")
def xxFlatten(rng):
    """将二维数组转换为一维数组"""
    result = [cell for row in rng for cell in row]
    return [[cell] for cell in result]


@xw.func
@xw.arg("hanzi", doc=": 汉字")
def xxPinyinInitial(hanzi: str) -> str:
    """汉字拼音首字母"""
    return py.pinyin(hanzi, style=py.Style.INITIALS, strict=False)


@xw.func
@xw.arg("s", doc=": 英文字符串")
def xxCaptalizeEveryWord(s: str) -> str:
    """将英文字符串中的每个单词首字母大写"""
    return " ".join(word.capitalize() for word in s.split())


@xw.func
@xw.arg("s", doc=": 英文字符串")
def xxSnakeCase(s: str) -> str:
    """将英文字符串转换为蛇形"""
    return "".join(word.capitalize() for word in re.split(r"[\s_]", s))


# for debug
if __name__ == "__main__":
    xw.serve()
