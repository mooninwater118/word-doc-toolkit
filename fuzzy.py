# -*- coding: utf-8 -*-
"""模糊匹配核心：日期等价 + 空格容错

设计要点：
- 日期按「年月日齐全」识别，支持阿拉伯数字、两位年份、中文数字
- 归一化成 (年, 月, 日) 元组后比较是否相等
- 搜索词里的字符之间允许插入空白（含全角空格）
- 匹配结果给出文本索引，调用方负责换算成 Word Range 位置
"""
import re

# ---------- 中文数字 ----------
_CN_DIGIT = {
    '〇': 0, '零': 0, '○': 0,
    '一': 1, '二': 2, '三': 3, '四': 4,
    '五': 5, '六': 6, '七': 7, '八': 8, '九': 9,
}
_CN_YEAR_CHARS = '〇零○一二三四五六七八九'


def cn_to_int(s):
    """中文数字或阿拉伯数字 -> int。失败返回 None"""
    if not s:
        return None
    s = s.strip()
    if s.isdigit():
        return int(s)
    s = s.replace('○', '〇').replace('零', '〇')
    if '十' in s:
        head, _, tail = s.partition('十')
        if head and head not in _CN_DIGIT:
            return None
        if tail and tail not in _CN_DIGIT:
            return None
        left = _CN_DIGIT[head] if head else 1
        right = _CN_DIGIT[tail] if tail else 0
        return left * 10 + right
    if all(c in _CN_DIGIT for c in s):
        return int(''.join(str(_CN_DIGIT[c]) for c in s))
    return None


# ---------- 日期正则 ----------
# 年：4位数字 / 2位数字 / 2~4位中文数字
_YEAR = r'(?:\d{4}|\d{2}|[' + _CN_YEAR_CHARS + r']{2,4})'
# 月、日：1~2位数字 / 1~3位中文数字
_MD = r'(?:\d{1,2}|[一二三四五六七八九十]{1,3})'
# 分隔符：年/月 用汉字或 . - /，两侧可有空白
_SEP_Y = r'\s*[年.\-/]\s*'
_SEP_M = r'\s*[月.\-/]\s*'

# 非捕获版本的日期核心（供外部包裹捕获组用）
# 尾部用 (?:\s*[日号])? —— 只有真的跟了「日/号」才把前面的空白一起匹配，
# 否则不吃日期后面的空格（避免替换后空格消失）
DATE_CORE = (
    r'(?<!\d)' + _YEAR + _SEP_Y + _MD + _SEP_M + _MD +
    r'(?:\s*[日号])?(?!\d)'
)

# 带捕获组的版本（解析用）
_DATE_PARSE = re.compile(
    r'(?<!\d)(' + _YEAR + r')' + _SEP_Y +
    r'(' + _MD + r')' + _SEP_M +
    r'(' + _MD + r')(?:\s*[日号])?(?!\d)'
)

DATE_RE = re.compile(DATE_CORE)


def parse_date(s):
    """从字符串里识别一个日期，返回 (年, 月, 日)；不是合法日期返回 None"""
    if not s:
        return None
    m = _DATE_PARSE.search(s)
    if not m:
        return None
    y = cn_to_int(m.group(1))
    mo = cn_to_int(m.group(2))
    d = cn_to_int(m.group(3))
    if y is None or mo is None or d is None:
        return None
    if y < 100:          # 两位年份按 20xx 处理
        y += 2000
    if not (1 <= mo <= 12):
        return None
    if not (1 <= d <= 31):
        return None
    return (y, mo, d)


# ---------- 空格容错 ----------
_SPACE_GAP = r'[\s\u3000]*'


def flex(text):
    """把一串字面量转成「字符之间允许任意空白」的正则片段"""
    if not text:
        return ''
    return _SPACE_GAP.join(re.escape(c) for c in text)


# ---------- 构建匹配模式 ----------
class FuzzyPattern:
    """模糊模式：封装正则 + 每个日期组应有的归一值"""

    def __init__(self, keyword, fuzzy_date=True, fuzzy_space=True):
        # 首尾空格不参与匹配：容错只发生在字符「之间」，不扩展到两端
        keyword = (keyword or '').strip()
        self.keyword = keyword
        self.dates = []
        parts = []
        pos = 0
        if fuzzy_date:
            for m in _DATE_PARSE.finditer(keyword):
                d = parse_date(m.group(0))
                seg = keyword[pos:m.start()]
                parts.append(flex(seg) if fuzzy_space else re.escape(seg))
                if d is None:
                    # 形似日期但不合法（如 2025年13月40日）：退回字面匹配。
                    # 否则 expect 会是 None，而匹配到的非法日期也是 None，
                    # 两者相等就放行了（检查报告 S16）。
                    parts.append(flex(m.group(0)) if fuzzy_space else re.escape(m.group(0)))
                else:
                    parts.append('(' + DATE_CORE + ')')
                    self.dates.append(d)
                pos = m.end()
        tail = keyword[pos:]
        parts.append(flex(tail) if fuzzy_space else re.escape(tail))
        self.regex = re.compile(''.join(parts))
        self.space = fuzzy_space
        self.date = fuzzy_date

    def finditer(self, text):
        """返回 [(start, end)]，已做跨段落过滤和日期等价校验"""
        out = []
        for m in self.regex.finditer(text):
            seg = m.group(0)
            if '\r' in seg or '\x07' in seg or '\x0b' in seg or '\x0c' in seg:
                continue
            ok = True
            for i, expect in enumerate(self.dates):
                got = parse_date(m.group(i + 1))
                if got != expect:
                    ok = False
                    break
            if ok:
                out.append((m.start(), m.end()))
        return out


# ---------- 位置换算 ----------
# Word 的 Content.Text 里，这些字符在字符串中占位但在 Range 位置里不占
# 目前实测确认的只有单元格标记 \x07。若以后发现 \x01（图片）、\x13-\x15（域代码）
# 也有同样行为，往这个集合里加即可。
NO_POS_CHARS = {'\x07'}


def build_pos_map(text):
    """返回数组 pm，pm[i] = 文本索引 i 对应的 Range 偏移"""
    n = len(text)
    pm = [0] * (n + 1)
    rp = 0
    for i in range(n):
        pm[i] = rp
        if text[i] not in NO_POS_CHARS:
            rp += 1
    pm[n] = rp
    return pm


def diagnose_chars(text):
    """统计文本中的不可见字符，用于排查哪些字符可能导致位置偏差"""
    from collections import Counter
    c = Counter(ch for ch in text if ord(ch) < 32 or ch == '\x07')
    return {('0x%02x' % ord(k)): v for k, v in sorted(c.items())}
