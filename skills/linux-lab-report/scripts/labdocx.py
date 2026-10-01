# -*- coding: utf-8 -*-
"""labdocx —— 《Linux操作系统配置与管理》实验报告的生成工具。

设计前提：报告的「外壳」（外层 4 行模板表格、页眉、样式表、numbering.xml）
必须原样保留，只替换正文。所以这里不新建 docx，而是**在 assets/template.docx
上做原位改写**：拿到每个段落的 XML 元素，改文字、改颜色、克隆、删除、插入。

核心思路：模板里每种元素都已经有一个「活样板」——站点头、编号任务、项目符号
命令行、蓝色提示、红色注意、灰色截图要求、绿色占位符、空行。生成内容时不手写
pPr，而是**深拷贝对应的样板再换文字**，这样编号、缩进、行距、字体全部自然带上，
不会出现「颜色对了但编号丢了」这类问题。

用法见 SKILL.md。
"""
import copy
import os
import sys

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_TEMPLATE = os.path.join(os.path.dirname(HERE), 'assets', 'template.docx')

# ---------------------------------------------------------------- 颜色语言
K = '000000'    # 正文
G = '00B050'    # 命令
Y = '92D050'    # 截图占位符
GY = '808080'   # 截图要求（灰色斜体）
BL = '0070C0'   # 【提示】【正确】——蓝色
RD = 'C00000'   # 【注意】【错误】——红色

# w:rPr 的子元素有固定顺序，w:color 必须插在正确位置，否则 Word 会报文档损坏。
RPR_ORDER = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike',
             'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid',
             'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz',
             'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign',
             'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath']

# 样板段落编号（照 references/skeleton.md，从 assets/template.docx 数出来的）
# 2026-10-01 模板被重做过一次：占位站里补成了完整的一套样板（站点头 + 任务 +
# 3 条命令 + 【正确】+【提示】+【注意】+ 截图要求 + 占位符 + 空行），所以它后面
# 的所有编号整体前移。前移量在速查卡那段是 2、之后是 3。改模板后一定要重新
# 跑一遍 dump.py 核对这张表。
TPL_INDEX = {
    'title': 1,        # 大标题第二段（sz36 华文中宋）
    'num': 6,          # 「实验报告 N」
    'subtitle': 7,     # 项目名
    'goal': 29,        # 实验目的条目（style a9 + numId 5）
    'env': 34,         # 实验设备条目（style a9 + numId 6）
    'intro': 42,       # 实验步骤引言（首行缩进 2 字）
    'card_header': 43,  # 【开工前…速查卡】居中
    'cmd_header': 63,  # 【本次实验常用命令】居中
    'stat': 96,        # 站点标题（华文中宋加粗）
    'task': 97,        # 编号任务（numId 3）
    'cmd': 98,         # 命令行（style a9 + numId 4 项目符号）
    'note': 100,       # 蓝色提示（左缩进 420）
    'warn': 103,       # 红色注意（左缩进 420）
    'shot': 104,       # 灰色斜体截图要求（左缩进 420）
    'ph': 105,         # 绿色斜体占位符（居中）
    'blank': 106,      # 空行（居中）
    'err_head': 108,   # 「附：本次实验常见报错速查」
    'ask_head': 130,   # 实验体会说明（灰色）
    'ask_q': 131,      # 实验体会问题（蓝色）
}
# 报错速查表锚点：站点内容插在它前面
ERR_ANCHOR = 108
# 模板里的占位站（会被生成的内容整段替换掉）
DUMMY_SITES = (96, 107)   # 含首含尾

# 三张嵌套表：速查卡 / 常用命令 / 报错速查
TABLE_INDEX = {'card': 0, 'cmdlist': 1, 'errors': 2}


# ================================================================ 底层工具
def ensure_child(rPr, tag):
    """在 rPr 里取到 <w:tag>，没有就按 RPR_ORDER 的次序插一个进去。

    w:rPr 的子元素顺序是 schema 定死的，插错位置 Word 会说文档损坏。所以不能
    简单地 append，得找到第一个「排在它后面」的元素，插在它前面。
    """
    el = rPr.find(qn('w:' + tag))
    if el is not None:
        return el
    el = OxmlElement('w:' + tag)
    target = RPR_ORDER.index(tag)
    for ch in rPr:
        name = ch.tag.split('}')[-1]
        if name in RPR_ORDER and RPR_ORDER.index(name) > target:
            ch.addprevious(el)
            return el
    rPr.append(el)
    return el


def set_color(rPr, color):
    """设置 w:color。color 为 None 表示「不动」——保留样板本身的颜色。

    这样 clone 的语义就是「照抄这个样板，只换文字」，不会因为传 None
    而把样板里本来就对的颜色抹掉。
    """
    if color is None:
        return
    ensure_child(rPr, 'color').set(qn('w:val'), color)


# 三张表格（速查卡 / 常用命令 / 报错速查）统一的字体：中文仿宋、英文 Consolas、
# 五号（sz 以半磅为单位，21 就是 10.5 磅）。
TBL_EA = '仿宋'
TBL_LATIN = 'Consolas'
TBL_SZ = '21'


def set_font(rPr, ea=TBL_EA, latin=TBL_LATIN, sz=TBL_SZ):
    """把 run 的字体钉死成「中文仿宋 + 英文 Consolas + 五号」。

    为什么要专门钉一遍：表头在模板里本来就写好了这几个属性，但**数据行是空段落**。
    set_cell 往空段落里写文字时找不到任何 run 可以当样板，只能造一个空的 rPr，
    于是字体全靠继承——ascii 从段落标记那儿拿到 Consolas，可 eastAsia 谁都没写，
    一路掉回 docDefaults 里的宋体。结果是同一张表里表头仿宋、数据行宋体，
    放大一看就不齐。这里不赌继承，直接写在 run 上。

    只动字体和字号，加粗不动——表头该粗还是粗。
    """
    f = ensure_child(rPr, 'rFonts')
    f.set(qn('w:ascii'), latin)
    f.set(qn('w:hAnsi'), latin)
    f.set(qn('w:cs'), latin)
    f.set(qn('w:eastAsia'), ea)
    for tag in ('sz', 'szCs'):
        ensure_child(rPr, tag).set(qn('w:val'), sz)
    return rPr


def normalize_table_font(tbl):
    """把一张表里每个单元格的每个 run 都过一遍 set_font。

    表头和数据行一起过，所以整张表是同一个字体，不依赖模板当时写成什么样。
    """
    for tr in tbl.findall(qn('w:tr')):
        for tc in tr.findall(qn('w:tc')):
            for p in tc.findall(qn('w:p')):
                for r in p.findall(qn('w:r')):
                    rPr = r.find(qn('w:rPr'))
                    if rPr is None:
                        # rPr 必须是 w:r 的第一个子元素，排在 w:t 前面
                        rPr = OxmlElement('w:rPr')
                        r.insert(0, rPr)
                    set_font(rPr)


def rewrite(p, segments, run_idx=0):
    """把段落内容换成 [(文本, 颜色 or None), ...]，沿用第 run_idx 个 run 的字体。

    颜色为 None 表示沿用样板原来的颜色（不是取消着色）。
    run_idx=0 用第一个 run 当模板；多 run 段落里各 run 格式可能不同，取错会丢格式。
    """
    runs = p.findall(qn('w:r'))
    rpr = None
    if runs:
        tpl = runs[min(run_idx, len(runs) - 1)]
        rpr = tpl.find(qn('w:rPr'))
    for r in runs:
        p.remove(r)
    for text, color in segments:
        r = OxmlElement('w:r')
        newrpr = copy.deepcopy(rpr) if rpr is not None else OxmlElement('w:rPr')
        set_color(newrpr, color)
        r.append(newrpr)
        t = OxmlElement('w:t')
        t.text = text
        t.set(qn('xml:space'), 'preserve')
        r.append(t)
        p.append(r)
    return p


def settext(p, text, idx=-1):
    """只改文字，保留原 run 的字体。

    多 run 段落里各 run 的 rPr 常常不一样（比如「（」「1」「）」三个 run），
    这时 rewrite(run_idx=0) 会取到「（」的格式而丢掉加粗。settext 直接留下
    第 idx 个 run（默认最后一个），格式最稳。小标题、抬头用这个。
    """
    runs = p.findall(qn('w:r'))
    if not runs:
        r = OxmlElement('w:r')
        p.append(r)
        runs = [r]
    keep = runs[idx] if idx >= 0 else runs[-1]
    for r in runs:
        if r is not keep:
            p.remove(r)
    for tag in ('w:t', 'w:br', 'w:tab'):
        for el in keep.findall(qn(tag)):
            keep.remove(el)
    t = OxmlElement('w:t')
    t.text = text
    t.set(qn('xml:space'), 'preserve')
    keep.append(t)
    return p


def clone(tpl_p, segments, run_idx=0):
    """深拷贝一个段落样板，再把内容换成 segments。"""
    p = copy.deepcopy(tpl_p)
    rewrite(p, segments, run_idx)
    return p


def ptext(p):
    return ''.join(t.text or '' for t in p.findall('.//' + qn('w:t')))


def build_pmap(doc):
    """按「body → tbl → tr → tc → p」顺序给每个 w:p 编号，返回 {序号: 元素}。

    这个顺序和 skeleton.md 一致，所以可以照着那张表取样板段落。
    """
    pmap = {}
    gi = [0]

    def walk(el):
        for child in el:
            if child.tag == qn('w:p'):
                pmap[gi[0]] = child
                gi[0] += 1
            elif child.tag == qn('w:tbl'):
                for tr in child.findall(qn('w:tr')):
                    for tcx in tr.findall(qn('w:tc')):
                        walk(tcx)

    walk(doc.element.body)
    return pmap


def tbl_rows(tbl):
    return tbl.findall(qn('w:tr'))


def set_cell(tc, text, color=None):
    """改写表格单元格文字（只留第一个段落，沿用其字体）。"""
    ps = tc.findall(qn('w:p'))
    for extra in ps[1:]:
        tc.remove(extra)
    rewrite(ps[0], [(text, color)])


def row_cells(tr):
    return tr.findall(qn('w:tc'))


def list_sig(p):
    """段落的「块签名」：决定它属于哪个多段落块的全部 pPr 特征之和。

    用来判断「下一个段落还是不是这一款」。签名必须够宽，否则认不出边界：
    实验目的/设备靠 pStyle + numId 区分，而实验体会的问题只有 ind left=420，
    既没有 pStyle 也没有 numPr——只比这两个的话它们签名都是空的，等于没有边界，
    就会一路吃过界把「教师签名」那行改掉。
    """
    pPr = p.find(qn('w:pPr'))
    if pPr is None:
        return None
    bits = []
    ps = pPr.find(qn('w:pStyle'))
    if ps is not None:
        bits.append('style=' + str(ps.get(qn('w:val'))))
    npr = pPr.find(qn('w:numPr'))
    if npr is not None:
        nid = npr.find(qn('w:numId'))
        ilv = npr.find(qn('w:ilvl'))
        bits.append('num=' + (nid.get(qn('w:val')) if nid is not None else '?'))
        bits.append('ilvl=' + (ilv.get(qn('w:val')) if ilv is not None else '?'))
    ind = pPr.find(qn('w:ind'))
    if ind is not None:
        for a in ('left', 'hanging', 'firstLine', 'firstLineChars'):
            v = ind.get(qn('w:' + a))
            if v:
                bits.append('ind.%s=%s' % (a, v))
    jc = pPr.find(qn('w:jc'))
    if jc is not None:
        bits.append('jc=' + str(jc.get(qn('w:val'))))
    return '|'.join(bits) if bits else None


def siblings_from(el, n):
    """把 el 所在的「同款段落块」伸缩到 n 条，返回这 n 个元素。

    三件事，缺一不可：

    1. **按签名收边**。只吃和 el 同款（同 pStyle / 同编号 / 同缩进）的后续段落。
       多吃了就会出事——实验目的在模板里只有 3 个槽位，写第 4、5 条时如果一路
       沿着 getnext() 取，取到的就是空行和「二、实验设备及环境」标题，把标题
       改成实验目的、连它的华文中宋一起丢掉。
    2. **把同款的多余段落删掉**，不只是 `out` 里那几个。模板给了 3 个目的槽位，
       只写 2 条时必须把第 3 条删掉，否则文档里会留着
       「【此处填写实验目的第 3 条…】」这行占位文字交到学生手上。
       所以这里一路收到底，而不是收够 n 就停。
    3. **不够就克隆**最后一条补上（克隆继承 numId，编号自然接得上）。
    """
    sig = list_sig(el)
    out = [el]
    cur = el
    while True:
        nxt = cur.getnext()
        while nxt is not None and nxt.tag != qn('w:p'):
            nxt = nxt.getnext()
        if nxt is None:
            break
        if sig is None or list_sig(nxt) != sig:
            break          # 走出这个块了
        out.append(nxt)
        cur = nxt
    while len(out) > n:
        el_to_drop = out.pop()
        el_to_drop.getparent().remove(el_to_drop)
    while len(out) < n:
        newp = copy.deepcopy(out[-1])
        out[-1].addnext(newp)
        out.append(newp)
    return out


def fit_rows(tbl, n):
    """把表格的数据行数（不含表头）调整到 n，返回全部行。"""
    rows = tbl_rows(tbl)
    while len(rows) - 1 < n:
        tbl.append(copy.deepcopy(rows[-1]))
        rows = tbl_rows(tbl)
    while len(rows) - 1 > n:
        rows[-1].getparent().remove(rows[-1])
        rows = tbl_rows(tbl)
    return tbl_rows(tbl)


def ins_before(ref_el, items):
    for el in items:
        ref_el.addprevious(el)


# ================================================================ body 积木
# 站点正文用这些函数拼，比裸写 (kind, color, segments) 可读得多。

def cmd(syntax, what=''):
    """命令行：'输入 <绿色命令> —— 这条命令在做什么'"""
    segs = [('输入 ', K), (' ' + syntax + ' ', G)]
    if what:
        segs += [('—— ', K), (what, K)]
    return ('cmd', None, segs)


def ok(text):
    """【正确】——按下去该看到什么。给学生一个可核对的结果，他们能自己判断对错。"""
    return ('note', BL, [('【正确】', BL), (text, BL)])


def win(text):
    """【提示】Windows 对照——把新概念挂到学生已经会的东西上。"""
    return ('note', BL, [('【提示】', BL), ('Windows ', BL), ('对照：', BL), (text, BL)])


def tip(text):
    return ('note', BL, [('【提示】', BL), (text, BL)])


def warn(text):
    """【注意】——最容易踩的坑。提前讲掉，比课后一个个答疑省时间。"""
    return ('note', RD, [('【注意】', RD), (text, RD)])


def wrong(text):
    return ('note', RD, [('【错误】', RD), (text, RD)])


def shot(text):
    """截图要求——写清「画面里必须能看到什么」，截图本身就成了自检动作。"""
    return ('shot', None, [('截图要求：' + text, GY)])


def gap():
    return ('blank', None, [('', None)])


# ================================================================ 主流程
def verify_template(doc):
    """核对 TPL_INDEX 和当前 template.docx 是否还对得上。

    模板一旦被改动（加一段、删一行），它后面所有段落的编号就整体错位。而错位的
    后果是**静默**的：不报错、能生成、文档也打得开，只是占位站没被删掉、站点标题
    整批丢失、内容插进了错误的位置。2026-10-01 模板重做时就踩过这一回——五份报告
    同时生成坏文档，靠 grader 才查出来。

    所以开工前先验几个**内容固定**的地标（章节标题、报错表标题、体会引导语、
    截图要求那几个带固定颜色的）。对不上就直接抛错，让人去跑 dump.py 重新核对
    references/skeleton.md 里的那张表，而不是把一份坏文档交出去。
    """
    pmap = build_pmap(doc)

    def text_at(idx):
        el = pmap.get(idx)
        return ptext(el) if el is not None else None

    def color_at(idx):
        el = pmap.get(idx)
        if el is None:
            return None
        for r in el.findall(qn('w:r')):
            rPr = r.find(qn('w:rPr'))
            if rPr is not None:
                c = rPr.find(qn('w:color'))
                if c is not None:
                    return c.get(qn('w:val'))
        return None

    # (编号, 应该含有的字样, 说明)  —— 这些字是模板自带的固定文案，不会随实验变
    landmarks = [
        (28, '实验目的', '一、实验目的'),
        (33, '实验设备', '二、实验设备及环境'),
        (41, '实验步骤', '三、实验步骤'),
        (TPL_INDEX['card_header'], '速查卡', '速查卡标题'),
        (TPL_INDEX['cmd_header'], '常用命令', '常用命令表标题'),
        (TPL_INDEX['stat'], '站点标题', '站点占位标题'),
        (TPL_INDEX['task'], '这一步要做什么', '任务占位句'),
        (TPL_INDEX['warn'], '最容易踩的坑', '【注意】样板'),
        (TPL_INDEX['ph'], '截图占位符', '【截图占位符】样板'),
        (TPL_INDEX['err_head'], '报错速查', '附：本次实验常见报错速查'),
        (TPL_INDEX['ask_head'], '请围绕', '实验体会引导语'),
        (TPL_INDEX['ask_q'], '问题', '体会问题样板'),
    ]
    bad = []
    for idx, needle, label in landmarks:
        t = text_at(idx)
        if t is None or needle not in t:
            bad.append('%s（应含「%s」）：编号 %d 实际是 %r'
                       % (label, needle, idx, (t or '')[:30]))

    # 截图要求那条还必须带灰斜体，【截图占位符】必须带浅绿——它们靠颜色区分，
    # 光看文字是分别不出来的。
    shot_idx = TPL_INDEX['shot']
    t = text_at(shot_idx)
    if t is None or '截图要求' not in t:
        bad.append('截图要求样板：编号 %d 实际是 %r' % (shot_idx, (t or '')[:30]))
    elif color_at(shot_idx) not in ('808080', None):
        bad.append('截图要求样板的颜色不对：%s' % color_at(shot_idx))
    if color_at(TPL_INDEX['ph']) not in ('92D050', None):
        bad.append('【截图占位符】样板的颜色不对：%s' % color_at(TPL_INDEX['ph']))

    if bad:
        raise ValueError(
            'TPL_INDEX 和当前 assets/template.docx 对不上了——模板多半被改过，'
            '段落编号整体错位。跑一遍 `python scripts/dump.py assets/template.docx '
            'skeleton.txt`，照着 references/skeleton.md 的对照表重新核一遍编号。\n对不上的地标：\n  '
            + '\n  '.join(bad))


def check_marks(spec):
    """扫一遍 spec 里所有会被写进正文的文字，找出 markdown 记号就报错。

    为什么要在生成**之前**拦一道：spec 是写成 Python 字符串的，写的时候手会
    习惯性敲 **加粗**、`代码`、## 标题。这些记号在 Word 里不会变成格式，而是
    原样打出来，变成正文中间突兀的星号、井号、反引号。

    这个错最恶劣的地方是**它不报错、也看不出来**：文档能正常打开，颜色、编号、
    站点全对，只有把正文从头读一遍才会发现。所以别指望记得住——我在这个 skill
    的五份报告里连犯过三次，所以让脚本在生成时直接拦住。

    文档字符串本身和以 # 开头的注释行不扫，那里面写 markdown 是正常的。
    扫的只是会进入文档的字段值。

    顺带也扫控制字符——这一类是 2026-10-01 写 Shell 脚本那一讲时撞上的。
    spec 里要写终端颜色码，很自然会敲 '\\033[0;32m'；可在 Python 字符串里
    单个反斜杠的 \\033 是**转义序列**，它不会变成四个字符，而是变成一个真的
    ESC 字符（0x1B）。这种字符进不了 XML，python-docx 保存时直接抛
    「All strings must be XML compatible: ... no NULL bytes or control
    characters」——错在很深的地方，既不说是哪个字段，也看不出跟颜色码有关。
    写成 '\\\\033' 才是字面的反斜杠加 033。这里提前拦一下，把话说清楚。
    """
    bad = []
    ctrl = []

    def scan(value, where):
        if isinstance(value, str):
            if '**' in value or '##' in value or '`' in value:
                bad.append('%s: %s' % (where, value.strip()[:50]))
            for ch in value:
                if ord(ch) < 32 and ch not in '\t\n':
                    ctrl.append('%s: 含控制字符 %r（U+%04X）'
                                % (where, ch, ord(ch)))
                    break
        elif isinstance(value, (list, tuple)):
            for i, v in enumerate(value):
                scan(v, '%s[%d]' % (where, i))
        elif isinstance(value, dict):
            for k, v in value.items():
                scan(v, '%s.%s' % (where, k))

    for key, val in spec.items():
        scan(val, key)

    if bad:
        raise ValueError(
            'spec 里有 markdown 记号（** ## `），这些会原样打进 Word 变成正文里的'
            '星号/井号/反引号。删掉它们再用中文书名号或引号表达强调：\n  ' +
            '\n  '.join(bad[:12]))

    if ctrl:
        raise ValueError(
            'spec 里有控制字符，它们进不了 XML，保存时会报一句看不懂的 lxml 错误。'
            '最常见的原因是写了 \\033 这类终端颜色码：在 Python 字符串里，单个反斜杠的'
            '\\033 会被解释成一个真的 ESC 字符，要写成 \\\\033 才能得到字面的反斜杠加 033。'
            '（数据里如果确实需要真换行，用 \\n 是允许的。）出问题的字段：\n  ' +
            '\n  '.join(ctrl[:12]))


def build_report(spec, out_path, template=DEFAULT_TEMPLATE):
    """按 spec 生成一份实验报告，返回输出路径。spec 字段见 SKILL.md。"""
    check_marks(spec)
    doc = docx.Document(template)
    verify_template(doc)
    pmap = build_pmap(doc)
    T = {k: pmap[v] for k, v in TPL_INDEX.items()}

    outer = doc.tables[0]
    body_tc = outer.rows[3].cells[0]._tc
    tabs = [el for el in body_tc if el.tag == qn('w:tbl')]
    TBL = {k: tabs[v] for k, v in TABLE_INDEX.items()}

    # ---- 抬头 ----
    title = spec.get('title')
    if title:
        settext(T['title'], 'Linux' + title)
        rewrite(T['subtitle'], [(title, K)])
    if spec.get('num') is not None:
        rewrite(T['num'], [('实验报告 ', K), (str(spec['num']), K)])

    # ---- 一、实验目的 ----
    goals = spec.get('goals')
    if goals:
        for p, g in zip(siblings_from(T['goal'], len(goals)), goals):
            rewrite(p, [(g, K)])

    # ---- 二、实验设备及环境 ----
    env = spec.get('env')
    if env:
        for p, e in zip(siblings_from(T['env'], len(env)), env):
            rewrite(p, [(e, K)])

    # ---- 实验步骤引言 ----
    if spec.get('intro'):
        rewrite(T['intro'], [(spec['intro'], K)])

    # ---- 速查卡 ----
    if spec.get('card_header'):
        settext(T['card_header'], spec['card_header'])
    card = spec.get('card')
    if card is not None:
        for tr, row in zip(fit_rows(TBL['card'], len(card))[1:], card):
            for tc, v in zip(row_cells(tr), row):
                set_cell(tc, v)

    # ---- 常用命令表 ----
    if spec.get('cmd_header'):
        settext(T['cmd_header'], spec['cmd_header'])
    cl = spec.get('cmdlist')
    if cl is not None:
        for tr, row in zip(fit_rows(TBL['cmdlist'], len(cl))[1:], cl):
            for tc, v in zip(row_cells(tr), row):
                set_cell(tc, v)

    # ---- 站点 ----
    anchor = pmap[ERR_ANCHOR]
    inserted = []
    for site in spec.get('sites', []):
        inserted.append(clone(T['blank'], [('', None)]))
        inserted.append(clone(T['stat'], [(site['name'], None)]))
        for task in site['tasks']:
            inserted.append(clone(T['task'], [(task['text'], K)]))
            for kind, color, segs in task.get('body', []):
                tpl = T['cmd'] if kind == 'cmd' else \
                      T['shot'] if kind == 'shot' else \
                      T['blank'] if kind == 'blank' else \
                      T['warn'] if color == RD else T['note']
                inserted.append(clone(tpl, segs))
                if kind == 'shot':
                    # 每条截图要求下面紧跟一个占位符，学生照着往里贴图。
                    # 写在脚本里而不是让调用方手写，是为了不漏——漏一个，
                    # 学生就少交一张图，而且没人会发现。
                    inserted.append(clone(T['ph'], [('【截图占位符】', Y)]))
    for n in range(DUMMY_SITES[0], DUMMY_SITES[1] + 1):
        el = pmap[n]
        el.getparent().remove(el)
    ins_before(anchor, inserted)

    # ---- 报错速查表 ----
    errors = spec.get('errors')
    if errors is not None:
        for tr, row in zip(fit_rows(TBL['errors'], len(errors))[1:], errors):
            for tc, v in zip(row_cells(tr), row):
                set_cell(tc, v)

    # ---- 四、实验体会 ----
    qs = spec.get('questions')
    if qs:
        # 模板那句引导语写死了「下面两个问题」。题目数不是 2 时必须改，否则等于
        # 告诉学生「还有题你没看到」。这个问题光扫结构查不出来——引导语和问题各自
        # 都长得正常，只有把两边数一数才看得出对不上。
        settext(T['ask_head'],
                '（请围绕下面%d个问题展开，不必写成模板套话）' % len(qs))
        for p, q in zip(siblings_from(T['ask_q'], len(qs)), qs):
            rewrite(p, [(q, BL)])

    # ---- 三张表的字体统一 ----
    # 放在最后、三张表都填完之后跑：这时候 fit_rows 已经把数据行补齐了，
    # 表头和数据行能一起过一遍，整张表字体一致。
    for tbl in TBL.values():
        normalize_table_font(tbl)

    out_path = os.path.abspath(out_path)
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    doc.save(out_path)
    return out_path


# ================================================================ CLI
if __name__ == '__main__':
    if len(sys.argv) < 3:
        print('用法: python labdocx.py <spec.py> <out.docx>')
        raise SystemExit(2)
    # spec 文件里会写 `from labdocx import cmd, ok, ...`，所以要把本脚本所在目录
    # 放进 sys.path——否则在别的 cwd 下执行会 ImportError。
    if HERE not in sys.path:
        sys.path.insert(0, HERE)
    ns = {}
    with open(sys.argv[1], encoding='utf-8') as f:
        exec(f.read(), ns)
    print(build_report(ns['SPEC'], sys.argv[2]))
