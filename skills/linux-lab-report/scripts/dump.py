# -*- coding: utf-8 -*-
"""把实验报告 docx 的结构导成纯文本，用来核对生成结果。

用法:
    python dump.py 报告.docx [输出.txt]

输出每段的：全局序号、pPr 摘要（样式/编号/对齐/缩进）、每个 run 的 [颜色][加粗][斜体]文字。
表格单独列出表头和每行内容。中文字符在 Windows 终端里容易乱码，所以默认写文件而不打印。
"""
import io
import os
import sys

import docx
from docx.oxml.ns import qn


def w(t):
    return qn('w:' + t)


def para_info(p_el, doc):
    pPr = p_el.find(w('pPr'))
    bits = []
    if pPr is not None:
        ps = pPr.find(w('pStyle'))
        if ps is not None:
            bits.append('st=' + ps.get(w('val')))
        npr = pPr.find(w('numPr'))
        if npr is not None:
            ni = npr.find(w('numId'))
            if ni is not None:
                bits.append('num=' + ni.get(w('val')))
        jc = pPr.find(w('jc'))
        if jc is not None:
            bits.append('jc=' + jc.get(w('val')))
        ind = pPr.find(w('ind'))
        if ind is not None:
            for a in ('left', 'hanging', 'firstLineChars'):
                if ind.get(w(a)):
                    bits.append('%s=%s' % (a, ind.get(w(a))))
        sp = pPr.find(w('spacing'))
        if sp is not None and sp.get(w('line')):
            bits.append('line=' + sp.get(w('line')))
    runs = []
    for r in p_el.findall(w('r')):
        rPr = r.find(w('rPr'))
        col, bold, it, ea, sz = '-', '', '', '', ''
        if rPr is not None:
            c = rPr.find(w('color'))
            col = c.get(w('val')) if c is not None else '-'
            bold = 'B' if rPr.find(w('b')) is not None else ''
            it = 'I' if rPr.find(w('i')) is not None else ''
            f = rPr.find(w('rFonts'))
            ea = f.get(w('eastAsia')) if f is not None else ''
            s = rPr.find(w('sz'))
            sz = s.get(w('val')) if s is not None else ''
        t = ''.join(x.text or '' for x in r.findall(w('t')))
        tag = '%s%s%s%s%s' % (col, bold, it, '/s' + sz if sz else '', '/' + ea if ea else '')
        runs.append('[%s]%s' % (tag, t))
    body = ''.join(runs) if runs else '<empty>'
    return '%-34s :: %s' % (' '.join(bits), body)


def dump(path, out=None):
    doc = docx.Document(path)
    lines = []
    gi = [0]

    def walk(el, depth):
        for ch in el:
            if ch.tag == w('p'):
                p = docx.text.paragraph.Paragraph(ch, doc)
                lines.append('%*sP%-4d %s' % (depth * 2, '', gi[0], para_info(ch, doc)))
                gi[0] += 1
            elif ch.tag == w('tbl'):
                trs = ch.findall(w('tr'))
                ncols = len(trs[0].findall(w('tc'))) if trs else 0
                lines.append('%*s[TABLE rows=%d cols=%d]' % (depth * 2, '', len(trs), ncols))
                for tr in trs:
                    cells = []
                    for tc in tr.findall(w('tc')):
                        cells.append(''.join(x.text or '' for x in tc.findall('.//' + w('t'))))
                    lines.append('%*s  | %s' % (depth * 2, '', ' || '.join(cells)))
                for tr in trs:
                    for tc in tr.findall(w('tc')):
                        walk(tc, depth + 1)

    walk(doc.element.body, 0)
    text = '\n'.join(lines)
    if out:
        io.open(out, 'w', encoding='utf-8').write(text)
        return len(lines), gi[0], out
    return len(lines), gi[0], text


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(2)
    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) > 2 else os.path.splitext(src)[0] + '.dump.txt'
    n, gp, dst = dump(src, dst)
    print('lines=%d paragraphs=%d -> %s' % (n, gp, dst))
