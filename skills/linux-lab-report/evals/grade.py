# -*- coding: utf-8 -*-
"""程序化核对实验报告是否符合 linux-lab-report 规范。

用法:
    python grade.py <报告.docx> [输出.json]

把每一项断言的结果写成 {"text": ..., "passed": bool, "evidence": ...}。
"""
import io
import json
import os
import re
import sys

import docx
from docx.oxml.ns import qn


def w(t):
    return qn('w:' + t)


# emoji 常用区段。注意不要包含 U+2190-U+21FF（含 √ U+221A）——那是模板表格里
# 的勾选符号，不是 emoji。
EMOJI = re.compile(
    '[\U0001F300-\U0001FAFF\U0001F000-\U0001F2FF'
    '\U00002600-\U000026FF\U00002700-\U000027BF\U0000FE0F]')


class Report:
    def __init__(self, path):
        self.path = path
        self.doc = docx.Document(path)
        self.paras = []          # (p_el, text, [(color,attrs,text), ...], pPr摘要)
        self.tables = []         # (nested? rows, header, rows_text)
        self._scan()
        self.all_text = '\n'.join(p['text'] for p in self.paras)

    def _scan(self):
        doc = self.doc

        def ppr_bits(p_el):
            pPr = p_el.find(w('pPr'))
            b = {}
            if pPr is not None:
                npr = pPr.find(w('numPr'))
                if npr is not None and npr.find(w('numId')) is not None:
                    b['numId'] = npr.find(w('numId')).get(w('val'))
                jc = pPr.find(w('jc'))
                if jc is not None:
                    b['jc'] = jc.get(w('val'))
                ps = pPr.find(w('pStyle'))
                if ps is not None:
                    b['style'] = ps.get(w('val'))
                ind = pPr.find(w('ind'))
                if ind is not None and ind.get(w('left')):
                    b['left'] = ind.get(w('left'))
            return b

        def walk(el):
            for ch in el:
                if ch.tag == w('p'):
                    runs = []
                    for r in ch.findall(w('r')):
                        rPr = r.find(w('rPr'))
                        col = it = bold = ea = None
                        if rPr is not None:
                            c = rPr.find(w('color'))
                            col = c.get(w('val')) if c is not None else None
                            it = rPr.find(w('i')) is not None
                            bold = rPr.find(w('b')) is not None
                            f = rPr.find(w('rFonts'))
                            ea = f.get(w('eastAsia')) if f is not None else None
                        t = ''.join(x.text or '' for x in r.findall(w('t')))
                        runs.append({'color': col, 'italic': it, 'bold': bold,
                                     'ea': ea, 'text': t})
                    txt = ''.join(r['text'] for r in runs)
                    self.paras.append({'el': ch, 'text': txt, 'runs': runs,
                                       'pPr': ppr_bits(ch)})
                elif ch.tag == w('tbl'):
                    trs = ch.findall(w('tr'))
                    grid = []
                    for tr in trs:
                        grid.append([''.join(x.text or '' for x in tc.findall('.//' + w('t')))
                                     for tc in tr.findall(w('tc'))])
                    ncols = len(grid[0]) if grid else 0
                    self.tables.append({'el': ch, 'rows': len(trs), 'cols': ncols, 'grid': grid})
                    for tr in trs:
                        for tc in tr.findall(w('tc')):
                            walk(tc)

        walk(doc.element.body)

    # ---- 便捷查询 ----
    def count(self, needle):
        return self.all_text.count(needle)

    def site_titles(self):
        """站点标题：中文括号编号开头、并且是加粗华文中宋的段落。"""
        out = []
        for i, p in enumerate(self.paras):
            m = re.match(r'^（\s*(\d+)\s*）(.+)$', p['text'].strip())
            if m and any(r['bold'] for r in p['runs']):
                out.append((i, int(m.group(1)), p['text'].strip()))
        return out

    def runs_matching(self, pred):
        out = []
        for p in self.paras:
            for r in p['runs']:
                if r['text'].strip() and pred(r):
                    out.append(r['text'].strip())
        return out

    def outer_table(self):
        """外层模板表格：body 下的第一个表格。"""
        return self.tables[0] if self.tables else None


def grade(path):
    rep = Report(path)
    res = []

    def add(text, passed, evidence=''):
        res.append({'text': text, 'passed': bool(passed), 'evidence': str(evidence)[:300]})

    # 1. 外层模板表格
    ot = rep.outer_table()
    ok1 = ot is not None and ot['rows'] == 4
    add('外层模板表格保留（4 行，正文在跨列单元格内）', ok1,
        'rows=%s cols=%s' % (ot['rows'], ot['cols']) if ot else '没有表格')

    # 2. 页眉
    hdr_text = ''
    for s in rep.doc.sections:
        for p in s.header.paragraphs:
            hdr_text += p.text
    add('页眉《Linux操作系统配置与管理》实验报告书保留',
        '实验报告书' in hdr_text or 'Linux操作系统配置与管理' in hdr_text, hdr_text[:80])

    # 3. 站点编号用中文括号且连续
    sites = rep.site_titles()
    nums = [n for _, n, _ in sites]
    add('站点标题用中文括号编号且从 1 连续',
        len(sites) >= 3 and nums == list(range(1, len(nums) + 1)),
        'site nums=%s' % nums)

    # 4. 站点标题为华文中宋加粗
    ea_ok = all(any(r['ea'] == '华文中宋' for r in rep.paras[i]['runs'])
                for i, _, _ in sites) if sites else False
    add('站点标题字体为华文中宋加粗', ea_ok and len(sites) >= 3,
        'sites=%d' % len(sites))

    # 5. 【正确】自检 —— 每个站点至少一条
    n_ok = rep.count('【正确】')
    add('【正确】自检行数量不少于站点数', n_ok >= len(sites) and n_ok >= 3,
        '【正确】x%d, 站点 x%d' % (n_ok, len(sites)))

    # 6. Windows 对照 —— 清理现场这类纯收尾站点本来就没有 Windows 对应物，
    #    所以要求「站点数 - 1」，留一个合理余量。
    n_win = len(re.findall(r'Windows\s*对照', rep.all_text))
    add('每条命令给出【提示】Windows 对照', n_win >= max(1, len(sites) - 1),
        'Windows 对照 x%d, 站点 x%d' % (n_win, len(sites)))

    # 7. 截图要求 + 占位符配对
    n_shot = rep.count('截图要求')
    n_ph = rep.count('【截图占位符】')
    add('每条截图要求都配有【截图占位符】',
        n_shot >= 1 and n_shot == n_ph, '截图要求 x%d, 占位符 x%d' % (n_shot, n_ph))

    # 8. 三张表（速查卡 / 常用命令 / 报错速查）
    has_card = '速查卡' in rep.all_text
    has_cmd = '常用命令' in rep.all_text
    has_err = '报错速查' in rep.all_text or '常见报错' in rep.all_text
    add('含速查卡、常用命令表、报错速查三张表', has_card and has_cmd and has_err,
        'card=%s cmd=%s err=%s' % (has_card, has_cmd, has_err))

    # 9. 颜色体系
    cmd_green = rep.runs_matching(lambda r: r['color'] == '00B050')
    warn_red = rep.runs_matching(lambda r: r['color'] == 'C00000')
    shot_gray = [t for t in rep.runs_matching(lambda r: r['color'] == '808080' and r['italic'])]
    ph_green = [t for t in rep.runs_matching(lambda r: r['color'] == '92D050')]
    add('命令用绿色 00B050', len(cmd_green) >= 5, 'x%d' % len(cmd_green))
    add('【注意】/【错误】用红色 C00000', len(warn_red) >= 2, 'x%d' % len(warn_red))
    add('截图要求用灰色 808080 斜体', len(shot_gray) >= 1, 'x%d' % len(shot_gray))
    add('截图占位符用浅绿 92D050', len(ph_green) >= 1, 'x%d' % len(ph_green))

    # 10. 无 emoji
    bad = EMOJI.findall(rep.all_text)
    add('正文没有 emoji', not bad, 'found=%s' % (bad[:8],))

    # 11. 文件名短且不带下划线
    #     注意这里查的是**学生自己要敲的文件名**，不是命令名。有些 Linux 命令本身
    #     就带下划线（xfs_growfs、mkfs.xfs 之类），它们是工具名、学生改不了，不能
    #     算违规——早先的版本把它们误判成文件名。
    REAL_CMDS = {
        'cd', 'ls', 'pwd', 'mkdir', 'touch', 'cp', 'mv', 'rm', 'cat', 'echo',
        'chmod', 'chown', 'chgrp', 'ln', 'df', 'du', 'stat', 'head', 'tail',
        'grep', 'ps', 'top', 'kill', 'jobs', 'sleep', 'useradd', 'usermod',
        'userdel', 'groupadd', 'groupdel', 'passwd', 'journalctl', 'find',
        'less', 'more', 'wc', 'sort', 'uniq', 'id', 'groups', 'sudo', 'su',
        'curl', 'wget', 'systemctl', 'mount', 'umount', 'fdisk', 'lsblk',
        'partprobe', 'blkid', 'pvcreate', 'vgcreate', 'lvcreate', 'pvs',
        'vgs', 'lvs', 'vgextend', 'lvextend', 'lvremove', 'vgremove',
        'pvremove', 'xfs_growfs', 'mkfs.xfs', 'mkfs.ext4', 'logger', 'date',
        'whoami', 'dnf', 'yum', 'rpm', 'tar', 'vi', 'vim', 'nano', 'hostname',
        'ip', 'ping', 'ss', 'netstat', 'dmesg', 'lscpu', 'free', 'uname',
        # 网络与远程这一块用到的工具名
        'ssh', 'scp', 'sftp', 'ssh-keygen', 'ssh-copy-id', 'ssh-keyscan',
        'nmcli', 'sshd', 'ping6', 'traceroute', 'nslookup', 'dig',
    }
    #     还有一类同样「学生改不了」的东西：协议和服务自己规定的文件名、配置项、
    #     系统目录。sshd_config、authorized_keys、id_rsa.pub 是 SSH 定死的，
    #     PermitRootLogin 是 sshd 的配置键，ipv4.ignore-auto-dns 是 nmcli 的参数名，
    #     network-scripts 是系统目录——它们长、带下划线、带点，但一个字母都不能改。
    #     网络配置和 SSH 这两讲整篇都在跟这些名字打交道，不豁免的话，
    #     规则会从头到尾报红，真正的信号（学生自造的长文件名）反而被淹掉。
    SYS_IDENTS = {
        # ssh / sshd 定死的文件名
        'sshd_config', 'ssh_config', 'authorized_keys', 'known_hosts',
        'id_rsa', 'id_rsa.pub', 'id_ecdsa', 'id_ecdsa.pub',
        'id_ed25519', 'id_ed25519.pub', 'openssh-server', 'openssh-clients',
        # sshd 的配置键
        'PermitRootLogin', 'PasswordAuthentication', 'PubkeyAuthentication',
        'ChallengeResponseAuthentication', 'prohibit-password',
        'X11Forwarding', 'AllowUsers', 'MaxAuthTries',
        # nmcli 的参数名
        'connection.id', 'connection.autoconnect', 'ipv4.method',
        'ipv4.addresses', 'ipv4.gateway', 'ipv4.dns', 'ipv4.ignore-auto-dns',
        'ipv6.method', 'ipv6.addresses',
        # 系统目录和 NetworkManager 自己的名字
        'network-scripts', 'NetworkManager', 'network-scripts/',
    }
    #     域名也不是文件名。判断「点后面跟着一个常见顶级域」比逐个域名列入白名单
    #     更经得住用，换个例子（www.qq.com、mirrors.aliyun.com）照样认得出来。
    DOMAIN = re.compile(r'\.(com|cn|org|net|edu|gov|io)(\.|$)', re.I)

    idents = set()
    for t in cmd_green:
        for tok in re.findall(r'[A-Za-z][A-Za-z0-9_\.\-]*', t):
            if tok.lower() in REAL_CMDS:
                continue
            if tok in SYS_IDENTS:
                continue
            if DOMAIN.search(tok):
                continue
            if tok.startswith('-'):
                continue
            idents.add(tok)
    unders = sorted(t for t in idents if '_' in t)
    longs = sorted(t for t in idents if len(t) > 12)
    add('命令中的文件名短且不含下划线', not unders and not longs,
        'underscore=%s long=%s' % (unders[:5], longs[:5]))

    # 12. 实验体会问具体问题
    has_ask = '实验体会' in rep.all_text
    n_q = rep.count('？')
    generic = '谈谈你的收获' in rep.all_text or '谈谈本次实验的收获' in rep.all_text
    add('实验体会提出具体问题（非「谈谈收获」式空题）',
        has_ask and n_q >= 2 and not generic, '问号 x%d generic=%s' % (n_q, generic))

    # 13. 模板占位文字必须被清干净。
    #     模板里实验目的有 3 个槽位、设备 6 个、体会 2 个，内容多了靠克隆补；
    #     反过来，内容比槽位少时，如果只收够 n 条就停手，多余槽位会带着
    #     「【此处填写实验目的第 3 条…】」原样留在文档里交到学生手上。这个错很隐蔽
    #     ——自己按名单数一遍站点、表格全都对得上，但文档里就多着一行占位提示。
    leftovers = [t for t in ('【此处填写', '【问题 1', '【问题 2', '【问题 3')
                 if t in rep.all_text]
    add('模板占位文字已全部替换', not leftovers, 'leftover=%s' % (leftovers,))

    # 14. 体会引导语里的题数要和实际题数一致。
    #     模板写死「下面两个问题」。问了 4 题却留着这句，学生会以为漏看了两题。
    #     这类「两处各自正常、只有比较才看得出」的错误，和占位文字一样属于
    #     结构检查查不出来的那一类。
    #     数题时有两个坑：一是不能只看「以？结尾」（写成「……说说你怎么挑。」照样
    #     是道题），二是不能全文数蓝色——【正确】【提示】那些行也是蓝的，一数就是
    #     四十几条。正确的范围是「四、实验体会」标题之后的那一段。
    m = re.search(r'请围绕下面\s*(\S+)\s*个问题', rep.all_text)
    zh = {'一': 1, '两': 2, '二': 2, '三': 3, '四': 4, '五': 5, '六': 6}
    tok = m.group(1) if m else None
    said = None
    if tok is not None:
        said = zh.get(tok)
        if said is None and tok.isdigit():
            said = int(tok)
    start = None
    for i, p in enumerate(rep.paras):
        if '实验体会' in p['text'] and p['text'].strip().startswith('四'):
            start = i
            break
    n_qq = 0
    if start is not None:
        for p in rep.paras[start + 1:]:
            runs = [r for r in p['runs'] if r['text'].strip()]
            if p['text'].strip() and runs and all(r['color'] == '0070C0' for r in runs):
                n_qq += 1
    if m is None:
        add('体会引导语的题数与实际题数一致', False, '全文找不到「请围绕下面 N 个问题」这句引导语')
    else:
        add('体会引导语的题数与实际题数一致', said == n_qq,
            '引导语说 %s 题，实际蓝色题目 %d 条' % (tok, n_qq))

    # 15. 正文里不能有 markdown 记号。
    #     spec 是写成 Python 字符串的，人在写的时候手会习惯性敲 **加粗**、
    #     `代码`、## 标题——但这些东西不会被解析，会原样打进 Word 里，
    #     变成正文中突兀的星号。这一条我在三份报告里都犯过，说明它不是偶然，
    #     所以交给脚本查，别指望记得住。
    #     模板抬头那行本身带 25** / *** / 2512****（班级姓名学号的占位），要放过。
    bad_md = []
    for p in rep.paras:
        t = p['text']
        if not t.strip():
            continue
        if '班级：' in t and '学号：' in t:
            continue                      # 模板自己的抬头占位行
        for tok in ('**', '##', '`'):
            if tok in t:
                bad_md.append(t.strip()[:40])
                break
    add('正文没有混进 markdown 记号（** ## `）', not bad_md,
        'found=%s' % (bad_md[:5],))

    # 16. 三张表格的字体要统一成「中文仿宋 + 英文 Consolas + 五号」。
    #     这条查的是**每个 run 上显式写了没有**，不是「渲染出来像不像」——
    #     因为出问题的正是「没写、靠继承」：表头在模板里写全了，数据行却是空段落，
    #     往里面填字时只能造一个空 rPr，ascii 还能从段落标记继承到 Consolas，
    #     eastAsia 谁都声明过，一路掉回 docDefaults 里的宋体。表头仿宋、数据宋体，
    #     同一张表两种中文字体，不放大看根本注意不到。
    tbl_bad = []
    tbl_runs = 0
    for t in rep.tables:
        if t['cols'] != 3:
            continue                  # 三张表都是 3 列，外层大表是 5 列
        for tr in t['el'].findall(w('tr')):
            for tc in tr.findall(w('tc')):
                for p in tc.findall(w('p')):
                    for r in p.findall(w('r')):
                        tbl_runs += 1
                        rPr = r.find(w('rPr'))
                        f = rPr.find(w('rFonts')) if rPr is not None else None
                        sz = rPr.find(w('sz')) if rPr is not None else None
                        got = (f.get(w('eastAsia')) if f is not None else None,
                               f.get(w('ascii')) if f is not None else None,
                               sz.get(w('val')) if sz is not None else None)
                        if got != ('仿宋', 'Consolas', '21'):
                            tbl_bad.append('%s|%s|%s' % got)
    add('三张表格字体为仿宋 + Consolas + 五号',
        tbl_runs > 0 and not tbl_bad,
        '检查 %d 个 run，不合格 %d 个%s' % (tbl_runs, len(tbl_bad),
                                        ('：' + ', '.join(tbl_bad[:3])) if tbl_bad else ''))

    return res


def main():
    if len(sys.argv) < 2:
        print(__doc__)
        raise SystemExit(2)
    path = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else None
    res = grade(path)
    n_pass = sum(1 for r in res if r['passed'])
    payload = {'expectations': res}
    if out:
        io.open(out, 'w', encoding='utf-8').write(json.dumps(payload, ensure_ascii=False, indent=2))
    print('%s: %d/%d passed -> %s' % (os.path.basename(path), n_pass, len(res), out or 'stdout'))
    for r in res:
        print('  %s %s  (%s)' % ('PASS' if r['passed'] else 'FAIL', r['text'], r['evidence'][:90]))


if __name__ == '__main__':
    main()
