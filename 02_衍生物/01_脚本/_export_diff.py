# -*- coding: utf-8 -*-
"""第二步 · 导出逐行改写对照（源 vs 验收成品）

用途：替代手工维护的「措辞对照表」。措辞规则的真源是脚本本身
      （_step2_fill.py 的 rewrite_desc + 常量表），本脚本只是把它跑一遍后
      导出人类可读的逐行对照，供验收核对。
      文档不再手工维护，避免与规则漂移。

输入：01_源文件/普通版.xlsx（源，只读）
      03_验收物/普通版_第二步已处理.xlsx（成品）
输出：02_衍生物/02_中间产物/第二步_逐行改写对照.md
"""
import os
import io
import difflib
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SRC = os.path.join(ROOT, '01_源文件', '普通版.xlsx')
FIN = os.path.join(ROOT, '03_验收物', '普通版_第二步已处理.xlsx')
OUT = os.path.join(ROOT, '02_衍生物', '02_中间产物', '第二步_逐行改写对照.md')

# (源分栏名, 成品分栏名, 配时列)
SHEETS = [
    ('Intensive ', 'Intensive ', 12),
    ('Glass ', 'Glass ', 12),
    ('90Mins', '89Mins', 12),
    ('Rapid', 'Rapid', 7),
    ('Soak', 'Soak', 6),
    ('Self-cleaning', 'Self-cleaning', 6),
]


def norm(v):
    return '∅' if v in (None, '') else str(v).replace('\n', '/').strip()


def rows(ws, cd):
    out = []
    for r in range(5, ws.max_row + 1):
        a, c, d = ws.cell(r, 1).value, ws.cell(r, 3).value, ws.cell(r, 4).value
        v = ws.cell(r, cd).value
        if a is None and c is None and d is None and v is None:
            continue
        out.append((r, a, norm(c), norm(d), v))
    return out


def fill_tag(ws, r, cd):
    """返回该行的标记摘要：填充色 + 红字列"""
    fills, reds = set(), []
    for cc in range(3, cd + 2):
        f = ws.cell(r, cc).fill
        if f and f.patternType and f.fgColor and f.fgColor.type == 'rgb':
            v = str(f.fgColor.value)
            fills.add('绿' if v == 'FF92D050' else '黄' if v == 'FFFFFF00' else v)
        fo = ws.cell(r, cc).font
        if fo.color and fo.color.type == 'rgb' and str(fo.color.value) == 'FFFF0000':
            reds.append(openpyxl.utils.get_column_letter(cc))
    tag = ''
    if fills:
        tag += ''.join(sorted(fills)) + '底'
    if reds:
        tag += ' 红:' + ','.join(reds)
    return tag or '—'


def main():
    src = openpyxl.load_workbook(SRC)
    fin = openpyxl.load_workbook(FIN)
    L = ['# 第二步 · 逐行改写对照（源 → 验收成品）', '',
         '> 由 `02_衍生物/01_脚本/_export_diff.py` 自动生成，请勿手工编辑。',
         '> 只列出**有变化的行**。基准＝`01_源文件/普通版.xlsx`，对照＝`03_验收物/普通版_第二步已处理.xlsx`。',
         '> 标记列说明：绿底=结构性改动，黄底=参数/描述级改动，`红:列`＝红字所在列。', '']
    total = 0

    for sname, fname, cd in SHEETS:
        ws_s, ws_f = src[sname], fin[fname]
        rs, rf = rows(ws_s, cd), rows(ws_f, cd)
        key_s = [(x[2], x[3], x[4]) for x in rs]
        key_f = [(x[2], x[3], x[4]) for x in rf]
        sm = difflib.SequenceMatcher(a=key_s, b=key_f, autojunk=False)

        # 收集变化：replace / delete / insert
        changes = []
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == 'equal':
                continue
            if tag == 'replace':
                # 两侧长度不等时，多出来的一侧按增/删处理
                for k in range(max(i2 - i1, j2 - j1)):
                    a = rs[i1 + k] if i1 + k < i2 else None
                    b = rf[j1 + k] if j1 + k < j2 else None
                    if a is None:
                        changes.append(('增', None, b))
                    elif b is None:
                        changes.append(('删', a, None))
                    else:
                        changes.append(('改', a, b))
            elif tag == 'delete':
                for k in range(i1, i2):
                    changes.append(('删', rs[k], None))
            elif tag == 'insert':
                for k in range(j1, j2):
                    changes.append(('增', None, rf[k]))

        L.append('## %s → %s' % (sname.strip(), fname.strip()))
        L.append('')
        L.append('源 %d 行　成品 %d 行　变化 %d 处' % (len(rs), len(rf), len(changes)))
        L.append('')
        if not changes:
            L.append('（无变化）')
            L.append('')
            continue

        L.append('| 操作 | 序号 | 描述（源 → 成品） | 参数（源 → 成品） | 配时 | 成品标记 |')
        L.append('|---|---|---|---|---|---|')
        for op, a, b in changes:
            if op == '删':
                L.append('| 删 | %s | ~~%s~~ | ~~%s~~ | %s | — |'
                         % (a[1], a[2], a[3], a[4]))
                continue
            if op == '增':
                L.append('| 增 | %s | **%s** | **%s** | %s | %s |'
                         % (b[1], b[2], b[3], b[4], fill_tag(ws_f, b[0], cd)))
                continue
            ds = a[2] if a[2] != b[2] else a[2]
            dp = a[3] if a[3] != b[3] else a[3]
            dv = '%s→%s' % (a[4], b[4]) if a[4] != b[4] else str(b[4])
            L.append('| 改 | %s | %s | %s | %s | %s |'
                     % (b[1], ds if a[2] == b[2] else '%s → **%s**' % (a[2], b[2]),
                        dp if a[3] == b[3] else '%s → **%s**' % (a[3], b[3]),
                        dv, fill_tag(ws_f, b[0], cd)))
        L.append('')
        total += len(changes)

    L.insert(5, '**变化合计：%d 处**' % total)
    L.insert(6, '')
    io.open(OUT, 'w', encoding='utf-8').write('\n'.join(L))
    print('已导出：%s' % OUT)
    for sname, fname, cd in SHEETS:
        ws_s, ws_f = src[sname], fin[fname]
        rs, rf = rows(ws_s, cd), rows(ws_f, cd)
        ks = [(x[2], x[3], x[4]) for x in rs]
        kf = [(x[2], x[3], x[4]) for x in rf]
        sm = difflib.SequenceMatcher(a=ks, b=kf, autojunk=False)
        n = sum(1 for t, *_ in sm.get_opcodes() if t != 'equal')
        print('  %-14s 源%3d行 成品%3d行 变化段%2d' % (fname, len(rs), len(rf), n))
    print('  合计 %d 处变化' % total)


if __name__ == '__main__':
    main()
