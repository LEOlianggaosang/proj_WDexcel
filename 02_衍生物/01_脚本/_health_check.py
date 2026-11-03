# -*- coding: utf-8 -*-
"""第二步 · 全表体检

查什么：
  ① 分栏覆盖：源里有、成品里没处理到的分栏
  ② 序号缺口
  ③ 完整性：有参数/配时却无描述
  ④ 双列区：每组「累计终值」是否等于第 tot_row 行标称
  ⑤ 双列区：执行行是否漏值
  ⑥ 阶段（B 列）是否覆盖所有数据行
  ⑦ 填充双向比对：
       A 类 = 规则说该标、实际没标        → 漏标
       B 类 = 实际标了、规则说不标        → **规则盲区**（人工标记了我没破译的语义）

用法：python 02_衍生物/01_脚本/_health_check.py
"""
import os
import sys
import openpyxl
from openpyxl.utils import get_column_letter as gl

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import _step2_fill as F

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SRC = os.path.join(ROOT, '01_源文件', '普通版.xlsx')
FIN = os.path.join(ROOT, '03_验收物', '普通版_第二步已处理.xlsx')

SHEETS = ['Intensive ', 'Glass ', '89Mins', 'Rapid', 'Soak', 'Self-cleaning']

# 布局：cd=配时列, motor=双速电机列, r0=数据起始行, ngrp=分组数, tot_row=总时长行
LAYOUT = {
    'Intensive ':    dict(sw0=5, cd=12, motor=13, r0=5, ngrp=5, tot_row=4),
    'Glass ':        dict(sw0=5, cd=12, motor=13, r0=5, ngrp=5, tot_row=4),
    '89Mins':        dict(sw0=5, cd=12, motor=13, r0=5, ngrp=5, tot_row=4),
    'Rapid':         dict(sw0=5, cd=7,  motor=8,  r0=5, ngrp=2, tot_row=4),
    'Soak':          dict(sw0=5, cd=6,  motor=7,  r0=4, ngrp=1, tot_row=3),
    'Self-cleaning': dict(sw0=5, cd=6,  motor=7,  r0=5, ngrp=1, tot_row=4),
}


def has_fill(cell):
    f = cell.fill
    return bool(f and f.patternType and f.fgColor and f.fgColor.type == 'rgb'
                and str(f.fgColor.value) in ('FF92D050', 'FFFFFF00'))


def has_red(cell):
    fo = cell.font
    return bool(fo.color and fo.color.type == 'rgb' and str(fo.color.value) == 'FFFF0000')


def main():
    src = openpyxl.load_workbook(SRC)
    fin = openpyxl.load_workbook(FIN)
    issues = []

    # ① 分栏覆盖
    done = set(SHEETS)
    missing = [n for n in src.sheetnames if n not in done]
    issues.append(('①分栏覆盖',
                   '源 %d 个分栏，已处理 %d 个' % (len(src.sheetnames), len(SHEETS)) +
                   ('；未处理：%s（范围外，按 B1 不动）' % '、'.join(missing) if missing else ''),
                   missing))

    for name in SHEETS:
        ws = fin[name]
        cfg = LAYOUT[name]
        cd, motor, r0 = cfg['cd'], cfg['motor'], cfg['r0']
        ngrp, tot_row, sw0 = cfg['ngrp'], cfg['tot_row'], cfg['sw0']

        lr = r0
        for r in range(r0, ws.max_row + 1):
            if isinstance(ws.cell(r, 1).value, (int, float)):
                lr = r

        # ② 序号缺口
        seq = [ws.cell(r, 1).value for r in range(r0, lr + 1)]
        nums = [int(x) for x in seq if isinstance(x, (int, float))]
        gaps = []
        if nums:
            full = set(range(min(nums), max(nums) + 1))
            gaps = sorted(full - set(nums))
        if gaps:
            issues.append(('②序号缺口', '%s 缺序号 %s' % (name, gaps), gaps))

        # ③ 完整性
        bad = [r for r in range(r0, lr + 1)
               if not ws.cell(r, 3).value
               and (ws.cell(r, 4).value not in (None, '') or isinstance(ws.cell(r, cd).value, (int, float)))]
        if bad:
            issues.append(('③无描述', '%s 行 %s' % (name, bad), bad))

        # ⑥ 阶段覆盖（B 列是合并单元格，只有首行有值 → 按合并区判定）
        covered = set()
        for m in ws.merged_cells.ranges:
            if m.min_col <= 2 <= m.max_col:
                covered.update(range(m.min_row, m.max_row + 1))
        nostage = [r for r in range(r0, lr + 1)
                   if not ws.cell(r, 2).value and r not in covered and ws.cell(r, 3).value]
        if nostage:
            issues.append(('⑥无阶段', '%s 行 %s' % (name, nostage), nostage))

        # ④⑤ 双列区
        for g in range(ngrp):
            swc = sw0 + g
            c1 = motor + 1 + g * 2
            c2 = c1 + 1
            total = ws.cell(tot_row, swc).value
            if not isinstance(total, (int, float)):
                continue
            cum, miss = 0, []
            for r in range(r0, lr + 1):
                if ws.cell(r, swc).value in (None, ''):
                    continue
                v = ws.cell(r, cd).value
                v = v if isinstance(v, (int, float)) else 0
                cum += v
                if ws.cell(r, c1).value is None and ws.cell(r, c2).value is None:
                    miss.append(r)
            if cum != int(total):
                issues.append(('④时长不符', '%s 组%d 标称%s 实算%s' % (name, g, total, cum), [cum]))
            if miss:
                issues.append(('⑤双列漏值', '%s 组%d 行 %s' % (name, g, miss), miss))

        # ⑦ 填充双向比对
        missA, blindB = [], []
        for r in range(r0, lr + 1):
            desc = ws.cell(r, 3).value
            if not desc:
                continue
            kind, fill, span, kw = F.classify(desc)
            cols = range(3, motor + 1) if span == 'ROW' else (range(3, 5) if span == 'CD' else [])
            act_fill = any(has_fill(ws.cell(r, c)) for c in cols) if cols else False
            act_red = any(has_red(ws.cell(r, c)) for c in range(3, motor + 1))
            want_fill = fill is not None
            want_red = kw is not None or (kind == 'DRY_STATIC' and act_red)
            # B 类：实际标了，规则说不该标
            if (act_fill or act_red) and kind is None:
                blindB.append((r, str(desc)[:26].replace('\n', '/'),
                               '底' if act_fill else '', '红' if act_red else ''))
            elif want_fill and not act_fill:
                missA.append((r, str(desc)[:26], '缺底'))
            elif want_red and not act_red:
                missA.append((r, str(desc)[:26], '缺红'))
        if missA:
            issues.append(('⑦A漏标', '%s %d 处：%s' % (name, len(missA), missA[:4]), missA))
        if blindB:
            issues.append(('⑦B规则盲区', '%s %d 处：%s' % (name, len(blindB), blindB[:6]), blindB))

    print('=' * 70)
    print('第二步 · 全表体检')
    print('=' * 70)
    if not issues:
        print('  ✅ 无异常')
    for tag, desc, payload in issues:
        mark = '⚪' if tag.startswith('①') else '🔴' if tag.startswith(('②', '③', '④', '⑤', '⑥')) else '🟡'
        print('  %s %-12s %s' % (mark, tag, desc))
    print()
    print('图例：🔴 数据缺陷  🟡 规则待统一  ⚪ 说明')


if __name__ == '__main__':
    main()
