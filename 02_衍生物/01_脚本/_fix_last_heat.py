# -*- coding: utf-8 -*-
"""
柜机时序表 · 成品一温达跳步按分组修复（V28，2026-11-03 用户纠错）
================================================================
用户口径：89分钟洗对温达跳步的定义搞错了。温达跳步与热漂核心机制有关——
**对于所有叠加功能/分组，到达干燥前的最后一次定温加热都做成温达跳步**，
目的是加热到温度就迅速进入干燥，利用余温提升干燥效率。
（对照样板：Glass/Intensive 漂洗尾段每分组最后一行加热全部是温达跳步+纯T参数。）

V16 旧实现错在取 `rows[-1]`（整表最后一行加热）——只转了 89Mins 的 ExtraDry 行。
正确判定（本补丁）：漂洗段内某「温达剩洗」行，若它在 E–K 任一开关列（含
MoreRin/MorePre）上是**该列最后一次加热行**（其后到干燥段前无更晚的同列加热行），
则转为温达跳步。预期命中：89Mins r42–r46（Energy / MoreRin+MorePre / Null /
StrWash / Express）；Glass、Intensive、Self-cleaning 已符合（各分组末行均跳步）；
Rapid/Soak 无干燥段不适用。

转换内容（照 V16 跳步标法，与既有 r47 完全一致）：
  · C 富文本按 RT_PATTERNS 跳步模式重建：[黑]'洗涤加热\\n' + [红]'（定温加热-温达跳步）'
  · D 参数 T/秒 → 纯 T（如 T66/840 → T66）
  · C–M 先清填充/红字，再上黄 C–D；L 配时不变（总时长守恒）、开关不变
成品一原位补丁（accept 非幂等不可重跑）；成品二随后重跑 _step3_door.py 派生。
"""
import os
import re

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock, InlineFont
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.colors import Color

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DST = os.path.join(ROOT, '03_验收物', '普通版_第二步已处理.xlsx')

RED, YELLOW, BLACK = 'FFFF0000', 'FFFFFF00', 'FF000000'

CFG = {
    'Intensive ':    dict(sw0=5, cd=12, motor=13, r0=5),
    'Glass ':        dict(sw0=5, cd=12, motor=13, r0=5),
    '89Mins':        dict(sw0=5, cd=12, motor=13, r0=5),
    'Rapid':         dict(sw0=5, cd=7,  motor=8,  r0=5),
    'Soak':          dict(sw0=5, cd=6,  motor=7,  r0=4),
    'Self-cleaning': dict(sw0=5, cd=6,  motor=7,  r0=5),
}
N_SW = 7          # E–K：Null, ExtraDry, Energy, Express, StrWash, MoreRin, MorePre


def norm(s):
    return re.sub(r'\s+', '', str(s or ''))


def data_rows(ws, r0):
    lr = r0
    for r in range(r0, ws.max_row + 1):
        if ws.cell(r, 3).value not in (None, '') or ws.cell(r, 4).value not in (None, ''):
            lr = r
    return [r for r in range(r0, lr + 1)
            if ws.cell(r, 3).value not in (None, '') or ws.cell(r, 4).value not in (None, '')]


def stage_map(ws, r0):
    st = {}
    for m in ws.merged_cells.ranges:
        if m.min_col == 2 and m.min_row >= r0:
            v = ws.cell(m.min_row, 2).value
            for r in range(m.min_row, m.max_row + 1):
                st[r] = v
    for r in range(r0, ws.max_row + 1):
        v = ws.cell(r, 2).value
        if v not in (None, ''):
            st.setdefault(r, v)
    return st


def make_rt():
    return CellRichText(
        TextBlock(InlineFont(rFont='微软雅黑', color=Color(rgb=BLACK)), '洗涤加热\n'),
        TextBlock(InlineFont(rFont='微软雅黑', color=Color(rgb=RED)), '（定温加热-温达跳步）'),
    )


def main():
    wb = openpyxl.load_workbook(DST, rich_text=True)
    converted = []
    for n, cfg in CFG.items():
        ws = wb[n]
        rows = data_rows(ws, cfg['r0'])
        st = stage_map(ws, cfg['r0'])
        # 表内必须有干燥段（Rapid/Soak 无干燥，规则不适用）
        dry_rows = [r for r in rows if '干燥' in norm(st.get(r))]
        if not dry_rows:
            print('%-14s 无干燥段，跳过' % n)
            continue
        dry0 = dry_rows[0]
        heat_rows = [r for r in rows
                     if r < dry0 and '加热' in norm(ws.cell(r, 3).value)
                     and '分配' not in norm(ws.cell(r, 3).value)]
        for r in [x for x in heat_rows if '温达剩洗' in norm(ws.cell(x, 3).value)]:
            on = [c for c in range(cfg['sw0'], cfg['sw0'] + N_SW)
                  if ws.cell(r, c).value not in (None, '')]
            # 该行是某开关列「最后一次加热行」？其后到干燥段前该列无更晚加热行
            is_last = any(
                not any(ws.cell(r2, c).value not in (None, '')
                        for r2 in heat_rows if r2 > r)
                for c in on)
            if not is_last:
                continue
            old_param = str(ws.cell(r, 4).value)
            t = re.match(r'^(T\d+)', old_param)
            if not t:
                print('  ⚠ %s r%d 参数 %r 不匹配 T 格式，跳过' % (n, r, old_param))
                continue
            # C：富文本按跳步模式重建
            ws.cell(r, 3).value = make_rt()
            # D：纯 T
            ws.cell(r, 4).value = t.group(1)
            # C–M 清旧标记（剩洗标法黄 C–M + 红 D–M），再上跳步标法黄 C–D
            for c in range(3, cfg['motor'] + 1):
                cell = ws.cell(r, c)
                cell.fill = PatternFill()
                f = cell.font
                if f.color and f.color.type == 'rgb' and str(f.color.value) == RED:
                    cell.font = Font(name=f.name or '微软雅黑', size=f.size or 10,
                                     bold=f.bold, italic=f.italic)
            for c in (3, 4):
                ws.cell(r, c).fill = PatternFill('solid', fgColor=Color(rgb=YELLOW))
            converted.append('%s r%d %s → %s（配时 %s 不变）'
                             % (n, r, old_param, t.group(1),
                                ws.cell(r, cfg['cd']).value))
    wb.save(DST)
    print('转换 %d 行：' % len(converted))
    for x in converted:
        print('  ' + x)
    print('已保存 ->', DST)
    return 0 if converted else 1


if __name__ == '__main__':
    raise SystemExit(main())
