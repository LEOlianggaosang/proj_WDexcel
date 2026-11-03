# -*- coding: utf-8 -*-
"""
柜机时序表 · 第二步 v3 第二阶段：双列区（衍生物）重建
=====================================================
输入：02_衍生物/02_中间产物/普通版_第二步_本体已改写.xlsx
输出：03_验收物/普通版_第二步已处理.xlsx

规范（第二步确立，不再复刻源表的手工痕迹）：
  · 双列区起点 = 双速电机列 + 1（辅助列已在第一阶段清空）
  · 每组 2 列：第 1 列累计分钟，第 2 列倒计时串 `本步耗时/剩余(H:MM)`
  · 第 3 行：组名写在**第 2 列**，第 1 列留空但保留框线
  · 第 4 行：起值行 `0/H:MM`，绿字
  · 该组不执行 → 两列置空 + 灰底（两者必须同时出现）
  · 端行统一 `1/end`，绿字
  · 双列区字体统一：微软雅黑 10 号
"""
import os
import re
from copy import copy

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter as gl

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SRC = os.path.join(ROOT, '02_衍生物', '02_中间产物', '普通版_第二步_本体已改写.xlsx')
DST = os.path.join(ROOT, '03_验收物', '普通版_第二步已处理.xlsx')

C_GREEN = 'FF92D050'
GRAY_FILL = PatternFill(patternType='solid', fgColor=Color(theme=0, tint=-0.15))
NO_FILL = PatternFill(patternType=None)
THIN = Side(style='thin', color='FF808080')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
FONT = Font(name='微软雅黑', size=10)

# 每个分栏的精确布局。auto-probe 只用于校验，不作为依据
# （三张表表头位置各不相同：Soak 数据起于第 4 行、总时长在第 3 行；Rapid 组名是大写 NULL）
CFG2 = {
    'Intensive ':    dict(sw0=5, cd=12, motor=13, r0=5, ngrp=5, tot_row=4),
    'Glass ':        dict(sw0=5, cd=12, motor=13, r0=5, ngrp=5, tot_row=4),
    '89Mins':        dict(sw0=5, cd=12, motor=13, r0=5, ngrp=5, tot_row=4),
    'Rapid':         dict(sw0=5, cd=7,  motor=8,  r0=5, ngrp=2, tot_row=4),
    'Soak':          dict(sw0=5, cd=6,  motor=7,  r0=4, ngrp=1, tot_row=3),
    'Self-cleaning': dict(sw0=5, cd=6,  motor=7,  r0=5, ngrp=1, tot_row=4),
}

# Glass 的 Express 标称 115 经 A3 反推判定为笔误，真值 105（P1 已批准）
OVERRIDE_TOTAL = {('Glass ', 'Express'): 105}


def norm(s):
    return re.sub(r'\s+', '', str(s or ''))


def probe(ws):
    """自动定位开关矩阵起点 / 配时列 / 双速电机列，避免硬编码"""
    sw0 = cd = motor = None
    for c in range(1, ws.max_column + 1):
        v3, v2 = norm(ws.cell(3, c).value), norm(ws.cell(2, c).value)
        if v3 == 'Null' and sw0 is None:
            sw0 = c
        if v3 == '倒计时' and cd is None:
            cd = c
        if '双速电机' in (v2 + v3) and motor is None:
            motor = c
    return sw0, cd, motor


def last_data_row(ws, r0=5):
    r = r0
    for i in range(r0, ws.max_row + 1):
        if isinstance(ws.cell(i, 1).value, (int, float)):
            r = i
    return r


def fmt_hmm(m):
    m = int(max(m, 0))
    return '%d:%02d' % (m // 60, m % 60)


def build(ws, sheet, log):
    cfg = CFG2.get(sheet)
    if cfg is None:
        log.append((sheet, '-', '⚠无配置', '', '跳过'))
        return
    sw0, cd, motor = cfg['sw0'], cfg['cd'], cfg['motor']
    r0, ngrp, tot_row = cfg['r0'], cfg['ngrp'], cfg['tot_row']

    # 位置自检：自动探测结果应与显式配置一致
    p = probe(ws)
    if p[1] is not None and p[1] != cd:
        log.append((sheet, '-', '⚠配时列不一致', '配置%s 探测%s' % (cd, p[1]), '用配置'))

    lr = last_data_row(ws, r0)
    out0 = motor + 1
    grp_row = r0 - 2          # 组名行
    start_row = r0 - 1        # 起值行

    # 彻底清空旧双列区（含残留辅助列与 Rapid 的半成品 I/J）
    for c in range(out0, ws.max_column + 1):
        for r in range(1, ws.max_row + 1):
            cell = ws.cell(r, c)
            cell.value = None
            cell.font = Font(name='微软雅黑', size=10)
            cell.fill = NO_FILL
            cell.border = Border()

    for g in range(ngrp):
        swc = sw0 + g
        c1 = out0 + g * 2          # 累计列
        c2 = c1 + 1                # 倒计时列
        label = ws.cell(3, swc).value
        if label in (None, '') or isinstance(label, (int, float)):
            label = 'Null'          # Soak 该行存的是总时长，无组名文字
        label = str(label).replace('\n', '')

        total = OVERRIDE_TOTAL.get((sheet, label), ws.cell(tot_row, swc).value)

        # 组名行：组名写第 2 列，第 1 列留空但保留框线
        ws.cell(grp_row, c1).value = None
        ws.cell(grp_row, c1).border = BORDER
        ws.cell(grp_row, c2).value = label
        ws.cell(grp_row, c2).font = FONT
        ws.cell(grp_row, c2).border = BORDER

        if not isinstance(total, (int, float)):
            log.append((sheet, label, '⚠无总时长', str(total), '跳过'))
            continue
        key = (sheet, label)
        if key in OVERRIDE_TOTAL:
            ws.cell(tot_row, swc).value = total      # 写回修正后的标称值
        total = int(total)

        cum = 0
        seq = []
        for r in range(r0, lr + 1):
            on = ws.cell(r, swc).value not in (None, '')
            v = ws.cell(r, cd).value
            v = v if isinstance(v, (int, float)) else 0
            if on:
                cum += v
            seq.append((r, on, cum))

        # 起值行
        ws.cell(start_row, c1).border = BORDER
        ws.cell(start_row, c2).border = BORDER
        ws.cell(start_row, c2).value = '0/%s' % fmt_hmm(total)
        ws.cell(start_row, c2).font = Font(name='微软雅黑', size=10, color=C_GREEN)

        for i, (r, on, c) in enumerate(seq):
            v = ws.cell(r, cd).value
            v = v if isinstance(v, (int, float)) else 0
            is_last = (i == len(seq) - 1)
            if not on:
                ws.cell(r, c1).value = None
                ws.cell(r, c2).value = None
                ws.cell(r, c1).fill = GRAY_FILL
                ws.cell(r, c2).fill = GRAY_FILL
            else:
                ws.cell(r, c1).fill = NO_FILL
                ws.cell(r, c2).fill = NO_FILL
                ws.cell(r, c1).value = c if v else None
                # 终行 = 本步耗时/end（2026-11-03 纠错：不再是固定 1/end）
                ws.cell(r, c2).value = ('%d/end' % v) if is_last else '%d/%s' % (v, fmt_hmm(total - c))
            for cc in (c1, c2):
                cell = ws.cell(r, cc)
                edge = (cc == c2 and on and is_last)
                cell.font = Font(name='微软雅黑', size=10, color=C_GREEN if edge else None)
                cell.border = BORDER

        log.append((sheet, label, '双列区', '列%s-%s' % (gl(c1), gl(c2)),
                    '总时长%s，累计终值%s%s' % (
                        total, seq[-1][2] if seq else 'NA',
                        ' ✅已修正' if key in OVERRIDE_TOTAL else '')))


def main():
    wb = openpyxl.load_workbook(SRC)
    log = []
    for sheet in CFG2:
        if sheet not in wb.sheetnames:
            log.append((sheet, '-', '⚠表不存在', '', '跳过'))
            continue
        build(wb[sheet], sheet, log)

    # 六个已维护分栏标签统一标绿
    for sheet in CFG2:
        if sheet in wb.sheetnames:
            wb[sheet].sheet_properties.tabColor = C_GREEN

    os.makedirs(os.path.dirname(DST), exist_ok=True)
    wb.save(DST)
    print('=== 双列区重建完成 ->', DST)
    for x in log:
        print('  ', x)


if __name__ == '__main__':
    main()
