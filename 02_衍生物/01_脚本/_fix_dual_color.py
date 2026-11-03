# -*- coding: utf-8 -*-
"""
柜机时序表 · 成品一双列区配色原位修复（V26，2026-11-03 用户纠错）
================================================================
背景：accept 的 dual_zone 旧版把倒计时红字挂在「C 列黄填充」上，
导致：静止（再生间隔）预留行（黄填充、非定温加热）倒计时被错标红。
正确语义（用户 2026-11-03 口径）：
  · 倒计时列 c2：绿 = end；蓝 = 描述含「进水」（实机进水提醒开水）；
    红 = 描述含「定温加热」（实机加热提醒变温）；其余黑
  · 累计列 c1：恒黑（源表规范，Glass r21：累计18黑 / 倒计时红）
为什么原位补丁而不重跑 accept：accept 是对成品一的非幂等原位修复链
（V20 会再插一行预留行），重跑有结构风险；本补丁只重着色，不动结构。
"""
import os
import re

import openpyxl
from openpyxl.styles import Font
from openpyxl.styles.colors import Color

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DST = os.path.join(ROOT, '03_验收物', '普通版_第二步已处理.xlsx')

GREEN, RED, BLUE = 'FF92D050', 'FFFF0000', 'FF00B0F0'

CFG = {
    'Intensive ':    dict(sw0=5, cd=12, motor=13, r0=5, ngrp=5),
    'Glass ':        dict(sw0=5, cd=12, motor=13, r0=5, ngrp=5),
    '89Mins':        dict(sw0=5, cd=12, motor=13, r0=5, ngrp=5),
    'Rapid':         dict(sw0=5, cd=7,  motor=8,  r0=5, ngrp=2),
    'Soak':          dict(sw0=5, cd=6,  motor=7,  r0=4, ngrp=1),
    'Self-cleaning': dict(sw0=5, cd=6,  motor=7,  r0=5, ngrp=1),
}


def norm(s):
    return re.sub(r'\s+', '', str(s or ''))


def data_rows(ws, r0):
    lr = r0
    for r in range(r0, ws.max_row + 1):
        if ws.cell(r, 3).value not in (None, '') or ws.cell(r, 4).value not in (None, ''):
            lr = r
    return [r for r in range(r0, lr + 1)
            if ws.cell(r, 3).value not in (None, '') or ws.cell(r, 4).value not in (None, '')]


def main():
    wb = openpyxl.load_workbook(DST, rich_text=True)
    stats = {'旧红→黑': 0, '保持红': 0, '蓝': 0, '绿': 0, '黑': 0, '累计列去色': 0}
    for n in CFG:
        ws = wb[n]
        cfg = CFG[n]
        out0, ngrp, r0 = cfg['motor'] + 1, cfg['ngrp'], cfg['r0']
        rows = data_rows(ws, r0)
        for g in range(ngrp):
            c1, c2 = out0 + g * 2, out0 + g * 2 + 1
            for r in rows:
                # 累计列恒黑
                cell1 = ws.cell(r, c1)
                fc = cell1.font.color
                if fc and fc.type == 'rgb' and str(fc.value) in (RED, BLUE, GREEN):
                    cell1.font = Font(name='微软雅黑', size=10)
                    stats['累计列去色'] += 1
                # 倒计时列按描述重着色
                v2 = ws.cell(r, c2).value
                if v2 in (None, ''):
                    continue
                d = norm(ws.cell(r, 3).value)
                old = ws.cell(r, c2).font.color
                old = str(old.value) if old and old.type == 'rgb' else None
                if str(v2).endswith('/end'):
                    col = GREEN
                    stats['绿'] += 1
                elif '进水' in d:
                    col = BLUE
                    stats['蓝'] += 1
                elif '定温加热' in d:
                    col = RED
                    stats['保持红'] += 1
                else:
                    col = None
                    stats['黑'] += 1
                if col != old and old == RED and col is None:
                    stats['旧红→黑'] += 1
                ws.cell(r, c2).font = Font(name='微软雅黑', size=10,
                                           color=Color(rgb=col) if col else None)
    wb.save(DST)
    print('成品一双列区配色修复 ->', DST)
    print('  %s' % stats)
    return 0


if __name__ == '__main__':
    raise SystemExit(main())
