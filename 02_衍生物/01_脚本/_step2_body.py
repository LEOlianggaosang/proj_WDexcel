# -*- coding: utf-8 -*-
"""
柜机时序表 · 第二步重建 v3（本体变换阶段）
==========================================
输入：01_源文件/普通版.xlsx（只读，永不被写）
输出：02_衍生物/02_中间产物/普通版_第二步_本体已改写.xlsx

执行顺序（依赖不可调换）：
  1. 措辞改写（旧→新，电子可调类同步改参数与配时）
  2. 删 MINI 变体续行（U2：参数被吸收进拆出来的新行）
  3. 删漂洗第 2 次分配器（U1，按序号定位，不受行号偏移影响）
  4. 拆行（洗涤分配加热（主洗）→ 洗涤分配(主洗) + 洗涤加热（定时加热））
  5. 90Mins 特设：倒数第二步冷风配时 1→0；名称与总时长 90→89
  6. 清空辅助列
  7. 标签色（Q2）+ Glass autoFilter 收窄（U3）
  8. 自校验：每组总时长能否由「开关矩阵 × 配时列」反推（A3 裁决）

规则一律外置为顶部常量，新增情况补一条即可，不改逻辑。
"""
import os
import re
from copy import copy

import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter as gl

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SRC = os.path.join(ROOT, '01_源文件', '普通版.xlsx')
DST = os.path.join(ROOT, '02_衍生物', '02_中间产物', '普通版_第二步_本体已改写.xlsx')

# ---------------------------------------------------------------- 配色
YELLOW_FILL = PatternFill(patternType='solid', fgColor=Color(rgb='FFFFFF00'))
TAB_GREEN = 'FF92D050'

# ---------------------------------------------------------------- 措辞映射
RENAME = {
    '洗涤分配加热（主洗）': ('SPLIT', None),
    '洗涤分配加热（漂洗）': '洗涤加热分配（漂洗)分配器：开10秒停8秒累计动作6次',
    '洗涤加热分配（漂洗)': '洗涤加热分配（漂洗)分配器：开10秒停8秒累计动作6次',
    '再生(电子可调)': '再生进水风机（非再生跳步）',
    '再生（电子可调）': '再生进水风机（非再生跳步）',
    '进水(电子可调)': '冷风进水排水（边进边排-进完跳步）（非再生跳步）',
    '进水（电子可调）': '冷风进水排水（边进边排-进完跳步）（非再生跳步）',
}
# 电子可调类改写时，参数与倒计时同步改为模板值（裁决 B）
ELECTRIC_NEW = {
    '再生进水风机（非再生跳步）': ('0.5L', 0),
    '冷风进水排水（边进边排-进完跳步）（非再生跳步）': ('2L', 1),
}
SPLIT_ROWS = {
    '洗涤分配加热（主洗）': (('洗涤分配(主洗)', 1, 0), ('洗涤加热（定时加热）', 90, 1)),
}

# ---------------------------------------------------------------- 分栏配置
CFG = {
    '90Mins':        dict(cd=12, motor=13, aux=14, sw0=5, ngrp=7, rewrite=True),
    'Self-cleaning': dict(cd=6,  motor=7,  aux=8,  sw0=5, ngrp=1, rewrite=True),
}

# U1 批准删除：漂洗第 2 次分配器（按「序号」定位，避免行号漂移）
DEL_RINSE_2ND_SEQ = {'90Mins': 36, 'Self-cleaning': 21}

# 90Mins 业务变更：名称 90 → 89
NINETY_TO_EIGHTYNINE = True
RENAME_SHEET_90 = True


def norm(s):
    return re.sub(r'\s+', '', str(s or ''))


def last_data_row(ws):
    r = 5
    for i in range(5, ws.max_row + 1):
        if isinstance(ws.cell(i, 1).value, (int, float)):
            r = i
    return r


def rows_of_seq(ws, seq, cd_col):
    out = []
    for i in range(5, ws.max_row + 1):
        if ws.cell(i, 1).value == seq:
            out.append(i)
    return out


def transform_body(ws, cfg, sheet):
    cd = cfg['cd']
    log = []
    lr = last_data_row(ws)

    # ---- 1. 措辞改写
    for r in range(5, lr + 1):
        s = norm(ws.cell(r, 3).value)
        if not s:
            continue
        tgt = RENAME.get(s)
        if tgt is None:
            continue
        if isinstance(tgt, tuple):
            continue
        old = ws.cell(r, 3).value
        ws.cell(r, 3).value = tgt
        if tgt in ELECTRIC_NEW:
            p, c = ELECTRIC_NEW[tgt]
            oldp, oldc = ws.cell(r, 4).value, ws.cell(r, cd).value
            ws.cell(r, 4).value = p
            ws.cell(r, cd).value = c
            log.append((sheet, r, '改写+数据', '%s|%s|cd%s' % (norm(old), oldp, oldc),
                        '%s|%s|cd%s' % (tgt, p, c)))
        else:
            log.append((sheet, r, '改写', norm(old), tgt))

    # ---- 2. 删 MINI 变体续行（描述空、参数非空）
    lr = last_data_row(ws)
    mini = [r for r in range(5, lr + 1)
            if not norm(ws.cell(r, 3).value) and ws.cell(r, 4).value not in (None, '')]
    for r in sorted(mini, reverse=True):
        log.append((sheet, r, '删MINI续行', str(ws.cell(r, 4).value), ''))
        ws.delete_rows(r)

    # ---- 3. 删漂洗第 2 次分配器（U1，按序号）
    seq = DEL_RINSE_2ND_SEQ.get(sheet)
    if seq is not None:
        cand = rows_of_seq(ws, seq, cd)
        for r in sorted(cand, reverse=True):
            log.append((sheet, r, '删第2次分配器',
                        '%s|%s' % (str(ws.cell(r, 3).value), ws.cell(r, 4).value), ''))
            ws.delete_rows(r)

    # ---- 4. 拆行
    lr = last_data_row(ws)
    for key, rows in SPLIT_ROWS.items():
        hits = [r for r in range(5, lr + 1) if norm(ws.cell(r, 3).value) == key]
        for r in sorted(hits, reverse=True):
            ws.insert_rows(r + 1)
            src, dst = r, r + 1
            for c in range(1, ws.max_column + 1):
                sc, dc_ = ws.cell(src, c), ws.cell(dst, c)
                dc_.value = sc.value
                dc_.number_format = sc.number_format
                dc_.font = copy(sc.font)
                dc_.fill = copy(sc.fill)
                dc_.border = copy(sc.border)
                dc_.alignment = copy(sc.alignment)
            (n1, p1, c1), (n2, p2, c2) = rows
            ws.cell(src, 3).value = n1
            ws.cell(src, 4).value = p1
            ws.cell(src, cd).value = c1
            ws.cell(dst, 3).value = n2
            ws.cell(dst, 4).value = p2
            ws.cell(dst, cd).value = c2
            ws.cell(dst, 1).value = ws.cell(src, 1).value  # 同序号
            log.append((sheet, r, '拆行', key, '%s + %s' % (n1, n2)))
    return log


def ninety_to_eightynine(ws, cfg, sheet):
    """90Mins 业务变更：倒数第二步冷风 1→0；菜单名与各组总时长统一 -1

    「倒数第二步配时 1→0」是所有分组都执行的物理步骤，
    因此每一组的总时长都要 -1，不能只改 Null（否则 A3 反推必然失败）。
    """
    cd = cfg['cd']
    sw0 = cfg['sw0']
    ngrp = cfg['ngrp']
    log = []
    lr = last_data_row(ws)

    # 倒数第二步
    tgt = lr - 1
    desc = norm(ws.cell(tgt, 3).value)
    if '冷风进水排水' in desc:
        old = ws.cell(tgt, cd).value
        ws.cell(tgt, cd).value = 0
        log.append((sheet, tgt, '倒数第二步配时', 'cd%s' % old, 'cd0'))
    else:
        log.append((sheet, tgt, '⚠倒数第二步非冷风', desc, '未改'))

    # 菜单名：第 2 行里任何含 90 的单元格
    for c in range(1, ws.max_column + 1):
        v = ws.cell(2, c).value
        if v is not None and '90' in str(v):
            old = str(v)
            ws.cell(2, c).value = old.replace('90', '89')
            log.append((sheet, 2, '菜单名(%s)' % gl(c), old, str(ws.cell(2, c).value)))

    # 各组总时长统一 -1
    for g in range(ngrp):
        col = sw0 + g
        old_total = ws.cell(4, col).value
        if isinstance(old_total, (int, float)):
            ws.cell(4, col).value = int(old_total) - 1
            log.append((sheet, 4, '总时长%s' % ws.cell(3, col).value,
                        str(old_total), str(int(old_total) - 1)))
    return log


def clear_aux(ws, cfg):
    """清空辅助列（SUM 中转），双列区之后会重建到这里"""
    aux = cfg['aux']
    if aux is None:
        return
    lr = last_data_row(ws)
    for r in range(1, ws.max_row + 1):
        for c in range(aux, ws.max_column + 1):
            cell = ws.cell(r, c)
            cell.value = None
            cell.fill = PatternFill(patternType=None)


def verify(ws, cfg, sheet):
    """A3 裁决：每组总时长是否可由「开关矩阵 × 配时列」反推"""
    cd, sw0, ngrp = cfg['cd'], cfg['sw0'], cfg['ngrp']
    lr = last_data_row(ws)
    out = []
    for g in range(ngrp):
        swc = sw0 + g
        label = ws.cell(3, swc).value
        nominal = ws.cell(4, swc).value
        s = 0
        for r in range(5, lr + 1):
            on = ws.cell(r, swc).value
            if on in (None, ''):
                continue
            v = ws.cell(r, cd).value
            if isinstance(v, (int, float)):
                s += v
        out.append((sheet, label, nominal, s, (s == nominal)))
    return out


def main():
    wb = openpyxl.load_workbook(SRC)
    logs, checks = [], []

    for sheet, cfg in CFG.items():
        ws = wb[sheet]
        if cfg['rewrite']:
            logs += transform_body(ws, cfg, sheet)
        if sheet == '90Mins' and NINETY_TO_EIGHTYNINE:
            logs += ninety_to_eightynine(ws, cfg, sheet)
        clear_aux(ws, cfg)
        checks += verify(ws, cfg, sheet)

    # Q2 标签色：Auto 清红；维护范围内的六个分栏 + Self-cleaning 标绿
    if 'Auto' in wb.sheetnames:
        wb['Auto'].sheet_properties.tabColor = None
    if 'Self-cleaning' in wb.sheetnames:
        wb['Self-cleaning'].sheet_properties.tabColor = TAB_GREEN

    # 90Mins → 89Mins（工作表标签同步改名，放在校验之前）
    old_st = None
    if RENAME_SHEET_90 and '90Mins' in wb.sheetnames:
        old_st = wb['90Mins'].title
        wb['90Mins'].title = '89Mins'
        print('工作表改名：%s -> 89Mins' % old_st)

    # U3：Glass 的 autoFilter 收窄到 M 列（不再覆盖到双列区）
    for name in wb.sheetnames:
        ws = wb[name]
        af = ws.auto_filter
        if af.ref and re.match(r'^A\d+:([A-Z]+)\d+$', af.ref):
            a, _, rest = af.ref.partition(':')
            m = re.match(r'^([A-Z]+)(\d+)$', rest)
            if m and m.group(1) != 'M':
                nrow = a[1:]
                ws.auto_filter.ref = 'A%s:M%s' % (nrow, m.group(2))

    os.makedirs(os.path.dirname(DST), exist_ok=True)
    wb.save(DST)

    print('=== 本体变换完成 ->', DST)
    print('\n--- 变更日志 (%d) ---' % len(logs))
    for x in logs:
        print('  ', x)
    print('\n--- A3 总时长反推校验（计算值 vs 标称）---')
    for c in checks:
        flag = 'OK ' if c[4] else '✗差分'
        print('  [%s] %s / %s : 标称=%s 实算=%s' % (flag, c[0], c[1], c[2], c[3]))


if __name__ == '__main__':
    main()
