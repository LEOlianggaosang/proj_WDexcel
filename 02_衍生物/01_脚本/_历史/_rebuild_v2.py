# -*- coding: utf-8 -*-
"""
柜机时序表 · 第二步重建脚本 v2
================================
基于源表《普通版.xlsx》整体重建，不复用 v1 成品（避免二次污染）。

处理顺序（依赖顺序，不可调换）：
  1. 本体措辞改写（旧措辞 → 新措辞）
  2. 拆行（洗涤分配加热（主洗） → 洗涤分配(主洗) + 洗涤加热（定时加热））
  3. 删行（MINI 变体续行、漂洗第 2 次分配器及其 MINI 行）
  4. 清空辅助列（SUM 中转列，已被双列区取代）
  5. 重建双列区（紧接双速电机列）
  6. 视觉标记（红/蓝字、绿/黄/灰填充）
  7. 分栏标签标绿

规则一律外置为顶部常量，新增情况补一条即可，不改逻辑。
"""
import os
import re
from copy import copy

import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter as gl

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SRC = os.path.join(ROOT, '01_源文件', '普通版.xlsx')
DST = os.path.join(ROOT, '03_验收物', '普通版_第二步已处理.xlsx')

# ---------------------------------------------------------------- 配色
C_RED = 'FFFF0000'      # 红字：定温加热 / 需特别提示
C_BLUE = 'FF00B0F0'     # 蓝字：进水类
C_GREEN = 'FF92D050'    # 绿字：端点
C_YELLOW = 'FFFFFF00'   # 黄填充：加热类改动小
C_GRAY_FILL = Color(theme=0, tint=-0.15)  # 灰填充：该组无此步骤
GREEN_FILL = PatternFill(patternType='solid', fgColor=Color(rgb='FF92D050'))
YELLOW_FILL = PatternFill(patternType='solid', fgColor=Color(rgb='FFFFFF00'))
GRAY_FILL = PatternFill(patternType='solid', fgColor=C_GRAY_FILL)
TAB_GREEN = Color(theme=6, tint=0.0)  # 与 Intensive/Glass/Rapid 现有标签一致

# ---------------------------------------------------------------- 措辞映射
# 旧措辞（去空白后） → 新措辞。None 表示保持原名。
RENAME = {
    '洗涤分配加热（主洗）': ('SPLIT', None),          # 特殊：拆两行
    '洗涤分配加热（漂洗）': '洗涤加热分配（漂洗)分配器：开10秒停8秒累计动作6次',
    '洗涤加热分配（漂洗)': '洗涤加热分配（漂洗)分配器：开10秒停8秒累计动作6次',
    '再生(电子可调)': '再生进水风机（非再生跳步）',
    '再生（电子可调）': '再生进水风机（非再生跳步）',
    '进水(电子可调)': '冷风进水排水（边进边排-进完跳步）（非再生跳步）',
    '进水（电子可调）': '冷风进水排水（边进边排-进完跳步）（非再生跳步）',
}
# 电子可调类改写时，参数与倒计时同步改为模板值
ELECTRIC_NEW = {
    '再生进水风机（非再生跳步）': ('0.5L', 0),
    '冷风进水排水（边进边排-进完跳步）（非再生跳步）': ('2L', 1),
}

# ---------------------------------------------------------------- 分栏配置
# name: (源倒计时列, 双速电机列, 辅助列起始, 分组数, 是否需措辞改写)
CFG = {
    'Intensive ':     dict(cd=12, motor=13, aux=14, ngrp=5, rewrite=False),
    'Glass ':         dict(cd=12, motor=13, aux=None, ngrp=5, rewrite=False),
    'Rapid':          dict(cd=7,  motor=8,  aux=None, ngrp=2, rewrite=False),
    '90Mins':         dict(cd=12, motor=13, aux=14, ngrp=5, rewrite=True),
    'Soak':           dict(cd=6,  motor=7,  aux=None, ngrp=1, rewrite=False),
    'Self-cleaning':  dict(cd=6,  motor=7,  aux=8, ngrp=1, rewrite=True),
}
GROUPS = ['Null', 'ExtraDry', 'Energy(1/2)', 'Express', 'StrWash']
LABEL = {'Energy(1/2)': 'Energy(1/2)'}

# 需删除的「漂洗第 2 次分配器」行（原表行号，从后往前删）
DEL_RINSE_2ND = {'90Mins': [42, 43], 'Self-cleaning': [27, 28]}


def norm(s):
    return re.sub(r'\s+', '', str(s or ''))


def fmt_hmm(minutes):
    return '%d:%02d' % (minutes // 60, minutes % 60)


def skeleton(cell):
    """取与语义着色无关的固定属性"""
    fc = cell.font.color
    dc = None
    if fc is not None:
        dc = fc.rgb if getattr(fc, 'type', None) == 'rgb' else ('theme%s:%s' % (fc.theme, fc.tint))
    return dict(nf=cell.number_format, name=cell.font.name, sz=cell.font.sz,
                b=cell.font.b, i=cell.font.i, u=cell.font.u, strike=cell.font.strike,
                default_color=dc, border=copy(cell.border), align=copy(cell.alignment))


def apply(cell, sk, color=None, fill=None, nf=None):
    final = color if color is not None else sk.get('default_color')
    if isinstance(final, str) and final.startswith('theme'):
        t, _, ti = final[5:].partition(':')
        final = Color(theme=int(t), tint=float(ti))
    cell.number_format = nf if nf is not None else sk['nf']
    cell.font = Font(name=sk['name'], size=sk['sz'], b=sk['b'], i=sk['i'],
                     u=sk['u'], strike=sk['strike'], color=final)
    cell.border = copy(sk['border'])
    cell.alignment = copy(sk['align'])
    cell.fill = fill if fill is not None else PatternFill()


# ================================================================ 1-3 本体变换
def transform_body(ws, cfg, sheet):
    """措辞改写 + 拆行 + 删行。返回 (新行映射, 变更日志)"""
    cd = cfg['cd']
    log = []

    # --- 1. 措辞改写（先做，因为删行判定依赖描述）
    for r in range(5, ws.max_row + 1):
        s = norm(ws.cell(r, 3).value)
        if not s:
            continue
        tgt = RENAME.get(s)
        if tgt is None:
            continue
        if isinstance(tgt, tuple) and tgt[0] == 'SPLIT':
            continue  # 拆行在下一步统一处理
        old = ws.cell(r, 3).value
        ws.cell(r, 3).value = tgt
        if tgt in ELECTRIC_NEW:
            p, c = ELECTRIC_NEW[tgt]
            oldp, oldc = ws.cell(r, 4).value, ws.cell(r, cd).value
            ws.cell(r, 4).value = p
            ws.cell(r, cd).value = c
            log.append((sheet, r, '改写+改数据', '%s|%s' % (old, oldp), '%s|%s' % (tgt, p)))
        else:
            log.append((sheet, r, '改写', str(old), tgt))

    # --- 2. 拆行：洗涤分配加热（主洗） → 洗涤分配(主洗) + 洗涤加热（定时加热）
    splits = [r for r in range(5, ws.max_row + 1) if norm(ws.cell(r, 3).value) == '洗涤分配加热（主洗）']
    for r in sorted(splits, reverse=True):
        ws.insert_rows(r + 1)
        src = r
        dst = r + 1
        # 复制整行样式与格式
        for c in range(1, ws.max_column + 1):
            sc, dc_ = ws.cell(src, c), ws.cell(dst, c)
            dc_.value = sc.value
            dc_.number_format = sc.number_format
            dc_.font = copy(sc.font)
            dc_.fill = copy(sc.fill)
            dc_.border = copy(sc.border)
            dc_.alignment = copy(sc.alignment)
        # 改写两行
        ws.cell(src, 3).value = '洗涤分配(主洗)'
        ws.cell(src, 4).value = 1
        ws.cell(src, cd).value = 0
        ws.cell(dst, 3).value = '洗涤加热（定时加热）'
        ws.cell(dst, 4).value = 90
        ws.cell(dst, cd).value = 1
        log.append((sheet, r, '拆行', '洗涤分配加热（主洗）',
                    '洗涤分配(主洗) + 洗涤加热（定时加热）'))

    # --- 3. 删行（从后往前）：MINI 续行 + 漂洗第 2 次分配器
    del_rows = set()
    for r in range(5, ws.max_row + 1):
        if not norm(ws.cell(r, 3).value) and ws.cell(r, 4).value not in (None, ''):
            del_rows.add(r)  # MINI 变体续行
    for r in DEL_RINSE_2ND.get(sheet, []):
        del_rows.add(r)
    for r in sorted(del_rows, reverse=True):
        ws.delete_rows(r)
        log.append((sheet, r, '删行', 'MINI续行/漂洗第2次分配器', ''))

    return log


def main():
    wb = openpyxl.load_workbook(SRC)
    logs = []
    for sheet, cfg in CFG.items():
        ws = wb[sheet]
        if cfg['rewrite']:
            logs += transform_body(ws, cfg, sheet)
    os.makedirs(os.path.dirname(DST), exist_ok=True)
    wb.save(DST)
    print('本体变换完成 ->', DST)
    print('变更 %d 条：' % len(logs))
    for x in logs:
        print('   ', x)


if __name__ == '__main__':
    main()
