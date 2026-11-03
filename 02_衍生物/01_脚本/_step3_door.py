# -*- coding: utf-8 -*-
"""
柜机时序表 · 第三步：自动门版（第二份成品）
============================================
输入：03_验收物/普通版_第二步已处理.xlsx   ← 第一份成品，冻结只读（B4 红线）
输出：03_验收物/自动门_第三步已处理.xlsx   ← 第二份成品，独立文件，绝不覆盖第一份

规则源：02_衍生物/00_知识/第三步_自动门规则.md
方案源：02_衍生物/02_中间产物/再生间隔_改法方案.md（v3 定稿，2026-11-03）

策略组（2026-11-03 设计更新）：
  · 再生间隔策略：从干燥段风机拆 N 分钟改为「静止（再生间隔）」，当前 N=10 分钟（600s）
    —— 取法为「开关公因式提取」：X 必须是所有分组都有的风机时长，重组为
       满开关一步 + 余下分开关步；若无满开关风机步，用
       A∪B=全开关、A∩B={ED} 的恒等式重组（89Mins 即此情形）
  · 自动门再生前移策略：把再生间隔再拆为 静止（再生间隔）(600−120) + 自动开门 120（固定 2 分钟）
    顺序：0.5L → 再生间隔 → 前移步骤 → 自动开门（先 0.5l 再 再生间隔 再 前移 再 自动开门）
  · 自动门开关切换策略：自动开门 → 「自动开门/静止」；风机 → 「自动开门静止/风机」；
    底部红字备注「自动门开关切换：自动开门/静止、自动开门静止/风机」

要点：
  · 改的四张（Intensive / Glass / 89Mins / Self-cleaning）：干燥段重排 + 描述改写 + 新标记 + 底部备注
  · Rapid / Soak：内容不改，仅清除第二步的更改强调（红字与填充）
  · 标记只作用于本体区 A..motor，与双列区无关（Q3）
  · 守恒：原风机总时长 = 静止（再生间隔）480 + 自动开门 120 + 剩余风机
  · 89Mins +2 分钟特殊处理已取消（公因式提取后各组天然守恒）
"""
import os
import re

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock, InlineFont
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter as gl

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SRC = os.path.join(ROOT, '03_验收物', '普通版_第二步已处理.xlsx')      # 第一份成品（只读）
DST = os.path.join(ROOT, '03_验收物', '自动门_第三步已处理.xlsx')      # 第二份成品

GREEN, YELLOW, RED, BLUE = 'FF92D050', 'FFFFFF00', 'FFFF0000', 'FF00B0F0'
THIN = Side(style='thin')
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NO_BORDER = Border()
NO_FILL = PatternFill(patternType=None)
GRAY_FILL = PatternFill(patternType='solid', fgColor=Color(theme=0, tint=-0.15))
FONT = Font(name='微软雅黑', size=10)

CFG = {
    'Intensive ':    dict(sw0=5, cd=12, motor=13, r0=5, ngrp=5, tot_row=4, swn=7),
    'Glass ':        dict(sw0=5, cd=12, motor=13, r0=5, ngrp=5, tot_row=4, swn=7),
    '89Mins':        dict(sw0=5, cd=12, motor=13, r0=5, ngrp=5, tot_row=4, swn=7),
    'Rapid':         dict(sw0=5, cd=7,  motor=8,  r0=5, ngrp=2, tot_row=4, swn=2),
    'Soak':          dict(sw0=5, cd=6,  motor=7,  r0=4, ngrp=1, tot_row=3, swn=1),
    'Self-cleaning': dict(sw0=5, cd=6,  motor=7,  r0=5, ngrp=1, tot_row=4, swn=1),
}
DOOR = ['Intensive ', 'Glass ', '89Mins', 'Self-cleaning']   # 要改的四张
PASS = ['Rapid', 'Soak']                                      # 仅清标记

# ============================================================ 富文本（主文红 + 括号段红）
RT_RW = [('再生进水\n', RED), ('（非再生跳步）', RED)]
RT_ID = [('进水排水\n', RED), ('（边进边排-进完跳步）', RED), ('\n', RED), ('（非再生跳步）', RED)]

# ============================================================ 干燥段新排布（唯一数据源，改这里即可回滚）
# 顺序：热漂后静置 → 排水 → 再生进水0.5L → 静止（再生间隔）480 → 前移步骤 → 自动开门/静止120 → 风机
# sw：'all'=全组全开 | int=沿用原干燥段第 n 行的开关矩阵 | ('or',a,b)=并集 | ('and',a,b)=交集
# mark：None=无 | 'Y'=黄（前移区） | 'G'=绿（后移区）；red=True → C 列描述整格红字
#
# 索引对照（输入＝含「静止（再生间隔）」预留行的成品一干燥段）：
#   Intensive   0静止600 1排水45 2再生0.5L 3预留0[all] 4风机2100[all] 5风机1200[ED] 6~9 前移
#   Glass       0静止600 1排水45 2再生0.5L 3预留0[all] 4风机1200[all] 5风机1500[5组]
#               6风机1200[Energy] 7风机1200[ED] 8~11 前移
#   89Mins      0排水30 1再生0.5L 2预留0[all] 3风机900[6组] 4风机600[ED,Express] 5风机600[ED] 6~9 前移
#   Self-clean  0排水30 1再生0.5L 2预留0[all] 3风机1200[all] 4~7 前移
DRY = {
    'Intensive ': [
        dict(desc='静止', param=600, cd=10, sw=0, mark=None, red=False),
        dict(desc='排水', param=45, cd=1, sw=1, mark=None, red=False),
        dict(desc='再生进水（非再生跳步）', param='0.5L', cd=0, sw=2, mark='Y', red=True, rt=RT_RW),
        dict(desc='静止（再生间隔）', param=480, cd=8, sw=3, mark='Y', red=True),
        dict(desc='进水排水（边进边排-进完跳步）（非再生跳步）', param='2L', cd=1, sw=6,
             mark='Y', red=True, rt=RT_ID),
        dict(desc='排水', param=10, cd=0, sw=7, mark='Y', red=False),
        dict(desc='进水排水（边进边排-进完跳步）（非再生跳步）', param='2L', cd=1, sw=8,
             mark='Y', red=True, rt=RT_ID),
        dict(desc='排水', param=45, cd=1, sw=9, mark='Y', red=False),
        dict(desc='自动开门/静止', param=120, cd=2, sw='all', mark='G', red=True),
        dict(desc='自动开门静止/风机', param=1500, cd=25, sw=4, mark='G', red=True),
        dict(desc='自动开门静止/风机', param=1200, cd=20, sw=5, mark='G', red=True),
    ],
    'Glass ': [
        dict(desc='静止', param=600, cd=10, sw=0, mark=None, red=False),
        dict(desc='排水', param=45, cd=1, sw=1, mark=None, red=False),
        dict(desc='再生进水（非再生跳步）', param='0.5L', cd=0, sw=2, mark='Y', red=True, rt=RT_RW),
        dict(desc='静止（再生间隔）', param=480, cd=8, sw=3, mark='Y', red=True),
        dict(desc='进水排水（边进边排-进完跳步）（非再生跳步）', param='2L', cd=1, sw=8,
             mark='Y', red=True, rt=RT_ID),
        dict(desc='排水', param=10, cd=0, sw=9, mark='Y', red=False),
        dict(desc='进水排水（边进边排-进完跳步）（非再生跳步）', param='2L', cd=1, sw=10,
             mark='Y', red=True, rt=RT_ID),
        dict(desc='排水', param=45, cd=1, sw=11, mark='Y', red=False),
        dict(desc='自动开门/静止', param=120, cd=2, sw='all', mark='G', red=True),
        dict(desc='自动开门静止/风机', param=600, cd=10, sw='all', mark='G', red=True),
        dict(desc='自动开门静止/风机', param=1500, cd=25, sw=5, mark='G', red=True),
        dict(desc='自动开门静止/风机', param=1200, cd=20, sw=6, mark='G', red=True),
        dict(desc='自动开门静止/风机', param=1200, cd=20, sw=7, mark='G', red=True),
    ],
    '89Mins': [
        dict(desc='排水', param=30, cd=0, sw=0, mark=None, red=False),
        dict(desc='再生进水（非再生跳步）', param='0.5L', cd=0, sw=1, mark='Y', red=True, rt=RT_RW),
        dict(desc='静止（再生间隔）', param=480, cd=8, sw=2, mark='Y', red=True),
        dict(desc='进水排水（边进边排-进完跳步）（非再生跳步）', param='2L', cd=1, sw=6,
             mark='Y', red=True, rt=RT_ID),
        dict(desc='排水', param=10, cd=0, sw=7, mark='Y', red=False),
        dict(desc='进水排水（边进边排-进完跳步）（非再生跳步）', param='2L', cd=0, sw=8,
             mark='Y', red=True, rt=RT_ID),
        dict(desc='排水', param=45, cd=1, sw=9, mark='Y', red=False),
        dict(desc='自动开门/静止', param=120, cd=2, sw='all', mark='G', red=True),
        # 公因式提取：600·A + 600·B = 600·全开关 + 600·{ED}
        #   风机900[A=6组] −600 → 300[A]；风机600[B=ED,Express] −600 → 0（行删除）
        #   风机600[C=ED] 吸收交集部分 → 1200[ED]（原两行 ED 合并）
        dict(desc='自动开门静止/风机', param=300, cd=5, sw=3, mark='G', red=True),
        dict(desc='自动开门静止/风机', param=1200, cd=20, sw=5, mark='G', red=True),
    ],
    'Self-cleaning': [
        dict(desc='排水', param=30, cd=0, sw=0, mark=None, red=False),
        dict(desc='再生进水（非再生跳步）', param='0.5L', cd=0, sw=1, mark='Y', red=True, rt=RT_RW),
        dict(desc='静止（再生间隔）', param=480, cd=8, sw=2, mark='Y', red=True),
        dict(desc='进水排水（边进边排-进完跳步）（非再生跳步）', param='2L', cd=1, sw=4,
             mark='Y', red=True, rt=RT_ID),
        dict(desc='排水', param=10, cd=0, sw=5, mark='Y', red=False),
        dict(desc='进水排水（边进边排-进完跳步）（非再生跳步）', param='2L', cd=1, sw=6,
             mark='Y', red=True, rt=RT_ID),
        dict(desc='排水', param=45, cd=1, sw=7, mark='Y', red=False),
        dict(desc='自动开门/静止', param=120, cd=2, sw='all', mark='G', red=True),
        dict(desc='自动开门静止/风机', param=600, cd=10, sw=3, mark='G', red=True),
    ],
}

# 89Mins Express +2 分钟特殊处理已取消（公因式提取后天然守恒）
EXPRESS_TOTAL = {}
NOTE = '自动门开关切换：自动开门/静止、自动开门静止/风机'

LOG = []


def say(*a):
    s = ' '.join(str(x) for x in a)
    print(s)
    LOG.append(s)


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
    return st


def make_rt(blocks):
    return CellRichText(*[TextBlock(InlineFont(rFont='微软雅黑', color=Color(rgb=c)), t)
                          for t, c in blocks])


def fmt_hmm(m):
    m = int(max(m, 0))
    return '%d:%02d' % (m // 60, m % 60)


# ================================================================ 【1】清除第二步的更改强调（Q3）
def clear_marks(wb):
    say('\n【1】清除第二步更改强调（六表本体区 A..motor：填充 + 单元格红字 + 富文本红块）')
    say('    （V27 2026-11-03 用户纠错：定温加热行不再跳过——倒计时红随描述在【5】重生，')
    say('     与本体区标记无关；富文本红块属第二步更改强调，一并转纯文本清除）')
    tot = 0
    rt = 0
    for n in CFG:
        if n not in wb.sheetnames:
            continue
        ws = wb[n]
        cfg = CFG[n]
        rows = data_rows(ws, cfg['r0'])
        for r in range(cfg['r0'], (rows[-1] if rows else cfg['r0']) + 1):
            for c in range(1, cfg['motor'] + 1):
                cell = ws.cell(r, c)
                # 富文本（V8 重建行）里的红块：cell.font.color 查不到，必须整体转纯文本
                if isinstance(cell.value, CellRichText):
                    cell.value = str(cell.value)
                    fo = cell.font
                    cell.font = Font(name=fo.name or '微软雅黑', size=fo.size or 10,
                                     bold=fo.bold, italic=fo.italic)
                    rt += 1
                f = cell.fill
                if f and f.patternType:
                    cell.fill = NO_FILL
                    tot += 1
                fc = cell.font.color
                if fc and fc.type == 'rgb' and str(fc.value) == RED:
                    fo = cell.font
                    cell.font = Font(name=fo.name or '微软雅黑', size=fo.size or 10,
                                     bold=fo.bold, italic=fo.italic)
                    tot += 1
        # 底部第二步红字备注（若有残留）一并清除
        if rows:
            for r in range(rows[-1] + 1, ws.max_row + 1):
                v = ws.cell(r, 1).value
                if v not in (None, ''):
                    fo = ws.cell(r, 1).font
                    if fo.color and fo.color.type == 'rgb' and str(fo.color.value) == RED \
                            and '自动门' not in norm(v):
                        ws.cell(r, 1).value = None
                        tot += 1
    say('  清除填充/红字 %d 处；富文本转纯文本 %d 格' % (tot, rt))


# ================================================================ 【2】干燥段重建
def rebuild_dry(wb):
    say('\n【2】干燥段重建（四表；守恒：原风机 = 静止（再生间隔）480 + 自动开门120 + 剩余风机）')
    for n in DOOR:
        ws = wb[n]
        cfg = CFG[n]
        plan = DRY[n]
        swn = cfg['swn']
        st = stage_map(ws, cfg['r0'])
        rows = [r for r in data_rows(ws, cfg['r0'])
                if st.get(r) and '干燥' in str(st[r])]
        if not rows:
            say('  %-14s ⚠未找到干燥段' % n)
            continue
        ds, lr = rows[0], rows[-1]

        # 保存原干燥段（开关/行高）
        old = []
        for r in rows:
            old.append(dict(
                sw=[ws.cell(r, cfg['sw0'] + g).value for g in range(swn)],
                h=ws.row_dimensions[r].height))

        need = len(plan)
        if need < len(rows):
            say('  %-14s ⚠计划 %d 行 < 现有 %d 行，需人工确认' % (n, need, len(rows)))
            continue

        # —— 必须先拆合并区再插行：openpyxl 的 insert_rows 会挪动合并区，
        #    插完再拆会落空，随后写单元格就会撞上只读的 MergedCell
        for m in list(ws.merged_cells.ranges):
            if m.min_col == 2 and m.min_row >= ds:
                ws.unmerge_cells(str(m))
        for r in range(ds, lr + 1):
            ws.cell(r, 2).value = None

        if need > len(rows):
            ws.insert_rows(ds, need - len(rows))
            say('  %-14s 插入 %d 行（干燥段 %d → %d 行）' % (n, need - len(rows), len(rows), need))

        newlr = ds + need - 1
        for i, p in enumerate(plan):
            r = ds + i
            # 开关
            sw = p['sw']
            if sw == 'all':
                vals = [1] * swn
            elif isinstance(sw, tuple) and sw[0] == 'or':
                vals = [1 if any(old[k]['sw'][g] for k in sw[1:]) else None for g in range(swn)]
            elif isinstance(sw, tuple) and sw[0] == 'and':
                vals = [1 if all(old[k]['sw'][g] for k in sw[1:]) else None for g in range(swn)]
            else:
                vals = list(old[sw]['sw'])
            for g, v in enumerate(vals):
                ws.cell(r, cfg['sw0'] + g).value = v

            # 内容
            ws.cell(r, 3).value = make_rt(p['rt']) if p.get('rt') else p['desc']
            ws.cell(r, 4).value = p['param']
            ws.cell(r, cfg['cd']).value = p['cd']
            ws.cell(r, cfg['motor']).value = '静止'

            # 标记（只在本体区）
            fil = {'Y': YELLOW, 'G': GREEN}.get(p['mark'])
            for c in range(1, cfg['motor'] + 1):
                cell = ws.cell(r, c)
                cell.fill = PatternFill('solid', fgColor=Color(rgb=fil)) if fil else NO_FILL
                cell.font = Font(name='微软雅黑', size=10,
                                 color=Color(rgb=RED) if p['red'] and c == 3 else None)
                cell.border = THIN_BORDER
                cell.alignment = Alignment(horizontal='center', vertical='center',
                                           wrap_text=True if c == 3 else False)
            h = old[sw]['h'] if isinstance(sw, int) else (old[0]['h'] if old else None)
            if h is not None:
                ws.row_dimensions[r].height = h

        # B 列干燥合并区重建（此时旧合并区已拆、值已清）
        for r in range(ds, newlr + 1):
            ws.cell(r, 2).value = None
            ws.cell(r, 2).border = THIN_BORDER
        ws.cell(ds, 2).value = '干燥'
        ws.cell(ds, 2).alignment = Alignment(horizontal='center', vertical='center')
        if newlr > ds:
            ws.merge_cells(start_row=ds, start_column=2, end_row=newlr, end_column=2)

        say('  %-14s r%d–r%d 干燥段 %d 行；静止（再生间隔）480 + 自动开门/静止 120'
            % (n, ds, newlr, need))


# ================================================================ 【3】序号重排
def renumber(wb):
    say('\n【3】序号重排（四表）')
    for n in DOOR:
        ws = wb[n]
        cfg = CFG[n]
        for m in list(ws.merged_cells.ranges):
            if m.min_col == 1 and m.min_row >= cfg['r0']:
                ws.unmerge_cells(str(m))
        rows = data_rows(ws, cfg['r0'])
        for i, r in enumerate(rows, 1):
            ws.cell(r, 1).value = i
            ws.cell(r, 1).alignment = Alignment(horizontal='center', vertical='center')
        say('  %-14s 1..%d' % (n, len(rows)))


# ================================================================ 【4】标称修正 + 底部备注
def totals_and_note(wb):
    say('\n【4】标称校验（不再需要 +2 修正）+ 底部备注')
    for n, (g, v) in EXPRESS_TOTAL.items():
        ws = wb[n]
        cfg = CFG[n]
        c = cfg['sw0'] + g
        old = ws.cell(cfg['tot_row'], c).value
        ws.cell(cfg['tot_row'], c).value = v
        say('  %-14s %s组 %s → %s' % (n, ws.cell(3, c).value, old, v))
    for n in DOOR:
        ws = wb[n]
        rows = data_rows(ws, CFG[n]['r0'])
        r = rows[-1] + 1
        ws.cell(r, 1).value = NOTE
        ws.cell(r, 1).font = Font(name='微软雅黑', size=10, color=Color(rgb=RED))
        say('  %-14s r%d 备注：%s' % (n, r, NOTE))


# ================================================================ 【5】双列区重建（按新本体）
def build_dual(wb):
    say('\n【5】双列区重建（按新本体重算累计/倒计时）')
    for n in CFG:
        if n not in wb.sheetnames:
            continue
        ws = wb[n]
        cfg = CFG[n]
        out0, ngrp, r0 = cfg['motor'] + 1, cfg['ngrp'], cfg['r0']
        rows = data_rows(ws, r0)
        if not rows:
            continue
        lr = rows[-1]
        for c in range(out0, ws.max_column + 1):
            for r in range(1, ws.max_row + 1):
                cell = ws.cell(r, c)
                cell.value = None
                cell.font = FONT
                cell.fill = NO_FILL
                cell.border = NO_BORDER
        for g in range(ngrp):
            swc = cfg['sw0'] + g
            c1 = out0 + g * 2
            c2 = c1 + 1
            label = ws.cell(3, swc).value
            label = str(label).replace('\n', '') if label not in (None, '') else 'Null'
            total = ws.cell(cfg['tot_row'], swc).value
            ws.cell(r0 - 2, c1).value = None
            ws.cell(r0 - 2, c1).border = THIN_BORDER
            ws.cell(r0 - 2, c2).value = label
            ws.cell(r0 - 2, c2).font = FONT
            ws.cell(r0 - 2, c2).border = THIN_BORDER
            if not isinstance(total, (int, float)):
                continue
            total = int(total)
            ws.cell(r0 - 1, c1).border = THIN_BORDER
            ws.cell(r0 - 1, c2).border = THIN_BORDER
            ws.cell(r0 - 1, c2).value = '0/%s' % fmt_hmm(total)
            ws.cell(r0 - 1, c2).font = Font(name='微软雅黑', size=10, color=GREEN)

            cum = 0
            seq = []
            for r in rows:
                on = ws.cell(r, swc).value not in (None, '')
                v = ws.cell(r, cfg['cd']).value
                v = v if isinstance(v, (int, float)) else 0
                if on:
                    cum += v
                seq.append((r, on, cum))

            # 终行必须是「该组最后执行的行」，不是全表最后一行：
            # 后移的自动开门静止/风机只对部分组开，末行对多数组不执行，
            # 按全表末行判 is_last 会让这些组的 '1/end' 丢失（第二步遗留缺陷，本步修正）
            last_on = None
            for i, (r, on, c) in enumerate(seq):
                if on:
                    last_on = i

            # V9 颜色 + V10 去框线
            for i, (r, on, c) in enumerate(seq):
                v = ws.cell(r, cfg['cd']).value
                v = v if isinstance(v, (int, float)) else 0
                is_last = (i == last_on)
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
                    ws.cell(r, c2).value = ('%d/end' % v) if is_last \
                        else '%d/%s' % (v, fmt_hmm(total - c))
                d = norm(ws.cell(r, 3).value)
                v2 = ws.cell(r, c2).value
                # V26 颜色规则（用户 2026-11-03 纠错）——只给倒计时列 c2 上色，累计列 c1 恒黑：
                #   绿 = end；蓝 = 描述含「进水」（实机进水提醒开水）；
                #   红 = 描述含「定温加热」（实机加热提醒变温），不随黄填充
                col = None
                if v2 in (None, ''):
                    col = None
                elif is_last or str(v2).endswith('/end'):
                    col = GREEN
                elif '进水' in d:
                    col = BLUE
                elif '定温加热' in d:
                    col = RED
                for cc in (c1, c2):
                    cell = ws.cell(r, cc)
                    c_color = Color(rgb=col) if (col and cc == c2) else None
                    cell.font = Font(name='微软雅黑', size=10, color=c_color)
                    cell.border = NO_BORDER if r > cfg['tot_row'] else THIN_BORDER
                    cell.alignment = Alignment(horizontal='center', vertical='center')
        say('  %-14s 双列区 %d 组' % (n, ngrp))


# ================================================================ 【6】终检
def verify(wb):
    say('\n【6】终检')
    ok = True
    for n in CFG:
        if n not in wb.sheetnames:
            continue
        ws = wb[n]
        cfg = CFG[n]
        rows = data_rows(ws, cfg['r0'])
        # A3 守恒
        bad = []
        for g in range(cfg['ngrp']):
            swc = cfg['sw0'] + g
            total = ws.cell(cfg['tot_row'], swc).value
            if not isinstance(total, (int, float)):
                continue
            s = 0
            for r in rows:
                if ws.cell(r, swc).value not in (None, ''):
                    v = ws.cell(r, cfg['cd']).value
                    s += v if isinstance(v, (int, float)) else 0
            if s != int(total):
                bad.append('%s组 %d≠%d' % (ws.cell(3, swc).value, s, int(total)))
        # 序号连续
        seqbad = [i for i, r in enumerate(rows, 1) if ws.cell(r, 1).value != i]
        # 自动开门
        door = [r for r in rows if '自动开门' in norm(ws.cell(r, 3).value)]
        say('  %-14s 守恒%s 序号%s 自动开门%d行 %s' % (
            n, '✅' if not bad else '❌' + str(bad),
            '✅' if not seqbad else '❌', len(door),
            '备注%s' % ('✅' if ws.cell(rows[-1] + 1, 1).value == NOTE else '—')
            if n in DOOR else ''))
        ok = ok and not bad and not seqbad
        if n in DOOR:
            ok = ok and bool(door) and ws.cell(rows[-1] + 1, 1).value == NOTE
    say('  总判定：%s' % ('✅ 全部通过' if ok else '❌ 有未通过项'))
    return ok


def main():
    wb = openpyxl.load_workbook(SRC, rich_text=True)
    say('第三步 · 自动门版')
    say('输入（只读）：%s' % SRC)
    say('输出：%s' % DST)

    clear_marks(wb)
    rebuild_dry(wb)
    renumber(wb)
    totals_and_note(wb)
    build_dual(wb)
    for n in CFG:
        if n in wb.sheetnames:
            wb[n].sheet_properties.tabColor = GREEN
    ok = verify(wb)

    os.makedirs(os.path.dirname(DST), exist_ok=True)
    wb.save(DST)
    say('\n已保存 -> %s' % DST)
    return 0 if ok else 1


if __name__ == '__main__':
    raise SystemExit(main())
