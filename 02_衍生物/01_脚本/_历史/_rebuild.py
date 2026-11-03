# -*- coding: utf-8 -*-
"""
重建 普通版.xlsx 的「已维护双列区」：
  1. Intensive / Glass  补 Express（整块重建，顺序 Null/ExtraDry/Energy/Express/StrWash）
  2. 90Mins             新建 5 组双列区
  3. Self-cleaning      重建 Null 双列（清空原 H 列的 =E*F 公式）

数据规则（从 Intensive / Glass 样板反推）：
  - 每组占 2 列：第1列=累计时长(数字)，第2列=倒计时串「本步耗时/剩余」
  - 第3行写组名(第1列)，第4行写起始倒计时「0/总时长」(第2列)
  - 该组不执行 -> 两列空值；执行且耗时=0 -> 第1列空、第2列「0/剩余」
  - 累计到达总时长 -> 「本步耗时/end」
  - MINI 变体续行是同一工序的参数选项，不产生独立倒计时

视觉规则（用户确认并已逐行验证零例外）：
  - 蓝色 FF00B0F0 = 进水强调
  - 红色 FFFF0000 = 加热强调（仅「定温加热」，定时加热不着红）
  - 绿色 FF92D050 = 首行起始 / 末行 end
  - 灰色填充 theme0 tint-0.15 = 该组无此步骤（含 MINI 变体行）
"""
import json
import os
from copy import copy

import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.colors import Color
from openpyxl.utils import get_column_letter

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SRC = os.path.join(ROOT, '01_源文件', '普通版.xlsx')
DST = os.path.join(ROOT, '03_验收物', '_历史', '普通版_Express已处理.xlsx')
REPORT = os.path.join(ROOT, '04_衍生物', '中间产物', '_diffs.json')

FONT = Font(name='微软雅黑', size=11)

GROUPS = ['Null', 'ExtraDry', 'Energy', 'Express', 'StrWash']
LABEL = {'Energy': 'Energy(1/2)'}

CONFIG = {
    'Intensive ': dict(r0=5, r1=67, sw0=5, cd=12, out=21, nsheets=5),
    'Glass ':     dict(r0=5, r1=58, sw0=5, cd=12, out=14, nsheets=5),
    '90Mins':     dict(r0=5, r1=60, sw0=5, cd=12, out=21, nsheets=5),
    'Self-cleaning': dict(r0=5, r1=38, sw0=5, cd=6,  out=8,  nsheets=1),
}

# Glass 的 Express 标称 115 经交叉验证判定为笔误，真值 105（详见诊断报告）
OVERRIDE_TOTAL = {('Glass ', 'Express'): 105}

# 字体骨架（字体名/字号/边框/对齐）的来源组；Express 借用 Energy
STYLE_SRC = {
    'Intensive ':    {0: 21, 1: 23, 2: 25, 3: 25, 4: 27},
    'Glass ':        {0: 14, 1: 16, 2: 18, 3: 18, 4: 20},
    '90Mins':        {0: 21, 1: 23, 2: 25, 3: 25, 4: 27},
    'Self-cleaning': {0: 8},
}
# 本表无样板时从哪张表借（同一工作簿，样式表共享）
STYLE_FROM = {'90Mins': 'Intensive '}

COLOR_INTAKE = 'FF00B0F0'
COLOR_HEAT = 'FFFF0000'
COLOR_EDGE = 'FF92D050'
NO_FILL = PatternFill(patternType=None)
GRAY_FILL = PatternFill(patternType='solid', fgColor=Color(theme=0, tint=-0.15))

DIFFS = []


def fmt_hmm(minutes):
    minutes = int(minutes)
    return '%d:%02d' % (minutes // 60, minutes % 60)


def is_mini_row(ws, r):
    """同一工序的参数变体行（如 90（大）/ 1（MINI）），A列无序号"""
    if ws.cell(r, 1).value not in (None, ''):
        return False
    d = ws.cell(r, 4).value
    return isinstance(d, str) and 'MINI' in d


def clear_block(ws, c_from, c_to, r_from, r_to):
    for r in range(r_from, r_to + 1):
        for c in range(c_from, c_to + 1):
            cell = ws.cell(r, c)
            cell.value = None
            cell.fill = NO_FILL


def row_color(step_name):
    """按步骤语义决定字体颜色。只认定温加热为红，定时加热保持默认色。"""
    s = str(step_name or '')
    if '进水' in s:
        return COLOR_INTAKE
    if '定温加热' in s:
        return COLOR_HEAT
    return None


def dominant_color(ws_src, sc, rows):
    """取该源列的众数字体色作为「默认色」。避免把行4的绿色误当成基准色。"""
    from collections import Counter
    cnt = Counter()
    for r in rows:
        c = ws_src.cell(r, sc)
        if c.value is None:
            continue
        fc = c.font.color
        if fc is None:
            cnt[None] += 1
        elif getattr(fc, 'type', None) == 'rgb':
            cnt[fc.rgb] += 1
        else:
            cnt[('theme', fc.theme, fc.tint)] += 1
    return cnt.most_common(1)[0][0] if cnt else None


def own_color(cell):
    """单元格自身的字体色（用于组名行等不该回落到众数的位置）"""
    fc = cell.font.color
    if fc is None:
        return None
    if getattr(fc, 'type', None) == 'rgb':
        return fc.rgb
    return ('theme', fc.theme, fc.tint)


def skeleton(cell, default_color=None):
    """字体骨架：只取与语义着色无关的固定属性"""
    return dict(name=cell.font.name, sz=cell.font.sz,
                b=cell.font.b, i=cell.font.i, u=cell.font.u, strike=cell.font.strike,
                default_color=default_color,
                border=copy(cell.border), align=copy(cell.alignment))


def apply(cell, sk, color=None, fill=None, nf='General'):
    """color=None 时回落到该组默认色；nf 需逐行传入（原表同行混用 General/h:mm）"""
    final = color if color is not None else sk.get('default_color')
    if isinstance(final, tuple):
        final = Color(theme=final[1], tint=final[2])
    cell.number_format = nf
    cell.font = Font(name=sk['name'], size=sk['sz'], b=sk['b'], i=sk['i'],
                     u=sk['u'], strike=sk['strike'], color=final)
    cell.border = copy(sk['border'])
    cell.alignment = copy(sk['align'])
    cell.fill = fill if fill is not None else NO_FILL


def rebuild(wb, ws, cfg, sheet):
    r0, r1 = cfg['r0'], cfg['r1']
    groups = GROUPS[:cfg['nsheets']]
    out = cfg['out']
    c_to = out + len(groups) * 2 - 1
    ws_src = wb[STYLE_FROM.get(sheet, sheet)]
    sw0, cd = cfg['sw0'], cfg['cd']

    # ① 骨架快照必须早于清值：Intensive 的 Null 组源列与目标列完全重叠
    data_rows = list(range(r0, r1 + 1))
    base, head3, nfmap = {}, {}, {}
    for gi in range(len(groups)):
        bc = STYLE_SRC[sheet][gi]
        for off in (0, 1):
            sc = bc + off
            # 样本必须取普通数据行：第3/4行是组名行与起始行，样式特殊，不能当基准
            sample = None
            for r in data_rows:
                if ws_src.cell(r, sc).value is not None:
                    sample = ws_src.cell(r, sc)
                    break
            if sample is None:
                sample = ws_src.cell(r0, sc)
            base[(gi, off)] = skeleton(sample, dominant_color(ws_src, sc, data_rows))
            # number_format 逐行搬运：原表同行混用 General / h:mm，是手工编辑痕迹
            nfmap[(gi, off)] = {r: ws_src.cell(r, sc).number_format
                                for r in [3, 4] + data_rows}
            head3[(gi, off)] = skeleton(ws_src.cell(3, sc), own_color(ws_src.cell(3, sc)))

    clear_block(ws, out, c_to, 3, ws.max_row)
    mini_rows = [r for r in range(r0, r1 + 1) if is_mini_row(ws, r)]

    for gi, gname in enumerate(groups):
        sw_col = sw0 + gi
        c1 = out + gi * 2
        c2 = c1 + 1

        total = ws.cell(4, sw_col).value
        total = int(total)
        final_total = OVERRIDE_TOTAL.get((sheet, gname), total)
        if final_total != total:
            DIFFS.append(dict(sheet=sheet, group=gname, kind='总时长修正',
                              old=total, new=final_total,
                              note='第4行标称与实求和不一致，已按交叉验证结果修正'))
        total = final_total

        ws.cell(3, c1).value = LABEL.get(gname, gname)
        apply(ws.cell(3, c1), head3[(gi, 0)], nf=nfmap[(gi, 0)].get(3, 'General'))

        ws.cell(4, c2).value = '0/' + fmt_hmm(total)
        apply(ws.cell(4, c2), base[(gi, 1)], color=COLOR_EDGE, nf=nfmap[(gi, 1)].get(4, 'General'))

        cum = 0
        step_color = {}
        for r in range(r0, r1 + 1):
            if r not in step_color:
                step_color[r] = row_color(ws.cell(r, 3).value)
            col = step_color[r]
            if r in mini_rows:
                apply(ws.cell(r, c1), base[(gi, 0)], fill=GRAY_FILL, nf=nfmap[(gi, 0)].get(r, 'General'))
                apply(ws.cell(r, c2), base[(gi, 1)], fill=GRAY_FILL, nf=nfmap[(gi, 1)].get(r, 'General'))
                continue
            sw = ws.cell(r, sw_col).value
            if sw in (None, ''):
                apply(ws.cell(r, c1), base[(gi, 0)], fill=GRAY_FILL, nf=nfmap[(gi, 0)].get(r, 'General'))
                apply(ws.cell(r, c2), base[(gi, 1)], fill=GRAY_FILL, nf=nfmap[(gi, 1)].get(r, 'General'))
                continue
            L = ws.cell(r, cd).value
            L = 0 if L in (None, '') else int(L)
            rem_before = total - cum
            head = fmt_hmm(rem_before) if rem_before > 0 else 'end'
            if L == 0:
                ws.cell(r, c2).value = '0/' + head
                apply(ws.cell(r, c2), base[(gi, 1)], color=col, nf=nfmap[(gi, 1)].get(r, 'General'))
                continue
            cum += L
            rem = total - cum
            is_end = rem <= 0
            ws.cell(r, c1).value = cum
            apply(ws.cell(r, c1), base[(gi, 0)], nf=nfmap[(gi, 0)].get(r, 'General'))
            ws.cell(r, c2).value = '%d/%s' % (L, 'end' if is_end else fmt_hmm(rem))
            apply(ws.cell(r, c2), base[(gi, 1)], color=(COLOR_EDGE if is_end else col),
                  nf=nfmap[(gi, 1)].get(r, 'General'))

        DIFFS.append(dict(sheet=sheet, group=gname, kind='双列生成',
                          total=total, end_value=cum,
                          note='累计终值 %d %s 总时长' % (cum, '==' if cum == total else '!=')))
        if cum != total:
            DIFFS.append(dict(sheet=sheet, group=gname, kind='!! 累计与总时长不符',
                              old=total, new=cum, note='需人工复核'))

    # ② 列宽：按源列搬运
    for gi in range(len(groups)):
        bc = STYLE_SRC[sheet][gi]
        for off in (0, 1):
            w = ws_src.column_dimensions[get_column_letter(bc + off)]
            dstL = get_column_letter(out + gi * 2 + off)
            ws.column_dimensions[dstL].width = \
                w.width if w.width is not None else (10.0 if off == 0 else 13.0)
            ws.column_dimensions[dstL].hidden = bool(w.hidden)

    return mini_rows


def fix_helper_mini(ws, cfg, mini_rows, sheet):
    """修正 N–T 辅助列：MINI 变体行被重复计入，导致 SUM 比标称大 2"""
    if cfg.get('out') == 8:
        return []
    cols = list(range(14, 21))  # N..T
    removed = []
    for r in mini_rows:
        vals = [ws.cell(r, c).value for c in cols]
        for c, v in zip(cols, vals):
            if isinstance(v, str) and v.startswith('='):
                ws.cell(r, c).value = None
                removed.append('%s!%s%d' % (sheet, get_column_letter(c), r))
    return removed


def main():
    wb = openpyxl.load_workbook(SRC)
    all_mini = {}
    for sheet, cfg in CONFIG.items():
        ws = wb[sheet]
        mini_rows = rebuild(wb, ws, cfg, sheet)
        all_mini[sheet] = mini_rows
        removed = fix_helper_mini(ws, cfg, mini_rows, sheet)
        if removed:
            DIFFS.append(dict(sheet=sheet, group='(辅助列 N–T)', kind='MINI 续行去重',
                              note='清除 %d 个单元格公式，SUM 恢复为标称值' % len(removed),
                              cells=removed[:12]))

    ws = wb['Glass ']
    if ws['S1'].value is not None:
        DIFFS.append(dict(sheet='Glass ', group='-', kind='清除游离值',
                          old=ws['S1'].value, new=None, note='S1 孤立的 0'))
        ws['S1'].value = None

    os.makedirs(os.path.dirname(DST), exist_ok=True)
    wb.save(DST)
    with open(REPORT, 'w', encoding='utf-8') as f:
        json.dump(dict(diffs=DIFFS, mini={k: v for k, v in all_mini.items()}),
                  f, ensure_ascii=False, indent=2)
    print('saved ->', DST)
    for d in DIFFS:
        print('  ', d['sheet'], '|', d['group'], '|', d['kind'], '|', d.get('note', ''))


if __name__ == '__main__':
    main()
