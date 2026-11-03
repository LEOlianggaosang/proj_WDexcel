# -*- coding: utf-8 -*-
"""
第二步验收修复（2026-11-02 用户验收反馈 13 项）
================================================
输入/输出：03_验收物/普通版_第二步已处理.xlsx（原地改写；源文件永不写）
前置：_step2_body.py → _step2_cols.py → _step2_fill.py 已跑完

验收裁决台账（来源：用户 2026-11-02 验收反馈）
  V1  双列区视觉 OK（组名第 2 列、起值行绿字——保持）
  V2  温达跳步标法以 Glass 为主（C–D 无红；成品现状已一致，本脚本只校验）
  V3  Self-cleaning 静止 1200 照 89Mins 特例：黄填充 C–D + 红字 C,D
  V4  序号重排（删行后跳号），字体统一
  V5  新增行字体统一（宋体 11→微软雅黑 10）、框线补全
  V6  B 列阶段重新合并并框起来（delete_rows 后合并错位）
  V7  取消全部隐藏行/列（含范围外表——用户明示"所有"）
  V8  C 列行内格式重建：回车 + 行内红字 + wrap + 行高 33（参照源模板）
  V9  双列区红蓝字恢复：倒计时列颜色跟随本体行（进水→蓝 / 黄填充→红 / 端点→绿）
  V10 双列区数据行去框线（仅名称行 r3 与起值行 r4 保留）
  V11 半载三列：N 列隐藏残留（旧辅助列）→ 取消；清 Y/Z/IZ 残留列宽；双列区列宽统一
  V12 Self-cleaning autoFilter 收缩到 G（原 A4:IV38 被误缩到 M，覆盖双列区）
  V13 Rapid 注释删除（"半载功能只适用于WHR客户"×2）
  V14 风机描述清除"（非交替）"字样
  V15 漂洗分配参数 '60（大）' → 60（只留数字，对齐模板）
  V16 干燥前最后一次加热做成温达跳步（89Mins r47、Self-cleaning r27：参数→纯T，cd 不变）
  V18 89Mins 漂洗第一次加热：定时加热 480 → 温达剩洗 T60/480（用户裁决 2026-11-02 复验轮）
      谱系：旧版本(45机) 90Mins 该行原为 T65/480 温达剩洗，恢复60机时被误改成定时加热；
      裁决温度 T60、总时长 480 不变、配时 8 不变；标法照 r16 温达剩洗（黄 C–M + 红字 D–M）
  V20 干燥段「静止拆分策略」回退（用户 2026-11-03 设计更新）
      · 废止「干燥从风机改为静止 + 静止拆分」策略
      · 干燥段（再生之后、冷风之前）的「静止」一律改回「风机」，时长/配时不变
      · 改回后相邻且开关矩阵完全相同的两步重新合并
      · 回退涉及的红字/黄填充一律清除（描述改回后与旧版本一致）
      · 静止拆分策略 → 再生间隔策略：从干燥段风机拆 N 分钟改为「静止（再生间隔）」
        REGEN_INTERVAL_MIN = 0（可调），当前不拆 → 干燥段不产生新标记
      · V3（Self-cleaning 静止特例）因前提消失而失效，已停用
"""
import os, re, difflib
from copy import deepcopy
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.styles.colors import Color
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.cell.text import InlineFont
from openpyxl.utils import get_column_letter as gl

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DST = os.path.join(ROOT, '03_验收物', '普通版_第二步已处理.xlsx')
SRC = os.path.join(ROOT, '01_源文件', '普通版.xlsx')

GREEN, YELLOW, RED, BLUE, BLACK = 'FF92D050', 'FFFFFF00', 'FFFF0000', 'FF00B0F0', 'FF000000'
THIN = Side(style='thin')
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
NO_BORDER = Border()
FONT = Font(name='微软雅黑', size=10)

CFG = {
    'Intensive ':    dict(sw0=5, cd=12, motor=13, r0=5, ngrp=5, tot_row=4),
    'Glass ':        dict(sw0=5, cd=12, motor=13, r0=5, ngrp=5, tot_row=4),
    '89Mins':        dict(sw0=5, cd=12, motor=13, r0=5, ngrp=5, tot_row=4),
    'Rapid':         dict(sw0=5, cd=7,  motor=8,  r0=5, ngrp=2, tot_row=4),
    'Soak':          dict(sw0=5, cd=6,  motor=7,  r0=4, ngrp=1, tot_row=3),
    'Self-cleaning': dict(sw0=5, cd=6,  motor=7,  r0=5, ngrp=1, tot_row=4),
}
# 源表名对照（对齐 B 列阶段用）
SRC_NAME = {'89Mins': '90Mins', 'Self-cleaning': 'Self-cleaning',
            'Intensive ': 'Intensive ', 'Glass ': 'Glass '}
# 开关矩阵组数（E 列起；双列区组数 ngrp 与此不同）
NSW = {'Intensive ': 7, 'Glass ': 7, '89Mins': 7, 'Rapid': 2, 'Soak': 1, 'Self-cleaning': 1}

# 范围外表的 autoFilter 源值（V12：被我此前误统一改成 M，须恢复；注意 'Normal ' 带尾随空格）
SRC_AF = {'Auto': 'A4:P70', 'all in one': 'A4:HN62', 'Normal ': 'A4:T64',
          'WHR 1hour': 'A4:R51', '1Hour': 'C2:D51'}

# V8 富文本模式（描述 norm 后匹配 → [(文本, 颜色), ...]）
# 颜色逐模式镜像源表：剩洗/定时/漂洗分配首块=继承色(None)；跳步/再生/冷风首块=显式黑
RT_K, RT_R, RT_B = None, RED, BLACK
RT_PATTERNS = {
    '洗涤加热（定温加热-温达剩洗）': [('洗涤加热', RT_K), ('\n（定温加热-温达剩洗）', RT_R)],
    '洗涤加热（定温加热-温达跳步）': [('洗涤加热\n', RT_B), ('（定温加热-温达跳步）', RT_R)],
    '洗涤加热（定时加热）': [('洗涤加热', RT_K), ('\n（定时加热）', RT_R)],
    '再生进水风机（非再生跳步）': [('再生进水风机\n', RT_B), ('（非再生跳步）', RT_R)],
    '冷风进水排水（边进边排-进完跳步）（非再生跳步）':
        [('冷风进水排水\n', RT_B), ('（边进边排-进完跳步）', RT_R), ('\n', RT_B), ('（非再生跳步）', RT_R)],
    '洗涤加热分配（漂洗)分配器：开10秒停8秒累计动作6次':
        [('洗涤加热分配（漂洗)\n', RT_K), ('分配器：开10秒停8秒累计动作6次', RT_R)],
}

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


def make_rt(blocks):
    parts = []
    for txt, col in blocks:
        parts.append(TextBlock(InlineFont(rFont='微软雅黑',
                                          color=Color(rgb=col) if col else None), txt))
    return CellRichText(*parts)


# ================================================================ 范围外表富文本回拷
def restore_outscope(wb):
    """链路各阶段普通模式加载保存，把范围外表的富文本拍平成纯文本（违反 B2）。
    范围外表行位置从未变过，按坐标从源回拷即可。"""
    say('\n【0b】范围外表富文本回拷（从源按坐标恢复）')
    src = openpyxl.load_workbook(SRC, rich_text=True)
    n_rt = 0
    for s in src.worksheets:
        if s.title in CFG or s.title not in wb.sheetnames:
            continue                      # 范围内六表走 RT_PATTERNS 重建；已删表跳过
        ws = wb[s.title]
        for row in s.iter_rows():
            for cell in row:
                if isinstance(cell.value, CellRichText):
                    ws.cell(cell.row, cell.column).value = deepcopy(cell.value)
                    n_rt += 1
    say('  回拷富文本 ×%d（范围外表，8 张表）' % n_rt)
    return n_rt


# ================================================================ V18 漂洗第一次加热
V18 = {'sheet': '89Mins', 'match_desc': '洗涤加热（定时加热）', 'match_d': '480',
       'desc': '洗涤加热（定温加热-温达剩洗）', 'param': 'T60/480'}


def fix_rinse_heat(wb):
    say('\n【1b】V18 漂洗第一次加热：定时加热 → 温达剩洗（T60/480，配时不变）')
    n = V18['sheet']
    ws = wb[n]
    motor = CFG[n]['motor']
    hit = 0
    for r in data_rows(ws, CFG[n]['r0']):
        d = norm(ws.cell(r, 3).value)
        if d == V18['match_desc'] and str(ws.cell(r, 4).value) == V18['match_d']:
            ws.cell(r, 3).value = V18['desc']       # 富文本由【5】按模式重建
            ws.cell(r, 4).value = V18['param']
            # 温达剩洗标法（照 r16）：黄填充 C–M + 红字 D–M
            for c in range(3, motor + 1):
                ws.cell(r, c).fill = PatternFill('solid', fgColor=Color(rgb=YELLOW))
            for c in range(4, motor + 1):
                f = ws.cell(r, c).font
                ws.cell(r, c).font = Font(name=f.name or '微软雅黑', size=f.size or 10,
                                           color=Color(rgb=RED))
            hit += 1
            say('  %s r%-3d 序号%s：480 → T60/480（黄 C–M + 红字 D–M）'
                % (n, r, ws.cell(r, 1).value))
    if hit != 1:
        say('  ⚠ 命中 %d 行（应为 1）' % hit)
    return hit


# ================================================================ V19 尾部红字备注
def clear_tail_notes(wb):
    """V19（2026-11-02 时序优化）：清除时序表区域**下方**的红字备注。
    判定：数据末行之后 + 本体区（A..双速电机列）+ 红字 + 有文字 → 清空值、字体复原、
    整行空则重置行高。不删行（避免牵动合并区/筛选范围）。"""
    say('\n【1c】V19 清除时序表区域下方红字备注（维护范围内六表）')
    n_clr = 0
    for n in CFG:
        ws = wb[n]
        motor = CFG[n]['motor']
        rows = data_rows(ws, CFG[n]['r0'])
        if not rows:
            continue
        lr = rows[-1]
        for r in range(lr + 1, ws.max_row + 1):
            for c in range(1, motor + 1):
                cell = ws.cell(r, c)
                if cell.value in (None, ''):
                    continue
                f = cell.font
                if f.color and f.color.type == 'rgb' and str(f.color.value) == RED:
                    txt = str(cell.value)[:34]
                    cell.value = None
                    cell.font = Font(name=f.name or '微软雅黑', size=f.size or 10)
                    n_clr += 1
                    say('  %s r%-3d %s%d 清除：%s' % (n, r, gl(c), r, txt))
            if all(ws.cell(r, c).value in (None, '') for c in range(1, motor + 1)):
                if r in ws.row_dimensions:
                    ws.row_dimensions[r].height = None
    say('  清除红字备注 %d 处（不删行）' % n_clr)
    return n_clr


# ================================================================ V20 干燥段静止回退
REGEN_INTERVAL_MIN = 0        # 再生间隔策略：从干燥段风机拆 N 分钟为「静止（再生间隔）」
V20_SHEETS = ['Intensive ', 'Glass ', '89Mins', 'Self-cleaning']


def dry_block(ws, r0):
    """干燥段风机/静止块 = 「再生进水」行之后、其后第一个「冷风进水排水」行之前（不含两端）。
    用内容定位而不用 B 列合并区——合并区在本阶段尚未重建，行号可能已漂移。"""
    rows = data_rows(ws, r0)
    regen = next((r for r in rows if norm(ws.cell(r, 3).value).startswith('再生进水')), None)
    if regen is None:
        return []
    cold = next((r for r in rows if r > regen and '冷风进水排水' in norm(ws.cell(r, 3).value)), None)
    if cold is None:
        return []
    return [r for r in rows if regen < r < cold]


def revert_dry_static(wb):
    say('\n【1d】V20 干燥段静止回退为风机 + 相邻同开关合并（再生间隔 N=%d 分钟）'
        % REGEN_INTERVAL_MIN)
    touched = []
    for n in V20_SHEETS:
        ws = wb[n]
        cfg = CFG[n]
        cd, motor, r0, sw0 = cfg['cd'], cfg['motor'], cfg['r0'], cfg['sw0']
        nsw = NSW[n]

        def mask(r):
            return tuple(ws.cell(r, sw0 + g).value for g in range(nsw))

        blk = dry_block(ws, r0)
        if not blk:
            say('  %-14s 干燥段无风机/静止块，跳过' % n)
            continue
        # --- 1) 静止 → 风机（时长/配时不变）；块内标记一律清除
        n_st = 0
        for r in blk:
            if norm(ws.cell(r, 3).value) == '静止':
                ws.cell(r, 3).value = '风机'
                n_st += 1
            for c in range(3, motor + 1):
                cell = ws.cell(r, c)
                if cell.fill and cell.fill.patternType:
                    cell.fill = PatternFill()
                f = cell.font
                if f.color and f.color.type == 'rgb' and str(f.color.value) == RED:
                    cell.font = Font(name=f.name or '微软雅黑', size=f.size or 10)
        # --- 2) 相邻同开关矩阵的两步合并（每合并一次就重算块，行号会变）
        n_mg = 0
        while True:
            blk = dry_block(ws, r0)
            hit = False
            for i in range(len(blk) - 1, 0, -1):
                a, b = blk[i - 1], blk[i]
                if mask(a) == mask(b):
                    pa, pb = ws.cell(a, 4).value or 0, ws.cell(b, 4).value or 0
                    ca, cb = ws.cell(a, cd).value or 0, ws.cell(b, cd).value or 0
                    ws.cell(a, 4).value = pa + pb
                    ws.cell(a, cd).value = ca + cb
                    ws.delete_rows(b)
                    say('  %-14s 合并 r%d+r%d → 风机%s（cd %s+%s=%s）[开关矩阵相同]'
                        % (n, a, b, pa + pb, ca, cb, ca + cb))
                    n_mg += 1
                    hit = True
                    break
            if not hit:
                break
        say('  %-14s 静止→风机 %d 行；合并 %d 处；块内标记已清除'
            % (n, n_st, n_mg))
        if n_mg:
            touched.append(n)
    if REGEN_INTERVAL_MIN:
        say('  ⚠ REGEN_INTERVAL_MIN=%d —— 拆分逻辑尚未接入，需先在 V20 实现' % REGEN_INTERVAL_MIN)
    else:
        say('  再生间隔策略 N=0 → 不拆出「静止（再生间隔）」，干燥段无新增标记')
    return touched


# ================================================================ 值层修复
def fix_values(wb):
    say('\n【1】值层修复（V14 风机 / V15 漂洗参数 —— 全工作簿；V16 干燥前最后加热跳步）')
    # V14/V15：用户验收原话「全部清除」「只需要保留数字」→ 对全工作簿生效
    # （AGENT B2「范围外不动」对这两单项按验收裁决豁免，已记入变更台账）
    n14 = n15 = 0
    for n in wb.sheetnames:
        ws = wb[n]
        for r in range(1, ws.max_row + 1):
            v = ws.cell(r, 3).value
            # V14 风机：清除"（非交替）"字样（写回纯文本）
            if v is not None and '风机' in str(v) and '（非交替）' in str(v):
                ws.cell(r, 3).value = '风机'
                n14 += 1
                say('  [风机]   %s r%-3d → 风机' % (n, r))
            # V15 漂洗分配参数只留数字（'60（大）' → 60）
            d = norm(v)
            p = str(ws.cell(r, 4).value or '')
            if '漂洗' in d and '分配' in d:
                m = re.match(r'^(\d+(?:\.\d+)?)（大）$', p)
                if m:
                    num = m.group(1)
                    ws.cell(r, 4).value = float(num) if '.' in num else int(num)
                    n15 += 1
                    say('  [参数]   %s r%-3d %s → %s' % (n, r, p, ws.cell(r, 4).value))
    say('  V14 风机 ×%d；V15 参数 ×%d（全工作簿）' % (n14, n15))
    # V16 干燥前最后一次加热 → 温达跳步（cd 不变，总时长守恒）
    for n, src_cd in [('89Mins', 12), ('Self-cleaning', 6)]:
        ws = wb[n]
        rows = [r for r in data_rows(ws, CFG[n]['r0']) if '加热' in norm(ws.cell(r, 3).value)]
        r = rows[-1]
        old = ws.cell(r, 4).value
        t = re.match(r'^(T\d+)', str(old))
        if t:
            ws.cell(r, 3).value = '洗涤加热（定温加热-温达跳步）'
            ws.cell(r, 4).value = t.group(1)
            say('  [跳步]   %s r%-3d %s → %s（cd=%s 不变）' % (
                n, r, old, t.group(1), ws.cell(r, src_cd).value))
            # 跳步标法（Glass）：黄 C–D 无红 → 先清整行再上黄
            for c in range(3, CFG[n]['motor'] + 1):
                ws.cell(r, c).fill = PatternFill()
                f = ws.cell(r, c).font
                if f.color and f.color.type == 'rgb' and str(f.color.value) == RED:
                    ws.cell(r, c).font = Font(name=f.name, size=f.size)
            for c in (3, 4):
                ws.cell(r, c).fill = PatternFill('solid', fgColor=Color(rgb=YELLOW))


# ================================================================ 序号重排 + A 合并全拆
def renumber(wb, sheets=None):
    say('\n【2】序号重排 + 拆 A 列合并（V4）')
    for n in (sheets or ['89Mins', 'Self-cleaning']):
        ws = wb[n]
        for m in list(ws.merged_cells.ranges):
            if m.min_col == 1 and m.min_row >= CFG[n]['r0']:
                ws.unmerge_cells(str(m))
        rows = data_rows(ws, CFG[n]['r0'])
        for i, r in enumerate(rows, 1):
            ws.cell(r, 1).value = i
        say('  %s 序号 1..%d（%d 行）' % (n, len(rows), len(rows)))


# ================================================================ B 列阶段重建
def rebuild_stage(wb, sheets=None):
    say('\n【3】B 列阶段合并重建（V6，difflib 对齐源表）')
    src = openpyxl.load_workbook(SRC)
    for n in (sheets or ['89Mins', 'Self-cleaning']):
        ws, sw = wb[n], src[SRC_NAME[n]]
        cd = CFG[n]['cd']
        r0 = CFG[n]['r0']
        # 源阶段标签（B 列合并区展开）
        stage_src = {}
        for m in sw.merged_cells.ranges:
            if m.min_col == 2 and m.min_row >= r0:
                v = sw.cell(m.min_row, 2).value
                for r in range(m.min_row, m.max_row + 1):
                    stage_src[r] = v

        def key(w, r):
            return (norm(w.cell(r, 3).value), str(w.cell(r, 4).value), w.cell(r, cd).value)

        srows = [r for r in range(r0, sw.max_row + 1) if sw.cell(r, 3).value]
        drows = data_rows(ws, r0)
        a = [key(sw, r) for r in srows]
        b = [key(ws, r) for r in drows]
        sm = difflib.SequenceMatcher(a=a, b=b, autojunk=False)
        stage = {}
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == 'equal':
                for k in range(i2 - i1):
                    stage[drows[j1 + k]] = stage_src.get(srows[i1 + k])
        last = None
        for r in drows:                      # 未对齐行向下填充
            if stage.get(r):
                last = stage[r]
            else:
                stage[r] = last
        # 拆旧合并、清值
        for m in list(ws.merged_cells.ranges):
            if m.min_col == 2 and m.min_row >= r0:
                ws.unmerge_cells(str(m))
        for r in range(r0, drows[-1] + 1):
            ws.cell(r, 2).value = None
        # 重新合并
        spans, i = [], 0
        while i < len(drows):
            s = stage[drows[i]]
            j = i
            while j + 1 < len(drows) and stage[drows[j + 1]] == s:
                j += 1
            spans.append((drows[i], drows[j], s))
            i = j + 1
        for r1, r2, s in spans:
            ws.cell(r1, 2).value = s
            if r2 > r1:
                ws.merge_cells('B%d:B%d' % (r1, r2))
        for r in range(r0, drows[-1] + 1):
            ws.cell(r, 2).border = THIN_BORDER
        say('  %s 阶段=%s' % (n, ['%s(%d-%d)' % (s, a, b) for a, b, s in spans]))


# ================================================================ Self-cleaning 静止特例
def special_static(wb):
    say('\n【4】Self-cleaning 静止特例（V3）—— V20 回退后已失效，停用')
    n = 'Self-cleaning'
    ws = wb[n]
    rows = [r for r in dry_block(ws, CFG[n]['r0']) if norm(ws.cell(r, 3).value) == '静止']
    if not rows:
        say('  %s 干燥段已无静止（V20 已回退为风机）→ V3 跳过' % n)
        return
    r = rows[-1]
    for c in (3, 4):
        ws.cell(r, c).fill = PatternFill('solid', fgColor=Color(rgb=YELLOW))
    for c in (3, 4):
        f = ws.cell(r, c).font
        ws.cell(r, c).font = Font(name=f.name or '微软雅黑', size=f.size or 10, color=Color(rgb=RED))
    say('  %s r%-3d 静止 → 黄填充 C–D + 红字 C,D' % (n, r))


# ================================================================ C 列富文本重建
def rebuild_rich_text(wb):
    say('\n【5】C 列富文本重建（V8：回车+行内红+wrap+行高33）')
    n_rt = 0
    for n in CFG:
        ws = wb[n]
        for r in data_rows(ws, CFG[n]['r0']):
            d = norm(ws.cell(r, 3).value)
            blocks = RT_PATTERNS.get(d)
            if not blocks:
                continue                          # 非富文本模式行（幂等：已建的也会按模式重建）
            ws.cell(r, 3).value = make_rt(blocks)
            # V17：源表 C 列一律「水平+垂直居中 + 自动换行」，此前漏了 horizontal
            ws.cell(r, 3).alignment = Alignment(horizontal='center', vertical='center',
                                                wrap_text=True)
            if r in ws.row_dimensions:
                ws.row_dimensions[r].height = 33.0
            else:
                ws.row_dimensions[r].height = 33.0
            n_rt += 1
    say('  富文本行 ×%d（六表）' % n_rt)


# ================================================================ 字体/框线统一
def unify_format(wb):
    say('\n【6】字体与框线统一（V4/V5）')
    n_fix = 0
    for n in ['89Mins', 'Self-cleaning']:
        ws = wb[n]
        motor = CFG[n]['motor']
        for r in data_rows(ws, CFG[n]['r0']):
            for c in range(1, motor + 1):
                cell = ws.cell(r, c)
                f = cell.font
                if f.name != '微软雅黑' or f.size not in (10, 10.0):
                    cell.font = Font(name='微软雅黑', size=10,
                                     bold=f.bold, color=f.color if f.color else None)
                    n_fix += 1
                b = cell.border
                if not all(x and x.style for x in (b.left, b.right, b.top, b.bottom)):
                    cell.border = THIN_BORDER
                    n_fix += 1
    say('  修补 %d 处（宋体11/无边框）' % n_fix)


# ================================================================ 双列区：去框线 + 红蓝绿
def dual_zone(wb):
    say('\n【7】双列区：数据行去框线（V10）+ 倒计时列颜色（V9）')
    for n in CFG:
        ws = wb[n]
        cfg = CFG[n]
        out0, ngrp, r0 = cfg['motor'] + 1, cfg['ngrp'], cfg['r0']
        rows = data_rows(ws, r0)
        if not rows:
            continue
        lr = rows[-1]
        stats = {'蓝': 0, '红': 0, '绿': 0}
        for g in range(ngrp):
            c1 = out0 + g * 2
            c2 = c1 + 1
            for i, r in enumerate(rows):
                is_last = (i == len(rows) - 1)
                # V10 去框线（r3 名称行 / r4 起值行 保留）
                if r > cfg['tot_row']:
                    ws.cell(r, c1).border = NO_BORDER
                    ws.cell(r, c2).border = NO_BORDER
                # V9 颜色（仅倒计时列 c2；累计列黑）
                d = norm(ws.cell(r, 3).value)
                v2 = ws.cell(r, c2).value
                if v2 in (None, ''):         # 灰底无值行
                    continue
                col = None
                if is_last or str(v2).endswith('/end'):
                    col = GREEN
                    stats['绿'] += 1
                elif '进水' in d:            # 进水/再生进水风机/冷风进水排水
                    col = BLUE
                    stats['蓝'] += 1
                else:
                    f = ws.cell(r, 3).fill
                    if f and f.patternType and f.fgColor and f.fgColor.type == 'rgb' \
                            and str(f.fgColor.value) == YELLOW:
                        col = RED            # 温达剩洗/跳步/静止特例
                        stats['红'] += 1
                ws.cell(r, c2).font = Font(name='微软雅黑', size=10,
                                           color=Color(rgb=col) if col else None)
                ws.cell(r, c1).font = Font(name='微软雅黑', size=10)
        say('  %-14s 去%dx2框线；颜色 蓝%d 红%d 绿%d' % (
            n, len(rows), stats['蓝'], stats['红'], stats['绿']))


# ================================================================ 居中统一（V17）
def fix_alignment(wb):
    """V17（2026-11-02 用户指出"居中忘记调整"）：
    · 富文本重建行（V8）丢了 horizontal → 源表 C 列 100% 是 center/center/wrap
    · 新增行 / 拆行（V6 阶段重建、V4 序号）新建单元格没有对齐
    · 双列区是新建区域，源表旧双列区主流也是居中
    统一补 horizontal=center + vertical=center，wrap_text 保留原值。"""
    say('\n【7b】居中统一（V17：水平+垂直居中，保留 wrap）')
    n_fix = 0
    for n in CFG:
        ws = wb[n]
        cfg = CFG[n]
        motor, r0 = cfg['motor'], cfg['r0']
        rows = data_rows(ws, r0)
        if not rows:
            continue
        dual_rows = sorted(set(rows) | {max(r0 - 2, 1), max(r0 - 1, 1)})
        targets = [(r, c) for r in rows for c in range(1, motor + 1)]
        targets += [(r, c) for r in dual_rows
                    for c in range(motor + 1, motor + 1 + cfg['ngrp'] * 2)]
        for r, c in targets:
            cell = ws.cell(r, c)
            a = cell.alignment
            if a.horizontal == 'center' and a.vertical == 'center':
                continue
            cell.alignment = Alignment(horizontal='center', vertical='center',
                                       wrap_text=a.wrap_text,
                                       textRotation=a.textRotation, indent=a.indent)
            n_fix += 1
    say('  居中补齐 %d 格（六表：本体区 + 双列区）' % n_fix)
    return n_fix


# ================================================================ 隐藏 / autoFilter / 列宽 / 注释
def housekeeping(wb):
    say('\n【8】取消全部隐藏行列（V7）+ autoFilter（V12）+ 列宽（V11）+ 注释（V13）')
    # V7 隐藏
    nh = 0
    for n in wb.sheetnames:
        ws = wb[n]
        for d in ws.row_dimensions.values():
            if d.hidden:
                d.hidden = False
                nh += 1
        for k, d in list(ws.column_dimensions.items()):
            if d.hidden:
                d.hidden = False
                nh += 1
    say('  取消隐藏：%d 处（含范围外表——用户明示"所有"）' % nh)

    # V12 autoFilter：范围内→motor 列；范围外→恢复源值
    for n in CFG:
        ws = wb[n]
        m = gl(CFG[n]['motor'])
        last = ws.max_row
        ws.auto_filter.ref = 'A%d:%s%d' % (CFG[n]['tot_row'], m, last)
    for n, ref in SRC_AF.items():
        if n in wb.sheetnames:
            wb[n].auto_filter.ref = ref
    say('  autoFilter：六表收缩到本体末列；范围外表恢复源值')

    # V11 列宽：清残留（Y/Z/IZ），双列区统一 10
    for n in CFG:
        ws = wb[n]
        cfg = CFG[n]
        out0, ngrp = cfg['motor'] + 1, cfg['ngrp']
        for k in list(ws.column_dimensions.keys()):
            c = openpyxl.utils.column_index_from_string(k)
            if c >= out0 + ngrp * 2:
                del ws.column_dimensions[k]
        for g in range(ngrp):
            for c in range(out0 + g * 2, out0 + g * 2 + 2):
                letter = gl(c)
                ws.column_dimensions[letter].width = 10.0
                ws.column_dimensions[letter].hidden = False
    say('  列宽：双列区统一 10；残留宽度已清')

    # V13 注释
    nc = 0
    for n in wb.sheetnames:
        ws = wb[n]
        for row in ws.iter_rows():
            for cell in row:
                if cell.comment:
                    cell.comment = None
                    nc += 1
    say('  删除注释：%d 个' % nc)


# ================================================================ 校验
def verify(wb):
    say('\n【9】终检')
    ok = True
    # A3 时长
    for n in CFG:
        ws, cfg = wb[n], CFG[n]
        cd, sw0, ngrp = cfg['cd'], cfg['sw0'], cfg['ngrp']
        rows = data_rows(ws, cfg['r0'])
        for g in range(ngrp):
            swc = sw0 + g
            total = ws.cell(cfg['tot_row'], swc).value
            if not isinstance(total, (int, float)):
                continue
            s = sum(ws.cell(r, cd).value for r in rows
                    if ws.cell(r, swc).value not in (None, '')
                    and isinstance(ws.cell(r, cd).value, (int, float)))
            flag = 'OK' if s == total else '❌'
            ok = ok and s == total
            say('  [%s] %s/%s 标称=%s 实算=%s' % (flag, n, ws.cell(3, swc).value, total, s))
    # 序号连续
    for n in ['89Mins', 'Self-cleaning']:
        ws = wb[n]
        seq = [ws.cell(r, 1).value for r in data_rows(ws, CFG[n]['r0'])]
        good = seq == list(range(1, len(seq) + 1))
        ok = ok and good
        say('  [%s] %s 序号 1..%d' % ('OK' if good else '❌', n, len(seq)))
    # 完整性
    for n in CFG:
        ws, cd = wb[n], CFG[n]['cd']
        bad = [r for r in data_rows(ws, CFG[n]['r0'])
               if (ws.cell(r, 4).value not in (None, '') or ws.cell(r, cd).value not in (None, ''))
               and not ws.cell(r, 3).value]
        ok = ok and not bad
        say('  [%s] %s 完整性' % ('OK' if not bad else '❌ 行%s' % bad, n))
    # 隐藏清零
    hid = sum(1 for n in wb.sheetnames for d in wb[n].row_dimensions.values() if d.hidden) \
        + sum(1 for n in wb.sheetnames for d in wb[n].column_dimensions.values() if d.hidden)
    ok = ok and hid == 0
    say('  [%s] 隐藏残留 %d' % ('OK' if hid == 0 else '❌', hid))
    # 富文本计数
    n_rt = sum(1 for n in CFG for r in data_rows(wb[n], CFG[n]['r0'])
               if isinstance(wb[n].cell(r, 3).value, CellRichText))
    say('  [OK] 富文本行 %d（范围内六表）' % n_rt)
    # 范围外表富文本保全（与源逐坐标一致）
    src = openpyxl.load_workbook(SRC, rich_text=True)
    n_out_src = n_out_ok = 0
    for s in src.worksheets:
        if s.title in CFG or s.title not in wb.sheetnames:
            continue
        for row in s.iter_rows():
            for cell in row:
                if isinstance(cell.value, CellRichText):
                    n_out_src += 1
                    pv = wb[s.title].cell(cell.row, cell.column).value
                    if isinstance(pv, CellRichText):
                        n_out_ok += 1
    good = n_out_ok >= n_out_src
    ok = ok and good
    say('  [%s] 范围外表富文本 %d/%d' % ('OK' if good else '❌', n_out_ok, n_out_src))
    # autoFilter 恢复
    af_ok = all(wb[n].auto_filter.ref == ref for n, ref in SRC_AF.items() if n in wb.sheetnames)
    ok = ok and af_ok
    say('  [%s] 范围外表 autoFilter 恢复' % ('OK' if af_ok else '❌'))
    # V17 居中
    bad_al = []
    for n in CFG:
        ws, cfg = wb[n], CFG[n]
        rows = data_rows(ws, cfg['r0'])
        for r, c in [(r, c) for r in rows for c in range(1, cfg['motor'] + 1)] + \
                    [(r, c) for r in rows
                     for c in range(cfg['motor'] + 1, cfg['motor'] + 1 + cfg['ngrp'] * 2)]:
            a = ws.cell(r, c).alignment
            if (a.horizontal, a.vertical) != ('center', 'center'):
                bad_al.append('%s!%s%d' % (n, gl(c), r))
    ok = ok and not bad_al
    say('  [%s] 居中残留 %d%s' % ('OK' if not bad_al else '❌', len(bad_al),
                                 ' ' + str(bad_al[:6]) if bad_al else ''))
    say('\n终检：%s' % ('✅ 全部通过' if ok else '❌ 有失败项'))
    return ok


def main():
    wb = openpyxl.load_workbook(DST, rich_text=True)
    say('=' * 74)
    say('第二步验收修复（13 项）')
    say('=' * 74)
    restore_outscope(wb)
    fix_values(wb)
    fix_rinse_heat(wb)
    clear_tail_notes(wb)
    touched = revert_dry_static(wb)          # V20：会删行，必须在序号重排与阶段重建之前
    renumber(wb, ['89Mins', 'Self-cleaning'] + touched)
    rebuild_stage(wb, ['89Mins', 'Self-cleaning'] + touched)
    special_static(wb)
    rebuild_rich_text(wb)
    unify_format(wb)
    dual_zone(wb)
    fix_alignment(wb)
    housekeeping(wb)
    ok = verify(wb)
    wb.save(DST)
    say('\n已保存：%s' % DST)

    rp = os.path.join(ROOT, '02_衍生物', '02_中间产物', '第二步_验收修复日志.md')
    with open(rp, 'w', encoding='utf-8') as f:
        f.write('# 第二步 · 验收修复日志（2026-11-02）\n\n```\n')
        f.write('\n'.join(LOG))
        f.write('\n```\n')
    print('日志：%s' % rp)
    return 0 if ok else 1


if __name__ == '__main__':
    raise SystemExit(main())
