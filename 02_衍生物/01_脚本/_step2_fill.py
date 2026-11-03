# -*- coding: utf-8 -*-
"""
第二步补完：修补缺陷 + 措辞改写 + 填充/红字（按「改动原因」规则）

══════════════════════════════════════════════════════════════
本轮实证得出的「标记语义」——不是装饰，是改动原因的编码：

  填充色 = 改动的性质
    绿 = 结构性改动：拆行、新增动作定义、动作方式重写、参数性质变化
    黄 = 参数/描述级改动：参数形式被改写、描述被规范化
    无 = 纯数值微调（水量 L、洗涤秒数、干燥段秒数重排）

  填充横向范围 = 改动的波及范围
    C–M(整行) = 参数被改，波及开关矩阵/配时/电机
    C–D       = 只有描述与参数被改

  红字 = 被改动的那个具体值，逐格点名
    C        = 描述被改
    D        = 参数被改
    D + 配时  = 参数与配时都被改
    D–M      = 加热参数形式被改写，整段数值需重看
══════════════════════════════════════════════════════════════

另修补上一轮遗留缺陷：
  · 6 行丢失「序号 + 描述」（只剩参数与配时）——求和校验发现不了
  · 4 处 MINI 合并格把正确序号遮住

输入/输出：03_验收物/普通版_第二步已处理.xlsx（原地改写；源文件永不写）
"""
import os, re, copy
import openpyxl
from openpyxl.styles import PatternFill, Font, Color
from openpyxl.utils import get_column_letter as gl

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
DST = os.path.join(ROOT, '03_验收物', '普通版_第二步已处理.xlsx')
SRC = os.path.join(ROOT, '01_源文件', '普通版.xlsx')   # 回归基准：人工维护的原始状态

# ---------------------------------------------------------------- 裁决常量（D2：一行可回滚）
GREEN = 'FF92D050'
YELLOW = 'FFFFFF00'
RED = 'FFFF0000'
BLACK = 'FF000000'

# 各分栏列位置（D3：一律显式配置）
CFG = {
    'Intensive ':    dict(cd=12, motor=13),
    'Glass ':        dict(cd=12, motor=13),
    '89Mins':        dict(cd=12, motor=13),
    'Rapid':         dict(cd=7,  motor=8),
    'Soak':          dict(cd=6,  motor=7),
    'Self-cleaning': dict(cd=6,  motor=7),
}

# 修补表：上一轮丢失「序号 + 描述」的行。依据《第二步_本体措辞改写对照表》与源文件反查得出。
# 行号 -> (序号, 描述)；描述 None = 按参数形式自动推断
REPAIR = {
    '89Mins':        {15: (10, '洗涤加热（定时加热）'), 40: (35, None), 43: (39, None)},
    'Self-cleaning': {11: (6,  '洗涤加热（定时加热）'), 25: (20, None), 28: (24, '洗涤')},
}
# 需要拆开的 MINI 合并格（合并后序号被上一行遮住）：(首行, 末行)
UNMERGE_A = {
    '89Mins':        [(39, 40), (42, 43)],
    'Self-cleaning': [(24, 25), (27, 28)],
}
# 拆开后需要补回的序号（依源文件：89Mins 应为 35/39，Self-cleaning 应为 20/24）
SET_A = {
    '89Mins':        {40: 35, 43: 39},
    'Self-cleaning': {25: 20, 28: 24},
}
# 89Mins 干燥段特例：再生之后的「静止」是「不补风机」的独立模式，
# 逻辑与其他分栏不同 → 黄填充 + 红字双标（用户裁决 D）
SPECIAL_STATIC = {'89Mins': (51, 52, 53)}

LOG = []


def say(*a):
    s = ' '.join(str(x) for x in a)
    print(s)
    LOG.append(s)


# ---------------------------------------------------------------- 判定
def temp_form(param):
    s = re.sub(r'\s+', '', str(param or ''))
    if re.fullmatch(r'T\d+/\d+', s):
        return '温达剩洗'
    if re.fullmatch(r'T\d+', s):
        return '温达跳步'
    return None


def classify(desc):
    """返回 (类别, 填充色, 填充末列偏移, 红字列生成函数关键词)"""
    d = re.sub(r'\s+', '', str(desc or ''))
    if not d:
        return None, None, None, None
    if '洗涤加热分配' in d and '分配器' in d:
        return 'DISPENSER', GREEN, 'ROW', None
    if re.match(r'^洗涤分配', d):
        return 'DISPENSER_MAIN', GREEN, 'ROW', 'C'
    if '定时加热' in d:
        return 'HEAT_TIMED', GREEN, 'ROW', None
    if '温达剩洗' in d:
        return 'HEAT_TEMP_FULL', YELLOW, 'ROW', 'D_M'
    if '温达跳步' in d:
        # ⚠ 模板不统一：Intensive 标 C–M+红，Glass 只标 C–D 无红。本轮两张待处理表无此行，
        #    暂按「波及范围」理论取 C–D。见报告待统一项。
        return 'HEAT_TEMP_DESC', YELLOW, 'CD', None
    if '再生进水风机' in d:
        return 'REGEN', GREEN, 'ROW', 'D_CD'
    if '冷风进水排水' in d:
        return 'COOL', GREEN, 'ROW', 'D_CD'
    if d == '静止（再生间隔）':
        # 再生间隔策略预留行（V20，配时可为 0）：黄 C–D + 红字 C,D（用户 2026-11-03 纠错）
        return 'REGEN_INTERVAL', YELLOW, 'CD', 'CD'
    if d == '静止' or re.match(r'^风机', d):
        # 干燥段的静止/风机：模板标法不统一（Intensive r61/62 红、Glass r48 不红）→ 判为「不托管」，
        # 保留原样不机械改写。89Mins 的静止另有用户裁决，走 SPECIAL_STATIC。
        return 'DRY_STATIC', None, None, None
    return None, None, None, None


def rewrite_desc(desc, param):
    d = re.sub(r'\s+', '', str(desc or ''))
    if d != '洗涤加热':
        return desc, False
    f = temp_form(param)
    if f:
        return '洗涤加热（定温加热-%s）' % f, True
    if re.fullmatch(r'\d+', re.sub(r'\s+', '', str(param or ''))):
        return '洗涤加热（定时加热）', True
    return desc, False


def set_font(cell, color):
    """只改颜色，其余字体属性原样重建。
    注意：cell.font 读出的是 StyleProxy，直接 copy/deepcopy 会踩 RecursionError，
    必须显式重建 Font 对象。"""
    f = cell.font
    cell.font = Font(name=f.name, size=f.size, bold=f.bold, italic=f.italic,
                     underline=f.underline, strike=f.strike,
                     vertAlign=f.vertAlign, color=Color(rgb=color))


def last_data_row(ws):
    r = 5
    for i in range(5, ws.max_row + 1):
        if isinstance(ws.cell(i, 1).value, (int, float)) or ws.cell(i, 4).value not in (None, ''):
            r = i
    return r


def red_cols_of(kind, keyword, cd, motor):
    if keyword is None:
        return []
    if keyword == 'C':
        return [3]
    if keyword == 'CD':
        return [3, 4]
    if keyword == 'D_CD':
        return [4, cd]
    if keyword == 'D_M':
        return list(range(4, motor + 1))
    return []


# ---------------------------------------------------------------- 阶段一：修补 + 措辞
def repair_and_rename(ws, sheet):
    # 1a. 拆合并格：C 列一律拆（每行要有自己的描述）；
    #     A 列只拆 UNMERGE_A 列出的 MINI 遗留对（拆行对保持合并＝同一序号，依对照表）
    a_pairs = set(UNMERGE_A.get(sheet, []))
    rep_rows = set(REPAIR.get(sheet, {}))
    for m in list(ws.merged_cells.ranges):
        if m.min_row < 5:                       # 表头区不动
            continue
        if m.min_col == 3 and any(m.min_row <= r <= m.max_row for r in rep_rows):
            ws.unmerge_cells(str(m))
            say('  [拆C合并] %s %s' % (sheet, m))
        elif m.min_col == 1 and (m.min_row, m.max_row) in a_pairs:
            ws.unmerge_cells(str(m))
            say('  [拆A合并] %s %s' % (sheet, m))

    # 1b. 补序号（仅限 A 列已拆开的 MINI 遗留行）
    for r, seq in sorted(SET_A.get(sheet, {}).items()):
        cell = ws.cell(r, 1)
        if not isinstance(cell, openpyxl.cell.cell.MergedCell) and cell.value is None:
            cell.value = seq
            say('  [补序号] %s r%-3d → %s' % (sheet, r, seq))

    # 1c. 补描述
    for r, (seq, desc) in sorted(REPAIR.get(sheet, {}).items()):
        if desc is None:
            desc, _ = rewrite_desc('洗涤加热', ws.cell(r, 4).value)
        if not ws.cell(r, 3).value:
            ws.cell(r, 3).value = desc
            say('  [补描述] %s r%-3d → %s' % (sheet, r, desc))

    # 1c. 裸「洗涤加热」→ 新版描述
    lr = last_data_row(ws)
    n = 0
    for r in range(5, lr + 1):
        old = ws.cell(r, 3).value
        new, ch = rewrite_desc(old, ws.cell(r, 4).value)
        if ch:
            ws.cell(r, 3).value = new
            n += 1
    say('  [措辞]   %s 洗涤加热 → 新版描述：%d 行' % (sheet, n))


# ---------------------------------------------------------------- 阶段二：填充 + 红字
def clear_visual(ws, sheet):
    motor = CFG[sheet]['motor']
    n = 0
    for r in range(5, last_data_row(ws) + 1):
        for cc in range(3, motor + 1):
            cell = ws.cell(r, cc)
            if cell.fill and cell.fill.patternType:
                cell.fill = PatternFill()
                n += 1
    # 红字不在此统一清除：干燥段的静止/风机属「不托管」行，其红字是人工标的，必须保留。
    # 托管行的红字由 paint() 逐格判定（该红的补上、不该红的清掉）。
    return n


def paint(ws, sheet, check_only=False):
    cfg = CFG[sheet]
    cd, motor = cfg['cd'], cfg['motor']
    lr = last_data_row(ws)
    special = SPECIAL_STATIC.get(sheet, ())
    diffs = []

    for r in range(5, lr + 1):
        desc = ws.cell(r, 3).value
        kind, fill, span, kw = classify(desc)

        # 89Mins 干燥段特例：必须在「不托管」判断之前生效，否则会被 continue 跳过
        if r in special and re.sub(r'\s+', '', str(desc or '')) == '静止':
            kind, fill, span, kw = 'SPECIAL_STATIC', YELLOW, 'CD', 'CD'

        # 不托管行（普通步骤、干燥段静止/风机）：填充已统一清除，红字保持手工原样
        if kind is None or kind == 'DRY_STATIC':
            continue

        fill_end = motor if span == 'ROW' else (4 if span == 'CD' else None)
        reds = red_cols_of(kind, kw, cd, motor)

        for cc in range(3, motor + 1):
            cell = ws.cell(r, cc)

            # ---- 填充
            want_fill = fill if (fill and fill_end and cc <= fill_end) else None
            cur_fill = ''
            if cell.fill and cell.fill.patternType and cell.fill.fgColor \
                    and cell.fill.fgColor.type == 'rgb':
                cur_fill = str(cell.fill.fgColor.value)
            if want_fill:
                if not check_only:
                    cell.fill = PatternFill('solid', fgColor=want_fill)
                if cur_fill != want_fill:
                    diffs.append((r, gl(cc), '填充', cur_fill or '无', want_fill))
            elif cur_fill in (GREEN, YELLOW):
                if not check_only:
                    cell.fill = PatternFill()
                diffs.append((r, gl(cc), '填充', cur_fill, '无'))

            # ---- 红字
            cur_red = False
            if cell.font and cell.font.color and cell.font.color.type == 'rgb':
                cur_red = str(cell.font.color.value) == RED
            if cc in reds:
                if not cur_red:
                    if not check_only:
                        set_font(cell, RED)
                    diffs.append((r, gl(cc), '红字', '无', RED))
            elif cur_red:
                if not check_only:
                    set_font(cell, BLACK)
                diffs.append((r, gl(cc), '红字', RED, '清'))

    return diffs


def integrity(ws, sheet):
    cd = CFG[sheet]['cd']
    bad = []
    for r in range(5, ws.max_row + 1):
        d, v, c = ws.cell(r, 4).value, ws.cell(r, cd).value, ws.cell(r, 3).value
        if (d not in (None, '') or v not in (None, '')) and not c:
            bad.append(r)
    return bad


# ---------------------------------------------------------------- 主流程
def main():
    wb = openpyxl.load_workbook(DST)
    say('=' * 74)
    say('第二步补完：修补 + 措辞改写 + 填充/红字')
    say('=' * 74)

    # 回归基准必须是**源文件里的模板分栏**（人工维护的原始状态），
    # 不能拿成品——成品已被上一轮落笔改成规则的样子，那会变成自证。
    say('\n【0】模板回归：规则能否复现人工标记（基准＝源文件，差异应为 0）')
    src = openpyxl.load_workbook(SRC)
    for sheet in ['Intensive ', 'Glass ', 'Rapid']:
        d = paint(src[sheet], sheet, check_only=True)
        say('  %-14s 差异 %-3d %s' % (sheet, len(d),
            '✅ 规则可复现' if not d else '⚠ ' + str(d[:3])))

    say('\n【1】修补丢失行 + 措辞改写')
    for sheet in ['89Mins', 'Self-cleaning']:
        repair_and_rename(wb[sheet], sheet)

    say('\n【2】清除既有填充/红字')
    for sheet in CFG:
        say('  %-14s 清除 %d 处' % (sheet, clear_visual(wb[sheet], sheet)))

    say('\n【3】按「改动原因」重标')
    for sheet in CFG:
        d = paint(wb[sheet], sheet, check_only=False)
        say('  %-14s 落笔 %d 处' % (sheet, len(d)))

    say('\n【4】完整性：有参数/配时却无描述 = 缺陷')
    ok = True
    for sheet in CFG:
        bad = integrity(wb[sheet], sheet)
        say('  %-14s %s' % (sheet, '✅' if not bad else '❌ 行 %s' % bad))
        ok = ok and not bad

    say('\n【5】序号序列')
    for sheet in ['89Mins', 'Self-cleaning']:
        seq = [wb[sheet].cell(r, 1).value for r in range(5, wb[sheet].max_row + 1)]
        seq = [x for x in seq if isinstance(x, (int, float))]
        say('  %-14s %s' % (sheet, seq))

    wb.save(DST)
    say('\n已保存：%s' % DST)
    say('完整性：%s' % ('✅ 通过' if ok else '❌ 仍有缺陷'))

    rp = os.path.join(ROOT, '02_衍生物', '02_中间产物', '第二步_填充与改写日志.md')
    with open(rp, 'w', encoding='utf-8') as f:
        f.write('# 第二步 · 填充与措辞改写日志\n\n```\n')
        f.write('\n'.join(LOG))
        f.write('\n```\n')
    print('日志：%s' % rp)


if __name__ == '__main__':
    main()
