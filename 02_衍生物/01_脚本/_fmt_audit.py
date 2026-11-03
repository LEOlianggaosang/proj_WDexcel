# -*- coding: utf-8 -*-
"""源表 vs 成品：逐格格式保真审计（不涉及数据值正确性）"""
import os
import openpyxl

ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
SRC = os.path.join(ROOT, '01_源文件', '普通版.xlsx')
DST = os.path.join(ROOT, '03_验收物', '_历史', '普通版_Express已处理.xlsx')


def cell_sig(c):
    """把单元格的全部视觉属性抽成可比对的签名"""
    f, a = c.font, c.alignment
    fill = c.fill
    fg = fill.fgColor
    def norm(v):
        # Indexed 型颜色在 openpyxl 里会格式化成报错串，视觉上等价于 00000000
        if isinstance(v, str) and 'Values must be of type' in v:
            return '00000000'
        return v

    fgv = fg.rgb if getattr(fg, 'type', None) == 'rgb' else ('theme%s:%s' % (fg.theme, fg.tint))
    bg = fill.bgColor
    bgv = bg.rgb if getattr(bg, 'type', None) == 'rgb' else ('theme%s:%s' % (bg.theme, bg.tint))

    fgv = norm(fgv)
    bgv = norm(bgv)
    b = c.border
    bds = []
    for s in (b.left, b.right, b.top, b.bottom):
        col = s.color
        colv = col.rgb if col is not None and getattr(col, 'type', None) == 'rgb' \
            else ('theme' if col is not None else None)
        bds.append((s.style, colv))
    fcol = f.color
    return (
        c.number_format,
        f.name, f.sz, bool(f.b), bool(f.i), bool(f.u),
        fcol.rgb if fcol is not None and getattr(fcol, 'type', None) == 'rgb' else
        ('theme%s:%s' % (fcol.theme, fcol.tint) if fcol is not None else None),
        f.strike,
        fill.patternType, fgv, bgv,
        a.horizontal, a.vertical, a.wrap_text, a.textRotation, a.indent,
        tuple(bds),
    )


SRC_WB = openpyxl.load_workbook(SRC)
DST_WB = openpyxl.load_workbook(DST)


def main():
    src, dst = SRC_WB, DST_WB
    report = []
    for name in src.sheetnames:
        ws, wd = src[name], dst[name]
        diffs = []

        # 表级属性
        tbl = []
        if sorted(str(m) for m in ws.merged_cells.ranges) != sorted(str(m) for m in wd.merged_cells.ranges):
            so, do = set(map(str, ws.merged_cells.ranges)), set(map(str, wd.merged_cells.ranges))
            tbl.append('合并单元格 %s: 少了%s 多了%s' % (name, sorted(so - do), sorted(do - so)))
        for attr in ('freeze_panes',):
            if getattr(ws, attr) != getattr(wd, attr):
                tbl.append('%s: %r -> %r' % (attr, getattr(ws, attr), getattr(wd, attr)))
        for attr in ('orientation', 'fitToWidth', 'paperSize'):
            if getattr(ws.page_setup, attr) != getattr(wd.page_setup, attr):
                tbl.append('page_setup.%s: %r -> %r' % (attr, getattr(ws.page_setup, attr),
                                                        getattr(wd.page_setup, attr)))
        if ws.auto_filter.ref != wd.auto_filter.ref:
            tbl.append('自动筛选范围: %r -> %r' % (ws.auto_filter.ref, wd.auto_filter.ref))
        if len(ws._images) != len(wd._images):
            tbl.append('图片数量: %d -> %d' % (len(ws._images), len(wd._images)))
        sv_attrs = ('zoomScale', 'showGridLines', 'tabSelected')
        for attr in sv_attrs:
            if getattr(ws.sheet_view, attr) != getattr(wd.sheet_view, attr):
                tbl.append('sheet_view.%s: %r -> %r' % (attr, getattr(ws.sheet_view, attr),
                                                        getattr(wd.sheet_view, attr)))
        try:
            cs = sum(len(v) for v in ws.conditional_formatting._cf_rules.values())
            cd = sum(len(v) for v in wd.conditional_formatting._cf_rules.values())
            if cs != cd:
                tbl.append('条件格式条数: %d -> %d' % (cs, cd))
        except Exception as e:
            tbl.append('条件格式无法比对: %s' % e)
        try:
            if len(ws.data_validations.dataValidation) != len(wd.data_validations.dataValidation):
                tbl.append('数据验证: %d -> %d' % (len(ws.data_validations.dataValidation),
                                                  len(wd.data_validations.dataValidation)))
        except Exception:
            pass

        # 列宽 / 隐藏（只比两表都有的列）
        maxc = min(ws.max_column, wd.max_column)
        def width_of(sheet_obj, letter):
            # 列可用 min..max 范围定义，直接索引才能解析到继承宽度
            try:
                d = sheet_obj.column_dimensions[letter]
                return (round(d.width, 2) if d.width else None, bool(d.hidden))
            except Exception:
                return (None, False)

        cold = []
        for ci in range(1, maxc + 1):
            L = openpyxl.utils.get_column_letter(ci)
            wa, ha = width_of(ws, L)
            wb_, hb = width_of(wd, L)
            if wa != wb_:
                cold.append('%s宽 %s->%s' % (L, wa, wb_))
            if ha != hb:
                cold.append('%s隐藏 %s->%s' % (L, ha, hb))

        # 逐格样式
        maxr = min(ws.max_row, wd.max_row)
        cell_diffs = []
        for r in range(1, maxr + 1):
            for ci in range(1, maxc + 1):
                cs_, cd_ = ws.cell(r, ci), wd.cell(r, ci)
                if cs_.value in (None, '') and cd_.value in (None, ''):
                    continue  # 空格子的样式不影响观感
                sa, sb = cell_sig(cs_), cell_sig(cd_)
                if sa != sb:
                    L = openpyxl.utils.get_column_letter(ci)
                    changed = [i for i, (x, y) in enumerate(zip(sa, sb)) if x != y]
                    cell_diffs.append('%s%d 变化字段%s: %r -> %r'
                                      % (L, r, changed, [sa[i] for i in changed], [sb[i] for i in changed]))
        report.append((name, tbl, cold, cell_diffs,
                       (ws.max_column, wd.max_column), (ws.max_row, wd.max_row)))
    return report


def aligned_report():
    """按「组」对齐比较：源表原有的 4 组迁移后是否保真（含格式）。
    坐标比对无意义，因为整块重排后列位置变了。"""
    MIGRATE = {
        'Intensive ': [('Null', 21, 21), ('ExtraDry', 23, 23),
                       ('Energy', 25, 25), ('StrWash', 27, 29)],
        'Glass ':     [('Null', 14, 14), ('ExtraDry', 16, 16),
                       ('Energy', 18, 18), ('StrWash', 20, 22)],
    }
    print('=== 原有 4 组「迁移对齐」后的差异（值 + 格式）===')
    for sheet, pairs in MIGRATE.items():
        ws, wd = SRC_WB[sheet], DST_WB[sheet]
        for gname, oc, nc in pairs:
            vd = fd = 0
            samples = []
            for off in (0, 1):
                for r in range(3, ws.max_row + 1):
                    so, do = ws.cell(r, oc + off), wd.cell(r, nc + off)
                    if so.value != do.value:
                        vd += 1
                        if len(samples) < 4:
                            samples.append('%s%d 值 %r->%r'
                                           % (openpyxl.utils.get_column_letter(oc + off), r,
                                              so.value, do.value))
                    if cell_sig(so) != cell_sig(do):
                        fd += 1
            print('  %-12s %-10s 列%s->%s  值差异%d  格式差异%d  %s'
                  % (sheet, gname, openpyxl.utils.get_column_letter(oc),
                     openpyxl.utils.get_column_letter(nc), vd, fd, samples))


if __name__ == '__main__':
    rep = main()
    FIELDS = ['数字格式', '字体名', '字号', '粗体', '斜体', '下划线', '字体色', '删除线',
              '填充模式', '前景色', '背景色', '水平对齐', '垂直对齐', '自动换行', '旋转', '缩进', '四边边框']
    total = 0
    for name, tbl, cold, cell_diffs, cols, rows in rep:
        flag = 'OK ' if not (tbl or cold or cell_diffs) else '差异'
        print('%-22s %s  列%d->%d 行%d->%d' % (name, flag, cols[0], cols[1], rows[0], rows[1]))
        for t in tbl:
            print('    [表级] ' + t)
        for c in cold[:12]:
            print('    [列] ' + c)
        if len(cold) > 12:
            print('    [列] ...共 %d 项' % len(cold))
        for d in cell_diffs[:8]:
            print('    [单元格] ' + d)
        if len(cell_diffs) > 8:
            print('    [单元格] ...共 %d 项' % len(cell_diffs))
        total += len(tbl) + len(cold) + len(cell_diffs)
    print()
    print('差异项合计:', total)
    print()
    aligned_report()
