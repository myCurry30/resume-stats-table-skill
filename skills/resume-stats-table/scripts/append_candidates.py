#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""校招统计表：增量追加新增候选人到已生成的表格。

场景：已有 <X>届校招人员信息统计表.xlsx（由 build_next_year.py 生成），后续又来了新增简历，
希望把新增的简历信息自动填写进同一份表 —— 而不是重建全表。

功能：
- 读取【已生成】的 Excel（非固定模板），解析主 sheet(第一个) 已有的候选人；
- 读取新增数据 data.json（只含新增的那几份简历）；
- 按「姓名」或「联系方式」去重：已在表中的跳过，真正新增的追加到主表数据末尾下一空行；
- 序号自动续号（从已有最大序号 +1 开始）；
- 沿用主表已有数据行的样式/列结构，保证视觉一致；
- 只有主表被修改，其余 sheet（面试名单/面试安排/211&985院校名单/录用填写表）原样保留；
- **绝不覆盖原文件**：缺省时自动另存一个新副本（如 `2027届..._新增.xlsx`），原文件保持不变。

仅用标准库。用法:
    python append_candidates.py <已生成.xlsx> <新增data.json>
    可选第三参数指定输出名；缺省自动生成 `<原名>_新增.xlsx`（重名自动加序号）。
"""
import sys, os, re, zipfile
from html import unescape as _unescape

NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'

# 主表数据行 各列的单元格样式（取自真实 2027届 表的数据行；若实际模板列样式不同，脚本会尝试从已有数据行推断）
DEFAULT_COL_STYLE = {
    2: '22', 3: '23', 4: '24', 5: '23', 6: '23', 7: '23', 8: '23', 9: '23', 10: '23',
    11: '25', 12: '26', 13: '27', 14: '28', 15: '23', 16: '29', 17: '30', 18: '31', 19: '32',
}
# 主表数据行 的行级属性（取自真实 2027届 表）
ROW_ATTR = 'spans="2:19" s="51" customFormat="1" ht="24.75" customHeight="1"'


def esc(s):
    return ('' if s is None else str(s)).replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('"', '&quot;')


def col_num(ltrs):
    n = 0
    for ch in ltrs:
        n = n * 26 + (ord(ch) - 64)
    return n


def col_letter(n):
    s = ''
    n = n + 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def get_rows(xml):
    out = []
    for rm in re.finditer(r'<row r="(\d+)"[^>]*?(?:/>|>(.*?)</row>)', xml, re.S):
        out.append((int(rm.group(1)), rm.group(0)))
    return out


def read_xlsx(xlsx):
    z = zipfile.ZipFile(xlsx)
    entries = {n: z.read(n) for n in z.namelist()}
    shared = []
    if 'xl/sharedStrings.xml' in entries:
        m = re.findall(r'<si>(.*?)</si>', entries['xl/sharedStrings.xml'].decode('utf-8'), re.S)
        def inner(si):
            return ''.join(re.findall(r'<t[^>]*>(.*?)</t>', si, re.S))
        shared = [inner(x) for x in m]
    return entries, shared


def read_sheet_map(wbx, relx):
    relmap = dict(re.findall(r'<Relationship Id="(rId\d+)"[^>]*Target="([^"]*)"', relx))
    ordered = []
    files = {}
    for name, rid in re.findall(r'<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"', wbx):
        name = _unescape(name)
        ordered.append(name)
        tgt = relmap.get(rid, '')
        files[name] = 'xl/' + tgt.lstrip('/')
    return ordered, files


def get_cell_val(full_row, col_letter, shared):
    """取某行里指定列的原始值。col_letter 形如 'C'。
    - t="s" 的单元格：<v> 是共享字符串索引 → 解析成字符串；
    - 其余数值/内联单元格：直接返回 <v> 文本。"""
    m = re.search(r'<c r="%s\d+"(?:[^>]*)>(.*?)</c>' % col_letter, full_row, re.S)
    if not m:
        return None
    body = m.group(1)
    vm = re.search(r'<v>(.*?)</v>', body)
    if not vm:
        return None
    raw = vm.group(1)
    # 内联字符串
    tm = re.search(r'<t[^>]*>(.*?)</t>', body)
    if 'inlineStr' in body:
        return tm.group(1) if tm else ''
    # 共享字符串：只有 t="s" 才把 <v> 当索引
    if re.search(r't="s"', m.group(0)):
        try:
            return shared[int(raw)]
        except (ValueError, IndexError):
            return raw
    return raw


def auto_output_name(input_path):
    """为原文件自动生成一个新副本名（绝不覆盖原文件）。
    如 `2027届校招人员信息统计表.xlsx` → `2027届校招人员信息统计表_新增.xlsx`，
    若已存在则追加序号 `_新增2`、`_新增3`..."""
    base, ext = os.path.splitext(input_path)
    cand = f'{base}_新增{ext}'
    k = 2
    while os.path.exists(cand):
        cand = f'{base}_新增{k}{ext}'
        k += 1
    return cand


def append(template_xlsx, data_path, output=None):
    """output 缺省时自动生成新副本名（有后缀 `_新增`），原文件绝不改动。"""
    if output is None or output == template_xlsx:
        # 不允许覆盖原文件：缺省或误传成原文件时，一律另存新副本
        output = auto_output_name(template_xlsx)
        print(f'不覆盖原文件，输出到新副本: {output}')

    with open(data_path, encoding='utf-8') as f:
        data = json_load(f)
    new_cands = data.get('candidates', [])
    col_map = data.get('column_map', {})

    entries, shared = read_xlsx(template_xlsx)
    wbx = entries['xl/workbook.xml'].decode('utf-8')
    relx = entries['xl/_rels/workbook.xml.rels'].decode('utf-8')
    ordered, sheet_files = read_sheet_map(wbx, relx)
    main_sheet = ordered[0]
    main_sp = sheet_files[main_sheet]
    main_xml = entries[main_sp].decode('utf-8')

    # ---- 解析主表已有数据行 ----
    rows = get_rows(main_xml)
    # 找到表头行以外、C列(姓名)有值的行作为已有数据
    existing = []   # (行号, 序号, 姓名, 联系方式)
    max_seq = 0
    for rnum, full in rows:
        if rnum <= 3:
            continue
        name = get_cell_val(full, 'C', shared)
        phone = get_cell_val(full, 'F', shared)
        seq = get_cell_val(full, 'B', shared)
        if name:
            existing.append((rnum, seq, name, phone))
            try:
                if seq and int(seq) > max_seq:
                    max_seq = int(seq)
            except (ValueError, TypeError):
                pass
    existing_names = {n for _, _, n, _ in existing}
    existing_phones = {p for _, _, _, p in existing if p}
    print(f'已有候选人 {len(existing)} 名，最大序号={max_seq}，覆盖行 {[r for r,_,_,_ in existing]}')

    # ---- 过滤新增 ----
    to_add = []
    skipped = 0
    for c in new_cands:
        nm = str(c.get('姓名', '')).strip()
        ph = str(c.get('联系方式', '')).strip()
        if nm in existing_names or (ph and ph in existing_phones):
            skipped += 1
            print(f'  跳过(已存在): {nm or ph}')
            continue
        to_add.append(c)
    if not to_add:
        print('无新增候选人，未做任何修改。')
        return output

    # ---- 确定写入起始行与样式 ----
    start_row = max([r for r, _, _, _ in existing] or [0]) + 1
    if start_row <= 3:
        start_row = 4
    # 从已有的最后一条数据行继承列样式（若可推断），否则用默认
    col_style = {}
    if existing:
        last_full = dict((r, f) for r, f, _, _ in existing)[max([r for r, _, _, _ in existing])]
        for cm in re.finditer(r'<c r="([A-Z]+)(\d+)"(?: s="(\d+)")?', last_full):
            col_style[col_num(cm.group(1))] = cm.group(3)
    for ci, st in DEFAULT_COL_STYLE.items():
        col_style.setdefault(ci, st)

    # ---- 追加 data 行 ----
    str_index = {s: i for i, s in enumerate(shared)}
    def add_str(s):
        if s is None:
            return None
        if s in str_index:
            return str_index[s]
        str_index[s] = len(shared)
        shared.append(s)
        return len(shared) - 1

    new_rows_xml = []
    seq = max_seq
    for c in to_add:
        r = start_row
        start_row += 1
        seq += 1
        # 序号（自动续号，除非数据里给了明确的序号）
        cells = {}
        cells[2] = seq  # B 序号
        used = {2}
        for fname, col in col_map.items():
            ci = col_num(col)
            if ci in used:
                continue
            if col not in cells and fname in c:
                cells[ci] = c[fname]
                used.add(ci)
        # 明确提供序号则覆盖
        if '序号' in c:
            cells[2] = c['序号']
        # 生成单元格
        out_cells = []
        for ci in range(2, 20):  # B..S
            v = cells.get(ci)
            sattr = f' s="{col_style.get(ci)}"' if col_style.get(ci) else ''
            ref = f'{col_letter(ci-1)}{r}'
            if v is None or v == '' or v == 'null':
                out_cells.append(f'<c r="{ref}"{sattr} t="s"/>')
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                out_cells.append(f'<c r="{ref}"{sattr}><v>{int(v)}</v></c>')
            else:
                idx = add_str(str(v))
                out_cells.append(f'<c r="{ref}"{sattr} t="s"><v>{idx}</v></c>')
        new_rows_xml.append(f'<row r="{r}" {ROW_ATTR}>' + ''.join(out_cells) + '</row>')

    # ---- 写回主表 ----
    m = re.search(r'(<sheetData>)(.*?)(</sheetData>)', main_xml, re.S)
    new_body = m.group(2) + ''.join(new_rows_xml)
    new_main = main_xml[:m.start(2)] + new_body + main_xml[m.end(2):]
    last_row = start_row - 1
    new_main = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:S{last_row}"/>', new_main, count=1)
    entries[main_sp] = new_main.encode('utf-8')

    # ---- 重建 sharedStrings ----
    ss = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<sst {NS} count="{len(shared)}" uniqueCount="{len(shared)}">'
          + ''.join(f'<si><t>{esc(t)}</t></si>' for t in shared) + '</sst>')
    entries['xl/sharedStrings.xml'] = ss.encode('utf-8')

    # ---- 写出（另存新副本） ----
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, d in entries.items():
            z.writestr(n, d)
    print(f'WRITTEN {output} 新增={len(to_add)} 跳过={skipped} 新数据到 R{last_row}')
    return output


def json_load(f):
    import json
    return json.load(f)


if __name__ == '__main__':
    if len(sys.argv) < 3:
        sys.exit('用法: python append_candidates.py <已生成.xlsx> <新增data.json> [输出.xlsx(缺省自动另存新副本)]')
    out = append(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else None)
