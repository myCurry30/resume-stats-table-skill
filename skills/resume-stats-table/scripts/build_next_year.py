#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""校招统计表生成（基于固定模板）。
- 以【固定模板.xlsx】为底（而非读取上一届）；
- 复制模板，保留所有 sheet 的表头行/合并单元格/列宽/样式/结构，清空具体人员数据；
- 在主 sheet 表头后写入候选人数据（来自 data.json）；
- 按岗位地点(location)重命名第一个 sheet 为「{location}岗」、面试名单 sheet 为「{location}面试名单」；
- 按用户确认的面试安排给「面试安排」sheet 表头 D2/E2 填时间（不确定则填“待定”）；
- 可选：清理其它 sheet 数据（只留表头）、保留时间框架清空姓名、整 sheet 保留（静态字典）。
仅用标准库。用法:
    python build_next_year.py <固定模板.xlsx> <输出.xlsx> <data.json> [config.json]
data.json 见 references/data_schema.md。config.json 见默认值 CONFIG_DEFAULTS。
"""
import sys, os, shutil, json, re, zipfile
from html import unescape as _unescape

NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'

CONFIG_DEFAULTS = {
    "location": "杭州",              # 岗位地点：用于重命名第一个 sheet 为「{location}岗」及意向城市
    "data_start_row": 4,             # 主 sheet 数据起始行（紧接表头）
    # 面试安排 sheet 表头：D2/E2 待填值（用户给的日期，如 {"D":"9.19（周五）","E":"9.20（周六）"}；不确定则填"待定"）
    "interview_schedule": {"D": "待定", "E": "待定"},
    # 表头行保留规则（键为固定模板里的原始 sheet 名；如缺省则以该 sheet 表头行数为准）
    "keep_header_rows": {"杭州岗": [2, 3], "杭州面试名单": [2], "面试安排": [2], "录用填写表": [1]},
    # 这些 sheet：保留时间框架行（keep_header_rows 之外的行）但清空姓名等列（D 及以后）
    "clear_tail_name_columns_in_sheets": ["面试安排"],
    # 这些 sheet 整表原样保留（静态字典等）
    "keep_whole_sheets": ["211&985院校名单"],
}

def esc(s):
    return ('' if s is None else str(s)).replace('&','&amp;').replace('<','&lt;').replace('>','&gt;').replace('"','&quot;')

def read_xlsx(xlsx):
    z = zipfile.ZipFile(xlsx)
    entries = {n: z.read(n) for n in z.namelist()}
    shared = []
    if 'xl/sharedStrings.xml' in entries:
        m = re.findall(r'<si>(.*?)</si>', entries['xl/sharedStrings.xml'].decode('utf-8'), re.S)
        def inner(si):
            return ''.join(re.findall(r'<t[^>]*>(.*?)</t>', si, re.S))
        shared = [inner(s) for s in m]
    return entries, shared

def col_num(ltrs):
    n = 0
    for ch in ltrs: n = n*26 + (ord(ch)-64)
    return n
def col_letter(n):
    s = ''; n = n+1
    while n > 0:
        n, r = divmod(n-1, 26)
        s = chr(65+r) + s
    return s

def get_rows(xml):
    """返回 [(行号, 完整row标签) ...]，正确处理自闭合空行 <row .../> 与普通 <row ...>...</row>。"""
    out = []
    for rm in re.finditer(r'<row r="(\d+)"[^>]*?(?:/>|>(.*?)</row>)', xml, re.S):
        out.append((int(rm.group(1)), rm.group(0)))
    return out

def read_sheet_map(wbx, relx):
    """返回 (ordered_sheet_names, name->filepath)"""
    relmap = dict(re.findall(r'<Relationship Id="(rId\d+)"[^>]*Target="([^"]*)"', relx))
    # 保持 workbook 中的顺序
    ordered = []
    files = {}
    for name, rid in re.findall(r'<sheet name="([^"]*)"[^>]*r:id="(rId\d+)"', wbx):
        name = _unescape(name)
        ordered.append(name)
        tgt = relmap.get(rid, '')
        files[name] = 'xl/' + tgt.lstrip('/')
    return ordered, files

def patch_sheet(xml, keep_rows, clear_tail_names=False, data_rows=None):
    m = re.search(r'(<sheetData>)(.*?)(</sheetData>)', xml, re.S)
    if not m:
        return xml
    before, after, body = xml[:m.start(2)], xml[m.end(2):], m.group(2)
    kept = []
    for rnum, full in get_rows(body):
        if rnum in keep_rows:
            kept.append(full)
        elif clear_tail_names:
            def repl(cm):
                ref = cm.group(1)
                if ref[0] in 'DEFGHIJKLMNOPQRSTUVWXYZ':
                    col = ref[0]; rr = ref[1:]
                    st = re.search(r's="(\d+)"', cm.group(0))
                    sattr = f' s="{st.group(1)}"' if st else ''
                    return f'<c r="{col}{rr}"{sattr} t="s"/>'
                return cm.group(0)
            kept.append(re.sub(r'<c r="([A-Z]+\d+)"[^>]*>.*?</c>', repl, full))
    new_body = ''.join(kept)
    if data_rows:
        new_body += ''.join(data_rows)
    xml2 = before + new_body + after
    last = max([r for r, _ in get_rows(body)] + ([style_row_num(data_rows)] if data_rows else []))
    return re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="A1:S{last}"/>', xml2, count=1)

def style_row_num(data_rows):
    if not data_rows: return 1
    r = re.search(r'<row r="(\d+)"', data_rows[0])
    return int(r.group(1)) if r else 1

def rename_sheets(wbx, location, ordered):
    """把第一个 sheet 改为 {location}岗；名为 杭州XX 的 sheet 中的 杭州 替换为 location。
    注意 group(1)=name, group(2)=sheetId(数字), group(3)=r:id(如 rId1)。"""
    new_wbx = wbx
    def repl_sheets(match):
        raw_name = match.group(1)
        name = _unescape(raw_name)
        sheetId = match.group(2)
        rid = match.group(3)
        new_name = name
        if ordered and name == ordered[0]:
            new_name = f"{location}岗"
        elif name.startswith("杭州"):
            new_name = location + name[2:]
        return f'<sheet name="{esc(new_name)}" sheetId="{sheetId}" r:id="{rid}"'
    # 匹配 <sheet name="..." sheetId="N" r:id="rIdX"/>
    return re.sub(r'<sheet name="([^"]*)" sheetId="(\d+)" r:id="(rId\d+)"', repl_sheets, wbx)

def fill_interview_header(xml, schedule):
    """把「面试安排」sheet 的表头 D2/E2（原 (待填写)）改为用户时间/待定，用内联字符串。
    兼容两种单元格写法：<c r="D2" ...>...</c> 与自闭合 <c r="D2" .../>。"""
    def repl_cell(cm):
        ref = cm.group(1) or cm.group(2)
        col = ref[0]; rr = ref[1:]
        if col in schedule and rr == '2':
            st = re.search(r's="(\d+)"', cm.group(0))
            sattr = f' s="{st.group(1)}"' if st else ''
            val = esc(schedule[col])
            return f'<c r="{col}{rr}"{sattr} t="inlineStr"><is><t>{val}</t></is></c>'
        return cm.group(0)
    # 匹配单个单元格：要么 <c r="X2"...>...</c>，要么 <c r="X2" .../>
    return re.sub(r'<c r="([A-Z]+2)"(?=[^>]*>)[^>]*>.*?</c>|<c r="([A-Z]+2)"[^>]*/>', repl_cell, xml)

def build(template, output, data_path, config_path=None):
    cfg = dict(CONFIG_DEFAULTS)
    if config_path:
        with open(config_path, encoding='utf-8') as f:
            cfg.update(json.load(f))
    with open(data_path, encoding='utf-8') as f:
        data = json.load(f)
    candidates = data['candidates']
    location = cfg.get('location', '杭州')

    shutil.copy2(template, output)
    entries, shared = read_xlsx(output)
    str_index = {}
    def add_str(s):
        if s is None: return None
        if s in str_index: return str_index[s]
        str_index[s] = len(shared); shared.append(s); return len(shared)-1

    wbx = entries['xl/workbook.xml'].decode('utf-8')
    relx = entries['xl/_rels/workbook.xml.rels'].decode('utf-8')
    ordered, sheet_files = read_sheet_map(wbx, relx)

    # 1) 重命名 sheet（按地点）
    new_wbx = rename_sheets(wbx, location, ordered)
    entries['xl/workbook.xml'] = new_wbx.encode('utf-8')
    # 重读映射（key 已是新名）
    ordered2, sheet_files2 = read_sheet_map(new_wbx, relx)

    main_sheet = ordered2[0]  # 第一个 sheet 即主表
    main_orig = ordered[0]    # 主表在固定模板里的原始名（重命名前），用于表头行/结构查找

    # 2) 主 sheet 样式来源：第一个含数据的非表头行
    main_xml = entries[sheet_files2[main_sheet]].decode('utf-8')
    hdr = set(cfg['keep_header_rows'].get(main_orig, cfg['keep_header_rows'].get(main_sheet, [2,3])))
    # 兜底：若 config 用原 sheet 名记录的 keep_header_rows，这里 main_sheet 是新名，需用原第一个名匹配
    src_style = {}
    src = None
    for rnum, full in get_rows(main_xml):
        if rnum not in hdr and re.search(r'<c r="[A-Z]"', full):
            src = full; break
    if src:
        for cm in re.finditer(r'<c r="([A-Z]+)\d+"(?:[^>]*)>', src):
            col = cm.group(1)
            st = re.search(r's="(\d+)"', cm.group(0))
            src_style[col] = st.group(1) if st else None

    def cell(col, row, val):
        ref = f'{col}{row}'
        s = src_style.get(col)
        sattr = f' s="{s}"' if s else ''
        if val is None or val == '' or val == 'null':
            return f'<c r="{ref}"{sattr} t="s"/>'
        if isinstance(val, (int, float)):
            return f'<c r="{ref}"{sattr}><v>{int(val)}</v></c>'
        idx = add_str(str(val))
        return f'<c r="{ref}"{sattr} t="s"><v>{idx}</v></c>'

    def emit_data_rows(start):
        rows_out = []
        col_map = data.get('column_map', {})
        for i, cand in enumerate(candidates, start=1):
            r = start + i - 1
            cells = []; used = set()
            if '序号' in cand:
                cells.append(cell('B', r, cand['序号'])); used.add('B')
            for fname, col in col_map.items():
                if col in used: continue
                if fname in cand:
                    cells.append(cell(col, r, cand[fname])); used.add(col)
            for ci in range(col_num('B'), col_num('S')+1):
                cl = col_letter(ci-1)
                if cl not in used:
                    cells.append(cell(cl, r, None))
            rows_out.append(f'<row r="{r}" customFormat="1" ht="24.75" customHeight="1">' + ''.join(cells) + '</row>')
        return rows_out
    data_rows = emit_data_rows(cfg['data_start_row'])

    # 3) 各 sheet 处理
    keep_whole = set(cfg.get('keep_whole_sheets', []))
    clear_tail = set(cfg.get('clear_tail_name_columns_in_sheets', []))
    # 表头行：优先用新名；对重命名过的主表回退到原始名 main_orig
    def keep_rows_for(name):
        if name in cfg['keep_header_rows']:
            return set(cfg['keep_header_rows'][name])
        if name == main_sheet and main_orig in cfg['keep_header_rows']:
            return set(cfg['keep_header_rows'][main_orig])
        return set()
    # keep_whole 当 sheet 名不变时直接命中（如 211&985院校名单）；若地点使名变化则不会命中，按普通 sheet 处理
    for name, sp in sheet_files2.items():
        xml = entries[sp].decode('utf-8')
        if name in keep_whole:
            continue
        if name == main_sheet:
            xml = patch_sheet(xml, keep_rows_for(name), data_rows=data_rows)
        elif name in clear_tail:
            xml = patch_sheet(xml, keep_rows_for(name), clear_tail_names=True)
        else:
            xml = patch_sheet(xml, keep_rows_for(name))
        entries[sp] = xml.encode('utf-8')

    # 4) 面试安排 sheet：填表头 D2/E2 时间（或待定）
    schedule = cfg.get('interview_schedule') or {}
    # 定位 面试安排 sheet（名字含「面试安排」）
    sched_sp = None
    for name, sp in sheet_files2.items():
        if '面试安排' in name:
            sched_sp = sp; break
    if sched_sp and schedule:
        entries[sched_sp] = fill_interview_header(entries[sched_sp].decode('utf-8'), schedule).encode('utf-8')

    # 5) 重建 sharedStrings
    ss = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          f'<sst {NS} count="{len(shared)}" uniqueCount="{len(shared)}">'
          + ''.join(f'<si><t>{esc(t)}</t></si>' for t in shared) + '</sst>')
    entries['xl/sharedStrings.xml'] = ss.encode('utf-8')

    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as z:
        for n, d in entries.items():
            z.writestr(n, d)
    print(f'WRITTEN {output} ({os.path.getsize(output)} bytes) candidates={len(candidates)} main_sheet={main_sheet}')

if __name__ == '__main__':
    if len(sys.argv) < 3:
        sys.exit("用法: python build_next_year.py <固定模板.xlsx> <输出.xlsx> <data.json> [config.json]")
    if len(sys.argv) < 4:
        sys.exit("需要 data.json")
    build(sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4] if len(sys.argv) > 4 else None)
