#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""读取校招模板 xlsx 的【所有 sheet】结构：表头行、列名、合并单元格、数据起始行。
仅用标准库。用法:
    python read_template.py <模板.xlsx> [输出dump.txt]
输出（可省略则打印到 stdout；中文建议重定向到 UTF-8 文件查看）：
    每个 sheet 的 行数、合并单元格、前若干行的单元格值(坐标=值)。
"""
import sys, zipfile, re
from xml.etree import ElementTree as ET

NSM = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'

def col_num(letters):
    n = 0
    for ch in letters:
        n = n*26 + (ord(ch)-64)
    return n

def dump(xlsx, out_path=None, max_rows=30):
    z = zipfile.ZipFile(xlsx)
    shared = []
    if 'xl/sharedStrings.xml' in z.namelist():
        root = ET.fromstring(z.read('xl/sharedStrings.xml'))
        for si in root.iter(NSM+'si'):
            shared.append(''.join(t.text or '' for t in si.iter(NSM+'t')))
    wb = ET.fromstring(z.read('xl/workbook.xml'))
    sheets = wb.findall(NSM+'sheets/'+NSM+'sheet')
    rels = ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))
    RNS = 'http://schemas.openxmlformats.org/package/2006/relationships'
    relmap = {r.get('Id'): r.get('Target') for r in rels.findall('{%s}Relationship' % RNS)}

    lines = []
    for sheet in sheets:
        rid = sheet.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        target = relmap[rid]
        sp = 'xl/' + target.lstrip('/')
        root = ET.fromstring(z.read(sp))
        rows = root.findall('.//'+NSM+'sheetData/'+NSM+'row')
        merges = root.find(NSM+'mergeCells')
        mstr = ''
        if merges is not None:
            mstr = '; '.join(mc.get('ref') for mc in merges.findall(NSM+'mergeCell'))
        lines.append(f"===== SHEET: {sheet.get('name')} ({sp}) rows={len(rows)} =====")
        lines.append('MERGES: ' + mstr)
        for row in rows[:max_rows]:
            rnum = row.get('r')
            cells = []
            for c in row.findall(NSM+'c'):
                ref = c.get('r'); t = c.get('t'); v = c.find(NSM+'v'); is_ = c.find(NSM+'is')
                if v is not None and v.text is not None:
                    val = shared[int(v.text)] if t == 's' else v.text
                elif is_ is not None:
                    val = ''.join(x.text or '' for x in is_.iter(NSM+'t'))
                else:
                    val = ''
                if val != '':
                    cells.append(f"{ref}={val}")
            lines.append('R'+rnum+': '+' | '.join(cells))
        lines.append('')
    text = '\n'.join(lines)
    if out_path:
        with open(out_path, 'w', encoding='utf-8') as f:
            f.write(text)
        print('written', out_path)
    else:
        sys.stdout.buffer.write(text.encode('utf-8'))

if __name__ == '__main__':
    if len(sys.argv) < 2:
        sys.exit("用法: python read_template.py <模板.xlsx> [输出dump.txt]")
    dump(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
