#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""提取 DOCX 文档内容并保留结构

用法:
  python extract_docx.py <文件路径> [输出文件]

如果不指定输出文件，结果输出到 stdout。
"""

import zipfile
import re
import sys
from pathlib import Path


def extract_docx(filepath: str) -> str:
    """提取 DOCX 内容为 Markdown，保留标题层级和表格结构"""
    output = []

    with zipfile.ZipFile(filepath, 'r') as z:
        with z.open('word/document.xml') as f:
            content = f.read().decode('utf-8')

        # 按段落提取
        paragraphs = re.findall(r'<w:p[^>]*>(.*?)</w:p>', content, re.DOTALL)

        for para in paragraphs:
            # 提取文本
            texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', para)
            para_text = ''.join(texts).strip()

            if not para_text:
                continue

            # 检查样式
            style_match = re.search(r'w:pStyle w:val="([^"]*)"', para)
            if not style_match:
                style_match = re.search(r"w:pStyle w:val='([^']*)'", para)

            if style_match:
                style = style_match.group(1)
                heading_map = {
                    'Heading1': 1, 'Heading2': 2, 'Heading3': 3,
                    'Heading4': 4, 'Heading5': 5, 'Heading6': 6,
                    'Title': 1, 'Subtitle': 2
                }
                if style in heading_map:
                    level = heading_map[style]
                    output.append(f"{'#' * level} {para_text}\n")
                    continue

            # 检查章节标题模式
            if re.match(r'^第[一二三四五六七八九十]+章\s+', para_text):
                output.append(f"# {para_text}\n")
            elif re.match(r'^\d+\.\d+\.\d+\s+', para_text):
                output.append(f"### {para_text}\n")
            elif re.match(r'^\d+\.\d+\s+', para_text):
                output.append(f"## {para_text}\n")
            else:
                output.append(para_text)

    return '\n\n'.join(output)


def extract_docx_tables(filepath: str) -> list:
    """提取 DOCX 中的表格内容"""
    tables = []

    with zipfile.ZipFile(filepath, 'r') as z:
        with z.open('word/document.xml') as f:
            content = f.read().decode('utf-8')

        table_pattern = re.findall(r'<w:tbl[^>]*>(.*?)</w:tbl>', content, re.DOTALL)
        for table_xml in table_pattern:
            rows = re.findall(r'<w:tr[^>]*>(.*?)</w:tr>', table_xml, re.DOTALL)
            table_data = []
            for row in rows:
                cells = re.findall(r'<w:tc[^>]*>(.*?)</w:tc>', row, re.DOTALL)
                row_data = []
                for cell in cells:
                    texts = re.findall(r'<w:t[^>]*>([^<]*)</w:t>', cell)
                    row_data.append(''.join(texts).strip())
                if any(row_data):
                    table_data.append(row_data)
            if table_data:
                tables.append(table_data)

    return tables


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python extract_docx.py <文件路径> [输出文件]")
        sys.exit(1)

    filepath = sys.argv[1]
    if not Path(filepath).exists():
        print(f"错误：文件不存在 - {filepath}")
        sys.exit(1)

    markdown = extract_docx(filepath)

    if len(sys.argv) >= 3:
        output_file = sys.argv[2]
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(markdown)
        print(f"提取完成：{len(markdown)} 字符")
        print(f"已保存到 {output_file}")
    else:
        print(markdown)
