#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文档审核脚本 v7.0 — Document Review Script

双层架构：
  L1 量化层（本脚本）：格式检查、文字错误检测、结构分析、数据提取
  L2 语义层（AI 模型）：逻辑冲突、事实核查、逻辑混乱（由 SKILL.md 驱动）

支持格式：.docx .pdf .pptx .txt .md
审核模式：quick / standard / deep
输出格式：markdown / json / word

用法:
  python review_document.py document.docx
  python review_document.py report.pdf --mode deep --output word --output-file report.docx
  python review_document.py doc1.docx doc2.pdf --batch --output json
  python review_document.py new.docx --diff old.docx
  python review_document.py doc.docx --exclude text_errors --glossary terms.json
"""

import argparse
import json
import os
import re
import subprocess
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any


# ──────────────────────────────────────────────
# 内容提取器
# ──────────────────────────────────────────────

class ContentExtractor:
    """多格式文档内容提取，支持多级回退"""

    @staticmethod
    def extract(filepath: str) -> Tuple[str, str]:
        """提取文档内容，返回 (内容, 格式类型)"""
        path = Path(filepath)
        if not path.exists():
            raise FileNotFoundError(f"文件不存在: {filepath}")

        ext = path.suffix.lower()
        extractors = {
            '.md': ContentExtractor._extract_text,
            '.txt': ContentExtractor._extract_text,
            '.docx': ContentExtractor._extract_docx,
            '.pdf': ContentExtractor._extract_pdf,
            '.pptx': ContentExtractor._extract_pptx,
        }

        extractor = extractors.get(ext)
        if not extractor:
            # 未知格式，尝试作为文本读取
            print(f"警告：未知格式 {ext}，尝试作为文本读取")
            return ContentExtractor._extract_text(filepath), ext

        try:
            content = extractor(filepath)
            if not content or not content.strip():
                print(f"警告：{ext} 提取内容为空，尝试原始方式")
                if ext in ('.docx', '.pdf', '.pptx'):
                    return ContentExtractor._extract_text_fallback(filepath), ext
            return content, ext
        except Exception as e:
            print(f"错误：提取失败 ({ext}): {e}")
            # 最后回退：尝试当文本读
            try:
                return ContentExtractor._extract_text_fallback(filepath), ext
            except Exception:
                return "", ext

    @staticmethod
    def _detect_encoding(filepath: str) -> str:
        """自动检测文件编码"""
        # 优先尝试 chardet
        try:
            import chardet
            with open(filepath, 'rb') as f:
                raw = f.read(65536)  # 读前 64KB 足够检测
            result = chardet.detect(raw)
            encoding = result.get('encoding', 'utf-8')
            if encoding:
                return encoding
        except ImportError:
            pass
        except Exception:
            pass

        # 回退：依次尝试常见编码
        for enc in ['utf-8', 'gbk', 'gb18030', 'utf-16', 'big5']:
            try:
                with open(filepath, 'r', encoding=enc) as f:
                    f.read(4096)
                return enc
            except (UnicodeDecodeError, UnicodeError):
                continue
        return 'utf-8'

    @staticmethod
    def _extract_text(filepath: str) -> str:
        """提取纯文本/Markdown"""
        encoding = ContentExtractor._detect_encoding(filepath)
        with open(filepath, 'r', encoding=encoding, errors='replace') as f:
            return f.read()

    @staticmethod
    def _extract_text_fallback(filepath: str) -> str:
        """文本回退方案"""
        encoding = ContentExtractor._detect_encoding(filepath)
        with open(filepath, 'r', encoding=encoding, errors='replace') as f:
            return f.read()

    @staticmethod
    def _extract_docx(filepath: str) -> str:
        """提取 Word 文档 — 三级回退：pandoc → python-docx → 原始XML"""
        # Level 1: pandoc
        try:
            result = subprocess.run(
                ['pandoc', '--track-changes=all', filepath, '-t', 'markdown'],
                capture_output=True, text=True, encoding='utf-8', timeout=60
            )
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass
        except Exception:
            pass

        # Level 2: python-docx
        try:
            from docx import Document
            doc = Document(filepath)
            parts = []
            for para in doc.paragraphs:
                text = para.text.strip()
                if text:
                    # 保留标题样式
                    if para.style and para.style.name:
                        style_name = para.style.name
                        heading_map = {
                            'Heading 1': '# ', 'Heading 2': '## ', 'Heading 3': '### ',
                            'Heading 4': '#### ', 'Heading 5': '##### ', 'Heading 6': '###### ',
                            'Title': '# ', 'Subtitle': '## '
                        }
                        prefix = heading_map.get(style_name, '')
                        parts.append(f"{prefix}{text}")
                    else:
                        parts.append(text)

            # 提取表格
            for table in doc.tables:
                for row in table.rows:
                    cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                    if cells:
                        parts.append(' | '.join(cells))
            return '\n\n'.join(parts)
        except ImportError:
            print("提示：未安装 python-docx，使用原始 XML 提取")
        except Exception as e:
            print(f"提示：python-docx 提取失败 ({e})，使用原始 XML")

        # Level 3: 原始 XML
        return ContentExtractor._extract_docx_raw(filepath)

    @staticmethod
    def _extract_docx_raw(filepath: str) -> str:
        """从 DOCX 原始 XML 提取文本"""
        import zipfile
        import xml.etree.ElementTree as ET
        try:
            with zipfile.ZipFile(filepath, 'r') as z:
                with z.open('word/document.xml') as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    ns = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
                    texts = []
                    for t in root.iter(f'{ns}t'):
                        if t.text:
                            texts.append(t.text)
                    return ' '.join(texts)
        except Exception as e:
            print(f"错误：DOCX 原始提取失败 - {e}")
            return ""

    @staticmethod
    def _extract_pdf(filepath: str) -> str:
        """提取 PDF 文档 — 三级回退：pymupdf → pdftotext → OCR"""
        # Level 1: pymupdf
        try:
            import fitz
            doc = fitz.open(filepath)
            text_parts = []
            for i, page in enumerate(doc):
                page_text = page.get_text()
                if page_text.strip():
                    text_parts.append(f"=== 第 {i+1} 页 ===\n{page_text}")
                else:
                    # 空页面，可能是扫描版
                    text_parts.append(f"=== 第 {i+1} 页 ===\n[此页无可提取文本，可能为扫描图片]")
            doc.close()
            full_text = '\n\n'.join(text_parts)

            # 检测是否大量页面为空（扫描版 PDF）
            empty_pages = sum(1 for p in text_parts if "可能为扫描图片" in p)
            total_pages = len(text_parts)
            if total_pages > 0 and empty_pages / total_pages > 0.5:
                print(f"提示：检测到 {empty_pages}/{total_pages} 页为扫描图片，尝试 OCR...")
                ocr_text = ContentExtractor._ocr_pdf(filepath)
                if ocr_text:
                    return ocr_text
            return full_text
        except ImportError:
            print("提示：未安装 pymupdf，尝试 pdftotext")
        except Exception as e:
            print(f"提示：pymupdf 提取失败 ({e})，尝试 pdftotext")

        # Level 2: pdftotext
        try:
            result = subprocess.run(
                ['pdftotext', '-layout', filepath, '-'],
                capture_output=True, text=True, encoding='utf-8', timeout=120
            )
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass
        except Exception:
            pass

        # Level 3: OCR
        print("提示：常规提取失败，尝试 OCR...")
        return ContentExtractor._ocr_pdf(filepath)

    @staticmethod
    def _ocr_pdf(filepath: str) -> str:
        """OCR 提取 PDF 文本"""
        try:
            import fitz
            import pytesseract
            from PIL import Image
            import io

            doc = fitz.open(filepath)
            text_parts = []
            for i, page in enumerate(doc):
                # 渲染页面为图片
                pix = page.get_pixmap(dpi=200)
                img_data = pix.tobytes("png")
                image = Image.open(io.BytesIO(img_data))
                # OCR 识别
                text = pytesseract.image_to_string(image, lang='chi_sim+eng')
                if text.strip():
                    text_parts.append(f"=== 第 {i+1} 页 (OCR) ===\n{text}")
            doc.close()
            if text_parts:
                return '\n\n'.join(text_parts)
        except ImportError:
            print("提示：OCR 需要 pytesseract 和 pillow，请运行 pip install pytesseract pillow")
        except Exception as e:
            print(f"提示：OCR 失败 - {e}")
        return ""

    @staticmethod
    def _extract_pptx(filepath: str) -> str:
        """提取 PPT 文档内容"""
        try:
            from pptx import Presentation
            prs = Presentation(filepath)
            parts = []
            for i, slide in enumerate(prs.slides):
                slide_texts = []
                slide_texts.append(f"=== 第 {i+1} 页 ===")

                for shape in slide.shapes:
                    # 文本框
                    if hasattr(shape, 'text') and shape.text.strip():
                        slide_texts.append(shape.text.strip())

                    # 表格
                    if shape.has_table:
                        table = shape.table
                        for row in table.rows:
                            cells = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                            if cells:
                                slide_texts.append(' | '.join(cells))

                # 备注
                if slide.has_notes_slide:
                    notes = slide.notes_slide.notes_text_frame.text.strip()
                    if notes:
                        slide_texts.append(f"[备注] {notes}")

                parts.append('\n'.join(slide_texts))
            return '\n\n'.join(parts)
        except ImportError:
            print("提示：未安装 python-pptx，无法提取 PPT")
            return ""
        except Exception as e:
            print(f"错误：PPT 提取失败 - {e}")
            return ""


# ──────────────────────────────────────────────
# 文档结构分析器
# ──────────────────────────────────────────────

class DocumentParser:
    """解析文档结构，提取章节、数据、时间线等"""

    def __init__(self, content: str):
        self.content = content
        self.lines = content.split('\n')
        self.sections: List[Dict] = []
        self.data_points: List[Dict] = []
        self.timeline: List[Dict] = []
        self.terms: Dict[str, List[str]] = {}

    def parse(self) -> None:
        """执行全部解析"""
        self._parse_sections()
        self._extract_data_points()
        self._extract_timeline()

    def _parse_sections(self) -> None:
        """解析章节结构"""
        current_section = {"title": "引言", "level": 0, "content": [], "line_start": 0, "line_end": 0}

        for i, line in enumerate(self.lines):
            match = re.match(r'^(#{1,6})\s+(.+)$', line.strip())
            if match:
                level = len(match.group(1))
                title = match.group(2).strip()

                if current_section["content"]:
                    current_section["line_end"] = i - 1
                    self.sections.append(current_section)

                current_section = {
                    "title": title, "level": level,
                    "content": [], "line_start": i, "line_end": i
                }
            else:
                current_section["content"].append(line)

        if current_section["content"]:
            current_section["line_end"] = len(self.lines) - 1
            self.sections.append(current_section)

    def _extract_data_points(self) -> None:
        """提取文档中的数值数据"""
        # 匹配百分比、金额、数量等
        patterns = [
            (r'(\d+(?:\.\d+)?)\s*%', 'percentage'),
            (r'(\d+(?:,\d{3})*(?:\.\d+)?)\s*[万亿]元', 'amount'),
            (r'(\d+(?:,\d{3})*)\s*(?:人|名|位)', 'count'),
            (r'(\d+(?:\.\d+)?)\s*倍', 'ratio'),
        ]

        for section in self.sections:
            text = '\n'.join(section["content"])
            for pattern, dtype in patterns:
                for match in re.finditer(pattern, text):
                    self.data_points.append({
                        "type": dtype,
                        "value": match.group(1),
                        "raw": match.group(0),
                        "section": section["title"],
                        "context": text[max(0, match.start()-20):match.end()+20].strip()
                    })

    def _extract_timeline(self) -> None:
        """提取时间节点"""
        date_patterns = [
            r'(\d{4}年\d{1,2}月\d{1,2}日)',
            r'(\d{4}-\d{1,2}-\d{1,2})',
            r'(\d{4}年\d{1,2}月)',
            r'(\d{4}年)',
            r'(Q[1-4]\s*\d{4})',
            r'(\d{4}年[上下]半年)',
        ]

        for section in self.sections:
            text = '\n'.join(section["content"])
            for pattern in date_patterns:
                for match in re.finditer(pattern, text):
                    self.timeline.append({
                        "date": match.group(1),
                        "section": section["title"],
                        "context": text[max(0, match.start()-15):match.end()+15].strip()
                    })

    def get_summary(self) -> Dict:
        """获取文档结构摘要，供 AI 进一步分析"""
        return {
            "total_sections": len(self.sections),
            "total_chars": len(self.content),
            "section_titles": [
                {"title": s["title"], "level": s["level"], "chars": len('\n'.join(s["content"]))}
                for s in self.sections
            ],
            "data_points_count": len(self.data_points),
            "data_points": self.data_points[:50],  # 限制数量
            "timeline_count": len(self.timeline),
            "timeline": self.timeline[:30],
        }


# ──────────────────────────────────────────────
# L1 量化检查器
# ──────────────────────────────────────────────

class QuantitativeChecker:
    """L1 量化层检查器：格式、文字错误、结构分析"""

    VAGUE_WORDS = [
        "可能", "大概", "也许", "左右", "一些", "某些", "相关", "等", "等等",
        "基本上", "大体上", "总体上", "相对", "比较", "较为", "相当",
        "似乎", "好像", "仿佛", "某种意义上", "某种程度上",
        "觉得", "认为", "应该", "挺好", "更好", "不错"
    ]

    # 高频错别字对照表
    COMMON_TYPOS = {
        "已得": "赢得", "象限": "选项", "决对": "绝对", "布署": "部署",
        "事绩": "业绩", "收届": "受到", "桃战": "挑战", "邦定": "绑定",
        "针对于": "针对", "大大的": "大大", "做为": "作为",
        "既使": "即使", "那怕": "哪怕", "竟争": "竞争",
        "按装": "安装", "甘拜下风": "甘拜下风", "-default": "default",
        "帐号": "账号", "帐户": "账户", "登陆": "登录",
        "辨识": "辨识", "纷至踏来": "纷至沓来",
        "默守成规": "墨守成规", "一愁莫展": "一筹莫展",
    }

    def __init__(self, parser: DocumentParser, content: str,
                 exclude: List[str] = None, glossary: Dict = None):
        self.parser = parser
        self.content = content
        self.exclude = exclude or []
        self.glossary = glossary or {}
        self.issues = {
            "format": [],
            "text_errors": [],
            "logic": [],
            "clarity": [],
            "feasibility": []
        }
        self.scores = {
            "format": 100, "logic": 100,
            "clarity": 100, "feasibility": 100
        }

    def run_all(self) -> None:
        """执行所有 L1 检查"""
        if "format" not in self.exclude:
            self.check_format()
        if "text_errors" not in self.exclude:
            self.check_text_errors()
        if "logic" not in self.exclude:
            self.check_logic_basic()
        if "clarity" not in self.exclude:
            self.check_clarity()
        if "feasibility" not in self.exclude:
            self.check_feasibility()
        if "glossary" not in self.exclude and self.glossary:
            self.check_terminology()
        self._calculate_scores()

    def check_format(self) -> None:
        """检查文档格式"""
        issues = []

        # 1. 标题层级跳跃
        prev_level = 0
        for section in self.parser.sections:
            level = section["level"]
            if level > 0:
                if prev_level > 0 and level > prev_level + 1:
                    issues.append({
                        "type": "format",
                        "severity": "中",
                        "description": f"标题层级跳跃（从 H{prev_level} 直接到 H{level}）",
                        "location": f"{section['title']}",
                        "suggestion": f"建议使用 H{prev_level + 1} 作为中间层级",
                        "status": "待处理"
                    })
                prev_level = level

        # 2. 段落长度
        for section in self.parser.sections:
            para_text = '\n'.join(section["content"])
            paras = [p for p in para_text.split('\n\n') if len(p.strip()) > 50]
            for para in paras:
                if len(para) > 400:
                    issues.append({
                        "type": "format",
                        "severity": "低",
                        "description": f"段落过长（{len(para)} 字）",
                        "location": f"{section['title']}",
                        "suggestion": "建议拆分为 200-300 字的短段落",
                        "status": "待处理"
                    })

        # 3. 结构完整性
        has_intro = any(
            kw in s["title"] for s in self.parser.sections
            for kw in ["引言", "简介", "背景", "概述"]
        )
        has_conclusion = any(
            kw in s["title"] for s in self.parser.sections
            for kw in ["结论", "总结", "展望", "结语"]
        )

        if not has_intro and len(self.parser.sections) > 2:
            issues.append({
                "type": "format", "severity": "中",
                "description": "文档缺少引言或背景说明",
                "location": "文档开头",
                "suggestion": "添加引言章节，说明文档目的和背景",
                "status": "待处理"
            })

        if not has_conclusion and len(self.parser.sections) > 2:
            issues.append({
                "type": "format", "severity": "中",
                "description": "文档缺少结论或总结",
                "location": "文档结尾",
                "suggestion": "添加结论章节，总结核心观点和建议",
                "status": "待处理"
            })

        # 4. 列表格式一致性
        list_markers = set()
        for section in self.parser.sections:
            text = '\n'.join(section["content"])
            for m in re.finditer(r'^[\s]*([-*+])\s+', text, re.MULTILINE):
                list_markers.add(m.group(1))
            for m in re.finditer(r'^[\s]*(\d+\.)\s+', text, re.MULTILINE):
                list_markers.add(m.group(1))

        if len(list_markers) > 1:
            issues.append({
                "type": "format", "severity": "低",
                "description": f"列表标记不统一（混用 {', '.join(list_markers)}）",
                "location": "多处",
                "suggestion": "统一使用一种列表标记",
                "status": "待处理"
            })

        self.issues["format"] = issues

    def check_text_errors(self) -> None:
        """检查文字错误"""
        issues = []

        # 1. 高频错别字
        for wrong, correct in self.COMMON_TYPOS.items():
            if wrong in self.content:
                # 找到位置
                idx = self.content.find(wrong)
                context = self.content[max(0, idx-20):idx+len(wrong)+20]
                section_name = self._find_section(idx)
                issues.append({
                    "type": "text_errors", "severity": "中",
                    "description": f"错别字：「{wrong}」应为「{correct}」",
                    "location": section_name,
                    "context": context.strip(),
                    "suggestion": f"将「{wrong}」改为「{correct}」",
                    "status": "待处理"
                })

        # 2. 标点错误
        # 中文文档中混用英文标点
        cn_text = re.findall(r'[\u4e00-\u9fff]+', self.content)
        if cn_text:
            # 检查中文后跟英文逗号/句号
            for match in re.finditer(r'[\u4e00-\u9fff]+([,.;:!?])', self.content):
                punct = match.group(1)
                cn_punct_map = {',': '，', '.': '。', ';': '；', ':': '：', '!': '！', '?': '？'}
                correct = cn_punct_map.get(punct)
                if correct:
                    section_name = self._find_section(match.start())
                    issues.append({
                        "type": "text_errors", "severity": "低",
                        "description": f"中文语境中使用了英文标点「{punct}」",
                        "location": section_name,
                        "context": match.group(0),
                        "suggestion": f"建议改为中文标点「{correct}」",
                        "status": "待处理"
                    })

        # 3. 引号不匹配
        quote_pairs = [('「', '」'), ('『', '』'), ('"', '"'), ('"', '"')]
        for open_q, close_q in quote_pairs:
            open_count = self.content.count(open_q)
            close_count = self.content.count(close_q)
            if open_count != close_count:
                issues.append({
                    "type": "text_errors", "severity": "中",
                    "description": f"引号不匹配：「{open_q}」{open_count}个 vs 「{close_q}」{close_count}个",
                    "location": "全文",
                    "suggestion": "检查引号是否成对出现",
                    "status": "待处理"
                })

        # 4. 重复词
        for match in re.finditer(r'(.{2,4})\1{1,}', self.content):
            word = match.group(1)
            # 排除正常重复（如"来了来了"、"很多很多"等合理用法）
            if word not in ['很多', '非常', '确实', '真的', '一直', '还是', '可以', '这个']:
                section_name = self._find_section(match.start())
                issues.append({
                    "type": "text_errors", "severity": "低",
                    "description": f"疑似重复词：「{word}{word}」",
                    "location": section_name,
                    "context": match.group(0),
                    "suggestion": f"检查是否为笔误或冗余重复",
                    "status": "待处理"
                })

        self.issues["text_errors"] = issues

    def check_logic_basic(self) -> None:
        """L1 初步逻辑检查（论据匹配、过渡检测）

        注意：深度逻辑冲突检测由 L2 AI 语义层完成。
        此处仅做初步筛选，为 AI 提供线索。
        """
        issues = []

        argument_indicators = ["应该", "必须", "建议", "需要", "因此", "所以", "由此可见"]
        evidence_indicators = ["数据", "显示", "表明", "根据", "研究", "统计", "例如", "比如", "调查"]

        for section in self.parser.sections:
            text = '\n'.join(section["content"])
            has_argument = any(word in text for word in argument_indicators)
            has_evidence = any(word in text for word in evidence_indicators)

            if has_argument and not has_evidence and len(text) > 100:
                issues.append({
                    "type": "logic", "severity": "高",
                    "description": "提出论点但缺乏数据或案例支撑",
                    "location": f"{section['title']}",
                    "suggestion": "补充相关数据、调研结果或实际案例来支撑论点",
                    "status": "待处理",
                    "hint_for_ai": f"章节「{section['title']}」包含论点性表述但缺少证据关键词，建议 AI 重点检查"
                })

        # 数据一致性初步检查
        data_by_value = {}
        for dp in self.parser.data_points:
            key = dp["value"]
            if key not in data_by_value:
                data_by_value[key] = []
            data_by_value[key].append(dp)

        # 同一数值在不同章节出现是正常的；不同数值描述同一指标才有问题
        # 这里为 AI 提取所有数据点供交叉验证
        if len(self.parser.data_points) > 0:
            issues.append({
                "type": "logic", "severity": "信息",
                "description": f"文档包含 {len(self.parser.data_points)} 个数值数据点，建议 AI 交叉验证一致性",
                "location": "全文",
                "suggestion": "检查同一指标在不同位置的数值是否一致",
                "status": "待处理",
                "hint_for_ai": "见 data_points 字段"
            })

        self.issues["logic"] = issues

    def check_clarity(self) -> None:
        """检查表述清晰度"""
        issues = []
        vague_locations = []

        for section in self.parser.sections:
            text = '\n'.join(section["content"])

            # 1. 模糊表述
            for word in self.VAGUE_WORDS:
                count = text.count(word)
                if count > 0:
                    vague_locations.append({
                        "section": section["title"],
                        "word": word,
                        "count": count
                    })

            # 2. 句子长度
            sentences = re.split(r'[。！？.!?\n]', text)
            long_sentences = [s.strip() for s in sentences if len(s.strip()) > 60]
            very_long = [s.strip() for s in sentences if len(s.strip()) > 100]

            if very_long:
                issues.append({
                    "type": "clarity", "severity": "中",
                    "description": f"存在{len(very_long)}个超长句（超过100字）",
                    "location": f"{section['title']}",
                    "suggestion": "将超长句拆分为 2-3 个短句",
                    "status": "待处理"
                })
            elif len(long_sentences) > 3:
                issues.append({
                    "type": "clarity", "severity": "低",
                    "description": f"存在{len(long_sentences)}个长句（超过60字）",
                    "location": f"{section['title']}",
                    "suggestion": "考虑拆分长句以提高可读性",
                    "status": "待处理"
                })

        # 汇总模糊表述
        if len(vague_locations) > 3:
            total_count = sum(v["count"] for v in vague_locations)
            sample = vague_locations[:5]
            sample_str = ', '.join(f"{v['section']}节(使用'{v['word']}'×{v['count']})" for v in sample)
            issues.append({
                "type": "clarity", "severity": "中",
                "description": f"使用较多模糊词汇（共{total_count}处），包括：{sample_str}",
                "location": "多处章节",
                "suggestion": "将模糊表述改为具体数据或明确陈述",
                "status": "待处理"
            })

        self.issues["clarity"] = issues

    def check_feasibility(self) -> None:
        """检查方案可行性"""
        issues = []
        text = self.content.lower()

        # 1. 执行步骤
        action_keywords = ["步骤", "流程", "计划", "方案", "实施", "执行", "操作"]
        detail_keywords = ["第一", "第二", "首先", "然后", "接着", "最后", "step", "phase", "阶段"]
        has_actions = any(w in text for w in action_keywords)
        has_details = any(w in text for w in detail_keywords)

        if not has_actions or not has_details:
            issues.append({
                "type": "feasibility", "severity": "高",
                "description": "文档未明确说明具体的执行步骤或操作流程",
                "location": "全文",
                "suggestion": "添加详细的实施计划，包括具体步骤、责任人和时间节点",
                "status": "待处理"
            })

        # 2. 资源需求
        resource_keywords = ["资源", "预算", "成本", "人员", "时间", "设备", "资金"]
        vague_resource = ["一些", "某些", "相关", "大概", "左右", "即将", "不久"]
        has_resources = any(w in text for w in resource_keywords)
        has_vague_resources = any(w in text for w in vague_resource)

        if not has_resources or (has_vague_resources and not has_resources):
            issues.append({
                "type": "feasibility", "severity": "高",
                "description": "未说明所需资源，或仅使用模糊表述",
                "location": "全文",
                "suggestion": "补充具体资源需求清单",
                "status": "待处理"
            })

        # 3. 时间规划
        time_keywords = ["202", "203", "个月", "周", "天", "日", "季度", "半年", "一年", "Q1", "Q2", "Q3", "Q4"]
        vague_time = ["将来", "未来", "不久", "即将", "到时候", "适时", "择机"]
        has_time = any(w in text for w in time_keywords)
        has_vague_time = any(w in text for w in vague_time)

        if not has_time or has_vague_time:
            issues.append({
                "type": "feasibility", "severity": "中",
                "description": "时间规划不明确",
                "location": "全文",
                "suggestion": "使用具体时间节点",
                "status": "待处理"
            })

        # 4. 风险评估
        risk_keywords = ["风险", "挑战", "困难", "障碍", "应对", "预案", "mitigation"]
        if not any(w in text for w in risk_keywords):
            issues.append({
                "type": "feasibility", "severity": "中",
                "description": "缺少风险评估和应对预案",
                "location": "全文",
                "suggestion": "添加风险分析章节",
                "status": "待处理"
            })

        # 5. 成功指标
        metric_keywords = ["指标", "目标", "kpi", "衡量", "评估", "验收", "成功标准", "%", "提升", "降低"]
        if not any(w in text for w in metric_keywords):
            issues.append({
                "type": "feasibility", "severity": "中",
                "description": "未定义可量化的成功指标",
                "location": "全文",
                "suggestion": "设定可量化的目标",
                "status": "待处理"
            })

        self.issues["feasibility"] = issues

    def check_terminology(self) -> None:
        """检查术语一致性（使用自定义术语表）"""
        issues = []
        terms = self.glossary.get("terms", {})
        preferred = self.glossary.get("preferred", {})

        for canonical, variants in terms.items():
            found_variants = set()
            for variant in variants:
                if variant.lower() in self.content.lower():
                    found_variants.add(variant)

            # 如果同一个概念出现了多种写法
            if len(found_variants) > 1:
                preferred_form = preferred.get(canonical, canonical)
                issues.append({
                    "type": "text_errors", "severity": "中",
                    "description": f"术语「{canonical}」存在多种写法：{', '.join(found_variants)}",
                    "location": "多处",
                    "suggestion": f"统一使用「{preferred_form}」",
                    "status": "待处理"
                })

        self.issues["text_errors"].extend(issues)

    def _find_section(self, char_index: int) -> str:
        """根据字符位置找到所在章节"""
        cumulative = 0
        for section in self.parser.sections:
            section_len = len('\n'.join(section["content"]))
            if cumulative + section_len > char_index:
                return section["title"]
            cumulative += section_len
        return "未知位置"

    def _calculate_scores(self) -> None:
        """根据问题计算评分"""
        penalty = {"高": 15, "中": 8, "低": 3, "信息": 0}
        type_to_score = {
            "format": "format",
            "text_errors": "clarity",  # 文字错误影响清晰度
            "logic": "logic",
            "clarity": "clarity",
            "feasibility": "feasibility"
        }

        for issue_type, issues in self.issues.items():
            score_key = type_to_score.get(issue_type, "format")
            for issue in issues:
                sev = issue.get("severity", "低")
                self.scores[score_key] -= penalty.get(sev, 0)

        for key in self.scores:
            self.scores[key] = max(0, min(100, self.scores[key]))


# ──────────────────────────────────────────────
# 增量审核
# ──────────────────────────────────────────────

class DiffReviewer:
    """增量审核：对比新旧版本，仅审核变更部分"""

    @staticmethod
    def diff(old_content: str, new_content: str) -> Dict:
        """对比两个版本的内容差异"""
        old_lines = old_content.split('\n')
        new_lines = new_content.split('\n')

        added = []
        removed = []
        modified = []

        # 简单的逐行对比（非最优，但对文档审核足够）
        old_set = set(l.strip() for l in old_lines if l.strip())
        new_set = set(l.strip() for l in new_lines if l.strip())

        added_lines = new_set - old_set
        removed_lines = old_set - new_set
        matched_added = set()
        matched_removed = set()

        # 找出修改的行（内容相近的增删对）
        for new_line in list(added_lines):
            for old_line in list(removed_lines):
                if old_line in matched_removed:
                    continue
                if DiffReviewer._similarity(old_line, new_line) > 0.5:
                    modified.append({"old": old_line, "new": new_line})
                    matched_added.add(new_line)
                    matched_removed.add(old_line)
                    break

        added = list(added_lines - matched_added)
        removed = list(removed_lines - matched_removed)

        return {
            "added": added,
            "removed": removed,
            "modified": modified,
            "added_count": len(added),
            "removed_count": len(removed),
            "modified_count": len(modified),
            "has_changes": len(added) + len(removed) + len(modified) > 0
        }

    @staticmethod
    def _similarity(a: str, b: str) -> float:
        """简单的字符串相似度计算"""
        if not a or not b:
            return 0
        # 使用字符集交集比例
        set_a = set(a)
        set_b = set(b)
        intersection = set_a & set_b
        union = set_a | set_b
        return len(intersection) / len(union) if union else 0


# ──────────────────────────────────────────────
# 报告生成器
# ──────────────────────────────────────────────

class ReportGenerator:
    """生成审核报告"""

    @staticmethod
    def generate(reviewer, mode: str, output_format: str = "markdown") -> str:
        """生成审核报告"""
        if output_format == "json":
            return ReportGenerator._generate_json(reviewer, mode)
        elif output_format == "word":
            return ReportGenerator._generate_word(reviewer, mode)
        else:
            return ReportGenerator._generate_markdown(reviewer, mode)

    @staticmethod
    def _generate_markdown(reviewer, mode: str) -> str:
        lines = []
        lines.append("# 文档审核报告\n")
        lines.append(f"**审核文档**: {reviewer.document_path}")
        lines.append(f"**审核时间**: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
        lines.append(f"**审核模式**: {mode}")
        lines.append(f"**文档规模**: {len(reviewer.content)} 字符 / {len(reviewer.parser.sections)} 章节\n")

        # 评分概览
        total_score = sum(reviewer.checker.scores.values()) / 4
        lines.append("## 📊 评分概览\n")
        lines.append("| 维度 | 评分 | 等级 |")
        lines.append("|------|------|------|")
        dim_names = {'format': '格式规范性', 'logic': '内容逻辑性', 'clarity': '表述清晰度', 'feasibility': '方案可行性'}
        for key in ['format', 'logic', 'clarity', 'feasibility']:
            score = reviewer.checker.scores[key]
            grade = ReportGenerator._get_grade(score)
            lines.append(f"| {dim_names[key]} | {score} | {grade} |")
        grade = ReportGenerator._get_grade(total_score)
        lines.append(f"| **综合** | **{round(total_score, 1)}** | **{grade}** |\n")

        # 问题概览
        lines.append("## 📋 问题概览\n")
        lines.append("| 问题类型 | 高 | 中 | 低 | 合计 |")
        lines.append("|----------|---|---|---|------|")
        type_names = {
            "format": "格式问题", "text_errors": "文字错误",
            "logic": "逻辑问题", "clarity": "表述问题", "feasibility": "可行性问题"
        }
        total_high = total_mid = total_low = total_count = 0
        for itype in ["format", "text_errors", "logic", "clarity", "feasibility"]:
            issues = reviewer.checker.issues.get(itype, [])
            high = sum(1 for i in issues if i.get("severity") == "高")
            mid = sum(1 for i in issues if i.get("severity") == "中")
            low = sum(1 for i in issues if i.get("severity") == "低")
            count = len(issues)
            lines.append(f"| {type_names[itype]} | {high} | {mid} | {low} | {count} |")
            total_high += high
            total_mid += mid
            total_low += low
            total_count += count
        lines.append(f"| **合计** | **{total_high}** | **{total_mid}** | **{total_low}** | **{total_count}** |\n")

        # 详细问题
        lines.append("## 🔍 详细问题\n")
        for itype in ["format", "text_errors", "logic", "clarity", "feasibility"]:
            issues = reviewer.checker.issues.get(itype, [])
            if issues:
                lines.append(f"### {type_names[itype]}\n")
                for i, issue in enumerate(issues, 1):
                    sev = issue.get("severity", "低")
                    lines.append(f"#### [{sev}] {issue['description']}")
                    lines.append(f"- **位置**: {issue.get('location', '未知')}")
                    if issue.get("context"):
                        lines.append(f"- **上下文**: {issue['context']}")
                    lines.append(f"- **建议**: {issue.get('suggestion', '')}")
                    lines.append(f"- **状态**: {issue.get('status', '待处理')}\n")

        # AI 分析提示
        lines.append("## 🤖 AI 分析数据\n")
        lines.append("以下数据供 AI 进行 L2 语义层分析：\n")
        summary = reviewer.parser.get_summary()
        lines.append(f"- 章节数: {summary['total_sections']}")
        lines.append(f"- 数据点: {summary['data_points_count']}")
        lines.append(f"- 时间节点: {summary['timeline_count']}")
        if summary["data_points"]:
            lines.append("\n**数据点摘要**（前10个）:")
            for dp in summary["data_points"][:10]:
                lines.append(f"  - [{dp['section']}] {dp['raw']}（上下文: {dp['context'][:40]}）")
        if summary["timeline"]:
            lines.append("\n**时间线**（前10个）:")
            for tl in summary["timeline"][:10]:
                lines.append(f"  - [{tl['section']}] {tl['date']}（上下文: {tl['context'][:40]}）")
        lines.append("")

        # 改进建议
        all_suggestions = []
        for issues in reviewer.checker.issues.values():
            for issue in issues:
                if issue.get("severity") in ("高", "中") and issue.get("suggestion") not in all_suggestions:
                    all_suggestions.append(issue["suggestion"])

        if all_suggestions:
            lines.append("## 💡 综合改进建议\n")
            high_suggestions = []
            mid_suggestions = []
            for issues in reviewer.checker.issues.values():
                for issue in issues:
                    if issue.get("severity") == "高" and issue.get("suggestion"):
                        high_suggestions.append(issue["suggestion"])
                    elif issue.get("severity") == "中" and issue.get("suggestion"):
                        mid_suggestions.append(issue["suggestion"])

            if high_suggestions:
                lines.append("### 必须修改（高优先级）")
                for i, s in enumerate(high_suggestions[:5], 1):
                    lines.append(f"{i}. {s}")
            if mid_suggestions:
                lines.append("\n### 建议修改（中优先级）")
                for i, s in enumerate(mid_suggestions[:5], 1):
                    lines.append(f"{i}. {s}")
            lines.append("")

        # 总结
        lines.append("## 🎯 总结\n")
        if total_score >= 90:
            lines.append(f"本文档质量优秀（综合评分{round(total_score, 1)}分），整体结构清晰，内容完整。\n")
        elif total_score >= 80:
            lines.append(f"本文档整体质量良好（综合评分{round(total_score, 1)}分），结构清晰，内容完整。\n")
        elif total_score >= 60:
            lines.append(f"本文档基本框架完整（综合评分{round(total_score, 1)}分），但在多个方面需要改进。\n")
        else:
            lines.append(f"本文档存在较多问题（综合评分{round(total_score, 1)}分），需要全面修订。\n")

        if total_high > 0:
            lines.append(f"**优先处理**: 发现{total_high}个高严重程度问题，建议优先修复。")

        min_dim = min(reviewer.checker.scores, key=reviewer.checker.scores.get)
        lines.append(f"**重点改进方向**: {dim_names[min_dim]} 是当前最薄弱的环节（{reviewer.checker.scores[min_dim]}分），建议重点关注。")

        return '\n'.join(lines)

    @staticmethod
    def _generate_json(reviewer, mode: str) -> str:
        total_score = sum(reviewer.checker.scores.values()) / 4
        return json.dumps({
            "document": reviewer.document_path,
            "timestamp": datetime.now().isoformat(),
            "mode": mode,
            "scores": reviewer.checker.scores,
            "total_score": round(total_score, 1),
            "grade": ReportGenerator._get_grade(total_score),
            "issues": reviewer.checker.issues,
            "summary": reviewer.parser.get_summary()
        }, ensure_ascii=False, indent=2)

    @staticmethod
    def _generate_word(reviewer, mode: str) -> str:
        """生成 Word 文档报告"""
        try:
            from docx import Document
            from docx.shared import Pt, RGBColor, Inches
            from docx.enum.text import WD_ALIGN_PARAGRAPH

            doc = Document()

            # 标题
            title = doc.add_heading('文档审核报告', 0)
            title.alignment = WD_ALIGN_PARAGRAPH.CENTER

            # 基本信息
            doc.add_paragraph(f"审核文档: {reviewer.document_path}")
            doc.add_paragraph(f"审核时间: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
            doc.add_paragraph(f"审核模式: {mode}")
            doc.add_paragraph(f"文档规模: {len(reviewer.content)} 字符 / {len(reviewer.parser.sections)} 章节")

            # 评分概览
            doc.add_heading('📊 评分概览', 1)
            total_score = sum(reviewer.checker.scores.values()) / 4
            dim_names = {'format': '格式规范性', 'logic': '内容逻辑性', 'clarity': '表述清晰度', 'feasibility': '方案可行性'}
            table = doc.add_table(rows=6, cols=3)
            table.style = 'Table Grid'
            table.cell(0, 0).text = '维度'
            table.cell(0, 1).text = '评分'
            table.cell(0, 2).text = '等级'
            for i, key in enumerate(['format', 'logic', 'clarity', 'feasibility']):
                table.cell(i+1, 0).text = dim_names[key]
                table.cell(i+1, 1).text = str(reviewer.checker.scores[key])
                table.cell(i+1, 2).text = ReportGenerator._get_grade(reviewer.checker.scores[key])
            table.cell(5, 0).text = '综合'
            table.cell(5, 1).text = str(round(total_score, 1))
            table.cell(5, 2).text = ReportGenerator._get_grade(total_score)

            # 详细问题
            doc.add_heading('🔍 详细问题', 1)
            type_names = {
                "format": "格式问题", "text_errors": "文字错误",
                "logic": "逻辑问题", "clarity": "表述问题", "feasibility": "可行性问题"
            }
            for itype in ["format", "text_errors", "logic", "clarity", "feasibility"]:
                issues = reviewer.checker.issues.get(itype, [])
                if issues:
                    doc.add_heading(type_names[itype], 2)
                    for i, issue in enumerate(issues, 1):
                        sev = issue.get("severity", "低")
                        p = doc.add_paragraph()
                        run = p.add_run(f"[{sev}] {issue['description']}")
                        run.bold = True
                        if sev == "高":
                            run.font.color.rgb = RGBColor(0xA3, 0x2D, 0x2D)
                        elif sev == "中":
                            run.font.color.rgb = RGBColor(0x85, 0x4F, 0x0B)

                        doc.add_paragraph(f"位置: {issue.get('location', '未知')}")
                        if issue.get("context"):
                            doc.add_paragraph(f"上下文: {issue['context']}")
                        doc.add_paragraph(f"建议: {issue.get('suggestion', '')}")
                        doc.add_paragraph("")

            # 总结
            doc.add_heading('🎯 总结', 1)
            if total_score >= 80:
                doc.add_paragraph(f"本文档整体质量良好（综合评分{round(total_score, 1)}分）。")
            elif total_score >= 60:
                doc.add_paragraph(f"本文档基本框架完整（综合评分{round(total_score, 1)}分），需要改进。")
            else:
                doc.add_paragraph(f"本文档存在较多问题（综合评分{round(total_score, 1)}分），需要全面修订。")

            # 保存到临时文件，返回内容
            import io
            buffer = io.BytesIO()
            doc.save(buffer)
            return buffer.getvalue()  # 返回二进制内容

        except ImportError:
            print("提示：未安装 python-docx，无法生成 Word 报告，回退到 Markdown")
            return ReportGenerator._generate_markdown(reviewer, mode)

    @staticmethod
    def _get_grade(score: float) -> str:
        if score >= 90: return "优"
        if score >= 75: return "良"
        if score >= 60: return "中"
        return "差"


# ──────────────────────────────────────────────
# 主审核器
# ──────────────────────────────────────────────

class DocumentReviewer:
    """文档审核器 — 协调 L1 量化层和 L2 语义层"""

    def __init__(self, document_path: str, mode: str = "standard",
                 exclude: List[str] = None, glossary: Dict = None,
                 output_format: str = "markdown"):
        self.document_path = document_path
        self.mode = mode
        self.exclude = exclude or []
        self.glossary = glossary or {}
        self.output_format = output_format
        self.content = ""
        self.file_type = ""
        self.parser: Optional[DocumentParser] = None
        self.checker: Optional[QuantitativeChecker] = None

    def review(self) -> str:
        """执行完整审核流程"""
        print(f"正在审核文档：{self.document_path}")
        print(f"审核模式：{self.mode}")

        # 1. 提取内容
        print("  → 提取文档内容...")
        self.content, self.file_type = ContentExtractor.extract(self.document_path)
        if not self.content.strip():
            return "审核失败：无法提取文档内容"
        print(f"  → 文档长度：{len(self.content)} 字符（格式: {self.file_type}）")

        # 2. 解析结构
        print("  → 解析文档结构...")
        self.parser = DocumentParser(self.content)
        self.parser.parse()
        print(f"  → 章节数量：{len(self.parser.sections)}")
        print(f"  → 数据点：{len(self.parser.data_points)}")
        print(f"  → 时间节点：{len(self.parser.timeline)}")

        # 3. L1 量化检查
        print("  → 执行 L1 量化检查...")
        self.checker = QuantitativeChecker(
            self.parser, self.content, self.exclude, self.glossary
        )

        # 根据模式决定检查范围
        if self.mode == "quick":
            # 快速模式：仅格式+文字错误
            self.checker.exclude.extend(["logic", "feasibility"])

        self.checker.run_all()

        # 4. 生成报告
        print("  → 生成审核报告...")
        report = ReportGenerator.generate(self, self.mode, self.output_format)
        print("\n✅ 审核完成！")
        return report


class BatchReviewer:
    """批量审核器"""

    def __init__(self, file_paths: List[str], mode: str = "standard",
                 exclude: List[str] = None, glossary: Dict = None):
        self.file_paths = file_paths
        self.mode = mode
        self.exclude = exclude or []
        self.glossary = glossary or {}
        self.results: List[Dict] = []

    def review_all(self) -> Dict:
        """审核所有文档，生成汇总报告"""
        print(f"批量审核 {len(self.file_paths)} 个文档...\n")

        for i, filepath in enumerate(self.file_paths, 1):
            print(f"[{i}/{len(self.file_paths)}] {filepath}")
            reviewer = DocumentReviewer(filepath, self.mode, self.exclude, self.glossary)
            try:
                report = reviewer.review()
                total_score = sum(reviewer.checker.scores.values()) / 4 if reviewer.checker else 0

                # 统计问题
                issue_count = 0
                high_count = 0
                for issues in reviewer.checker.issues.values() if reviewer.checker else []:
                    issue_count += len(issues)
                    high_count += sum(1 for i in issues if i.get("severity") == "高")

                self.results.append({
                    "file": filepath,
                    "score": round(total_score, 1),
                    "grade": ReportGenerator._get_grade(total_score),
                    "total_issues": issue_count,
                    "high_issues": high_count,
                    "report": report
                })
            except Exception as e:
                self.results.append({
                    "file": filepath,
                    "error": str(e),
                    "score": 0,
                    "grade": "差",
                    "total_issues": 0,
                    "high_issues": 0
                })
            print()

        return self._generate_summary()

    def _generate_summary(self) -> Dict:
        """生成汇总报告"""
        # 跨文档一致性检查
        cross_doc_issues = []

        # 检查不同文档中是否引用了相同数据但数值不同
        all_data = {}
        for result in self.results:
            if "error" in result:
                continue
            # 简化版：检查文档名中包含相同关键词的数据点
            # 实际跨文档一致性检查需要 AI 语义层完成

        return {
            "total_documents": len(self.file_paths),
            "successful": sum(1 for r in self.results if "error" not in r),
            "failed": sum(1 for r in self.results if "error" in r),
            "results": [
                {
                    "file": r["file"],
                    "score": r["score"],
                    "grade": r["grade"],
                    "total_issues": r["total_issues"],
                    "high_issues": r["high_issues"]
                }
                for r in self.results
            ],
            "ranking": sorted(
                [{"file": r["file"], "score": r["score"]} for r in self.results if "error" not in r],
                key=lambda x: x["score"], reverse=True
            )
        }


# ──────────────────────────────────────────────
# 命令行入口
# ──────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description='文档审核工具 v7.0 — 双层架构全维度审核',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 单文件标准审核
  python review_document.py document.docx

  # 深度审核，输出 Word 报告
  python review_document.py report.pdf --mode deep --output word --output-file report.docx

  # 批量审核
  python review_document.py doc1.docx doc2.pdf --batch --output json

  # 增量审核（仅审核修改部分）
  python review_document.py new.docx --diff old.docx

  # 排除指定检查 + 术语词典
  python review_document.py doc.docx --exclude text_errors --glossary terms.json
        """
    )

    parser.add_argument('document_paths', nargs='+', help='待审核文档路径（支持多个）')
    parser.add_argument('--mode', '-m', choices=['quick', 'standard', 'deep'],
                        default='standard', help='审核模式（默认: standard）')
    parser.add_argument('--output', '-o', choices=['markdown', 'json', 'word'],
                        default='markdown', help='输出格式（默认: markdown）')
    parser.add_argument('--output-file', help='将报告保存到指定文件')
    parser.add_argument('--exclude', help='排除的检查类型（逗号分隔）')
    parser.add_argument('--glossary', help='术语词典文件路径（JSON）')
    parser.add_argument('--diff', help='增量审核：指定旧版本文件路径')
    parser.add_argument('--batch', action='store_true', help='批量审核模式')

    args = parser.parse_args()

    # 解析排除列表
    exclude = args.exclude.split(',') if args.exclude else []

    # 加载术语词典
    glossary = {}
    if args.glossary:
        try:
            with open(args.glossary, 'r', encoding='utf-8') as f:
                glossary = json.load(f)
        except Exception as e:
            print(f"警告：无法加载术语词典 - {e}")

    # 增量审核模式
    if args.diff:
        print("增量审核模式\n")
        old_content, _ = ContentExtractor.extract(args.diff)
        new_content, _ = ContentExtractor.extract(args.document_paths[0])
        diff_result = DiffReviewer.diff(old_content, new_content)

        print(f"新增: {diff_result['added_count']} 行")
        print(f"删除: {diff_result['removed_count']} 行")
        print(f"修改: {diff_result['modified_count']} 行\n")

        if not diff_result["has_changes"]:
            print("未检测到变更。")
            return 0

        # 对变更部分执行审核
        changed_content = '\n'.join(diff_result["added"]) + '\n' + \
                         '\n'.join(m["new"] for m in diff_result["modified"])
        reviewer = DocumentReviewer(args.document_paths[0], args.mode, exclude, glossary, args.output)
        reviewer.content = changed_content
        reviewer.parser = DocumentParser(changed_content)
        reviewer.parser.parse()
        reviewer.checker = QuantitativeChecker(reviewer.parser, changed_content, exclude, glossary)
        reviewer.checker.run_all()

        report = ReportGenerator.generate(reviewer, args.mode + " (增量)", args.output)
        _output_report(report, args.output_file)
        return 0

    # 批量审核模式
    if args.batch or len(args.document_paths) > 1:
        batch = BatchReviewer(args.document_paths, args.mode, exclude, glossary)
        summary = batch.review_all()

        if args.output == 'json':
            report = json.dumps(summary, ensure_ascii=False, indent=2)
        else:
            report = _format_batch_summary(summary)

        _output_report(report, args.output_file)
        return 0

    # 单文件审核
    reviewer = DocumentReviewer(args.document_paths[0], args.mode, exclude, glossary, args.output)
    report = reviewer.review()

    if args.output == 'word':
        # Word 输出需要特殊处理
        if args.output_file:
            if isinstance(report, bytes):
                with open(args.output_file, 'wb') as f:
                    f.write(report)
            else:
                with open(args.output_file, 'w', encoding='utf-8') as f:
                    f.write(report)
            print(f"\n报告已保存到：{args.output_file}")
        else:
            print("提示：Word 格式需要指定 --output-file")
    else:
        _output_report(report, args.output_file)

    return 0


def _output_report(report: str, output_file: str = None):
    """输出报告"""
    if output_file:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f"\n报告已保存到：{output_file}")
    else:
        print("\n" + "=" * 60)
        print(report)


def _format_batch_summary(summary: Dict) -> str:
    """格式化批量审核汇总报告"""
    lines = []
    lines.append("# 批量审核汇总报告\n")
    lines.append(f"**审核时间**: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append(f"**文档总数**: {summary['total_documents']}")
    lines.append(f"**成功**: {summary['successful']} / **失败**: {summary['failed']}\n")

    lines.append("## 📊 审核结果\n")
    lines.append("| # | 文档 | 评分 | 等级 | 问题数 | 高优先级 |")
    lines.append("|---|------|------|------|--------|---------|")
    for i, r in enumerate(summary["results"], 1):
        lines.append(f"| {i} | {Path(r['file']).name} | {r['score']} | {r['grade']} | {r['total_issues']} | {r['high_issues']} |")

    lines.append("\n## 🏆 质量排名\n")
    for i, r in enumerate(summary["ranking"], 1):
        lines.append(f"{i}. {Path(r['file']).name} — {r['score']}分")

    return '\n'.join(lines)


if __name__ == '__main__':
    sys.exit(main())
