# -*- coding: utf-8 -*-
"""
错题本工作流自动化辅助工具
用于辅助精准定位试题、渲染答案页以及生成更新 Markdown 与 Word 错题本。
"""
import os
import re
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

SUBJECT_MAP = {
    '马原': ('试题/1-马原-试题.pdf', '答案/1-马原-答案.pdf', '马克思主义基本原理'),
    '毛中特': ('试题/2-毛中特-试题.pdf', '答案/2-毛中特-答案.pdf', '毛泽东思想和中国特色社会主义理论体系概论'),
    '新思想': ('试题/3-新思想-试题.pdf', '答案/3-新思想-答案.pdf', '习近平新时代中国特色社会主义思想概论'),
    '史纲': ('试题/4-近现代史-试题.pdf', '答案/4-近现代史-答案.pdf', '中国近现代史纲要'),
    '近现代史': ('试题/4-近现代史-试题.pdf', '答案/4-近现代史-答案.pdf', '中国近现代史纲要'),
    '思修': ('试题/5-思修-试题.pdf', '答案/5-思修-答案.pdf', '思想道德与法治'),
}

def resolve_subject(sub_name):
    for k, v in SUBJECT_MAP.items():
        if k in sub_name:
            return v
    # 默认尝试马原
    return SUBJECT_MAP['马原']

def extract_question(subject_pdf, q_num):
    """从试题 PDF 中提取指定题号的题目文本"""
    doc = fitz.open(subject_pdf)
    q_str = str(q_num)
    target_pattern = re.compile(rf'(?:^|\n)\s*{q_str}\s*\.\s*(.+)', re.DOTALL)
    
    # 找到包含该题的页面
    found_page = None
    for p_idx in range(len(doc)):
        text = doc[p_idx].get_text()
        # 寻找诸如 "1." 或 "\n1."
        if re.search(rf'(?:^|\n)\s*{q_str}\s*\.', text):
            found_page = p_idx
            break
            
    if found_page is None:
        return None, "未找到该题号"
        
    # 题目可能跨页，合并当前页及下一页
    full_text = doc[found_page].get_text()
    if found_page + 1 < len(doc):
        full_text += "\n" + doc[found_page + 1].get_text()
        
    # 切割题目：从目标题号开始，直到下一题的题号 (如 "\n2.")
    next_q = str(int(q_num) + 1)
    split_pattern = rf'(?:^|\n)(\s*{q_str}\s*\..+?)(?=(?:\n\s*{next_q}\s*\.|\n[一二三四五]、|\Z))'
    m = re.search(split_pattern, full_text, re.DOTALL)
    if m:
        question_text = m.group(1).strip()
    else:
        question_text = full_text[:800] # 退避截取
        
    return found_page + 1, question_text

def render_answer_pages(answer_pdf, page_range=(1, 5), output_dir="temp_answers"):
    """渲染答案指定页数并保存为图片"""
    os.makedirs(output_dir, exist_ok=True)
    doc = fitz.open(answer_pdf)
    rendered = []
    for p in page_range:
        if 1 <= p <= len(doc):
            pix = doc[p-1].get_pixmap(dpi=150)
            img_path = os.path.join(output_dir, f"page_{p}.png")
            pix.save(img_path)
            rendered.append(img_path)
    return rendered

if __name__ == '__main__':
    print("Helper workflow ready.")
