#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
雨课堂题库工具箱 GUI
整合：
1. 雨课堂自动抓取题库到 Excel
2. Excel 题库排版生成 PDF（背诵版 / 练习版）
3. 两个 PDF 题库差异比对
"""

import hashlib
try:
    _original_md5 = hashlib.md5
    def _patched_md5(*args, **kwargs):
        kwargs.pop('usedforsecurity', None)
        return _original_md5(*args, **kwargs)
    hashlib.md5 = _patched_md5
except Exception:
    pass

import io
import os
import re
import shutil
import sys
import time
import threading
import queue
import traceback
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Callable, Optional, Tuple

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

try:
    import pandas as pd
except Exception:
    pd = None

try:
    import pdfplumber
except Exception:
    pdfplumber = None

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, KeepTogether, Flowable, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.lib.pdfencrypt import StandardEncryption
except Exception:
    A4 = None

try:
    from selenium import webdriver
    from selenium.common.exceptions import TimeoutException, WebDriverException
    from selenium.webdriver.common.by import By
    from selenium.webdriver.chrome.service import Service as ChromeService
    from selenium.webdriver.edge.service import Service as EdgeService
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from webdriver_manager.chrome import ChromeDriverManager
    from webdriver_manager.microsoft import EdgeChromiumDriverManager
except Exception:
    webdriver = None
    TimeoutException = Exception
    WebDriverException = Exception

# OCR 依赖（优先 easyocr，备选 rapidocr，最后 pytesseract）
# easyocr 原生支持日语假名，最适合本项目的日语题库场景
EasyOCR = None
RapidOCR = None
pytesseract = None
PIL_Image = None

try:
    import easyocr as _easyocr
    EasyOCR = _easyocr
except Exception:
    pass

if EasyOCR is None:
    try:
        from rapidocr_onnxruntime import RapidOCR as _RapidOCR
        RapidOCR = _RapidOCR
    except Exception:
        pass

if EasyOCR is None and RapidOCR is None:
    try:
        import pytesseract as _pytesseract
        from PIL import Image as _PIL_Image
        pytesseract = _pytesseract
        PIL_Image = _PIL_Image
    except Exception:
        pass

APP_TITLE = "雨课堂考试题与练习题增强版题库工具箱"
DEFAULT_HEADER = "长江雨课堂日语题库"
DEFAULT_TITLE = "《2026日语》题库"
YUKETANG_URL = "https://www.yuketang.cn/v2/web/index"


def require_modules(modules: List[Tuple[object, str]]):
    missing = [name for module, name in modules if module is None]
    if missing:
        raise RuntimeError("缺少依赖：" + ", ".join(missing) + "\n请在虚拟环境中安装：python -m pip install pandas openpyxl selenium webdriver-manager reportlab pdfplumber")


def safe_xml_text(text: object) -> str:
    s = "" if text is None else str(text)
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    return s.replace("\n", "<br/>")


class HorizontalLine(Flowable):
    def __init__(self, width=440):
        Flowable.__init__(self)
        self.width = width

    def draw(self):
        self.canv.setStrokeColor(colors.lightgrey)
        self.canv.setLineWidth(0.5)
        self.canv.line(0, 0, self.width, 0)


def get_system_font_path() -> Optional[str]:
    candidates = [
        "SimHei.ttf",
        "/System/Library/Fonts/PingFang.ttc",
        "/System/Library/Fonts/Supplemental/Songti.ttc",
        "/System/Library/Fonts/STHeiti Medium.ttc",
        "/System/Library/Fonts/STHeiti Light.ttc",
        "/Library/Fonts/Songti.ttc",
        "/Library/Fonts/Arial Unicode.ttf",
        "C:\\Windows\\Fonts\\simhei.ttf",
        "C:\\Windows\\Fonts\\msyh.ttf",
    ]
    for p in candidates:
        if os.path.exists(p):
            return p
    return None


def make_header_drawer(header_text: str):
    def draw_header(canvas, doc):
        canvas.saveState()
        try:
            canvas.setFont('ChineseFont', 9)
        except Exception:
            canvas.setFont('Helvetica', 9)
        page_width, page_height = A4
        x_pos = page_width - 2 * cm
        y_pos = page_height - 1.0 * cm
        canvas.drawRightString(x_pos, y_pos, header_text)
        canvas.setLineWidth(0.5)
        canvas.setStrokeColor(colors.grey)
        canvas.line(2 * cm, y_pos - 0.2 * cm, page_width - 2 * cm, y_pos - 0.2 * cm)
        canvas.restoreState()
    return draw_header


def create_pdf_file(filename: str, single_choice: list, multi_choice: list, judgment_choice: list,
                    font_name: str, mode: str, header_text: str, title_text: str,
                    encrypt_pdf: bool, log: Callable[[str], None]):
    content_width = A4[0] - 4 * cm
    encrypt_config = None
    if encrypt_pdf:
        encrypt_config = StandardEncryption(
            userPassword="",
            ownerPassword="AndyRONG921",
            canPrint=1,
            canModify=0,
            canCopy=1,
            canAnnotate=0,
        )

    doc = SimpleDocTemplate(
        filename,
        pagesize=A4,
        rightMargin=2 * cm,
        leftMargin=2 * cm,
        topMargin=2 * cm,
        bottomMargin=2 * cm,
        encrypt=encrypt_config,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(name='ExamTitle', parent=styles['Heading1'], fontName=font_name, fontSize=20, alignment=1, spaceAfter=20)
    section_style = ParagraphStyle(name='SectionHeader', parent=styles['Heading2'], fontName=font_name, fontSize=15, spaceBefore=15, spaceAfter=10, textColor=colors.HexColor("#2c3e50"), borderPadding=5)
    question_style = ParagraphStyle(name='QuestionText', parent=styles['Normal'], fontName=font_name, fontSize=11, leading=18, spaceAfter=8)
    option_style = ParagraphStyle(name='OptionText', parent=styles['Normal'], fontName=font_name, fontSize=10.5, leftIndent=15, leading=16, textColor=colors.HexColor("#34495e"))
    answer_style = ParagraphStyle(name='AnswerText', parent=styles['Normal'], fontName=font_name, fontSize=10, textColor=colors.HexColor("#1e8449"), leftIndent=15, spaceBefore=5, spaceAfter=5, backColor=colors.HexColor("#e8f8f5"), borderPadding=3)

    story = [Paragraph(safe_xml_text(title_text), title_style), Spacer(1, 0.5 * cm)]
    single_answers, multi_answers, judgment_answers = [], [], []

    def append_question_section(section_title: str, questions: list, answer_list: list, answer_label: str = "正确答案"):
        story.append(Paragraph(section_title, section_style))
        story.append(HorizontalLine())
        story.append(Spacer(1, 0.3 * cm))
        for i, q in enumerate(questions, 1):
            q_elements = [Paragraph(f"<b>{i}.</b> {safe_xml_text(q['title'])}", question_style)]
            for opt in q.get('options', []):
                q_elements.append(Paragraph(safe_xml_text(opt), option_style))
            if mode == 'inline':
                q_elements.append(Paragraph(f"<b>【{answer_label}】 {safe_xml_text(q['answer'])}</b>", answer_style))
            else:
                answer_list.append(f"{i}.{q['answer']}")
            q_elements += [Spacer(1, 0.35 * cm), HorizontalLine(), Spacer(1, 0.35 * cm)]
            story.append(KeepTogether(q_elements))

    if single_choice:
        append_question_section(f"一、单选题 (共 {len(single_choice)} 题)", single_choice, single_answers)
    if multi_choice:
        if single_choice:
            story.append(PageBreak())
        append_question_section(f"二、多选题 (共 {len(multi_choice)} 题)", multi_choice, multi_answers)
    if judgment_choice:
        if single_choice or multi_choice:
            story.append(PageBreak())
        append_question_section(f"三、判断题 (共 {len(judgment_choice)} 题)", judgment_choice, judgment_answers, "答案")

    if mode == 'end':
        story.append(PageBreak())
        story.append(Paragraph("参考答案", title_style))
        story.append(HorizontalLine())
        story.append(Spacer(1, 0.5 * cm))
        matrix_style = TableStyle([
            ('FONTNAME', (0, 0), (-1, -1), font_name),
            ('FONTSIZE', (0, 0), (-1, -1), 11),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
        ])

        def make_answer_table(data_list, cols_count):
            table_data, row = [], []
            for item in data_list:
                row.append(item)
                if len(row) == cols_count:
                    table_data.append(row)
                    row = []
            if row:
                while len(row) < cols_count:
                    row.append("")
                table_data.append(row)
            table = Table(table_data, colWidths=[content_width / cols_count] * cols_count)
            table.setStyle(matrix_style)
            return table

        if single_answers:
            story.append(Paragraph("<b>一、单选题答案</b>", section_style))
            story.append(make_answer_table(single_answers, 8))
            story.append(Spacer(1, 0.5 * cm))
        if multi_answers:
            story.append(Paragraph("<b>二、多选题答案</b>", section_style))
            story.append(make_answer_table(multi_answers, 5))
            story.append(Spacer(1, 0.5 * cm))
        if judgment_answers:
            story.append(Paragraph("<b>三、判断题答案</b>", section_style))
            story.append(make_answer_table(judgment_answers, 6))
            story.append(Spacer(1, 0.5 * cm))

    log(f"📄 正在写入 PDF：{filename}")
    doc.build(story, onFirstPage=make_header_drawer(header_text), onLaterPages=make_header_drawer(header_text))
    log(f"✅ 已生成：{filename}")


def generate_exam_pdf(excel_path: str, output_dir: str, base_name: str, header_text: str,
                      title_text: str, make_inline: bool, make_practice: bool,
                      encrypt_pdf: bool, log: Callable[[str], None]) -> List[str]:
    require_modules([(pd, 'pandas/openpyxl'), (A4, 'reportlab')])
    if not os.path.exists(excel_path):
        raise FileNotFoundError(f"找不到 Excel 文件：{excel_path}")
    os.makedirs(output_dir, exist_ok=True)

    font_path = get_system_font_path()
    if not font_path:
        raise RuntimeError("未找到中文字体。可安装 SimHei.ttf，或使用 macOS 自带 PingFang/Songti。")
    try:
        pdfmetrics.registerFont(TTFont('ChineseFont', font_path))
    except Exception:
        pdfmetrics.registerFont(TTFont('ChineseFont', font_path, subfontIndex=0))
    log(f"✅ 中文字体：{font_path}")

    df = pd.read_excel(excel_path).fillna("")
    required = {'题目', '答案'}
    missing = required - set(df.columns)
    if missing:
        raise RuntimeError(f"Excel 缺少必要列：{', '.join(missing)}")

    single_choice, multi_choice, judgment = [], [], []
    for _, row in df.iterrows():
        ans = str(row.get('答案', '')).strip()
        clean_ans = ans.replace(" ", "").replace(",", "")
        title = str(row.get('题目', '')).strip()
        if not title:
            continue
        q = {'title': title, 'options': [], 'answer': ans}
        for label in ['A', 'B', 'C', 'D', 'E', 'F', 'G']:
            val = str(row.get(label, '')).strip() if label in row else ''
            if val:
                q['options'].append(f"{label}. {val}")
        if any(k in clean_ans for k in ["正确", "错误", "对", "错"]):
            judgment.append(q)
        elif len(clean_ans) > 1:
            multi_choice.append(q)
        else:
            single_choice.append(q)

    log(f"📚 题目统计：单选 {len(single_choice)} | 多选 {len(multi_choice)} | 判断 {len(judgment)}")
    if not (single_choice or multi_choice or judgment):
        raise RuntimeError("Excel 中没有解析出题目。")

    outputs = []
    clean_base = base_name.strip() or "雨课堂毛概题库"
    if make_inline:
        path = os.path.join(output_dir, f"{clean_base}_背诵版.pdf")
        create_pdf_file(path, single_choice, multi_choice, judgment, 'ChineseFont', 'inline', header_text, title_text, encrypt_pdf, log)
        outputs.append(path)
    if make_practice:
        path = os.path.join(output_dir, f"{clean_base}_练习版.pdf")
        create_pdf_file(path, single_choice, multi_choice, judgment, 'ChineseFont', 'end', header_text, title_text, encrypt_pdf, log)
        outputs.append(path)
    return outputs


def load_existing_data(filepath: str, log: Callable[[str], None]) -> Dict[str, dict]:
    require_modules([(pd, 'pandas/openpyxl')])
    if not os.path.exists(filepath):
        log("✨ 未检测到旧题库，将创建新文件。")
        return {}
    try:
        df = pd.read_excel(filepath).fillna("")
        existing = {}
        for _, row in df.iterrows():
            q_text = str(row.get('题目', '')).strip()
            if q_text:
                existing[q_text] = row.to_dict()
        log(f"✅ 已加载历史题目：{len(existing)} 道")
        return existing
    except Exception as e:
        log(f"⚠️ 读取旧文件失败，将重新开始：{e}")
        return {}


def save_to_excel(data: Dict[str, dict], save_path: str, log: Callable[[str], None]):
    require_modules([(pd, 'pandas/openpyxl')])
    os.makedirs(os.path.dirname(save_path) or '.', exist_ok=True)
    df = pd.DataFrame(data.values())
    cols = ["题目", "答案", "A", "B", "C", "D", "E", "F"]
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols]
    df.to_excel(save_path, index=False)
    log(f"📁 Excel 已更新：{save_path}")


def clean_option_text(text: str) -> str:
    return re.sub(r'^[A-F][\.\s、．]*', '', text or '').strip()


def is_ocr_available() -> bool:
    """检查是否有可用的 OCR 引擎。"""
    return EasyOCR is not None or RapidOCR is not None or pytesseract is not None


def ocr_recognize_text(image_data) -> str:
    """对图片数据执行 OCR 识别，返回识别出的文本。

    Args:
        image_data: PIL Image 对象 或 PNG bytes 数据
    Returns:
        识别出的文本字符串
    """
    if EasyOCR is not None:
        import numpy as np
        if hasattr(image_data, 'convert'):
            img_array = np.array(image_data.convert('RGB'))
        elif isinstance(image_data, (bytes, bytearray)):
            import cv2
            img_array = cv2.imdecode(
                np.frombuffer(image_data, np.uint8), cv2.IMREAD_COLOR
            )
            if img_array is None:
                raise RuntimeError("cv2.imdecode 解码图片失败")
        else:
            img_array = np.array(image_data)
        reader = EasyOCR.Reader(['ja', 'en'], gpu=False)
        results = reader.readtext(img_array, detail=0)
        return "\n".join(results)
    elif RapidOCR is not None:
        import numpy as np
        if hasattr(image_data, 'convert'):
            img_array = np.array(image_data.convert('RGB'))
        elif isinstance(image_data, (bytes, bytearray)):
            import cv2
            img_array = cv2.imdecode(
                np.frombuffer(image_data, np.uint8), cv2.IMREAD_COLOR
            )
            if img_array is None:
                raise RuntimeError("cv2.imdecode 解码图片失败")
        else:
            img_array = np.array(image_data)
        engine = RapidOCR(text_score=0.3)
        result, _ = engine(img_array)
        if result:
            return "\n".join([line[1] for line in result])
        return ""
    elif pytesseract is not None:
        if not hasattr(image_data, 'convert'):
            image_data = PIL_Image.open(io.BytesIO(image_data))
        return pytesseract.image_to_string(image_data, lang='chi_sim+eng+jpn')
    else:
        raise RuntimeError("无可用的 OCR 引擎。请安装 easyocr 或 rapidocr-onnxruntime 或 pytesseract。")


def ocr_screenshot_element(driver, element=None) -> str:
    """对页面或指定元素截图并 OCR 识别，返回文本。

    Args:
        driver: Selenium WebDriver 实例
        element: 可选，指定要截图的 WebElement；为 None 时截全页面
    Returns:
        OCR 识别出的文本
    """
    if element is not None:
        png_data = element.screenshot_as_png
    else:
        png_data = driver.get_screenshot_as_png()
    # ocr_recognize_text 内部会自动处理 bytes → numpy array 的转换
    return ocr_recognize_text(png_data)


def parse_ocr_text_to_question(ocr_text: str, log: Callable[[str], None]) -> Optional[dict]:
    """将 OCR 识别出的原始文本解析为题目结构化数据。

    OCR 文本格式预期类似：
        1.单选题（1分）
        下列哪个是正确的？
        A.选项一
        B.选项二
        C.选项三
        D.选项四
        正确答案：A

    Args:
        ocr_text: OCR 识别的完整文本
        log: 日志回调
    Returns:
        解析后的题目字典，或 None
    """
    if not ocr_text or len(ocr_text.strip()) < 5:
        log("⚠️ OCR 文本过短，无法解析。")
        return None

    text = ocr_text.strip()

    # --- 提取题型 ---
    question_type_str = ""
    type_patterns = [
        r'(单选题|多选题|判断题|填空题|简答题)',
        r'(single.?choice|multiple.?choice|true.?false)',
    ]
    for pat in type_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            question_type_str = m.group(1)
            break

    # --- 提取正确答案 ---
    answer = "未知"
    ans_patterns = [
        r'正确答案[：:]\s*([A-Za-z\s,一-鿿぀-ヿ･-ﾟ]+)',
        r'答案[：:]\s*([A-Za-z\s,一-鿿぀-ヿ･-ﾟ]+)',
        r'参考答案[：:]\s*([A-Za-z\s,一-鿿぀-ヿ･-ﾟ]+)',
    ]
    for pat in ans_patterns:
        m = re.search(pat, text)
        if m:
            answer = m.group(1).replace(" ", "").replace(",", "").strip()
            break

    # --- 提取选项 ---
    # 去掉题型行和答案行，只保留题干+选项
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    cleaned_lines = []
    for line in lines:
        # 跳过题型标注行（如 "1.单选题（1分）"）
        if re.match(r'^\d+[\.\s、．]*(单选题|多选题|判断题|填空题|简答题)', line):
            continue
        # 跳过答案行
        if re.match(r'^(正确答案|答案|参考答案)[：:]', line):
            continue
        cleaned_lines.append(line)

    # 从清理后的文本中拆分题干和选项
    title_text = ""
    options_raw = []
    option_started = False
    title_parts = []

    for line in cleaned_lines:
        # 策略1：行首匹配选项（A. / B. / ...）
        opt_match = re.match(r'^([A-Fa-f])[\.\s、．]\s*(.+)', line)
        if opt_match:
            option_started = True
            options_raw.append(f"{opt_match.group(1).upper()}. {opt_match.group(2).strip()}")
            continue

        # 策略2：一行中包含多个选项（OCR 可能把 A B C D 输出在同一行）
        # 匹配形如 "A.xxx B.xxx C.xxx D.xxx" 或 "A、xxx B、xxx ..."
        multi_opts = re.findall(r'([A-Fa-f])[\.\s、．]\s*([^\sA-Fa-f]+(?:[^\nA-Fa-f]*[^\sA-Fa-f])?)', line)
        if len(multi_opts) >= 2:
            option_started = True
            for label, content in multi_opts:
                options_raw.append(f"{label.upper()}. {content.strip()}")
            continue

        if not option_started:
            title_parts.append(line)

    # 如果仍未提取到选项，尝试从整个文本中用正则兜底
    if not options_raw:
        all_opts = re.findall(r'([A-Fa-f])[\.\s、．]\s*([^A-Fa-f\n]{2,})', text)
        if all_opts:
            for label, content in all_opts:
                options_raw.append(f"{label.upper()}. {content.strip()}")

    title_text = " ".join(title_parts).strip()
    # 去掉题号前缀
    title_text = re.sub(r'^[\d]+[\.\s、．）\)]\s*', '', title_text).strip()
    # 去掉残留的题型文字
    title_text = re.sub(r'^(单选题|多选题|判断题|填空题|简答题)\s*[\(（]?\s*\d*\s*分?\s*[\)）]?\s*', '', title_text).strip()

    # --- 组装结果 ---
    labels = ['A', 'B', 'C', 'D', 'E', 'F']
    cleaned_opts = [clean_option_text(o) for o in options_raw]

    is_judgment = (len(cleaned_opts) == 2 and
                   any(k in "".join(cleaned_opts) for k in ["正确", "错误", "对", "错"]))

    item = {"题目": title_text}
    if is_judgment:
        if answer != "未知":
            item["答案"] = "正确" if any(k in answer for k in ["A", "正确", "对"]) else "错误"
        else:
            item["答案"] = "未知"
        for label in labels:
            item[label] = ""
    else:
        item["答案"] = answer
        for i, label in enumerate(labels):
            item[label] = cleaned_opts[i] if i < len(cleaned_opts) else ""

    # 推断题型
    if not question_type_str:
        if is_judgment:
            question_type_str = "判断题"
        elif len(answer) > 1 and answer not in ("未知",):
            question_type_str = "多选题"
        else:
            question_type_str = "单选题"
    item["题型"] = question_type_str

    if not title_text or len(title_text) < 2:
        log("⚠️ OCR 解析出的题干过短，跳过。")
        return None

    log(f"📝 [OCR] 提取成功：[{question_type_str}] {title_text[:50]}... → 答案: {answer}")
    return item


def scrape_current_practice_question_ocr(driver, log: Callable[[str], None]) -> Optional[dict]:
    """使用 OCR 截图方案从练习题页面提取当前显示的题目。

    与 scrape_current_practice_question 的区别：不依赖 DOM 文本，
    而是对题目区域截图后用 OCR 识别，适用于题目以图片/Canvas 渲染的场景。
    """
    try:
        # 优先尝试定位题目容器元素进行截图（比全页面截图更精准）
        container_selectors = [
            ".question-body", ".problem-body", ".question-content",
            ".question-wrapper", ".problem-wrapper", ".exam-question",
            ".practice-question", ".item-body", ".question_area",
            ".questionArea", ".ques-container",
        ]
        container = None
        for sel in container_selectors:
            try:
                elems = driver.find_elements(By.CSS_SELECTOR, sel)
                if elems and elems[0].is_displayed():
                    container = elems[0]
                    break
            except Exception:
                continue

        if container is not None:
            log("📷 [OCR] 对题目容器元素截图识别...")
            ocr_text = ocr_screenshot_element(driver, container)
        else:
            log("📷 [OCR] 未找到题目容器，对整个页面截图识别...")
            ocr_text = ocr_screenshot_element(driver)

        if not ocr_text or len(ocr_text.strip()) < 5:
            log("⚠️ [OCR] 未识别到有效文本。")
            return None

        log(f"📄 [OCR] 识别文本预览：{ocr_text[:120]}...")
        return parse_ocr_text_to_question(ocr_text, log)

    except Exception as e:
        log(f"❌ [OCR] 截图识别失败：{e}")
        return None


def xpath_class_contains(class_name: str) -> str:
    return f"contains(concat(' ', normalize-space(@class), ' '), ' {class_name} ')"


def scrape_current_practice_question(driver, log: Callable[[str], None]) -> Optional[dict]:
    """从练习题页面提取当前显示的单道题目。

    练习题页面结构（单题模式）：
    - 左侧导航显示进度（如 11/60题）
    - 题目主体：题号+题型、题干文本、选项列表
    - 答案解析区：本题得分、正确答案
    - 底部导航：上一题 / 已提交 / 下一题
    """
    try:
        # 等待题目内容加载（多种策略 fallback）
        question_type_str = ""

        # --- 策略1：尝试常见练习题页面 class ---
        container_selectors = [
            ".question-body", ".problem-body", ".question-content",
            ".question-wrapper", ".problem-wrapper", ".exam-question",
            ".practice-question", ".item-body", ".question_area",
            ".questionArea", ".ques-container",
        ]
        container = None
        for sel in container_selectors:
            try:
                elems = driver.find_elements(By.CSS_SELECTOR, sel)
                if elems and elems[0].is_displayed():
                    container = elems[0]
                    break
            except Exception:
                continue

        if container is None:
            # fallback：直接用 body
            container = driver.find_element(By.TAG_NAME, "body")

        full_text = container.text.strip()
        if not full_text or len(full_text) < 5:
            log("⚠️ 当前页面未检测到题目内容，可能页面尚未加载。")
            return None

        # --- 提取题型 ---
        # 常见格式："1.单选题 (1分)" 或 "单选题" 或 "【单选题】"
        type_patterns = [
            r'(单选题|多选题|判断题|填空题|简答题)',
            r'(single.?choice|multiple.?choice|true.?false)',
        ]
        for pat in type_patterns:
            m = re.search(pat, full_text, re.IGNORECASE)
            if m:
                question_type_str = m.group(1)
                break

        # --- 提取题干文本 ---
        # 尝试从专门的题干元素提取
        title_selectors = [
            ".question-title", ".problem-title", ".question-text",
            ".problem-text", ".stem", ".question_stem",
            ".item-title", ".ques-title", "h4", "h3",
        ]
        title_text = ""
        for sel in title_selectors:
            try:
                elems = container.find_elements(By.CSS_SELECTOR, sel)
                for el in elems:
                    t = el.text.strip()
                    if t and len(t) > 2:
                        title_text = t
                        break
                if title_text:
                    break
            except Exception:
                continue

        if not title_text:
            # fallback：取 full_text 的前半部分（到第一个选项 A. 之前）
            option_match = re.search(r'\n\s*[A-Fa-f][\.\s、．]', full_text)
            if option_match:
                title_text = full_text[:option_match.start()].strip()
            else:
                # 取第一行非空文本
                lines = [l.strip() for l in full_text.split('\n') if l.strip()]
                title_text = lines[0] if lines else full_text[:200]

        # 去掉题号前缀，如 "1." "（1）" "1、"
        title_text = re.sub(r'^[\d]+[\.\s、．）\)]\s*', '', title_text).strip()
        # 去掉题型前缀
        title_text = re.sub(r'^(单选题|多选题|判断题|填空题|简答题)\s*[\(（]?\s*\d*\s*分?\s*[\)）]?\s*', '', title_text).strip()

        # --- 提取选项 ---
        option_selectors = [
            ".option-item", ".optionItem", ".option-wrapper",
            ".el-radio", ".el-checkbox",
            ".radioText", ".checkboxText",
            ".el-radio__label", ".el-checkbox__label",
            ".answer-option", ".choice-item",
        ]
        option_elements = []
        for sel in option_selectors:
            try:
                elems = container.find_elements(By.CSS_SELECTOR, sel)
                if elems and len(elems) >= 2:
                    option_elements = elems
                    break
            except Exception:
                continue

        options_raw = []
        if option_elements:
            for el in option_elements:
                opt_text = el.text.strip()
                if opt_text:
                    options_raw.append(opt_text)

        # 如果没有找到选项元素，从 full_text 中正则提取
        if not options_raw:
            opt_pattern = re.findall(r'([A-Fa-f])[\.\s、．]\s*(.+?)(?=\n[A-Fa-f][\.\s、．]|\n*$)', full_text, re.DOTALL)
            if opt_pattern:
                options_raw = [f"{label}. {text.strip()}" for label, text in opt_pattern]

        # --- 提取正确答案 ---
        answer = "未知"
        # 策略1：在页面全文中正则匹配
        ans_patterns = [
            r'正确答案[：:]\s*([A-Za-z\s,一-鿿぀-ヿ･-ﾟ]+)',
            r'答案[：:]\s*([A-Za-z\s,一-鿿぀-ヿ･-ﾟ]+)',
            r'参考答案[：:]\s*([A-Za-z\s,一-鿿぀-ヿ･-ﾟ]+)',
        ]
        for pat in ans_patterns:
            m = re.search(pat, full_text)
            if m:
                answer = m.group(1).replace(" ", "").replace(",", "").strip()
                break

        # 策略2：尝试从答案解析区的专门元素提取
        if answer == "未知":
            ans_selectors = [
                ".answer-detail", ".analysis", ".answer-analysis",
                ".correct-answer", ".answer-text", ".answer_key",
                ".right-answer", ".answer-result",
            ]
            for sel in ans_selectors:
                try:
                    elems = container.find_elements(By.CSS_SELECTOR, sel)
                    for el in elems:
                        el_text = el.text.strip()
                        for pat in ans_patterns:
                            m = re.search(pat, el_text)
                            if m:
                                answer = m.group(1).replace(" ", "").replace(",", "").strip()
                                break
                        if answer != "未知":
                            break
                    if answer != "未知":
                        break
                except Exception:
                    continue

        # 策略3：通过绿色对勾（√）标记判断已选/正确选项
        if answer == "未知":
            try:
                # 查找带有成功/选中状态的选项
                checked_selectors = [
                    ".is-checked", ".is-success", ".selected",
                    ".correct", ".right", ".active",
                ]
                for sel in checked_selectors:
                    checked = container.find_elements(By.CSS_SELECTOR, sel)
                    if checked:
                        labels = []
                        for c in checked:
                            ct = c.text.strip()
                            m = re.match(r'^([A-Fa-f])', ct)
                            if m:
                                labels.append(m.group(1).upper())
                        if labels:
                            answer = "".join(sorted(labels))
                            break
            except Exception:
                pass

        # --- 组装结果 ---
        labels = ['A', 'B', 'C', 'D', 'E', 'F']
        # 清理选项文本
        cleaned_opts = [clean_option_text(o) for o in options_raw]

        is_judgment = (len(cleaned_opts) == 2 and
                       any(k in "".join(cleaned_opts) for k in ["正确", "错误", "对", "错"]))

        item = {"题目": title_text}
        if is_judgment:
            if answer != "未知":
                item["答案"] = "正确" if any(k in answer for k in ["A", "正确", "对"]) else "错误"
            else:
                item["答案"] = "未知"
            for label in labels:
                item[label] = ""
        else:
            item["答案"] = answer
            for i, label in enumerate(labels):
                item[label] = cleaned_opts[i] if i < len(cleaned_opts) else ""

        # 推断题型（如果前面没提取到）
        if not question_type_str:
            if is_judgment:
                question_type_str = "判断题"
            elif len(answer) > 1 and answer not in ("未知",):
                question_type_str = "多选题"
            else:
                question_type_str = "单选题"

        item["题型"] = question_type_str

        if not title_text or len(title_text) < 2:
            log("⚠️ 提取到的题干过短，可能定位有误，跳过本题。")
            return None

        log(f"📝 提取成功：[{question_type_str}] {title_text[:50]}... → 答案: {answer}")
        return item

    except Exception as e:
        log(f"❌ 提取当前题目失败：{e}")
        return None


def click_next_question(driver, log: Callable[[str], None]) -> bool:
    """点击练习题页面底部的"下一题"按钮。

    返回 True 表示成功点击并等待页面更新；
    返回 False 表示按钮不可用（已到最后一题）或未找到。
    """
    try:
        # 多策略查找"下一题"按钮
        next_btn_xpaths = [
            # 直接匹配包含"下一题"文本的按钮
            "//button[contains(normalize-space(.), '下一题')]",
            "//a[contains(normalize-space(.), '下一题')]",
            "//*[contains(normalize-space(.), '下一题')]/ancestor-or-self::*[self::button or self::a or @role='button'][1]",
            # Element UI 按钮
            f"//button[{xpath_class_contains('el-button')} and contains(normalize-space(.), '下一题')]",
            # 下方导航栏的按钮（通常是最后一个"下一题"）
            f"//*[{xpath_class_contains('el-button')}]//span[contains(normalize-space(.), '下一题')]/ancestor::button[1]",
        ]

        next_btn = None
        for xpath in next_btn_xpaths:
            try:
                elems = driver.find_elements(By.XPATH, xpath)
                for el in elems:
                    if el.is_displayed():
                        next_btn = el
                        break
                if next_btn:
                    break
            except Exception:
                continue

        if next_btn is None:
            log('⚠️ 未找到「下一题」按钮，可能已到最后一题或页面结构不匹配。')
            return False

        # 检查按钮是否 disabled
        btn_class = next_btn.get_attribute("class") or ""
        btn_disabled = next_btn.get_attribute("disabled")
        if "is-disabled" in btn_class or "disabled" in btn_class or btn_disabled:
            log('🏁 「下一题」按钮已禁用，已到达最后一题。')
            return False

        if not next_btn.is_enabled():
            log('🏁 「下一题」按钮不可点击，已到达最后一题。')
            return False

        # 记住当前题目文本，用于判断页面是否已切换
        try:
            old_text = driver.find_element(By.TAG_NAME, "body").text[:200]
        except Exception:
            old_text = ""

        # 点击「下一题」
        scroll_and_click(driver, next_btn)
        log('➡️ 已点击「下一题」。')

        # 等待页面内容变化（最多等待 10 秒）
        try:
            WebDriverWait(driver, 10).until(lambda d: d.find_element(By.TAG_NAME, "body").text[:200] != old_text)
        except TimeoutException:
            # 页面可能没变化（比如相邻题目类似），不视为错误
            log("ℹ️ 等待页面刷新超时，继续尝试提取。")
            time.sleep(1)

        time.sleep(0.5)
        return True

    except Exception as e:
        log(f'❌ 点击「下一题」失败：{e}')
        return False


def wait_for_clickable_xpath(driver, xpaths: List[str], timeout: int = 20, prefer_last: bool = False):
    def find_clickable(_driver):
        for xpath in xpaths:
            elements = _driver.find_elements(By.XPATH, xpath)
            if prefer_last:
                elements = list(reversed(elements))
            for element in elements:
                if element.is_displayed() and element.is_enabled():
                    return element
        return False

    return WebDriverWait(driver, timeout).until(find_clickable)


def scroll_and_click(driver, element):
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", element)
    time.sleep(0.5)
    try:
        element.click()
    except WebDriverException:
        driver.execute_script("arguments[0].click();", element)


def auto_click_flow(driver, log: Callable[[str], None], auto_pause: bool) -> bool:
    try:
        original_window = driver.current_window_handle
        log("🔍 寻找“再次答题”按钮...")
        retry_button = wait_for_clickable_xpath(driver, [
            "//*[contains(normalize-space(.), '再次答题') and "
            f"{xpath_class_contains('linkkk')}]",
            "//*[contains(normalize-space(.), '再次答题')]/ancestor-or-self::*["
            "self::button or self::a or @role='button' or "
            f"{xpath_class_contains('btn_qq')} or {xpath_class_contains('linkkk')}][1]",
        ])
        scroll_and_click(driver, retry_button)
        log("✅ 已点击再次答题")

        try:
            WebDriverWait(driver, 8).until(lambda d: len(d.window_handles) > 1)
            for handle in driver.window_handles:
                if handle != original_window:
                    driver.switch_to.window(handle)
                    break
        except TimeoutException:
            log("ℹ️ 未检测到新窗口，继续在当前页面操作")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.TAG_NAME, "body")))
        time.sleep(1)

        start_button = wait_for_clickable_xpath(driver, [
            "//button[contains(normalize-space(.), '开始') and "
            f"{xpath_class_contains('el-button--primary')}]",
            "//*[contains(normalize-space(.), '开始')]/ancestor-or-self::*["
            "self::button or self::a or @role='button'][1]",
        ])
        scroll_and_click(driver, start_button)
        log("✅ 已点击开始")
        time.sleep(1)

        submit1_button = wait_for_clickable_xpath(driver, [
            "//button[contains(normalize-space(.), '交卷') and "
            f"{xpath_class_contains('tosubmit')}]",
            "//button[contains(normalize-space(.), '交卷') and "
            f"{xpath_class_contains('el-button--primary')}]",
        ])
        scroll_and_click(driver, submit1_button)
        log("✅ 第一次交卷")
        time.sleep(1)

        submit2_button = wait_for_clickable_xpath(driver, [
            "//*[(" + xpath_class_contains('el-message-box') + " or " + xpath_class_contains('el-dialog') + ") "
            "and not(contains(@style, 'display: none'))]//button[.//span[normalize-space() = '交卷']]",
            "//button[.//span[normalize-space() = '交卷'] and "
            f"{xpath_class_contains('el-button--default')} and {xpath_class_contains('el-button--medium')}]",
            "//button[.//span[normalize-space() = '交卷'] and "
            f"({xpath_class_contains('el-button--medium')} or {xpath_class_contains('el-button--primary')})]",
            "//*[contains(normalize-space(.), '交卷')]/ancestor-or-self::*["
            "self::button or self::a or @role='button'][1]",
        ], prefer_last=True)
        scroll_and_click(driver, submit2_button)
        log("✅ 确认交卷")
        time.sleep(4 if not auto_pause else 2)

        view_button = wait_for_clickable_xpath(driver, [
            "//*[contains(normalize-space(.), '查看试卷') and "
            f"{xpath_class_contains('btn_qq')}]",
            "//*[contains(normalize-space(.), '查看试卷')]/ancestor-or-self::*["
            "self::button or self::a or @role='button' or "
            f"{xpath_class_contains('btn_qq')}][1]",
            "//*[" + xpath_class_contains('btn_qq') + " and " + xpath_class_contains('btn--nopass') + "]",
        ])
        scroll_and_click(driver, view_button)
        log("✅ 已点击查看试卷")
        time.sleep(2)
        return True
    except Exception as e:
        log(f"❌ 自动流程失败：{e}")
        return False


def find_driver_path(env_names: List[str], executable_name: str, candidates: List[str]) -> Optional[str]:
    for env_name in env_names:
        value = os.environ.get(env_name)
        if value and os.path.isfile(value) and os.access(value, os.X_OK):
            return value

    path_value = shutil.which(executable_name)
    if path_value:
        return path_value

    for candidate in candidates:
        if os.path.isfile(candidate) and os.access(candidate, os.X_OK):
            return candidate
    return None


def create_webdriver(browser_name: str, attach_existing: bool = False, debugger_address: str = "127.0.0.1:9222"):
    browser = (browser_name or "chrome").lower()
    if browser == "edge":
        options = webdriver.EdgeOptions()
        if attach_existing:
            options.add_experimental_option("debuggerAddress", debugger_address)
        else:
            options.add_experimental_option('excludeSwitches', ['enable-automation'])
            options.add_experimental_option('useAutomationExtension', False)
        driver_path = find_driver_path(
            ["MSEDGEDRIVER_PATH", "EDGEDRIVER_PATH"],
            "msedgedriver",
            [
                str(Path(sys.executable).resolve().parent / "msedgedriver"),
                str(Path(sys.executable).resolve().parent / "msedgedriver.exe"),
                str(Path(__file__).resolve().parent / "yuketang_env" / "bin" / "msedgedriver"),
                str(Path(__file__).resolve().parent / "yuketang_env" / "Scripts" / "msedgedriver.exe"),
                "/usr/local/bin/msedgedriver",
                "/opt/homebrew/bin/msedgedriver",
                str(Path.home() / "msedgedriver"),
                str(Path.home() / "msedgedriver.exe"),
                str(Path.home() / "Downloads" / "msedgedriver"),
                str(Path.home() / "Downloads" / "msedgedriver.exe"),
            ],
        )
        try:
            service = EdgeService(driver_path or EdgeChromiumDriverManager().install())
        except Exception as e:
            raise RuntimeError(
                "无法启动 Edge：未找到本地 msedgedriver，且当前网络无法下载 EdgeDriver。\n"
                "如果选择了“连接已打开浏览器”，请确认 Edge 是用 --remote-debugging-port=9222 启动的。\n"
                "也可以设置环境变量 MSEDGEDRIVER_PATH 指向 msedgedriver 文件。\n"
                f"原始错误：{e}"
            ) from e
        return webdriver.Edge(service=service, options=options)

    options = webdriver.ChromeOptions()
    if attach_existing:
        options.add_experimental_option("debuggerAddress", debugger_address)
    else:
        options.add_experimental_option('excludeSwitches', ['enable-automation'])
        options.add_experimental_option('useAutomationExtension', False)
    driver_path = find_driver_path(
        ["CHROMEDRIVER_PATH"],
        "chromedriver",
        [
            str(Path(sys.executable).resolve().parent / "chromedriver"),
            str(Path(sys.executable).resolve().parent / "chromedriver.exe"),
            str(Path(__file__).resolve().parent / "yuketang_env" / "bin" / "chromedriver"),
            str(Path(__file__).resolve().parent / "yuketang_env" / "Scripts" / "chromedriver.exe"),
            "/usr/local/bin/chromedriver",
            "/opt/homebrew/bin/chromedriver",
            str(Path.home() / "chromedriver"),
            str(Path.home() / "chromedriver.exe"),
            str(Path.home() / "Downloads" / "chromedriver"),
            str(Path.home() / "Downloads" / "chromedriver.exe"),
        ],
    )
    try:
        service = ChromeService(driver_path or ChromeDriverManager().install())
    except Exception as e:
        raise RuntimeError(
            "无法启动 Chrome：未找到本地 chromedriver，且当前网络无法下载 ChromeDriver。\n"
            "如果选择了“连接已打开浏览器”，请确认 Chrome 是用 --remote-debugging-port=9222 启动的。\n"
            "也可以设置环境变量 CHROMEDRIVER_PATH 指向 chromedriver 文件。\n"
            f"原始错误：{e}"
        ) from e
    return webdriver.Chrome(service=service, options=options)


def scrape_current_result_page(driver, question_db: Dict[str, dict], save_path: str,
                               log: Callable[[str], None]) -> bool:
    if len(driver.window_handles) > 1:
        driver.switch_to.window(driver.window_handles[-1])

    blocks = driver.find_elements(By.CLASS_NAME, "result_item")
    if not blocks:
        log("⚠️ 没找到题目，可能当前还不是“查看试卷/结果页”，或页面未加载完成。")
        time.sleep(3)
        return False

    new_count = 0
    for block in blocks:
        try:
            q_text = block.find_element(By.CSS_SELECTOR, ".item-body h4").text.strip()
            if not q_text or q_text in question_db:
                continue
            opt_eles = block.find_elements(By.CSS_SELECTOR, ".radioText, .checkboxText")
            if not opt_eles:
                opt_eles = block.find_elements(By.CSS_SELECTOR, ".el-radio__label, .el-checkbox__label")
            opts = [clean_option_text(o.text.strip()) for o in opt_eles if o.text.strip()]
            full_text = block.text
            ans_match = re.search(r"正确答案[：:]\s*([A-Za-z\s,\u4e00-\u9fff\u3040-\u30ff\uff65-\uff9f]+)", full_text)
            raw_ans = ans_match.group(1).replace(" ", "").replace(",", "").strip() if ans_match else "未知"
            item = {"题目": q_text}
            labels = ['A', 'B', 'C', 'D', 'E', 'F']
            is_judgment = len(opts) == 2 and any(k in "".join(opts) for k in ["正确", "错误", "对", "错"])
            if is_judgment:
                item["答案"] = "正确" if any(k in raw_ans for k in ["A", "正确", "对"]) else "错误"
                for label in labels:
                    item[label] = ""
            else:
                item["答案"] = raw_ans
                for i, label in enumerate(labels):
                    item[label] = opts[i] if i < len(opts) else ""
            question_db[q_text] = item
            new_count += 1
        except Exception:
            continue

    log(f"✅ 本页新增：{new_count} 题 | 题库总计：{len(question_db)} 题")
    if new_count > 0:
        save_to_excel(question_db, save_path, log)
    else:
        log("💤 本页题目已全部收录。")
    return True


def run_practice_spider(save_path: str, browser_name: str, browser_start_mode: str,
                        log: Callable[[str], None], stop_event: threading.Event,
                        driver_state: Optional[dict] = None, use_ocr: bool = False):
    # 练习题模式：逐题提取，自动点击「下一题」，直到最后一题。
    require_modules([(pd, 'pandas/openpyxl'), (webdriver, 'selenium/webdriver-manager')])
    if use_ocr and not is_ocr_available():
        log("⚠️ OCR 不可用（未安装 easyocr/rapidocr/pytesseract），回退到 DOM 模式。")
        use_ocr = False
    attach_existing = browser_start_mode == 'attach'
    driver = create_webdriver(browser_name, attach_existing=attach_existing)
    if driver_state is not None:
        driver_state["driver"] = driver
    question_db = load_existing_data(save_path, log)

    try:
        if attach_existing:
            log("🔗 已连接到当前打开的浏览器。")
        else:
            driver.get(YUKETANG_URL)
            log("🌐 已直接打开雨课堂首页。")

        if attach_existing:
            messagebox.showinfo("开始抓取",
                "请确认当前浏览器页面已停在练习题页面（能看到题目和「下一题」按钮）。\n"
                "准备好后点击确定，程序开始逐题抓取。")
        else:
            messagebox.showinfo("开始抓取",
                "请在浏览器中完成登录，并进入练习题页面（能看到题目和「下一题」按钮）。\n"
                "准备好后点击确定，程序开始逐题抓取。")

        question_num = 0
        consecutive_failures = 0
        max_consecutive_failures = 3  # 连续失败次数上限

        while not stop_event.is_set():
            question_num += 1
            log(f"\n--- 第 {question_num} 题 ---")

            # 提取当前题目（根据模式选择 DOM 或 OCR 方案）
            if use_ocr:
                item = scrape_current_practice_question_ocr(driver, log)
            else:
                item = scrape_current_practice_question(driver, log)
            if item is None:
                consecutive_failures += 1
                if consecutive_failures >= max_consecutive_failures:
                    log(f"⚠️ 连续 {consecutive_failures} 次提取失败，可能页面已异常，停止抓取。")
                    break
                log(f"⚠️ 提取失败（连续失败 {consecutive_failures}/{max_consecutive_failures}），尝试点击下一题继续...")
            else:
                consecutive_failures = 0
                q_key = item.get("题目", "")
                if q_key and q_key not in question_db:
                    # 保存时去掉内部的「题型」字段（不写入 Excel）
                    save_item = {k: v for k, v in item.items() if k != "题型"}
                    question_db[q_key] = save_item
                    log(f"📥 新增入库，题库总计：{len(question_db)} 题")
                elif q_key in question_db:
                    log(f"⏭ 题目已存在，跳过。题库总计：{len(question_db)} 题")
                else:
                    log("⚠️ 题干为空，跳过。")

                # 每 10 题保存一次
                if question_num % 10 == 0:
                    save_to_excel(question_db, save_path, log)

            # 点击「下一题」
            if not click_next_question(driver, log):
                log("🏁 已到达最后一题或无法继续，抓取结束。")
                break

            # 等待一小段时间让页面稳定
            time.sleep(0.8)

        # 最终保存
        save_to_excel(question_db, save_path, log)
        log(f"🎉 练习题抓取结束，共处理 {question_num} 题，题库总量：{len(question_db)} 题")

    finally:
        try:
            driver.quit()
        except Exception:
            pass
        if driver_state is not None and driver_state.get("driver") is driver:
            driver_state["driver"] = None


def run_auto_spider(save_path: str, max_cycles: int, browser_name: str, browser_start_mode: str,
                    crawl_mode: str, auto_pause: bool,
                    log: Callable[[str], None], stop_event: threading.Event,
                    driver_state: Optional[dict] = None, use_ocr: bool = False):
    require_modules([(pd, 'pandas/openpyxl'), (webdriver, 'selenium/webdriver-manager')])

    # 练习题模式：调用独立的 run_practice_spider
    if crawl_mode == 'practice':
        run_practice_spider(save_path, browser_name, browser_start_mode, log, stop_event, driver_state, use_ocr=use_ocr)
        return

    attach_existing = browser_start_mode == 'attach'
    driver = create_webdriver(browser_name, attach_existing=attach_existing)
    if driver_state is not None:
        driver_state["driver"] = driver
    question_db = load_existing_data(save_path, log)

    try:
        if attach_existing:
            log("🔗 已连接到当前打开的浏览器，不再打开雨课堂首页。")
        else:
            driver.get(YUKETANG_URL)
            log("🌐 已直接打开雨课堂首页。")

        if crawl_mode == 'current':
            messagebox.showinfo("开始抓取", "请在浏览器中完成登录，并进入已经能看到题目和正确答案的「查看试卷/结果页」。\n准备好后点击确定，程序会直接抓取当前页。")
            scrape_current_result_page(driver, question_db, save_path, log)
            save_to_excel(question_db, save_path, log)
            log(f"🎉 抓取结束，最终题库总量：{len(question_db)} 题")
            return

        if attach_existing:
            messagebox.showinfo("开始抓取", "请确认当前浏览器页面已经停在可点击「再次答题」的界面。\n准备好后点击确定，程序开始循环抓取。")
        else:
            messagebox.showinfo("开始抓取", "请在浏览器中完成登录，并进入可点击「再次答题」的页面。\n准备好后点击确定，程序开始循环抓取。")

        for cycle in range(1, max_cycles + 1):
            if stop_event.is_set():
                log("⏹ 已收到停止请求。")
                break
            log(f"\n--- 第 {cycle}/{max_cycles} 轮抓取 ---")
            if not auto_click_flow(driver, log, auto_pause):
                if stop_event.is_set():
                    log("⏹ 抓取已强制停止。")
                    break
                if auto_pause:
                    log("⚠️ 自动点击流程失败，已停止自动循环。请手动处理后重新开始。")
                    break
                else:
                    log("⚠️ 自动流程失败，本轮跳过。")
                    continue

            try:
                if not scrape_current_result_page(driver, question_db, save_path, log):
                    continue
                if len(driver.window_handles) > 1:
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                time.sleep(1)
            except Exception as e:
                if stop_event.is_set():
                    log("⏹ 抓取已强制停止。")
                    break
                log(f"❌ 抓取错误：{e}")
                if auto_pause:
                    log("⚠️ 抓取过程中出错，已停止自动循环。请手动恢复页面后重新开始。")
                    break
        if not stop_event.is_set():
            save_to_excel(question_db, save_path, log)
            log(f"🎉 抓取结束，最终题库总量：{len(question_db)} 题")
    finally:
        try:
            driver.quit()
        except Exception:
            pass
        if driver_state is not None and driver_state.get("driver") is driver:
            driver_state["driver"] = None


def clean_text_for_compare(text: str) -> str:
    text = re.sub(r'^\d+\.\s*', '', text or '')
    return re.sub(r'[^\u4e00-\u9fff\u3040-\u30ff\uff65-\uff9fa-zA-Z0-9]', '', text)


def extract_questions_from_pdf(pdf_path: str, log: Callable[[str], None]) -> Dict[str, str]:
    require_modules([(pdfplumber, 'pdfplumber')])
    questions = {}
    log(f"正在读取：{pdf_path}")
    full_lines = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue
            for line in text.split('\n'):
                if any(k in line for k in ["适用学期", "整理人", "PAGE"]):
                    continue
                if any(k in line for k in ["参考答案", "单选题答案", "多选题答案", "判断题答案"]):
                    break
                full_lines.append(line)

    current = []
    pattern = re.compile(r'^(\d+)\.\s*(.*)')

    def save_current():
        if not current:
            return
        full = "\n".join(current)
        opt = re.search(r'[A-G]\.', full)
        stem = full[:opt.start()] if opt else full
        fp = clean_text_for_compare(stem)
        if len(fp) > 5:
            questions[fp] = full

    for line in full_lines:
        line = line.strip()
        if not line:
            continue
        m = pattern.match(line)
        if m and len(m.group(2)) > 5:
            save_current()
            current = [line]
        elif current:
            current.append(line)
    save_current()
    log(f"解析完成：{len(questions)} 道题")
    return questions


def compare_pdfs(pdf1: str, pdf2: str, output_txt: str, log: Callable[[str], None]) -> str:
    if not os.path.exists(pdf1) or not os.path.exists(pdf2):
        raise FileNotFoundError("请确认两个 PDF 文件都存在。")
    q1 = extract_questions_from_pdf(pdf1, log)
    q2 = extract_questions_from_pdf(pdf2, log)
    unique1 = [text for fp, text in q1.items() if fp not in q2]
    unique2 = [text for fp, text in q2.items() if fp not in q1]
    os.makedirs(os.path.dirname(output_txt) or '.', exist_ok=True)
    with open(output_txt, 'w', encoding='utf-8') as f:
        f.write("=== 对比报告 ===\n")
        f.write(f"文件1: {pdf1}\n文件2: {pdf2}\n\n")
        f.write(f"【仅在 文件1 中出现的题目】 (共 {len(unique1)} 题):\n")
        f.write("=" * 50 + "\n")
        for i, q in enumerate(unique1, 1):
            f.write(f"[{i}] {q}\n" + "-" * 30 + "\n")
        f.write("\n\n")
        f.write(f"【仅在 文件2 中出现的题目】 (共 {len(unique2)} 题):\n")
        f.write("=" * 50 + "\n")
        for i, q in enumerate(unique2, 1):
            f.write(f"[{i}] {q}\n" + "-" * 30 + "\n")
    log(f"✅ 对比完成：文件1独有 {len(unique1)} 题，文件2独有 {len(unique2)} 题")
    log(f"📄 结果已保存：{output_txt}")
    return output_txt


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("980x720")
        self.minsize(900, 640)
        self.log_queue = queue.Queue()
        self.stop_event = threading.Event()
        self.worker = None
        self.driver_state = {"driver": None}
        self._build_ui()
        self.after(120, self._drain_log)

    def _build_ui(self):
        root = ttk.Frame(self, padding=10)
        root.pack(fill='both', expand=True)
        self.nb = ttk.Notebook(root)
        self.nb.pack(fill='both', expand=True)
        self._build_spider_tab()
        self._build_pdf_tab()
        self._build_compare_tab()
        log_frame = ttk.LabelFrame(root, text="运行日志", padding=8)
        log_frame.pack(fill='both', expand=False, pady=(10, 0))
        self.log_text = tk.Text(log_frame, height=10, wrap='word')
        self.log_text.pack(side='left', fill='both', expand=True)
        scroll = ttk.Scrollbar(log_frame, command=self.log_text.yview)
        scroll.pack(side='right', fill='y')
        self.log_text.configure(yscrollcommand=scroll.set)

    def row_file(self, parent, label, var, is_dir=False, filetypes=None):
        row = ttk.Frame(parent)
        row.pack(fill='x', pady=4)
        ttk.Label(row, text=label, width=16).pack(side='left')
        ttk.Entry(row, textvariable=var).pack(side='left', fill='x', expand=True, padx=5)
        def browse():
            if is_dir:
                path = filedialog.askdirectory()
            else:
                path = filedialog.askopenfilename(filetypes=filetypes or [("所有文件", "*")])
            if path:
                var.set(path)
        ttk.Button(row, text="选择", command=browse).pack(side='left')

    def _build_spider_tab(self):
        tab = ttk.Frame(self.nb, padding=12)
        self.nb.add(tab, text="1. 自动抓取")
        self.spider_save = tk.StringVar(value=str(Path.home() / "Desktop" / "雨课堂.xlsx"))
        self.spider_cycles = tk.IntVar(value=300)
        self.browser_name = tk.StringVar(value='edge')
        self.browser_start_mode = tk.StringVar(value='new')
        self.crawl_mode = tk.StringVar(value='auto')
        self.auto_pause = tk.BooleanVar(value=True)
        self.row_file(tab, "保存 Excel", self.spider_save, False, [("Excel", "*.xlsx")])
        row = ttk.Frame(tab); row.pack(fill='x', pady=4)
        ttk.Label(row, text="最大循环次数", width=16).pack(side='left')
        ttk.Spinbox(row, from_=1, to=10000, textvariable=self.spider_cycles, width=10).pack(side='left', padx=5)
        ttk.Checkbutton(row, text="自动暂停：出错时弹窗等待手动处理", variable=self.auto_pause).pack(side='left', padx=12)
        row2 = ttk.LabelFrame(tab, text="浏览器", padding=8); row2.pack(fill='x', pady=8)
        ttk.Radiobutton(row2, text="Edge", variable=self.browser_name, value='edge').pack(anchor='w')
        ttk.Radiobutton(row2, text="Chrome", variable=self.browser_name, value='chrome').pack(anchor='w')
        row3 = ttk.LabelFrame(tab, text="启动方式", padding=8); row3.pack(fill='x', pady=8)
        ttk.Radiobutton(row3, text="启动新浏览器并打开雨课堂首页", variable=self.browser_start_mode, value='new').pack(anchor='w')
        ttk.Radiobutton(row3, text="连接已打开的浏览器，跳过浏览器启动和打开首页", variable=self.browser_start_mode, value='attach').pack(anchor='w')
        row4 = ttk.LabelFrame(tab, text="抓取流程", padding=8); row4.pack(fill='x', pady=8)
        ttk.Radiobutton(row4, text="自动点击“再次答题 / 开始 / 交卷 / 查看试卷”并循环抓取", variable=self.crawl_mode, value="auto").pack(anchor="w")
        ttk.Radiobutton(row4, text="我已进入“查看试卷/结果页”，直接抓取当前页", variable=self.crawl_mode, value="current").pack(anchor="w")
        ttk.Radiobutton(row4, text="练习题模式：逐题抓取，自动点击「下一题」翻页", variable=self.crawl_mode, value="practice").pack(anchor="w")
        row5 = ttk.Frame(tab); row5.pack(fill='x', pady=4)
        self.use_ocr = tk.BooleanVar(value=False)
        ocr_hint = "（OCR 可用）" if is_ocr_available() else "（需安装 easyocr：pip install easyocr）"
        ttk.Checkbutton(row5, text=f"使用 OCR 截图识别题目（适用于图片/Canvas 渲染的题目）{ocr_hint}", variable=self.use_ocr).pack(anchor='w')
        btns = ttk.Frame(tab); btns.pack(fill='x', pady=10)
        ttk.Button(btns, text="开始自动抓取", command=self.start_spider).pack(side='left')
        ttk.Button(btns, text="停止抓取", command=self.stop_task).pack(side='left', padx=8)
        ttk.Label(tab, text="提示：连接已打开浏览器时，需要先用 --remote-debugging-port=9222 启动同一个浏览器，并停在目标页面。", foreground="#555").pack(anchor='w', pady=6)

    def _build_pdf_tab(self):
        tab = ttk.Frame(self.nb, padding=12)
        self.nb.add(tab, text="2. 生成 PDF")
        self.pdf_excel = tk.StringVar(value=str(Path.home() / "Desktop" / "雨课堂.xlsx"))
        self.pdf_outdir = tk.StringVar(value=str(Path.home() / "Desktop"))
        self.pdf_base = tk.StringVar(value="雨课堂毛概题库")
        self.pdf_header = tk.StringVar(value=DEFAULT_HEADER)
        self.pdf_title = tk.StringVar(value=DEFAULT_TITLE)
        self.make_inline = tk.BooleanVar(value=True)
        self.make_practice = tk.BooleanVar(value=True)
        self.encrypt_pdf = tk.BooleanVar(value=True)
        self.row_file(tab, "Excel 文件", self.pdf_excel, False, [("Excel", "*.xlsx")])
        self.row_file(tab, "输出文件夹", self.pdf_outdir, True)
        for label, var in [("PDF 文件名前缀", self.pdf_base), ("页眉文字", self.pdf_header), ("标题文字", self.pdf_title)]:
            row = ttk.Frame(tab); row.pack(fill='x', pady=4)
            ttk.Label(row, text=label, width=16).pack(side='left')
            ttk.Entry(row, textvariable=var).pack(side='left', fill='x', expand=True, padx=5)
        row = ttk.Frame(tab); row.pack(fill='x', pady=6)
        ttk.Checkbutton(row, text="输出背诵版 PDF", variable=self.make_inline).pack(side='left')
        ttk.Checkbutton(row, text="输出练习版 PDF", variable=self.make_practice).pack(side='left', padx=12)
        ttk.Checkbutton(row, text="使用原脚本 PDF 防修改设置", variable=self.encrypt_pdf).pack(side='left')
        ttk.Button(tab, text="生成 PDF", command=self.start_pdf).pack(anchor='w', pady=10)

    def _build_compare_tab(self):
        tab = ttk.Frame(self.nb, padding=12)
        self.nb.add(tab, text="3. PDF 比对")
        self.c_pdf1 = tk.StringVar()
        self.c_pdf2 = tk.StringVar()
        self.c_out = tk.StringVar(value=str(Path.home() / "Desktop" / "差异题目汇总.txt"))
        self.row_file(tab, "PDF 文件1", self.c_pdf1, False, [("PDF", "*.pdf")])
        self.row_file(tab, "PDF 文件2", self.c_pdf2, False, [("PDF", "*.pdf")])
        row = ttk.Frame(tab); row.pack(fill='x', pady=4)
        ttk.Label(row, text="输出 TXT", width=16).pack(side='left')
        ttk.Entry(row, textvariable=self.c_out).pack(side='left', fill='x', expand=True, padx=5)
        def saveas():
            path = filedialog.asksaveasfilename(defaultextension='.txt', filetypes=[('文本文件', '*.txt')])
            if path:
                self.c_out.set(path)
        ttk.Button(row, text="选择", command=saveas).pack(side='left')
        ttk.Button(tab, text="开始比对", command=self.start_compare).pack(anchor='w', pady=10)

    def log(self, msg: str):
        self.log_queue.put(str(msg))

    def _drain_log(self):
        while True:
            try:
                msg = self.log_queue.get_nowait()
            except queue.Empty:
                break
            self.log_text.insert('end', msg + '\n')
            self.log_text.see('end')
        if self.worker and not self.worker.is_alive():
            self.worker = None
            self.driver_state["driver"] = None
            self.stop_event.clear()
        self.after(120, self._drain_log)

    def run_in_thread(self, target):
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("任务运行中", "当前已有任务在运行，请先等待或停止。")
            return
        self.stop_event.clear()
        self.worker = threading.Thread(target=target, daemon=True)
        self.worker.start()

    def stop_task(self):
        self.stop_event.set()
        driver = self.driver_state.get("driver")
        if driver is not None:
            try:
                driver.quit()
                self.log("⏹ 已强制关闭浏览器会话，正在终止抓取任务。")
            except Exception as e:
                self.log(f"⏹ 已请求强制停止，但关闭浏览器会话时出错：{e}")
        else:
            self.log("⏹ 已请求停止。当前没有可关闭的浏览器会话。")

    def start_spider(self):
        def job():
            try:
                run_auto_spider(
                    self.spider_save.get(),
                    int(self.spider_cycles.get()),
                    self.browser_name.get(),
                    self.browser_start_mode.get(),
                    self.crawl_mode.get(),
                    self.auto_pause.get(),
                    self.log,
                    self.stop_event,
                    self.driver_state,
                    use_ocr=self.use_ocr.get(),
                )
            except Exception:
                if self.stop_event.is_set():
                    self.log("⏹ 抓取任务已终止。")
                else:
                    self.log("❌ 任务失败：\n" + traceback.format_exc())
        self.run_in_thread(job)

    def start_pdf(self):
        def job():
            try:
                outputs = generate_exam_pdf(
                    self.pdf_excel.get(), self.pdf_outdir.get(), self.pdf_base.get(), self.pdf_header.get(), self.pdf_title.get(),
                    self.make_inline.get(), self.make_practice.get(), self.encrypt_pdf.get(), self.log
                )
                self.log("🎉 PDF 生成完成：" + ", ".join(outputs))
            except Exception:
                self.log("❌ PDF 生成失败：\n" + traceback.format_exc())
        self.run_in_thread(job)

    def start_compare(self):
        def job():
            try:
                compare_pdfs(self.c_pdf1.get(), self.c_pdf2.get(), self.c_out.get(), self.log)
            except Exception:
                self.log("❌ 比对失败：\n" + traceback.format_exc())
        self.run_in_thread(job)


if __name__ == '__main__':
    app = App()
    app.mainloop()
