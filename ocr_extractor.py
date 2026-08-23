# -*- coding: utf-8 -*-
"""
ocr_extractor.py
-----------------
استخراج جداول من صفحات PDF ممسوحة ضوئيًا (Scanned) لا تحتوي طبقة نص حقيقية.

الفكرة:
بدل عمل OCR على الصفحة كلها دفعة واحدة (وده بيدي نتيجة ركيكة لأي جدول
لأنه بيخلط بين الأعمدة والصفوف)، بنعمل الآتي:

1) نحوّل صفحة الـ PDF لصورة عالية الدقة (300 DPI فأكتر).
2) نستخدم OpenCV لاكتشاف خطوط الجدول نفسها (الخطوط الأفقية والرأسية) لأن
   هذا النوع من الجداول الهندسية دائمًا محاط بخطوط واضحة.
3) من تقاطع الخطوط، نبني إحداثيات كل خلية في الجدول (Grid).
4) نعمل OCR منفصل على كل خلية على حدة (تسريسكت Tesseract) بدل الصفحة كلها،
   وده بيرفع الدقة بشكل كبير لأن كل خلية بتترا منفردة من غير تشويش من
   الخلايا المجاورة.
5) نرجع نفس شكل الناتج اللي بيرجعه pdfplumber.extract_tables() (list of
   rows) عشان يشتغل مباشرة مع باقي الكود في extractor.py.

المتطلبات على مستوى النظام (System-level, مش pip فقط):
    - tesseract-ocr  (على Ubuntu/Debian: sudo apt-get install tesseract-ocr tesseract-ocr-ara)
    - poppler-utils  (لازم لـ pdf2image: sudo apt-get install poppler-utils)

المتطلبات على مستوى pip: pytesseract, pdf2image, opencv-python, numpy, Pillow
(موجودين في requirements.txt)
"""

from dataclasses import dataclass
from typing import List, Optional

import cv2
import numpy as np
import pytesseract
from pdf2image import convert_from_path

# لو عايز تدعم OCR عربي كمان (نادر في رسومات إنشائية لكن احتياطي) غيّرها لـ "eng+ara"
OCR_LANG = "eng"

DPI = 350  # كل ما زادت الدقة كل ما تحسنت جودة قراءة الأرقام الصغيرة، لكن بتبطئ المعالجة


@dataclass
class OcrTable:
    page_number: int
    rows: List[List[str]]


def _detect_grid_lines(gray_img: np.ndarray):
    """يكتشف الخطوط الأفقية والرأسية في صورة بالأبيض والأسود ويرجع إحداثياتهم."""
    # عتبة ثنائية عكسية: الخطوط السودة بتبقى بيضا على خلفية سودا
    bw = cv2.adaptiveThreshold(
        gray_img, 255, cv2.ADAPTIVE_THRESH_MEAN_C, cv2.THRESH_BINARY_INV, 15, -2
    )

    h, w = gray_img.shape

    # اكتشاف الخطوط الأفقية
    horiz_size = max(w // 30, 20)
    horiz_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (horiz_size, 1))
    horizontal = cv2.erode(bw, horiz_kernel)
    horizontal = cv2.dilate(horizontal, horiz_kernel)

    # اكتشاف الخطوط الرأسية
    vert_size = max(h // 30, 20)
    vert_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, vert_size))
    vertical = cv2.erode(bw, vert_kernel)
    vertical = cv2.dilate(vertical, vert_kernel)

    return horizontal, vertical


def _lines_to_positions(mask: np.ndarray, axis: int, min_gap: int = 10) -> List[int]:
    """
    يحول قناع الخطوط (horizontal أو vertical) إلى قائمة إحداثيات (y لو أفقي، x لو رأسي)
    بعد تجميع الخطوط القريبة من بعضها في خط واحد.
    """
    projection = mask.sum(axis=axis)  # axis=1 للأفقي (نجمع على الصفوف), axis=0 للرأسي
    positions = np.where(projection > projection.max() * 0.3)[0]
    if len(positions) == 0:
        return []

    grouped = [int(positions[0])]
    for p in positions[1:]:
        if p - grouped[-1] > min_gap:
            grouped.append(int(p))
        else:
            grouped[-1] = int((grouped[-1] + p) / 2)
    return grouped


def _ocr_cell(image: np.ndarray, x1: int, y1: int, x2: int, y2: int, pad: int = 3) -> str:
    h, w = image.shape[:2]
    x1, y1 = max(x1 + pad, 0), max(y1 + pad, 0)
    x2, y2 = min(x2 - pad, w), min(y2 - pad, h)
    if x2 <= x1 or y2 <= y1:
        return ""
    cell = image[y1:y2, x1:x2]
    # تكبير الخلية شوية بيحسن قراءة الأرقام الصغيرة
    cell = cv2.resize(cell, None, fx=2, fy=2, interpolation=cv2.INTER_CUBIC)
    text = pytesseract.image_to_string(
        cell, lang=OCR_LANG, config="--psm 6"
    )
    return " ".join(text.split())  # تنضيف المسافات والأسطر الزيادة


def extract_table_from_scanned_page(pil_image, page_number: int) -> Optional[OcrTable]:
    """يستخرج جدول (لو موجود) من صورة صفحة واحدة عن طريق كشف الخطوط + OCR لكل خلية."""
    img = cv2.cvtColor(np.array(pil_image), cv2.COLOR_RGB2BGR)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    horizontal, vertical = _detect_grid_lines(gray)

    y_positions = _lines_to_positions(horizontal, axis=1, min_gap=15)  # خطوط أفقية -> إحداثيات y
    x_positions = _lines_to_positions(vertical, axis=0, min_gap=15)    # خطوط رأسية -> إحداثيات x

    if len(y_positions) < 2 or len(x_positions) < 2:
        return None  # مفيش جدول واضح بخطوط في الصفحة دي

    rows: List[List[str]] = []
    for ry in range(len(y_positions) - 1):
        y1, y2 = y_positions[ry], y_positions[ry + 1]
        if y2 - y1 < 8:  # سطر رفيع جدًا، على الأغلب خط مزدوج وليس صف حقيقي
            continue
        row_cells = []
        for rx in range(len(x_positions) - 1):
            x1, x2 = x_positions[rx], x_positions[rx + 1]
            if x2 - x1 < 8:
                continue
            text = _ocr_cell(gray, x1, y1, x2, y2)
            row_cells.append(text)
        if any(c for c in row_cells):
            rows.append(row_cells)

    if not rows:
        return None

    return OcrTable(page_number=page_number, rows=rows)


def find_tables_via_ocr(pdf_path: str) -> List[OcrTable]:
    """يحوّل كل صفحات الـ PDF لصور ويحاول استخراج جدول من كل صفحة."""
    tables: List[OcrTable] = []
    pages = convert_from_path(pdf_path, dpi=DPI)
    for i, page_img in enumerate(pages, start=1):
        t = extract_table_from_scanned_page(page_img, i)
        if t:
            tables.append(t)
    return tables


if __name__ == "__main__":
    import sys
    if len(sys.argv) < 2:
        print("الاستخدام: python ocr_extractor.py input.pdf")
        sys.exit(1)
    result = find_tables_via_ocr(sys.argv[1])
    for t in result:
        print(f"--- صفحة {t.page_number}: {len(t.rows)} صف ---")
        for r in t.rows:
            print(r)
