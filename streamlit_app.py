# -*- coding: utf-8 -*-
"""
streamlit_app.py
-----------------
واجهة ويب بسيطة: اسحب ملف PDF (أو أكتر من ملف) واحصل فورًا على ملف Excel
فيه جدول Bill of Material وجدول Structure Quantity.

تشغيل:
    pip install -r requirements.txt
    streamlit run streamlit_app.py
"""

import io
import zipfile

import streamlit as st

from extractor import (
    export_to_excel,
    find_tables_with_ocr_fallback,
    parse_qty_table,
    pick_bom_table,
    pick_qty_table,
)

st.set_page_config(page_title="تحويل جداول Bill of Material من PDF إلى Excel", layout="wide")

st.title("📐 تحويل جداول Bill of Material من PDF إلى Excel")
st.caption(
    "ارفع ملف/ملفات PDF لرسومات الحديد (Steel Structure Drawings) وسيتم استخراج "
    "جدول BILL OF MATERIAL OF STRUCTURE وجدول STRUCTURE QUANTITY REQUIRED تلقائيًا "
    "وتحويلهما إلى ملف Excel."
)

uploaded_files = st.file_uploader("ارفع ملف أو أكثر بصيغة PDF", type=["pdf"], accept_multiple_files=True)
force_ocr = st.checkbox(
    "إجبار استخدام OCR لكل الملفات (استخدمها لو متأكد إن الملفات صور ممسوحة/سكان)",
    value=False,
)
st.caption(
    "افتراضيًا: البرنامج يحاول القراءة من طبقة النص أولًا، ولو الملف طلع صورة ممسوحة "
    "(مفيهوش طبقة نص) هيتحول تلقائيًا لاستخدام OCR بدون ما تعمل حاجة."
)

if uploaded_files:
    results = []  # (filename, xlsx_bytes)

    for uf in uploaded_files:
        st.divider()
        st.subheader(f"📄 {uf.name}")

        # نحفظ الملف مؤقتًا على القرص لأن pdfplumber و pdf2image محتاجين path حقيقي
        pdf_bytes = uf.read()
        tmp_pdf_path = f"/tmp/_upload_{uf.name}"
        with open(tmp_pdf_path, "wb") as f:
            f.write(pdf_bytes)

        with st.spinner("جاري تحليل الملف واستخراج الجداول (قد يستغرق OCR وقتًا أطول)..."):
            tables = find_tables_with_ocr_fallback(tmp_pdf_path, force_ocr=force_ocr)
            bom = pick_bom_table(tables)
            qty_table = pick_qty_table(tables)
            qty_info = parse_qty_table(qty_table) if qty_table else None

        col1, col2 = st.columns(2)

        with col1:
            st.markdown("**Bill of Material**")
            if bom and bom.rows:
                st.dataframe(bom.rows, use_container_width=True, hide_index=True)
            else:
                st.warning("لم يتم العثور على جدول Bill of Material.")

        with col2:
            st.markdown("**Structure Quantity**")
            if qty_info:
                st.json(qty_info)
            else:
                st.warning("لم يتم العثور على جدول Structure Quantity.")

        # تصدير Excel في الذاكرة
        out_buffer = io.BytesIO()
        tmp_xlsx_path = f"/tmp/{uf.name}.xlsx"
        export_to_excel(bom, qty_info, tmp_xlsx_path, sheet_title=uf.name.replace(".pdf", ""))
        with open(tmp_xlsx_path, "rb") as f:
            xlsx_bytes = f.read()

        results.append((uf.name.replace(".pdf", ".xlsx"), xlsx_bytes))

        st.download_button(
            label=f"⬇️ تحميل Excel لملف {uf.name}",
            data=xlsx_bytes,
            file_name=uf.name.replace(".pdf", ".xlsx"),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    if len(results) > 1:
        st.divider()
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for fname, data in results:
                zf.writestr(fname, data)
        st.download_button(
            label="⬇️ تحميل كل الملفات كـ ZIP",
            data=zip_buffer.getvalue(),
            file_name="bill_of_material_export.zip",
            mime="application/zip",
        )
else:
    st.info("ارفع ملفات PDF من الأعلى للبدء.")
