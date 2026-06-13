import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import tempfile
import os
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pdf2docx import Converter
import unicodedata
from streamlit_gsheets import GSheetsConnection

# -----------------------------------------------------------------------------
# 1. إعدادات النظام
# -----------------------------------------------------------------------------
st.set_page_config(page_title="نظام المقارنة السحابي المطور", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: var(--primary-color); color: white; width: 100%; font-weight: bold; border-radius: 8px; border: none; height: 45px; }
    .report-box { background-color: var(--secondary-background-color); color: var(--text-color); padding: 15px; border-radius: 8px; border-right: 5px solid var(--primary-color); text-align: right; margin-bottom: 15px; box-shadow: 0px 2px 5px rgba(0,0,0,0.05); }
    .report-box h4 { color: var(--text-color); margin: 5px 0; }
    </style>
""", unsafe_allow_html=True)

try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception:
    pass

# -----------------------------------------------------------------------------
# 2. محركات المعالجة الذكية والتحويل التلقائي
# -----------------------------------------------------------------------------
def fix_arabic_text(text):
    if not text: return text
    text = unicodedata.normalize('NFKC', text)
    if any('\u0600' <= char <= '\u06FF' for char in text):
        words = text.split()
        fixed_words = [word[::-1] if any('\u0600' <= c <= '\u06FF' for c in word) else word for word in words]
        return " ".join(fixed_words[::-1])
    return text

def parse_row(cells, is_original_pdf=False):
    clean_cells = [str(c).strip().replace('\n', ' ') if c is not None else "" for c in cells]
    joined = "".join(clean_cells)
    
    if not any(clean_cells) or "المركز" in joined or "الوكيل" in joined or "اسم" in joined or "زكرمل" in joined: 
        return {}
        
    name_idx = -1
    max_len = 0
    for i, c in enumerate(clean_cells):
        if any('\u0600' <= char <= '\u06FF' for char in c) and not any(char.isdigit() for char in c):
            if len(c) > max_len: 
                max_len = len(c)
                name_idx = i
                
    if name_idx == -1: return {}
    
    raw_name = clean_cells[name_idx]
    # تطبيق التعديل فقط إذا كان الملف الأصلي PDF (لأن التحويل الداخلي قد ينقل الحروف مقلوبة)
    final_name = fix_arabic_text(raw_name) if is_original_pdf else raw_name

    card_num = "-"
    card_cands = [c for c in clean_cells if c.isdigit() and len(c) >= 5]
    if card_cands: card_num = card_cands[0]
    else: return {}

    small_before = [int(clean_cells[i]) for i in range(name_idx) if clean_cells[i].isdigit() and len(clean_cells[i]) < 5]
    small_after = [int(clean_cells[i]) for i in range(name_idx + 1, len(clean_cells)) if clean_cells[i].isdigit() and len(clean_cells[i]) < 5]
    
    seq, total, eligible, withheld = "-", 0, 0, 0
    
    if len(small_before) >= 2:
        if len(small_before) >= 3: withheld, eligible, total = small_before[-3], small_before[-2], small_before[-1]
        else: withheld, eligible, total = 0, small_before[-2], small_before[-1]
        seq = small_after[0] if small_after else "-"
    elif len(small_after) >= 2:
        if len(small_after) >= 3: total, eligible, withheld = small_after[0], small_after[1], small_after[2]
        else: total, eligible, withheld = small_after[0], small_after[1], 0
        seq = small_before[0] if small_before else "-"
    else: return {}
        
    return {card_num: {"seq": seq, "name": final_name, "total": total, "eligible": eligible, "withheld": withheld}}

def extract_clean_records(file_obj, is_pdf=False):
    """الآن: إذا كان PDF يحوله لوورد في الخلفية، ثم يقرأه كلوورد!"""
    records = {}
    
    if is_pdf:
        # 1. إنشاء ملفات مؤقتة للتحويل الصامت
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(file_obj.read())
            pdf_path = tmp_pdf.name
            
        docx_path = pdf_path.replace('.pdf', '.docx')
        
        # 2. عملية التحويل الآلية (من PDF إلى Word)
        cv = Converter(pdf_path)
        cv.convert(docx_path, start=0, end=None)
        cv.close()
        
        # 3. قراءة الملف الناتج كأنه ملف Word طبيعي
        doc = Document(docx_path)
        
        # 4. تنظيف الخادم ومسح الملفات المؤقتة
        os.remove(pdf_path)
        os.remove(docx_path)
    else:
        # إذا تم رفع Word مباشرة (مثل الملفات التي قمت بتحويلها بـ iLovePDF مسبقاً)
        doc = Document(file_obj)
        
    # قراءة الجداول (سواء كان Word أصلي أو محول من الـ PDF)
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            records.update(parse_row(cells, is_original_pdf=is_pdf))
            
    return records

def compare_records(old_data, new_data):
    results, counters = [], {"total_fam": 0, "eligible_fam": 0, "withheld_fam": 0, "added_fam": 0, "deleted_fam": 0, "net_total": 0, "net_eligible": 0, "net_withheld": 0}
    all_cards = set(old_data.keys()).union(set(new_data.keys()))
    for card in all_cards:
        if card in old_data and card in new_data:
            old_v, new_v = old_data[card], new_data[card]
            if old_v["total"] != new_v["total"] or old_v["eligible"] != new_v["eligible"] or old_v["withheld"] != new_v["withheld"]:
                if old_v["total"] != new_v["total"]: counters["total_fam"] += 1; counters["net_total"] += (new_v["total"] - old_v["total"])
                if old_v["eligible"] != new_v["eligible"]: counters["eligible_fam"] += 1; counters["net_eligible"] += (new_v["eligible"] - old_v["eligible"])
                if old_v["withheld"] != new_v["withheld"]: counters["withheld_fam"] += 1; counters["net_withheld"] += (new_v["withheld"] - old_v["withheld"])
                results.append({"التسلسل": new_v["seq"], "رقم البطاقة": card, "الاسم (سابقاً)": old_v["name"], "الاسم (حالياً)": new_v["name"], "الكلية (سابقاً)": old_v["total"], "الكلية (حالياً)": new_v["total"], "المستحقة (سابقاً)": old_v["eligible"], "المستحقة (حالياً)": new_v["eligible"], "المحجوبين (سابقاً)": old_v["withheld"], "المحجوبين (حالياً)": new_v["withheld"]})
        elif card in old_data and card not in new_data:
            old_v = old_data[card]
            counters["deleted_fam"] += 1; counters["net_total"] -= old_v["total"]; counters["net_eligible"] -= old_v["eligible"]; counters["net_withheld"] -= old_v["withheld"]
            results.append({"التسلسل": old_v["seq"], "رقم البطاقة": card, "الاسم (سابقاً)": old_v["name"], "الاسم (حالياً)": "❌ (محذوف / منقول)", "الكلية (سابقاً)": old_v["total"], "الكلية (حالياً)": 0, "المستحقة (سابقاً)": old_v["eligible"], "المستحقة (حالياً)": 0, "المحجوبين (سابقاً)": old_v["withheld"], "المحجوبين (حالياً)": 0})
        elif card not in old_data and card in new_data:
            new_v = new_data[card]
            counters["added_fam"] += 1; counters["net_total"] += new_v["total"]; counters["net_eligible"] += new_v["eligible"]; counters["net_withheld"] += new_v["withheld"]
            results.append({"التسلسل": new_v["seq"], "رقم البطاقة": card, "الاسم (سابقاً)": "✨ (مضاف حديثاً)", "الاسم (حالياً)": new_v["name"], "الكلية (سابقاً)": 0, "الكلية (حالياً)": new_v["total"], "المستحقة (سابقاً)": 0, "المستحقة (حالياً)": new_v["eligible"], "المحجوبين (سابقاً)": 0, "المحجوبين (حالياً)": new_v["withheld"]})
    return results, counters

def create_word_table_report(df, title):
    doc = Document(); heading = doc.add_heading(title, level=1); heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cols = list(df.columns)[::-1]; table = doc.add_table(rows=1, cols=len(cols)); table.style = 'Table Grid'
    for i, col in enumerate(cols): table.rows[0].cells[i].text = str(col)
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, col in enumerate(cols): row_cells[i].text = str(row[col])
    buffer = BytesIO(); doc.save(buffer); buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# 3. واجهة التطبيق الرئيسية 
# -----------------------------------------------------------------------------
tab1, tab2 = st.tabs(["🔎 إجراء مقارنة ذكية", "📜 الأرشيف التاريخي"])

with tab1:
    st.markdown("<h3 style='text-align: right;'>لوحة المطابقة (تحويل آلي مدمج)</h3>", unsafe_allow_html=True)
    st.info("💡 ارفع ملفات الـ PDF وسيقوم النظام بتحويلها لـ Word بالخلفية فوراً قبل مقارنتها!")
    
    col1, col2 = st.columns(2)
    with col1: new_file = st.file_uploader("الملف الجديد", type=['pdf', 'docx'], key="n_f")
    with col2: old_file = st.file_uploader("الملف القديم", type=['pdf', 'docx'], key="o_f")

    if st.button("🚀 تشغيل الفحص والمطابقة"):
        if old_file and new_file:
            with st.spinner('جاري التحويل الآلي ومطابقة البيانات...'):
                try:
                    is_old_pdf = old_file.name.lower().endswith('.pdf')
                    old_data = extract_clean_records(old_file, is_pdf=is_old_pdf)
                    
                    is_new_pdf = new_file.name.lower().endswith('.pdf')
                    new_data = extract_clean_records(new_file, is_pdf=is_new_pdf)

                    results, counters = compare_records(old_data, new_data)
                    
                    if results:
                        st.session_state['c_results'] = pd.DataFrame(results)
                        st.session_state['c_counters'] = counters
                        st.session_state['c_filename'] = new_file.name.rsplit('.', 1)[0]
                        st.success("✅ تمت عملية التحويل والقراءة والمطابقة بنجاح!")
                    else: st.info("تطابق كامل بين الملفين.")
                except Exception as e: st.error(f"حدث خطأ: {e}")
        else: st.warning("يرجى رفع الملفات أولاً.")

    if 'c_results' in st.session_state:
        df_res = st.session_state['c_results']
        cnt = st.session_state['c_counters']
        
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f"<div class='report-box'>🔹 الكلية: {cnt['net_total']:+d}</div>", unsafe_allow_html=True)
        with c2: st.markdown(f"<div class='report-box'>🔹 المستحقة: {cnt['net_eligible']:+d}</div>", unsafe_allow_html=True)
        with c3: st.markdown(f"<div class='report-box'>🔹 مضاف: {cnt['added_fam']} | محذوف: {cnt['deleted_fam']}</div>", unsafe_allow_html=True)
        
        st.dataframe(df_res, use_container_width=True, hide_index=True)
        st.download_button("📥 تحميل التقرير (Word)", data=create_word_table_report(df_res, st.session_state['c_filename']), file_name=f"{st.session_state['c_filename']}.docx")

with tab2:
    if st.button("🔄 جلب الأرشيف"):
        try:
            st.dataframe(conn.read(worksheet="Sheet1", ttl=0), use_container_width=True, hide_index=True)
        except Exception as ex: st.error(ex)
