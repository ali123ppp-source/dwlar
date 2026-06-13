import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import pdfplumber
import arabic_reshaper
from bidi.algorithm import get_display
from streamlit_gsheets import GSheetsConnection

# -----------------------------------------------------------------------------
# 1. إعدادات النظام والمظهر
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

# -----------------------------------------------------------------------------
# 2. إنشاء الاتصال بجوجل شيت 
# -----------------------------------------------------------------------------
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception:
    pass

# -----------------------------------------------------------------------------
# 3. محركات المعالجة الذكية (قراءة البيانات من PDF و Word مباشرة)
# -----------------------------------------------------------------------------
def fix_arabic_text(text):
    """دالة لتعديل الحروف العربية المقلوبة لتظهر بشكل سليم"""
    if not text:
        return text
    if any('\u0600' <= char <= '\u06FF' or '\uFE70' <= char <= '\uFEFF' or '\uFB50' <= char <= '\uFDFF' for char in text):
        reshaped_text = arabic_reshaper.reshape(text)
        return get_display(reshaped_text)
    return text

def parse_row(cells, is_pdf=False):
    """دالة تحليل ذكية تفهم ترتيب الأعمدة العشوائي في الـ PDF والوورد"""
    clean_cells = [str(c).strip().replace('\n', ' ') if c is not None else "" for c in cells]
    joined = "".join(clean_cells)
    
    # تجاهل العناوين والصفوف الفارغة
    if not any(clean_cells) or "المركز" in joined or "الوكيل" in joined or "اسم" in joined or "زكرمل" in joined: 
        return {}
        
    name_idx = -1
    max_len = 0
    # 1. البحث عن عمود الاسم
    for i, c in enumerate(clean_cells):
        if any('\u0600' <= char <= '\u06FF' or '\uFE70' <= char <= '\uFEFF' or '\uFB50' <= char <= '\uFDFF' for char in c) and not any(char.isdigit() for char in c):
            if len(c) > max_len: 
                max_len = len(c)
                name_idx = i
                
    if name_idx == -1: return {}
    
    # تعديل الاسم العربي المقلوب فقط بعد التعرف عليه
    raw_name = clean_cells[name_idx]
    final_name = fix_arabic_text(raw_name) if is_pdf else raw_name

    # 2. البحث عن رقم البطاقة (رقم طويل يتكون من 5 خانات أو أكثر)
    card_num = "-"
    card_cands = [c for c in clean_cells if c.isdigit() and len(c) >= 5]
    if card_cands:
        card_num = card_cands[0]
    else:
        return {}

    # 3. استخراج الأرقام الصغيرة (الكلية، المستحقة، المحجوبين، والتسلسل)
    # الأرقام قبل الاسم
    small_before = [int(clean_cells[i]) for i in range(name_idx) if clean_cells[i].isdigit() and len(clean_cells[i]) < 5]
    # الأرقام بعد الاسم
    small_after = [int(clean_cells[i]) for i in range(name_idx + 1, len(clean_cells)) if clean_cells[i].isdigit() and len(clean_cells[i]) < 5]
    
    seq = "-"
    total = eligible = withheld = 0
    
    # 4. توزيع الأرقام بناءً على ترتيب الأعمدة (يسار ليمين أو يمين ليسار)
    if len(small_before) >= 2:
        # ترتيب الـ PDF غالباً
        if len(small_before) >= 3:
            withheld, eligible, total = small_before[-3], small_before[-2], small_before[-1]
        else:
            withheld, eligible, total = 0, small_before[-2], small_before[-1]
        seq = small_after[0] if small_after else "-"
        
    elif len(small_after) >= 2:
        # ترتيب الوورد غالباً
        if len(small_after) >= 3:
            total, eligible, withheld = small_after[0], small_after[1], small_after[2]
        else:
            total, eligible, withheld = small_after[0], small_after[1], 0
        seq = small_before[0] if small_before else "-"
    else:
        return {}
        
    return {card_num: {"seq": seq, "name": final_name, "total": total, "eligible": eligible, "withheld": withheld}}

def extract_clean_records(file_obj, is_pdf=False):
    """قراءة الجداول مباشرة من الملفات"""
    records = {}
    if is_pdf:
        with pdfplumber.open(file_obj) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    for row in table:
                        # تمرير الصف خام بدون تغيير
                        records.update(parse_row(row, is_pdf=True))
    else:
        doc = Document(file_obj)
        for table in doc.tables:
            for row in table.rows:
                cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
                records.update(parse_row(cells, is_pdf=False))
    return records

def compare_records(old_data, new_data):
    results = []
    counters = {"total_fam": 0, "eligible_fam": 0, "withheld_fam": 0, "added_fam": 0, "deleted_fam": 0, "net_total": 0, "net_eligible": 0, "net_withheld": 0}
    all_cards = set(old_data.keys()).union(set(new_data.keys()))
    for card in all_cards:
        if card in old_data and card in new_data:
            old_v, new_v = old_data[card], new_data[card]
            diff_total = old_v["total"] != new_v["total"]; diff_elig = old_v["eligible"] != new_v["eligible"]; diff_with = old_v["withheld"] != new_v["withheld"]
            if diff_total or diff_elig or diff_with:
                if diff_total: counters["total_fam"] += 1; counters["net_total"] += (new_v["total"] - old_v["total"])
                if diff_elig: counters["eligible_fam"] += 1; counters["net_eligible"] += (new_v["eligible"] - old_v["eligible"])
                if diff_with: counters["withheld_fam"] += 1; counters["net_withheld"] += (new_v["withheld"] - old_v["withheld"])
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
    doc = Document()
    heading = doc.add_heading(title, level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cols = list(df.columns)[::-1]
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = 'Table Grid'
    for i, col in enumerate(cols): table.rows[0].cells[i].text = str(col)
    for _, row in df.iterrows():
        row_cells = table.add_row().cells
        for i, col in enumerate(cols): row_cells[i].text = str(row[col])
    buffer = BytesIO(); doc.save(buffer); buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# 4. واجهة التطبيق الرئيسية (تبويبات)
# -----------------------------------------------------------------------------
tab1, tab2 = st.tabs(["🔎 إجراء مقارنة ذكية (يدعم PDF و Word)", "📜 الأرشيف التاريخي"])

with tab1:
    st.markdown("<h3 style='text-align: right;'>لوحة المطابقة الرقمية التلقائية</h3>", unsafe_allow_html=True)
    st.info("💡 يمكنك رفع ملفات الـ PDF أو Word مباشرة! سيقوم النظام بقراءة البيانات من الداخل في لمح البصر.")
    
    col1, col2 = st.columns(2)
    with col1: new_file = st.file_uploader("الملف الجديد (PDF أو Word)", type=['pdf', 'docx'], key="n_f")
    with col2: old_file = st.file_uploader("الملف القديم (PDF أو Word)", type=['pdf', 'docx'], key="o_f")

    if st.button("🚀 تشغيل الفحص والمطابقة"):
        if old_file and new_file:
            with st.spinner('جاري قراءة البيانات ومطابقتها...'):
                try:
                    # قراءة الملف القديم
                    is_old_pdf = old_file.name.lower().endswith('.pdf')
                    old_data = extract_clean_records(old_file, is_pdf=is_old_pdf)
                    
                    # قراءة الملف الجديد
                    is_new_pdf = new_file.name.lower().endswith('.pdf')
                    new_data = extract_clean_records(new_file, is_pdf=is_new_pdf)

                    # المطابقة المباشرة
                    results, counters = compare_records(old_data, new_data)
                    
                    if results:
                        df_results = pd.DataFrame(results)[["التسلسل", "رقم البطاقة", "الاسم (سابقاً)", "الاسم (حالياً)", "الكلية (سابقاً)", "الكلية (حالياً)", "المستحقة (سابقاً)", "المستحقة (حالياً)", "المحجوبين (سابقاً)", "المحجوبين (حالياً)"]]
                        st.session_state['c_results'] = df_results
                        st.session_state['c_counters'] = counters
                        st.session_state['c_filename'] = new_file.name.rsplit('.', 1)[0]
                        st.success("✅ تمت عملية القراءة والمطابقة بنجاح وفي ثوانٍ!")
                    else: 
                        st.info("تطابق كامل بين الملفين، لا توجد فروقات.")
                except Exception as e:
                    st.error(f"حدث خطأ أثناء قراءة الملفات: {e}")
        else: st.warning("يرجى رفع الملفات أولاً.")

    # عرض النتائج
    if 'c_results' in st.session_state:
        df_res = st.session_state['c_results']
        cnt = st.session_state['c_counters']
        
        st.markdown("<h4 style='text-align: right;'>📊 خلاصة المتغيرات الحالية</h4>", unsafe_allow_html=True)
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f"<div class='report-box'>🔹 حركة الكلية<br><h4>{cnt['total_fam']} عائلة</h4>الصافي: {cnt['net_total']:+d}</div>", unsafe_allow_html=True)
        with c2: st.markdown(f"<div class='report-box'>🔹 حركة المستحقة<br><h4>{cnt['eligible_fam']} عائلة</h4>الصافي: {cnt['net_eligible']:+d}</div>", unsafe_allow_html=True)
        with c3: st.markdown(f"<div class='report-box'>🔹 الحالات<br><h4>مضاف: {cnt['added_fam']} | محذوف: {cnt['deleted_fam']}</h4></div>", unsafe_allow_html=True)
        
        # 💾 زر حفظ النتائج في جوجل شيت
        if st.button("💾 ترحيل وحفظ هذه العملية في أرشيف جوجل شيت السحابي"):
            with st.spinner("جاري الترحيل للسحابة..."):
                try:
                    existing_df = conn.read(worksheet="Sheet1", ttl=0)
                    new_row = pd.DataFrame([{
                        "التاريخ": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "الوكيل / المستخدم": "مستخدم (بدون تسجيل)",
                        "اسم الملف المفحوص": st.session_state['c_filename'],
                        "حركة الكلية (عائلات)": cnt['total_fam'],
                        "صافي الأفراد (كلية)": cnt['net_total'],
                        "حركة المستحقة (عائلات)": cnt['eligible_fam'],
                        "صافي الأفراد (مستحقة)": cnt['net_eligible'],
                        "العائلات المضافة": cnt['added_fam'],
                        "العائلات المحذوفة": cnt['deleted_fam']
                    }])
                    updated_df = pd.concat([existing_df, new_row], ignore_index=True)
                    conn.update(worksheet="Sheet1", data=updated_df)
                    st.success("🚀 تم ترحيل البيانات بنجاح إلى أرشيف جوجل شيت!")
                except Exception as ex:
                    st.error(f"فشل الاتصال بالشيت، تأكد من إعدادات Secrets: {ex}")

        st.dataframe(df_res, use_container_width=True, hide_index=True)
        
        st.markdown("---")
        st.markdown("<h4 style='text-align: right;'>📥 قسم التحميلات</h4>", unsafe_allow_html=True)
        
        word_report = create_word_table_report(df_res, f"تقرير - {st.session_state['c_filename']}")
        st.download_button("📥 تحميل جدول الفروقات كـ Word", data=word_report, file_name=f"تقرير_{st.session_state['c_filename']}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

with tab2:
    st.markdown("<h3 style='text-align: right;'>📜 العمليات المؤرشفة في جوجل شيت</h3>", unsafe_allow_html=True)
    if st.button("🔄 تحديث وجلب البيانات الحالية من جوجل شيت"):
        try:
            archive_df = conn.read(worksheet="Sheet1", ttl=0)
            if not archive_df.empty: st.dataframe(archive_df, use_container_width=True, hide_index=True)
            else: st.info("الأرشيف فارغ حالياً.")
        except Exception as ex: st.error(f"لم نتمكن من قراءة الشيت، تأكد من الإعدادات: {ex}")
