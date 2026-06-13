import streamlit as st
import pandas as pd
import os
import requests
from datetime import datetime
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
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
# 2. إنشاء الاتصال بجوجل شيت ومفتاح iLovePDF
# -----------------------------------------------------------------------------
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception:
    pass

# -----------------------------------------------------------------------------
# 3. محركات المعالجة والتحويل (PDF -> Word -> Data)
# -----------------------------------------------------------------------------
def convert_pdf_to_word_in_memory(pdf_bytes, filename):
    """دالة الاتصال المباشر بخوادم iLovePDF بدون استخدام مكتبات خارجية معقدة"""
    public_key = st.secrets.get("ILOVEPDF_PUBLIC_KEY", "")
    if not public_key:
        raise Exception("مفتاح iLovePDF غير متوفر في إعدادات Secrets.")
    
    try:
        # 1. المصادقة للحصول على تصريح الدخول (Token)
        auth_res = requests.post("https://api.ilovepdf.com/v1/auth", data={"public_key": public_key}).json()
        token = auth_res.get("token")
        headers = {"Authorization": f"Bearer {token}"}
        
        # 2. بدء مهمة تحويل جديدة
        start_res = requests.get("https://api.ilovepdf.com/v1/start/pdfword", headers=headers).json()
        server = start_res.get("server")
        task_id = start_res.get("task")
        
        # 3. رفع الملف إلى سيرفرهم
        upload_url = f"https://{server}/v1/upload"
        files = {"file": (filename, pdf_bytes, "application/pdf")}
        upload_res = requests.post(upload_url, headers=headers, data={"task": task_id}, files=files).json()
        server_filename = upload_res.get("server_filename")
        
        # 4. إعطاء أمر التحويل
        process_url = f"https://{server}/v1/process"
        process_data = {
            "task": task_id,
            "tool": "pdfword",
            "files[0][server_filename]": server_filename,
            "files[0][filename]": filename
        }
        requests.post(process_url, headers=headers, data=process_data)
        
        # 5. تحميل الملف الجاهز إلى نظامك
        download_url = f"https://{server}/v1/download/{task_id}"
        download_res = requests.get(download_url, headers=headers)
        
        output_docx_name = filename.rsplit('.', 1)[0] + ".docx"
        return BytesIO(download_res.content), output_docx_name
        
    except Exception as e:
        raise Exception(f"فشل الاتصال المباشر بـ iLovePDF: {e}")

def extract_clean_records(file_obj):
    doc = Document(file_obj)
    records = {}
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            if not any(cells) or "المركز" in "".join(cells) or "الوكيل" in "".join(cells) or "اسم رب" in "".join(cells): continue
            name_idx = -1
            max_len = 0
            for i, c in enumerate(cells):
                if any('\u0600' <= char <= '\u06FF' for char in c) and not any(char.isdigit() for char in c):
                    if len(c) > max_len: max_len = len(c); name_idx = i
            if name_idx == -1: continue
            card_indices = [i for i, c in enumerate(cells) if c.isdigit() and len(c) >= 5]
            if not card_indices: continue
            card_num = cells[card_indices[1]] if len(card_indices) >= 2 else cells[card_indices[0]]
            seq = "-"
            for i in range(len(cells)-1, card_indices[-1], -1):
                if cells[i].isdigit(): seq = cells[i]; break
            digit_cells = [int(cells[i]) for i in range(name_idx) if cells[i].isdigit()]
            if len(digit_cells) >= 3: withheld, eligible, total = digit_cells[0], digit_cells[1], digit_cells[2]
            elif len(digit_cells) == 2: withheld, eligible, total = 0, digit_cells[0], digit_cells[1]
            else: continue
            records[card_num] = {"seq": seq, "name": cells[name_idx], "total": total, "eligible": eligible, "withheld": withheld}
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
    st.info("💡 يمكنك رفع ملفات الـ PDF أو Word مباشرة! سيقوم النظام بالمطابقة فوراً.")
    
    col1, col2 = st.columns(2)
    with col1: new_file = st.file_uploader("الملف الجديد (PDF أو Word)", type=['pdf', 'docx'], key="n_f")
    with col2: old_file = st.file_uploader("الملف القديم (PDF أو Word)", type=['pdf', 'docx'], key="o_f")

    if st.button("🚀 تشغيل الفحص والتحويل السحابي"):
        if old_file and new_file:
            with st.spinner('جاري التحويل والمطابقة السحابية... قد يستغرق الأمر ثوانٍ معدودة...'):
                try:
                    # معالجة الملف القديم
                    if old_file.name.endswith('.pdf'):
                        old_docx_io, old_docx_name = convert_pdf_to_word_in_memory(old_file.getvalue(), old_file.name)
                        st.session_state['download_old'] = {"data": old_docx_io.getvalue(), "name": old_docx_name}
                    else:
                        old_docx_io = old_file
                        
                    # معالجة الملف الجديد
                    if new_file.name.endswith('.pdf'):
                        new_docx_io, new_docx_name = convert_pdf_to_word_in_memory(new_file.getvalue(), new_file.name)
                        st.session_state['download_new'] = {"data": new_docx_io.getvalue(), "name": new_docx_name}
                    else:
                        new_docx_io = new_file

                    # المطابقة المباشرة
                    old_data = extract_clean_records(old_docx_io)
                    new_data = extract_clean_records(new_docx_io)
                    results, counters = compare_records(old_data, new_data)
                    
                    if results:
                        df_results = pd.DataFrame(results)[["التسلسل", "رقم البطاقة", "الاسم (سابقاً)", "الاسم (حالياً)", "الكلية (سابقاً)", "الكلية (حالياً)", "المستحقة (سابقاً)", "المستحقة (حالياً)", "المحجوبين (سابقاً)", "المحجوبين (حالياً)"]]
                        st.session_state['c_results'] = df_results
                        st.session_state['c_counters'] = counters
                        st.session_state['c_filename'] = new_file.name.rsplit('.', 1)[0]
                        st.success("✅ تمت عملية التحويل والمطابقة بنجاح!")
                    else: 
                        st.info("تطابق كامل بين الملفين، لا توجد فروقات.")
                except Exception as e:
                    st.error(f"حدث خطأ أثناء الاتصال السحابي للتحويل: {e}")
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
        
        # أزرار تحميل الملفات التي تم تحويلها من PDF (اختياري للوكيل)
        col_dl1, col_dl2, col_dl3 = st.columns(3)
        with col_dl1:
            if 'download_new' in st.session_state:
                st.download_button("📥 تحميل الملف الجديد كـ Word", data=st.session_state['download_new']["data"], file_name=st.session_state['download_new']["name"], mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        with col_dl2:
            if 'download_old' in st.session_state:
                st.download_button("📥 تحميل الملف القديم كـ Word", data=st.session_state['download_old']["data"], file_name=st.session_state['download_old']["name"], mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        with col_dl3:
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
