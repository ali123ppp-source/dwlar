import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import tempfile
import os
import glob
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from streamlit_gsheets import GSheetsConnection
import jwt
import requests
import time

# -----------------------------------------------------------------------------
# 1. إعدادات النظام والتصميم
# -----------------------------------------------------------------------------
st.set_page_config(page_title="نظام المقارنة السحابي المطور", layout="wide", initial_sidebar_state="expanded")

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
# 2. مفاتيح iLovePDF الثابتة والمحمية
# -----------------------------------------------------------------------------
ILOVEPDF_PUBLIC_KEY = "project_public_75f6c7c1da4d691f31ccc5fb88fc930e_E5Ka36ebed69263364c32d87466e6ee49be79"
ILOVEPDF_SECRET_KEY = "secret_key_4a929d6d2a01df607c76146f1962dbbd_BkBfG51d077fc768a97e36a6649fd2ee702bc"

# -----------------------------------------------------------------------------
# 3. محرك الاتصال المطور بـ iLovePDF API (ذكي ومقاوم للأخطاء)
# -----------------------------------------------------------------------------
def convert_with_ilovepdf_api(pdf_path, out_dir):
    try:
        pub_key_clean = ILOVEPDF_PUBLIC_KEY.strip()
        sec_key_clean = ILOVEPDF_SECRET_KEY.strip()

        if "ضع_هنا" in pub_key_clean or not pub_key_clean:
            raise Exception("لم يتم إدخال مفاتيح iLovePDF داخل كود السيرفر بعد.")

        token = None
        
        # [خطوة 1] محاولة جلب التوكن مباشرة من سيرفر iLovePDF لضمان تخطي مشاكل الوقت
        try:
            auth_resp = requests.post("https://api.ilovepdf.com/v1/auth", json={"public_key": pub_key_clean}, timeout=5)
            if auth_resp.status_code == 200:
                token = auth_resp.json().get("token")
        except Exception:
            pass
            
        # [خطوة 2] حل احتياطي: التوقيع الذاتي المحلي في حال تعذر الاتصال بسيرفر التوثيق (مع إضافة الميقات exp)
        if not token:
            current_time = int(time.time())
            payload = {
                "iss": pub_key_clean,
                "iat": current_time,
                "exp": current_time + 3600  # صلاحية التوكن ساعة واحدة
            }
            token = jwt.encode(payload, sec_key_clean, algorithm="HS256")
            if isinstance(token, bytes):
                token = token.decode('utf-8')

        headers = {"Authorization": f"Bearer {token}"}
        
        # [خطوة 3] فتح اتصال لبدء مهمة تحويل PDF to Word
        start_resp = requests.get("https://api.ilovepdf.com/v1/start/pdfword", headers=headers).json()
        
        # إذا لم يقبل السيرفر الاتصال، نقرأ السبب الحقيقي بدقة ونعرضه للمستخدم
        if "server" not in start_resp:
            error_details = start_resp.get("error", {})
            msg = error_details.get("message", "التوكن مرفوض أو الحساب انتهى رصيده المتاح.")
            code = error_details.get("code", "غير معروف")
            raise Exception(f"استجابة iLovePDF: {msg} (كود الخطأ: {code})")
            
        server, task = start_resp["server"], start_resp["task"]
        
        # [خطوة 4] رفع ملف الـ PDF
        with open(pdf_path, 'rb') as f:
            upload_resp = requests.post(f"https://{server}/v1/upload", headers=headers, data={'task': task}, files={'file': f}).json()
            
        # [خطوة 5] معالجة التحويل الذكي
        process_data = {
            "task": task, "tool": "pdfword",
            "files": [{"server_filename": upload_resp["server_filename"], "filename": "converted.docx"}]
        }
        requests.post(f"https://{server}/v1/process", headers=headers, json=process_data)
        
        # [خطوة 6] تنزيل الملف النهائي كـ Word
        download_resp = requests.get(f"https://{server}/v1/download/{task}", headers=headers)
        docx_path = os.path.join(out_dir, "converted.docx")
        
        with open(docx_path, 'wb') as f:
            f.write(download_resp.content)
            
        return docx_path
    except Exception as e:
        raise Exception(f"فشل الاتصال الآمن بخوادم التحويل: {e}")

# -----------------------------------------------------------------------------
# 4. محركات المعالجة الذكية للجداول والبيانات
# -----------------------------------------------------------------------------
def parse_row(cells):
    clean_cells = [str(c).strip().replace('\n', ' ') if c is not None else "" for c in cells]
    joined = "".join(clean_cells)
    
    if not any(clean_cells) or "المركز" in joined or "الوكيل" in joined or "اسم" in joined or "زكرمل" in joined: 
        return {}
        
    name_idx, max_len = -1, 0
    for i, c in enumerate(clean_cells):
        if any('\u0600' <= char <= '\u06FF' for char in c) and not any(char.isdigit() for char in c):
            if len(c) > max_len: max_len, name_idx = len(c), i
                
    if name_idx == -1: return {}
    
    final_name = clean_cells[name_idx]

    card_cands = [c for c in clean_cells if c.isdigit() and len(c) >= 5]
    card_num = card_cands[0] if card_cands else "-"
    if card_num == "-": return {}

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

def extract_clean_records(file_obj, is_pdf):
    records = {}
    
    if is_pdf:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(file_obj.read())
            pdf_path = tmp_pdf.name
            
        out_dir = tempfile.mkdtemp()
        try:
            docx_path = convert_with_ilovepdf_api(pdf_path, out_dir)
            doc = Document(docx_path)
        finally:
            if os.path.exists(pdf_path): os.remove(pdf_path)
            for f in glob.glob(os.path.join(out_dir, "*")): os.remove(f)
            if os.path.exists(out_dir): os.rmdir(out_dir)
    else:
        doc = Document(file_obj)
        
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            records.update(parse_row(cells))
            
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
# 5. واجهة التطبيق الرئيسية المخصصة للهواتف
# -----------------------------------------------------------------------------
tab1, tab2 = st.tabs(["🔎 إجراء مقارنة ذكية", "📜 الأرشيف التاريخي"])

with tab1:
    st.markdown("<h3 style='text-align: right;'>لوحة المطابقة التلقائية الذكية</h3>", unsafe_allow_html=True)
    st.info("📱 مخصص للهواتف: ارفع الملفات مباشرة (PDF أو Word)، وسيقوم النظام بكل شيء تلقائياً.")
    
    col1, col2 = st.columns(2)
    with col1: new_file = st.file_uploader("الملف الجديد (PDF أو Word)", type=['pdf', 'docx'], key="n_f")
    with col2: old_file = st.file_uploader("الملف القديم (PDF أو Word)", type=['pdf', 'docx'], key="o_f")

    if st.button("🚀 تشغيل الفحص والمطابقة"):
        if old_file and new_file:
            with st.spinner('جاري تحويل الملفات وقراءة الجداول سحابياً... يرجى الانتظار ثوانٍ'):
                try:
                    old_is_pdf = old_file.name.lower().endswith('.pdf')
                    old_data = extract_clean_records(old_file, is_pdf=old_is_pdf)
                    
                    new_is_pdf = new_file.name.lower().endswith('.pdf')
                    new_data = extract_clean_records(new_file, is_pdf=new_is_pdf)

                    results, counters = compare_records(old_data, new_data)
                    
                    if results:
                        st.session_state['c_results'] = pd.DataFrame(results)
                        st.session_state['c_counters'] = counters
                        st.session_state['c_filename'] = new_file.name.rsplit('.', 1)[0]
                        st.success("✅ تمت المعالجة والمطابقة بنجاح تام!")
                    else: st.info("تطابق كامل ومثالي بين الملفين.")
                except Exception as e: st.error(str(e))
        else: st.warning("يرجى رفع الملفين أولاً للبدء.")

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
