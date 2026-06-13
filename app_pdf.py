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
# 2. محرك الاتصال بـ iLovePDF API (المباشر والمضمون)
# -----------------------------------------------------------------------------
def convert_with_ilovepdf_api(pdf_path, public_key, secret_key, out_dir):
    try:
        # تنظيف المفاتيح من أي مسافات أو أسطر فارغة منسوخة بالخطأ
        pub_key_clean = public_key.strip()
        sec_key_clean = secret_key.strip()

        # 1. إنشاء توثيق مشفر (JWT) للاتصال الآمن
        payload = {"jti": str(time.time()), "iss": pub_key_clean, "iat": int(time.time())}
        token = jwt.encode(payload, sec_key_clean, algorithm="HS256")
        headers = {"Authorization": f"Bearer {token}"}
        
        # 2. فتح خط اتصال لمهمة (PDF to Word)
        start_resp = requests.get("https://api.ilovepdf.com/v1/start/pdfword", headers=headers).json()
        
        # التأكد من قبول السيرفر للمفاتيح
        if "server" not in start_resp:
            error_msg = start_resp.get("error", "مفاتيح غير صالحة")
            raise Exception(f"رفض السيرفر الاتصال! يرجى التأكد من المفاتيح. (رسالة السيرفر: {error_msg})")
            
        server, task = start_resp["server"], start_resp["task"]
        
        # 3. رفع الملف بسرعة
        with open(pdf_path, 'rb') as f:
            upload_resp = requests.post(f"https://{server}/v1/upload", headers=headers, data={'task': task}, files={'file': f}).json()
            
        # 4. معالجة التحويل
        process_data = {
            "task": task, "tool": "pdfword",
            "files": [{"server_filename": upload_resp["server_filename"], "filename": "converted.docx"}]
        }
        requests.post(f"https://{server}/v1/process", headers=headers, json=process_data)
        
        # 5. تنزيل الملف كـ Word
        download_resp = requests.get(f"https://{server}/v1/download/{task}", headers=headers)
        docx_path = os.path.join(out_dir, "converted.docx")
        
        with open(docx_path, 'wb') as f:
            f.write(download_resp.content)
            
        return docx_path
    except Exception as e:
        raise Exception(f"فشل الاتصال بخوادم iLovePDF: {e}")

# -----------------------------------------------------------------------------
# 3. محركات المعالجة الذكية
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

def extract_clean_records(file_obj, is_pdf, pub_key=None, sec_key=None):
    records = {}
    
    if is_pdf:
        if not pub_key or not sec_key:
            raise ValueError("يرجى إدخال مفاتيح iLovePDF في القائمة الجانبية أولاً لتحويل ملفات الـ PDF.")
            
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
            tmp_pdf.write(file_obj.read())
            pdf_path = tmp_pdf.name
            
        out_dir = tempfile.mkdtemp()
        try:
            # إرسال الملف للتحويل
            docx_path = convert_with_ilovepdf_api(pdf_path, pub_key, sec_key, out_dir)
            doc = Document(docx_path)
        finally:
            # التنظيف وحذف الملفات المؤقتة
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
# 4. واجهة التطبيق الرئيسية 
# -----------------------------------------------------------------------------
with st.sidebar:
    st.header("⚙️ إعدادات محرك iLovePDF")
    st.info("لتفعيل التحويل السريع للـ PDF، يرجى إدخال المفاتيح من حسابك المطور.")
    i_pub_key = st.text_input("Public Key:", type="password")
    i_sec_key = st.text_input("Secret Key:", type="password")
    st.markdown("---")

tab1, tab2 = st.tabs(["🔎 إجراء مقارنة ذكية", "📜 الأرشيف التاريخي"])

with tab1:
    st.markdown("<h3 style='text-align: right;'>لوحة المطابقة (بمحرك iLovePDF السريع)</h3>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1: new_file = st.file_uploader("الملف الجديد", type=['pdf', 'docx'], key="n_f")
    with col2: old_file = st.file_uploader("الملف القديم", type=['pdf', 'docx'], key="o_f")

    if st.button("🚀 تشغيل الفحص والمطابقة"):
        if old_file and new_file:
            with st.spinner('جاري الاتصال بخوادم iLovePDF ومطابقة البيانات...'):
                try:
                    old_is_pdf = old_file.name.lower().endswith('.pdf')
                    old_data = extract_clean_records(old_file, is_pdf=old_is_pdf, pub_key=i_pub_key, sec_key=i_sec_key)
                    
                    new_is_pdf = new_file.name.lower().endswith('.pdf')
                    new_data = extract_clean_records(new_file, is_pdf=new_is_pdf, pub_key=i_pub_key, sec_key=i_sec_key)

                    results, counters = compare_records(old_data, new_data)
                    
                    if results:
                        st.session_state['c_results'] = pd.DataFrame(results)
                        st.session_state['c_counters'] = counters
                        st.session_state['c_filename'] = new_file.name.rsplit('.', 1)[0]
                        st.success("✅ تمت المعالجة بنجاح عبر خوادم iLovePDF!")
                    else: st.info("تطابق كامل بين الملفين.")
                except Exception as e: st.error(str(e))
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
