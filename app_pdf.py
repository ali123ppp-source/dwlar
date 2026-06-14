import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from streamlit_gsheets import GSheetsConnection
import datetime

# -----------------------------------------------------------------------------
# 1. إعدادات النظام والتصميم الديناميكي المطور للموبايل
# -----------------------------------------------------------------------------
st.set_page_config(page_title="نظام التدقيق الشامل لبيانات الوكلاء", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #2E7D32; color: white; width: 100%; font-weight: bold; border-radius: 8px; border: none; height: 48px; font-size: 16px; }
    .report-box { background-color: var(--secondary-background-color); color: var(--text-color); padding: 12px; border-radius: 8px; border-right: 5px solid #2E7D32; text-align: right; margin-bottom: 10px; font-weight: bold; }
    .instruction-box { background-color: #FFF3E0; color: #E65100; padding: 15px; border-radius: 8px; border-right: 5px solid #EF6C00; text-align: right; margin-bottom: 20px; font-size: 15px; line-height: 1.6; }
    .ilove-link { color: #E65100 !important; font-weight: bold; text-decoration: underline !important; background-color: #FFE0B2; padding: 3px 8px; border-radius: 4px; }
    
    /* تنسيقات كروت العوائل الملونة للمفتشين والوكلاء */
    .month-card { background-color: var(--background-color); padding: 12px; border-radius: 6px; border: 1px solid #e0e0e0; margin-bottom: 8px; text-align: right; }
    .badge-total { background-color: #0288D1; color: white; padding: 3px 10px; border-radius: 4px; font-weight: bold; display: inline-block; }
    .badge-eligible { background-color: #388E3C; color: white; padding: 3px 10px; border-radius: 4px; font-weight: bold; display: inline-block; }
    .badge-withheld { background-color: #D32F2F; color: white; padding: 3px 10px; border-radius: 4px; font-weight: bold; display: inline-block; }
    </style>
""", unsafe_allow_html=True)

# الاتصال بقوكول شيت للأرشفة التاريخية
conn = None
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception:
    st.error("⚠️ فشل الاتصال بقاعدة البيانات السحابية (Google Sheets)، تحقق من الإعدادات.")

# -----------------------------------------------------------------------------
# 2. محرك الذاكرة المستمرة للمتصفح (حماية ضد عمل Refresh)
# -----------------------------------------------------------------------------
@st.cache_resource
def get_persistent_memory():
    return {"df": None, "counters": None, "filename": None}

browser_memory = get_persistent_memory()

# استعادة تلقائية فورية للبيانات عند تحديث الصفحة
if 'c_results' not in st.session_state and browser_memory["df"] is not None:
    st.session_state['c_results'] = browser_memory["df"]
    st.session_state['c_counters'] = browser_memory["counters"]
    st.session_state['c_filename'] = browser_memory["filename"]
    st.toast("🔄 تم استعادة سجلات الوكيل بالكامل تلقائياً من ذاكرة المتصفح!", icon="⚡")

# -----------------------------------------------------------------------------
# 3. محرك تحليل السجلات واستخراج الأسماء الرباعية والتسلسل
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
    
    final_quad_name = clean_cells[name_idx] # الاسم الرباعي

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
        
    return {card_num: {"seq": seq, "name": final_quad_name, "total": total, "eligible": eligible, "withheld": withheld}}

def extract_clean_records(file_obj):
    records = {}
    doc = Document(file_obj)
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            records.update(parse_row(cells))
    return records

def compare_records(old_data, new_data):
    results, counters = [], {"added_fam": 0, "deleted_fam": 0, "unchanged_fam": 0, "modified_fam": 0, "net_total": 0, "net_eligible": 0, "net_withheld": 0}
    all_cards = set(old_data.keys()).union(set(new_data.keys()))
    
    for card in all_cards:
        if card in old_data and card in new_data:
            old_v, new_v = old_data[card], new_data[card]
            if old_v["total"] != new_v["total"] or old_v["eligible"] != new_v["eligible"] or old_v["withheld"] != new_v["withheld"]:
                counters["modified_fam"] += 1
                counters["net_total"] += (new_v["total"] - old_v["total"])
                counters["net_eligible"] += (new_v["eligible"] - old_v["eligible"])
                counters["net_withheld"] += (new_v["withheld"] - old_v["withheld"])
                results.append({"التسلسل": new_v["seq"], "رقم البطاقة": card, "الاسم الرباعي (سابقاً)": old_v["name"], "الاسم الرباعي (حالياً)": new_v["name"], "الكلية (سابقاً)": old_v["total"], "الكلية (حالياً)": new_v["total"], "المستحقة (سابقاً)": old_v["eligible"], "المستحقة (حالياً)": new_v["eligible"], "المحجوبين (سابقاً)": old_v["withheld"], "المحجوبين (حالياً)": new_v["withheld"], "الحالة": "🟡 قيد معدل الحصص"})
            else:
                counters["unchanged_fam"] += 1
                results.append({"التسلسل": new_v["seq"], "رقم البطاقة": card, "الاسم الرباعي (سابقاً)": old_v["name"], "الاسم الرباعي (حالياً)": new_v["name"], "الكلية (سابقاً)": old_v["total"], "الكلية (حالياً)": new_v["total"], "المستحقة (سابقاً)": old_v["eligible"], "المستحقة (حالياً)": new_v["eligible"], "المحجوبين (سابقاً)": old_v["withheld"], "المحجوبين (حالياً)": new_v["withheld"], "الحالة": "✅ متطابق (بدون تغيير)"})
        elif card in old_data and card not in new_data:
            old_v = old_data[card]
            counters["deleted_fam"] += 1; counters["net_total"] -= old_v["total"]; counters["net_eligible"] -= old_v["eligible"]; counters["net_withheld"] -= old_v["withheld"]
            results.append({"التسلسل": old_v["seq"], "رقم البطاقة": card, "الاسم الرباعي (سابقاً)": old_v["name"], "الاسم الرباعي (حالياً)": "❌ (محذوف / منقول)", "الكلية (سابقاً)": old_v["total"], "الكلية (حالياً)": 0, "المستحقة (سابقاً)": old_v["eligible"], "المستحقة (حالياً)": 0, "المحجوبين (سابقاً)": old_v["withheld"], "المحجوبين (حالياً)": 0, "الحالة": "🔴 محذوف من الوجبة"})
        elif card not in old_data and card in new_data:
            new_v = new_data[card]
            counters["added_fam"] += 1; counters["net_total"] += new_v["total"]; counters["net_eligible"] += new_v["eligible"]; counters["net_withheld"] += new_v["withheld"]
            results.append({"التسلسل": new_v["seq"], "رقم البطاقة": card, "الاسم الرباعي (سابقاً)": "✨ (مضاف حديثاً)", "الاسم الرباعي (حالياً)": new_v["name"], "الكلية (سابقاً)": 0, "الكلية (حالياً)": new_v["total"], "المستحقة (سابقاً)": 0, "المستحقة (حالياً)": new_v["eligible"], "المحجوبين (سابقاً)": 0, "المحجوبين (حالياً)": new_v["withheld"], "الحالة": "🟢 قيد مضاف جديد"})
            
    return results, counters

# -----------------------------------------------------------------------------
# 4. محرك تقرير الـ Word للمتغير الحالي فقط (بدون ذكر حقول السابق) مع التسلسل
# -----------------------------------------------------------------------------
def create_current_only_word_report(df, file_title):
    doc = Document()
    heading = doc.add_heading(f"كشف البيانات الحالية للوكيل: {file_title}", level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    report_rows = []
    for _, row in df.iterrows():
        quad_name = row["الاسم الرباعي (حالياً)"]
        if "❌" in quad_name:
            quad_name = row["الاسم الرباعي (سابقاً)"]
            
        report_rows.append({
            "التسلسل": row["التسلسل"],
            "رقم البطاقة": row["رقم البطاقة"],
            "الاسم الرباعي الحالي": quad_name,
            "الكلية الحالية": row["الكلية (حالياً)"],
            "المستحقة الحالية": row["المستحقة (حالياً)"],
            "المحجوبين حالياً": row["المحجوبين (حالياً)"],
            "حالة القيد": row["الحالة"]
        })
        
    report_df = pd.DataFrame(report_rows)
    cols = list(report_df.columns)[::-1]
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = 'Table Grid'
    
    for i, col in enumerate(cols):
        table.rows[0].cells[i].text = str(col)
        
    for _, r_row in report_df.iterrows():
        row_cells = table.add_row().cells
        for i, col in enumerate(cols):
            row_cells[i].text = str(r_row[col])
            
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# 5. واجهة التطبيق التفاعلية والتحكم
# -----------------------------------------------------------------------------
tab1, tab2 = st.tabs(["🔎 تدقيق بيانات الوكيل الكاملة", "📜 الأرشيف التاريخي"])

with tab1:
    st.markdown("<h3 style='text-align: right;'>المنظومة الرقمية لعرض وجرد بيانات الوكيل بالكامل</h3>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1: new_file = st.file_uploader("الملف الجديد (الشهر الحالي)", type=['docx'], key="n_f")
    with col2: old_file = st.file_uploader("الملف القديم (الشهر السابق)", type=['docx'], key="o_f")

    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn1:
        run_match = st.button("🚀 جلب وعرض كافّة بيانات الوكيل")
    with col_btn2:
        clear_mem = st.button("🗑️ تفريغ الذاكرة لبدء وكيل جديد")
        if clear_mem:
            browser_memory["df"] = None
            browser_memory["counters"] = None
            browser_memory["filename"] = None
            for k in ['c_results', 'c_counters', 'c_filename']:
                if k in st.session_state: del st.session_state[k]
            st.rerun()

    if run_match:
        if old_file and new_file:
            with st.spinner('جاري قراءة الملفات وجرد الأسماء الرباعية بالتسلسلات...'):
                try:
                    old_data = extract_clean_records(old_file)
                    new_data = extract_clean_records(new_file)
                    results, counters = compare_records(old_data, new_data)
                    
                    if results:
                        st.session_state['c_results'] = pd.DataFrame(results).sort_values(by="التسلسل")
                        st.session_state['c_counters'] = counters
                        st.session_state['c_filename'] = new_file.name.rsplit('.', 1)[0]
                        
                        browser_memory["df"] = st.session_state['c_results']
                        browser_memory["counters"] = st.session_state['c_counters']
                        browser_memory["filename"] = st.session_state['c_filename']
                        
                        st.success("✅ تم جلب السجلات وتأمينها بالمتصفح ضد التحديث والمسح تلقائياً.")
                    else:
                        st.info("لم يتم العثور على سجلات صالحة.")
                except Exception as e:
                    st.error(f"حدث خطأ أثناء القراءة: {str(e)}")
        else:
            st.warning("يرجى تزويد النظام بملفات الـ Word لوكيلك أولاً.")

    # -------------------------------------------------------------------------
    # 6. لوحة التدقيق والفرز الإحصائي الشامل بالأسماء والتسلسلات
    # -------------------------------------------------------------------------
    if 'c_results' in st.session_state:
        df_res = st.session_state['c_results']
        cnt = st.session_state['c_counters']
        
        # 🔐 حماية إضافية ضد الـ KeyError في حال قراءة كاش قديم باستخدام .get()
        net_total = cnt.get('net_total', 0)
        added_fam = cnt.get('added_fam', 0)
        deleted_fam = cnt.get('deleted_fam', 0)
        unchanged_fam = cnt.get('unchanged_fam', 0)
        modified_fam = cnt.get('modified_fam', 0)
        
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f"<div class='report-box'>🔹 إجمالي سجلات الوكيل: {len(df_res)}</div>", unsafe_allow_html=True)
        with c2: st.markdown(f"<div class='report-box'>🔹 صافي المستحقة: {net_total:+d}</div>", unsafe_allow_html=True)
        with c3: st.markdown(f"<div class='report-box'>🔹 مضاف: {added_fam} | محذوف: {deleted_fam} | متطابق: {unchanged_fam}</div>", unsafe_allow_html=True)
        
        st.markdown("### 🔐 ترحيل سحابي دائم")
        if st.button("💾 حفظ بيانات هذا الوكيل بالكامل في Google Sheets"):
            if conn is not None:
                with st.spinner('جاري الإرسال للسحابة...'):
                    try:
                        try: existing_df = pd.DataFrame(conn.read(worksheet="Sheet1", ttl=0))
                        except Exception: existing_df = pd.DataFrame()
                        
                        df_to_save = df_res.copy()
                        df_to_save["وقت الحفظ"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                        df_to_save["اسم الوكيل/الملف"] = st.session_state['c_filename']
                        
                        final_archive = pd.concat([existing_df, df_to_save], ignore_index=True)
                        conn.update(worksheet="Sheet1", data=final_archive)
                        st.success("🌟 تم حفظ الكشف كاملاً داخل ملف قوقل شيت بنجاح.")
                    except Exception as ex:
                        st.error(f"فشل الاتصال: {str(ex)}")
        
        st.markdown("---")
        st.markdown("<h4 style='text-align: right;'>🎯 لوحة البحث والفرز الذكي بأسماء الوكيل الحالية</h4>", unsafe_allow_html=True)
        
        filter_options = [
            f"📋 عرض كافّة سجلات الوكيل بالكامل ({len(df_res)})",
            f"✅ العوائل المستقرة والمتطابقة فقط ({unchanged_fam})",
            f"🟡 العوائل التي تغيرت حصصها التموينية ({modified_fam})",
            f"✨ العوائل المضافة حديثاً للوكيل ({added_fam})",
            f"❌ العوائل المحذوفة أو المنقولة من الوكيل ({deleted_fam})"
        ]
        selected_filter = st.selectbox("🎯 اختر الفئة المراد تدقيقها للوكيل:", filter_options)
        
        search_q = st.text_input("👤 اكتب الاسم الرباعي للمواطن أو رقم البطاقة للبحث الفوري:", "")
        
        if "✅" in selected_filter:
            filtered_df = df_res[df_res["الحالة"] == "✅ متطابق (بدون تغيير)"]
        elif "🟡" in selected_filter:
            filtered_df = df_res[df_res["الحالة"] == "🟡 قيد معدل الحصص"]
        elif "✨" in selected_filter:
            filtered_df = df_res[df_res["الحالة"] == "🟢 قيد مضاف جديد"]
        elif "❌" in selected_filter:
            filtered_df = df_res[df_res["الحالة"] == "🔴 محذوف من الوجبة"]
        else:
            filtered_df = df_res

        if search_q:
            filtered_df = filtered_df[
                filtered_df["الاسم الرباعي (حالياً)"].str.contains(search_q, na=False) | 
                filtered_df["الاسم الرباعي (سابقاً)"].str.contains(search_q, na=False) |
                filtered_df["رقم البطاقة"].str.contains(search_q, na=False)
            ]
            
        st.markdown(f"<p style='text-align: right; color: gray;'>عدد النتائج الحالية: {len(filtered_df)} مواطن</p>", unsafe_allow_html=True)
        
        for idx, row in filtered_df.iterrows():
            display_name = row["الاسم الرباعي (سابقاً)"] if "❌" in row["الاسم الرباعي (حالياً)"] else row["الاسم الرباعي (حالياً)"]
            status_badge = row["الحالة"]
            
            box_title = f"ت تسلسل: [{row['التسلسل']}] | {status_badge} | {display_name} (رقم البطاقة: {row['رقم البطاقة']})"
            
            with st.expander(box_title):
                col_old, col_new = st.columns(2)
                with col_old:
                    st.markdown("<div class='month-card'><b>📅 السجلات التموينية السابقة:</b><br><br>"
                                f"▪️ إجمالي الكلية: {row['الكلية (سابقاً)']}<br>"
                                f"▪️ المستحقة الفعلية: {row['المستحقة (سابقاً)']}<br>"
                                f"▪️ أفراد الحجب: {row['المحجوبين (سابقاً)']}"
                                "</div>", unsafe_allow_html=True)
                with col_new:
                    st.markdown("<div class='month-card'><b>🌟 السجلات التموينية الحالية (النهائية):</b><br><br>"
                                f"▪️ إجمالي الكلية: <span class='badge-total'>{row['الكلية (حالياً)']}</span><br><br>"
                                f"▪️ المستحقة الفعلية: <span class='badge-eligible'>{row['المستحقة (حالياً)']}</span><br><br>"
                                f"▪️ أفراد الحجب: <span class='badge-withheld'>{row['المحجوبين (حالياً)']}</span>"
                                "</div>", unsafe_allow_html=True)
        
        st.markdown("---")
        word_data = create_current_only_word_report(df_res, st.session_state['c_filename'])
        st.download_button(
            label="📥 تحميل كشف الحصص الحالية للوكيل (Word)", 
            data=word_data, 
            file_name=f"كشف_بيانات_الوكيل_الحالية_{st.session_state['c_filename']}.docx"
        )

with tab2:
    if st.button("🔄 جلب وتحديث الأرشيف المركزي السحابي"):
        try:
            if conn is not None:
                st.dataframe(conn.read(worksheet="Sheet1", ttl=0), use_container_width=True, hide_index=True)
            else:
                st.error("الاتصال بقاعدة البيانات السحابية مقطوع.")
        except Exception as ex: 
            st.error(ex)
