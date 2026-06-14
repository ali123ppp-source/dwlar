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
st.set_page_config(page_title="نظام التدقيق الرقمي والمطابقة الذكية", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #2E7D32; color: white; width: 100%; font-weight: bold; border-radius: 8px; border: none; height: 48px; font-size: 16px; }
    .report-box { background-color: var(--secondary-background-color); color: var(--text-color); padding: 12px; border-radius: 8px; border-right: 5px solid #2E7D32; text-align: right; margin-bottom: 10px; font-weight: bold; }
    .instruction-box { background-color: #FFF3E0; color: #E65100; padding: 15px; border-radius: 8px; border-right: 5px solid #EF6C00; text-align: right; margin-bottom: 20px; font-size: 15px; line-height: 1.6; }
    .ilove-link { color: #E65100 !important; font-weight: bold; text-decoration: underline !important; background-color: #FFE0B2; padding: 3px 8px; border-radius: 4px; }
    
    /* تنسيقات كروت العوائل الملونة للمفتشين */
    .month-card { background-color: var(--background-color); padding: 12px; border-radius: 6px; border: 1px solid #e0e0e0; margin-bottom: 8px; text-align: right; }
    .badge-total { background-color: #0288D1; color: white; padding: 3px 10px; border-radius: 4px; font-weight: bold; display: inline-block; }
    .badge-eligible { background-color: #388E3C; color: white; padding: 3px 10px; border-radius: 4px; font-weight: bold; display: inline-block; }
    .badge-withheld { background-color: #D32F2F; color: white; padding: 3px 10px; border-radius: 4px; font-weight: bold; display: inline-block; }
    </style>
""", unsafe_allow_html=True)

# الاتصال بقوقل شيت وتأمين التوصيل للأرشيف التاريخي
conn = None
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception:
    st.error("⚠️ فشل الاتصال بقاعدة البيانات السحابية (Google Sheets)، تحقق من الإعدادات.")

# -----------------------------------------------------------------------------
# 2. محرك الذاكرة المستمرة (لحفظ البيانات من المسح عند تحديث الصفحة)
# -----------------------------------------------------------------------------
@st.cache_resource
def get_persistent_memory():
    return {"df": None, "counters": None, "filename": None}

browser_memory = get_persistent_memory()

# استعادة تلقائية فورية للبيانات في حال حدوث Refresh للمتصفح
if 'c_results' not in st.session_state and browser_memory["df"] is not None:
    st.session_state['c_results'] = browser_memory["df"]
    st.session_state['c_counters'] = browser_memory["counters"]
    st.session_state['c_filename'] = browser_memory["filename"]
    st.toast("🔄 تم استعادة البيانات والاسم الرباعي تلقائياً من ذاكرة المتصفح!", icon="⚡")

# -----------------------------------------------------------------------------
# 3. محركات المعالجة الذكية واستخراج الاسم الرباعي وحقول الحصص
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
    
    final_quad_name = clean_cells[name_idx] # الاسم الرباعي المستخلص

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
    results, counters = [], {"added_fam": 0, "deleted_fam": 0, "net_total": 0, "net_eligible": 0, "net_withheld": 0}
    all_cards = set(old_data.keys()).union(set(new_data.keys()))
    
    for card in all_cards:
        if card in old_data and card in new_data:
            old_v, new_v = old_data[card], new_data[card]
            if old_v["total"] != new_v["total"] or old_v["eligible"] != new_v["eligible"] or old_v["withheld"] != new_v["withheld"]:
                counters["net_total"] += (new_v["total"] - old_v["total"])
                counters["net_eligible"] += (new_v["eligible"] - old_v["eligible"])
                counters["net_withheld"] += (new_v["withheld"] - old_v["withheld"])
                results.append({"التسلسل": new_v["seq"], "رقم البطاقة": card, "الاسم الرباعي (سابقاً)": old_v["name"], "الاسم الرباعي (حالياً)": new_v["name"], "الكلية (سابقاً)": old_v["total"], "الكلية (حالياً)": new_v["total"], "المستحقة (سابقاً)": old_v["eligible"], "المستحقة (حالياً)": new_v["eligible"], "المحجوبين (سابقاً)": old_v["withheld"], "المحجوبين (حالياً)": new_v["withheld"]})
        elif card in old_data and card not in new_data:
            old_v = old_data[card]
            counters["deleted_fam"] += 1; counters["net_total"] -= old_v["total"]; counters["net_eligible"] -= old_v["eligible"]; counters["net_withheld"] -= old_v["withheld"]
            results.append({"التسلسل": old_v["seq"], "رقم البطاقة": card, "الاسم الرباعي (سابقاً)": old_v["name"], "الاسم الرباعي (حالياً)": "❌ (محذوف / منقول)", "الكلية (سابقاً)": old_v["total"], "الكلية (حالياً)": 0, "المستحقة (سابقاً)": old_v["eligible"], "المستحقة (حالياً)": 0, "المحجوبين (سابقاً)": old_v["withheld"], "المحجوبين (حالياً)": 0})
        elif card not in old_data and card in new_data:
            new_v = new_data[card]
            counters["added_fam"] += 1; counters["net_total"] += new_v["total"]; counters["net_eligible"] += new_v["eligible"]; counters["net_withheld"] += new_v["withheld"]
            results.append({"التسلسل": new_v["seq"], "رقم البطاقة": card, "الاسم الرباعي (سابقاً)": "✨ (مضاف حديثاً)", "الاسم الرباعي (حالياً)": new_v["name"], "الكلية (سابقاً)": 0, "الكلية (حالياً)": new_v["total"], "المستحقة (سابقاً)": 0, "المستحقة (حالياً)": new_v["eligible"], "المحجوبين (سابقاً)": 0, "المحجوبين (حالياً)": new_v["withheld"]})
    return results, counters

# -----------------------------------------------------------------------------
# 4. محرك التقارير المخصص للمتغير الحالي فقط (بدون ذكر حقول السابق)
# -----------------------------------------------------------------------------
def create_current_only_word_report(df, file_title):
    doc = Document()
    heading = doc.add_heading(f"تقرير المتغيرات الحالية للملف: {file_title}", level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # بناء هيكلية جدول المتغير الحالي فقط
    report_rows = []
    for _, row in df.iterrows():
        status = "تعديل حصة"
        quad_name = row["الاسم الرباعي (حالياً)"]
        
        if "✨" in row["الاسم الرباعي (سابقاً)"]:
            status = "مضاف حديثاً"
        elif "❌" in row["الاسم الرباعي (حالياً)"]:
            status = "محذوف / منقول"
            quad_name = row["الاسم الرباعي (سابقاً)"] # استخدام الاسم الرباعي المتوفر قبل الحذف
            
        report_rows.append({
            "التسلسل": row["التسلسل"],
            "رقم البطاقة": row["رقم البطاقة"],
            "الاسم الرباعي": quad_name,
            "الحصة الكلية الحالية": row["الكلية (حالياً)"],
            "الحصة المستحقة الحالية": row["المستحقة (حالياً)"],
            "المحجوبين حالياً": row["المحجوبين (حالياً)"],
            "حالة القيد": status
        })
        
    report_df = pd.DataFrame(report_rows)
    
    # عكس اتجاه الأعمدة لتظهر بشكل صحيح متناسق مع لغة الـ Word اليمينية RTL
    cols = list(report_df.columns)[::-1]
    table = doc.add_table(rows=1, cols=len(cols))
    table.style = 'Table Grid'
    
    # كتابة الهيدر
    for i, col in enumerate(cols):
        table.rows[0].cells[i].text = str(col)
        
    # تعبئة الصفوف بالمتغير الحالي فقط
    for _, r_row in report_df.iterrows():
        row_cells = table.add_row().cells
        for i, col in enumerate(cols):
            row_cells[i].text = str(r_row[col])
            
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# 5. واجهة التطبيق التفاعلية
# -----------------------------------------------------------------------------
tab1, tab2 = st.tabs(["🔎 مطابقة ملفات Word المتطورة", "📜 الأرشيف التاريخي والسحابي"])

with tab1:
    st.markdown("<h3 style='text-align: right;'>منظومة الفرز والتدقيق الذكي بالأسماء الرباعية</h3>", unsafe_allow_html=True)
    
    st.markdown("""
    <div class='instruction-box'>
        💡 <b>ملاحظات الاستخدام والتدقيق الفوري:</b><br>
        • نظام البحث بالأسماء يدعم كتابة أي جزء من <b>الاسم الرباعي</b> للمواطن.<br>
        • تقرير الـ Word المستخرج مصمم ليعرض <b>الحصص والمتغيرات الحالية فقط</b> لحماية السرية واختصار البيانات السابقة غير الضرورية.
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1: new_file = st.file_uploader("الملف الجديد (الشهر الحالي)", type=['docx'], key="n_f")
    with col2: old_file = st.file_uploader("الملف القديم (الشهر السابق)", type=['docx'], key="o_f")

    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn1:
        run_match = st.button("🚀 بدء تحليل ومطابقة الفروقات الحالية")
    with col_btn2:
        clear_mem = st.button("🗑️ تفريغ المتصفح والبدء بملف جديد")
        if clear_mem:
            browser_memory["df"] = None
            browser_memory["counters"] = None
            browser_memory["filename"] = None
            for k in ['c_results', 'c_counters', 'c_filename']:
                if k in st.session_state: del st.session_state[k]
            st.rerun()

    if run_match:
        if old_file and new_file:
            with st.spinner('جاري فحص الأسماء الرباعية واحتساب فروقات الحصص...'):
                try:
                    old_data = extract_clean_records(old_file)
                    new_data = extract_clean_records(new_file)
                    results, counters = compare_records(old_data, new_data)
                    
                    if results:
                        st.session_state['c_results'] = pd.DataFrame(results)
                        st.session_state['c_counters'] = counters
                        st.session_state['c_filename'] = new_file.name.rsplit('.', 1)[0]
                        
                        browser_memory["df"] = st.session_state['c_results']
                        browser_memory["counters"] = st.session_state['c_counters']
                        browser_memory["filename"] = st.session_state['c_filename']
                        
                        st.success("✅ تمت المعالجة! تم حفظ النتائج في ذاكرة المتصفح النشطة.")
                    else:
                        st.info("👍 لا توجد أي متغيرات أو فروقات، الملفين متطابقين تماماً.")
                except Exception as e:
                    st.error(f"حدث خطأ أثناء قراءة المستند: {str(e)}")
        else:
            st.warning("يرجى تزويد النظام بملفات الـ Word لبدء المطابقة الرقمية.")

    # -------------------------------------------------------------------------
    # 6. لوحة التحكم الفوري والفلاتر الذكية المتكاملة
    # -------------------------------------------------------------------------
    if 'c_results' in st.session_state:
        df_res = st.session_state['c_results']
        cnt = st.session_state['c_counters']
        
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f"<div class='report-box'>🔹 صافي المتغير الكلي: {cnt['net_total']:+d}</div>", unsafe_allow_html=True)
        with c2: st.markdown(f"<div class='report-box'>🔹 صافي المستحقة الحالية: {cnt['net_eligible']:+d}</div>", unsafe_allow_html=True)
        with c3: st.markdown(f"<div class='report-box'>🔹 المضاف: {cnt['added_fam']} عائلة | المحذوف: {cnt['deleted_fam']} عائلة</div>", unsafe_allow_html=True)
        
        st.markdown("### 🔐 الترحيل للأرشيف السحابي")
        if st.button("💾 حفظ البيانات الحالية في قوقل شيت بصورة دائمة"):
            if conn is not None:
                with st.spinner('جاري إرسال حقول الأسماء الرباعية للسحابة...'):
                    try:
                        try: existing_df = pd.DataFrame(conn.read(worksheet="Sheet1", ttl=0))
                        except Exception: existing_df = pd.DataFrame()
                        
                        df_to_save = df_res.copy()
                        df_to_save["تاريخ الفحص التلقائي"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                        df_to_save["الملف المرجعي"] = st.session_state['c_filename']
                        
                        final_archive = pd.concat([existing_df, df_to_save], ignore_index=True)
                        conn.update(worksheet="Sheet1", data=final_archive)
                        st.success("🌟 تم تأمين قفل البيانات داخل ملف قوقل شيت المركزي بنجاح.")
                    except Exception as ex:
                        st.error(f"عطل أثناء الاتصال السحابي: {str(ex)}")
        
        st.markdown("---")
        st.markdown("<h4 style='text-align: right;'>🎯 لوحة التدقيق الإحصائي والبحث المدمج</h4>", unsafe_allow_html=True)
        
        total_all = len(df_res)
        total_added = len(df_res[df_res["الاسم الرباعي (سابقاً)"] == "✨ (مضاف حديثاً)"])
        total_deleted = len(df_res[df_res["الاسم الرباعي (حالياً)"] == "❌ (محذوف / منقول)"])
        
        modified_only = df_res[(df_res["الاسم الرباعي (سابقاً)"] != "✨ (مضاف حديثاً)") & (df_res["الاسم الرباعي (حالياً)"] != "❌ (محذوف / منقول)")]
        total_eligible_changed = len(modified_only[modified_only["المستحقة (سابقاً)"] != modified_only["المستحقة (حالياً)"]])
        total_withheld_changed = len(modified_only[modified_only["المحجوبين (سابقاً)"] != modified_only["المحجوبين (حالياً)"]])
        total_quota_changed = len(modified_only[modified_only["الكلية (سابقاً)"] != modified_only["الكلية (حالياً)"]])
        
        filter_options = [
            f"📋 كشف جميع المتغيرات العامة ({total_all})",
            f"✨ فلتر العوائل المضافة حديثاً ({total_added})",
            f"❌ فلتر العوائل المحذوفة والمنقولة ({total_deleted})",
            f"🟢 فلتر الإحصاء: المتغير مستحقها ({total_eligible_changed})",
            f"🔴 فلتر الإحصاء: المتغير محجوبها ({total_withheld_changed})",
            f"🔵 فلتر الإحصاء: المتغير حصتها الكلية ({total_quota_changed})"
        ]
        
        selected_filter = st.selectbox("🎯 تصفية حسب نوع التغير الإحصائي:", filter_options)
        
        # 👤 محرك البحث الذكي بالأسماء الرباعية وأرقام البطاقات داخل النظام
        search_q = st.text_input("👤 ابحث هنا بالاسم الرباعي للمواطن أو رقم البطاقة التموينية:", "")
        
        if "✨" in selected_filter:
            filtered_df = df_res[df_res["الاسم الرباعي (سابقاً)"] == "✨ (مضاف حديثاً)"]
        elif "❌" in selected_filter:
            filtered_df = df_res[df_res["الاسم الرباعي (حالياً)"] == "❌ (محذوف / منقول)"]
        elif "🟢" in selected_filter:
            filtered_df = modified_only[modified_only["المستحقة (سابقاً)"] != modified_only["المستحقة (حالياً)"]]
        elif "🔴" in selected_filter:
            filtered_df = modified_only[modified_only["المحجوبين (سابقاً)"] != modified_only["المحجوبين (حالياً)"]]
        elif "🔵" in selected_filter:
            filtered_df = modified_only[modified_only["الكلية (سابقاً)"] != modified_only["الكلية (حالياً)"]]
        else:
            filtered_df = df_res

        if search_q:
            filtered_df = filtered_df[
                filtered_df["الاسم الرباعي (حالياً)"].str.contains(search_q, na=False) | 
                filtered_df["الاسم الرباعي (سابقاً)"].str.contains(search_q, na=False) |
                filtered_df["رقم البطاقة"].str.contains(search_q, na=False)
            ]
            
        st.markdown(f"<p style='text-align: right; color: gray;'>السجلات المطابقة للبحث والفلتر الحركي: {len(filtered_df)} مواطن</p>", unsafe_allow_html=True)
        
        for idx, row in filtered_df.iterrows():
            if "❌" in row["الاسم الرباعي (حالياً)"]:
                display_name = row["الاسم الرباعي (سابقاً)"]
                status_badge = "🔴 محذوف من الوجبة"
            elif "✨" in row["الاسم الرباعي (سابقاً)"]:
                display_name = row["الاسم الرباعي (حالياً)"]
                status_badge = "🟢 قيد مضاف جديد"
            else:
                display_name = row["الاسم الرباعي (حالياً)"]
                status_badge = "🟡 قيد معدل الحصص"
                
            box_title = f"{status_badge} | {display_name} (رقم البطاقة: {row['رقم البطاقة']})"
            
            with st.expander(box_title):
                col_old, col_new = st.columns(2)
                with col_old:
                    st.markdown("<div class='month-card'><b>📅 الحصص السابقة (الشهر الماضي):</b><br><br>"
                                f"▪️ إجمالي الكلية: {row['الكلية (سابقاً)']}<br>"
                                f"▪️ الأفراد المستحقين: {row['المستحقة (سابقاً)']}<br>"
                                f"▪️ الأفراد المحجوبين: {row['المحجوبين (سابقاً)']}"
                                "</div>", unsafe_allow_html=True)
                with col_new:
                    st.markdown("<div class='month-card'><b>🌟 الحصص الحالية (المتغير الحالي):</b><br><br>"
                                f"▪️ إجمالي الكلية: <span class='badge-total'>{row['الكلية (حالياً)']}</span><br><br>"
                                f"▪️ الأفراد المستحقين: <span class='badge-eligible'>{row['المستحقة (حالياً)']}</span><br><br>"
                                f"▪️ الأفراد المحجوبين: <span class='badge-withheld'>{row['المحجوبين (حالياً)']}</span>"
                                "</div>", unsafe_allow_html=True)
        
        st.markdown("---")
        # زر استخراج تقرير الورد المبني على الفروقات والمتغيرات الحالية فقط
        word_data = create_current_only_word_report(df_res, st.session_state['c_filename'])
        st.download_button(
            label="📥 تحميل تقرير المتغيرات الحالية فقط (Word)", 
            data=word_data, 
            file_name=f"تقرير_المتغيرات_الحالية_{st.session_state['c_filename']}.docx"
        )

with tab2:
    if st.button("🔄 تحديث وعرض شيت الأرشيف"):
        try:
            if conn is not None:
                st.dataframe(conn.read(worksheet="Sheet1", ttl=0), use_container_width=True, hide_index=True)
            else:
                st.error("الاتصال السحابي بقاعدة البيانات غير نشط.")
        except Exception as ex: 
            st.error(ex)
