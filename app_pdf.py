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
st.set_page_config(page_title="نظام المطابقة الذكي المؤمن المستمر", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #2E7D32; color: white; width: 100%; font-weight: bold; border-radius: 8px; border: none; height: 48px; font-size: 16px; }
    .report-box { background-color: var(--secondary-background-color); color: var(--text-color); padding: 12px; border-radius: 8px; border-right: 5px solid #2E7D32; text-align: right; margin-bottom: 10px; font-weight: bold; }
    .instruction-box { background-color: #FFF3E0; color: #E65100; padding: 15px; border-radius: 8px; border-right: 5px solid #EF6C00; text-align: right; margin-bottom: 20px; font-size: 15px; line-height: 1.6; }
    .ilove-link { color: #E65100 !important; font-weight: bold; text-decoration: underline !important; background-color: #FFE0B2; padding: 3px 8px; border-radius: 4px; }
    
    /* تنسيقات كروت العوائل الملونة */
    .month-card { background-color: var(--background-color); padding: 12px; border-radius: 6px; border: 1px solid #e0e0e0; margin-bottom: 8px; text-align: right; }
    .badge-total { background-color: #0288D1; color: white; padding: 3px 10px; border-radius: 4px; font-weight: bold; display: inline-block; }
    .badge-eligible { background-color: #388E3C; color: white; padding: 3px 10px; border-radius: 4px; font-weight: bold; display: inline-block; }
    .badge-withheld { background-color: #D32F2F; color: white; padding: 3px 10px; border-radius: 4px; font-weight: bold; display: inline-block; }
    </style>
""", unsafe_allow_html=True)

# الاتصال بقوقل شيت وتأمين التوصيل للأرشيف
conn = None
try:
    conn = st.connection("gsheets", type=GSheetsConnection)
except Exception:
    st.error("⚠️ فشل الاتصال بقاعدة البيانات السحابية (Google Sheets)، تحقق من الإعدادات.")

# -----------------------------------------------------------------------------
# 2. محرك الذاكرة المستمرة (لحفظ البيانات من المسح عند التحديث)
# -----------------------------------------------------------------------------
@st.cache_resource
def get_persistent_memory():
    # مستودع سحابي مؤقت خاص بمتصفح المستخدم لعدم فقدان البيانات عند الـ Refresh
    return {"df": None, "counters": None, "filename": None}

browser_memory = get_persistent_memory()

# خاصية الاستعادة التلقائية الذكية فور تحديث الصفحة
if 'c_results' not in st.session_state and browser_memory["df"] is not None:
    st.session_state['c_results'] = browser_memory["df"]
    st.session_state['c_counters'] = browser_memory["counters"]
    st.session_state['c_filename'] = browser_memory["filename"]
    st.toast("🔄 تم استعادة آخر قراءة للبيانات تلقائياً من ذاكرة المتصفح الحية!", icon="⚡")

# -----------------------------------------------------------------------------
# 3. محركات المعالجة الذكية لجداول الـ Word
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
# 4. واجهة التطبيق الرئيسية واللوحات التفاعلية
# -----------------------------------------------------------------------------
tab1, tab2 = st.tabs(["🔎 مطابقة ملفات Word", "📜 الأرشيف التاريخي"])

with tab1:
    st.markdown("<h3 style='text-align: right;'>نظام المطابقة والفرز الإحصائي الذكي (المحمي من التحديث)</h3>", unsafe_allow_html=True)
    
    st.markdown("""
    <div class='instruction-box'>
        💡 <b>ميزة حماية البيانات نشطة حالياً:</b><br>
        1. إذا قمت بعمل تحديث (Refresh) للصفحة من الموبايل، <b>لن تفقد البيانات المرفوعة والمطابقة سابقاً</b> بفضل ذاكرة النظام المستمرة.<br>
        2. لتحويل الملفات أولاً: <a href='https://www.ilovepdf.com/pdf_to_word' target='_blank' class='ilove-link'>الانتقال إلى أداة iLovePDF</a> ثم ارفع الملفات بالأسفل.
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1: new_file = st.file_uploader("الملف الجديد بصيغة (Word)", type=['docx'], key="n_f")
    with col2: old_file = st.file_uploader("الملف القديم بصيغة (Word)", type=['docx'], key="o_f")

    col_btn1, col_btn2 = st.columns([3, 1])
    with col_btn1:
        run_match = st.button("🚀 تشغيل المطابقة الآن")
    with col_btn2:
        # زر إضافي لمسح الذاكرة وبدء ملفات جديدة كلياً عند الحاجة
        clear_mem = st.button("🗑️ تفريغ الذاكرة للبدء من جديد")
        if clear_mem:
            browser_memory["df"] = None
            browser_memory["counters"] = None
            browser_memory["filename"] = None
            for k in ['c_results', 'c_counters', 'c_filename']:
                if k in st.session_state: del st.session_state[k]
            st.rerun()

    if run_match:
        if old_file and new_file:
            with st.spinner('جاري قراءة الجداول وتحليل الفروقات ديناميكياً...'):
                try:
                    old_data = extract_clean_records(old_file)
                    new_data = extract_clean_records(new_file)
                    results, counters = compare_records(old_data, new_data)
                    
                    if results:
                        # التخزين في الجلسة الحالية
                        st.session_state['c_results'] = pd.DataFrame(results)
                        st.session_state['c_counters'] = counters
                        st.session_state['c_filename'] = new_file.name.rsplit('.', 1)[0]
                        
                        # قفل البيانات فوراً في الذاكرة المستمرة للمتصفح لمنع مسحها عند التحديث
                        browser_memory["df"] = st.session_state['c_results']
                        browser_memory["counters"] = st.session_state['c_counters']
                        browser_memory["filename"] = st.session_state['c_filename']
                        
                        st.success("✅ تمت المطابقة بنجاح! وتم تأمين البيانات في المتصفح تلقائياً ضد التحديث.")
                    else:
                        st.info("👍 تطابق كامل ومثالي، لا توجد فروقات بين الملفين.")
                except Exception as e:
                    st.error(f"حدث خطأ أثناء قراءة البيانات: {str(e)}")
        else:
            st.warning("يرجى رفع الملفين بصيغة Word أولاً للبدء.")

    # -------------------------------------------------------------------------
    # 5. لوحة الفلاتر والفرز الإحصائي الذكي للموبايل عند توفر البيانات
    # -------------------------------------------------------------------------
    if 'c_results' in st.session_state:
        df_res = st.session_state['c_results']
        cnt = st.session_state['c_counters']
        
        # عرض الإحصائيات العامة المباشرة
        c1, c2, c3 = st.columns(3)
        with c1: st.markdown(f"<div class='report-box'>🔹 صافي الكلية: {cnt['net_total']:+d}</div>", unsafe_allow_html=True)
        with c2: st.markdown(f"<div class='report-box'>🔹 صافي المستحقة: {cnt['net_eligible']:+d}</div>", unsafe_allow_html=True)
        with c3: st.markdown(f"<div class='report-box'>🔹 المضاف: {cnt['added_fam']} | المحذوف: {cnt['deleted_fam']}</div>", unsafe_allow_html=True)
        
        # 🔐 ميزة الحفظ والتأمين السحابي الاختياري للأرشفة الدائمة
        st.markdown("### 🔐 ترحيل دائم للأرشيف")
        if st.button("💾 اضغط هنا لحفظ وتأمين هذه البيانات في الأرشيف السحابي فوراً"):
            if conn is not None:
                with st.spinner('جاري ترحيل وتأمين البيانات في Google Sheets...'):
                    try:
                        try: existing_df = pd.DataFrame(conn.read(worksheet="Sheet1", ttl=0))
                        except Exception: existing_df = pd.DataFrame()
                        
                        df_to_save = df_res.copy()
                        df_to_save["تاريخ الفحص"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                        df_to_save["اسم الملف المرجعي"] = st.session_state['c_filename']
                        
                        final_archive = pd.concat([existing_df, df_to_save], ignore_index=True)
                        conn.update(worksheet="Sheet1", data=final_archive)
                        st.success("🌟 ممتاز! تم قفل وحفظ البيانات داخل قوقل شيت بنجاح تام.")
                    except Exception as ex:
                        st.error(f"فشل الاتصال بقوقل شيت أثناء الترحيل: {str(ex)}")
        
        st.markdown("---")
        st.markdown("<h4 style='text-align: right;'>🎯 لوحة الفرز والتدقيق الذكي للمفتشين</h4>", unsafe_allow_html=True)
        
        total_all = len(df_res)
        total_added = len(df_res[df_res["الاسم (سابقاً)"] == "✨ (مضاف حديثاً)"])
        total_deleted = len(df_res[df_res["الاسم (حالياً)"] == "❌ (محذوف / منقول)"])
        
        modified_only = df_res[(df_res["الاسم (سابقاً)"] != "✨ (مضاف حديثاً)") & (df_res["الاسم (حالياً)"] != "❌ (محذوف / منقول)")]
        total_eligible_changed = len(modified_only[modified_only["المستحقة (سابقاً)"] != modified_only["المستحقة (حالياً)"]])
        total_withheld_changed = len(modified_only[modified_only["المحجوبين (سابقاً)"] != modified_only["المحجوبين (حالياً)"]])
        total_quota_changed = len(modified_only[modified_only["الكلية (سابقاً)"] != modified_only["الكلية (حالياً)"]])
        
        filter_options = [
            f"📋 عرض جميع الحالات المتغيرة والمحدثة ({total_all})",
            f"✨ العوائل المضافة حديثاً فقط ({total_added})",
            f"❌ العوائل المحذوفة أو المنقولة ({total_deleted})",
            f"🟢 العوائل التي تغير عدد مستحقيها ({total_eligible_changed})",
            f"🔴 العوائل التي تغير عدد أفرادها المحجوبين ({total_withheld_changed})",
            f"🔵 العوائل التي تغيرت حصتها الكلية الإجمالية ({total_quota_changed})"
        ]
        
        selected_filter = st.selectbox("🔍 اختر الفئة المراد تدقيقها من القائمة الإحصائية:", filter_options)
        search_q = st.text_input("👤 ابحث عن اسم مواطن محدد أو رقم بطاقة داخل الفئة المختارة:", "")
        
        if "✨" in selected_filter:
            filtered_df = df_res[df_res["الاسم (سابقاً)"] == "✨ (مضاف حديثاً)"]
        elif "❌" in selected_filter:
            filtered_df = df_res[df_res["الاسم (حالياً)"] == "❌ (محذوف / منقول)"]
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
                filtered_df["الاسم (حالياً)"].str.contains(search_q, na=False) | 
                filtered_df["الاسم (سابقاً)"].str.contains(search_q, na=False) |
                filtered_df["رقم البطاقة"].str.contains(search_q, na=False)
            ]
            
        st.markdown(f"<p style='text-align: right; color: gray;'>عدد السجلات المطابقة الحالية: {len(filtered_df)} مواطن</p>", unsafe_allow_html=True)
        
        for idx, row in filtered_df.iterrows():
            if "❌" in row["الاسم (حالياً)"]:
                display_name = row["الاسم (سابقاً)"]
                status_badge = "🔴 محذوف / منقول"
            elif "✨" in row["الاسم (سابقاً)"]:
                display_name = row["الاسم (حالياً)"]
                status_badge = "🟢 مضاف حديثاً"
            else:
                display_name = row["الاسم (حالياً)"]
                status_badge = "🟡 تم تعديل حصته"
                
            box_title = f"{status_badge} | {display_name} (البطاقة: {row['رقم البطاقة']})"
            
            with st.expander(box_title):
                col_old, col_new = st.columns(2)
                
                with col_old:
                    st.markdown("<div class='month-card'><b>📅 بيانات الشهر السابق:</b><br><br>"
                                f"▪️ الحصة الكلية: {row['الكلية (سابقاً)']}<br>"
                                f"▪️ الحصة المستحقة: {row['المستحقة (سابقاً)']}<br>"
                                f"▪️ الأفراد المحجوبين: {row['المحجوبين (سابقاً)']}"
                                "</div>", unsafe_allow_html=True)
                                
                with col_new:
                    st.markdown("<div class='month-card'><b>🌟 بيانات الشهر الحالي (الملونة):</b><br><br>"
                                f"▪️ الحصة الكلية: <span class='badge-total'>{row['الكلية (حالياً)']}</span><br><br>"
                                f"▪️ الحصة المستحقة: <span class='badge-eligible'>{row['المستحقة (حالياً)']}</span><br><br>"
                                f"▪️ الأفراد المحجوبين: <span class='badge-withheld'>{row['المحجوبين (حالياً)']}</span>"
                                "</div>", unsafe_allow_html=True)
        
        st.markdown("---")
        st.download_button("📥 تحميل تقرير الفروقات كاملاً بملف (Word)", data=create_word_table_report(df_res, st.session_state['c_filename']), file_name=f"تقرير_فروقات_{st.session_state['c_filename']}.docx")

with tab2:
    if st.button("🔄 جلب الأرشيف وتحديث القائمة"):
        try:
            if conn is not None:
                st.dataframe(conn.read(worksheet="Sheet1", ttl=0), use_container_width=True, hide_index=True)
            else:
                st.error("لا يمكن جلب البيانات، الاتصال السحابي مقطوع.")
        except Exception as ex: 
            st.error(ex)
