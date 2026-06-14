import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from streamlit_gsheets import GSheetsConnection
import datetime

# -----------------------------------------------------------------------------
# 1. الواجهة السينمائية الفاخرة والتصميم المعاصر (Modern Dark/Light Hybrid UI)
# -----------------------------------------------------------------------------
st.set_page_config(page_title="منظومة الجرد التمويني الذكية", layout="wide", initial_sidebar_state="collapsed")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Cairo:wght@300;400;600;700&display=swap');
    
    * { font-family: 'Cairo', sans-serif; text-align: right; direction: rtl; }
    
    /* خلفيات وتأثيرات زجاجية معاصرة */
    .stApp { background: radial-gradient(circle at 50% 50%, #f8f9fa 0%, #e9ecef 100%); }
    
    /* صناديق الإحصائيات الفخمة KPI Cards */
    .kpi-container { display: flex; gap: 15px; margin-bottom: 25px; justify-content: space-between; flex-wrap: wrap; }
    .kpi-card { flex: 1; min-width: 220px; background: rgba(255, 255, 255, 0.85); backdrop-filter: blur(10px); padding: 20px; border-radius: 16px; border: 1px solid rgba(255,255,255,0.5); box-shadow: 0 8px 32px rgba(31, 38, 135, 0.05); transition: transform 0.3s ease; }
    .kpi-card:hover { transform: translateY(-5px); }
    .kpi-title { font-size: 14px; color: #6c757d; font-weight: 600; }
    .kpi-value { font-size: 24px; font-weight: 700; color: #212529; margin-top: 5px; }
    
    /* كروت العوائل الديناميكية المذهلة */
    .family-card-wrapper { background: white; padding: 18px; border-radius: 16px; margin-bottom: 15px; box-shadow: 0 4px 12px rgba(0,0,0,0.03); border-right: 8px solid #cbd5e1; transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1); }
    .family-card-wrapper:hover { transform: scale(1.01); box-shadow: 0 12px 24px rgba(0,0,0,0.08); }
    
    /* الألوان الديناميكية الفاخرة حسب الحالة */
    .card-matched { border-right-color: #2ec4b6 !important; background: linear-gradient(90deg, #ffffff 90%, #e6f9f8 100%); }
    .card-modified { border-right-color: #ff9f1c !important; background: linear-gradient(90deg, #ffffff 90%, #fff5e6 100%); }
    .card-added { border-right-color: #011627 !important; background: linear-gradient(90deg, #ffffff 90%, #e6ecef 100%); }
    .card-deleted { border-right-color: #e71d36 !important; background: linear-gradient(90deg, #ffffff 90%, #fde8ea 100%); }
    
    /* كروت تفاصيل الحصص الداخلية الصغيرة */
    .quota-box { background: #f8f9fa; border-radius: 12px; padding: 10px; text-align: center; border: 1px solid #f1f3f5; }
    .quota-label { font-size: 12px; color: #868e96; display: block; }
    .quota-num { font-size: 18px; font-weight: 700; display: block; }
    
    .q-total { color: #4dadf7; }
    .q-eligible { color: #40c057; }
    .q-withheld { color: #ff6b6b; }
    
    /* أزرار عصرية أنيقة */
    div.stButton > button { background: linear-gradient(135deg, #011627 0%, #2ec4b6 100%); color: white; border-radius: 12px; border: none; height: 50px; font-weight: 600; font-size: 16px; transition: all 0.3s ease; box-shadow: 0 4px 15px rgba(1, 22, 39, 0.15); }
    div.stButton > button:hover { background: linear-gradient(135deg, #2ec4b6 0%, #011627 100%); transform: translateY(-2px); box-shadow: 0 6px 20px rgba(1, 22, 39, 0.25); }
    
    /* تجميل الـ Expander الخاص بستريمليت */
    .streamlit-expanderHeader { background-color: transparent !important; border: none !important; font-weight: 600 !important; font-size: 16px !important; }
    </style>
""", unsafe_allow_html=True)

# الاتصال بقوقل شيت للأرشفة التاريخية
conn = None
try: conn = st.connection("gsheets", type=GSheetsConnection)
except Exception: st.error("⚠️ فشل الاتصال بقاعدة البيانات السحابية (Google Sheets).")

# -----------------------------------------------------------------------------
# 2. ذاكرة النظام الذكية لمنع الفقدان عند التحديث
# -----------------------------------------------------------------------------
@st.cache_resource
def get_persistent_memory(): return {"df": None, "counters": None, "filename": None}
browser_memory = get_persistent_memory()

if 'c_results' not in st.session_state and browser_memory["df"] is not None:
    st.session_state['c_results'] = browser_memory["df"]
    st.session_state['c_counters'] = browser_memory["counters"]
    st.session_state['c_filename'] = browser_memory["filename"]

# -----------------------------------------------------------------------------
# 3. محرك القراءة المرن والمطور (الحل الحقيقي لمشاكل تداخل حقول الوورد)
# -----------------------------------------------------------------------------
def advanced_flexible_parse(cells):
    """
    محرك ذكي لا يعتمد على الترتيب الثابت للأعمدة، بل يقوم بتشريح محتوى الخلايا برمجياً
    وفصل الأسماء عن الأرقام، ثم ترتيب الحصص تنازلياً لضمان عدم اختلاط البيانات التالفة.
    """
    clean_cells = [str(c).strip().replace('\n', ' ') for c in cells if c is not None]
    joined_text = " ".join(clean_cells)
    
    # استبعاد صفوف العناوين والترويسات والأسطر الفارغة
    if not clean_cells or any(word in joined_text for word in ["المركز", "الوكيل", "اسم الوكيل", "زكرمل", "توقيع"]):
        return {}

    # 1. استخراج الرقم التمويني (رقم البطاقة المكون من 5 أرقام فما فوق)
    card_candidates = [c for c in clean_cells if c.isdigit() and len(c) >= 5]
    if not card_candidates:
        return {}
    card_num = card_candidates[0]

    # 2. استخراج الاسم الرباعي (الخلية النصية الأطول الخالية من الأرقام)
    name_candidates = [c for c in clean_cells if any('\u0600' <= char <= '\u06FF' for char in c) and not any(char.isdigit() for char in c)]
    if not name_candidates:
        return {}
    quad_name = max(name_candidates, key=len)

    # 3. تجميع كل الأرقام الصغيرة الأخرى المتوفرة في الصف (الحصص والتسلسل)
    all_numbers = []
    for c in clean_cells:
        # استخراج الأرقام حتى لو كانت مدمجة مع نصوص تنظيفية
        nums = [int(s) for s in c.split() if s.isdigit() and len(s) < 5]
        all_numbers.extend(nums)

    if len(all_numbers) < 2:
        return {}  # بيانات غير كاملة للحصص في هذا الصف

    # حل معضلة الإزاحة والتداخل (الذكاء الرياضي):
    # دائماً في النظام التمويني: الكلي >= المستحق >= المحجوب
    quota_numbers = sorted([n for n in all_numbers if n < 40], reverse=True) # استبعاد التسلسلات الكبيرة إن وجدت
    
    # تأمين قراءة الأرقام الثلاثية حتى لو اختفت أحدها
    total = quota_numbers[0] if len(quota_numbers) >= 1 else 0
    eligible = quota_numbers[1] if len(quota_numbers) >= 2 else total
    withheld = quota_numbers[2] if len(quota_numbers) >= 3 else (total - eligible)
    if withheld < 0: withheld = 0

    # محاولة تخمين التسلسل (الرقم المتبقي الذي لم يدخل في الحصص)
    remaining_nums = [n for n in all_numbers if n not in quota_numbers]
    seq = remaining_nums[0] if remaining_nums else "-"

    return {card_num: {"seq": seq, "name": quad_name, "total": total, "eligible": eligible, "withheld": withheld}}

def extract_clean_records(file_obj):
    records = {}
    doc = Document(file_obj)
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text for cell in row.cells]
            records.update(advanced_flexible_parse(cells))
    return records

# -----------------------------------------------------------------------------
# 4. محرك المقارنة الرصين عالي الدقة (فروقات الأسماء والحصص والإضافة والحذف)
# -----------------------------------------------------------------------------
def compare_records(old_data, new_data):
    results = []
    counters = {"added_fam": 0, "deleted_fam": 0, "unchanged_fam": 0, "modified_fam": 0, "net_total": 0}
    all_cards = set(old_data.keys()).union(set(new_data.keys()))
    
    for card in sorted(all_cards):
        if card in old_data and card in new_data:
            o, n = old_data[card], new_data[card]
            
            # رصد دقيق لأي اختلاف في الحقول الحصصية أو الاسم الرباعي
            has_diff = (o["total"] != n["total"] or o["eligible"] != n["eligible"] or 
                        o["withheld"] != n["withheld"] or o["name"] != n["name"])
            
            if has_diff:
                counters["modified_fam"] += 1
                counters["net_total"] += (n["eligible"] - o["eligible"])
                
                changes = []
                if o["name"] != n["name"]: changes.append("تعديل اسم رب الأسرة")
                if o["total"] != n["total"]: changes.append(f"تغير الكلي ({o['total']} ➔ {n['total']})")
                if o["eligible"] != n["eligible"]: changes.append(f"تغير المستحق ({o['eligible']} ➔ {n['eligible']})")
                if o["withheld"] != n["withheld"]: changes.append(f"تغير المحجوب ({o['withheld']} ➔ {n['withheld']})")
                
                results.append({
                    "التسلسل": n["seq"], "رقم البطاقة": card,
                    "الاسم الرباعي (سابقاً)": o["name"], "الاسم الرباعي (حالياً)": n["name"],
                    "الكلية (سابقاً)": o["total"], "الكلية (حالياً)": n["total"],
                    "المستحقة (سابقاً)": o["eligible"], "المستحقة (حالياً)": n["eligible"],
                    "المحجوبين (سابقاً)": o["withheld"], "المحجوبين (حالياً)": n["withheld"],
                    "الحالة": "🟡 قيد معدل الحصص", "تفاصيل": " | ".join(changes)
                })
            else:
                counters["unchanged_fam"] += 1
                results.append({
                    "التسلسل": n["seq"], "رقم البطاقة": card,
                    "الاسم الرباعي (سابقاً)": o["name"], "الاسم الرباعي (حالياً)": n["name"],
                    "الكلية (سابقاً)": o["total"], "الكلية (حالياً)": n["total"],
                    "المستحقة (سابقاً)": o["eligible"], "المستحقة (حالياً)": n["eligible"],
                    "المحجوبين (سابقاً)": o["withheld"], "المحجوبين (حالياً)": n["withheld"],
                    "الحالة": "✅ متطابق (بدون تغيير)", "تفاصيل": "البيانات متطابقة تماماً"
                })
                
        elif card in old_data:
            o = old_data[card]
            counters["deleted_fam"] += 1
            counters["net_total"] -= o["eligible"]
            results.append({
                "التسلسل": o["seq"], "رقم البطاقة": card,
                "الاسم الرباعي (سابقاً)": o["name"], "الاسم الرباعي (حالياً)": "❌ (محذوف / منقول)",
                "الكلية (سابقاً)": o["total"], "الكلية (حالياً)": 0,
                "المستحقة (سابقاً)": o["eligible"], "المستحقة (حالياً)": 0,
                "المحجوبين (سابقاً)": o["withheld"], "المحجوبين (حالياً)": 0,
                "الحالة": "🔴 محذوف من الوجبة", "تفاصيل": "تم رفع أو نقل العائلة بالكامل من كشوفات الوكيل"
            })
            
        elif card not in old_data and card in new_data:
            n = new_data[card]
            counters["added_fam"] += 1
            counters["net_total"] += n["eligible"]
            results.append({
                "التسلسل": n["seq"], "رقم البطاقة": card,
                "الاسم الرباعي (سابقاً)": "✨ (مضاف حديثاً)", "الاسم الرباعي (حالياً)": n["name"],
                "الكلية (سابقاً)": 0, "الكلية (حالياً)": n["total"],
                "المستحقة (سابقاً)": 0, "المستحقة (حالياً)": n["eligible"],
                "المحجوبين (سابقاً)": 0, "المحجوبين (حالياً)": n["withheld"],
                "الحالة": "🟢 قيد مضاف جديد", "تفاصيل": "عائلة جديدة نزلت في وجبة هذا الشهر"
            })
            
    return results, counters

# -----------------------------------------------------------------------------
# 5. بناء واجهة العرض الرسومية والتحكم المعاصر
# -----------------------------------------------------------------------------
tab1, tab2 = st.tabs(["🔎 منظومة التدقيق والمقارنة الذكية", "📜 الأرشيف السحابي المركزي"])

with tab1:
    st.markdown("<h2 style='text-align: right; font-weight: 700; color: #011627;'>نظام الجرد والمطابقة الفوري للوكلاء</h2>", unsafe_allow_html=True)
    st.markdown("<p style='text-align: right; color: #6c757d; margin-bottom:30px;'>ارفع ملفات الوورد للشهرين الحالي والسابق للمطابقة الإلكترونية الفورية برصانة مطلقة.</p>", unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1: new_file = st.file_uploader("📂 كشف الشهر الحالي (الجديد)", type=['docx'])
    with col2: old_file = st.file_uploader("📂 كشف الشهر السابق (القديم)", type=['docx'])

    btn_space1, btn_space2 = st.columns([3, 1])
    with btn_space1: run_match = st.button("⚡ بدء معالجة ومطابقة السجلات الآن")
    with btn_space2:
        if st.button("🗑️ تصفير الذاكرة التخزينية"):
            browser_memory["df"], browser_memory["counters"], browser_memory["filename"] = None, None, None
            for k in ['c_results', 'c_counters', 'c_filename']:
                if k in st.session_state: del st.session_state[k]
            st.rerun()

    if run_match and old_file and new_file:
        with st.spinner('جاري تشريح المستندات وترتيب الحصص بدقة متناهية...'):
            old_records = extract_clean_records(old_file)
            new_records = extract_clean_records(new_file)
            res_list, cnt_dict = compare_records(old_records, new_records)
            
            if res_list:
                st.session_state['c_results'] = pd.DataFrame(res_list).sort_values(by="التسلسل")
                st.session_state['c_counters'] = cnt_dict
                st.session_state['c_filename'] = new_file.name.rsplit('.', 1)[0]
                
                browser_memory["df"] = st.session_state['c_results']
                browser_memory["counters"] = cnt_dict
                browser_memory["filename"] = st.session_state['c_filename']
                st.rerun()

    # -------------------------------------------------------------------------
    # 6. لوحة العرض السينمائية (عرض النتائج بتصميم فخم ومستقر)
    # -------------------------------------------------------------------------
    if 'c_results' in st.session_state:
        df_res = st.session_state['c_results']
        cnt = st.session_state['c_counters']
        
        # حقن كروت الإحصائيات الفاخرة الموزعة بدقة المعاصرة
        st.markdown(f"""
            <div class='kpi-container'>
                <div class='kpi-card'><div class='kpi-title'>👥 إجمالي العوائل المدرجة</div><div class='kpi-value'>{len(df_res)} عائلة</div></div>
                <div class='kpi-card'><div class='kpi-title'>📈 صافي الفروقات الحصصية</div><div class='kpi-value' style='color:#40c057;'>{cnt['net_total']:+d} حصة</div></div>
                <div class='kpi-card'><div class='kpi-title'>✨ مضافة / ❌ محذوفة</div><div class='kpi-value'>{cnt['added_fam']}+ | {cnt['deleted_fam']}-</div></div>
                <div class='kpi-card'><div class='kpi-title'>🟡 عوائل معدلة الحصص</div><div class='kpi-value' style='color:#ff9f1c;'>{cnt['modified_fam']} عائلة</div></div>
            </div>
        """, unsafe_allow_html=True)
        
        # شريط البحث والفلترة السريعة لسهولة الوصول
        st.markdown("### 🎯 محرك الفرز السريع والبحث الذكي")
        f_col, s_col = st.columns([2, 2])
        with f_col:
            filter_choice = st.selectbox("عرض الفئة المطلوبة:", [
                "📋 كشف السجلات الكامل للوكيل", "✅ العوائل المستقرة والمتطابقة فقط", 
                "🟡 العوائل التي جرى عليها تعديل", "✨ العوائل المضافة حديثاً"، "❌ العوائل المحذوفة والمنقولة"
            ])
        with s_col:
            search_query = st.text_input("👤 ابحث باسم المواطن أو رقم البطاقة التموينية:")

        # تصفية البيانات برمجياً طبقاً للفلتر
        if "المستقرة" in filter_choice: filtered_df = df_res[df_res["الحالة"].str.contains("متطابق")]
        elif "تعديل" in filter_choice: filtered_df = df_res[df_res["الحالة"].str.contains("معدل")]
        elif "المضافة" in filter_choice: filtered_df = df_res[df_res["الحالة"].str.contains("مضاف")]
        elif "المحذوفة" in filter_choice: filtered_df = df_res[df_res["الحالة"].str.contains("محذوف")]
        else: filtered_df = df_res

        if search_query:
            filtered_df = filtered_df[
                filtered_df["الاسم الرباعي (حالياً)"].astype(str).str.contains(search_query) |
                filtered_df["الاسم الرباعي (سابقاً)"].astype(str).str.contains(search_query) |
                filtered_df["رقم البطاقة"].astype(str).str.contains(search_query)
            ]

        st.markdown(f"<p style='color:#868e96; font-size:14px; margin-bottom:15px;'>عدد العوائل المطابقة حالياً: {len(filtered_df)}</p>", unsafe_allow_html=True)

        # 👤 حلقة رسم كروت العوائل بالطراز العصري الرفيع
        for _, row in filtered_df.iterrows():
            status = row["الحالة"]
            name_now, name_old = row["الاسم الرباعي (حالياً)"], row["الاسم الرباعي (سابقاً)"]
            display_name = name_old if "❌" in name_now else name_now
            
            # تحديد نوع التغليف اللوني الخارجي للكرت بشكل ديناميكي أنيق
            card_cls = "card-matched"
            if "معدل" in status: card_cls = "card-modified"
            elif "مضاف" in status: card_cls = "card-added"
            elif "محذوف" in status: card_cls = "card-deleted"
            
            st.markdown(f"<div class='family-card-wrapper {card_cls}'>", unsafe_allow_html=True)
            
            # توليد عنوان الكرت الفخم مع الأيقونة والاسم التميز
            box_header = f"🔹 ت: {row['التسلسل']} | {display_name} (رقم البطاقة: {row['رقم البطاقة']})"
            
            with st.expander(box_header):
                st.markdown(f"<p style='font-size:13px; color:#495057; margin-bottom:15px;'><b>🔍 تشخيص النظام:</b> {status} — {row['تفاصيل']}</p>", unsafe_allow_html=True)
                
                left_panel, right_panel = st.columns(2)
                with left_panel:
                    st.markdown("<p style='font-size:13px; font-weight:600; color:#868e96; margin-bottom:8px; border-bottom:1px solid #dee2e6;'>📅 موقف الشهر السابق</p>", unsafe_allow_html=True)
                    sub1, sub2, sub3 = st.columns(3)
                    sub1.markdown(f"<div class='quota-box'><span class='quota-label'>👥 كلي</span><span class='quota-num q-total'>{row['الكلية (سابقاً)']}</span></div>", unsafe_allow_html=True)
                    sub2.markdown(f"<div class='quota-box'><span class='quota-label'>✅ مستحق</span><span class='quota-num q-eligible'>{row['المستحقة (سابقاً)']}</span></div>", unsafe_allow_html=True)
                    sub3.markdown(f"<div class='quota-box'><span class='quota-label'>🚫 محجوب</span><span class='quota-num q-withheld'>{row['المحجوبين (سابقاً)']}</span></div>", unsafe_allow_html=True)
                
                with right_panel:
                    st.markdown("<p style='font-size:13px; font-weight:600; color:#011627; margin-bottom:8px; border-bottom:1px solid #cbd5e1;'>🌟 موقف الشهر الحالي (النهائي)</p>", unsafe_allow_html=True)
                    sub4, sub5, sub6 = st.columns(3)
                    sub4.markdown(f"<div class='quota-box'><span class='quota-label'>👥 كلي</span><span class='quota-num q-total'>{row['الكلية (حالياً)']}</span></div>", unsafe_allow_html=True)
                    sub5.markdown(f"<div class='quota-box'><span class='quota-label'>✅ مستحق</span><span class='quota-num q-eligible'>{row['المستحقة (حالياً)']}</span></div>", unsafe_allow_html=True)
                    sub6.markdown(f"<div class='quota-box'><span class='quota-label'>🚫 محجوب</span><span class='quota-num q-withheld'>{row['المحجوبين (حالياً)']}</span></div>", unsafe_allow_html=True)
            
            st.markdown("</div>", unsafe_allow_html=True)

with tab2:
    st.markdown("### ☁️ الترحيل المركزي والأرشيف السحابي")
    if st.button("💾 ترحيل وحفظ كشوفات هذا الوكيل الحالية إلى Google Sheets"):
        if conn is not None and 'c_results' in st.session_state:
            with st.spinner('جاري مزامنة السجلات مع السحابة...'):
                try:
                    try: existing_df = pd.DataFrame(conn.read(worksheet="Sheet1", ttl=0))
                    except Exception: existing_df = pd.DataFrame()
                    
                    df_to_save = st.session_state['c_results'].copy()
                    df_to_save["تاريخ الأرشفة"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                    df_to_save["اسم الوكيل"] = st.session_state['c_filename']
                    
                    final_df = pd.concat([existing_df, df_to_save], ignore_index=True)
                    conn.update(worksheet="Sheet1", data=final_df)
                    st.success("🌟 تم مزامنة وأرشفة كشف الوكيل بنجاح تام!")
                except Exception as ex: st.error(f"فشل الترحيل: {ex}")
