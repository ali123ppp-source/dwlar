import streamlit as st
import pandas as pd
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Pt, Cm, RGBColor
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import os
import re
from datetime import datetime

# مكتبات الـ PDF
try:
    import arabic_reshaper
    from bidi.algorithm import get_display
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib import colors
    LIBS_READY = True
except ImportError:
    LIBS_READY = False

# -----------------------------------------------------------------------------
# إعدادات الواجهة الأساسية
# -----------------------------------------------------------------------------
st.set_page_config(page_title="نظام المقارنة الشامل للوكلاء", layout="wide")

st.markdown("""
    <div style='background-color: #E8F8F5; padding: 10px; border-radius: 8px; text-align: center; border: 2px solid #1ABC9C; margin-bottom: 20px;'>
        <h4 style='color: #16A085; margin: 0;'>🔗 لتحويل ملفات PDF إلى Word بدقة عالية وبشكل مجاني</h4>
        <a href='https://www.ilovepdf.com/ar/pdf_to_word' target='_blank' style='font-size: 18px; font-weight: bold; color: #2980B9; text-decoration: none;'>اضغط هنا للانتقال إلى موقع iLovePDF</a>
    </div>
""", unsafe_allow_html=True)

st.markdown("""
    <style>
    th, td { text-align: right !important; dir: rtl !important; }
    div.stButton > button { background-color: #2C3E50; color: white; width: 100%; font-weight: bold; border-radius: 8px;}
    .report-box { background-color: #ECF0F1; padding: 15px; border-radius: 8px; border-right: 5px solid #2C3E50; text-align: right; margin-bottom: 10px;}
    .net-diff { font-size: 16px; font-weight: bold; margin-top: 5px; color: #2C3E50; border-top: 1px solid #BDC3C7; padding-top: 5px;}
    .stat-inc { font-size: 14px; color: #27AE60; font-weight: bold; }
    .stat-dec { font-size: 14px; color: #C0392B; font-weight: bold; }
    .compare-table { width: 100%; border-collapse: collapse; margin-top: 10px; }
    .compare-table th, .compare-table td { border: 1px solid #ddd; padding: 8px; text-align: center !important; }
    .compare-table th { background-color: #2C3E50; color: white; }
    .auto-detect-box { background-color: #FCF3CF; padding: 10px; border-radius: 5px; border-right: 5px solid #F1C40F; color: #7D6608; font-weight: bold; }
    </style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='text-align: right;'>نظام المقارنة الشامل والذكي 📄🔎</h1>", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# دالة الاستشعار الذكي للتواريخ (تم تحديثها لقراءة تواريخ الوورد المكتوبة بالإنجليزية)
# -----------------------------------------------------------------------------
def extract_document_date(file_obj):
    doc = Document(file_obj)
    texts = []
    
    # سحب النصوص
    for p in doc.paragraphs:
        if p.text.strip(): texts.append(p.text.strip())
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if cell.text.strip(): texts.append(cell.text.strip())
                
    texts.reverse() # البحث من الأسفل للأعلى
    
    # نمط للتواريخ الرقمية العادية (مثل 12/05/2023)
    date_pattern_num = re.compile(r'\b(\d{2,4}[/.-]\d{1,2}[/.-]\d{2,4})\b')
    # نمط للتواريخ النصية (مثل Saturday, March 14, 2026)
    date_pattern_eng = re.compile(r'([A-Za-z]+,\s+[A-Za-z]+\s+\d{1,2},\s+\d{4})')
    
    for t in texts:
        # 1. البحث عن الصيغة الإنجليزية أولاً (مثل ملفك المرفق)
        matches_eng = date_pattern_eng.findall(t)
        if matches_eng:
            for match_str in matches_eng:
                try:
                    parsed = datetime.strptime(match_str, '%A, %B %d, %Y')
                    file_obj.seek(0)
                    return parsed, match_str
                except ValueError:
                    continue

        # 2. البحث عن الصيغة الرقمية
        matches_num = date_pattern_num.findall(t)
        if matches_num:
            for match_str in matches_num:
                formats = ['%Y/%m/%d', '%d/%m/%Y', '%m/%d/%Y', '%Y-%m-%d', '%d-%m-%Y', '%Y.%m.%d', '%d.%m.%Y']
                for fmt in formats:
                    try:
                        parsed = datetime.strptime(match_str, fmt)
                        file_obj.seek(0)
                        return parsed, match_str
                    except ValueError:
                        continue
                        
    file_obj.seek(0)
    return None, None

# -----------------------------------------------------------------------------
# دالات المساعدة للـ PDF والـ Word
# -----------------------------------------------------------------------------
def fix_arabic(text):
    if not text: return ""
    try:
        return get_display(arabic_reshaper.reshape(str(text)))
    except Exception:
        return str(text)

def set_cell_shading(cell, color_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}"/>')
    tcPr.append(shd)

def set_cell_rotation(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    text_dir = parse_xml(f'<w:textDirection {nsdecls("w")} w:val="btLr"/>')
    tcPr.append(text_dir)

def repeat_header_row(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = parse_xml(f'<w:tblHeader {nsdecls("w")}/>')
    trPr.append(tblHeader)

# -----------------------------------------------------------------------------
# محرك الاستخراج
# -----------------------------------------------------------------------------
def extract_clean_records(file_obj):
    doc = Document(file_obj)
    records = {}
    for table in doc.tables:
        for row in table.rows:
            cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
            if not any(cells) or "المركز" in "".join(cells) or "الوكيل" in "".join(cells) or "اسم رب" in "".join(cells):
                continue
            name_idx = -1
            max_len = 0
            for i, c in enumerate(cells):
                if any('\u0600' <= char <= '\u06FF' for char in c) and not any(char.isdigit() for char in c):
                    if len(c) > max_len:
                        max_len = len(c)
                        name_idx = i
            if name_idx == -1: continue
            
            card_indices = [i for i, c in enumerate(cells) if c.isdigit() and len(c) >= 5]
            if not card_indices: continue
            
            card_num = cells[card_indices[1]] if len(card_indices) >= 2 else cells[card_indices[0]]
            old_card_num = cells[card_indices[0]] if len(card_indices) >= 2 else "-"
            
            seq = "-"
            for i in range(len(cells)-1, card_indices[-1], -1):
                if cells[i].isdigit():
                    seq = cells[i]
                    break
            
            digit_cells = [int(cells[i]) for i in range(name_idx) if cells[i].isdigit()]
            if len(digit_cells) >= 3:
                withheld, eligible, total = digit_cells[0], digit_cells[1], digit_cells[2]
            elif len(digit_cells) == 2:
                withheld, eligible, total = 0, digit_cells[0], digit_cells[1]
            else:
                continue
                
            records[card_num] = {
                "seq": seq, "name": cells[name_idx], "total": total, 
                "eligible": eligible, "withheld": withheld, "old_card": old_card_num
            }
    return records

# -----------------------------------------------------------------------------
# محرك المقارنة
# -----------------------------------------------------------------------------
def compare_records(old_data, new_data):
    results = []
    counters = {
        "total_fam": 0, "eligible_fam": 0, "withheld_fam": 0, 
        "added_fam": 0, "deleted_fam": 0,
        "inc_total": 0, "dec_total": 0, "net_total": 0,
        "inc_eligible": 0, "dec_eligible": 0, "net_eligible": 0,
        "inc_withheld": 0, "dec_withheld": 0, "net_withheld": 0
    }
    all_cards = set(old_data.keys()).union(set(new_data.keys()))
    
    for card in all_cards:
        if card in old_data and card in new_data:
            old_v, new_v = old_data[card], new_data[card]
            diff_total = old_v["total"] != new_v["total"]
            diff_elig = old_v["eligible"] != new_v["eligible"]
            diff_with = old_v["withheld"] != new_v["withheld"]
            
            if diff_total or diff_elig or diff_with:
                if diff_total: 
                    counters["total_fam"] += 1
                    diff = new_v["total"] - old_v["total"]
                    counters["net_total"] += diff
                    if diff > 0: counters["inc_total"] += diff
                    else: counters["dec_total"] += abs(diff)
                if diff_elig: 
                    counters["eligible_fam"] += 1
                    diff = new_v["eligible"] - old_v["eligible"]
                    counters["net_eligible"] += diff
                    if diff > 0: counters["inc_eligible"] += diff
                    else: counters["dec_eligible"] += abs(diff)
                if diff_with: 
                    counters["withheld_fam"] += 1
                    diff = new_v["withheld"] - old_v["withheld"]
                    counters["net_withheld"] += diff
                    if diff > 0: counters["inc_withheld"] += diff
                    else: counters["dec_withheld"] += abs(diff)
                
                results.append({
                    "card": card, "seq_orig": new_v["seq"], "card_old": new_v["old_card"],
                    "name": new_v["name"], "total": new_v["total"], "eligible": new_v["eligible"],
                    "withheld": new_v["withheld"], "status": "modified"
                })
                
        elif card in old_data and card not in new_data:
            old_v = old_data[card]
            counters["deleted_fam"] += 1
            counters["dec_total"] += old_v["total"]; counters["net_total"] -= old_v["total"]
            counters["dec_eligible"] += old_v["eligible"]; counters["net_eligible"] -= old_v["eligible"]
            counters["dec_withheld"] += old_v["withheld"]; counters["net_withheld"] -= old_v["withheld"]
            
            results.append({
                "card": card, "seq_orig": old_v["seq"], "card_old": old_v["old_card"],
                "name": old_v["name"], "total": old_v["total"], "eligible": old_v["eligible"],
                "withheld": old_v["withheld"], "status": "deleted"
            })
            
        elif card not in old_data and card in new_data:
            new_v = new_data[card]
            counters["added_fam"] += 1
            counters["inc_total"] += new_v["total"]; counters["net_total"] += new_v["total"]
            counters["inc_eligible"] += new_v["eligible"]; counters["net_eligible"] += new_v["eligible"]
            counters["inc_withheld"] += new_v["withheld"]; counters["net_withheld"] += new_v["withheld"]
            
            results.append({
                "card": card, "seq_orig": new_v["seq"], "card_old": new_v["old_card"],
                "name": new_v["name"], "total": new_v["total"], "eligible": new_v["eligible"],
                "withheld": new_v["withheld"], "status": "added"
            })
    return results, counters

# -----------------------------------------------------------------------------
# منشئ PDF الاحترافي
# -----------------------------------------------------------------------------
def create_advanced_pdf_report(results, title, counters, new_data):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=landscape(A4), rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    story = []
    
    font_path = "C:\\Windows\\Fonts\\arial.ttf"
    if not os.path.exists(font_path): font_path = "arial.ttf"
    try: pdfmetrics.registerFont(TTFont('ArabicArial', font_path))
    except Exception: pass
    
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('TitleStyle', fontName='ArabicArial', fontSize=18, alignment=1, textColor=colors.HexColor('#1A365D'), spaceAfter=20)
    cell_style = ParagraphStyle('CellStyle', fontName='ArabicArial', fontSize=13, alignment=1, textColor=colors.black)
    header_style = ParagraphStyle('HeaderStyle', fontName='ArabicArial', fontSize=14, alignment=1, textColor=colors.black, fontStyle='bold')
    
    headers = ["ت", "البطاقة القديم", "اسم رب الاسرة", "التسلسل الاصلي", "الكلي", "المستحق", "المحجوب", "ملاحظة"]
    reversed_headers = headers[::-1]
    
    table_data = [[Paragraph(fix_arabic(h), header_style) for h in reversed_headers]]
    
    for idx, r in enumerate(results, start=1):
        status_txt = ""
        if r["status"] == "added": status_txt = "مضاف حديثا"
        elif r["status"] == "deleted": status_txt = "محذوف"
        else: status_txt = "معدل"
        
        row_data = [
            str(idx), str(r["card_old"]), r["name"], str(r["seq_orig"]), 
            str(r["total"]), str(r["eligible"]), str(r["withheld"]), status_txt
        ]
        row_cells = [Paragraph(fix_arabic(cell), cell_style) for cell in row_data[::-1]]
        table_data.append(row_cells)
        
    col_widths = [70, 60, 60, 60, 90, 200, 100, 40]
    
    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#EAECEE')),
        ('BACKGROUND', (4, 1), (4, -1), colors.HexColor('#FEF9E7')),
        ('BACKGROUND', (3, 1), (3, -1), colors.HexColor('#EBF5FB')),
        ('BACKGROUND', (2, 1), (2, -1), colors.HexColor('#E8F8F5')),
        ('BACKGROUND', (1, 1), (1, -1), colors.HexColor('#FDEDEC')),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor('#BDC3C7')),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
    ]))
    
    story.append(Paragraph(fix_arabic(title), title_style))
    story.append(t)
    
    doc.build(story)
    buffer.seek(0)
    return buffer

# -----------------------------------------------------------------------------
# إدارة حالة المتصفح (Session State)
# -----------------------------------------------------------------------------
if 'app_state' not in st.session_state:
    st.session_state.app_state = None

if st.session_state.app_state is not None:
    st.markdown("<hr>", unsafe_allow_html=True)
    if st.button("🔄 تصفير النظام وبدء مقارنة ملفات جديدة (اضغط هنا)", type="primary"):
        st.session_state.app_state = None
        st.rerun()

# -----------------------------------------------------------------------------
# الواجهة الرئيسية (رفع الملفات المتكامل)
# -----------------------------------------------------------------------------
if st.session_state.app_state is None:
    st.markdown("<h3 style='text-align: right;'>📂 ارفع ملفي الشهرين معاً (وسيقوم النظام بتصنيفها تلقائياً حسب التاريخ)</h3>", unsafe_allow_html=True)
    
    uploaded_files = st.file_uploader("قم بتحديد أو سحب الملفين هنا", type=['docx'], accept_multiple_files=True)
    
    old_file_determined, new_file_determined = None, None
    old_date_str, new_date_str = "", ""
    
    if uploaded_files:
        if len(uploaded_files) == 2:
            file1, file2 = uploaded_files
            date1, str1 = extract_document_date(file1)
            date2, str2 = extract_document_date(file2)
            
            if date1 and date2:
                if date1 > date2:
                    new_file_determined, old_file_determined = file1, file2
                    new_date_str, old_date_str = str1, str2
                else:
                    new_file_determined, old_file_determined = file2, file1
                    new_date_str, old_date_str = str2, str1
                    
                st.markdown(f"""
                <div class='auto-detect-box' dir='rtl'>
                    ✅ تم التعرف على التواريخ بنجاح!<br>
                    📄 <b>الملف القديم:</b> {old_file_determined.name} (بتاريخ {old_date_str})<br>
                    📄 <b>الملف الجديد:</b> {new_file_determined.name} (بتاريخ {new_date_str})
                </div>
                <br>
                """, unsafe_allow_html=True)
            else:
                st.warning("⚠️ لم يتمكن النظام من العثور على تواريخ قياسية أسفل أحد الملفين. سيتم اعتبارهما بالترتيب الذي رفعتهما به كحل بديل.")
                old_file_determined, new_file_determined = file1, file2
                
        elif len(uploaded_files) > 2:
            st.error("❌ يرجى رفع ملفين فقط.")
        else:
            st.info("⏳ بانتظار رفع الملف الثاني...")

    # زر البدء يعتمد على نجاح التحليل
    if st.button("بدء المقارنة الدقيقة وتوليد التقارير"):
        if old_file_determined and new_file_determined:
            with st.spinner('جاري قراءة الملفات ومطابقة القيود وتوليد الجداول...'):
                old_data = extract_clean_records(old_file_determined)
                new_data = extract_clean_records(new_file_determined)
                results, counters = compare_records(old_data, new_data)
                
                st.session_state.app_state = {
                    "results": sorted(results, key=lambda x: str(x.get("name", ""))),
                    "counters": counters,
                    "old_data": old_data,
                    "new_data": new_data,
                    "base_name": new_file_determined.name.rsplit('.', 1)[0]
                }
                st.rerun()
        else:
            st.warning("يرجى رفع الملفين أولاً.")

# -----------------------------------------------------------------------------
# عرض النتائج المحفوظة
# -----------------------------------------------------------------------------
if st.session_state.app_state is not None:
    state = st.session_state.app_state
    results = state["results"]
    counters = state["counters"]
    old_data = state["old_data"]
    new_data = state["new_data"]
    base_name = state["base_name"]

    if results:
        # الإحصائيات
        st.markdown("<h3 style='text-align: right; margin-top: 10px;'>📊 إحصائية الفروقات الحركية للأفراد</h3>", unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)
        with c1: 
            st.markdown(f"""<div class='report-box'>
                حركة الكلية<br><h2>{counters['total_fam']} عائلة</h2>
                <div class='stat-inc'>مضافة: +{counters['inc_total']}</div>
                <div class='stat-dec'>محذوف: -{counters['dec_total']}</div>
                <div class='net-diff'>الصافي: {counters['net_total']:+d}</div>
            </div>""", unsafe_allow_html=True)
        with c2: 
            st.markdown(f"""<div class='report-box'>
                حركة المستحقة<br><h2>{counters['eligible_fam']} عائلة</h2>
                <div class='stat-inc'>مضافة: +{counters['inc_eligible']}</div>
                <div class='stat-dec'>محذوف: -{counters['dec_eligible']}</div>
                <div class='net-diff'>الصافي: {counters['net_eligible']:+d}</div>
            </div>""", unsafe_allow_html=True)
        with c3: 
            st.markdown(f"""<div class='report-box'>
                حركة المحجوبين<br><h2>{counters['withheld_fam']} عائلة</h2>
                <div class='stat-inc'>مضافة: +{counters['inc_withheld']}</div>
                <div class='stat-dec'>محذوف: -{counters['dec_withheld']}</div>
                <div class='net-diff'>الصافي: {counters['net_withheld']:+d}</div>
            </div>""", unsafe_allow_html=True)
        with c4: 
            st.markdown(f"""<div class='report-box'>
                إضافة/حذف عوائل<br><h2>{counters['added_fam'] + counters['deleted_fam']} عائلة</h2>
                <div class='stat-inc'>تمت إضافتها: +{counters['added_fam']}</div>
                <div class='stat-dec'>تم حذفها: -{counters['deleted_fam']}</div>
                <div class='net-diff'>حركة السجلات</div>
            </div>""", unsafe_allow_html=True)

        st.markdown("<h3 style='text-align: right; color: #2C3E50; margin-top: 20px;'>📋 تفاصيل الأسماء المتغيرة (انقر على الاسم للتوسيع والمقارنة)</h3>", unsafe_allow_html=True)
        
        for r in results:
            card = r["card"]
            old_v = old_data.get(card, {"total": "-", "eligible": "-", "withheld": "-"})
            new_v = new_data.get(card, {"total": "-", "eligible": "-", "withheld": "-"})
            
            if r["status"] == "added":
                status_icon = "🟢 مضاف حديثاً"
            elif r["status"] == "deleted":
                status_icon = "🔴 محذوف من السجل"
            else:
                status_icon = "🟡 تعديل في الأفراد"
                
            with st.expander(f"{status_icon} | {r['name']} | بطاقة: {r['card_old']}"):
                st.markdown(f"""
                <table class='compare-table' dir='rtl'>
                    <tr>
                        <th style='width: 33%;'>الحقل</th>
                        <th style='width: 33%;'>في الشهر السابق (القديم)</th>
                        <th style='width: 33%;'>في الشهر الحالي (الجديد)</th>
                    </tr>
                    <tr>
                        <td><b>الأفراد الكلية</b></td>
                        <td>{old_v['total']}</td>
                        <td>{new_v['total']}</td>
                    </tr>
                    <tr>
                        <td><b>الأفراد المستحقة</b></td>
                        <td>{old_v['eligible']}</td>
                        <td>{new_v['eligible']}</td>
                    </tr>
                    <tr>
                        <td><b>الأفراد المحجوبين</b></td>
                        <td>{old_v['withheld']}</td>
                        <td>{new_v['withheld']}</td>
                    </tr>
                </table>
                """, unsafe_allow_html=True)

        st.markdown("<hr>", unsafe_allow_html=True)
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            if LIBS_READY:
                pdf_report = create_advanced_pdf_report(results, f"تقرير المتغيرات النهائي - {base_name}", counters, new_data)
                st.download_button(
                    label="📥 تحميل تقرير PDF مباشر وجاهز",
                    data=pdf_report,
                    file_name=f"متغيرات_PDF_{base_name}.pdf",
                    mime="application/pdf",
                )
            else:
                st.error("مكتبات الـ PDF غير مكتملة على الخادم. راجع المكتبات.")
        with col_dl2:
            st.info("💡 تم حفظ بياناتك محلياً في المتصفح. لتدقيق وكيل آخر، اضغط على زر 'التصفير' باللون الأحمر في الأعلى.")
            
    else:
        st.success("🎉 لا توجد فروقات أو متغيرات عددية للأفراد بين الملفين!")
