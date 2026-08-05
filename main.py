import os
import streamlit as st
import pandas as pd
from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
from docx.shared import Inches, Pt, RGBColor
import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from PIL import Image as PILImage

# صفحہ کی سیٹنگ
st.set_page_config(
    page_title="Government High School Khan Pur Maral - Paper Generator",
    page_icon="📝",
    layout="centered"
)

# شاندار CSS اسٹائلنگ
st.markdown(
    """
    <style>
    .main {
        background-color: #f8fafc;
    }
    .stSelectbox {
        background-color: #ffffff;
        border-radius: 8px;
    }
    div.stButton > button {
        background: linear-gradient(135deg, #059669 0%, #047857 100%);
        color: white;
        font-weight: bold;
        font-size: 16px;
        border-radius: 12px;
        padding: 12px 0px;
        width: 100% !important;
        border: none;
        box-shadow: 0 4px 12px rgba(5, 150, 105, 0.35);
        transition: 0.3s;
        display: block;
        margin: 0 auto;
    }
    div.stButton > button:hover {
        background: linear-gradient(135deg, #047857 0%, #065f46 100%);
        box-shadow: 0 6px 16px rgba(5, 150, 105, 0.45);
    }
    .header-box-1 {
        background: linear-gradient(135deg, #1e3a8a 0%, #2563eb 100%);
        padding: 15px;
        border-radius: 12px;
        text-align: center;
        color: white;
        box-shadow: 0 4px 10px rgba(30, 58, 138, 0.2);
        margin-bottom: 12px;
        font-family: Arial, sans-serif;
        font-weight: 700;
        font-size: 20px;
        letter-spacing: 1px;
    }
    .header-box-2 {
        background: linear-gradient(135deg, #0f172a 0%, #334155 100%);
        padding: 20px;
        border-radius: 12px;
        text-align: center;
        color: white;
        box-shadow: 0 4px 12px rgba(15, 23, 42, 0.25);
        margin-bottom: 12px;
        font-family: 'Arial Black', Gadget, sans-serif;
        font-weight: 900;
        font-size: 34px;
        letter-spacing: 2px;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
    }
    .header-box-3 {
        background: linear-gradient(135deg, #7c3aed 0%, #4f46e5 100%);
        padding: 15px;
        border-radius: 12px;
        text-align: center;
        color: white;
        box-shadow: 0 4px 10px rgba(124, 58, 237, 0.2);
        margin-bottom: 25px;
        font-family: 'Jameel Noori Nastaleeq', Arial, sans-serif;
        font-size: 20px;
        font-style: italic;
    }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div class="header-box-1">Government High School Khan Pur Maral</div>
    <div class="header-box-2">Paper Generator</div>
    <div class="header-box-3">محمد نواز شاہد 03014675646</div>
    """,
    unsafe_allow_html=True
)

excel_path = "Question_Bank.xlsx"

def load_excel_data(path):
    if not os.path.exists(path):
        return None
    try:
        xls = pd.ExcelFile(path)
        sheets = xls.sheet_names
        data_structure = {}

        for sheet in sheets:
            if "_" in sheet:
                cls, subj = sheet.split("_", 1)
                cls_name = cls + " Class"
            else:
                cls_name = "Other"
                subj = sheet

            df = pd.read_excel(path, sheet_name=sheet)
            units_data = {}
            if "Unit" in df.columns and "Topic_No" in df.columns:
                for unit_val in df["Unit"].dropna().unique():
                    topics = (
                        df[df["Unit"] == unit_val]["Topic_No"]
                        .dropna()
                        .astype(str)
                        .str.strip()
                        .unique()
                        .tolist()
                    )
                    units_data[str(unit_val).strip()] = topics

            if cls_name not in data_structure:
                data_structure[cls_name] = {}

            data_structure[cls_name][subj] = {
                "sheet": sheet,
                "units_data": units_data,
            }
        return data_structure
    except Exception as e:
        st.error(f"اكسل فائل پڑھنے میں مسئلہ پیش آیا:\n{e}")
        return None

data_structure = load_excel_data(excel_path)

if data_structure is None:
    st.error(f"مطلوبہ ایکسل فائل '{excel_path}' فولڈر میں نہیں ملی!")
else:
    if "saved_sections" not in st.session_state:
        st.session_state.saved_sections = set()
    if "section_data_records" not in st.session_state:
        st.session_state.section_data_records = {}
    if "start_process" not in st.session_state:
        st.session_state.start_process = False

    # مرکزی "Generat your paper" بٹن
    if not st.session_state.start_process:
        st.markdown("<br>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if st.button("🚀 Generat your paper"):
                st.session_state.start_process = True
                st.rerun()
    else:
        st.markdown("### 🏛️ ادارے کی معلومات اور لوگو")
        institution_name = st.text_input("اپنے ادارے کا نام لکھیں:", value="", placeholder="یہاں اپنے اسکول یا اکیڈمی کا نام لکھیں...")
        
        logo_choice = st.radio("کیا آپ اپنا لوگو بھی شامل کرنا چاہتے ہیں؟", ["نہیں (No)", "ہاں (Yes)"], horizontal=True)
        
        uploaded_logo = None
        proceed_to_classes = True

        if logo_choice == "ہاں (Yes)":
            uploaded_logo = st.file_uploader("لوگو فائل اپ لوڈ کریں (PNG یا JPG):", type=["png", "jpg", "jpeg"])
            if uploaded_logo is None:
                proceed_to_classes = False
                st.info("💡 براہ کرم آگے بڑھنے کے لیے لوگو اپ لوڈ کریں یا 'نہیں (No)' منتخب کریں۔")

        if proceed_to_classes:
            st.markdown("---")
            
            classes = list(data_structure.keys())
            selected_cls = st.selectbox("1. کلاس منتخب کریں:", [""] + classes, key="sel_cls")

            selected_subj = ""
            if selected_cls != "":
                subj_list = list(data_structure[selected_cls].keys())
                selected_subj = st.selectbox("2. مضمون منتخب کریں:", [""] + subj_list, key="sel_subj")

            selected_mode = ""
            if selected_subj != "":
                selected_mode = st.selectbox(
                    "3. پیپر موڈ منتخب کریں:",
                    ["", "سنگل یونٹ (Single Unit)", "ملٹی یونٹ (Multi Unit)", "فل بک (Full Book)"],
                    key="sel_mode"
                )

            selected_unit = ""
            selected_topics = []
            if selected_mode != "" and "سنگل یونٹ" in selected_mode:
                units_data = data_structure[selected_cls][selected_subj]["units_data"]
                unit_str_list = [str(u) for u in units_data.keys()]
                selected_unit = st.selectbox("4. یونٹ منتخب کریں:", [""] + unit_str_list, key="sel_unit")

                if selected_unit != "":
                    topics = units_data.get(str(selected_unit).strip(), [])
                    if topics:
                        st.markdown("##### 📌 ٹاپکس منتخب کریں:")
                        select_all = st.checkbox("سلیکٹ آل (Select All)", key="chk_select_all")
                        for topic in topics:
                            t_key = f"top_{selected_cls}_{selected_subj}_{selected_unit}_{topic}"
                            is_checked = st.checkbox(str(topic), value=select_all, key=t_key)
                            if is_checked:
                                selected_topics.append(topic)

            ready_to_show = False
            if selected_cls != "" and selected_subj != "" and selected_mode != "":
                if "سنگل یونٹ" in selected_mode:
                    if selected_unit != "" and len(selected_topics) > 0:
                        ready_to_show = True
                else:
                    ready_to_show = True

            if ready_to_show:
                st.markdown("---")
                st.markdown("### 📚 سوالات کے سیکشنز اور انتخاب")

                sheet_info = data_structure[selected_cls][selected_subj]
                actual_sheet = sheet_info["sheet"]

                try:
                    df_full = pd.read_excel(excel_path, sheet_name=actual_sheet)

                    section_types = [
                        (1, "ایم سی کیوز (MCQs - کثیر الانتخابی سوالات)"),
                        (2, "مختصر سوالات (Short Questions - SQ)"),
                        (5, "طویل سوالات (Long Questions - LQ)")
                    ]

                    for q_id, title_text in section_types:
                        with st.expander(f"{title_text}"):
                            filtered_df = df_full.copy()
                            if "Q_Type" in filtered_df.columns:
                                filtered_df["Q_Type"] = filtered_df["Q_Type"].astype(str).str.strip().str.upper()
                                if q_id == 1:
                                    filtered_df = filtered_df[filtered_df["Q_Type"] == "MCQ"]
                                elif q_id == 2:
                                    filtered_df = filtered_df[filtered_df["Q_Type"] == "SQ"]
                                elif q_id == 5:
                                    filtered_df = filtered_df[filtered_df["Q_Type"] == "LQ"]

                            if selected_unit and "سنگل یونٹ" in selected_mode:
                                if "Unit" in filtered_df.columns:
                                    filtered_df = filtered_df[filtered_df["Unit"].astype(str).str.strip() == str(selected_unit).strip()]
                                if selected_topics and "Topic_No" in filtered_df.columns:
                                    filtered_df["Topic_No_Str"] = filtered_df["Topic_No"].astype(str).str.strip()
                                    filtered_df = filtered_df[filtered_df["Topic_No_Str"].isin([str(t) for t in selected_topics])]

                            selected_records_for_q = []
                            for idx, row in enumerate(filtered_df.iterrows(), 1):
                                r_data = row[1]
                                q_text = str(r_data.get("Q_Text", "")) if pd.notna(r_data.get("Q_Text")) else ""
                                expr = str(r_data.get("Expression", "")) if pd.notna(r_data.get("Expression")) else ""
                                display_str = f"{idx}. {expr} {q_text}".strip()

                                if q_id == 1:
                                    opt_a = str(r_data.get("Option_A", "")) if pd.notna(r_data.get("Option_A")) else ""
                                    opt_b = str(r_data.get("Option_B", "")) if pd.notna(r_data.get("Option_B")) else ""
                                    opt_c = str(r_data.get("Option_C", "")) if pd.notna(r_data.get("Option_C")) else ""
                                    opt_d = str(r_data.get("Option_D", "")) if pd.notna(r_data.get("Option_D")) else ""
                                    if opt_a or opt_b or opt_c or opt_d:
                                        display_str += f"\n   (A) {opt_a}    (B) {opt_b}    (C) {opt_c}    (D) {opt_d}"

                                if st.checkbox(display_str, key=f"q_{q_id}_{idx}"):
                                    selected_records_for_q.append(r_data)

                            if st.button(f"محفوظ کریں ({title_text})", key=f"save_btn_{q_id}"):
                                st.session_state.section_data_records[q_id] = selected_records_for_q
                                st.session_state.saved_sections.add(q_id)
                                st.success(f"✅ {len(selected_records_for_q)} سوالات کامیابی سے محفوظ ہو گئے ہیں!")

                    st.markdown("<br>", unsafe_allow_html=True)

                    if st.button("📥 Generate Paper (ورڈ میں مکمل پرچہ بنائیں)"):
                        if not st.session_state.section_data_records:
                            st.warning("⚠️ براہ کرم پہلے کم از کم کسی ایک سیکشن کو حل کر کے محفوظ کریں!")
                        else:
                            try:
                                doc = Document()
                                
                                for section in doc.sections:
                                    section.page_width = Inches(8.27)
                                    section.page_height = Inches(11.69)
                                    section.top_margin = Inches(0.4)
                                    section.bottom_margin = Inches(0.4)
                                    section.left_margin = Inches(0.4)
                                    section.right_margin = Inches(0.4)
                                    
                                    body = section._sectPr
                                    bidi = OxmlElement('w:bidi')
                                    body.append(bidi)

                                style = doc.styles['Normal']
                                font = style.font
                                font.name = 'Jameel Noori Nastaleeq'
                                font.size = Pt(11)

                                total_paper_marks = 0
                                for q_idx, records in st.session_state.section_data_records.items():
                                    cnt = len(records)
                                    if q_idx == 1:
                                        total_paper_marks += cnt * 1
                                    elif q_idx == 2:
                                        total_paper_marks += cnt * 2
                                    elif q_idx == 5:
                                        total_paper_marks += cnt * 4

                                raw_minutes = total_paper_marks * 1.5
                                remainder = raw_minutes % 10
                                if remainder >= 5:
                                    total_minutes = int(raw_minutes + (10 - remainder))
                                else:
                                    total_minutes = int(raw_minutes - remainder)

                                sel_class_display = selected_cls.replace(" Class", "").strip()
                                sel_subj = selected_subj

                                class_urdu_map = {
                                    "9th": "نہم", "10th": "دہم",
                                    "8th": "آٹھویں", "7th": "ساتویں"
                                }
                                sel_class_urdu = class_urdu_map.get(sel_class_display, sel_class_display)

                                subject_urdu_map = {
                                    "Chemistry": "کیمسٹری", "Physics": "فزکس", "Biology": "بائیولوجی",
                                    "Mathematics": "ریاضی", "Math": "ریاضی", "English": "انگریزی", 
                                    "Urdu": "اردو", "Islamiyat": "اسلامیات", "Pak_Studies": "مطالعہ پاکستان", 
                                    "Computer": "کمپیوٹر سائنس"
                                }
                                sel_subj_urdu = subject_urdu_map.get(sel_subj, sel_subj)
                                sel_unit_val = selected_unit if selected_unit else "Full_Book"

                                # --- پہلا ہیڈر ٹیبل (۸ کالمز: لوگو بائیں، سکول نام درمیان، مضمون/کلاس دائیں) ---
                                head_table = doc.add_table(rows=1, cols=8)
                                head_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                                head_table.autofit = False

                                col_width = Inches(7.47 / 8.0)
                                for i in range(8):
                                    head_table.columns[i].width = col_width

                                trPr = head_table.rows[0]._tr.get_or_add_trPr()
                                trHeight = parse_xml(r'<w:trHeight {} w:val="450" w:hRule="atLeast"/>'.format(nsdecls('w')))
                                trPr.append(trHeight)

                                cells = head_table.rows[0].cells
                                cell_left = cells[0]
                                cell_mid = cells[1].merge(cells[6])
                                cell_right = cells[7]

                                # [0] بائیں سیل (لوگو کے لیے)
                                cell_left.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER
                                p_left = cell_left.paragraphs[0]
                                p_left.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                                p_left.paragraph_format.bidi = True
                                p_left.paragraph_format.space_after = Pt(0)
                                p_left.paragraph_format.space_before = Pt(0)

                                if uploaded_logo is not None:
                                    temp_logo_path = "temp_uploaded_logo.png"
                                    with open(temp_logo_path, "wb") as f:
                                        f.write(uploaded_logo.getbuffer())
                                    p_left.add_run().add_picture(temp_logo_path, width=Inches(0.45))

                                # [1-6] درمیان والا مرج شدہ سیل (اسکول کا نام)
                                cell_mid.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER
                                p_mid = cell_mid.paragraphs[0]
                                p_mid.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                                p_mid.paragraph_format.bidi = True
                                p_mid.paragraph_format.space_after = Pt(0)
                                p_mid.paragraph_format.space_before = Pt(0)

                                final_inst_name = institution_name.strip() if institution_name.strip() else "GOVERNMENT HIGH SCHOOL KHAN PUR MARAL MULTAN"
                                run_m1 = p_mid.add_run(f"{final_inst_name.upper()}\n")
                                run_m1.font.name = 'Arial'
                                run_m1.font.size = Pt(11)
                                run_m1.bold = True
                                run_m1.font.color.rgb = RGBColor(30, 58, 138)

                                run_m2 = p_mid.add_run("TOPIC WISE TEST")
                                run_m2.font.name = 'Arial'
                                run_m2.font.size = Pt(8.5)
                                run_m2.bold = True
                                run_m2.font.color.rgb = RGBColor(100, 100, 100)

                                # [7] دائیں سیل (مضمون اور کلاس)
                                cell_right.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER
                                p_right = cell_right.paragraphs[0]
                                p_right.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                                p_right.paragraph_format.bidi = True
                                p_right.paragraph_format.space_after = Pt(0)
                                p_right.paragraph_format.space_before = Pt(0)
                                
                                run_r = p_right.add_run(f"{sel_subj_urdu} ({sel_class_urdu})")
                                run_r.font.name = 'Jameel Noori Nastaleeq'
                                run_r.font.size = Pt(12)
                                run_r.bold = True
                                run_r.font.color.rgb = RGBColor(30, 58, 138)

                                for c in [cell_left, cell_mid, cell_right]:
                                    shd = parse_xml(r'<w:shd {} w:fill="F1F5F9"/>'.format(nsdecls('w')))
                                    c._tc.get_or_add_tcPr().append(shd)
                                    
                                    tcMar = parse_xml(r'''
                                        <w:tcMar {} >
                                            <w:top w:w="60" w:type="dxa"/>
                                            <w:bottom w:w="60" w:type="dxa"/>
                                            <w:left w:w="60" w:type="dxa"/>
                                            <w:right w:w="60" w:type="dxa"/>
                                        </w:tcMar>
                                    '''.format(nsdecls('w')))
                                    c._tc.get_or_add_tcPr().append(tcMar)

                                    tcBorders = parse_xml(r'''
                                        <w:tcBorders {} >
                                            <w:top w:val="single" w:sz="8" w:space="0" w:color="1E3A8A"/>
                                            <w:left w:val="single" w:sz="8" w:space="0" w:color="1E3A8A"/>
                                            <w:bottom w:val="single" w:sz="8" w:space="0" w:color="1E3A8A"/>
                                            <w:right w:val="single" w:sz="8" w:space="0" w:color="1E3A8A"/>
                                        </w:tcBorders>
                                    '''.format(nsdecls('w')))
                                    c._tc.get_or_add_tcPr().append(tcBorders)

                                # --- دوسرا ہیڈر ٹیبل (معلومات: نمبرز، یونٹ، ٹاپکس، وقت) ---
                                topics_str_formatted = " + ".join([str(t) for t in selected_topics]) if selected_topics else "All Topics"
                                unit_info_str = f"Unit: {sel_unit_val}" if sel_unit_val != "Full_Book" else "Full Book"
                                
                                info_table = doc.add_table(rows=1, cols=4)
                                info_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                                info_table.autofit = False
                                
                                info_widths = [Inches(1.4), Inches(1.5), Inches(3.2), Inches(1.4)]
                                for idx, w in enumerate(info_widths):
                                    info_table.columns[idx].width = w

                                trPr_info = info_table.rows[0]._tr.get_or_add_trPr()
                                trHeight_info = parse_xml(r'<w:trHeight {} w:val="380" w:hRule="atLeast"/>'.format(nsdecls('w')))
                                trPr_info.append(trHeight_info)

                                info_cells = info_table.rows[0].cells
                                info_texts = [
                                    f"Marks: {total_paper_marks}",
                                    unit_info_str,
                                    f"Topics: {topics_str_formatted}",
                                    f"Time: {total_minutes} Mins"
                                ]

                                for idx, text in enumerate(info_texts):
                                    c = info_cells[idx]
                                    c.width = info_widths[idx]
                                    c.vertical_alignment = docx.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER
                                    p_c = c.paragraphs[0]
                                    p_c.alignment = docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER
                                    p_c.paragraph_format.space_after = Pt(0)
                                    p_c.paragraph_format.space_before = Pt(0)
                                    
                                    run_c = p_c.add_run(text)
                                    run_c.font.name = 'Arial'
                                    run_c.font.size = Pt(9.5)
                                    run_c.bold = True
                                    run_c.font.color.rgb = RGBColor(15, 23, 42)
                                    
                                    shd_info = parse_xml(r'<w:shd {} w:fill="E2E8F0"/>'.format(nsdecls('w')))
                                    c._tc.get_or_add_tcPr().append(shd_info)

                                    tcMar_info = parse_xml(r'''
                                        <w:tcMar {} >
                                            <w:top w:w="50" w:type="dxa"/>
                                            <w:bottom w:w="50" w:type="dxa"/>
                                            <w:left w:w="50" w:type="dxa"/>
                                            <w:right w:w="50" w:type="dxa"/>
                                        </w:tcMar>
                                    '''.format(nsdecls('w')))
                                    c._tc.get_or_add_tcPr().append(tcMar_info)

                                    tcBorders_info = parse_xml(r'''
                                        <w:tcBorders {} >
                                            <w:top w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>
                                            <w:left w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>
                                            <w:bottom w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>
                                            <w:right w:val="single" w:sz="6" w:space="0" w:color="CBD5E1"/>
                                        </w:tcBorders>
                                    '''.format(nsdecls('w')))
                                    c._tc.get_or_add_tcPr().append(tcBorders_info)

                                # --- حتمی اور فرکشن پاور پارسر ---
                                def parse_math_token(parent_element, token_str):
                                    t = token_str.strip()
                                    if not t:
                                        return

                                    if '^' in t:
                                        split_idx = t.find('^')
                                        base_str = t[:split_idx].strip()
                                        rest_after_caret = t[split_idx+1:].strip()
                                        
                                        exp_str = ""
                                        if rest_after_caret.startswith('(') and rest_after_caret.endswith(')'):
                                            exp_str = rest_after_caret[1:-1].strip()
                                        elif rest_after_caret.startswith('('):
                                            b_count = 0
                                            end_b_idx = -1
                                            for b_i, b_char in enumerate(rest_after_caret):
                                                if b_char == '(':
                                                    b_count += 1
                                                elif b_char == ')':
                                                    b_count -= 1
                                                    if b_count == 0:
                                                        end_b_idx = b_i
                                                        break
                                            if end_b_idx != -1:
                                                exp_str = rest_after_caret[1:end_b_idx].strip()
                                            else:
                                                exp_str = rest_after_caret.strip()
                                        else:
                                            exp_str = rest_after_caret.strip()

                                        m_sSup = OxmlElement('m:sSup')
                                        m_e = OxmlElement('m:e')
                                        m_sup = OxmlElement('m:sup')
                                        
                                        parse_sub_expression(m_e, base_str)
                                        
                                        if '/' in exp_str and not exp_str.lower().startswith('sqrt'):
                                            f_parts = exp_str.split('/', 1)
                                            num_p = f_parts[0].strip()
                                            den_p = f_parts[1].strip()
                                            
                                            m_f = OxmlElement('m:f')
                                            m_num = OxmlElement('m:num')
                                            m_den = OxmlElement('m:den')
                                            
                                            parse_sub_expression(m_num, num_p)
                                            parse_sub_expression(m_den, den_p)
                                            
                                            m_f.append(m_num)
                                            m_f.append(m_den)
                                            m_sup.append(m_f)
                                        else:
                                            parse_sub_expression(m_sup, exp_str)
                                        
                                        m_sSup.append(m_e)
                                        m_sSup.append(m_sup)
                                        parent_element.append(m_sSup)
                                        return

                                    if t.startswith('(') and t.endswith(')'):
                                        inner_content = t[1:-1].strip()
                                        m_d = OxmlElement('m:d')
                                        m_e = OxmlElement('m:e')
                                        parse_sub_expression(m_e, inner_content)
                                        m_d.append(m_e)
                                        parent_element.append(m_d)
                                        return

                                    if '/' in t and not t.lower().startswith('sqrt') and not t.lower().startswith('log') and not t.lower().startswith('ln'):
                                        bracket_lvl = 0
                                        split_idx = -1
                                        for idx_c, char_c in enumerate(t):
                                            if char_c == '(':
                                                bracket_lvl += 1
                                            elif char_c == ')':
                                                bracket_lvl -= 1
                                            elif char_c == '/' and bracket_lvl == 0:
                                                split_idx = idx_c
                                                break
                                        
                                        if split_idx != -1:
                                            num_str = t[:split_idx].strip()
                                            den_str = t[split_idx+1:].strip()
                                            
                                            if num_str.startswith('(') and num_str.endswith(')'):
                                                num_str = num_str[1:-1].strip()
                                            if den_str.startswith('(') and den_str.endswith(')'):
                                                den_str = den_str[1:-1].strip()
                                            
                                            m_f = OxmlElement('m:f')
                                            m_num = OxmlElement('m:num')
                                            m_den = OxmlElement('m:den')
                                            
                                            parse_sub_expression(m_num, num_str)
                                            parse_sub_expression(m_den, den_str)
                                            
                                            m_f.append(m_num)
                                            m_f.append(m_den)
                                            parent_element.append(m_f)
                                            return

                                    if 'sqrt' in t.lower():
                                        idx_sqrt = t.lower().find('sqrt')
                                        if idx_sqrt > 0:
                                            prefix = t[:idx_sqrt].strip()
                                            if prefix:
                                                m_r = OxmlElement('m:r')
                                                m_t = OxmlElement('m:t')
                                                m_t.text = prefix
                                                m_r.append(m_t)
                                                parent_element.append(m_r)
                                        
                                        rest = t[idx_sqrt + 4:].strip()
                                        if rest.startswith('('):
                                            open_brackets = 0
                                            split_idx = -1
                                            for idx_char, char_val in enumerate(rest):
                                                if char_val == '(':
                                                    open_brackets += 1
                                                elif char_val == ')':
                                                    open_brackets -= 1
                                                    if open_brackets == 0:
                                                        split_idx = idx_char
                                                        break
                                            if split_idx != -1:
                                                inner = rest[1:split_idx].strip()
                                                deg_val = ""
                                                if ',' in inner:
                                                    p_in = inner.split(',')
                                                    deg_val = p_in[-1].strip()
                                                    inner = ",".join(p_in[:-1]).strip()
                                                
                                                m_rad = OxmlElement('m:rad')
                                                m_deg = OxmlElement('m:deg')
                                                if deg_val:
                                                    m_r_deg = OxmlElement('m:r')
                                                    m_t_deg = OxmlElement('m:t')
                                                    m_t_deg.text = deg_val
                                                    m_r_deg.append(m_t_deg)
                                                    m_deg.append(m_r_deg)
                                                
                                                m_rad.append(m_deg)
                                                m_e = OxmlElement('m:e')
                                                parse_sub_expression(m_e, inner)
                                                m_rad.append(m_e)
                                                parent_element.append(m_rad)
                                                
                                                remainder_text = rest[split_idx+1:].strip()
                                                if remainder_text:
                                                    parse_sub_expression(parent_element, remainder_text)
                                            else:
                                                inner = rest[1:].strip()
                                                m_rad = OxmlElement('m:rad')
                                                m_deg = OxmlElement('m:deg')
                                                m_rad.append(m_deg)
                                                m_e = OxmlElement('m:e')
                                                parse_sub_expression(m_e, inner)
                                                m_rad.append(m_e)
                                                parent_element.append(m_rad)
                                        else:
                                            m_r = OxmlElement('m:r')
                                            m_t = OxmlElement('m:t')
                                            m_t.text = t
                                            m_r.append(m_t)
                                            parent_element.append(m_r)
                                    elif ('log' in t.lower() or 'ln' in t.lower()) and '(' in t:
                                        op_type = 'ln' if 'ln' in t.lower() else 'log'
                                        idx_l = t.lower().find(op_type)
                                        prefix = t[:idx_l].strip()
                                        if prefix:
                                            parse_sub_expression(parent_element, prefix)
                                        
                                        rest = t[idx_l + len(op_type):].strip()
                                        
                                        base_val = ""
                                        if op_type == 'log' and rest and rest[0].isdigit():
                                            match_digits = ""
                                            for ch in rest:
                                                if ch.isdigit():
                                                    match_digits += ch
                                                else:
                                                    break
                                            if match_digits:
                                                base_val = match_digits
                                                rest = rest[len(match_digits):].strip()

                                        if rest.startswith('('):
                                            open_brackets = 0
                                            split_idx = -1
                                            for idx_char, char_val in enumerate(rest):
                                                if char_val == '(':
                                                    open_brackets += 1
                                                elif char_val == ')':
                                                    open_brackets -= 1
                                                    if open_brackets == 0:
                                                        split_idx = idx_char
                                                        break
                                            if split_idx != -1:
                                                inner = rest[1:split_idx].strip()
                                                
                                                if ',' in inner:
                                                    parts_log = inner.split(',')
                                                    base_val = parts_log[-1].strip().replace("base=", "").strip()
                                                    inner = ",".join(parts_log[:-1]).strip()

                                                if op_type == 'log':
                                                    if base_val:
                                                        m_sSub = OxmlElement('m:sSub')
                                                        m_e_base = OxmlElement('m:e')
                                                        m_r_log = OxmlElement('m:r')
                                                        m_t_log = OxmlElement('m:t')
                                                        m_t_log.text = "log"
                                                        m_r_log.append(m_t_log)
                                                        m_e_base.append(m_r_log)
                                                        m_sSub.append(m_e_base)
                                                        
                                                        m_sub_elem = OxmlElement('m:sub')
                                                        m_r_sub = OxmlElement('m:r')
                                                        m_t_sub = OxmlElement('m:t')
                                                        m_t_sub.text = base_val
                                                        m_r_sub.append(m_t_sub)
                                                        m_sub_elem.append(m_r_sub)
                                                        m_sSub.append(m_sub_elem)
                                                        parent_element.append(m_sSub)
                                                    else:
                                                        m_r_log = OxmlElement('m:r')
                                                        m_t_log = OxmlElement('m:t')
                                                        m_t_log.text = "log"
                                                        m_r_log.append(m_t_log)
                                                        parent_element.append(m_r_log)
                                                else:
                                                    m_r_ln = OxmlElement('m:r')
                                                    m_t_ln = OxmlElement('m:t')
                                                    m_t_ln.text = "ln"
                                                    m_r_ln.append(m_t_ln)
                                                    parent_element.append(m_r_ln)

                                                parse_sub_expression(parent_element, inner)
                                                
                                                remainder = rest[split_idx+1:].strip()
                                                if remainder:
                                                    parse_sub_expression(parent_element, remainder)
                                    else:
                                        m_r = OxmlElement('m:r')
                                        m_t = OxmlElement('m:t')
                                        m_t.text = t
                                        m_r.append(m_t)
                                        parent_element.append(m_r)

                                def parse_sub_expression(parent_element, expr_str):
                                    s = expr_str.strip()
                                    if not s:
                                        return
                                    
                                    if '=' in s:
                                        parts = s.split('=', 1)
                                        left_expr = parts[0].strip()
                                        right_expr = parts[1].strip()
                                        
                                        parse_sub_expression(parent_element, left_expr)
                                        
                                        m_r_eq = OxmlElement('m:r')
                                        m_t_eq = OxmlElement('m:t')
                                        m_t_eq.text = " = "
                                        m_r_eq.append(m_t_eq)
                                        parent_element.append(m_r_eq)
                                        
                                        parse_sub_expression(parent_element, right_expr)
                                        return

                                    if '*' in s:
                                        parts = s.split('*')
                                        for p in parts:
                                            parse_sub_expression(parent_element, p.strip())
                                        return

                                    if '^' in s and not s.lower().startswith('sqrt') and not s.lower().startswith('log') and not s.lower().startswith('ln'):
                                        split_idx = s.find('^')
                                        base_str = s[:split_idx].strip()
                                        rest_after_caret = s[split_idx+1:].strip()
                                        
                                        exp_str = ""
                                        if rest_after_caret.startswith('(') and rest_after_caret.endswith(')'):
                                            exp_str = rest_after_caret[1:-1].strip()
                                        elif rest_after_caret.startswith('('):
                                            b_count = 0
                                            end_b_idx = -1
                                            for b_i, b_char in enumerate(rest_after_caret):
                                                if b_char == '(':
                                                    b_count += 1
                                                elif b_char == ')':
                                                    b_count -= 1
                                                    if b_count == 0:
                                                        end_b_idx = b_i
                                                        break
                                            if end_b_idx != -1:
                                                exp_str = rest_after_caret[1:end_b_idx].strip()
                                            else:
                                                exp_str = rest_after_caret.strip()
                                        else:
                                            exp_str = rest_after_caret.strip()

                                        m_sSup = OxmlElement('m:sSup')
                                        m_e = OxmlElement('m:e')
                                        m_sup = OxmlElement('m:sup')
                                        parse_sub_expression(m_e, base_str)
                                        
                                        if '/' in exp_str and not exp_str.lower().startswith('sqrt'):
                                            f_parts = exp_str.split('/', 1)
                                            num_p = f_parts[0].strip()
                                            den_p = f_parts[1].strip()
                                            
                                            m_f = OxmlElement('m:f')
                                            m_num = OxmlElement('m:num')
                                            m_den = OxmlElement('m:den')
                                            
                                            parse_sub_expression(m_num, num_p)
                                            parse_sub_expression(m_den, den_p)
                                            
                                            m_f.append(m_num)
                                            m_f.append(m_den)
                                            m_sup.append(m_f)
                                        else:
                                            parse_sub_expression(m_sup, exp_str)
                                        
                                        m_sSup.append(m_e)
                                        m_sSup.append(m_sup)
                                        parent_element.append(m_sSup)
                                        return

                                    if '/' in s and not s.lower().startswith('sqrt') and not s.lower().startswith('log') and not s.lower().startswith('ln'):
                                        bracket_lvl = 0
                                        split_idx = -1
                                        for idx_c, char_c in enumerate(s):
                                            if char_c == '(':
                                                bracket_lvl += 1
                                            elif char_c == ')':
                                                bracket_lvl -= 1
                                            elif char_c == '/' and bracket_lvl == 0:
                                                split_idx = idx_c
                                                break
                                        
                                        if split_idx != -1:
                                            num_str = s[:split_idx].strip()
                                            den_str = s[split_idx+1:].strip()
                                            
                                            if num_str.startswith('(') and num_str.endswith(')'):
                                                num_str = num_str[1:-1].strip()
                                            if den_str.startswith('(') and den_str.endswith(')'):
                                                den_str = den_str[1:-1].strip()
                                            
                                            m_f = OxmlElement('m:f')
                                            m_num = OxmlElement('m:num')
                                            m_den = OxmlElement('m:den')
                                            
                                            parse_sub_expression(m_num, num_str)
                                            parse_sub_expression(m_den, den_str)
                                            
                                            m_f.append(m_num)
                                            m_f.append(m_den)
                                            parent_element.append(m_f)
                                            return

                                    if '+' in s and not s.lower().startswith('sqrt') and not s.lower().startswith('log') and not s.lower().startswith('ln'):
                                        parts = []
                                        current_part = ""
                                        bracket_count = 0
                                        for char_item in s:
                                            if char_item == '(':
                                                bracket_count += 1
                                                current_part += char_item
                                            elif char_item == ')':
                                                bracket_count -= 1
                                                current_part += char_item
                                            elif char_item == '+' and bracket_count == 0:
                                                parts.append(current_part)
                                                current_part = ""
                                            else:
                                                current_part += char_item
                                        if current_part:
                                            parts.append(current_part)
                                        
                                        for i, p in enumerate(parts):
                                            if i > 0:
                                                m_op = OxmlElement('m:r')
                                                m_opt = OxmlElement('m:t')
                                                m_opt.text = '+'
                                                m_op.append(m_opt)
                                                parent_element.append(m_op)
                                            parse_math_token(parent_element, p)
                                    else:
                                        parse_math_token(parent_element, s)

                                def append_smart_text_or_math(paragraph, text_str):
                                    if not text_str or not str(text_str).strip():
                                        return

                                    text_val = str(text_str).strip()
                                    has_math = any(op in text_val.lower() for op in ['sqrt', '/', '^', 'log', 'ln', '*'])

                                    if has_math:
                                        m_math = OxmlElement('m:oMath')
                                        parse_sub_expression(m_math, text_val)
                                        paragraph._p.append(m_math)
                                    else:
                                        run = paragraph.add_run(text_val)
                                        run.font.name = 'Jameel Noori Nastaleeq'

                                for q_idx in sorted(st.session_state.section_data_records.keys()):
                                    if st.session_state.section_data_records[q_idx]:
                                        p_heading = doc.add_paragraph()
                                        p_heading.paragraph_format.bidi = True
                                        p_heading.paragraph_format.space_before = Pt(2)
                                        p_heading.paragraph_format.space_after = Pt(0)
                                        
                                        run_h = p_heading.add_run()
                                        run_h.font.name = 'Jameel Noori Nastaleeq'
                                        run_h.bold = True
                                        run_h.font.size = Pt(11)
                                        
                                        q_count = len(st.session_state.section_data_records[q_idx])

                                        if q_idx == 1:
                                            run_h.text = f"سوال نمبر 1\tدرست جواب پر دائرہ لگائیں۔\t({q_count} x 1 = {q_count}):"
                                            p_heading.paragraph_format.tab_stops.add_tab_stop(Inches(0.5))
                                            p_heading.paragraph_format.tab_stops.add_tab_stop(Inches(3.5))
                                        elif q_idx == 2:
                                            total_marks = q_count * 2
                                            run_h.text = f"سوال نمبر 2\tدرج ذیل مختصر سوالات کے جوابات دیں۔\t({q_count} x 2 = {total_marks}):"
                                            p_heading.paragraph_format.tab_stops.add_tab_stop(Inches(0.5))
                                            p_heading.paragraph_format.tab_stops.add_tab_stop(Inches(3.5))
                                        elif q_idx == 5:
                                            total_marks = q_count * 4
                                            run_h.text = f"سوال نمبر 3\tدرج ذیل طویل سوالات کے جوابات دیں۔\t({q_count} x 4 = {total_marks}):"
                                            p_heading.paragraph_format.tab_stops.add_tab_stop(Inches(0.5))
                                            p_heading.paragraph_format.tab_stops.add_tab_stop(Inches(3.5))

                                        pPr_h = p_heading._p.get_or_add_pPr()
                                        pBdr_h = OxmlElement('w:pBdr')
                                        bottom_h = OxmlElement('w:bottom')
                                        bottom_h.set(qn('w:val'), 'single')
                                        bottom_h.set(qn('w:sz'), '6')
                                        bottom_h.set(qn('w:space'), '1')
                                        bottom_h.set(qn('w:color'), '000000')
                                        pBdr_h.append(bottom_h)
                                        pPr_h.append(pBdr_h)

                                        for sub_i, r_data in enumerate(st.session_state.section_data_records[q_idx], 1):
                                            p_q = doc.add_paragraph()
                                            p_q.paragraph_format.bidi = True
                                            p_q.paragraph_format.space_before = Pt(0)
                                            p_q.paragraph_format.space_after = Pt(0)
                                            
                                            run_num = p_q.add_run(f"({sub_i})\t")
                                            run_num.font.name = 'Jameel Noori Nastaleeq'
                                            run_num.bold = True
                                            p_q.paragraph_format.tab_stops.add_tab_stop(Inches(0.8))

                                            expr = r_data.get("Expression", "")
                                            if pd.notna(expr) and str(expr).strip() != "":
                                                append_smart_text_or_math(p_q, str(expr))

                                            q_text_prefix = ""
                                            if pd.notna(r_data.get("Q_Text")):
                                                q_text_prefix += str(r_data["Q_Text"])
                                            
                                            if q_text_prefix:
                                                run_qt = p_q.add_run(f"   {q_text_prefix}")
                                                run_qt.font.name = 'Jameel Noori Nastaleeq'

                                            if q_idx == 1:
                                                opt_a = str(r_data.get("Option_A", "")) if pd.notna(r_data.get("Option_A")) else ""
                                                opt_b = str(r_data.get("Option_B", "")) if pd.notna(r_data.get("Option_B")) else ""
                                                opt_c = str(r_data.get("Option_C", "")) if pd.notna(r_data.get("Option_C")) else ""
                                                opt_d = str(r_data.get("Option_D", "")) if pd.notna(r_data.get("Option_D")) else ""

                                                if opt_a or opt_b or opt_c or opt_d:
                                                    table = doc.add_table(rows=1, cols=8)
                                                    table.alignment = docx.enum.table.WD_TABLE_ALIGNMENT.CENTER
                                                    
                                                    tblPr = table._tbl.tblPr
                                                    bidi = OxmlElement('w:bidiVisual')
                                                    tblPr.append(bidi)
                                                    
                                                    col_widths = [Inches(0.4), Inches(2.0), Inches(0.4), Inches(2.0), Inches(0.4), Inches(2.0), Inches(0.4), Inches(2.0)]
                                                    for i, width in enumerate(col_widths):
                                                        table.columns[i].width = width

                                                    hdr_cells = table.rows[0].cells
                                                    options_data = [("(A)", opt_a), ("(B)", opt_b), ("(C)", opt_c), ("(D)", opt_d)]
                                                    
                                                    for idx, (label, val) in enumerate(options_data):
                                                        cell_lbl = hdr_cells[idx * 2]
                                                        cell_val = hdr_cells[idx * 2 + 1]
                                                        
                                                        cell_lbl.text = label
                                                        
                                                        if val and str(val).strip():
                                                            cell_val.text = ""
                                                            p_val = cell_val.paragraphs[0]
                                                            p_val.paragraph_format.bidi = True
                                                            p_val.paragraph_format.space_after = Pt(0)
                                                            p_val.paragraph_format.space_before = Pt(0)
                                                            append_smart_text_or_math(p_val, str(val))
                                                        else:
                                                            cell_val.text = ""

                                                        for p in cell_lbl.paragraphs:
                                                            p.paragraph_format.bidi = True
                                                            p.paragraph_format.space_after = Pt(0)
                                                            p.paragraph_format.space_before = Pt(0)
                                                            for run in p.runs:
                                                                run.font.name = 'Jameel Noori Nastaleeq'
                                                        
                                                        cell_lbl.width = col_widths[idx * 2]
                                                        cell_val.width = col_widths[idx * 2 + 1]

                                target_folder = r"F:\Automated_Paper_App"
                                os.makedirs(target_folder, exist_ok=True)

                                if selected_topics:
                                    topics_str = "+".join(selected_topics)
                                    file_naming_part = f"Unit_{sel_unit_val}_Topics_{topics_str}"
                                elif sel_unit_val and sel_unit_val != "Full_Book":
                                    file_naming_part = f"Unit_{sel_unit_val}"
                                else:
                                    file_naming_part = f"Full_Book"

                                for char in ['\\', '/', ':', '*', '?', '"', '<', '>', '|']:
                                    file_naming_part = file_naming_part.replace(char, '')

                                file_name = f"{sel_class_display}_{sel_subj}_{file_naming_part}.docx".replace(" ", "_")
                                file_path = os.path.join(target_folder, file_name)

                                doc.save(file_path)

                                with open(file_path, "rb") as f:
                                    btn = st.download_button(
                                        label="📥 Download Word Paper File (فائل ڈاؤن لوڈ کریں)",
                                        data=f,
                                        file_name=file_name,
                                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                        on_click=st.session_state.clear
                                    )

                            except Exception as e:
                                st.error(f"ورڈ فائل بنانے میں مسئلہ پیش آیا:\n{e}")

                except Exception as e:
                    st.error(f"ڈیٹا لوڈ کرنے میں مسئلہ پیش آیا: {e}")