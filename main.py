import os
import pandas as pd
import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, Cm
import docx

# صفحہ کی سیٹنگ
st.set_page_config(
    page_title="Government High School Khan Pur Maral - Paper Generator",
    page_icon="📝",
    layout="centered"
)

# اسٹائلنگ اور ہیڈر
st.markdown(
    """
    <div style="background-color: #1e3a8a; padding: 12px; border-radius: 5px; text-align: center;">
        <h2 style="color: white; margin: 0; font-family: Arial;">آٹومیٹڈ امتحانی پرچہ جنریٹر سسٹم</h2>
        <p style="color: #e2e8f0; margin: 5px 0 0 0; font-size: 14px;">Government High School Khan Pur Maral</p>
    </div>
    """,
    unsafe_allow_html=True
)

excel_path = "Question_Bank.xlsx"

@st.cache_data
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
    st.error(f"مطلوبہ ایکسل فائل '{excel_path}' فولڈر میں نہیں ملی! براہ کرم یقینی بنائیں کہ فائل گিট ہب پر موجود ہے۔")
else:
    st.write("### پرچہ کی ترتیبات")
    
    # کلاس سلیکشن
    classes = list(data_structure.keys())
    selected_cls = st.selectbox("1. کلاس منتخب کریں:", [""] + classes)
    
    selected_subj = ""
    if selected_cls:
        subj_list = list(data_structure[selected_cls].keys())
        selected_subj = st.selectbox("2. مضمون منتخب کریں:", [""] + subj_list)
        
    selected_mode = ""
    if selected_subj:
        selected_mode = st.selectbox(
            "3. پیپر موڈ منتخب کریں:",
            ["", "سنگل یونٹ (Single Unit)", "ملٹی یونٹ (Multi Unit)", "فل بک (Full Book)"]
        )
        
    selected_unit = ""
    selected_topics = []
    if selected_mode and ("سنگل یونٹ" in selected_mode or "Single Unit" in selected_mode):
        units_data = data_structure[selected_cls][selected_subj]["units_data"]
        unit_str_list = [str(u) for u in units_data.keys()]
        selected_unit = st.selectbox("4. یونٹ منتخب کریں:", [""] + unit_str_list)
        
        if selected_unit:
            topics = units_data.get(str(selected_unit).strip(), [])
            if topics:
                st.write("ٹاپکس منتخب کریں:")
                select_all = st.checkbox("سلیکٹ آل (Select All)")
                selected_topics = []
                for topic in topics:
                    is_checked = st.checkbox(str(topic), value=select_all, key=f"top_{topic}")
                    if is_checked:
                        selected_topics.append(topic)

    # مارکس سلیکشن
    marks_option = st.selectbox("پرچے کے کل نمبر منتخب کریں:", ["30 نمبر", "25 نمبر", "40 نمبر"], index=0)

    if selected_cls and selected_subj and selected_mode:
        st.markdown("---")
        st.write("### سوالات کے سیکشنز")
        
        sheet_info = data_structure[selected_cls][selected_subj]
        actual_sheet = sheet_info["sheet"]
        
        try:
            df_full = pd.read_excel(excel_path, sheet_name=actual_sheet)
            
            # سیشن اسٹیٹ میں سوالات محفوظ کرنے کے لیے
            if "saved_sections" not in st.session_state:
                st.session_state.saved_sections = {}

            section_types = [
                (1, "ایم سی کیوز (MCQs - کثیر الانتخابی سوالات)"),
                (2, "مختصر سوالات (Short Questions - SQ)"),
                (5, "طویل سوالات (Long Questions - LQ)")
            ]

            for q_id, title_text in section_types:
                with st.expander(title_text):
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
                        display_str = f"{idx}. {q_text} {expr}"
                        
                        if st.checkbox(display_str, key=f"q_{q_id}_{idx}"):
                            selected_records_for_q.append(r_data)
                    
                    if st.button(f"محفوظ کریں ({title_text})", key=f"save_btn_{q_id}"):
                        st.session_state.saved_sections[q_id] = selected_records_for_q
                        st.success(f"{len(selected_records_for_q)} سوالات محفوظ ہو گئے!")

            # ورڈ فائل جنریشن کا بٹن
            if st.button("Generate Paper (ورڈ میں مکمل پرچہ بنائیں)"):
                if not st.session_state.saved_sections:
                    st.warning("براہ کرم پہلے کم از کم کسی ایک سیکشن کے سوالات منتخب کر کے محفوظ کریں!")
                else:
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

                    p_head = doc.add_paragraph()
                    p_head.alignment = 1
                    p_head.paragraph_format.bidi = True
                    run_head = p_head.add_run("Government High School Khan Pur Maral\nآٹومیٹڈ امتحانی پرچہ")
                    run_head.font.name = 'Jameel Noori Nastaleeq'
                    run_head.font.size = Pt(14)
                    run_head.bold = True

                    total_paper_marks = 0
                    for q_idx, records in st.session_state.saved_sections.items():
                        cnt = len(records)
                        if q_idx == 1: total_paper_marks += cnt * 1
                        elif q_idx == 2: total_paper_marks += cnt * 2
                        elif q_idx == 5: total_paper_marks += cnt * 4

                    raw_minutes = total_paper_marks * 1.5
                    remainder = raw_minutes % 10
                    total_minutes = int(raw_minutes + (10 - remainder)) if remainder >= 5 else int(raw_minutes - remainder)

                    sel_class_display = selected_cls.replace(" Class", "").strip()
                    subject_urdu_map = {
                        "Chemistry": "کیمسٹری", "Physics": "فزکس", "Biology": "بائیولوجی",
                        "Mathematics": "ریاضی", "English": "انگریزی", "Urdu": "اردو",
                        "Islamiyat": "اسلامیات", "Pak_Studies": "مطالعہ پاکستان", "Computer": "کمپیوٹر سائنس"
                    }
                    sel_subj_display = subject_urdu_map.get(selected_subj, selected_subj)
                    sel_unit_display = selected_unit if selected_unit else "Full_Book"

                    p_sub1 = doc.add_paragraph()
                    p_sub1.alignment = 1
                    p_sub1.paragraph_format.bidi = True
                    sub_text1 = f"کلاس: {sel_class_display}\tیونٹ: {sel_unit_display}\tمضمون: {sel_subj_display}"
                    run_sub1 = p_sub1.add_run(sub_text1)
                    run_sub1.font.name = 'Jameel Noori Nastaleeq'
                    run_sub1.font.size = Pt(10)
                    run_sub1.bold = True

                    p_sub2 = doc.add_paragraph()
                    p_sub2.alignment = 1
                    p_sub2.paragraph_format.bidi = True
                    sub_text2 = f"کل نمبر: {total_paper_marks}\t\tکل وقت: {total_minutes} منٹ"
                    run_sub2 = p_sub2.add_run(sub_text2)
                    run_sub2.font.name = 'Jameel Noori Nastaleeq'
                    run_sub2.font.size = Pt(10)
                    run_sub2.bold = True

                    for q_idx in sorted(st.session_state.saved_sections.keys()):
                        records = st.session_state.saved_sections[q_idx]
                        if records:
                            p_heading = doc.add_paragraph()
                            p_heading.paragraph_format.bidi = True
                            run_h = p_heading.add_run()
                            run_h.font.name = 'Jameel Noori Nastaleeq'
                            run_h.bold = True
                            q_count = len(records)
                            
                            if q_idx == 1:
                                run_h.text = f"سوال نمبر 1\tدرست جواب پر دائرہ لگائیں۔\t({q_count} x 1 = {q_count}):"
                            elif q_idx == 2:
                                run_h.text = f"سوال نمبر 2\tدرج ذیل مختصر سوالات کے جوابات دیں۔\t({q_count} x 2 = {q_count * 2}):"
                            elif q_idx == 5:
                                run_h.text = f"سوال نمبر 3\tدرج ذیل طویل سوالات کے جوابات دیں۔\t({q_count} x 4 = {q_count * 4}):"

                            for sub_i, r_data in enumerate(records, 1):
                                p_q = doc.add_paragraph()
                                p_q.paragraph_format.bidi = True
                                run_num = p_q.add_run(f"({sub_i})\t")
                                run_num.font.name = 'Jameel Noori Nastaleeq'
                                run_num.bold = True
                                q_text = str(r_data.get("Q_Text", "")) if pd.notna(r_data.get("Q_Text")) else ""
                                run_qt = p_q.add_run(f"  {q_text}")
                                run_qt.font.name = 'Jameel Noori Nastaleeq'

                    output_file = "Generated_Paper.docx"
                    doc.save(output_file)

                    with open(output_file, "rb") as f:
                        st.download_button(
                            label="📥 ڈاؤن لوڈ ورڈ فائل (Download Paper)",
                            data=f,
                            file_name=output_file,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                    st.success("پرچہ کامیابی سے تیار ہو چکا ہے! اوپر دیے گئے بٹن سے ڈاؤن لوڈ کریں۔")

        except Exception as e:
            st.error(f"ڈیٹا لوڈ کرنے میں خرابی پیش آئی: {e}")