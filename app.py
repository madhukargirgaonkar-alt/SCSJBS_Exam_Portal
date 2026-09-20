import streamlit as st
import pandas as pd
import docx
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import random
import time
import re
import os
import urllib.parse
import unicodedata
import uuid

st.set_page_config(page_title="श्री छत्रपती शिवाजी ज्युनिअर बेसिक स्कूल सगरोळी ऑनलाईन परीक्षा महाप्रणाली", layout="wide")

# ==============================================================================
# आधुनिक व स्वच्छ Custom CSS Design
# ==============================================================================
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@600;700;800&family=Mukta:wght@600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Mukta', sans-serif;
    }
    
    .portal-header {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%);
        color: white;
        padding: 24px;
        border-radius: 16px;
        text-align: center;
        box-shadow: 0 10px 25px rgba(30, 58, 138, 0.2);
        margin-bottom: 24px;
    }
    .portal-header h2 {
        color: white !important;
        margin: 0;
        font-family: 'Poppins', sans-serif;
        font-weight: 800;
        letter-spacing: 0.5px;
    }
    .portal-header p {
        color: #E0E7FF;
        margin: 8px 0 0 0;
        font-size: 20px;
        font-weight: 700;
        font-family: 'Poppins', sans-serif;
    }

    div.stTabs {
        display: flex;
        justify-content: center;
        width: 100%;
    }
    div.stTabs [data-baseweb="tab-list"] {
        display: flex;
        gap: 15px;
        justify-content: center !important;
        background-color: #F8FAFC;
        padding: 12px;
        border-radius: 16px;
        box-shadow: 0 6px 15px rgba(0, 0, 0, 0.06);
        border: 1px solid #E2E8F0;
        margin: 0 auto 25px auto;
    }
    div.stTabs [data-baseweb="tab"] {
        height: 52px;
        background-color: #FFFFFF;
        border-radius: 12px;
        padding: 0 28px;
        font-weight: 700;
        font-family: 'Poppins', sans-serif;
        color: #1E3A8A;
        border: 1px solid #CBD5E1;
        box-shadow: 0 3px 8px rgba(0,0,0,0.03);
        transition: all 0.3s ease;
    }
    div.stTabs [data-baseweb="tab"] svg {
        display: none !important;
    }
    div.stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%) !important;
        color: white !important;
        border: none !important;
        box-shadow: 0 6px 18px rgba(37, 99, 235, 0.35) !important;
    }

    .stButton>button {
        background: linear-gradient(135deg, #2563EB 0%, #1D4ED8 100%);
        color: white;
        font-family: 'Poppins', sans-serif;
        font-size: 16px !important;
        font-weight: 700 !important;
        letter-spacing: 0.5px;
        border-radius: 14px;
        padding: 16px 20px;
        border: none;
        box-shadow: 0 6px 16px rgba(37, 99, 235, 0.25);
        transition: all 0.3s ease-in-out;
        width: 100%;
    }
    .stButton>button:hover {
        transform: translateY(-3px);
        box-shadow: 0 10px 22px rgba(37, 99, 235, 0.4);
        background: linear-gradient(135deg, #1D4ED8 0%, #1E40AF 100%);
    }
    </style>
""", unsafe_allow_html=True)

# ==============================================================================
# १. GOOGLE SHEETS डेटाबेस कनेक्शन व सेफ डेटा रीडर
# ==============================================================================
@st.cache_resource
def get_db():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    
    if os.path.exists("credentials.json"):
        creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
    else:
        try:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        except Exception as e:
            st.error("क्रेडेंशियल फाईल सापडली नाही! कृपया credentials.json फाईल तपासा.")
            raise e
    
    client = gspread.authorize(creds)
    ss = client.open("School_Exam_Results")
    
    def get_or_create(title, headers):
        try:
            return ss.worksheet(title)
        except:
            ws = ss.add_worksheet(title=title, rows="3000", cols="18")
            ws.append_row(headers)
            return ws

    try:
        ws_res = ss.worksheet("Results")
        current_headers = ws_res.row_values(1)
        expected_headers = ["Class", "Division", "Roll_No", "Student_Name", "Exam_Type", "Subject", "Score", "Total", "Date_Time"]
        if not current_headers or current_headers[:len(expected_headers)] != expected_headers:
            ws_res.insert_row(expected_headers, 1)
    except:
        ws_res = ss.add_worksheet(title="Results", rows="3000", cols="15")
        ws_res.append_row(["Class", "Division", "Roll_No", "Student_Name", "Exam_Type", "Subject", "Score", "Total", "Date_Time"])

    # Teachers शीटमध्ये Session_Token साठी अतिरिक्त कॉलम सांभाळणे
    ws_teach = get_or_create("Teachers", ["Teacher_Name", "Mobile", "Assigned_Class", "Assigned_Division", "Username", "Password", "Date_Created", "Session_Token"])
    ws_stud = get_or_create("Students_Credentials", ["Class", "Division", "Roll_No", "Student_Name", "Mobile", "Username", "Password"])
    ws_q = get_or_create("Questions_Bank", ["Class", "Division", "Exam_Type", "Subject", "Marks", "Duration_Minutes", "Question", "Option_A", "Option_B", "Option_C", "Option_D", "Correct_Option", "Added_By"])
    ws_set = get_or_create("Admin_Settings", ["Setting_Name", "Status"])

    return ws_res, ws_teach, ws_stud, ws_q, ws_set

def safe_get_dataframe(worksheet):
    try:
        data = worksheet.get_all_values()
        if not data or len(data) <= 1:
            return pd.DataFrame()
        
        raw_headers = data[0]
        counts = {}
        clean_headers = []
        for i, h in enumerate(raw_headers):
            h_str = str(h).strip()
            if not h_str:
                h_str = f"Col_{i+1}"
            if h_str in counts:
                counts[h_str] += 1
                clean_headers.append(f"{h_str}_{counts[h_str]}")
            else:
                counts[h_str] = 0
                clean_headers.append(h_str)
        
        rows = data[1:]
        padded_rows = []
        header_len = len(clean_headers)
        for r in rows:
            if len(r) < header_len:
                r = r + [''] * (header_len - len(r))
            else:
                r = r[:header_len]
            padded_rows.append(r)
            
        return pd.DataFrame(padded_rows, columns=clean_headers)
    except Exception as ex:
        st.error(f"डेटा वाचताना त्रुटी: {ex}")
        return pd.DataFrame()

# ==============================================================================
# २. स्मार्ट वर्ग व प्रश्ननिहाय गुण तपासणारा Word पार्सर
# ==============================================================================
def clean_marathi(txt):
    if not txt: 
        return ""
    txt = unicodedata.normalize("NFC", str(txt))
    return txt.replace('\u200b', '').replace('\ufeff', '').strip()

def get_class_num(cls_str):
    s = str(cls_str).strip()
    m = re.search(r'\d+', s)
    if m:
        return m.group()
    word_map = {
        "पहिली": "1", "दुसरी": "2", "तिसरी": "3", "तीसरी": "3", 
        "चौथी": "4", "पाचवी": "5", "सहावी": "6", "सातवी": "7"
    }
    for w, n in word_map.items():
        if w in s:
            return n
    return s

def get_clean_div(div_str):
    d = str(div_str).replace("तुकडी", "").strip()
    return d[0] if d else ""

def parse_questions_docx(doc_file, assigned_class, assigned_div, exam_type, final_subject, duration_min, teacher_name):
    doc = docx.Document(doc_file)
    lines = [clean_marathi(p.text) for p in doc.paragraphs if clean_marathi(p.text)]
    q_list, cur_q = [], {}
    cur_sub = final_subject

    for line in lines:
        sub_match = re.search(r"(?:विषय|Subject)\s*[:\-–]\s*(.+)", line, re.IGNORECASE)
        if sub_match:
            detected_sub = sub_match.group(1).strip()
            if detected_sub:
                cur_sub = detected_sub
            continue

        if line.startswith("प्र.") or line.startswith("Q.") or re.match(r"^\d+\.", line):
            if cur_q: 
                q_list.append(cur_q)
            
            q_marks = 1 
            m_match = re.search(r"\[\s*(?:गुण|Marks)\s*[:\-–]?\s*(\d+)\]", line, re.IGNORECASE)
            if m_match:
                try:
                    q_marks = int(m_match.group(1))
                except:
                    q_marks = 1
                line_clean = re.sub(r"\[\s*(?:गुण|Marks)\s*[:\-–]?\s*\d+\]", "", line).strip()
            else:
                line_clean = line

            q_txt = re.sub(r"^(प्र\.\s*\d+\.?|Q\.\s*\d+\.?|\d+\.)", "", line_clean).strip()
            
            cur_q = {
                "Class": assigned_class,
                "Division": assigned_div,
                "Exam_Type": exam_type,
                "Subject": cur_sub,
                "Marks": q_marks,
                "Duration_Minutes": duration_min,
                "Question": q_txt,
                "Option_A": "", "Option_B": "", "Option_C": "", "Option_D": "",
                "Correct_Option": "", "Added_By": teacher_name
            }
        elif line.startswith("A.") or line.startswith("A)"): cur_q["Option_A"] = line[2:].strip()
        elif line.startswith("B.") or line.startswith("B)"): cur_q["Option_B"] = line[2:].strip()
        elif line.startswith("C.") or line.startswith("C)"): cur_q["Option_C"] = line[2:].strip()
        elif line.startswith("D.") or line.startswith("D)"): cur_q["Option_D"] = line[2:].strip()
        elif "उत्तर:" in line or "Ans:" in line:
            ans = line.split(":")[-1].strip().upper()
            mp = {"A": "Option_A", "B": "Option_B", "C": "Option_C", "D": "Option_D"}
            if ans in mp: 
                cur_q["Correct_Option"] = cur_q[mp[ans]]

    if cur_q: 
        q_list.append(cur_q)
    return pd.DataFrame(q_list)

# ==============================================================================
# ३. क्रेडेंशियल व WhatsApp जनरेटर
# ==============================================================================
def gen_teacher_creds(name, mob):
    clean = re.sub(r'[^a-zA-Z]', '', name).lower() or "tchr"
    last4 = str(mob).strip().replace(".0", "")[-4:]
    return f"t_{clean[:5]}_{last4}", f"Tchr@{last4}"

def gen_student_creds_custom(name, cls_name, div_name, roll):
    clean_n = re.sub(r'[^a-zA-Z]', '', str(name).split()[0]).lower() if str(name).strip() else "st"
    if len(clean_n) < 2:
        clean_n = (clean_n + "xx")[:2]
    
    c_code = get_class_num(cls_name)
    div_letter = get_clean_div(div_name)
    div_map = {"अ": "a", "ब": "b", "क": "c", "ड": "d"}
    div_l = div_map.get(div_letter, div_letter.lower() if div_letter else "x")
    
    try:
        r_int = int(roll)
    except:
        r_int = 1
        
    username = f"{clean_n}_{c_code}{div_l}_{r_int:02d}"
    password = f"{clean_n[:2]}{r_int:02d}"
    return username, password

def make_teacher_wa_url(mobile, name, cls, div, username, password):
    m = str(mobile).strip().replace("+", "").replace(".0", "")
    if len(m) == 10:
        m = f"91{m}"
    
    msg = (
        f"*श्री छत्रपती शिवाजी ज्युनिअर बेसिक स्कूल सगरोळी ऑनलाईन परीक्षा महाप्रणाली — शिक्षक प्रवेश*\n"
        f"--------------------------------\n"
        f"नमस्कार *{name}* सर/मॅडम,\n"
        f"आपले शिक्षक खाते तयार झाले आहे.\n\n"
        f"नियुक्त वर्ग: {cls} ({div})\n"
        f"युझरनेम: {username}\n"
        f"पासवर्ड: {password}\n\n"
        f"पोर्टल लिंक: http://localhost:8501\n"
        f"--------------------------------\n"
        f"— मुख्याध्यापक / ॲडमिन"
    )
    encoded_msg = urllib.parse.quote(msg.encode('utf-8'))
    return f"https://web.whatsapp.com/send?phone={m}&text={encoded_msg}"

def make_student_wa_url(mob, name, cls, div, roll, exam_type, subject, score, tot):
    m = str(mob).strip().replace("+", "").replace(".0", "")
    if len(m) == 10:
        m = f"91{m}"
    try:
        pct = round((float(score) / float(tot)) * 100, 1) if float(tot) > 0 else 0
    except:
        pct = 0
    
    msg = (
        f"*{exam_type} — निकाल पत्रक*\n"
        f"--------------------------------\n"
        f"विद्यार्थ्याचे नाव: {name}\n"
        f"वर्ग: {cls} ({div}) | हजेरी क्र.: {roll}\n"
        f"परीक्षा: {exam_type}\n"
        f"विषय: {subject}\n"
        f"मिळालेले गुण: {score} पैकी {tot}\n"
        f"एकूण गुण: {tot} | टक्केवारी: {pct}%\n"
        f"--------------------------------\n"
        f"अभिनंदन! नियमित अभ्यास चालू ठेवावा.\n"
        f"— वर्गशिक्षक"
    )
    encoded_msg = urllib.parse.quote(msg.encode('utf-8'))
    return f"https://web.whatsapp.com/send?phone={m}&text={encoded_msg}"

CLASSES = ["इयत्ता पहिली", "इयत्ता दुसरी", "इयत्ता तिसरी", "इयत्ता चौथी", "इयत्ता पाचवी", "इयत्ता सहावी", "इयत्ता सातवी"]
DIVISIONS = ["अ", "ब", "क", "ड"]
BASE_SUBJECTS = ["मराठी", "गणित", "इंग्रजी", "परिसर अभ्यास", "हिंदी", "सामान्य विज्ञान", "इतिहास", "भूगोल", "संगणक"]
EXAM_TYPES = [
    "घटक चाचणी १",
    "घटक चाचणी २",
    "प्रथम सत्र परीक्षा",
    "घटक चाचणी ३",
    "घटक चाचणी ४",
    "द्वितीय सत्र परीक्षा"
]

# ==============================================================================
# ४. मुख्य स्क्रीनवर मध्यभागी स्टायलिश कार्ड पॅनल
# ==============================================================================
st.markdown("""
    <div class="portal-header">
        <h2>🏫 श्री छत्रपती शिवाजी ज्युनिअर बेसिक स्कूल, सगरोळी</h2>
        <p>ऑनलाईन परीक्षा महाप्रणाली</p>
    </div>
""", unsafe_allow_html=True)

if 'portal_mode' not in st.session_state:
    st.session_state['portal_mode'] = "विद्यार्थी परीक्षा प्रणाली"

st.markdown("<h4 style='text-align: center; color: #1E3A8A; margin-bottom: 15px; font-family: Poppins, sans-serif;'>📌 कृपया आपले पोर्टल निवडा</h4>", unsafe_allow_html=True)

col_c1, col_c2, col_c3 = st.columns([1, 1, 1])

with col_c1:
    if st.button("🎓 विद्यार्थी परीक्षा प्रणाली", use_container_width=True):
        st.session_state['portal_mode'] = "विद्यार्थी परीक्षा प्रणाली"
        st.session_state['t_auth'] = False
        st.session_state['admin_auth'] = False
        st.rerun()

with col_c2:
    if st.button("👨‍🏫 वर्गशिक्षक लॉगिन पोर्टल", use_container_width=True):
        st.session_state['portal_mode'] = "वर्गशिक्षक लॉगिन पोर्टल"
        st.session_state['admin_auth'] = False
        st.rerun()

with col_c3:
    if st.button("⚙️ मुख्य प्रशासक (Admin)", use_container_width=True):
        st.session_state['portal_mode'] = "मुख्य प्रशासक"
        st.session_state['t_auth'] = False
        st.rerun()

st.write("---")

mode = st.session_state['portal_mode']

# ==============================================================================
# 1. विद्यार्थी परीक्षा प्रणाली
# ==============================================================================
if mode == "विद्यार्थी परीक्षा प्रणाली":
    if st.session_state.get('s_done', False):
        st.balloons()
        st.success("✅ तुमची परीक्षा यशस्वीरीत्या सबमिट झाली आहे!")
        st.info("धन्यवाद! तुमची उत्तरे सुरक्षित जमा झाली आहेत. निकाल लवकरच वर्गशिक्षकांकडून जाहीर केला जाईल.")
        if st.button("पुढील विद्यार्थ्यासाठी लॉगिन", key="std_restart"):
            st.session_state['s_done'] = False
            st.session_state['s_auth'] = False
            st.session_state['exam_on'] = False
            st.rerun()
        st.stop()

    if not st.session_state.get('s_auth', False):
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            with st.container(border=True):
                st.markdown("<h3 style='text-align: center; color: #1E3A8A; font-family: Poppins, sans-serif; margin-bottom: 20px;'>🎓 विद्यार्थी लॉगिन</h3>", unsafe_allow_html=True)
                u_in = st.text_input("🆔 User ID (उदा. aman_3a_01):", key="std_u_input").strip()
                p_in = st.text_input("🔒 पासवर्ड (Password):", type="password", key="std_p_input").strip()
                st.write("")
                if st.button("विद्यार्थी लॉगिन करा", key="std_login_btn"):
                    try:
                        _, _, ws_stud, _, _ = get_db()
                        df_stud = safe_get_dataframe(ws_stud)
                        st_match = None
                        if not df_stud.empty and "Username" in df_stud.columns and "Password" in df_stud.columns:
                            m_row = df_stud[(df_stud['Username'].astype(str).str.strip().str.lower() == u_in.lower()) & 
                                            (df_stud['Password'].astype(str).str.strip().str.lower() == p_in.lower())]
                            if not m_row.empty: st_match = m_row.iloc[0].to_dict()

                        if st_match:
                            st.session_state['s_auth'] = True
                            st.session_state['s_info'] = st_match
                            st.rerun()
                        else:
                            st.error("चुकीचा आयडी किंवा पासवर्ड!")
                    except Exception as ex:
                        st.error(f"लॉगिन त्रुटी: {ex}")
    else:
        st_data = st.session_state['s_info']
        s_cls = str(st_data.get('Class', '')).strip()
        s_div = get_clean_div(st_data.get('Division', ''))
        s_cls_num = get_class_num(s_cls)
        s_name_display = st_data.get('Student_Name') or st_data.get('Name') or st_data.get('नाव') or 'विद्यार्थी'
        
        c_head1, c_head2 = st.columns([4, 1])
        with c_head1:
            st.success(f"विद्यार्थी: **{s_name_display}** | वर्ग: **{s_cls} ({s_div})** | हजेरी क्र.: **{st_data.get('Roll_No')}**")
        with c_head2:
            if not st.session_state.get('exam_on', False):
                if st.button("बाहेर पडा", key="std_logout"):
                    st.session_state['s_auth'] = False
                    st.rerun()

        webcam_enabled = False
        try:
            _, _, _, _, ws_set = get_db()
            set_df = safe_get_dataframe(ws_set)
            if not set_df.empty and "Setting_Name" in set_df.columns:
                match_s = set_df[set_df['Setting_Name'] == "Webcam_Proctoring"]
                if not match_s.empty:
                    webcam_enabled = (str(match_s.iloc[0].get("Status", "OFF")) == "ON")
        except:
            pass

        if webcam_enabled:
            st.warning("🛡️ [प्रॉक्टरींग सुरू]: ॲडमिनच्या निर्देशानुसार परीक्षेदरम्यान तुमचा वेबकॅमेरा चालू असणे बंधनकारक आहे.")
            st.camera_input("📷 तुमची लाईव्ह उपस्थिती नोंदवण्यासाठी कॅमेराकडे पहा:")

        _, _, ws_q, _, _ = get_db()
        df_q = safe_get_dataframe(ws_q)
        
        if df_q.empty or "Question" not in df_q.columns:
            st.warning("⚠️ प्रश्न पत्रिकेत अद्याप कोणतेही प्रश्न उपलब्ध नाहीत.")
            st.stop()

        if "Exam_Type" not in df_q.columns: df_q["Exam_Type"] = "सत्र परीक्षा"
        if "Subject" not in df_q.columns: df_q["Subject"] = "सामान्य ज्ञान"
        if "Marks" not in df_q.columns: df_q["Marks"] = 1
        if "Duration_Minutes" not in df_q.columns: df_q["Duration_Minutes"] = "30"
        if "Division" not in df_q.columns: df_q["Division"] = ""
        if "Class" not in df_q.columns: df_q["Class"] = ""

        df_q['c_num'] = df_q['Class'].apply(get_class_num)
        df_q['d_clean'] = df_q['Division'].apply(get_clean_div)
        
        class_q_df = df_q[(df_q['c_num'] == s_cls_num) & ((df_q['d_clean'] == s_div) | (df_q['d_clean'] == ""))]
        if class_q_df.empty:
            st.warning(f"तुमच्या वर्गासाठी (**{s_cls} - {s_div}**) सध्या कोणतीही परीक्षा उपलब्ध नाही.")
            st.stop()

        class_q_df['Exam_Sub_Tag'] = class_q_df['Exam_Type'] + " — विषय: " + class_q_df['Subject']
        avail_exam_tags = class_q_df['Exam_Sub_Tag'].unique().tolist()

        sel_stud_exam_tag = st.selectbox("📌 परीक्षा व विषय निवडा:", avail_exam_tags, key="std_exam_select")
        
        cls_qs = class_q_df[class_q_df['Exam_Sub_Tag'] == sel_stud_exam_tag]
        cur_exam_name = cls_qs['Exam_Type'].iloc[0]
        cur_subject_name = cls_qs['Subject'].iloc[0]
        
        try:
            cls_qs['Marks_int'] = pd.to_numeric(cls_qs['Marks'], errors='coerce').fillna(1)
            custom_total_marks = int(cls_qs['Marks_int'].sum())
        except:
            custom_total_marks = len(cls_qs)

        try:
            custom_duration_min = int(float(cls_qs['Duration_Minutes'].iloc[0]))
        except:
            custom_duration_min = 30 if "घटक चाचणी" in cur_exam_name else 120

        exam_limit_sec = custom_duration_min * 60
        st.info(f"📋 परीक्षा: **{cur_exam_name}** | 📖 विषय: **{cur_subject_name}** | 🎯 एकूण गुण: **{custom_total_marks}** | ⏳ वेळ: **{custom_duration_min} मिनिटे**")

        if not st.session_state.get('exam_on', False):
            if st.button(f"🚀 {cur_subject_name} परीक्षा सुरू करा", key="start_exam_btn"):
                st.session_state['exam_on'] = True
                st.session_state['exam_type_running'] = cur_exam_name
                st.session_state['subject_running'] = cur_subject_name
                st.session_state['total_marks_running'] = custom_total_marks
                st.session_state['exam_time_limit'] = exam_limit_sec
                st.session_state['t_start'] = time.time()
                q_list = cls_qs.to_dict('records')
                random.shuffle(q_list)
                for item in q_list:
                    opts = [item.get('Option_A', ''), item.get('Option_B', ''), item.get('Option_C', ''), item.get('Option_D', '')]
                    random.shuffle(opts)
                    item['shuffled_opts'] = opts
                st.session_state['s_qs'] = q_list
                st.rerun()

        if st.session_state.get('exam_on', False):
            cur_qs = st.session_state['s_qs']
            cur_exam_name = st.session_state.get('exam_type_running', 'परीक्षा')
            cur_subject_name = st.session_state.get('subject_running', 'सामान्य ज्ञान')
            tot_marks = st.session_state.get('total_marks_running', len(cur_qs))
            time_limit = st.session_state.get('exam_time_limit', 1800)
            rem_sec = int(time_limit - (time.time() - st.session_state['t_start']))

            def save_silent(ans_map):
                earned_score = 0
                for idx, q in enumerate(cur_qs):
                    if ans_map.get(idx) == q.get('Correct_Option'):
                        try: earned_score += int(q.get('Marks', 1))
                        except: earned_score += 1

                actual_name = st_data.get('Student_Name') or st_data.get('Name') or st_data.get('नाव') or f"विद्यार्थी {st_data.get('Roll_No', '')}"
                try:
                    ws_res, _, _, _, _ = get_db()
                    ws_res.append_row([s_cls, s_div, int(st_data.get('Roll_No', 0)), actual_name, cur_exam_name, cur_subject_name, earned_score, tot_marks, time.strftime("%Y-%m-%d %H:%M:%S")])
                except Exception: pass
                st.session_state['exam_on'] = False
                st.session_state['s_done'] = True
                st.rerun()

            if rem_sec <= 0:
                st.error("वेळ संपली आहे! परीक्षा आपोआप सबमिट होत आहे...")
                time.sleep(1)
                save_silent({})
            else:
                hours, rem = divmod(rem_sec, 3600)
                m, s = divmod(rem, 60)
                time_display = f"{hours:02d} तास {m:02d} मि. {s:02d} से." if hours > 0 else f"{m:02d}:{s:02d}"

                st.warning(f"📝 चालू परीक्षा: **{cur_exam_name}** ({cur_subject_name}) | 🎯 एकूण गुण: **{tot_marks}** | ⏳ शिल्लक वेळ: **{time_display}**")

                ans_dict = {}
                with st.form("st_exam"):
                    for idx, q in enumerate(cur_qs):
                        q_m = q.get('Marks', 1)
                        st.markdown(f"**प्रश्न {idx+1}: {q.get('Question', '')}** *({q_m} गुण)*")
                        ans_dict[idx] = st.radio(
                            f"उत्तर निवडा (प्र. {idx+1}):", 
                            q.get('shuffled_opts', []), 
                            index=None, 
                            key=f"q_{idx}"
                        )
                        st.write("---")
                    if st.form_submit_button("🏁 परीक्षा सबमिट करा"):
                        save_silent(ans_dict)

# ==============================================================================
# 2. वर्गशिक्षक लॉगिन पोर्टल (सिंगल ॲक्टिव्ह सेशन व ऑटो लॉगआऊट वैशिष्ट्यासह)
# ==============================================================================
elif mode == "वर्गशिक्षक लॉगिन पोर्टल":
    # सुरुवातीला पार्श्वभूमीवर सत्र (Session Token) तपासा
    if st.session_state.get('t_auth', False):
        try:
            _, ws_teach, _, _, _ = get_db()
            df_t = safe_get_dataframe(ws_teach)
            current_user = st.session_state['t_info'].get('Username', '')
            current_token = st.session_state.get('t_session_token', '')
            
            if not df_t.empty and "Username" in df_t.columns and "Session_Token" in df_t.columns:
                match_row = df_t[df_t['Username'].astype(str).str.strip() == str(current_user).strip()]
                if not match_row.empty:
                    db_token = str(match_row.iloc[0].get('Session_Token', '')).strip()
                    # जर डेटाबेसवरील टोकन आपल्या सत्राच्या टोकनपेक्षा वेगळे असेल, तर याचा अर्थ दुसऱ्या ठिकाणाहून लॉगिन झाले आहे!
                    if db_token and db_token != current_token:
                        st.warning("⚠️ तुमचे खाते अन्य ठिकाणाहून (Other Device/Tab) लॉगिन केले गेले आहे. त्यामुळे हे सत्र बंद (Auto Logged Out) होत आहे.")
                        st.session_state['t_auth'] = False
                        st.session_state.pop('t_info', None)
                        st.session_state.pop('t_session_token', None)
                        time.sleep(2)
                        st.rerun()
        except Exception:
            pass

    if not st.session_state.get('t_auth', False):
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            with st.container(border=True):
                st.markdown("<h3 style='text-align: center; color: #1E3A8A; font-family: Poppins, sans-serif; margin-bottom: 20px;'>🔑 वर्ग शिक्षक लॉगिन</h3>", unsafe_allow_html=True)
                u = st.text_input("👤 शिक्षकांचे युझरनेम (Username):", key="t_u_input")
                p = st.text_input("🔒 पासवर्ड (Password):", type="password", key="t_p_input")
                st.write("")
                if st.button("वर्ग शिक्षक लॉगिन करा", key="teacher_login_btn"):
                    try:
                        _, ws_teach, _, _, _ = get_db()
                        df_t = safe_get_dataframe(ws_teach)
                        matched = None
                        matched_row_idx = None
                        if not df_t.empty and "Username" in df_t.columns and "Password" in df_t.columns:
                            for idx, row in df_t.iterrows():
                                if str(row.get('Username', '')).strip() == u.strip() and str(row.get('Password', '')).strip() == p.strip():
                                    matched = row.to_dict()
                                    matched_row_idx = idx + 2 # Google Sheets 1-indexed plus header
                                    break
                        
                        if matched:
                            # नवीन युनिक सेशन टोकन तयार करणे
                            new_token = str(uuid.uuid4())
                            matched['Session_Token'] = new_token
                            
                            # जर शीट्समध्ये Session_Token कॉलम नसेल किंवा अपडेट करायचे असेल
                            cell = ws_teach.find(u.strip())
                            if cell:
                                # युझरनेम सापडले, आता त्याच ओळीत Session_Token (समजा ८ वा कॉलम) अपडेट करणे
                                headers = ws_teach.row_values(1)
                                if "Session_Token" in headers:
                                    tok_col_idx = headers.index("Session_Token") + 1
                                    ws_teach.update_cell(cell.row, tok_col_idx, new_token)
                                else:
                                    ws_teach.append_row([matched.get('Teacher_Name'), matched.get('Mobile'), matched.get('Assigned_Class'), matched.get('Assigned_Division'), u, p, time.strftime("%Y-%m-%d"), new_token])

                            st.session_state['t_auth'] = True
                            st.session_state['t_info'] = matched
                            st.session_state['t_session_token'] = new_token
                            st.success("✅ यशस्वी लॉगिन! प्रवेश देत आहे...")
                            time.sleep(1)
                            st.rerun()
                        else:
                            st.error("चुकीचे युझरनेम किंवा पासवर्ड!")
                    except Exception as ex:
                        st.error(f"डेटाबेस त्रुटी: {ex}")
    else:
        t_info = st.session_state['t_info']
        my_cls = t_info.get('Assigned_Class', '')
        my_div = get_clean_div(t_info.get('Assigned_Division', ''))
        my_cls_num = get_class_num(my_cls)
        
        head_col1, head_col2 = st.columns([4, 1])
        with head_col1:
            st.success(f"स्वागत आहे: **{t_info.get('Teacher_Name', 'शिक्षक')}** सर/मॅडम | नियुक्त वर्ग: **{my_cls} - {my_div}**")
        with head_col2:
            if st.button("🔴 लॉगआऊट", key="teacher_logout"):
                st.session_state['t_auth'] = False
                st.session_state.pop('t_info', None)
                st.session_state.pop('t_session_token', None)
                st.rerun()

        t1, t2, t3 = st.tabs(["📝 प्रश्नसंच अपलोड व वेळ सेटिंग", "👥 विद्यार्थी व्यवस्थापन व पासवर्ड", "📊 वर्गाचा निकाल व WhatsApp वाटप"])

        with t1:
            st.markdown(f"#### ⚙️ **{my_cls} - {my_div}** साठी परीक्षेचे नियोजन व प्रश्न अपलोड")
            c_ex1, c_ex2 = st.columns(2)
            with c_ex1:
                sel_exam = st.selectbox("📌 १. परीक्षेचे नाव निवडा:", EXAM_TYPES, key="t_exam_sel")
            with c_ex2:
                sub_mode = st.radio("📌 २. विषय कसा निवडायचा?", ["यादीतून निवडा", "स्वतः नवीन विषय टाईप करा"], horizontal=True, key="t_sub_mode")

            if sub_mode == "यादीतून निवडा":
                final_subject = st.selectbox("विषय निवडा:", BASE_SUBJECTS, key="t_base_sub")
            else:
                final_subject = st.text_input("✍️ नवीन विषयाचे नाव येथे टाईप करा:", value="चित्रकला", key="t_custom_sub").strip()
                if not final_subject: final_subject = "सामान्य ज्ञान"

            def_time = 30 if "घटक चाचणी" in sel_exam else 120
            st.write("---")
            
            up_doc = st.file_uploader("Word फाईल (.docx) निवडा", type=["docx"], key="doc_up_teacher")
            
            calc_marks = 0
            if up_doc:
                temp_df = parse_questions_docx(up_doc, my_cls, my_div, sel_exam, final_subject, 30, t_info.get('Teacher_Name', 'शिक्षक'))
                calc_marks = int(temp_df['Marks'].astype(int).sum()) if not temp_df.empty else 0

            col_dur, col_tot_m = st.columns(2)
            with col_dur:
                set_duration = st.number_input("⏱️ परीक्षेची एकूण वेळ (मिनिटांमध्ये):", min_value=5, max_value=180, value=def_time, step=5, key="t_duration")
            with col_tot_m:
                custom_total_marks_input = st.number_input("🎯 परीक्षेचे एकूण गुण (Total Marks):", min_value=1, max_value=500, value=(calc_marks if calc_marks > 0 else 20), step=1, key="t_total_marks_input")

            if up_doc:
                df_q = parse_questions_docx(up_doc, my_cls, my_div, sel_exam, final_subject, set_duration, t_info.get('Teacher_Name', 'शिक्षक'))
                
                st.write(f"एकूण प्रश्न: **{len(df_q)}** | स्वयंचलित मोजलेले गुण: **{calc_marks}** | ठरवलेले एकूण गुण: **{custom_total_marks_input}** | विषय: **{final_subject}**")
                st.dataframe(df_q[["Exam_Type", "Subject", "Marks", "Duration_Minutes", "Question", "Option_A", "Option_B", "Option_C", "Option_D", "Correct_Option"]])
                
                if st.button(f"💾 {my_cls} ({my_div}) चे '{sel_exam}' प्रश्न सेव्ह करा", key="save_q_btn"):
                    try:
                        _, _, _, ws_q, _ = get_db()
                        cols_order = ["Class", "Division", "Exam_Type", "Subject", "Marks", "Duration_Minutes", "Question", "Option_A", "Option_B", "Option_C", "Option_D", "Correct_Option", "Added_By"]
                        ws_q.append_rows(df_q[cols_order].values.tolist())
                        st.success(f"✅ प्रश्नपत्रिका सुरक्षित सेव्ह झाली! (एकूण गुण: {custom_total_marks_input})")
                    except Exception as ex:
                        st.error(f"त्रुटी: {ex}")

        with t2:
            st.markdown(f"#### 👥 **{my_cls} - {my_div}** नोंदणीकृत विद्यार्थी यादी")
            _, _, ws_stud, _, _ = get_db()
            df_curr_s = safe_get_dataframe(ws_stud)
            existing_for_class = pd.DataFrame()
            if not df_curr_s.empty and "Class" in df_curr_s.columns and "Division" in df_curr_s.columns:
                df_curr_s['c_num'] = df_curr_s['Class'].apply(get_class_num)
                df_curr_s['d_clean'] = df_curr_s['Division'].apply(get_clean_div)
                existing_for_class = df_curr_s[(df_curr_s['c_num'] == my_cls_num) & (df_curr_s['d_clean'] == my_div)]

            if not existing_for_class.empty:
                st.success(f"✅ या वर्गात **{len(existing_for_class)}** विद्यार्थी सेव्ह आहेत.")
                display_cols = [c for c in ["Roll_No", "Student_Name", "Username", "Password", "Mobile"] if c in existing_for_class.columns]
                st.dataframe(existing_for_class[display_cols].sort_values(by="Roll_No"), use_container_width=True)
                csv_c = existing_for_class[display_cols].to_csv(index=False, encoding="utf-8-sig").encode('utf-8-sig')
                st.download_button("📥 विद्यार्थ्यांचे ID-Password Excel डाउनलोड करा", data=csv_c, file_name=f"{my_cls}_{my_div}_Passwords.csv", key="dl_stud_pass")
            else:
                st.info("या वर्गासाठी अद्याप एकही विद्यार्थी सेव्ह नाही.")

            with st.expander("➕ नवीन विद्यार्थ्यांची Excel फाईल जोडा"):
                up_st = st.file_uploader("Excel फाईल निवडा", type=["xlsx"], key="st_up_excel")
                if up_st:
                    df_in = pd.read_excel(up_st)
                    if {"Roll_No", "Name", "Mobile"}.issubset(df_in.columns):
                        s_creds = []
                        for _, r in df_in.iterrows():
                            u_id, p_wd = gen_student_creds_custom(r['Name'], my_cls, my_div, r['Roll_No'])
                            s_creds.append({
                                "Class": my_cls, "Division": my_div, "Roll_No": int(r['Roll_No']),
                                "Student_Name": r['Name'], "Mobile": str(r['Mobile']), "Username": u_id, "Password": p_wd
                            })
                        df_out = pd.DataFrame(s_creds)
                        st.dataframe(df_out[["Roll_No", "Student_Name", "Username", "Password", "Mobile"]])
                        if st.button("ही सर्व खाती डेटाबेसमध्ये جوडा", key="add_stud_db"):
                            try:
                                rows_to_add = df_out[["Class", "Division", "Roll_No", "Student_Name", "Mobile", "Username", "Password"]].values.tolist()
                                ws_stud.append_rows(rows_to_add)
                                st.success("✅ सर्व विद्यार्थी सेव्ह झाले!")
                                time.sleep(1)
                                st.rerun()
                            except Exception as ex:
                                st.error(f"त्रुटी: {ex}")

        with t3:
            st.markdown(f"#### 📊 **{my_cls} - {my_div}** निकाल आणि WhatsApp वाटप")
            filter_exam = st.selectbox("📌 निकाल पाहण्यासाठी परीक्षा निवडा:", EXAM_TYPES, key="res_exam_sel_t3")
            try:
                ws_res, _, ws_stud, _, _ = get_db()
                df_r = safe_get_dataframe(ws_res)
                df_s = safe_get_dataframe(ws_stud)
                
                if not df_r.empty and "Class" in df_r.columns:
                    if "Exam_Type" not in df_r.columns: df_r["Exam_Type"] = "सत्र परीक्षा"
                    if "Subject" not in df_r.columns: df_r["Subject"] = "सामान्य ज्ञान"
                    if "Division" not in df_r.columns: df_r["Division"] = ""
                    
                    df_r['c_num'] = df_r['Class'].apply(get_class_num)
                    df_r['d_clean'] = df_r['Division'].apply(get_clean_div)
                    cls_res = df_r[(df_r['c_num'] == my_cls_num) & (df_r['d_clean'] == my_div) & (df_r['Exam_Type'].astype(str).str.strip() == filter_exam.strip())]
                    
                    if not cls_res.empty:
                        if not df_s.empty and "Class" in df_s.columns:
                            df_s['c_num'] = df_s['Class'].apply(get_class_num)
                            df_s['d_clean'] = df_s['Division'].apply(get_clean_div)

                        for _, r in cls_res.iterrows():
                            s_roll = str(r.get('Roll_No', '')).strip()
                            s_score = r.get('Score', 0)
                            s_total = r.get('Total', 0)
                            s_sub = str(r.get('Subject', 'सामान्य ज्ञान')).strip()
                            s_name = str(r.get('Student_Name') or r.get('Name') or r.get('नाव') or '').strip()
                            
                            mob = ""
                            if not df_s.empty and "Roll_No" in df_s.columns:
                                m_row = df_s[(df_s['c_num'] == my_cls_num) & (df_s['d_clean'] == my_div) & (df_s['Roll_No'].astype(str).str.strip() == s_roll)]
                                if len(m_row) > 0:
                                    first_m = m_row.iloc[0]
                                    reg_name = str(first_m.get('Student_Name') or first_m.get('Name') or first_m.get('नाव') or '').strip()
                                    if reg_name: s_name = reg_name
                                    mob = str(first_m.get('Mobile', '')).strip()
                            
                            if not s_name or s_name == "विद्यार्थी": s_name = f"विद्यार्थी (ह.क्र. {s_roll})"

                            c1, c2, c3, c4 = st.columns([1, 2, 2, 2])
                            c1.write(f"ह.क्र. **{s_roll}**")
                            c2.write(f"**{s_name}**\n\n*(विषय: {s_sub})*")
                            c3.write(f"गुण: **{s_score} पैकी {s_total}**")
                            
                            with c4:
                                if mob:
                                    wa_href = make_student_wa_url(mob, s_name, my_cls, my_div, s_roll, filter_exam, s_sub, s_score, s_total)
                                    st.markdown(f'<a href="{wa_href}" target="_blank" style="background-color:#25D366;color:white;padding:5px 12px;text-decoration:none;border-radius:6px;font-weight:bold;">📲 निकाल पाठवा</a>', unsafe_allow_html=True)
                                else:
                                    st.caption("मोबाईल नं. नाही")
                            st.write("---")
                    else:
                        st.info(f"'{filter_exam}' साठी कोणताही निकाल जमा झालेला नाही.")
                else:
                    st.info("डेटाबेसमध्ये अद्याप कोणताही निकाल नाही.")
            except Exception as ex:
                st.error(f"त्रुटी: {ex}")

# ==============================================================================
# 3. मुख्य प्रशासक
# ==============================================================================
elif mode == "मुख्य प्रशासक":
    if not st.session_state.get('admin_auth', False):
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            with st.container(border=True):
                st.markdown("<h3 style='text-align: center; color: #1E3A8A; font-family: Poppins, sans-serif; margin-bottom: 20px;'>🔐 ॲडमिन लॉगिन</h3>", unsafe_allow_html=True)
                admin_pwd = st.text_input("🔑 मास्टर पासवर्ड प्रविष्ट करा:", type="password", key="admin_pwd_input")
                st.write("")
                if st.button("ॲडमिन लॉगिन करा", key="admin_login_btn"):
                    if admin_pwd == "admin@123":
                        st.session_state['admin_auth'] = True
                        st.rerun()
                    else:
                        st.error("चुकीचा मास्टर पासवर्ड!")
    else:
        ahead1, ahead2 = st.columns([4, 1])
        with ahead1:
            st.success("✅ ॲडमिन सुरक्षित सत्रात कार्यरत आहे.")
        with ahead2:
            if st.button("🔴 लॉगआऊट", key="admin_logout"):
                st.session_state['admin_auth'] = False
                st.rerun()

        a1, a2, a3 = st.tabs(["👨‍🏫 शिक्षक निर्मिती", "📊 मास्टर निकाल", "🛡️ परीक्षा सुरक्षा (Webcam Control)"])

        with a1:
            st.subheader("👨‍🏫 नवीन शिक्षक खाते निर्मिती")
            t_n = st.text_input("शिक्षकांचे पूर्ण नाव:", key="adm_t_name").strip()
            t_m = st.text_input("मोबाईल नंबर (१० अंकी):", max_chars=10, key="adm_t_mob").strip()
            
            col_tc, col_td = st.columns(2)
            with col_tc:
                t_c = st.selectbox("इयत्ता निवडा:", CLASSES, key="adm_t_cls")
            with col_td:
                t_d = st.selectbox("तुकडी निवडा:", DIVISIONS, key="adm_t_div")
            
            st.write("")
            if st.button("⚡ शिक्षक ऑटो खाते बनवा", key="create_teacher_btn"):
                if t_n and len(t_m) == 10 and t_m.isdigit():
                    try:
                        _, ws_teach, _, _, _ = get_db()
                        df_t = safe_get_dataframe(ws_teach)
                        matched_by_mob, matched_by_class = None, None
                        if not df_t.empty:
                            if "Mobile" in df_t.columns:
                                m_match = df_t[df_t['Mobile'].astype(str).str.replace('.0','').str.strip() == str(t_m)]
                                if not m_match.empty: matched_by_mob = m_match.iloc[0].to_dict()
                            if "Assigned_Class" in df_t.columns and "Assigned_Division" in df_t.columns:
                                df_t['c_num'] = df_t['Assigned_Class'].apply(get_class_num)
                                df_t['d_clean'] = df_t['Assigned_Division'].apply(get_clean_div)
                                c_match = df_t[(df_t['c_num'] == get_class_num(t_c)) & (df_t['d_clean'] == t_d)]
                                if not c_match.empty: matched_by_class = c_match.iloc[0].to_dict()

                        if matched_by_mob:
                            st.error(f"⚠️ या मोबाईलवर ({t_m}) खाते आधीच आहे!")
                        elif matched_by_class:
                            st.warning(f"⚠️ {t_c} ({t_d}) साठी आधीच शिक्षक नियुक्त आहेत.")
                        else:
                            u, p = gen_teacher_creds(t_n, t_m)
                            ws_teach.append_row([t_n, str(t_m), t_c, t_d, u, p, time.strftime("%Y-%m-%d"), ""])
                            st.success(f"✅ नवीन खाते तयार झाले! (युझरनेम: {u} | पासवर्ड: {p})")
                            time.sleep(1)
                            st.rerun()
                    except Exception as ex:
                        st.error(f"त्रुटी: {ex}")
                else:
                    st.warning("कृपया अचूक नाव आणि १० अंकी मोबाईल नंबर प्रविष्ट करा.")

        with a2:
            st.subheader("📈 सर्व परीक्षांचा मास्टर निकाल (Centralized Master Reports)")
            try:
                ws_res, _, _, _, _ = get_db()
                df_all = safe_get_dataframe(ws_res)
                if not df_all.empty:
                    if "Student_Name" in df_all.columns:
                        df_all = df_all[df_all['Student_Name'].astype(str).str.strip() != "Student_Name"]
                        df_all = df_all[df_all['Student_Name'].astype(str).str.strip() != ""]

                    for col in ["Class", "Division", "Roll_No", "Student_Name", "Exam_Type", "Subject", "Score", "Total", "Date_Time"]:
                        if col not in df_all.columns: df_all[col] = ""

                    fc1, fc2, fc3 = st.columns(3)
                    with fc1:
                        opt_cls = ["सर्व वर्ग"] + sorted([str(x) for x in df_all['Class'].unique() if str(x).strip()])
                        sel_f_cls = st.selectbox("वर्ग निवडा:", opt_cls, key="adm_cls_sel_master")
                    with fc2:
                        opt_exam = ["सर्व परीक्षा"] + sorted([str(x) for x in df_all['Exam_Type'].unique() if str(x).strip()])
                        sel_f_exam = st.selectbox("परीक्षा निवडा:", opt_exam, key="adm_ex_sel_master")
                    with fc3:
                        opt_sub = ["सर्व विषय"] + sorted([str(x) for x in df_all['Subject'].unique() if str(x).strip()])
                        sel_f_sub = st.selectbox("विषय निवडा:", opt_sub, key="adm_sub_sel_master")

                    filtered_df = df_all.copy()
                    if sel_f_cls != "सर्व वर्ग": filtered_df = filtered_df[filtered_df['Class'].astype(str) == sel_f_cls]
                    if sel_f_exam != "सर्व परीक्षा": filtered_df = filtered_df[filtered_df['Exam_Type'].astype(str) == sel_f_exam]
                    if sel_f_sub != "सर्व विषय": filtered_df = filtered_df[filtered_df['Subject'].astype(str) == sel_f_sub]

                    def calc_pct(row):
                        try:
                            sc = float(row['Score'])
                            tot = float(row['Total'])
                            if tot > 0: return f"{(sc / tot * 100):.1f}%"
                        except: pass
                        return "-"

                    filtered_df['Percentage (%)'] = filtered_df.apply(calc_pct, axis=1)
                    cols_to_show = ["Class", "Division", "Roll_No", "Student_Name", "Exam_Type", "Subject", "Score", "Total", "Percentage (%)", "Date_Time"]
                    final_show_cols = [c for c in cols_to_show if c in filtered_df.columns]
                    
                    st.dataframe(filtered_df[final_show_cols], use_container_width=True)
                    csv_m = filtered_df[final_show_cols].to_csv(index=False, encoding="utf-8-sig").encode('utf-8-sig')
                    st.download_button("📥 मास्टर निकाल (Excel) डाउनलोड करा", data=csv_m, file_name="Master_Exam_Result.csv", mime="text/csv", key="dl_master_csv")
                else:
                    st.info("कोणताही निकाल उपलब्ध नाही.")
            except Exception as ex:
                st.error(f"त्रुटी: {ex}")

        with a3:
            st.subheader("🛡️ परीक्षा सुरक्षा व वेबकॅमेरा नियंत्रण (Anti-Cheat Proctoring)")
            try:
                _, _, _, _, ws_set = get_db()
                set_df = safe_get_dataframe(ws_set)
                current_cam_status = "OFF"
                if not set_df.empty and "Setting_Name" in set_df.columns:
                    match_s = set_df[set_df['Setting_Name'] == "Webcam_Proctoring"]
                    if not match_s.empty: current_cam_status = str(match_s.iloc[0].get("Status", "OFF"))

                cam_toggle = st.toggle("🎥 परीक्षा देताना विद्यार्थ्यांचा वेबकॅमेरा (Webcam) ऑन ठेवा", value=(current_cam_status == "ON"), key="adm_cam_toggle")
                if st.button("💾 कॅमेरा सेटिंग सेव्ह करा", key="save_cam_btn"):
                    new_status = "ON" if cam_toggle else "OFF"
                    try:
                        cell = ws_set.find("Webcam_Proctoring")
                        if cell: ws_set.update_cell(cell.row, 2, new_status)
                        else: ws_set.append_row(["Webcam_Proctoring", new_status])
                        st.success(f"✅ सेटिंग यशस्वीरीत्या सेव्ह झाली! सद्यस्थिती: कॅमेरा {new_status}")
                    except Exception as e:
                        st.error(f"त्रुटी: {e}")
            except Exception as ex:
                st.error(f"त्रुटी: {ex}")