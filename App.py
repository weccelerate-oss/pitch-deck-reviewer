"""
Pitch Deck Review System - Premium Edition
מערכת סקירת מצגות AI - מהדורת פרימיום

A beautifully designed Streamlit application for analyzing presentations
using Google Gemini AI with premium UI/UX and full Hebrew RTL support.
"""

import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from docx import Document
import google.generativeai as genai
import json
import io


# ============================================================
# API Configuration - הגדר את המפתח שלך כאן
# ============================================================
GEMINI_API_KEY = "AIzaSyBJstgLpy_6W8OkQTD6t8HmfYTLL1sTLXE"  # <-- הכנס את המפתח שלך כאן

# Configure Gemini API
genai.configure(api_key=GEMINI_API_KEY)


# ============================================================
# Page Configuration
# ============================================================
st.set_page_config(
    page_title="מערכת סקירת מצגות AI",
    page_icon="✨",
    layout="wide",
    initial_sidebar_state="expanded"
)


# ============================================================
# Premium CSS Styling
# ============================================================
st.markdown("""
<style>
/* ===== IMPORTS ===== */
@import url('https://fonts.googleapis.com/css2?family=Heebo:wght@300;400;500;600;700;800&display=swap');

/* ===== ROOT VARIABLES ===== */
:root {
    --primary-gradient: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    --secondary-gradient: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
    --success-gradient: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
    --card-bg: rgba(255, 255, 255, 0.95);
    --shadow-light: 0 4px 20px rgba(0, 0, 0, 0.08);
    --shadow-medium: 0 8px 40px rgba(0, 0, 0, 0.12);
    --shadow-heavy: 0 20px 60px rgba(0, 0, 0, 0.15);
    --border-radius: 16px;
    --transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
}

/* ===== GLOBAL STYLES ===== */
.stApp {
    direction: rtl;
    font-family: 'Heebo', sans-serif !important;
    background: linear-gradient(135deg, #f5f7fa 0%, #e4e8ec 100%);
}

/* ===== MAIN CONTAINER ===== */
.main .block-container {
    direction: rtl;
    text-align: right;
    padding: 2rem 3rem;
    max-width: 1400px;
}

/* ===== SIDEBAR STYLING ===== */
[data-testid="stSidebar"] {
    direction: rtl;
    background: linear-gradient(180deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
}

[data-testid="stSidebar"] * {
    color: #ffffff !important;
    text-align: right;
}

[data-testid="stSidebar"] .stMarkdown h2,
[data-testid="stSidebar"] .stMarkdown h3 {
    color: #ffffff !important;
    font-weight: 600;
    border-bottom: 2px solid rgba(255,255,255,0.1);
    padding-bottom: 0.5rem;
    margin-bottom: 1rem;
}

[data-testid="stSidebar"] .stSelectbox > div > div {
    background: rgba(255,255,255,0.1);
    border: 1px solid rgba(255,255,255,0.2);
    border-radius: 10px;
}

[data-testid="stSidebar"] hr {
    border-color: rgba(255,255,255,0.1);
    margin: 1.5rem 0;
}

/* ===== PREMIUM HEADER ===== */
.premium-header {
    background: var(--primary-gradient);
    border-radius: 20px;
    padding: 2.5rem 3rem;
    margin-bottom: 2rem;
    box-shadow: var(--shadow-heavy);
    position: relative;
    overflow: hidden;
}

.premium-header::before {
    content: '';
    position: absolute;
    top: -50%;
    right: -50%;
    width: 100%;
    height: 200%;
    background: radial-gradient(circle, rgba(255,255,255,0.1) 0%, transparent 70%);
    animation: shimmer 3s ease-in-out infinite;
}

@keyframes shimmer {
    0%, 100% { transform: translateX(0) translateY(0); }
    50% { transform: translateX(10%) translateY(-10%); }
}

.premium-header h1 {
    color: white !important;
    font-size: 2.8rem !important;
    font-weight: 800 !important;
    margin: 0 !important;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
    position: relative;
    z-index: 1;
}

.premium-header p {
    color: rgba(255,255,255,0.9) !important;
    font-size: 1.2rem !important;
    margin-top: 0.5rem !important;
    position: relative;
    z-index: 1;
}

/* ===== CARD STYLING ===== */
.premium-card {
    background: var(--card-bg);
    border-radius: var(--border-radius);
    padding: 1.8rem;
    box-shadow: var(--shadow-light);
    border: 1px solid rgba(255,255,255,0.8);
    transition: var(--transition);
    margin-bottom: 1.5rem;
}

.premium-card:hover {
    box-shadow: var(--shadow-medium);
    transform: translateY(-2px);
}

.premium-card-header {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-bottom: 1rem;
    padding-bottom: 1rem;
    border-bottom: 2px solid #f0f2f5;
}

.premium-card-icon {
    width: 48px;
    height: 48px;
    border-radius: 12px;
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 1.5rem;
    background: var(--primary-gradient);
    box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4);
}

.premium-card-title {
    font-size: 1.3rem;
    font-weight: 700;
    color: #1a1a2e;
    margin: 0;
}

/* ===== UPLOAD AREA ===== */
.upload-zone {
    background: linear-gradient(135deg, #f8f9ff 0%, #f0f4ff 100%);
    border: 2px dashed #667eea;
    border-radius: 16px;
    padding: 2rem;
    text-align: center;
    transition: var(--transition);
}

.upload-zone:hover {
    border-color: #764ba2;
    background: linear-gradient(135deg, #f0f4ff 0%, #e8edff 100%);
}

[data-testid="stFileUploader"] {
    direction: rtl;
}

[data-testid="stFileUploader"] section {
    direction: rtl;
    border: 2px dashed #667eea;
    border-radius: 12px;
    padding: 1.5rem;
    background: linear-gradient(135deg, #fafbff 0%, #f5f7ff 100%);
    transition: var(--transition);
}

[data-testid="stFileUploader"] section:hover {
    border-color: #764ba2;
    background: linear-gradient(135deg, #f5f7ff 0%, #eef1ff 100%);
}

/* ===== BUTTONS ===== */
.stButton > button {
    direction: rtl;
    background: var(--primary-gradient) !important;
    color: white !important;
    border: none !important;
    border-radius: 12px !important;
    padding: 0.8rem 2rem !important;
    font-weight: 600 !important;
    font-size: 1.1rem !important;
    transition: var(--transition) !important;
    box-shadow: 0 4px 15px rgba(102, 126, 234, 0.4) !important;
}

.stButton > button:hover {
    transform: translateY(-2px) !important;
    box-shadow: 0 8px 25px rgba(102, 126, 234, 0.5) !important;
}

.stButton > button:active {
    transform: translateY(0) !important;
}

.stButton > button[disabled] {
    background: linear-gradient(135deg, #ccc 0%, #aaa 100%) !important;
    box-shadow: none !important;
}

/* ===== DOWNLOAD BUTTONS ===== */
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #11998e 0%, #38ef7d 100%) !important;
    box-shadow: 0 4px 15px rgba(17, 153, 142, 0.4) !important;
}

[data-testid="stDownloadButton"] > button:hover {
    box-shadow: 0 8px 25px rgba(17, 153, 142, 0.5) !important;
}

/* ===== ALERTS ===== */
.stAlert {
    direction: rtl;
    text-align: right;
    border-radius: 12px !important;
    border: none !important;
}

[data-testid="stAlert"] {
    border-radius: 12px;
    padding: 1rem 1.5rem;
}

/* Success Alert */
.stSuccess, [data-testid="stAlert"][data-baseweb*="positive"] {
    background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%) !important;
    border-right: 4px solid #28a745 !important;
}

/* Warning Alert */
.stWarning, [data-testid="stAlert"][data-baseweb*="warning"] {
    background: linear-gradient(135deg, #fff3cd 0%, #ffeeba 100%) !important;
    border-right: 4px solid #ffc107 !important;
}

/* Info Alert */
.stInfo, [data-testid="stAlert"][data-baseweb*="info"] {
    background: linear-gradient(135deg, #e7f3ff 0%, #cce5ff 100%) !important;
    border-right: 4px solid #667eea !important;
}

/* Error Alert */
.stError, [data-testid="stAlert"][data-baseweb*="negative"] {
    background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%) !important;
    border-right: 4px solid #dc3545 !important;
}

/* ===== DATA EDITOR ===== */
[data-testid="stDataEditor"] {
    border-radius: 16px !important;
    overflow: hidden;
    box-shadow: var(--shadow-light);
}

[data-testid="stDataEditor"] > div {
    border-radius: 16px;
}

/* ===== METRICS ===== */
[data-testid="stMetric"] {
    background: var(--card-bg);
    border-radius: 16px;
    padding: 1.2rem;
    box-shadow: var(--shadow-light);
    transition: var(--transition);
}

[data-testid="stMetric"]:hover {
    transform: translateY(-3px);
    box-shadow: var(--shadow-medium);
}

[data-testid="stMetricValue"] {
    font-size: 2rem !important;
    font-weight: 700 !important;
    background: var(--primary-gradient);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}

[data-testid="stMetricLabel"] {
    font-weight: 600 !important;
    color: #666 !important;
}

/* ===== EXPANDER ===== */
.streamlit-expanderHeader {
    direction: rtl;
    text-align: right;
    background: linear-gradient(135deg, #f8f9ff 0%, #f0f4ff 100%) !important;
    border-radius: 12px !important;
    font-weight: 600 !important;
    padding: 1rem 1.5rem !important;
    transition: var(--transition);
}

.streamlit-expanderHeader:hover {
    background: linear-gradient(135deg, #f0f4ff 0%, #e8edff 100%) !important;
}

.streamlit-expanderContent {
    direction: rtl;
    text-align: right;
    background: #fafbff;
    border-radius: 0 0 12px 12px;
    padding: 1.5rem !important;
}

/* ===== TEXT INPUTS ===== */
.stTextInput > div > div > input,
.stTextArea > div > div > textarea {
    direction: rtl;
    text-align: right;
    border-radius: 10px !important;
    border: 2px solid #e0e5ec !important;
    padding: 0.8rem 1rem !important;
    transition: var(--transition);
}

.stTextInput > div > div > input:focus,
.stTextArea > div > div > textarea:focus {
    border-color: #667eea !important;
    box-shadow: 0 0 0 3px rgba(102, 126, 234, 0.2) !important;
}

/* ===== SELECTBOX ===== */
.stSelectbox > div > div {
    direction: rtl;
    text-align: right;
    border-radius: 10px !important;
}

/* ===== SPINNER ===== */
.stSpinner > div {
    border-top-color: #667eea !important;
}

/* ===== SECTION HEADERS ===== */
.section-header {
    display: flex;
    align-items: center;
    gap: 12px;
    margin: 2rem 0 1.5rem 0;
    padding-bottom: 0.8rem;
    border-bottom: 3px solid;
    border-image: var(--primary-gradient) 1;
}

.section-header h2 {
    font-size: 1.6rem;
    font-weight: 700;
    color: #1a1a2e;
    margin: 0;
}

.section-icon {
    font-size: 1.8rem;
}

/* ===== STATUS BADGES ===== */
.status-badge {
    display: inline-flex;
    align-items: center;
    gap: 6px;
    padding: 0.4rem 1rem;
    border-radius: 20px;
    font-size: 0.9rem;
    font-weight: 600;
}

.status-pending {
    background: linear-gradient(135deg, #fff3cd 0%, #ffeeba 100%);
    color: #856404;
}

.status-resolved {
    background: linear-gradient(135deg, #d4edda 0%, #c3e6cb 100%);
    color: #155724;
}

.status-liked {
    background: linear-gradient(135deg, #f8d7da 0%, #f5c6cb 100%);
    color: #721c24;
}

.status-delete {
    background: linear-gradient(135deg, #e2e3e5 0%, #d6d8db 100%);
    color: #383d41;
}

/* ===== DIVIDER ===== */
hr {
    border: none;
    height: 2px;
    background: linear-gradient(90deg, transparent 0%, #667eea 50%, transparent 100%);
    margin: 2rem 0;
}

/* ===== ANIMATIONS ===== */
@keyframes fadeInUp {
    from {
        opacity: 0;
        transform: translateY(20px);
    }
    to {
        opacity: 1;
        transform: translateY(0);
    }
}

.animate-fade-in {
    animation: fadeInUp 0.5s ease-out;
}

/* ===== SCROLLBAR ===== */
::-webkit-scrollbar {
    width: 8px;
    height: 8px;
}

::-webkit-scrollbar-track {
    background: #f1f1f1;
    border-radius: 4px;
}

::-webkit-scrollbar-thumb {
    background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
    border-radius: 4px;
}

::-webkit-scrollbar-thumb:hover {
    background: linear-gradient(135deg, #5a6fd6 0%, #6a4190 100%);
}

/* ===== RESPONSIVE ===== */
@media (max-width: 768px) {
    .premium-header {
        padding: 1.5rem;
    }
    
    .premium-header h1 {
        font-size: 2rem !important;
    }
    
    .main .block-container {
        padding: 1rem;
    }
}
</style>
""", unsafe_allow_html=True)


# ============================================================
# Text Extraction Functions
# ============================================================

def extract_text_from_pptx(file_bytes: bytes) -> list[dict]:
    """Extract text from each slide of a PowerPoint presentation."""
    pptx_stream = io.BytesIO(file_bytes)
    prs = Presentation(pptx_stream)
    
    slides_data = []
    
    for slide_num, slide in enumerate(prs.slides, start=1):
        slide_text_parts = []
        
        for shape in slide.shapes:
            if hasattr(shape, "text") and shape.text.strip():
                slide_text_parts.append(shape.text.strip())
        
        slide_text = "\n".join(slide_text_parts)
        
        slides_data.append({
            "slide_number": slide_num,
            "text": slide_text
        })
    
    return slides_data


def extract_text_from_docx(file_bytes: bytes) -> str:
    """Extract text from a DOCX file."""
    docx_stream = io.BytesIO(file_bytes)
    doc = Document(docx_stream)
    
    paragraphs = [para.text for para in doc.paragraphs if para.text.strip()]
    return "\n\n".join(paragraphs)


def extract_text_from_txt(file_bytes: bytes) -> str:
    """Extract text from a TXT file with multiple encoding support."""
    for encoding in ["utf-8", "utf-8-sig", "windows-1255", "iso-8859-8", "latin-1"]:
        try:
            return file_bytes.decode(encoding)
        except UnicodeDecodeError:
            continue
    
    return file_bytes.decode("utf-8", errors="replace")


# ============================================================
# AI Analysis Functions
# ============================================================

def analyze_slides(slides_data: list[dict], context_text: str, model_name: str = "gemini-2.0-flash") -> list[dict]:
    """Analyze presentation slides using Google Gemini AI."""
    
    max_context_length = 3000
    if len(context_text) > max_context_length:
        context_text = context_text[:max_context_length] + "\n... (הטקסט קוצר)"
    
    slides_content = "\n\n".join([
        f"--- שקף {s['slide_number']} ---\n{s['text'][:500]}"
        for s in slides_data
    ])
    
    prompt = f"""אתה מומחה לסקירת מצגות עסקיות. נתח את המצגת הבאה.

## הקשר הפרויקט:
{context_text}

## תוכן המצגת:
{slides_content}

## המשימה:
תן ביקורת קצרה בעברית לכל שקף (2-3 משפטים בלבד).

## פורמט - החזר JSON בלבד:
[
  {{"slide_number": 1, "original_text": "טקסט קצר", "ai_comment": "הערה קצרה", "status": "לביצוע"}},
  {{"slide_number": 2, "original_text": "טקסט קצר", "ai_comment": "הערה קצרה", "status": "לביצוע"}}
]

חשוב: הערות קצרות בלבד, JSON תקין, בעברית.
"""

    model = genai.GenerativeModel(model_name)
    
    generation_config = genai.GenerationConfig(
        response_mime_type="application/json",
        temperature=0.3,
        max_output_tokens=4096
    )
    
    response = model.generate_content(
        prompt,
        generation_config=generation_config
    )
    
    response_text = response.text.strip()
    
    try:
        result = json.loads(response_text)
    except json.JSONDecodeError:
        if response_text.endswith(','):
            response_text = response_text[:-1]
        
        open_brackets = response_text.count('[') - response_text.count(']')
        open_braces = response_text.count('{') - response_text.count('}')
        
        if open_braces > 0:
            last_complete = response_text.rfind('},')
            if last_complete > 0:
                response_text = response_text[:last_complete + 1]
        
        response_text = response_text + ']' * open_brackets
        
        try:
            result = json.loads(response_text)
        except json.JSONDecodeError:
            result = [
                {
                    "slide_number": s["slide_number"],
                    "original_text": s["text"][:200] + "..." if len(s["text"]) > 200 else s["text"],
                    "ai_comment": "⚠️ לא ניתן היה לנתח שקף זה - נסה שוב",
                    "status": "לביצוע"
                }
                for s in slides_data
            ]
    
    return result


# ============================================================
# PPTX Modification Functions
# ============================================================

def add_comments_to_pptx(original_pptx_bytes: bytes, analyzed_data: list[dict]) -> bytes:
    """Add AI comments to the Speaker Notes of a PowerPoint presentation."""
    pptx_stream = io.BytesIO(original_pptx_bytes)
    prs = Presentation(pptx_stream)
    
    analysis_map = {item["slide_number"]: item for item in analyzed_data}
    
    for slide_idx, slide in enumerate(prs.slides):
        slide_number = slide_idx + 1
        
        if slide_number not in analysis_map:
            continue
        
        analysis = analysis_map[slide_number]
        status = analysis.get("status", "")
        ai_comment = analysis.get("ai_comment", "")
        
        if status in ["נפתר", "למחוק"]:
            continue
        
        if not ai_comment.strip():
            continue
        
        status_indicator = "🔴 לביצוע" if status == "לביצוע" else "💚 אהבתי"
        formatted_comment = f"{status_indicator} | AI Insight:\n{ai_comment}"
        
        if slide.has_notes_slide:
            notes_slide = slide.notes_slide
        else:
            notes_slide = slide.notes_slide
        
        notes_text_frame = notes_slide.notes_text_frame
        
        existing_notes = notes_text_frame.text if notes_text_frame.text else ""
        
        if existing_notes.strip():
            new_notes = f"{existing_notes}\n\n{'─'*40}\n{formatted_comment}"
        else:
            new_notes = formatted_comment
        
        notes_text_frame.clear()
        p = notes_text_frame.paragraphs[0]
        p.text = new_notes
    
    output_stream = io.BytesIO()
    prs.save(output_stream)
    output_stream.seek(0)
    
    return output_stream.getvalue()


def create_excel_report(analyzed_data: list[dict]) -> bytes:
    """Create an Excel report from the analysis data."""
    df = pd.DataFrame(analyzed_data)
    
    df = df.rename(columns={
        "slide_number": "מספר שקף",
        "original_text": "טקסט מקורי",
        "ai_comment": "הערת AI",
        "status": "סטטוס"
    })
    
    output = io.BytesIO()
    
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="ניתוח מצגת")
        
        worksheet = writer.sheets["ניתוח מצגת"]
        
        for idx, col in enumerate(df.columns):
            max_length = max(
                df[col].astype(str).apply(len).max(),
                len(col)
            )
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[chr(65 + idx)].width = adjusted_width
    
    output.seek(0)
    return output.getvalue()


# ============================================================
# Main Application
# ============================================================

def main():
    # --------------------------------------------------------
    # Premium Header
    # --------------------------------------------------------
    st.markdown("""
    <div class="premium-header">
        <h1>✨ מערכת סקירת מצגות AI</h1>
        <p>נתח את המצגת שלך באמצעות בינה מלאכותית וקבל תובנות מקצועיות תוך שניות</p>
    </div>
    """, unsafe_allow_html=True)
    
    # --------------------------------------------------------
    # How to Use - Expander
    # --------------------------------------------------------
    with st.expander("📖 מדריך למשתמש - לחץ כאן להוראות", expanded=False):
        st.markdown("""
        <div style="padding: 1rem;">
        
        ### 🚀 5 צעדים פשוטים לניתוח המצגת שלך:
        
        ---
        
        **1️⃣ העלאת המצגת**
        > העלה את קובץ ה-PowerPoint (PPTX) שברצונך לנתח
        
        **2️⃣ העלאת קובץ ההקשר**  
        > העלה קובץ טקסט (TXT) או מסמך Word (DOCX) עם הקשר הפרויקט, דרישות או מידע רלוונטי
        
        **3️⃣ הפעלת הניתוח**
        > לחץ על כפתור "נתח מצגת" והמתן לתוצאות (עד דקה)
        
        **4️⃣ סקירה ועריכה**
        > עבור על ההערות בטבלה, ערוך לפי הצורך ועדכן סטטוסים
        
        **5️⃣ הורדת התוצאות**
        > הורד את המצגת המעודכנת עם ההערות ב-Speaker Notes
        
        ---
        
        ### 📊 מקרא סטטוסים:
        
        | סימן | סטטוס | משמעות | יתווסף למצגת? |
        |------|-------|--------|---------------|
        | ⏳ | לביצוע | הערה פעילה שדורשת טיפול | ✅ כן |
        | ❤️ | אהבתי | הערה חיובית ששווה לשמור | ✅ כן |
        | ✅ | נפתר | טופל - אפשר להמשיך הלאה | ❌ לא |
        | 🗑️ | למחוק | לא רלוונטי - התעלם | ❌ לא |
        
        </div>
        """, unsafe_allow_html=True)
    
    # --------------------------------------------------------
    # Sidebar
    # --------------------------------------------------------
    with st.sidebar:
        st.markdown("## ⚙️ לוח בקרה")
        
        st.markdown("---")
        
        # API Status
        st.markdown("### 🔐 סטטוס חיבור")
        if GEMINI_API_KEY and GEMINI_API_KEY != "YOUR_API_KEY_HERE":
            st.success("✅ API מחובר")
        else:
            st.error("❌ נדרש מפתח API")
        
        st.markdown("---")
        
        # Model Selection
        st.markdown("### 🤖 מודל AI")
        model_choice = st.selectbox(
            "בחר מודל",
            options=["gemini-2.0-flash", "gemini-1.5-pro-latest", "gemini-1.5-flash-latest"],
            index=0,
            help="Flash = מהיר | Pro = מדויק יותר",
            label_visibility="collapsed"
        )
        
        model_descriptions = {
            "gemini-2.0-flash": "⚡ מהיר וחדש",
            "gemini-1.5-pro-latest": "🎯 מדויק ומעמיק",
            "gemini-1.5-flash-latest": "🚀 קל ומהיר"
        }
        st.caption(model_descriptions.get(model_choice, ""))
        
        st.markdown("---")
        
        # Stats
        st.markdown("### 📈 סטטיסטיקות")
        
        if "analysis_results" in st.session_state and st.session_state["analysis_results"]:
            df = pd.DataFrame(st.session_state["analysis_results"])
            status_counts = df["status"].value_counts()
            
            st.metric("סה״כ שקפים", len(df))
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("⏳ לביצוע", status_counts.get("לביצוע", 0))
                st.metric("✅ נפתר", status_counts.get("נפתר", 0))
            with col2:
                st.metric("❤️ אהבתי", status_counts.get("אהבתי", 0))
                st.metric("🗑️ למחוק", status_counts.get("למחוק", 0))
        else:
            st.info("📊 הנתונים יופיעו לאחר הניתוח")
        
        st.markdown("---")
        
        st.markdown("""
        <div style="text-align: center; opacity: 0.7; font-size: 0.8rem; margin-top: 2rem;">
            <p>🛠️ נבנה עם Streamlit</p>
            <p>🤖 מונע ע״י Google Gemini</p>
        </div>
        """, unsafe_allow_html=True)
    
    # --------------------------------------------------------
    # File Upload Section
    # --------------------------------------------------------
    st.markdown("""
    <div class="section-header">
        <span class="section-icon">📁</span>
        <h2>העלאת קבצים</h2>
    </div>
    """, unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("""
        <div class="premium-card">
            <div class="premium-card-header">
                <div class="premium-card-icon">📑</div>
                <h3 class="premium-card-title">מצגת לניתוח</h3>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        pptx_file = st.file_uploader(
            "העלה קובץ PowerPoint",
            type=["pptx"],
            key="pptx_uploader",
            help="גרור קובץ PPTX או לחץ לבחירה",
            label_visibility="collapsed"
        )
        
        if pptx_file:
            st.success(f"✅ {pptx_file.name}")
    
    with col2:
        st.markdown("""
        <div class="premium-card">
            <div class="premium-card-header">
                <div class="premium-card-icon">💬</div>
                <h3 class="premium-card-title">קובץ הקשר</h3>
            </div>
        </div>
        """, unsafe_allow_html=True)
        
        context_file = st.file_uploader(
            "העלה קובץ הקשר",
            type=["txt", "docx"],
            key="context_uploader",
            help="גרור קובץ TXT או DOCX או לחץ לבחירה",
            label_visibility="collapsed"
        )
        
        if context_file:
            st.success(f"✅ {context_file.name}")
    
    # Process uploaded files
    slides_data = None
    context_text = None
    
    if pptx_file is not None:
        pptx_file.seek(0)
        pptx_bytes = pptx_file.read()
        st.session_state["original_pptx_bytes"] = pptx_bytes
        slides_data = extract_text_from_pptx(pptx_bytes)
        st.session_state["slides_data"] = slides_data
    
    if context_file is not None:
        context_bytes = context_file.read()
        
        if context_file.name.endswith(".docx"):
            context_text = extract_text_from_docx(context_bytes)
        else:
            context_text = extract_text_from_txt(context_bytes)
        
        st.session_state["context_text"] = context_text
    
    if slides_data is None and "slides_data" in st.session_state:
        slides_data = st.session_state["slides_data"]
    
    if context_text is None and "context_text" in st.session_state:
        context_text = st.session_state["context_text"]
    
    # --------------------------------------------------------
    # Analysis Section
    # --------------------------------------------------------
    st.markdown("""
    <div class="section-header">
        <span class="section-icon">🔬</span>
        <h2>ניתוח AI</h2>
    </div>
    """, unsafe_allow_html=True)
    
    # Validation
    missing_items = []
    api_configured = GEMINI_API_KEY and GEMINI_API_KEY != "YOUR_API_KEY_HERE"
    if not api_configured:
        missing_items.append("🔑 מפתח API")
    if slides_data is None:
        missing_items.append("📑 קובץ מצגת")
    if context_text is None:
        missing_items.append("💬 קובץ הקשר")
    
    can_analyze = len(missing_items) == 0
    
    if missing_items:
        st.warning(f"⚠️ חסרים: {' • '.join(missing_items)}")
    else:
        st.success("✅ הכל מוכן! לחץ על הכפתור להתחלת הניתוח")
    
    # Analyze Button - Centered
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        analyze_button = st.button(
            "🚀 נתח מצגת",
            disabled=not can_analyze,
            type="primary",
            width="stretch"
        )
    
    # Run analysis
    if analyze_button and can_analyze:
        with st.spinner("🔄 מנתח את המצגת... אנא המתן"):
            try:
                analysis_results = analyze_slides(
                    slides_data,
                    context_text,
                    model_name=model_choice
                )
                st.session_state["analysis_results"] = analysis_results
                st.success("🎉 הניתוח הושלם בהצלחה!")
                st.balloons()
            except json.JSONDecodeError as e:
                st.error(f"❌ שגיאה בפענוח התשובה: {e}")
            except Exception as e:
                error_msg = str(e)
                if "429" in error_msg or "quota" in error_msg.lower():
                    st.error("❌ חריגה ממכסת השימוש. נסה שוב מאוחר יותר או החלף מפתח API.")
                elif "404" in error_msg:
                    st.error("❌ המודל לא נמצא. נסה מודל אחר מהרשימה.")
                else:
                    st.error(f"❌ שגיאה: {e}")
    
    # --------------------------------------------------------
    # Results Dashboard
    # --------------------------------------------------------
    if "analysis_results" in st.session_state and st.session_state["analysis_results"]:
        st.markdown("""
        <div class="section-header">
            <span class="section-icon">📋</span>
            <h2>תוצאות הניתוח</h2>
        </div>
        """, unsafe_allow_html=True)
        
        analysis_df = pd.DataFrame(st.session_state["analysis_results"])
        
        # Data Editor
        edited_df = st.data_editor(
            analysis_df,
            use_container_width=True,
            hide_index=True,
            num_rows="fixed",
            column_config={
                "slide_number": st.column_config.NumberColumn(
                    "🔢 שקף",
                    width="small",
                    disabled=True
                ),
                "original_text": st.column_config.TextColumn(
                    "📄 טקסט מקורי",
                    width="medium",
                    disabled=True
                ),
                "ai_comment": st.column_config.TextColumn(
                    "💬 הערת AI",
                    width="large",
                    disabled=False
                ),
                "status": st.column_config.SelectboxColumn(
                    "📊 סטטוס",
                    width="small",
                    options=["לביצוע", "נפתר", "אהבתי", "למחוק"],
                    required=True
                )
            },
            key="analysis_editor"
        )
        
        st.session_state["analysis_results"] = edited_df.to_dict("records")
        
        # --------------------------------------------------------
        # Download Section
        # --------------------------------------------------------
        st.markdown("""
        <div class="section-header">
            <span class="section-icon">📥</span>
            <h2>הורדת תוצאות</h2>
        </div>
        """, unsafe_allow_html=True)
        
        status_counts = edited_df["status"].value_counts()
        pending = status_counts.get("לביצוע", 0)
        liked = status_counts.get("אהבתי", 0)
        
        st.info(f"""
        **📊 סיכום:** {pending + liked} הערות יתווספו למצגת (לביצוע + אהבתי)
        """)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            if "original_pptx_bytes" in st.session_state:
                try:
                    modified_pptx = add_comments_to_pptx(
                        st.session_state["original_pptx_bytes"],
                        st.session_state["analysis_results"]
                    )
                    
                    st.download_button(
                        "📊 הורד מצגת",
                        data=modified_pptx,
                        file_name="reviewed_presentation.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        width="stretch",
                        type="primary"
                    )
                except Exception as e:
                    st.error(f"❌ {e}")
        
        with col2:
            try:
                excel_data = create_excel_report(st.session_state["analysis_results"])
                
                st.download_button(
                    "📑 הורד Excel",
                    data=excel_data,
                    file_name="analysis_checklist.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch"
                )
            except Exception as e:
                st.error(f"❌ {e}")
        
        with col3:
            json_data = json.dumps(
                st.session_state["analysis_results"],
                ensure_ascii=False,
                indent=2
            )
            st.download_button(
                "🔧 הורד JSON",
                data=json_data,
                file_name="analysis_results.json",
                mime="application/json",
                width="stretch"
            )
    
    # --------------------------------------------------------
    # Debug Section
    # --------------------------------------------------------
    with st.expander("🔧 תצוגה טכנית (למפתחים)"):
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### 📝 תוכן ההקשר")
            if context_text:
                preview = context_text[:1500] + "..." if len(context_text) > 1500 else context_text
                st.text_area("תצוגה מקדימה", value=preview, height=200, disabled=True, label_visibility="collapsed")
                st.caption(f"📊 {len(context_text):,} תווים")
            else:
                st.info("טרם נטען קובץ")
        
        with col2:
            st.markdown("#### 📊 שקפים שזוהו")
            if slides_data:
                st.dataframe(
                    pd.DataFrame(slides_data).rename(columns={"slide_number": "מס׳", "text": "טקסט"}),
                    width="stretch",
                    hide_index=True,
                    height=200
                )
                st.caption(f"📊 {len(slides_data)} שקפים")
            else:
                st.info("טרם נטענה מצגת")


# ============================================================
# Entry Point
# ============================================================

if __name__ == "__main__":
    main()