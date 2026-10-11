import streamlit as st
import json
from io import BytesIO
import google.generativeai as genai
from pptx import Presentation

# ==========================================
# AQ... TOKENINGIZNI SHU YERGA YOZING
# ==========================================
GEMINI_API_KEY = "AQ.Ab8RN6Luk6ZAsRo1wbOlhy4sYOroqv9IZRFC8egA51eliQ8UJw"  # O'zingizdagi AQ... tokenni shu yerga qo'ying

# Konfiguratsiya qilish
genai.configure(api_key=GEMINI_API_KEY.strip())

# ==========================================
# PAGE CONFIG
# ==========================================
st.set_page_config(
    page_title="NeoBiology AI",
    page_icon="🧬",
    layout="wide"
)

st.title("🧬 NeoBiology AI")

topic = st.text_input(
    "Mavzu kiriting:",
    "Prokariot hujayra va mikrobiologiya"
)

# ==========================================
# GEMINI FUNCTION
# ==========================================
def generate_slides(topic):
    # Eski va barqaror modeldan foydalanamiz
    model = genai.GenerativeModel('gemini-1.5-flash')

    prompt = f"""
    '{topic}' mavzusida 10 ta slayd tayyorla.
    Faqat JSON qaytar.
    {{
      "slides":[
        {{
          "title":"Sarlavha",
          "content":"Matn"
        }}
      ]
    }}
    """

    response = model.generate_content(prompt)
    text = response.text.strip()
    
    text = text.replace("```json", "")
    text = text.replace("```", "")

    return json.loads(text)

# ==========================================
# PPT FUNCTION
# ==========================================
def create_ppt(data):
    prs = Presentation()

    for slide_data in data["slides"]:
        slide = prs.slides.add_slide(
            prs.slide_layouts[1]
        )
        slide.shapes.title.text = slide_data["title"]
        slide.placeholders[1].text = slide_data["content"]

    buffer = BytesIO()
    prs.save(buffer)
    return buffer.getvalue()

# ==========================================
# BUTTON & EXECUTION
# ==========================================
if st.button("🚀 PPT yaratish"):
    try:
        st.info("Gemini ishlamoqda...")

        data = generate_slides(topic)
        ppt_file = create_ppt(data)

        st.success("✅ Tayyor!")

        st.download_button(
            label="📥 PPT yuklab olish",
            data=ppt_file,
            file_name="presentation.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )

    except Exception as e:
        st.error(f"Xatolik yuz berdi: {str(e)}")
