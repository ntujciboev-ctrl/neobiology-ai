import streamlit as st
import json
import os
from io import BytesIO
from google import genai
from pptx import Presentation

# ==========================================
# AQ... BILAN BOSHLANUVCHI API KALITINGIZ
# ==========================================
GEMINI_API_KEY = "AQ.Ab8RN6KdRyvoAOGvkcIZ_4MNaFIwm21JeJunpZwXpW7cjNP7WQ"

# Tokenni muhit o'zgaruvchisiga o'rnatamiz
os.environ["GEMINI_API_KEY"] = GEMINI_API_KEY.strip()

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
    # Klientni avtomatik ravishda muhitdagi kalit orqali ishga tushiramiz
    client = genai.Client()

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

    response = client.models.generate_content(
        model="gemini-3.8-flash",
        contents=prompt
    )

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
