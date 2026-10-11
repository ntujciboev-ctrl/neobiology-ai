import streamlit as st
import json
from io import BytesIO
import streamlit as st

st.error("TEST_2026")
from pptx import Presentation
from pptx.util import Pt

from google import genai


# ==================================================
# FAQAT SHU YERDA API KALITNI YOZING
# ==================================================
GEMINI_API_KEY = "AQ.Ab8RN6IzZgl3gy-fNXbYZePMEPU3DFwwaQhyBDEVdHRaSAMceQ"


# ==================================================
# SAHIFA SOZLAMALARI
# ==================================================
st.set_page_config(
    page_title="NeoBiology Multimedia AI",
    page_icon="🧬",
    layout="wide"
)

# ==================================================
# LOGOTIP
# logo.png faylni app.py bilan bir papkaga joylashtiring
# ==================================================
try:
    st.image("logo.png", width=120)
except:
    st.write("🧬")

st.title("NeoBiology Multimedia AI")
st.write("Gemini AI yordamida avtomatik PPT yaratish tizimi")

# ==================================================
# MAVZU
# ==================================================
topic = st.text_input(
    "Prezentatsiya mavzusi:",
    "Yurak qon tomir tizimi"
)

# ==================================================
# GEMINI ORQALI MAZMUN YARATISH
# ==================================================
def generate_slides(topic):

    client = genai.Client(
        api_key=GEMINI_API_KEY
    )

    prompt = f"""
    '{topic}' mavzusida 10 ta slayd yarating.

    Faqat JSON formatida qaytaring.

    {{
      "slides":[
        {{
          "title":"Slayd sarlavhasi",
          "content":"2-4 ta qisqa punkt"
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


# ==================================================
# PPT YARATISH
# ==================================================
def create_ppt(slides):

    prs = Presentation()

    for slide_item in slides:

        slide = prs.slides.add_slide(
            prs.slide_layouts[1]
        )

        slide.shapes.title.text = slide_item["title"]

        slide.placeholders[1].text = slide_item["content"]

        title = slide.shapes.title.text_frame

        for paragraph in title.paragraphs:
            paragraph.font.size = Pt(24)

    ppt_buffer = BytesIO()

    prs.save(ppt_buffer)

    return ppt_buffer.getvalue()


# ==================================================
# TUGMA
# ==================================================
if st.button("🚀 PPT yaratish"):

    if not GEMINI_API_KEY.strip():

        st.error(
            "API kalit kiritilmagan."
        )

    else:

        try:

            with st.spinner(
                "Gemini AI slaydlarni yaratmoqda..."
            ):

                data = generate_slides(topic)

                ppt_file = create_ppt(
                    data["slides"]
                )

            st.success(
                "✅ PPT muvaffaqiyatli yaratildi!"
            )

            st.download_button(
                label="📥 PPT yuklab olish",
                data=ppt_file,
                file_name=f"{topic}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

        except json.JSONDecodeError:

            st.error(
                "Gemini JSON formatida javob qaytarmadi."
            )

        except Exception as e:

            st.error(
                f"Xatolik: {e}"
            )
