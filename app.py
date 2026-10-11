import streamlit as st
import json
from io import BytesIO

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

from google import genai

# ==================================================
# FAQAT SHU YERDA API KALITNI YOZING
# ==================================================
GEMINI_API_KEY = "AQ.Ab8RN6IzZgl3gy-fNXbYZePMEPU3DFwwaQhyBDEVdHRaSAMceQ"

# ==================================================
# SOZLAMALAR
# ==================================================
st.set_page_config(
    page_title="NeoBiology Multimedia AI",
    page_icon="🧬",
    layout="wide"
)

# ==================================================
# LOGOTIP
# ==================================================
try:
    st.image("logo.png", width=100)
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
# PPT YARATISH
# ==================================================
def create_presentation(slides):

    prs = Presentation()

    for idx, item in enumerate(slides):

        slide = prs.slides.add_slide(
            prs.slide_layouts[1]
        )

        slide.shapes.title.text = item["title"]

        slide.placeholders[1].text = item["content"]

    ppt_io = BytesIO()
    prs.save(ppt_io)

    return ppt_io.getvalue()

# ==================================================
# GEMINI
# ==================================================
def generate_slides(topic):

    client = genai.Client(
        api_key=GEMINI_API_KEY
    )

    prompt = f"""
    Mavzu: {topic}

    10 ta slayd tuz.

    Faqat JSON qaytar.

    {{
      "slides":[
        {{
          "title":"...",
          "content":"..."
        }}
      ]
    }}
    """

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt
    )

    text = response.text

    text = text.replace("```json", "")
    text = text.replace("```", "")

    return json.loads(text)

# ==================================================
# BUTTON
# ==================================================
if st.button("🚀 PPT yaratish"):

    if not GEMINI_API_KEY.startswith("AIza"):

        st.error(
            "Haqiqiy Gemini API kalit kiriting."
        )

    else:

        try:

            with st.spinner(
                "Slaydlar yaratilmoqda..."
            ):

                data = generate_slides(topic)

                ppt_data = create_presentation(
                    data["slides"]
                )

            st.success(
                "✅ PPT tayyor!"
            )

            st.download_button(
                label="📥 PPT yuklab olish",
                data=ppt_data,
                file_name=f"{topic}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

        except Exception as e:

            st.error(
                f"Xatolik: {e}"
            )
