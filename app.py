import streamlit as st
import json
from io import BytesIO

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE

from google import genai


# ==========================
# GEMINI API
# ==========================

try:
    GEMINI_API_KEY = st.secrets["AQ.Ab8RN6IzZgl3gy-fNXbYZePMEPU3DFwwaQhyBDEVdHRaSAMceQ"]
except:
    GEMINI_API_KEY = "AQ.Ab8RN6IzZgl3gy-fNXbYZePMEPU3DFwwaQhyBDEVdHRaSAMceQ"

# ==========================
# PAGE SETTINGS
# ==========================

st.set_page_config(
    page_title="NeoBiology AI",
    page_icon="🧬",
    layout="wide"
)

st.title("🧬 NeoBiology Multimedia AI")
st.write("Gemini AI yordamida avtomatik PPT yaratish tizimi")

topic = st.text_input(
    "Prezentatsiya mavzusi:",
    "Yurak qon tomir tizimi"
)


# ==========================
# HELPERS
# ==========================

def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip("#")

    return RGBColor(
        int(hex_color[0:2], 16),
        int(hex_color[2:4], 16),
        int(hex_color[4:6], 16)
    )


def generate_content(topic):

    client = genai.Client(
        api_key=GEMINI_API_KEY
    )

    prompt = f"""
    Mavzu: {topic}

    Ushbu mavzu uchun 10 ta slaydli taqdimot yarat.

    Quyidagi JSON formatdan chiqma:

    {{
      "colors": {{
        "background_dark": "#1E293B",
        "accent_color": "#22C55E",
        "text_dark": "#0F172A"
      }},
      "slides": [
        {{
          "title": "Sarlavha",
          "content": "1-qator\\n2-qator"
        }}
      ]
    }}

    Faqat JSON qaytar.
    """

    response = client.models.generate_content(
        model="gemini-2.5-flash",
        contents=prompt
    )

    raw = response.text.strip()

    raw = raw.replace("```json", "")
    raw = raw.replace("```", "")

    return json.loads(raw)


def create_presentation(data):

    prs = Presentation()

    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    blank = prs.slide_layouts[6]

    colors = data["colors"]

    BG = hex_to_rgb(colors["background_dark"])
    ACCENT = hex_to_rgb(colors["accent_color"])
    TEXT = hex_to_rgb(colors["text_dark"])

    for i, slide_data in enumerate(data["slides"]):

        slide = prs.slides.add_slide(blank)

        title = slide_data["title"]
        content = slide_data["content"]

        bg = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            0,
            0,
            prs.slide_width,
            prs.slide_height
        )

        if i == 0:

            bg.fill.solid()
            bg.fill.fore_color.rgb = BG
            bg.line.fill.background()

            title_box = slide.shapes.add_textbox(
                Inches(1),
                Inches(2.5),
                Inches(11),
                Inches(2)
            )

            p = title_box.text_frame.paragraphs[0]

            p.text = title
            p.alignment = PP_ALIGN.CENTER

            p.font.size = Pt(40)
            p.font.bold = True
            p.font.color.rgb = ACCENT

        else:

            bg.fill.solid()
            bg.fill.fore_color.rgb = RGBColor(
                248, 250, 252
            )

            bg.line.fill.background()

            title_box = slide.shapes.add_textbox(
                Inches(0.8),
                Inches(0.4),
                Inches(10),
                Inches(1)
            )

            tp = title_box.text_frame.paragraphs[0]

            tp.text = title
            tp.font.size = Pt(28)
            tp.font.bold = True
            tp.font.color.rgb = BG

            body_box = slide.shapes.add_textbox(
                Inches(0.8),
                Inches(1.5),
                Inches(8),
                Inches(4)
            )

            tf = body_box.text_frame

            lines = content.split("\n")

            first = True

            for line in lines:

                if not line.strip():
                    continue

                if first:
                    p = tf.paragraphs[0]
                    first = False
                else:
                    p = tf.add_paragraph()

                p.text = f"• {line}"

                p.font.size = Pt(18)
                p.font.color.rgb = TEXT

    ppt_buffer = BytesIO()

    prs.save(ppt_buffer)

    return ppt_buffer.getvalue()


# ==========================
# UI
# ==========================

if st.button("🚀 PPT yaratish"):

    if not GEMINI_API_KEY:

        st.error(
            "Streamlit secrets ichiga GEMINI_API_KEY qo'shing."
        )

    else:

        try:

            with st.spinner("Gemini AI ishlamoqda..."):

                data = generate_content(topic)

                ppt_data = create_presentation(data)

            st.success("✅ PPT tayyor!")

            st.download_button(
                label="📥 PPT yuklab olish",
                data=ppt_data,
                file_name=f"{topic}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

        except json.JSONDecodeError:

            st.error(
                "Gemini JSON formatda javob qaytarmadi."
            )

        except Exception as e:

            st.error(
                f"Xatolik: {e}"
            )
