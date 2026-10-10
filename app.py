import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from google import genai
import json

# API KALITINGIZNI SHU YERGA QO'YING:
GEMINI_API_KEY = "AQ.Ab8RN6IzZgl3gy-fNXbYZePMEPU3DFwwaQhyBDEVdHRaSAMceQ"

# --- SERVER UCHUN MUKAMMAL LOGOTIP LOGIKASI ---
# Dunyodagi har qanday brauzer yuklay oladigan xavfsiz va barqaror havola
LOGO_URL = "LOGO_URL = "https://githubusercontent.com"
"

st.set_page_config(page_title="NeoBiology AI", page_icon=LOGO_URL, layout="wide")

# Interfeys qismi
col1, col2 = st.columns([1, 8])

with col1:
    st.image(LOGO_URL, width=100) # Sayt tepasidagi dumaloq yashil rasm

with col2:
    st.markdown("<h1 style='margin:0; padding:0;'>NeoBiology — Professional 3D AI Presentation Maker</h1>", unsafe_allow_html=True)
    st.write("Sarlavhalari Schoolbell, matnlari esa Comic Sans shriftida bo'lgan 3D ramkali zamonaviy platforma.")

topic = st.text_input("Prezentatsiya mavzusini kiriting:", "Prokariot hujayra va mikrobiologiya")

def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return RGBColor(*(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))

if st.button("3D Formatda Yaratish 🚀"):
    st.info("🤖 NeoBiology AI tizimi mukammal 3D dizayn va matnlarni siz uchun tayyorlamoqda...")
    
    try:
        client = genai.Client(api_key=GEMINI_API_KEY)
        
        prompt = f"""
        Mavzu: '{topic}'.
        Ushbu mavzu bo'yicha kamida 10 ta, ko'pi bilan 12 ta slayddan iborat juda batafsil taqdimot matnini tayyorla.
        Mavzu tabiatidan kelib chiqib 3 ta mos rang (HEX formatida) tanla:
        1. 'background_dark' (To'q fon uchun)
        2. 'accent_color' (Yorqin elementlar va ramkalar uchun)
        3. 'text_dark' (Matnlar uchun asosiy quyuq rang)

        Har bir slayd uchun aniq 'title' (Sarlavha) va 'content' (kamida 3-4 ta qisqa nuqtadan iborat matn) bo'lsin.
        Javobni faqat va faqat quyidagi JSON formatida qaytar, hech qanday qo'shimcha gap yozma:
        {{
            "colors": {{
                "background_dark": "#HEX_KOD",
                "accent_color": "#HEX_KOD",
                "text_dark": "#HEX_KOD"
            }},
            "slides": [
                {{"title": "1-slayd sarlavhasi", "content": "Tezis 1\\nTezis 2"}}
            ]
        }}
        Javob toza o'zbek tilida va qat'iy JSON formatida bo'lishi shart.
        """
        
        response = client.models.generate_content(model='gemini-3.8-flash', contents=prompt)
        raw_text = response.text.strip().replace("```json", "").replace("```", "")
        data = json.loads(raw_text)
        
        color_config = data["colors"]
        BG_DARK = hex_to_rgb(color_config["background_dark"])
        ACCENT = hex_to_rgb(color_config["accent_color"])
        TXT_DARK = hex_to_rgb(color_config["text_dark"])
        
        SHADOW_COLOR = RGBColor(203, 213, 225) 
        LIGHT_GRAY = RGBColor(248, 250, 252)    
        
        prs = Presentation()
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
        blank_layout = prs.slide_layouts[6]
        
        for index, slide_data in enumerate(data["slides"]):
            slide = prs.slides.add_slide(blank_layout)
            title_text = str(slide_data["title"])
            body_text = str(slide_data["content"])
            
            if index == 0 or index == len(data["slides"]) - 1:
                bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
                bg.fill.solid()
                bg.fill.fore_color.rgb = BG_DARK
                bg.line.fill.background()
                
                shadow = slide.shapes.add_shape(1, Inches(0.45), Inches(0.45), prs.slide_width - Inches(0.8), prs.slide_height - Inches(0.8))
                shadow.fill.background()
                shadow.line.color.rgb = RGBColor(30, 41, 59)
                shadow.line.width = Pt(3)
                
                border = slide.shapes.add_shape(1, Inches(0.4), Inches(0.4), prs.slide_width - Inches(0.8), prs.slide_height - Inches(0.8))
                border.fill.background()
                border.line.color.rgb = ACCENT
                border.line.width = Pt(3)
                
                t_box = slide.shapes.add_textbox(Inches(1.5), Inches(2.8), Inches(10.33), Inches(2))
                tf = t_box.text_frame
                tf.word_wrap = True
                p = tf.paragraphs
                p.text = title_text
                p.font.size = Pt(48)
                p.font.name = "Schoolbell"
                p.font.color.rgb = ACCENT
                p.alignment = PP_ALIGN.CENTER
            else:
                bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
                bg.fill.solid()
                bg.fill.fore_color.rgb = LIGHT_GRAY
                bg.line.fill.background()
                
                shadow_border = slide.shapes.add_shape(1, Inches(0.55), Inches(0.55), prs.slide_width - Inches(1.0), prs.slide_height - Inches(1.0))
                shadow_border.fill.background()
                shadow_border.line.color.rgb = SHADOW_COLOR
                shadow_border.line.width = Pt(2.5)
                
                main_border = slide.shapes.add_shape(1, Inches(0.5), Inches(0.5), prs.slide_width - Inches(1.0), prs.slide_height - Inches(1.0))
                main_border.fill.background()
                main_border.line.color.rgb = ACCENT
                main_border.line.width = Pt(2.5)
                
                t_box = slide.shapes.add_textbox(Inches(1.2), Inches(0.6), Inches(11), Inches(1))
                tf = t_box.text_frame
                tf.word_wrap = True
                p = tf.paragraphs
                p.text = title_text
                p.font.size = Pt(34)
                p.font.name = "Schoolbell"
                p.font.color.rgb = BG_DARK
                
                b_box = slide.shapes.add_textbox(Inches(1.2), Inches(2.2), Inches(11), Inches(4.5))
                tf_b = b_box.text_frame
                tf_b.word_wrap = True
                
                lines = [l.strip() for l in body_text.split("\n") if l.strip()]
                for idx, line in enumerate(lines):
                    p_b = tf_b.paragraphs if idx == 0 else tf_b.add_paragraph()
                    p_b.text = line.replace("- ", "• ").replace("*", "")
                    p_b.font.size = Pt(19)
                    p_b.font.name = "Comic Sans MS"
                    p_b.font.color.rgb = TXT_DARK
                    p_b.space_after = Pt(14)
                    p_b.line_spacing = 1.2
                    
                try:
                    # Slayd ichiga ham logotipni joylash
                    import requests
                    from io import BytesIO
                    res = requests.get(LOGO_URL)
                    slide.shapes.add_picture(BytesIO(res.content), Inches(11.8), Inches(0.8), width=Inches(0.9))
                except:
                    pass
        
        file_name = "neobiology_presentation.pptx"
        prs.save(file_name)
        st.success("🎉 Taqdimot muvaffaqiyatli tayyorlandi!")
        
        with open(file_name, "rb") as file:
            st.download_button(label="Yuklab olish (.pptx)", data=file, file_name=file_name, mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    except Exception as e:
        st.error(f"Xatolik: {e}")
