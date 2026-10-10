import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.chart import XL_CHART_TYPE
from pptx.chart.data import CategoryChartData
from google import genai
import json
import os
import requests
from io import BytesIO

# --- API KALITNI STREAMLIT SECRETS ORQALI XAVFSIZ O'QISH ---
try:
    GEMINI_API_KEY = st.secrets["AQ.Ab8RN6IzZgl3gy-fNXbYZePMEPU3DFwwaQhyBDEVdHRaSAMceQ"]
except:
    GEMINI_API_KEY = "KALIT_SOZLANMAGAN"

# --- LOGOTIP NOMINI AVTOMATIK ANIQLASH ---
logo_file = None
possible_logos = ["logo.png", "logo.png.png", "logo.PNG", "logo.jpg"]
for name in possible_logos:
    if os.path.exists(name):
        logo_file = name
        break

if logo_file:
    st.set_page_config(page_title="NeoBiology Multimedia AI", page_icon=logo_file, layout="wide")
else:
    st.set_page_config(page_title="NeoBiology Multimedia AI", page_icon="🧬", layout="wide")

# ---- INTERFEYS DIZAYNI ----
col1, col2 = st.columns()
with col1:
    if logo_file:
        st.image(logo_file, width=90)
    else:
        st.markdown("<h1 style='color: #10b981; margin: 0;'>🧬</h1>", unsafe_allow_html=True)
with col2:
    st.markdown("<h1 style='margin: 0; padding: 0;'>NeoBiology — Multimedia & 3D AI Presentation Maker</h1>", unsafe_allow_html=True)
    st.write("Sarlavhalari Schoolbell, matnlari Comic Sans, rasmlar va diagrammalarni avtomatik joylovchi super platforma.")

topic = st.text_input("Prezentatsiya mavzusini kiriting:", "Yurak qon-tomir tizimi va uning statistikasi")

def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return RGBColor(*(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))

if st.button("Multimedia 3D Slayd Yaratish 🚀"):
    if GEMINI_API_KEY == "KALIT_SOZLANMAGAN" or GEMINI_API_KEY == "":
        st.error("Iltimos, avval Streamlit Cloud panelidagi 'Manage app' -> 'Settings' -> 'Secrets' qismiga GEMINI_API_KEY kalitingizni kiriting!")
    else:
        st.info("🤖 NeoBiology AI matnlar tuzmoqda, rasmlar qidirmoqda va diagrammalar chizmoqda...")
        
        try:
            client = genai.Client(api_key=GEMINI_API_KEY)
            
            prompt = f"""
            Mavzu: '{topic}'.
            Ushbu mavzu bo'yicha kamida 10 ta slayddan iborat multimedia taqdimotini tayyorla.
            Mavzu tabiatidan kelib chiqib 3 ta mos rang (HEX formatida) tanla:
            1. 'background_dark' (To'q fon uchun)
            2. 'accent_color' (Yorqin elementlar va ramkalar uchun)
            3. 'text_dark' (Matnlar uchun)

            Har bir slayd uchun aniq 'title', 'content', 'image_keyword' (slaydga mos inglizcha 1-2 so'zli rasm qidiruv kaliti, masalan 'heart cells') va 'has_chart' (true/false - agarda slaydda statistika, taqqoslash yoki foizlar bo'lsa true qil) bo'lsin.
            Chart ma'lumotlari uchun 'chart_data' kirit, masalan: {{"categories": ["Erkaklar", "Ayollar"], "values": [45, 55]}}. Agarda has_chart false bo'lsa buni null qil.
            Javobni faqat quyidagi JSON formatida qaytar, hech qanday qo'shimcha gap yozma:
            {{
                "colors": {{
                    "background_dark": "#HEX",
                    "accent_color": "#HEX",
                    "text_dark": "#HEX"
                }},
                "slides": [
                    {{
                        "title": "Slayd sarlavhasi", 
                        "content": "Tezis 1\\nTezis 2",
                        "image_keyword": "heart anatomy",
                        "has_chart": false,
                        "chart_data": null
                    }}
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
                title_text = str(slide_data["title"]).replace("**", "").replace("#", "")
                body_text = str(slide_data["content"])
                img_keyword = slide_data.get("image_keyword", "biology")
                has_chart = slide_data.get("has_chart", False)
                
                # ---- 3D RAMKALAR VA FON ----
                if index == 0 or index == len(data["slides"]) - 1:
                    bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
                    bg.fill.solid(); bg.fill.fore_color.rgb = BG_DARK; bg.line.fill.background()
                    
                    shadow = slide.shapes.add_shape(1, Inches(0.45), Inches(0.45), prs.slide_width - Inches(0.8), prs.slide_height - Inches(0.8))
                    shadow.fill.background(); shadow.line.color.rgb = RGBColor(30, 41, 59); shadow.line.width = Pt(3)
                    
                    border = slide.shapes.add_shape(1, Inches(0.4), Inches(0.4), prs.slide_width - Inches(0.8), prs.slide_height - Inches(0.8))
                    border.fill.background(); border.line.color.rgb = ACCENT; border.line.width = Pt(3)
                    
                    t_box = slide.shapes.add_textbox(Inches(1.5), Inches(2.8), Inches(10.33), Inches(2))
                    p = t_box.text_frame.paragraphs[0]
                    p.text = title_text; p.font.size = Pt(48); p.font.bold = True; p.font.name = "Schoolbell"; p.font.color.rgb = ACCENT; p.alignment = PP_ALIGN.CENTER
                else:
                    bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
                    bg.fill.solid(); bg.fill.fore_color.rgb = LIGHT_GRAY; bg.line.fill.background()
                    
                    shadow_border = slide.shapes.add_shape(1, Inches(0.55), Inches(0.55), prs.slide_width - Inches(1.0), prs.slide_height - Inches(1.0))
                    shadow_border.fill.background(); shadow_border.line.color.rgb = SHADOW_COLOR; shadow_border.line.width = Pt(2.5)
                    
                    main_border = slide.shapes.add_shape(1, Inches(0.5), Inches(0.5), prs.slide_width - Inches(1.0), prs.slide_height - Inches(1.0))
                    main_border.fill.background(); main_border.line.color.rgb = ACCENT; main_border.line.width = Pt(2.5)
                    
                    t_box = slide.shapes.add_textbox(Inches(1.2), Inches(0.6), Inches(11), Inches(1))
                    p = t_box.text_frame.paragraphs[0]
                    p.text = title_text; p.font.size = Pt(34); p.font.bold = True; p.font.name = "Schoolbell"; p.font.color.rgb = BG_DARK
                    
                    # JURNAL MULTIMEDIA VERSTKASI: Matn chapda
                    b_box = slide.shapes.add_textbox(Inches(1.2), Inches(2.2), Inches(6.5), Inches(4.5))
                    tf_b = b_box.text_frame; tf_b.word_wrap = True
                    
                    lines = [l.strip() for l in body_text.split("\n") if l.strip()]
                    for idx, line in enumerate(lines):
                        p_b = tf_b.paragraphs[0] if idx == 0 else tf_b.add_paragraph()
                        p_b.text = line.replace("- ", "• ").replace("*", "")
                        p_b.font.size = Pt(18); p_b.font.name = "Comic Sans MS"; p_b.font.color.rgb = TXT_DARK; p_b.space_after = Pt(12)
                    
                    # ---- MULTIMEDIA 1: AQLLI RASM QO'SHISH ----
                    if not has_chart:
                        try:
                            img_url = "https://unsplash.com"
                            search_url = f"https://unsplash.com{img_keyword}&per_page=1&client_id=p_Yv4wVv-XqM-HwG5HqRfZC-CqO-E498_r6VpXh1Y2Y"
                            res_search = requests.get(search_url, timeout=5).json()
                            if res_search.get("results"):
                                img_url = res_search["results"][0]["urls"]["small"]
                            
                            img_res = requests.get(img_url, timeout=5)
                            if img_res.status_code == 200:
                                img_stream = BytesIO(img_res.content)
                                slide.shapes.add_picture(img_stream, Inches(8.2), Inches(2.2), width=Inches(4.2), height=Inches(4.2))
                        except:
                            pass
                    
                    # ---- MULTIMEDIA 2: AQLLI DYNAMIC DIAGRAMMA CHIZISH ----
                    if has_chart and slide_data.get("chart_data"):
                        try:
                            c_data = slide_data["chart_data"]
                            chart_data = CategoryChartData()
                            chart_data.categories = c_data["categories"]
                            chart_data.add_series('Ko\'rsatkich', tuple(c_data["values"]))
                            
                            x, y, cx, cy = Inches(8.0), Inches(2.2), Inches(4.5), Inches(4.2)
                            chart = slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data).chart
                            chart.has_legend = False
                        except:
                            pass
                            
                    # ---- LOGOTIPNI JOYLASHTIRISH ----
                       finally
