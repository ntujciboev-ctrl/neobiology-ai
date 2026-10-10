import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from google import genai
import json
import requests
from io import BytesIO

# Toza va ishchi logotip havolasi
LOGO_URL = "https://githubusercontent.com"

st.set_page_config(page_title="NeoBiology Multimedia AI", page_icon="🧬", layout="wide")

# ---- INTERFEYS DIZAYNI ----
col1, col2 = st.columns([1, 4])
with col1:
    try:
        st.image(LOGO_URL, width=90)
    except:
        st.write("🧬 LOGO")
with col2:
    st.markdown("<h1 style='margin: 0; padding: 0;'>NeoBiology — Multimedia & 3D AI Presentation Maker</h1>", unsafe_allow_html=True)
    st.write("Sarlavhalari Schoolbell, matnlari Comic Sans, rasmlar va diagrammalarni avtomatik joylovchi platforma.")

# API kalit yozish uchun yagona joy interfeysning o'zida
GEMINI_API_KEY = st.text_input("AQ.Ab8RN6IzZgl3gy-fNXbYZePMEPU3DFwwaQhyBDEVdHRaSAMceQ:", type="password", help="Kalitni https://google.com saytidan olishingiz mumkin.")

topic = st.text_input("Prezentatsiya mavzusini kiriting:", "Yurak qon-tomir tizimi va uning statistikasi")

def hex_to_rgb(hex_str):
    hex_str = hex_str.lstrip('#')
    return RGBColor(*(int(hex_str[i:i+2], 16) for i in (0, 2, 4)))

# Streamlit sessiyasida fayl holatlarini saqlash
if "file_ready" not in st.session_state:
    st.session_state.file_ready = False
if "ppt_data" not in st.session_state:
    st.session_state.ppt_data = None

if st.button("Multimedia 3D Slayd Yaratish 🚀"):
    if not GEMINI_API_KEY.strip():
        st.error("Iltimos, avval Google Gemini API kalitingizni tepadagi maydonga kiriting!")
    else:
        st.info("🤖 NeoBiology AI matnlar tuzmoqda, rasmlar qidirmoqda va slaydlar tayyorlamoqda...")
        st.session_state.file_ready = False
        
        try:
            client = genai.Client(api_key=GEMINI_API_KEY.strip())
            
            prompt = f"""
            Mavzu: '{topic}'.
            Ushbu mavzu bo'yicha kamida 10 ta slayddan iborat o'quv multimedia taqdimotini tayyorla.
            Mavzu tabiatidan kelib chiqib 3 ta mos rang (HEX formatida) tanla:
            1. 'background_dark' (To'q fon uchun)
            2. 'accent_color' (Yorqin elementlar va ramkalar uchun)
            3. 'text_dark' (Matnlar uchun)

            Har bir slayd uchun aniq 'title', 'content', 'image_keyword' (slaydga mos inglizcha 1-2 so'zli rasm qidiruv kaliti) bo'lsin.
            Javobni faqat quyidagi JSON formatida qaytar, hech qanday qo'shimcha gap yoki markdown belgilari yozma:
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
                        "image_keyword": "heart anatomy"
                    }}
                ]
            }}
            Javob toza o'zbek tilida va qat'iy JSON formatida bo'lishi shart.
            """
            
            # Barqaror va eng oxirgi Gemini modeli ishlatildi
            response = client.models.generate_content(model='gemini-2.5-flash', contents=prompt)
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
            # TO'G'RILANDI: Bo'sh slayd andozasi uchun indeks (6) to'g'ri ko'rsatildi
            blank_layout = prs.slide_layouts[6]
            
            for index, slide_data in enumerate(data["slides"]):
                slide = prs.slides.add_slide(blank_layout)
                title_text = str(slide_data["title"]).replace("**", "").replace("#", "")
                body_text = str(slide_data["content"])
                img_keyword = slide_data.get("image_keyword", "biology")
                
                # Titul va yakuniy slaydlar dizayni
                if index == 0 or index == len(data["slides"]) - 1:
                    bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
                    bg.fill.solid(); bg.fill.fore_color.rgb = BG_DARK; bg.line.fill.background()
                    
                    shadow = slide.shapes.add_shape(1, Inches(0.45), Inches(0.45), prs.slide_width - Inches(0.9), prs.slide_height - Inches(0.9))
                    shadow.fill.background(); shadow.line.color.rgb = RGBColor(30, 41, 59); shadow.line.width = Pt(3)
                    
                    border = slide.shapes.add_shape(1, Inches(0.4), Inches(0.4), prs.slide_width - Inches(0.8), prs.slide_height - Inches(0.8))
                    border.fill.background(); border.line.color.rgb = ACCENT; border.line.width = Pt(3)
                    
                    t_box = slide.shapes.add_textbox(Inches(1.5), Inches(2.8), Inches(10.33), Inches(2))
                    p = t_box.text_frame.paragraphs[0]
                    p.text = title_text; p.font.size = Pt(48); p.font.bold = True; p.font.name = "Schoolbell"; p.font.color.rgb = ACCENT; p.alignment = PP_ALIGN.CENTER
                
                # Ichki slaydlar dizayni
                else:
                    bg = slide.shapes.add_shape(1, 0, 0, prs.slide_width, prs.slide_height)
                    bg.fill.solid(); bg.fill.fore_color.rgb = LIGHT_GRAY; bg.line.fill.background()
                    
                    shadow_border = slide.shapes.add_shape(1, Inches(0.55), Inches(0.55), prs.slide_width - Inches(1.1), prs.slide_height - Inches(1.1))
                    shadow_border.fill.background(); shadow_border.line.color.rgb = SHADOW_COLOR; shadow_border.line.width = Pt(2.5)
                    
                    main_border = slide.shapes.add_shape(1, Inches(0.5), Inches(0.5), prs.slide_width - Inches(1.0), prs.slide_height - Inches(1.0))
                    main_border.fill.background(); main_border.line.color.rgb = ACCENT; main_border.line.width = Pt(2.5)
                    
                    # Sarlavha
                    t_box = slide.shapes.add_textbox(Inches(1.2), Inches(0.6), Inches(11), Inches(1))
                    p = t_box.text_frame.paragraphs[0]
                    p.text = title_text; p.font.size = Pt(34); p.font.bold = True; p.font.name = "Schoolbell"; p.font.color.rgb = BG_DARK
                    
                    # Matnli qism
                    b_box = slide.shapes.add_textbox(Inches(1.2), Inches(2.2), Inches(6.0), Inches(4.5))
                    tf_b = b_box.text_frame; tf_b.word_wrap = True
                    
                    lines = [l.strip() for l in body_text.split("\n") if l.strip()]
                    for idx, line in enumerate(lines):
                        p_b = tf_b.paragraphs[0] if idx == 0 else tf_b.add_paragraph()
                        p_b.text = line.replace("- ", "• ").replace("*", "")
                        p_b.font.size = Pt(18); p_b.font.name = "Comic Sans MS"; p_b.font.color.rgb = TXT_DARK; p_b.space_after = Pt(12)
                    
                    # Rasm yuklash va joylashtirish
                    try:
                        img_url = "https://unsplash.com"
                        search_url = f"https://unsplash.com{img_keyword}&per_page=1&client_id=p_Yv4wVv-XqM-HwG5HqRfZC-CqO-E498_r6VpXh1Y2Y"
                        
                        res_search = requests.get(search_url, timeout=5)
                        if res_search.status_code == 200:
                            res_json = res_search.json()
                            if res_json.get("results"):
                                img_url = res_json["results"][0]["urls"]["small"]
                        
                        img_res = requests.get(img_url, timeout=5)
                        if img_res.status_code == 200:
                            img_bytes = BytesIO(img_res.content)
                            slide.shapes.add_picture(img_bytes, Inches(7.8), Inches(2.2), width=Inches(4.5))
                    except Exception as e:
                        pass

            # Faylni xotiraga yozish
            ppt_io = BytesIO()
            prs.save(ppt_io)
            st.session_state.ppt_data = ppt_io.getvalue()
            st.session_state.file_ready = True
            st.success("🎉 Prezentatsiya muvaffaqiyatli yaratildi!")

        except Exception as e:
            st.error(f"Xato yuz berdi: {str(e)}")

# Faylni yuklab olish vidjeti
if st.session_state.file_ready and st.session_state.ppt_data is not None:
    st.download_button(
        label="Prezentatsiyani yuklab olish (.pptx)", 
        data=st.session_state.ppt_data, 
        file_name=f"{topic.replace(' ', '_')}.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
