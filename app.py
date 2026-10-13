import streamlit as st
import json
import requests
from io import BytesIO
from pptx import Presentation

# ==========================================
# 1. SAHIFA SOZLAMALARI VA SEO / QIDIRUV TIZIMLARI
# ==========================================
st.set_page_config(
    page_title="NeoBiology AI - Biologiya va Mikrobiologiya uchun sun'iy intellekt",
    page_icon="🧬",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Qidiruv tizimlari (Google, Yandex) uchun meta teglar va logotipni chiqarish
st.markdown("""
    <head>
        <meta name="description" content="NeoBiology AI - Biologiya fanlari, prokariot hujayralar va mikrobiologiya bo'yicha sun'iy intellekt yordamida taqdimot va hujjatlar yaratish platformasi.">
        <meta name="keywords" content="NeoBiology, Biologiya AI, Slayd yaratish, PPT, PDF, sun'iy intellekt">
        <meta name="author" content="NeoBiology Team">
    </head>
""", unsafe_allow_html=True)

# ==========================================
# 2. LOGOTIP VA SARLAVHA
# ==========================================
col1, col2 = st.columns([1, 6])
with col1:
    # Desktop papkangizdagi logo.png faylini chaqiramiz
    try:
        st.image("logo.png", width=100)
    except:
        st.write("🧬")

with col2:
    st.title("NeoBiology AI")
    st.caption("Canva, Gamma va Genspark kabi ilg‘or sun'iy intellekt imkoniyatlariga ega platforma")

# ==========================================
# API KALIT VA GEMINI REST API SO'ROVI (AQ... kalitlar uchun moslashgan)
# ==========================================
# O'zingizning AQ... bilan boshlanuvchi kalitingizni shu yerga yozing
GEMINI_API_KEY = "AQ.Ab8RN6KdRyvoAOGvkcIZ_4MNaFIwm21JeJunpZwXpW7cjNP7WQ"

def generate_slides_with_aq_key(topic):
    url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent"
    
    headers = {
        "x-goog-api-key": GEMINI_API_KEY.strip(),
        "Content-Type": "application/json"
    }
    
    prompt = f"""
    '{topic}' mavzusida 10 ta slayd uchun ma'lumot tayyorla.
    Faqatgina toza JSON formatida javob ber, boshqa matn yozma.
    Format quyidagicha bo'lsin:
    {{
      "slides":[
        {{
          "title":"Sarlavha",
          "content":"Matn qismi..."
        }}
      ]
    }}
    """
    
    data = {
        "contents": [{
            "parts": [{"text": prompt}]
        }]
    }
    
    response = requests.post(url, headers=headers, json=data)
    
    if response.status_code != 200:
        raise Exception(f"API xatosi ({response.status_code}): {response.text}")
        
    res_json = response.json()
    text = res_json["candidates"][0]["content"]["parts"][0]["text"].strip()
    
    # Markdown belgilarini tozalash
    text = text.replace("```json", "").replace("```", "").strip()
    return json.loads(text)

# ==========================================
# 3. PPT FAYL YARATISH FUNKSIYASI
# ==========================================
def create_ppt(data):
    prs = Presentation()
    for slide_data in data["slides"]:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        slide.shapes.title.text = slide_data["title"]
        slide.placeholders[1].text = slide_data["content"]

    buffer = BytesIO()
    prs.save(buffer)
    return buffer.getvalue()

# ==========================================
# INTERFEYS VA BOSHQARUV
# ==========================================
topic = st.text_input(
    "Mavzu kiriting (Masalan: Prokariot hujayra va mikrobiologiya):",
    "Prokariot hujayra va mikrobiologiya"
)

if st.button("🚀 Taqdimot (PPT) yaratish"):
    if not GEMINI_API_KEY or GEMINI_API_KEY == "AQ...":
        st.warning("Iltimos, koddagi GEMINI_API_KEY o'zgaruvchisiga o'z kalitingizni kiriting!")
    else:
        try:
            with st.spinner("NeoBiology AI ma'lumotlarni tahlil qilmoqda va slaydlarni shakllantirmoqda..."):
                data = generate_slides_with_aq_key(topic)
                ppt_file = create_ppt(data)

            st.success("✅ Taqdimot muvaffaqiyatli tayyorlandi!")

            st.download_button(
                label="📥 Yuqori sifatli PPT faylni yuklab olish",
                data=ppt_file,
                file_name="NeoBiology_Presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

        except Exception as e:
            st.error(f"Xatolik yuz berdi: {str(e)}")
