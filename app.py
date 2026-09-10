import streamlit as st
import google.generativeai as genai
from gtts import gTTS
from docx import Document
from pptx import Presentation
import requests
import io

# 1. Sayfa Yapılandırması ve Başlık
st.set_page_config(
    page_title="IdeaStorm Robot Core",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. Özel Tasarım ve CSS Stilleri
st.markdown("""
    <style>
    .main {
        background-color: #0e1117;
    }
    .stApp header {
        background-color: transparent;
    }
    .hero-title {
        font-size: 2.6rem;
        font-weight: 800;
        background: linear-gradient(90deg, #4F46E5, #06B6D4);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.2rem;
    }
    .hero-subtitle {
        font-size: 1.1rem;
        color: #9CA3AF;
        margin-bottom: 2rem;
    }
    .feature-card {
        background-color: #1E293B;
        border: 1px solid #334155;
        border-radius: 12px;
        padding: 1.2rem;
        margin-bottom: 1rem;
    }
    </style>
""", unsafe_allow_allow_html=True)

# Başlık Alanı
st.markdown('<div class="hero-title">🤖 IdeaStorm Robot Core</div>', unsafe_allow_html=True)
st.markdown('<div class="hero-subtitle">Otonom Yapay Zeka Ajanı • Kodlama, Dokümantasyon ve Tek Tık Canlı Yayın</div>', unsafe_allow_html=True)

# 3. Gemini API Bağlantısı (Secrets Üzerinden Arka Planda)
if "GEMINI_API_KEY" in st.secrets:
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
else:
    st.warning("⚠️ Sistem API Key bulunamadı. Lütfen Streamlit Secrets alanına GEMINI_API_KEY tanımlayın.")

# 4. Sol Menü Tasarımı (Kullanıcı Dostu Vercel Bağlantısı)
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/bot.png", width=70)
    st.header("🚀 Canlı Yayın Modülü")
    st.write("Ürettiğiniz web projelerini anında kendi Vercel hesabınızda yayınlayın.")
    
    st.markdown("[🔗 Ücretsiz Vercel Token Al](https://vercel.com/account/tokens)")
    vercel_token = st.sidebar.text_input("Vercel Token (Opsiyonel)", type="password", help="Projeleri canlıya almak için kendi Vercel token'ınızı girin.")

    st.divider()
    st.caption("IdeaStorm v2.0 • Sürdürülebilir BYOK Altyapısı")

# 5. Yardımcı Fonksiyonlar (Doküman, Ses, Vercel API)
def create_word_doc(text):
    doc = Document()
    doc.add_heading('IdeaStorm - Proje Çıktısı', 0)
    doc.add_paragraph(text)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def create_ppt_doc(text):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "IdeaStorm Proje Özeti"
    slide.placeholders[1].text = text[:600]
    bio = io.BytesIO()
    prs.save(bio)
    return bio.getvalue()

def text_to_audio(text):
    tts = gTTS(text=text[:300], lang='tr')
    bio = io.BytesIO()
    tts.write_to_fp(bio)
    return bio.getvalue()

def deploy_to_vercel(project_name, html_content, token):
    url = "https://api.vercel.com/v13/deployments"
    headers = {"Authorization": f"Bearer {token}"}
    payload = {
        "name": project_name,
        "files": [{"file": "index.html", "data": html_content}],
        "projectSettings": {"framework": None}
    }
    res = requests.post(url, headers=headers, json=payload)
    return res.json()

# 6. Sohbet Geçmişi ve Arayüz Akışı
if "messages" not in st.session_state:
    st.session_state.messages = []

for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

if prompt := st.chat_input("IdeaStorm'a bir proje fikri ver veya web sitesi oluşturmasını iste..."):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        try:
            # Sistem Talimatı (System Prompt)
            system_instruction = (
                "Sen IdeaStorm Robot Core mimarisinin otonom yapay zeka ajanısın. "
                "Kullanıcılara modern web projeleri, temiz kodlar, detaylı teknik raporlar üreten "
                "üst düzey bir geliştirici ve analistsin. Yanıtların net, estetik ve işlevsel olmalıdır."
            )
            model = genai.GenerativeModel(
                'gemini-1.5-flash',
                system_instruction=system_instruction
            )
            
            with st.spinner("IdeaStorm düşünüyor ve üretiyor..."):
                response = model.generate_content(prompt)
                res_text = response.text
                st.markdown(res_text)
                
                # Sesli Yanıt Paneli
                st.divider()
                st.subheader("🔊 Sesli Özet")
                audio_bytes = text_to_audio(res_text)
                st.audio(audio_bytes, format='audio/mp3')

                # Doküman İndirme Butonları
                st.subheader("📥 Doküman Çıktıları")
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="📄 Word Raporu İndir (.docx)",
                        data=create_word_doc(res_text),
                        file_name="IdeaStorm_Rapor.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                with col2:
                    st.download_button(
                        label="📊 PPT Sunumu İndir (.pptx)",
                        data=create_ppt_doc(res_text),
                        file_name="IdeaStorm_Sunum.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )

                # Vercel Otomatik Dağıtım Mekanizması
                if "<html>" in res_text.lower() or "<!doctype html>" in res_text.lower():
                    st.divider()
                    st.subheader("🚀 Vercel Canlı Yayın")
                    if vercel_token:
                        if st.button("🌐 Web Sitesini Kendi Vercel Hesabında Yayınla", type="primary", use_container_width=True):
                            with st.spinner("Proje paketleniyor ve Vercel'e aktarılıyor..."):
                                deploy_res = deploy_to_vercel("ideastorm-generated-app", res_text, vercel_token)
                                if "url" in deploy_res:
                                    st.balloons()
                                    st.success(f"🎉 Projeniz Canlıda! Link: https://{deploy_res['url']}")
                                else:
                                    st.error("Vercel bağlantı hatası. Lütfen Vercel Token'ınızı kontrol edin.")
                    else:
                        st.info("💡 Bu web kodunu tek tıkla canlıya almak için sol menüden Vercel Token'ınızı girebilirsiniz.")

            st.session_state.messages.append({"role": "assistant", "content": res_text})
        except Exception as e:
            st.error(f"Bir hata oluştu: {e}")
