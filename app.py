import streamlit as st
import google.generativeai as genai
from gtts import gTTS
from docx import Document
from pptx import Presentation
import requests
import io
import re

# ==========================================
# 1. SAYFA YAPILANDIRMASI & TEMA (UI/UX)
# ==========================================
st.set_page_config(
    page_title="IdeaStorm Robot Core v2.5",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS Stilleri
st.markdown("""
    <style>
    /* Ana Arka Plan ve Tipografi */
    .stApp {
        background-color: #0b0f19;
        font-family: 'Inter', sans-serif;
    }
    
    /* Hero Başlık Tasarımı */
    .hero-container {
        padding: 1.5rem 0rem;
        margin-bottom: 1.5rem;
        border-bottom: 1px solid rgba(255, 255, 255, 0.08);
    }
    .hero-title {
        font-size: 2.8rem;
        font-weight: 800;
        background: linear-gradient(135deg, #6366F1 0%, #A855F7 50%, #EC4899 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        letter-spacing: -0.02em;
    }
    .hero-subtitle {
        font-size: 1.1rem;
        color: #9CA3AF;
        margin-top: 0.4rem;
    }
    
    /* Rozet (Badge) Tasarımları */
    .status-badge {
        display: inline-block;
        padding: 0.25rem 0.75rem;
        border-radius: 9999px;
        font-size: 0.8rem;
        font-weight: 600;
        margin-right: 0.5rem;
    }
    .badge-active { background-color: rgba(16, 185, 129, 0.15); color: #10B981; border: 1px solid rgba(16, 185, 129, 0.3); }
    .badge-byok { background-color: rgba(99, 102, 241, 0.15); color: #818CF8; border: 1px solid rgba(99, 102, 241, 0.3); }
    
    /* Kart Yapıları */
    .feature-card {
        background: #111827;
        border: 1px solid #1F2937;
        border-radius: 12px;
        padding: 1rem;
        margin-top: 1rem;
    }
    </style>
""", unsafe_allow_html=True)

# Başlık Alanı
st.markdown("""
    <div class="hero-container">
        <span class="status-badge badge-active">● Gemini 1.5 Flash Active</span>
        <span class="status-badge badge-byok">⚡ BYOK Architecture</span>
        <h1 class="hero-title">🤖 IdeaStorm Robot Core</h1>
        <p class="hero-subtitle">Otonom Yazılım Mimarisi, Raporlama ve Tek Tık Vercel Dağıtım Ajanı</p>
    </div>
""", unsafe_allow_html=True)

# ==========================================
# 2. GEMINI API VE KONFİGÜRASYON
# ==========================================
GEMINI_KEY = None

# Secrets Kontrolü (Önce Streamlit Secrets'a bakar)
if "GEMINI_API_KEY" in st.secrets:
    GEMINI_KEY = st.secrets["GEMINI_API_KEY"]
elif "gemini_api_key" in st.secrets:
    GEMINI_KEY = st.secrets["gemini_api_key"]

if GEMINI_KEY:
    genai.configure(api_key=GEMINI_KEY)
else:
    st.warning("⚠️ Arka plan Gemini API Anahtarı algılanamadı. Sol menüden kendi API anahtarınızı girerek başlatabilirsiniz.")

# ==========================================
# 3. YAN MENÜ (SIDEBAR) & BYOK AYARLARI
# ==========================================
with st.sidebar:
    st.image("https://img.icons8.com/fluency/96/bot.png", width=64)
    st.title("IdeaStorm Paneli")
    st.caption("v2.5 Pro Edition")
    st.divider()

    # Yedek Gemini Key Girişi
    if not GEMINI_KEY:
        st.subheader("🔑 AI Erişimi")
        user_gemini_key = st.text_input("Gemini API Key", type="password", help="Google AI Studio'dan aldığınız key.")
        if user_gemini_key:
            genai.configure(api_key=user_gemini_key)
            GEMINI_KEY = user_gemini_key
        st.markdown("[🔑 Ücretsiz Gemini Key Al](https://aistudio.google.com/app/apikey)")
        st.divider()

    # Vercel Deployment Entegrasyonu
    st.subheader("🚀 Canlı Yayın (Vercel)")
    vercel_token = st.text_input("Vercel Access Token", type="password", help="Kendi Vercel hesabınızda yayınlamak için girin.")
    st.markdown("[🔗 Vercel Token Oluştur](https://vercel.com/account/tokens)")

    st.divider()
    st.info("💡 **Bilgi:** Üretilen kodlar BYOK modeliyle doğrudan sizin Vercel hesabınızda canlıya alınır. Sunucu ücreti ödemezsiniz.")

# ==========================================
# 4. YARDIMCI MÜHENDİSLİK FONKSİYONLARI
# ==========================================
def extract_html_content(text):
    """Metin içerisindeki HTML kod bloğunu tespit edip temizler."""
    html_pattern = r"```html\s*(.*?)\s*```"
    match = re.search(html_pattern, text, re.DOTALL)
    if match:
        return match.group(1)
    if "<html>" in text.lower() or "<!doctype html>" in text.lower():
        return text
    return None

def create_word_doc(content):
    """Yanıtı profesyonel Word belgesine dönüştürür."""
    doc = Document()
    doc.add_heading('IdeaStorm Robot Core - Teknik Rapor', 0)
    
    # Paragraflara bölerek yazma
    for line in content.split('\n'):
        if line.startswith('# '):
            doc.add_heading(line.replace('# ', ''), level=1)
        elif line.startswith('## '):
            doc.add_heading(line.replace('## ', ''), level=2)
        elif line.strip():
            doc.add_paragraph(line)
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

def create_ppt_doc(content):
    """Yanıtı otomatik slayt sunumuna çevirir."""
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = "IdeaStorm Proje Analizi"
    
    # Ilk 800 karakteri özete koy
    summary_text = content[:800] + "..." if len(content) > 800 else content
    slide.placeholders[1].text = summary_text
    
    bio = io.BytesIO()
    prs.save(bio)
    return bio.getvalue()

def text_to_audio(text):
    """Yanıtın sesli özetini (MP3) üretir."""
    clean_text = re.sub(r'<[^>]+>', '', text)  # HTML taglerini temizle
    clean_text = clean_text.replace('*', '').replace('`', '')
    tts = gTTS(text=clean_text[:350], lang='tr', slow=False)
    bio = io.BytesIO()
    tts.write_to_fp(bio)
    return bio.getvalue()

def deploy_to_vercel(project_name, html_code, token):
    """Vercel REST API v13 üzerinden tek tıkla canlı yayın yapar."""
    url = "https://api.vercel.com/v13/deployments"
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    payload = {
        "name": project_name,
        "files": [
            {
                "file": "index.html",
                "data": html_code
            }
        ],
        "projectSettings": {
            "framework": None
        }
    }
    response = requests.post(url, headers=headers, json=payload)
    return response.json()

# ==========================================
# 5. SOHBET VE AI AKIŞI
# ==========================================
if "messages" not in st.session_state:
    st.session_state.messages = []

# Geçmiş Mesajları Ekrana Basma
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# Kullanıcı Komut Girişi
if prompt := st.chat_input("Bir web projesi fikri ver veya kod yazmasını iste..."):
    if not GEMINI_KEY:
        st.error("Lütfen önce bir Gemini API Key sağlayın.")
        st.stop()

    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        try:
            # Gemini 1.5 Flash Model Yapılandırması
            system_instruction = (
                "Sen IdeaStorm Robot Core yapay zeka mimarısın. "
                "Görevin kullanıcılara eksiksiz, modern, modern CSS stillerine sahip web projeleri, "
                "temiz yazılım mimarileri ve teknik belgeler üretmektir. "
                "Eğer kullanıcı bir web sitesi veya arayüz isterse, koda başlarken tek bir tam HTML dosyası "
                "içinde CSS ve JavaScript dahil olacak şekilde ```html ... ``` blokları arasında yanıt ver."
            )
            
            model = genai.GenerativeModel(
                model_name='gemini-3.6-flash',
                system_instruction=system_instruction,
                generation_config={
                    "temperature": 0.7,
                    "top_p": 0.95,
                    "max_output_tokens": 8192
                }
            )

            with st.spinner("🤖 IdeaStorm projeyi analiz ediyor ve üretiyor..."):
                response = model.generate_content(prompt)
                res_text = response.text
                
                st.markdown(res_text)

                # --- ÇIKTI MODÜLLERİ ---
                st.divider()
                
                # 1. Sesli Özet
                with st.expander("🔊 Sesli Özeti Dinle", expanded=False):
                    audio_data = text_to_audio(res_text)
                    st.audio(audio_data, format='audio/mp3')

                # 2. Doküman İndirme Alanı
                col1, col2 = st.columns(2)
                with col1:
                    st.download_button(
                        label="📄 Word Raporu (.docx)",
                        data=create_word_doc(res_text),
                        file_name="IdeaStorm_Rapor.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                with col2:
                    st.download_button(
                        label="📊 PPT Sunumu (.pptx)",
                        data=create_ppt_doc(res_text),
                        file_name="IdeaStorm_Sunum.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )

                # 3. Vercel Otomatik Yayın Alanı
                html_code = extract_html_content(res_text)
                if html_code:
                    st.markdown("---")
                    st.subheader("🚀 Otonom Vercel Dağıtımı")
                    
                    if vercel_token:
                        if st.button("🌐 Projeyi Canlıya Al (Vercel Deploy)", type="primary", use_container_width=True):
                            with st.spinner("Vercel API ile iletişim kuruluyor ve sunucusuz dağıtım yapılıyor..."):
                                deploy_res = deploy_to_vercel("ideastorm-web-app", html_code, vercel_token)
                                
                                if "url" in deploy_res:
                                    site_url = f"https://{deploy_res['url']}"
                                    st.balloons()
                                    st.success(f"🎉 **Proje Canlıda!** [Projenizi Görüntüleyin]({site_url})")
                                    st.code(site_url, language="text")
                                else:
                                    st.error("Vercel dağıtım hatası! Lütfen Vercel Token'ınızı kontrol edin.")
                    else:
                        st.info("💡 Üretilen bu web projesini canlıya almak için sol menüye Vercel Token ekleyebilirsiniz.")

            st.session_state.messages.append({"role": "assistant", "content": res_text})

        except Exception as e:
            st.error(f"Sistem Hatası: {e}")
