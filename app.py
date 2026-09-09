import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime

# ==========================================
# 1. SAYFA VE MODERN SAAS TEMA KONFİGÜRASYONU
# ==========================================
st.set_page_config(
    page_title="IdeaStorm AI (IGS) v4.0 | DROX TEAM",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Dark SaaS Custom CSS
st.markdown("""
<style>
    /* Ana Arka Plan */
    .stApp {
        background-color: #0b0f19;
        color: #f1f5f9;
    }
    
    /* Yan Menü Tasarımı */
    section[data-testid="stSidebar"] {
        background-color: #111827;
        border-right: 1px solid #1f2937;
    }
    
    /* Özel Başlık Kartı */
    .hero-container {
        background: linear-gradient(135deg, #1e1b4b 0%, #31104b 50%, #0f172a 100%);
        border: 1px solid #4338ca;
        border-radius: 16px;
        padding: 2rem;
        margin-bottom: 2rem;
        box-shadow: 0 10px 25px -5px rgba(99, 102, 241, 0.2);
    }
    
    /* Kart Tasarımları */
    .saas-card {
        background-color: #1f2937;
        border: 1px solid #374151;
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1rem;
        transition: transform 0.2s ease, border-color 0.2s ease;
    }
    .saas-card:hover {
        border-color: #6366f1;
    }
    
    /* Buton Tasarımları */
    .stButton>button {
        background: linear-gradient(135deg, #6366f1 0%, #a855f7 100%);
        color: #ffffff;
        border: none;
        border-radius: 8px;
        padding: 0.6rem 1.4rem;
        font-weight: 600;
        box-shadow: 0 4px 12px rgba(99, 102, 241, 0.3);
        transition: all 0.3s ease;
        width: 100%;
    }
    .stButton>button:hover {
        opacity: 0.95;
        box-shadow: 0 6px 20px rgba(168, 85, 247, 0.4);
    }
    
    /* Rozetler (Badges) */
    .badge {
        background: #312e81;
        color: #a5b4fc;
        padding: 4px 10px;
        border-radius: 20px;
        font-size: 0.8rem;
        font-weight: 600;
        display: inline-block;
        margin-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. GEMINI API VE TEK MODEL YAPILANDIRMASI
# ==========================================
api_key = st.secrets.get("GEMINI_API_KEY", None)

if api_key:
    genai.configure(api_key=api_key)
    # Çakışmaları önlemek için tek ve kararlı model tanımı
    model = genai.GenerativeModel('gemini-3.6-flash')
else:
    model = None

# Oturum Durumları (Session State)
if "chat_sessions" not in st.session_state:
    st.session_state.chat_sessions = {"Genel Sohbet": []}
if "current_session" not in st.session_state:
    st.session_state.current_session = "Genel Sohbet"
if "editor_code" not in st.session_state:
    st.session_state.editor_code = """<!DOCTYPE html>
<html>
<head>
    <style>
        body { background: #0f172a; color: white; font-family: sans-serif; display: flex; justify-content: center; align-items: center; height: 100vh; margin: 0; }
        .card { background: #1e293b; padding: 2rem; border-radius: 12px; border: 1px solid #334155; text-align: center; }
        button { background: #6366f1; color: white; border: none; padding: 10px 20px; border-radius: 6px; cursor: pointer; }
    </style>
</head>
<body>
    <div class="card">
        <h2>🚀 IdeaStorm Live Canvas</h2>
        <p>Sol taraftaki editörden kodu değiştirin veya AI'ya yeni bir uygulama yaptırın!</p>
        <button onclick="alert('IdeaStorm AI Çalışıyor!')">Tıkla</button>
    </div>
</body>
</html>"""

# ==========================================
# 3. PROFESYONEL WORD (.DOCX) DÖNÜŞTÜRÜCÜ
# ==========================================
def build_docx_file(title, raw_text):
    doc = Document()
    
    # Başlık
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_title = p_title.add_run(f"{title}\n")
    run_title.font.size = Pt(22)
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(99, 102, 241)
    
    # Alt Bilgi
    p_meta = doc.add_paragraph()
    p_meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_meta = p_meta.add_run(f"DROX TEAM | IdeaStorm AI (IGS) - {datetime.date.today()}\n\n")
    run_meta.font.size = Pt(10)
    run_meta.font.italic = True
    run_meta.font.color.rgb = RGBColor(100, 116, 139)

    # İçerik Ayrıştırma
    lines = raw_text.split("\n")
    for line in lines:
        clean_line = line.strip()
        if not clean_line:
            continue
            
        if clean_line.startswith("### "):
            h = doc.add_heading(clean_line.replace("### ", ""), level=3)
            h.style.font.color.rgb = RGBColor(168, 85, 247)
        elif clean_line.startswith("## "):
            h = doc.add_heading(clean_line.replace("## ", ""), level=2)
            h.style.font.color.rgb = RGBColor(99, 102, 241)
        elif clean_line.startswith("# "):
            h = doc.add_heading(clean_line.replace("# ", ""), level=1)
            h.style.font.color.rgb = RGBColor(15, 23, 42)
        elif clean_line.startswith("* ") or clean_line.startswith("- "):
            p = doc.add_paragraph(style='List Bullet')
            p.add_run(clean_line[2:].replace("**", ""))
        else:
            p = doc.add_paragraph()
            parts = clean_line.split("**")
            for i, part in enumerate(parts):
                run = p.add_run(part)
                if i % 2 == 1:
                    run.font.bold = True

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer.getvalue()

# ==========================================
# 4. YAN MENÜ (SIDEBAR)
# ==========================================
with st.sidebar:
    st.markdown("### ⚡ IdeaStorm AI")
    st.caption("DROX TEAM | Enterprise Edition v4.0")
    st.divider()

    st.markdown("#### ⚙️ Çalışma Modu")
    ai_mode = st.selectbox(
        "AI Rolü:",
        [
            "🤖 Otonom Ürün Mimarı", 
            "💻 Full-Stack Web Geliştirici", 
            "🏆 TEKNOFEST & TÜBİTAK Danışmanı",
            "📚 Eğitim ve Öğrenme Koçu"
        ]
    )
    
    st.divider()

    st.markdown("#### 📌 Modüller")
    app_page = st.radio(
        "Navigasyon:",
        [
            "💬 Akıllı Asistan & Chat", 
            "💻 Low-Code Web Stüdyosu", 
            "📊 Otonom Pazar Analizi",
            "🏆 Yarışma & Rapor Motoru"
        ]
    )
    
    st.divider()
    if not api_key:
        st.error("⚠️ API Key Eksik!\n.streamlit/secrets.toml dosyasına GEMINI_API_KEY ekleyin.")

# ==========================================
# MODÜL 1: AKILLI ASİSTAN & CHAT
# ==========================================
if app_page == "💬 Akıllı Asistan & Chat":
    st.markdown(f"""
    <div class="hero-container">
        <span class="badge">{ai_mode}</span>
        <h1 style="margin: 0; font-size: 2.2rem;">IdeaStorm Otonom Asistan</h1>
        <p style="color: #94a3b8; margin-top: 0.5rem;">Fikirlerinizi tartışın, mimari tasarlayın ve anında indirin.</p>
    </div>
    """, unsafe_allow_html=True)

    col_s1, col_s2 = st.columns([3, 1])
    with col_s1:
        selected_session = st.selectbox("Sohbet Odası:", list(st.session_state.chat_sessions.keys()))
        st.session_state.current_session = selected_session
    with col_s2:
        new_room = st.text_input("Yeni Oda", placeholder="Oda Adı")
        if st.button("➕ Aç") and new_room:
            if new_room not in st.session_state.chat_sessions:
                st.session_state.chat_sessions[new_room] = []
                st.session_state.current_session = new_room
                st.rerun()

    current_messages = st.session_state.chat_sessions[st.session_state.current_session]

    # Mesaj Geçmişi
    for idx, msg in enumerate(current_messages):
        with st.chat_message(msg["role"]):
            st.write(msg["content"])
            if msg["role"] == "assistant":
                # Kırmızı çizgi ve tırnak hatasını önlemek için adımlar değişkenlere atandı
                step_num = idx + 1
                doc_title = f"Yanit_{step_num}"
                docx_b = build_docx_file(doc_title, msg["content"])
                btn_key = f"dl_btn_{st.session_state.current_session}_{step_num}"
                
                st.download_button(
                    label="📄 Word (.docx) İndir",
                    data=docx_b,
                    file_name=f"IdeaStorm_{doc_title}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=btn_key
                )

    # Girdi
    if prompt := st.chat_input("Bir fikir sorun veya görevi tanımlayın..."):
        current_messages.append({"role": "user", "content": prompt})
        st.rerun()

    if current_messages and current_messages[-1]["role"] == "user":
        last_prompt = current_messages[-1]["content"]
        with st.chat_message("assistant"):
            if model:
                with st.spinner("IdeaStorm AI Yanıtlıyor..."):
                    try:
                        sys_prompt = f"Sen {ai_mode} olarak görev yapıyorsun. Düzenli, başlıklar içeren profesyonel yanıtlar ver.\n\nKullanıcı: "
                        response = model.generate_content(sys_prompt + last_prompt)
                        
                        if response and response.text:
                            current_messages.append({"role": "assistant", "content": response.text})
                            st.rerun()
                        else:
                            st.error("Modelden yanıt alınamadı.")
                    except Exception as e:
                        st.error(f"Hata Oluştu: {e}")
            else:
                st.warning("Lütfen API anahtarınızı tanımlayın.")

# ==========================================
# MODÜL 2: LOW-CODE WEB STÜDYOSU (CANLI EDITÖR)
# ==========================================
elif app_page == "💻 Low-Code Web Stüdyosu":
    st.markdown("""
    <div class="hero-container">
        <span class="badge">Live Web Studio</span>
        <h1 style="margin: 0; font-size: 2.2rem;">Low-Code Prototip Stüdyosu</h1>
        <p style="color: #94a3b8; margin-top: 0.5rem;">İstediğiniz uygulamayı anlatın, canlı çalışan koda dönüştürün.</p>
    </div>
    """, unsafe_allow_html=True)

    col_prompt, col_btn = st.columns([4, 1])
    with col_prompt:
        app_desc = st.text_input("Nasıl bir web uygulaması istersiniz?", value="Karanlık temalı, sayaçlı ve butonlu modern bir Pomodoro Zamanlayıcı")
    with col_btn:
        st.write("")
        st.write("")
        generate_code = st.button("🚀 Üret")

    if generate_code and model:
        with st.spinner("Uygulama kodu yazılıyor..."):
            try:
                p = f"SADECE çalışan, modern görünümlü, CSS ve JS içeren TEK BİR HTML kodu yaz. Başka hiçbir açıklama metni ekleme. İstek: {app_desc}"
                response = model.generate_content(p)
                if response and response.text:
                    raw_html = response.text
                    clean_html = raw_html.replace("```html", "").replace("```", "").strip()
                    st.session_state.editor_code = clean_html
                    st.success("Uygulama başarıyla oluşturuldu!")
                else:
                    st.error("Kod üretilemedi.")
            except Exception as e:
                st.error(f"Kod Üretim Hatası: {e}")

    st.divider()

    # İki Panelli Düzenleyici & Önizleme
    col_code, col_preview = st.columns([1, 1])

    with col_code:
        st.subheader("📝 Kod Düzenleyici")
        updated_code = st.text_area("HTML / CSS / JS", value=st.session_state.editor_code, height=450)
        st.session_state.editor_code = updated_code
        
        st.download_button(
            label="💾 .HTML Dosyası Olarak İndir",
            data=st.session_state.editor_code,
            file_name="ideastorm_app.html",
            mime="text/html"
        )

    with col_preview:
        st.subheader("🖥️ Canlı Önizleme")
        st.components.v1.html(st.session_state.editor_code, height=450, scrolling=True)

# ==========================================
# MODÜL 3: OTONOM PAZAR ANALİZİ
# ==========================================
elif app_page == "📊 Otonom Pazar Analizi":
    st.markdown("""
    <div class="hero-container">
        <span class="badge">Market Intelligence</span>
        <h1 style="margin: 0; font-size: 2.2rem;">Girişim & Pazar Analiz Matrisi</h1>
        <p style="color: #94a3b8; margin-top: 0.5rem;">Proje fikrinizi girin; SWOT, rakip analizi ve gelir modelini saniyeler içinde alın.</p>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns(2)
    with col1:
        p_name = st.text_input("Proje / Girişim Adı:", value="IdeaStorm AI")
    with col2:
        p_target = st.text_input("Hedef Kitle:", value="Yazılımcılar, Öğrenciler, Girişimciler")

    p_concept = st.text_area("Projenizin Detaylı Açıklaması:", value="Yapay zeka desteğiyle fikirleri otonom web uygulamalarına ve resmi yarışma raporlarına dönüştüren platform.")

    if st.button("📊 Kapsamlı Analizi Başlat") and model:
        with st.spinner("Pazar verileri ve mimari analiz ediliyor..."):
            try:
                prompt_m = f"""
                Proje Adı: {p_name}
                Hedef Kitle: {p_target}
                Açıklama: {p_concept}

                Lütfen şu başlıklarda detaylı bir Girişim Analiz Raporu oluştur:
                ## 1. SWOT Analizi (Güçlü, Zayıf Yönler, Fırsatlar, Tehditler)
                ## 2. Hedef Pazar ve Rakip Analizi
                ## 3. Gelir Modeli ve Büyüme Stratejisi
                ## 4. Önerilen Teknoloji Yığını (Tech Stack)
                """
                response = model.generate_content(prompt_m)
                if response and response.text:
                    res = response.text
                    st.markdown(res)
                    
                    doc_b = build_docx_file(f"{p_name} Pazar Analiz Raporu", res)
                    st.download_button(
                        label="📥 Analiz Raporunu Word (.docx) Olarak İndir",
                        data=doc_b,
                        file_name=f"{p_name}_Pazar_Analizi.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"Analiz Hatası: {e}")

# ==========================================
# MODÜL 4: YARIŞMA & RAPOR MOTORU
# ==========================================
elif app_page == "🏆 Yarışma & Rapor Motoru":
    st.markdown("""
    <div class="hero-container">
        <span class="badge">Official Report Engine</span>
        <h1 style="margin: 0; font-size: 2.2rem;">TEKNOFEST & TÜBİTAK Raporlayıcı</h1>
        <p style="color: #94a3b8; margin-top: 0.5rem;">Resmi başvuru standartlarında teknik raporlar ve proje metinleri hazırlayın.</p>
    </div>
    """, unsafe_allow_html=True)

    c1, c2 = st.columns(2)
    with c1:
        t_name = st.text_input("Proje Adı", value="IdeaStorm AI")
    with c2:
        t_type = st.selectbox("Yarışma / Program", ["TEKNOFEST İnsanlık Yararına Teknoloji", "TÜBİTAK 2204 Ortaokul/Lise Projeleri", "Kuluçka / İTÜ Çekirdek Başvurusu"])

    t_problem = st.text_area("Çözülen Problem ve Hedef:", value="Proje geliştirme süreçlerinde fikir aşamasından prototip ve raporlama aşamasına geçişteki zaman kaybı ve rehberlik eksikliği.")

    if st.button("📋 Resmi Raporu Üret") and model:
        with st.spinner("Resmi başvuru formatında rapor hazırlanıyor..."):
            try:
                prompt_r = f"""
                Program: {t_type}
                Proje Adı: {t_name}
                Problem: {t_problem}

                Resmi başvuru standartlarına uygun olarak şu bölümleri içeren detaylı bir proje raporu hazırla:
                ## 1. Proje Özeti ve Amacı
                ## 2. Çözüm Ürettiği Problem ve Yenilikçi (Özgün) Yönü
                ## 3. Kullanılacak Yöntem ve Teknik Mimarisi
                ## 4. Yaygın Etki ve Hedef Kitle
                ## 5. Uygulama Takvimi ve Bütçe Planı
                """
                response = model.generate_content(prompt_r)
                if response and response.text:
                    res = response.text
                    st.markdown(res)
                    
                    doc_b = build_docx_file(f"{t_name} {t_type} Raporu", res)
                    st.download_button(
                        label="📥 Başvuru Raporunu Word (.docx) Olarak İndir",
                        data=doc_b,
                        file_name=f"{t_name}_Resmi_Rapor.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            except Exception as e:
                st.error(f"Rapor Üretim Hatası: {e}")