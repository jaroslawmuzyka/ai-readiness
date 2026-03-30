import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from openpyxl.styles import PatternFill, Font
from openpyxl.utils import get_column_letter
import io
import os
import requests
import cairosvg
import traceback
from PIL import Image

# 1. Page Configuration
st.set_page_config(
    page_title="AI Readiness Tool",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for Premium Look
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        background-color: #4F81BD;
        color: white;
        font-weight: bold;
    }
    .stDownloadButton>button {
        width: 100%;
        border-radius: 5px;
        height: 3em;
        background-color: #2ecc71;
        color: white;
        font-weight: bold;
    }
    div[data-testid="stExpander"] {
        border: 1px solid #e9ecef;
        border-radius: 10px;
        background-color: white;
    }
    h1, h2, h3 {
        color: #2c3e50;
    }
    </style>
    """, unsafe_allow_html=True)

# --- LOGIN MODULE ---
def check_password():
    """Returns `True` if the user had the correct password."""
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == st.secrets.get("APP_PASSWORD", "admin123"): # Default for local dev
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # don't store password
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # First run, show input for password.
        st.text_input(
            "Podaj hasło dostępu:", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.info("Dane potrzebne do zalogowania znajdują się w Monday. Kontakt: jaroslaw.muzyka@performance-group.pl")
        return False
    elif not st.session_state["password_correct"]:
        # Password not correct, show input + error.
        st.text_input(
            "Podaj hasło dostępu:", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.error("😕 Niepoprawne hasło")
        st.info("Dane potrzebne do zalogowania znajdują się w Monday. Kontakt: jaroslaw.muzyka@performance-group.pl")
        return False
    else:
        # Password correct.
        return True

if not check_password():
    st.stop()

# --- DATABASE / KNOWLEDGE ---
commentary_db = {
    "Czy dodany jest plik robots.txt?": "Plik robots.txt to podstawowy mechanizm kontroli dostępu dla botów. Jego brak może prowadzić do crawlowania niepożądanych zasobów (co może wpłynąć na gorszą indeksację kluczowych podstron), a jego błędna konfiguracja może całkowicie zablokować dostęp dla botów np dla GPTBot.",
    "Czy strona jest możliwa do crawlowania przez wyszukiwarki?": "Modele AI opierają się na w dużej mierze na danych zindeksowanych przez wyszukiwarki. Jeśli strona jest niedostępna dla robotów, jej treść może być niewidoczna dla LLMów i nie zostanie wykorzystana w odpowiedziach.", 
    "Czy strona jest indeksowalna?": "Brak możliwości zindeksowania strony (np. przez tag 'noindex') może wpływać na wykluczenie jej treści z bazy wiedzy modeli AI.", 
    "Czy dodana jest mapa strony XML?": "Mapa strony to przewodnik dla robotów, ułatwiający im szybkie odnalezienie wszystkich kluczowych podstron, które powinny zostać przeanalizowane.", 
    "Czy dodawany jest znacznik <lastmod> w mapie witryny XML?": "Znacznik <lastmod> informuje o dacie ostatniej modyfikacji. Modele językowe preferują aktualne informacje, a ten znacznik może pomóc im priorytetyzować nowe treści.", 
    "Czy dodane są dedykowane mapy XML ze zdjęciami i filmami?": "Odpowiedzi AI coraz częściej zawierają multimedia. Dedykowane mapy mogą pomóc modelom AI odkrywać i poprawnie interpretować te zasoby.", 
    "Czy mapa strony XML została przesłana do Google Search Console?": "Przesłanie mapy do GSC przyspiesza proces odkrywania i indeksowania treści przez boty Google, w tym te zasilające AI.", 
    "Czy mapa strony XML została przesłana do Bing Webmaster Tools?": "Jest to kluczowe dla widoczności w ekosystemie Microsoft, w tym w wyszukiwarce Bing i Copilot.", 
    "Czy mapa strony XML jest dodana do robots.txt?": "Umieszczenie ścieżki do mapy strony w pliku robots.txt to standardowa praktyka, która ułatwia botom jej odnalezienie.", 
    "Czy mapa strony XML zawiera tylko adresy z kodem 200, kanoniczne, nie zawiera adresów noindex?": "Czysta, wolna od błędów i niepożądanych adresów mapa strony pozwala botom AI efektywniej wykorzystać czas na analizę wartościowych treści, zamiast marnować go na niepożądane przez nas adresy.", 
    "Czy plik robots.txt pozwala agentom LLM i robotom wyszukiwarek na crawlowanie strony?": "Zablokowanie botów np GPTBot może całkowicie uniemożliwiść im wykorzystanie treści na naszej stronie.", 
    "Czy dodany jest plik llm.txt?": "Obecność pliku llm.txt to nowy sygnał dla systemów AI, który może wspierać wykorzystywanie treści w LLM.", 
    "Czy plik llm.txt jest poprawnie sformatowany i zawiera najważniejsze podstrony?": "Plik llm.txt może precyzyjnie kontrolować, które dane mogą być używane do trenowania modeli AI.", 
    "Czy dodany jest certyfikat SSL?": "Protokół HTTPS (SSL) jest fundamentalnym sygnałem zaufania. Strony bez certyfikatu są uznawane za niezabezpieczone co może mieć wpływ na obecność w odpowiedziach AI.", 
    "Czy strona wykorzystuje Breadcrumb?": "Breadcrumby (nawigacja okruszkowa) poza lepszym doświadczeniem użytkowników na stronie może pomagać LLM zrozumieć hierarchię i relacje między podstronami, co jest kluczowe dla generowania trafnych odpowiedzi.",
    "social_media_general": "Profile społecznościowe budują tożsamość i autorytet marki. Modele AI mogą wykorzystywać je do potwierdzenia, że marka jest prawdziwa, aktywna i jest ekspertem w swojej dziedzinie, co zwiększa jej wiarygodność. Regularna aktywność to sygnał, że informacje o marce są aktualne.",
    "linkbuilding_general": "Linki przychodzące z innych stron to jeden z najważniejszych sygnałów autorytetu w Internecie. Modele AI, podobnie jak tradycyjne wyszukiwarki, preferują informacje pochodzące z wiarygodnych, dobrze podlinkowanych źródeł.",
    "tresci_general": "Treść jest fundamentem, na którym opierają się modele AI. Musi być ona nie tylko unikalna i merytoryczna, ale także aktualna, zaufana i strukturyzowana w sposób, który ułatwia maszynom szybkie znalezienie konkretnych odpowiedzi. Elementy takie jak daty, linki do źródeł czy sekcje FAQ bezpośrednio wpływają na to, czy treść zostanie uznana za wiarygodne źródło dla generowanych odpowiedzi.",
    "non_indexable": "Modele językowe mogą nie wykorzystywać treści ze stron, które są zablokowane przed indeksowaniem (np. przez tag 'noindex'). Informacje zawarte na tych podstronach mogą być dla AI niewidoczne.", 
    "core_web_vitals": "Google traktuje Core Web Vitals jako kluczowe wskaźniki jakości i użyteczności strony. Słabe wyniki CWV mogą obniżyć postrzeganą jakość witryny i zmniejszyć jej szansę na bycie cytowaną.", 
    "gsc_indexation": "Status w Google Search Console API to potwierdzenie, czy Google zna i indeksuje daną stronę. Jeśli strona nie jest zaindeksowana, może być poza zasięgiem mechanizmów AI od Google.", 
    "4xx_errors": "Błędy 4xx (np. 404 Not Found) oznaczają niedziałające linki. Prowadzą one boty AI w ślepe zaułki, marnując budżet na indeksowanie i sygnalizując, że strona jest słabo utrzymana. Modele AI mogą nie odwoływać się do niepewnych, niedziałających źródeł.", 
    "3xx_redirects": "Wewnętrzne przekierowania spowalniają pracę robotów, zużywając niepotrzebnie zasoby na podążanie za łańcuchem odnośników. Preferowane są bezpośrednie linki do docelowych zasobów.", 
    "meta_description": "Meta descriptions dostarczają modelom AI zwięzłego podsumowania zawartości strony. Pomaga to w szybszym zrozumieniu kontekstu i może być wykorzystane do generowania fragmentów odpowiedzi.",
    "js_content": "Roboty wyszukiwarek i AI preferują treści dostępne bezpośrednio w kodzie HTML. Jeśli kluczowe informacje pojawiają się dopiero po wykonaniu skryptów JavaScript, może to opóźnić ich indeksowanie lub, w przypadku mniej zaawansowanych botów, całkowicie uniemożliwić ich odczytanie.", 
    "schema_data": "Dane strukturalne (Schema) to 'język ojczysty' sztucznej inteligencji. Używanie znaczników Schema (np. Article, FAQPage, Organization) pozwala precyzyjnie opisać zawartość strony w sposób zrozumiały dla maszyn. To jeden z najważniejszych czynników, który pozwala AI zrozumieć kontekst i fakty, co bezpośrednio przekłada się na jakość generowanych odpowiedzi.",
}

# --- HELPER FUNCTIONS ---
def read_data_file(file):
    if file is None: return None
    filename = file.name
    content = file.getvalue()
    
    if filename.lower().endswith('.csv'):
        encodings = ['utf-8-sig', 'utf-16', 'utf-8', 'latin-1']
        separators = [';', ',']
        for encoding in encodings:
            for sep in separators:
                try:
                    return pd.read_csv(io.BytesIO(content), sep=sep, encoding=encoding)
                except: continue
        return None
    elif filename.lower().endswith('.xlsx'):
        try:
            return pd.read_excel(io.BytesIO(content))
        except: return None
    return None

def set_cell_shading(cell, fill_color):
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:val'), 'clear')
    shading_elm.set(qn('w:color'), 'auto')
    shading_elm.set(qn('w:fill'), fill_color)
    cell._tc.get_or_add_tcPr().append(shading_elm)

def add_hyperlink(paragraph, text, url):
    part = paragraph.part
    r_id = part.relate_to(url, "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink", is_external=True)
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)
    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')
    new_run.append(rPr)
    new_run.text = text
    hyperlink.append(new_run)
    r = paragraph.add_run()
    r._r.append(hyperlink)
    r.font.color.rgb = RGBColor(0x05, 0x63, 0xC1)
    r.font.underline = True
    return hyperlink

def add_strategic_commentary(document, key, commentary_db):
    if key in commentary_db:
        document.add_paragraph(commentary_db[key])
    document.add_paragraph()

# Main UI
st.title("🚀 AI Readiness Report Generator")
st.markdown("Automatyczne narzędzie do audytu gotowości witryny na zmiany w ekosystemie AI (SGE/AIO).")

tabs = st.tabs(["📁 Dane Podstawowe", "🔧 Audyt Techniczny", "✍️ Treści i Social", "🤖 Opinie i LLM", "📄 Generuj Raport"])

# --- TAB 1: Dane Podstawowe ---
with tabs[0]:
    col1, col2 = st.columns([1, 1])
    with col1:
        analyzed_url = st.text_input("Adres analizowanej strony:", placeholder="https://example.com", key="url_input")
        client_name = st.text_input("Nazwa Klienta:", placeholder="Firma X", key="client_input")
    with col2:
        logo_file = st.file_uploader("Logo Klienta (PNG, JPG, SVG):", type=['png', 'jpg', 'jpeg', 'svg'])
        if logo_file:
            st.image(logo_file, width=150)

# --- TAB 2: Audyt Techniczny ---
with tabs[1]:
    st.subheader("Checklista Techniczna")
    tech_q = [
        "Czy strona jest dodana w Google Search Console?",
        "Czy strona jest dodana w Bing Webmaster Tools?",
        "Czy dodany jest plik robots.txt?",
        "Czy plik robots.txt pozwala agentom LLM i robotom wyszukiwarek na crawlowanie strony?",
        "Czy strona jest możliwa do crawlowania przez wyszukiwarki?",
        "Czy strona jest indeksowalna?",
        "Czy dodana jest mapa strony XML?",
        "Czy dodawany jest znacznik <lastmod> w mapie witryny XML?",
        "Czy dodane są dedykowane mapy XML ze zdjęciami i filmami?",
        "Czy mapa strony XML jest dodana do robots.txt?",
        "Czy mapa strony XML zawiera tylko adresy z kodem 200, kanoniczne, nie zawiera adresów noindex?",
        "Czy dodany jest certyfikat SSL?",
        "Czy dodane jest 'Organization' schema z adresami 'sameAs' kierującymi do profili społecznościowych?",
        "Czy dodane jest 'Article' schema na wpisach blogowych?",
        "Czy dodane jest 'Author' schema na wpisach blogowych i podlinkowane do profili autorów?",
        "Czy autorzy mają stworzone dedykowane podstrony z 'ProfilePage' schema?",
        "Czy dodane jest 'Breadcrumb' schema?",
        "Czy dodany jest plik llm.txt?",
        "Czy plik llm.txt jest poprawnie sformatowany i zawiera najważniejsze podstrony?"
    ]
    tech_answers = {}
    cols = st.columns(2)
    for i, q in enumerate(tech_q):
        tech_answers[q] = cols[i%2].selectbox(q, ["Tak", "Nie", "Do wdrożenia", "Nie dotyczy"], key=f"tech_{i}")

    st.divider()
    st.subheader("Pliki z narzędzi")
    col1, col2 = st.columns(2)
    robots_file = col1.file_uploader("Plik robots.txt (opcjonalnie):", type=['txt'])
    sf_file = col1.file_uploader("Plik z audytem Screaming Frog (Internal All):", type=['csv', 'xlsx'])
    gsc_img = col2.file_uploader("Screen z Google Search Console:", type=['png', 'jpg', 'jpeg'])
    ga_img = col2.file_uploader("Screen z Google Analytics (Ruch AI):", type=['png', 'jpg', 'jpeg'])
    
    col3, col4 = st.columns(2)
    ahrefs_file = col3.file_uploader("Ahrefs (AI Overview):", type=['csv', 'xlsx'])
    senuto_file = col3.file_uploader("Senuto (AI Overview):", type=['csv', 'xlsx'])
    js_file = col4.file_uploader("SF (JS Content Analysis):", type=['csv', 'xlsx'])
    schema_file = col4.file_uploader("SF (Structured Data):", type=['csv', 'xlsx'])

# --- TAB 3: Treści i Social ---
with tabs[2]:
    st.subheader("Audyt Treści")
    content_q = [
        "Czy Twoje najważniejsze strony zostały zaktualizowane w ciągu ostatnich 6 miesięcy?",
        "Czy widoczne są daty publikacji i ostatniej modyfikacji?",
        "Czy podstrony zawierają unikalną treść?",
        "Czy kluczowe strony zawierają sekcje FAQ?",
        "Czy dodane są linki do źródeł naukowych, raportów branżowych albo źródeł pierwotnych?"
    ]
    content_answers = {}
    for q in content_q:
        content_answers[q] = st.selectbox(q, ["Tak", "Nie", "Częściowo"], key=f"cont_{q}")

    st.divider()
    st.subheader("Social Media")
    social_q = [
        "Czy marka ma utworzony profil społecznościowy na Facebook? (podaj link)",
        "Czy marka ma utworzony profil społecznościowy na Instagram? (podaj link)",
        "Czy marka ma utworzony profil społecznościowy na Tiktok? (podaj link)",
        "Czy marka ma utworzony profil społecznościowy na Youtube? (podaj link)",
        "Czy w ciągu ostatniego miesiaca został dodany materiał na Facebook?",
        "Czy w ciągu ostatniego miesiaca został dodany materiał na Instagram?",
        "Czy w ciągu ostatniego miesiaca został dodany materiał na Tiktok?",
        "Czy w ciągu ostatniego miesiaca został dodany materiał na Youtube?"
    ]
    social_answers = {}
    for q in social_q:
        if "podaj link" in q:
            social_answers[q] = st.text_input(q, key=f"soc_{q}")
        else:
            social_answers[q] = st.selectbox(q, ["Tak", "Nie"], key=f"soc_{q}")

    st.divider()
    st.subheader("Linkbuilding")
    lb_q = [
        "Czy autorytet strony wyrażony DR rośnie lub jest stabilny?",
        "Czy stronie stale przybywa linków przychodzących?",
        "Czy linki przychodzące kierują do stron 404?"
    ]
    lb_answers = {}
    for q in lb_q:
        lb_answers[q] = st.selectbox(q, ["Tak", "Nie", "Wymaga analizy"], key=f"lb_{q}")
    lb_img = st.file_uploader("Screen z Ahrefs (Backlink profile):", type=['png', 'jpg', 'jpeg'])

# --- TAB 4: Opinie i LLM ---
with tabs[3]:
    st.info("Poniższe sekcje służą do ręcznego wpisania wyników testów w modelach LLM.")
    llm_models = ["ChatGPT", "Gemini", "Claude", "Perplexity", "Grok", "DeepSeek", "Bing (Copilot)"]
    
    st.subheader("Jakie są opinie o marce?")
    opinie_data = {}
    for model in llm_models:
        opinie_data[model] = st.text_area(f"Opinia w {model}:", placeholder="Wklej tutaj odpowiedź modelu...", height=100, key=f"op_{model}")
        
    st.divider()
    st.subheader("Zgodność informacji")
    zgodnosc_data = {}
    for model in llm_models:
        zgodnosc_data[model] = st.text_area(f"Zgodność w {model}:", placeholder="Czy informacje są prawdziwe?", height=100, key=f"zg_{model}")

# --- TAB 5: Generuj Raport ---
with tabs[4]:
    st.subheader("Finalizacja")
    if st.button("🚀 GENERUJ RAPORT (DOCX + XLSX)", type="primary"):
        if not analyzed_url:
            st.error("Podaj adres URL strony!")
        else:
            with st.spinner("Przetwarzanie danych i generowanie plików..."):
                try:
                    # 1. Image preprocessing
                    logo_bytes = None
                    if logo_file:
                        if logo_file.name.lower().endswith('.svg'):
                            logo_bytes = cairosvg.svg2png(bytestring=logo_file.getvalue())
                        else:
                            logo_bytes = logo_file.getvalue()

                    # 2. DOCX Generation
                    doc = Document()
                    # Header with Logo
                    p = doc.add_paragraph()
                    if logo_bytes:
                        try:
                            logo_io = io.BytesIO(logo_bytes)
                            p.add_run().add_picture(logo_io, width=Inches(1.5))
                        except Exception as logo_err:
                            st.warning(f"Problem z logiem: {logo_err}")
                    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                    
                    doc.add_heading('Raport AI Readiness', level=0)
                    doc.add_paragraph(f"Raport dla domeny: {analyzed_url}", 'Subtitle')
                    
                    doc.add_heading('Wstęp', level=1)
                    intro_text = ("Niniejszy dokument stanowi analizę gotowości witryny na zmiany w sposobie funkcjonowania wyszukiwarek internetowych, "
                                  "związane z wprowadzeniem generatywnych modeli językowych (LLM) i odpowiedzi AI (AI Overviews).")
                    doc.add_paragraph(intro_text)

                    # 1. Analiza widoczności
                    if gsc_img or ga_img:
                        doc.add_heading('1. Analiza widoczności i ruchu', level=1)
                        if gsc_img:
                            doc.add_heading('1.1. Widoczność w Google Search Console', level=2)
                            doc.add_picture(io.BytesIO(gsc_img.getvalue()), width=Inches(6.0))
                            doc.add_paragraph("<tutaj_dodaj_komentarz>").runs[0].font.italic = True
                        if ga_img:
                            doc.add_heading('1.2. Ruch z LLM w Google Analytics 4', level=2)
                            doc.add_picture(io.BytesIO(ga_img.getvalue()), width=Inches(6.0))
                            doc.add_paragraph("<tutaj_dodaj_komentarz>").runs[0].font.italic = True

                    # Helper for sections (Nested definition for scoped access or move outside)
                    def build_q_and_a_section(document, title, answers, commentary_db, robots_content=""):
                        document.add_heading(title, level=1)
                        for question, status in answers.items():
                            short_title = str(question).replace("(podaj link)", "").strip()
                            document.add_heading(short_title, level=2)
                            status_text = str(status).lower()
                            icon = "✅" if "tak" in status_text else "❌" if "nie" in status_text or not status_text else "➡️"
                            p = document.add_paragraph()
                            if str(status).startswith('http'):
                                p.add_run(f'{icon} '); add_hyperlink(p, status, status)
                            else:
                                p.add_run(f'{icon} {status if status else "Nie podano"}')
                            
                            if question == "Czy dodany jest plik robots.txt?":
                                add_strategic_commentary(document, question, commentary_db)
                                if robots_content:
                                    p_robots = document.add_paragraph(); p_robots.add_run("Zawartość pliku robots.txt:").bold = True
                                    p_quote = document.add_paragraph(robots_content); p_quote.style = 'Quote'
                            else:
                                add_strategic_commentary(document, question, commentary_db)

                    def add_styled_table(document, df, title):
                        document.add_heading(title, level=3)
                        if df is None or df.empty:
                            document.add_paragraph("Nie znaleziono danych spełniających kryteria."); return
                        table = document.add_table(rows=1, cols=len(df.columns)); table.style = 'Table Grid'
                        hdr_cells = table.rows[0].cells
                        for i, column_name in enumerate(df.columns):
                            p = hdr_cells[i].paragraphs[0]; run = p.add_run(str(column_name)); run.font.bold = True; run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF); set_cell_shading(hdr_cells[i], "4F81BD")
                        for index, row in df.head(10).iterrows(): # Limit to top 10 for visibility
                            row_cells = table.add_row().cells
                            for i, cell_value in enumerate(row): row_cells[i].text = str(cell_value)
                        for row in table.rows:
                            for cell in row.cells:
                                for p in cell.paragraphs:
                                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                                    for run in p.runs:
                                        if not run.font.bold: run.font.size = Pt(8)
                        document.add_paragraph()

                    # Sections
                    robots_text = robots_file.getvalue().decode('utf-8', errors='ignore') if robots_file else ""
                    build_q_and_a_section(doc, '2. Crawling i Indeksowanie', tech_answers, commentary_db, robots_text)
                    
                    doc.add_heading('3. Treści', level=1)
                    add_strategic_commentary(doc, 'tresci_general', commentary_db)
                    for q, a in content_answers.items():
                        doc.add_heading(str(q), level=2)
                        status_icon = "✅" if "tak" in str(a).lower() else "❌" if "nie" in str(a).lower() else "➡️"
                        doc.add_paragraph(f"{status_icon} {a}")
                    
                    doc.add_heading('4. Social Media', level=1)
                    add_strategic_commentary(doc, 'social_media_general', commentary_db)
                    for q, a in social_answers.items():
                        doc.add_heading(str(q).replace("(podaj link)", "").strip(), level=2)
                        p = doc.add_paragraph()
                        if str(a).startswith('http'):
                            p.add_run("🔗 "); add_hyperlink(p, str(a), str(a))
                        else:
                            status_icon = "✅" if "tak" in str(a).lower() else "➡️"
                            p.add_run(f"{status_icon} {a}")

                    # Ahrefs / Senuto Data
                    if ahrefs_file:
                        df_ahrefs = read_data_file(ahrefs_file)
                        if df_ahrefs is not None:
                            doc.add_heading('6. Analiza potencjału w AI Overviews (Ahrefs)', level=1)
                            if 'Current URL inside' in df_ahrefs.columns:
                                df_ahrefs_ai = df_ahrefs[df_ahrefs['Current URL inside'] == 'AI Overview'].sort_values(by='Organic traffic', ascending=False)
                                add_styled_table(doc, df_ahrefs_ai.head(5), "Top 5 fraz z AI Overview")

                    # Screaming Frog Data
                    if sf_file:
                        df_sf = read_data_file(sf_file)
                        if df_sf is not None:
                            doc.add_heading('7. Analiza techniczna (Screaming Frog)', level=1)
                            
                            # non-indexable
                            if 'Indexability' in df_sf.columns:
                                non_idx = df_sf[df_sf['Indexability'] == 'Non-Indexable'][['Address', 'Indexability Status', 'Status Code']]
                                add_styled_table(doc, non_idx, f"Strony nieindeksowalne ({len(non_idx)})")
                                add_strategic_commentary(doc, 'non_indexable', commentary_db)
                            
                            # CWV
                            cwv_cols = ['Largest Contentful Paint Time (ms)', 'Cumulative Layout Shift', 'First Contentful Paint Time (ms)']
                            if all(c in df_sf.columns for c in cwv_cols):
                                doc.add_heading('7.2. Core Web Vitals', level=2)
                                add_strategic_commentary(doc, 'core_web_vitals', commentary_db)
                                # Simple summary table for brevity in report
                                doc.add_paragraph("Szczegóły znajdują się w załączonym pliku Excel.")

                            # Errors 4xx
                            if 'Status Code' in df_sf.columns:
                                err4xx = df_sf[(df_sf['Status Code'] >= 400) & (df_sf['Status Code'] < 500)]
                                if not err4xx.empty:
                                    add_styled_table(doc, err4xx[['Address', 'Status Code']].head(10), f"Błędy 4xx ({len(err4xx)})")
                                add_strategic_commentary(doc, '4xx_errors', commentary_db)
                                
                                err3xx = df_sf[(df_sf['Status Code'] >= 300) & (df_sf['Status Code'] < 400)]
                                if not err3xx.empty:
                                    add_styled_table(doc, err3xx[['Address', 'Redirect URL']].head(10), f"Przekierowania 3xx ({len(err3xx)})")
                                add_strategic_commentary(doc, '3xx_redirects', commentary_db)

                    if schema_file:
                        df_schema = read_data_file(schema_file)
                        if df_schema is not None:
                            doc.add_heading('7.7. Dane Strukturalne (Schema)', level=2)
                            add_styled_table(doc, df_schema.head(10), "Podgląd danych Schema")
                            add_strategic_commentary(doc, 'schema_data', commentary_db)

                    # LLM Sections
                    doc.add_heading('8. Opinie LLM', level=1)
                    for model, text in opinie_data.items():
                        if text:
                            doc.add_heading(f"Opinia: {model}", level=2)
                            doc.add_paragraph(text)

                    doc.add_heading('9. Zgodność informacji', level=1)
                    for model, text in zgodnosc_data.items():
                        if text:
                            doc.add_heading(f"Zgodność: {model}", level=2)
                            doc.add_paragraph(text)
                    
                    # 3. XLSX Generation
                    xlsx_buffer = io.BytesIO()
                    with pd.ExcelWriter(xlsx_buffer, engine='openpyxl') as writer:
                        # Summary Sheet
                        all_summary = {**tech_answers, **content_answers, **social_answers, **lb_answers}
                        pd.DataFrame(list(all_summary.items()), columns=['Pytanie', 'Odpowiedź']).to_excel(writer, sheet_name='Podsumowanie', index=False)
                        
                        if sf_file:
                            df_sf = read_data_file(sf_file)
                            if df_sf is not None:
                                df_sf.head(100).to_excel(writer, sheet_name='Audyt Screaming Frog', index=False)
                        if ahrefs_file:
                            df_ahrefs = read_data_file(ahrefs_file)
                            if df_ahrefs is not None:
                                df_ahrefs.to_excel(writer, sheet_name='Ahrefs AIO', index=False)

                    # Final export
                    doc_io = io.BytesIO()
                    doc.save(doc_io)
                    doc_io.seek(0)
                    xlsx_buffer.seek(0)
                    
                    st.success("Raport wygenerowany!")
                    
                    col1, col2 = st.columns(2)
                    col1.download_button(
                        label="📥 Pobierz Raport DOCX",
                        data=doc_io,
                        file_name=f"Raport_AI_Readiness_{client_name if client_name else 'Klient'}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    col2.download_button(
                        label="📥 Pobierz Dane XLSX",
                        data=xlsx_buffer.getvalue(),
                        file_name=f"Dane_AI_Readiness_{client_name if client_name else 'Klient'}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except Exception as e:
                    st.error(f"Błąd podczas generowania: {e}")
                    st.code(traceback.format_exc())

st.sidebar.markdown(f"""
### Status Projektu
**Domena:** `{analyzed_url if analyzed_url else 'Brak'}`
**Klient:** `{client_name if client_name else 'Brak'}`
""")
if not analyzed_url:
    st.sidebar.warning("Uzupełnij dane w pierwszej zakładce!")
else:
    st.sidebar.success("Projekt zainicjowany.")

st.sidebar.divider()
st.sidebar.subheader("💡 Przykładowe Narzędzie")
st.sidebar.markdown("Zobacz jak wygląda gotowy raport lub przetestuj narzędzie używając przykładowych danych.")

def load_example_data():
    st.session_state["url_input"] = "https://oralb.pl"
    st.session_state["client_input"] = "Oral-B"
    # Tech
    for i in range(19):
        st.session_state[f"tech_{i}"] = "Tak"
    # Content
    content_questions = [
        "Czy Twoje najważniejsze strony zostały zaktualizowane w ciągu ostatnich 6 miesięcy?",
        "Czy widoczne są daty publikacji i ostatniej modyfikacji?",
        "Czy podstrony zawierają unikalną treść?",
        "Czy kluczowe strony zawierają sekcje FAQ?",
        "Czy dodane są linki do źródeł naukowych, raportów branżowych albo źródeł pierwotnych?"
    ]
    for q in content_questions:
        st.session_state[f"cont_{q}"] = "Tak"
    # Social Medias
    social_questions = [
        "Czy marka ma utworzony profil społecznościowy na Facebook? (podaj link)",
        "Czy marka ma utworzony profil społecznościowy na Instagram? (podaj link)",
        "Czy marka ma utworzony profil społecznościowy na Tiktok? (podaj link)",
        "Czy marka ma utworzony profil społecznościowy na Youtube? (podaj link)",
        "Czy w ciągu ostatniego miesiaca został dodany materiał na Facebook?",
        "Czy w ciągu ostatniego miesiaca został dodany materiał na Instagram?",
        "Czy w ciągu ostatniego miesiaca został dodany materiał na Tiktok?",
        "Czy w ciągu ostatniego miesiaca został dodany materiał na Youtube?"
    ]
    for q in social_questions:
        if "podaj link" in q:
            st.session_state[f"soc_{q}"] = "https://facebook.com/oralb" if "Facebook" in q else "https://instagram.com/oralb"
        else:
            st.session_state[f"soc_{q}"] = "Tak"
    # Linkbuilding
    lb_questions = [
        "Czy autorytet strony wyrażony DR rośnie lub jest stabilny?",
        "Czy stronie stale przybywa linków przychodzących?",
        "Czy linki przychodzące kierują do stron 404?"
    ]
    for q in lb_questions:
        st.session_state[f"lb_{q}"] = "Tak"
    # LLM Options
    llm_models_list = ["ChatGPT", "Gemini", "Claude", "Perplexity", "Grok", "DeepSeek", "Bing (Copilot)"]
    for model in llm_models_list:
        st.session_state[f"op_{model}"] = "Marka Oral-B według modelu AI jest powszechnie uznawana za jednego z najpopularniejszych liderów zaufania z dziedziny pielęgnacji jamy ustnej. Rekomendowana przez dentystów na całym świecie..."
        st.session_state[f"zg_{model}"] = "Większość wygenerowanych danych technicznych jest w pełnej zgodności ze stanem faktycznym i rynkową wiedzą ekspercką."

if st.sidebar.button("✨ Uzupełnij formularz", on_click=load_example_data):
    st.sidebar.success("Formularz wypełniony danymi Oral-B!")

try:
    with open("example/Raport_AI_Readiness_Ekspercki (19).docx", "rb") as f:
        docx_bytes = f.read()
    st.sidebar.download_button(
        label="📄 Pobierz wzór Raportu (DOCX)",
        data=docx_bytes,
        file_name="Przykladowy_Raport_AI_Readiness_OralB.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
except Exception as e:
    st.sidebar.error(f"Brak pliku przykładowego DOCX w /example")

try:
    with open("example/Analiza_Techniczna_Wyniki (18).xlsx", "rb") as f:
        xlsx_bytes = f.read()
    st.sidebar.download_button(
        label="📊 Pobierz wzór Danych (XLSX)",
        data=xlsx_bytes,
        file_name="Przykladowa_Analiza_Techniczna_OralB.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
except Exception as e:
    st.sidebar.error(f"Brak pliku przykładowego XLSX w /example")