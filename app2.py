import streamlit as st
import pandas as pd
from rapidfuzz import process, fuzz
import io
import openpyxl

# 1. Konfiguracja strony (musi być ZAWSZE pierwsza)
st.set_page_config(
    page_title="URL Matcher",
    page_icon="🔒",
    layout="wide"
)

# --- MODUŁ LOGOWANIA ---
def check_password():
    """Zwraca `True` jeśli użytkownik podał poprawne hasło."""

    def password_entered():
        """Sprawdza czy wpisane hasło zgadza się z tym w sekretach."""
        if st.session_state["password"] == st.secrets["APP_PASSWORD"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Nie przechowujemy hasła
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        # Pierwsze uruchomienie, pokaż formularz
        st.text_input(
            "Podaj hasło dostępu:", 
            type="password", 
            on_change=password_entered, 
            key="password"
        )
        st.info("Dane potrzebne do zalogowania znajdują się w Monday. Kontakt: jaroslaw.muzyka@performance-group.pl")
        return False
    elif not st.session_state["password_correct"]:
        # Hasło błędne
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
        # Hasło poprawne
        return True

if not check_password():
    st.stop()  # Zatrzymuje aplikację, jeśli brak autoryzacji

# =========================================================
# WŁAŚCIWA APLIKACJA (Kod wykonuje się tylko po zalogowaniu)
# =========================================================

# --- Funkcja przetwarzająca (zoptymalizowana i cache'owana) ---
@st.cache_data(show_spinner=False)
def process_file(file_bytes, filename):
    """
    Funkcja przetwarza plik Excel lub CSV w pamięci.
    """
    try:
        url1_data = []
        candidates = []
        
        if filename.endswith('.csv'):
            # Próba odczytu pliku CSV
            # Zakładamy, że pierwszy wiersz to nagłówki i używamy pandas
            df = pd.read_csv(io.BytesIO(file_bytes), header=0, sep=None, engine='python')
            
            if df.shape[1] < 2:
                return None, "Plik CSV musi zawierać co najmniej dwie kolumny (A i B)."
            
            # Pobieranie URL1 (z pominięciem NaN/pustych)
            all_url1_values = df.iloc[:, 0].dropna().astype(str).str.strip().tolist()
            url1_data = [{'URL1': u} for u in all_url1_values if u]
            
            # Pobieranie URL2
            all_url2_values = df.iloc[:, 1].dropna().astype(str).str.strip().tolist()
            candidates = list(set([u for u in all_url2_values if u]))
            
        else:
            # Wczytanie skoroszytu z data_only=True dla plików Excel
            workbook = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
            sheet = workbook.active
            
            # Pobieranie danych (pomijamy nagłówek - row 1)
            # 1. Kandydaci z kolumny B (URL2)
            all_url2_values = [
                str(cell.value).strip() 
                for cell in sheet['B'][1:] 
                if cell.value is not None and str(cell.value).strip() != ''
            ]
            candidates = list(set(all_url2_values)) # Unikalne wartości
            
            # 2. Cele z kolumny A (URL1)
            for cell_a in sheet['A'][1:]:
                if cell_a.value is not None and str(cell_a.value).strip() != '':
                    url1_data.append({
                        'URL1': str(cell_a.value).strip()
                    })
        
        if not url1_data:
            return None, "Nie znaleziono danych w pierwszej kolumnie."
            
        return url1_data, candidates

    except Exception as e:
        return None, f"Błąd podczas odczytu pliku: {e}"

def run_matching(url1_data, candidates):
    results = []
    total_items = len(url1_data)
    
    stop_btn_placeholder = st.empty()
    stop_btn_placeholder.button("🛑 Zatrzymaj analizę", key="stop_analysis_btn")

    # Pasek postępu
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    for idx, item in enumerate(url1_data):
        url1 = item['URL1']
        
        matches = process.extract(url1, candidates, scorer=fuzz.WRatio, limit=3)
        
        row = {
            'URL 1 (do przekierowania)': url1
        }
        
        for i, match in enumerate(matches):
            url_match, similarity, _ = match
            row[f'Propozycja {i+1}'] = url_match
            row[f'% {i+1}'] = similarity
            
        results.append(row)
        
        # Aktualizacja paska postępu
        if idx % 10 == 0 or idx == total_items - 1:
            progress = (idx + 1) / total_items
            progress_bar.progress(progress)
            status_text.text(f"Przetwarzanie: {idx + 1}/{total_items}")

    progress_bar.empty()
    status_text.empty()
    stop_btn_placeholder.empty()
    
    df = pd.DataFrame(results)
    
    if '% 1' in df.columns:
        df = df.sort_values(by='% 1', ascending=False)
        
    for i in range(1, 4):
        col_name = f'% {i}'
        if col_name in df.columns:
            df[col_name] = df[col_name].apply(lambda x: f"{x:.2f}")
            
    return df

def color_percentage(val):
    """Funkcja do kolorowania wyników procentowych w zależności od siły dopasowania."""
    try:
        f_val = float(val)
        if f_val >= 85:
            # Subtelny zielony dla świetnego dopasowania
            return 'background-color: rgba(46, 204, 113, 0.2); font-weight: bold;'
        elif f_val >= 50:
            # Żółty/Pomarańczowy dla umiarkowanego
            return 'background-color: rgba(241, 196, 15, 0.2);'
        else:
            # Subtelny czerwony dla słabego
            return 'background-color: rgba(231, 76, 60, 0.1); color: #c0392b;'
    except:
        return ''

# --- Interfejs Użytkownika (UI) ---

st.title("🚀 URL Matcher")
st.markdown("""
Narzędzie rozwiązuje problem masowego mapowania przekierowań 301, znajdując najbardziej zbliżone adresy URL z aktualnej oferty dla starych lub wygasłych linków.
""")

with st.expander("📖 Instrukcja"):
    st.markdown("""
    1. Wgraj plik XLSX lub CSV z dwiema kolumnami:
        * **Kolumna A:** lista adresów do przekierowania (np. adresy z kodem 404)
        * **Kolumna B:** lista wszystkich aktualnych adresów URL (kandydaci do przekierowań)
    2. Uruchom skrypt analizy podobieństwa i poczekaj na wyniki analizy.  
    3. Pobierz wynik z propozycjami dopasowań. 
    """)

with st.expander("💡 Przykład zastosowania"):
    st.markdown("""
    W sklepie internetowym zostały usunięte produkty. Adresy URL odpowiadają teraz kodem 404. Celem jest przekierowanie starych adresów na nowe produkty, dostępne aktualnie w ofercie sklepu. Pobieramy listę wszystkich aktualnie dostępnych produktów np. z sitemap.xml lub z crawla i porównujemy je w narzędziu z adresami z kodem 404.
    """)

st.markdown("**Przykładowy plik:**")

try:
    with open("przykladowy-plik.xlsx", "rb") as f:
        przykladowy_plik_bytes = f.read()
    st.download_button(
        label="📥 Pobierz przykładowy plik (przykladowy-plik.xlsx)",
        data=przykladowy_plik_bytes,
        file_name="przykladowy-plik.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
except FileNotFoundError:
    st.warning("Przykładowy plik (przykladowy-plik.xlsx) nie został znaleziony w głównym katalogu.")

st.info("Wymagany format pliku: Excel (.xlsx) lub plik wartości oddzielonych przecinkami (.csv). Kolumna A: URL do sprawdzenia, Kolumna B: Baza adresów.")

uploaded_file = st.file_uploader("Wgraj plik z adresami (.xlsx, .csv)", type=['xlsx', 'csv'])

if uploaded_file is not None:
    if st.button("Uruchom analizę", type="primary"):
        
        with st.spinner("Wczytywanie i analiza pliku..."):
            bytes_data = uploaded_file.getvalue()
            data_extracted, candidates_or_error = process_file(bytes_data, uploaded_file.name)
            
            if data_extracted is None:
                st.error(f"Wystąpił błąd: {candidates_or_error}")
            else:
                # Krok 2: Uruchomienie dopasowywania
                final_df = run_matching(data_extracted, candidates_or_error)
                
                # --- WYLICZENIE STATYSTYK ---
                if '% 1' in final_df.columns:
                    high_matches = sum(1 for x in final_df['% 1'].astype(float) if x >= 85)
                    medium_matches = sum(1 for x in final_df['% 1'].astype(float) if 50 <= x < 85)
                    low_matches = sum(1 for x in final_df['% 1'].astype(float) if x < 50)
                else:
                    high_matches = medium_matches = low_matches = 0
                
                st.success(f"Analiza zakończona! Przetworzono {len(data_extracted)} adresów na podstawie {len(candidates_or_error)} kandydatów ze wskazanej bazy.")

                st.subheader("Statystyki dopasowań")
                col1, col2, col3, col4 = st.columns(4)
                col1.metric("Wszystkie adresy", len(final_df))
                col2.metric("🟢 Świetne dop. (≥85%)", high_matches)
                col3.metric("🟡 Umiarkowane (50-84%)", medium_matches)
                col4.metric("🔴 Słabe (<50%)", low_matches)

                # --- WYŚWIETLENIE WYNIKÓW Z KOLORAMI ---
                st.subheader("Podgląd wyników")
                
                # Formatowanie kolumn z procentami
                subset_cols = [col for col in ['% 1', '% 2', '% 3'] if col in final_df.columns]
                styled_df = final_df.head(100).style.map(color_percentage, subset=subset_cols)
                
                st.dataframe(styled_df, use_container_width=True)
                
                # Krok 4: Eksport do Excela
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                    final_df.to_excel(writer, index=False, sheet_name='Wyniki')
                
                st.download_button(
                    label="📥 Pobierz wyniki (.xlsx)",
                    data=buffer.getvalue(),
                    file_name=f"wyniki_{uploaded_file.name}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
