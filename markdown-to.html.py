# @title 🛠️ GENERATOR RAPORTÓW (Uruchom i wgraj pliki .md)
# 1. Instalacja bibliotek
!pip install markdown jinja2 > /dev/null

import markdown
import os
import shutil
from google.colab import files
from jinja2 import Template
from datetime import datetime
import re

# ==========================================
# 🎨 KONFIGURACJA KOLORÓW (EDYTUJ TUTAJ)
# ==========================================
# Wpisz tutaj swoje kody HEX
PRIMARY_VAL   = "#003366"  # Główny kolor (np. Granat)
SECONDARY_VAL = "#006699"  # Drugi kolor (np. Niebieski)
TERTIARY_VAL  = "#FF9900"  # Akcent (np. Pomarańcz)
LIGHT_VAL     = "#F5F7FA"  # Tło (Jasny szary)

# ==========================================
# ⚙️ LOGIKA BRANDOWA (Z Twojego skryptu)
# ==========================================

def lighten_color(hex_color, amount=0.3):
    """Rozjaśnia kolor hex o określoną wartość (wg Twojego wzoru)"""
    hex_color = hex_color.lstrip('#')
    rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
    new_rgb = tuple(int(c + (255 - c) * amount) for c in rgb)
    return '#{:02x}{:02x}{:02x}'.format(*new_rgb)

def get_brand_colors():
    """Zwraca słownik z kolorami brandowymi i ich pochodnymi"""
    return {
        'primary': PRIMARY_VAL,
        'secondary': SECONDARY_VAL,
        'tertiary': TERTIARY_VAL,
        'light': LIGHT_VAL,
        'primary_light': lighten_color(PRIMARY_VAL, 0.9), # Bardzo jasne tło dla tabel
        'secondary_light': lighten_color(SECONDARY_VAL, 0.3),
        'tertiary_light': lighten_color(TERTIARY_VAL, 0.3),
        'border_color': lighten_color(PRIMARY_VAL, 0.8)
    }

# Pobieramy paletę
colors = get_brand_colors()

# ==========================================
# 📄 SZABLON HTML + CSS (Manrope & Styles)
# ==========================================

html_template = """
<!DOCTYPE html>
<html lang="pl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{{ title }}</title>
    <!-- Import Manrope Font -->
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
    <link href="https://fonts.googleapis.com/css2?family=Manrope:wght@300;400;500;700;800&display=swap" rel="stylesheet">

    <style>
        :root {
            --primary: {{ c.primary }};
            --secondary: {{ c.secondary }};
            --tertiary: {{ c.tertiary }};
            --light: {{ c.light }};
            --primary-light: {{ c.primary_light }};
            --border-color: {{ c.border_color }};
        }

        body {
            font-family: 'Manrope', sans-serif;
            background-color: var(--light);
            color: #2d3436;
            line-height: 1.6;
            margin: 0;
            padding: 40px 20px;
        }

        .container {
            max-width: 900px;
            margin: 0 auto;
            background: white;
            padding: 60px;
            border-radius: 16px;
            box-shadow: 0 15px 35px rgba(0,0,0,0.08);
            border-top: 8px solid var(--primary);
        }

        /* --- TYPOGRAFIA --- */
        h1, h2, h3, h4, h5, h6 {
            font-family: 'Manrope', sans-serif;
            color: var(--primary);
            margin-top: 1.5em;
            margin-bottom: 0.5em;
            font-weight: 800;
            letter-spacing: -0.02em;
        }

        h1 {
            font-size: 2.8rem;
            padding-bottom: 20px;
            border-bottom: 2px solid var(--light);
            margin-top: 0;
            color: var(--primary);
        }

        h2 {
            font-size: 1.8rem;
            margin-top: 2em;
            padding-left: 15px;
            border-left: 6px solid var(--tertiary);
            color: var(--secondary);
        }

        h3 {
            font-size: 1.4rem;
            color: var(--primary);
            font-weight: 700;
        }

        p { margin-bottom: 1.2em; font-weight: 400; color: #4a4a4a; }
        strong { font-weight: 800; color: var(--primary); }

        ul, ol { margin-bottom: 1.5em; padding-left: 25px; }
        li { margin-bottom: 0.5em; }

        a {
            color: var(--secondary);
            text-decoration: none;
            font-weight: 700;
            border-bottom: 2px solid rgba(0,0,0,0.1);
            transition: 0.2s;
        }
        a:hover { color: var(--tertiary); border-color: var(--tertiary); }

        /* --- TABELE (Kluczowe dla raportu) --- */
        table {
            width: 100%;
            border-collapse: separate;
            border-spacing: 0;
            margin: 30px 0;
            font-size: 0.95rem;
            border-radius: 8px;
            overflow: hidden;
            border: 1px solid var(--border-color);
        }

        thead {
            background-color: var(--primary);
            color: white;
        }

        th {
            padding: 16px;
            text-align: left;
            font-weight: 700;
            text-transform: uppercase;
            font-size: 0.8rem;
            letter-spacing: 0.05em;
        }

        td {
            padding: 14px 16px;
            border-bottom: 1px solid var(--border-color);
            vertical-align: top;
        }

        tr:last-child td { border-bottom: none; }
        tr:nth-child(even) { background-color: var(--primary-light); }
        tr:hover { background-color: rgba(0,0,0,0.02); }

        /* --- CODE BLOCKS --- */
        pre {
            background-color: #263238;
            color: #eceff1;
            padding: 20px;
            border-radius: 8px;
            overflow-x: auto;
            font-family: 'Consolas', monospace;
            font-size: 0.9rem;
            margin: 20px 0;
            border-left: 5px solid var(--tertiary);
        }

        code {
            font-family: 'Consolas', monospace;
            background-color: var(--light);
            padding: 2px 6px;
            border-radius: 4px;
            color: #d63031;
            font-size: 0.9em;
            border: 1px solid #e1e1e1;
        }

        pre code {
            background-color: transparent;
            color: inherit;
            border: none;
            padding: 0;
        }

        /* --- BLOCKQUOTES --- */
        blockquote {
            background: var(--light);
            border-left: 6px solid var(--secondary);
            margin: 30px 0;
            padding: 20px 30px;
            font-style: italic;
            color: var(--primary);
            border-radius: 0 8px 8px 0;
        }

        /* --- HR --- */
        hr {
            border: 0;
            height: 2px;
            background: var(--light);
            margin: 50px 0;
        }

        /* --- FOOTER --- */
        .footer {
            margin-top: 60px;
            padding-top: 20px;
            border-top: 1px solid var(--border-color);
            font-size: 0.8rem;
            color: #999;
            text-align: center;
            font-family: 'Manrope', sans-serif;
        }
    </style>
</head>
<body>
    <div class="container">
        {{ content }}
        <div class="footer">
            Raport wygenerowany: {{ date }} | Style zgodne z Brand Guidelines
        </div>
    </div>
</body>
</html>
"""

# ==========================================
# 🚀 GŁÓWNA FUNKCJA KONWERTUJĄCA
# ==========================================

def process_markdown_files():
    # 1. Wyczyść poprzednie
    output_dir = "gotowe_raporty_html"
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    os.makedirs(output_dir)

    # 2. Upload
    print("⬆️ WGRAJ PLIKI .MD TERAZ:")
    uploaded = files.upload()

    if not uploaded:
        print("❌ Nie wybrano plików.")
        return

    print("\n🔄 Przetwarzanie...")

    files_processed = 0

    for filename, content in uploaded.items():
        if not filename.lower().endswith('.md'):
            continue

        # Dekodowanie treści
        try:
            text_content = content.decode('utf-8')
        except UnicodeDecodeError:
            text_content = content.decode('latin-1')

        # Konwersja MD -> HTML
        # Używamy rozszerzeń do tabel (tables), spisu treści (toc) i bloków kodu (fenced_code)
        html_body = markdown.markdown(
            text_content,
            extensions=['tables', 'fenced_code', 'toc', 'nl2br', 'sane_lists']
        )

        # Renderowanie szablonu Jinja
        template = Template(html_template)
        final_html = template.render(
            content=html_body,
            title=filename.replace('.md', '').replace('_', ' ').title(),
            c=colors,
            date=datetime.now().strftime("%d-%m-%Y %H:%M")
        )

        # Zapis
        output_filename = os.path.join(output_dir, filename.replace('.md', '.html'))
        with open(output_filename, 'w', encoding='utf-8') as f:
            f.write(final_html)

        print(f"✅ Utworzono: {output_filename}")
        files_processed += 1

    # 3. Pakowanie i pobieranie
    if files_processed > 0:
        shutil.make_archive("raporty", 'zip', output_dir)
        print(f"\n🎉 Sukces! Przetworzono plików: {files_processed}.")
        print("📦 Pobieranie ZIP...")
        files.download("raporty.zip")
    else:
        print("⚠️ Nie znaleziono plików .md w wgranym zestawie.")

# Uruchomienie
if __name__ == "__main__":
    process_markdown_files()