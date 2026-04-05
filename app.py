import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import json
import io
import os
import pandas as pd
from PIL import Image
import pypdf

def set_document_formatting(doc):
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Calibri'
    font.size = Pt(12)
    
    # Also set for parsed styles if needed, but Normal covers most
    for style in doc.styles:
        if hasattr(style, 'font'):
            style.font.name = 'Calibri'


def extract_content_from_files(uploaded_files):
    """
    Extracts text and images from a LIST of uploaded files.
    Returns: (text_content, list_of_images)
    """
    text_content = ""
    images = []
    
    if not uploaded_files:
        return text_content, images
        
    # Ensure list
    if not isinstance(uploaded_files, list):
        uploaded_files = [uploaded_files]

    for uploaded_file in uploaded_files:

        try:
            if uploaded_file.type.startswith('image'):
                image = Image.open(uploaded_file)
                images.append(image)
                text_content += f"\n[Obrázek: {uploaded_file.name}]\n"
            
            elif "spreadsheet" in uploaded_file.type or uploaded_file.name.endswith('.xlsx'):
                df = pd.read_excel(uploaded_file)
                text_content += f"\n--- {uploaded_file.name} ---\n" + df.to_markdown(index=False) + "\n"
            
            elif "wordprocessing" in uploaded_file.type or uploaded_file.name.endswith('.docx'):
                doc_file = io.BytesIO(uploaded_file.getvalue())
                doc = Document(doc_file)
                full_text = []
                for para in doc.paragraphs:
                    full_text.append(para.text)
                text_content += f"\n--- {uploaded_file.name} ---\n" + '\n'.join(full_text) + "\n"

            elif "pdf" in uploaded_file.type or uploaded_file.name.endswith('.pdf'):
                reader = pypdf.PdfReader(uploaded_file)
                pdf_text = ""
                for page in reader.pages:
                    pdf_text += page.extract_text() + "\n"
                text_content += f"\n--- {uploaded_file.name} ---\n" + pdf_text + "\n"
            
            elif uploaded_file.type == "text/plain" or uploaded_file.name.endswith('.txt') or uploaded_file.name.endswith('.csv'):
                stringio = io.StringIO(uploaded_file.getvalue().decode("utf-8"))
                text_content += f"\n--- {uploaded_file.name} ---\n" + stringio.read() + "\n"

        except Exception as e:
            st.error(f"Chyba při zpracování souboru {uploaded_file.name}: {e}")
        
    return text_content, images

def generate_lab_report(api_key, model_name, topic, inputs_map, is_handwritten=False):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)

    theory_length = "ZKRÁCENÝ ROZSAH: Maximálně půl strany A4! Piš velmi stručně a jen to nejdůležitější, protože to bude student přepisovat ručně." if is_handwritten else "Rozsah: cca 1.5 strany A4, min. 1000 slov"
    conclusion_length = "ZKRÁCENÝ ROZSAH: Maximálně půl strany A4! Stručné zhodnocení." if is_handwritten else "Fakta a detailní analýza"

    system_prompt = f"""
    Jsi student 3. ročníku SPŠE (Střední průmyslová škola elektrotechnická). 
    Tvým úkolem je napsat školní laboratorní protokol (elaborát).
    
    Téma: {topic}

    DŮLEŽITÉ PRAVIDLO PRO CHYBĚJÍCÍ ZDROJE:
    Pokud u jakékoliv sekce (Teorie, Postup, Závěr) zjistíš, že nebyly poskytnuty ŽÁDNÉ podklady (žádný text ani relevantní nápověda), tvojí povinností je tuto sekci SAMOSTATNĚ VYGENEROVAT podle nejlepších znalostí k danému tématu. Na úplný začátek této dovygenerované sekce však MUSÍŠ přidat přesně tuto větu velkými písmeny:
    "NEBYL PŘILOŽEN ZDROJ INFORMACÍ, AI TI TO VYGENEROVALA. ZKONTROLUJ SI TO!\\n\\n"

    POSTUPUJ PODLE TĚCHTO SEKCI:

    1. TEORIE ({theory_length})
       - Vycházej z přiloženého textu/osnovy:
       {inputs_map.get('theory_text', '')}
       - ODDĚLENĚ NAHRANÉ GRAFICKÉ PRŮBĚHY:
       {inputs_map.get('waveforms_text', '')}
       - Pokud není nic přiloženo, sekci kompletně vygeneruj a nezapomeň na povinnou větu: "NEBYL PŘILOŽEN ZDROJ INFORMACÍ...".
       - Text musí být odborný a vyčerpávající. Vysvětli fyzikální principy, vzorce, odvození a souvislosti (při ručním psaní pouze to nezbytné).
       - KRITICKY DŮLEŽITÉ: Učitelé velmi potrpí na grafické průběhy. Vyhodnoť odděleně nahrané obrázky grafických průběhů (případně ty v zadání) a v teoretickém úvodu je detailně textově popiš a zanalyzuj. Vysvětli, co znamenají.

    2. POSTUP MĚŘENÍ
       - PŘEPIŠ přiložený text pracovního postupu do 1. OSOBY MINULÉHO ČASU (např. změň "Změřte napětí" na "Změřil jsem napětí").
       - Zdrojový text postupu:
       {inputs_map.get('procedure_text', '')}
       - Pokud není přiložen postup, logicky jej k tématu dovygeneruj a nezapomeň na povinnou větu: "NEBYL PŘILOŽEN ZDROJ INFORMACÍ...".
       - Pokud jsou přiloženy obrázky, odkazuj se na ně textem (např. "jak je vidět na obrázku 1").

    3. PŘÍKLAD VÝPOČTU
       - Na základě naměřených dat ({inputs_map.get('data_text', '')}) a teorie vytvoř JEDEN KONKRÉTNÍ PŘÍKLAD výpočtu.
       - Uveď vzorec, dosaď konkrétní naměřené hodnoty (např. U=10V, I=2A) a vypočítej výsledek. Výpočet musí být fyzikálně správný.

    4. ZÁVĚR ({conclusion_length})
       - Vycházej z naměřených hodnot ({inputs_map.get('data_text', '')}) a přiložené osnovy:
       {inputs_map.get('conclusion_text', '')}
       - Pokud osnova závěru chybí (nebo chybí data), závěr dovygeneruj obecněji na základě tématu a teorie, a dej na začátek povinnou větu: "NEBYL PŘILOŽEN ZDROJ INFORMACÍ...".
       - Zhodnoť měření technicky a kriticky. CITUJ KONKRÉTNÍ HODNOTY z naměřených dat (pokud vůbec nějaká jsou). Porovnej s teorií.

    DALŠÍ VSTUPY:
    --- ZADÁNÍ ---
    {inputs_map.get('assignment_text', '')}
    
    --- POUŽITÉ PŘÍSTROJE ---
    {inputs_map.get('instruments_text', '')}

    DŮLEŽITÉ: Rovnice piš jako prostý text (R=U/I). U všech desetinných čísel v textu používej výhradně desetinnou čárku, nikoliv tečku (např. 0,5 místo 0.5). Pro znak násobení ve výpočtech VŽDY používej tečku (.), nepoužívej hvězdičku (*).
    
    Vygeneruj výstup POUZE jako validní JSON s následující strukturou:
    {{
        "teorie": "...",
        "postup": "...",
        "priklad_vypoctu": "...",
        "zaver": "..."
    }}
    """
    
    content_parts = [system_prompt]
    
    # Add all images found in inputs
    for img_list in inputs_map.get('images_lists', []):
        content_parts.extend(img_list)

    try:
        response = model.generate_content(content_parts)
        text_response = response.text
        if "```json" in text_response:
            text_response = text_response.split("```json")[1].split("```")[0]
        elif "```" in text_response:
            text_response = text_response.split("```")[1].split("```")[0]
        return json.loads(text_response)
    except Exception as e:
        st.error(f"Chyba při generování s AI: {e}")
        return None

def fill_template_docx(template_path, topic, inputs_map, ai_content):
    if not os.path.exists(template_path):
        st.error(f"Šablona {template_path} nenalezena!")
        doc = Document()
    else:
        doc = Document(template_path)

    # Map headers to content
    sections_map = {
        "Teoretický úvod": {"text": ai_content.get('teorie', ''), "images": []},
        "Schéma zapojení": {"text": "", "images": inputs_map.get('schema_images', [])},
        "Postup měření": {"text": ai_content.get('postup', ''), "images": []},
        "Naměřené a vypočítané hodnoty": {"text": "", "images": inputs_map.get('data_images', [])},
        "Příklad výpočtu": {"text": ai_content.get('priklad_vypoctu', ''), "images": []},
        "Závěr": {"text": ai_content.get('zaver', ''), "images": []}
    }
    
    def add_content_after(paragraph, text, images):
        parent = paragraph._element.getparent()
        index = parent.index(paragraph._element)
        
        if text:
            new_p = doc.add_paragraph()
            new_p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            runner = new_p.add_run(text)
            runner.font.name = 'Calibri'
            runner.font.size = Pt(12)
            parent.insert(index + 1, new_p._element)
            index += 1

        if images:
            for img in images:
                img_stream = io.BytesIO()
                img.save(img_stream, format='PNG')
                img_stream.seek(0)
                
                new_p = doc.add_paragraph()
                new_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = new_p.add_run()
                # Resize to reasonable width (e.g. 400pt)
                run.add_picture(img_stream, width=Pt(400))
                
                parent.insert(index + 1, new_p._element)
                index += 1

    # Iterate copy of paragraphs
    for para in list(doc.paragraphs):
        cleaned = para.text.strip()
        if cleaned in sections_map:
            add_content_after(para, sections_map[cleaned]["text"], sections_map[cleaned]["images"])
            
    bio = io.BytesIO()
    doc.save(bio)
    return bio

st.set_page_config(page_title="AI Lab Report Generator", layout="wide", page_icon="⚡")

# Moderní stylování pro lepší vzhled a přívětivost pro studenty
st.markdown("""
<style>
    /* Odstranění marginů a paddingů pro mobilní responzivitu */
    @media (max-width: 768px) {
        .block-container {
            padding-top: 1rem;
            padding-bottom: 1rem;
            padding-left: 1rem;
            padding-right: 1rem;
        }
    }

    /* Vylepšení nadpisů s responzivní velikostí písma (clamp) */
    h1, h2, h3 {
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
    }
    
    .main-title {
        text-align: center;
        font-size: clamp(2.5em, 6vw, 4em);
        font-weight: 800;
        background: -webkit-linear-gradient(45deg, #00f2fe, #4facfe, #00f2fe);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-size: 200% auto;
        margin-bottom: 0.1em;
        animation: shine 3s linear infinite;
    }
    
    @keyframes shine {
      to {
        background-position: 200% center;
      }
    }

    .sub-title {
        text-align: center;
        color: #a0aec0;
        font-size: clamp(1em, 3vw, 1.2em);
        margin-bottom: 2em;
    }
    
    /* Vylepšené Tlačítko pro generování - moderní glassmorphism & gradients */
    .stButton>button {
        background: linear-gradient(90deg, #00f2fe 0%, #4facfe 100%);
        color: #0e1117 !important;
        border-radius: 12px;
        border: none;
        padding: 14px 24px;
        font-weight: 800;
        font-size: 1.15em;
        transition: all 0.3s ease;
        width: 100%;
        box-shadow: 0 4px 15px rgba(0, 242, 254, 0.4);
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .stButton>button:hover {
        background: linear-gradient(90deg, #4facfe 0%, #00f2fe 100%);
        color: #ffffff !important;
        transform: translateY(-3px);
        box-shadow: 0 8px 25px rgba(0, 242, 254, 0.6);
    }
    
    /* Boxy pro nahrávání - jemnější okraje s hover efektem */
    div[data-testid="stFileUploader"] {
        background: rgba(30, 37, 48, 0.6);
        padding: 15px;
        border-radius: 12px;
        box-shadow: inset 0 2px 4px rgba(255, 255, 255, 0.05), 0 4px 10px rgba(0,0,0,0.2);
        border: 1px solid #2d3748;
        transition: all 0.3s ease-in-out;
        backdrop-filter: blur(10px);
    }
    div[data-testid="stFileUploader"]:hover {
        border-color: #00f2fe;
        box-shadow: inset 0 2px 4px rgba(255, 255, 255, 0.05), 0 4px 15px rgba(0, 242, 254, 0.15);
    }
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 class='main-title'>⚡ AI Generátor Protokolů</h1>", unsafe_allow_html=True)
st.markdown("<p class='sub-title'>Tvoje záchrana pro laborky na SPŠE. Rychle, moderně a bez stresu.</p>", unsafe_allow_html=True)

with st.expander("🔑 Nastavení & API (Google Gemini)", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        api_key = st.text_input("Google Gemini API Key", type="password", help="Získejte svůj klíč zdarma na https://aistudio.google.com/")
    with col2:
        model_options = {
            "Gemini 3 Flash (Preview)": "gemini-3-flash-preview",
            "Gemini 2.5 Flash (Stable)": "gemini-2.5-flash"
        }
        selected_label = st.radio(
            "Vyberte model AI:",
            options=list(model_options.keys()),
            index=1,
            help="Zobrazeny jsou pouze modely, které aktuálně fungují spolehlivě."
        )
        model_choice = model_options[selected_label]

    if not api_key:
        st.warning("⚠️ Nezapomeň zadat svůj API klíč pro pokračování.")

with st.form("lab_report_form"):
    st.markdown("### 📌 Základní informace")
    topic = st.text_input("Téma měření", placeholder="Např. Oživování a měření na stabilizovaném zdroji...", help="Téma, které se propíše do hlavičky protokolu.")
    is_handwritten = st.toggle("📝 Píšu tenhle elaborát ručně! (Zkrátí teorii a závěr na max. půl A4)", value=False, help="Zapni, abys nedostal do generování kilometry textu a nemusel jsi ho celý ručně přepisovat. AI Tě ušetří.")
    
    st.markdown("---")
    st.markdown("### 📂 Podklady pro AI (až 10 souborů na sekci!)")
    st.markdown("Můžeš nahrát tabulky, pdf zadání ze školy, fotky z mobilu z měření, cokoliv máš.")
    
    assignment_file = st.file_uploader("Zadání úlohy", type=['txt', 'docx', 'pdf', 'png', 'jpg', 'jpeg'], key="assignment", accept_multiple_files=True, help="Třeba fotka zadání (max 10 souborů)")
    
    instruments_file = st.file_uploader("Seznam použitých přístrojů", type=['txt', 'docx', 'pdf', 'xlsx', 'png', 'jpg', 'jpeg'], key="instruments", accept_multiple_files=True, help="Odkud AI vyčte přístroje (max 10 souborů)")
    
    data_files = st.file_uploader("Tabulky naměřených hodnot", type=['xlsx', 'csv', 'txt', 'pdf', 'png', 'jpg', 'jpeg'], key="data", accept_multiple_files=True, help="Hodně pomůže Excel nebo čitelná fotka hodnot (max 10 souborů)")

    theory_file = st.file_uploader("Podklady k teorii", type=['txt', 'docx', 'pdf', 'png', 'jpg', 'jpeg'], key="theory", accept_multiple_files=True, help="Třeba screenshoty skript nebo prezentace (max 10 souborů)")

    waveforms_file = st.file_uploader("Grafické průběhy k teorii (Novinka!)", type=['png', 'jpg', 'jpeg', 'pdf'], key="waveforms", accept_multiple_files=True, help="Nahraj fotky nebo PDF průběhů. AI je rozpozná a řádně vysvětlí v teoretické části! (max 10 souborů)")

    procedure_file = st.file_uploader("Pracovní postup", type=['txt', 'docx', 'pdf', 'png', 'jpg', 'jpeg'], key="procedure", accept_multiple_files=True, help="Materiál, z kterého AI přepíše postup do min. času (max 10 souborů)")

    conclusion_file = st.file_uploader("Osnova pro závěr", type=['txt', 'docx', 'pdf', 'png', 'jpg', 'jpeg'], key="conclusion", accept_multiple_files=True, help="Zadání toho, co musí být v závěru uvedeno (max 10 souborů)")

    schema_files = st.file_uploader("Schéma zapojení pro protokol", type=['png', 'jpg', 'jpeg'], key="schema", accept_multiple_files=True, help="Obrázky, které se vloží do sekce Schéma zapojení (max 10 souborů)")

    st.markdown("<br>", unsafe_allow_html=True)
    submitted = st.form_submit_button("🚀 Vygenerovat nadupaný protokol")

if submitted:
    # Kontrola maximálně 10 souborů na sekci
    files_groups = [assignment_file, instruments_file, data_files, theory_file, procedure_file, conclusion_file, schema_files]
    over_limit = any(f is not None and len(f) > 10 for f in files_groups if isinstance(f, list))
    
    if over_limit:
        st.error("❌ Nahrál jsi pod jednu sekci víc než 10 souborů! AI to nezvládne zpracovat. Vrať se a limituj je na max 10.")
    elif not api_key:
        st.error("❌ Nejdřív doplň svůj Google Gemini API klíč!")
    elif not topic:
        st.error("❌ Musíš uvést Téma měření!")
    else:
        with st.spinner("Zpracovávám soubory a generuji protokol..."):
            
            # Extract content from files
            # Note: extract_content_from_files handles lists now
            assignment_text, assignment_images = extract_content_from_files(assignment_file)
            instruments_text, instruments_images = extract_content_from_files(instruments_file)
            data_text, data_images = extract_content_from_files(data_files)
            theory_text, theory_images = extract_content_from_files(theory_file)
            waveforms_text, waveforms_images = extract_content_from_files(waveforms_file)
            procedure_text, procedure_images = extract_content_from_files(procedure_file)
            conclusion_text, conclusion_images = extract_content_from_files(conclusion_file)
            
            # Extract schema images specially
            _, schema_images_list = extract_content_from_files(schema_files)

            # If no file uploaded, handle empty text
            if not assignment_text and not assignment_images: assignment_text = ""
            if not instruments_text and not instruments_images: instruments_text = ""
            if not data_text and not data_images: data_text = ""

            # Prepare Input Map for AI
            inputs_map = {
                'assignment_text': assignment_text,
                'instruments_text': instruments_text,
                'data_text': data_text,
                'theory_text': theory_text,
                'waveforms_text': waveforms_text,
                'procedure_text': procedure_text,
                'conclusion_text': conclusion_text,
                'schema_images': schema_images_list,
                'data_images': data_images,
                'images_lists': [
                    assignment_images, instruments_images, data_images, 
                    theory_images, waveforms_images, procedure_images, conclusion_images
                ]
            }

            # Generate Logic
            ai_data = generate_lab_report(api_key, model_choice, topic, inputs_map, is_handwritten)
            
            if ai_data:
                st.balloons()
                st.success("🎉 Úspěšně vygenerováno! Tady je tvůj základ záchrany.")
                
                # Preview
                st.subheader("Náhled obsahu:")
                with st.expander("Teorie", expanded=False):
                    st.markdown(ai_data.get('teorie', ''))
                with st.expander("Postup měření (náhled)", expanded=True):
                    st.markdown(ai_data.get('postup', ''))
                with st.expander("Příklad výpočtu", expanded=True):
                    st.markdown(ai_data.get('priklad_vypoctu', ''))
                with st.expander("Závěr", expanded=False):
                    st.markdown(ai_data.get('zaver', ''))
                
                # Generate DOCX from TEMPLATE
                template_path = "Graficka_Osnova.docx"
                docx_file = fill_template_docx(
                    template_path,
                    topic, 
                    inputs_map,
                    ai_data
                )
                
                st.download_button(
                    label="📥 Stáhnout protokol (.docx)",
                    data=docx_file.getvalue(),
                    file_name="laboratorni_protokol.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
