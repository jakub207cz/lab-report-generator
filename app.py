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
    font.name = 'Times New Roman'
    font.size = Pt(12)
    
    # Also set for parsed styles if needed, but Normal covers most
    for style in doc.styles:
        if hasattr(style, 'font'):
            style.font.name = 'Times New Roman'


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

def generate_lab_report(api_key, model_name, topic, inputs_map):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)

    system_prompt = """
    Jsi student 3. ročníku SPŠE (Střední průmyslová škola elektrotechnická). 
    Tvým úkolem je napsat školní laboratorní protokol (elaborát).
    
    Téma: {topic}

    POSTUPUJ PODLE TĚCHTO SEKCI:

    1. TEORIE (Rozsah: cca 1.5 strany A4, min. 1000 slov)
       - Vycházej z přiloženého textu/osnovy:
       {theory_text}
       - Pokud není nic přiloženo, napiš velmi podrobnou teorii k tématu. Vysvětli fyzikální principy, vzorce, odvození a souvislosti.
       - Text musí být odborný a vyčerpávající.

    2. POSTUP MĚŘENÍ
       - PŘEPIŠ přiložený text pracovního postupu do 1. OSOBY MINULÉHO ČASU (např. změň "Změřte napětí" na "Změřil jsem napětí").
       - Zdrojový text postupu:
       {procedure_text}
       - Pokud jsou přiloženy obrázky, odkazuj se na ně textem (např. "jak je vidět na obrázku 1").

    3. PŘÍKLAD VÝPOČTU
       - Na základě naměřených dat ({data_text}) a teorie vytvoř JEDEN KONKRÉTNÍ PŘÍKLAD výpočtu.
       - Uveď vzorec, dosaď konkrétní naměřené hodnoty (např. U=10V, I=2A) a vypočítej výsledek.
       - Výpočet musí být fyzikálně správný.

    4. ZÁVĚR (Fakta a Analýza)
       - Vycházej POUZE z naměřených hodnot ({data_text}) a přiložené osnovy:
       {conclusion_text}
       - Zhodnoť měření technicky a kriticky.
       - CITUJ KONKRÉTNÍ HODNOTY z naměřených dat. Např. "Naměřil jsem napětí 5.2 V, což odpovídá..."
       - NEVYMÝŠLEJ SI ŽÁDNÉ HODNOTY. Pokud data chybí, napiš, že nebylo možné vyhodnotit.
       - Porovnej naměřené hodnoty s teoretickými předpoklady.

    DALŠÍ VSTUPY:
    --- ZADÁNÍ ---
    {assignment_text}
    
    --- POUŽITÉ PŘÍSTROJE ---
    {instruments_text}
    
    --- NAMĚŘENÉ HODNOTY ---
    {data_text}

    DŮLEŽITÉ: Rovnice piš jako prostý text (R=U/I).
    
    Vygeneruj výstup POUZE jako validní JSON s následující strukturou:
    {{
        "teorie": "...",
        "postup": "...",
        "priklad_vypoctu": "...",
        "zaver": "..."
    }}
    """

    formatted_prompt = system_prompt.format(
        topic=topic,
        theory_text=inputs_map.get('theory_text', ''),
        procedure_text=inputs_map.get('procedure_text', ''),
        conclusion_text=inputs_map.get('conclusion_text', ''),
        assignment_text=inputs_map.get('assignment_text', ''),
        instruments_text=inputs_map.get('instruments_text', ''),
        data_text=inputs_map.get('data_text', '')
    )
    
    content_parts = [formatted_prompt]
    
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
            new_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            runner = new_p.add_run(text)
            runner.font.name = 'Times New Roman'
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

st.set_page_config(page_title="AI Lab Report Generator", layout="centered")

st.title("⚡ Generátor Laboratorních Protokolů (SPŠE)")

with st.expander("ℹ️ Nastavení & API", expanded=True):
    col1, col2 = st.columns(2)
    with col1:
        api_key = st.text_input("Google Gemini API Key", type="password", help="Získejte svůj klíč na https://aistudio.google.com/")
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
        st.warning("Zadejte svůj API klíč pro pokračování.")

with st.form("lab_report_form"):
    # 1. Téma
    topic = st.text_input("Téma měření", placeholder="Např. Měření zatěžovací charakteristiky zdroje")
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("### 📝 Zadání")
        assignment_file = st.file_uploader("Nahrát zadání (Text/Word/PDF/Img)", type=['txt', 'docx', 'pdf', 'png', 'jpg', 'jpeg'], key="assignment", accept_multiple_files=False)
        
        st.markdown("### 🔌 Přístroje")
        instruments_file = st.file_uploader("Seznam přístrojů (Text/Word/PDF/Img)", type=['txt', 'docx', 'pdf', 'xlsx', 'png', 'jpg', 'jpeg'], key="instruments", accept_multiple_files=False)
        
        st.markdown("### 📊 Naměřená data")
        data_files = st.file_uploader("Tabulka hodnot (Excel/CSV/Text/PDF/Img) - Možno více", type=['xlsx', 'csv', 'txt', 'pdf', 'png', 'jpg', 'jpeg'], key="data", accept_multiple_files=True)

    with col_b:
        st.markdown("### 📖 Teorie (Osnova)")
        theory_file = st.file_uploader("Podklady k teorii (Text/Word/PDF/Img)", type=['txt', 'docx', 'pdf', 'png', 'jpg', 'jpeg'], key="theory", accept_multiple_files=False)

        st.markdown("### 🔧 Postup (Pro přepis)")
        procedure_file = st.file_uploader("Pracovní postup (Text/Word/PDF/Img)", type=['txt', 'docx', 'pdf', 'png', 'jpg', 'jpeg'], key="procedure", accept_multiple_files=False)

        st.markdown("### 🏁 Závěr (Osnova)")
        conclusion_file = st.file_uploader("Osnova závěru (Text/Word/PDF/Img)", type=['txt', 'docx', 'pdf', 'png', 'jpg', 'jpeg'], key="conclusion", accept_multiple_files=False)

    st.markdown("### 🖼️ Schéma zapojení (Možno více)")
    schema_files = st.file_uploader("Obrázek schématu (PNG/JPG)", type=['png', 'jpg', 'jpeg'], key="schema", accept_multiple_files=True)

    submitted = st.form_submit_button("Generovat protokol")

if submitted:
    if not api_key:
        st.error("Chybí API klíč!")
    elif not topic:
        st.error("Vyplňte prosím Téma měření.")
    else:
        with st.spinner("Zpracovávám soubory a generuji protokol..."):
            
            # Extract content from files
            # Note: extract_content_from_files handles lists now
            assignment_text, assignment_images = extract_content_from_files(assignment_file)
            instruments_text, instruments_images = extract_content_from_files(instruments_file)
            data_text, data_images = extract_content_from_files(data_files)
            theory_text, theory_images = extract_content_from_files(theory_file)
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
                'procedure_text': procedure_text,
                'conclusion_text': conclusion_text,
                'schema_images': schema_images_list,
                'data_images': data_images,
                'images_lists': [
                    assignment_images, instruments_images, data_images, 
                    theory_images, procedure_images, conclusion_images
                ]
            }

            # Generate Logic
            ai_data = generate_lab_report(api_key, model_choice, topic, inputs_map)
            
            if ai_data:
                st.success("Generování dokončeno!")
                
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
