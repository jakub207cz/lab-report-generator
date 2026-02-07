# ⚡ AI Generátor Laboratorních Protokolů

Aplikace pro SPŠE (Střední průmyslová škola elektrotechnická), která automatizuje tvorbu laboratorních protokolů. Na základě nahraného zadání, naměřených dat, teorie a obrázků vygeneruje kompletní protokol ve formátu `.docx`.

## ✨ Funkce
- **Pokročilé vstupy:** Podporuje nahrávání Word, PDF, Textu, Excelu i Obrázků.
- **AI Analýza:** Vytváří Teoretický úvod, Přepisuje postup do min. času, Analyzuje naměřená data a tvoří Závěr.
- **Příklad Výpočtu:** Automaticky navrhne a spočítá jeden vzorový příklad na základě dat.
- **Word Šablona:** Vše se vkládá do připravené šablony `Graficka_Osnova.docx`.

## 🚀 Jak spustit lokálně

1.  Nainstalujte Python.
2.  Nainstalujte závislosti:
    ```bash
    pip install -r requirements.txt
    ```
3.  Spusťte aplikaci:
    ```bash
    streamlit run app.py
    ```

## 🔑 API Klíč
Aplikace využívá **Google Gemini AI**.
Pro fungování musíte do aplikace vložit svůj **Gemini API Key**.
- Získejte ho zdarma zde: [Google AI Studio](https://aistudio.google.com/)

## ☁️ Jak nasadit na internet
Tato aplikace je ideální pro **Streamlit Community Cloud**.
Podrobný návod na nasazení najdete v souboru [JAK_NASADIT.md](JAK_NASADIT.md).
