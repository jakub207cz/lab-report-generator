# Proč Endora nestačí a jak nasadit aplikaci ZDARMA na Streamlit Cloud

**Krátká odpověď:** 
Endora (a většina klasických free hostingů) je určena pro **PHP a statické stránky**. Streamlit aplikace je ale běžící program v Pythonu, který potřebuje neustále běžet na serveru. Na Endoře by to technicky **nefungovalo** nebo by to bylo extrémně složité a pomalé.

**Nejlepší řešení:**
Použijte **Streamlit Community Cloud**. Je to oficiální hosting přímo od tvůrců Streamlit, je **zcela zdarma**, velmi rychlý a určený přesně pro tyto aplikace.

---

## 🚀 Postup nasazení (krok za krokem)

### 1. Příprava souborů
Ujistěte se, že máte ve složce tyto soubory:
- `app.py` (váš hlavní kód)
- `requirements.txt` (seznam knihoven, aktualizoval jsem ho)
- `Graficka_Osnova.docx` (vaše šablona)

### 2. Nahrání na GitHub (Nutnost)
Streamlit Cloud si bere kód z GitHubu.
1.  Jděte na [GitHub.com](https://github.com/) a přihlašte se (nebo zaregistrujte).
2.  Vytvořte **New Repository** (např. `lab-report-generator`).
3.  Nahrajte do něj výše zmíněné soubory (můžete použít tlačítko "Upload files" na webu GitHubu nebo Git příkazy).

### 3. Propojení se Streamlit Cloud
1.  Jděte na [share.streamlit.io](https://share.streamlit.io/).
2.  Přihlašte se přes GitHub.
3.  Klikněte na **"New app"**.
4.  Vyberte váš repozitář (`lab-report-generator`).
5.  Branch: `main` (nebo `master`).
6.  Main file path: `app.py`.
7.  Klikněte na **"Deploy!"**.

### 4. Nastavení API Klíče (Tajné!)
Aby aplikace fungovala a nikdo neukradl váš klíč:
1.  Na stránce vaší běžící aplikace klikněte vpravo dole na **"Manage app"** (nebo v nastavení aplikace na dashboardu).
2.  Jděte do sekce **"Settings"** -> **"Secrets"**.
3.  Do pole vložte tento text (s vaším klíčem):

```toml
GOOGLE_API_KEY = "VÁŠ_SKUTEČNÝ_GOOGLE_API_KLÍČ"
```

4.  Uložte. Aplikace se restartuje a bude fungovat!

---

💡 **Tip:** Pokud chcete, aby aplikace brala klíč automaticky ze Secrets místo inputu, můžeme `app.py` lehce upravit, aby se podíval, jestli klíč existuje v `st.secrets`. Chcete to?
