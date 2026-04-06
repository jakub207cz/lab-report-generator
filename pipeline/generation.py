from __future__ import annotations

import json
import re
from typing import Any, Dict, List

import google.generativeai as genai
from pydantic import BaseModel, Field

from pipeline.schemas import LabReportData


def _strip_markdown_fences(raw_text: str) -> str:
    text_response = raw_text.strip()
    if "```json" in text_response:
        text_response = text_response.split("```json", 1)[1].split("```", 1)[0]
    elif "```" in text_response:
        text_response = text_response.split("```", 1)[1].split("```", 1)[0]
    return text_response.strip()


def _extract_json_object(raw_text: str) -> str:
    text = _strip_markdown_fences(raw_text)
    start = text.find("{")
    end = text.rfind("}")
    if start >= 0 and end > start:
        return text[start : end + 1]
    return text


class _TheoryProcedureData(BaseModel):
    teorie: str = Field(default="")
    postup: str = Field(default="")
    priklad_vypoctu: str = Field(default="")
    image_references: List[str] = Field(default_factory=list)


class _ConclusionData(BaseModel):
    zaver: str = Field(default="")
    image_references: List[str] = Field(default_factory=list)


def _generate_structured_part(model: genai.GenerativeModel, prompt: str, schema_model: type[BaseModel]) -> BaseModel:
    response = model.generate_content([prompt])
    raw_json_text = _extract_json_object(response.text)

    try:
        parsed = json.loads(raw_json_text)
        return schema_model.model_validate(parsed)
    except Exception:
        repair_prompt = (
            "Uprav následující text na VALIDNÍ JSON přesně dle schema. "
            "Vrať pouze JSON bez komentářů a bez markdownu.\n\n"
            f"SCHEMA:\n{json.dumps(schema_model.model_json_schema(), ensure_ascii=False)}\n\n"
            f"TEXT:\n{response.text}"
        )
        repair_response = model.generate_content([repair_prompt])
        repaired_json_text = _extract_json_object(repair_response.text)
        parsed = json.loads(repaired_json_text)
        return schema_model.model_validate(parsed)


def _build_theory_and_procedure_prompt(topic: str, inputs_map: Dict[str, Any], is_handwritten: bool) -> str:
    theory_length = (
        "ZKRÁCENÝ ROZSAH: Maximálně půl strany A4! Piš velmi stručně a jen to nejdůležitější, protože to bude student přepisovat ručně."
        if is_handwritten
        else "Rozsah: cca 1.5 strany A4, min. 1000 slov"
    )

    image_catalog = inputs_map.get("image_catalog_text", "")

    return f"""
Jsi student 3. ročníku SPŠE (Střední průmyslová škola elektrotechnická).
Tvým úkolem je napsat část školního laboratorního protokolu (elaborátu).

Téma: {topic}

V tomto volání generuj POUZE sekce: TEORIE, POSTUP MĚŘENÍ a PŘÍKLAD VÝPOČTU.

DŮLEŽITÉ PRAVIDLO PRO CHYBĚJÍCÍ ZDROJE (TEORIE / POSTUP):
Pokud u sekce Teorie nebo Postup zjistíš, že nebyly poskytnuty ŽÁDNÉ podklady (žádný text ani relevantní nápověda), sekci samostatně vygeneruj podle nejlepších znalostí k danému tématu. Na úplný začátek takto dovygenerované sekce však MUSÍŠ přidat přesně tuto větu velkými písmeny:
\"NEBYL PŘILOŽEN ZDROJ INFORMACÍ, AI TI TO VYGENEROVALA. ZKONTROLUJ SI TO!\\n\\n\"

POSTUPUJ PODLE TĚCHTO SEKCÍ:

1. TEORIE ({theory_length})
   - Vycházej z extrahovaného textu sekce "Teoretický úvod" ze zadání:
   {inputs_map.get('assignment_theory_text', '')}
   - Dále využij přiložený text/osnovu od uživatele (kolonka Teorie):
   {inputs_map.get('theory_text', '')}
   - ODDĚLENĚ NAHRANÉ GRAFICKÉ PRŮBĚHY:
   {inputs_map.get('waveforms_text', '')}
   - Pokud není nic přiloženo, sekci kompletně vygeneruj a nezapomeň na povinnou větu.
   - Text musí být odborný a vyčerpávající.

2. POSTUP MĚŘENÍ
   - Vycházej ze zdrojového textu postupu. Primárně jej přepiš do 1. osoby jednotného čísla a minulého času.
   - Přepisuj pouze to, co je ve zdrojovém postupu. Nevymýšlej nové kroky, neměň pořadí kroků ani technické hodnoty.
   - Styl musí být osobní: např. "připojil jsem", "nastavil jsem", "změřil jsem", "vypočítal jsem".
   - Zdrojový text postupu:
   {inputs_map.get('procedure_text', '')}

3. PŘÍKLAD VÝPOČTU
   - Na základě naměřených dat ({inputs_map.get('data_text', '')}) a teorie vytvoř více než jeden příklad výpočtu.

DALŠÍ VSTUPY:
--- ZADÁNÍ ---
{inputs_map.get('assignment_text', '')}

--- POUŽITÉ PŘÍSTROJE ---
{inputs_map.get('instruments_text', '')}

--- DOSTUPNÉ OBRÁZKY (image_id) ---
{image_catalog}

DŮLEŽITÉ:
- Rovnice piš jako prostý text (R=U/I).
- U desetinných čísel používej desetinnou čárku.
- Pokud použiješ obrázek v textu, uveď jeho ID do pole image_references.

Vrať POUZE validní JSON dle tohoto schema (a vyplň jen relevantní pole):
{json.dumps(_TheoryProcedureData.model_json_schema(), ensure_ascii=False)}
"""


def _build_conclusion_prompt(topic: str, inputs_map: Dict[str, Any], is_handwritten: bool) -> str:
    conclusion_length = (
        "ZKRÁCENÝ ROZSAH: Maximálně půl strany A4! Stručné zhodnocení."
        if is_handwritten
        else "Fakta a detailní analýza"
    )
    image_catalog = inputs_map.get("image_catalog_text", "")

    return f"""
Jsi student 3. ročníku SPŠE (Střední průmyslová škola elektrotechnická).
Tvým úkolem je napsat POUZE sekci ZÁVĚR laboratorního protokolu.

Téma: {topic}

1. ZÁVĚR ({conclusion_length})
   - Primárně vycházej z extrahovaného textu sekce "Závěr" ze zadání:
   {inputs_map.get('assignment_conclusion_text', '')}
   - Dále využij uživatelskou osnovu/poznámky pro závěr:
   {inputs_map.get('conclusion_text', '')}
   - Naměřená data:
   {inputs_map.get('data_text', '')}
   - Volitelně zohledni i grafické průběhy (pokud jsou k dispozici):
   {inputs_map.get('waveforms_text', '')}

2. PRAVIDLO PRO CHYBĚJÍCÍ ZDROJE ZÁVĚRU
   - Pokud není k závěru žádný podklad, závěr samostatně vygeneruj podle tématu.
   - Na začátek této sekce musíš přidat přesně:
   \"NEBYL PŘILOŽEN ZDROJ INFORMACÍ, AI TI TO VYGENEROVALA. ZKONTROLUJ SI TO!\\n\\n\"

DALŠÍ VSTUPY:
--- ZADÁNÍ ---
{inputs_map.get('assignment_text', '')}

--- DOSTUPNÉ OBRÁZKY (image_id) ---
{image_catalog}

DŮLEŽITÉ:
- Rovnice piš jako prostý text (R=U/I).
- U desetinných čísel používej desetinnou čárku.
- Pokud použiješ obrázek v textu, uveď jeho ID do pole image_references.

Vrať POUZE validní JSON dle tohoto schema:
{json.dumps(_ConclusionData.model_json_schema(), ensure_ascii=False)}
"""


def generate_lab_report(
    api_key: str,
    model_name: str,
    topic: str,
    inputs_map: Dict[str, Any],
    is_handwritten: bool = False,
) -> LabReportData:
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel(model_name)

    theory_prompt = _build_theory_and_procedure_prompt(topic=topic, inputs_map=inputs_map, is_handwritten=is_handwritten)
    theory_parts = [theory_prompt]
    for img_list in inputs_map.get("images_lists", []):
        theory_parts.extend(img_list)

    theory_data = _generate_structured_part(
        model=model,
        prompt="\n".join([str(part) for part in theory_parts if isinstance(part, str)]) if all(isinstance(part, str) for part in theory_parts) else theory_prompt,
        schema_model=_TheoryProcedureData,
    )

    conclusion_prompt = _build_conclusion_prompt(topic=topic, inputs_map=inputs_map, is_handwritten=is_handwritten)
    conclusion_parts: List[Any] = [conclusion_prompt]
    conclusion_parts.extend(inputs_map.get("data_images", []))
    conclusion_parts.extend(inputs_map.get("waveforms_images", []))

    # Pokud je seznam jen text, sloučíme ho. Jinak pošleme multimodálně (text + obrázky).
    if all(isinstance(part, str) for part in conclusion_parts):
        conclusion_data = _generate_structured_part(
            model=model,
            prompt="\n".join([str(part) for part in conclusion_parts]),
            schema_model=_ConclusionData,
        )
    else:
        response = model.generate_content(conclusion_parts)
        raw_json_text = _extract_json_object(response.text)
        try:
            parsed = json.loads(raw_json_text)
            conclusion_data = _ConclusionData.model_validate(parsed)
        except Exception:
            repair_prompt = (
                "Uprav následující text na VALIDNÍ JSON přesně dle schema. "
                "Vrať pouze JSON bez komentářů a bez markdownu.\n\n"
                f"SCHEMA:\n{json.dumps(_ConclusionData.model_json_schema(), ensure_ascii=False)}\n\n"
                f"TEXT:\n{response.text}"
            )
            repair_response = model.generate_content([repair_prompt])
            repaired_json_text = _extract_json_object(repair_response.text)
            parsed = json.loads(repaired_json_text)
            conclusion_data = _ConclusionData.model_validate(parsed)

    source_procedure = (inputs_map.get("procedure_text", "") or "").strip()

    report = LabReportData(
        teorie=theory_data.teorie,
        postup=theory_data.postup,
        priklad_vypoctu=theory_data.priklad_vypoctu,
        zaver=conclusion_data.zaver,
        image_references=list(dict.fromkeys((theory_data.image_references or []) + (conclusion_data.image_references or []))),
    )

    if source_procedure:
        rewrite_prompt = f"""
Přepiš následující pracovní postup do 1. osoby jednotného čísla a minulého času v češtině.

PRAVIDLA:
- Zachovej věcný obsah, pořadí kroků, názvy veličin a všechny číselné hodnoty.
- Nic nového nevymýšlej ani nepřidávej.
- Piš přirozeně stylem "já jsem ..." / "změřil jsem ...".
- Vrať pouze čistý text bez markdownu, bez nadpisu a bez JSONu.

ZDROJOVÝ POSTUP:
{source_procedure}
"""
        try:
            rewritten_procedure = _strip_markdown_fences(model.generate_content([rewrite_prompt]).text).strip()
            if rewritten_procedure:
                report.postup = rewritten_procedure
        except Exception:
            # Když přepis selže, použije se postup z hlavního generování.
            pass

    return report
