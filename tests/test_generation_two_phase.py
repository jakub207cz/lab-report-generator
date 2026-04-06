from __future__ import annotations

import json

from pipeline.generation import generate_lab_report


class _FakeResponse:
    def __init__(self, text: str) -> None:
        self.text = text


class _FakeModel:
    def __init__(self, outputs: list[str]) -> None:
        self._outputs = outputs
        self.calls: list[list[object]] = []

    def generate_content(self, content_parts: list[object]) -> _FakeResponse:
        self.calls.append(content_parts)
        return _FakeResponse(self._outputs.pop(0))


def test_generate_lab_report_two_phase_calls(monkeypatch):
    first_payload = {
        "teorie": "Teorie z 1. callu",
        "postup": "Postup z 1. callu",
        "priklad_vypoctu": "Výpočet z 1. callu",
        "image_references": ["IMG-001"],
    }
    second_payload = {
        "zaver": "Závěr z 2. callu",
        "image_references": ["IMG-002"],
    }

    fake_model = _FakeModel(
        outputs=[
            json.dumps(first_payload, ensure_ascii=False),
            json.dumps(second_payload, ensure_ascii=False),
            "Připojil jsem obvod a změřil jsem hodnoty.",
        ]
    )

    monkeypatch.setattr("pipeline.generation.genai.configure", lambda api_key: None)
    monkeypatch.setattr("pipeline.generation.genai.GenerativeModel", lambda _: fake_model)

    report = generate_lab_report(
        api_key="dummy",
        model_name="dummy-model",
        topic="Testovací téma",
        inputs_map={
            "assignment_text": "Zadání",
            "assignment_theory_text": "Extrahovaný teoretický úvod",
            "assignment_conclusion_text": "Extrahovaný závěr",
            "theory_text": "Uživatelská teorie",
            "procedure_text": "Zapojte obvod.",
            "data_text": "Naměřené hodnoty",
            "conclusion_text": "Osnova závěru",
            "waveforms_text": "Grafické průběhy",
            "image_catalog_text": "- IMG-001: a.png\n- IMG-002: b.png",
            "images_lists": [],
            "data_images": [],
            "waveforms_images": [],
        },
        is_handwritten=False,
    )

    assert len(fake_model.calls) == 3
    assert "POUZE sekce: TEORIE, POSTUP MĚŘENÍ a PŘÍKLAD VÝPOČTU" in fake_model.calls[0][0]
    assert "POUZE sekci ZÁVĚR" in fake_model.calls[1][0]
    assert "Přepiš následující pracovní postup" in fake_model.calls[2][0]

    assert report.teorie == "Teorie z 1. callu"
    assert report.postup == "Připojil jsem obvod a změřil jsem hodnoty."
    assert report.priklad_vypoctu == "Výpočet z 1. callu"
    assert report.zaver == "Závěr z 2. callu"
    assert report.image_references == ["IMG-001", "IMG-002"]
