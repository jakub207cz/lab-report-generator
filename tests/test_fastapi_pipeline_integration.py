from __future__ import annotations

from PIL import Image
from fastapi.testclient import TestClient

import fastapi_server
from pipeline.schemas import ImageAsset, LabReportData, QualityGateResult


def test_generate_endpoint_uses_pipeline_and_returns_docx(monkeypatch) -> None:
    calls = {"generation": 0, "quality": 0}

    def fake_generate(api_key, model_name, topic, inputs_map, is_handwritten):
        calls["generation"] += 1
        assert api_key == "test-key"
        assert topic == "Test tema"
        assert isinstance(inputs_map, dict)
        return LabReportData(
            teorie="Test teorie",
            postup="Test postup",
            priklad_vypoctu="Test vypocet",
            zaver="Test zaver",
            image_references=[],
        )

    def fake_quality(report, image_registry):
        calls["quality"] += 1
        assert report.teorie == "Test teorie"
        assert isinstance(image_registry, dict)
        return QualityGateResult(status="PASS", issues=[])

    monkeypatch.setattr(fastapi_server, "generate_lab_report_structured", fake_generate)
    monkeypatch.setattr(fastapi_server, "run_quality_check", fake_quality)

    client = TestClient(fastapi_server.app)
    response = client.post(
        "/api/generate",
        data={
            "topic": "Test tema",
            "username": "Tester",
            "is_handwritten": "false",
            "model_name": "gemini-2.5-flash",
            "api_key": "test-key",
        },
        files=[
            (
                "assignment_files",
                ("zadani.txt", b"Text zadani", "text/plain"),
            )
        ],
    )

    assert response.status_code == 200
    assert response.headers["x-quality-status"] == "PASS"
    assert (
        response.headers["content-type"]
        == "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
    assert response.content.startswith(b"PK")
    assert calls["generation"] == 1
    assert calls["quality"] == 1


def test_generate_endpoint_uses_assignment_fallback_images(monkeypatch) -> None:
    calls = {"generation": 0}

    async def fake_assignment_fallback(assignment_files, image_counter_start):
        schema_img = Image.new("RGB", (16, 16), color="white")
        wave_img = Image.new("RGB", (16, 16), color="black")
        return (
            [schema_img],
            [ImageAsset(image_id="IMG-900", filename="schema-fallback.png", section="schema")],
            [wave_img],
            [ImageAsset(image_id="IMG-901", filename="wave-fallback.png", section="waveforms")],
            image_counter_start,
        )

    def fake_generate(api_key, model_name, topic, inputs_map, is_handwritten):
        calls["generation"] += 1
        assert len(inputs_map["schema_images"]) == 1
        assert len(inputs_map["waveforms_images"]) == 1
        assert inputs_map["schema_image_ids"] == ["IMG-900"]
        assert inputs_map["waveforms_image_ids"] == ["IMG-901"]
        return LabReportData(
            teorie="Test teorie",
            postup="Test postup",
            priklad_vypoctu="Test vypocet",
            zaver="Test zaver",
            image_references=[],
        )

    def fake_quality(report, image_registry):
        return QualityGateResult(status="PASS", issues=[])

    monkeypatch.setattr(fastapi_server, "extract_assignment_fallback_images", fake_assignment_fallback)
    monkeypatch.setattr(fastapi_server, "generate_lab_report_structured", fake_generate)
    monkeypatch.setattr(fastapi_server, "run_quality_check", fake_quality)

    client = TestClient(fastapi_server.app)
    response = client.post(
        "/api/generate",
        data={
            "topic": "Test fallback",
            "username": "Tester",
            "is_handwritten": "false",
            "model_name": "gemini-2.5-flash",
            "api_key": "test-key",
        },
        files=[
            (
                "assignment_files",
                ("zadani.txt", b"Text zadani", "text/plain"),
            )
        ],
    )

    assert response.status_code == 200
    assert calls["generation"] == 1


def test_generate_endpoint_reroutes_xlsx_chart_images_to_waveforms(monkeypatch) -> None:
    calls = {"generation": 0}

    async def fake_assignment_fallback(assignment_files, image_counter_start):
        return ([], [], [], [], image_counter_start)

    async def fake_assignment_sections(_assignment_files):
        return {}

    async def fake_extract_content(files, section_name, image_counter_start):
        if section_name == "data":
            img = Image.new("RGB", (16, 16), color="blue")
            asset = ImageAsset(image_id="IMG-123", filename="xlsx-chart-openpyxl-001.png", section="data")
            return (
                "[Graf z XLSX IMG-123: xlsx-chart-openpyxl-001.png]",
                [img],
                [asset],
                [],
                image_counter_start,
            )
        return "", [], [], [], image_counter_start

    def fake_generate(api_key, model_name, topic, inputs_map, is_handwritten):
        calls["generation"] += 1
        assert inputs_map["data_images"] == []
        assert len(inputs_map["waveforms_images"]) == 1
        assert inputs_map["waveforms_image_ids"] == ["IMG-123"]
        assert inputs_map["data_image_ids"] == []
        return LabReportData(
            teorie="Test teorie",
            postup="Test postup",
            priklad_vypoctu="Test vypocet",
            zaver="Test zaver",
            image_references=[],
        )

    def fake_quality(report, image_registry):
        return QualityGateResult(status="PASS", issues=[])

    monkeypatch.setattr(fastapi_server, "extract_assignment_fallback_images", fake_assignment_fallback)
    monkeypatch.setattr(fastapi_server, "extract_assignment_sections_from_files", fake_assignment_sections)
    monkeypatch.setattr(fastapi_server, "extract_content_from_uploadfiles", fake_extract_content)
    monkeypatch.setattr(fastapi_server, "generate_lab_report_structured", fake_generate)
    monkeypatch.setattr(fastapi_server, "run_quality_check", fake_quality)

    client = TestClient(fastapi_server.app)
    response = client.post(
        "/api/generate",
        data={
            "topic": "Test reroute",
            "username": "Tester",
            "is_handwritten": "false",
            "model_name": "gemini-2.5-flash",
            "api_key": "test-key",
        },
    )

    assert response.status_code == 200
    assert calls["generation"] == 1
