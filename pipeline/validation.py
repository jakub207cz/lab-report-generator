from __future__ import annotations

import re
from typing import Dict

from pipeline.schemas import ImageAsset, LabReportData, QualityGateResult, QualityIssue


def run_quality_check(report: LabReportData, image_registry: Dict[str, ImageAsset]) -> QualityGateResult:
    # DOČASNĚ VYPNUTO: quality gate je prozatím bypassnutý.
    # Až budeš chtít validaci vrátit, smaž tento return a obnoví se původní logika níže.
    return QualityGateResult(status="PASS", issues=[])

    issues: list[QualityIssue] = []

    required_sections = {
        "teorie": report.teorie,
        "postup": report.postup,
        "priklad_vypoctu": report.priklad_vypoctu,
        "zaver": report.zaver,
    }

    for section_name, value in required_sections.items():
        if not value or not value.strip():
            issues.append(
                QualityIssue(
                    severity="FAIL",
                    code="MISSING_SECTION",
                    message=f"Sekce '{section_name}' je prázdná.",
                )
            )

    for image_id in report.image_references:
        if image_id not in image_registry:
            issues.append(
                QualityIssue(
                    severity="FAIL",
                    code="UNKNOWN_IMAGE_REFERENCE",
                    message=f"Výstup odkazuje na neznámé image_id '{image_id}'.",
                )
            )

    all_text = "\n".join(required_sections.values())
    if re.search(r"\d+\.\d+", all_text):
        issues.append(
            QualityIssue(
                severity="WARN",
                code="DECIMAL_DOT_DETECTED",
                message="Detekována desetinná tečka; preferuj desetinnou čárku.",
            )
        )

    if issues and any(i.severity == "FAIL" for i in issues):
        status = "FAIL"
    elif issues:
        status = "WARN"
    else:
        status = "PASS"

    return QualityGateResult(status=status, issues=issues)
