from __future__ import annotations

from typing import Any, Dict, List, Literal, Optional

from pydantic import BaseModel, Field


class ImageAsset(BaseModel):
    image_id: str
    filename: str
    section: str


class LabReportData(BaseModel):
    teorie: str = Field(default="")
    postup: str = Field(default="")
    priklad_vypoctu: str = Field(default="")
    zaver: str = Field(default="")
    image_references: List[str] = Field(default_factory=list)


class QualityIssue(BaseModel):
    severity: Literal["WARN", "FAIL"]
    code: str
    message: str


class QualityGateResult(BaseModel):
    status: Literal["PASS", "WARN", "FAIL"]
    issues: List[QualityIssue] = Field(default_factory=list)


class DocumentChunk(BaseModel):
    text: str
    type: Literal["paragraph", "table", "figure", "heading"]
    source_file: str
    page: Optional[int] = None
    section_hint: Optional[str] = None
    confidence: float = 1.0


class TableData(BaseModel):
    headers: List[str]
    rows: List[List[Any]]
    units: Optional[List[str]] = None
    source_file: str
    page: Optional[int] = None
    sheet_name: Optional[str] = None
    section_hint: Optional[str] = None


class FigureData(BaseModel):
    figure_id: str
    source_file: str
    page: Optional[int] = None
    slide: Optional[int] = None
    section_hint: Optional[str] = None
    ocr_text: Optional[str] = None
    confidence: float = 1.0


class NormalizedIngestionResult(BaseModel):
    chunks: List[DocumentChunk] = Field(default_factory=list)
    tables: List[TableData] = Field(default_factory=list)
    figures: List[FigureData] = Field(default_factory=list)
    metadata: Dict[str, Any] = Field(default_factory=dict)
