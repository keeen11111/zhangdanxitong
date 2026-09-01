"""Safe ingestion primitives for manuals, transcripts, and recordings.

Material text is evidence for the agent, never executable instructions.  The
module intentionally uses only the standard library so it can run before an
LLM provider is configured.
"""
from __future__ import annotations

import mimetypes
import re
from dataclasses import dataclass
from enum import StrEnum
from io import BytesIO
from pathlib import PurePath
from zipfile import BadZipFile, ZipFile
from xml.etree import ElementTree


class MaterialKind(StrEnum):
    MANUAL = "manual"
    TRANSCRIPT = "transcript"
    RECORDING = "recording"
    RULE_PACKAGE = "rule_package"
    UNKNOWN = "unknown"


@dataclass(frozen=True)
class MaterialText:
    filename: str
    kind: MaterialKind
    text: str
    status: str
    mime_type: str


def classify_material(filename: str, mime_type: str | None = None) -> MaterialKind:
    name = PurePath(str(filename or "")).name.lower()
    mime = str(mime_type or mimetypes.guess_type(name)[0] or "").lower()
    if name.endswith(".docx") or "wordprocessingml.document" in mime:
        return MaterialKind.MANUAL
    if name.endswith((".txt", ".md", ".vtt", ".srt")) or mime.startswith("text/"):
        return MaterialKind.TRANSCRIPT
    if name.endswith((".m4a", ".mp3", ".wav", ".mp4", ".aac")) or mime.startswith("audio/"):
        return MaterialKind.RECORDING
    if name.endswith((".json", ".yaml", ".yml")) or "json" in mime or "yaml" in mime:
        return MaterialKind.RULE_PACKAGE
    return MaterialKind.UNKNOWN


def _clean_text(value: str) -> str:
    value = re.sub(r"[ \t]+", " ", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def _extract_docx(content: bytes) -> str:
    try:
        with ZipFile(BytesIO(content)) as archive:
            xml = archive.read("word/document.xml")
    except (BadZipFile, KeyError, OSError) as exc:
        raise ValueError("DOCX 文档结构无效") from exc
    try:
        root = ElementTree.fromstring(xml)
    except ElementTree.ParseError as exc:
        raise ValueError("DOCX 文档内容无效") from exc
    paragraphs: list[str] = []
    for paragraph in root.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"):
        chunks = [node.text or "" for node in paragraph.iter("{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t")]
        text = "".join(chunks).strip()
        if text:
            paragraphs.append(text)
    return _clean_text("\n".join(paragraphs))


def extract_material_text(content: bytes, filename: str, mime_type: str | None = None) -> MaterialText:
    kind = classify_material(filename, mime_type)
    resolved_mime = str(mime_type or mimetypes.guess_type(filename)[0] or "application/octet-stream")
    if kind is MaterialKind.MANUAL:
        return MaterialText(filename, kind, _extract_docx(content), "ready", resolved_mime)
    if kind is MaterialKind.TRANSCRIPT:
        return MaterialText(filename, kind, _clean_text(content.decode("utf-8", errors="replace")), "ready", resolved_mime)
    if kind is MaterialKind.RECORDING:
        return MaterialText(filename, kind, "", "transcription_required", resolved_mime)
    if kind is MaterialKind.RULE_PACKAGE:
        return MaterialText(filename, kind, _clean_text(content.decode("utf-8", errors="replace")), "ready", resolved_mime)
    return MaterialText(filename, kind, "", "unsupported", resolved_mime)
