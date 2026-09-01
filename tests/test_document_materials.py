from core.document_agent.materials import (
    MaterialKind,
    classify_material,
    extract_material_text,
)


def test_classifies_supported_materials_without_trusting_extension_only() -> None:
    assert classify_material("0-操作手册.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document") is MaterialKind.MANUAL
    assert classify_material("会议.txt", "text/plain") is MaterialKind.TRANSCRIPT
    assert classify_material("会议.m4a", "audio/mp4") is MaterialKind.RECORDING
    assert classify_material("rules.json", "application/json") is MaterialKind.RULE_PACKAGE


def test_extracts_docx_paragraphs_and_keeps_source_boundaries() -> None:
    # Minimal DOCX-like OOXML payload used to keep this test dependency-free.
    import io
    from zipfile import ZIP_DEFLATED, ZipFile

    xml = (
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        '<w:body><w:p><w:r><w:t>J列有数值时采用 J 列</w:t></w:r></w:p>'
        '<w:p><w:r><w:t>两列均有值时需要确认</w:t></w:r></w:p></w:body></w:document>'
    ).encode()
    buffer = io.BytesIO()
    with ZipFile(buffer, "w", ZIP_DEFLATED) as archive:
        archive.writestr("word/document.xml", xml)

    result = extract_material_text(buffer.getvalue(), "manual.docx")
    assert result.kind is MaterialKind.MANUAL
    assert "J列有数值时采用 J 列" in result.text
    assert "两列均有值时需要确认" in result.text


def test_audio_material_is_explicitly_pending_transcription() -> None:
    result = extract_material_text(b"audio-bytes", "会议录音.m4a")
    assert result.kind is MaterialKind.RECORDING
    assert result.status == "transcription_required"
    assert result.text == ""
