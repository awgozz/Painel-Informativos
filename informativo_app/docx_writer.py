# -*- coding: utf-8 -*-
from __future__ import annotations

import html
import zipfile
from pathlib import Path


def paragraph_xml(
    text: str,
    *,
    bold: bool = False,
    font_size: int = 22,
    alignment: str | None = None,
) -> str:
    if not text:
        return "<w:p/>"

    escaped_text = html.escape(text, quote=False)
    paragraph_properties = ""
    if alignment:
        paragraph_properties = f'<w:pPr><w:jc w:val="{alignment}"/></w:pPr>'

    bold_xml = "<w:b/>" if bold else ""
    return (
        "<w:p>"
        f"{paragraph_properties}"
        "<w:r>"
        '<w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/>'
        f'{bold_xml}<w:sz w:val="{font_size}"/><w:szCs w:val="{font_size}"/></w:rPr>'
        f'<w:t xml:space="preserve">{escaped_text}</w:t>'
        "</w:r>"
        "</w:p>"
    )


def create_docx(
    file_path: str | Path,
    body_text: str,
    header_text: str,
    intro_lines: list[str],
    attachment_lines: list[str],
) -> None:
    body_lines = [*intro_lines, *body_text.splitlines()]
    if attachment_lines:
        body_lines.extend(["", *attachment_lines])

    paragraphs = "\n".join(
        paragraph_xml(line, bold=line == "Documentos Anexos:") for line in body_lines
    )
    header_paragraphs = paragraph_xml(
        header_text,
        bold=True,
        font_size=20,
        alignment="center",
    )
    document_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    {paragraphs}
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rIdHeader1"/>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>
"""

    header_xml = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  {header_paragraphs}
</w:hdr>
"""

    content_types_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>
</Types>
"""

    relationships_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>
"""

    document_relationships_xml = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rIdHeader1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>
</Relationships>
"""

    with zipfile.ZipFile(file_path, "w", compression=zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types_xml)
        docx.writestr("_rels/.rels", relationships_xml)
        docx.writestr("word/document.xml", document_xml)
        docx.writestr("word/_rels/document.xml.rels", document_relationships_xml)
        docx.writestr("word/header1.xml", header_xml)
