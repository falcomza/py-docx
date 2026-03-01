from __future__ import annotations


def insert_at_body_start(doc_xml: str, fragment: str) -> str:
    marker = "<w:body>"
    idx = doc_xml.find(marker)
    if idx == -1:
        raise ValueError("document.xml missing <w:body>")
    insert_pos = idx + len(marker)
    return doc_xml[:insert_pos] + fragment + doc_xml[insert_pos:]


def insert_at_body_end(doc_xml: str, fragment: str) -> str:
    marker = "</w:body>"
    idx = doc_xml.rfind(marker)
    if idx == -1:
        raise ValueError("document.xml missing </w:body>")
    sect_idx = doc_xml.rfind("<w:sectPr", 0, idx)
    if sect_idx != -1:
        last_p = doc_xml.rfind("</w:p>", 0, idx)
        if sect_idx > last_p:
            return doc_xml[:sect_idx] + fragment + doc_xml[sect_idx:]
    return doc_xml[:idx] + fragment + doc_xml[idx:]
