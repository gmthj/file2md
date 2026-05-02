import fitz


def convert_pdf(filepath: str) -> str:
    doc = fitz.open(filepath)
    parts = []
    for page_num, page in enumerate(doc, 1):
        text = page.get_text("text")
        if text.strip():
            parts.append(f"## Page {page_num}\n\n{text.strip()}")
    doc.close()
    return "\n\n---\n\n".join(parts)
