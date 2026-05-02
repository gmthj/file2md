import os
import shutil

from PIL import Image
import pytesseract

TESSERACT_PATHS = [
    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
]


def _set_tesseract_path():
    if not shutil.which("tesseract"):
        for path in TESSERACT_PATHS:
            if os.path.isfile(path):
                pytesseract.pytesseract.tesseract_cmd = path
                return


_set_tesseract_path()


def convert_image(filepath: str) -> str:
    img = Image.open(filepath)
    width, height = img.size
    fmt = img.format
    file_size = os.path.getsize(filepath)

    md = []
    md.append(f"**Image:** `{os.path.basename(filepath)}`")
    md.append(f"- **Dimensions:** {width} x {height} px")
    md.append(f"- **Format:** {fmt}")
    md.append(f"- **File size:** {file_size:,} bytes")

    try:
        text = pytesseract.image_to_string(img).strip()
        if text:
            md.append(f"\n### Extracted text\n\n{text}")
        else:
            md.append("\n_(No text detected in image)_")
    except pytesseract.TesseractNotFoundError:
        md.append("\n_(OCR skipped: Tesseract not installed)_")

    img.close()
    return "\n".join(md)
