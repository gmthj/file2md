import os

from .converters.pdf import convert_pdf
from .converters.docx import convert_docx
from .converters.pptx import convert_pptx
from .converters.xlsx import convert_xlsx
from .converters.image import convert_image

HANDLERS = {
    ".pdf": convert_pdf,
    ".docx": convert_docx,
    ".pptx": convert_pptx,
    ".xlsx": convert_xlsx,
    ".xlsm": convert_xlsx,
    ".jpg": convert_image,
    ".jpeg": convert_image,
    ".png": convert_image,
    ".bmp": convert_image,
    ".tiff": convert_image,
    ".tif": convert_image,
}

SUPPORTED_EXTENSIONS = set(HANDLERS.keys())


def convert(filepath: str) -> str:
    ext = os.path.splitext(filepath)[1].lower()
    handler = HANDLERS.get(ext)
    if not handler:
        raise ValueError(f"Unsupported file type: {ext}")
    return handler(filepath)


def is_supported(filepath: str) -> bool:
    return os.path.splitext(filepath)[1].lower() in SUPPORTED_EXTENSIONS
