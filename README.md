# file2md

Convert files to lightweight Markdown for efficient LLM token usage.

## Supported Formats

| Format | Extension | Notes |
|---|---|---|
| PDF | `.pdf` | Page-by-page text extraction |
| Word | `.docx` | Headings, paragraphs, lists, tables |
| PowerPoint | `.pptx` | Slides as sections, bullets as lists |
| Excel | `.xlsx`, `.xlsm` | Tables, formulas, VBA macros, shapes, images |
| Images | `.jpg`, `.jpeg`, `.png`, `.bmp`, `.tiff`, `.tif` | OCR text + metadata |

## Install

```bash
pip install -e .
```

### Image OCR

Image conversion uses [Tesseract](https://github.com/tesseract-ocr/tesseract). Install it separately for your OS:

```bash
# Windows
# https://github.com/UB-Mannheim/tesseract/wiki

# macOS
brew install tesseract

# Ubuntu
sudo apt install tesseract-ocr
```

## Usage

```bash
# Single file
file2md report.pdf

# Custom output path
file2md slides.pptx -o summary.md

# Print to stdout
file2md document.docx --stdout

# Directory — recursively converts all supported files
file2md ./folder/

# Mixed files and directories
file2md report.pdf ./data/ image.png
```

## Output

Converted `.md` files are written alongside the source file:

```
report.pdf  ->  report.md
slides.pptx ->  slides.md
```

### Excel (.xlsx / .xlsm)

The Excel converter captures:
- **Cell values** and **formulas** (displayed beneath values)
- **Shapes & controls** (buttons, rectangles, textboxes) with assigned macros
- **Embedded images** with file size and format
- **VBA macro source code** extracted from the workbook
