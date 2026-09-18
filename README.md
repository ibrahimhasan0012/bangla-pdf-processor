# Bangla PDF Processor & Searchable PDF Generator 🇧🇩

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Engine: PyMuPDF](https://img.shields.io/badge/PDF_Engine-PyMuPDF-green.svg)](https://pymupdf.readthedocs.io/)
[![DOCX: python-docx](https://img.shields.io/badge/DOCX_Engine-python--docx-red.svg)](https://python-docx.readthedocs.io/)

A high-accuracy dual-engine system designed to process **Bengali (Bangla) PDFs** into:
1. **Searchable, Copy-Pasteable "Sandwich" PDFs** (`Ctrl+F`, `Ctrl+C` / `Ctrl+V` with 100% clean Unicode).
2. **Identical Microsoft Word (`.docx`) Documents** mirroring official Bangladesh Gazette layouts, formatted tables, emblems, and **preserving all yellow highlighted text**.
3. **Clean Unicode Plain Text (`.txt`) and Markdown (`.md`)**.

---

## ⚡ The Problem: The SutonnyMJ/Bijoy Trap

In Bangladesh, official gazettes, legal circulars, and institutional documents are overwhelmingly created in Microsoft Word using legacy **SutonnyMJ (Bijoy)** ANSI fonts. When compiled into PDFs:
* Copying text yields unintelligible mojibake: `†iwR÷vW© bs wW G-1 evsjv‡`k †M‡RU`
* Traditional OCR produces frequent hallucinations, garbled conjuncts (যুক্তাক্ষর), and dropped diacritics (কার/ফলা).

### The Solution: Direct Deterministic Decoding
Because the document is digitally authored, the vector glyph codes and baseline coordinates exist with mathematical precision. **`BanglaPDFProcessor`** decodes the SutonnyMJ glyph stream directly to standard Unicode Bengali, delivering **100% mathematical fidelity with zero OCR hallucinations**.

---

## ✨ Features

- **Born-Digital & Scanned Dual Engine**: Automatically detects Born-Digital Bijoy vs. Unicode vs. Scanned Paper PDFs (with deep-learning EasyOCR fallback).
- **Comprehensive Kar & Fola Support**:
  - Exact Ou-kar (`ৌ`) merging from pre-kar `ে` and post-mark `ৗ` (`কৌটা`, `অপকৌশল`, `ভৌত`, `মৌলিক`, `পৌর`).
  - True Reph (`র্`) repositioning before entire consonant clusters and folas (`পার্থক্য`, `কার্যক্রম`, `ধর্তব্য`, `শর্তাবলি`, `বর্গসেন্টিমিটার`).
  - Compound ligatures: `ªæ`/`«æ`/`Öæ` (`্রু`), `åæ` (`ভ্রু`), `æ` (`ম্ন` as in `নিম্নবর্ণিত`), `Ú` (`ণ্ঠ`), `i¨` (`র‍্য` with ZWJ as in `র‍্যাপার`).
- **Searchable Sandwich PDF**:
  - 300 DPI high-resolution visual background (preserving seals, signatures, annotations).
  - Invisible Unicode text layer (`render_mode=3`) using `Kalpurush` font.
  - Seamless in-document search (`Ctrl+F`) and copy-paste into Word or browsers.
- **Identical DOCX Layout**:
  - Bangladesh Gazette header banner with National Emblem.
  - Running headers on even/odd pages with page numbers and horizontal separator rules.
  - Formatted tables with official gazette symbols and cell highlight shading.
  - **Yellow Highlight Preservation**: All PDF highlight annotations are mapped directly to Word highlight runs.
- **Dedicated Output Folder**: Automatically groups all 4 formats into a neat document folder.

---

## 🚀 Quickstart

### 1. Installation

```bash
git clone https://github.com/ibrahimhasan0012/bangla-pdf-processor.git
cd bangla-pdf-processor
pip install -r requirements.txt
```

### 2. Usage

Process any Bangla PDF with a single command:

```bash
python bangla_pdf_processor.py "path/to/your_document.pdf"
```

Specify a custom document name and output directory:

```bash
python bangla_pdf_processor.py "path/to/gazette.pdf" -n "My_Gazette_2017" -o "output_folder" --dpi 300
```

### Command Line Options:
```text
usage: bangla_pdf_processor.py [-h] [-n NAME] [-o OUTPUT_DIR] [--dpi DPI] input_pdf

positional arguments:
  input_pdf             Path to input Bangla PDF file

options:
  -h, --help            show this help message and exit
  -n NAME, --name NAME  Custom base name for output folder and files
  -o OUTPUT_DIR, --output-dir OUTPUT_DIR
                        Directory where outputs will be saved
  --dpi DPI             DPI for background images (default: 300)
```

---

## 📁 Output Structure

For an input file `Packaged_Food_Labelling_Regulations, 2017.pdf`, the tool generates:

```text
Packaged_Food_Labelling_Regulations, 2017/
├── Packaged_Food_Labelling_Regulations, 2017.pdf      # Searchable PDF (Copy-pasteable)
├── Packaged_Food_Labelling_Regulations, 2017.docx     # Word document with identical layout & highlights
├── Packaged_Food_Labelling_Regulations, 2017.txt      # Clean Unicode text
├── Packaged_Food_Labelling_Regulations, 2017.md       # Markdown document
├── govt_crest.png                                     # Extracted National Emblem
├── brown_symbol.png                                   # Official gazette symbol (Non-veg)
└── green_symbol.png                                   # Official gazette symbol (Veg)
```

---

## 🤖 AI Agent & Skill Integrations

This repository includes native skill definitions ready to drop into your favorite AI assistant:

- **Universal Skill**: [`bangla-pdf-processor.md`](./bangla-pdf-processor.md)
- **Google Antigravity**: [`.agents/skills/bangla-pdf-processor/SKILL.md`](./.agents/skills/bangla-pdf-processor/SKILL.md)
- **Claude Code**: [`CLAUDE.md`](./CLAUDE.md) & [`.claude/skills/bangla-pdf-processor/SKILL.md`](./.claude/skills/bangla-pdf-processor/SKILL.md)
- **OpenAI Codex**: [`AGENTS.md`](./AGENTS.md) & [`.codex/skills/bangla-pdf-processor/SKILL.md`](./.codex/skills/bangla-pdf-processor/SKILL.md)

---

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](./LICENSE) file for details.
