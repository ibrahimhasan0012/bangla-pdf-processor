---
name: bangla-pdf-processor
description: "High-accuracy dual-engine Bangla PDF processor: born-digital Bijoy/SutonnyMJ to Unicode decoding, dual-layer searchable PDF generation, layout-identical Word (.docx) creation with highlight preservation, and deep-learning OCR fallback."
---

# Bangla PDF Processor — Universal Skill Specification

This document defines the universal skill specification for processing, converting, and extracting Bangla (Bengali) PDF documents across all AI coding agents (Claude Code, OpenAI Codex, Google Antigravity, Cursor, Windsurf, Roo).

## Overview

Most official government gazettes, circulars, legal acts, and institutional documents in Bangladesh are compiled using legacy ANSI fonts (primarily **SutonnyMJ** and the **Bijoy keyboard layout**) rather than standard Unicode. When users attempt to copy text or extract it with standard PDF libraries, the result is unreadable mojibake (e.g. `†iwR÷vW© bs wW G-1`).

This skill provides the end-to-end procedure for:
1. **Intelligent PDF Classification**: Auto-detecting Born-Digital Bijoy vs. Born-Digital Unicode vs. Scanned Paper PDFs.
2. **Deterministic Glyph Decoding**: Direct, mathematical conversion from SutonnyMJ glyphs to standard Unicode Bengali (100% accuracy, zero OCR hallucinations or typos on conjuncts).
3. **Complex Script Reordering**: Correct handling of vowel signs (Pre-kars `ি`, `ে`, `ৈ`; composite `ো`, `ৌ`), consonant conjuncts (যুক্তাক্ষর), folas (য-ফলা, র-ফলা, ব-ফলা, ম-ফলা, ল-ফলা, ন-ফলা), and Reph (`র্`).
4. **Searchable Sandwich PDF Generation**: Creating a dual-layer PDF with high-resolution visual fidelity (300 DPI) and an invisible Unicode Bengali text layer (`render_mode=3`) supporting full selection, copy-paste (`Ctrl+C`/`Ctrl+V`), and search (`Ctrl+F`).
5. **Identical Word (.docx) Generation**: Mirroring the exact layout of official documents (headers, national emblems, running page numbers, borders, official tables) while preserving all PDF highlight annotations in standard yellow formatting.
6. **Fallback OCR**: Seamless fallback to deep-learning OCR (EasyOCR Bengali models) when processing scanned physical papers.

---

## When to Use This Skill

Activate this skill whenever:
- The user provides a Bangla PDF (gazette, legal notice, circular, book, or certificate) and wants to make it searchable or copy-pasteable.
- The user reports garbled / mojibake text when copying Bengali from a PDF.
- The user requests an editable Microsoft Word (`.docx`) file matching the original PDF layout and preserving highlighted text annotations.
- The user needs clean Bengali plain text (`.txt`) or Markdown (`.md`) extracts.

---

## Installation & Requirements

Install the necessary dependencies:

```bash
pip install pymupdf python-docx pillow
# Optional for scanned image PDFs:
pip install easyocr torch numpy
```

Recommended Unicode font on Windows / Linux:
- `Kalpurush.ttf` (Windows: `C:\Windows\Fonts\kalpurush.ttf`, Linux: `/usr/share/fonts/truetype/kalpurush/kalpurush.ttf`)
- Alternatively: `Siyamrupali.ttf`, `NotoSerifBengali-Regular.ttf`, or `Nirmala.ttc`.

---

## CLI Usage

Run the processor directly against any PDF file:

```powershell
python bangla_pdf_processor.py "path\to\document.pdf"
```

### Custom Naming & Output Directory:

```powershell
python bangla_pdf_processor.py "path\to\document.pdf" -n "Document_Name" -o "output_directory" --dpi 300
```

### Arguments:
- `input_pdf`: Path to the input PDF file.
- `-n, --name`: Custom base name for the output folder and deliverable files.
- `-o, --output-dir`: Custom target directory (default: dedicated `<base_name>/` folder).
- `--dpi`: Rendering resolution for the visual background layer (default: `300`).

---

## Generated Deliverables

Each run creates a dedicated output folder containing:

| File | Format | Description |
| :--- | :--- | :--- |
| `<name>.pdf` | Searchable PDF | Dual-layer PDF with exact visual fidelity; selectable, searchable (`Ctrl+F`), and copy-pasteable in Bengali. |
| `<name>.docx` | Microsoft Word | Layout-identical document with running headers, crests, tables, and **preserved yellow highlights**. |
| `<name>.txt` | Plain Text | Clean Unicode Bengali text organized page-by-page. |
| `<name>.md` | Markdown | Structured document with markdown headers and dividers. |

---

## Linguistic & Glyph Transformation Rules

### 1. The Ou-kar (`ৌ`) Rule
In SutonnyMJ, `ৌ` is typed as Pre-kar `†` (`ে`) before the consonant and `Š` (`ৗ` / `ৗ`) after it.
- **Rule**: Map `Š` $ightarrow$ `ৗ` (`ৗ`). During rearrangement, merge `ে` + consonant cluster + `ৗ` $ightarrow$ consonant cluster + `ৌ` (`ৌ`).
- Resolves words like `কৌটা`, `অপকৌশল`, `ভৌত`, `মৌলিক`, `পৌর`, `সৌরভ`.

### 2. Reph (`র্`) Reordering Before Folas & Clusters
In SutonnyMJ, Reph `©` is typed at the end of the consonant syllable. In Unicode, Reph is `র` + `্` (`র্`) placed at the **beginning** of the consonant cluster.
- **Rule**: Move `র` + `্` backward before the entire preceding consonant cluster, regardless of whether the succeeding consonant has a fola or halant.
- Resolves words like `পার্থক্য`, `কার্যক্রম`, `ধর্তব্য`, `শর্তাবলি`, `বর্গসেন্টিমিটার`, `আন্তর্জাতিক`, `সর্বনিম্ন`, `তদ্কর্তৃক`.

### 3. Compound Ligatures & Special U-kars
- `ªæ`, `«æ`, `Öæ` $ightarrow$ `্রু` (e.g. `দ্রুত`)
- `åæ`, `å“` $ightarrow$ `ভ্রু`
- Standalone `æ` $ightarrow$ `ম্ন` (e.g. `wbæewY©Z` $ightarrow$ `নিম্নবর্ণিত`)
- `Ú` $ightarrow$ `ণ্ঠ` (Murdhanya-Na conjunct as in `কণ্ঠ`, `উৎকণ্ঠা`)
- `i¨` $ightarrow$ `র` + `্` + `‍` (ZWJ) + `য` (`র‍্য` as in `র‍্যাপার` / wrapper)

### 4. Text Layer Space Preservation
Ensure space spans (`' '`) are explicitly inserted into the invisible text layer (`render_mode=3`) so that copied text preserves inter-word spacing.
