# Claude Code Guide — Bangla PDF Processor

This repository provides an automated engine for converting, extracting, and processing Bangla PDFs into copy-pasteable searchable PDFs and identical Word (.docx) documents with highlight preservation.

## Quick Commands
- Run processor on a PDF:
  ```bash
  python bangla_pdf_processor.py "path/to/document.pdf"
  ```
- Specify custom output name:
  ```bash
  python bangla_pdf_processor.py "path/to/document.pdf" -n "Document_Name"
  ```
- Run linguistic test suite:
  ```bash
  python -c "import scratch.test_cases"
  ```

## Architecture & Workflow
1. `bangla_pdf_processor.py` is the single source of truth containing:
   - `BijoyToUnicode`: Deterministic SutonnyMJ ANSI to Unicode conversion.
   - `BanglaPDFProcessor`: PyMuPDF layout parser, PDF classifier, searchable PDF builder, and DOCX builder.
2. Output folders are automatically named after the document base name and hold all four output formats (`.pdf`, `.docx`, `.txt`, `.md`).
3. Yellow highlights in PDFs are parsed from annotation quadpoints and mapped directly into Word formatting.
4. Official gazette tables (symbols, dimension tables) are rendered with native Word table XML.

## Coding Guidelines
- Maintain 100% Unicode standard compliance for Bengali.
- Do not commit personal documents, PDFs, or generated deliverables (enforced via `.gitignore`).
- Ensure all stdout outputs on Windows use UTF-8 (`sys.stdout.reconfigure(encoding='utf-8')`).
