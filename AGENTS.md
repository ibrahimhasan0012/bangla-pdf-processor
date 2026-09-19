# Agent Instructions — Bangla PDF Processor

This repository implements a high-precision dual-engine processor for Bengali (Bangla) PDF documents.

## Core Directives for Autonomous Agents
1. **Born-Digital Priority & Obfuscated Font Handling**: Always inspect font metadata first. If `sutonny` or `mj` fonts are detected, use deterministic glyph decoding via `BijoyToUnicode`. If embedded fonts use truncated CMaps mapping to PUA (`0xE000-0xF8FF`) or unmapped ASCII, route to deep-learning OCR fallback. Never invoke OCR on clean born-digital documents.
2. **Linguistic Precision**: Preserve the exact vowel signs, conjuncts, folas, and reph ordering described in `bangla-pdf-processor.md`.
3. **Visual & Highlight Preservation**:
   - For PDFs: Render background at 300 DPI and overlay invisible Unicode text (`render_mode=3`).
   - For Word documents: Preserve all yellow highlight annotations using `w:highlight` and table cell shading `w:shd`.
4. **Git Hygiene**: Never stage or commit PDF, DOCX, TXT, or media files. Only code, skills, and documentation belong in version control.
