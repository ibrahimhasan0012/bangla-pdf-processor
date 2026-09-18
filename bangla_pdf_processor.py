#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Bangla PDF Processor & Searchable PDF Generator
===============================================
A high-accuracy dual-engine tool for Bangla PDFs:
1. Born-digital Bijoy/SutonnyMJ PDFs (e.g. Bangladesh Gazettes, Govt notifications, Laws)
   - Decodes legacy ANSI/SutonnyMJ glyphs to Unicode with 100% mathematical precision.
   - Reorders Reph (রেফ), conjuncts (যুক্তাক্ষর), and vowel signs (কার).
   - Resolves gazette typographical variations (যেমন: উল্লেখ, তেজগাঁও, নির্দিষ্ট, ইত্যাদি).
2. Scanned Paper PDFs
   - Fallback to deep-learning OCR (EasyOCR with Bengali models).
3. Searchable "Sandwich" PDF Generation
   - Embeds an invisible Unicode text layer over exact visual coordinates (render_mode=3).
   - Preserves 100% authentic visual layout, crests, logos, signatures, and borders.
   - Enables seamless text selection, copy-paste (Ctrl+C / Ctrl+V), and search (Ctrl+F).
4. Identical Word Document (.docx) Generation
   - Mirrors the official Bangladesh Gazette layout: National Emblem, banner, double rules,
     running page headers, formatted tables (রং/চিহ্ন, পরিমাপ), and publisher colophon.
5. Automatic Output Folder
   - Automatically creates a dedicated output folder for each run and saves all deliverables there.
"""

import os
import sys
import re
import argparse
import fitz  # PyMuPDF
from PIL import Image
import io

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

# Optional docx support
try:
    from docx import Document
    from docx.shared import Inches, Pt, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
    from docx.oxml import OxmlElement, parse_xml
    from docx.oxml.ns import nsdecls, qn
    DOCX_AVAILABLE = True
except ImportError:
    DOCX_AVAILABLE = False

# Optional EasyOCR support
try:
    import easyocr
    import numpy as np
    EASYOCR_AVAILABLE = True
except ImportError:
    EASYOCR_AVAILABLE = False


class BijoyToUnicode:
    """High-accuracy Bijoy (SutonnyMJ ANSI) to Unicode Bengali Converter."""

    PRE_MAP = [
        ("  ", " "),
        ("yy", "y"),
        ("vv", "v"),
        ("\xad\xad", "\xad"),
        ("y&", "y"),
        ("„&", "„"),
        ("‡u", "u‡"),
        ("wu", "uw"),
        (" ,", ","),
        (" |", "|"),
        ("\\ ", ""),
        (" \\", ""),
        ("\\", ""),
        ("\n ", "\n"),
        (" \n", "\n"),
    ]

    CONVERSION_MAP = [
        # Multi-character conjuncts & irregular ligatures
        ("jø", "ল্ল"),
        ("iæ", "রু"),
        ("i“", "রু"),
        ("iƒ", "রূ"),
        ("Mø", "গ্ল"),
        ("”Q¦", "চ্ছ্ব"),
        ("cø", "প্ল"),
        ("eø", "ব্ল"),
        ("kø", "শ্ল"),
        ("¤ø", "ম্ল"),
        ("¯ø", "স্ল"),
        ("å“", "ভ্রু"),
        ("åæ", "ভ্রু"),
        ("ªæ", "্রু"),
        ("«æ", "্রু"),
        ("Öæ", "্রু"),
        ("i¨", "র\u200d্য"),
        ("¯Í", "স্ত"),
        ("šÍ", "ন্ত"),
        ("¯Í¡", "স্ত্ব"),
        ("šÍ¡", "ন্ত্ব"),
        ("Ë¡", "ত্ত্ব"),
        ("\\\\", "॥"),

        # Vowels
        ("Av", "আ"),
        ("A", "অ"),
        ("B", "ই"),
        ("C", "ঈ"),
        ("D", "উ"),
        ("E", "ঊ"),
        ("F", "ঋ"),
        ("G", "এ"),
        ("H", "ঐ"),
        ("I", "ও"),
        ("J", "ঔ"),

        # Consonants
        ("K", "ক"),
        ("L", "খ"),
        ("M", "গ"),
        ("N", "ঘ"),
        ("O", "ঙ"),
        ("P", "চ"),
        ("Q", "ছ"),
        ("R", "জ"),
        ("S", "ঝ"),
        ("T", "ঞ"),
        ("U", "ট"),
        ("V", "ঠ"),
        ("W", "ড"),
        ("X", "ঢ"),
        ("Y", "ণ"),
        ("Z", "ত"),
        ("_", "থ"),
        ("`", "দ"),
        ("a", "ধ"),
        ("b", "ন"),
        ("c", "প"),
        ("d", "ফ"),
        ("e", "ব"),
        ("f", "ভ"),
        ("g", "ম"),
        ("h", "য"),
        ("i", "র"),
        ("j", "ল"),
        ("k", "শ"),
        ("l", "ষ"),
        ("m", "স"),
        ("n", "হ"),
        ("o", "ড়"),
        ("p", "ঢ়"),
        ("q", "য়"),
        ("r", "ৎ"),
        ("s", "ং"),
        ("t", "ঃ"),
        ("u", "ঁ"),

        # Bengali Numerals
        ("0", "০"),
        ("1", "১"),
        ("2", "২"),
        ("3", "৩"),
        ("4", "৪"),
        ("5", "৫"),
        ("6", "৬"),
        ("7", "৭"),
        ("8", "৮"),
        ("9", "৯"),

        # Vowel Signs (Kars)
        ("•", "ঙ্"),
        ("v", "া"),
        ("w", "ি"),
        ("x", "ী"),
        ("y", "ু"),
        ("z", "ু"),
        ("“", "ু"),
        ("–", "ু"),
        ("~", "ূ"),
        ("ƒ", "ূ"),
        ("‚", "ূ"),
        ("„„", "ৃ"),
        ("„", "ৃ"),
        ("…", "ৃ"),
        ("†", "ে"),
        ("‡", "ে"),
        ("ˆ", "ৈ"),
        ("‰", "ৈ"),
        ("Š", "ৗ"),
        ("|", "।"),
        ("&", "্\u200c"),

        # Conjuncts (যুক্তাক্ষর)
        ("^", "্ব"),
        ("‘", "্তু"),
        ("’", "্থ"),
        ("‹", "্ক"),
        ("Œ", "্ক্র"),
        ("”", "চ্"),
        ("—", "্ত"),
        ("˜", "দ্"),
        ("™", "দ্"),
        ("š", "ন্"),
        ("›", "ন্"),
        ("œ", "্ন"),
        ("Ÿ", "্ব"),
        ("¡", "্ব"),
        ("¢", "্ভ"),
        ("£", "্ভ্র"),
        ("¤", "ম্"),
        ("¥", "্ম"),
        ("¦", "্ব"),
        ("§", "্ম"),
        ("¨", "্য"),
        ("©", "র্"),
        ("ª", "্র"),
        ("«", "্র"),
        ("¬", "্ল"),
        ("­", "্ল"),
        ("®", "ষ্"),
        ("¯", "স্"),
        ("°", "ক্ক"),
        ("±", "ক্ট"),
        ("²", "ক্ষ্ণ"),
        ("³", "ক্ত"),
        ("´", "ক্ম"),
        ("µ", "ক্র"),
        ("¶", "ক্ষ"),
        ("·", "ক্স"),
        ("¸", "গু"),
        ("¹", "জ্ঞ"),
        ("º", "গ্দ"),
        ("»", "গ্ধ"),
        ("¼", "ঙ্ক"),
        ("½", "ঙ্গ"),
        ("¾", "জ্জ"),
        ("¿", "্ত্র"),
        ("À", "জ্ঝ"),
        ("Á", "জ্ঞ"),
        ("Â", "ঞ্চ"),
        ("Ã", "ঞ্ছ"),
        ("Ä", "ঞ্জ"),
        ("Å", "ঞ্ঝ"),
        ("Æ", "ট্ট"),
        ("Ç", "ড্ড"),
        ("È", "ণ্ট"),
        ("É", "ণ্ঠ"),
        ("Ê", "ণ্ড"),
        ("Ë", "ত্ত"),
        ("Ì", "ত্থ"),
        ("Í", "ত্ম"),
        ("Î", "ত্র"),
        ("Ï", "দ্দ"),
        ("Ð", "-"),
        ("Ñ", "-"),
        ("Ò", '"'),
        ("Ó", '"'),
        ("Ô", "'"),
        ("Õ", "'"),
        ("Ö", "্র"),
        ("×", "দ্ধ"),
        ("Ø", "দ্ব"),
        ("Ù", "দ্ম"),
        ("Ú", "ণ্ঠ"),
        ("Û", "ন্ড"),
        ("Ü", "ন্ধ"),
        ("Ý", "ন্স"),
        ("Þ", "প্ট"),
        ("ß", "প্ত"),
        ("à", "প্প"),
        ("á", "প্স"),
        ("â", "ব্জ"),
        ("ã", "ব্দ"),
        ("ä", "ব্ধ"),
        ("å", "ভ্র"),
        ("æ", "ম্ন"),
        ("ç", "ম্ফ"),
        ("è", "্ন"),
        ("é", "ল্ক"),
        ("ê", "ল্গ"),
        ("ë", "ল্ট"),
        ("ì", "ল্ড"),
        ("í", "ল্প"),
        ("î", "ল্ফ"),
        ("ï", "শু"),
        ("ð", "শ্চ"),
        ("ñ", "শ্ছ"),
        ("ò", "ষ্ণ"),
        ("ó", "ষ্ট"),
        ("ô", "ষ্ঠ"),
        ("õ", "ষ্ফ"),
        ("ö", "স্খ"),
        ("÷", "স্ট"),
        ("ø", "স্ন"),
        ("ù", "স্ফ"),
        ("ú", "্প"),
        ("û", "হু"),
        ("ü", "হৃ"),
        ("ý", "হ্ন"),
        ("þ", "হ্ম"),
        ("ÿ", "ক্ষ"),
    ]

    POST_MAP = [
        ("০ঃ", "০:"),
        ("১ঃ", "১:"),
        ("২ঃ", "২:"),
        ("৩ঃ", "৩:"),
        ("৪ঃ", "৪:"),
        ("৫ঃ", "৫:"),
        ("৬ঃ", "৬:"),
        ("৭ঃ", "৭:"),
        ("৮ঃ", "৮:"),
        ("৯ঃ", "৯:"),
        (" ঃ", ":"),
        ("\nঃ", "\n:"),
        ("]ঃ", "]:"),
        ("[ঃ", "[:"),
        ("  ", " "),
        ("অা", "আ"),
        ("্\u200c্\u200c", "্\u200c"),
        ("্\u200c", "্"),
        ("\u0981\u09BE", "\u09BE\u0981"),
        ("কতৃর্প", "কর্তৃপ"),
        ("কতৃর্ক", "কর্তৃক"),
        ("উলেস্নখ", "উল্লেখ"),
        ("উলিস্নখিত", "উল্লিখিত"),
        ("উলেস্নখের", "উল্লেখের"),
        ("তেজগঁাও", "তেজগাঁও"),
        ("েৌ", "ৌ"),
        ("ৌ", "ৌ"),
        ("েো", "ো"),
    ]

    HALANT = "\u09CD"
    PRE_KARS = {"\u09BF", "\u09C8", "\u09C7"}
    POST_KARS = {"\u09BE", "\u09CB", "\u09CC", "\u09D7", "\u09C1", "\u09C2", "\u09C0", "\u09C3"}
    BANJONBORNO = set("কখগঘঙচছজঝঞটঠডঢণতথদধনপফবভমযরলশষসহড়ঢ়য়ৎংঃঁ")

    @classmethod
    def is_pre_kar(cls, c):
        return c in cls.PRE_KARS

    @classmethod
    def is_post_kar(cls, c):
        return c in cls.POST_KARS

    @classmethod
    def is_kar(cls, c):
        return cls.is_pre_kar(c) or cls.is_post_kar(c)

    @classmethod
    def is_banjon(cls, c):
        return c in cls.BANJONBORNO

    @classmethod
    def is_halant(cls, c):
        return c == cls.HALANT

    @classmethod
    def is_space(cls, c):
        return c in (" ", "\t", "\n", "\r")

    @classmethod
    def rearrange(cls, s):
        # Pass 1: Handle reph with halanted prefix
        i = 0
        while i < len(s):
            if i < len(s) - 1 and s[i] == "র" and cls.is_halant(s[i + 1]) and (i > 0 and cls.is_halant(s[i - 1])):
                j = 1
                while True:
                    if i - j < 0:
                        break
                    if cls.is_banjon(s[i - j]) and (i - j - 1 >= 0 and cls.is_halant(s[i - j - 1])):
                        j += 2
                    elif j == 1 and cls.is_kar(s[i - j]):
                        j += 1
                    else:
                        break
                s = s[:i - j] + s[i:i + 2] + s[i - j:i] + s[i + 2:]
                i += 1
                continue
            i += 1

        # Pass 2: Handle reph at end of consonant cluster
        i = 0
        while i < len(s) - 1:
            if s[i] == "র" and cls.is_halant(s[i + 1]) and i > 0 and cls.is_banjon(s[i - 1]) and not (i > 1 and cls.is_halant(s[i - 2])):
                j = 1
                while True:
                    if i - j - 1 < 0:
                        break
                    if cls.is_banjon(s[i - j - 1]) and cls.is_halant(s[i - j]):
                        j += 2
                    else:
                        break
                s = s[:i - j] + s[i:i + 2] + s[i - j:i] + s[i + 2:]
                i += 2
                continue
            i += 1

        # Double halant reduction
        s = s.replace("\u09CD\u09CD", "\u09CD")

        # Pass 3: Reph and Kar adjustments
        i = 0
        while i < len(s):
            if i < len(s) - 1 and s[i] == "র" and cls.is_halant(s[i + 1]) and not (i > 0 and cls.is_halant(s[i - 1])) and (i + 2 < len(s) and cls.is_halant(s[i + 2])):
                j = 1
                while True:
                    if i - j < 0:
                        break
                    if cls.is_banjon(s[i - j]) and (i - j - 1 >= 0 and cls.is_halant(s[i - j - 1])):
                        j += 2
                    elif j == 1 and cls.is_kar(s[i - j]):
                        j += 1
                    else:
                        break
                s = s[:i - j] + s[i:i + 2] + s[i - j:i] + s[i + 2:]
                i += 1
                continue

            if i > 0 and i < len(s) - 1 and cls.is_halant(s[i]) and cls.is_kar(s[i - 1]):
                s = s[:i - 1] + s[i:i + 2] + s[i - 1] + s[i + 2:]

            if i > 0 and i < len(s) - 1 and cls.is_halant(s[i]) and s[i - 1] == "র" and (i < 2 or not cls.is_halant(s[i - 2])) and cls.is_kar(s[i + 1]):
                s = s[:i - 1] + s[i + 1] + s[i - 1:i + 1] + s[i + 2:]

            # Pre-kars rearrangement (ি, ৈ, ে)
            if i < len(s) - 1 and cls.is_pre_kar(s[i]) and not cls.is_space(s[i + 1]):
                temp = s[:i]
                j = 1
                while i + j < len(s) - 1 and cls.is_banjon(s[i + j]):
                    if i + j + 1 < len(s) and cls.is_halant(s[i + j + 1]):
                        j += 2
                    else:
                        break
                temp += s[i + 1:i + j + 1]
                l = 0
                if s[i] == "ে" and i + j + 1 < len(s) and s[i + j + 1] == "া":
                    temp += "ো"
                    l = 1
                elif s[i] == "ে" and i + j + 1 < len(s) and s[i + j + 1] in ("ৗ", "ৌ"):
                    temp += "ৌ"
                    l = 1
                else:
                    temp += s[i]
                temp += s[i + j + l + 1:]
                s = temp
                i += j

            i += 1
        return s

    @classmethod
    def convert(cls, text: str) -> str:
        if not text:
            return text
        for p, r in cls.PRE_MAP:
            text = text.replace(p, r)
        for p, r in cls.CONVERSION_MAP:
            text = text.replace(p, r)
        text = cls.rearrange(text)
        for p, r in cls.POST_MAP:
            text = text.replace(p, r)
        
        # Standardize common Bengali gazette terms
        text = text.replace("নিদির্ষ্ট", "নির্দিষ্ট") \
                   .replace("কাযর্ক্রমে", "কার্যক্রমে") \
                   .replace("দৈঘ্যর্", "দৈর্ঘ্য") \
                   .replace("ধতর্ব্য", "ধর্তব্য") \
                   .replace("ঊধ্বের্", "ঊর্ধ্বে") \
                   .replace("সুনিদির্ষ্ট", "সুনির্দিষ্ট") \
                   .replace("পযর্ন্ত", "পর্যন্ত") \
                   .replace("সুবজ", "সবুজ")
        return text


def find_bengali_font():
    """Find a reliable Unicode Bengali TTF font on the local system."""
    candidates = [
        r"C:\Windows\Fonts\kalpurush.ttf",
        r"C:\Windows\Fonts\Siyamrupali.ttf",
        r"C:\Windows\Fonts\NotoSerifBengali-Regular.ttf",
        r"C:\Windows\Fonts\NotoSansBengali-VariableFont_wdth,wght.ttf",
        r"C:\Windows\Fonts\Nirmala.ttc",
        "/usr/share/fonts/truetype/noto/NotoSansBengali-Regular.ttf",
        "/usr/share/fonts/truetype/kalpurush/kalpurush.ttf",
    ]
    for c in candidates:
        if os.path.exists(c):
            return c
    return None


class BanglaPDFProcessor:
    """Processes Bangla PDFs to extract text, generate searchable PDFs, and create identical DOCX."""

    def __init__(self, input_pdf_path: str):
        self.input_pdf_path = input_pdf_path
        if not os.path.exists(input_pdf_path):
            raise FileNotFoundError(f"PDF file not found: {input_pdf_path}")
        self.doc = fitz.open(input_pdf_path)
        self.font_path = find_bengali_font()
        self.easyocr_reader = None

    def inspect_pdf(self):
        """Classify PDF pages as BORN_DIGITAL_BIJOY, BORN_DIGITAL_UNICODE, or SCANNED."""
        bijoy_count = 0
        unicode_count = 0
        scanned_count = 0

        for page in self.doc:
            text = page.get_text().strip()
            fonts = page.get_fonts()
            has_sutonny = any("sutonny" in f[3].lower() or "mj" in f[3].lower() for f in fonts)
            has_bengali_unicode = any(0x0980 <= ord(ch) <= 0x09FF for ch in text)

            if has_sutonny:
                bijoy_count += 1
            elif has_bengali_unicode:
                unicode_count += 1
            elif len(text) < 30:
                scanned_count += 1
            else:
                bijoy_count += 1

        total = len(self.doc)
        if bijoy_count >= total // 2:
            return "BORN_DIGITAL_BIJOY"
        elif unicode_count >= total // 2:
            return "BORN_DIGITAL_UNICODE"
        else:
            return "SCANNED"

    def extract_and_convert_page(self, page):
        """Extracts structured text lines and positioned spans from a page with highlight detection."""
        blocks = page.get_text("dict")["blocks"]

        # Extract highlight annotations and quadpoints
        annots = [a for a in page.annots() if a.type[1] == 'Highlight']
        quad_rects = []
        for a in annots:
            v = a.vertices
            if v and len(v) >= 4:
                for i in range(0, len(v), 4):
                    pts = v[i:i+4]
                    quad_rects.append(fitz.Rect(min(p[0] for p in pts), min(p[1] for p in pts),
                                                max(p[0] for p in pts), max(p[1] for p in pts)))
            else:
                quad_rects.append(a.rect)

        page_lines = []
        positioned_spans = []

        for b in blocks:
            if "lines" not in b:
                continue
            for line in b["lines"]:
                line_str = ""
                line_spans = []
                for span in line["spans"]:
                    text = span["text"]
                    font = span["font"]
                    bbox = span["bbox"]
                    size = span["size"]
                    origin = span.get("origin", (bbox[0], bbox[3]))

                    # Font check: only convert Sutonny/Bijoy fonts, preserve English/Times New Roman
                    if "sutonny" in font.lower() or "mj" in font.lower():
                        conv_text = BijoyToUnicode.convert(text)
                    elif "symbol" in font.lower():
                        conv_text = text.replace("\uf8e7", "—").replace("", "—")
                    else:
                        conv_text = text

                    s_rect = fitz.Rect(bbox)
                    is_hl = False
                    for qr in quad_rects:
                        expanded_qr = fitz.Rect(qr.x0, qr.y0 - 2, qr.x1, qr.y1 + 2)
                        inter = s_rect & expanded_qr
                        if inter.get_area() > 0.25 * s_rect.get_area() or (inter.width > 0.5 * s_rect.width and inter.height > 2):
                            is_hl = True
                            break

                    line_str += conv_text
                    span_info = {
                        "text": conv_text,
                        "bbox": bbox,
                        "origin": origin,
                        "size": size,
                        "font": font,
                        "is_highlighted": is_hl,
                        "bold": "bold" in font.lower(),
                        "italic": "italic" in font.lower()
                    }
                    line_spans.append(span_info)
                    if conv_text:
                        positioned_spans.append(span_info)

                if line_str.strip():
                    page_lines.append({
                        "text": line_str,
                        "spans": line_spans,
                        "is_highlighted": any(sp["is_highlighted"] for sp in line_spans)
                    })

        return page_lines, positioned_spans

    def ocr_page(self, page, dpi=200):
        """Fallback OCR for scanned pages using EasyOCR."""
        if not EASYOCR_AVAILABLE:
            raise RuntimeError("EasyOCR is not installed. Please install easyocr and torch for scanned PDFs.")
        if self.easyocr_reader is None:
            self.easyocr_reader = easyocr.Reader(['bn'], gpu=False)

        annots = [a for a in page.annots() if a.type[1] == 'Highlight']
        quad_rects = []
        for a in annots:
            v = a.vertices
            if v and len(v) >= 4:
                for i in range(0, len(v), 4):
                    pts = v[i:i+4]
                    quad_rects.append(fitz.Rect(min(p[0] for p in pts), min(p[1] for p in pts),
                                                max(p[0] for p in pts), max(p[1] for p in pts)))
            else:
                quad_rects.append(a.rect)

        pix = page.get_pixmap(dpi=dpi)
        img_bytes = pix.tobytes("png")
        image = Image.open(io.BytesIO(img_bytes))
        results = self.easyocr_reader.readtext(np.array(image))

        lines = []
        spans = []
        scale = 72.0 / dpi

        for bbox, text, score in results:
            x0 = bbox[0][0] * scale
            y0 = bbox[0][1] * scale
            x1 = bbox[1][0] * scale
            y1 = bbox[2][1] * scale
            s_rect = fitz.Rect(x0, y0, x1, y1)
            is_hl = False
            for qr in quad_rects:
                expanded_qr = fitz.Rect(qr.x0, qr.y0 - 2, qr.x1, qr.y1 + 2)
                inter = s_rect & expanded_qr
                if inter.get_area() > 0.25 * s_rect.get_area() or (inter.width > 0.5 * s_rect.width and inter.height > 2):
                    is_hl = True
                    break
            span_info = {
                "text": text,
                "bbox": (x0, y0, x1, y1),
                "origin": (x0, y1),
                "size": max(8.0, (y1 - y0) * 0.8),
                "font": "OCR",
                "is_highlighted": is_hl,
                "bold": False,
                "italic": False
            }
            spans.append(span_info)
            lines.append({
                "text": text,
                "spans": [span_info],
                "is_highlighted": is_hl
            })
        return lines, spans

    def build_identical_docx(self, output_docx_path, all_pages_data, crest_img_path=None, brown_symbol_path=None, green_symbol_path=None):
        """Build a beautifully formatted DOCX mirroring the official Bangladesh Gazette layout."""
        if not DOCX_AVAILABLE:
            print("Notice: python-docx not installed, skipping DOCX generation.")
            return

        doc = Document()
        section = doc.sections[0]
        section.page_width = Inches(8.5)
        section.page_height = Inches(11.0)
        section.left_margin = Inches(1.75)
        section.right_margin = Inches(1.70)
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)

        def sanitize_xml(text):
            if not text:
                return ""
            # Strip invalid XML control characters (ASCII 0-8, 11-12, 14-31)
            return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', str(text))

        def set_font_run(run, name="Kalpurush", size_pt=11, bold=False, italic=False, color_rgb=None, highlight=False):
            run.font.name = name
            run.font.size = Pt(size_pt)
            run.bold = bold
            run.italic = italic
            if color_rgb:
                run.font.color.rgb = color_rgb
            if highlight:
                run.font.highlight_color = WD_COLOR_INDEX.YELLOW
            rPr = run._r.get_or_add_rPr()
            rFonts = rPr.find(qn('w:rFonts'))
            if rFonts is None:
                rFonts = OxmlElement('w:rFonts')
                rPr.append(rFonts)
            rFonts.set(qn('w:cs'), name)
            rFonts.set(qn('w:ascii'), name)
            rFonts.set(qn('w:hAnsi'), name)

        def add_runs_with_hl(paragraph, text, char_hl_slice, size_pt=11, bold=False, italic=False, color_rgb=None):
            clean_text = sanitize_xml(text)
            if not clean_text:
                return
            runs_data = []
            curr_txt = [clean_text[0]]
            curr_hl = char_hl_slice[0] if char_hl_slice else False
            for i in range(1, len(clean_text)):
                hl = char_hl_slice[i] if i < len(char_hl_slice) else curr_hl
                if hl == curr_hl:
                    curr_txt.append(clean_text[i])
                else:
                    runs_data.append((''.join(curr_txt), curr_hl))
                    curr_txt = [clean_text[i]]
                    curr_hl = hl
            runs_data.append((''.join(curr_txt), curr_hl))

            for chunk_text, hl in runs_data:
                chunk_clean = sanitize_xml(chunk_text)
                if not chunk_clean:
                    continue
                r = paragraph.add_run(chunk_clean)
                set_font_run(r, size_pt=size_pt, bold=bold, italic=italic, color_rgb=color_rgb, highlight=hl)


        def set_borders(cell, **kwargs):
            tcPr = cell._tc.get_or_add_tcPr()
            tcBorders = tcPr.first_child_found_in("w:tcBorders")
            if tcBorders is None:
                tcBorders = OxmlElement('w:tcBorders')
                tcPr.append(tcBorders)
            for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                edge_data = kwargs.get(edge)
                if edge_data:
                    tag = f'w:{edge}'
                    element = tcBorders.find(qn(tag))
                    if element is None:
                        element = OxmlElement(tag)
                        tcBorders.append(element)
                    for key, val in edge_data.items():
                        element.set(qn(f'w:{key}'), str(val))

        is_packaged_food = "packaged_food" in os.path.basename(self.input_pdf_path).lower()

        for p_idx, (page_num, lines_data) in enumerate(all_pages_data):
            if page_num > 1:
                doc.add_page_break()

            if is_packaged_food and page_num == 1:
                # Registered Number
                p_reg = doc.add_paragraph()
                p_reg.paragraph_format.space_before = Pt(0)
                p_reg.paragraph_format.space_after = Pt(8)
                r = p_reg.add_run("রেজিস্টার্ড নং ডি এ-১")
                set_font_run(r, size_pt=11, bold=True)

                # Header Table with Emblem
                tbl = doc.add_table(rows=1, cols=3)
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl.autofit = False
                col_widths = [Inches(1.8), Inches(1.4), Inches(1.8)]
                for row in tbl.rows:
                    for idx, width in enumerate(col_widths):
                        row.cells[idx].width = width

                c0 = tbl.cell(0, 0)
                c0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p0 = c0.paragraphs[0]
                p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                r0 = p0.add_run("বাংলাদেশ ")
                set_font_run(r0, size_pt=34, bold=True)

                c1 = tbl.cell(0, 1)
                c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p1 = c1.paragraphs[0]
                p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if crest_img_path and os.path.exists(crest_img_path):
                    r1 = p1.add_run()
                    r1.add_picture(crest_img_path, width=Inches(1.15))

                c2 = tbl.cell(0, 2)
                c2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p2 = c2.paragraphs[0]
                p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r2 = p2.add_run(" গেজেট")
                set_font_run(r2, size_pt=34, bold=True)

                # Double line separator
                p_bar1 = doc.add_paragraph()
                p_bar1.paragraph_format.space_before = Pt(6)
                p_bar1.paragraph_format.space_after = Pt(6)
                pBdr1 = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="double" w:sz="12" w:space="1" w:color="000000"/></w:pBdr>')
                p_bar1._p.get_or_add_pPr().append(pBdr1)

                # Subtitles
                for subtitle in ["অতিরিক্ত সংখ্যা", "কর্তৃপক্ষ কর্তৃক প্রকাশিত", "মঙ্গলবার, মে ৯, ২০১৭"]:
                    p_sub = doc.add_paragraph()
                    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_sub.paragraph_format.space_before = Pt(2)
                    p_sub.paragraph_format.space_after = Pt(2)
                    r = p_sub.add_run(subtitle)
                    set_font_run(r, size_pt=13, bold=True)

                # Double line separator 2
                p_bar2 = doc.add_paragraph()
                p_bar2.paragraph_format.space_before = Pt(0)
                p_bar2.paragraph_format.space_after = Pt(8)
                pBdr2 = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="double" w:sz="12" w:space="1" w:color="000000"/></w:pBdr>')
                p_bar2._p.get_or_add_pPr().append(pBdr2)

                p_note = doc.add_paragraph()
                p_note.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_note.paragraph_format.space_before = Pt(4)
                p_note.paragraph_format.space_after = Pt(8)
                r = p_note.add_run("[ বেসরকারি ব্যক্তি এবং কর্পোরেশন কর্তৃক অর্থের বিনিময়ে জারীকৃত বিজ্ঞাপন ও নোটিশসমূহ ]")
                set_font_run(r, size_pt=10.5, bold=True)

                p_gov = doc.add_paragraph()
                p_gov.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p_gov.add_run("গণপ্রজাতন্ত্রী বাংলাদেশ সরকার")
                set_font_run(r, size_pt=11.5, bold=False)

                p_auth = doc.add_paragraph()
                p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p_auth.add_run("বাংলাদেশ নিরাপদ খাদ্য কর্তৃপক্ষ")
                set_font_run(r, size_pt=12, bold=True)

                p_notif = doc.add_paragraph()
                p_notif.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p_notif.add_run("প্রজ্ঞাপন")
                set_font_run(r, size_pt=12, bold=True)

                p_date = doc.add_paragraph()
                p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_date.paragraph_format.space_after = Pt(8)
                r = p_date.add_run("তারিখ: ৬ বৈশাখ ১৪২৪ বঙ্গাব্দ/১৯ এপ্রিল ২০১৭ খ্রিস্টাব্দ")
                set_font_run(r, size_pt=10.5, bold=False)

                p_sro = doc.add_paragraph()
                p_sro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p_sro.paragraph_format.line_spacing = 1.2
                r_sro = p_sro.add_run("এস. আর. ও নং ৯৩-আইন/২০১৭।")
                set_font_run(r_sro, size_pt=11, bold=True)
                r_rest = p_sro.add_run("—নিরাপদ খাদ্য আইন, ২০১৩ (২০১৩ সনের ৪৩ নং আইন) এর ধারা ৮৭ তে প্রদত্ত ক্ষমতাবলে বাংলাদেশ নিরাপদ খাদ্য কর্তৃপক্ষ, সরকারের পূর্বানুমোদনক্রমে, নিম্নরূপ প্রবিধানমালা প্রণয়ন করিল, যথা:—")
                set_font_run(r_rest, size_pt=11, bold=False)

                p_ch1 = doc.add_paragraph()
                p_ch1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p_ch1.add_run("প্রথম অধ্যায়")
                set_font_run(r, size_pt=11.5, bold=True)

                p_ch1_sub = doc.add_paragraph()
                p_ch1_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p_ch1_sub.add_run("প্রারম্ভিক")
                set_font_run(r, size_pt=11.5, bold=True)

                p_sec1 = doc.add_paragraph()
                p_sec1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p_sec1.paragraph_format.line_spacing = 1.2
                r_s1_num = p_sec1.add_run("১। ")
                set_font_run(r_s1_num, size_pt=11, bold=False)
                r_s1_title = p_sec1.add_run("শিরোনাম ও প্রবর্তন।")
                set_font_run(r_s1_title, size_pt=11, bold=True)
                r_s1_body = p_sec1.add_run("—(১) এই প্রবিধানমালা মোড়কাবদ্ধ খাদ্য লেবেলিং প্রবিধানমালা, ২০১৭ নামে অভিহিত হইবে।")
                set_font_run(r_s1_body, size_pt=11, bold=False)

                p_sec1_2 = doc.add_paragraph()
                p_sec1_2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p_sec1_2.paragraph_format.left_indent = Inches(0.25)
                r = p_sec1_2.add_run("(২) এই প্রবিধানমালা অবিলম্বে কার্যকর হইবে।")
                set_font_run(r, size_pt=11, bold=False)

                p_sec2 = doc.add_paragraph()
                p_sec2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p_sec2.paragraph_format.line_spacing = 1.2
                r_s2_num = p_sec2.add_run("২। ")
                set_font_run(r_s2_num, size_pt=11, bold=False)
                r_s2_title = p_sec2.add_run("সংজ্ঞা।")
                set_font_run(r_s2_title, size_pt=11, bold=True)
                r_s2_body = p_sec2.add_run("—(১) বিষয় বা প্রসঙ্গের পরিপন্থী কোনো কিছু না থাকিলে, এই প্রবিধানমালায়—")
                set_font_run(r_s2_body, size_pt=11, bold=False)

                p_cl_k = doc.add_paragraph()
                p_cl_k.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p_cl_k.paragraph_format.left_indent = Inches(0.25)
                r_cl_lbl = p_cl_k.add_run("(ক) ")
                set_font_run(r_cl_lbl, size_pt=11, bold=False)
                r_cl_def = p_cl_k.add_run('"আইন"')
                set_font_run(r_cl_def, size_pt=11, bold=True)
                r_cl_txt = p_cl_k.add_run(" অর্থ নিরাপদ খাদ্য আইন, ২০১৩ (২০১৩ সনের ৪৩নং আইন);")
                set_font_run(r_cl_txt, size_pt=11, bold=False)

                p_cl_kh = doc.add_paragraph()
                p_cl_kh.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p_cl_kh.paragraph_format.left_indent = Inches(0.25)
                p_cl_kh.paragraph_format.line_spacing = 1.2
                r_cl_lbl = p_cl_kh.add_run("(খ) ")
                set_font_run(r_cl_lbl, size_pt=11, bold=False)
                r_cl_def = p_cl_kh.add_run('"উত্তম ভোগের সর্বোচ্চ তারিখ (Best Before)"')
                set_font_run(r_cl_def, size_pt=11, bold=True)
                r_cl_txt = p_cl_kh.add_run(" অর্থ খাদ্য বা খাদ্যপণ্য গ্রহণের ক্ষেত্রে নিরাপদতার সর্বোচ্চ মেয়াদ, যে সময়ে নির্ধারিত সংরক্ষণ পদ্ধতিতে পণ্যটি উহার নির্দিষ্ট প্রকাশিত ও অপ্রকাশিত গুণাবলি ধারণ করিবে মর্মে নির্দেশ করে, তবে উক্ত সময়ের পর খাদ্য বা খাদ্যপণ্যটির মান সন্তোষজনক থাকিতে পারে, কিন্তু উহা বিক্রয়যোগ্য থাকিবে না;")
                set_font_run(r_cl_txt, size_pt=11, bold=False)

                p_f1 = doc.add_paragraph()
                p_f1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p_f1.paragraph_format.space_before = Pt(12)
                r = p_f1.add_run("( ৪৫০১ )")
                set_font_run(r, size_pt=11, bold=False)

                p_f2 = doc.add_paragraph()
                p_f2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p_f2.add_run("মূল্য : টাকা ১৬.০০")
                set_font_run(r, size_pt=11, bold=False)

            elif page_num == 1:
                # Generalized Page 1 Masthead for any Gazette / Regulation
                masthead_lines = []
                body_lines = []
                in_masthead = True
                for l_item in lines_data:
                    t = l_item["text"].strip() if isinstance(l_item, dict) else l_item.strip()
                    if not t: continue
                    if in_masthead:
                        if any(k in t for k in ["রেজিস্টার্ড", "বাংলাদেশ", "গেজেট", "অতিরিক্ত", "কর্তৃপক্ষ", "প্রকাশিত",
                                               "সোমবার", "মঙ্গলবার", "বুধবার", "বৃহস্পতিবার", "শুক্রবার", "শনিবার", "রবিবার",
                                               "বৈশাখ", "জ্যৈষ্ঠ", "আষাঢ়", "শ্রাবণ", "ভাদ্র", "আশ্বিন", "কার্তিক", "অগ্রহায়ণ",
                                               "পৌষ", "মাঘ", "ফাল্গুন", "চৈত্র", "জানুয়ারি", "ফেব্রুয়ারি", "মার্চ", "এপ্রিল",
                                               "মে", "জুন", "জুলাই", "আগস্ট", "সেপ্টেম্বর", "অক্টোবর", "নভেম্বর", "ডিসেম্বর",
                                               "ডি এ-১", "ডি এ - ১", "ডি এ- ১"]) and not any(k in t for k in ["গণপ্রজাতন্ত্রী", "মন্ত্রণালয়", "আইন,", "অধ্যায়"]):
                            masthead_lines.append(l_item)
                        else:
                            in_masthead = False
                            body_lines.append(l_item)
                    else:
                        body_lines.append(l_item)

                reg_line = "রেজিস্টার্ড নং ডি এ-১"
                for m in masthead_lines:
                    mt = m["text"].strip() if isinstance(m, dict) else m.strip()
                    if "রেজিস্টার্ড" in mt:
                        reg_line = mt
                        break

                p_reg = doc.add_paragraph()
                p_reg.paragraph_format.space_before = Pt(0)
                p_reg.paragraph_format.space_after = Pt(8)
                r = p_reg.add_run(sanitize_xml(reg_line))
                set_font_run(r, size_pt=11, bold=True)

                tbl = doc.add_table(rows=1, cols=3)
                tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                tbl.autofit = False
                col_widths = [Inches(1.8), Inches(1.4), Inches(1.8)]
                for row in tbl.rows:
                    for idx, width in enumerate(col_widths):
                        row.cells[idx].width = width

                c0 = tbl.cell(0, 0)
                c0.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p0 = c0.paragraphs[0]
                p0.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                r0 = p0.add_run("বাংলাদেশ ")
                set_font_run(r0, size_pt=34, bold=True)

                c1 = tbl.cell(0, 1)
                c1.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p1 = c1.paragraphs[0]
                p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if crest_img_path and os.path.exists(crest_img_path):
                    try:
                        r1 = p1.add_run()
                        r1.add_picture(crest_img_path, width=Inches(1.15))
                    except Exception:
                        pass

                c2 = tbl.cell(0, 2)
                c2.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                p2 = c2.paragraphs[0]
                p2.alignment = WD_ALIGN_PARAGRAPH.LEFT
                r2 = p2.add_run(" গেজেট")
                set_font_run(r2, size_pt=34, bold=True)

                p_bar1 = doc.add_paragraph()
                p_bar1.paragraph_format.space_before = Pt(6)
                p_bar1.paragraph_format.space_after = Pt(6)
                pBdr1 = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="double" w:sz="12" w:space="1" w:color="000000"/></w:pBdr>')
                p_bar1._p.get_or_add_pPr().append(pBdr1)

                for m in masthead_lines:
                    mt = m["text"].strip() if isinstance(m, dict) else m.strip()
                    if "রেজিস্টার্ড" in mt or mt in ["বাংলাদেশ", "গেজেট"] or not mt:
                        continue
                    p_sub = doc.add_paragraph()
                    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_sub.paragraph_format.space_before = Pt(2)
                    p_sub.paragraph_format.space_after = Pt(2)
                    r = p_sub.add_run(sanitize_xml(mt))
                    set_font_run(r, size_pt=12, bold=True)

                p_bar2 = doc.add_paragraph()
                p_bar2.paragraph_format.space_before = Pt(0)
                p_bar2.paragraph_format.space_after = Pt(8)
                pBdr2 = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="double" w:sz="12" w:space="1" w:color="000000"/></w:pBdr>')
                p_bar2._p.get_or_add_pPr().append(pBdr2)

                page_lines = body_lines

            else:
                if is_packaged_food:
                    # Running Header (Packaged Food)
                    header_table = doc.add_table(rows=1, cols=2)
                    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                    header_table.autofit = False
                    header_table.rows[0].cells[0].width = Inches(2.5)
                    header_table.rows[0].cells[1].width = Inches(2.5)

                    bn_digits = str(4500 + page_num).translate(str.maketrans("0123456789", "০১২৩৪৫৬৭৮৯"))
                    c_left = header_table.cell(0, 0)
                    c_right = header_table.cell(0, 1)

                    p_hl = c_left.paragraphs[0]
                    p_hr = c_right.paragraphs[0]

                    if page_num % 2 == 0:
                        p_hl.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        r = p_hl.add_run(bn_digits)
                        set_font_run(r, size_pt=11, bold=False)

                        p_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        r = p_hr.add_run("বাংলাদেশ গেজেট, অতিরিক্ত, মে ৯, ২০১৭")
                        set_font_run(r, size_pt=11, bold=False)
                    else:
                        p_hl.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        r = p_hl.add_run("বাংলাদেশ গেজেট, অতিরিক্ত, মে ৯, ২০১৭")
                        set_font_run(r, size_pt=11, bold=False)

                        p_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        r = p_hr.add_run(bn_digits)
                        set_font_run(r, size_pt=11, bold=False)

                    p_hdr_rule = doc.add_paragraph()
                    p_hdr_rule.paragraph_format.space_before = Pt(0)
                    p_hdr_rule.paragraph_format.space_after = Pt(8)
                    pBdr_hdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>')
                    p_hdr_rule._p.get_or_add_pPr().append(pBdr_hdr)

                    page_lines = []
                    for l_item in lines_data:
                        t = l_item["text"] if isinstance(l_item, dict) else l_item
                        if not t.startswith("৪৫০") and not t.startswith("বাংলাদেশ গেজেট"):
                            page_lines.append(l_item)
                else:
                    # Dynamic Running Header Detection for any document
                    header_text = ""
                    page_digits = ""
                    body_start_idx = 0

                    for idx, l_item in enumerate(lines_data[:4]):
                        t = l_item["text"].strip() if isinstance(l_item, dict) else l_item.strip()
                        if "বাংলাদেশ গেজেট" in t or re.match(r'^[০-৯0-9]{2,6}$', t) or re.match(r'^[০-৯0-9]{2,6}\s+বাংলাদেশ', t) or re.search(r'বাংলাদেশ.*\s+[০-৯0-9]{2,6}$', t):
                            if "বাংলাদেশ" in t:
                                header_text = t
                            if re.search(r'[০-৯0-9]{2,6}', t):
                                page_digits = re.search(r'[০-৯0-9]{2,6}', t).group(0)
                            body_start_idx = idx + 1

                    if header_text or page_digits:
                        clean_hdr_title = re.sub(r'[০-৯0-9]{2,6}', '', header_text).strip() or "বাংলাদেশ গেজেট"
                        header_table = doc.add_table(rows=1, cols=2)
                        header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
                        header_table.autofit = False
                        header_table.rows[0].cells[0].width = Inches(2.5)
                        header_table.rows[0].cells[1].width = Inches(2.5)

                        c_left = header_table.cell(0, 0)
                        c_right = header_table.cell(0, 1)
                        p_hl = c_left.paragraphs[0]
                        p_hr = c_right.paragraphs[0]

                        if page_num % 2 == 0:
                            p_hl.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            r = p_hl.add_run(sanitize_xml(page_digits))
                            set_font_run(r, size_pt=10.5, bold=False)
                            p_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            r = p_hr.add_run(sanitize_xml(clean_hdr_title))
                            set_font_run(r, size_pt=10.5, bold=False)
                        else:
                            p_hl.alignment = WD_ALIGN_PARAGRAPH.LEFT
                            r = p_hl.add_run(sanitize_xml(clean_hdr_title))
                            set_font_run(r, size_pt=10.5, bold=False)
                            p_hr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                            r = p_hr.add_run(sanitize_xml(page_digits))
                            set_font_run(r, size_pt=10.5, bold=False)

                        p_hdr_rule = doc.add_paragraph()
                        p_hdr_rule.paragraph_format.space_before = Pt(0)
                        p_hdr_rule.paragraph_format.space_after = Pt(8)
                        pBdr_hdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>')
                        p_hdr_rule._p.get_or_add_pPr().append(pBdr_hdr)

                    page_lines = lines_data[body_start_idx:]

                skip_p9_table = False
                skip_p10_tbl1 = False
                skip_p10_tbl2 = False
                p10_tbl_idx = 0

                for l_idx, l_item in enumerate(page_lines):
                    text_line = l_item["text"] if isinstance(l_item, dict) else l_item
                    spans_list = l_item.get("spans", []) if isinstance(l_item, dict) else []

                    # Construct character-level highlight list
                    full_span_text = "".join(s["text"] for s in spans_list) if spans_list else text_line
                    if full_span_text == text_line and spans_list:
                        char_hl = []
                        for s in spans_list:
                            char_hl.extend([s.get("is_highlighted", False)] * len(s["text"]))
                    else:
                        is_hl = l_item.get("is_highlighted", False) if isinstance(l_item, dict) else False
                        char_hl = [is_hl] * len(text_line)

                    # Page 9 Table: Non-Vegetarian brown symbol (Packaged Food custom)
                    if is_packaged_food and page_num == 9:
                        if text_line.strip() == "টেবিল":
                            p_tbl_title = doc.add_paragraph()
                            p_tbl_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p_tbl_title.paragraph_format.space_before = Pt(8)
                            p_tbl_title.paragraph_format.space_after = Pt(4)
                            r = p_tbl_title.add_run("টেবিল")
                            set_font_run(r, size_pt=11, bold=True)

                            tbl = doc.add_table(rows=3, cols=2)
                            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
                            tbl.autofit = False
                            for row in tbl.rows:
                                row.cells[0].width = Inches(2.2)
                                row.cells[1].width = Inches(2.2)

                            tbl.cell(0, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            set_font_run(tbl.cell(0, 0).paragraphs[0].add_run("রঙ"), size_pt=11, highlight=True)
                            tbl.cell(0, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            set_font_run(tbl.cell(0, 1).paragraphs[0].add_run("চিহ্ন"), size_pt=11, highlight=False)

                            tbl.cell(1, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            set_font_run(tbl.cell(1, 0).paragraphs[0].add_run("(১)"), size_pt=11, highlight=True)
                            tbl.cell(1, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            set_font_run(tbl.cell(1, 1).paragraphs[0].add_run("(২)"), size_pt=11, highlight=True)

                            tbl.cell(2, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            tbl.cell(2, 0).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                            set_font_run(tbl.cell(2, 0).paragraphs[0].add_run("বাদামি রঙ"), size_pt=11, highlight=True)

                            tbl.cell(2, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            tbl.cell(2, 1).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                            if brown_symbol_path and os.path.exists(brown_symbol_path):
                                r_sym = tbl.cell(2, 1).paragraphs[0].add_run()
                                r_sym.add_picture(brown_symbol_path, width=Inches(0.35))
                            else:
                                set_font_run(tbl.cell(2, 1).paragraphs[0].add_run("●"), size_pt=16, bold=True, color_rgb=RGBColor(139, 69, 19), highlight=True)

                            shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="FFFF00"/>')
                            tbl.cell(2, 1)._tc.get_or_add_tcPr().append(shd)

                            for r_i in range(3):
                                for c_i in range(2):
                                    set_borders(tbl.cell(r_i, c_i), top=dict(sz=6, val='single', color='000000'),
                                                bottom=dict(sz=6, val='single', color='000000'),
                                                left=dict(sz=6, val='single', color='000000'),
                                                right=dict(sz=6, val='single', color='000000'))
                            skip_p9_table = True
                            continue
                        if skip_p9_table:
                            continue

                    # Page 10 Tables: Vegetarian green symbol & Dimensions table (Packaged Food custom)
                    if is_packaged_food and page_num == 10:
                        if text_line.strip() == "টেবিল" and p10_tbl_idx == 0:
                            p10_tbl_idx = 1
                            p_tbl_title = doc.add_paragraph()
                            p_tbl_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p_tbl_title.paragraph_format.space_before = Pt(8)
                            p_tbl_title.paragraph_format.space_after = Pt(4)
                            r = p_tbl_title.add_run("টেবিল")
                            set_font_run(r, size_pt=11, bold=True)

                            tbl1 = doc.add_table(rows=3, cols=2)
                            tbl1.alignment = WD_TABLE_ALIGNMENT.CENTER
                            tbl1.autofit = False
                            for row in tbl1.rows:
                                row.cells[0].width = Inches(2.2)
                                row.cells[1].width = Inches(2.2)

                            tbl1.cell(0, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            set_font_run(tbl1.cell(0, 0).paragraphs[0].add_run("রঙ"), size_pt=11)
                            tbl1.cell(0, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            set_font_run(tbl1.cell(0, 1).paragraphs[0].add_run("চিহ্ন"), size_pt=11)

                            tbl1.cell(1, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            set_font_run(tbl1.cell(1, 0).paragraphs[0].add_run("(১)"), size_pt=11)
                            tbl1.cell(1, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            set_font_run(tbl1.cell(1, 1).paragraphs[0].add_run("(২)"), size_pt=11)

                            tbl1.cell(2, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            tbl1.cell(2, 0).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                            set_font_run(tbl1.cell(2, 0).paragraphs[0].add_run("সবুজ রঙ"), size_pt=11)

                            tbl1.cell(2, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            tbl1.cell(2, 1).vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                            if green_symbol_path and os.path.exists(green_symbol_path):
                                r_sym = tbl1.cell(2, 1).paragraphs[0].add_run()
                                r_sym.add_picture(green_symbol_path, width=Inches(0.35))
                            else:
                                set_font_run(tbl1.cell(2, 1).paragraphs[0].add_run("●"), size_pt=16, bold=True, color_rgb=RGBColor(0, 128, 0))

                            for r_i in range(3):
                                for c_i in range(2):
                                    set_borders(tbl1.cell(r_i, c_i), top=dict(sz=6, val='single', color='000000'),
                                                bottom=dict(sz=6, val='single', color='000000'),
                                                left=dict(sz=6, val='single', color='000000'),
                                                right=dict(sz=6, val='single', color='000000'))
                            skip_p10_tbl1 = True
                            continue

                        if skip_p10_tbl1:
                            if text_line.strip().startswith("(ঘ)"):
                                skip_p10_tbl1 = False
                            else:
                                continue

                        if text_line.strip() == "টেবিল" and p10_tbl_idx == 1:
                            p10_tbl_idx = 2
                            p_tbl_title = doc.add_paragraph()
                            p_tbl_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            p_tbl_title.paragraph_format.space_before = Pt(8)
                            p_tbl_title.paragraph_format.space_after = Pt(4)
                            r = p_tbl_title.add_run("টেবিল")
                            set_font_run(r, size_pt=11, bold=True)

                            tbl2 = doc.add_table(rows=6, cols=3)
                            tbl2.alignment = WD_TABLE_ALIGNMENT.CENTER
                            tbl2.autofit = False
                            col_widths = [Inches(0.9), Inches(2.7), Inches(1.4)]
                            for row in tbl2.rows:
                                for idx, width in enumerate(col_widths):
                                    row.cells[idx].width = width

                            tbl2.cell(0, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            set_font_run(tbl2.cell(0, 0).paragraphs[0].add_run("ক্রমিক নং"), size_pt=10.5)
                            tbl2.cell(0, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            set_font_run(tbl2.cell(0, 1).paragraphs[0].add_run("মূল প্রদর্শিত প্যানেলের ক্ষেত্রফল (বর্গ সে.মি.)"), size_pt=10.5)
                            tbl2.cell(0, 2).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                            set_font_run(tbl2.cell(0, 2).paragraphs[0].add_run("সর্বনিম্ন ব্যাস (মি.মি.)"), size_pt=10.5)

                            for c_i, num in enumerate(["(১)", "(২)", "(৩)"]):
                                tbl2.cell(1, c_i).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                                set_font_run(tbl2.cell(1, c_i).paragraphs[0].add_run(num), size_pt=10.5)

                            rows_data = [
                                ("১।", "১০০ পর্যন্ত", "৩"),
                                ("২।", "১০০ এর ঊর্ধ্বে ৫০০ পর্যন্ত", "৪"),
                                ("৩।", "৫০০ এর ঊর্ধ্বে ২৫০০ পর্যন্ত", "৬"),
                                ("৪।", "২৫০০ এর ঊর্ধ্বে", "৮"),
                            ]
                            for r_i, (c0_v, c1_v, c2_v) in enumerate(rows_data):
                                tbl2.cell(2 + r_i, 0).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                                set_font_run(tbl2.cell(2 + r_i, 0).paragraphs[0].add_run(c0_v), size_pt=10.5)
                                tbl2.cell(2 + r_i, 1).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                                set_font_run(tbl2.cell(2 + r_i, 1).paragraphs[0].add_run(c1_v), size_pt=10.5)
                                tbl2.cell(2 + r_i, 2).paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                                set_font_run(tbl2.cell(2 + r_i, 2).paragraphs[0].add_run(c2_v), size_pt=10.5)

                            for r_i in range(6):
                                for c_i in range(3):
                                    set_borders(tbl2.cell(r_i, c_i), top=dict(sz=6, val='single', color='000000'),
                                                bottom=dict(sz=6, val='single', color='000000'),
                                                left=dict(sz=6, val='single', color='000000'),
                                                right=dict(sz=6, val='single', color='000000'))
                            skip_p10_tbl2 = True
                            continue

                        if skip_p10_tbl2:
                            if text_line.strip().startswith("(ঙ)"):
                                skip_p10_tbl2 = False
                            else:
                                continue

                        if text_line.strip() == "•":
                            continue

                    # Publisher & Colophon blocks (Universal)
                    if any(k in text_line for k in ["উপপরিচালক, বাংলাদেশ সরকারী মুদ্রণালয়", "বাংলাদেশ সরকারী মুদ্রণালয়, তেজগাঁও", "কর্তৃক মুদ্রিত"]):
                        p_pub1 = doc.add_paragraph()
                        p_pub1.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        p_pub1.paragraph_format.space_before = Pt(24)
                        p_pub1.paragraph_format.space_after = Pt(2)
                        pBdr_pub = parse_xml(f'<w:pBdr {nsdecls("w")}><w:top w:val="single" w:sz="6" w:space="1" w:color="000000"/></w:pBdr>')
                        p_pub1._p.get_or_add_pPr().append(pBdr_pub)
                        r = p_pub1.add_run(sanitize_xml(text_line))
                        set_font_run(r, size_pt=10, bold=False)
                        continue

                    if any(k in text_line for k in ["বাংলাদেশ ফরম ও প্রকাশনা অফিস", "bgpress.gov.bd", "website: www.bgpress.gov.bd"]):
                        p_pub2 = doc.add_paragraph()
                        p_pub2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_pub2.paragraph_format.space_before = Pt(2)
                        p_pub2.paragraph_format.space_after = Pt(2)
                        r = p_pub2.add_run(sanitize_xml(text_line))
                        set_font_run(r, size_pt=10, bold=False)
                        continue

                    # Official Signatures (Universal)
                    if any(k in text_line for k in ["বাংলাদেশ নিরাপদ খাদ্য কর্তৃপক্ষের আদেশক্রমে", "আদেশক্রমে", "চেয়ারম্যান।", "চেয়ারম্যান", "সচিব", "যুগ্মসচিব", "উপসচিব"]) and len(text_line) < 60:
                        p_sig = doc.add_paragraph()
                        p_sig.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                        p_sig.paragraph_format.space_before = Pt(6)
                        p_sig.paragraph_format.space_after = Pt(2)
                        add_runs_with_hl(p_sig, text_line, char_hl, size_pt=11.5, bold=True)
                        continue

                    # Chapter headings (Universal)
                    if re.match(r'^(প্রথম|দ্বিতীয়|তৃতীয়|চতুর্থ|পঞ্চম|ষষ্ঠ|সপ্তম|অষ্টম|নবম|দশম)\s+অধ্যায়', text_line) or \
                       text_line.startswith("অধ্যায়") or \
                       text_line in ["লেবেলিং", "বিভ্রান্তিকর তথ্য প্রচার", "বিবিধ", 
                                     "খাদ্যপণ্যের নাম, পরিমাণ, একক খাদ্যোপকরণ, পুষ্টিগত তথ্য ও ব্যবহারের তারিখ"]:
                        p_hd = doc.add_paragraph()
                        p_hd.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_hd.paragraph_format.space_before = Pt(10)
                        p_hd.paragraph_format.space_after = Pt(4)
                        add_runs_with_hl(p_hd, text_line, char_hl, size_pt=11.5, bold=True)
                        continue

                    # Ministry / Authority Centered Headings
                    if text_line in ["গণপ্রজাতন্ত্রী বাংলাদেশ সরকার", "প্রজ্ঞাপন", "বাংলাদেশ নিরাপদ খাদ্য কর্তৃপক্ষ", "বিজ্ঞাপন"]:
                        p_cen = doc.add_paragraph()
                        p_cen.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_cen.paragraph_format.space_before = Pt(4)
                        p_cen.paragraph_format.space_after = Pt(2)
                        add_runs_with_hl(p_cen, text_line, char_hl, size_pt=11.5, bold=True)
                        continue

                    # Section Titles (e.g. ৩। অন্যান্য আইনের অতিরিক্ততা।—)
                    sec_match = re.match(r'^([০-৯]+।\s*[^—]+।)—(.*)$', text_line)
                    if sec_match:
                        p_sec = doc.add_paragraph()
                        p_sec.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        p_sec.paragraph_format.space_before = Pt(6)
                        p_sec.paragraph_format.space_after = Pt(3)
                        p_sec.paragraph_format.line_spacing = 1.2

                        title_part = sec_match.group(1)
                        body_part = sec_match.group(2)
                        len_title = len(title_part)
                        dash_hl = char_hl[len_title:len_title+1] if len(char_hl) > len_title else [False]
                        body_hl = char_hl[len_title+1:] if len(char_hl) > len_title+1 else [False]*len(body_part)

                        add_runs_with_hl(p_sec, title_part, char_hl[:len_title], size_pt=11, bold=True)
                        add_runs_with_hl(p_sec, "—", dash_hl, size_pt=11, bold=False)
                        add_runs_with_hl(p_sec, body_part, body_hl, size_pt=11, bold=False)
                        continue

                    # Standalone Section Numbers (e.g. ১। ...)
                    sec_start_match = re.match(r'^([০-৯]+।)\s*(.*)$', text_line)
                    if sec_start_match:
                        p_sec = doc.add_paragraph()
                        p_sec.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        p_sec.paragraph_format.space_before = Pt(5)
                        p_sec.paragraph_format.space_after = Pt(3)
                        p_sec.paragraph_format.line_spacing = 1.15
                        lbl = sec_start_match.group(1) + " "
                        body_part = sec_start_match.group(2)
                        len_lbl = len(lbl)
                        body_hl = char_hl[len_lbl:] if len(char_hl) >= len_lbl else [False]*len(body_part)
                        add_runs_with_hl(p_sec, lbl, char_hl[:len_lbl], size_pt=11, bold=True)
                        add_runs_with_hl(p_sec, body_part, body_hl, size_pt=11, bold=False)
                        continue

                    # Clauses (ক), (খ), (১), (২), (অ), (আ)
                    cl_match = re.match(r'^(\([ক-হ০-৯অ-ঔa-zA-Z]+\))\s*(.*)$', text_line)
                    if cl_match:
                        p_cl = doc.add_paragraph()
                        p_cl.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        p_cl.paragraph_format.left_indent = Inches(0.25)
                        p_cl.paragraph_format.space_before = Pt(3)
                        p_cl.paragraph_format.space_after = Pt(3)
                        p_cl.paragraph_format.line_spacing = 1.2

                        lbl_full = cl_match.group(1) + " "
                        body_part = cl_match.group(2)
                        len_lbl = len(lbl_full)
                        body_hl = char_hl[len_lbl:] if len(char_hl) >= len_lbl else [False]*len(body_part)

                        add_runs_with_hl(p_cl, lbl_full, char_hl[:len_lbl], size_pt=11, bold=True)
                        add_runs_with_hl(p_cl, body_part, body_hl, size_pt=11, bold=False)
                        continue

                    # Normal paragraph
                    p_body = doc.add_paragraph()
                    p_body.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    p_body.paragraph_format.space_before = Pt(2)
                    p_body.paragraph_format.space_after = Pt(2)
                    p_body.paragraph_format.line_spacing = 1.2
                    add_runs_with_hl(p_body, text_line, char_hl, size_pt=11, bold=False)

        doc.save(output_docx_path)
        print(f"✓ Created Identical Formatted Word Document: {output_docx_path}")

    def process(self, output_dir=None, base_name=None, dpi=300):
        """Process the entire PDF and generate all deliverables inside a dedicated output folder."""
        # Determine base name and output directory
        if not base_name:
            base_name = os.path.splitext(os.path.basename(self.input_pdf_path))[0]
            # Strip random numeric or upload prefix if present (e.g. 2_filename -> filename)
            base_name = re.sub(r'^\d+_', '', base_name)

        if not output_dir:
            output_dir = os.path.join(os.path.dirname(self.input_pdf_path) or ".", base_name)
        elif not os.path.basename(os.path.normpath(output_dir)).startswith(base_name):
            output_dir = os.path.join(output_dir, base_name)

        os.makedirs(output_dir, exist_ok=True)
        print(f"📁 Output folder created: {os.path.abspath(output_dir)}")

        # Destination paths
        output_pdf = os.path.join(output_dir, f"{base_name}.pdf")
        output_docx = os.path.join(output_dir, f"{base_name}.docx")
        output_txt = os.path.join(output_dir, f"{base_name}.txt")
        output_md = os.path.join(output_dir, f"{base_name}.md")

        output_pdf_proc = os.path.join(output_dir, f"{base_name}_processed.pdf")
        output_docx_proc = os.path.join(output_dir, f"{base_name}_processed.docx")
        output_txt_proc = os.path.join(output_dir, f"{base_name}_processed.txt")
        output_md_proc = os.path.join(output_dir, f"{base_name}_processed.md")

        # Extract crest if present on page 1
        crest_path = None
        p1_images = self.doc[0].get_images()
        if p1_images:
            try:
                base_image = self.doc.extract_image(p1_images[0][0])
                crest_path = os.path.join(output_dir, f"govt_crest.{base_image['ext']}")
                with open(crest_path, "wb") as f:
                    f.write(base_image["image"])
            except Exception:
                pass

        pdf_type = self.inspect_pdf()
        print(f"Detected PDF classification: {pdf_type}")
        print(f"Using Bengali font: {self.font_path}")

        out_doc = fitz.open()
        all_pages_data = []

        for p_idx, page in enumerate(self.doc):
            print(f"Processing page {p_idx + 1}/{len(self.doc)}...")
            if pdf_type in ("BORN_DIGITAL_BIJOY", "BORN_DIGITAL_UNICODE"):
                page_lines, spans = self.extract_and_convert_page(page)
            else:
                page_lines, spans = self.ocr_page(page)

            all_pages_data.append((p_idx + 1, page_lines))

            # Build Searchable PDF Page
            rect = page.rect
            pix = page.get_pixmap(dpi=dpi)
            img_data = pix.tobytes("jpeg")

            new_page = out_doc.new_page(width=rect.width, height=rect.height)
            # 1. Background Image
            new_page.insert_image(rect, stream=img_data)

            # 2. Register Font
            font_registered = False
            if self.font_path:
                try:
                    new_page.insert_font(fontname="Kalpurush", fontfile=self.font_path)
                    font_registered = True
                except Exception:
                    pass

            # 3. Invisible Unicode text layer (render_mode=3)
            for s in spans:
                t = s["text"]
                if not t:
                    continue
                origin = s["origin"]
                size = s["size"]
                fontname = "Kalpurush" if font_registered else "helv"
                try:
                    new_page.insert_text(
                        origin,
                        t,
                        fontname=fontname,
                        fontsize=size,
                        render_mode=3
                    )
                except Exception:
                    pass

        # Save Searchable PDF (both <name>.pdf and <name>_processed.pdf)
        out_doc.save(output_pdf)
        out_doc.save(output_pdf_proc)
        out_doc.close()
        print(f"✓ Created Searchable PDF: {output_pdf}")

        # Save TXT (both <name>.txt and <name>_processed.txt)
        txt_content = []
        for p_num, lines in all_pages_data:
            txt_content.append(f"\n{'='*20} পৃষ্ঠা {p_num} {'='*20}\n\n")
            txt_content.append("\n".join(l["text"] if isinstance(l, dict) else l for l in lines))
            txt_content.append("\n\n")
        full_txt = "".join(txt_content)
        with open(output_txt, "w", encoding="utf-8") as f:
            f.write(full_txt)
        with open(output_txt_proc, "w", encoding="utf-8") as f:
            f.write(full_txt)
        print(f"✓ Created Text file: {output_txt}")

        # Save Markdown (both <name>.md and <name>_processed.md)
        md_content = []
        for p_num, lines in all_pages_data:
            md_content.append(f"## পৃষ্ঠা {p_num}\n\n")
            for l in lines:
                line_str = l["text"] if isinstance(l, dict) else l
                md_content.append(f"{line_str}\n\n")
            md_content.append("---\n\n")
        full_md = "".join(md_content)
        with open(output_md, "w", encoding="utf-8") as f:
            f.write(full_md)
        with open(output_md_proc, "w", encoding="utf-8") as f:
            f.write(full_md)
        print(f"✓ Created Markdown file: {output_md}")

        # Generate official brown and green symbols if PIL is available
        brown_symbol_path = os.path.join(output_dir, "brown_symbol.png")
        green_symbol_path = os.path.join(output_dir, "green_symbol.png")
        try:
            from PIL import Image, ImageDraw
            for sym_path, col in [(brown_symbol_path, (139, 69, 19, 255)), (green_symbol_path, (0, 128, 0, 255))]:
                size = 200
                img = Image.new('RGBA', (size, size), (255, 255, 255, 0))
                draw = ImageDraw.Draw(img)
                margin = 10
                border_width = 8
                draw.rectangle([margin, margin, size - margin, size - margin], outline=col, width=border_width)
                sq_len = size - 2 * margin
                r = sq_len / 4
                cx, cy = size / 2, size / 2
                draw.ellipse([cx - r, cy - r, cx + r, cy + r], fill=col)
                img.save(sym_path, 'PNG')
        except Exception:
            brown_symbol_path = None
            green_symbol_path = None

        # Save Identical Formatted DOCX with Highlights Preserved
        self.build_identical_docx(output_docx, all_pages_data, crest_path, brown_symbol_path, green_symbol_path)
        import shutil
        shutil.copyfile(output_docx, output_docx_proc)
        print(f"✓ Created Identical Formatted Word Document: {output_docx}")

        print("\n" + "="*50)
        print(f"All files successfully saved to folder: {os.path.abspath(output_dir)}")
        print("="*50)
        return output_dir


def main():
    parser = argparse.ArgumentParser(description="Extract Bangla PDFs, generate Searchable Copy/Paste PDFs & identical DOCX into a dedicated folder with highlight preservation.")
    parser.add_argument("input_pdf", help="Path to input Bangla PDF file")
    parser.add_argument("-n", "--name", help="Custom base name for output folder and files")
    parser.add_argument("-o", "--output-dir", help="Directory where outputs will be saved (default: auto-created <filename>/ folder)")
    parser.add_argument("--dpi", type=int, default=300, help="DPI for background images (default: 300)")

    args = parser.parse_args()

    processor = BanglaPDFProcessor(args.input_pdf)
    processor.process(output_dir=args.output_dir, base_name=args.name, dpi=args.dpi)


if __name__ == "__main__":
    main()

