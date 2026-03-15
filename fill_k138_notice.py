# fill_k138_notice.py
# Version 0.2.3 (2025/12/08)
# Same base coordinates; clarified fine-tuning via GLOBAL_DX/DY and OFFSETS, includes Seizing Officer.

#region fill_k138_notice.py fill K138 form with dummy data

import io
import csv
import re
from PyPDF2 import PdfReader, PdfWriter, PageObject
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth

try:
    import fitz  # PyMuPDF
    HAVE_PYMUPDF = True
except ImportError:
    HAVE_PYMUPDF = False

# User preference: render K138 values in uppercase.
FORCE_ALL_CAPS = True

# ---------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------

TEMPLATE_PDF = r"K138 Stupefiant-Others - TEMPLATE-dummy.pdf"
OUTPUT_PDF   = r"K138.pdf"
# Only the first 4 K138 copies should be populated.
MAX_FILLED_PAGES = 4

# Form type notices to append to description block
CANNABIS_NOTICE = (
    "\n\n(Please note: Under the Cannabis Act, the above item is illegal to transport, import or export across Canada's borders, including by mail or courier. For more information please visit this website: https://www.canada.ca/en/services/health/campaigns/cannabis/border.html)\n"
    "-------------------------------------------------------------------------------------------------------\n"
    "(Attention : La Loi sur le cannabis interdit d'importer l'article ci-dessus au Canada ou de l'en exporter, y compris par la poste ou par messagerie. Pour en savoir plus : https://www.canada.ca/fr/services/sante/campagnes/cannabis/frontiere.html)"
)

KNIFE_NOTICE = (
    "\n\nCENTRIFUGAL KNIFE AS PER MEMORANDUM D19-13-2\n"
    "COUTEAU CENTRIFUGE AU MÉMORANDUM D19-13-2"
)

# Global nudge (if everything is slightly off in one direction)
# 1 point â‰ˆ 0.35 mm (2 mm â‰ˆ 6 points).
# Positive DX = move right; negative = move left.
# Positive DY = move up;   negative = move down.

# To find these values, Dmitry used k138_offset_tuner_app.py # Version 0.1.2 (2025/12/08)
GLOBAL_DX = -20
GLOBAL_DY = -40

# Per-field offsets (Dmitry 2026/02): fine-tune placement on K138 form
OFFSETS = {
    "notice_to":       (24, 20),    # Extra right shift to avoid first-character clipping
    "notice_date":     (0, 20),
    "seizure_date":    (20, -20),   # was (20, 0)
    "seizure_year_L":  (0, -20),    # not used â€“ always "20" as in 2026
    "seizure_year_R":  (0, -20),    # use "26"
    "seizure_location":(-20, -62),  # Move a bit further up
    "description":     (-30, -30),  # Shift right to avoid clipped first character
    "seizing_officer": (120, 90),
}

# GLOBAL_DX = 0
# GLOBAL_DY = 0

# # Per-field extra offsets (for local fine-tuning)
# # Example: OFFSETS["notice_to"] = (-5, +3)  moves the address
# # 5 points left and 3 points up, on top of GLOBAL_DX/DY.
# OFFSETS = {
#     "notice_to":       (0,   0),
#     "notice_date":     (0,   0),
#     "seizure_date":    (0,   0),
#     "seizure_year_L":  (0,   0),
#     "seizure_year_R":  (0,   0),
#     "seizure_location":(0,   0),
#     "description":     (0,   0),
#     "seizing_officer": (0,   0),
# }

# Dummy data to fill (later you can load/replace from CSV)
dummy_data = {
    # NOTICE OF SEIZURE â€“ To:
    "notice_to": (
        "John Doe\n"
        "123 RUE DUMMY, APT 1234\n"
        "MONTREAL QC A1B 2C3"
    ),

    # AVIS DE SAISIE â€“ Date
    "notice_date": "2025-11-25",

    # This refers to goods seized/detained on ...
    "seizure_date_line": "16 NOVEMBER / 16 NOVEMBRE",
    "seizure_year_left": "20",
    "seizure_year_right": "25",

    # ... at
    "seizure_location": (
        "MONTREAL POSTAL FACILITY, ETC LEO-BLANCHETTE / POSTAL CUSTOMS"
    ),

    # Description block
    "description_block": (
        "INVENTORY NO / NO. D'INVENTAIRE: ABXXXYYYZZZCO\n\n"
        "DECLARED/DÉCLARÉ: N/A\n\n"
        "ITEM SEIZED / MARCHANDISE SAISIE: "
        "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX.\n\n"
        "SEIZURE NUMBER / NUMÉRO DE SAISIE: 3952-25-1234"
    ),

    # Seizing Officer (bottom of page)
    "seizing_officer": "12345",
}

# ---------------------------------------------------------------------
# Optional: future CSV loader
# ---------------------------------------------------------------------

def _repair_mojibake_text(s: str) -> str:
    """Best-effort fix for text like 'FÃ‰VRIER' -> 'FÉVRIER'."""
    out = str(s or "")
    for _ in range(2):
        if not any(ch in out for ch in ("\u00C3", "\u00C2", "\u00E2", "\u0102")):
            break
        candidate = ""
        for enc in ("latin1", "cp1252"):
            try:
                candidate = out.encode(enc).decode("utf-8")
                break
            except Exception:
                candidate = ""
        if not candidate or candidate == out:
            break
        out = candidate
    return out


def _clean_pdf_text(s: str) -> str:
    """Normalize text before drawing to PDF to avoid mojibake artifacts."""
    out = _repair_mojibake_text(s or "")
    out = re.sub(r"[ \t]+", " ", out)
    out = out.strip()
    return out.upper() if FORCE_ALL_CAPS else out


def _clean_pdf_multiline_text(s: str) -> str:
    """Normalize multiline text while preserving line breaks."""
    out = _repair_mojibake_text(s or "")
    out = out.replace("\r\n", "\n").replace("\r", "\n")
    out = re.sub(r"[ \t]+", " ", out)
    out = re.sub(r"[ \t]*\n[ \t]*", "\n", out)
    out = out.strip()
    return out.upper() if FORCE_ALL_CAPS else out


def _compose_description_block_from_fields(data: dict, keep_bottom_gap: bool = True) -> str:
    """
    Build K138 description section with stable spacing:
    inventory, declared, item, seizure number (contiguous lines).
    """
    declared_single = _clean_pdf_text(
        re.sub(
            r"\s+",
            " ",
            _repair_mojibake_text((data.get("description_declared") or "").replace("|", " ")),
        )
    )
    lines = [
        f"INVENTORY NO / NO. D'INVENTAIRE: {_clean_pdf_text(data.get('description_inventory', ''))}",
        f"DECLARED/DÉCLARÉ: {declared_single}",
        f"ITEM SEIZED / MARCHANDISE SAISIE: {_clean_pdf_multiline_text(data.get('description_item', ''))}",
        f"SEIZURE NUMBER / NUMÉRO DE SAISIE: {_clean_pdf_text(data.get('description_seizure_number', ''))}",
    ]
    block = "\n".join(lines)
    if keep_bottom_gap:
        block += "\n"
    legal = _clean_pdf_multiline_text(data.get("legal_notice", ""))
    if legal:
        block = f"{block}\n{legal}"
    return block
# Dmitry: replaced vertical to horizontal 
# def load_data_from_csv(csv_path: str) -> dict:
#     """
#     Load field values from a CSV file.
#     Expected format: one row, with columns matching keys in dummy_data
#     (e.g., notice_to, notice_date, seizure_date_line, etc.).
#     """
#     with open(csv_path, newline="", encoding="utf-8") as f:
#         reader = csv.DictReader(f)
#         for row in reader:
#             return row  # first row only
#     return {}

def load_data_from_csv(csv_path: str) -> dict:
    """
    Load key/value CSV where each row is: field,value
    """
    def _get_col(row: dict, col_name: str) -> str | None:
        """Read CSV column value with BOM/whitespace/case-tolerant header matching."""
        direct = row.get(col_name)
        if direct is not None:
            return direct
        target = col_name.strip().lower()
        for k, v in row.items():
            if k is None:
                continue
            k_norm = str(k).replace("\ufeff", "").strip().lower()
            if k_norm == target:
                return v
        return None

    data = {}
    with open(csv_path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            field_name = _get_col(row, "field")
            if not field_name:
                continue
            field_value = _get_col(row, "value")
            clean_field = _repair_mojibake_text(str(field_name).strip())
            clean_value = "" if field_value is None else _repair_mojibake_text(str(field_value))
            data[clean_field] = clean_value

    # Reconstruct multiline fields (use .get to avoid KeyError on missing keys)
    data["notice_to"] = (data.get("notice_to") or "").replace(" | ", "\n")
    data["description_block"] = _compose_description_block_from_fields(data, keep_bottom_gap=True)

    return data


# ---------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------

def detect_form_type(template_path: str) -> str:
    """
    Detect the form type from the template PDF filename.
    Returns: "Cannabis-Stupefiant", "Knives-Arms", or "Stupefiant-Others"
    """
    template_lower = template_path.lower()
    if "cannabis" in template_lower:
        return "Cannabis-Stupefiant"
    elif "knife" in template_lower or "knives" in template_lower or "arms" in template_lower:
        return "Knives-Arms"
    else:
        return "Stupefiant-Others"


def detect_description_box_dimensions(template_path: str) -> tuple[float, float, float, float] | None:
    """
    Detect the description box dimensions from the PDF template.
    Returns (x_left, y_bottom, x_right, y_top) in ReportLab coordinates (bottom-left origin),
    or None if detection fails.
    
    In PyMuPDF: origin is top-left, y=0 at top, y increases downward
    In ReportLab: origin is bottom-left, y=0 at bottom, y increases upward
    Conversion: reportlab_y = page_height - pymupdf_y
    """
    if not HAVE_PYMUPDF:
        return None
    
    try:
        doc = fitz.open(template_path)
        page = doc[0]
        page_height = page.rect.height
        
        # Method 1: Look for form fields/widgets (best method)
        # Try widgets first (more reliable)
        for widget in page.widgets():
            field_name = widget.field_name
            rect = widget.rect
            # Check if this is the description field (often named "C" or similar)
            # Description box should be large (width > 400, height > 100)
            box_width = rect.width
            box_height = rect.height
            if box_width > 400 and box_height > 100 and rect.y0 > 300 and rect.y0 < 600:
                # Convert from PyMuPDF (top-left origin) to ReportLab (bottom-left origin)
                x_left = rect.x0
                x_right = rect.x1
                # In PyMuPDF: rect.y0 is top (smaller y), rect.y1 is bottom (larger y)
                # In ReportLab: we want y_bottom (smaller) and y_top (larger)
                y_bottom_reportlab = page_height - rect.y1  # PyMuPDF bottom -> ReportLab bottom
                y_top_reportlab = page_height - rect.y0     # PyMuPDF top -> ReportLab top
                doc.close()
                print(f"Found form field '{field_name}': {rect}")
                print(f"  PyMuPDF coords: top={rect.y0:.1f}, bottom={rect.y1:.1f}")
                print(f"  ReportLab coords: y_bottom={y_bottom_reportlab:.1f}, y_top={y_top_reportlab:.1f}")
                return (x_left, y_bottom_reportlab, x_right, y_top_reportlab)
        
        # Also check annotations as fallback
        for annot in page.annots():
            if annot.type[1] == "Widget":  # Form field
                rect = annot.rect
                box_width = rect.width
                box_height = rect.height
                if box_width > 400 and box_height > 100 and rect.y0 > 300 and rect.y0 < 600:
                    x_left = rect.x0
                    x_right = rect.x1
                    y_bottom_reportlab = page_height - rect.y1
                    y_top_reportlab = page_height - rect.y0
                    doc.close()
                    print(f"Found annotation: {rect}")
                    return (x_left, y_bottom_reportlab, x_right, y_top_reportlab)
        
        # Method 2: Look for text that indicates description box location
        # Search for keywords like "ITEM SEIZED" or "MARCHANDISE SAISIE"
        words = page.get_text("words")
        # Find the specific "ITEM SEIZED / MARCHANDISE SAISIE" label (usually around y=440-460)
        # Look for "ITEM" and "SEIZED" that are close together horizontally
        item_seized_words = []
        for w in words:
            text = w[4].upper()
            if "ITEM" in text or ("SEIZED" in text and "/" not in text):
                # Check if there's a nearby "SEIZED" or "MARCHANDISE"
                x0, y0, x1, y1 = w[0], w[1], w[2], w[3]
                # The label should be in the middle-upper part of page (around y=440-470)
                if 400 < y0 < 500:
                    item_seized_words.append(w)
        
        if item_seized_words:
            # Use the bottom-most edge of these label words
            max_y1_pymupdf = max(w[3] for w in item_seized_words)  # Bottom edge of label
            min_x = min(w[0] for w in item_seized_words)
            max_x = max(w[2] for w in item_seized_words)
            
            # Description box typically starts below this text
            # In PyMuPDF: y increases downward, so larger y = lower on page
            box_top_pymupdf = max_y1_pymupdf + 25  # Just below keywords
            # Estimate box height: typically 250-300 points
            box_height = 280
            box_bottom_pymupdf = box_top_pymupdf + box_height
            
            # Convert to ReportLab coordinates
            # ReportLab: y=0 at bottom, y increases upward
            # PyMuPDF: y=0 at top, y increases downward
            # Conversion: reportlab_y = page_height - pymupdf_y
            y_top_reportlab = page_height - box_top_pymupdf    # Top of box in ReportLab
            y_bottom_reportlab = page_height - box_bottom_pymupdf  # Bottom of box in ReportLab
            
            # Estimate box extends from left margin to near right margin
            x_left = 85  # Left margin  
            x_right = 520  # Right edge of box (conservative estimate)
            
            # Ensure box doesn't extend beyond page
            if box_bottom_pymupdf > page_height:
                box_bottom_pymupdf = page_height - 20  # Leave small margin
                box_height = box_bottom_pymupdf - box_top_pymupdf
            
            # Convert to ReportLab coordinates
            y_top_reportlab = page_height - box_top_pymupdf    # Top of box in ReportLab
            y_bottom_reportlab = page_height - box_bottom_pymupdf  # Bottom of box in ReportLab
            
            doc.close()
            print(f"Detected from text keywords: 'ITEM SEIZED' label at PyMuPDF y={max_y1_pymupdf:.1f}")
            print(f"Box extends from PyMuPDF y={box_top_pymupdf:.1f} to {box_bottom_pymupdf:.1f} (height={box_height:.1f})")
            print(f"ReportLab coords: y_bottom={y_bottom_reportlab:.1f}, y_top={y_top_reportlab:.1f}")
            return (x_left, y_bottom_reportlab, x_right, y_top_reportlab)
        
        # Fallback: if we found desc_words but not item_seized_words, use all desc_words
        desc_keywords = ["ITEM", "SEIZED", "MARCHANDISE", "SAISIE"]
        desc_words = [w for w in words if any(kw.lower() in w[4].lower() for kw in desc_keywords)]
        
        if desc_words:
            # Filter to middle section (likely the label)
            mid_words = [w for w in desc_words if 400 < w[1] < 500]
            if mid_words:
                max_y1_pymupdf = max(w[3] for w in mid_words)
                box_top_pymupdf = max_y1_pymupdf + 25
                box_height = 280
                box_bottom_pymupdf = min(box_top_pymupdf + box_height, page_height - 20)
                
                y_top_reportlab = page_height - box_top_pymupdf
                y_bottom_reportlab = page_height - box_bottom_pymupdf
                x_left = 85
                x_right = 520
                
                doc.close()
                print(f"Detected using fallback method")
                return (x_left, y_bottom_reportlab, x_right, y_top_reportlab)
        
        doc.close()
        return None
    except Exception as e:
        print(f"Warning: Could not detect box dimensions: {e}")
        import traceback
        traceback.print_exc()
        return None


def wrap_text_measured(text: str, max_width_points: float, font_name: str = "Helvetica", font_size: int = 9) -> list[str]:
    """
    Wrap text using actual ReportLab text width measurement to ensure it fits within bounds.
    
    Args:
        text: Text to wrap
        max_width_points: Maximum width in points
        font_name: Font name (default "Helvetica")
        font_size: Font size (default 9)
    
    Returns:
        List of wrapped lines
    """
    lines = []
    words = text.split(' ')
    current_line = []
    
    for word in words:
        # Test if adding this word would exceed the width
        test_line = ' '.join(current_line + [word]) if current_line else word
        test_width = stringWidth(test_line, font_name, font_size)
        
        if test_width <= max_width_points:
            current_line.append(word)
        else:
            # Current line is full, save it
            if current_line:
                lines.append(' '.join(current_line))
            
            # Check if the word itself is longer than max width
            word_width = stringWidth(word, font_name, font_size)
            if word_width > max_width_points:
                # Word is too long, split it character by character
                char_line = ""
                for char in word:
                    test_char_line = char_line + char
                    if stringWidth(test_char_line, font_name, font_size) <= max_width_points:
                        char_line = test_char_line
                    else:
                        if char_line:
                            lines.append(char_line)
                        char_line = char
                current_line = [char_line] if char_line else []
            else:
                current_line = [word]
    
    if current_line:
        lines.append(' '.join(current_line))
    
    return lines if lines else [text]


def apply_offsets(x: float, y: float, key: str) -> tuple[float, float]:
    """
    Apply GLOBAL_DX/DY and per-field OFFSETS[key] to a base (x, y).
    """
    dx, dy = OFFSETS.get(key, (0, 0))
    return x + GLOBAL_DX + dx, y + GLOBAL_DY + dy


def _format_notice_date_for_display(raw_date: str) -> str:
    """Format notice date as DD MONTH YYYY when possible."""
    txt = _clean_pdf_text(raw_date or "")
    if not txt:
        return ""
    try:
        from datetime import datetime
        dt = datetime.strptime(txt, "%Y-%m-%d")
        return dt.strftime("%d %B %Y")
    except Exception:
        return txt


def _build_description_block_text(data: dict) -> str:
    """Return normalized multi-line description text as rendered on K138."""
    has_structured_fields = any(
        _clean_pdf_text(data.get(k, ""))
        for k in (
            "description_inventory",
            "description_declared",
            "description_item",
            "description_seizure_number",
        )
    )
    if has_structured_fields:
        return _compose_description_block_from_fields(data, keep_bottom_gap=True)

    block = _clean_pdf_multiline_text(data.get("description_block", ""))
    # Keep one extra blank line after seizure number when no legal notice follows.
    if block and re.search(r"SEIZURE NUMBER / NUM[ÉE]RO DE SAISIE:", block, re.IGNORECASE):
        if not block.endswith("\n"):
            block = block + "\n"
    return block


def _k138_layout_points(width: float, height: float, box_dimensions: tuple[float, float, float, float] | None):
    """
    Return main text anchor points in ReportLab coordinates.
    """
    base_x_to = 100
    base_y_to_top = height - 210
    x_to, y_to_top = apply_offsets(base_x_to, base_y_to_top, "notice_to")

    base_x_date = width - 200
    x_date, y_date = apply_offsets(base_x_date, base_y_to_top, "notice_date")

    base_x_on = 170
    base_y_on = height - 275
    x_on, y_on = apply_offsets(base_x_on, base_y_on, "seizure_date")

    base_x_year_left = base_x_on + 200
    x_year_left, y_year_left = apply_offsets(base_x_year_left, base_y_on, "seizure_year_L")
    x_year_right, y_year_right = apply_offsets(base_x_year_left + 20, base_y_on, "seizure_year_R")

    base_y_at = base_y_on - 30
    x_at, y_at = apply_offsets(base_x_on, base_y_at, "seizure_location")

    if box_dimensions:
        x_left_box, y_bottom_box, x_right_box, y_top_box = box_dimensions
        left_padding = 10
        x_desc = x_left_box + left_padding
        y_desc_top = y_top_box - 15
        desc_box_reportlab = (x_left_box, y_bottom_box, x_right_box, y_top_box)
    else:
        base_x_desc = 90
        base_y_desc_top = base_y_at - 50
        x_desc, y_desc_top = apply_offsets(base_x_desc, base_y_desc_top, "description")
        desc_box_reportlab = (x_desc - 8, y_desc_top - 300, 590, y_desc_top + 16)

    base_x_officer = 165
    base_y_officer = 155
    x_officer, y_officer = apply_offsets(base_x_officer, base_y_officer, "seizing_officer")

    return {
        "notice_to": (x_to, y_to_top),
        "notice_date": (x_date, y_date),
        "seizure_date": (x_on, y_on),
        "seizure_year_L": (x_year_left, y_year_left),
        "seizure_year_R": (x_year_right, y_year_right),
        "seizure_location": (x_at, y_at),
        "description": (x_desc, y_desc_top),
        "seizing_officer": (x_officer, y_officer),
        "description_box": desc_box_reportlab,
    }


def _reportlab_rect_to_fitz(
    x_left: float,
    y_bottom: float,
    x_right: float,
    y_top: float,
    page_height: float,
):
    """
    Convert a ReportLab rectangle (bottom-left origin) to a PyMuPDF rectangle (top-left origin).
    """
    fx0 = min(x_left, x_right)
    fx1 = max(x_left, x_right)
    fy0 = page_height - max(y_bottom, y_top)
    fy1 = page_height - min(y_bottom, y_top)
    return fitz.Rect(fx0, fy0, fx1, fy1)


def _rect_overlap_area(a, b) -> float:
    ix0 = max(a.x0, b.x0)
    iy0 = max(a.y0, b.y0)
    ix1 = min(a.x1, b.x1)
    iy1 = min(a.y1, b.y1)
    if ix1 <= ix0 or iy1 <= iy0:
        return 0.0
    return (ix1 - ix0) * (iy1 - iy0)


def _rect_center_distance(a, b) -> float:
    acx = (a.x0 + a.x1) / 2.0
    acy = (a.y0 + a.y1) / 2.0
    bcx = (b.x0 + b.x1) / 2.0
    bcy = (b.y0 + b.y1) / 2.0
    return ((acx - bcx) ** 2 + (acy - bcy) ** 2) ** 0.5


def _k138_expected_widget_regions(width: float, height: float, box_dimensions: tuple[float, float, float, float] | None):
    """
    Build expected widget regions (PyMuPDF coordinates) for K138 field mapping.
    """
    p = _k138_layout_points(width, height, box_dimensions)
    x_to, y_to = p["notice_to"]
    x_date, y_date = p["notice_date"]
    x_on, y_on = p["seizure_date"]
    x_year_l, y_year_l = p["seizure_year_L"]
    x_year_r, y_year_r = p["seizure_year_R"]
    x_at, y_at = p["seizure_location"]
    x_desc, y_desc = p["description"]
    x_officer, y_officer = p["seizing_officer"]
    bx0, by0, bx1, by1 = p["description_box"]

    regions_reportlab = {
        "notice_to": (x_to - 10, y_to - 46, x_to + 300, y_to + 14),
        "notice_date": (x_date - 6, y_date - 10, x_date + 180, y_date + 12),
        "seizure_date": (x_on - 6, y_on - 10, x_on + 215, y_on + 12),
        "seizure_year_L": (x_year_l - 5, y_year_l - 10, x_year_l + 24, y_year_l + 12),
        "seizure_year_R": (x_year_r - 5, y_year_r - 10, x_year_r + 24, y_year_r + 12),
        "seizure_location": (x_at - 8, y_at - 10, x_at + 380, y_at + 12),
        "description": (bx0, by0, bx1, by1),
        "seizing_officer": (x_officer - 8, y_officer - 10, x_officer + 110, y_officer + 12),
        # Inventory/declared/item/seizure number often exist as dedicated fields in some templates.
        "description_inventory": (x_desc - 4, y_desc - 2, x_desc + 280, y_desc + 14),
        "description_declared": (x_desc - 4, y_desc - 30, x_desc + 280, y_desc - 14),
        "description_item": (x_desc - 4, y_desc - 58, x_desc + 420, y_desc - 34),
        "description_seizure_number": (x_desc - 4, y_desc - 86, x_desc + 300, y_desc - 62),
    }
    return {
        key: _reportlab_rect_to_fitz(x0, y0, x1, y1, height)
        for key, (x0, y0, x1, y1) in regions_reportlab.items()
    }


def _k138_field_values_for_widgets(data: dict) -> dict[str, str]:
    """Values to push into fillable K138 fields."""
    return {
        "notice_to": _clean_pdf_multiline_text(data.get("notice_to", "")),
        "notice_date": _clean_pdf_text(_format_notice_date_for_display(data.get("notice_date", ""))),
        "seizure_date": _clean_pdf_text(data.get("seizure_date_line", "")),
        "seizure_year_L": _clean_pdf_text(data.get("seizure_year_left", "")),
        "seizure_year_R": _clean_pdf_text(data.get("seizure_year_right", "")),
        "seizure_location": _clean_pdf_text(data.get("seizure_location", "")),
        "description": _build_description_block_text(data),
        "description_inventory": _clean_pdf_text(data.get("description_inventory", "")),
        "description_declared": _clean_pdf_text(data.get("description_declared", "")),
        "description_item": _clean_pdf_text(data.get("description_item", "")),
        "description_seizure_number": _clean_pdf_text(data.get("description_seizure_number", "")),
        "seizing_officer": _clean_pdf_text(data.get("seizing_officer", "")),
    }


def _ensure_k138_widgets_on_page(
    page,
    data: dict,
    box_dimensions: tuple[float, float, float, float] | None,
    page_index: int,
) -> int:
    """
    Create editable PDF text widgets for every K138 field region that doesn't already
    have a widget. Used when the template is a plain (non-form) PDF.

    Each widget gets:
    - A white background so it cleanly covers the burned-in overlay text beneath it,
      showing only the editable widget value (no double-text).
    - A thin grey border so the user can see the field boundary.
    - Pre-filled with the extracted value.
    - Unique field name per page to avoid AcroForm collisions across the 4 copies.

    Returns number of widgets created.
    """
    values = _k138_field_values_for_widgets(data)
    regions = _k138_expected_widget_regions(page.rect.width, page.rect.height, box_dimensions)

    # Find which regions already have a widget (skip those).
    covered: set[str] = set()
    for widget in list(page.widgets() or []):
        if not _is_text_widget(widget):
            continue
        key = _k138_key_for_widget(widget, regions)
        if key:
            covered.add(key)

    # Primary fields to expose as editable in the output.
    FIELD_KEYS = [
        "notice_to",
        "notice_date",
        "seizure_date",
        "seizure_location",
        "description",
        "seizing_officer",
    ]
    PAGE_SUFFIX = f"_p{page_index}"

    created = 0
    for key in FIELD_KEYS:
        if key in covered:
            continue
        region = regions.get(key)
        if region is None:
            continue
        val = values.get(key, "")
        is_multiline = key in ("notice_to", "description")
        try:
            widget = fitz.Widget()
            widget.field_type = fitz.PDF_WIDGET_TYPE_TEXT
            widget.field_name = f"K138_{key}{PAGE_SUFFIX}"
            widget.field_value = str(val or "")
            widget.field_flags = fitz.PDF_FIELD_IS_MULTILINE if is_multiline else 0
            widget.rect = fitz.Rect(region)
            widget.fill_color = (1, 1, 1)       # White background covers overlay text
            widget.text_color = (0, 0, 0)
            widget.text_fontsize = 8
            widget.border_width = 0.5
            widget.border_color = (0.6, 0.6, 0.6)  # Subtle grey border
            page.add_widget(widget)
            created += 1
        except Exception as e:
            print(f"  ! Could not create widget for {key}: {e}")
    return created


K138_WIDGET_NAME_MAP = {
    "A": "notice_to",
    "A1": "notice_date",
    "DATE": "notice_date",
    "DATE1": "notice_date",
    "DATE2": "notice_date",
    "B": "seizure_date",
    "B1": "seizure_year_R",
    "B2": "seizure_location",
    "C": "description",
    "D": "seizing_officer",
}


def _is_text_widget(widget) -> bool:
    """Return True when widget is a text field."""
    text_type = int(getattr(fitz, "PDF_WIDGET_TYPE_TEXT", 7))
    try:
        return int(getattr(widget, "field_type", 0) or 0) == text_type
    except Exception:
        return False


def _normalize_widget_name(name: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "", _repair_mojibake_text(name or "")).upper()


def _k138_key_for_widget(
    widget,
    regions: dict[str, "fitz.Rect"],
) -> str | None:
    """
    Map a K138 widget to logical value key by field name first, then geometry fallback.
    """
    norm_name = _normalize_widget_name(str(getattr(widget, "field_name", "") or ""))
    key = K138_WIDGET_NAME_MAP.get(norm_name)
    if key:
        return key
    # Some templates use custom date field names (e.g. DATE_GEN, A1_0, etc.).
    if "DATE" in norm_name or norm_name.startswith("A1"):
        return "notice_date"

    wrect = widget.rect
    best_key = None
    best_overlap = 0.0
    for candidate, region in regions.items():
        overlap = _rect_overlap_area(wrect, region)
        if overlap > best_overlap:
            best_overlap = overlap
            best_key = candidate
    if best_key and best_overlap > 0:
        return best_key

    best_dist = 1e9
    for candidate, region in regions.items():
        dist = _rect_center_distance(wrect, region)
        if dist < best_dist:
            best_dist = dist
            best_key = candidate
    if best_dist <= 140:
        return best_key
    return None


def _k138_widget_score_for_key(
    widget,
    key: str,
    regions: dict[str, "fitz.Rect"],
) -> float:
    """
    Rank candidate widgets for the same logical key.
    Higher score = better match.
    """
    region = regions.get(key)
    if region is None:
        return -1e9
    norm_name = _normalize_widget_name(str(getattr(widget, "field_name", "") or ""))
    explicit_map = K138_WIDGET_NAME_MAP.get(norm_name)
    name_bonus = 0.0
    if explicit_map == key:
        name_bonus = 1_000_000.0
    elif key == "notice_date" and ("DATE" in norm_name or norm_name.startswith("A1")):
        name_bonus = 900_000.0
    overlap = _rect_overlap_area(widget.rect, region)
    dist = _rect_center_distance(widget.rect, region)
    return name_bonus + (10.0 * overlap) - dist


def _k138_widget_keys_on_page(
    page,
    box_dimensions: tuple[float, float, float, float] | None,
) -> set[str]:
    """Return logical keys represented by widgets on this page."""
    widgets = list(page.widgets() or [])
    if not widgets:
        return set()
    regions = _k138_expected_widget_regions(page.rect.width, page.rect.height, box_dimensions)
    keys: set[str] = set()
    for widget in widgets:
        key = _k138_key_for_widget(widget, regions)
        if key:
            keys.add(key)
            continue
        # Extra guard: some templates use custom date field names that are not in map/geometry.
        norm_name = _normalize_widget_name(str(getattr(widget, "field_name", "") or ""))
        if "DATE" in norm_name or norm_name.startswith("A1"):
            keys.add("notice_date")
    return keys


def fill_k138_widget_fields(
    pdf_path: str,
    data: dict,
    pages_to_fill: int,
    box_dimensions: tuple[float, float, float, float] | None,
) -> int:
    """
    Populate existing PDF form widgets in K138 output so the PDF stays fillable.
    Returns number of widget fields updated.
    Always saves via temp-file + atomic replace so the file handle is fully closed
    before any filesystem rename (avoids Windows file-lock issues).
    """
    if not HAVE_PYMUPDF:
        return 0
    import os
    tmp_path = f"{pdf_path}._saving_tmp"
    if os.path.exists(tmp_path):
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
    updated = 0
    doc = fitz.open(pdf_path)
    try:
        max_pages = min(pages_to_fill, len(doc))
        for i in range(max_pages):
            updated += _fill_k138_widgets_on_page(doc[i], data, box_dimensions)
        if updated > 0:
            doc.save(tmp_path, incremental=False, encryption=fitz.PDF_ENCRYPT_KEEP)
    finally:
        doc.close()
    if updated > 0:
        os.replace(tmp_path, pdf_path)
    if os.path.exists(tmp_path):
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
    return updated


def _insert_text_reportlab_coords(page, x: float, y_reportlab: float, text: str, font_size: float = 9) -> None:
    """Insert text on a PyMuPDF page using ReportLab-style (bottom-left) y-coordinates."""
    txt = _clean_pdf_text(text)
    if not txt:
        return
    y_fitz = page.rect.height - y_reportlab
    page.insert_text((x, y_fitz), txt, fontsize=font_size, fontname="helv", color=(0, 0, 0))


def _draw_k138_overlay_on_page_fitz(
    page,
    data: dict,
    form_type: str,
    box_dimensions: tuple[float, float, float, float] | None,
    skip_keys: set[str] | None = None,
):
    """
    Draw the same K138 overlay content directly with PyMuPDF (keeps PDF form structure intact).
    """
    width = float(page.rect.width)
    height = float(page.rect.height)
    points = _k138_layout_points(width, height, box_dimensions)
    skip = set(skip_keys or set())

    # 1) Notice to
    x_to, y_to_top = points["notice_to"]
    if "notice_to" not in skip:
        notice_to_lines = _repair_mojibake_text(data.get("notice_to", "")).split("\n")
        for idx, line in enumerate(notice_to_lines):
            _insert_text_reportlab_coords(page, x_to, y_to_top - 11 * idx, line, 9)

    # 2) Notice date
    x_date, y_date = points["notice_date"]
    if "notice_date" not in skip:
        _insert_text_reportlab_coords(page, x_date, y_date, _format_notice_date_for_display(data.get("notice_date", "")), 9)

    # 3) Seizure date/year
    x_on, y_on = points["seizure_date"]
    if "seizure_date" not in skip:
        _insert_text_reportlab_coords(page, x_on, y_on, _repair_mojibake_text(data.get("seizure_date_line", "")), 9)
    x_year_right, y_year_right = points["seizure_year_R"]
    if "seizure_year_R" not in skip:
        _insert_text_reportlab_coords(page, x_year_right, y_year_right, data.get("seizure_year_right", ""), 9)

    # 4) Location
    x_at, y_at = points["seizure_location"]
    if "seizure_location" not in skip:
        _insert_text_reportlab_coords(page, x_at, y_at, data.get("seizure_location", ""), 9)

    # 5) Description block with wrapping
    if not (
        ("description" in skip)
        or ("description_inventory" in skip)
        or ("description_declared" in skip)
        or ("description_item" in skip)
        or ("description_seizure_number" in skip)
    ):
        description_block = _build_description_block_text(data)
        x_desc, y_desc_top = points["description"]
        if box_dimensions:
            x_left_box, _y_bottom_box, x_right_box, _y_top_box = box_dimensions
            max_width_points = (x_right_box - x_left_box) - 10 - 8
        else:
            max_width_points = 590 - x_desc

        is_cannabis_notice = form_type == "Cannabis-Stupefiant"
        notice_start_markers = ["(Please note:", "(Attention :"]
        in_notice_section = False
        line_idx = 0
        for original_line in _repair_mojibake_text(description_block).split("\n"):
            if not original_line.strip():
                line_idx += 1
                continue
            line_upper = original_line.strip().upper()
            is_declared_line = line_upper.startswith("DECLARED")
            if is_cannabis_notice and any(marker in original_line for marker in notice_start_markers):
                in_notice_section = True
            if in_notice_section and is_cannabis_notice:
                font_size_for_line = 7
                line_spacing = 10
            else:
                font_size_for_line = 9
                line_spacing = 12

            if is_declared_line:
                wrapped_lines = [original_line]
            else:
                wrapped_lines = wrap_text_measured(original_line, max_width_points, "Helvetica", int(font_size_for_line))
            for wrapped_line in wrapped_lines:
                if not is_declared_line:
                    line_width = stringWidth(wrapped_line, "Helvetica", font_size_for_line)
                    while line_width > max_width_points and wrapped_line:
                        wrapped_line = wrapped_line[:-1]
                        line_width = stringWidth(wrapped_line, "Helvetica", font_size_for_line)
                _insert_text_reportlab_coords(page, x_desc, y_desc_top - line_spacing * line_idx, wrapped_line, font_size_for_line)
                line_idx += 1

    # 6) Seizing officer
    x_officer, y_officer = points["seizing_officer"]
    if "seizing_officer" not in skip:
        _insert_text_reportlab_coords(page, x_officer, y_officer, data.get("seizing_officer", ""), 9)


def _fill_k138_widgets_on_page(
    page,
    data: dict,
    box_dimensions: tuple[float, float, float, float] | None,
) -> int:
    """Populate K138 widgets on a single page and return number of updates.
    All text widgets are made editable (read-only flag cleared) regardless of whether
    they receive a value, so the output PDF stays fully fillable by the user.
    """
    values = _k138_field_values_for_widgets(data)
    widgets = list(page.widgets() or [])
    if not widgets:
        return 0
    regions = _k138_expected_widget_regions(page.rect.width, page.rect.height, box_dimensions)
    updated = 0
    keyed_widgets: dict[str, list] = {}
    for widget in widgets:
        if not _is_text_widget(widget):
            continue
        try:
            # Always clear read-only flag so user can edit after generation.
            flags = int(getattr(widget, "field_flags", 0) or 0)
            if flags & 1:
                widget.field_flags = flags & ~1
                widget.update()
        except Exception:
            pass
        best_key = _k138_key_for_widget(widget, regions)
        if not best_key:
            continue
        keyed_widgets.setdefault(best_key, []).append(widget)

    for key, key_widgets in keyed_widgets.items():
        ranked = sorted(
            key_widgets,
            key=lambda w: _k138_widget_score_for_key(w, key, regions),
            reverse=True,
        )
        primary = ranked[0]
        val = values.get(key, "")
        try:
            primary.field_value = str(val or "")
            primary.update()
            updated += 1
        except Exception:
            pass

        # If multiple date widgets overlap/match, only keep one populated to avoid
        # double-rendered notice date on some K138 templates.
        if key == "notice_date" and len(ranked) > 1:
            for extra in ranked[1:]:
                try:
                    if str(getattr(extra, "field_value", "") or "").strip():
                        extra.field_value = ""
                        extra.update()
                        updated += 1
                except Exception:
                    continue
    return updated


def _save_fitz_doc(output_doc, output_path: str) -> None:
    """Save a PyMuPDF document to output_path.
    Always saves to a temp file first, then closes the doc, then atomically replaces.
    This avoids Windows file-lock issues from having the doc open during rename.
    """
    import os
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    tmp_path = f"{output_path}._saving_tmp"
    if os.path.exists(tmp_path):
        try:
            os.unlink(tmp_path)
        except Exception:
            pass
    try:
        output_doc.save(tmp_path, incremental=False, encryption=fitz.PDF_ENCRYPT_KEEP)
    finally:
        output_doc.close()
    os.replace(tmp_path, output_path)
    if os.path.exists(tmp_path):
        try:
            os.unlink(tmp_path)
        except Exception:
            pass


# ---------------------------------------------------------------------
# Overlay creation
# ---------------------------------------------------------------------

def create_overlay_page(width: float, height: float, data: dict, form_type: str = None, box_dimensions: tuple[float, float, float, float] | None = None):
    """
    Create a single-page PDF in memory with the text overlay for the K138.

    Coordinates tuned for a portrait letter-size K138 form (â‰ˆ612 x 792).
    Adjust either:
      - GLOBAL_DX / GLOBAL_DY to move everything,
      - OFFSETS[...] entries for each field.
    
    Args:
        width: Page width
        height: Page height
        data: Dictionary containing form data
        form_type: Form type ("Cannabis-Stupefiant", "Knives-Arms", or "Stupefiant-Others")
        box_dimensions: (x_left, y_bottom, x_right, y_top) of description box, or None to use defaults
    """
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(width, height))
    c.setFont("Helvetica", 9)

    # NOTE: origin (0,0) is bottom-left. height ~ 792 for letter.
    
    # Prepare description block (notice text is already included from .txt files via description_item)
    description_block = _build_description_block_text(data)
    # No need to add notices here - they're already included in description_item from the .txt files

    # --------------------------------------------------
    # 1) "To:" block (name + address)
    # --------------------------------------------------
    base_x_to = 100
    base_y_to_top = height - 210    # ~582 on a 792-pt page

    x_to, y_to_top = apply_offsets(base_x_to, base_y_to_top, "notice_to")
    notice_to_lines = _repair_mojibake_text(data.get("notice_to", "")).split("\n")
    for idx, line in enumerate(notice_to_lines):
        c.drawString(x_to, y_to_top - 11 * idx, _clean_pdf_text(line))  # Reduced from 13 to 11

    # --------------------------------------------------
    # 2) "Date" (right header area) â€“ Avis de saisie = date letter was generated
    # --------------------------------------------------
    base_x_date = width - 200
    base_y_date = base_y_to_top
    x_date, y_date = apply_offsets(base_x_date, base_y_date, "notice_date")
    # Format notice_date as DD MONTH YYYY for display (e.g. 05 February 2026)
    _nd_display = _format_notice_date_for_display(data.get("notice_date", ""))
    c.drawString(x_date, y_date, _clean_pdf_text(_nd_display))

    # --------------------------------------------------
    # 3) Seizure date (On / le)
    # --------------------------------------------------
    base_x_on = 170
    base_y_on = height - 275  # Moved up from 295 (20 points closer to top)

    x_on, y_on = apply_offsets(base_x_on, base_y_on, "seizure_date")
    c.drawString(x_on, y_on, _clean_pdf_text(data.get("seizure_date_line", "")))

    # Year split "20" "25"
    base_x_year_left = base_x_on + 200
    base_y_year = base_y_on

    x_year_left, y_year_left = apply_offsets(base_x_year_left, base_y_year, "seizure_year_L")
    x_year_right, y_year_right = apply_offsets(base_x_year_left + 20, base_y_year, "seizure_year_R")

    # c.drawString(x_year_left,  y_year_left,  data["seizure_year_left"]) Dmitry
    c.drawString(x_year_right, y_year_right, _clean_pdf_text(data.get("seizure_year_right", "")))

    # --------------------------------------------------
    # 4) Location ("at" / Ã ) â€“ MONTREAL POSTAL FACILITY...
    # --------------------------------------------------
    base_x_at = base_x_on
    base_y_at = base_y_on - 30  # Moved down (was 15) so Montreal text sits lower

    x_at, y_at = apply_offsets(base_x_at, base_y_at, "seizure_location")
    c.drawString(x_at, y_at, _clean_pdf_text(data.get("seizure_location", "")))

    # --------------------------------------------------
    # 5) Description block (large text area mid-page)
    # --------------------------------------------------
    # Use detected box dimensions if available, otherwise fall back to defaults
    if box_dimensions:
        x_left_box, y_bottom_box, x_right_box, y_top_box = box_dimensions
        # Use the detected box boundaries directly
        left_padding = 10
        right_padding = 8
        x_desc = x_left_box + left_padding
        # Add top padding to avoid overlapping with existing text in top-left corner
        top_padding = 15  # Points to move first line down
        y_desc_top = y_top_box - top_padding
        max_width_points = (x_right_box - x_left_box) - left_padding - right_padding
    else:
        # Fallback to manual calculation
        base_x_desc = 90
        base_y_desc_top = base_y_at - 50  # Reduced from 70 to 50 (closer to location)
        
        x_desc, y_desc_top = apply_offsets(base_x_desc, base_y_desc_top, "description")
        
        # Determine right edge of description box
        # Page width is 612, leave small margin on right (around 20-25 points)
        right_edge_x = 590  # Using more of the available box width
        max_width_points = right_edge_x - x_desc
    
    # Detect if we're in the cannabis notice section (for smaller font)
    is_cannabis_notice = form_type == "Cannabis-Stupefiant"
    notice_start_markers = ["(Please note:", "(Attention :"]
    in_notice_section = False
    
    # Process description block with text wrapping
    line_idx = 0
    for original_line in _repair_mojibake_text(description_block).split("\n"):
        if not original_line.strip():
            # Empty line - skip to next line
            line_idx += 1
            continue
        line_upper = original_line.strip().upper()
        is_declared_line = line_upper.startswith("DECLARED")
        
        # Detect start of cannabis notice
        if is_cannabis_notice and any(marker in original_line for marker in notice_start_markers):
            in_notice_section = True
        
        # Use smaller font for cannabis notice (7pt instead of 9pt)
        if in_notice_section and is_cannabis_notice:
            font_size_for_line = 7
            line_spacing = 10# Tighter spacing for smaller font - moved up
        else:
            font_size_for_line = 9
            line_spacing = 12
        
        # Set font size for this line
        c.setFont("Helvetica", font_size_for_line)
        
        # Keep DECLARED line on one line (no wrapping).
        if is_declared_line:
            wrapped_lines = [original_line]
        else:
            # Wrap long lines using actual text width measurement
            wrapped_lines = wrap_text_measured(original_line, max_width_points, "Helvetica", font_size_for_line)
        for wrapped_line in wrapped_lines:
            if not is_declared_line:
                # Double-check the line fits before drawing
                line_width = stringWidth(wrapped_line, "Helvetica", font_size_for_line)
                if line_width > max_width_points:
                    # If somehow still too wide, truncate
                    # This shouldn't happen but safety check
                    while line_width > max_width_points and wrapped_line:
                        wrapped_line = wrapped_line[:-1]
                        line_width = stringWidth(wrapped_line, "Helvetica", font_size_for_line)
            
            c.drawString(x_desc, y_desc_top - line_spacing * line_idx, _clean_pdf_text(wrapped_line))
            line_idx += 1
    
    # Reset font to default for remaining fields
    c.setFont("Helvetica", 9)

    # --------------------------------------------------
    # 6) Seizing Officer (bottom area)
    # --------------------------------------------------
    base_x_officer = 165
    base_y_officer = 155

    x_officer, y_officer = apply_offsets(base_x_officer, base_y_officer, "seizing_officer")
    c.drawString(x_officer, y_officer, _clean_pdf_text(data.get("seizing_officer", "")))

    # Finish overlay
    c.save()
    packet.seek(0)
    overlay_pdf = PdfReader(packet)
    return overlay_pdf.pages[0]


# ---------------------------------------------------------------------
# Main routine
# ---------------------------------------------------------------------

def fill_k138(
    template_path: str = TEMPLATE_PDF,
    output_path: str = OUTPUT_PDF,
    use_csv: bool = False,
    csv_path: str | None = None,
    form_type: str | None = None,
):
    """
    Load the K138 template, overlay the dummy (or CSV) data on each page,
    and write out a new filled PDF.

    Only the first 4 K138 pages are filled (official copies).
    Any page after page 4 is left unchanged.
    
    Args:
        template_path: Path to the K138 template PDF
        output_path: Path where the filled PDF will be saved
        use_csv: Whether to load data from CSV
        csv_path: Path to CSV file (if use_csv is True)
        form_type: Form type ("Cannabis-Stupefiant", "Knives-Arms", or "Stupefiant-Others").
                   If None, auto-detects from template_path filename.
    """

    # Detect form type if not provided
    if form_type is None:
        form_type = detect_form_type(template_path)
    else:
        # Normalize form type
        form_type_lower = form_type.lower()
        if "cannabis" in form_type_lower:
            form_type = "Cannabis-Stupefiant"
        elif "knife" in form_type_lower or "knives" in form_type_lower or "arms" in form_type_lower:
            form_type = "Knives-Arms"
        else:
            form_type = "Stupefiant-Others"

    # Decide data source
    if use_csv and csv_path:
        data_from_csv = load_data_from_csv(csv_path)
        # Ensure any missing keys fall back to dummy_data
        merged = dummy_data.copy()
        merged.update({k: v for k, v in data_from_csv.items() if v is not None})
        # Check if form_type is in CSV, override if present
        if "form_type" in data_from_csv and data_from_csv["form_type"]:
            form_type = data_from_csv["form_type"]
        data = merged
    else:
        data = dummy_data

    # Detect box dimensions from template
    box_dimensions = detect_description_box_dimensions(template_path)
    if box_dimensions:
        x_left, y_bottom, x_right, y_top = box_dimensions
        print(f"Detected description box: left={x_left:.1f}, bottom={y_bottom:.1f}, right={x_right:.1f}, top={y_top:.1f}")
        print(f"Box width: {x_right - x_left:.1f} points, height: {y_top - y_bottom:.1f} points")
    else:
        print("Warning: Could not detect box dimensions, using defaults")

    if HAVE_PYMUPDF:
        doc = fitz.open(template_path)
        try:
            num_pages = len(doc)
            pages_to_fill = min(MAX_FILLED_PAGES, num_pages)
            print(f"Filling {pages_to_fill} of {num_pages} page(s) in K138...")

            # Detect once whether this template has any existing form widgets.
            has_existing_widgets = any(
                list(doc[i].widgets() or [])
                for i in range(min(pages_to_fill, len(doc)))
            )

            widget_updates = 0
            widgets_created = 0
            for idx, page in enumerate(doc, start=1):
                if idx > pages_to_fill:
                    print(f"  - left page {idx}/{num_pages} unchanged")
                    continue

                if has_existing_widgets:
                    # Template already has form fields: fill them and draw overlay
                    # only for fields that have no matching widget (skip_keys handles this).
                    # Also create missing widgets to avoid double-rendered text (editable + static).
                    widget_updates += _fill_k138_widgets_on_page(page, data, box_dimensions)
                    widgets_created += _ensure_k138_widgets_on_page(page, data, box_dimensions, idx)
                    widget_updates += _fill_k138_widgets_on_page(page, data, box_dimensions)
                    present_widget_keys = _k138_widget_keys_on_page(page, box_dimensions)
                    _draw_k138_overlay_on_page_fitz(
                        page, data, form_type, box_dimensions,
                        skip_keys=present_widget_keys,
                    )
                else:
                    # Template has NO form fields (plain PDF).
                    # Create editable widgets first.
                    # The widget shows the pre-filled value; user can click and edit.
                    # Using a slight white-transparent fill so the widget is visually clean
                    # and doesn't obscure the printed text already on the template.
                    widgets_created += _ensure_k138_widgets_on_page(page, data, box_dimensions, idx)
                    # Draw static overlay only for fields not backed by widgets.
                    present_widget_keys = _k138_widget_keys_on_page(page, box_dimensions)
                    _draw_k138_overlay_on_page_fitz(
                        page, data, form_type, box_dimensions,
                        skip_keys=present_widget_keys,
                    )

                print(f"  - populated page {idx}/{num_pages}")

            _save_fitz_doc(doc, output_path)
            if has_existing_widgets:
                print(f"Updated {widget_updates} existing widget field(s).")
                if widgets_created > 0:
                    print(f"Created {widgets_created} missing widget field(s).")
            else:
                print(f"Created {widgets_created} editable field(s) over static overlay.")
            return
        except Exception:
            try:
                doc.close()
            except Exception:
                pass
            raise

    # Fallback path when PyMuPDF is unavailable (overlay only).
    reader = PdfReader(template_path)
    writer = PdfWriter()
    num_pages = len(reader.pages)
    pages_to_fill = min(MAX_FILLED_PAGES, num_pages)
    print(f"Filling {pages_to_fill} of {num_pages} page(s) in K138...")

    for idx, src_page in enumerate(reader.pages, start=1):
        width = float(src_page.mediabox.width)
        height = float(src_page.mediabox.height)

        out_page = PageObject.create_blank_page(width=width, height=height)
        out_page.merge_page(src_page)

        if idx <= pages_to_fill:
            overlay_page = create_overlay_page(width, height, data, form_type, box_dimensions)
            out_page.merge_page(overlay_page)
            print(f"  - populated page {idx}/{num_pages}")
        else:
            print(f"  - left page {idx}/{num_pages} unchanged")

        writer.add_page(out_page)

    with open(output_path, "wb") as f_out:
        writer.write(f_out)
    print("Warning: PyMuPDF unavailable; generated K138 may not remain fillable.")


if __name__ == "__main__":
    import os
    import sys
    
    # Check if called with command-line arguments (from extractor)
    if len(sys.argv) >= 3:
        # Called as: python fill_k138_notice.py <k138_pdf> <k138_values_csv>
        k138_pdf_path = sys.argv[1]
        k138_csv_path = sys.argv[2]
        # Use the same filename as input (overwrites or creates new file)
        output_path = k138_pdf_path
        
        print(f"Filling K138 from CSV...")
        print(f"Template: {k138_pdf_path}")
        print(f"CSV: {k138_csv_path}")
        print(f"Output: {output_path}")
        
        fill_k138(
            template_path=k138_pdf_path,
            output_path=output_path,
            use_csv=True,
            csv_path=k138_csv_path,
        )
        print(f"Created: {output_path}")
    else:
        # Standalone usage - check if CSV file exists
        csv_exists = os.path.exists("k138_values.csv")
        
        if csv_exists:
            # Example for CSV usage:
            print("Using CSV file: k138_values.csv")
            fill_k138(
                template_path=TEMPLATE_PDF,
                output_path="K138_filled_from_csv.pdf",
                use_csv=True,
                csv_path="k138_values.csv",
            )
            print("Created: K138_filled_from_csv.pdf")
        else:
            # Basic usage with dummy data:
            print("No CSV file found. Using dummy data with default template.")
            print(f"Template: {TEMPLATE_PDF}")
            print(f"Output: {OUTPUT_PDF}")
            fill_k138()
            print(f"Created: {OUTPUT_PDF}")

#endregion
