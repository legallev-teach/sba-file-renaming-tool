"""
SBA File Renamer — Core Engine
Shared by both the Tkinter desktop app and the Flask web app.
"""

import csv
import os
import re
import shutil
import logging
from pathlib import Path
from dataclasses import dataclass
from typing import Optional

logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# CSEC subject list (from CXC MoE Memorandum, May-June 2026 appendix)
# Architecture note: CAPE codes can be added in a second dict later.
# ---------------------------------------------------------------------------

CSEC_SUBJECTS: dict[str, str] = {
    "Additional Mathematics":                                        "01254090",
    "Caribbean History":                                             "01210090",
    "Economics":                                                     "01216090",
    "Electronic Document Preparation and Management":                "01251090",
    "English A":                                                     "01218090",
    "English B":                                                     "01219090",
    "Geography":                                                     "01225090",
    "Information Technology":                                        "01229090",
    "Mathematics":                                                   "01234090",
    "Office Administration":                                         "01237090",
    "Physical Education and Sport":                                  "01252090",
    "Principles of Accounts":                                        "01239090",
    "Principles of Business":                                        "01240090",
    "Religious Education":                                           "01241090",
    "Social Studies":                                                "01243090",
    "Theatre Arts":                                                  "01248090",
    "Human and Social Biology":                                      "01253090",
}

# Sorted list for display in dropdowns
CSEC_SUBJECT_NAMES: list[str] = sorted(CSEC_SUBJECTS.keys())

FILE_TYPES = ["SBA", "Cover Sheet", "Mark Scheme"]
DOC_NUMBERS = [1, 2, 3]

# Candidate number length constant — used everywhere instead of bare 10
CAND_NUM_LENGTH = 10
MOD_CODE_LENGTH = 8


# ---------------------------------------------------------------------------
# Naming rules (Section 5, CXC memorandum)
# ---------------------------------------------------------------------------

def build_filename(
    candidate_number: str,
    moderation_code: str,
    file_type: str,
    doc_number: int = 1,
    extension: str = ".pdf",
) -> str:
    """
    Return the required CXC filename (with extension).

    SBA sample:    {cand}{mod}-{n}          e.g. 1000750100012250901-1.pdf
    Cover Sheet:   {cand}{mod}CS            e.g. 100075010001225090CS.pdf
    Mark Scheme:   {cand}{mod}-{n}MS        e.g. 100075010001225090-1MS.pdf

    NOTE: The memorandum shows no separator between cand+mod and CS/MS suffixes.
    """
    base = f"{candidate_number}{moderation_code}"

    if file_type == "SBA":
        stem = f"{base}-{doc_number}"
    elif file_type == "Cover Sheet":
        stem = f"{base}CS"
    elif file_type == "Mark Scheme":
        stem = f"{base}-{doc_number}MS"
    else:
        raise ValueError(f"Unknown file type: {file_type!r}")

    return stem + extension


# ---------------------------------------------------------------------------
# Validation helpers
# ---------------------------------------------------------------------------

def validate_candidate_number(value: str) -> tuple[bool, str]:
    """Return (ok, error_message). ok=True means valid."""
    cleaned = value.strip().replace(" ", "")
    if not cleaned.isdigit():
        return False, "Candidate number must contain digits only."
    if len(cleaned) != CAND_NUM_LENGTH:
        return False, f"Candidate number must be exactly {CAND_NUM_LENGTH} digits (got {len(cleaned)})."
    return True, ""


def validate_moderation_code(value: str) -> tuple[bool, str]:
    cleaned = value.strip()
    if not cleaned.isdigit():
        return False, "Moderation code must be numeric."
    if len(cleaned) != MOD_CODE_LENGTH:
        return False, f"Moderation code must be {MOD_CODE_LENGTH} digits (got {len(cleaned)})."
    return True, ""


# ---------------------------------------------------------------------------
# Tesseract availability check
# ---------------------------------------------------------------------------

def check_tesseract_available() -> bool:
    """
    Return True if the Tesseract binary is reachable.
    Call once at startup; if False, disable OCR in the UI.
    """
    try:
        import pytesseract
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Text extraction from PDF (clean text first, OCR fallback)
# ---------------------------------------------------------------------------

def extract_title_page_text(pdf_path: str | Path, pages: int = 2) -> tuple[str, str | None]:
    """
    Try pypdf first (fast, works on text-based PDFs).
    Falls back to pytesseract OCR if no usable text found.

    Returns:
        (text, error_message)
        text          — extracted text, may be empty string
        error_message — human-readable failure reason, or None on success
    """
    pdf_path = Path(pdf_path)

    if not pdf_path.exists():
        return "", f"File not found: {pdf_path.name}"

    text, err = _extract_with_pypdf(pdf_path, pages)
    if err:
        logger.warning("pypdf failed on %s: %s", pdf_path.name, err)

    if len(text.strip()) > 30:          # enough real content
        return text, None

    # OCR fallback
    ocr_text, ocr_err = _extract_with_ocr(pdf_path, pages)
    if ocr_err:
        # Both methods failed — surface the reason
        reason = f"Text extraction: {err or 'no text found'}. OCR: {ocr_err}"
        return text, reason   # return whatever pypdf gave (possibly empty)

    return ocr_text, None


def _extract_with_pypdf(pdf_path: Path, pages: int) -> tuple[str, str | None]:
    """Returns (text, error_message_or_None)."""
    try:
        from pypdf import PdfReader
        reader = PdfReader(str(pdf_path))
        chunks = []
        for i, page in enumerate(reader.pages):
            if i >= pages:
                break
            chunks.append(page.extract_text() or "")
        return "\n".join(chunks), None
    except Exception as exc:
        return "", str(exc)


def _extract_with_ocr(pdf_path: Path, pages: int) -> tuple[str, str | None]:
    """
    Rasterise first N pages and run Tesseract.
    Returns (text, error_message_or_None).
    Cleans up temp files via context manager.
    """
    try:
        import tempfile
        from pdf2image import convert_from_path
        import pytesseract

        with tempfile.TemporaryDirectory() as tmpdir:
            images = convert_from_path(
                str(pdf_path),
                first_page=1,
                last_page=pages,
                dpi=200,
                output_folder=tmpdir,
            )
            text = "\n".join(pytesseract.image_to_string(img) for img in images)
        return text, None

    except FileNotFoundError:
        return "", "Tesseract binary not found. Install Tesseract or place it on PATH."
    except Exception as exc:
        return "", str(exc)


# ---------------------------------------------------------------------------
# Candidate number extraction from raw text
# ---------------------------------------------------------------------------

CAND_PATTERNS = [
    # "Candidate Number: 1234567890" or "Candidate No. 1234567890"
    re.compile(r"[Cc]andidate\s+(?:[Nn](?:umber|o\.?|um\.?))[:\s#]*(\d{10})"),
    # Bare 10-digit number (any leading digit — not just 1)
    re.compile(r"\b(\d{10})\b"),
]

CAND_NAME_PATTERNS = [
    re.compile(r"[Cc]andidate\s+[Nn]ame[:\s]+([A-Za-z ,.\-']+)"),
    re.compile(r"[Nn]ame[:\s]+([A-Za-z ,.\-']+)"),
]


def extract_candidate_info(text: str) -> dict[str, str]:
    """
    Try to pull candidate number (and optionally name) from title-page text.
    Returns {"number": "...", "name": "..."}.  Values may be empty strings.
    """
    number = ""
    name = ""

    for pattern in CAND_PATTERNS:
        m = pattern.search(text)
        if m:
            number = m.group(1)
            break

    for pattern in CAND_NAME_PATTERNS:
        m = pattern.search(text)
        if m:
            name = m.group(1).strip().strip(",")
            break

    return {"number": number, "name": name}


# ---------------------------------------------------------------------------
# Name normalisation
# CXC master lists use "SURNAME, FIRSTNAME MIDDLENAME" format.
# Cover pages typically show "Firstname Surname" only.
# We store both forms so matching works either way.
# ---------------------------------------------------------------------------

def normalise_cxc_name(raw: str) -> tuple[str, str]:
    """
    Given a CXC-format name like "ABERDEEN, CHRIS OBED", return:
        ("ABERDEEN, CHRIS OBED",  ← original, title-cased
         "Chris Aberdeen")         ← reversed firstname-surname, no middle name

    If there is no comma (name is already in normal order), the second
    value is just the title-cased original.
    """
    raw = raw.strip()
    if "," in raw:
        surname_part, given_part = raw.split(",", 1)
        surname   = surname_part.strip().title()
        # Take only the first given name — drop middle names
        firstname = given_part.strip().split()[0].title() if given_part.strip() else ""
        original  = raw.title()
        reversed_ = f"{firstname} {surname}".strip()
    else:
        original  = raw.title()
        reversed_ = original
    return original, reversed_


def name_matches(extracted: str, original: str, reversed_: str) -> bool:
    """
    Case-insensitive check: does the name extracted from a cover page
    match either form of the master list name?

    Students rarely write middle names on their SBA cover pages, so we
    compare on first name + surname only, ignoring middle names on both sides.

    Matching strategy (tried in order, stops at first match):
    1. Exact match against full original or reversed form (catches the rare
       student who did write their full name).
    2. First-name + surname match: extract just those two tokens from both
       sides and compare — "Christopher Ishmael" matches
       "Ishmael, Christopher Obediah Marcus" because both share the same
       first name (Christopher) and surname (Ishmael).
    """
    ext    = extracted.strip().lower()
    orig_l = original.strip().lower()
    rev_l  = reversed_.strip().lower()

    # Pass 1: exact match
    if ext == orig_l or ext == rev_l:
        return True

    # Pass 2: first-name + surname only
    # reversed_ is always "Firstname Surname" from normalise_cxc_name
    rev_parts = rev_l.split()
    if len(rev_parts) < 2:
        return False
    master_first, master_sur = rev_parts[0], rev_parts[-1]

    # Cover page may be "Firstname Surname", "Firstname Middle Surname",
    # or even "SURNAME, FIRSTNAME" — try both orientations.
    ext_parts = ext.replace(",", "").split()
    if len(ext_parts) < 2:
        return False

    # Firstname ... Surname (most common)
    if ext_parts[0] == master_first and ext_parts[-1] == master_sur:
        return True

    # Surname ... Firstname (less common)
    if ext_parts[0] == master_sur and ext_parts[-1] == master_first:
        return True

    return False


# ---------------------------------------------------------------------------
# Master list verification
# Call this after the user enters (or the app scans) a candidate number,
# if a master list has been loaded.  Returns a VerifyResult — always check
# .warnings before proceeding to rename.
# ---------------------------------------------------------------------------

@dataclass
class VerifyResult:
    """
    Result of checking a candidate number + optional name against the master list.

    Attributes:
        found         — True if the candidate number exists in the master list.
        master_name   — The name stored in the master list for this number
                        (empty string if not found or master list had no names).
        name_match    — True if the cover-page name matches the master list name.
                        None if no cover-page name was available for comparison,
                        or if the master list has no names.
        warnings      — Human-readable warning strings; empty list means all clear.
    """
    found: bool
    master_name: str = ""
    name_match: Optional[bool] = None
    warnings: list = None

    def __post_init__(self):
        if self.warnings is None:
            self.warnings = []


def verify_candidate_against_master_list(
    candidate_number: str,
    master_list: dict[str, str],
    cover_page_name: str = "",
) -> VerifyResult:
    """
    Cross-check a candidate number (and optionally a cover-page name) against
    a loaded master list.

    Args:
        candidate_number  — 10-digit string entered by user or scanned from PDF.
        master_list       — dict returned by load_master_list()
                            {candidate_number_str: name_str}.
        cover_page_name   — Name extracted from the SBA cover page (may be "").

    Returns:
        VerifyResult with .warnings populated for any issues found.
        An empty .warnings list means the candidate checked out cleanly.

    The app should display warnings to the teacher but NOT block the rename —
    the teacher is the authority; the master list is a safety net.
    """
    warnings: list[str] = []
    num = candidate_number.strip().replace(" ", "")

    # ── Check 1: does the number exist at all? ─────────────────────────────
    if num not in master_list:
        warnings.append(
            f"⚠️  Candidate number {num} was NOT found in the master list. "
            "Check for a typo — digits transposed or missing."
        )
        return VerifyResult(found=False, warnings=warnings)

    master_name = master_list[num]

    # ── Check 2: name comparison (only if both sides are available) ─────────
    name_match: Optional[bool] = None

    if cover_page_name.strip() and master_name.strip():
        original, reversed_ = normalise_cxc_name(master_name)
        name_match = name_matches(cover_page_name, original, reversed_)

        if not name_match:
            warnings.append(
                f"\u26a0\ufe0f  Name mismatch: cover page shows '{cover_page_name.strip()}' "
                f"but master list has '{master_name}'. "
                "Confirm this is the right candidate before renaming."
            )

    return VerifyResult(
        found=True,
        master_name=master_name,
        name_match=name_match,
        warnings=warnings,
    )


# ---------------------------------------------------------------------------
# Master list cross-check (Excel / CSV / PDF)
# Replaced pandas with csv + openpyxl to reduce build size by ~100 MB.
# ---------------------------------------------------------------------------

def load_master_list(path: str | Path) -> dict[str, str]:
    """
    Load an Excel or CSV candidate master list.
    Expects at minimum a column containing 10-digit candidate numbers.
    Returns {candidate_number_str: name_str}.

    Raises ValueError with a human-readable message if the file cannot be
    parsed or the required column is not found.
    """
    path = Path(path)
    suffix = path.suffix.lower()

    if suffix in (".xlsx", ".xlsm", ".xls"):
        rows, skipped = _load_excel(path)
    elif suffix == ".csv":
        rows, skipped = _load_csv(path)
    elif suffix == ".pdf":
        rows, skipped = _load_pdf_master_list(path)
    else:
        raise ValueError(f"Unsupported file type: {suffix}. Use .xlsx, .csv, or .pdf.")

    if skipped:
        logger.warning(
            "Master list: %d row(s) skipped (non-numeric or wrong-length candidate number).",
            skipped,
        )

    if not rows:
        raise ValueError(
            "No valid candidate numbers found in master list. "
            "Check that the file has a column of 10-digit candidate numbers."
        )

    return rows


def _load_excel(path: Path) -> tuple[dict[str, str], int]:
    """Load from .xlsx/.xls using openpyxl (lightweight, no pandas)."""
    try:
        import openpyxl
    except ImportError:
        raise ValueError(
            "openpyxl is required to read Excel files. "
            "Install it with: pip install openpyxl"
        )

    wb = openpyxl.load_workbook(str(path), read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)

    # First row = headers
    try:
        headers = [str(h).strip() if h is not None else "" for h in next(rows_iter)]
    except StopIteration:
        raise ValueError("Excel file appears to be empty.")

    cand_idx = _find_col_index(headers, ["Candidate Number", "CandNo", "Candidate No", "ID"])
    name_idx = _find_col_index(headers, ["Name", "Full Name", "Candidate Name"])

    if cand_idx is None:
        raise ValueError(
            "Master list must have a 'Candidate Number' column (or similar). "
            f"Columns found: {headers}"
        )

    result: dict[str, str] = {}
    skipped = 0
    for row in rows_iter:
        raw = row[cand_idx] if cand_idx < len(row) else None
        num = str(raw).strip().replace(" ", "") if raw is not None else ""
        if not num.isdigit() or len(num) != CAND_NUM_LENGTH:
            skipped += 1
            continue
        name = str(row[name_idx]).strip() if (name_idx is not None and name_idx < len(row)) else ""
        result[num] = name

    wb.close()
    return result, skipped


def _load_csv(path: Path) -> tuple[dict[str, str], int]:
    """Load from .csv using the standard library csv module."""
    result: dict[str, str] = {}
    skipped = 0

    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        if reader.fieldnames is None:
            raise ValueError("CSV file appears to be empty.")

        headers = [h.strip() for h in reader.fieldnames]
        cand_col = _find_col_name(headers, ["Candidate Number", "CandNo", "Candidate No", "ID"])
        name_col = _find_col_name(headers, ["Name", "Full Name", "Candidate Name"])

        if not cand_col:
            raise ValueError(
                "Master list must have a 'Candidate Number' column (or similar). "
                f"Columns found: {headers}"
            )

        for row in reader:
            num = row.get(cand_col, "").strip().replace(" ", "")
            if not num.isdigit() or len(num) != CAND_NUM_LENGTH:
                skipped += 1
                continue
            name = row.get(name_col, "").strip() if name_col else ""
            result[num] = name

    return result, skipped


def _load_pdf_master_list(path: Path) -> tuple[dict[str, str], int]:
    """
    Load a CXC-issued PDF master list.

    pypdf extracts the table with columns in reverse order from the visual layout:
        DD/MM/YYYYM/F  SURNAME, FIRSTNAMES  CANDNUMBER  Y/N

    For example:
        16/09/2009M ISHMAEL, CHRISTOPHER OBEDIAH MARCUS 1600120013 Y

    The candidate number sits at the END of each row, not the beginning.
    Names are in CXC format: SURNAME, FIRSTNAME MIDDLENAMES (all caps).
    Students may have multiple middle names — the full name is preserved.

    pypdf quirks handled here:
    - Gender may touch the date or have irregular spacing after it
    - Rows may be split across lines (text is normalised before matching)
    - A Y/N eligibility flag follows the candidate number — used as anchor

    Rules:
    - Any line matching the date+gender+name+10-digit+Y/N pattern is a candidate row.
    - Duplicate candidate numbers are silently overwritten (same candidate
      may appear once per subject in multi-subject master lists).
    """
    try:
        from pypdf import PdfReader
    except ImportError:
        raise ValueError(
            "pypdf is required to read PDF master lists. "
            "Install it with: pip install pypdf"
        )

    result: dict[str, str] = {}
    skipped = 0

    try:
        reader = PdfReader(str(path))
    except Exception as exc:
        raise ValueError(f"Could not open PDF: {exc}")

    # Actual pypdf output (confirmed from live extraction):
    #   16/09/2009MABERDEEN, CHRISTOPHER OBEDIAH ISHMAEEL1600120013 Y
    #
    # Key facts:
    #   - NO space between gender [MF] and the start of the name
    #   - NO space between the end of the name and the 10-digit number
    #   - A space + Y/N eligibility flag follows the number — reliable right anchor
    #   - Names are ALL CAPS, may contain comma, spaces, hyphens, apostrophes
    #
    # Strategy: anchor left on date+gender, anchor right on number+space+Y/N,
    # capture everything in between as the name.
    ROW_RE = re.compile(
        r"\d{2}/\d{2}/\d{4}[MF]"       # date + gender (no space after — confirmed)
        r"([A-Z][A-Z ,.'\-]+?)"           # name: uppercase letters, comma, space, hyphen
        r"(\d{10})\s+[YN]",              # 10-digit number (no space before) + Y/N flag
    )

    for page in reader.pages:
        try:
            text = page.extract_text() or ""
        except Exception:
            continue

        for m in ROW_RE.finditer(text):
            name = m.group(1).strip().rstrip(",")
            num  = m.group(2).strip()

            if len(num) != CAND_NUM_LENGTH:
                skipped += 1
                continue

            # Store the full title-cased name in CXC format (SURNAME, FIRSTNAMES).
            # Do NOT pre-reverse here — verify_candidate_against_master_list
            # calls normalise_cxc_name itself, which handles multiple middle names.
            result[num] = name.title()

    if not result:
        raise ValueError(
            "No candidate rows found in the PDF. "
            "Make sure this is a CXC master list with 10-digit candidate numbers."
        )

    return result, skipped


def _find_col_index(headers: list[str], candidates: list[str]) -> int | None:
    """Case-insensitive column index search."""
    lower_headers = [h.lower() for h in headers]
    for candidate in candidates:
        try:
            return lower_headers.index(candidate.lower())
        except ValueError:
            continue
    return None


def _find_col_name(headers: list[str], candidates: list[str]) -> str | None:
    """Case-insensitive column name search."""
    lower_map = {h.lower(): h for h in headers}
    for candidate in candidates:
        match = lower_map.get(candidate.lower())
        if match:
            return match
    return None


# ---------------------------------------------------------------------------
# File renaming / copying
# ---------------------------------------------------------------------------

@dataclass
class RenameJob:
    source_path: Path
    candidate_number: str
    moderation_code: str
    file_type: str            # "SBA" | "Cover Sheet" | "Mark Scheme"
    doc_number: int = 1
    output_dir: Optional[Path] = None


@dataclass
class RenameResult:
    ok: bool
    source: Path
    dest: Optional[Path] = None
    new_name: str = ""
    error: str = ""


def process_job(job: RenameJob, overwrite: bool = False) -> RenameResult:
    """
    Copy source file to output_dir with the CXC-required name.

    Args:
        job:       The rename job descriptor.
        overwrite: If False (default) and the destination already exists,
                   return an error rather than silently overwriting.
    """
    # Validate inputs
    ok, err = validate_candidate_number(job.candidate_number)
    if not ok:
        return RenameResult(ok=False, source=job.source_path, error=err)
    ok, err = validate_moderation_code(job.moderation_code)
    if not ok:
        return RenameResult(ok=False, source=job.source_path, error=err)

    # Verify source still exists
    if not job.source_path.exists():
        return RenameResult(
            ok=False,
            source=job.source_path,
            error=f"Source file no longer exists: {job.source_path.name}",
        )

    extension = job.source_path.suffix.lower() or ".pdf"
    new_name = build_filename(
        job.candidate_number,
        job.moderation_code,
        job.file_type,
        job.doc_number,
        extension,
    )

    out_dir = job.output_dir or job.source_path.parent
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        return RenameResult(
            ok=False,
            source=job.source_path,
            error=f"Cannot create output folder: {exc}",
        )

    dest = out_dir / new_name

    # Overwrite guard — prevents silent data loss
    if dest.exists() and not overwrite:
        return RenameResult(
            ok=False,
            source=job.source_path,
            new_name=new_name,
            error=f"Destination already exists: {new_name}",
        )

    try:
        shutil.copy2(job.source_path, dest)
    except OSError as exc:
        return RenameResult(
            ok=False,
            source=job.source_path,
            new_name=new_name,
            error=f"Copy failed: {exc}",
        )

    return RenameResult(ok=True, source=job.source_path, dest=dest, new_name=new_name)
