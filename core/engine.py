"""
SBA File Renamer — Core Engine
Shared by both the Tkinter desktop app and the Flask web app.
"""

import os
import re
import shutil
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

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
    Return the required CXC filename (without extension by default).

    SBA sample:    {cand}{mod}-{n}          e.g. 1000750100012250901-1
    Cover Sheet:   {cand}{mod}CS            e.g. 10007501000122509CS    ← wait, spec says no dash
    Mark Scheme:   {cand}{mod}-{n}MS        e.g. 100075010001225090-1MS

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
    if len(cleaned) != 10:
        return False, f"Candidate number must be exactly 10 digits (got {len(cleaned)})."
    return True, ""


def validate_moderation_code(value: str) -> tuple[bool, str]:
    if not value.strip().isdigit():
        return False, "Moderation code must be numeric."
    if len(value.strip()) != 8:
        return False, f"Moderation code must be 8 digits (got {len(value.strip())})."
    return True, ""


# ---------------------------------------------------------------------------
# Text extraction from PDF (clean text first, OCR fallback)
# ---------------------------------------------------------------------------

def extract_title_page_text(pdf_path: str | Path, pages: int = 2) -> str:
    """
    Try pypdf first (fast, works on text-based PDFs).
    Falls back to pytesseract OCR if no usable text found.
    Returns extracted text or empty string.
    """
    pdf_path = Path(pdf_path)
    text = _extract_with_pypdf(pdf_path, pages)
    if len(text.strip()) > 30:          # enough real content
        return text

    # OCR fallback
    try:
        text = _extract_with_ocr(pdf_path, pages)
    except Exception:
        pass
    return text


def _extract_with_pypdf(pdf_path: Path, pages: int) -> str:
    try:
        from pypdf import PdfReader
        reader = PdfReader(str(pdf_path))
        chunks = []
        for i, page in enumerate(reader.pages):
            if i >= pages:
                break
            chunks.append(page.extract_text() or "")
        return "\n".join(chunks)
    except Exception:
        return ""


def _extract_with_ocr(pdf_path: Path, pages: int) -> str:
    """Rasterise first N pages and run Tesseract."""
    try:
        from pdf2image import convert_from_path
        import pytesseract
        images = convert_from_path(str(pdf_path), first_page=1, last_page=pages, dpi=200)
        return "\n".join(pytesseract.image_to_string(img) for img in images)
    except Exception:
        return ""


# ---------------------------------------------------------------------------
# Candidate number extraction from raw text
# ---------------------------------------------------------------------------

CAND_PATTERNS = [
    # "Candidate Number: 1234567890" or "Candidate No. 1234567890"
    re.compile(r"[Cc]andidate\s+(?:[Nn](?:umber|o\.?|um\.?))[:\s#]*(\d{10})"),
    # Bare 10-digit number that looks like a CXC cand number (starts with 1)
    re.compile(r"\b(1\d{9})\b"),
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
# Master list cross-check (Excel / CSV)
# ---------------------------------------------------------------------------

def load_master_list(path: str | Path) -> dict[str, str]:
    """
    Load an Excel or CSV candidate master list.
    Expects at minimum columns: 'Candidate Number' and (optionally) 'Name'.
    Returns {candidate_number_str: name_str}.
    """
    import pandas as pd
    path = Path(path)
    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        df = pd.read_excel(path, dtype=str)
    else:
        df = pd.read_csv(path, dtype=str)

    df.columns = [c.strip() for c in df.columns]

    # Flexible column detection
    cand_col = _find_col(df, ["Candidate Number", "CandNo", "Candidate No", "ID"])
    name_col = _find_col(df, ["Name", "Full Name", "Candidate Name"])

    if not cand_col:
        raise ValueError("Master list must have a 'Candidate Number' column (or similar).")

    result = {}
    for _, row in df.iterrows():
        num = str(row[cand_col]).strip().replace(" ", "")
        if not num or not num.isdigit():
            continue
        nm = str(row[name_col]).strip() if name_col else ""
        result[num] = nm
    return result


def _find_col(df, candidates: list[str]) -> Optional[str]:
    """Case-insensitive column search."""
    lower_map = {c.lower(): c for c in df.columns}
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


def process_job(job: RenameJob) -> RenameResult:
    """Copy source file to output_dir with the CXC-required name."""
    # Validate
    ok, err = validate_candidate_number(job.candidate_number)
    if not ok:
        return RenameResult(ok=False, source=job.source_path, error=err)
    ok, err = validate_moderation_code(job.moderation_code)
    if not ok:
        return RenameResult(ok=False, source=job.source_path, error=err)

    extension = job.source_path.suffix.lower() or ".pdf"
    new_name = build_filename(
        job.candidate_number,
        job.moderation_code,
        job.file_type,
        job.doc_number,
        extension,
    )

    out_dir = job.output_dir or job.source_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)
    dest = out_dir / new_name

    shutil.copy2(job.source_path, dest)
    return RenameResult(ok=True, source=job.source_path, dest=dest, new_name=new_name)
