# SBA File Renamer

A desktop tool for renaming CSEC SBA PDF files to the format required by the Caribbean Examinations Council (CXC).

Built for teachers at Caribbean Union College Secondary School, Maracas, St. Joseph, Trinidad — and freely available to any school doing the same work.

---

## What It Does

CXC requires SBA files to be submitted with a specific filename format based on the candidate number, subject moderation code, document type, and document number. Doing this manually for a class of 30+ students is error-prone and time-consuming.

This tool:
- Extracts candidate numbers automatically from PDF title pages
- Falls back to OCR if the PDF is image-based (requires Tesseract)
- Falls back to filename pattern matching if extraction fails
- Optionally cross-references a master candidate list (Excel or CSV)
- Previews the output filename before any files are touched
- Copies (never moves) files to a standardised name in your chosen output folder

---

## Filename Format

| Document Type | Format | Example |
|---|---|---|
| SBA | `{candidate}{modcode}-{n}.pdf` | `1000750100012250901-1.pdf` |
| Cover Sheet | `{candidate}{modcode}CS.pdf` | `100075010001225090CS.pdf` |
| Mark Scheme | `{candidate}{modcode}-{n}MS.pdf` | `100075010001225090-1MS.pdf` |

Based on the CXC MoE Memorandum for May–June 2026 CSEC submissions.

---

## Requirements

### Python version
Python 3.10 or later.

### Python packages
Install with:
```
pip install -r requirements.txt
```

### Tesseract OCR (optional)
Only needed if your PDFs are image-based (scanned rather than digitally created).

- **Windows:** Download the installer from [UB-Mannheim/tesseract](https://github.com/UB-Mannheim/tesseract/wiki) and add it to your PATH.
- **macOS:** `brew install tesseract`
- **Linux:** `sudo apt install tesseract-ocr`

If Tesseract is not installed, the app will warn you on startup and continue using text-based extraction only.

---

## Installation

### Option A — Run from source
```bash
git clone https://github.com/YOUR_USERNAME/sba-file-renamer.git
cd sba-file-renamer
pip install -r requirements.txt
python desktop_app.py
```

### Option B — Windows executable
Download `SBARenamer.exe` from the [Releases](../../releases) page. No Python installation required.

> **Note:** The `.exe` does not include Tesseract. If you need OCR, install Tesseract separately and ensure it is on your PATH.

---

## Building the Executable (Windows)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name SBARenamer desktop_app.py
```

The `.exe` will be in the `dist/` folder.

---

## Project Structure

```
sba-file-renamer/
├── core/
│   └── engine.py          # All business logic — naming, validation, extraction, file ops
├── desktop_app.py         # Tkinter desktop UI
├── requirements.txt
├── README.md
└── LICENSE
```

The engine has no UI dependencies and can be reused in a CLI or web context.

---

## Supported Subjects

All CSEC subjects listed in the CXC MoE Memorandum (May–June 2026 appendix), including:

Additional Mathematics · Caribbean History · Economics · EDPM · English A · English B · Geography · Information Technology · Mathematics · Office Administration · Physical Education and Sport · Principles of Accounts · Principles of Business · Religious Education · Social Studies · Theatre Arts · Human and Social Biology

---

## Master Candidate List

You can optionally load an Excel (`.xlsx`) or CSV file containing your candidates. The file must have a column of 10-digit candidate numbers. A column named `Name` (or similar) is optional but enables name confirmation in the UI.

Supported column names for candidate number: `Candidate Number`, `CandNo`, `Candidate No`, `ID`  
Supported column names for name: `Name`, `Full Name`, `Candidate Name`

---

## Usage Notes

- Files are **copied**, not moved. Your originals are never touched.
- If an output file already exists, the job will fail with a clear error rather than overwriting silently.
- The batch settings (subject, file type, document number) apply to all files. If your class has multiple document types, process them in separate batches.
- Files marked **?** in the list need a candidate number entered manually before they will be included in the rename run.
- Files marked **✗** had a scan error (e.g. corrupted PDF). Hover over the status bar for details, or try the ⟳ Scan button after checking the file.

---

## License

MIT — see [LICENSE](LICENSE).  
Free to use, modify, and share. Credit appreciated but not required.
