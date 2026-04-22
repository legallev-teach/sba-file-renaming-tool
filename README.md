# SBA File Renamer
**Caribbean Union College Secondary — May/June CSEC Examinations**

Renames SBA PDF files to the exact format required by CXC's Online Registration System (ORS).

---

## Naming format (from CXC memorandum E:15/12/13)

| File type    | Pattern                             | Example                          |
|--------------|-------------------------------------|----------------------------------|
| SBA sample   | `{CandNo}{ModCode}-{n}.pdf`         | `100075010001225090-1.pdf`       |
| Cover Sheet  | `{CandNo}{ModCode}CS.pdf`           | `10007501000122509CS.pdf`        |
| Mark Scheme  | `{CandNo}{ModCode}-{n}MS.pdf`       | `100075010001225090-1MS.pdf`     |

- **CandNo** — 10 digits (from ORS candidate registration)
- **ModCode** — 8 digits (subject-specific, see Appendix in memorandum)
- **n** — document number: `-1` for first file, `-2` for second, `-3` for third

---

## Project structure

```
sba_renamer/
├── core/
│   └── engine.py          # Shared logic: subjects, naming, extraction, renaming
├── desktop/
│   └── desktop_app.py     # Tkinter GUI → package to .exe with PyInstaller
├── web/
│   ├── web_app.py          # Flask backend
│   └── templates/
│       └── index.html     # Web UI
├── requirements.txt
└── README.md
```

---

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

**OCR support** (optional — only needed for scanned PDFs):
- Install [Poppler](https://poppler.freedesktop.org/) (provides `pdftoppm` used by `pdf2image`)
- Install [Tesseract-OCR](https://github.com/UB-Mannheim/tesseract/wiki)
- On Windows, ensure both are on your PATH

### 2a. Run the desktop app (Tkinter)

```bash
python desktop/desktop_app.py
```

### 2b. Run the web app (Flask)

```bash
python web/web_app.py
```

Then open `http://localhost:5000` in a browser.
For school-wide access, open port 5000 on the server firewall and share the server's local IP.

---

## Packaging to .exe (Windows laptop)

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "SBA_Renamer" desktop/desktop_app.py
```

Output: `dist/SBA_Renamer.exe` — no Python installation required on target machine.

---

## CSEC subjects supported (online-moderated, from memorandum appendix)

| Subject | Moderation Code |
|---------|----------------|
| Additional Mathematics | 01254090 |
| Caribbean History | 01210090 |
| Economics | 01216090 |
| Electronic Document Preparation and Management | 01251090 |
| English A | 01218090 |
| English B | 01219090 |
| Geography | 01225090 |
| Human and Social Biology | 01253090 |
| Information Technology | 01229090 |
| Mathematics | 01234090 |
| Office Administration | 01237090 |
| Physical Education and Sport | 01252090 |
| Principles of Accounts | 01239090 |
| Principles of Business | 01240090 |
| Religious Education | 01241090 |
| Social Studies | 01243090 |
| Theatre Arts | 01248090 |

> CAPE subject codes can be added to `core/engine.py` later — the architecture is designed for it.

---

## Adding CAPE support later

In `core/engine.py`, add a second dictionary:

```python
CAPE_SUBJECTS: dict[str, dict] = {
    "Accounting Unit 1": {"code": "02101090"},
    "Accounting Unit 2": {"code": "02201090"},
    # ... etc
}
```

Then update the UI dropdowns to offer a CSEC/CAPE toggle.

---

## Known limitations / deferred features

- **Group SBAs** — not yet supported. Will need a multi-candidate input mode.
- **Bulk batch from spreadsheet** — deferred. Could read a CSV of all candidates and process in one run.
- **CAPE subjects** — architecture ready, codes not yet entered.
