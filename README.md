# Document Formatter – Python (pywin32)

## Description

This Python script reads a Microsoft Word `.docx` file, applies hardcoded formatting rules, and generates a corrected `.docx` file.

It uses `pywin32` to control Microsoft Word via COM automation. It does **not** use `python-docx`, AI, or LLMs.

---

## Tasks Performed

### Task 1: Language & Spelling
- Converts "organize", "organizes", etc. → British English equivalents (e.g., "organise").
- Replaces "eg" with "for example" (only outside quotation marks).
- Skips content inside quotes.
- Preserves proper nouns such as "World Health Organization".

### Task 2: Names and Acronyms
- First mention: full name (e.g., Dr Manmohan Singh).
- Later mentions: shortened to "Dr Singh".
- If two people share a last name (e.g., Nawaz and Shehbaz Sharif), full names are retained.
- Adds periods to initials (e.g., Franklin D Roosevelt → Franklin D. Roosevelt).

---

## Requirements

- Windows OS
- Microsoft Word (installed and activated)
- Python 3.x
- `pywin32` library

## Install required package
```bash
pip install pywin32
```

### Run the script with your input and output Word files
```bash
python format_text.py input.docx output.docx
```
# Example:
```bash
python format_text.py input.docx output.docx
```
We can also use the absolute path considering both the input.docx and output.docx are in the same folder and same path as the main folder





