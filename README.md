# Document Formatter – Python (pywin32)

## Description

This Python script reads a Microsoft Word `.docx` file, processes its entire text content using hardcoded rules, and generates a new `.docx` file with corrections applied.

It uses the `pywin32` library to automate Microsoft Word through COM. No AI models, APIs, or external libraries like `python-docx` are used.

The formatting rules are:
- Apply UK spelling (e.g., "organize" → "organise")
- Replace "eg" with "for example" (only outside quotes)
- Skip modifications inside quotation marks
- Preserve proper nouns like "World Health Organization"
- Add periods to initials (e.g., "D" → "D.")
- Shorten names after their first full mention (e.g., "Dr Manmohan Singh" → "Dr Singh")
- If two people share the same last name (e.g., Nawaz and Shehbaz Sharif), keep full names to avoid confusion

The corrected document is saved as a new Word file with the prefix `Corrected:` at the top.

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
python "absolute_path\format_text.py" "absolute_path\input.docx" "absolute_path\output.docx"
```
# Demo Video 

Link:- https://drive.google.com/file/d/13ABZX9jGPfv3Uxqact7c8L2quGOjXNJA/view?usp=sharing

We can also use the absolute path here to run the script considering both the input.docx and output.docx are in the same folder and same path as the main folder

