# Document Formatter – Python (pywin32)

## Description

This Python script reads a Microsoft Word `.docx` file, applies hardcoded formatting rules, and saves the corrected version as another `.docx` file.

This solution uses the `pywin32` library to automate Microsoft Word. It does not use `python-docx` or any AI/LLM APIs.

---

## Tasks Performed

### Task 1: Language & Spelling

- Converts American English (e.g., "organize") to British English ("organise").
- Replaces "eg" with "for example" (only outside quotes).
- Preserves text inside quotation marks.
- Leaves proper nouns like "World Health Organization" unchanged.

### Task 2: Names and Acronyms

- First mention: full name with title (e.g., Dr Manmohan Singh).
- Subsequent mentions: shortened to title + last name (e.g., Dr Singh).
- If two people share the same last name, full names are used to avoid confusion.
- Periods are added to initials (e.g., Franklin D. Roosevelt).

---

## How to Run

### 1. Requirements

- **Windows OS**
- **Microsoft Word** (installed)
- Python 3.x
- Install `pywin32`:

```bash
pip install pywin32


### 2. Run the script in terminal:

'''bash
python format_text.py input.docx output.docx

### 3. Open output.docx to see the corrected text.

Open Word → paste the final **"Corrected:"** content (you pasted above) → Save as:


