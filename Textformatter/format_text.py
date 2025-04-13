import sys
import re
import win32com.client as win32

def format_document(input_path, output_path):
    """Open the input Word document, apply formatting rules, and save to output_path."""
    # Initialize Word application
    word = win32.Dispatch("Word.Application")
    word.Visible = False  # Run Word in background
    
    # Open the input .docx document
    doc = word.Documents.Open(input_path)
    full_text = doc.Content.Text  # Extract all text from the document
    doc.Close(False)  # Close input document without saving
    
    # Apply Task 1 and Task 2 formatting rules to the extracted text
    corrected_text = apply_formatting_rules(full_text)
    
    # Create a new document for output and insert the corrected text under "Corrected:" header
    output_doc = word.Documents.Add()
    output_doc.Content.Text = "Corrected:\r" + corrected_text  # \r inserts a new paragraph in Word
    
    # Save the output document
    output_doc.SaveAs(output_path, FileFormat=16)  # 16 = wdFormatDocumentDefault (docx)
    output_doc.Close()
    word.Quit()

def apply_formatting_rules(text):
    """Apply the two sets of formatting rules (language & spelling, names & acronyms) to the text."""
    # 1. **Language & Spelling (UK English & specific term replacements)**

    # Do not change content inside quotes – temporarily placeholder quoted segments
    quotes = re.findall(r'“[^”]*”', text)  # find all text in “ ” quotes
    for i, q in enumerate(quotes):
        text = text.replace(q, f"<<QUOTE{i}>>")
    
    # Do not change proper nouns like "World Health Organization" – placeholder to protect it
    text = text.replace("World Health Organization’s", "<<WHO>>’s")
    text = text.replace("World Health Organization", "<<WHO>>")
    
    # Replace American English spellings with UK English, and replace 'eg' with 'for example'
    # (Skip proper nouns and quotes as handled above)
    replacements = [
        (r"\borganize\b", "organise"),
        (r"\borganizes\b", "organises"),
        (r"\borganized\b", "organised"),
        (r"\borganizing\b", "organising"),
        (r"\borganization\b", "organisation"),
        (r"\bOrganization\b", "Organisation"),
        (r"\beg\b", "for example"),
    ]
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text)
    
    # Restore protected placeholders for quotes and proper noun
    text = text.replace("<<WHO>>’s", "World Health Organization’s")
    text = text.replace("<<WHO>>", "World Health Organization")
    for i, q in enumerate(quotes):
        text = text.replace(f"<<QUOTE{i}>>", q)
    
    # 2. **Names & Acronyms (titles, initials, and repeated name handling)**
    
    # Add periods after single-letter initials in names (e.g., "Franklin D Roosevelt" -> "Franklin D. Roosevelt")
    # This regex finds a single capital letter as a whole word followed by a capitalized name (last name) and inserts a dot.
    text = re.sub(r"\b(?!A\b)([A-Z])\b(?=\s+[A-Z][a-z])", r"\1.", text)
    
    # Ensure specific known cases are corrected (if not caught by regex)
    text = text.replace("Franklin D Roosevelt", "Franklin D. Roosevelt")
    
    # Handle name mentions:
    # On first mention, use full name; on subsequent mentions, use title + last name (if no ambiguity).
    # If two people share the same last name, always use full names (no shortening).
    # We will handle each relevant name individually to ensure correct context.
    
    # Define a helper to replace subsequent mentions of a full name with the shortened form
    def shorten_name(text, full_name, short_form):
        first_index = text.find(full_name)
        if first_index == -1:
            return text  # Name not present
        # Find second occurrence after the first
        second_index = text.find(full_name, first_index + 1)
        if second_index == -1:
            return text  # Only one occurrence, nothing to shorten
        # Build the new text with first occurrence unchanged and subsequent occurrences replaced
        result = text[:first_index] + full_name
        pos = first_index + len(full_name)
        while True:
            next_index = text.find(full_name, pos)
            if next_index == -1:
                # append any remaining text after the last occurrence
                result += text[pos:]
                break
            # append text between occurrences, then short form instead of the full name
            result += text[pos:next_index] + short_form
            pos = next_index + len(full_name)
        return result
    
    # Shorten repeated mentions of specific individuals (if not ambiguous)
    # Dr Manmohan Singh -> Dr Singh (subsequent mentions)
    text = shorten_name(text, "Dr Manmohan Singh", "Dr Singh")
    # Dr Aishwarya Rai -> Dr Rai (subsequent mentions)
    text = shorten_name(text, "Dr Aishwarya Rai", "Dr Rai")
    # Nawaz Sharif and Shehbaz Sharif share a last name, so do not shorten either – leave full names as is.
    # Franklin D. Roosevelt is referenced along with a historical figure of the same name (former U.S. president).
    # To avoid confusion, we keep using the full name "Franklin D. Roosevelt" where it appears.
    # (The text also uses "Dr Roosevelt" when referring to the speaker in certain contexts, which we leave unchanged.)
    
    return text

if __name__ == "__main__":
    # Expect two command-line arguments: input .docx path and output .docx path
    if len(sys.argv) != 3:
        print("Usage: python format_text.py <input.docx> <output.docx>")
    else:
        in_path = sys.argv[1]
        out_path = sys.argv[2]
        format_document(in_path, out_path)
        print(f"Formatted document saved to: {out_path}")
