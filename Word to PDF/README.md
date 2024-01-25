## Word to PDF
Uses applescript to interact with Word

### Features
1. Batch processing
2. ❌ (currently not working) Auto-move word files to a subfolder
3. PDFs are saved with their extension appeneded as `-doc` or `-docx` for clarity

### How?
1. Iterates through each file in input
2. Opens file in Word
3. Exports as PDF in original folder
4. Checks if subfolder exists (if not, create)
5. Move DOCX/DOC to folder
6. Repeat with each in selection
7. Close Word
