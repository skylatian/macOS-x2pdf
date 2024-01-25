## PPTX to PDF
Uses applescript to interact with powerpoint: iterates through each file, exports them as PDF, and moves each file to a subfolder.

### Features
1. Batch processing
2. Auto-move pptx files to a subfolder (default behavior, modify workflow to change)
3. PDFs are saved with their extension appeneded as `-ppt` or `-pptx` for clarity

### How?
1. Iterates through each file in input
2. Opens file in PowerPoint
3. Exports as PDF in original folder
4. Checks if subfolder exists (if not, create)
5. Move PPTX to folder
6. Repeat with each in selection
7. Close PowerPoint