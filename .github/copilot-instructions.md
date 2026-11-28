# Processing Word transcripts

Whenever the user asks to process files in `temp_word/`:

1. First run `python scripts/process_word.py`.
2. Review only the Persian (`fa`) rows reported by the script as containing empty `()`. Use the Arabic (`ar`) text from that same row to replace each empty pair with the appropriate honorific: `(علیه‌السلام)`, `(علیهم‌السلام)`, `(صلی‌الله‌علیه‌وآله)`, or `(رحمه‌الله)`. The row numbers are 1-based transcript data rows and exclude the table header. Do not review or modify honorifics in unreported rows.
3. Validate that all generated sync JSON files and `data.json` parse correctly and that every generated record has the `time`, `ar`, and `fa` fields.
4. Only after processing and validation succeed, remove the temporary `.docx` files from `temp_word/`.
