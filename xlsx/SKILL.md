---
name: xlsx
description: Create professional Microsoft Excel (.XLSX) files with consistent styling, proper formatting, and production-ready code. Use this skill whenever the user asks to create, generate, or build Excel files, spreadsheets, workbooks, or mentions .xlsx files. Make sure to use this skill for any Excel file generation task even if the user doesn't explicitly say "xlsx skill".
---

# XLSX Skill

This skill provides a deterministic, predictable way to create professional Excel files. Instead of improvising code each time, it uses a standardized approach with consistent styling and formatting.

## When to use
- User asks to create an Excel/spreadsheet file
- User needs a professional report in .xlsx format
- User wants consistent formatting (colors, number formats, headers)
- User needs repeatable, controlled Excel generation

## Standard workflow
1. Understand the data structure needed (columns, rows, calculations)
2. Use the provided template/script to generate the file
3. Apply professional styling (header colors, borders, number formatting)
4. Save as .xlsx with a descriptive filename

## Styling standards
- **Headers**: Bold white text on dark blue background (#1F4E79), borders
- **Data cells**: Clean borders, appropriate number formats
- **Currency**: Use $#,##0.00 or similar
- **Percentages**: Use 0.00%
- **Dates**: Use YYYY-MM-DD or locale-appropriate format
- **Totals/summary rows**: Bold, with top border

## How to generate
Use Python with openpyxl. The script `scripts/create_xlsx.py` provides a reusable function.

### Quick usage
```python
from scripts.create_xlsx import create_professional_excel
create_professional_excel(data, filename="report.xlsx", title="Report Title")
```

## File naming
Use descriptive names like `sales_report.xlsx`, `analysis.xlsx`, etc.

## Notes
- Always produce a valid .xlsx file
- Keep formatting consistent across all generations
- Prefer the bundled script over writing ad-hoc code
- Make the output predictable and reproducible
