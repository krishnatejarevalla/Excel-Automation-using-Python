# Excel Automation with Python - AI Coding Agent Instructions

## Project Overview
This is a Python project focused on Excel automation using the `openpyxl` library. The codebase enables programmatic manipulation of Excel workbooks and worksheets without requiring Excel to be installed.

## Key Dependencies
- **openpyxl**: Primary library for reading/writing Excel files (.xlsx format)
  - Use for loading workbooks: `workbook = openpyxl.load_workbook('file.xlsx')`
  - Access worksheets via: `sheet = workbook['sheet_name']` or `workbook.active`
  - Modify cells: `sheet['A1'] = value` or `sheet.cell(row=1, column=1).value = value`

## Architecture & File Structure
- `Task.py`: Main entry point - currently contains minimal setup for openpyxl imports
- Single-file structure: All automation logic should extend Task.py or create companion modules as complexity grows

## Development Patterns

### Excel Operations
- **Reading**: Always use `load_workbook()` with `data_only=True` for cell values only, `data_only=False` for formulas
- **Writing**: Use `sheet.cell()` for programmatic access or direct indexing like `sheet['A1']`
- **Iteration**: Loop through cells via `sheet.iter_rows()` or access specific ranges

### Code Structure
- Start with imports at top (openpyxl and standard library modules)
- Keep workbook operations within context where possible
- Always close/save workbooks: `workbook.save('output.xlsx')`

## Common Workflows
1. **Load & Read**: `load_workbook()` → iterate rows/cells → extract data
2. **Modify & Save**: `load_workbook()` → find cells/ranges → update values → `save()`
3. **Create New**: `Workbook()` → add sheets → populate → `save()`

## Integration Points
- File system: Reads/writes `.xlsx` files
- No external APIs or services currently integrated
- All operations are synchronous and file-based

## Testing & Execution
- Run via: `python Task.py`
- Test with sample Excel files in same directory or subdirectories
- Verify output files are created/modified correctly after execution
