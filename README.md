# EPR Sales Automation

## Configuration (edit `config.json`)
```json
{
  "inputExcel": "ERP_automations.xlsx",
  "sheetName": "Main Sheet",
  "outputExcel": "ERP_automations_output.xlsx",
  "max_rows": 0
}
```

Set `max_rows` to a positive number to limit how many rows are processed. Use `0` or remove the key to process all rows.

## Log file names (match the output file name)
All log files are created using the output file base name so they are easy to track:

- `ERP_automations_output_log.csv` (all processed rows: filled + failed)
- `ERP_automations_output_filled.csv` (only filled rows)

## Notes
- Only rows that are filled successfully are written to the filled output CSV.
- Logs append across runs.
