# Copilot Instructions for Non-Value Add Rework Tracking Dashboard

## Project Overview
This dashboard analyzes rework data from CSV files, focusing on operational insights for manufacturing facilities. The main files are:
- `index.html`: UI and dashboard layout
- `app.js`: Core logic for data parsing, filtering, metrics, and chart rendering
- `NVA Data for Dashboard.csv`: Example data file

## Data Flow & Architecture
- CSV data is uploaded via the dashboard UI and parsed in `app.js`.
- Flexible column mapping: The dashboard auto-detects column names (case-insensitive, see README for mappings).
- Data is stored in browser `localStorage` for persistence.
- Metrics, charts, and tables are dynamically generated from parsed data.

## Key Patterns & Conventions
- **Column Flexibility**: Always use column name normalization when referencing CSV fields.
- **Date Handling**: Prefer "Hold Date" over "Production Date" if both are present.
- **Metrics Calculation**: If "Rework %" is missing, calculate as `Cases Reworked / Cases Produced * 100`.
- **Missing Data**: Show "N/A" for metrics if required fields are absent.
- **UI Updates**: All dashboard updates are triggered by CSV upload or filter changes.

## Developer Workflows
- Open `index.html` in a browser to run the dashboard.
- No build step required; all logic is in JS/HTML.
- Debug by editing `app.js` and refreshing the browser.
- No automated tests; manual validation via dashboard UI.

## Integration Points
- No external dependencies (vanilla JS, HTML, CSV parsing).
- Data is not sent to any backend; all processing is client-side.

## Example Patterns
- To add a new filter, update the filter UI in `index.html` and handle logic in `app.js`.
- To support new CSV columns, extend the column mapping logic in `app.js`.

## Reference
- See `README.md` for CSV format, column flexibility, and usage instructions.

---

If any section is unclear or missing, please provide feedback for improvement.