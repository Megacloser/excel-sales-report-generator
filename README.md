# Sales Report Generator

A Python tool that turns a raw sales CSV export into a polished, ready-to-share Excel report — automatically grouping totals by customer and formatting the workbook for presentation. Comes with both a command-line script and a Streamlit web app.

## Features

- Reads a CSV of transactions (`Customer, Product, Amount, Date`)
- Column headers are matched case-insensitively and accept common synonyms (e.g. `Customer`/`customer`/`Client`/`Клиент`, `Product`/`Item`/`Товар`, `Amount`/`Sum`/`Сумма`, `Date`/`Дата`)
- Flexible date parsing — accepts `YYYY-MM-DD`, `DD.MM.YYYY`, or `MM/DD/YYYY`, including mixed formats within the same file
- Groups and aggregates order count and total spend per customer
- Produces a two-sheet Excel workbook:
  - **Customer Report** — branded banner, KPI summary line, summary table with rank and % of total, totals row, currency formatting, frozen header, autofilter, and a bar chart of spend by customer
  - **Sales Data** — the full formatted transaction log
- Clean, presentation-ready styling: colored headers, zebra-striped rows, auto-sized columns, currency formatting
- Web UI (`app.py`) for uploading a CSV, previewing the results, and downloading the Excel report — no command line needed
- Zero manual Excel work required — run one command and get a finished report

## Requirements

- Python 3.9+
- `pandas`
- `openpyxl`
- `streamlit` (for the web app)

Install dependencies:

```bash
pip install -r requirements.txt
```

## Usage

### Command line

```bash
python generate_sales_report.py [input_csv] [output_xlsx]
```

Defaults to `sample_sales_data.csv` → `sales_report.xlsx` if no arguments are given.

Example:

```bash
python generate_sales_report.py sample_sales_data.csv sales_report.xlsx
```

### Web app

```bash
streamlit run app.py
```

Opens a browser UI to upload a CSV, click **Process**, preview the customer summary and raw data, and download the generated Excel report. Invalid files (missing columns, empty file, bad encoding, unparseable CSV) are caught and reported with a clear error message.

## Sample Data

`sample_sales_data.csv` contains 20 example transactions across 5 customers, useful for testing the script out of the box.

## Project Structure

```
excel-report/
├── generate_sales_report.py   # Core script + report-building functions
├── app.py                     # Streamlit web app
├── requirements.txt           # Python dependencies
├── sample_sales_data.csv      # Example input data
├── sales_report.xlsx          # Generated report (output)
└── README.md
```

## Possible Extensions

- Support multiple input files / batch processing
- Add date-range filtering or monthly breakdowns
- Export a PDF summary alongside the Excel file
- Deploy the Streamlit app to Streamlit Community Cloud for a live demo link
