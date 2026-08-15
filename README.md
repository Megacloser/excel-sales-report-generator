# Sales Report Generator

A Python script that turns a raw sales CSV export into a polished, ready-to-share Excel report — automatically grouping totals by customer and formatting the workbook for presentation.

## Features

- Reads a CSV of transactions (`Customer, Product, Amount, Date`)
- Groups and aggregates order count and total spend per customer
- Produces a two-sheet Excel workbook:
  - **Customer Report** — summary table, totals row, currency formatting, frozen header, autofilter, and a bar chart of spend by customer
  - **Sales Data** — the full formatted transaction log
- Clean, presentation-ready styling: colored headers, borders, auto-sized columns, currency formatting
- Zero manual Excel work required — run one command and get a finished report

## Requirements

- Python 3.9+
- `pandas`
- `openpyxl`

Install dependencies:

```bash
pip install pandas openpyxl
```

## Usage

```bash
python generate_sales_report.py [input_csv] [output_xlsx]
```

Defaults to `sample_sales_data.csv` → `sales_report.xlsx` if no arguments are given.

Example:

```bash
python generate_sales_report.py sample_sales_data.csv sales_report.xlsx
```

## Sample Data

`sample_sales_data.csv` contains 20 example transactions across 5 customers, useful for testing the script out of the box.

## Project Structure

```
excel-report/
├── generate_sales_report.py   # Main script
├── sample_sales_data.csv      # Example input data
├── sales_report.xlsx          # Generated report (output)
└── README.md
```

## Possible Extensions

- Support multiple input files / batch processing
- Add date-range filtering or monthly breakdowns
- Export a PDF summary alongside the Excel file
- Add a simple CLI with `argparse` for more options
