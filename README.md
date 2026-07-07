# Sales Report Automation Pipeline

## Executive summary

Sales Report Automation Pipeline is a sanitized portfolio version of a real-world operational reporting workflow. It automates SQL Server extraction, applies business validations, generates Excel files, sends them through Outlook, and records operational logs for traceability.

## Business problem

Operational reporting teams often need to prepare recurring sales files under strict timing and quality rules. Manual extraction, filtering, validation, deduplication, file creation, and email delivery can introduce delays and errors, especially when reports depend on weekday-specific date logic and record status checks.

## Solution overview

This project uses Python to run an end-to-end reporting pipeline. It resolves the correct reporting date, extracts segment-level data from SQL Server, blocks delivery when records are still pending, removes rejected or observed records, deduplicates customer identifiers, validates required observations, exports clean Excel files, sends the files by Outlook, and logs each step.

`DNI` is kept only as a generic example of a customer identifier field. In a production implementation, this can be renamed to match the organization's anonymized or internal customer key.

## Key features

- SQL Server extraction
- Business rule validation
- Duplicate handling by customer identifier
- Excel generation
- Outlook email automation
- Operational logging
- Weekday/date logic for recurring reports

## Tech stack

- Python
- Pandas
- pyodbc
- openpyxl
- pywin32
- SQL Server
- Outlook

## Workflow

1. Resolve the target reporting date.
2. Stop execution on Sundays.
3. Process Saturday data on Mondays.
4. Process previous-day data from Tuesday to Saturday.
5. Extract records from SQL Server for each configured segment.
6. Stop the pipeline if any record is still Pending.
7. Remove Rejected or Observed records.
8. Remove duplicate customer identifiers, keeping the latest record.
9. Validate that Observation is not empty or null.
10. Export Excel files without internal control columns.
11. Send the generated files through Outlook.
12. Register operational logs.

## Folder structure

```text
sales-report-automation-pipeline/
├── main.py
├── config.example.py
├── requirements.txt
├── README.md
├── .gitignore
└── docs/
    └── process_flow.md
```

## How to configure

Copy the example configuration file:

```powershell
Copy-Item config.example.py config.py
```

Then fill `config.py` with your local SQL Server connection values and Outlook recipients. `config.py` is intentionally ignored by Git and should never be committed.

## How to run

Activate the Conda environment:

```powershell
conda activate sales_report_env
```

Install dependencies if needed:

```powershell
pip install -r requirements.txt
```

Run the pipeline:

```powershell
python main.py
```

## Security notes

- `config.py` is ignored and must be created locally.
- `outputs/` and `logs/` are ignored.
- Generated `.xlsx` files are ignored.
- No credentials, real emails, internal server names, internal table names, company names, campaign names, or operational paths are included.

## Portfolio note

This repository is a sanitized portfolio version based on a real-world reporting automation use case.
