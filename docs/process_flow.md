# Process Flow

## 1. Date resolution

The pipeline calculates the reporting date before connecting to the database. Sundays are non-operational, Mondays process Saturday data, and Tuesday through Saturday process the previous calendar day.

## 2. SQL extraction

The process connects to SQL Server with a local `config.py` file and runs a parameterized query for each configured segment. The portfolio version uses generic table names: `SALES_VALIDATION_TABLE` and `CUSTOMER_MASTER_TABLE`.

## 3. Pending validation

After extraction, the pipeline checks the `validation_status` column. If any record is marked as `Pending`, the process stops before generating files or sending email.

## 4. Rejected/Observed removal

Records with `Rejected` or `Observed` status are removed from the exportable dataset. This keeps only records that passed the operational validation rules.

## 5. Deduplication

The cleaned dataset is deduplicated by customer identifier, represented by the example field `DNI`. When duplicates exist, the pipeline keeps the latest row in the extracted order.

## 6. Observation validation

The `Observation` column is required. Empty strings, whitespace-only values, and null values stop the process before file generation.

## 7. Excel export

Each segment is exported to an Excel file using the generic naming pattern `sales_report_YYYYMMDD_SEGMENT.xlsx`. Internal control columns such as validation status and segment code are excluded from the final file.

## 8. Email dispatch

Generated files are attached to an Outlook email and sent to recipients configured locally in `config.py`.

## 9. Logging

The pipeline writes operational logs to `logs/sales_report_pipeline.log`, including extraction counts, validation results, removed records, generated file paths, email status, and errors.
