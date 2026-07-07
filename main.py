from datetime import datetime, timedelta
from pathlib import Path
import logging

import pandas as pd
import pyodbc
import win32com.client as win32

from config import CONNECTION_STRING


LOG_DIR = Path("logs")
OUTPUT_DIR = Path("outputs")
SEGMENTS = ["SEGMENT_A", "SEGMENT_B"]

LOG_DIR.mkdir(exist_ok=True)

logging.basicConfig(
    filename=LOG_DIR / "sales_report_pipeline.log",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    encoding="utf-8",
)


def get_target_date(today: datetime | None = None):
    today = today or datetime.today()

    # Sunday = 6. The reporting process is not operational on Sundays.
    if today.weekday() == 6:
        raise SystemExit("Process stopped: Sunday is not an operational day.")

    # Monday = 0. Monday processes Saturday data.
    if today.weekday() == 0:
        return (today - timedelta(days=2)).date()

    # Tuesday to Saturday process the previous day.
    return (today - timedelta(days=1)).date()


def get_period_key(target_date):
    return int(target_date.strftime("%Y%m"))


def get_query():
    return """
    SELECT
        A.report_date,
        A.customer_identifier AS DNI,
        B.customer_document_id,
        B.last_name,
        B.second_last_name,
        B.first_name,
        A.email,
        A.phone_number,
        B.eligible_amount,
        B.interest_rate,
        CONVERT(INT, REPLACE(A.installment_count, ' installments', '')) AS installment_count,
        A.agent_code,
        A.Observation,
        'YES' AS offer_flag,
        A.validation_status,
        B.segment_code
    FROM SALES_VALIDATION_TABLE A
    INNER JOIN CUSTOMER_MASTER_TABLE B
        ON A.customer_identifier = B.customer_identifier
        AND A.report_date BETWEEN B.valid_from
        AND ISNULL(B.valid_to, CONVERT(DATE, GETDATE()))
    WHERE A.period_key = ?
      AND A.report_date = ?
      AND B.segment_code = ?
    ORDER BY A.validation_status;
    """


def extract_sales_data(conn, target_date, segment):
    query = get_query()
    period_key = get_period_key(target_date)

    df = pd.read_sql(query, conn, params=[period_key, target_date, segment])
    extracted_rows = len(df)

    print(f"Extracted records for {segment}: {extracted_rows}")
    logging.info(f"Extracted records for {segment}: {extracted_rows}")

    return df


def validate_pending_status(dataframes_by_segment):
    total_pending = 0
    pending_by_segment = {}

    for segment, df in dataframes_by_segment.items():
        if df.empty or "validation_status" not in df.columns:
            pending_count = 0
        else:
            statuses = df["validation_status"].astype("string").str.strip().str.upper()
            pending_count = (statuses == "PENDING").sum()

        total_pending += pending_count
        pending_by_segment[segment] = pending_count

        print(f"Pending records detected for {segment}: {pending_count}")
        logging.info(f"Pending records detected for {segment}: {pending_count}")

    if total_pending > 0:
        for segment, pending_count in pending_by_segment.items():
            logging.warning(f"Pending records detected for {segment}: {pending_count}")

        message = (
            f"WARNING: Process stopped. There are {total_pending} records "
            "with validation_status = Pending. Files are not exported and email is not sent."
        )

        print(message)
        logging.warning(message)
        return False

    return True


def remove_rejected_observed_and_duplicates(dataframes_by_segment):
    clean_dataframes = {}

    for segment, df in dataframes_by_segment.items():
        extracted_rows = len(df)

        if df.empty:
            removed_status_rows = 0
            removed_duplicates = 0
            final_rows = 0
            clean_dataframes[segment] = df.copy()
        else:
            statuses = df["validation_status"].astype("string").str.strip().str.upper()
            df_without_rejected = df[~statuses.isin(["REJECTED", "OBSERVED"])].copy()

            removed_status_rows = extracted_rows - len(df_without_rejected)
            rows_before_deduplication = len(df_without_rejected)
            df_final = df_without_rejected.drop_duplicates(subset=["DNI"], keep="last")

            removed_duplicates = rows_before_deduplication - len(df_final)
            final_rows = len(df_final)
            clean_dataframes[segment] = df_final

        print(f"Rejected/Observed records removed for {segment}: {removed_status_rows}")
        print(f"Duplicate customer identifiers removed for {segment}: {removed_duplicates}")
        print(f"Final records for {segment}: {final_rows}")

        logging.info(f"Extracted records for {segment}: {extracted_rows}")
        logging.info(f"Rejected/Observed records removed for {segment}: {removed_status_rows}")
        logging.info(f"Duplicate customer identifiers removed for {segment}: {removed_duplicates}")
        logging.info(f"Final records for {segment}: {final_rows}")

    return clean_dataframes


def validate_observations(dataframes_by_segment):
    total_invalid = 0

    for segment, df in dataframes_by_segment.items():
        if df.empty:
            print(f"Invalid observations for {segment}: 0")
            logging.info(f"Invalid observations for {segment}: 0")
            continue

        if "Observation" not in df.columns:
            message = (
                f"ERROR: Observation column does not exist in {segment} data. "
                "The process will stop."
            )
            print(message)
            logging.error(message)
            return False

        invalid_observation = (
            df["Observation"].isna()
            | df["Observation"].astype("string").str.strip().eq("")
        )

        invalid_rows = df[invalid_observation]
        invalid_count = len(invalid_rows)

        print(f"Invalid observations for {segment}: {invalid_count}")
        logging.info(f"Invalid observations for {segment}: {invalid_count}")

        if invalid_count > 0:
            total_invalid += invalid_count

            message = (
                f"WARNING: {segment} has {invalid_count} records "
                "with empty or null Observation values."
            )
            print(message)
            logging.warning(message)

            if "DNI" in invalid_rows.columns:
                invalid_identifiers = (
                    invalid_rows["DNI"].astype("string").fillna("MISSING_ID").tolist()
                )

                print(f"Customer identifiers with invalid Observation: {invalid_identifiers}")
                logging.warning(
                    f"Customer identifiers with invalid Observation in {segment}: "
                    f"{invalid_identifiers}"
                )

    if total_invalid > 0:
        message = (
            f"Process stopped: there are {total_invalid} records with empty or null "
            "Observation values. Files are not generated and email is not sent."
        )
        print(message)
        logging.warning(message)
        return False

    logging.info("Observation validation completed successfully.")
    return True


def export_excel(df, target_date, segment):
    OUTPUT_DIR.mkdir(exist_ok=True)

    date_token = target_date.strftime("%Y%m%d")
    file_name = f"sales_report_{date_token}_{segment}.xlsx"
    file_path = OUTPUT_DIR / file_name

    internal_columns = ["validation_status", "segment_code"]
    df_export = df.drop(columns=internal_columns, errors="ignore")

    df_export.to_excel(file_path, index=False)

    return file_path


def send_outlook_email(file_paths, target_date):
    from config import EMAIL_CC, EMAIL_SUBJECT_PREFIX, EMAIL_TO

    date_token = target_date.strftime("%Y%m%d")

    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)

    mail.To = EMAIL_TO
    mail.CC = EMAIL_CC
    mail.Subject = f"{EMAIL_SUBJECT_PREFIX} {date_token}"

    mail.Body = (
        "Hello,\n\n"
        "Please find attached the generated sales report files.\n\n"
        f"Reporting date: {target_date}\n\n"
        "Regards."
    )

    for file_path in file_paths:
        mail.Attachments.Add(str(file_path.resolve()))

    mail.Send()


def main():
    try:
        target_date = get_target_date()
        extracted_dataframes = {}
        generated_files = []

        print("Starting sales report automation pipeline")
        print(f"Target date: {target_date}")

        logging.info("Starting sales report automation pipeline")
        logging.info(f"Target date: {target_date}")

        with pyodbc.connect(CONNECTION_STRING) as conn:
            for segment in SEGMENTS:
                print(f"Processing {segment}...")
                logging.info(f"Processing {segment}...")

                df = extract_sales_data(conn, target_date, segment)
                extracted_dataframes[segment] = df

                if df.empty:
                    message = f"Warning: {segment} has no records for {target_date}."
                    print(message)
                    logging.warning(message)

        statuses_are_valid = validate_pending_status(extracted_dataframes)

        if not statuses_are_valid:
            return

        clean_dataframes = remove_rejected_observed_and_duplicates(extracted_dataframes)

        observations_are_valid = validate_observations(clean_dataframes)

        if not observations_are_valid:
            return

        for segment, df in clean_dataframes.items():
            file_path = export_excel(df, target_date, segment)
            generated_files.append(file_path)

            print(f"{segment}: {len(df)} records exported.")
            print(f"Generated file: {file_path}")

            logging.info(f"{segment}: {len(df)} records exported.")
            logging.info(f"Generated file: {file_path}")

        if not generated_files:
            message = "No files were generated. Email will not be sent."
            print(message)
            logging.warning(message)
            return

        send_outlook_email(generated_files, target_date)

        print("Email sent successfully.")
        print("Process completed successfully.")

        logging.info("Email sent successfully.")
        logging.info("Process completed successfully.")

    except SystemExit as controlled_exit:
        print(controlled_exit)
        logging.warning(controlled_exit)

    except Exception as error:
        print("Process error.")
        print(error)
        logging.exception("Error in sales report automation pipeline")


if __name__ == "__main__":
    main()
