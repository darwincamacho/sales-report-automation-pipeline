SERVER = "YOUR_SQL_SERVER"
DATABASE = "YOUR_DATABASE"
USER = "YOUR_USERNAME"
PASSWORD = "YOUR_PASSWORD"

CONNECTION_STRING = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    f"SERVER={SERVER};"
    f"DATABASE={DATABASE};"
    f"UID={USER};"
    f"PWD={PASSWORD};"
    "TrustServerCertificate=yes;"
)

EMAIL_TO = "recipient@example.com"
EMAIL_CC = "copy@example.com"
EMAIL_SUBJECT_PREFIX = "sales report"
