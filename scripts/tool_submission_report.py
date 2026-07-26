#!/usr/bin/env python3
import io
import json
import logging
import os
import sys
import smtplib
from email.message import EmailMessage
from datetime import datetime, timezone
from pathlib import Path
from typing import List

import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload
from openpyxl.utils import get_column_letter

DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
SOURCE_COLUMNS = ["SubmissionDate", "STATE", "Cal_DIST", "F_Type", "Cal_Facility", "QDC"]
COLUMN_RENAME = {
    "SubmissionDate": "Date",
    "STATE": "State",
    "Cal_DIST": "District",
    "F_Type": "Facility_Type",
    "Cal_Facility": "Facility_Name",
    "QDC": "Investigator",
}
GROUP_COLUMNS = ["State", "District", "Investigator", "Facility_Type", "Facility_Name"]

TOOL_FILES = [
    {"tool_name": "DOD", "filename": "DOD_2026_WIDE.xlsx", "reader": "excel"},
    {"tool_name": "Maternal_Log", "filename": "Maternal Log_WIDE.xlsx", "reader": "excel"},
    {"tool_name": "FIS", "filename": "FIS CASE RECORDS 2026_WIDE.csv", "reader": "csv"},
    {"tool_name": "PMSMA", "filename": "PMSMA Client Interview_WIDE.xlsx", "reader": "excel"},
    {"tool_name": "SNCU", "filename": "SNCU Index Cases Observation 2026_WIDE.xlsx", "reader": "excel"},
]


def escape_drive_query_value(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise ValueError(f"Missing required environment variable: {name}")
    return value


def build_drive_credentials() -> Credentials:
    service_account_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    service_account_file = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "credential.json").strip()

    if service_account_json:
        try:
            credentials_info = json.loads(service_account_json)
        except json.JSONDecodeError as exc:
            raise ValueError("GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON") from exc
        return Credentials.from_service_account_info(credentials_info, scopes=DRIVE_SCOPES)

    if os.path.exists(service_account_file):
        return Credentials.from_service_account_file(service_account_file, scopes=DRIVE_SCOPES)

    raise ValueError(
        "Service account credentials not found. Set GOOGLE_SERVICE_ACCOUNT_JSON "
        "or provide GOOGLE_SERVICE_ACCOUNT_FILE (default: credential.json)."
    )


def find_file_id_by_name(drive_service, folder_id: str, filename: str) -> str:
    safe_name = escape_drive_query_value(filename)
    query = (
        f"'{folder_id}' in parents and trashed = false and "
        f"name = '{safe_name}'"
    )
    response = (
        drive_service.files()
        .list(q=query, spaces="drive", fields="files(id,name)", pageSize=10)
        .execute()
    )
    files = response.get("files", [])
    if not files:
        raise FileNotFoundError(f"File not found in Drive folder: {filename}")
    return files[0]["id"]


def download_drive_file_bytes(drive_service, file_id: str) -> bytes:
    request = drive_service.files().get_media(fileId=file_id)
    file_buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(file_buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return file_buffer.getvalue()


def read_tool_file(data: bytes, reader: str) -> pd.DataFrame:
    data_stream = io.BytesIO(data)
    if reader == "excel":
        return pd.read_excel(data_stream, usecols=SOURCE_COLUMNS)
    if reader == "csv":
        return pd.read_csv(data_stream, usecols=SOURCE_COLUMNS, low_memory=False)
    raise ValueError(f"Unsupported reader: {reader}")


def load_and_standardize_data(drive_service, source_folder_id: str) -> pd.DataFrame:
    frames: List[pd.DataFrame] = []

    for item in TOOL_FILES:
        tool_name = item["tool_name"]
        filename = item["filename"]
        logging.info("Reading %s", filename)
        file_id = find_file_id_by_name(drive_service, source_folder_id, filename)
        raw_bytes = download_drive_file_bytes(drive_service, file_id)
        frame = read_tool_file(raw_bytes, item["reader"]).rename(columns=COLUMN_RENAME)
        frame["Tool_Name"] = tool_name
        frames.append(frame)

    master_df = pd.concat(frames, ignore_index=True)
    master_df["Date"] = pd.to_datetime(master_df["Date"], errors="coerce").dt.normalize()
    master_df = master_df.dropna(subset=["Date"])
    return master_df


def _build_metric_series(df: pd.DataFrame, tool_name: str, label: str) -> pd.Series:
    return (
        df.loc[df["Tool_Name"] == tool_name]
        .groupby(GROUP_COLUMNS, dropna=False)
        .size()
        .rename(f"{label}_{tool_name}")
    )


def build_report_frame(master_df: pd.DataFrame) -> pd.DataFrame:
    if master_df.empty:
        raise ValueError("Master dataframe is empty after date parsing.")

    # Date logic:
    # 1) latest_date = absolute maximum date available in all tools.
    # 2) previous_date = nearest date strictly less than latest_date.
    latest_date = master_df["Date"].max()
    previous_date = master_df.loc[master_df["Date"] < latest_date, "Date"].max()

    metrics: List[pd.Series] = []
    for item in TOOL_FILES:
        tool_name = item["tool_name"]
        metrics.append(_build_metric_series(master_df, tool_name, "Total"))
        metrics.append(
            _build_metric_series(master_df.loc[master_df["Date"] == latest_date], tool_name, "Latest_Date")
        )
        if pd.notna(previous_date):
            metrics.append(
                _build_metric_series(master_df.loc[master_df["Date"] == previous_date], tool_name, "Previous_Date")
            )
        else:
            empty_previous = pd.Series(dtype="int64", name=f"Previous_Date_{tool_name}")
            metrics.append(empty_previous)

    report_df = pd.concat(metrics, axis=1).fillna(0).astype("int64").reset_index()

    ordered_columns = GROUP_COLUMNS.copy()
    for item in TOOL_FILES:
        tool_name = item["tool_name"]
        ordered_columns.extend(
            [f"Total_{tool_name}", f"Latest_Date_{tool_name}", f"Previous_Date_{tool_name}"]
        )
    report_df = report_df[ordered_columns]
    report_df = report_df.sort_values(GROUP_COLUMNS, kind="stable").reset_index(drop=True)

    report_df["Latest_Date_Used"] = latest_date.date().isoformat()
    report_df["Previous_Date_Used"] = (
        previous_date.date().isoformat() if pd.notna(previous_date) else ""
    )
    return report_df


def autosize_worksheet(worksheet) -> None:
    for col_idx, column_cells in enumerate(worksheet.columns, start=1):
        max_length = 0
        for cell in column_cells:
            value = "" if cell.value is None else str(cell.value)
            if len(value) > max_length:
                max_length = len(value)
        worksheet.column_dimensions[get_column_letter(col_idx)].width = min(max_length + 2, 60)


def export_report(report_df: pd.DataFrame, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        report_df.to_excel(writer, sheet_name="Submission_Report", index=False)
        autosize_worksheet(writer.book["Submission_Report"])


def parse_recipients() -> List[str]:
    recipients = [
        item.strip()
        for item in required_env("RECIPIENTS").split(",")
        if item.strip()
    ]
    if not recipients:
        raise ValueError("RECIPIENTS must contain at least one email address")
    return recipients


def send_email_with_attachment(
    sender_email: str,
    app_password: str,
    recipients: List[str],
    attachment_path: Path,
) -> None:
    msg = EmailMessage()
    msg["From"] = sender_email
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = os.getenv("MAIL_SUBJECT", "Daily Tool Submission Report").strip() or "Daily Tool Submission Report"
    msg.set_content("Please find attached the latest Tool Submission Report.")

    with open(attachment_path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=attachment_path.name,
        )

    with smtplib.SMTP("smtp.gmail.com", 587, timeout=60) as server:
        server.starttls()
        server.login(sender_email, app_password)
        server.send_message(msg)


def main() -> int:
    logging.basicConfig(
        level=os.getenv("LOG_LEVEL", "INFO").upper(),
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
    try:
        source_folder_id = required_env("DRIVE_FOLDER_ID")
        run_date_name = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        local_output = Path(os.getenv("LOCAL_OUTPUT_PATH", f"/tmp/{run_date_name}/Tool_Submission_Report.xlsx"))

        credentials = build_drive_credentials()
        drive_service = build("drive", "v3", credentials=credentials, cache_discovery=False)

        master_df = load_and_standardize_data(drive_service, source_folder_id)
        report_df = build_report_frame(master_df)
        export_report(report_df, local_output)

        logging.info("EMAIL ATTACHMENT MODE ACTIVATED")

        sender_email = required_env("SMTP_EMAIL")
        app_password = required_env("SMTP_APP_PASSWORD")
        recipients = parse_recipients()

        send_email_with_attachment(
            sender_email,
            app_password,
            recipients,
            local_output,
        )

        logging.info("Report generated and emailed: %s", local_output)
        return 0
    except (ValueError, FileNotFoundError) as exc:
        logging.error("Configuration/runtime error: %s", exc)
        return 2
    except HttpError as exc:
        logging.error("Google API request failed: %s", exc)
        return 3
    except smtplib.SMTPException as exc:
        logging.error("SMTP request failed: %s", exc)
        return 4
    except Exception as exc:
        logging.exception("Unexpected failure: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
