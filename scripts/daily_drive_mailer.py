#!/usr/bin/env python3
import io
import json
import logging
import os
import smtplib
import sys
from datetime import datetime, timezone
from email.message import EmailMessage
from typing import List, Tuple

import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]


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


def parse_recipients() -> List[str]:
    recipients = [item.strip() for item in required_env("RECIPIENTS").split(",") if item.strip()]
    if not recipients:
        raise ValueError("RECIPIENTS must contain at least one email address")
    return recipients


def find_latest_excel_file(drive_service, folder_id: str) -> dict:
    query = (
        f"'{folder_id}' in parents and trashed = false and "
        "(mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' "
        "or mimeType = 'application/vnd.ms-excel')"
    )
    response = (
        drive_service.files()
        .list(
            q=query,
            spaces="drive",
            fields="files(id,name,modifiedTime,mimeType)",
            orderBy="modifiedTime desc",
            pageSize=1,
        )
        .execute()
    )
    files = response.get("files", [])
    if not files:
        raise FileNotFoundError(f"No Excel files found in folder: {folder_id}")
    return files[0]


def download_file_bytes(drive_service, file_id: str) -> bytes:
    request = drive_service.files().get_media(fileId=file_id)
    file_buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(file_buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return file_buffer.getvalue()


def analyze_excel_content(file_bytes: bytes, filename: str) -> Tuple[str, str]:
    excel_data = pd.ExcelFile(io.BytesIO(file_bytes))
    lines = [
        f"Daily Excel processing result",
        f"Timestamp (UTC): {datetime.now(timezone.utc).isoformat()}",
        f"Source file: {filename}",
        f"Sheets found: {len(excel_data.sheet_names)}",
        "",
    ]

    for sheet in excel_data.sheet_names:
        frame = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet)
        lines.append(
            f"- {sheet}: rows={len(frame)}, columns={len(frame.columns)}"
        )

    body = "\n".join(lines)
    subject = os.getenv("MAIL_SUBJECT", "Daily Excel Processing Report").strip() or "Daily Excel Processing Report"
    return subject, body


def send_gmail_smtp(sender_email: str, app_password: str, recipients: List[str], subject: str, body: str) -> None:
    message = EmailMessage()
    message["From"] = sender_email
    message["To"] = ", ".join(recipients)
    message["Subject"] = subject
    message.set_content(body)

    with smtplib.SMTP("smtp.gmail.com", 587, timeout=60) as server:
        server.starttls()
        server.login(sender_email, app_password)
        server.send_message(message)


def main() -> int:
    logging.basicConfig(
        level=os.getenv("LOG_LEVEL", "INFO").upper(),
        format="%(asctime)s [%(levelname)s] %(message)s",
    )

    try:
        credentials = build_drive_credentials()
        drive_service = build("drive", "v3", credentials=credentials, cache_discovery=False)

        folder_id = required_env("DRIVE_FOLDER_ID")
        sender_email = required_env("SMTP_EMAIL")
        app_password = required_env("SMTP_APP_PASSWORD")
        recipients = parse_recipients()

        logging.info("Looking for latest Excel file in folder %s", folder_id)
        latest_file = find_latest_excel_file(drive_service, folder_id)
        logging.info(
            "Using file: %s (id=%s, modified=%s)",
            latest_file.get("name"),
            latest_file.get("id"),
            latest_file.get("modifiedTime"),
        )

        file_bytes = download_file_bytes(drive_service, latest_file["id"])
        subject, body = analyze_excel_content(file_bytes, latest_file["name"])
        send_gmail_smtp(sender_email, app_password, recipients, subject, body)
        logging.info("Email sent to %s", ", ".join(recipients))
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
