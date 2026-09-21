#!/usr/bin/env python3
"""Generate the Assam GAVB Excel report from Google Drive CSV files and email it through Gmail."""

import io
import json
import logging
import os
import re
import smtplib
import ssl
import sys
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path
from typing import Optional, Sequence

import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

SCRIPT_VERSION = "2026-09-21-assam-only-gmail-v1"
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
TARGET_STATE = "ASSAM"

TOOL_FILES = [
    {"tool_name": "DOD", "filename": "DOD_2026_WIDE.csv", "reader": "csv"},
    {"tool_name": "Maternal_Log", "filename": "Maternal Log_WIDE.csv", "reader": "csv"},
    {"tool_name": "FIS", "filename": "FIS CASE RECORDS 2026_WIDE.csv", "reader": "csv"},
    {"tool_name": "PMSMA", "filename": "PMSMA Client Interview_WIDE.csv", "reader": "csv"},
    {"tool_name": "SNCU", "filename": "SNCU Index Cases Observation 2026_WIDE.csv", "reader": "csv"},
    {"tool_name": "Referral_Services", "filename": "GABV IDD Referral_WIDE.csv", "reader": "csv"},
    {"tool_name": "Digital_System", "filename": "GAVB IDD Digital System_WIDE.csv", "reader": "csv"},
    {"tool_name": "Exit_Interview", "filename": "GAVB IDD Exit Interview_WIDE.csv", "reader": "csv"},
    {"tool_name": "HR", "filename": "GAVB IDD HR_WIDE.csv", "reader": "csv"},
    {"tool_name": "Labour_Room_Readiness", "filename": "GAVB IDD Labour Room Readiness_WIDE.csv", "reader": "csv"},
    {"tool_name": "Supply_Chain", "filename": "GAVB IDD Supply Chain_WIDE.csv", "reader": "csv"},
]

TOOL_LABELS = {
    "DOD": "DOD", "Maternal_Log": "Maternal Log", "FIS": "FIS", "PMSMA": "PMSMA",
    "SNCU": "SNCU", "Referral_Services": "Referral Services", "Digital_System": "Digital System",
    "Exit_Interview": "Exit Interview", "HR": "HR", "Labour_Room_Readiness": "Labour Room Readiness",
    "Supply_Chain": "Supply Chain",
}
TOOL_ORDER = [x["tool_name"] for x in TOOL_FILES]

COLUMN_ALIASES = {
    "state": ["STATE", "State", "state", "Cal_STATE", "state_name"],
    "investigator": ["QDC", "Investigator", "Nurse", "Nurse_Name", "Nursing_Consultant",
                     "Name of Nursing Consultants", "Name of Nurses", "collector_name"],
    "facility_type": ["F_Type", "Facility_Type", "Facility Type", "facilitytype"],
    "facility_level": ["Facility_Level", "Facility Level", "Level", "DH_Below_DH"],
    "submission_date": ["SubmissionDate", "Submission Date", "submission_date", "SubmissionDateTime",
                        "Submission_Time", "starttime", "endtime"],
}

DH_VALUES = {"dh", "district hospital", "district_hospital", "district hospital dh"}
DAY_START_HOUR, DAY_END_HOUR = 9, 18
DEFAULT_STATE_SPOC = {"ASSAM": "Nikhil Kumar", "Assam": "Nikhil Kumar"}


def required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise ValueError(f"Missing required environment variable: {name}")
    return value


def optional_env(*names: str) -> str:
    for name in names:
        value = os.getenv(name, "").strip()
        if value:
            return value
    return ""


def escape_drive_query_value(value: str) -> str:
    return value.replace("\\", "\\\\").replace("'", "\\'")


def build_drive_credentials() -> Credentials:
    service_account_json = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    service_account_file = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "credential.json").strip()
    if service_account_json:
        try:
            info = json.loads(service_account_json)
        except json.JSONDecodeError as exc:
            raise ValueError("GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON") from exc
        return Credentials.from_service_account_info(info, scopes=DRIVE_SCOPES)
    if os.path.exists(service_account_file):
        return Credentials.from_service_account_file(service_account_file, scopes=DRIVE_SCOPES)
    raise ValueError("Service account credentials not found. Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_SERVICE_ACCOUNT_FILE.")


def find_file_id_by_name(drive_service, folder_id: str, filename: str) -> str:
    query = f"'{folder_id}' in parents and trashed = false and name = '{escape_drive_query_value(filename)}'"
    response = drive_service.files().list(q=query, spaces="drive", fields="files(id,name)", pageSize=10).execute()
    files = response.get("files", [])
    if not files:
        raise FileNotFoundError(f"File not found in Drive folder: {filename}")
    return files[0]["id"]


def download_drive_file_bytes(drive_service, file_id: str) -> bytes:
    request = drive_service.files().get_media(fileId=file_id)
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buffer.getvalue()


def read_tool_file(data: bytes, reader: str) -> pd.DataFrame:
    if reader == "csv":
        return pd.read_csv(io.BytesIO(data), low_memory=False, encoding="utf-8-sig")
    raise ValueError(f"Unsupported reader: {reader}")


def normalize(value) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def normalize_state(value) -> str:
    return str(value).strip().casefold()


def clean_text(series):
    return series.astype("string").str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})


def find_column(columns: Sequence[object], aliases: Sequence[str]) -> Optional[str]:
    lookup = {normalize(c): str(c) for c in columns}
    for alias in aliases:
        if normalize(alias) in lookup:
            return lookup[normalize(alias)]
    return None


def parse_submission_datetime(series):
    text = clean_text(series)
    parsed = pd.to_datetime(text, format="%d/%m/%Y, %H:%M:%S", errors="coerce")
    unresolved = parsed.isna() & text.notna()
    if unresolved.any():
        try:
            parsed.loc[unresolved] = pd.to_datetime(text.loc[unresolved], format="mixed", dayfirst=True, errors="coerce")
        except TypeError:
            parsed.loc[unresolved] = pd.to_datetime(text.loc[unresolved], dayfirst=True, errors="coerce")
    return parsed


def flatten_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    if isinstance(out.columns, pd.MultiIndex):
        out.columns = [" | ".join(str(x).strip() for x in tup if str(x).strip()) for tup in out.columns.to_flat_index()]
    else:
        out.columns = [str(c) for c in out.columns]
    return out


def standardize_frame(raw, tool_name, filename):
    detected = {key: find_column(raw.columns, aliases) for key, aliases in COLUMN_ALIASES.items()}
    if not detected["state"] or not detected["investigator"]:
        raise ValueError(f"{filename}: missing State or Investigator/QDC column. Available headers: {list(raw.columns)}")
    frame = raw.copy()
    frame["__State"] = clean_text(frame[detected["state"]])
    frame["__Investigator"] = clean_text(frame[detected["investigator"]])
    level_source = detected["facility_level"] or detected["facility_type"]
    if level_source:
        dh_set = {normalize(v) for v in DH_VALUES}
        frame["__Level"] = clean_text(frame[level_source]).map(lambda v: "DH" if normalize(v) in dh_set else "Below DH")
    else:
        frame["__Level"] = "Unclassified"
    if detected["submission_date"]:
        frame["__DateTime"] = parse_submission_datetime(frame[detected["submission_date"]])
        invalid_dates = int(frame["__DateTime"].isna().sum())
    else:
        frame["__DateTime"] = pd.NaT
        invalid_dates = "Not supplied"
    log = {"Tool": TOOL_LABELS[tool_name], "File": filename, "Rows_Read": len(raw),
           "State_Column": detected["state"], "Investigator_Column": detected["investigator"],
           "Facility_Level_Column": level_source or "", "Submission_Date_Column": detected["submission_date"] or "",
           "Invalid_Submission_Dates": invalid_dates, "Status": "OK"}
    return frame, log


def load_and_standardize_data(drive_service, source_folder_id):
    frames, logs = {}, []
    for item in TOOL_FILES:
        name, filename = item["tool_name"], item["filename"]
        logging.info("Reading %s", filename)
        try:
            file_id = find_file_id_by_name(drive_service, source_folder_id, filename)
            raw = read_tool_file(download_drive_file_bytes(drive_service, file_id), item["reader"])
            frames[name], log = standardize_frame(raw, name, filename)
            logs.append(log)
        except Exception as exc:
            logging.exception("Could not process %s", filename)
            logs.append({"Tool": TOOL_LABELS[name], "File": filename, "Rows_Read": 0, "Status": f"ERROR: {exc}"})
    if not frames:
        raise ValueError("No Google Drive CSV file could be processed.")
    return frames, pd.DataFrame(logs)


def filter_state(frame, state):
    return frame.loc[frame["__State"].fillna("").str.strip().str.casefold() == normalize_state(state)].copy()


def get_investigators(state_frames):
    return sorted({str(v).strip() for f in state_frames.values() for v in f["__Investigator"].dropna().unique()
                   if str(v).strip()}, key=str.casefold)


def count_by_investigator(frame, investigators):
    if frame is None or frame.empty:
        return [0] * len(investigators)
    grouped = frame.dropna(subset=["__Investigator"]).groupby("__Investigator").size()
    lookup = {str(name).strip().casefold(): int(count) for name, count in grouped.items()}
    return [lookup.get(name.casefold(), 0) for name in investigators]


def build_nurse_wise(state_frames, investigators):
    report = pd.DataFrame({"Name of Nursing Consultants": investigators})
    for tool in TOOL_ORDER:
        report[f"# {TOOL_LABELS[tool]}"] = count_by_investigator(state_frames.get(tool), investigators)
    total = {report.columns[0]: "Grand Total", **{c: int(report[c].sum()) for c in report.columns[1:]}}
    return pd.concat([report, pd.DataFrame([total])], ignore_index=True)


def build_dh_below_dh(state_frames, investigators):
    report = pd.DataFrame({"Name of Nurses": investigators})
    for tool in TOOL_ORDER:
        frame = state_frames.get(tool)
        for level in ("Below DH", "DH"):
            subset = None if frame is None else frame.loc[frame["__Level"] == level]
            report[f"{TOOL_LABELS[tool]} - {level}"] = count_by_investigator(subset, investigators)
    total = {"Name of Nurses": "Grand Total", **{c: int(report[c].sum()) for c in report.columns[1:]}}
    return pd.concat([report, pd.DataFrame([total])], ignore_index=True)


def fis_shift_counts(frame, investigators, day):
    if frame is None or frame.empty:
        return [0] * len(investigators)
    hours = frame["__DateTime"].dt.hour
    valid = frame["__DateTime"].notna()
    day_mask = valid & (hours >= DAY_START_HOUR) & (hours < DAY_END_HOUR)
    return count_by_investigator(frame.loc[day_mask if day else valid & ~day_mask], investigators)


def build_summary(state, state_frames, investigators, spoc_map):
    report = pd.DataFrame({"Name of Nursing Consultants": investigators, "State": [state] * len(investigators),
                           "State SPOC": [spoc_map.get(state, spoc_map.get(state.title(), ""))] * len(investigators)})
    for tool in TOOL_ORDER:
        report[TOOL_LABELS[tool]] = count_by_investigator(state_frames.get(tool), investigators)
    pos = report.columns.get_loc("FIS") + 1
    report.insert(pos, "FIS-Day (9AM-6PM)", fis_shift_counts(state_frames.get("FIS"), investigators, True))
    report.insert(pos + 1, "FIS-Night (6PM-9AM)", fis_shift_counts(state_frames.get("FIS"), investigators, False))
    return report


def source_sheet_name(filename):
    return re.sub(r"[\\/*?:\[\]]", "_", Path(filename).stem).strip()[:31] or "Raw_Data"


def raw_for_export(frame):
    return flatten_columns(frame.drop(columns=[c for c in frame.columns if str(c).startswith("__")], errors="ignore"))


def sample_widths(df, max_rows=50, cap=35):
    widths, sample = [], df.head(max_rows)
    for column in df.columns:
        maximum = len(str(column))
        lengths = sample[column].dropna().astype(str).str.len()
        if not lengths.empty:
            maximum = max(maximum, int(lengths.max()))
        widths.append(min(max(maximum + 2, 10), cap))
    return widths


def format_sheet(ws, df, header_format, total_format=None, header_row=0, freeze=(1, 0), cap=24):
    ws.freeze_panes(*freeze)
    if len(df.columns):
        ws.autofilter(header_row, 0, header_row + len(df), len(df.columns) - 1)
    ws.set_row(header_row, 38, header_format)
    for index, width in enumerate(sample_widths(df, 100, cap)):
        ws.set_column(index, index, width)
    if total_format is not None:
        ws.set_row(header_row + len(df), None, total_format)


def create_assam_workbook(frames, processing_log, output_folder, spoc_map):
    state = TARGET_STATE
    state_frames = {tool: filter_state(frame, state) for tool, frame in frames.items()}
    investigators = get_investigators(state_frames)
    if not investigators:
        raise ValueError("No Assam investigators or records were found.")
    output_folder.mkdir(parents=True, exist_ok=True)
    report_date = datetime.now().strftime("%d-%m-%Y")
    output_path = output_folder / f"ASSAM_Tool_Submission_Report_{report_date}.xlsx"
    logging.info("Starting workbook for ASSAM")
    nurse_df = flatten_columns(build_nurse_wise(state_frames, investigators))
    dh_df = flatten_columns(build_dh_below_dh(state_frames, investigators))
    summary_df = flatten_columns(build_summary(state, state_frames, investigators, spoc_map))

    with pd.ExcelWriter(output_path, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        wb = writer.book
        header = wb.add_format({"bold": True, "bg_color": "#B7DEE8", "border": 1, "align": "center",
                                "valign": "vcenter", "text_wrap": True})
        total = wb.add_format({"bold": True, "bg_color": "#B7DEE8", "border": 1})
        title = wb.add_format({"bold": True, "font_size": 14, "bottom": 1})
        nurse_df.to_excel(writer, sheet_name="Nurse Wise", index=False)
        format_sheet(writer.sheets["Nurse Wise"], nurse_df, header, total, 0, (1, 1), 32)
        dh_df.to_excel(writer, sheet_name="DH & Below DH", index=False)
        format_sheet(writer.sheets["DH & Below DH"], dh_df, header, total, 0, (1, 1), 24)
        summary_df.to_excel(writer, sheet_name="Summary", index=False, startrow=2)
        ws = writer.sheets["Summary"]
        ws.write(0, 0, f"GAVB Facility Tool Data collection status as on {report_date}", title)
        format_sheet(ws, summary_df, header, None, 2, (3, 1), 28)
        for item in TOOL_FILES:
            tool = item["tool_name"]
            if tool not in state_frames:
                continue
            raw = raw_for_export(state_frames[tool])
            sheet_name = source_sheet_name(item["filename"])
            logging.info("ASSAM: writing %s (%s rows x %s columns)", sheet_name, len(raw), len(raw.columns))
            raw.to_excel(writer, sheet_name=sheet_name, index=False)
            format_sheet(writer.sheets[sheet_name], raw, header, None, 0, (1, 0), 24)
        state_log = flatten_columns(processing_log.assign(State_Workbook=state))
        state_log.to_excel(writer, sheet_name="Processing Log", index=False)
        format_sheet(writer.sheets["Processing Log"], state_log, header, None, 0, (1, 0), 45)
    logging.info("Created Assam workbook: %s", output_path)
    return output_path


def parse_state_spoc():
    mapping = DEFAULT_STATE_SPOC.copy()
    raw = os.getenv("STATE_SPOC_JSON", "").strip()
    if raw:
        supplied = json.loads(raw)
        if not isinstance(supplied, dict):
            raise ValueError("STATE_SPOC_JSON must be a JSON object")
        mapping.update({str(k).strip(): str(v).strip() for k, v in supplied.items()})
    return mapping


def send_report_email(attachment: Path):
    sender = optional_env("GMAIL_USERNAME", "EMAIL_USERNAME")
    password = optional_env("GMAIL_APP_PASSWORD", "EMAIL_PASSWORD")
    recipients_raw = optional_env("REPORT_RECIPIENTS", "EMAIL_TO")
    if not sender or not password or not recipients_raw:
        raise ValueError("Email settings missing. Set GMAIL_USERNAME, GMAIL_APP_PASSWORD, and REPORT_RECIPIENTS.")
    recipients = [x.strip() for x in re.split(r"[,;]", recipients_raw) if x.strip()]
    if not recipients:
        raise ValueError("REPORT_RECIPIENTS does not contain a valid recipient address.")
    message = EmailMessage()
    message["From"] = optional_env("EMAIL_FROM") or sender
    message["To"] = ", ".join(recipients)
    message["Subject"] = optional_env("EMAIL_SUBJECT") or f"Assam GAVB Tool Submission Report - {datetime.now():%d-%m-%Y}"
    message.set_content(optional_env("EMAIL_BODY") or
                        "Dear Team,\n\nPlease find attached the latest Assam GAVB Tool Submission Report.\n\nRegards,\nGAVB Reporting Automation")
    with attachment.open("rb") as handle:
        message.add_attachment(handle.read(), maintype="application",
                               subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               filename=attachment.name)
    host = optional_env("SMTP_HOST") or "smtp.gmail.com"
    port = int(optional_env("SMTP_PORT") or "465")
    logging.info("Sending Assam report email to %s recipient(s) through %s:%s", len(recipients), host, port)
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(host, port, context=context, timeout=60) as smtp:
        smtp.login(sender, password)
        smtp.send_message(message)
    logging.info("Email sent successfully to: %s", ", ".join(recipients))


def main():
    logging.basicConfig(level=os.getenv("LOG_LEVEL", "INFO").upper(),
                        format="%(asctime)s [%(levelname)s] %(message)s")
    try:
        logging.info("Running report script version: %s", SCRIPT_VERSION)
        source_folder_id = required_env("DRIVE_FOLDER_ID")
        run_date = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        output_folder = Path(os.getenv("LOCAL_OUTPUT_FOLDER", f"/tmp/{run_date}/state_reports"))
        credentials = build_drive_credentials()
        drive_service = build("drive", "v3", credentials=credentials, cache_discovery=False)
        frames, processing_log = load_and_standardize_data(drive_service, source_folder_id)
        report_path = create_assam_workbook(frames, processing_log, output_folder, parse_state_spoc())
        send_report_email(report_path)
        print(f"Created and emailed Assam workbook: {report_path}")
        return 0
    except (ValueError, FileNotFoundError) as exc:
        logging.error("Configuration/runtime error: %s", exc)
        return 2
    except HttpError as exc:
        logging.error("Google API request failed: %s", exc)
        return 3
    except smtplib.SMTPAuthenticationError:
        logging.exception("Gmail authentication failed. Check GMAIL_USERNAME and GMAIL_APP_PASSWORD.")
        return 4
    except (smtplib.SMTPException, OSError) as exc:
        logging.exception("Email delivery failed: %s", exc)
        return 5
    except Exception as exc:
        logging.exception("Unexpected failure: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())

