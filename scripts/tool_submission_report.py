#!/usr/bin/env python3
"""Create GAVB reports for every detected state and email them in one ZIP file."""

import io
import json
import logging
import os
import re
import smtplib
import ssl
import sys
import zipfile
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path
from typing import Optional, Sequence

import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

SCRIPT_VERSION = "2026-09-21-all-states-smtp-v3"
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

TOOLS = [
    ("DOD", "DOD_2026_WIDE.csv"),
    ("Maternal_Log", "Maternal Log_WIDE.csv"),
    ("FIS", "FIS CASE RECORDS 2026_WIDE.csv"),
    ("PMSMA", "PMSMA Client Interview_WIDE.csv"),
    ("SNCU", "SNCU Index Cases Observation 2026_WIDE.csv"),
    ("Referral_Services", "GABV IDD Referral_WIDE.csv"),
    ("Digital_System", "GAVB IDD Digital System_WIDE.csv"),
    ("Exit_Interview", "GAVB IDD Exit Interview_WIDE.csv"),
    ("HR", "GAVB IDD HR_WIDE.csv"),
    ("Labour_Room_Readiness", "GAVB IDD Labour Room Readiness_WIDE.csv"),
    ("Supply_Chain", "GAVB IDD Supply Chain_WIDE.csv"),
]

LABELS = {
    "DOD": "DOD",
    "Maternal_Log": "Maternal Log",
    "FIS": "FIS",
    "PMSMA": "PMSMA",
    "SNCU": "SNCU",
    "Referral_Services": "Referral Services",
    "Digital_System": "Digital System",
    "Exit_Interview": "Exit Interview",
    "HR": "HR",
    "Labour_Room_Readiness": "Labour Room Readiness",
    "Supply_Chain": "Supply Chain",
}

ALIASES = {
    "state": ["STATE", "State", "state", "Cal_STATE", "state_name"],
    "investigator": [
        "QDC", "Investigator", "Nurse", "Nurse_Name", "Nursing_Consultant",
        "Name of Nursing Consultants", "Name of Nurses", "collector_name",
    ],
    "facility_type": ["F_Type", "Facility_Type", "Facility Type", "facilitytype"],
    "facility_level": ["Facility_Level", "Facility Level", "Level", "DH_Below_DH"],
    "submission_date": [
        "SubmissionDate", "Submission Date", "submission_date", "SubmissionDateTime",
        "Submission_Time", "starttime", "endtime",
    ],
}

DH_VALUES = {"dh", "district hospital", "district_hospital", "district hospital dh"}
DAY_START_HOUR = 9
DAY_END_HOUR = 18
DEFAULT_STATE_SPOC = {"Assam": "Nikhil Kumar"}


def env(*names: str, required: bool = False) -> str:
    for name in names:
        value = os.getenv(name, "").strip()
        if value:
            return value
    if required:
        raise ValueError("Missing required environment variable. Accepted name(s): " + ", ".join(names))
    return ""


def normalize(value) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def normalize_state(value) -> str:
    return str(value).strip().casefold()


def clean(series):
    return series.astype("string").str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})


def find_column(columns: Sequence[object], names: Sequence[str]) -> Optional[str]:
    lookup = {normalize(column): str(column) for column in columns}
    return next((lookup[normalize(name)] for name in names if normalize(name) in lookup), None)


def flatten(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    if isinstance(result.columns, pd.MultiIndex):
        result.columns = [
            " | ".join(str(value).strip() for value in values if str(value).strip())
            for values in result.columns.to_flat_index()
        ]
    else:
        result.columns = [str(column) for column in result.columns]
    return result


def build_credentials() -> Credentials:
    raw = env("GOOGLE_SERVICE_ACCOUNT_JSON")
    file_path = env("GOOGLE_SERVICE_ACCOUNT_FILE") or "credential.json"
    if raw:
        try:
            info = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise ValueError("GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON") from exc
        return Credentials.from_service_account_info(info, scopes=DRIVE_SCOPES)
    if os.path.exists(file_path):
        return Credentials.from_service_account_file(file_path, scopes=DRIVE_SCOPES)
    raise ValueError("Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_SERVICE_ACCOUNT_FILE")


def drive_file_id(service, folder_id: str, filename: str) -> str:
    safe_name = filename.replace("\\", "\\\\").replace("'", "\\'")
    query = f"'{folder_id}' in parents and trashed = false and name = '{safe_name}'"
    files = (
        service.files()
        .list(q=query, spaces="drive", fields="files(id,name)", pageSize=10)
        .execute()
        .get("files", [])
    )
    if not files:
        raise FileNotFoundError(f"File not found: {filename}")
    return files[0]["id"]


def download(service, file_id: str) -> bytes:
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, service.files().get_media(fileId=file_id))
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buffer.getvalue()


def parse_datetime(series):
    text = clean(series)
    parsed = pd.to_datetime(text, format="%d/%m/%Y, %H:%M:%S", errors="coerce")
    unresolved = parsed.isna() & text.notna()
    if unresolved.any():
        try:
            parsed.loc[unresolved] = pd.to_datetime(
                text.loc[unresolved], format="mixed", dayfirst=True, errors="coerce"
            )
        except TypeError:
            parsed.loc[unresolved] = pd.to_datetime(
                text.loc[unresolved], dayfirst=True, errors="coerce"
            )
    return parsed


def standardize(raw: pd.DataFrame, tool: str, filename: str):
    detected = {key: find_column(raw.columns, aliases) for key, aliases in ALIASES.items()}
    if not detected["state"] or not detected["investigator"]:
        raise ValueError(f"{filename}: missing State or Investigator/QDC column")

    frame = raw.copy()
    frame["__State"] = clean(frame[detected["state"]])
    frame["__Investigator"] = clean(frame[detected["investigator"]])

    level_source = detected["facility_level"] or detected["facility_type"]
    if level_source:
        dh_values = {normalize(value) for value in DH_VALUES}
        frame["__Level"] = clean(frame[level_source]).map(
            lambda value: "DH" if normalize(value) in dh_values else "Below DH"
        )
    else:
        frame["__Level"] = "Unclassified"

    if detected["submission_date"]:
        frame["__DateTime"] = parse_datetime(frame[detected["submission_date"]])
        invalid_dates = int(frame["__DateTime"].isna().sum())
    else:
        frame["__DateTime"] = pd.NaT
        invalid_dates = "Not supplied"

    log = {
        "Tool": LABELS[tool],
        "File": filename,
        "Rows_Read": len(raw),
        "State_Column": detected["state"],
        "Investigator_Column": detected["investigator"],
        "Facility_Level_Column": level_source or "",
        "Submission_Date_Column": detected["submission_date"] or "",
        "Invalid_Submission_Dates": invalid_dates,
        "Status": "OK",
    }
    return frame, log


def load_all(service, folder_id: str):
    frames = {}
    logs = []
    for tool, filename in TOOLS:
        logging.info("Reading %s", filename)
        try:
            file_bytes = download(service, drive_file_id(service, folder_id, filename))
            raw = pd.read_csv(io.BytesIO(file_bytes), low_memory=False, encoding="utf-8-sig")
            frames[tool], log = standardize(raw, tool, filename)
            logs.append(log)
        except Exception as exc:
            logging.exception("Could not process %s", filename)
            logs.append({
                "Tool": LABELS[tool],
                "File": filename,
                "Rows_Read": 0,
                "Status": f"ERROR: {exc}",
            })
    if not frames:
        raise ValueError("No Google Drive CSV file could be processed")
    return frames, pd.DataFrame(logs)


def canonical_state_names(frames) -> list[str]:
    groups = {}
    for frame in frames.values():
        for value in frame["__State"].dropna().unique():
            display = str(value).strip()
            if display:
                groups.setdefault(normalize_state(display), []).append(display)
    # This merges values such as ASSAM and Assam into one state workbook.
    states = []
    for values in groups.values():
        states.append(max(values, key=lambda value: (sum(v == value for v in values), len(value))))
    return sorted(states, key=str.casefold)


def selected_states(frames) -> list[str]:
    states = canonical_state_names(frames)
    requested = env("STATES")
    if requested:
        wanted = {normalize_state(value) for value in requested.split(",") if value.strip()}
        states = [state for state in states if normalize_state(state) in wanted]
    excluded = {normalize_state(value) for value in env("EXCLUDED_STATES").split(",") if value.strip()}
    if excluded:
        states = [state for state in states if normalize_state(state) not in excluded]
    return states


def filter_state(frame: pd.DataFrame, state: str) -> pd.DataFrame:
    mask = frame["__State"].fillna("").str.strip().str.casefold() == normalize_state(state)
    return frame.loc[mask].copy()


def investigators(frames) -> list[str]:
    return sorted(
        {
            str(value).strip()
            for frame in frames.values()
            for value in frame["__Investigator"].dropna().unique()
            if str(value).strip()
        },
        key=str.casefold,
    )


def counts(frame, names):
    if frame is None or frame.empty:
        return [0] * len(names)
    grouped = frame.dropna(subset=["__Investigator"]).groupby("__Investigator").size()
    lookup = {str(name).strip().casefold(): int(count) for name, count in grouped.items()}
    return [lookup.get(name.casefold(), 0) for name in names]


def nurse_report(frames, names):
    report = pd.DataFrame({"Name of Nursing Consultants": names})
    for tool, _ in TOOLS:
        report[f"# {LABELS[tool]}"] = counts(frames.get(tool), names)
    total = {report.columns[0]: "Grand Total", **{c: int(report[c].sum()) for c in report.columns[1:]}}
    return pd.concat([report, pd.DataFrame([total])], ignore_index=True)


def dh_report(frames, names):
    report = pd.DataFrame({"Name of Nurses": names})
    for tool, _ in TOOLS:
        frame = frames.get(tool)
        for level in ("Below DH", "DH"):
            subset = None if frame is None else frame.loc[frame["__Level"] == level]
            report[f"{LABELS[tool]} - {level}"] = counts(subset, names)
    total = {"Name of Nurses": "Grand Total", **{c: int(report[c].sum()) for c in report.columns[1:]}}
    return pd.concat([report, pd.DataFrame([total])], ignore_index=True)


def shift_counts(frame, names, day: bool):
    if frame is None or frame.empty:
        return [0] * len(names)
    valid = frame["__DateTime"].notna()
    hour = frame["__DateTime"].dt.hour
    day_mask = valid & (hour >= DAY_START_HOUR) & (hour < DAY_END_HOUR)
    return counts(frame.loc[day_mask if day else valid & ~day_mask], names)


def parse_spoc_mapping() -> dict:
    mapping = DEFAULT_STATE_SPOC.copy()
    raw = env("STATE_SPOC_JSON")
    if raw:
        supplied = json.loads(raw)
        if not isinstance(supplied, dict):
            raise ValueError("STATE_SPOC_JSON must be a JSON object")
        mapping.update({str(key).strip(): str(value).strip() for key, value in supplied.items()})
    return {normalize_state(key): value for key, value in mapping.items()}


def summary_report(state, frames, names, spoc_mapping):
    report = pd.DataFrame({
        "Name of Nursing Consultants": names,
        "State": [state] * len(names),
        "State SPOC": [spoc_mapping.get(normalize_state(state), "")] * len(names),
    })
    for tool, _ in TOOLS:
        report[LABELS[tool]] = counts(frames.get(tool), names)
    position = report.columns.get_loc("FIS") + 1
    report.insert(position, "FIS-Day (9AM-6PM)", shift_counts(frames.get("FIS"), names, True))
    report.insert(position + 1, "FIS-Night (6PM-9AM)", shift_counts(frames.get("FIS"), names, False))
    return report


def safe_filename(value: str) -> str:
    return re.sub(r'[<>:"/\\|?*]+', "_", value).strip() or "Unknown_State"


def sheet_name(filename: str) -> str:
    return re.sub(r"[\\/*?:\[\]]", "_", Path(filename).stem).strip()[:31] or "Raw_Data"


def sample_widths(df, cap):
    sample = df.head(100)
    result = []
    for column in df.columns:
        lengths = sample[column].dropna().astype(str).str.len()
        maximum = max(len(str(column)), int(lengths.max()) if not lengths.empty else 0)
        result.append(min(max(maximum + 2, 10), cap))
    return result


def format_sheet(ws, df, header, total=None, row=0, freeze=(1, 0), cap=24):
    ws.freeze_panes(*freeze)
    if len(df.columns):
        ws.autofilter(row, 0, row + len(df), len(df.columns) - 1)
    ws.set_row(row, 38, header)
    for index, width in enumerate(sample_widths(df, cap)):
        ws.set_column(index, index, width)
    if total is not None:
        ws.set_row(row + len(df), None, total)


def create_state_report(state, frames, process_log, output_dir, spoc_mapping):
    state_frames = {tool: filter_state(frame, state) for tool, frame in frames.items()}
    names = investigators(state_frames)
    if not names:
        logging.warning("Skipping %s because no investigators were found", state)
        return None

    output_dir.mkdir(parents=True, exist_ok=True)
    report_date = datetime.now().strftime("%d-%m-%Y")
    path = output_dir / f"{safe_filename(state).upper()}_Tool_Submission_Report_{report_date}.xlsx"
    logging.info("Starting workbook for %s", state)

    nurse_df = flatten(nurse_report(state_frames, names))
    dh_df = flatten(dh_report(state_frames, names))
    summary_df = flatten(summary_report(state, state_frames, names, spoc_mapping))

    with pd.ExcelWriter(path, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        workbook = writer.book
        header = workbook.add_format({
            "bold": True, "bg_color": "#B7DEE8", "border": 1,
            "align": "center", "valign": "vcenter", "text_wrap": True,
        })
        total = workbook.add_format({"bold": True, "bg_color": "#B7DEE8", "border": 1})
        title = workbook.add_format({"bold": True, "font_size": 14, "bottom": 1})

        nurse_df.to_excel(writer, sheet_name="Nurse Wise", index=False)
        format_sheet(writer.sheets["Nurse Wise"], nurse_df, header, total, cap=32)

        dh_df.to_excel(writer, sheet_name="DH & Below DH", index=False)
        format_sheet(writer.sheets["DH & Below DH"], dh_df, header, total, cap=24)

        summary_df.to_excel(writer, sheet_name="Summary", index=False, startrow=2)
        summary_ws = writer.sheets["Summary"]
        summary_ws.write(0, 0, f"GAVB Facility Tool Data collection status as on {report_date}", title)
        format_sheet(summary_ws, summary_df, header, row=2, freeze=(3, 1), cap=28)

        for tool, filename in TOOLS:
            if tool not in state_frames:
                continue
            raw = flatten(
                state_frames[tool].drop(
                    columns=[c for c in state_frames[tool].columns if str(c).startswith("__")],
                    errors="ignore",
                )
            )
            raw_sheet = sheet_name(filename)
            logging.info("%s: writing %s (%s rows x %s columns)", state, raw_sheet, len(raw), len(raw.columns))
            raw.to_excel(writer, sheet_name=raw_sheet, index=False)
            format_sheet(writer.sheets[raw_sheet], raw, header, cap=24)

        state_log = flatten(process_log.assign(State_Workbook=state))
        state_log.to_excel(writer, sheet_name="Processing Log", index=False)
        format_sheet(writer.sheets["Processing Log"], state_log, header, cap=45)

    logging.info("Created state workbook: %s", path)
    return path


def create_zip(report_paths, output_dir):
    date = datetime.now().strftime("%d-%m-%Y")
    zip_path = output_dir / f"ALL_STATES_GAVB_Reports_{date}.zip"
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for report_path in report_paths:
            archive.write(report_path, arcname=report_path.name)
    logging.info("Created ZIP package: %s", zip_path)
    return zip_path


def send_email(attachment: Path, states: list[str]):
    sender = env("SMTP_EMAIL", "GMAIL_USERNAME", "EMAIL_USERNAME", required=True)
    password = env("SMTP_APP_PASSWORD", "GMAIL_APP_PASSWORD", "EMAIL_PASSWORD", required=True)
    raw_recipients = env("RECIPIENTS", "REPORT_RECIPIENTS", "EMAIL_TO", required=True)
    recipients = [value.strip() for value in re.split(r"[,;]", raw_recipients) if value.strip()]
    if not recipients:
        raise ValueError("RECIPIENTS is empty")

    message = EmailMessage()
    message["From"] = sender
    message["To"] = ", ".join(recipients)
    message["Subject"] = env("MAIL_SUBJECT", "EMAIL_SUBJECT") or (
        f"All States GAVB Tool Submission Reports - {datetime.now():%d-%m-%Y}"
    )
    state_list = ", ".join(states)
    message.set_content(
        env("MAIL_BODY", "EMAIL_BODY")
        or (
            "Dear Team,\n\n"
            "Please find attached the latest GAVB Tool Submission Reports for all detected states.\n\n"
            f"States included: {state_list}\n\n"
            "Regards,\nGAVB Reporting Automation"
        )
    )

    with attachment.open("rb") as handle:
        message.add_attachment(
            handle.read(),
            maintype="application",
            subtype="zip",
            filename=attachment.name,
        )

    max_attachment_mb = float(env("MAX_EMAIL_ATTACHMENT_MB") or "24")
    attachment_mb = attachment.stat().st_size / (1024 * 1024)
    if attachment_mb > max_attachment_mb:
        raise ValueError(
            f"ZIP attachment is {attachment_mb:.2f} MB, above MAX_EMAIL_ATTACHMENT_MB={max_attachment_mb:.2f}. "
            "Reduce raw sheets, split the reports, or increase the configured limit only if your mail account supports it."
        )

    logging.info(
        "Email configuration detected: SMTP_EMAIL=%s, SMTP_APP_PASSWORD=%s, RECIPIENTS=%s, MAIL_SUBJECT=%s",
        bool(sender), bool(password), bool(recipients), bool(message["Subject"]),
    )
    logging.info("Sending %s reports in ZIP attachment %.2f MB", len(states), attachment_mb)

    smtp_host = env("SMTP_HOST") or "smtp.gmail.com"
    smtp_port = int(env("SMTP_PORT") or "465")
    with smtplib.SMTP_SSL(
        smtp_host,
        smtp_port,
        context=ssl.create_default_context(),
        timeout=60,
    ) as smtp:
        smtp.login(sender, password)
        smtp.send_message(message)
    logging.info("Email sent successfully to %s recipient(s)", len(recipients))


def main():
    logging.basicConfig(
        level=(env("LOG_LEVEL") or "INFO").upper(),
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
    try:
        logging.info("Running report script version: %s", SCRIPT_VERSION)
        folder_id = env("DRIVE_FOLDER_ID", required=True)
        output_dir = Path(
            env("LOCAL_OUTPUT_FOLDER")
            or f"/tmp/{datetime.now(timezone.utc):%Y-%m-%d}/state_reports"
        )

        service = build("drive", "v3", credentials=build_credentials(), cache_discovery=False)
        frames, process_log = load_all(service, folder_id)
        states = selected_states(frames)
        if not states:
            raise ValueError("No reportable states were found")

        logging.info("States selected for report generation: %s", ", ".join(states))
        spoc_mapping = parse_spoc_mapping()
        reports = []
        completed_states = []
        for state in states:
            report = create_state_report(state, frames, process_log, output_dir, spoc_mapping)
            if report is not None:
                reports.append(report)
                completed_states.append(state)

        if not reports:
            raise ValueError("No state workbook was created")

        zip_path = create_zip(reports, output_dir)
        send_email(zip_path, completed_states)

        print("Created state workbooks:")
        for report in reports:
            print(f" - {report}")
        print(f"Email attachment: {zip_path}")
        return 0

    except (ValueError, FileNotFoundError, json.JSONDecodeError) as exc:
        logging.error("Configuration/runtime error: %s", exc)
        return 2
    except HttpError as exc:
        logging.error("Google API request failed: %s", exc)
        return 3
    except smtplib.SMTPAuthenticationError:
        logging.exception("Gmail authentication failed. Check SMTP_EMAIL and SMTP_APP_PASSWORD")
        return 4
    except (smtplib.SMTPException, OSError) as exc:
        logging.exception("Email delivery failed: %s", exc)
        return 5
    except Exception as exc:
        logging.exception("Unexpected failure: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())

