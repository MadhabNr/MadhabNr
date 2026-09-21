#!/usr/bin/env python3
"""IDD Extension field-activity reporting for 0-5, 6-11, and Separate CD tools.

Creates one workbook per state, packages all workbooks in a ZIP, and emails the ZIP.
The outputs are operational review signals only and must not be used as employee ratings.
"""

import io
import json
import logging
import math
import os
import re
import smtplib
import ssl
import sys
import tempfile
import zipfile
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path

import numpy as np
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

SCRIPT_VERSION = "2026-09-21-idd-extension-exact-columns-v2"
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
MIN_MOVEMENT_KM = float(os.getenv("MIN_MOVEMENT_KM", "5"))

TOOL_FILES = [
    ("Mother 0-5", "IDD_Extension_Mother_0_5_WIDE.xlsx"),
    ("Mother 6-11", "IDD Extension Mother 6-11_WIDE.xlsx"),
    ("Separate CD", "IDD_Extension_Separate_CD_Tool_WIDE.xlsx"),
]

ALIASES = {
    "state": ["STATE"],
    "fi": ["QDC_Name"],
    "submission": ["SubmissionDate", "Submission Date", "submission_date", "SubmissionDateTime", "Submission_Time", "endtime", "EndTime", "starttime", "StartTime"],
    "latitude": ["Qgeo-Latitude"],
    "longitude": ["Qgeo-Longitude"],
    "awc_code": ["Cal_AWC", "QAWC_code", "QAWC_3"],
    "awc_name": ["QAWC_name"],
    "district": ["DIST"],
    "block": ["BLOCK"],
    "division": ["DIVISION"],
    "sample": ["Qsample"],
    "serial": ["Qserial"],
    "mobile": ["QDC_Mobile"],
}


def env(*names, required=False):
    for name in names:
        value = os.getenv(name, "").strip()
        if value:
            return value
    if required:
        raise ValueError("Missing required environment variable. Accepted name(s): " + ", ".join(names))
    return ""


def norm(value):
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def clean_text(series):
    return series.astype("string").str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA})


def find_column(columns, aliases):
    lookup = {norm(column): str(column) for column in columns}
    for alias in aliases:
        if norm(alias) in lookup:
            return lookup[norm(alias)]
    return None


def build_credentials():
    raw = env("GOOGLE_SERVICE_ACCOUNT_JSON")
    file_path = env("GOOGLE_SERVICE_ACCOUNT_FILE") or "credential.json"
    if raw:
        return Credentials.from_service_account_info(json.loads(raw), scopes=DRIVE_SCOPES)
    if os.path.exists(file_path):
        return Credentials.from_service_account_file(file_path, scopes=DRIVE_SCOPES)
    raise ValueError("Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_SERVICE_ACCOUNT_FILE")


def escape_drive_value(value):
    return value.replace("\\", "\\\\").replace("'", "\\'")


def find_drive_file(service, folder_id, filename):
    query = f"'{folder_id}' in parents and trashed = false and name = '{escape_drive_value(filename)}'"
    files = service.files().list(q=query, spaces="drive", fields="files(id,name,modifiedTime)", pageSize=10).execute().get("files", [])
    if not files:
        raise FileNotFoundError(f"File not found in Drive folder: {filename}")
    return files[0]


def download_bytes(service, file_id):
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, service.files().get_media(fileId=file_id))
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buffer.getvalue()


def detect_header_and_read(data, filename):
    """Find a likely header row in the first sheet, then read the complete workbook sheet."""
    preview = pd.read_excel(io.BytesIO(data), sheet_name=0, header=None, nrows=20, engine="openpyxl")
    target_tokens = {norm(x) for key in ("fi", "state", "submission") for x in ALIASES[key]}
    best_row, best_score = 0, -1
    for row_no in range(len(preview)):
        values = {norm(v) for v in preview.iloc[row_no].tolist() if pd.notna(v)}
        score = len(values & target_tokens)
        if score > best_score:
            best_row, best_score = row_no, score
    frame = pd.read_excel(io.BytesIO(data), sheet_name=0, header=best_row, engine="openpyxl")
    frame.columns = [str(c).strip() for c in frame.columns]
    logging.info("%s: detected header row %s with %s rows x %s columns", filename, best_row + 1, len(frame), len(frame.columns))
    return frame, best_row + 1


def parse_datetime(series):
    text = clean_text(series)
    parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
    return parsed


def parse_coordinate(series, is_latitude):
    values = pd.to_numeric(series, errors="coerce")
    values = values.where(values.between(-90, 90) if is_latitude else values.between(-180, 180))
    return values


def extract_geo_pair(frame, detected):
    lat_col, lon_col = detected.get("latitude"), detected.get("longitude")
    if lat_col and lon_col and lat_col != lon_col:
        return parse_coordinate(frame[lat_col], True), parse_coordinate(frame[lon_col], False)

    geo_col = lat_col or lon_col
    if geo_col:
        split = clean_text(frame[geo_col]).str.extract(r"^\s*([-+]?\d+(?:\.\d+)?)\s*[,; ]\s*([-+]?\d+(?:\.\d+)?)")
        return parse_coordinate(split[0], True), parse_coordinate(split[1], False)
    return pd.Series(np.nan, index=frame.index), pd.Series(np.nan, index=frame.index)


def standardize(raw, tool, filename, header_row):
    detected = {key: find_column(raw.columns, aliases) for key, aliases in ALIASES.items()}
    required_missing = [key for key in ("state", "fi", "submission", "latitude", "longitude") if not detected[key]]
    if required_missing:
        raise ValueError(
            f"{filename}: missing required column(s): {', '.join(required_missing)}. "
            "Expected exact core fields include STATE, QDC_Name, Qgeo-Latitude and Qgeo-Longitude, "
            "plus a submission date/time field."
        )

    frame = raw.copy()
    frame["__Tool"] = tool
    frame["__State"] = clean_text(frame[detected["state"]])
    frame["__FI"] = clean_text(frame[detected["fi"]])
    frame["__DateTime"] = parse_datetime(frame[detected["submission"]])
    frame["__Date"] = frame["__DateTime"].dt.normalize()
    frame["__Latitude"], frame["__Longitude"] = extract_geo_pair(frame, detected)
    for key in ("awc_code", "awc_name", "district", "block", "division", "sample", "serial", "mobile"):
        frame[f"__{key.title().replace('_', '')}"] = clean_text(frame[detected[key]]) if detected[key] else pd.NA

    log = {
        "Tool": tool,
        "File": filename,
        "Header_Row": header_row,
        "Rows_Read": len(raw),
        "State_Column": detected["state"],
        "FI_Column": detected["fi"],
        "Submission_Column": detected["submission"],
        "Latitude_Column": detected["latitude"] or "",
        "Longitude_Column": detected["longitude"] or "",
        "AWC_Code_Column": detected["awc_code"] or "",
        "AWC_Name_Column": detected["awc_name"] or "",
        "District_Column": detected["district"] or "",
        "Block_Column": detected["block"] or "",
        "Division_Column": detected["division"] or "",
        "Valid_Dates": int(frame["__Date"].notna().sum()),
        "Valid_GPS": int((frame["__Latitude"].notna() & frame["__Longitude"].notna()).sum()),
        "Status": "OK",
    }
    return frame, log


def load_tools(service, folder_id):
    frames, logs = {}, []
    for tool, filename in TOOL_FILES:
        logging.info("Reading %s", filename)
        try:
            file_info = find_drive_file(service, folder_id, filename)
            raw_bytes = download_bytes(service, file_info["id"])
            raw, header_row = detect_header_and_read(raw_bytes, filename)
            frames[tool], log = standardize(raw, tool, filename, header_row)
            log["Drive_Modified_Time"] = file_info.get("modifiedTime", "")
            logs.append(log)
        except Exception as exc:
            logging.exception("Could not process %s", filename)
            logs.append({"Tool": tool, "File": filename, "Rows_Read": 0, "Status": f"ERROR: {exc}"})
    if not frames:
        raise ValueError("None of the three IDD Extension files could be processed")
    return frames, pd.DataFrame(logs)


def haversine_km(lat1, lon1, lat2, lon2):
    if any(pd.isna(v) for v in (lat1, lon1, lat2, lon2)):
        return np.nan
    radius = 6371.0088
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp = math.radians(lat2 - lat1)
    dl = math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return radius * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def normalize_state(value):
    return str(value).strip().casefold()


def get_states(frames):
    groups = {}
    for frame in frames.values():
        for value in frame["__State"].dropna().unique():
            display = str(value).strip()
            if display:
                groups.setdefault(normalize_state(display), []).append(display)
    states = [max(values, key=lambda v: (values.count(v), len(v))) for values in groups.values()]
    selected = env("STATES")
    if selected:
        wanted = {normalize_state(v) for v in selected.split(",") if v.strip()}
        states = [s for s in states if normalize_state(s) in wanted]
    excluded = {normalize_state(v) for v in env("EXCLUDED_STATES").split(",") if v.strip()}
    return sorted([s for s in states if normalize_state(s) not in excluded], key=str.casefold)


def combine_state_data(frames, state):
    parts = []
    for frame in frames.values():
        mask = frame["__State"].fillna("").str.strip().str.casefold() == normalize_state(state)
        parts.append(frame.loc[mask].copy())
    return pd.concat(parts, ignore_index=True, sort=False) if parts else pd.DataFrame()


def daily_counts(data):
    valid = data.dropna(subset=["__FI", "__Date"])
    pivot = valid.pivot_table(index=["__FI", "__Date"], columns="__Tool", values="__State", aggfunc="size", fill_value=0).reset_index()
    for tool, _ in TOOL_FILES:
        if tool not in pivot.columns:
            pivot[tool] = 0
    pivot["Total Submissions"] = pivot[[tool for tool, _ in TOOL_FILES]].sum(axis=1)
    pivot = pivot.rename(columns={"__FI": "FI Name", "__Date": "Submission Date"})
    return pivot.sort_values(["FI Name", "Submission Date"])


def fi_latest_summary(daily):
    rows = []
    for fi, group in daily.groupby("FI Name", dropna=False):
        group = group.sort_values("Submission Date")
        latest = group.iloc[-1]
        previous = group.iloc[-2] if len(group) > 1 else None
        row = {
            "FI Name": fi,
            "Latest Available Date": latest["Submission Date"],
            "Latest Total Submissions": int(latest["Total Submissions"]),
            "Previous Active Date": previous["Submission Date"] if previous is not None else pd.NaT,
            "Previous Total Submissions": int(previous["Total Submissions"]) if previous is not None else 0,
            "Change vs Previous Active Date": int(latest["Total Submissions"] - (previous["Total Submissions"] if previous is not None else 0)),
        }
        for tool, _ in TOOL_FILES:
            row[f"Latest {tool}"] = int(latest[tool])
            row[f"Previous {tool}"] = int(previous[tool]) if previous is not None else 0
        rows.append(row)
    return pd.DataFrame(rows).sort_values("FI Name") if rows else pd.DataFrame()


def daily_location_table(data):
    gps = data.dropna(subset=["__FI", "__Date", "__Latitude", "__Longitude"]).copy()
    if gps.empty:
        return pd.DataFrame(columns=["FI Name", "Submission Date", "Daily Latitude", "Daily Longitude", "GPS Submissions", "Distinct AWC Codes"])
    result = gps.groupby(["__FI", "__Date"], as_index=False).agg(
        **{
            "Daily Latitude": ("__Latitude", "median"),
            "Daily Longitude": ("__Longitude", "median"),
            "GPS Submissions": ("__Latitude", "size"),
            "Distinct AWC Codes": ("__AwcCode", lambda s: int(s.dropna().nunique())),
        }
    ).rename(columns={"__FI": "FI Name", "__Date": "Submission Date"})
    result = result.sort_values(["FI Name", "Submission Date"])
    result["Previous Active Date"] = result.groupby("FI Name")["Submission Date"].shift(1)
    result["Previous Latitude"] = result.groupby("FI Name")["Daily Latitude"].shift(1)
    result["Previous Longitude"] = result.groupby("FI Name")["Daily Longitude"].shift(1)
    result["Distance from Previous Active Day (km)"] = result.apply(
        lambda r: haversine_km(r["Previous Latitude"], r["Previous Longitude"], r["Daily Latitude"], r["Daily Longitude"]), axis=1
    )
    result["Movement Signal"] = np.select(
        [result["Distance from Previous Active Day (km)"].isna(), result["Distance from Previous Active Day (km)"] >= MIN_MOVEMENT_KM],
        ["No prior valid GPS day", f">= {MIN_MOVEMENT_KM:g} km"],
        default=f"< {MIN_MOVEMENT_KM:g} km - review context",
    )
    return result


def weekly_movement(location_daily):
    if location_daily.empty:
        return pd.DataFrame()
    data = location_daily.copy()
    iso = data["Submission Date"].dt.isocalendar()
    data["ISO Year"] = iso.year.astype(int)
    data["ISO Week"] = iso.week.astype(int)
    data["Week Start"] = data["Submission Date"] - pd.to_timedelta(data["Submission Date"].dt.weekday, unit="D")
    valid = data.dropna(subset=["Distance from Previous Active Day (km)"])
    if valid.empty:
        return pd.DataFrame()
    result = valid.groupby(["FI Name", "ISO Year", "ISO Week", "Week Start"], as_index=False).agg(
        Active_Days=("Submission Date", "nunique"),
        GPS_Days=("GPS Submissions", "size"),
        Average_Distance_km=("Distance from Previous Active Day (km)", "mean"),
        Median_Distance_km=("Distance from Previous Active Day (km)", "median"),
        Maximum_Distance_km=("Distance from Previous Active Day (km)", "max"),
        Days_At_Least_5_km=("Distance from Previous Active Day (km)", lambda s: int((s >= MIN_MOVEMENT_KM).sum())),
        Comparisons=("Distance from Previous Active Day (km)", "count"),
    )
    result["Percent Comparisons >= 5 km"] = np.where(result["Comparisons"] > 0, result["Days_At_Least_5_km"] / result["Comparisons"], np.nan)
    result["Weekly Review Signal"] = np.select(
        [result["Comparisons"] < 2, result["Percent Comparisons >= 5 km"] >= 0.5],
        ["Insufficient comparisons", "Movement pattern present"],
        default="Review route/AWC allocation",
    )
    return result.sort_values(["FI Name", "Week Start"])


def awc_visit_summary(data):
    valid = data.dropna(subset=["__FI", "__Date"]).copy()
    valid["AWC Identifier"] = clean_text(valid["__AwcCode"]).fillna(clean_text(valid["__AwcName"]))
    result = valid.groupby("__FI", as_index=False).agg(
        First_Submission_Date=("__Date", "min"),
        Latest_Submission_Date=("__Date", "max"),
        Active_Days=("__Date", "nunique"),
        Total_Submissions=("__Tool", "size"),
        Distinct_AWC_Identifiers=("AWC Identifier", lambda s: int(s.dropna().nunique())),
        Valid_GPS_Submissions=("__Latitude", lambda s: int(s.notna().sum())),
    ).rename(columns={"__FI": "FI Name"})
    result["GPS Coverage %"] = np.where(result["Total_Submissions"] > 0, result["Valid_GPS_Submissions"] / result["Total_Submissions"], np.nan)
    return result.sort_values("FI Name")


def duplicate_gps_review(data):
    valid = data.dropna(subset=["__Latitude", "__Longitude", "__FI", "__Date"]).copy()
    if valid.empty:
        return pd.DataFrame()
    valid["Rounded Latitude"] = valid["__Latitude"].round(5)
    valid["Rounded Longitude"] = valid["__Longitude"].round(5)
    grouped = valid.groupby(["__FI", "Rounded Latitude", "Rounded Longitude"], as_index=False).agg(
        Submission_Count=("__Tool", "size"),
        Active_Days=("__Date", "nunique"),
        First_Date=("__Date", "min"),
        Latest_Date=("__Date", "max"),
        Distinct_AWC_Codes=("__AwcCode", lambda s: int(s.dropna().nunique())),
    )
    grouped = grouped[(grouped["Submission_Count"] >= 3) | (grouped["Active_Days"] >= 2)]
    return grouped.rename(columns={"__FI": "FI Name"}).sort_values(["Submission_Count", "Active_Days"], ascending=False)


def raw_for_export(data):
    internal = [c for c in data.columns if str(c).startswith("__")]
    return data.drop(columns=internal, errors="ignore")


def safe_name(value):
    return re.sub(r'[<>:"/\\|?*]+', "_", str(value)).strip() or "Unknown_State"


def format_sheet(writer, sheet, df, header_row=0, percent_columns=None, date_columns=None):
    ws = writer.sheets[sheet]
    wb = writer.book
    header = wb.add_format({"bold": True, "bg_color": "#1F4E78", "font_color": "white", "text_wrap": True, "align": "center", "valign": "vcenter"})
    percent_fmt = wb.add_format({"num_format": "0.0%"})
    date_fmt = wb.add_format({"num_format": "dd-mm-yyyy"})
    warning_fmt = wb.add_format({"bg_color": "#FCE4D6"})
    ws.freeze_panes(header_row + 1, 1 if len(df.columns) > 1 else 0)
    ws.set_row(header_row, 36, header)
    if len(df.columns):
        ws.autofilter(header_row, 0, header_row + len(df), len(df.columns) - 1)
    sample = df.head(100)
    for i, col in enumerate(df.columns):
        lengths = sample[col].dropna().astype(str).str.len()
        width = min(max(len(str(col)) + 2, int(lengths.max()) + 2 if not lengths.empty else 10), 35)
        cell_format = percent_fmt if percent_columns and col in percent_columns else date_fmt if date_columns and col in date_columns else None
        ws.set_column(i, i, width, cell_format)
    for phrase in ("review", "insufficient"):
        ws.conditional_format(header_row + 1, 0, header_row + len(df), max(len(df.columns) - 1, 0), {
            "type": "text", "criteria": "containing", "value": phrase, "format": warning_fmt
        })


def create_state_workbook(state, data, processing_log, output_dir):
    output_dir.mkdir(parents=True, exist_ok=True)
    report_date = datetime.now().strftime("%d-%m-%Y")
    path = output_dir / f"{safe_name(state).upper()}_IDD_Extension_Field_Activity_{report_date}.xlsx"

    daily = daily_counts(data)
    latest = fi_latest_summary(daily)
    locations = daily_location_table(data)
    weekly = weekly_movement(locations)
    awc = awc_visit_summary(data)
    repeated = duplicate_gps_review(data)

    note = pd.DataFrame({
        "Important Use Note": [
            "This workbook contains operational data-quality and field-coverage signals, not employee performance ratings.",
            "A distance below 5 km does not prove non-travel or incorrect work. Nearby AWCs, repeat household visits, GPS error, offline submission, and route allocation can all explain the pattern.",
            "Use review signals for supportive verification with route plans, AWC allocation lists, supervisor notes, and consented GPS-monitoring policy.",
            "Today's date is defined separately for each FI as that FI's latest available valid submission date. Previous day means the FI's immediately previous active submission date.",
            "Daily movement uses the median valid Qgeo-Latitude and Qgeo-Longitude coordinate for an FI-day and compares it with the previous active FI-day.",
            "AWC identity uses Cal_AWC first when available, then QAWC_code, then QAWC_3; QAWC_name is used when no code is available.",
        ]
    })

    with pd.ExcelWriter(path, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        note.to_excel(writer, sheet_name="Read Me", index=False)
        format_sheet(writer, "Read Me", note)

        latest.to_excel(writer, sheet_name="FI Latest vs Previous", index=False)
        format_sheet(writer, "FI Latest vs Previous", latest, date_columns={"Latest Available Date", "Previous Active Date"})

        daily.to_excel(writer, sheet_name="FI Daily Submissions", index=False)
        format_sheet(writer, "FI Daily Submissions", daily, date_columns={"Submission Date"})

        locations.to_excel(writer, sheet_name="Daily GPS Movement", index=False)
        format_sheet(writer, "Daily GPS Movement", locations, date_columns={"Submission Date", "Previous Active Date"})

        weekly.to_excel(writer, sheet_name="Weekly Movement", index=False)
        format_sheet(writer, "Weekly Movement", weekly, percent_columns={"Percent Comparisons >= 5 km"}, date_columns={"Week Start"})

        awc.to_excel(writer, sheet_name="FI Coverage Summary", index=False)
        format_sheet(writer, "FI Coverage Summary", awc, percent_columns={"GPS Coverage %"}, date_columns={"First_Submission_Date", "Latest_Submission_Date"})

        repeated.to_excel(writer, sheet_name="Repeated GPS Review", index=False)
        format_sheet(writer, "Repeated GPS Review", repeated, date_columns={"First_Date", "Latest_Date"})

        processing_log.assign(State_Workbook=state).to_excel(writer, sheet_name="Processing Log", index=False)
        format_sheet(writer, "Processing Log", processing_log.assign(State_Workbook=state))

        for tool, _ in TOOL_FILES:
            tool_data = data.loc[data["__Tool"] == tool]
            raw = raw_for_export(tool_data)
            sheet = re.sub(r"[\\/*?:\[\]]", "_", f"Raw {tool}")[:31]
            raw.to_excel(writer, sheet_name=sheet, index=False)
            format_sheet(writer, sheet, raw)

    logging.info("Created state workbook: %s", path)
    return path


def create_zip(paths, output_dir):
    zip_path = output_dir / f"IDD_Extension_Field_Activity_All_States_{datetime.now():%d-%m-%Y}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for path in paths:
            archive.write(path, arcname=path.name)
    return zip_path


def send_email(attachment, states):
    sender = env("SMTP_EMAIL", "GMAIL_USERNAME", "EMAIL_USERNAME", required=True)
    password = env("SMTP_APP_PASSWORD", "GMAIL_APP_PASSWORD", "EMAIL_PASSWORD", required=True)
    recipients_raw = env("RECIPIENTS", "REPORT_RECIPIENTS", "EMAIL_TO", required=True)
    recipients = [x.strip() for x in re.split(r"[,;]", recipients_raw) if x.strip()]
    if not recipients:
        raise ValueError("RECIPIENTS is empty")

    size_mb = attachment.stat().st_size / (1024 * 1024)
    maximum = float(env("MAX_EMAIL_ATTACHMENT_MB") or "24")
    if size_mb > maximum:
        raise ValueError(f"ZIP attachment is {size_mb:.2f} MB, above configured limit {maximum:.2f} MB")

    message = EmailMessage()
    message["From"] = sender
    message["To"] = ", ".join(recipients)
    message["Subject"] = env("MAIL_SUBJECT", "EMAIL_SUBJECT") or f"IDD Extension Field Activity Report - {datetime.now():%d-%m-%Y}"
    message.set_content(env("MAIL_BODY", "EMAIL_BODY") or (
        "Dear Team,\n\nPlease find attached the IDD Extension field-activity reports. "
        "The workbook provides submission, GPS coverage, and travel-review signals for operational follow-up only. "
        "It should not be used as a standalone employee performance assessment.\n\n"
        f"States included: {', '.join(states)}\n\nRegards,\nGAVB Reporting Automation"
    ))
    with attachment.open("rb") as handle:
        message.add_attachment(handle.read(), maintype="application", subtype="zip", filename=attachment.name)

    with smtplib.SMTP_SSL(env("SMTP_HOST") or "smtp.gmail.com", int(env("SMTP_PORT") or "465"), context=ssl.create_default_context(), timeout=60) as smtp:
        smtp.login(sender, password)
        smtp.send_message(message)
    logging.info("Email sent successfully to %s recipient(s)", len(recipients))


def main():
    logging.basicConfig(level=(env("LOG_LEVEL") or "INFO").upper(), format="%(asctime)s [%(levelname)s] %(message)s")
    try:
        logging.info("Running report script version: %s", SCRIPT_VERSION)
        folder_id = env("DRIVE_FOLDER_ID", required=True)
        output_dir = Path(env("LOCAL_OUTPUT_FOLDER") or f"/tmp/{datetime.now(timezone.utc):%Y-%m-%d}/idd_extension_reports")
        service = build("drive", "v3", credentials=build_credentials(), cache_discovery=False)
        frames, process_log = load_tools(service, folder_id)
        states = get_states(frames)
        if not states:
            raise ValueError("No reportable state values were found")

        reports, completed_states = [], []
        for state in states:
            data = combine_state_data(frames, state)
            if data.empty:
                continue
            reports.append(create_state_workbook(state, data, process_log, output_dir))
            completed_states.append(state)

        if not reports:
            raise ValueError("No state workbook was created")
        zip_path = create_zip(reports, output_dir)
        send_email(zip_path, completed_states)
        print("Created reports:")
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
