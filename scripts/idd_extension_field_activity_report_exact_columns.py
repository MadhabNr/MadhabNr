#!/usr/bin/env python3
"""Simple cumulative IDD Extension summary by FI for all states."""

import io
import json
import logging
import math
import os
import re
import smtplib
import ssl
import sys
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

SCRIPT_VERSION = "2026-09-25-idd-simple-summary-v9"
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
TRAVEL_THRESHOLD_KM = float(os.getenv("WEEKLY_TRAVEL_KM", "10"))

TOOL_FILES = [
    ("Mother 0-5", "IDD_Extension_Mother_0_5_WIDE.xlsx", "Qgeo-Latitude", "Qgeo-Longitude"),
    ("Mother 6-11", "IDD Extension Mother 6-11_WIDE.xlsx", "QX1-Latitude", "QX1-Longitude"),
    ("Separate CD", "IDD_Extension_Separate_CD_Tool_WIDE.xlsx", "Qgeo-Latitude", "Qgeo-Longitude"),
]
TOOL_NAMES = [tool for tool, _, _, _ in TOOL_FILES]

ALIASES = {
    "state": ["STATE", "State", "state", "Cal_STATE", "state_name"],
    "fi": ["QDC_Name", "QDC Name", "Field Investigator Name", "Investigator", "QDC", "collector_name"],
    "date": [
        "SubmissionDate", "Submission Date", "submission_date", "SubmissionDateTime",
        "Submission_Time", "endtime", "EndTime", "starttime", "StartTime",
    ],
    "cal_awc": ["Cal_AWC"],
    "current_awc": ["QAWC_3"],
    "given_awc": ["QAWC_code"],
    "awc_name": ["QAWC_name"],
}


def env(*names, required=False):
    for name in names:
        value = os.getenv(name, "").strip()
        if value:
            return value
    if required:
        raise ValueError("Missing environment variable: " + " or ".join(names))
    return ""


def normalize(value):
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def clean(series):
    return series.astype("string").str.strip().replace(
        {"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA}
    )


def find_column(columns, aliases):
    lookup = {normalize(column): str(column) for column in columns}
    for alias in aliases:
        if normalize(alias) in lookup:
            return lookup[normalize(alias)]
    return None


def credentials():
    raw = env("GOOGLE_SERVICE_ACCOUNT_JSON")
    path = env("GOOGLE_SERVICE_ACCOUNT_FILE") or "credential.json"
    if raw:
        return Credentials.from_service_account_info(json.loads(raw), scopes=DRIVE_SCOPES)
    if os.path.exists(path):
        return Credentials.from_service_account_file(path, scopes=DRIVE_SCOPES)
    raise ValueError("Google service account credentials not found")


def find_file_id(service, folder_id, filename):
    safe = filename.replace("\\", "\\\\").replace("'", "\\'")
    query = f"'{folder_id}' in parents and trashed = false and name = '{safe}'"
    files = service.files().list(
        q=query, spaces="drive", fields="files(id,name)", pageSize=10
    ).execute().get("files", [])
    if not files:
        raise FileNotFoundError(f"File not found: {filename}")
    return files[0]["id"]


def download(service, file_id):
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, service.files().get_media(fileId=file_id))
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buffer.getvalue()


def read_excel(data, filename):
    preview = pd.read_excel(io.BytesIO(data), sheet_name=0, header=None, nrows=25, engine="openpyxl")
    targets = {normalize("STATE"), normalize("QDC_Name")}
    best_row, best_score = 0, -1
    for row_number in range(len(preview)):
        row_tokens = {
            normalize(value)
            for value in preview.iloc[row_number].tolist()
            if pd.notna(value)
        }
        score = len(row_tokens & targets)
        if score > best_score:
            best_row, best_score = row_number, score
    frame = pd.read_excel(io.BytesIO(data), sheet_name=0, header=best_row, engine="openpyxl")
    frame.columns = [str(column).strip() for column in frame.columns]
    logging.info("%s: header row %s, %s rows x %s columns", filename, best_row + 1, len(frame), len(frame.columns))
    return frame, best_row + 1


def parse_dates(series):
    text = clean(series)
    parsed = pd.to_datetime(text, format="ISO8601", errors="coerce")
    unresolved = parsed.isna() & text.notna()
    if unresolved.any():
        try:
            parsed.loc[unresolved] = pd.to_datetime(
                text.loc[unresolved], format="mixed", dayfirst=True, errors="coerce"
            )
        except (TypeError, ValueError):
            parsed.loc[unresolved] = pd.to_datetime(
                text.loc[unresolved], dayfirst=True, errors="coerce"
            )
    return parsed


def awc_identifier(frame, detected):
    result = pd.Series(pd.NA, index=frame.index, dtype="string")
    for key in ("cal_awc", "current_awc", "given_awc", "awc_name"):
        column = detected.get(key)
        if column:
            result = result.fillna(clean(frame[column]))
    return result.str.replace(r"\.0$", "", regex=True).str.strip().str.casefold()


def standardize(raw, tool, filename, latitude_name, longitude_name, header_row):
    detected = {key: find_column(raw.columns, names) for key, names in ALIASES.items()}
    latitude_col = find_column(raw.columns, [latitude_name])
    longitude_col = find_column(raw.columns, [longitude_name])
    required = {
        "STATE": detected["state"],
        "QDC_Name": detected["fi"],
        "submission date/time": detected["date"],
        latitude_name: latitude_col,
        longitude_name: longitude_col,
    }
    missing = [label for label, column in required.items() if not column]
    if missing:
        raise ValueError(f"{filename}: missing {', '.join(missing)}")
    if not any(detected[key] for key in ("cal_awc", "current_awc", "given_awc", "awc_name")):
        raise ValueError(f"{filename}: no AWC identifier field found")

    frame = raw.copy()
    frame["__Tool"] = tool
    frame["__State"] = clean(frame[detected["state"]])
    frame["__FI"] = clean(frame[detected["fi"]])
    frame["__FIKey"] = frame["__FI"].str.casefold()
    frame["__DateTime"] = parse_dates(frame[detected["date"]])
    frame["__Date"] = frame["__DateTime"].dt.normalize()
    frame["__AWCKey"] = awc_identifier(frame, detected)
    frame["__Latitude"] = pd.to_numeric(frame[latitude_col], errors="coerce").where(lambda values: values.between(-90, 90))
    frame["__Longitude"] = pd.to_numeric(frame[longitude_col], errors="coerce").where(lambda values: values.between(-180, 180))

    log = {
        "Tool": tool,
        "File": filename,
        "Rows Read": len(raw),
        "Unique FIs": int(frame["__FIKey"].dropna().nunique()),
        "Start Date": frame["__Date"].min(),
        "Latest Date": frame["__Date"].max(),
        "Valid AWC IDs": int(frame["__AWCKey"].notna().sum()),
        "Status": "OK",
    }
    return frame, log


def load_tools(service, folder_id):
    frames, logs = {}, []
    for tool, filename, latitude, longitude in TOOL_FILES:
        logging.info("Reading %s", filename)
        try:
            raw, header_row = read_excel(download(service, find_file_id(service, folder_id, filename)), filename)
            frames[tool], log = standardize(raw, tool, filename, latitude, longitude, header_row)
            logs.append(log)
        except Exception as exc:
            logging.exception("Could not process %s", filename)
            logs.append({"Tool": tool, "File": filename, "Rows Read": 0, "Status": f"ERROR: {exc}"})
    missing = sorted(set(TOOL_NAMES) - set(frames))
    if missing:
        raise ValueError("Report stopped because tools failed: " + ", ".join(missing))
    return frames, pd.DataFrame(logs)


def get_states(frames):
    groups = {}
    for frame in frames.values():
        for value in frame["__State"].dropna().unique():
            display = str(value).strip()
            if display:
                groups.setdefault(display.casefold(), []).append(display)
    state_values = [max(values, key=lambda value: (values.count(value), len(value))) for values in groups.values()]
    selected = {value.strip().casefold() for value in env("STATES").split(",") if value.strip()}
    excluded = {value.strip().casefold() for value in env("EXCLUDED_STATES").split(",") if value.strip()}
    if selected:
        state_values = [state for state in state_values if state.casefold() in selected]
    return sorted([state for state in state_values if state.casefold() not in excluded], key=str.casefold)


def combine_state(frames, state):
    return pd.concat(
        [
            frame.loc[
                frame["__State"].fillna("").str.strip().str.casefold() == state.casefold()
            ].copy()
            for frame in frames.values()
        ],
        ignore_index=True,
        sort=False,
    )


def fi_roster(data):
    valid = data.dropna(subset=["__FIKey", "__FI"])
    return (
        valid.groupby(["__FIKey", "__FI"]).size().rename("Records").reset_index()
        .sort_values(["__FIKey", "Records", "__FI"], ascending=[True, False, True])
        .drop_duplicates("__FIKey")[["__FIKey", "__FI"]]
        .rename(columns={"__FI": "FI Name"})
    )


def latest_and_previous_dates(data):
    dates = data["__Date"].dropna()
    if dates.empty:
        raise ValueError("No valid submission dates found")
    latest = dates.max()
    return latest, latest - pd.Timedelta(days=1)


def cumulative_interviews(data, date_cutoff, roster):
    selected = data.loc[data["__Date"].notna() & (data["__Date"] <= date_cutoff)]
    totals = selected.groupby("__FIKey").size().rename("Total Interviews")
    separate_cd = selected.loc[selected["__Tool"] == "Separate CD"].groupby("__FIKey").size().rename("Separate CD")
    result = roster.set_index("__FIKey")
    result["Total Interviews"] = totals.reindex(result.index, fill_value=0).astype(int)
    result["Separate CD"] = separate_cd.reindex(result.index, fill_value=0).astype(int)
    return result


def awc_completion(data, date_cutoff, roster):
    """AWC is complete only when both Mother 0-5 and Mother 6-11 exist for same FI and AWC."""
    selected = data.loc[
        data["__Date"].notna()
        & (data["__Date"] <= date_cutoff)
        & data["__FIKey"].notna()
        & data["__AWCKey"].notna()
        & data["__Tool"].isin(["Mother 0-5", "Mother 6-11"])
    ]
    result = roster.set_index("__FIKey")
    if selected.empty:
        result["AWC with Two Tools Completed"] = 0
        result["AWC with Less Than Two Tools"] = 0
        return result

    tool_presence = (
        selected.drop_duplicates(["__FIKey", "__AWCKey", "__Tool"])
        .groupby(["__FIKey", "__AWCKey"])["__Tool"]
        .nunique()
    )
    complete = tool_presence.eq(2).groupby(level=0).sum()
    incomplete = tool_presence.lt(2).groupby(level=0).sum()
    result["AWC with Two Tools Completed"] = complete.reindex(result.index, fill_value=0).astype(int)
    result["AWC with Less Than Two Tools"] = incomplete.reindex(result.index, fill_value=0).astype(int)
    return result


def simple_status(data, latest, previous):
    roster = fi_roster(data)
    current = cumulative_interviews(data, latest, roster)
    prior = cumulative_interviews(data, previous, roster)
    awc = awc_completion(data, latest, roster)

    result = roster.set_index("__FIKey")
    result["Total Submission Till Latest Date"] = current["Total Interviews"]
    result["Total Submission Till Previous Date"] = prior["Total Interviews"]
    result["New Submission on Latest Date"] = result["Total Submission Till Latest Date"] - result["Total Submission Till Previous Date"]
    result["AWC with Two Tools Completed"] = awc["AWC with Two Tools Completed"]
    result["AWC with Less Than Two Tools"] = awc["AWC with Less Than Two Tools"]
    result["Separate CD Submission"] = current["Separate CD"]
    return result.reset_index(drop=True).sort_values("FI Name", key=lambda values: values.str.casefold()).reset_index(drop=True)


def haversine(lat1, lon1, lat2, lon2):
    if any(pd.isna(value) for value in (lat1, lon1, lat2, lon2)):
        return np.nan
    radius = 6371.0088
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    delta_phi = math.radians(lat2 - lat1)
    delta_lambda = math.radians(lon2 - lon1)
    a = math.sin(delta_phi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(delta_lambda / 2) ** 2
    return radius * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def fi_weekly_travel(data):
    """One row per FI per ISO week, showing cumulative distance between daily median GPS locations."""
    gps = data.dropna(subset=["__FIKey", "__FI", "__Date", "__Latitude", "__Longitude"]).copy()
    if gps.empty:
        return pd.DataFrame(columns=["FI Name", "Week Start", "Weekly Travel km", "Travelled At Least 10 km"])

    display = (
        gps.groupby(["__FIKey", "__FI"]).size().rename("Count").reset_index()
        .sort_values(["__FIKey", "Count", "__FI"], ascending=[True, False, True])
        .drop_duplicates("__FIKey")
        .set_index("__FIKey")["__FI"]
    )
    daily = (
        gps.groupby(["__FIKey", "__Date"], as_index=False)
        .agg(Latitude=("__Latitude", "median"), Longitude=("__Longitude", "median"))
        .sort_values(["__FIKey", "__Date"])
    )
    daily["Previous Latitude"] = daily.groupby("__FIKey")["Latitude"].shift()
    daily["Previous Longitude"] = daily.groupby("__FIKey")["Longitude"].shift()
    daily["Distance km"] = daily.apply(
        lambda row: haversine(
            row["Previous Latitude"],
            row["Previous Longitude"],
            row["Latitude"],
            row["Longitude"],
        ),
        axis=1,
    )
    iso = daily["__Date"].dt.isocalendar()
    daily["ISO Year"] = iso.year.astype(int)
    daily["ISO Week"] = iso.week.astype(int)
    daily["Week Start"] = daily["__Date"] - pd.to_timedelta(daily["__Date"].dt.weekday, unit="D")

    weekly = daily.groupby(["__FIKey", "ISO Year", "ISO Week", "Week Start"], as_index=False).agg(
        **{
            "Weekly Travel km": ("Distance km", "sum"),
            "Valid Day-to-Day Comparisons": ("Distance km", "count"),
            "GPS Active Days": ("__Date", "nunique"),
        }
    )
    weekly["FI Name"] = weekly["__FIKey"].map(display)
    weekly["Travelled At Least 10 km"] = np.where(
        weekly["Valid Day-to-Day Comparisons"] > 0,
        weekly["Weekly Travel km"] >= TRAVEL_THRESHOLD_KM,
        pd.NA,
    )
    weekly["10 km Status"] = np.select(
        [
            weekly["Valid Day-to-Day Comparisons"] == 0,
            weekly["Weekly Travel km"] >= TRAVEL_THRESHOLD_KM,
        ],
        ["Not enough GPS days", "Yes"],
        default="No",
    )
    return weekly[
        [
            "FI Name", "ISO Year", "ISO Week", "Week Start", "GPS Active Days",
            "Valid Day-to-Day Comparisons", "Weekly Travel km", "10 km Status",
        ]
    ].sort_values(["Week Start", "FI Name"], ascending=[False, True])


def safe_filename(value):
    return re.sub(r'[<>:"/\\|?*]+', "_", str(value)).strip() or "Unknown_State"


def format_sheet(writer, sheet_name, dataframe, title=None, date_columns=None, decimal_columns=None, zero_columns=None):
    worksheet = writer.sheets[sheet_name]
    workbook = writer.book
    header_row = 2 if title else 0
    header = workbook.add_format({
        "bold": True,
        "bg_color": "#B7DEE8",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "text_wrap": True,
    })
    title_format = workbook.add_format({"bold": True, "font_size": 14})
    date_format = workbook.add_format({"num_format": "dd-mm-yyyy"})
    decimal_format = workbook.add_format({"num_format": "0.0"})
    zero_format = workbook.add_format({"bg_color": "#F4CCCC"})
    if title:
        worksheet.write(0, 0, title, title_format)
    worksheet.set_row(header_row, 38, header)
    worksheet.freeze_panes(header_row + 1, 1)
    if len(dataframe.columns):
        worksheet.autofilter(header_row, 0, header_row + len(dataframe), len(dataframe.columns) - 1)
    for index, column in enumerate(dataframe.columns):
        lengths = dataframe.head(100)[column].dropna().astype(str).str.len()
        width = min(max(len(str(column)) + 2, int(lengths.max()) + 2 if not lengths.empty else 10), 34)
        cell_format = date_format if date_columns and column in date_columns else decimal_format if decimal_columns and column in decimal_columns else None
        worksheet.set_column(index, index, width, cell_format)
        if zero_columns and column in zero_columns and len(dataframe):
            worksheet.conditional_format(
                header_row + 1,
                index,
                header_row + len(dataframe),
                index,
                {"type": "cell", "criteria": "==", "value": 0, "format": zero_format},
            )


def create_workbook(state, data, logs, output_folder):
    latest, previous = latest_and_previous_dates(data)
    summary = simple_status(data, latest, previous)
    travel = fi_weekly_travel(data)
    logging.info(
        "%s: cumulative through %s; previous through %s; FI count=%s",
        state,
        latest.date(),
        previous.date(),
        len(summary),
    )

    output_folder.mkdir(parents=True, exist_ok=True)
    path = output_folder / f"{safe_filename(state).upper()}_IDD_Extension_Simple_Summary_{datetime.now():%d-%m-%Y}.xlsx"
    with pd.ExcelWriter(path, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        summary.to_excel(writer, sheet_name="FI Summary", index=False, startrow=2)
        format_sheet(
            writer,
            "FI Summary",
            summary,
            title=f"FI cumulative status through {latest:%d-%m-%Y}; previous cut-off {previous:%d-%m-%Y} | {state}",
            zero_columns={
                "New Submission on Latest Date",
                "AWC with Two Tools Completed",
                "Separate CD Submission",
            },
        )

        travel.to_excel(writer, sheet_name="FI Weekly 10km Travel", index=False, startrow=2)
        format_sheet(
            writer,
            "FI Weekly 10km Travel",
            travel,
            title=f"Weekly travel by FI; threshold {TRAVEL_THRESHOLD_KM:g} km | {state}",
            date_columns={"Week Start"},
            decimal_columns={"Weekly Travel km"},
        )

        process = logs.assign(
            State_Workbook=state,
            Latest_Database_Date=latest,
            Previous_Cutoff_Date=previous,
            FI_Count=len(summary),
        )
        process.to_excel(writer, sheet_name="Processing Log", index=False)
        format_sheet(
            writer,
            "Processing Log",
            process,
            date_columns={"Start Date", "Latest Date", "Latest_Database_Date", "Previous_Cutoff_Date"},
        )

        for tool in TOOL_NAMES:
            raw = data.loc[data["__Tool"] == tool].drop(
                columns=[column for column in data.columns if str(column).startswith("__")],
                errors="ignore",
            )
            sheet_name = re.sub(r"[\\/*?:\[\]]", "_", f"Raw {tool}")[:31]
            raw.to_excel(writer, sheet_name=sheet_name, index=False)
            format_sheet(writer, sheet_name, raw)
    logging.info("Created workbook: %s", path)
    return path


def create_zip(paths, output_folder):
    path = output_folder / f"IDD_Extension_Simple_Summary_All_States_{datetime.now():%d-%m-%Y}.zip"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for item in paths:
            archive.write(item, arcname=item.name)
    return path


def send_email(attachment, states):
    sender = env("SMTP_EMAIL", "GMAIL_USERNAME", required=True)
    password = env("SMTP_APP_PASSWORD", "GMAIL_APP_PASSWORD", required=True)
    recipients = [
        value.strip()
        for value in re.split(r"[,;]", env("RECIPIENTS", "REPORT_RECIPIENTS", required=True))
        if value.strip()
    ]
    size_mb = attachment.stat().st_size / (1024 * 1024)
    limit_mb = float(env("MAX_EMAIL_ATTACHMENT_MB") or "24")
    if size_mb > limit_mb:
        raise ValueError(f"ZIP is {size_mb:.2f} MB, above {limit_mb:.2f} MB")
    message = EmailMessage()
    message["From"] = sender
    message["To"] = ", ".join(recipients)
    message["Subject"] = env("MAIL_SUBJECT") or f"IDD Extension Simple Summary - {datetime.now():%d-%m-%Y}"
    message.set_content(
        f"Dear Team,\n\nPlease find attached the simplified IDD Extension FI summary for: {', '.join(states)}.\n\nRegards,\nGAVB Reporting Automation"
    )
    with attachment.open("rb") as handle:
        message.add_attachment(handle.read(), maintype="application", subtype="zip", filename=attachment.name)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context(), timeout=60) as smtp:
        smtp.login(sender, password)
        smtp.send_message(message)
    logging.info("Email sent successfully")


def main():
    logging.basicConfig(
        level=(env("LOG_LEVEL") or "INFO").upper(),
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
    try:
        logging.info("Running report script version: %s", SCRIPT_VERSION)
        output_folder = Path(
            env("LOCAL_OUTPUT_FOLDER")
            or f"/tmp/{datetime.now(timezone.utc):%Y-%m-%d}/idd_extension_simple"
        )
        service = build("drive", "v3", credentials=credentials(), cache_discovery=False)
        frames, logs = load_tools(service, env("DRIVE_FOLDER_ID", required=True))
        report_paths, completed_states = [], []
        for state in get_states(frames):
            data = combine_state(frames, state)
            if not data.empty:
                report_paths.append(create_workbook(state, data, logs, output_folder))
                completed_states.append(state)
        if not report_paths:
            raise ValueError("No reports created")
        zip_path = create_zip(report_paths, output_folder)
        send_email(zip_path, completed_states)
        print("Created reports:")
        for report_path in report_paths:
            print(" -", report_path)
        print("Email attachment:", zip_path)
        return 0
    except (ValueError, FileNotFoundError, json.JSONDecodeError) as exc:
        logging.error("Configuration/runtime error: %s", exc)
        return 2
    except HttpError as exc:
        logging.error("Google API error: %s", exc)
        return 3
    except smtplib.SMTPAuthenticationError:
        logging.exception("Gmail authentication failed")
        return 4
    except (smtplib.SMTPException, OSError) as exc:
        logging.exception("Email delivery failed: %s", exc)
        return 5
    except Exception as exc:
        logging.exception("Unexpected failure: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())

