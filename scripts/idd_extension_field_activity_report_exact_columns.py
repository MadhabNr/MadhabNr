#!/usr/bin/env python3
"""State-wise IDD Extension submission summary using each FI's own latest active date."""

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

SCRIPT_VERSION = "2026-09-25-idd-per-fi-latest-corrected-v5"
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
WEEKLY_TRAVEL_KM = float(os.getenv("WEEKLY_TRAVEL_KM", "10"))

TOOL_FILES = [
    ("Mother 0-5", "IDD_Extension_Mother_0_5_WIDE.xlsx"),
    ("Mother 6-11", "IDD Extension Mother 6-11_WIDE.xlsx"),
    ("Separate CD", "IDD_Extension_Separate_CD_Tool_WIDE.xlsx"),
]
TOOL_ORDER = [tool for tool, _ in TOOL_FILES]

COMMON_ALIASES = {
    "state": ["STATE", "State", "state", "Cal_STATE", "state_name"],
    "fi": ["QDC_Name", "QDC Name", "Field Investigator Name", "Investigator", "QDC", "collector_name"],
    "submission": [
        "SubmissionDate", "Submission Date", "submission_date", "SubmissionDateTime",
        "Submission_Time", "endtime", "EndTime", "starttime", "StartTime",
    ],
}

GPS_COLUMNS = {
    "Mother 0-5": ("Qgeo-Latitude", "Qgeo-Longitude"),
    "Mother 6-11": ("QX1-Latitude", "QX1-Longitude"),
    "Separate CD": ("Qgeo-Latitude", "Qgeo-Longitude"),
}


def env(*names, required=False):
    for name in names:
        value = os.getenv(name, "").strip()
        if value:
            return value
    if required:
        raise ValueError("Missing required environment variable. Accepted name(s): " + ", ".join(names))
    return ""


def normalize(value):
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def normalize_state(value):
    return str(value).strip().casefold()


def clean_text(series):
    return series.astype("string").str.strip().replace(
        {"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA}
    )


def find_column(columns, aliases):
    lookup = {normalize(column): str(column) for column in columns}
    for alias in aliases:
        if normalize(alias) in lookup:
            return lookup[normalize(alias)]
    return None


def build_credentials():
    raw = env("GOOGLE_SERVICE_ACCOUNT_JSON")
    credential_file = env("GOOGLE_SERVICE_ACCOUNT_FILE") or "credential.json"
    if raw:
        return Credentials.from_service_account_info(json.loads(raw), scopes=DRIVE_SCOPES)
    if os.path.exists(credential_file):
        return Credentials.from_service_account_file(credential_file, scopes=DRIVE_SCOPES)
    raise ValueError("Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_SERVICE_ACCOUNT_FILE")


def escape_drive_value(value):
    return value.replace("\\", "\\\\").replace("'", "\\'")


def find_drive_file(service, folder_id, filename):
    query = f"'{folder_id}' in parents and trashed = false and name = '{escape_drive_value(filename)}'"
    files = service.files().list(
        q=query, spaces="drive", fields="files(id,name)", pageSize=10
    ).execute().get("files", [])
    if not files:
        raise FileNotFoundError(f"File not found in Drive folder: {filename}")
    return files[0]["id"]


def download_drive_file(service, file_id):
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, service.files().get_media(fileId=file_id))
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buffer.getvalue()


def detect_header_and_read(data, filename):
    preview = pd.read_excel(
        io.BytesIO(data), sheet_name=0, header=None, nrows=25, engine="openpyxl"
    )
    target_names = {normalize("STATE"), normalize("QDC_Name")}
    best_row, best_score = 0, -1
    for row_number in range(len(preview)):
        values = {
            normalize(value)
            for value in preview.iloc[row_number].tolist()
            if pd.notna(value)
        }
        score = len(target_names & values)
        if score > best_score:
            best_row, best_score = row_number, score
    frame = pd.read_excel(
        io.BytesIO(data), sheet_name=0, header=best_row, engine="openpyxl"
    )
    frame.columns = [str(column).strip() for column in frame.columns]
    logging.info(
        "%s: detected header row %s with %s rows x %s columns",
        filename, best_row + 1, len(frame), len(frame.columns),
    )
    return frame, best_row + 1


def parse_submission_datetime(series):
    text = clean_text(series)
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


def parse_coordinate(series, latitude):
    values = pd.to_numeric(series, errors="coerce")
    return values.where(values.between(-90, 90) if latitude else values.between(-180, 180))


def standardize(raw, tool, filename, header_row):
    state_col = find_column(raw.columns, COMMON_ALIASES["state"])
    fi_col = find_column(raw.columns, COMMON_ALIASES["fi"])
    submission_col = find_column(raw.columns, COMMON_ALIASES["submission"])
    latitude_expected, longitude_expected = GPS_COLUMNS[tool]
    latitude_col = find_column(raw.columns, [latitude_expected])
    longitude_col = find_column(raw.columns, [longitude_expected])

    missing = []
    if not state_col:
        missing.append("STATE")
    if not fi_col:
        missing.append("QDC_Name")
    if not submission_col:
        missing.append("submission date/time")
    if not latitude_col:
        missing.append(latitude_expected)
    if not longitude_col:
        missing.append(longitude_expected)
    if missing:
        raise ValueError(f"{filename}: missing required column(s): {', '.join(missing)}")

    frame = raw.copy()
    frame["__Tool"] = tool
    frame["__State"] = clean_text(frame[state_col])
    frame["__FI"] = clean_text(frame[fi_col])
    frame["__FIKey"] = frame["__FI"].str.casefold()
    frame["__DateTime"] = parse_submission_datetime(frame[submission_col])
    frame["__Date"] = frame["__DateTime"].dt.normalize()
    frame["__Latitude"] = parse_coordinate(frame[latitude_col], True)
    frame["__Longitude"] = parse_coordinate(frame[longitude_col], False)

    log = {
        "Tool": tool,
        "File": filename,
        "Header Row": header_row,
        "Rows Read": len(raw),
        "State Column": state_col,
        "FI Column": fi_col,
        "Submission Column": submission_col,
        "Latitude Column": latitude_col,
        "Longitude Column": longitude_col,
        "Valid FI Names": int(frame["__FI"].notna().sum()),
        "Valid Dates": int(frame["__Date"].notna().sum()),
        "Valid GPS": int((frame["__Latitude"].notna() & frame["__Longitude"].notna()).sum()),
        "Status": "OK",
    }
    return frame, log


def load_tools(service, folder_id):
    frames, logs = {}, []
    for tool, filename in TOOL_FILES:
        logging.info("Reading %s", filename)
        try:
            raw_bytes = download_drive_file(service, find_drive_file(service, folder_id, filename))
            raw, header_row = detect_header_and_read(raw_bytes, filename)
            frames[tool], log = standardize(raw, tool, filename, header_row)
            logs.append(log)
        except Exception as exc:
            logging.exception("Could not process %s", filename)
            logs.append({"Tool": tool, "File": filename, "Rows Read": 0, "Status": f"ERROR: {exc}"})

    missing_tools = sorted(set(TOOL_ORDER) - set(frames))
    if missing_tools:
        raise ValueError(
            "Report generation stopped because these tools could not be processed: "
            + ", ".join(missing_tools)
        )
    return frames, pd.DataFrame(logs)


def get_states(frames):
    grouped = {}
    for frame in frames.values():
        for value in frame["__State"].dropna().unique():
            display = str(value).strip()
            if display:
                grouped.setdefault(normalize_state(display), []).append(display)
    states = [max(values, key=lambda value: (values.count(value), len(value))) for values in grouped.values()]

    selected = env("STATES")
    if selected:
        wanted = {normalize_state(value) for value in selected.split(",") if value.strip()}
        states = [state for state in states if normalize_state(state) in wanted]
    excluded = {normalize_state(value) for value in env("EXCLUDED_STATES").split(",") if value.strip()}
    return sorted([state for state in states if normalize_state(state) not in excluded], key=str.casefold)


def combine_state_data(frames, state):
    parts = []
    for frame in frames.values():
        mask = frame["__State"].fillna("").str.strip().str.casefold() == normalize_state(state)
        parts.append(frame.loc[mask].copy())
    return pd.concat(parts, ignore_index=True, sort=False)


def fi_latest_status_table(data):
    """Include every FI, using the latest valid date separately for each FI."""
    valid = data.dropna(subset=["__FI", "__FIKey", "__Date"]).copy()
    if valid.empty:
        return pd.DataFrame()

    # Keep the most frequently used display spelling for each normalized FI name.
    display_names = (
        valid.groupby(["__FIKey", "__FI"]).size().rename("n").reset_index()
        .sort_values(["__FIKey", "n", "__FI"], ascending=[True, False, True])
        .drop_duplicates("__FIKey")
        .set_index("__FIKey")["__FI"]
    )

    latest_dates = valid.groupby("__FIKey")["__Date"].max().rename("Latest Submission Date")
    latest_rows = valid.merge(latest_dates, left_on="__FIKey", right_index=True)
    latest_rows = latest_rows.loc[latest_rows["__Date"] == latest_rows["Latest Submission Date"]]

    counts = latest_rows.pivot_table(
        index="__FIKey", columns="__Tool", values="__State", aggfunc="size", fill_value=0
    )
    counts = counts.reindex(columns=TOOL_ORDER, fill_value=0)
    result = counts.join(latest_dates).reset_index()
    result.insert(0, "Name of Field Investigators", result["__FIKey"].map(display_names))
    result = result.drop(columns="__FIKey")
    result["Total Submissions"] = result[TOOL_ORDER].sum(axis=1).astype(int)
    result = result[["Name of Field Investigators", "Latest Submission Date", *TOOL_ORDER, "Total Submissions"]]
    result = result.sort_values("Name of Field Investigators", key=lambda s: s.str.casefold()).reset_index(drop=True)

    total = {
        "Name of Field Investigators": "Total Submissions",
        "Latest Submission Date": pd.NaT,
        **{tool: int(result[tool].sum()) for tool in TOOL_ORDER},
        "Total Submissions": int(result["Total Submissions"].sum()),
    }
    return pd.concat([result, pd.DataFrame([total])], ignore_index=True)


def fi_latest_vs_previous_table(data):
    """For every FI and tool, compare that FI's latest active date with previous active date."""
    valid = data.dropna(subset=["__FI", "__FIKey", "__Date"]).copy()
    if valid.empty:
        return pd.DataFrame()

    display_names = (
        valid.groupby(["__FIKey", "__FI"]).size().rename("n").reset_index()
        .sort_values(["__FIKey", "n", "__FI"], ascending=[True, False, True])
        .drop_duplicates("__FIKey")
        .set_index("__FIKey")["__FI"]
    )
    daily = valid.groupby(["__FIKey", "__Date", "__Tool"]).size().rename("Count").reset_index()
    all_dates = daily[["__FIKey", "__Date"]].drop_duplicates().sort_values(["__FIKey", "__Date"])
    date_lists = all_dates.groupby("__FIKey")["__Date"].apply(list)

    rows = []
    for fi_key, dates in date_lists.items():
        latest_date = dates[-1]
        previous_date = dates[-2] if len(dates) > 1 else pd.NaT
        row = {
            "Name of Field Investigators": display_names.get(fi_key, fi_key),
            "Latest Submission Date": latest_date,
            "Previous Active Date": previous_date,
        }
        latest_total, previous_total = 0, 0
        for tool in TOOL_ORDER:
            latest_count = int(daily.loc[
                (daily["__FIKey"] == fi_key) & (daily["__Date"] == latest_date) & (daily["__Tool"] == tool), "Count"
            ].sum())
            previous_count = int(daily.loc[
                (daily["__FIKey"] == fi_key) & (daily["__Date"] == previous_date) & (daily["__Tool"] == tool), "Count"
            ].sum()) if pd.notna(previous_date) else 0
            row[f"Latest {tool}"] = latest_count
            row[f"Previous {tool}"] = previous_count
            latest_total += latest_count
            previous_total += previous_count
        row["Latest Total"] = latest_total
        row["Previous Total"] = previous_total
        row["Change"] = latest_total - previous_total
        row["% Change"] = (latest_total - previous_total) / previous_total if previous_total else np.nan
        rows.append(row)
    return pd.DataFrame(rows).sort_values(
        "Name of Field Investigators", key=lambda s: s.str.casefold()
    ).reset_index(drop=True)


def tool_date_comparison_table(data):
    """Each tool uses its own latest and previous active dates, avoiding sparse cross-tool dates."""
    valid = data.dropna(subset=["__Date"]).copy()
    rows = []
    for tool in TOOL_ORDER:
        tool_data = valid.loc[valid["__Tool"] == tool]
        dates = sorted(tool_data["__Date"].dropna().unique())
        latest_date = dates[-1] if dates else pd.NaT
        previous_date = dates[-2] if len(dates) > 1 else pd.NaT
        latest_count = int((tool_data["__Date"] == latest_date).sum()) if pd.notna(latest_date) else 0
        previous_count = int((tool_data["__Date"] == previous_date).sum()) if pd.notna(previous_date) else 0
        change = latest_count - previous_count
        rows.append({
            "Tool": tool,
            "Latest Active Date": latest_date,
            "Latest Date Count": latest_count,
            "Previous Active Date": previous_date,
            "Previous Active Date Count": previous_count,
            "Change": change,
            "% Change": change / previous_count if previous_count else np.nan,
        })
    return pd.DataFrame(rows)


def haversine_km(lat1, lon1, lat2, lon2):
    if any(pd.isna(value) for value in (lat1, lon1, lat2, lon2)):
        return np.nan
    radius = 6371.0088
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    delta_phi = math.radians(lat2 - lat1)
    delta_lambda = math.radians(lon2 - lon1)
    a = math.sin(delta_phi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(delta_lambda / 2) ** 2
    return radius * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def weekly_travel_summary(data):
    gps = data.dropna(subset=["__FIKey", "__Date", "__Latitude", "__Longitude"]).copy()
    if gps.empty:
        return pd.DataFrame(columns=[
            "ISO Year", "ISO Week", "Week Start", "FIs with Valid GPS Comparisons",
            "FIs Travelled At Least 10 km", "Percent FIs Travelled >= 10 km",
        ])
    daily = gps.groupby(["__FIKey", "__Date"], as_index=False).agg(
        Latitude=("__Latitude", "median"), Longitude=("__Longitude", "median")
    ).sort_values(["__FIKey", "__Date"])
    daily["Previous Latitude"] = daily.groupby("__FIKey")["Latitude"].shift()
    daily["Previous Longitude"] = daily.groupby("__FIKey")["Longitude"].shift()
    daily["Distance km"] = daily.apply(
        lambda row: haversine_km(
            row["Previous Latitude"], row["Previous Longitude"], row["Latitude"], row["Longitude"]
        ), axis=1,
    )
    iso = daily["__Date"].dt.isocalendar()
    daily["ISO Year"] = iso.year.astype(int)
    daily["ISO Week"] = iso.week.astype(int)
    daily["Week Start"] = daily["__Date"] - pd.to_timedelta(daily["__Date"].dt.weekday, unit="D")

    fi_week = daily.groupby(["__FIKey", "ISO Year", "ISO Week", "Week Start"], as_index=False).agg(
        Weekly_Travel_km=("Distance km", "sum"),
        Valid_Distance_Comparisons=("Distance km", "count"),
    )
    fi_week = fi_week.loc[fi_week["Valid_Distance_Comparisons"] > 0].copy()
    fi_week["Travelled At Least Threshold"] = fi_week["Weekly_Travel_km"] >= WEEKLY_TRAVEL_KM
    if fi_week.empty:
        return pd.DataFrame()

    summary = fi_week.groupby(["ISO Year", "ISO Week", "Week Start"], as_index=False).agg(
        **{
            "FIs with Valid GPS Comparisons": ("__FIKey", "nunique"),
            "FIs Travelled At Least 10 km": ("Travelled At Least Threshold", "sum"),
        }
    )
    summary["Percent FIs Travelled >= 10 km"] = (
        summary["FIs Travelled At Least 10 km"] / summary["FIs with Valid GPS Comparisons"]
    )
    return summary.sort_values("Week Start", ascending=False)


def safe_filename(value):
    return re.sub(r'[<>:"/\\|?*]+', "_", str(value)).strip() or "Unknown_State"


def format_sheet(writer, sheet_name, dataframe, title=None, date_columns=None, percent_columns=None, zero_columns=None):
    worksheet = writer.sheets[sheet_name]
    workbook = writer.book
    header_row = 2 if title else 0
    header_format = workbook.add_format({
        "bold": True, "bg_color": "#B7DEE8", "border": 1,
        "align": "center", "valign": "vcenter", "text_wrap": True,
    })
    title_format = workbook.add_format({"bold": True, "font_size": 14})
    date_format = workbook.add_format({"num_format": "dd-mm-yyyy"})
    percent_format = workbook.add_format({"num_format": "+0.0%;-0.0%;-"})
    zero_format = workbook.add_format({"bg_color": "#F4CCCC"})
    total_format = workbook.add_format({"bold": True, "top": 1})

    if title:
        worksheet.write(0, 0, title, title_format)
    worksheet.set_row(header_row, 38, header_format)
    worksheet.freeze_panes(header_row + 1, 1)
    if len(dataframe.columns):
        worksheet.autofilter(header_row, 0, header_row + len(dataframe), len(dataframe.columns) - 1)

    sample = dataframe.head(100)
    for index, column in enumerate(dataframe.columns):
        lengths = sample[column].dropna().astype(str).str.len()
        width = min(max(len(str(column)) + 2, int(lengths.max()) + 2 if not lengths.empty else 10), 34)
        column_format = None
        if date_columns and column in date_columns:
            column_format = date_format
        elif percent_columns and column in percent_columns:
            column_format = percent_format
        worksheet.set_column(index, index, width, column_format)
        if zero_columns and column in zero_columns and len(dataframe):
            worksheet.conditional_format(
                header_row + 1, index, header_row + len(dataframe), index,
                {"type": "cell", "criteria": "==", "value": 0, "format": zero_format},
            )
    if len(dataframe) and str(dataframe.iloc[-1, 0]) == "Total Submissions":
        worksheet.set_row(header_row + len(dataframe), None, total_format)


def create_state_workbook(state, data, logs, output_dir):
    fi_latest = fi_latest_status_table(data)
    fi_change = fi_latest_vs_previous_table(data)
    tool_change = tool_date_comparison_table(data)
    weekly = weekly_travel_summary(data)
    if fi_latest.empty:
        raise ValueError(f"{state}: no valid FI and submission-date records")

    output_dir.mkdir(parents=True, exist_ok=True)
    report_date = datetime.now().strftime("%d-%m-%Y")
    path = output_dir / f"{safe_filename(state).upper()}_IDD_Extension_Submission_Summary_{report_date}.xlsx"

    with pd.ExcelWriter(path, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        fi_latest.to_excel(writer, sheet_name="FI Latest Tool Status", index=False, startrow=2)
        format_sheet(
            writer, "FI Latest Tool Status", fi_latest,
            title=f"FI-wise submissions on each FI's latest active date | {state}",
            date_columns={"Latest Submission Date"},
            zero_columns=set(TOOL_ORDER),
        )

        fi_change.to_excel(writer, sheet_name="FI Latest vs Previous", index=False, startrow=2)
        format_sheet(
            writer, "FI Latest vs Previous", fi_change,
            title=f"FI-wise latest active date compared with previous active date | {state}",
            date_columns={"Latest Submission Date", "Previous Active Date"},
            percent_columns={"% Change"},
        )

        tool_change.to_excel(writer, sheet_name="Tool Latest vs Previous", index=False, startrow=2)
        format_sheet(
            writer, "Tool Latest vs Previous", tool_change,
            title=f"Tool-wise latest active date compared with previous active date | {state}",
            date_columns={"Latest Active Date", "Previous Active Date"},
            percent_columns={"% Change"},
            zero_columns={"Latest Date Count"},
        )

        weekly.to_excel(writer, sheet_name="Weekly 10km Travel", index=False)
        format_sheet(
            writer, "Weekly 10km Travel", weekly,
            date_columns={"Week Start"},
            percent_columns={"Percent FIs Travelled >= 10 km"},
        )

        state_log = logs.assign(State_Workbook=state)
        state_log.to_excel(writer, sheet_name="Processing Log", index=False)
        format_sheet(writer, "Processing Log", state_log)

        for tool, _ in TOOL_FILES:
            tool_data = data.loc[data["__Tool"] == tool]
            raw = tool_data.drop(columns=[column for column in tool_data.columns if str(column).startswith("__")], errors="ignore")
            sheet_name = re.sub(r"[\\/*?:\[\]]", "_", f"Raw {tool}")[:31]
            raw.to_excel(writer, sheet_name=sheet_name, index=False)
            format_sheet(writer, sheet_name, raw)

    logging.info("Created state workbook: %s", path)
    return path


def create_zip(report_paths, output_dir):
    zip_path = output_dir / f"IDD_Extension_Submission_Summary_All_States_{datetime.now():%d-%m-%Y}.zip"
    with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for report_path in report_paths:
            archive.write(report_path, arcname=report_path.name)
    return zip_path


def send_email(attachment, states):
    sender = env("SMTP_EMAIL", "GMAIL_USERNAME", "EMAIL_USERNAME", required=True)
    password = env("SMTP_APP_PASSWORD", "GMAIL_APP_PASSWORD", "EMAIL_PASSWORD", required=True)
    recipients_raw = env("RECIPIENTS", "REPORT_RECIPIENTS", "EMAIL_TO", required=True)
    recipients = [value.strip() for value in re.split(r"[,;]", recipients_raw) if value.strip()]
    if not recipients:
        raise ValueError("RECIPIENTS is empty")

    attachment_size_mb = attachment.stat().st_size / (1024 * 1024)
    maximum_size_mb = float(env("MAX_EMAIL_ATTACHMENT_MB") or "24")
    if attachment_size_mb > maximum_size_mb:
        raise ValueError(
            f"ZIP attachment is {attachment_size_mb:.2f} MB, above configured limit {maximum_size_mb:.2f} MB"
        )

    message = EmailMessage()
    message["From"] = sender
    message["To"] = ", ".join(recipients)
    message["Subject"] = env("MAIL_SUBJECT", "EMAIL_SUBJECT") or f"IDD Extension Submission Summary - {datetime.now():%d-%m-%Y}"
    message.set_content(
        env("MAIL_BODY", "EMAIL_BODY")
        or f"Dear Team,\n\nPlease find attached the corrected IDD Extension submission summary for: {', '.join(states)}.\n\nRegards,\nGAVB Reporting Automation"
    )
    with attachment.open("rb") as handle:
        message.add_attachment(handle.read(), maintype="application", subtype="zip", filename=attachment.name)

    with smtplib.SMTP_SSL(
        env("SMTP_HOST") or "smtp.gmail.com",
        int(env("SMTP_PORT") or "465"),
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
            or f"/tmp/{datetime.now(timezone.utc):%Y-%m-%d}/idd_extension_summary"
        )
        service = build("drive", "v3", credentials=build_credentials(), cache_discovery=False)
        frames, logs = load_tools(service, folder_id)
        states = get_states(frames)
        if not states:
            raise ValueError("No reportable state values were found")

        reports, completed_states = [], []
        for state in states:
            data = combine_state_data(frames, state)
            if data.empty:
                continue
            reports.append(create_state_workbook(state, data, logs, output_dir))
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
