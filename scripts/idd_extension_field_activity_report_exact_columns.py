#!/usr/bin/env python3
"""Cumulative IDD Extension FI and AWC reporting for all states."""

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

SCRIPT_VERSION = "2026-09-25-idd-cumulative-awc-v8"
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
WEEKLY_TRAVEL_KM = float(os.getenv("WEEKLY_TRAVEL_KM", "10"))

TOOLS = [
    ("Mother 0-5", "IDD_Extension_Mother_0_5_WIDE.xlsx", "Qgeo-Latitude", "Qgeo-Longitude"),
    ("Mother 6-11", "IDD Extension Mother 6-11_WIDE.xlsx", "QX1-Latitude", "QX1-Longitude"),
    ("Separate CD", "IDD_Extension_Separate_CD_Tool_WIDE.xlsx", "Qgeo-Latitude", "Qgeo-Longitude"),
]
TOOL_NAMES = [tool for tool, _, _, _ in TOOLS]

ALIASES = {
    "state": ["STATE", "State", "state", "Cal_STATE", "state_name"],
    "fi": ["QDC_Name", "QDC Name", "Field Investigator Name", "Investigator", "QDC", "collector_name"],
    "date": [
        "SubmissionDate", "Submission Date", "submission_date", "SubmissionDateTime",
        "Submission_Time", "endtime", "EndTime", "starttime", "StartTime",
    ],
    "awc_calculated": ["Cal_AWC"],
    "awc_given": ["QAWC_code"],
    "awc_current": ["QAWC_3"],
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


def norm(value):
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def clean(series):
    return series.astype("string").str.strip().replace(
        {"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA}
    )


def find_col(columns, aliases):
    lookup = {norm(column): str(column) for column in columns}
    for alias in aliases:
        if norm(alias) in lookup:
            return lookup[norm(alias)]
    return None


def credentials():
    raw = env("GOOGLE_SERVICE_ACCOUNT_JSON")
    path = env("GOOGLE_SERVICE_ACCOUNT_FILE") or "credential.json"
    if raw:
        return Credentials.from_service_account_info(json.loads(raw), scopes=SCOPES)
    if os.path.exists(path):
        return Credentials.from_service_account_file(path, scopes=SCOPES)
    raise ValueError("Google service account credentials not found")


def drive_file_id(service, folder_id, filename):
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
    targets = {norm("STATE"), norm("QDC_Name")}
    best_row, best_score = 0, -1
    for row_number in range(len(preview)):
        values = {norm(v) for v in preview.iloc[row_number].tolist() if pd.notna(v)}
        score = len(values & targets)
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


def combine_awc_identifier(frame, detected):
    """Priority: Cal_AWC, current QAWC_3, given QAWC_code, then normalized QAWC_name."""
    result = pd.Series(pd.NA, index=frame.index, dtype="string")
    for key in ("awc_calculated", "awc_current", "awc_given", "awc_name"):
        column = detected.get(key)
        if column:
            candidate = clean(frame[column])
            result = result.fillna(candidate)
    result = result.str.replace(r"\.0$", "", regex=True).str.strip().str.casefold()
    return result


def standardize(raw, tool, filename, lat_name, lon_name, header_row):
    detected = {key: find_col(raw.columns, aliases) for key, aliases in ALIASES.items()}
    lat_col = find_col(raw.columns, [lat_name])
    lon_col = find_col(raw.columns, [lon_name])
    required = {
        "STATE": detected["state"], "QDC_Name": detected["fi"],
        "submission date/time": detected["date"], lat_name: lat_col, lon_name: lon_col,
    }
    missing = [label for label, column in required.items() if not column]
    if missing:
        raise ValueError(f"{filename}: missing {', '.join(missing)}")
    if not any(detected[key] for key in ("awc_calculated", "awc_current", "awc_given", "awc_name")):
        raise ValueError(f"{filename}: no AWC identifier found; expected Cal_AWC, QAWC_code, QAWC_3 or QAWC_name")

    frame = raw.copy()
    frame["__Tool"] = tool
    frame["__State"] = clean(frame[detected["state"]])
    frame["__FI"] = clean(frame[detected["fi"]])
    frame["__FIKey"] = frame["__FI"].str.casefold()
    frame["__DateTime"] = parse_dates(frame[detected["date"]])
    frame["__Date"] = frame["__DateTime"].dt.normalize()
    frame["__AWCKey"] = combine_awc_identifier(frame, detected)
    frame["__Latitude"] = pd.to_numeric(frame[lat_col], errors="coerce").where(lambda s: s.between(-90, 90))
    frame["__Longitude"] = pd.to_numeric(frame[lon_col], errors="coerce").where(lambda s: s.between(-180, 180))

    return frame, {
        "Tool": tool, "File": filename, "Header Row": header_row, "Rows Read": len(raw),
        "State Column": detected["state"], "FI Column": detected["fi"],
        "Submission Column": detected["date"], "Latitude Column": lat_col,
        "Longitude Column": lon_col,
        "Cal AWC Column": detected["awc_calculated"] or "",
        "Current AWC Column": detected["awc_current"] or "",
        "Given AWC Column": detected["awc_given"] or "",
        "AWC Name Column": detected["awc_name"] or "",
        "Unique FIs": int(frame["__FIKey"].dropna().nunique()),
        "Valid Dates": int(frame["__Date"].notna().sum()),
        "Valid AWC Identifiers": int(frame["__AWCKey"].notna().sum()),
        "Start Date Found": frame["__Date"].min(),
        "Latest Date Found": frame["__Date"].max(),
        "Status": "OK",
    }


def load_all(service, folder_id):
    frames, logs = {}, []
    for tool, filename, lat_name, lon_name in TOOLS:
        logging.info("Reading %s", filename)
        try:
            raw, header_row = read_excel(download(service, drive_file_id(service, folder_id, filename)), filename)
            frames[tool], log = standardize(raw, tool, filename, lat_name, lon_name, header_row)
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
    states = [max(values, key=lambda x: (values.count(x), len(x))) for values in groups.values()]
    selected = {x.strip().casefold() for x in env("STATES").split(",") if x.strip()}
    excluded = {x.strip().casefold() for x in env("EXCLUDED_STATES").split(",") if x.strip()}
    if selected:
        states = [state for state in states if state.casefold() in selected]
    return sorted([state for state in states if state.casefold() not in excluded], key=str.casefold)


def state_data(frames, state):
    return pd.concat([
        frame.loc[frame["__State"].fillna("").str.strip().str.casefold() == state.casefold()].copy()
        for frame in frames.values()
    ], ignore_index=True, sort=False)


def fi_roster(data):
    valid = data.dropna(subset=["__FIKey", "__FI"])
    return (
        valid.groupby(["__FIKey", "__FI"]).size().rename("Records").reset_index()
        .sort_values(["__FIKey", "Records", "__FI"], ascending=[True, False, True])
        .drop_duplicates("__FIKey")[["__FIKey", "__FI"]]
        .rename(columns={"__FI": "Name of Field Investigators"})
    )


def report_dates(data):
    dates = data["__Date"].dropna()
    if dates.empty:
        raise ValueError("No valid dates found")
    latest = dates.max()
    return latest, latest - pd.Timedelta(days=1)


def cumulative_submission_matrix(data, through_date, roster):
    selected = data.loc[data["__Date"].notna() & (data["__Date"] <= through_date)]
    counts = selected.groupby(["__FIKey", "__Tool"]).size().unstack(fill_value=0)
    counts = counts.reindex(index=roster["__FIKey"], columns=TOOL_NAMES, fill_value=0)
    result = roster.merge(counts.reset_index(), on="__FIKey", how="left")
    for tool in TOOL_NAMES:
        result[tool] = result[tool].fillna(0).astype(int)
    result["Total Interviews"] = result[TOOL_NAMES].sum(axis=1).astype(int)
    return result


def awc_metrics(data, through_date, roster):
    selected = data.loc[
        data["__Date"].notna() & (data["__Date"] <= through_date)
        & data["__FIKey"].notna() & data["__AWCKey"].notna()
    ].copy()
    base = roster.set_index("__FIKey")
    if selected.empty:
        base["Unique AWCs Visited"] = 0
        base["AWCs with At Least 2 Interviews"] = 0
        base["AWCs with Both Mother Interviews"] = 0
        return base.reset_index()

    unique_awc = selected.groupby("__FIKey")["__AWCKey"].nunique()
    awc_record_counts = selected.groupby(["__FIKey", "__AWCKey"]).size()
    two_interviews = awc_record_counts.ge(2).groupby(level=0).sum()

    mother = selected.loc[selected["__Tool"].isin(["Mother 0-5", "Mother 6-11"])]
    mother_pair = mother.groupby(["__FIKey", "__AWCKey"])["__Tool"].nunique().eq(2)
    both_mother = mother_pair.groupby(level=0).sum()

    base["Unique AWCs Visited"] = unique_awc.reindex(base.index, fill_value=0).astype(int)
    base["AWCs with At Least 2 Interviews"] = two_interviews.reindex(base.index, fill_value=0).astype(int)
    base["AWCs with Both Mother Interviews"] = both_mother.reindex(base.index, fill_value=0).astype(int)
    return base.reset_index()


def cumulative_status(data, latest):
    roster = fi_roster(data)
    submissions = cumulative_submission_matrix(data, latest, roster)
    awc = awc_metrics(data, latest, roster).drop(columns="Name of Field Investigators")
    result = submissions.merge(awc, on="__FIKey", how="left").drop(columns="__FIKey")
    result = result.sort_values("Name of Field Investigators", key=lambda s: s.str.casefold()).reset_index(drop=True)
    numeric = [*TOOL_NAMES, "Total Interviews", "Unique AWCs Visited", "AWCs with At Least 2 Interviews", "AWCs with Both Mother Interviews"]
    total = {"Name of Field Investigators": "Total / Sum", **{column: int(result[column].sum()) for column in numeric}}
    return pd.concat([result, pd.DataFrame([total])], ignore_index=True)


def cumulative_comparison(data, latest, previous):
    roster = fi_roster(data)
    current = cumulative_submission_matrix(data, latest, roster).set_index("__FIKey")
    prior = cumulative_submission_matrix(data, previous, roster).set_index("__FIKey")
    current_awc = awc_metrics(data, latest, roster).set_index("__FIKey")
    prior_awc = awc_metrics(data, previous, roster).set_index("__FIKey")
    result = roster.set_index("__FIKey")
    result["Current Total Interviews"] = current["Total Interviews"]
    result["Previous Total Interviews"] = prior["Total Interviews"]
    result["Interview Change"] = result["Current Total Interviews"] - result["Previous Total Interviews"]
    result["Interview % Change"] = np.where(
        result["Previous Total Interviews"] > 0,
        result["Interview Change"] / result["Previous Total Interviews"],
        np.where(result["Current Total Interviews"] == 0, 0.0, np.nan),
    )
    for metric in ("Unique AWCs Visited", "AWCs with At Least 2 Interviews", "AWCs with Both Mother Interviews"):
        result[f"Current {metric}"] = current_awc[metric]
        result[f"Previous {metric}"] = prior_awc[metric]
        result[f"Change in {metric}"] = current_awc[metric] - prior_awc[metric]
    return result.reset_index(drop=True).sort_values("Name of Field Investigators", key=lambda s: s.str.casefold())


def tool_comparison(data, latest, previous):
    rows = []
    for tool in TOOL_NAMES:
        tool_data = data.loc[(data["__Tool"] == tool) & data["__Date"].notna()]
        current = int((tool_data["__Date"] <= latest).sum())
        prior = int((tool_data["__Date"] <= previous).sum())
        change = current - prior
        rows.append({
            "Tool": tool,
            "Start Date Found": tool_data["__Date"].min() if not tool_data.empty else pd.NaT,
            f"Cumulative Through {latest:%d-%m-%Y}": current,
            f"Cumulative Through {previous:%d-%m-%Y}": prior,
            "Change": change,
            "% Change": change / prior if prior else (0.0 if current == 0 else np.nan),
        })
    return pd.DataFrame(rows)


def haversine(lat1, lon1, lat2, lon2):
    if any(pd.isna(v) for v in (lat1, lon1, lat2, lon2)):
        return np.nan
    radius = 6371.0088
    phi1, phi2 = math.radians(lat1), math.radians(lat2)
    dphi, dlambda = math.radians(lat2 - lat1), math.radians(lon2 - lon1)
    a = math.sin(dphi / 2) ** 2 + math.cos(phi1) * math.cos(phi2) * math.sin(dlambda / 2) ** 2
    return radius * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def weekly_travel(data):
    gps = data.dropna(subset=["__FIKey", "__Date", "__Latitude", "__Longitude"]).copy()
    if gps.empty:
        return pd.DataFrame()
    daily = gps.groupby(["__FIKey", "__Date"], as_index=False).agg(
        Latitude=("__Latitude", "median"), Longitude=("__Longitude", "median")
    ).sort_values(["__FIKey", "__Date"])
    daily["Previous Latitude"] = daily.groupby("__FIKey")["Latitude"].shift()
    daily["Previous Longitude"] = daily.groupby("__FIKey")["Longitude"].shift()
    daily["Distance km"] = daily.apply(
        lambda row: haversine(row["Previous Latitude"], row["Previous Longitude"], row["Latitude"], row["Longitude"]), axis=1
    )
    iso = daily["__Date"].dt.isocalendar()
    daily["ISO Year"], daily["ISO Week"] = iso.year.astype(int), iso.week.astype(int)
    daily["Week Start"] = daily["__Date"] - pd.to_timedelta(daily["__Date"].dt.weekday, unit="D")
    fi_week = daily.groupby(["__FIKey", "ISO Year", "ISO Week", "Week Start"], as_index=False).agg(
        Weekly_Travel_km=("Distance km", "sum"), Comparisons=("Distance km", "count")
    )
    fi_week = fi_week.loc[fi_week["Comparisons"] > 0].copy()
    if fi_week.empty:
        return pd.DataFrame()
    fi_week["Reached 10 km"] = fi_week["Weekly_Travel_km"] >= WEEKLY_TRAVEL_KM
    result = fi_week.groupby(["ISO Year", "ISO Week", "Week Start"], as_index=False).agg(
        **{
            "FIs with Valid GPS Comparisons": ("__FIKey", "nunique"),
            "FIs Travelled At Least 10 km": ("Reached 10 km", "sum"),
        }
    )
    result["Percent FIs Travelled >= 10 km"] = result["FIs Travelled At Least 10 km"] / result["FIs with Valid GPS Comparisons"]
    return result.sort_values("Week Start", ascending=False)


def safe(value):
    return re.sub(r'[<>:"/\\|?*]+', "_", str(value)).strip() or "Unknown_State"


def format_sheet(writer, name, df, title=None, date_columns=None, percent_columns=None, zero_columns=None):
    ws, wb = writer.sheets[name], writer.book
    header_row = 2 if title else 0
    header = wb.add_format({"bold": True, "bg_color": "#B7DEE8", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
    title_format = wb.add_format({"bold": True, "font_size": 14})
    date_format = wb.add_format({"num_format": "dd-mm-yyyy"})
    percent_format = wb.add_format({"num_format": "+0.0%;-0.0%;-"})
    zero_format = wb.add_format({"bg_color": "#F4CCCC"})
    if title:
        ws.write(0, 0, title, title_format)
    ws.set_row(header_row, 38, header)
    ws.freeze_panes(header_row + 1, 1)
    if len(df.columns):
        ws.autofilter(header_row, 0, header_row + len(df), len(df.columns) - 1)
    for index, column in enumerate(df.columns):
        lengths = df.head(100)[column].dropna().astype(str).str.len()
        width = min(max(len(str(column)) + 2, int(lengths.max()) + 2 if not lengths.empty else 10), 38)
        fmt = date_format if date_columns and column in date_columns else percent_format if percent_columns and column in percent_columns else None
        ws.set_column(index, index, width, fmt)
        if zero_columns and column in zero_columns and len(df):
            ws.conditional_format(header_row + 1, index, header_row + len(df), index, {"type": "cell", "criteria": "==", "value": 0, "format": zero_format})


def create_workbook(state, data, logs, output_dir):
    latest, previous = report_dates(data)
    status = cumulative_status(data, latest)
    comparison = cumulative_comparison(data, latest, previous)
    tool_summary = tool_comparison(data, latest, previous)
    travel = weekly_travel(data)
    logging.info("%s: cumulative through %s; previous through %s; FI roster=%s", state, latest.date(), previous.date(), len(status) - 1)

    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f"{safe(state).upper()}_IDD_Extension_Cumulative_AWC_Summary_{datetime.now():%d-%m-%Y}.xlsx"
    with pd.ExcelWriter(path, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        status.to_excel(writer, sheet_name="Cumulative FI and AWC", index=False, startrow=2)
        format_sheet(
            writer, "Cumulative FI and AWC", status,
            title=f"Cumulative FI submissions and AWC coverage through latest database date: {latest:%d-%m-%Y} | {state}",
            zero_columns=set(TOOL_NAMES),
        )
        comparison.to_excel(writer, sheet_name="Cumulative Comparison", index=False, startrow=2)
        format_sheet(
            writer, "Cumulative Comparison", comparison,
            title=f"Cumulative totals through {latest:%d-%m-%Y} compared with {previous:%d-%m-%Y} | {state}",
            percent_columns={"Interview % Change"},
        )
        tool_summary.to_excel(writer, sheet_name="Tool Cumulative Comparison", index=False, startrow=2)
        format_sheet(
            writer, "Tool Cumulative Comparison", tool_summary,
            title=f"Tool cumulative totals through {latest:%d-%m-%Y} compared with {previous:%d-%m-%Y} | {state}",
            date_columns={"Start Date Found"}, percent_columns={"% Change"},
        )
        travel.to_excel(writer, sheet_name="Weekly 10km Travel", index=False)
        format_sheet(writer, "Weekly 10km Travel", travel, date_columns={"Week Start"}, percent_columns={"Percent FIs Travelled >= 10 km"})
        process = logs.assign(State_Workbook=state, Latest_Database_Date=latest, Previous_Date=previous, FI_Roster_Count=len(status) - 1)
        process.to_excel(writer, sheet_name="Processing Log", index=False)
        format_sheet(writer, "Processing Log", process, date_columns={"Start Date Found", "Latest Date Found", "Latest_Database_Date", "Previous_Date"})
        for tool in TOOL_NAMES:
            raw = data.loc[data["__Tool"] == tool].drop(columns=[c for c in data.columns if str(c).startswith("__")], errors="ignore")
            name = re.sub(r"[\\/*?:\[\]]", "_", f"Raw {tool}")[:31]
            raw.to_excel(writer, sheet_name=name, index=False)
            format_sheet(writer, name, raw)
    logging.info("Created workbook: %s", path)
    return path


def make_zip(paths, output_dir):
    path = output_dir / f"IDD_Extension_Cumulative_AWC_Summary_All_States_{datetime.now():%d-%m-%Y}.zip"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for item in paths:
            archive.write(item, arcname=item.name)
    return path


def send_email(attachment, state_names):
    sender = env("SMTP_EMAIL", "GMAIL_USERNAME", required=True)
    password = env("SMTP_APP_PASSWORD", "GMAIL_APP_PASSWORD", required=True)
    recipients = [x.strip() for x in re.split(r"[,;]", env("RECIPIENTS", "REPORT_RECIPIENTS", required=True)) if x.strip()]
    size_mb = attachment.stat().st_size / (1024 * 1024)
    limit_mb = float(env("MAX_EMAIL_ATTACHMENT_MB") or "24")
    if size_mb > limit_mb:
        raise ValueError(f"ZIP is {size_mb:.2f} MB, above {limit_mb:.2f} MB")
    message = EmailMessage()
    message["From"], message["To"] = sender, ", ".join(recipients)
    message["Subject"] = env("MAIL_SUBJECT") or f"IDD Extension Cumulative AWC Summary - {datetime.now():%d-%m-%Y}"
    message.set_content(f"Dear Team,\n\nPlease find attached the cumulative IDD Extension FI and AWC summary for: {', '.join(state_names)}.\n\nRegards,\nGAVB Reporting Automation")
    with attachment.open("rb") as handle:
        message.add_attachment(handle.read(), maintype="application", subtype="zip", filename=attachment.name)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context(), timeout=60) as smtp:
        smtp.login(sender, password)
        smtp.send_message(message)
    logging.info("Email sent successfully")


def main():
    logging.basicConfig(level=(env("LOG_LEVEL") or "INFO").upper(), format="%(asctime)s [%(levelname)s] %(message)s")
    try:
        logging.info("Running report script version: %s", SCRIPT_VERSION)
        output_dir = Path(env("LOCAL_OUTPUT_FOLDER") or f"/tmp/{datetime.now(timezone.utc):%Y-%m-%d}/idd_extension_cumulative_awc")
        service = build("drive", "v3", credentials=credentials(), cache_discovery=False)
        frames, logs = load_all(service, env("DRIVE_FOLDER_ID", required=True))
        reports, completed = [], []
        for state in get_states(frames):
            data = state_data(frames, state)
            if not data.empty:
                reports.append(create_workbook(state, data, logs, output_dir))
                completed.append(state)
        if not reports:
            raise ValueError("No reports created")
        zip_path = make_zip(reports, output_dir)
        send_email(zip_path, completed)
        print("Created reports:")
        for report in reports:
            print(" -", report)
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
