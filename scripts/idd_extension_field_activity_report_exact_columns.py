#!/usr/bin/env python3
"""IDD Extension submission summary for Mother 0-5, Mother 6-11, and Separate CD."""

import io, json, logging, math, os, re, smtplib, ssl, sys, zipfile
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path

import numpy as np
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

SCRIPT_VERSION = "2026-09-25-idd-submission-summary-v4"
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
WEEKLY_TRAVEL_KM = float(os.getenv("WEEKLY_TRAVEL_KM", "10"))

TOOL_FILES = [
    ("Mother 0-5", "IDD_Extension_Mother_0_5_WIDE.xlsx"),
    ("Mother 6-11", "IDD Extension Mother 6-11_WIDE.xlsx"),
    ("Separate CD", "IDD_Extension_Separate_CD_Tool_WIDE.xlsx"),
]

ALIASES = {
    "state": ["STATE"],
    "fi": ["QDC_Name"],
    "submission": ["SubmissionDate", "Submission Date", "submission_date", "SubmissionDateTime", "Submission_Time", "endtime", "EndTime", "starttime", "StartTime"],
    "awc_name": ["QAWC_name"],
    "awc_code": ["Cal_AWC", "QAWC_code", "QAWC_3"],
    "district": ["DIST"],
    "block": ["BLOCK"],
}

GPS_ALIASES = {
    "Mother 0-5": (["Qgeo-Latitude"], ["Qgeo-Longitude"]),
    "Mother 6-11": (["QX1-Latitude"], ["QX1-Longitude"]),
    "Separate CD": (["Qgeo-Latitude"], ["Qgeo-Longitude"]),
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


def clean(series):
    return series.astype("string").str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA})


def find_column(columns, aliases):
    lookup = {norm(column): str(column) for column in columns}
    for alias in aliases:
        if norm(alias) in lookup:
            return lookup[norm(alias)]
    return None


def credentials():
    raw = env("GOOGLE_SERVICE_ACCOUNT_JSON")
    credential_file = env("GOOGLE_SERVICE_ACCOUNT_FILE") or "credential.json"
    if raw:
        return Credentials.from_service_account_info(json.loads(raw), scopes=DRIVE_SCOPES)
    if os.path.exists(credential_file):
        return Credentials.from_service_account_file(credential_file, scopes=DRIVE_SCOPES)
    raise ValueError("Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_SERVICE_ACCOUNT_FILE")


def escape_drive_value(value):
    return value.replace("\\", "\\\\").replace("'", "\\'")


def find_file(service, folder_id, filename):
    query = f"'{folder_id}' in parents and trashed = false and name = '{escape_drive_value(filename)}'"
    files = service.files().list(q=query, spaces="drive", fields="files(id,name)", pageSize=10).execute().get("files", [])
    if not files:
        raise FileNotFoundError(f"File not found in Drive folder: {filename}")
    return files[0]["id"]


def download(service, file_id):
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, service.files().get_media(fileId=file_id))
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buffer.getvalue()


def detect_header_and_read(data, filename):
    preview = pd.read_excel(io.BytesIO(data), sheet_name=0, header=None, nrows=20, engine="openpyxl")
    targets = {norm("STATE"), norm("QDC_Name")}
    scores = []
    for row_number in range(len(preview)):
        values = {norm(value) for value in preview.iloc[row_number].tolist() if pd.notna(value)}
        scores.append((len(targets & values), row_number))
    header_row = max(scores)[1] if scores else 0
    frame = pd.read_excel(io.BytesIO(data), sheet_name=0, header=header_row, engine="openpyxl")
    frame.columns = [str(column).strip() for column in frame.columns]
    logging.info("%s: detected header row %s with %s rows x %s columns", filename, header_row + 1, len(frame), len(frame.columns))
    return frame, header_row + 1


def parse_dates(series):
    text = clean(series)
    parsed = pd.to_datetime(text, format="ISO8601", errors="coerce")
    unresolved = parsed.isna() & text.notna()
    if unresolved.any():
        try:
            parsed.loc[unresolved] = pd.to_datetime(text.loc[unresolved], format="mixed", dayfirst=True, errors="coerce")
        except (TypeError, ValueError):
            parsed.loc[unresolved] = pd.to_datetime(text.loc[unresolved], dayfirst=True, errors="coerce")
    return parsed


def coordinate(series, latitude):
    values = pd.to_numeric(series, errors="coerce")
    return values.where(values.between(-90, 90) if latitude else values.between(-180, 180))


def standardize(raw, tool, filename, header_row):
    detected = {key: find_column(raw.columns, aliases) for key, aliases in ALIASES.items()}
    lat_col = find_column(raw.columns, GPS_ALIASES[tool][0])
    lon_col = find_column(raw.columns, GPS_ALIASES[tool][1])
    missing = [key for key in ("state", "fi", "submission") if not detected[key]]
    if missing:
        raise ValueError(f"{filename}: missing required column(s): {', '.join(missing)}")
    if not lat_col or not lon_col:
        raise ValueError(f"{filename}: missing GPS columns. Expected {GPS_ALIASES[tool][0][0]} and {GPS_ALIASES[tool][1][0]}")

    frame = raw.copy()
    frame["__Tool"] = tool
    frame["__State"] = clean(frame[detected["state"]])
    frame["__FI"] = clean(frame[detected["fi"]])
    frame["__DateTime"] = parse_dates(frame[detected["submission"]])
    frame["__Date"] = frame["__DateTime"].dt.normalize()
    frame["__Latitude"] = coordinate(frame[lat_col], True)
    frame["__Longitude"] = coordinate(frame[lon_col], False)
    frame["__AWC"] = clean(frame[detected["awc_code"]]) if detected["awc_code"] else (clean(frame[detected["awc_name"]]) if detected["awc_name"] else pd.NA)

    log = {
        "Tool": tool, "File": filename, "Header Row": header_row, "Rows Read": len(frame),
        "State Column": detected["state"], "FI Column": detected["fi"],
        "Submission Column": detected["submission"], "Latitude Column": lat_col,
        "Longitude Column": lon_col, "Valid Dates": int(frame["__Date"].notna().sum()),
        "Valid GPS": int((frame["__Latitude"].notna() & frame["__Longitude"].notna()).sum()), "Status": "OK",
    }
    return frame, log


def load_all(service, folder_id):
    frames, logs = {}, []
    for tool, filename in TOOL_FILES:
        logging.info("Reading %s", filename)
        try:
            raw, header_row = detect_header_and_read(download(service, find_file(service, folder_id, filename)), filename)
            frames[tool], log = standardize(raw, tool, filename, header_row)
            logs.append(log)
        except Exception as exc:
            logging.exception("Could not process %s", filename)
            logs.append({"Tool": tool, "File": filename, "Rows Read": 0, "Status": f"ERROR: {exc}"})
    missing_tools = sorted(set(tool for tool, _ in TOOL_FILES) - set(frames))
    if missing_tools:
        raise ValueError("Report stopped because these tools could not be processed: " + ", ".join(missing_tools))
    return frames, pd.DataFrame(logs)


def normalize_state(value):
    return str(value).strip().casefold()


def states_from_frames(frames):
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


def state_data(frames, state):
    parts = []
    for frame in frames.values():
        mask = frame["__State"].fillna("").str.strip().str.casefold() == normalize_state(state)
        parts.append(frame.loc[mask].copy())
    return pd.concat(parts, ignore_index=True, sort=False)


def latest_fi_tool_table(data):
    valid = data.dropna(subset=["__FI", "__Date"])
    if valid.empty:
        return pd.DataFrame()
    latest_date = valid["__Date"].max()
    latest = valid.loc[valid["__Date"] == latest_date]
    table = latest.pivot_table(index="__FI", columns="__Tool", values="__State", aggfunc="size", fill_value=0).reset_index()
    table = table.rename(columns={"__FI": "Name of Field Investigators"})
    for tool, _ in TOOL_FILES:
        if tool not in table.columns:
            table[tool] = 0
    table["Total Submissions"] = table[[tool for tool, _ in TOOL_FILES]].sum(axis=1)
    total = {"Name of Field Investigators": "Total Submissions", **{tool: int(table[tool].sum()) for tool, _ in TOOL_FILES}, "Total Submissions": int(table["Total Submissions"].sum())}
    table = pd.concat([table, pd.DataFrame([total])], ignore_index=True)
    return table, latest_date


def latest_vs_previous_tool_table(data):
    valid = data.dropna(subset=["__Date"])
    dates = sorted(valid["__Date"].unique())
    if not dates:
        return pd.DataFrame(), pd.NaT, pd.NaT
    latest_date = dates[-1]
    previous_date = dates[-2] if len(dates) > 1 else pd.NaT
    rows = []
    for tool, _ in TOOL_FILES:
        latest_count = int(((valid["__Date"] == latest_date) & (valid["__Tool"] == tool)).sum())
        previous_count = int(((valid["__Date"] == previous_date) & (valid["__Tool"] == tool)).sum()) if pd.notna(previous_date) else 0
        change = latest_count - previous_count
        change_pct = (change / previous_count) if previous_count else (0.0 if latest_count == 0 else np.nan)
        rows.append({
            "Tool": tool, "Latest Date Count": latest_count, "Previous Active Date Count": previous_count,
            "Change": change, "% Change": change_pct,
        })
    return pd.DataFrame(rows), latest_date, previous_date


def haversine(lat1, lon1, lat2, lon2):
    if any(pd.isna(v) for v in (lat1, lon1, lat2, lon2)):
        return np.nan
    radius = 6371.0088
    p1, p2 = math.radians(lat1), math.radians(lat2)
    dp, dl = math.radians(lat2 - lat1), math.radians(lon2 - lon1)
    a = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return radius * 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))


def weekly_travel_summary(data):
    gps = data.dropna(subset=["__FI", "__Date", "__Latitude", "__Longitude"]).copy()
    if gps.empty:
        return pd.DataFrame()
    daily = gps.groupby(["__FI", "__Date"], as_index=False).agg(Latitude=("__Latitude", "median"), Longitude=("__Longitude", "median"))
    daily = daily.sort_values(["__FI", "__Date"])
    daily["Previous Latitude"] = daily.groupby("__FI")["Latitude"].shift()
    daily["Previous Longitude"] = daily.groupby("__FI")["Longitude"].shift()
    daily["Distance km"] = daily.apply(lambda row: haversine(row["Previous Latitude"], row["Previous Longitude"], row["Latitude"], row["Longitude"]), axis=1)
    iso = daily["__Date"].dt.isocalendar()
    daily["ISO Year"] = iso.year.astype(int)
    daily["ISO Week"] = iso.week.astype(int)
    daily["Week Start"] = daily["__Date"] - pd.to_timedelta(daily["__Date"].dt.weekday, unit="D")
    fi_week = daily.groupby(["__FI", "ISO Year", "ISO Week", "Week Start"], as_index=False).agg(
        Weekly_Travel_km=("Distance km", "sum"), GPS_Active_Days=("__Date", "nunique"), Valid_Distance_Comparisons=("Distance km", "count")
    )
    fi_week["Travelled >= 10 km"] = np.where(fi_week["Valid_Distance_Comparisons"] > 0, fi_week["Weekly_Travel_km"] >= WEEKLY_TRAVEL_KM, False)
    summary = fi_week.groupby(["ISO Year", "ISO Week", "Week Start"], as_index=False).agg(
        FIs_with_Valid_GPS_Comparisons=("__FI", "nunique"),
        FIs_Travelled_At_Least_10_km=("Travelled >= 10 km", "sum"),
    )
    summary["Percent FIs Travelled >= 10 km"] = np.where(summary["FIs_with_Valid_GPS_Comparisons"] > 0, summary["FIs_Travelled_At_Least_10_km"] / summary["FIs_with_Valid_GPS_Comparisons"], np.nan)
    return summary.sort_values("Week Start", ascending=False)


def safe_filename(value):
    return re.sub(r'[<>:"/\\|?*]+', "_", str(value)).strip() or "Unknown_State"


def format_sheet(writer, sheet_name, df, title=None, date_columns=None, percent_columns=None, zero_highlight_columns=None):
    ws = writer.sheets[sheet_name]
    wb = writer.book
    header_row = 2 if title else 0
    header = wb.add_format({"bold": True, "bg_color": "#B7DEE8", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
    title_fmt = wb.add_format({"bold": True, "font_size": 14})
    date_fmt = wb.add_format({"num_format": "dd-mm-yyyy"})
    percent_fmt = wb.add_format({"num_format": "+0.0%;-0.0%;-"})
    zero_fmt = wb.add_format({"bg_color": "#F4CCCC"})
    if title:
        ws.write(0, 0, title, title_fmt)
    ws.set_row(header_row, 36, header)
    ws.freeze_panes(header_row + 1, 1)
    if len(df.columns):
        ws.autofilter(header_row, 0, header_row + len(df), len(df.columns) - 1)
    for index, column in enumerate(df.columns):
        lengths = df.head(100)[column].dropna().astype(str).str.len()
        width = min(max(len(str(column)) + 2, int(lengths.max()) + 2 if not lengths.empty else 10), 34)
        column_format = date_fmt if date_columns and column in date_columns else percent_fmt if percent_columns and column in percent_columns else None
        ws.set_column(index, index, width, column_format)
        if zero_highlight_columns and column in zero_highlight_columns and len(df):
            ws.conditional_format(header_row + 1, index, header_row + len(df), index, {"type": "cell", "criteria": "==", "value": 0, "format": zero_fmt})


def create_workbook(state, data, logs, output_dir):
    fi_table, latest_date = latest_fi_tool_table(data)
    comparison, comparison_latest, previous_date = latest_vs_previous_tool_table(data)
    weekly = weekly_travel_summary(data)
    if fi_table.empty:
        raise ValueError(f"{state}: no valid FI/date records")

    report_date = datetime.now().strftime("%d-%m-%Y")
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f"{safe_filename(state).upper()}_IDD_Extension_Submission_Summary_{report_date}.xlsx"
    with pd.ExcelWriter(path, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        fi_table.to_excel(writer, sheet_name="Latest FI Tool Status", index=False, startrow=2)
        format_sheet(writer, "Latest FI Tool Status", fi_table,
                     title=f"Total Submissions status as on: {latest_date:%d-%m-%Y} | {state}",
                     zero_highlight_columns={tool for tool, _ in TOOL_FILES})

        comparison.to_excel(writer, sheet_name="Latest vs Previous", index=False, startrow=2)
        previous_text = f"{previous_date:%d-%m-%Y}" if pd.notna(previous_date) else "Not available"
        format_sheet(writer, "Latest vs Previous", comparison,
                     title=f"Tool submission comparison: {comparison_latest:%d-%m-%Y} vs {previous_text} | {state}",
                     percent_columns={"% Change"}, zero_highlight_columns={"Latest Date Count"})

        weekly.to_excel(writer, sheet_name="Weekly 10km Travel", index=False)
        format_sheet(writer, "Weekly 10km Travel", weekly, date_columns={"Week Start"}, percent_columns={"Percent FIs Travelled >= 10 km"})

        logs.assign(State_Workbook=state).to_excel(writer, sheet_name="Processing Log", index=False)
        format_sheet(writer, "Processing Log", logs.assign(State_Workbook=state))

        for tool, _ in TOOL_FILES:
            raw = data.loc[data["__Tool"] == tool].drop(columns=[c for c in data.columns if str(c).startswith("__")], errors="ignore")
            sheet = re.sub(r"[\\/*?:\[\]]", "_", f"Raw {tool}")[:31]
            raw.to_excel(writer, sheet_name=sheet, index=False)
            format_sheet(writer, sheet, raw)
    logging.info("Created state workbook: %s", path)
    return path


def create_zip(paths, output_dir):
    path = output_dir / f"IDD_Extension_Submission_Summary_All_States_{datetime.now():%d-%m-%Y}.zip"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for report in paths:
            archive.write(report, arcname=report.name)
    return path


def send_email(attachment, states):
    sender = env("SMTP_EMAIL", "GMAIL_USERNAME", "EMAIL_USERNAME", required=True)
    password = env("SMTP_APP_PASSWORD", "GMAIL_APP_PASSWORD", "EMAIL_PASSWORD", required=True)
    recipients = [x.strip() for x in re.split(r"[,;]", env("RECIPIENTS", "REPORT_RECIPIENTS", "EMAIL_TO", required=True)) if x.strip()]
    size_mb = attachment.stat().st_size / (1024 * 1024)
    maximum = float(env("MAX_EMAIL_ATTACHMENT_MB") or "24")
    if size_mb > maximum:
        raise ValueError(f"ZIP attachment is {size_mb:.2f} MB, above configured limit {maximum:.2f} MB")
    message = EmailMessage()
    message["From"], message["To"] = sender, ", ".join(recipients)
    message["Subject"] = env("MAIL_SUBJECT", "EMAIL_SUBJECT") or f"IDD Extension Submission Summary - {datetime.now():%d-%m-%Y}"
    message.set_content(env("MAIL_BODY", "EMAIL_BODY") or f"Dear Team,\n\nPlease find attached the IDD Extension submission summary for: {', '.join(states)}.\n\nRegards,\nGAVB Reporting Automation")
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
        output_dir = Path(env("LOCAL_OUTPUT_FOLDER") or f"/tmp/{datetime.now(timezone.utc):%Y-%m-%d}/idd_extension_summary")
        service = build("drive", "v3", credentials=credentials(), cache_discovery=False)
        frames, logs = load_all(service, folder_id)
        states = states_from_frames(frames)
        reports, completed = [], []
        for state in states:
            data = state_data(frames, state)
            if not data.empty:
                reports.append(create_workbook(state, data, logs, output_dir))
                completed.append(state)
        if not reports:
            raise ValueError("No state workbook was created")
        zip_path = create_zip(reports, output_dir)
        send_email(zip_path, completed)
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
