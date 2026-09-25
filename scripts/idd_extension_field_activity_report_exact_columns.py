#!/usr/bin/env python3
"""All-state cumulative IDD Extension reports with latest-vs-previous cumulative comparison."""
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

SCRIPT_VERSION = "2026-09-25-idd-cumulative-latest-vs-previous-v7"
SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]
TOOLS = [
    ("Mother 0-5", "IDD_Extension_Mother_0_5_WIDE.xlsx", "Qgeo-Latitude", "Qgeo-Longitude"),
    ("Mother 6-11", "IDD Extension Mother 6-11_WIDE.xlsx", "QX1-Latitude", "QX1-Longitude"),
    ("Separate CD", "IDD_Extension_Separate_CD_Tool_WIDE.xlsx", "Qgeo-Latitude", "Qgeo-Longitude"),
]
TOOL_NAMES = [x[0] for x in TOOLS]
ALIASES = {
    "state": ["STATE", "State", "state", "Cal_STATE", "state_name"],
    "fi": ["QDC_Name", "QDC Name", "Field Investigator Name", "Investigator", "QDC", "collector_name"],
    "date": ["SubmissionDate", "Submission Date", "submission_date", "SubmissionDateTime", "Submission_Time", "endtime", "EndTime", "starttime", "StartTime"],
}
TRAVEL_KM = float(os.getenv("WEEKLY_TRAVEL_KM", "10"))


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
    return series.astype("string").str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA})


def find_col(columns, aliases):
    lookup = {norm(c): str(c) for c in columns}
    return next((lookup[norm(a)] for a in aliases if norm(a) in lookup), None)


def creds():
    raw = env("GOOGLE_SERVICE_ACCOUNT_JSON")
    path = env("GOOGLE_SERVICE_ACCOUNT_FILE") or "credential.json"
    if raw:
        return Credentials.from_service_account_info(json.loads(raw), scopes=SCOPES)
    if os.path.exists(path):
        return Credentials.from_service_account_file(path, scopes=SCOPES)
    raise ValueError("Google service account credentials not found")


def drive_id(service, folder_id, filename):
    safe = filename.replace("\\", "\\\\").replace("'", "\\'")
    query = f"'{folder_id}' in parents and trashed = false and name = '{safe}'"
    files = service.files().list(q=query, spaces="drive", fields="files(id,name)", pageSize=10).execute().get("files", [])
    if not files:
        raise FileNotFoundError(filename)
    return files[0]["id"]


def download(service, file_id):
    buffer = io.BytesIO()
    job = MediaIoBaseDownload(buffer, service.files().get_media(fileId=file_id))
    done = False
    while not done:
        _, done = job.next_chunk()
    return buffer.getvalue()


def read_excel(data, filename):
    preview = pd.read_excel(io.BytesIO(data), sheet_name=0, header=None, nrows=25, engine="openpyxl")
    targets = {norm("STATE"), norm("QDC_Name")}
    best = max(
        ((len({norm(v) for v in preview.iloc[i].tolist() if pd.notna(v)} & targets), i) for i in range(len(preview))),
        default=(0, 0),
    )[1]
    df = pd.read_excel(io.BytesIO(data), sheet_name=0, header=best, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    logging.info("%s: header row %s, %s rows x %s columns", filename, best + 1, len(df), len(df.columns))
    return df, best + 1


def parse_dates(series):
    text = clean(series)
    parsed = pd.to_datetime(text, format="ISO8601", errors="coerce")
    missing = parsed.isna() & text.notna()
    if missing.any():
        try:
            parsed.loc[missing] = pd.to_datetime(text.loc[missing], format="mixed", dayfirst=True, errors="coerce")
        except (TypeError, ValueError):
            parsed.loc[missing] = pd.to_datetime(text.loc[missing], dayfirst=True, errors="coerce")
    return parsed


def standardize(raw, tool, filename, lat_name, lon_name, header_row):
    state_col = find_col(raw.columns, ALIASES["state"])
    fi_col = find_col(raw.columns, ALIASES["fi"])
    date_col = find_col(raw.columns, ALIASES["date"])
    lat_col = find_col(raw.columns, [lat_name])
    lon_col = find_col(raw.columns, [lon_name])
    missing = [label for label, col in [("STATE", state_col), ("QDC_Name", fi_col), ("date", date_col), (lat_name, lat_col), (lon_name, lon_col)] if not col]
    if missing:
        raise ValueError(f"{filename}: missing {', '.join(missing)}")
    df = raw.copy()
    df["__Tool"] = tool
    df["__State"] = clean(df[state_col])
    df["__FI"] = clean(df[fi_col])
    df["__FIKey"] = df["__FI"].str.casefold()
    df["__DateTime"] = parse_dates(df[date_col])
    df["__Date"] = df["__DateTime"].dt.normalize()
    df["__Latitude"] = pd.to_numeric(df[lat_col], errors="coerce").where(lambda s: s.between(-90, 90))
    df["__Longitude"] = pd.to_numeric(df[lon_col], errors="coerce").where(lambda s: s.between(-180, 180))
    return df, {
        "Tool": tool, "File": filename, "Header Row": header_row, "Rows Read": len(raw),
        "State Column": state_col, "FI Column": fi_col, "Submission Column": date_col,
        "Latitude Column": lat_col, "Longitude Column": lon_col,
        "Unique FIs": int(df["__FIKey"].nunique()), "Valid Dates": int(df["__Date"].notna().sum()),
        "Start Date Found": df["__Date"].min(), "Latest Date Found": df["__Date"].max(),
        "Status": "OK",
    }


def load_all(service, folder_id):
    frames, logs = {}, []
    for tool, filename, lat, lon in TOOLS:
        logging.info("Reading %s", filename)
        try:
            raw, header = read_excel(download(service, drive_id(service, folder_id, filename)), filename)
            frames[tool], log = standardize(raw, tool, filename, lat, lon, header)
            logs.append(log)
        except Exception as exc:
            logging.exception("Could not process %s", filename)
            logs.append({"Tool": tool, "File": filename, "Rows Read": 0, "Status": f"ERROR: {exc}"})
    missing = sorted(set(TOOL_NAMES) - set(frames))
    if missing:
        raise ValueError("Stopped because tools failed: " + ", ".join(missing))
    return frames, pd.DataFrame(logs)


def states(frames):
    groups = {}
    for df in frames.values():
        for value in df["__State"].dropna().unique():
            display = str(value).strip()
            if display:
                groups.setdefault(display.casefold(), []).append(display)
    values = [max(v, key=lambda x: (v.count(x), len(x))) for v in groups.values()]
    selected = {x.strip().casefold() for x in env("STATES").split(",") if x.strip()}
    excluded = {x.strip().casefold() for x in env("EXCLUDED_STATES").split(",") if x.strip()}
    if selected:
        values = [x for x in values if x.casefold() in selected]
    return sorted([x for x in values if x.casefold() not in excluded], key=str.casefold)


def state_frame(frames, state):
    return pd.concat([
        df.loc[df["__State"].fillna("").str.strip().str.casefold() == state.casefold()].copy()
        for df in frames.values()
    ], ignore_index=True, sort=False)


def roster(data):
    valid = data.dropna(subset=["__FIKey", "__FI"])
    return (
        valid.groupby(["__FIKey", "__FI"]).size().rename("n").reset_index()
        .sort_values(["__FIKey", "n", "__FI"], ascending=[True, False, True])
        .drop_duplicates("__FIKey")[["__FIKey", "__FI"]]
        .rename(columns={"__FI": "Name of Field Investigators"})
    )


def report_dates(data):
    dates = data["__Date"].dropna()
    if dates.empty:
        raise ValueError("No valid dates")
    latest = dates.max()
    return latest, latest - pd.Timedelta(days=1)


def cumulative_matrix(data, through_date, fi_roster):
    selected = data.loc[data["__Date"].notna() & (data["__Date"] <= through_date)]
    counts = selected.groupby(["__FIKey", "__Tool"]).size().unstack(fill_value=0)
    counts = counts.reindex(index=fi_roster["__FIKey"], columns=TOOL_NAMES, fill_value=0)
    counts.index.name = "__FIKey"
    result = fi_roster.merge(counts.reset_index(), on="__FIKey", how="left")
    for tool in TOOL_NAMES:
        result[tool] = result[tool].fillna(0).astype(int)
    result["Total Submissions"] = result[TOOL_NAMES].sum(axis=1).astype(int)
    return result


def cumulative_fi_status(data, latest):
    r = cumulative_matrix(data, latest, roster(data)).drop(columns="__FIKey")
    r = r.sort_values("Name of Field Investigators", key=lambda s: s.str.casefold()).reset_index(drop=True)
    total = {"Name of Field Investigators": "Total Submissions", **{t: int(r[t].sum()) for t in TOOL_NAMES}, "Total Submissions": int(r["Total Submissions"].sum())}
    return pd.concat([r, pd.DataFrame([total])], ignore_index=True)


def cumulative_fi_change(data, latest, previous):
    people = roster(data)
    current = cumulative_matrix(data, latest, people).set_index("__FIKey")
    prior = cumulative_matrix(data, previous, people).set_index("__FIKey")
    out = people.set_index("__FIKey")
    for tool in TOOL_NAMES:
        out[f"Current {tool}"] = current[tool]
        out[f"Previous {tool}"] = prior[tool]
    out["Current Total"] = current["Total Submissions"]
    out["Previous Total"] = prior["Total Submissions"]
    out["Change"] = out["Current Total"] - out["Previous Total"]
    out["% Change"] = np.where(out["Previous Total"] > 0, out["Change"] / out["Previous Total"], np.where(out["Current Total"] == 0, 0.0, np.nan))
    return out.reset_index(drop=True).sort_values("Name of Field Investigators", key=lambda s: s.str.casefold())


def cumulative_tool_change(data, latest, previous):
    rows = []
    for tool in TOOL_NAMES:
        d = data.loc[(data["__Tool"] == tool) & data["__Date"].notna()]
        start = d["__Date"].min() if not d.empty else pd.NaT
        current = int((d["__Date"] <= latest).sum())
        prior = int((d["__Date"] <= previous).sum())
        change = current - prior
        rows.append({
            "Tool": tool, "Start Date Found": start,
            f"Cumulative Through {latest:%d-%m-%Y}": current,
            f"Cumulative Through {previous:%d-%m-%Y}": prior,
            "Change": change, "% Change": change / prior if prior else (0.0 if current == 0 else np.nan),
        })
    return pd.DataFrame(rows)


def haversine(a, b, c, d):
    if any(pd.isna(x) for x in (a, b, c, d)):
        return np.nan
    r = 6371.0088
    p1, p2, dp, dl = map(math.radians, (a, c, c - a, d - b))
    x = math.sin(dp / 2) ** 2 + math.cos(p1) * math.cos(p2) * math.sin(dl / 2) ** 2
    return r * 2 * math.atan2(math.sqrt(x), math.sqrt(1 - x))


def weekly_travel(data):
    gps = data.dropna(subset=["__FIKey", "__Date", "__Latitude", "__Longitude"]).copy()
    if gps.empty:
        return pd.DataFrame()
    daily = gps.groupby(["__FIKey", "__Date"], as_index=False).agg(Lat=("__Latitude", "median"), Lon=("__Longitude", "median")).sort_values(["__FIKey", "__Date"])
    daily["PrevLat"] = daily.groupby("__FIKey")["Lat"].shift()
    daily["PrevLon"] = daily.groupby("__FIKey")["Lon"].shift()
    daily["Distance km"] = daily.apply(lambda x: haversine(x.PrevLat, x.PrevLon, x.Lat, x.Lon), axis=1)
    iso = daily["__Date"].dt.isocalendar()
    daily["ISO Year"], daily["ISO Week"] = iso.year.astype(int), iso.week.astype(int)
    daily["Week Start"] = daily["__Date"] - pd.to_timedelta(daily["__Date"].dt.weekday, unit="D")
    fi = daily.groupby(["__FIKey", "ISO Year", "ISO Week", "Week Start"], as_index=False).agg(Weekly_km=("Distance km", "sum"), Comparisons=("Distance km", "count"))
    fi = fi.loc[fi.Comparisons > 0].copy()
    if fi.empty:
        return pd.DataFrame()
    fi["Reached"] = fi.Weekly_km >= TRAVEL_KM
    out = fi.groupby(["ISO Year", "ISO Week", "Week Start"], as_index=False).agg(**{"FIs with Valid GPS Comparisons": ("__FIKey", "nunique"), "FIs Travelled At Least 10 km": ("Reached", "sum")})
    out["Percent FIs Travelled >= 10 km"] = out["FIs Travelled At Least 10 km"] / out["FIs with Valid GPS Comparisons"]
    return out.sort_values("Week Start", ascending=False)


def safe(value):
    return re.sub(r'[<>:"/\\|?*]+', "_", str(value)).strip() or "Unknown_State"


def format_sheet(writer, name, df, title=None, dates=None, percents=None, zeros=None):
    ws, wb = writer.sheets[name], writer.book
    row = 2 if title else 0
    hf = wb.add_format({"bold": True, "bg_color": "#B7DEE8", "border": 1, "align": "center", "valign": "vcenter", "text_wrap": True})
    tf = wb.add_format({"bold": True, "font_size": 14})
    dfmt, pfmt, zfmt = wb.add_format({"num_format": "dd-mm-yyyy"}), wb.add_format({"num_format": "+0.0%;-0.0%;-"}), wb.add_format({"bg_color": "#F4CCCC"})
    if title:
        ws.write(0, 0, title, tf)
    ws.set_row(row, 38, hf)
    ws.freeze_panes(row + 1, 1)
    if len(df.columns):
        ws.autofilter(row, 0, row + len(df), len(df.columns) - 1)
    for i, col in enumerate(df.columns):
        lengths = df.head(100)[col].dropna().astype(str).str.len()
        width = min(max(len(str(col)) + 2, int(lengths.max()) + 2 if not lengths.empty else 10), 35)
        fmt = dfmt if dates and col in dates else pfmt if percents and col in percents else None
        ws.set_column(i, i, width, fmt)
        if zeros and col in zeros and len(df):
            ws.conditional_format(row + 1, i, row + len(df), i, {"type": "cell", "criteria": "==", "value": 0, "format": zfmt})


def workbook(state, data, logs, output):
    latest, previous = report_dates(data)
    status = cumulative_fi_status(data, latest)
    fi_change = cumulative_fi_change(data, latest, previous)
    tool_change = cumulative_tool_change(data, latest, previous)
    travel = weekly_travel(data)
    logging.info("%s: latest database date=%s, previous date=%s, FI roster=%s", state, latest.date(), previous.date(), len(status) - 1)
    output.mkdir(parents=True, exist_ok=True)
    path = output / f"{safe(state).upper()}_IDD_Extension_Cumulative_Summary_{datetime.now():%d-%m-%Y}.xlsx"
    with pd.ExcelWriter(path, engine="xlsxwriter", engine_kwargs={"options": {"strings_to_urls": False}}) as w:
        status.to_excel(w, sheet_name="Cumulative FI Status", index=False, startrow=2)
        format_sheet(w, "Cumulative FI Status", status, f"Cumulative submissions through latest database date: {latest:%d-%m-%Y} | {state}", zeros=set(TOOL_NAMES))
        fi_change.to_excel(w, sheet_name="FI Cumulative Comparison", index=False, startrow=2)
        format_sheet(w, "FI Cumulative Comparison", fi_change, f"FI cumulative totals through {latest:%d-%m-%Y} vs {previous:%d-%m-%Y} | {state}", percents={"% Change"})
        tool_change.to_excel(w, sheet_name="Tool Cumulative Comparison", index=False, startrow=2)
        format_sheet(w, "Tool Cumulative Comparison", tool_change, f"Tool cumulative totals through {latest:%d-%m-%Y} vs {previous:%d-%m-%Y} | {state}", dates={"Start Date Found"}, percents={"% Change"})
        travel.to_excel(w, sheet_name="Weekly 10km Travel", index=False)
        format_sheet(w, "Weekly 10km Travel", travel, dates={"Week Start"}, percents={"Percent FIs Travelled >= 10 km"})
        log = logs.assign(State_Workbook=state, Latest_Database_Date=latest, Previous_Date=previous, FI_Roster_Count=len(status) - 1)
        log.to_excel(w, sheet_name="Processing Log", index=False)
        format_sheet(w, "Processing Log", log, dates={"Start Date Found", "Latest Date Found", "Latest_Database_Date", "Previous_Date"})
        for tool in TOOL_NAMES:
            raw = data.loc[data.__Tool == tool].drop(columns=[c for c in data.columns if str(c).startswith("__")], errors="ignore")
            name = re.sub(r"[\\/*?:\[\]]", "_", f"Raw {tool}")[:31]
            raw.to_excel(w, sheet_name=name, index=False)
            format_sheet(w, name, raw)
    return path


def make_zip(paths, output):
    path = output / f"IDD_Extension_Cumulative_Summary_All_States_{datetime.now():%d-%m-%Y}.zip"
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, compresslevel=6) as z:
        for item in paths:
            z.write(item, arcname=item.name)
    return path


def send(attachment, state_names):
    sender = env("SMTP_EMAIL", "GMAIL_USERNAME", required=True)
    password = env("SMTP_APP_PASSWORD", "GMAIL_APP_PASSWORD", required=True)
    recipients = [x.strip() for x in re.split(r"[,;]", env("RECIPIENTS", "REPORT_RECIPIENTS", required=True)) if x.strip()]
    size = attachment.stat().st_size / (1024 * 1024)
    limit = float(env("MAX_EMAIL_ATTACHMENT_MB") or "24")
    if size > limit:
        raise ValueError(f"ZIP is {size:.2f} MB, above {limit:.2f} MB")
    msg = EmailMessage()
    msg["From"], msg["To"] = sender, ", ".join(recipients)
    msg["Subject"] = env("MAIL_SUBJECT") or f"IDD Extension Cumulative Summary - {datetime.now():%d-%m-%Y}"
    msg.set_content(f"Dear Team,\n\nPlease find attached the cumulative IDD Extension report for: {', '.join(state_names)}.\n\nRegards,\nGAVB Reporting Automation")
    with attachment.open("rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="zip", filename=attachment.name)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=ssl.create_default_context(), timeout=60) as smtp:
        smtp.login(sender, password)
        smtp.send_message(msg)


def main():
    logging.basicConfig(level=(env("LOG_LEVEL") or "INFO").upper(), format="%(asctime)s [%(levelname)s] %(message)s")
    try:
        logging.info("Running report script version: %s", SCRIPT_VERSION)
        output = Path(env("LOCAL_OUTPUT_FOLDER") or f"/tmp/{datetime.now(timezone.utc):%Y-%m-%d}/idd_extension_cumulative")
        service = build("drive", "v3", credentials=creds(), cache_discovery=False)
        frames, logs = load_all(service, env("DRIVE_FOLDER_ID", required=True))
        report_paths, done = [], []
        for state in states(frames):
            data = state_frame(frames, state)
            if not data.empty:
                report_paths.append(workbook(state, data, logs, output))
                done.append(state)
        if not report_paths:
            raise ValueError("No reports created")
        zipped = make_zip(report_paths, output)
        send(zipped, done)
        print("Created reports:")
        for item in report_paths:
            print(" -", item)
        print("Email attachment:", zipped)
        return 0
    except (ValueError, FileNotFoundError, json.JSONDecodeError) as exc:
        logging.error("Configuration/runtime error: %s", exc); return 2
    except HttpError as exc:
        logging.error("Google API error: %s", exc); return 3
    except smtplib.SMTPAuthenticationError:
        logging.exception("Gmail authentication failed"); return 4
    except (smtplib.SMTPException, OSError) as exc:
        logging.exception("Email delivery failed: %s", exc); return 5
    except Exception as exc:
        logging.exception("Unexpected failure: %s", exc); return 1


if __name__ == "__main__":
    sys.exit(main())

