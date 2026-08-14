#!/usr/bin/env python3
"""Fast state-wise GAVB reports from CSV files stored in Google Drive.

The Google Drive authentication, exact filename lookup, byte download, and
read_tool_file structure follow the original script. Performance improvement is
limited to Excel generation: XlsxWriter writes each workbook once, report sheets
receive formatting, and raw sheets receive only lightweight header formatting.
"""
import io
import json
import logging
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Dict, List, Optional, Sequence, Tuple

import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

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
    "DOD": "DOD", "Maternal_Log": "Maternal Log", "FIS": "FIS",
    "PMSMA": "PMSMA", "SNCU": "SNCU",
    "Referral_Services": "Referral Services", "Digital_System": "Digital System",
    "Exit_Interview": "Exit Interview", "HR": "HR",
    "Labour_Room_Readiness": "Labour Room Readiness", "Supply_Chain": "Supply Chain",
}
TOOL_ORDER = [x["tool_name"] for x in TOOL_FILES]

COLUMN_ALIASES = {
    "state": ["STATE", "State", "state", "Cal_STATE", "state_name"],
    "investigator": ["QDC", "Investigator", "Nurse", "Nurse_Name", "Nursing_Consultant",
                     "Name of Nursing Consultants", "Name of Nurses", "collector_name"],
    "facility_type": ["F_Type", "Facility_Type", "Facility Type", "facilitytype"],
    "facility_level": ["Facility_Level", "Facility Level", "Level", "DH_Below_DH"],
    "submission_date": ["SubmissionDate", "Submission Date", "submission_date",
                        "SubmissionDateTime", "Submission_Time", "starttime", "endtime"],
}
DH_VALUES = {"dh", "district hospital", "district_hospital", "district hospital dh"}
DAY_START_HOUR, DAY_END_HOUR = 9, 18
DEFAULT_STATE_SPOC = {"Assam": "Nikhil Kumar"}

# -----------------------------------------------------------------------------
# GOOGLE DRIVE READING: same structure as the original code
# -----------------------------------------------------------------------------
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
    query = f"'{folder_id}' in parents and trashed = false and name = '{safe_name}'"
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
    if reader == "csv":
        return pd.read_csv(data_stream, low_memory=False, encoding="utf-8-sig")
    raise ValueError(f"Unsupported reader: {reader}")

# -----------------------------------------------------------------------------
# STANDARDIZATION
# -----------------------------------------------------------------------------
def normalize(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def clean_text(series: pd.Series) -> pd.Series:
    return series.astype("string").str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})


def find_column(columns: Sequence[object], aliases: Sequence[str]) -> Optional[str]:
    lookup = {normalize(c): str(c) for c in columns}
    return next((lookup[normalize(a)] for a in aliases if normalize(a) in lookup), None)


def parse_submission_datetime(series: pd.Series) -> pd.Series:
    text = clean_text(series)
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


def standardize_frame(raw: pd.DataFrame, tool: str, filename: str):
    found = {k: find_column(raw.columns, v) for k, v in COLUMN_ALIASES.items()}
    if not found["state"] or not found["investigator"]:
        raise ValueError(f"{filename}: missing State or Investigator/QDC column")
    frame = raw.copy()
    frame["__State"] = clean_text(frame[found["state"]])
    frame["__Investigator"] = clean_text(frame[found["investigator"]])
    source = found["facility_level"] or found["facility_type"]
    if source:
        dh = {normalize(x) for x in DH_VALUES}
        frame["__Level"] = clean_text(frame[source]).map(
            lambda x: "DH" if normalize(x) in dh else "Below DH"
        )
    else:
        frame["__Level"] = "Unclassified"
    frame["__DateTime"] = (
        parse_submission_datetime(frame[found["submission_date"]])
        if found["submission_date"] else pd.NaT
    )
    log = {
        "Tool": TOOL_LABELS[tool], "File": filename, "Rows_Read": len(raw),
        "State_Column": found["state"], "Investigator_Column": found["investigator"],
        "Facility_Level_Column": source or "",
        "Submission_Date_Column": found["submission_date"] or "",
        "Invalid_Submission_Dates": int(frame["__DateTime"].isna().sum())
            if found["submission_date"] else "Not supplied",
        "Status": "OK",
    }
    return frame, log


def load_and_standardize_data(drive_service, folder_id: str):
    frames, logs = {}, []
    for item in TOOL_FILES:
        tool, filename = item["tool_name"], item["filename"]
        logging.info("Reading %s", filename)
        try:
            file_id = find_file_id_by_name(drive_service, folder_id, filename)
            raw_bytes = download_drive_file_bytes(drive_service, file_id)
            raw = read_tool_file(raw_bytes, item["reader"])
            frames[tool], log = standardize_frame(raw, tool, filename)
            logs.append(log)
        except Exception as exc:
            logging.exception("Could not process %s", filename)
            logs.append({"Tool": TOOL_LABELS[tool], "File": filename, "Rows_Read": 0,
                         "Status": f"ERROR: {exc}"})
    if not frames:
        raise ValueError("No Google Drive CSV file could be processed.")
    return frames, pd.DataFrame(logs)

# -----------------------------------------------------------------------------
# REPORT TABLES
# -----------------------------------------------------------------------------
def filter_state(frame, state):
    return frame.loc[frame["__State"].fillna("").str.casefold() == state.casefold()].copy()


def get_states(frames):
    return sorted({str(x).strip() for df in frames.values() for x in df["__State"].dropna().unique()
                   if str(x).strip()}, key=str.casefold)


def get_investigators(frames):
    return sorted({str(x).strip() for df in frames.values() for x in df["__Investigator"].dropna().unique()
                   if str(x).strip()}, key=str.casefold)


def counts(frame, investigators):
    if frame is None or frame.empty:
        return [0] * len(investigators)
    grouped = frame.dropna(subset=["__Investigator"]).groupby("__Investigator").size()
    lookup = {str(k).strip().casefold(): int(v) for k, v in grouped.items()}
    return [lookup.get(name.casefold(), 0) for name in investigators]


def nurse_wise(frames, investigators):
    out = pd.DataFrame({"Name of Nursing Consultants": investigators})
    for tool in TOOL_ORDER:
        out[f"# {TOOL_LABELS[tool]}"] = counts(frames.get(tool), investigators)
    total = {out.columns[0]: "Grand Total", **{c: int(out[c].sum()) for c in out.columns[1:]}}
    return pd.concat([out, pd.DataFrame([total])], ignore_index=True)


def dh_below(frames, investigators):
    data = {("", "Name of Nurses"): investigators}
    for tool in TOOL_ORDER:
        frame = frames.get(tool)
        for level in ("Below DH", "DH"):
            subset = None if frame is None else frame.loc[frame["__Level"] == level]
            data[(TOOL_LABELS[tool], level)] = counts(subset, investigators)
    out = pd.DataFrame(data)
    total = {out.columns[0]: "Grand Total", **{c: int(out[c].sum()) for c in out.columns[1:]}}
    return pd.concat([out, pd.DataFrame([total])], ignore_index=True)


def shift_counts(frame, investigators, day):
    if frame is None or frame.empty:
        return [0] * len(investigators)
    hours = frame["__DateTime"].dt.hour
    day_mask = (hours >= DAY_START_HOUR) & (hours < DAY_END_HOUR)
    return counts(frame.loc[day_mask if day else (~day_mask & hours.notna())], investigators)


def summary(state, frames, investigators, spoc):
    out = pd.DataFrame({
        "Name of Nursing Consultants": investigators,
        "State": [state] * len(investigators),
        "State SPOC": [spoc.get(state, "")] * len(investigators),
    })
    for tool in TOOL_ORDER:
        out[TOOL_LABELS[tool]] = counts(frames.get(tool), investigators)
    pos = out.columns.get_loc("FIS") + 1
    out.insert(pos, "FIS-Day (9AM-6PM)", shift_counts(frames.get("FIS"), investigators, True))
    out.insert(pos + 1, "FIS-Night (6PM-9AM)", shift_counts(frames.get("FIS"), investigators, False))
    return out

# -----------------------------------------------------------------------------
# FAST EXCEL OUTPUT
# -----------------------------------------------------------------------------
def safe_filename(value):
    return re.sub(r'[<>:"/\\|?*]+', "_", value).strip() or "Unknown_State"


def source_sheet_name(filename):
    return re.sub(r'[\\/*?:\[\]]', "_", Path(filename).stem).strip()[:31] or "Raw_Data"


def raw_export(frame):
    return frame.drop(columns=[c for c in frame.columns if str(c).startswith("__")], errors="ignore")


def sample_widths(df: pd.DataFrame, max_rows=50, cap=35):
    sample = df.head(max_rows)
    widths = []
    for col in df.columns:
        max_len = len(str(col))
        if col in sample.columns:
            lengths = sample[col].dropna().astype(str).str.len()
            if not lengths.empty:
                max_len = max(max_len, int(lengths.max()))
        widths.append(min(max(max_len + 2, 10), cap))
    return widths


def create_state_workbook(state, frames, log_df, output_folder, spoc):
    state_frames = {tool: filter_state(df, state) for tool, df in frames.items()}
    investigators = get_investigators(state_frames)
    if not investigators:
        logging.warning("Skipping %s: no investigator names", state)
        return None

    output_folder.mkdir(parents=True, exist_ok=True)
    report_date = datetime.now().strftime("%d-%m-%Y")
    output_path = output_folder / f"{safe_filename(state)}_Tool_Submission_Report_{report_date}.xlsx"
    logging.info("Starting workbook for %s", state)

    nurse_df = nurse_wise(state_frames, investigators)
    dh_df = dh_below(state_frames, investigators)
    summary_df = summary(state, state_frames, investigators, spoc)

    # XlsxWriter writes and formats in one pass. No openpyxl reopen/full-cell scan.
    with pd.ExcelWriter(output_path, engine="xlsxwriter",
                        engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        workbook = writer.book
        header_fmt = workbook.add_format({
            "bold": True, "bg_color": "#B7DEE8", "border": 1,
            "align": "center", "valign": "vcenter", "text_wrap": True,
        })
        total_fmt = workbook.add_format({"bold": True, "bg_color": "#B7DEE8", "border": 1})
        title_fmt = workbook.add_format({"bold": True, "font_size": 14, "bottom": 1})
        data_border_fmt = workbook.add_format({"border": 1})

        # Nurse Wise
        nurse_df.to_excel(writer, sheet_name="Nurse Wise", index=False)
        ws = writer.sheets["Nurse Wise"]
        ws.freeze_panes(1, 1); ws.autofilter(0, 0, len(nurse_df), len(nurse_df.columns)-1)
        ws.set_row(0, 38, header_fmt)
        for i, width in enumerate(sample_widths(nurse_df, 100, 32)):
            ws.set_column(i, i, width)
        ws.set_row(len(nurse_df), None, total_fmt)
        ws.conditional_format(1, 0, len(nurse_df), len(nurse_df.columns)-1,
                              {"type": "no_blanks", "format": data_border_fmt})

        # DH & Below DH
        dh_df.to_excel(writer, sheet_name="DH & Below DH", index=False, merge_cells=True)
        ws = writer.sheets["DH & Below DH"]
        ws.freeze_panes(2, 1)
        ws.set_row(0, 30, header_fmt); ws.set_row(1, 30, header_fmt)
        for i, width in enumerate(sample_widths(dh_df, 100, 22)):
            ws.set_column(i, i, width)
        ws.set_row(len(dh_df) + 1, None, total_fmt)

        # Summary
        summary_df.to_excel(writer, sheet_name="Summary", index=False, startrow=2)
        ws = writer.sheets["Summary"]
        ws.write(0, 0, f"GAVB Facility Tool Data collection status as on {report_date}", title_fmt)
        ws.freeze_panes(3, 1); ws.autofilter(2, 0, len(summary_df)+2, len(summary_df.columns)-1)
        ws.set_row(2, 40, header_fmt)
        for i, width in enumerate(sample_widths(summary_df, 100, 28)):
            ws.set_column(i, i, width)

        # Raw state data: header format, filter, freeze only. No full-sheet borders or autosize scan.
        for item in TOOL_FILES:
            tool = item["tool_name"]
            if tool not in state_frames:
                continue
            raw = raw_export(state_frames[tool])
            sheet = source_sheet_name(item["filename"])
            logging.info("%s: writing %s (%s rows x %s columns)", state, sheet, len(raw), len(raw.columns))
            raw.to_excel(writer, sheet_name=sheet, index=False)
            ws = writer.sheets[sheet]
            ws.freeze_panes(1, 0)
            if len(raw.columns):
                ws.autofilter(0, 0, len(raw), len(raw.columns)-1)
                ws.set_row(0, 30, header_fmt)
                for i, width in enumerate(sample_widths(raw, 25, 24)):
                    ws.set_column(i, i, width)

        log_df.assign(State_Workbook=state).to_excel(writer, sheet_name="Processing Log", index=False)
        ws = writer.sheets["Processing Log"]
        ws.freeze_panes(1, 0); ws.set_row(0, 30, header_fmt)
        for i, width in enumerate(sample_widths(log_df, 50, 45)):
            ws.set_column(i, i, width)

    logging.info("Created state workbook: %s", output_path)
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


def main() -> int:
    logging.basicConfig(level=os.getenv("LOG_LEVEL", "INFO").upper(),
                        format="%(asctime)s [%(levelname)s] %(message)s")
    try:
        source_folder_id = required_env("DRIVE_FOLDER_ID")
        run_date_name = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        output_folder = Path(os.getenv("LOCAL_OUTPUT_FOLDER", f"/tmp/{run_date_name}/state_reports"))

        credentials = build_drive_credentials()
        drive_service = build("drive", "v3", credentials=credentials, cache_discovery=False)
        frames, log_df = load_and_standardize_data(drive_service, source_folder_id)
        states = get_states(frames)

        selected = os.getenv("STATES", "").strip()
        if selected:
            wanted = {x.strip().casefold() for x in selected.split(",") if x.strip()}
            states = [x for x in states if x.casefold() in wanted]
        if not states:
            raise ValueError("No matching State values were found.")

        spoc = parse_state_spoc()
        created = []
        for state in states:
            path = create_state_workbook(state, frames, log_df, output_folder, spoc)
            if path:
                created.append(path)
        if not created:
            raise ValueError("No state workbook was created.")
        print("Created state workbooks:")
        for path in created:
            print(f" - {path}")
        return 0
    except (ValueError, FileNotFoundError) as exc:
        logging.error("Configuration/runtime error: %s", exc)
        return 2
    except HttpError as exc:
        logging.error("Google API request failed: %s", exc)
        return 3
    except Exception as exc:
        logging.exception("Unexpected failure: %s", exc)
        return 1


if __name__ == "__main__":
    sys.exit(main())
