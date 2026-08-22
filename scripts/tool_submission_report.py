#!/usr/bin/env python3
"""Fast state-wise GAVB reports from Google Drive CSV files."""

import io
import json
import logging
import os
import re
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Optional, Sequence

import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

SCRIPT_VERSION = "2026-08-21-flat-dh-columns-v3"
DRIVE_SCOPES = ["https://www.googleapis.com/auth/drive.readonly"]

TOOL_FILES = [
    {"tool_name": "DOD", "filename": "DOD_2026_WIDE.csv", "reader": "csv"},
    {"tool_name": "Maternal_Log", "filename": "Maternal Log_WIDE.csv", "reader": "csv"},
    {"tool_name": "FIS", "filename": "FIS CASE RECORDS 2026_WIDE.csv", "reader": "csv"},
    {"tool_name": "PMSMA", "filename": "PMSMA Client Interview_WIDE.csv", "reader": "csv"},
    {
        "tool_name": "SNCU",
        "filename": "SNCU Index Cases Observation 2026_WIDE.csv",
        "reader": "csv",
    },
    {"tool_name": "Referral_Services", "filename": "GABV IDD Referral_WIDE.csv", "reader": "csv"},
    {"tool_name": "Digital_System", "filename": "GAVB IDD Digital System_WIDE.csv", "reader": "csv"},
    {"tool_name": "Exit_Interview", "filename": "GAVB IDD Exit Interview_WIDE.csv", "reader": "csv"},
    {"tool_name": "HR", "filename": "GAVB IDD HR_WIDE.csv", "reader": "csv"},
    {
        "tool_name": "Labour_Room_Readiness",
        "filename": "GAVB IDD Labour Room Readiness_WIDE.csv",
        "reader": "csv",
    },
    {"tool_name": "Supply_Chain", "filename": "GAVB IDD Supply Chain_WIDE.csv", "reader": "csv"},
]

TOOL_LABELS = {
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

TOOL_ORDER = [x["tool_name"] for x in TOOL_FILES]

COLUMN_ALIASES = {
    "state": ["STATE", "State", "state", "Cal_STATE", "state_name"],
    "investigator": [
        "QDC",
        "Investigator",
        "Nurse",
        "Nurse_Name",
        "Nursing_Consultant",
        "Name of Nursing Consultants",
        "Name of Nurses",
        "collector_name",
    ],
    "facility_type": ["F_Type", "Facility_Type", "Facility Type", "facilitytype"],
    "facility_level": ["Facility_Level", "Facility Level", "Level", "DH_Below_DH"],
    "submission_date": [
        "SubmissionDate",
        "Submission Date",
        "submission_date",
        "SubmissionDateTime",
        "Submission_Time",
        "starttime",
        "endtime",
    ],
}

DH_VALUES = {"dh", "district hospital", "district_hospital", "district hospital dh"}
DAY_START_HOUR, DAY_END_HOUR = 9, 18
DEFAULT_STATE_SPOC = {"Assam": "Nikhil Kumar"}
EXCLUDED_STATES = {"bihar"}


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
            info = json.loads(service_account_json)
        except json.JSONDecodeError as exc:
            raise ValueError("GOOGLE_SERVICE_ACCOUNT_JSON is not valid JSON") from exc
        return Credentials.from_service_account_info(info, scopes=DRIVE_SCOPES)

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


def normalize(v) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(v).strip().lower())


def normalize_state(v) -> str:
    return str(v).strip().casefold()


def clean_text(s):
    return s.astype("string").str.strip().replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})


def find_column(columns: Sequence[object], aliases: Sequence[str]) -> Optional[str]:
    lookup = {normalize(c): str(c) for c in columns}
    return next((lookup[normalize(a)] for a in aliases if normalize(a) in lookup), None)


def parse_submission_datetime(series):
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


def standardize_frame(raw, tool_name, filename):
    detected = {k: find_column(raw.columns, a) for k, a in COLUMN_ALIASES.items()}

    if not detected["state"] or not detected["investigator"]:
        raise ValueError(
            f"{filename}: missing State or Investigator/QDC column. "
            f"Available headers: {list(raw.columns)}"
        )

    frame = raw.copy()
    frame["__State"] = clean_text(frame[detected["state"]])
    frame["__Investigator"] = clean_text(frame[detected["investigator"]])

    level_source = detected["facility_level"] or detected["facility_type"]
    if level_source:
        dh = {normalize(v) for v in DH_VALUES}
        frame["__Level"] = clean_text(frame[level_source]).map(
            lambda v: "DH" if normalize(v) in dh else "Below DH"
        )
    else:
        frame["__Level"] = "Unclassified"

    if detected["submission_date"]:
        frame["__DateTime"] = parse_submission_datetime(frame[detected["submission_date"]])
        invalid_dates = int(frame["__DateTime"].isna().sum())
    else:
        frame["__DateTime"] = pd.NaT
        invalid_dates = "Not supplied"

    log = {
        "Tool": TOOL_LABELS[tool_name],
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


def load_and_standardize_data(drive_service, source_folder_id):
    frames = {}
    logs = []

    for item in TOOL_FILES:
        tool_name = item["tool_name"]
        filename = item["filename"]
        logging.info("Reading %s", filename)

        try:
            file_id = find_file_id_by_name(drive_service, source_folder_id, filename)
            raw_bytes = download_drive_file_bytes(drive_service, file_id)
            raw_frame = read_tool_file(raw_bytes, item["reader"])
            frames[tool_name], log = standardize_frame(raw_frame, tool_name, filename)
            logs.append(log)
        except Exception as exc:
            logging.exception("Could not process %s", filename)
            logs.append(
                {
                    "Tool": TOOL_LABELS[tool_name],
                    "File": filename,
                    "Rows_Read": 0,
                    "Status": f"ERROR: {exc}",
                }
            )

    if not frames:
        raise ValueError("No Google Drive CSV file could be processed.")

    return frames, pd.DataFrame(logs)


def filter_state(frame, state):
    mask = frame["__State"].fillna("").str.strip().str.casefold() == normalize_state(state)
    return frame.loc[mask].copy()


def get_states(frames):
    return sorted(
        {
            str(v).strip()
            for f in frames.values()
            for v in f["__State"].dropna().unique()
            if str(v).strip()
        },
        key=str.casefold,
    )


def is_excluded_state(state):
    excluded = {normalize_state(v) for v in EXCLUDED_STATES}
    return normalize_state(state) in excluded


def get_investigators(state_frames):
    return sorted(
        {
            str(v).strip()
            for f in state_frames.values()
            for v in f["__Investigator"].dropna().unique()
            if str(v).strip()
        },
        key=str.casefold,
    )


def count_by_investigator(frame, investigators):
    if frame is None or frame.empty:
        return [0] * len(investigators)

    grouped = frame.dropna(subset=["__Investigator"]).groupby("__Investigator").size()
    lookup = {str(n).strip().casefold(): int(c) for n, c in grouped.items()}
    return [lookup.get(n.casefold(), 0) for n in investigators]


def build_nurse_wise(state_frames, investigators):
    report = pd.DataFrame({"Name of Nursing Consultants": investigators})

    for t in TOOL_ORDER:
        report[f"# {TOOL_LABELS[t]}"] = count_by_investigator(state_frames.get(t), investigators)

    total = {report.columns[0]: "Grand Total"}
    total.update({c: int(report[c].sum()) for c in report.columns[1:]})

    return pd.concat([report, pd.DataFrame([total])], ignore_index=True)


def build_dh_below_dh(state_frames, investigators):
    # CRITICAL FIX: ordinary columns, never MultiIndex.
    report = pd.DataFrame({"Name of Nurses": investigators})

    for t in TOOL_ORDER:
        frame = state_frames.get(t)
        for level in ("Below DH", "DH"):
            subset = None if frame is None else frame.loc[frame["__Level"] == level]
            report[f"{TOOL_LABELS[t]} - {level}"] = count_by_investigator(subset, investigators)

    total = {"Name of Nurses": "Grand Total"}
    total.update({c: int(report[c].sum()) for c in report.columns[1:]})

    return pd.concat([report, pd.DataFrame([total])], ignore_index=True)


def fis_shift_counts(frame, investigators, day):
    if frame is None or frame.empty:
        return [0] * len(investigators)

    hours = frame["__DateTime"].dt.hour
    valid = frame["__DateTime"].notna()
    mask = valid & (hours >= DAY_START_HOUR) & (hours < DAY_END_HOUR)

    return count_by_investigator(frame.loc[mask if day else valid & ~mask], investigators)


def build_summary(state, state_frames, investigators, spoc_map):
    report = pd.DataFrame(
        {
            "Name of Nursing Consultants": investigators,
            "State": [state] * len(investigators),
            "State SPOC": [spoc_map.get(state, spoc_map.get(state.title(), ""))] * len(investigators),
        }
    )

    for t in TOOL_ORDER:
        report[TOOL_LABELS[t]] = count_by_investigator(state_frames.get(t), investigators)

    pos = report.columns.get_loc("FIS") + 1
    report.insert(
        pos,
        "FIS-Day (9AM-6PM)",
        fis_shift_counts(state_frames.get("FIS"), investigators, True),
    )
    report.insert(
        pos + 1,
        "FIS-Night (6PM-9AM)",
        fis_shift_counts(state_frames.get("FIS"), investigators, False),
    )

    return report


def safe_filename(v):
    return re.sub(r'[<>:"/\\|?*]+', "_", v).strip() or "Unknown_State"


def source_sheet_name(filename):
    return re.sub(r"[\\/*?:\[\]]", "_", Path(filename).stem).strip()[:31] or "Raw_Data"


def raw_for_export(frame):
    return frame.drop(columns=[c for c in frame.columns if str(c).startswith("__")], errors="ignore")


def sample_widths(df, max_rows=50, cap=35):
    widths = []
    sample = df.head(max_rows)

    for c in df.columns:
        m = len(str(c))
        lengths = sample[c].dropna().astype(str).str.len()
        if not lengths.empty:
            m = max(m, int(lengths.max()))
        widths.append(min(max(m + 2, 10), cap))

    return widths


def format_sheet(ws, df, header_format, total_format=None, header_row=0, freeze=(1, 0), cap=24):
    ws.freeze_panes(*freeze)

    if len(df.columns):
        ws.autofilter(header_row, 0, header_row + len(df), len(df.columns) - 1)

    ws.set_row(header_row, 38, header_format)

    for i, w in enumerate(sample_widths(df, 100, cap)):
        ws.set_column(i, i, w)

    if total_format is not None:
        ws.set_row(header_row + len(df), None, total_format)


def create_state_workbook(state, frames, processing_log, output_folder, spoc_map):
    if is_excluded_state(state):
        return None

    state_frames = {t: filter_state(f, state) for t, f in frames.items()}
    investigators = get_investigators(state_frames)
    if not investigators:
        return None

    output_folder.mkdir(parents=True, exist_ok=True)
    report_date = datetime.now().strftime("%d-%m-%Y")
    output_path = output_folder / f"{safe_filename(state)}_Tool_Submission_Report_{report_date}.xlsx"
    logging.info("Starting workbook for %s", state)

    nurse_df = build_nurse_wise(state_frames, investigators)
    dh_df = build_dh_below_dh(state_frames, investigators)
    summary_df = build_summary(state, state_frames, investigators, spoc_map)

    assert not isinstance(dh_df.columns, pd.MultiIndex), (
        "Internal error: DH report must not have MultiIndex columns"
    )

    with pd.ExcelWriter(
        output_path,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False}},
    ) as writer:
        wb = writer.book
        hf = wb.add_format(
            {
                "bold": True,
                "bg_color": "#B7DEE8",
                "border": 1,
                "align": "center",
                "valign": "vcenter",
                "text_wrap": True,
            }
        )
        tf = wb.add_format({"bold": True, "bg_color": "#B7DEE8", "border": 1})
        title = wb.add_format({"bold": True, "font_size": 14, "bottom": 1})

        nurse_df.to_excel(writer, sheet_name="Nurse Wise", index=False)
        format_sheet(writer.sheets["Nurse Wise"], nurse_df, hf, tf, 0, (1, 1), 32)

        dh_df.to_excel(writer, sheet_name="DH & Below DH", index=False)
        format_sheet(writer.sheets["DH & Below DH"], dh_df, hf, tf, 0, (1, 1), 24)

        summary_df.to_excel(writer, sheet_name="Summary", index=False, startrow=2)
        ws = writer.sheets["Summary"]
        ws.write(0, 0, f"GAVB Facility Tool Data collection status as on {report_date}", title)
        format_sheet(ws, summary_df, hf, None, 2, (3, 1), 28)

        for item in TOOL_FILES:
            t = item["tool_name"]
            if t not in state_frames:
                continue

            raw = raw_for_export(state_frames[t])
            sn = source_sheet_name(item["filename"])
            logging.info(
                "%s: writing %s (%s rows x %s columns)",
                state,
                sn,
                len(raw),
                len(raw.columns),
            )
            raw.to_excel(writer, sheet_name=sn, index=False)
            format_sheet(writer.sheets[sn], raw, hf, None, 0, (1, 0), 24)

        state_log = processing_log.assign(State_Workbook=state)
        state_log.to_excel(writer, sheet_name="Processing Log", index=False)
        format_sheet(writer.sheets["Processing Log"], state_log, hf, None, 0, (1, 0), 45)

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


def main():
    logging.basicConfig(
        level=os.getenv("LOG_LEVEL", "INFO").upper(),
        format="%(asctime)s [%(levelname)s] %(message)s",
    )

    try:
        logging.info("Running report script version: %s", SCRIPT_VERSION)

        source_folder_id = required_env("DRIVE_FOLDER_ID")
        run_date_name = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        output_folder = Path(
            os.getenv("LOCAL_OUTPUT_FOLDER", f"/tmp/{run_date_name}/state_reports")
        )

        credentials = build_drive_credentials()
        drive_service = build("drive", "v3", credentials=credentials, cache_discovery=False)
        frames, processing_log = load_and_standardize_data(drive_service, source_folder_id)

        states = get_states(frames)

        for s in [x for x in states if is_excluded_state(x)]:
            logging.info("Excluded state from report generation: %s", s)

        states = [s for s in states if not is_excluded_state(s)]

        selected = os.getenv("STATES", "").strip()
        if selected:
            wanted = {normalize_state(v) for v in selected.split(",") if v.strip()}
            states = [s for s in states if normalize_state(s) in wanted]

        if not states:
            raise ValueError("No reportable State values were found after exclusions.")

        spoc_map = parse_state_spoc()
        created = [
            p
            for s in states
            if (p := create_state_workbook(s, frames, processing_log, output_folder, spoc_map))
        ]

        if not created:
            raise ValueError("No state workbook was created.")

        print("Created state workbooks:")
        for p in created:
            print(f" - {p}")

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
