#!/usr/bin/env python3
"""Fast state-wise GAVB reports from Google Drive CSV files.

BIHAR is excluded case-insensitively, including values such as Bihar, BIHAR,
bihar, or values with surrounding spaces.
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
TOOL_ORDER = [item["tool_name"] for item in TOOL_FILES]

COLUMN_ALIASES = {
    "state": ["STATE", "State", "state", "Cal_STATE", "state_name"],
    "investigator": [
        "QDC", "Investigator", "Nurse", "Nurse_Name", "Nursing_Consultant",
        "Name of Nursing Consultants", "Name of Nurses", "collector_name",
    ],
    "facility_type": ["F_Type", "Facility_Type", "Facility Type", "facilitytype"],
    "facility_level": ["Facility_Level", "Facility Level", "Level", "DH_Below_DH"],
    "submission_date": [
        "SubmissionDate", "Submission Date", "submission_date",
        "SubmissionDateTime", "Submission_Time", "starttime", "endtime",
    ],
}

DH_VALUES = {"dh", "district hospital", "district_hospital", "district hospital dh"}
DAY_START_HOUR = 9
DAY_END_HOUR = 18
DEFAULT_STATE_SPOC = {"Assam": "Nikhil Kumar"}

# States listed here will never receive a state-specific workbook.
# Matching is case-insensitive and ignores leading/trailing spaces.
EXCLUDED_STATES = {"bihar"}


# =============================================================================
# GOOGLE DRIVE READING STRUCTURE FROM THE ORIGINAL SCRIPT
# =============================================================================
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
    query = (
        f"'{folder_id}' in parents and trashed = false and "
        f"name = '{safe_name}'"
    )
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


# =============================================================================
# STANDARDIZATION
# =============================================================================
def normalize(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().lower())


def normalize_state(value: object) -> str:
    return str(value).strip().casefold()


def clean_text(series: pd.Series) -> pd.Series:
    return series.astype("string").str.strip().replace(
        {"": pd.NA, "nan": pd.NA, "None": pd.NA}
    )


def find_column(columns: Sequence[object], aliases: Sequence[str]) -> Optional[str]:
    lookup = {normalize(column): str(column) for column in columns}
    for alias in aliases:
        key = normalize(alias)
        if key in lookup:
            return lookup[key]
    return None


def parse_submission_datetime(series: pd.Series) -> pd.Series:
    text = clean_text(series)
    parsed = pd.to_datetime(
        text,
        format="%d/%m/%Y, %H:%M:%S",
        errors="coerce",
    )
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


def standardize_frame(
    raw: pd.DataFrame, tool_name: str, filename: str
) -> Tuple[pd.DataFrame, Dict[str, object]]:
    detected = {
        key: find_column(raw.columns, aliases)
        for key, aliases in COLUMN_ALIASES.items()
    }
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
        dh_values = {normalize(value) for value in DH_VALUES}
        frame["__Level"] = clean_text(frame[level_source]).map(
            lambda value: "DH" if normalize(value) in dh_values else "Below DH"
        )
    else:
        frame["__Level"] = "Unclassified"

    if detected["submission_date"]:
        frame["__DateTime"] = parse_submission_datetime(
            frame[detected["submission_date"]]
        )
    else:
        frame["__DateTime"] = pd.NaT

    log = {
        "Tool": TOOL_LABELS[tool_name],
        "File": filename,
        "Rows_Read": len(raw),
        "State_Column": detected["state"],
        "Investigator_Column": detected["investigator"],
        "Facility_Level_Column": level_source or "",
        "Submission_Date_Column": detected["submission_date"] or "",
        "Invalid_Submission_Dates": (
            int(frame["__DateTime"].isna().sum())
            if detected["submission_date"] else "Not supplied"
        ),
        "Status": "OK",
    }
    return frame, log


def load_and_standardize_data(drive_service, source_folder_id: str):
    frames: Dict[str, pd.DataFrame] = {}
    logs: List[Dict[str, object]] = []

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
            logs.append({
                "Tool": TOOL_LABELS[tool_name],
                "File": filename,
                "Rows_Read": 0,
                "Status": f"ERROR: {exc}",
            })

    if not frames:
        raise ValueError("No Google Drive CSV file could be processed.")
    return frames, pd.DataFrame(logs)


# =============================================================================
# REPORT TABLES
# =============================================================================
def filter_state(frame: pd.DataFrame, state: str) -> pd.DataFrame:
    return frame.loc[
        frame["__State"].fillna("").str.strip().str.casefold() == normalize_state(state)
    ].copy()


def get_states(frames: Dict[str, pd.DataFrame]) -> List[str]:
    states = {
        str(value).strip()
        for frame in frames.values()
        for value in frame["__State"].dropna().unique()
        if str(value).strip()
    }
    return sorted(states, key=str.casefold)


def is_excluded_state(state: str) -> bool:
    excluded = {normalize_state(value) for value in EXCLUDED_STATES}
    return normalize_state(state) in excluded


def get_investigators(state_frames: Dict[str, pd.DataFrame]) -> List[str]:
    investigators = {
        str(value).strip()
        for frame in state_frames.values()
        for value in frame["__Investigator"].dropna().unique()
        if str(value).strip()
    }
    return sorted(investigators, key=str.casefold)


def count_by_investigator(
    frame: Optional[pd.DataFrame], investigators: List[str]
) -> List[int]:
    if frame is None or frame.empty:
        return [0] * len(investigators)
    grouped = frame.dropna(subset=["__Investigator"]).groupby("__Investigator").size()
    lookup = {
        str(name).strip().casefold(): int(count)
        for name, count in grouped.items()
    }
    return [lookup.get(name.casefold(), 0) for name in investigators]


def build_nurse_wise(state_frames, investigators):
    report = pd.DataFrame({"Name of Nursing Consultants": investigators})
    for tool_name in TOOL_ORDER:
        report[f"# {TOOL_LABELS[tool_name]}"] = count_by_investigator(
            state_frames.get(tool_name), investigators
        )
    total = {report.columns[0]: "Grand Total"}
    total.update({column: int(report[column].sum()) for column in report.columns[1:]})
    return pd.concat([report, pd.DataFrame([total])], ignore_index=True)


def build_dh_below_dh(state_frames, investigators):
    data = {("", "Name of Nurses"): investigators}
    for tool_name in TOOL_ORDER:
        frame = state_frames.get(tool_name)
        for level in ("Below DH", "DH"):
            subset = None if frame is None else frame.loc[frame["__Level"] == level]
            data[(TOOL_LABELS[tool_name], level)] = count_by_investigator(
                subset, investigators
            )
    report = pd.DataFrame(data)
    total = {report.columns[0]: "Grand Total"}
    total.update({column: int(report[column].sum()) for column in report.columns[1:]})
    return pd.concat([report, pd.DataFrame([total])], ignore_index=True)


def fis_shift_counts(frame, investigators, day: bool):
    if frame is None or frame.empty:
        return [0] * len(investigators)
    hours = frame["__DateTime"].dt.hour
    day_mask = (hours >= DAY_START_HOUR) & (hours < DAY_END_HOUR)
    subset = frame.loc[day_mask if day else (~day_mask & hours.notna())]
    return count_by_investigator(subset, investigators)


def build_summary(state, state_frames, investigators, spoc_map):
    report = pd.DataFrame({
        "Name of Nursing Consultants": investigators,
        "State": [state] * len(investigators),
        "State SPOC": [spoc_map.get(state, "")] * len(investigators),
    })
    for tool_name in TOOL_ORDER:
        report[TOOL_LABELS[tool_name]] = count_by_investigator(
            state_frames.get(tool_name), investigators
        )
    fis_position = report.columns.get_loc("FIS") + 1
    report.insert(
        fis_position,
        "FIS-Day (9AM-6PM)",
        fis_shift_counts(state_frames.get("FIS"), investigators, True),
    )
    report.insert(
        fis_position + 1,
        "FIS-Night (6PM-9AM)",
        fis_shift_counts(state_frames.get("FIS"), investigators, False),
    )
    return report


# =============================================================================
# FAST EXCEL OUTPUT
# =============================================================================
def safe_filename(value: str) -> str:
    return re.sub(r'[<>:"/\\|?*]+', "_", value).strip() or "Unknown_State"


def source_sheet_name(filename: str) -> str:
    cleaned = re.sub(r'[\\/*?:\[\]]', "_", Path(filename).stem).strip()
    return cleaned[:31] or "Raw_Data"


def raw_for_export(frame: pd.DataFrame) -> pd.DataFrame:
    return frame.drop(
        columns=[column for column in frame.columns if str(column).startswith("__")],
        errors="ignore",
    )


def sample_widths(df: pd.DataFrame, max_rows: int = 50, cap: int = 35):
    sample = df.head(max_rows)
    widths = []
    for column in df.columns:
        max_length = len(str(column))
        lengths = sample[column].dropna().astype(str).str.len()
        if not lengths.empty:
            max_length = max(max_length, int(lengths.max()))
        widths.append(min(max(max_length + 2, 10), cap))
    return widths


def create_state_workbook(
    state, frames, processing_log, output_folder, spoc_map
) -> Optional[Path]:
    if is_excluded_state(state):
        logging.info("Skipping excluded state: %s", state)
        return None

    state_frames = {
        tool_name: filter_state(frame, state)
        for tool_name, frame in frames.items()
    }
    investigators = get_investigators(state_frames)
    if not investigators:
        logging.warning("Skipping %s because no investigator was found", state)
        return None

    output_folder.mkdir(parents=True, exist_ok=True)
    report_date = datetime.now().strftime("%d-%m-%Y")
    output_path = output_folder / (
        f"{safe_filename(state)}_Tool_Submission_Report_{report_date}.xlsx"
    )
    logging.info("Starting workbook for %s", state)

    nurse_df = build_nurse_wise(state_frames, investigators)
    dh_df = build_dh_below_dh(state_frames, investigators)
    summary_df = build_summary(state, state_frames, investigators, spoc_map)

    with pd.ExcelWriter(
        output_path,
        engine="xlsxwriter",
        engine_kwargs={"options": {"strings_to_urls": False}},
    ) as writer:
        workbook = writer.book
        header_format = workbook.add_format({
            "bold": True,
            "bg_color": "#B7DEE8",
            "border": 1,
            "align": "center",
            "valign": "vcenter",
            "text_wrap": True,
        })
        total_format = workbook.add_format({
            "bold": True, "bg_color": "#B7DEE8", "border": 1
        })
        title_format = workbook.add_format({
            "bold": True, "font_size": 14, "bottom": 1
        })

        nurse_df.to_excel(writer, sheet_name="Nurse Wise", index=False)
        worksheet = writer.sheets["Nurse Wise"]
        worksheet.freeze_panes(1, 1)
        worksheet.autofilter(0, 0, len(nurse_df), len(nurse_df.columns) - 1)
        worksheet.set_row(0, 38, header_format)
        for index, width in enumerate(sample_widths(nurse_df, 100, 32)):
            worksheet.set_column(index, index, width)
        worksheet.set_row(len(nurse_df), None, total_format)

        dh_df.to_excel(
            writer, sheet_name="DH & Below DH", index=False, merge_cells=True
        )
        worksheet = writer.sheets["DH & Below DH"]
        worksheet.freeze_panes(2, 1)
        worksheet.set_row(0, 30, header_format)
        worksheet.set_row(1, 30, header_format)
        for index, width in enumerate(sample_widths(dh_df, 100, 22)):
            worksheet.set_column(index, index, width)
        worksheet.set_row(len(dh_df) + 1, None, total_format)

        summary_df.to_excel(
            writer, sheet_name="Summary", index=False, startrow=2
        )
        worksheet = writer.sheets["Summary"]
        worksheet.write(
            0,
            0,
            f"GAVB Facility Tool Data collection status as on {report_date}",
            title_format,
        )
        worksheet.freeze_panes(3, 1)
        worksheet.autofilter(
            2, 0, len(summary_df) + 2, len(summary_df.columns) - 1
        )
        worksheet.set_row(2, 40, header_format)
        for index, width in enumerate(sample_widths(summary_df, 100, 28)):
            worksheet.set_column(index, index, width)

        for item in TOOL_FILES:
            tool_name = item["tool_name"]
            if tool_name not in state_frames:
                continue
            raw_frame = raw_for_export(state_frames[tool_name])
            sheet_name = source_sheet_name(item["filename"])
            logging.info(
                "%s: writing %s (%s rows x %s columns)",
                state,
                sheet_name,
                len(raw_frame),
                len(raw_frame.columns),
            )
            raw_frame.to_excel(writer, sheet_name=sheet_name, index=False)
            worksheet = writer.sheets[sheet_name]
            worksheet.freeze_panes(1, 0)
            if len(raw_frame.columns):
                worksheet.autofilter(
                    0, 0, len(raw_frame), len(raw_frame.columns) - 1
                )
                worksheet.set_row(0, 30, header_format)
                for index, width in enumerate(
                    sample_widths(raw_frame, 25, 24)
                ):
                    worksheet.set_column(index, index, width)

        state_log = processing_log.assign(State_Workbook=state)
        state_log.to_excel(writer, sheet_name="Processing Log", index=False)
        worksheet = writer.sheets["Processing Log"]
        worksheet.freeze_panes(1, 0)
        worksheet.set_row(0, 30, header_format)
        for index, width in enumerate(sample_widths(state_log, 50, 45)):
            worksheet.set_column(index, index, width)

    logging.info("Created state workbook: %s", output_path)
    return output_path


def parse_state_spoc() -> Dict[str, str]:
    mapping = DEFAULT_STATE_SPOC.copy()
    raw = os.getenv("STATE_SPOC_JSON", "").strip()
    if raw:
        supplied = json.loads(raw)
        if not isinstance(supplied, dict):
            raise ValueError("STATE_SPOC_JSON must be a JSON object")
        mapping.update({
            str(key).strip(): str(value).strip()
            for key, value in supplied.items()
        })
    return mapping


def main() -> int:
    logging.basicConfig(
        level=os.getenv("LOG_LEVEL", "INFO").upper(),
        format="%(asctime)s [%(levelname)s] %(message)s",
    )
    try:
        source_folder_id = required_env("DRIVE_FOLDER_ID")
        run_date_name = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        output_folder = Path(
            os.getenv(
                "LOCAL_OUTPUT_FOLDER",
                f"/tmp/{run_date_name}/state_reports",
            )
        )

        credentials = build_drive_credentials()
        drive_service = build(
            "drive", "v3", credentials=credentials, cache_discovery=False
        )
        frames, processing_log = load_and_standardize_data(
            drive_service, source_folder_id
        )

        states = get_states(frames)

        # Exclude Bihar before any workbook processing starts. This catches
        # Bihar, BIHAR, bihar, and values with surrounding spaces.
        excluded_found = [state for state in states if is_excluded_state(state)]
        for state in excluded_found:
            logging.info("Excluded state from report generation: %s", state)
        states = [state for state in states if not is_excluded_state(state)]

        selected_states = os.getenv("STATES", "").strip()
        if selected_states:
            wanted = {
                normalize_state(value)
                for value in selected_states.split(",")
                if value.strip()
            }
            states = [
                state for state in states
                if normalize_state(state) in wanted
            ]

        if not states:
            raise ValueError("No reportable State values were found after exclusions.")

        spoc_map = parse_state_spoc()
        created: List[Path] = []
        for state in states:
            output_path = create_state_workbook(
                state, frames, processing_log, output_folder, spoc_map
            )
            if output_path:
                created.append(output_path)

        if not created:
            raise ValueError("No state workbook was created.")

        print("Created state workbooks:")
        for output_path in created:
            print(f" - {output_path}")
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
