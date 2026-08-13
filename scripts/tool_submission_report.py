#!/usr/bin/env python3
"""Create state-wise Excel reports from GAVB WIDE CSV files."""
from __future__ import annotations
import json, logging, os, re, sys
from datetime import datetime
from pathlib import Path
from typing import Optional
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

INPUT_FOLDER = Path(os.getenv("INPUT_FOLDER", r"C:\Users\YOUR_USERNAME\OneDrive\GAVB_CSV_Files"))
OUTPUT_FOLDER = Path(os.getenv("OUTPUT_FOLDER", r"C:\Users\YOUR_USERNAME\OneDrive\State_Wise_Reports"))

TOOLS = [
 ("DOD", "DOD_2026_WIDE.csv", "DOD_2026_WIDE"),
 ("Maternal Log", "Maternal Log_WIDE.csv", "Maternal_Log_WIDE"),
 ("FIS", "FIS CASE RECORDS 2026_WIDE.csv", "FIS_CASE_RECORDS_2026_WIDE"),
 ("PMSMA", "PMSMA Client Interview_WIDE.csv", "PMSMA_Client_Interview_WIDE"),
 ("SNCU", "SNCU Index Cases Observation 2026_WIDE.csv", "SNCU_Index_Cases_Obs_2026"),
 ("Referral Services", "GABV IDD Referral_WIDE.csv", "IDD_Referral_WIDE"),
 ("Digital System", "GAVB IDD Digital System_WIDE.csv", "IDD_Digital_System_WIDE"),
 ("Exit Interview", "GAVB IDD Exit Interview_WIDE.csv", "IDD_Exit_Interview_WIDE"),
 ("HR", "GAVB IDD HR_WIDE.csv", "IDD_HR_WIDE"),
 ("Labour Room Readiness", "GAVB IDD Labour Room Readiness_WIDE.csv", "IDD_LR_Readiness_WIDE"),
 ("Supply Chain", "GAVB IDD Supply Chain_WIDE.csv", "IDD_Supply_Chain_WIDE"),
]
ORDER = [x[0] for x in TOOLS]
ALIASES = {
 "state": ["STATE","State","state","state_name","Cal_STATE"],
 "investigator": ["QDC","Investigator","Nurse","Nurse_Name","Nursing_Consultant","Name of Nursing Consultants","Name of Nurses","collector_name"],
 "district": ["Cal_DIST","District","DISTRICT","district_name"],
 "facility_type": ["F_Type","Facility_Type","Facility Type","facilitytype"],
 "facility_level": ["Facility_Level","Facility Level","Level","DH_Below_DH"],
 "submission_date": ["SubmissionDate","Submission Date","submission_date","SubmissionDateTime","Submission_Time","starttime","endtime"],
}
DH_VALUES = {"dh", "district hospital", "district_hospital", "district hospital dh"}
DEFAULT_SPOC = {"Assam": "Nikhil Kumar"}
DAY_START, DAY_END = 9, 18

BLUE = PatternFill("solid", fgColor="B7DEE8")
TOTAL = PatternFill("solid", fgColor="B7DEE8")
SIDE = Side(style="thin", color="000000")
BORDER = Border(left=SIDE,right=SIDE,top=SIDE,bottom=SIDE)

def norm(x): return re.sub(r"[^a-z0-9]+", "", str(x).strip().lower())
def clean(s): return s.astype("string").str.strip().replace({"":pd.NA,"nan":pd.NA,"None":pd.NA})
def col(columns, names):
    d={norm(x):str(x) for x in columns}
    return next((d[norm(x)] for x in names if norm(x) in d), None)

def parse_submission_datetime(series: pd.Series) -> pd.Series:
    """Parse values such as 27/01/2026, 19:01:46 safely and day-first."""
    text = clean(series)
    parsed = pd.to_datetime(text, format="%d/%m/%Y, %H:%M:%S", errors="coerce")
    missing = parsed.isna() & text.notna()
    if missing.any():
        # Tolerant fallback for valid variants (ISO, no seconds, etc.).
        try:
            fallback = pd.to_datetime(text.loc[missing], format="mixed", dayfirst=True, errors="coerce")
        except TypeError:  # Compatibility with older pandas versions.
            fallback = pd.to_datetime(text.loc[missing], dayfirst=True, errors="coerce")
        parsed.loc[missing] = fallback
    return parsed

def read_csv(path):
    last=None
    for enc in ("utf-8-sig","utf-8","cp1252","latin1"):
        try: return pd.read_csv(path, encoding=enc, sep=None, engine="python", low_memory=False), enc
        except Exception as e: last=e
    raise last

def standardize(raw, tool, filename):
    m={k:col(raw.columns,v) for k,v in ALIASES.items()}
    if not m["state"] or not m["investigator"]:
        raise ValueError(f"Missing State/Investigator column. Headers: {list(raw.columns)}")
    w=raw.copy()
    w["__State"]=clean(w[m["state"]]); w["__Investigator"]=clean(w[m["investigator"]])
    source=m["facility_level"] or m["facility_type"]
    if source:
        dh={norm(x) for x in DH_VALUES}
        w["__Level"]=clean(w[source]).map(lambda x:"DH" if norm(x) in dh else "Below DH")
    else: w["__Level"]="Unclassified"
    w["__DateTime"]=parse_submission_datetime(w[m["submission_date"]]) if m["submission_date"] else pd.NaT
    w["__Tool"]=tool
    log={"Tool":tool,"File":filename,"Rows_Read":len(raw),"State_Column":m["state"],
         "Investigator_Column":m["investigator"],"Facility_Level_Column":source or "",
         "Submission_Date_Column":m["submission_date"] or "",
         "Invalid_Submission_Dates":int(w["__DateTime"].isna().sum()) if m["submission_date"] else "Not supplied",
         "Status":"OK"}
    return w,log

def load_all():
    frames={}; logs=[]
    for tool,filename,_ in TOOLS:
        p=INPUT_FOLDER/filename
        if not p.exists(): logs.append({"Tool":tool,"File":filename,"Status":"FILE NOT FOUND"}); continue
        try:
            raw,enc=read_csv(p); frames[tool],lg=standardize(raw,tool,filename); lg["Encoding"]=enc; logs.append(lg)
        except Exception as e: logs.append({"Tool":tool,"File":filename,"Status":f"ERROR: {e}"})
    if not frames: raise ValueError("No CSV file could be processed")
    return frames,pd.DataFrame(logs)

def state_filter(df,state): return df[df["__State"].fillna("").str.casefold()==state.casefold()].copy()
def names(frames):
    return sorted({str(x).strip() for d in frames.values() for x in d["__Investigator"].dropna().unique()},key=str.casefold)
def counts(df,nms):
    if df is None or df.empty:return [0]*len(nms)
    c=df.dropna(subset=["__Investigator"]).groupby("__Investigator").size()
    look={str(k).strip().casefold():int(v) for k,v in c.items()}
    return [look.get(n.casefold(),0) for n in nms]

def nurse_table(frames,nms):
    out=pd.DataFrame({"Name of Nursing Consultants":nms})
    for t in ORDER: out[f"# {t}"]=counts(frames.get(t),nms)
    total={out.columns[0]:"Grand Total",**{c:int(out[c].sum()) for c in out.columns[1:]}}
    return pd.concat([out,pd.DataFrame([total])],ignore_index=True)

def dh_table(frames,nms):
    data={("","Name of Nurses"):nms}
    for t in ORDER:
        d=frames.get(t)
        for lev in ("Below DH","DH"):
            data[(t,lev)]=counts(None if d is None else d[d["__Level"]==lev],nms)
    out=pd.DataFrame(data); total={out.columns[0]:"Grand Total",**{c:int(out[c].sum()) for c in out.columns[1:]}}
    return pd.concat([out,pd.DataFrame([total])],ignore_index=True)

def shifts(df,nms,day=True):
    if df is None or df.empty:return [0]*len(nms)
    hour=df["__DateTime"].dt.hour
    mask=(hour>=DAY_START)&(hour<DAY_END)
    return counts(df[mask if day else (~mask & hour.notna())],nms)

def summary(state,frames,nms,spoc):
    out=pd.DataFrame({"Name of Nursing Consultants":nms,"State":state,"State SPOC":spoc.get(state,"")})
    for t in ORDER: out[t]=counts(frames.get(t),nms)
    pos=out.columns.get_loc("FIS")+1
    out.insert(pos,"FIS-Day (9AM-6PM)",shifts(frames.get("FIS"),nms,True))
    out.insert(pos+1,"FIS-Night (6PM-9AM)",shifts(frames.get("FIS"),nms,False))
    return out

def safe(x): return re.sub(r'[<>:"/\\|?*]+','_',x).strip() or "Unknown_State"
def sheet(x): return re.sub(r'[\\/*?:\[\]]','_',x)[:31]
def autosize(ws,limit=40):
    for i in range(1,ws.max_column+1):
        vals=[len(str(c.value)) if c.value is not None else 0 for c in ws[get_column_letter(i)]]
        ws.column_dimensions[get_column_letter(i)].width=min(max(vals+[8])+2,limit)
def style(path):
    wb=load_workbook(path)
    for ws in wb.worksheets:
        headers=(1,2) if ws.title=="DH & Below DH" else ((3,) if ws.title=="Summary" else (1,))
        ws.freeze_panes="A4" if ws.title=="Summary" else ("B3" if ws.title=="DH & Below DH" else "A2")
        for r in headers:
            for c in ws[r]: c.fill=BLUE;c.font=Font(bold=True);c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True);c.border=BORDER
        for row in ws.iter_rows():
            for c in row:c.border=BORDER
        if ws.title in ("Nurse Wise","DH & Below DH"):
            for c in ws[ws.max_row]:c.fill=TOTAL;c.font=Font(bold=True)
        autosize(ws)
    wb.save(path)

def create_book(state,frames,logs,spoc):
    sf={t:state_filter(d,state) for t,d in frames.items()}; nms=names(sf)
    if not nms: return None
    OUTPUT_FOLDER.mkdir(parents=True,exist_ok=True); stamp=datetime.now().strftime("%d-%m-%Y")
    path=OUTPUT_FOLDER/f"{safe(state)}_Tool_Submission_Report_{stamp}.xlsx"
    with pd.ExcelWriter(path,engine="openpyxl") as writer:
        nurse_table(sf,nms).to_excel(writer,sheet_name="Nurse Wise",index=False)
        dh_table(sf,nms).to_excel(writer,sheet_name="DH & Below DH",index=True,merge_cells=True)
        summary(state,sf,nms,spoc).to_excel(writer,sheet_name="Summary",index=False,startrow=2)
        for t,_,sn in TOOLS:
            if t in sf:
                sf[t].drop(columns=[c for c in sf[t].columns if c.startswith("__")]).to_excel(writer,sheet_name=sheet(sn),index=False)
        logs.assign(State_Workbook=state).to_excel(writer,sheet_name="Processing Log",index=False)
    wb=load_workbook(path);wb["DH & Below DH"].delete_cols(1);wb["Summary"]["A1"]=f"GAVB Facility Tool Data collection status as on {stamp}";wb.save(path)
    style(path);return path

def main():
    logging.basicConfig(level=logging.INFO,format="%(asctime)s [%(levelname)s] %(message)s")
    if not INPUT_FOLDER.exists(): raise FileNotFoundError(f"Input folder not found: {INPUT_FOLDER}")
    frames,logs=load_all(); states=sorted({str(x).strip() for d in frames.values() for x in d["__State"].dropna().unique()},key=str.casefold)
    wanted={x.strip().casefold() for x in os.getenv("STATES","").split(",") if x.strip()}
    if wanted: states=[x for x in states if x.casefold() in wanted]
    spoc=DEFAULT_SPOC.copy(); raw=os.getenv("STATE_SPOC_JSON","").strip()
    if raw: spoc.update(json.loads(raw))
    made=[create_book(s,frames,logs,spoc) for s in states]
    print("Created:");[print(" -",p) for p in made if p]
    return 0
if __name__=="__main__": sys.exit(main())
