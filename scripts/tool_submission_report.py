#!/usr/bin/env python3
"""Create the ASSAM GAVB report from Google Drive CSVs and send it through Gmail."""
import io, json, logging, os, re, smtplib, ssl, sys
from datetime import datetime, timezone
from email.message import EmailMessage
from pathlib import Path
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload

SCRIPT_VERSION="2026-09-21-assam-smtp-secret-names-v2"
DRIVE_SCOPES=["https://www.googleapis.com/auth/drive.readonly"]
TARGET_STATE="ASSAM"
TOOLS=[
("DOD","DOD_2026_WIDE.csv"),("Maternal_Log","Maternal Log_WIDE.csv"),
("FIS","FIS CASE RECORDS 2026_WIDE.csv"),("PMSMA","PMSMA Client Interview_WIDE.csv"),
("SNCU","SNCU Index Cases Observation 2026_WIDE.csv"),("Referral_Services","GABV IDD Referral_WIDE.csv"),
("Digital_System","GAVB IDD Digital System_WIDE.csv"),("Exit_Interview","GAVB IDD Exit Interview_WIDE.csv"),
("HR","GAVB IDD HR_WIDE.csv"),("Labour_Room_Readiness","GAVB IDD Labour Room Readiness_WIDE.csv"),
("Supply_Chain","GAVB IDD Supply Chain_WIDE.csv")]
LABELS={"DOD":"DOD","Maternal_Log":"Maternal Log","FIS":"FIS","PMSMA":"PMSMA","SNCU":"SNCU",
"Referral_Services":"Referral Services","Digital_System":"Digital System","Exit_Interview":"Exit Interview",
"HR":"HR","Labour_Room_Readiness":"Labour Room Readiness","Supply_Chain":"Supply Chain"}
ALIASES={
"state":["STATE","State","state","Cal_STATE","state_name"],
"investigator":["QDC","Investigator","Nurse","Nurse_Name","Nursing_Consultant","Name of Nursing Consultants","Name of Nurses","collector_name"],
"facility_type":["F_Type","Facility_Type","Facility Type","facilitytype"],
"facility_level":["Facility_Level","Facility Level","Level","DH_Below_DH"],
"submission_date":["SubmissionDate","Submission Date","submission_date","SubmissionDateTime","Submission_Time","starttime","endtime"]}
DH_VALUES={"dh","district hospital","district_hospital","district hospital dh"}

def env(*names,required=False):
    for n in names:
        v=os.getenv(n,"").strip()
        if v:return v
    if required: raise ValueError("Missing required environment variable. Accepted name(s): "+", ".join(names))
    return ""

def norm(v):return re.sub(r"[^a-z0-9]+","",str(v).strip().lower())
def clean(s):return s.astype("string").str.strip().replace({"":pd.NA,"nan":pd.NA,"None":pd.NA})
def find_col(cols,names):
    lookup={norm(c):str(c) for c in cols}
    return next((lookup[norm(x)] for x in names if norm(x) in lookup),None)
def flatten(df):
    x=df.copy()
    if isinstance(x.columns,pd.MultiIndex):x.columns=[" | ".join(str(v).strip() for v in t if str(v).strip()) for t in x.columns.to_flat_index()]
    else:x.columns=[str(c) for c in x.columns]
    return x

def credentials():
    raw=env("GOOGLE_SERVICE_ACCOUNT_JSON")
    file=env("GOOGLE_SERVICE_ACCOUNT_FILE") or "credential.json"
    if raw:return Credentials.from_service_account_info(json.loads(raw),scopes=DRIVE_SCOPES)
    if os.path.exists(file):return Credentials.from_service_account_file(file,scopes=DRIVE_SCOPES)
    raise ValueError("Set GOOGLE_SERVICE_ACCOUNT_JSON or GOOGLE_SERVICE_ACCOUNT_FILE")

def drive_file_id(service,folder,name):
    safe = name.replace(chr(92), chr(92) * 2).replace(chr(39), chr(92) + chr(39))
    q=f"'{folder}' in parents and trashed = false and name = '{safe}'"
    files=service.files().list(q=q,spaces="drive",fields="files(id,name)",pageSize=10).execute().get("files",[])
    if not files:raise FileNotFoundError(f"File not found: {name}")
    return files[0]["id"]
def download(service,file_id):
    b=io.BytesIO(); d=MediaIoBaseDownload(b,service.files().get_media(fileId=file_id)); done=False
    while not done:_,done=d.next_chunk()
    return b.getvalue()
def parse_dt(s):
    t=clean(s); p=pd.to_datetime(t,format="%d/%m/%Y, %H:%M:%S",errors="coerce"); m=p.isna()&t.notna()
    if m.any():
        try:p.loc[m]=pd.to_datetime(t.loc[m],format="mixed",dayfirst=True,errors="coerce")
        except TypeError:p.loc[m]=pd.to_datetime(t.loc[m],dayfirst=True,errors="coerce")
    return p

def standardize(raw,tool,filename):
    d={k:find_col(raw.columns,v) for k,v in ALIASES.items()}
    if not d["state"] or not d["investigator"]:raise ValueError(f"{filename}: missing State or Investigator/QDC column")
    f=raw.copy(); f["__State"]=clean(f[d["state"]]); f["__Investigator"]=clean(f[d["investigator"]])
    src=d["facility_level"] or d["facility_type"]
    if src:
        dh={norm(v) for v in DH_VALUES}; f["__Level"]=clean(f[src]).map(lambda v:"DH" if norm(v) in dh else "Below DH")
    else:f["__Level"]="Unclassified"
    f["__DateTime"]=parse_dt(f[d["submission_date"]]) if d["submission_date"] else pd.NaT
    log={"Tool":LABELS[tool],"File":filename,"Rows_Read":len(raw),"State_Column":d["state"],
         "Investigator_Column":d["investigator"],"Facility_Level_Column":src or "",
         "Submission_Date_Column":d["submission_date"] or "","Invalid_Submission_Dates":int(f["__DateTime"].isna().sum()) if d["submission_date"] else "Not supplied","Status":"OK"}
    return f,log

def load_all(service,folder):
    frames={}; logs=[]
    for tool,name in TOOLS:
        logging.info("Reading %s",name)
        try:
            raw=pd.read_csv(io.BytesIO(download(service,drive_file_id(service,folder,name))),low_memory=False,encoding="utf-8-sig")
            frames[tool],log=standardize(raw,tool,name); logs.append(log)
        except Exception as e:
            logging.exception("Could not process %s",name); logs.append({"Tool":LABELS[tool],"File":name,"Rows_Read":0,"Status":f"ERROR: {e}"})
    if not frames:raise ValueError("No Google Drive CSV file could be processed")
    return frames,pd.DataFrame(logs)
def assam(f):return f.loc[f["__State"].fillna("").str.strip().str.casefold()=="assam"].copy()
def investigators(frames):return sorted({str(v).strip() for f in frames.values() for v in f["__Investigator"].dropna().unique() if str(v).strip()},key=str.casefold)
def counts(f,names):
    if f is None or f.empty:return [0]*len(names)
    g=f.dropna(subset=["__Investigator"]).groupby("__Investigator").size(); q={str(k).strip().casefold():int(v) for k,v in g.items()}
    return [q.get(n.casefold(),0) for n in names]
def nurse_report(frames,names):
    r=pd.DataFrame({"Name of Nursing Consultants":names})
    for t,_ in TOOLS:r[f"# {LABELS[t]}"]=counts(frames.get(t),names)
    return pd.concat([r,pd.DataFrame([{r.columns[0]:"Grand Total",**{c:int(r[c].sum()) for c in r.columns[1:]}}])],ignore_index=True)
def dh_report(frames,names):
    r=pd.DataFrame({"Name of Nurses":names})
    for t,_ in TOOLS:
        for level in ("Below DH","DH"):
            f=frames.get(t); r[f"{LABELS[t]} - {level}"]=counts(None if f is None else f.loc[f["__Level"]==level],names)
    return pd.concat([r,pd.DataFrame([{"Name of Nurses":"Grand Total",**{c:int(r[c].sum()) for c in r.columns[1:]}}])],ignore_index=True)
def shifts(f,names,day):
    if f is None or f.empty:return [0]*len(names)
    valid=f["__DateTime"].notna(); h=f["__DateTime"].dt.hour; dm=valid&(h>=9)&(h<18)
    return counts(f.loc[dm if day else valid&~dm],names)
def summary(frames,names):
    r=pd.DataFrame({"Name of Nursing Consultants":names,"State":["ASSAM"]*len(names),"State SPOC":["Nikhil Kumar"]*len(names)})
    for t,_ in TOOLS:r[LABELS[t]]=counts(frames.get(t),names)
    p=r.columns.get_loc("FIS")+1; r.insert(p,"FIS-Day (9AM-6PM)",shifts(frames.get("FIS"),names,True)); r.insert(p+1,"FIS-Night (6PM-9AM)",shifts(frames.get("FIS"),names,False))
    return r
def sheet_name(filename):return re.sub(r"[\/*?:\[\]]","_",Path(filename).stem).strip()[:31] or "Raw_Data"
def widths(df,cap):
    s=df.head(100); out=[]
    for c in df.columns:
        z=s[c].dropna().astype(str).str.len(); m=max(len(str(c)),int(z.max()) if not z.empty else 0); out.append(min(max(m+2,10),cap))
    return out
def fmt(ws,df,h,total=None,row=0,freeze=(1,0),cap=24):
    ws.freeze_panes(*freeze)
    if len(df.columns):ws.autofilter(row,0,row+len(df),len(df.columns)-1)
    ws.set_row(row,38,h)
    for i,w in enumerate(widths(df,cap)):ws.set_column(i,i,w)
    if total is not None:ws.set_row(row+len(df),None,total)

def create_report(frames,process_log,outdir):
    sf={t:assam(f) for t,f in frames.items()}; names=investigators(sf)
    if not names:raise ValueError("No ASSAM records were found")
    outdir.mkdir(parents=True,exist_ok=True); date=datetime.now().strftime("%d-%m-%Y"); path=outdir/f"ASSAM_Tool_Submission_Report_{date}.xlsx"
    logging.info("Starting workbook for ASSAM")
    n=flatten(nurse_report(sf,names)); d=flatten(dh_report(sf,names)); s=flatten(summary(sf,names))
    with pd.ExcelWriter(path,engine="xlsxwriter",engine_kwargs={"options":{"strings_to_urls":False}}) as w:
        wb=w.book; h=wb.add_format({"bold":True,"bg_color":"#B7DEE8","border":1,"align":"center","valign":"vcenter","text_wrap":True}); total=wb.add_format({"bold":True,"bg_color":"#B7DEE8","border":1}); title=wb.add_format({"bold":True,"font_size":14,"bottom":1})
        n.to_excel(w,sheet_name="Nurse Wise",index=False); fmt(w.sheets["Nurse Wise"],n,h,total,cap=32)
        d.to_excel(w,sheet_name="DH & Below DH",index=False); fmt(w.sheets["DH & Below DH"],d,h,total,cap=24)
        s.to_excel(w,sheet_name="Summary",index=False,startrow=2); w.sheets["Summary"].write(0,0,f"GAVB Facility Tool Data collection status as on {date}",title); fmt(w.sheets["Summary"],s,h,row=2,freeze=(3,1),cap=28)
        for t,name in TOOLS:
            if t not in sf:continue
            raw=flatten(sf[t].drop(columns=[c for c in sf[t].columns if str(c).startswith("__")],errors="ignore")); sn=sheet_name(name)
            logging.info("ASSAM: writing %s (%s rows x %s columns)",sn,len(raw),len(raw.columns)); raw.to_excel(w,sheet_name=sn,index=False); fmt(w.sheets[sn],raw,h,cap=24)
        lg=flatten(process_log.assign(State_Workbook="ASSAM")); lg.to_excel(w,sheet_name="Processing Log",index=False); fmt(w.sheets["Processing Log"],lg,h,cap=45)
    logging.info("Created Assam workbook: %s",path); return path

def send_email(path):
    # Exact names shown in the repository secrets screenshot, with backward-compatible fallbacks.
    sender=env("SMTP_EMAIL","GMAIL_USERNAME","EMAIL_USERNAME",required=True)
    password=env("SMTP_APP_PASSWORD","GMAIL_APP_PASSWORD","EMAIL_PASSWORD",required=True)
    raw_to=env("RECIPIENTS","REPORT_RECIPIENTS","EMAIL_TO",required=True)
    recipients=[x.strip() for x in re.split(r"[,;]",raw_to) if x.strip()]
    if not recipients:raise ValueError("RECIPIENTS is empty")
    msg=EmailMessage(); msg["From"]=sender; msg["To"]=", ".join(recipients)
    msg["Subject"]=env("MAIL_SUBJECT","EMAIL_SUBJECT") or f"Assam GAVB Tool Submission Report - {datetime.now():%d-%m-%Y}"
    msg.set_content(env("MAIL_BODY", "EMAIL_BODY") or "Dear Team,\n\nPlease find attached the latest Assam GAVB Tool Submission Report.\n\nRegards,\nGAVB Reporting Automation")
    with path.open("rb") as f:msg.add_attachment(f.read(),maintype="application",subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",filename=path.name)
    logging.info("Email configuration detected: SMTP_EMAIL=%s, SMTP_APP_PASSWORD=%s, RECIPIENTS=%s, MAIL_SUBJECT=%s",bool(sender),bool(password),bool(recipients),bool(msg["Subject"]))
    with smtplib.SMTP_SSL(env("SMTP_HOST") or "smtp.gmail.com",int(env("SMTP_PORT") or "465"),context=ssl.create_default_context(),timeout=60) as smtp:
        smtp.login(sender,password); smtp.send_message(msg)
    logging.info("Email sent successfully to %s recipient(s)",len(recipients))

def main():
    logging.basicConfig(level=env("LOG_LEVEL").upper() or "INFO",format="%(asctime)s [%(levelname)s] %(message)s")
    try:
        logging.info("Running report script version: %s",SCRIPT_VERSION)
        folder=env("DRIVE_FOLDER_ID",required=True); out=Path(env("LOCAL_OUTPUT_FOLDER") or f"/tmp/{datetime.now(timezone.utc):%Y-%m-%d}/state_reports")
        service=build("drive","v3",credentials=credentials(),cache_discovery=False); frames,logs=load_all(service,folder)
        report=create_report(frames,logs,out); send_email(report); print(f"Created and emailed Assam workbook: {report}"); return 0
    except (ValueError,FileNotFoundError) as e:logging.error("Configuration/runtime error: %s",e); return 2
    except HttpError as e:logging.error("Google API request failed: %s",e); return 3
    except smtplib.SMTPAuthenticationError:logging.exception("Gmail authentication failed. Check SMTP_EMAIL and SMTP_APP_PASSWORD"); return 4
    except (smtplib.SMTPException,OSError) as e:logging.exception("Email delivery failed: %s",e); return 5
    except Exception as e:logging.exception("Unexpected failure: %s",e); return 1
if __name__=="__main__":sys.exit(main())

