#!/usr/bin/env python3
"""IDD Extension report with auditable date parsing.
Required: DRIVE_FOLDER_ID, GOOGLE_SERVICE_ACCOUNT_JSON (or GOOGLE_SERVICE_ACCOUNT_FILE),
SMTP_EMAIL/GMAIL_USERNAME, SMTP_APP_PASSWORD/GMAIL_APP_PASSWORD,
RECIPIENTS/REPORT_RECIPIENTS.
Dependencies: pandas>=2.0 numpy openpyxl xlsxwriter google-auth google-api-python-client
Optional:
 DATE_DAYFIRST=true (default; false for month-first text)
 REPORT_TIMEZONE=Asia/Kolkata (aware timestamps converted here)
 NAIVE_TIMEZONE=Asia/Kolkata (timezone assumed for timestamps without offsets)
 PREVIOUS_DATE_MODE=calendar (default) or available
 DATE_COLUMN or DATE_COLUMN_MOTHER_0_5 / DATE_COLUMN_MOTHER_6_11 /
 DATE_COLUMN_SEPARATE_CD: exact date column override.
No automatic fallback to start/end times when a submission date column exists.
Blank FI records are retained under [Missing FI Name]. No submission deduplication.
"""
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
from datetime import datetime, timezone, date
from email.message import EmailMessage
from email.utils import getaddresses
from pathlib import Path
import numpy as np
import pandas as pd
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

SCRIPT_VERSION = '2026-10-14-v11-date-audit'
DRIVE_SCOPES = ['https://www.googleapis.com/auth/drive.readonly']
TRAVEL_THRESHOLD_KM = float(os.getenv('WEEKLY_TRAVEL_KM', '10'))
TOOL_FILES = [
 ('Mother 0-5','IDD_Extension_Mother_0_5_WIDE.xlsx','Qgeo-Latitude','Qgeo-Longitude'),
 ('Mother 6-11','IDD Extension Mother 6-11_WIDE.xlsx','QX1-Latitude','QX1-Longitude'),
 ('Separate CD','IDD_Extension_Separate_CD_Tool_WIDE.xlsx','Qgeo-Latitude','Qgeo-Longitude')]
TOOL_NAMES = [t[0] for t in TOOL_FILES]
ALIASES = {
 'state':['STATE','State','state','Cal_STATE','state_name'],
 'fi':['QDC_Name','QDC Name','Field Investigator Name','Investigator','QDC','collector_name'],
 'date':['SubmissionDate','Submission Date','submission_date','SubmissionDateTime','Submission_Time','endtime','EndTime','starttime','StartTime'],
 'cal_awc':['Cal_AWC'],'current_awc':['QAWC_3'],'given_awc':['QAWC_code'],'awc_name':['QAWC_name']}


def env(*names, required=False):
    for name in names:
        value = os.getenv(name,'').strip()
        if value:
            return value
    if required:
        raise ValueError('Missing environment variable: ' + ' or '.join(names))
    return ''


def normalize(value):
    return re.sub(r'[^a-z0-9]+','',str(value).strip().lower())


def clean(series):
    return series.astype('string').str.strip().replace({'':pd.NA,'nan':pd.NA,'None':pd.NA,'<NA>':pd.NA})


def find_column(columns, aliases):
    lookup = {normalize(c):str(c) for c in columns}
    return next((lookup[normalize(a)] for a in aliases if normalize(a) in lookup),None)


def credentials():
    raw = env('GOOGLE_SERVICE_ACCOUNT_JSON')
    if raw:
        return Credentials.from_service_account_info(json.loads(raw),scopes=DRIVE_SCOPES)
    return Credentials.from_service_account_file(env('GOOGLE_SERVICE_ACCOUNT_FILE') or 'credential.json',scopes=DRIVE_SCOPES)


def find_file_id(service, folder_id, filename):
    safe = filename.replace('\\','\\\\').replace("'","\\'")
    files = service.files().list(q=f"'{folder_id}' in parents and trashed=false and name='{safe}'",
        spaces='drive',fields='files(id,name)',pageSize=10).execute().get('files',[])
    if not files:
        raise FileNotFoundError(filename)
    return files[0]['id']


def download(service, file_id):
    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer,service.files().get_media(fileId=file_id))
    done = False
    while not done:
        _,done = downloader.next_chunk()
    return buffer.getvalue()


def read_excel(data, filename):
    preview = pd.read_excel(io.BytesIO(data),header=None,nrows=25,engine='openpyxl')
    targets = {normalize(a) for k in ('state','fi') for a in ALIASES[k]}
    scores = [len({normalize(v) for v in row if pd.notna(v)} & targets) for row in preview.values]
    if not scores or max(scores)==0:
        raise ValueError(f'{filename}: cannot identify header row')
    header = int(np.argmax(scores))
    raw = pd.read_excel(io.BytesIO(data),header=header,engine='openpyxl')
    raw.columns = [str(c).strip() for c in raw.columns]
    raw['__SourceRow'] = np.arange(len(raw)) + header + 2
    return raw,header+1


def parse_one_date(value):
    """Return local naive timestamp, parse status, and ambiguity warning."""
    if pd.isna(value) or str(value).strip().lower() in ('','nan','none','nat','<na>'):
        return pd.NaT,'Blank date',''
    dayfirst = (env('DATE_DAYFIRST') or 'true').lower() not in ('false','0','no')
    warning = ''
    try:
        if isinstance(value,(datetime,date,pd.Timestamp,np.datetime64)):
            ts = pd.Timestamp(value)
            status = 'Native Excel/datetime'
        else:
            text = str(value).strip()
            if re.fullmatch(r'\d{5}(?:\.\d+)?',text):
                number = float(text)
                if not 20000 <= number <= 100000:
                    raise ValueError('Unsupported Excel serial')
                ts = pd.Timestamp('1899-12-30') + pd.to_timedelta(number,unit='D')
                status = 'Excel serial (1900 system)'
            else:
                match = re.match(r'^(\d{1,2})[/-](\d{1,2})[/-](\d{4})(?:\D|$)',text)
                if match and 1 <= int(match[1]) <= 12 and 1 <= int(match[2]) <= 12 and match[1] != match[2]:
                    warning = 'Ambiguous day/month; interpreted ' + ('day-first' if dayfirst else 'month-first')
                # Preserve ISO year-first interpretation independently of dayfirst.
                if re.match(r'^\d{4}[-/]\d{1,2}[-/]\d{1,2}',text):
                    ts = pd.Timestamp(pd.to_datetime(text,yearfirst=True,dayfirst=False,errors='raise'))
                else:
                    ts = pd.Timestamp(pd.to_datetime(text,dayfirst=dayfirst,errors='raise'))
                status = 'Parsed text'
        if pd.isna(ts):
            raise ValueError('Unparseable date')
        local_zone = env('REPORT_TIMEZONE') or 'Asia/Kolkata'
        if ts.tzinfo is None:
            ts = ts.tz_localize(env('NAIVE_TIMEZONE') or local_zone)
        ts = ts.tz_convert(local_zone).tz_localize(None)
        return ts,status,warning
    except (ValueError,TypeError,OverflowError):
        return pd.NaT,'Invalid date','Review original value; excluded from dated counts'


def parse_dates(series):
    return pd.Series([parse_one_date(v)[0] for v in series],index=series.index,dtype='datetime64[ns]')


def awc_identifier(frame, detected):
    result = pd.Series(pd.NA,index=frame.index,dtype='string')
    for key in ('cal_awc','current_awc','given_awc','awc_name'):
        if detected.get(key):
            result = result.fillna(clean(frame[detected[key]]))
    return result.str.replace(r'\.0$','',regex=True).str.strip().str.casefold()


def standardize(raw,tool,filename,latitude_name,longitude_name,header_row):
    detected = {k:find_column(raw.columns,v) for k,v in ALIASES.items()}
    override = env('DATE_COLUMN_'+re.sub(r'[^A-Z0-9]+','_',tool.upper()).strip('_'),'DATE_COLUMN')
    if override:
        if override not in raw.columns:
            raise ValueError(f'{filename}: requested date column {override!r} missing')
        detected['date'] = override
    lat,lon = find_column(raw.columns,[latitude_name]),find_column(raw.columns,[longitude_name])
    missing = [k for k in ('state','fi','date') if not detected[k]]
    if missing or not lat or not lon:
        raise ValueError(f'{filename}: missing required columns: {missing}, latitude={lat}, longitude={lon}')
    if not any(detected[k] for k in ('cal_awc','current_awc','given_awc','awc_name')):
        raise ValueError(f'{filename}: no AWC identifier')
    frame = raw.copy()
    if '__SourceRow' not in frame:
        frame['__SourceRow'] = np.arange(len(frame))+header_row+1
    frame['__Tool'],frame['__File'] = tool,filename
    frame['__State'] = clean(frame[detected['state']])
    frame['__MissingFI'] = clean(frame[detected['fi']]).isna()
    frame['__FI'] = clean(frame[detected['fi']]).fillna('[Missing FI Name]')
    frame['__FIKey'] = frame['__FI'].str.casefold()
    frame['__DateColumn'] = detected['date']
    frame['__RawDate'] = frame[detected['date']].astype('string')
    parsed = [parse_one_date(v) for v in frame[detected['date']]]
    frame['__DateTime'] = pd.Series([v[0] for v in parsed],index=frame.index,dtype='datetime64[ns]')
    frame['__DateStatus'] = [v[1] for v in parsed]
    frame['__DateWarning'] = [v[2] for v in parsed]
    frame['__Date'] = frame['__DateTime'].dt.normalize()
    frame['__AWCKey'] = awc_identifier(frame,detected)
    frame['__Latitude'] = pd.to_numeric(frame[lat],errors='coerce').where(lambda x:x.between(-90,90))
    frame['__Longitude'] = pd.to_numeric(frame[lon],errors='coerce').where(lambda x:x.between(-180,180))
    log = {'Tool':tool,'File':filename,'Rows Read':len(frame),'Header Row':header_row,
        'Date Column Used':detected['date'],'Valid Dates':int(frame['__Date'].notna().sum()),
        'Blank Dates':int(frame['__DateStatus'].eq('Blank date').sum()),
        'Invalid Dates':int(frame['__DateStatus'].eq('Invalid date').sum()),
        'Ambiguous Dates':int(frame['__DateWarning'].str.startswith('Ambiguous').sum()),
        'Missing FI Names':int(frame['__MissingFI'].sum()),'Missing States':int(frame['__State'].isna().sum()),
        'Start Date':frame['__Date'].min(),'Latest Date':frame['__Date'].max(),'Status':'OK'}
    logging.info('%s: date column=%s; valid=%s/%s; ambiguous=%s',tool,detected['date'],log['Valid Dates'],len(frame),log['Ambiguous Dates'])
    return frame,log


def load_tools(service,folder_id):
    frames,logs = {},[]
    for tool,filename,lat,lon in TOOL_FILES:
        raw,header = read_excel(download(service,find_file_id(service,folder_id,filename)),filename)
        frames[tool],log = standardize(raw,tool,filename,lat,lon,header)
        logs.append(log)
    return frames,pd.DataFrame(logs)


def get_states(frames):
    values = {}
    for frame in frames.values():
        for state in frame['__State'].dropna():
            values.setdefault(state.casefold(),state)
    selected = {v.strip().casefold() for v in env('STATES').split(',') if v.strip()}
    excluded = {v.strip().casefold() for v in env('EXCLUDED_STATES').split(',') if v.strip()}
    return sorted([v for k,v in values.items() if (not selected or k in selected) and k not in excluded],key=str.casefold)


def combine_state(frames,state):
    return pd.concat([f.loc[f['__State'].fillna('').str.casefold()==state.casefold()].copy() for f in frames.values()],ignore_index=True,sort=False)


def fi_roster(data):
    return (data.groupby(['__FIKey','__FI']).size().rename('Records').reset_index()
        .sort_values(['__FIKey','Records','__FI'],ascending=[True,False,True])
        .drop_duplicates('__FIKey')[['__FIKey','__FI']].rename(columns={'__FI':'FI Name'}))


def latest_and_previous_dates(data):
    dates = sorted(data['__Date'].dropna().unique())
    if not dates:
        raise ValueError('No valid submission dates; inspect source date values')
    latest = pd.Timestamp(dates[-1])
    mode = env('PREVIOUS_DATE_MODE') or 'calendar'
    if mode not in ('calendar','available'):
        raise ValueError('PREVIOUS_DATE_MODE must be calendar or available')
    previous = pd.Timestamp(dates[-2]) if mode=='available' and len(dates)>1 else latest-pd.Timedelta(days=1)
    return latest,previous


def cumulative_interviews(data,cutoff,roster):
    selected = data.loc[data['__Date'].notna() & data['__Date'].le(cutoff)]
    result = roster.set_index('__FIKey')
    result['Total Interviews'] = selected.groupby('__FIKey').size().reindex(result.index,fill_value=0).astype(int)
    result['Separate CD'] = selected.loc[selected['__Tool'].eq('Separate CD')].groupby('__FIKey').size().reindex(result.index,fill_value=0).astype(int)
    return result


def awc_completion(data,cutoff,roster):
    selected = data.loc[data['__Date'].notna() & data['__Date'].le(cutoff) & data['__AWCKey'].notna()
        & ~data['__MissingFI'] & data['__Tool'].isin(['Mother 0-5','Mother 6-11'])]
    result = roster.set_index('__FIKey')
    presence = selected.drop_duplicates(['__FIKey','__AWCKey','__Tool']).groupby(['__FIKey','__AWCKey'])['__Tool'].nunique()
    for label,mask in [('AWC with Two Tools Completed',presence.eq(2)),('AWC with Less Than Two Tools',presence.lt(2))]:
        result[label] = mask.groupby(level=0).sum().reindex(result.index,fill_value=0).astype(int)
    return result


def simple_status(data,latest,previous):
    roster = fi_roster(data)
    current,prior = cumulative_interviews(data,latest,roster),cumulative_interviews(data,previous,roster)
    awc = awc_completion(data,latest,roster)
    result = roster.set_index('__FIKey')
    result['Total Submission Till Latest Date'] = current['Total Interviews']
    result['Total Submission Till Previous Date'] = prior['Total Interviews']
    result['New Submission on Latest Date'] = data.loc[data['__Date'].eq(latest)].groupby('__FIKey').size().reindex(result.index,fill_value=0).astype(int)
    for col in ('AWC with Two Tools Completed','AWC with Less Than Two Tools'):
        result[col] = awc[col]
    result['Separate CD Submission'] = current['Separate CD']
    result['Rows with Blank or Invalid Date'] = data.loc[data['__Date'].isna()].groupby('__FIKey').size().reindex(result.index,fill_value=0).astype(int)
    assert result['Total Submission Till Latest Date'].sum()==int(data['__Date'].notna().sum())
    assert (result['Total Submission Till Latest Date']-result['Total Submission Till Previous Date']).equals(result['New Submission on Latest Date'])
    return result.reset_index(drop=True).sort_values('FI Name',key=lambda s:s.str.casefold()).reset_index(drop=True)


def haversine(lat1,lon1,lat2,lon2):
    if any(pd.isna(v) for v in (lat1,lon1,lat2,lon2)):
        return np.nan
    a = math.sin(math.radians(lat2-lat1)/2)**2+math.cos(math.radians(lat1))*math.cos(math.radians(lat2))*math.sin(math.radians(lon2-lon1)/2)**2
    a = min(1,max(0,a))
    return 6371.0088*2*math.atan2(math.sqrt(a),math.sqrt(1-a))


def fi_weekly_travel(data):
    cols = ['FI Name','ISO Year','ISO Week','Week Start','GPS Active Days','Valid Day-to-Day Comparisons','Weekly Travel km','10 km Status']
    gps = data.loc[~data['__MissingFI']].dropna(subset=['__Date','__Latitude','__Longitude']).copy()
    if gps.empty:
        return pd.DataFrame(columns=cols)
    display = fi_roster(gps).set_index('__FIKey')['FI Name']
    daily = gps.groupby(['__FIKey','__Date'],as_index=False).agg(Latitude=('__Latitude','median'),Longitude=('__Longitude','median')).sort_values(['__FIKey','__Date'])
    daily['PrevLat'] = daily.groupby('__FIKey')['Latitude'].shift()
    daily['PrevLon'] = daily.groupby('__FIKey')['Longitude'].shift()
    daily['Distance'] = daily.apply(lambda r:haversine(r['PrevLat'],r['PrevLon'],r['Latitude'],r['Longitude']),axis=1)
    iso = daily['__Date'].dt.isocalendar()
    daily['ISO Year'],daily['ISO Week'] = iso.year.astype(int),iso.week.astype(int)
    daily['Week Start'] = daily['__Date']-pd.to_timedelta(daily['__Date'].dt.weekday,unit='D')
    weekly = daily.groupby(['__FIKey','ISO Year','ISO Week','Week Start'],as_index=False).agg(**{
        'Weekly Travel km':('Distance','sum'),'Valid Day-to-Day Comparisons':('Distance','count'),'GPS Active Days':('__Date','nunique')})
    weekly['FI Name'] = weekly['__FIKey'].map(display)
    weekly['10 km Status'] = np.select([weekly['Valid Day-to-Day Comparisons'].eq(0),weekly['Weekly Travel km'].ge(TRAVEL_THRESHOLD_KM)],['Not enough GPS days','Yes'],default='No')
    return weekly[cols].sort_values(['Week Start','FI Name'],ascending=[False,True])


def safe_filename(value):
    return re.sub(r'[<>:"/\\|?*]+','_',str(value)).strip() or 'Unknown_State'


def write_sheet(writer,name,frame,title=None):
    start = 2 if title else 0
    frame.to_excel(writer,sheet_name=name,index=False,startrow=start)
    ws,book = writer.sheets[name],writer.book
    if title:
        ws.write(0,0,title,book.add_format({'bold':True,'font_size':14}))
    header = book.add_format({'bold':True,'bg_color':'#B7DEE8','border':1,'text_wrap':True,'valign':'vcenter'})
    for i,col in enumerate(frame.columns):
        ws.write(start,i,str(col),header)
        fmt = book.add_format({'num_format':'dd-mm-yyyy hh:mm:ss'}) if pd.api.types.is_datetime64_any_dtype(frame[col]) else None
        ws.set_column(i,i,min(42,max(16,len(str(col))+2)),fmt)
    ws.set_row(start,42)
    ws.freeze_panes(start+1,1)
    if len(frame.columns):
        ws.autofilter(start,0,start+len(frame),len(frame.columns)-1)


def audit_rows(data,latest=None,previous=None):
    mapping = {'__File':'Source File','__SourceRow':'Excel Row','__Tool':'Tool','__State':'State','__FI':'FI Name',
        '__MissingFI':'Missing FI Name','__DateColumn':'Date Column Used','__RawDate':'Original Date Value',
        '__DateTime':'Parsed Local DateTime','__Date':'Parsed Date','__DateStatus':'Parse Status','__DateWarning':'Review Note'}
    result = data[list(mapping)].rename(columns=mapping).copy()
    if latest is not None:
        result['Included Till Latest'] = data['__Date'].notna() & data['__Date'].le(latest)
        result['Included Till Previous'] = data['__Date'].notna() & data['__Date'].le(previous)
        result['On Latest Date'] = data['__Date'].eq(latest)
    return result


def create_workbook(state,data,logs,output_folder):
    latest,previous = latest_and_previous_dates(data)
    summary = simple_status(data,latest,previous)
    path = output_folder/f'{safe_filename(state).upper()}_IDD_Extension_Simple_Summary_{datetime.now():%d-%m-%Y}.xlsx'
    daily = data.loc[data['__Date'].notna()].groupby(['__Date','__Tool'],dropna=False).size().unstack(fill_value=0).reindex(columns=TOOL_NAMES,fill_value=0)
    daily['Total Submissions'] = daily.sum(axis=1)
    daily = daily.reset_index().rename(columns={'__Date':'Submission Date'})
    checks = []
    for tool in TOOL_NAMES:
        f = data.loc[data['__Tool'].eq(tool)]
        valid = int(f['__Date'].notna().sum())
        checks.append({'Tool':tool,'Rows Read':len(f),'Valid Dated Rows':valid,
            'Blank or Invalid Date':len(f)-valid,'Missing FI Names':int(f['__MissingFI'].sum()),
            'Till Previous Date':int(f['__Date'].le(previous).sum()),'On Latest Date':int(f['__Date'].eq(latest).sum()),
            'Till Latest Date':int(f['__Date'].le(latest).sum()),
            'Ambiguous Dates':int(f['__DateWarning'].str.startswith('Ambiguous').sum())})
    settings = pd.DataFrame({'Setting':['Latest date','Previous cutoff','Previous mode','Text day-first','Report timezone','Naive timezone','Counting rule','Missing FI rule','AWC missing FI rule','Serial date rule','Date selection rule'],
        'Value':[str(latest.date()),str(previous.date()),env('PREVIOUS_DATE_MODE') or 'calendar',env('DATE_DAYFIRST') or 'true',
        env('REPORT_TIMEZONE') or 'Asia/Kolkata',env('NAIVE_TIMEZONE') or 'Asia/Kolkata',
        'Every source row counts once; no deduplication; invalid dates excluded from dated totals',
        'Included under [Missing FI Name]','Not attributed to an FI for AWC completion',
        'Numeric serials assume Excel 1900 date system','First matching date alias or exact environment override; no per-row fallback']})
    audit = audit_rows(data,latest,previous)
    issues = audit.loc[data['__Date'].isna() | data['__MissingFI'] | data['__DateWarning'].ne('')]
    with pd.ExcelWriter(path,engine='xlsxwriter',engine_kwargs={'options':{'strings_to_urls':False,'strings_to_formulas':False}}) as writer:
        write_sheet(writer,'FI Summary',summary,f'{state} | through {latest:%d-%m-%Y}; previous {previous:%d-%m-%Y}')
        write_sheet(writer,'FI Weekly 10km Travel',fi_weekly_travel(data),f'Weekly GPS displacement | {state} | threshold {TRAVEL_THRESHOLD_KM:g} km')
        write_sheet(writer,'Daily Submission Counts',daily)
        write_sheet(writer,'Count Reconciliation',pd.DataFrame(checks))
        write_sheet(writer,'Date Audit',audit)
        write_sheet(writer,'Date Issues',issues)
        write_sheet(writer,'Report Settings',settings)
        write_sheet(writer,'Processing Log',logs.assign(State_Workbook=state,Latest_Database_Date=latest,Previous_Cutoff_Date=previous,FI_Count=len(summary)))
        for tool in TOOL_NAMES:
            raw = data.loc[data['__Tool'].eq(tool)].drop(columns=[c for c in data if c.startswith('__')])
            write_sheet(writer,('Raw '+tool)[:31],raw)
    return path


def create_zip(paths,output_folder):
    path = output_folder/f'IDD_Extension_Simple_Summary_All_States_{datetime.now():%d-%m-%Y}.zip'
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED,compresslevel=6) as archive:
        for item in paths:
            archive.write(item,arcname=item.name)
    return path


def parse_recipients(raw):
    recipients = []
    for _,address in getaddresses([raw.replace(';',',').replace('\r','').replace('\n',',')]):
        if not re.fullmatch(r'[^\s@<>;,]+@[^\s@<>;,]+\.[^\s@<>;,]+',address):
            raise ValueError('Invalid recipient address; use complete comma-separated email addresses')
        if address.casefold() not in {v.casefold() for v in recipients}:
            recipients.append(address)
    if not recipients:
        raise ValueError('No recipients')
    return recipients


def send_email(attachment,states):
    sender = env('SMTP_EMAIL','GMAIL_USERNAME',required=True)
    password = env('SMTP_APP_PASSWORD','GMAIL_APP_PASSWORD',required=True)
    recipients = parse_recipients(env('RECIPIENTS','REPORT_RECIPIENTS',required=True))
    if attachment.stat().st_size/(1024*1024)>float(env('MAX_EMAIL_ATTACHMENT_MB') or '24'):
        raise ValueError('ZIP exceeds configured attachment size limit')
    message = EmailMessage()
    message['From'],message['To'] = sender,', '.join(recipients)
    message['Subject'] = env('MAIL_SUBJECT') or f'IDD Extension Simple Summary - {datetime.now():%d-%m-%Y}'
    message.set_content(f"Dear Team,\n\nPlease find attached the IDD Extension FI summary for: {', '.join(states)}.\nDate audit and count reconciliation are included.\n\nRegards,\nGAVB Reporting Automation")
    message.add_attachment(attachment.read_bytes(),maintype='application',subtype='zip',filename=attachment.name)
    with smtplib.SMTP_SSL('smtp.gmail.com',465,context=ssl.create_default_context(),timeout=60) as smtp:
        smtp.login(sender,password)
        refused = smtp.send_message(message,from_addr=sender,to_addrs=recipients)
    if refused:
        raise smtplib.SMTPException('Some recipients refused; do not resend to accepted recipients: '+', '.join(refused))
    logging.info('SMTP accepted message for %s recipients; inbox delivery not guaranteed',len(recipients))


def main():
    logging.basicConfig(level=(env('LOG_LEVEL') or 'INFO').upper(),format='%(asctime)s [%(levelname)s] %(message)s')
    try:
        logging.info('Script version %s',SCRIPT_VERSION)
        output = Path(env('LOCAL_OUTPUT_FOLDER') or f'/tmp/{datetime.now(timezone.utc):%Y-%m-%d}/idd_extension_simple')
        output.mkdir(parents=True,exist_ok=True)
        service = build('drive','v3',credentials=credentials(),cache_discovery=False)
        frames,logs = load_tools(service,env('DRIVE_FOLDER_ID',required=True))
        paths,states = [],[]
        # Blank states cannot be assigned to a state workbook; preserve in a separate audit.
        unassigned = pd.concat([f.loc[f['__State'].isna()] for f in frames.values()],ignore_index=True,sort=False)
        if not unassigned.empty:
            audit_path = output/'IDD_Unassigned_State_Audit.xlsx'
            with pd.ExcelWriter(audit_path,engine='xlsxwriter',engine_kwargs={'options':{'strings_to_formulas':False,'strings_to_urls':False}}) as writer:
                write_sheet(writer,'Unassigned State Rows',audit_rows(unassigned))
            paths.append(audit_path)
            logging.warning('%s rows have no state; included in separate audit',len(unassigned))
        for state in get_states(frames):
            data = combine_state(frames,state)
            if data['__Date'].notna().any():
                paths.append(create_workbook(state,data,logs,output))
                states.append(state)
            else:
                # Preserve failures for inspection rather than silently omitting the state.
                path = output/f'{safe_filename(state).upper()}_DATE_ERROR.xlsx'
                with pd.ExcelWriter(path,engine='xlsxwriter',engine_kwargs={'options':{'strings_to_formulas':False,'strings_to_urls':False}}) as writer:
                    write_sheet(writer,'Date Issues',audit_rows(data))
                paths.append(path)
                states.append(state+' (no valid dates; audit only)')
        if not paths:
            raise ValueError('No reports created')
        attachment = create_zip(paths,output)
        send_email(attachment,states)
        print('Created reports:',*[str(p) for p in paths],sep='\n')
        return 0
    except Exception:
        logging.exception('Report failed')
        return 1


if __name__=='__main__':
    sys.exit(main())
