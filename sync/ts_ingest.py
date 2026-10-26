# Usage: python3 -I ts_ingest.py OUTDIR RAWJSON [RAWJSON ...]
# RAWJSON = saved result of Gmail get_message(messageFormat=RAW).
# Handles two attachment layouts:
#   "daily"  - ServiceTitan Timesheet Report (Technician Performance): one row per tech, Total Paid Time, Filters tab with the date range.
#   "detail" - legacy Technician Timesheet Summary for all employees: one row per employee per day with Start/End/Meal/Vac/Non-PTO/Hol.
# Writes one snapshot JSON per attachment into OUTDIR and prints "<doc_id> <summary>".
import sys,json,base64,email,io,re,os,datetime,warnings
from zoneinfo import ZoneInfo
warnings.filterwarnings("ignore")
import openpyxl
ET=ZoneInfo("America/New_York")
def dept(*vals):
    for v in vals:
        u=str(v or "").lower()
        if not u: continue
        if re.search(r"sewer|drain",u): return "Sewer"
        if "plumb" in u: return "Plumbing"
        if "elec" in u: return "Electrical"
        if re.search(r"hvac|heat|cool|air|furnace",u): return "HVAC"
    return "Other"
def pdate(s):
    m=re.match(r"(\d{1,2})/(\d{1,2})/(\d{2,4})",s.strip()); y=int(m.group(3)); y+=2000 if y<100 else 0
    return datetime.date(y,int(m.group(1)),int(m.group(2))).isoformat()
def mins(v):
    if v in (None,""): return 0
    if isinstance(v,datetime.timedelta): return round(v.total_seconds()/60)
    if isinstance(v,(int,float)): return round(v*60)
    s=str(v).strip()
    m=re.match(r"^(\d+):(\d{2})(?::\d{2})?$",s)
    if m: return int(m.group(1))*60+int(m.group(2))
    try: return round(float(s.replace(",",""))*60)
    except: return 0
def hhmm(v):
    if isinstance(v,datetime.datetime): v=v.time()
    if isinstance(v,datetime.time): return f"{v.hour}:{v.minute:02d}"
    if isinstance(v,(int,float)) and 0<=v<1: t=round(v*1440); return f"{t//60}:{t%60:02d}"
    s=str(v or "").strip(); m=re.match(r"(\d{1,2}):(\d{2})",s)
    return f"{int(m.group(1))}:{m.group(2)}" if m else ""
def tmin(s): h,m=s.split(":"); return int(h)*60+int(m)
def ddate(v):
    if isinstance(v,datetime.datetime): return v.date().isoformat()
    if isinstance(v,datetime.date): return v.isoformat()
    if isinstance(v,str) and re.match(r"\d{1,2}/\d{1,2}/\d{2,4}",v.strip()): return pdate(v)
    return None
def header(rows,test):
    for i,r in enumerate(rows[:15]):
        h=[str(x or "").strip().lower() for x in r]
        if test(h): return i,h
    return None,None
def parse_daily(wb,rows,captured_iso,source):
    hi,h=header(rows,lambda h: any(re.search(r"total paid time|total hours worked",t) for t in h) and any(re.fullmatch(r"name|technician( name)?|employee( name)?",t) for t in h))
    if hi is None: return None
    c=dict(name=next(j for j,t in enumerate(h) if re.fullmatch(r"name|technician( name)?|employee( name)?",t)),
           hours=next(j for j,t in enumerate(h) if re.search(r"total paid time|total hours worked",t)),
           trade=next((j for j,t in enumerate(h) if "trade" in t),-1), team=h.index("team") if "team" in h else -1,
           unit=next((j for j,t in enumerate(h) if t.endswith("business unit")),-1), id=next((j for j,t in enumerate(h) if t in("technician id","employee id")),-1))
    emps=[]
    for r in rows[hi+1:]:
        name=str(r[c["name"]] or "").strip()
        if not name or re.fullmatch(r"\d+",name) or re.match(r"(office|warehouse|parking lot)",name,re.I): continue
        g=lambda j: str(r[j] or "").strip() if j>=0 else ""
        emps.append(dict(name=name,id=g(c["id"]),unit=g(c["unit"]),team=g(c["team"]),dept=dept(g(c["trade"]),g(c["team"]),g(c["unit"])),min=mins(r[c["hours"]])))
    rs=re_=None
    for w in wb.worksheets:
        for r in w.iter_rows(values_only=True):
            if r and re.search(r"date range|from - to",str(r[0] or ""),re.I) and len(r)>1 and r[1]:
                p=re.split(r"\s+[-–]\s+",str(r[1]).strip()); rs=pdate(p[0]); re_=pdate(p[-1])
    if not rs: rs=re_=captured_iso[:10]
    return dict(kind="daily",captured_at=captured_iso,range_start=rs,range_end=re_,source=source,employees=emps)
def parse_detail(wb,captured_iso,source):
    ws=next((w for w in wb.worksheets if "timesheet" in w.title.lower()),wb.worksheets[0])
    rows=list(ws.iter_rows(values_only=True))
    hi,h=header(rows,lambda h: "start" in h and "end" in h and "hours worked" in h and "date" in h)
    if hi is None: return None
    col=lambda *names: next((h.index(n) for n in names if n in h),-1)
    c=dict(id=col("employee id"),name=col("employee","employee name","technician"),date=col("date"),start=col("start"),end=col("end"),
           meal=col("meal"),hw=col("hours worked"),reg=col("regular"),ot=col("overtime"),vac=col("vac","vacation"),
           nonpto=col("non-pto","non pto"),nonpto2=col("non pto") if "non-pto" in h else -1,hol=col("hol","holiday"),hol2=col("hol1"),pto=col("paid time off"),train=col("training"),train2=col("office training"))
    g=lambda r,k: r[c[k]] if c[k]>=0 else None
    out=[]
    for r in rows[hi+1:]:
        if not r or c["name"]<0: continue
        name=str(g(r,"name") or "").strip(); d=ddate(g(r,"date"))
        if not name or not d: continue
        out.append(dict(id=str(g(r,"id") or ""),name=name,date=d,start=hhmm(g(r,"start")),end=hhmm(g(r,"end")),meal=mins(g(r,"meal")),
            hw=mins(g(r,"hw")),reg=mins(g(r,"reg")),ot=mins(g(r,"ot")),vac=mins(g(r,"vac"))+mins(g(r,"pto")),
            nonpto=mins(g(r,"nonpto"))+mins(g(r,"nonpto2")),hol=mins(g(r,"hol"))+mins(g(r,"hol2")),train=mins(g(r,"train"))+mins(g(r,"train2"))))
    if not out: return None
    run_date=datetime.datetime.fromisoformat(captured_iso.replace("Z","+00:00")).astimezone(ET).date().isoformat()
    # The export time shows up as the End of everyone still on the clock: take the most common End on the run date.
    ends=[x["end"] for x in out if x["date"]==run_date and x["start"] and x["end"] and not x["vac"]]
    run_time=max(set(ends),key=ends.count) if ends else ""
    for x in out:
        x["booked"]=x["date"]>run_date
        x["open"]=bool(run_time and x["date"]==run_date and x["start"] and x["end"] and not x["vac"] and abs(tmin(x["end"])-tmin(run_time))<=5)
    dates=sorted(x["date"] for x in out)
    return dict(kind="detail",captured_at=captured_iso,run_date=run_date,run_time=run_time,range_start=dates[0],range_end=dates[-1],source=source,rows=out)
out=sys.argv[1]; os.makedirs(out,exist_ok=True)
for f in sys.argv[2:]:
    d=json.load(open(f)); raw=d["raw"]
    m=email.message_from_bytes(base64.urlsafe_b64decode(raw+"="*(-len(raw)%4)))
    # report time, rounded to the minute so duplicate copies of one email collapse
    t=datetime.datetime.fromtimestamp(int(d["internalDate"])/1000,datetime.timezone.utc).replace(second=0,microsecond=0)
    cap=t.strftime("%Y-%m-%dT%H:%M:00.000Z"); stamp=cap.replace(":","-").replace(".","-")
    for p in m.walk():
        fn=p.get_filename()
        if not (fn and fn.lower().endswith(".xlsx")): continue
        wb=openpyxl.load_workbook(io.BytesIO(p.get_payload(decode=True)),data_only=True)
        rows=list(wb.worksheets[0].iter_rows(values_only=True))
        s=parse_detail(wb,cap,fn) or parse_daily(wb,rows,cap,fn)
        if not s: print("SKIP unrecognized attachment",fn); continue
        if s["kind"]=="detail":
            doc="detail_"+s["range_start"]+"_"+stamp
            print(doc,len({r["name"] for r in s["rows"]}),"employees, run",s["run_date"],s["run_time"])
        else:
            doc=s["range_start"]+"_"+s["range_end"]+"_"+stamp
            print(doc,len([e for e in s["employees"] if e["min"]>0]),"with hours")
        json.dump(s,open(os.path.join(out,doc+".json"),"w"))
