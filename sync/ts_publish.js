// Merge snapshot JSON files into the Timesheet Report repo (data/weeks/<sunday>.json + data/index.json).
// API mode:   GH_TOKEN=... node ts_publish.js OUTDIR
// Local mode: node ts_publish.js OUTDIR --local REPODIR
const fs=require("fs"),path=require("path");
const REPO="mattsbaker1980-dev/LPH-Timesheet-Report", BRANCH="main";
const [outDir,flag,localDir]=process.argv.slice(2);
const local=flag==="--local";
const tok=process.env.GH_TOKEN;
if(!local&&!tok){ console.log("ERROR no GH_TOKEN"); process.exit(1); }
const H={Authorization:"Bearer "+tok,Accept:"application/vnd.github+json","User-Agent":"lph-timesheet"};
async function getFile(p){
  if(local){ const f=path.join(localDir,p); return fs.existsSync(f)?{json:JSON.parse(fs.readFileSync(f,"utf8")),sha:null}:null; }
  const r=await fetch(`https://api.github.com/repos/${REPO}/contents/${p}?ref=${BRANCH}&cb=${Date.now()}`,{headers:H});
  if(r.status===404) return null; if(!r.ok) throw new Error("GET "+p+" "+r.status+" "+await r.text());
  const j=await r.json(); return {json:JSON.parse(Buffer.from(j.content,"base64").toString("utf8")),sha:j.sha};
}
async function putFile(p,obj,sha,msg){
  const body=JSON.stringify(obj);
  if(local){ const f=path.join(localDir,p); fs.mkdirSync(path.dirname(f),{recursive:true}); fs.writeFileSync(f,body); return; }
  const r=await fetch(`https://api.github.com/repos/${REPO}/contents/${p}`,{method:"PUT",headers:H,body:JSON.stringify({message:msg,branch:BRANCH,content:Buffer.from(body).toString("base64"),...(sha?{sha}:{})})});
  if(!r.ok) throw new Error("PUT "+p+" "+r.status+" "+await r.text());
}
const sunday=s=>{ const [y,m,d]=s.split("-").map(Number); const t=new Date(Date.UTC(y,m-1,d)); t.setUTCDate(t.getUTCDate()-t.getUTCDay()); return t.toISOString().slice(0,10); };
(async()=>{
  const files=fs.readdirSync(outDir).filter(f=>f.endsWith(".json"));
  const byWeek={};
  for(const f of files){
    const s=JSON.parse(fs.readFileSync(path.join(outDir,f),"utf8"));
    if(!s.range_start||s.range_start!==s.range_end||!Array.isArray(s.employees)) continue; // daily reports only
    s.id=f.replace(/\.json$/,"");
    (byWeek[sunday(s.range_start)]=byWeek[sunday(s.range_start)]||[]).push(s);
  }
  let added=0; const touched=[];
  for(const wk of Object.keys(byWeek).sort()){
    const cur=await getFile(`data/weeks/${wk}.json`);
    const snaps=cur?.json?.snapshots||[]; const have=new Set(snaps.map(s=>s.id));
    const fresh=byWeek[wk].filter(s=>!have.has(s.id));
    if(!fresh.length) continue;
    snaps.push(...fresh); snaps.sort((a,b)=>a.id.localeCompare(b.id));
    await putFile(`data/weeks/${wk}.json`,{week_start:wk,snapshots:snaps},cur?.sha,`Timesheet data: ${fresh.length} report(s) for week of ${wk}`);
    added+=fresh.length; touched.push(wk+" (+"+fresh.length+")");
  }
  const idx=await getFile("data/index.json");
  const weeks=new Set(idx?.json?.weeks||[]); Object.keys(byWeek).forEach(w=>weeks.add(w));
  if(added||!idx) await putFile("data/index.json",{updated_at:new Date().toISOString(),weeks:[...weeks].sort()},idx?.sha,"Timesheet data: update index");
  console.log(added?`ADDED ${added} report(s): ${touched.join(", ")}`:"NOTHING NEW");
})().catch(e=>{ console.log("ERROR "+e.message); process.exit(1); });
