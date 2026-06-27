import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── Supabase config ──────────────────────────────────────────────────────────
const SUPA_URL = "https://qfywqmysopmgchncunsr.supabase.co";
const SUPA_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InFmeXdxbXlzb3BtZ2NobmN1bnNyIiwicm9sZSI6ImFub24iLCJpYXQiOjE3ODI0NDc2NzcsImV4cCI6MjA5ODAyMzY3N30.i2wX6vx1UJbtcV-UZGF2AhhmvS-GQSrlIsQXpt-1KvI";

const supa = {
  async get(table, filters = "") {
    const r = await fetch(`${SUPA_URL}/rest/v1/${table}?${filters}`, {
      headers: { apikey: SUPA_KEY, Authorization: `Bearer ${SUPA_KEY}` }
    });
    return r.json();
  },
  async upsert(table, data) {
    const r = await fetch(`${SUPA_URL}/rest/v1/${table}`, {
      method: "POST",
      headers: {
        apikey: SUPA_KEY, Authorization: `Bearer ${SUPA_KEY}`,
        "Content-Type": "application/json", Prefer: "resolution=merge-duplicates,return=representation"
      },
      body: JSON.stringify(data)
    });
    return r.json();
  },
  async del(table, filter) {
    await fetch(`${SUPA_URL}/rest/v1/${table}?${filter}`, {
      method: "DELETE",
      headers: { apikey: SUPA_KEY, Authorization: `Bearer ${SUPA_KEY}` }
    });
  }
};

// ─── Rate limiter ─────────────────────────────────────────────────────────────
const rl = (() => {
  const w = {};
  return (key, max = 10, ms = 60000) => {
    const now = Date.now();
    w[key] = (w[key] || []).filter(t => now - t < ms);
    if (w[key].length >= max) return false;
    w[key].push(now); return true;
  };
})();

const clean = s => String(s).replace(/[<>"'`]/g, "").trim().slice(0, 300);

// ─── Constants ────────────────────────────────────────────────────────────────
const GL = ["A","A-","B+","B","B-","C+","C","C-","D+","D","D-","E"];
const GP = { A:12,"A-":11,"B+":10,B:9,"B-":8,"C+":7,C:6,"C-":5,"D+":4,D:3,"D-":2,E:1 };

const SUBJECTS = [
  { id:"math",  label:"Mathematics",      papers:["pp1","pp2"],       compulsory:true },
  { id:"eng",   label:"English",          papers:["pp1","pp2","pp3"], compulsory:true },
  { id:"swa",   label:"Kiswahili",        papers:["pp1","pp2","pp3"], compulsory:true },
  { id:"bio",   label:"Biology",          papers:["pp1","pp2","pp3"], compulsory:true },
  { id:"chem",  label:"Chemistry",        papers:["pp1","pp2","pp3"], compulsory:true },
  { id:"phyc",  label:"Physics",          papers:["pp1","pp2","pp3"], compulsory:false, group:"A" },
  { id:"geo",   label:"Geography",        papers:["pp1","pp2"],       compulsory:false, group:"A" },
  { id:"hist",  label:"History",          papers:["pp1","pp2"],       compulsory:false, group:"A" },
  { id:"cre",   label:"C.R.E",           papers:["pp1","pp2"],       compulsory:false, group:"A" },
  { id:"bst",   label:"Business Studies", papers:["pp1","pp2"],       compulsory:false, group:"B" },
  { id:"agric", label:"Agriculture",      papers:["pp1","pp2"],       compulsory:false, group:"B" },
  { id:"comp",  label:"Computer Studies", papers:["pp1","pp2"],       compulsory:false, group:"B" },
];
const SM = Object.fromEntries(SUBJECTS.map(s => [s.id, s]));

const defGB = () => {
  const d = { A:75,"A-":70,"B+":65,B:60,"B-":55,"C+":50,C:45,"C-":40,"D+":35,D:30,"D-":25,E:0 };
  const b = {}; SUBJECTS.forEach(s => { b[s.id] = { ...d }; }); return b;
};
const defTB = () => ({ A:73,"A-":61,"B+":53,B:47,"B-":41,"C+":35,C:29,"C-":23,"D+":17,D:11,"D-":5,E:0 });
const defFM = () => { const f = {}; SUBJECTS.forEach(s => { f[s.id] = s.papers.length===3?"pp1+pp2+pp3":"pp1+pp2"; }); return f; };

// ─── Helpers ──────────────────────────────────────────────────────────────────
function evalFm(fm, marks) {
  try {
    const e = fm.replace(/pp1/gi, Number(marks?.pp1)||0).replace(/pp2/gi, Number(marks?.pp2)||0).replace(/pp3/gi, Number(marks?.pp3)||0);
    const r = Function('"use strict";return('+e+')')();
    return isNaN(r) ? 0 : Math.round(r*100)/100;
  } catch { return 0; }
}
function getGrade(total, bounds) {
  for (const g of GL) if (total >= (bounds[g]??0)) return g;
  return "E";
}
function validateChoice(chosen) {
  const A = chosen.filter(id => ["phyc","geo","hist","cre"].includes(id));
  const B = chosen.filter(id => ["bst","agric","comp"].includes(id));
  const e = [];
  if (A.length<2) e.push("Choose at least 2 from Group A.");
  if (A.length>3) e.push("Choose at most 3 from Group A.");
  if (A.includes("geo")&&A.includes("hist")) e.push("Geography and History cannot be combined.");
  if (!A.includes("phyc")&&!A.includes("cre")) e.push("At least one of Physics or CRE is required.");
  if (B.length>1) e.push("At most 1 subject from Group B.");
  return e;
}
function computeResults(student, gb, tb, fm) {
  const GB = gb||defGB(), TB = tb||defTB(), FM = fm||defFM();
  const compIds = SUBJECTS.filter(s=>s.compulsory).map(s=>s.id);
  const allIds = [...compIds, ...(student.subjects||[])];
  const res = {};
  allIds.forEach(id => {
    const total = evalFm(FM[id]||"pp1+pp2", student.marks?.[id]||{});
    const grade = getGrade(total, GB[id]||{});
    res[id] = { total, grade, points: GP[grade]||1 };
  });
  const eP = res["eng"]?.points||0, sP = res["swa"]?.points||0;
  const langUsed = eP>=sP?"eng":"swa";
  const nonLang = allIds.filter(id=>id!=="eng"&&id!=="swa");
  const top6 = [...nonLang].sort((a,b)=>(res[b]?.points||0)-(res[a]?.points||0)).slice(0,6);
  const totalPts = (res[langUsed]?.points||0)+top6.reduce((s,id)=>s+(res[id]?.points||0),0);
  return { res, langUsed, totalPts, totalGrade: getGrade(totalPts,TB), top5: top6, allIds };
}
const gColor = g => ["A","A-"].includes(g)?"#15803d":["B+","B","B-"].includes(g)?"#1e40af":["C+","C","C-"].includes(g)?"#b45309":"#b91c1c";

// ─── UI Bits ──────────────────────────────────────────────────────────────────
const KenyaFlag = ({ size=32 }) => (
  <svg width={size} height={Math.round(size*0.67)} viewBox="0 0 90 60" style={{borderRadius:3,boxShadow:"0 1px 4px #0004",flexShrink:0}}>
    <rect width="90" height="60" fill="#006600"/>
    <rect y="15" width="90" height="30" fill="#BB0000"/>
    <rect y="22" width="90" height="16" fill="#000"/>
    <rect y="24" width="90" height="12" fill="#fff"/>
    <rect y="26" width="90" height="8" fill="#000"/>
    <ellipse cx="45" cy="30" rx="10" ry="14" fill="#BB0000" stroke="#fff" strokeWidth="1.5"/>
    <line x1="45" y1="16" x2="45" y2="44" stroke="#fff" strokeWidth="2"/>
  </svg>
);

const BG = () => (
  <div style={{position:"fixed",inset:0,zIndex:0,overflow:"hidden",pointerEvents:"none"}}>
    <div style={{position:"absolute",inset:0,background:"linear-gradient(160deg,#003d1a 0%,#0a2a4a 45%,#1a0a00 100%)"}}/>
    <svg style={{position:"absolute",inset:0,width:"100%",height:"100%",opacity:0.05}}>
      <defs><pattern id="pg" width="60" height="60" patternUnits="userSpaceOnUse"><path d="M 60 0 L 0 0 0 60" fill="none" stroke="#fff" strokeWidth="0.5"/></pattern></defs>
      <rect width="100%" height="100%" fill="url(#pg)"/>
    </svg>
    <div style={{position:"absolute",top:-120,right:-120,width:450,height:450,borderRadius:"50%",background:"rgba(187,0,0,0.07)"}}/>
    <div style={{position:"absolute",bottom:-100,left:-100,width:380,height:380,borderRadius:"50%",background:"rgba(0,102,0,0.09)"}}/>
  </div>
);

function Toast({msg,type,onClose}){
  useEffect(()=>{const t=setTimeout(onClose,3500);return()=>clearTimeout(t);},[onClose]);
  return <div style={{position:"fixed",bottom:24,right:24,zIndex:9999,background:type==="error"?"#b91c1c":"#15803d",color:"#fff",borderRadius:10,padding:"12px 20px",fontWeight:600,fontSize:14,boxShadow:"0 4px 20px #0005"}}>{msg}</div>;
}

function Spinner(){return <div style={{display:"flex",justifyContent:"center",padding:40}}><div style={{width:36,height:36,border:"4px solid #e2e8f0",borderTop:"4px solid #c8102e",borderRadius:"50%",animation:"spin 1s linear infinite"}}/></div>;}

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [page, setPage] = useState(()=>window.location.hash==="#/superadmin"?"admin":"results");
  useEffect(()=>{
    const h=()=>setPage(window.location.hash==="#/superadmin"?"admin":"results");
    window.addEventListener("hashchange",h); return()=>window.removeEventListener("hashchange",h);
  },[]);

  const [toast, setToast] = useState(null);
  const showToast = useCallback((msg,type="success")=>setToast({msg,type}),[]);

  const [authed,setAuthed] = useState(()=>sessionStorage.getItem("kcse_auth")==="1");
  const [authPw,setAuthPw] = useState("");
  const [authErr,setAuthErr] = useState("");
  const ADMIN_PW = "Admin@2025";

  const handleLogin = () => {
    if(!rl("login",5,60000)){setAuthErr("Too many attempts. Wait 1 minute.");return;}
    if(authPw===ADMIN_PW){setAuthed(true);sessionStorage.setItem("kcse_auth","1");setAuthErr("");}
    else setAuthErr("Incorrect password.");
  };

  return (
    <div style={{minHeight:"100vh",position:"relative",fontFamily:"'Segoe UI',system-ui,sans-serif"}}>
      <style>{`@keyframes spin{to{transform:rotate(360deg)}}@keyframes fadeIn{from{opacity:0;transform:translateY(10px)}to{opacity:1;transform:translateY(0)}}`}</style>
      <BG/>
      {toast&&<Toast msg={toast.msg} type={toast.type} onClose={()=>setToast(null)}/>}
      {page==="results"&&<ResultsPage showToast={showToast}/>}
      {page==="admin"&&!authed&&<LoginPage pw={authPw} setPw={setAuthPw} err={authErr} onLogin={handleLogin} onBack={()=>{window.location.hash="";}}/>}
      {page==="admin"&&authed&&<AdminPage showToast={showToast} onLogout={()=>{setAuthed(false);sessionStorage.removeItem("kcse_auth");window.location.hash="";}}/>}
    </div>
  );
}

// ─── Login ────────────────────────────────────────────────────────────────────
function LoginPage({pw,setPw,err,onLogin,onBack}){
  return(
    <div style={{position:"relative",zIndex:1,minHeight:"100vh",display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:"rgba(255,255,255,0.97)",borderRadius:18,padding:44,width:360,boxShadow:"0 12px 60px #0009",animation:"fadeIn .4s ease"}}>
        <div style={{textAlign:"center",marginBottom:28}}><KenyaFlag size={48}/>
          <div style={{fontSize:22,fontWeight:900,color:"#0a2a4a",marginTop:14}}>Superadmin Portal</div>
          <div style={{fontSize:13,color:"#64748b",marginTop:4}}>KCSE Results Management System</div>
        </div>
        {err&&<div style={{background:"#fee2e2",color:"#b91c1c",borderRadius:7,padding:"8px 12px",marginBottom:12,fontSize:13}}>⚠ {err}</div>}
        <label style={{fontSize:11,fontWeight:700,color:"#475569",display:"block",marginBottom:4,textTransform:"uppercase"}}>Password</label>
        <input type="password" value={pw} onChange={e=>setPw(e.target.value)} onKeyDown={e=>e.key==="Enter"&&onLogin()}
          style={{width:"100%",padding:"11px 13px",border:"2px solid #e2e8f0",borderRadius:8,fontSize:15,boxSizing:"border-box",marginBottom:16}}/>
        <button onClick={onLogin} style={{width:"100%",background:"#c8102e",color:"#fff",border:"none",borderRadius:9,padding:13,fontWeight:700,fontSize:16,cursor:"pointer"}}>Sign In</button>
        <button onClick={onBack} style={{width:"100%",background:"transparent",color:"#94a3b8",border:"none",marginTop:10,cursor:"pointer",fontSize:13,padding:8}}>← Back to Results</button>
      </div>
    </div>
  );
}

// ─── Results Page ─────────────────────────────────────────────────────────────
function ResultsPage({showToast}){
  const [name,setName]=useState("");
  const [idx,setIdx]=useState("");
  const [result,setResult]=useState(null);
  const [err,setErr]=useState("");
  const [loading,setLoading]=useState(false);
  const [deployed,setDeployed]=useState(null);

  useEffect(()=>{
    supa.get("deployment","id=eq.1&select=deployed").then(d=>{
      if(d&&d[0]) setDeployed(d[0].deployed);
    });
  },[]);

  const lookup = async () => {
    if(!rl("lookup",20,60000)){setErr("Too many lookups. Please wait a moment.");return;}
    const cleanName=clean(name).toLowerCase();
    const cleanIdx=clean(idx).toLowerCase();
    if(!cleanName||!cleanIdx){setErr("Please enter both your name and index number.");return;}
    setLoading(true);setErr("");setResult(null);
    try{
      const rows = await supa.get("students",`index_no=ilike.${encodeURIComponent(cleanIdx)}&select=*`);
      if(!rows||!rows.length){setErr("No student found with that index number.");setLoading(false);return;}
      const student=rows[0];
      if(!student.name.toLowerCase().includes(cleanName)){
        setErr("Name does not match our records. Please check and try again.");setLoading(false);return;
      }
      const [gbRes,tbRes,fmRes]=await Promise.all([
        supa.get("settings","key=eq.grade_bounds&select=value"),
        supa.get("settings","key=eq.total_bounds&select=value"),
        supa.get("settings","key=eq.formulas&select=value"),
      ]);
      const gb=gbRes?.[0]?.value||defGB();
      const tb=tbRes?.[0]?.value||defTB();
      const fm=fmRes?.[0]?.value||defFM();
      const mapped={...student,subjects:student.subjects||[],marks:student.marks||{}};
      setResult({student:mapped,...computeResults(mapped,gb,tb,fm)});
    }catch(e){setErr("Connection error. Please try again.");}
    setLoading(false);
  };

  return(
    <div style={{position:"relative",zIndex:1,minHeight:"100vh",paddingBottom:50}}>
      <div style={{background:"rgba(10,42,74,0.93)",backdropFilter:"blur(10px)",padding:"20px 28px",display:"flex",alignItems:"center",gap:16,borderBottom:"3px solid #c8102e",boxShadow:"0 2px 20px #0008"}}>
        <KenyaFlag size={40}/>
        <div style={{flex:1}}>
          <div style={{fontSize:19,fontWeight:900,color:"#fff"}}>KNEC · KCSE Results Portal</div>
          <div style={{fontSize:11,color:"#94a3b8"}}>Kenya National Examinations Council</div>
        </div>
        <div style={{display:"flex",gap:6,alignItems:"center"}}>
          <div style={{width:8,height:8,borderRadius:"50%",background:deployed?"#22c55e":"#f59e0b",boxShadow:`0 0 8px ${deployed?"#22c55e55":"#f59e0b55"}`}}/>
          <span style={{fontSize:11,color:"#94a3b8"}}>{deployed?"Results Live":"Pending"}</span>
        </div>
      </div>

      <div style={{background:"linear-gradient(90deg,rgba(200,16,46,0.85) 0%,rgba(10,42,74,0.7) 100%)",padding:"14px 28px",display:"flex",gap:24,alignItems:"center",borderBottom:"1px solid rgba(255,255,255,0.1)"}}>
        {[["🇰🇪","Kenya"],["🏫","All Schools"],["📋","KCSE 2025"]].map(([i,l])=>(
          <div key={l} style={{display:"flex",gap:8,alignItems:"center"}}><span style={{fontSize:18}}>{i}</span><span style={{color:"#fff",fontSize:13,fontWeight:600}}>{l}</span></div>
        ))}
      </div>

      <div style={{maxWidth:700,margin:"40px auto",padding:"0 16px"}}>
        {!deployed&&deployed!==null&&(
          <div style={{background:"rgba(245,158,11,0.15)",border:"1px solid #f59e0b",borderRadius:12,padding:"14px 20px",marginBottom:20,color:"#fde68a",fontSize:13,fontWeight:600,textAlign:"center"}}>
            ⏳ Results have not been published yet. Please check back later.
          </div>
        )}

        <div style={{background:"rgba(255,255,255,0.97)",borderRadius:16,padding:32,boxShadow:"0 10px 50px #0008",marginBottom:24,animation:"fadeIn .4s ease"}}>
          <div style={{fontSize:19,fontWeight:800,color:"#0a2a4a",marginBottom:4}}>Check Your Results</div>
          <div style={{fontSize:13,color:"#64748b",marginBottom:20}}>Enter your registered name and index number to view your results.</div>

          <div style={{marginBottom:12}}>
            <label style={{fontSize:11,fontWeight:700,color:"#475569",display:"block",marginBottom:4,textTransform:"uppercase"}}>Full Name (as registered)</label>
            <input value={name} onChange={e=>setName(e.target.value)} placeholder="e.g. John Kamau"
              style={{width:"100%",padding:"12px 14px",border:"2px solid #e2e8f0",borderRadius:8,fontSize:15,boxSizing:"border-box"}}/>
          </div>
          <div style={{marginBottom:16}}>
            <label style={{fontSize:11,fontWeight:700,color:"#475569",display:"block",marginBottom:4,textTransform:"uppercase"}}>Index Number</label>
            <input value={idx} onChange={e=>setIdx(e.target.value)} onKeyDown={e=>e.key==="Enter"&&lookup()} placeholder="e.g. 12345678"
              style={{width:"100%",padding:"12px 14px",border:"2px solid #e2e8f0",borderRadius:8,fontSize:15,boxSizing:"border-box"}}/>
          </div>

          <button onClick={lookup} disabled={loading||!deployed} style={{width:"100%",background:(!deployed||loading)?"#94a3b8":"#c8102e",color:"#fff",border:"none",borderRadius:9,padding:13,fontWeight:700,fontSize:15,cursor:(!deployed||loading)?"not-allowed":"pointer"}}>
            {loading?"Searching...":"Search Results"}
          </button>
          {err&&<div style={{color:"#b91c1c",marginTop:10,fontSize:13,fontWeight:600}}>⚠ {err}</div>}
        </div>

        {result&&<ResultSlip {...result}/>}
      </div>
    </div>
  );
}

function ResultSlip({student,res,langUsed,totalPts,totalGrade,top5,allIds}){
  const langLabel=SM[langUsed]?.label;

  // The other language (not used in total calculation)
  const langNotUsed=langUsed==="eng"?"swa":"eng";

  // All subjects except math and the language used for total
  // This includes: the other language + all optional subjects taken
  const remaining=allIds
    .filter(id=>id!=="math"&&id!==langUsed)
    .map(id=>({id,label:SM[id]?.label,grade:res[id]?.grade,pts:res[id]?.points||0}))
    .sort((a,b)=>b.pts-a.pts); // descending by points

  return(
    <div style={{background:"rgba(255,255,255,0.98)",borderRadius:16,padding:32,boxShadow:"0 10px 50px #0008",animation:"fadeIn .4s ease"}}>
      {/* Header */}
      <div style={{borderBottom:"3px double #c8102e",paddingBottom:18,marginBottom:20,display:"flex",justifyContent:"space-between",alignItems:"flex-start",flexWrap:"wrap",gap:12}}>
        <div>
          <div style={{fontSize:10,letterSpacing:2,color:"#c8102e",fontWeight:700,textTransform:"uppercase"}}>Kenya National Examinations Council</div>
          <div style={{fontSize:24,fontWeight:900,color:"#0a2a4a",marginTop:5}}>{student.name}</div>
          <div style={{fontSize:14,color:"#475569",marginTop:3}}>{student.school}</div>
          <div style={{fontSize:12,color:"#94a3b8",marginTop:2}}>Index No: <strong style={{color:"#334155"}}>{student.index_no}</strong></div>
        </div>
        <div style={{textAlign:"right"}}>
          <div style={{fontSize:10,color:"#94a3b8",marginBottom:4,letterSpacing:1,textTransform:"uppercase"}}>Overall Grade</div>
          <div style={{fontSize:56,fontWeight:900,color:"#15803d",lineHeight:1}}>{totalGrade}</div>
        </div>
      </div>

      {/* Subject table — all 8 subjects, Subject + Grade only, all grades green */}
      <table style={{width:"100%",borderCollapse:"collapse",fontSize:14}}>
        <thead><tr style={{background:"#0a2a4a"}}>
          <th style={{padding:"10px 14px",textAlign:"left",color:"#fff",fontWeight:700}}>Subject</th>
          <th style={{padding:"10px 14px",textAlign:"center",color:"#fff",fontWeight:700}}>Grade</th>
        </tr></thead>
        <tbody>
          {/* Row 1: Mathematics — always first */}
          <SRow label="Mathematics" grade={res["math"]?.grade} alt={false}/>
          {/* Row 2: Language used for total calculation — always second */}
          <SRow label={langLabel} grade={res[langUsed]?.grade} alt/>
          {/* Rows 3–8: All remaining 6 subjects in descending grade order */}
          {remaining.map((r,i)=><SRow key={r.id} label={r.label} grade={r.grade} alt={i%2===0}/>)}
        </tbody>
      </table>

      <div style={{marginTop:20,padding:"10px 14px",background:"#f8fafc",borderRadius:8,fontSize:11,color:"#94a3b8",textAlign:"center"}}>
        Official results issued by KNEC.
      </div>
    </div>
  );
}

function SRow({label,grade,alt}){
  return(
    <tr style={{background:alt?"#f8fafc":"#fff",borderBottom:"1px solid #f1f5f9"}}>
      <td style={{padding:"12px 14px",fontWeight:600,color:"#1e293b"}}>{label}</td>
      <td style={{padding:"12px 14px",textAlign:"center"}}>
        <span style={{fontSize:17,fontWeight:900,color:"#15803d"}}>{grade||"-"}</span>
      </td>
    </tr>
  );
}

// ─── Admin Page ───────────────────────────────────────────────────────────────
function AdminPage({showToast,onLogout}){
  const [view,setView]=useState("students");
  const [gb,setGb]=useState(null);
  const [tb,setTb]=useState(null);
  const [fm,setFm]=useState(null);
  const [students,setStudents]=useState([]);
  const [deployed,setDeployed]=useState(false);
  const [loading,setLoading]=useState(true);

  // Load all settings from Supabase
  useEffect(()=>{
    (async()=>{
      setLoading(true);
      try{
        const [gbR,tbR,fmR,stR,depR]=await Promise.all([
          supa.get("settings","key=eq.grade_bounds&select=value"),
          supa.get("settings","key=eq.total_bounds&select=value"),
          supa.get("settings","key=eq.formulas&select=value"),
          supa.get("students","select=*&order=created_at.asc"),
          supa.get("deployment","id=eq.1&select=deployed"),
        ]);
        setGb(gbR?.[0]?.value||defGB());
        setTb(tbR?.[0]?.value||defTB());
        setFm(fmR?.[0]?.value||defFM());
        setStudents(stR||[]);
        setDeployed(depR?.[0]?.deployed||false);
      }catch(e){showToast("Failed to load data from database.","error");}
      setLoading(false);
    })();
  },[]);

  const saveSetting = async (key,value) => {
    await supa.upsert("settings",[{key,value}]);
  };

  const toggleDeploy = async () => {
    const next=!deployed;
    await supa.upsert("deployment",[{id:1,deployed:next,deployed_at:new Date().toISOString()}]);
    setDeployed(next);
    showToast(next?"✅ Results published! Students can now view their grades.":"⏸ Results hidden from students.");
  };

  const navItems=[
    {v:"students",label:"👤 Students"},
    {v:"grades",label:"📊 Grade Boundaries"},
    {v:"formulas",label:"🧮 Formulas"},
    {v:"total",label:"🏆 Total Boundaries"},
  ];

  return(
    <div style={{position:"relative",zIndex:1,minHeight:"100vh",paddingBottom:50}}>
      <div style={{background:"rgba(10,42,74,0.95)",backdropFilter:"blur(10px)",padding:"16px 28px",display:"flex",alignItems:"center",gap:16,borderBottom:"3px solid #c8102e",boxShadow:"0 2px 20px #0008"}}>
        <KenyaFlag size={36}/>
        <div style={{flex:1}}>
          <div style={{fontSize:18,fontWeight:900,color:"#fff"}}>KCSE Admin Panel</div>
          <div style={{fontSize:11,color:"#94a3b8"}}>Superadmin · Restricted</div>
        </div>
        {/* Deploy toggle */}
        <button onClick={toggleDeploy} style={{background:deployed?"#15803d":"#f59e0b",color:"#fff",border:"none",borderRadius:8,padding:"8px 16px",fontWeight:700,fontSize:13,cursor:"pointer",marginRight:8}}>
          {deployed?"✅ Results Live":"🚀 Publish Results"}
        </button>
        <a href="#" onClick={e=>{e.preventDefault();window.location.hash="";}} style={{color:"#94a3b8",fontSize:12,textDecoration:"none",marginRight:8}}>← Public</a>
        <button onClick={onLogout} style={{background:"#c8102e",color:"#fff",border:"none",borderRadius:7,padding:"7px 16px",fontWeight:700,fontSize:13,cursor:"pointer"}}>Logout</button>
      </div>

      <div style={{maxWidth:1200,margin:"0 auto",padding:"24px 16px"}}>
        {/* Deploy status banner */}
        <div style={{background:deployed?"rgba(21,128,61,0.15)":"rgba(245,158,11,0.15)",border:`1px solid ${deployed?"#15803d":"#f59e0b"}`,borderRadius:10,padding:"12px 20px",marginBottom:20,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{color:deployed?"#86efac":"#fde68a",fontWeight:600,fontSize:13}}>
            {deployed?"✅ Results are currently LIVE — students can view grades.":"⏸ Results are HIDDEN — students cannot view grades yet."}
          </span>
          <button onClick={toggleDeploy} style={{background:deployed?"#b91c1c":"#15803d",color:"#fff",border:"none",borderRadius:6,padding:"6px 14px",fontWeight:700,fontSize:12,cursor:"pointer"}}>
            {deployed?"Hide Results":"Publish Results"}
          </button>
        </div>

        <div style={{display:"flex",gap:8,marginBottom:24,flexWrap:"wrap"}}>
          {navItems.map(({v,label})=>(
            <button key={v} onClick={()=>setView(v)} style={{background:view===v?"#c8102e":"rgba(255,255,255,0.93)",color:view===v?"#fff":"#0a2a4a",border:"none",borderRadius:9,padding:"9px 20px",fontWeight:700,fontSize:13,cursor:"pointer",boxShadow:"0 2px 10px #0003"}}>
              {label}
            </button>
          ))}
        </div>

        {loading?<Spinner/>:<>
          {view==="students"&&<StudentsPanel students={students} setStudents={setStudents} gb={gb} tb={tb} fm={fm} showToast={showToast}/>}
          {view==="grades"&&gb&&<GradePanel gb={gb} setGb={setGb} saveSetting={saveSetting} showToast={showToast}/>}
          {view==="formulas"&&fm&&<FormulaPanel fm={fm} setFm={setFm} saveSetting={saveSetting} showToast={showToast}/>}
          {view==="total"&&tb&&<TotalPanel tb={tb} setTb={setTb} saveSetting={saveSetting} showToast={showToast}/>}
        </>}
      </div>
    </div>
  );
}

// ─── Students Panel ───────────────────────────────────────────────────────────
function StudentsPanel({students,setStudents,gb,tb,fm,showToast}){
  const empty={name:"",school:"",index_no:"",subjects:[],marks:{}};
  const [form,setForm]=useState(empty);
  const [errors,setErrors]=useState([]);
  const [editId,setEditId]=useState(null);
  const [search,setSearch]=useState("");
  const [saving,setSaving]=useState(false);
  const fileRef=useRef();

  const toggleSub=id=>setForm(f=>{const on=f.subjects.includes(id);return{...f,subjects:on?f.subjects.filter(x=>x!==id):[...f.subjects,id]};});

  const setMark=(subId,paper,val)=>setForm(prev=>{
    const sm={...(prev.marks[subId]||{})};
    if(val===""||val===undefined)delete sm[paper]; else sm[paper]=val;
    return{...prev,marks:{...prev.marks,[subId]:sm}};
  });

  const save=async()=>{
    const e=[];
    if(!form.name.trim())e.push("Full name required.");
    if(!form.school.trim())e.push("School required.");
    if(!form.index_no.trim())e.push("Index number required.");
    const comp=SUBJECTS.filter(s=>s.compulsory).map(s=>s.id);
    if([...comp,...form.subjects].length>8)e.push("Maximum 8 subjects.");
    e.push(...validateChoice(form.subjects));
    if(e.length){setErrors(e);return;}
    setSaving(true);
    try{
      const rec={name:clean(form.name),school:clean(form.school),index_no:clean(form.index_no),subjects:form.subjects,marks:form.marks};
      if(editId){rec.id=editId;}
      const res=await supa.upsert("students",[rec]);
      if(res&&res[0]){
        if(editId)setStudents(ss=>ss.map(s=>s.id===editId?res[0]:s));
        else setStudents(ss=>[...ss,res[0]]);
        showToast(editId?"Student updated.":"Student saved.");
        setForm(empty);setEditId(null);setErrors([]);
      }else{showToast("Save failed. Check for duplicate index number.","error");}
    }catch{showToast("Database error.","error");}
    setSaving(false);
  };

  const del=async(id,name)=>{
    if(!window.confirm("Delete "+name+"?"))return;
    await supa.del("students","id=eq."+id);
    setStudents(ss=>ss.filter(s=>s.id!==id));
    showToast("Deleted.");
  };

  const startEdit=s=>{
    setForm({name:s.name,school:s.school||"",index_no:s.index_no,subjects:s.subjects||[],marks:JSON.parse(JSON.stringify(s.marks||{}))});
    setEditId(s.id);window.scrollTo({top:0,behavior:"smooth"});
  };

  // File import
  const handleFile=async e=>{
    const file=e.target.files[0];if(!file)return;
    try{
      const data=await file.arrayBuffer();
      const wb=XLSX.read(data);
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{defval:""});
      if(!rows.length){showToast("File is empty.","error");e.target.value="";return;}
      const toInsert=[];
      for(const row of rows){
        const name=String(row.name||row.Name||row.FULL_NAME||"").trim();
        const school=String(row.school||row.School||"").trim();
        const index_no=String(row.index||row.Index||row.index_no||row["Index No"]||"").trim();
        if(!name||!index_no)continue;
        const marks={};const subjects=[];
        SUBJECTS.forEach(sub=>{
          sub.papers.forEach(p=>{
            const val=row[sub.id+"_"+p];
            if(val!==""&&val!==undefined&&val!==null){
              if(!marks[sub.id])marks[sub.id]={};
              marks[sub.id][p]=Number(val);
              if(!sub.compulsory&&!subjects.includes(sub.id))subjects.push(sub.id);
            }
          });
        });
        toInsert.push({name:clean(name),school:clean(school||"Unknown"),index_no:clean(index_no),subjects,marks});
      }
      if(!toInsert.length){showToast("No valid rows. Check column names.","error");e.target.value="";return;}
      const res=await supa.upsert("students",toInsert);
      const newStudents=Array.isArray(res)?res:[];
      setStudents(ss=>{const ids=new Set(ss.map(s=>s.id));return[...ss,...newStudents.filter(s=>!ids.has(s.id))];});
      showToast("Imported "+toInsert.length+" student(s).");
    }catch(err){showToast("File error: "+err.message,"error");}
    e.target.value="";
  };

  const comp=SUBJECTS.filter(s=>s.compulsory).map(s=>s.id);
  const allFormSubs=[...comp,...form.subjects];
  const filtered=students.filter(s=>
    s.name?.toLowerCase().includes(search.toLowerCase())||
    s.index_no?.toLowerCase().includes(search.toLowerCase())||
    s.school?.toLowerCase().includes(search.toLowerCase())
  );

  return(
    <div style={{display:"grid",gridTemplateColumns:"minmax(0,1fr) minmax(0,1fr)",gap:20,alignItems:"start"}}>
      {/* Form */}
      <div style={{background:"rgba(255,255,255,0.97)",borderRadius:14,padding:24,boxShadow:"0 6px 30px #0006"}}>
        <h3 style={{margin:"0 0 16px",color:"#0a2a4a",fontSize:16}}>{editId?"✏️ Edit Student":"➕ Add Student"}</h3>
        {errors.map((e,i)=><div key={i} style={{background:"#fee2e2",color:"#b91c1c",borderRadius:7,padding:"7px 12px",marginBottom:6,fontSize:12,fontWeight:600}}>⚠ {e}</div>)}

        {[["Full Name","name"],["School","school"],["Index Number","index_no"]].map(([lbl,key])=>(
          <div key={key} style={{marginBottom:11}}>
            <label style={{display:"block",fontSize:11,fontWeight:700,color:"#475569",marginBottom:3,textTransform:"uppercase"}}>{lbl}</label>
            <input value={form[key]} onChange={e=>setForm(f=>({...f,[key]:e.target.value}))}
              style={{width:"100%",padding:"9px 11px",border:"2px solid #e2e8f0",borderRadius:7,fontSize:14,boxSizing:"border-box"}}/>
          </div>
        ))}

        {/* File import */}
        <div style={{marginBottom:14}}>
          <div style={{fontSize:11,fontWeight:700,color:"#475569",marginBottom:5,textTransform:"uppercase"}}>Import from File</div>
          <div style={{border:"2px dashed #c8102e",borderRadius:9,padding:14,background:"#fff8f8",cursor:"pointer",textAlign:"center"}}
            onClick={()=>fileRef.current.click()}
            onDragOver={e=>e.preventDefault()}
            onDrop={e=>{e.preventDefault();const f=e.dataTransfer.files[0];if(f){const dt=new DataTransfer();dt.items.add(f);fileRef.current.files=dt.files;handleFile({target:fileRef.current});}}}>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{display:"none"}}/>
            <div style={{fontSize:22,marginBottom:4}}>📎</div>
            <div style={{fontSize:13,color:"#c8102e",fontWeight:700}}>Click or drag to upload</div>
            <div style={{fontSize:11,color:"#94a3b8",marginTop:4}}>.xlsx · .xls · .csv</div>
            <div style={{fontSize:10,color:"#cbd5e1",marginTop:6,fontFamily:"monospace"}}>Columns: name, school, index_no, math_pp1, math_pp2, eng_pp1...</div>
          </div>
        </div>

        {/* Optional subjects */}
        <div style={{marginBottom:12}}>
          <div style={{fontSize:11,fontWeight:700,color:"#475569",marginBottom:6,textTransform:"uppercase"}}>Optional Subjects</div>
          <div style={{background:"#f0fdf4",borderRadius:8,padding:"10px 12px",marginBottom:8,border:"1px solid #d1fae5"}}>
            <div style={{fontSize:11,fontWeight:700,color:"#15803d",marginBottom:6}}>Group A · choose 2–3</div>
            <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
              {["phyc","geo","hist","cre"].map(id=>{const on=form.subjects.includes(id);return(
                <button key={id} onClick={()=>toggleSub(id)} style={{background:on?"#15803d":"#fff",color:on?"#fff":"#334155",border:"1.5px solid "+(on?"#15803d":"#d1fae5"),borderRadius:20,padding:"4px 13px",fontSize:12,cursor:"pointer",fontWeight:on?700:400}}>{SM[id].label}</button>
              );})}
            </div>
          </div>
          <div style={{background:"#eff6ff",borderRadius:8,padding:"10px 12px",border:"1px solid #bfdbfe"}}>
            <div style={{fontSize:11,fontWeight:700,color:"#1e40af",marginBottom:6}}>Group B · at most 1</div>
            <div style={{display:"flex",flexWrap:"wrap",gap:5}}>
              {["bst","agric","comp"].map(id=>{const on=form.subjects.includes(id);return(
                <button key={id} onClick={()=>toggleSub(id)} style={{background:on?"#1e40af":"#fff",color:on?"#fff":"#334155",border:"1.5px solid "+(on?"#1e40af":"#bfdbfe"),borderRadius:20,padding:"4px 13px",fontSize:12,cursor:"pointer",fontWeight:on?700:400}}>{SM[id].label}</button>
              );})}
            </div>
          </div>
        </div>

        {/* Marks */}
        <div style={{fontSize:11,fontWeight:700,color:"#475569",marginBottom:8,textTransform:"uppercase"}}>Marks Entry</div>
        {allFormSubs.map(id=>{
          const sub=SM[id];
          return(
            <div key={id} style={{marginBottom:10,background:"#f8fafc",borderRadius:9,padding:"11px 13px",border:"1.5px solid #e2e8f0"}}>
              <div style={{fontSize:12,fontWeight:700,color:"#0a2a4a",marginBottom:7,display:"flex",alignItems:"center",gap:6}}>
                {sub.label}
                {sub.compulsory&&<span style={{fontSize:10,background:"#c8102e",color:"#fff",borderRadius:10,padding:"1px 8px"}}>Compulsory</span>}
              </div>
              <div style={{display:"flex",gap:8}}>
                {sub.papers.map(p=>(
                  <div key={p} style={{flex:1}}>
                    <label style={{fontSize:10,color:"#94a3b8",display:"block",marginBottom:3,fontWeight:700,textTransform:"uppercase"}}>{p}</label>
                    <input type="number" min="0" max="200"
                      value={form.marks[id]?.[p]??""}
                      onChange={e=>setMark(id,p,e.target.value===""?"":e.target.value)}
                      style={{width:"100%",padding:"8px 6px",fontSize:14,textAlign:"center",border:"2px solid #e2e8f0",borderRadius:7,boxSizing:"border-box",background:"#fff"}}/>
                  </div>
                ))}
              </div>
            </div>
          );
        })}

        <div style={{display:"flex",gap:8,marginTop:16}}>
          <button onClick={save} disabled={saving} style={{flex:1,background:saving?"#94a3b8":"#c8102e",color:"#fff",border:"none",borderRadius:9,padding:12,fontWeight:700,fontSize:14,cursor:saving?"not-allowed":"pointer"}}>
            {saving?"Saving...":(editId?"Update Student":"Save Student")}
          </button>
          {editId&&<button onClick={()=>{setForm(empty);setEditId(null);setErrors([]);}} style={{background:"#f1f5f9",color:"#475569",border:"none",borderRadius:9,padding:"12px 18px",cursor:"pointer",fontWeight:600}}>Cancel</button>}
        </div>
      </div>

      {/* List */}
      <div style={{background:"rgba(255,255,255,0.97)",borderRadius:14,padding:24,boxShadow:"0 6px 30px #0006"}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:12}}>
          <h3 style={{margin:0,color:"#0a2a4a",fontSize:16}}>Students ({students.length})</h3>
          <span style={{fontSize:12,color:"#64748b"}}>Showing {filtered.length}</span>
        </div>
        <input value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search name, index, school..."
          style={{width:"100%",padding:"9px 12px",border:"2px solid #e2e8f0",borderRadius:8,fontSize:13,boxSizing:"border-box",marginBottom:12}}/>
        {filtered.length===0&&<div style={{color:"#94a3b8",fontSize:13,textAlign:"center",padding:30}}>{students.length===0?"No students yet. Add or import one.":"No matches."}</div>}
        <div style={{maxHeight:520,overflowY:"auto"}}>
          {filtered.map(s=>(
            <div key={s.id} style={{background:"#f8fafc",borderRadius:9,padding:"12px 14px",marginBottom:8,display:"flex",justifyContent:"space-between",alignItems:"center",border:"1px solid #e2e8f0"}}>
              <div>
                <div style={{fontWeight:700,color:"#1e293b",fontSize:13}}>{s.name}</div>
                <div style={{fontSize:11,color:"#64748b",marginTop:2}}>{s.school} · <code style={{background:"#e2e8f0",borderRadius:3,padding:"0 4px"}}>{s.index_no}</code></div>
              </div>
              <div style={{display:"flex",gap:6}}>
                <button onClick={()=>startEdit(s)} style={{background:"#dbeafe",color:"#1d4ed8",border:"none",borderRadius:6,padding:"5px 12px",cursor:"pointer",fontSize:11,fontWeight:700}}>Edit</button>
                <button onClick={()=>del(s.id,s.name)} style={{background:"#fee2e2",color:"#b91c1c",border:"none",borderRadius:6,padding:"5px 12px",cursor:"pointer",fontSize:11,fontWeight:700}}>Delete</button>
              </div>
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

// ─── Grade Bounds ─────────────────────────────────────────────────────────────
function GradePanel({gb,setGb,saveSetting,showToast}){
  const set=(id,g,v)=>setGb(b=>({...b,[id]:{...b[id],[g]:Number(v)}}));
  const save=async()=>{await saveSetting("grade_bounds",gb);showToast("Grade boundaries saved.");};
  return(
    <div style={{background:"rgba(255,255,255,0.97)",borderRadius:14,padding:24,boxShadow:"0 6px 30px #0006"}}>
      <h3 style={{margin:"0 0 6px",color:"#0a2a4a"}}>Subject Grade Boundaries</h3>
      <p style={{color:"#64748b",fontSize:13,marginBottom:18}}>Minimum mark for each grade per subject.</p>
      <div style={{overflowX:"auto"}}>
        <table style={{borderCollapse:"collapse",width:"100%",fontSize:12}}>
          <thead><tr style={{background:"#0a2a4a"}}>
            <th style={{padding:"10px 12px",textAlign:"left",color:"#fff",fontWeight:700,minWidth:130}}>Subject</th>
            {GL.map(g=><th key={g} style={{padding:"10px 5px",color:["A","A-"].includes(g)?"#86efac":["B+","B","B-"].includes(g)?"#93c5fd":["C+","C","C-"].includes(g)?"#fde68a":"#fca5a5",fontWeight:700}}>{g}</th>)}
          </tr></thead>
          <tbody>
            {SUBJECTS.map((sub,ri)=>(
              <tr key={sub.id} style={{background:ri%2===0?"#fff":"#f8fafc"}}>
                <td style={{padding:"7px 12px",fontWeight:600,color:"#334155",whiteSpace:"nowrap"}}>{sub.label}</td>
                {GL.map(g=>(
                  <td key={g} style={{padding:"4px 2px"}}>
                    <input type="number" min="0" max="500" value={gb[sub.id]?.[g]??0} onChange={e=>set(sub.id,g,e.target.value)}
                      style={{width:46,padding:"4px 4px",border:"1.5px solid #e2e8f0",borderRadius:5,fontSize:11,textAlign:"center"}}/>
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <button onClick={save} style={{marginTop:16,background:"#c8102e",color:"#fff",border:"none",borderRadius:8,padding:"10px 26px",fontWeight:700,cursor:"pointer"}}>Save to Database</button>
    </div>
  );
}

// ─── Formulas ─────────────────────────────────────────────────────────────────
function FormulaPanel({fm,setFm,saveSetting,showToast}){
  const save=async()=>{await saveSetting("formulas",fm);showToast("Formulas saved.");};
  return(
    <div style={{background:"rgba(255,255,255,0.97)",borderRadius:14,padding:24,boxShadow:"0 6px 30px #0006",maxWidth:700}}>
      <h3 style={{margin:"0 0 4px",color:"#0a2a4a"}}>Mark Calculation Formulas</h3>
      <p style={{color:"#64748b",fontSize:13,marginBottom:20}}>Variables: <code style={{background:"#f1f5f9",padding:"1px 7px",borderRadius:4}}>pp1</code> <code style={{background:"#f1f5f9",padding:"1px 7px",borderRadius:4}}>pp2</code> <code style={{background:"#f1f5f9",padding:"1px 7px",borderRadius:4}}>pp3</code></p>
      {SUBJECTS.map(sub=>(
        <div key={sub.id} style={{display:"flex",alignItems:"center",gap:12,marginBottom:10}}>
          <label style={{width:160,fontWeight:600,color:"#334155",fontSize:13,flexShrink:0}}>{sub.label}</label>
          <input value={fm[sub.id]||""} onChange={e=>setFm(f=>({...f,[sub.id]:e.target.value}))}
            style={{flex:1,padding:"8px 12px",border:"2px solid #e2e8f0",borderRadius:7,fontSize:13,fontFamily:"monospace"}}/>
          <span style={{fontSize:11,color:"#94a3b8",flexShrink:0}}>({sub.papers.join(", ")})</span>
        </div>
      ))}
      <button onClick={save} style={{marginTop:12,background:"#c8102e",color:"#fff",border:"none",borderRadius:8,padding:"10px 26px",fontWeight:700,cursor:"pointer"}}>Save to Database</button>
    </div>
  );
}

// ─── Total Bounds ─────────────────────────────────────────────────────────────
function TotalPanel({tb,setTb,saveSetting,showToast}){
  const save=async()=>{await saveSetting("total_bounds",tb);showToast("Total boundaries saved.");};
  return(
    <div style={{background:"rgba(255,255,255,0.97)",borderRadius:14,padding:24,boxShadow:"0 6px 30px #0006",maxWidth:460}}>
      <h3 style={{margin:"0 0 4px",color:"#0a2a4a"}}>Total Grade Boundaries</h3>
      <p style={{color:"#64748b",fontSize:13,marginBottom:20}}>Minimum total points per overall grade. Max = 84.</p>
      {GL.map(g=>(
        <div key={g} style={{display:"flex",alignItems:"center",gap:14,marginBottom:10}}>
          <span style={{width:36,fontWeight:900,fontSize:20,color:gColor(g)}}>{g}</span>
          <input type="number" min="0" max="84" value={tb[g]??0} onChange={e=>setTb(b=>({...b,[g]:Number(e.target.value)}))}
            style={{width:80,padding:"8px 10px",border:"2px solid #e2e8f0",borderRadius:7,fontSize:14,textAlign:"center"}}/>
          <span style={{fontSize:12,color:"#94a3b8"}}>points and above</span>
        </div>
      ))}
      <button onClick={save} style={{marginTop:12,background:"#c8102e",color:"#fff",border:"none",borderRadius:8,padding:"10px 26px",fontWeight:700,cursor:"pointer"}}>Save to Database</button>
    </div>
  );
}
