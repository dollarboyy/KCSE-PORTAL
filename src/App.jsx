import { useState, useEffect, useRef, useCallback } from "react";
import * as XLSX from "xlsx";

// ─── Rate limiting / DDoS protection (client-side layer) ─────────────────────
const rateLimiter = (() => {
  const windows = {};
  return {
    check(key, maxReqs = 10, windowMs = 60000) {
      const now = Date.now();
      if (!windows[key]) windows[key] = [];
      windows[key] = windows[key].filter(t => now - t < windowMs);
      if (windows[key].length >= maxReqs) return false;
      windows[key].push(now);
      return true;
    }
  };
})();

const sanitize = (str) => String(str).replace(/[<>"'`]/g, "").trim().slice(0, 200);

// ─── Constants ────────────────────────────────────────────────────────────────
const GRADE_LABELS = ["A","A-","B+","B","B-","C+","C","C-","D+","D","D-","E"];
const GRADE_POINTS = { A:12,"A-":11,"B+":10,B:9,"B-":8,"C+":7,C:6,"C-":5,"D+":4,D:3,"D-":2,E:1 };

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

const SUB_MAP = Object.fromEntries(SUBJECTS.map(s => [s.id, s]));

const defaultGradeBounds = () => {
  const b = {};
  const d = { A:75,"A-":70,"B+":65,B:60,"B-":55,"C+":50,C:45,"C-":40,"D+":35,D:30,"D-":25,E:0 };
  SUBJECTS.forEach(s => { b[s.id] = { ...d }; });
  return b;
};

const defaultTotalBounds = () =>
  ({ A:73,"A-":61,"B+":53,B:47,"B-":41,"C+":35,C:29,"C-":23,"D+":17,D:11,"D-":5,E:0 });

const defaultFormulas = () => {
  const f = {};
  SUBJECTS.forEach(s => { f[s.id] = s.papers.length === 3 ? "pp1 + pp2 + pp3" : "pp1 + pp2"; });
  return f;
};

// ─── Helpers ──────────────────────────────────────────────────────────────────
function evalFormula(formula, marks) {
  try {
    const expr = formula
      .replace(/pp1/gi, Number(marks?.pp1) || 0)
      .replace(/pp2/gi, Number(marks?.pp2) || 0)
      .replace(/pp3/gi, Number(marks?.pp3) || 0);
    const result = Function('"use strict"; return (' + expr + ')')();
    return isNaN(result) ? 0 : Math.round(result * 100) / 100;
  } catch { return 0; }
}

function getGrade(total, bounds) {
  for (const g of GRADE_LABELS) {
    if (total >= (bounds[g] ?? 0)) return g;
  }
  return "E";
}

function validateChoice(chosen) {
  const A = chosen.filter(id => ["phyc","geo","hist","cre"].includes(id));
  const B = chosen.filter(id => ["bst","agric","comp"].includes(id));
  const errors = [];
  if (A.length < 2) errors.push("Choose at least 2 from Group A (Physics, Geography, History, CRE).");
  if (A.length > 3) errors.push("Choose at most 3 from Group A.");
  if (A.includes("geo") && A.includes("hist")) errors.push("Geography and History cannot be combined.");
  if (!A.includes("phyc") && !A.includes("cre")) errors.push("At least one of Physics or CRE must be chosen.");
  if (B.length > 1) errors.push("At most 1 subject from Group B (Business, Agriculture, Computer Studies).");
  return errors;
}

function computeResults(student, gradeBounds, totalBounds, formulas) {
  const gb = gradeBounds || defaultGradeBounds();
  const tb = totalBounds || defaultTotalBounds();
  const fm = formulas || defaultFormulas();
  const compIds = SUBJECTS.filter(s => s.compulsory).map(s => s.id);
  const allIds = [...compIds, ...(student.subjects || [])];
  const res = {};

  allIds.forEach(id => {
    const total = evalFormula(fm[id] || "pp1 + pp2", student.marks?.[id] || {});
    const grade = getGrade(total, gb[id] || {});
    res[id] = { total, grade, points: GRADE_POINTS[grade] || 1 };
  });

  const engPts = res["eng"]?.points || 0;
  const swaPts = res["swa"]?.points || 0;
  const langUsed = engPts >= swaPts ? "eng" : "swa";

  const nonLang = allIds.filter(id => id !== "eng" && id !== "swa");
  const sortedNonLang = [...nonLang].sort((a, b) => (res[b]?.points || 0) - (res[a]?.points || 0));
  const top5 = sortedNonLang.slice(0, 5);

  const totalPts = (res[langUsed]?.points || 0) + top5.reduce((s, id) => s + (res[id]?.points || 0), 0);
  const totalGrade = getGrade(totalPts, tb);

  return { res, langUsed, totalPts, totalGrade, top5, allIds };
}

const gradeColor = (g) => {
  if (["A","A-"].includes(g)) return "#15803d";
  if (["B+","B","B-"].includes(g)) return "#1e40af";
  if (["C+","C","C-"].includes(g)) return "#b45309";
  return "#b91c1c";
};

// ─── Kenya Flag SVG ───────────────────────────────────────────────────────────
const KenyaFlag = ({ size = 32 }) => (
  <svg width={size} height={Math.round(size * 0.67)} viewBox="0 0 90 60" style={{ borderRadius:3, boxShadow:"0 1px 4px #0004", flexShrink:0 }}>
    <rect width="90" height="60" fill="#006600"/>
    <rect y="15" width="90" height="30" fill="#BB0000"/>
    <rect y="22" width="90" height="16" fill="#000"/>
    <rect y="24" width="90" height="12" fill="#fff"/>
    <rect y="26" width="90" height="8" fill="#000"/>
    <ellipse cx="45" cy="30" rx="10" ry="14" fill="#BB0000" stroke="#fff" strokeWidth="1.5"/>
    <line x1="45" y1="16" x2="45" y2="44" stroke="#fff" strokeWidth="2"/>
  </svg>
);

// ─── Background ───────────────────────────────────────────────────────────────
const PageBackground = () => (
  <div style={{ position:"fixed", inset:0, zIndex:0, overflow:"hidden", pointerEvents:"none" }}>
    <div style={{ position:"absolute", inset:0, background:"linear-gradient(160deg, #003d1a 0%, #0a2a4a 45%, #1a0a00 100%)" }}/>
    <svg style={{ position:"absolute", inset:0, width:"100%", height:"100%", opacity:0.05 }}>
      <defs>
        <pattern id="pg" width="60" height="60" patternUnits="userSpaceOnUse">
          <path d="M 60 0 L 0 0 0 60" fill="none" stroke="#fff" strokeWidth="0.5"/>
        </pattern>
      </defs>
      <rect width="100%" height="100%" fill="url(#pg)"/>
    </svg>
    <div style={{ position:"absolute", top:-120, right:-120, width:450, height:450, borderRadius:"50%", background:"rgba(187,0,0,0.07)" }}/>
    <div style={{ position:"absolute", bottom:-100, left:-100, width:380, height:380, borderRadius:"50%", background:"rgba(0,102,0,0.09)" }}/>
    <div style={{ position:"absolute", top:"35%", left:"55%", width:220, height:220, borderRadius:"50%", background:"rgba(255,255,255,0.025)" }}/>
  </div>
);

// ─── Toast ────────────────────────────────────────────────────────────────────
function Toast({ msg, type, onClose }) {
  useEffect(() => { const t = setTimeout(onClose, 3500); return () => clearTimeout(t); }, [onClose]);
  return (
    <div style={{
      position:"fixed", bottom:24, right:24, zIndex:9999,
      background: type === "error" ? "#b91c1c" : "#15803d",
      color:"#fff", borderRadius:10, padding:"12px 20px",
      fontWeight:600, fontSize:14, boxShadow:"0 4px 20px #0005",
      animation:"slideIn .3s ease"
    }}>{msg}</div>
  );
}

// ─── App ──────────────────────────────────────────────────────────────────────
export default function App() {
  const [page, setPage] = useState(() =>
    window.location.hash === "#/superadmin" ? "admin" : "results"
  );

  useEffect(() => {
    const h = () => setPage(window.location.hash === "#/superadmin" ? "admin" : "results");
    window.addEventListener("hashchange", h);
    return () => window.removeEventListener("hashchange", h);
  }, []);

  const [students, setStudents] = useState(() => {
    try { return JSON.parse(sessionStorage.getItem("kcse_students") || "[]"); } catch { return []; }
  });
  const [gradeBounds, setGradeBounds] = useState(() => {
    try { return JSON.parse(sessionStorage.getItem("kcse_gbounds") || "null") || defaultGradeBounds(); } catch { return defaultGradeBounds(); }
  });
  const [totalBounds, setTotalBounds] = useState(() => {
    try { return JSON.parse(sessionStorage.getItem("kcse_tbounds") || "null") || defaultTotalBounds(); } catch { return defaultTotalBounds(); }
  });
  const [formulas, setFormulas] = useState(() => {
    try { return JSON.parse(sessionStorage.getItem("kcse_formulas") || "null") || defaultFormulas(); } catch { return defaultFormulas(); }
  });

  useEffect(() => { sessionStorage.setItem("kcse_students", JSON.stringify(students)); }, [students]);
  useEffect(() => { sessionStorage.setItem("kcse_gbounds", JSON.stringify(gradeBounds)); }, [gradeBounds]);
  useEffect(() => { sessionStorage.setItem("kcse_tbounds", JSON.stringify(totalBounds)); }, [totalBounds]);
  useEffect(() => { sessionStorage.setItem("kcse_formulas", JSON.stringify(formulas)); }, [formulas]);

  const [toast, setToast] = useState(null);
  const showToast = useCallback((msg, type = "success") => setToast({ msg, type }), []);

  const [authed, setAuthed] = useState(() => sessionStorage.getItem("kcse_auth") === "1");
  const [authPw, setAuthPw] = useState("");
  const [authErr, setAuthErr] = useState("");
  const ADMIN_PW = "Admin@2025";

  const handleLogin = () => {
    if (!rateLimiter.check("login", 5, 60000)) {
      setAuthErr("Too many attempts. Please wait 1 minute."); return;
    }
    if (authPw === ADMIN_PW) {
      setAuthed(true); sessionStorage.setItem("kcse_auth", "1"); setAuthErr("");
    } else {
      setAuthErr("Incorrect password.");
    }
  };

  return (
    <div style={{ minHeight:"100vh", position:"relative", fontFamily:"'Segoe UI',system-ui,sans-serif" }}>
      <style>{`
        @keyframes slideIn { from { transform:translateX(40px);opacity:0 } to { transform:translateX(0);opacity:1 } }
        @keyframes fadeIn  { from { opacity:0;transform:translateY(10px) } to { opacity:1;transform:translateY(0) } }
        input[type=number]::-webkit-inner-spin-button { opacity:.6 }
      `}</style>
      <PageBackground />
      {toast && <Toast msg={toast.msg} type={toast.type} onClose={() => setToast(null)} />}
      {page === "results" && (
        <ResultsPage students={students} gradeBounds={gradeBounds} totalBounds={totalBounds} formulas={formulas} />
      )}
      {page === "admin" && !authed && (
        <LoginPage pw={authPw} setPw={setAuthPw} err={authErr} onLogin={handleLogin}
          onBack={() => { window.location.hash = ""; }} />
      )}
      {page === "admin" && authed && (
        <AdminPage
          students={students} setStudents={setStudents}
          gradeBounds={gradeBounds} setGradeBounds={setGradeBounds}
          totalBounds={totalBounds} setTotalBounds={setTotalBounds}
          formulas={formulas} setFormulas={setFormulas}
          showToast={showToast}
          onLogout={() => { setAuthed(false); sessionStorage.removeItem("kcse_auth"); window.location.hash = ""; }}
        />
      )}
    </div>
  );
}

// ─── Login ────────────────────────────────────────────────────────────────────
function LoginPage({ pw, setPw, err, onLogin, onBack }) {
  return (
    <div style={{ position:"relative", zIndex:1, minHeight:"100vh", display:"flex", alignItems:"center", justifyContent:"center" }}>
      <div style={{ background:"rgba(255,255,255,0.97)", borderRadius:18, padding:44, width:360, boxShadow:"0 12px 60px #0009", animation:"fadeIn .4s ease" }}>
        <div style={{ textAlign:"center", marginBottom:28 }}>
          <KenyaFlag size={48} />
          <div style={{ fontSize:22, fontWeight:900, color:"#0a2a4a", marginTop:14, letterSpacing:.5 }}>Superadmin Portal</div>
          <div style={{ fontSize:13, color:"#64748b", marginTop:4 }}>KCSE Results Management System</div>
        </div>
        {err && <div style={{ background:"#fee2e2", color:"#b91c1c", borderRadius:7, padding:"8px 12px", marginBottom:12, fontSize:13 }}>⚠ {err}</div>}
        <label style={{ fontSize:11, fontWeight:700, color:"#475569", display:"block", marginBottom:4, textTransform:"uppercase", letterSpacing:.5 }}>Password</label>
        <input type="password" value={pw} onChange={e => setPw(e.target.value)}
          onKeyDown={e => e.key === "Enter" && onLogin()}
          style={{ width:"100%", padding:"11px 13px", border:"2px solid #e2e8f0", borderRadius:8, fontSize:15, boxSizing:"border-box", marginBottom:16 }} />
        <button onClick={onLogin} style={{
          width:"100%", background:"#c8102e", color:"#fff", border:"none",
          borderRadius:9, padding:13, fontWeight:700, fontSize:16, cursor:"pointer"
        }}>Sign In</button>
        <button onClick={onBack} style={{
          width:"100%", background:"transparent", color:"#94a3b8", border:"none",
          marginTop:10, cursor:"pointer", fontSize:13, padding:8
        }}>← Back to Results</button>
      </div>
    </div>
  );
}

// ─── Results Page ─────────────────────────────────────────────────────────────
function ResultsPage({ students, gradeBounds, totalBounds, formulas }) {
  const [idx, setIdx] = useState("");
  const [result, setResult] = useState(null);
  const [err, setErr] = useState("");
  const [blocked, setBlocked] = useState(false);

  const lookup = () => {
    if (!rateLimiter.check("lookup", 20, 60000)) {
      setBlocked(true); setErr("Too many lookups. Please wait a moment."); return;
    }
    const clean = sanitize(idx);
    if (!clean) { setErr("Please enter an index number."); return; }
    const s = students.find(st => st.index.toLowerCase() === clean.toLowerCase());
    if (!s) { setResult(null); setErr("No student found with that index number."); return; }
    setErr(""); setBlocked(false);
    setResult({ student: s, ...computeResults(s, gradeBounds, totalBounds, formulas) });
  };

  return (
    <div style={{ position:"relative", zIndex:1, minHeight:"100vh", paddingBottom:50 }}>
      {/* Header */}
      <div style={{
        background:"rgba(10,42,74,0.93)", backdropFilter:"blur(10px)",
        padding:"20px 28px", display:"flex", alignItems:"center", gap:16,
        borderBottom:"3px solid #c8102e", boxShadow:"0 2px 20px #0008"
      }}>
        <KenyaFlag size={40} />
        <div style={{ flex:1 }}>
          <div style={{ fontSize:19, fontWeight:900, color:"#fff", letterSpacing:.5 }}>KNEC · KCSE Results Portal</div>
          <div style={{ fontSize:11, color:"#94a3b8", marginTop:1 }}>Kenya National Examinations Council — Official Results System</div>
        </div>
        <div style={{ display:"flex", gap:6, alignItems:"center" }}>
          <div style={{ width:8, height:8, borderRadius:"50%", background:"#22c55e", boxShadow:"0 0 8px #22c55e55" }}/>
          <span style={{ fontSize:11, color:"#94a3b8" }}>Online</span>
        </div>
      </div>

      {/* Hero band */}
      <div style={{
        background:"linear-gradient(90deg, rgba(200,16,46,0.85) 0%, rgba(10,42,74,0.7) 100%)",
        padding:"18px 28px", display:"flex", gap:32, alignItems:"center",
        borderBottom:"1px solid rgba(255,255,255,0.1)"
      }}>
        {[["🇰🇪","Kenya"],["🏫","All Schools"],["📋","KCSE 2025"]].map(([icon,label]) => (
          <div key={label} style={{ display:"flex", gap:8, alignItems:"center" }}>
            <span style={{ fontSize:18 }}>{icon}</span>
            <span style={{ color:"#fff", fontSize:13, fontWeight:600 }}>{label}</span>
          </div>
        ))}
      </div>

      <div style={{ maxWidth:700, margin:"40px auto", padding:"0 16px" }}>
        {/* Search */}
        <div style={{
          background:"rgba(255,255,255,0.97)", borderRadius:16, padding:32,
          boxShadow:"0 10px 50px #0008", marginBottom:24, animation:"fadeIn .4s ease"
        }}>
          <div style={{ fontSize:19, fontWeight:800, color:"#0a2a4a", marginBottom:4 }}>Check Your Results</div>
          <div style={{ fontSize:13, color:"#64748b", marginBottom:20 }}>Enter your KCSE index number below to view your official results slip.</div>
          <div style={{ display:"flex", gap:10 }}>
            <input value={idx} onChange={e => setIdx(e.target.value)}
              onKeyDown={e => e.key === "Enter" && lookup()}
              placeholder="e.g. 12345678"
              disabled={blocked}
              style={{ flex:1, padding:"13px 15px", border:"2px solid #e2e8f0", borderRadius:9, fontSize:15, boxSizing:"border-box" }}
            />
            <button onClick={lookup} disabled={blocked} style={{
              background: blocked ? "#94a3b8" : "#c8102e", color:"#fff", border:"none",
              borderRadius:9, padding:"13px 26px", fontWeight:700, fontSize:15, cursor: blocked ? "not-allowed" : "pointer",
              whiteSpace:"nowrap", transition:"background .2s"
            }}>Search</button>
          </div>
          {err && <div style={{ color:"#b91c1c", marginTop:10, fontSize:13, fontWeight:600 }}>⚠ {err}</div>}
        </div>

        {result && <ResultSlip {...result} />}
      </div>
    </div>
  );
}

function ResultSlip({ student, res, langUsed, totalPts, totalGrade, top5, allIds }) {
  const langLabel = SUB_MAP[langUsed]?.label;
  const otherRows = top5
    .filter(id => id !== "math")
    .map(id => ({ id, label: SUB_MAP[id]?.label, grade: res[id]?.grade, pts: res[id]?.points || 0 }))
    .sort((a, b) => b.pts - a.pts);
  const notCounted = allIds.filter(id => id !== "math" && id !== langUsed && !top5.includes(id));

  return (
    <div style={{ background:"rgba(255,255,255,0.98)", borderRadius:16, padding:32, boxShadow:"0 10px 50px #0008", animation:"fadeIn .4s ease" }}>
      <div style={{
        borderBottom:"3px double #c8102e", paddingBottom:18, marginBottom:20,
        display:"flex", justifyContent:"space-between", alignItems:"flex-start", flexWrap:"wrap", gap:12
      }}>
        <div>
          <div style={{ fontSize:10, letterSpacing:2, color:"#c8102e", fontWeight:700, textTransform:"uppercase" }}>Kenya National Examinations Council</div>
          <div style={{ fontSize:24, fontWeight:900, color:"#0a2a4a", marginTop:5 }}>{student.name}</div>
          <div style={{ fontSize:14, color:"#475569", marginTop:3 }}>{student.school}</div>
          <div style={{ fontSize:12, color:"#94a3b8", marginTop:2 }}>Index No: <strong style={{ color:"#334155" }}>{student.index}</strong></div>
        </div>
        <div style={{ textAlign:"right" }}>
          <div style={{ fontSize:10, color:"#94a3b8", marginBottom:4, letterSpacing:1, textTransform:"uppercase" }}>Overall Grade</div>
          <div style={{ fontSize:56, fontWeight:900, color:gradeColor(totalGrade), lineHeight:1 }}>{totalGrade}</div>
          <div style={{ fontSize:13, color:"#64748b", fontWeight:600, marginTop:2 }}>{totalPts} / 84 pts</div>
        </div>
      </div>

      <table style={{ width:"100%", borderCollapse:"collapse", fontSize:14 }}>
        <thead>
          <tr style={{ background:"#0a2a4a" }}>
            <th style={{ padding:"10px 14px", textAlign:"left", color:"#fff", fontWeight:700 }}>Subject</th>
            <th style={{ padding:"10px 14px", textAlign:"center", color:"#fff", fontWeight:700 }}>Grade</th>
            <th style={{ padding:"10px 14px", textAlign:"center", color:"#fff", fontWeight:700 }}>Points</th>
          </tr>
        </thead>
        <tbody>
          <SubRow label="Mathematics" grade={res["math"]?.grade} pts={res["math"]?.points} alt={false} />
          <SubRow
            label={<>{langLabel} <span style={{ fontSize:11, color:"#b45309", marginLeft:6 }}>(used for total)</span></>}
            grade={res[langUsed]?.grade} pts={res[langUsed]?.points} alt={true} highlight
          />
          {otherRows.map((r, i) => (
            <SubRow key={r.id} label={r.label} grade={r.grade} pts={r.pts} alt={i % 2 === 0} />
          ))}
        </tbody>
      </table>

      {notCounted.length > 0 && (
        <div style={{ marginTop:16, paddingTop:14, borderTop:"1px dashed #e2e8f0" }}>
          <div style={{ fontSize:11, color:"#94a3b8", fontWeight:700, textTransform:"uppercase", letterSpacing:1, marginBottom:8 }}>Not counted in total</div>
          {notCounted.map(id => (
            <div key={id} style={{ display:"flex", justifyContent:"space-between", padding:"5px 0", fontSize:13, color:"#94a3b8" }}>
              <span>{SUB_MAP[id]?.label}</span>
              <span style={{ fontWeight:700 }}>{res[id]?.grade || "-"}</span>
            </div>
          ))}
        </div>
      )}
      <div style={{ marginTop:18, padding:"10px 14px", background:"#f8fafc", borderRadius:8, fontSize:11, color:"#94a3b8", textAlign:"center" }}>
        This is a portal preview. Official results are issued by KNEC.
      </div>
    </div>
  );
}

function SubRow({ label, grade, pts, alt, highlight }) {
  return (
    <tr style={{ background: highlight ? "#fffbeb" : alt ? "#f8fafc" : "#fff", borderBottom:"1px solid #f1f5f9" }}>
      <td style={{ padding:"11px 14px", fontWeight:600, color:"#1e293b" }}>{label}</td>
      <td style={{ padding:"11px 14px", textAlign:"center" }}>
        <span style={{ fontSize:17, fontWeight:900, color:gradeColor(grade || "E") }}>{grade || "-"}</span>
      </td>
      <td style={{ padding:"11px 14px", textAlign:"center", color:"#64748b", fontWeight:700 }}>{pts || 0}</td>
    </tr>
  );
}

// ─── Admin Page ───────────────────────────────────────────────────────────────
function AdminPage({ students, setStudents, gradeBounds, setGradeBounds, totalBounds, setTotalBounds, formulas, setFormulas, showToast, onLogout }) {
  const [view, setView] = useState("students");
  const navItems = [
    { v:"students", label:"👤 Students" },
    { v:"grades",   label:"📊 Grade Boundaries" },
    { v:"formulas", label:"🧮 Formulas" },
    { v:"total",    label:"🏆 Total Boundaries" },
  ];

  return (
    <div style={{ position:"relative", zIndex:1, minHeight:"100vh", paddingBottom:50 }}>
      <div style={{
        background:"rgba(10,42,74,0.95)", backdropFilter:"blur(10px)",
        padding:"16px 28px", display:"flex", alignItems:"center", gap:16,
        borderBottom:"3px solid #c8102e", boxShadow:"0 2px 20px #0008"
      }}>
        <KenyaFlag size={36} />
        <div style={{ flex:1 }}>
          <div style={{ fontSize:18, fontWeight:900, color:"#fff" }}>KCSE Admin Panel</div>
          <div style={{ fontSize:11, color:"#94a3b8" }}>Superadmin · Restricted Access</div>
        </div>
        <a href="#" onClick={e => { e.preventDefault(); window.location.hash = ""; }} style={{ color:"#94a3b8", fontSize:13, textDecoration:"none", marginRight:8 }}>← Public Site</a>
        <button onClick={onLogout} style={{ background:"#c8102e", color:"#fff", border:"none", borderRadius:7, padding:"7px 18px", fontWeight:700, fontSize:13, cursor:"pointer" }}>Logout</button>
      </div>

      <div style={{ maxWidth:1200, margin:"0 auto", padding:"24px 16px" }}>
        <div style={{ display:"flex", gap:8, marginBottom:24, flexWrap:"wrap" }}>
          {navItems.map(({ v, label }) => (
            <button key={v} onClick={() => setView(v)} style={{
              background: view === v ? "#c8102e" : "rgba(255,255,255,0.93)",
              color: view === v ? "#fff" : "#0a2a4a",
              border:"none", borderRadius:9, padding:"9px 20px",
              fontWeight:700, fontSize:13, cursor:"pointer",
              boxShadow:"0 2px 10px #0003", transition:"all .2s"
            }}>{label}</button>
          ))}
        </div>

        {view === "students" && <StudentsPanel students={students} setStudents={setStudents} gradeBounds={gradeBounds} totalBounds={totalBounds} formulas={formulas} showToast={showToast} />}
        {view === "grades"   && <GradeBoundsPanel gradeBounds={gradeBounds} setGradeBounds={setGradeBounds} showToast={showToast} />}
        {view === "formulas" && <FormulasPanel formulas={formulas} setFormulas={setFormulas} showToast={showToast} />}
        {view === "total"    && <TotalBoundsPanel totalBounds={totalBounds} setTotalBounds={setTotalBounds} showToast={showToast} />}
      </div>
    </div>
  );
}

// ─── Students Panel ───────────────────────────────────────────────────────────
function StudentsPanel({ students, setStudents, gradeBounds, totalBounds, formulas, showToast }) {
  const emptyForm = { name:"", school:"", index:"", subjects:[], marks:{} };
  const [form, setForm] = useState(emptyForm);
  const [errors, setErrors] = useState([]);
  const [editIdx, setEditIdx] = useState(null);
  const [search, setSearch] = useState("");
  const fileRef = useRef();

  const toggleSub = (id) => {
    setForm(f => {
      const on = f.subjects.includes(id);
      return { ...f, subjects: on ? f.subjects.filter(x => x !== id) : [...f.subjects, id] };
    });
  };

  // FIX: use functional update so we always write to the latest marks state
  const setMark = (subId, paper, val) => {
    setForm(prev => {
      const subMarks = { ...(prev.marks[subId] || {}) };
      if (val === "" || val === undefined) {
        delete subMarks[paper];
      } else {
        subMarks[paper] = val;
      }
      return { ...prev, marks: { ...prev.marks, [subId]: subMarks } };
    });
  };

  const save = () => {
    const errs = [];
    if (!form.name.trim()) errs.push("Full name is required.");
    if (!form.school.trim()) errs.push("School name is required.");
    if (!form.index.trim()) errs.push("Index number is required.");
    const compIds = SUBJECTS.filter(s => s.compulsory).map(s => s.id);
    if ([...compIds, ...form.subjects].length > 8) errs.push("Maximum 8 subjects allowed.");
    errs.push(...validateChoice(form.subjects));
    if (errs.length) { setErrors(errs); return; }

    const student = {
      name: sanitize(form.name), school: sanitize(form.school), index: sanitize(form.index),
      subjects: form.subjects, marks: form.marks
    };
    if (editIdx !== null) {
      setStudents(ss => ss.map((s, i) => i === editIdx ? student : s));
      showToast("Student updated."); setEditIdx(null);
    } else {
      if (students.find(s => s.index === student.index)) { setErrors(["Index number already exists."]); return; }
      setStudents(ss => [...ss, student]);
      showToast("Student saved.");
    }
    setForm(emptyForm); setErrors([]);
  };

  const cancel = () => { setForm(emptyForm); setEditIdx(null); setErrors([]); };

  const startEdit = (realIdx) => {
    const s = students[realIdx];
    setForm({ name:s.name, school:s.school, index:s.index, subjects:[...(s.subjects||[])], marks: JSON.parse(JSON.stringify(s.marks||{})) });
    setEditIdx(realIdx);
    window.scrollTo({ top:0, behavior:"smooth" });
  };

  const del = (realIdx) => {
    if (window.confirm("Delete " + students[realIdx].name + "?")) {
      setStudents(ss => ss.filter((_, i) => i !== realIdx));
      showToast("Student deleted.");
    }
  };

  // File import
  const handleFile = async (e) => {
    const file = e.target.files[0]; if (!file) return;
    try {
      const data = await file.arrayBuffer();
      const wb = XLSX.read(data);
      const ws = wb.Sheets[wb.SheetNames[0]];
      const rows = XLSX.utils.sheet_to_json(ws, { defval:"" });
      if (!rows.length) { showToast("File is empty.", "error"); e.target.value = ""; return; }

      const imported = [];
      for (const row of rows) {
        const name  = String(row.name || row.Name || row.FULL_NAME || "").trim();
        const school= String(row.school || row.School || row.SCHOOL || "").trim();
        const index = String(row.index || row.Index || row.INDEX || row.index_no || row["Index No"] || "").trim();
        if (!name || !index) continue;

        const marks = {}; const subjects = [];
        SUBJECTS.forEach(sub => {
          sub.papers.forEach(p => {
            const col = sub.id + "_" + p;
            const val = row[col];
            if (val !== "" && val !== undefined && val !== null) {
              if (!marks[sub.id]) marks[sub.id] = {};
              marks[sub.id][p] = Number(val);
              if (!sub.compulsory && !subjects.includes(sub.id)) subjects.push(sub.id);
            }
          });
        });
        imported.push({ name:sanitize(name), school:sanitize(school||"Unknown"), index:sanitize(index), subjects, marks });
      }
      if (!imported.length) { showToast("No valid rows found. Check column format.", "error"); e.target.value = ""; return; }
      const existing = new Set(students.map(s => s.index));
      const newOnes = imported.filter(s => !existing.has(s.index));
      setStudents(ss => [...ss, ...newOnes]);
      showToast("Imported " + newOnes.length + " student(s). " + (imported.length - newOnes.length) + " skipped (duplicate index).");
    } catch (err) {
      showToast("File error: " + err.message, "error");
    }
    e.target.value = "";
  };

  const compIds = SUBJECTS.filter(s => s.compulsory).map(s => s.id);
  const allFormSubs = [...compIds, ...form.subjects];

  const filtered = students.filter(s =>
    s.name.toLowerCase().includes(search.toLowerCase()) ||
    s.index.toLowerCase().includes(search.toLowerCase()) ||
    (s.school || "").toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div style={{ display:"grid", gridTemplateColumns:"minmax(0,1fr) minmax(0,1fr)", gap:20, alignItems:"start" }}>
      {/* Form card */}
      <div style={{ background:"rgba(255,255,255,0.97)", borderRadius:14, padding:24, boxShadow:"0 6px 30px #0006" }}>
        <h3 style={{ margin:"0 0 16px", color:"#0a2a4a", fontSize:16 }}>{editIdx !== null ? "✏️ Edit Student" : "➕ Add Student"}</h3>

        {errors.map((e, i) => (
          <div key={i} style={{ background:"#fee2e2", color:"#b91c1c", borderRadius:7, padding:"7px 12px", marginBottom:6, fontSize:12, fontWeight:600 }}>⚠ {e}</div>
        ))}

        {/* Basic info */}
        {[["Full Name","name","text"],["School","school","text"],["Index Number","index","text"]].map(([lbl, key, type]) => (
          <div key={key} style={{ marginBottom:11 }}>
            <label style={{ display:"block", fontSize:11, fontWeight:700, color:"#475569", marginBottom:3, textTransform:"uppercase", letterSpacing:.5 }}>{lbl}</label>
            <input type={type} value={form[key]}
              onChange={e => setForm(f => ({ ...f, [key]: e.target.value }))}
              style={{ width:"100%", padding:"9px 11px", border:"2px solid #e2e8f0", borderRadius:7, fontSize:14, boxSizing:"border-box" }} />
          </div>
        ))}

        {/* File upload */}
        <div style={{ marginBottom:14 }}>
          <div style={{ fontSize:11, fontWeight:700, color:"#475569", marginBottom:5, textTransform:"uppercase", letterSpacing:.5 }}>Import from File</div>
          <div style={{ border:"2px dashed #c8102e", borderRadius:9, padding:"14px", background:"#fff8f8", cursor:"pointer", textAlign:"center" }}
            onClick={() => fileRef.current.click()}
            onDragOver={e => e.preventDefault()}
            onDrop={e => { e.preventDefault(); const f = e.dataTransfer.files[0]; if(f) { const dt = new DataTransfer(); dt.items.add(f); fileRef.current.files = dt.files; handleFile({ target: fileRef.current }); } }}>
            <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" onChange={handleFile} style={{ display:"none" }} />
            <div style={{ fontSize:22, marginBottom:4 }}>📎</div>
            <div style={{ fontSize:13, color:"#c8102e", fontWeight:700 }}>Click or drag to upload</div>
            <div style={{ fontSize:11, color:"#94a3b8", marginTop:4 }}>.xlsx · .xls · .csv</div>
            <div style={{ fontSize:10, color:"#cbd5e1", marginTop:6, fontFamily:"monospace" }}>
              Columns: name, school, index, math_pp1, math_pp2, eng_pp1, eng_pp2, eng_pp3, swa_pp1...
            </div>
          </div>
        </div>

        {/* Optional subjects */}
        <div style={{ marginBottom:12 }}>
          <div style={{ fontSize:11, fontWeight:700, color:"#475569", marginBottom:6, textTransform:"uppercase", letterSpacing:.5 }}>Optional Subjects</div>
          <div style={{ background:"#f0fdf4", borderRadius:8, padding:"10px 12px", marginBottom:8, border:"1px solid #d1fae5" }}>
            <div style={{ fontSize:11, fontWeight:700, color:"#15803d", marginBottom:6 }}>Group A · choose 2–3</div>
            <div style={{ display:"flex", flexWrap:"wrap", gap:5 }}>
              {["phyc","geo","hist","cre"].map(id => {
                const on = form.subjects.includes(id);
                return <button key={id} onClick={() => toggleSub(id)} style={{
                  background: on ? "#15803d" : "#fff", color: on ? "#fff" : "#334155",
                  border: "1.5px solid " + (on ? "#15803d" : "#d1fae5"),
                  borderRadius:20, padding:"4px 13px", fontSize:12, cursor:"pointer", fontWeight:on?700:400
                }}>{SUB_MAP[id].label}</button>;
              })}
            </div>
          </div>
          <div style={{ background:"#eff6ff", borderRadius:8, padding:"10px 12px", border:"1px solid #bfdbfe" }}>
            <div style={{ fontSize:11, fontWeight:700, color:"#1e40af", marginBottom:6 }}>Group B · at most 1</div>
            <div style={{ display:"flex", flexWrap:"wrap", gap:5 }}>
              {["bst","agric","comp"].map(id => {
                const on = form.subjects.includes(id);
                return <button key={id} onClick={() => toggleSub(id)} style={{
                  background: on ? "#1e40af" : "#fff", color: on ? "#fff" : "#334155",
                  border: "1.5px solid " + (on ? "#1e40af" : "#bfdbfe"),
                  borderRadius:20, padding:"4px 13px", fontSize:12, cursor:"pointer", fontWeight:on?700:400
                }}>{SUB_MAP[id].label}</button>;
              })}
            </div>
          </div>
        </div>

        {/* Marks — ALL subjects, compulsory shown always */}
        <div style={{ fontSize:11, fontWeight:700, color:"#475569", marginBottom:8, textTransform:"uppercase", letterSpacing:.5 }}>Marks Entry</div>
        {allFormSubs.map(id => {
          const sub = SUB_MAP[id];
          return (
            <div key={id} style={{ marginBottom:10, background:"#f8fafc", borderRadius:9, padding:"11px 13px", border:"1.5px solid #e2e8f0" }}>
              <div style={{ fontSize:12, fontWeight:700, color:"#0a2a4a", marginBottom:7, display:"flex", alignItems:"center", gap:6 }}>
                {sub.label}
                {sub.compulsory && <span style={{ fontSize:10, background:"#c8102e", color:"#fff", borderRadius:10, padding:"1px 8px", fontWeight:600 }}>Compulsory</span>}
              </div>
              <div style={{ display:"flex", gap:8 }}>
                {sub.papers.map(p => (
                  <div key={p} style={{ flex:1 }}>
                    <label style={{ fontSize:10, color:"#94a3b8", display:"block", marginBottom:3, fontWeight:700, textTransform:"uppercase" }}>{p}</label>
                    <input
                      type="number" min="0" max="200"
                      value={form.marks[id]?.[p] ?? ""}
                      onChange={e => setMark(id, p, e.target.value === "" ? "" : e.target.value)}
                      style={{
                        width:"100%", padding:"8px 6px", fontSize:14, textAlign:"center",
                        border:"2px solid #e2e8f0", borderRadius:7, boxSizing:"border-box",
                        background:"#fff"
                      }}
                    />
                  </div>
                ))}
              </div>
            </div>
          );
        })}

        <div style={{ display:"flex", gap:8, marginTop:16 }}>
          <button onClick={save} style={{ flex:1, background:"#c8102e", color:"#fff", border:"none", borderRadius:9, padding:12, fontWeight:700, fontSize:14, cursor:"pointer" }}>
            {editIdx !== null ? "Update Student" : "Save Student"}
          </button>
          {editIdx !== null && (
            <button onClick={cancel} style={{ background:"#f1f5f9", color:"#475569", border:"none", borderRadius:9, padding:"12px 18px", cursor:"pointer", fontWeight:600 }}>Cancel</button>
          )}
        </div>
      </div>

      {/* Student list */}
      <div style={{ background:"rgba(255,255,255,0.97)", borderRadius:14, padding:24, boxShadow:"0 6px 30px #0006" }}>
        <div style={{ display:"flex", justifyContent:"space-between", alignItems:"center", marginBottom:12 }}>
          <h3 style={{ margin:0, color:"#0a2a4a", fontSize:16 }}>Students ({students.length})</h3>
          {students.length > 0 && <span style={{ fontSize:12, color:"#64748b" }}>Showing {filtered.length}</span>}
        </div>
        <input value={search} onChange={e => setSearch(e.target.value)}
          placeholder="Search name, index, school..."
          style={{ width:"100%", padding:"9px 12px", border:"2px solid #e2e8f0", borderRadius:8, fontSize:13, boxSizing:"border-box", marginBottom:12 }} />
        {filtered.length === 0 && (
          <div style={{ color:"#94a3b8", fontSize:13, textAlign:"center", padding:30 }}>
            {students.length === 0 ? "No students yet. Add or import one." : "No matches."}
          </div>
        )}
        <div style={{ maxHeight:520, overflowY:"auto", paddingRight:2 }}>
          {filtered.map((s) => {
            const realIdx = students.indexOf(s);
            return (
              <div key={s.index + realIdx} style={{ background:"#f8fafc", borderRadius:9, padding:"12px 14px", marginBottom:8, display:"flex", justifyContent:"space-between", alignItems:"center", border:"1px solid #e2e8f0" }}>
                <div>
                  <div style={{ fontWeight:700, color:"#1e293b", fontSize:13 }}>{s.name}</div>
                  <div style={{ fontSize:11, color:"#64748b", marginTop:2 }}>{s.school} · <code style={{ background:"#e2e8f0", borderRadius:3, padding:"0 4px" }}>{s.index}</code></div>
                  <div style={{ fontSize:11, color:"#94a3b8", marginTop:1 }}>{[...SUBJECTS.filter(sub=>sub.compulsory).map(sub=>sub.id), ...(s.subjects||[])].length} subjects</div>
                </div>
                <div style={{ display:"flex", gap:6 }}>
                  <button onClick={() => startEdit(realIdx)} style={{ background:"#dbeafe", color:"#1d4ed8", border:"none", borderRadius:6, padding:"5px 12px", cursor:"pointer", fontSize:11, fontWeight:700 }}>Edit</button>
                  <button onClick={() => del(realIdx)} style={{ background:"#fee2e2", color:"#b91c1c", border:"none", borderRadius:6, padding:"5px 12px", cursor:"pointer", fontSize:11, fontWeight:700 }}>Delete</button>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// ─── Grade Bounds Panel ───────────────────────────────────────────────────────
function GradeBoundsPanel({ gradeBounds, setGradeBounds, showToast }) {
  const set = (subId, grade, val) =>
    setGradeBounds(b => ({ ...b, [subId]: { ...b[subId], [grade]: Number(val) } }));
  return (
    <div style={{ background:"rgba(255,255,255,0.97)", borderRadius:14, padding:24, boxShadow:"0 6px 30px #0006" }}>
      <h3 style={{ margin:"0 0 6px", color:"#0a2a4a" }}>Subject Grade Boundaries</h3>
      <p style={{ color:"#64748b", fontSize:13, marginBottom:18 }}>Set the minimum total mark for each grade, per subject.</p>
      <div style={{ overflowX:"auto" }}>
        <table style={{ borderCollapse:"collapse", width:"100%", fontSize:12 }}>
          <thead>
            <tr style={{ background:"#0a2a4a" }}>
              <th style={{ padding:"10px 12px", textAlign:"left", color:"#fff", fontWeight:700, minWidth:130 }}>Subject</th>
              {GRADE_LABELS.map(g => (
                <th key={g} style={{ padding:"10px 5px", color:["A","A-"].includes(g)?"#86efac":["B+","B","B-"].includes(g)?"#93c5fd":["C+","C","C-"].includes(g)?"#fde68a":"#fca5a5", fontWeight:700 }}>{g}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {SUBJECTS.map((sub, ri) => (
              <tr key={sub.id} style={{ background: ri%2===0?"#fff":"#f8fafc" }}>
                <td style={{ padding:"7px 12px", fontWeight:600, color:"#334155", whiteSpace:"nowrap", fontSize:12 }}>{sub.label}</td>
                {GRADE_LABELS.map(g => (
                  <td key={g} style={{ padding:"4px 2px" }}>
                    <input type="number" min="0" max="500"
                      value={gradeBounds[sub.id]?.[g] ?? 0}
                      onChange={e => set(sub.id, g, e.target.value)}
                      style={{ width:46, padding:"4px 4px", border:"1.5px solid #e2e8f0", borderRadius:5, fontSize:11, textAlign:"center" }} />
                  </td>
                ))}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      <button onClick={() => showToast("Grade boundaries saved.")} style={{ marginTop:16, background:"#c8102e", color:"#fff", border:"none", borderRadius:8, padding:"10px 26px", fontWeight:700, cursor:"pointer" }}>Save</button>
    </div>
  );
}

// ─── Formulas Panel ───────────────────────────────────────────────────────────
function FormulasPanel({ formulas, setFormulas, showToast }) {
  return (
    <div style={{ background:"rgba(255,255,255,0.97)", borderRadius:14, padding:24, boxShadow:"0 6px 30px #0006", maxWidth:700 }}>
      <h3 style={{ margin:"0 0 4px", color:"#0a2a4a" }}>Mark Calculation Formulas</h3>
      <p style={{ color:"#64748b", fontSize:13, marginBottom:20 }}>
        Variables: <code style={{ background:"#f1f5f9", padding:"1px 7px", borderRadius:4 }}>pp1</code> <code style={{ background:"#f1f5f9", padding:"1px 7px", borderRadius:4 }}>pp2</code> <code style={{ background:"#f1f5f9", padding:"1px 7px", borderRadius:4 }}>pp3</code>&nbsp;&nbsp;
        Example: <code style={{ background:"#f1f5f9", padding:"1px 7px", borderRadius:4 }}>pp1 + pp2 * 0.6</code>
      </p>
      {SUBJECTS.map(sub => (
        <div key={sub.id} style={{ display:"flex", alignItems:"center", gap:12, marginBottom:10 }}>
          <label style={{ width:160, fontWeight:600, color:"#334155", fontSize:13, flexShrink:0 }}>{sub.label}</label>
          <input value={formulas[sub.id] || ""}
            onChange={e => setFormulas(f => ({ ...f, [sub.id]: e.target.value }))}
            style={{ flex:1, padding:"8px 12px", border:"2px solid #e2e8f0", borderRadius:7, fontSize:13, fontFamily:"monospace" }} />
          <span style={{ fontSize:11, color:"#94a3b8", flexShrink:0 }}>({sub.papers.join(", ")})</span>
        </div>
      ))}
      <button onClick={() => showToast("Formulas saved.")} style={{ marginTop:12, background:"#c8102e", color:"#fff", border:"none", borderRadius:8, padding:"10px 26px", fontWeight:700, cursor:"pointer" }}>Save</button>
    </div>
  );
}

// ─── Total Bounds Panel ───────────────────────────────────────────────────────
function TotalBoundsPanel({ totalBounds, setTotalBounds, showToast }) {
  return (
    <div style={{ background:"rgba(255,255,255,0.97)", borderRadius:14, padding:24, boxShadow:"0 6px 30px #0006", maxWidth:460 }}>
      <h3 style={{ margin:"0 0 4px", color:"#0a2a4a" }}>Total Grade Boundaries</h3>
      <p style={{ color:"#64748b", fontSize:13, marginBottom:20 }}>Minimum total points for each overall grade. Maximum = 84 points.</p>
      {GRADE_LABELS.map(g => (
        <div key={g} style={{ display:"flex", alignItems:"center", gap:14, marginBottom:10 }}>
          <span style={{ width:36, fontWeight:900, fontSize:20, color:gradeColor(g) }}>{g}</span>
          <input type="number" min="0" max="84"
            value={totalBounds[g] ?? 0}
            onChange={e => setTotalBounds(b => ({ ...b, [g]: Number(e.target.value) }))}
            style={{ width:80, padding:"8px 10px", border:"2px solid #e2e8f0", borderRadius:7, fontSize:14, textAlign:"center" }} />
          <span style={{ fontSize:12, color:"#94a3b8" }}>points and above</span>
        </div>
      ))}
      <button onClick={() => showToast("Total boundaries saved.")} style={{ marginTop:12, background:"#c8102e", color:"#fff", border:"none", borderRadius:8, padding:"10px 26px", fontWeight:700, cursor:"pointer" }}>Save</button>
    </div>
  );
}
