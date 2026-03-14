import { useState, useCallback, useEffect, useRef } from "react";
import { Plus, X, LayoutList, Columns2, GripVertical, Printer, Save,
  User, HandMetal, Music2, Piano, Mic2, Sparkles, Megaphone,
  ClipboardList, Plane, BookOpen, Calendar, Tag, Star,
  FileText, Bell, Lock, Building2, RefreshCw,
  ChevronLeft, ChevronRight, MessageSquare,
  Landmark, ExternalLink, RotateCcw, Settings2,
  Check, Info, AlertTriangle, XCircle,
  ArrowDown, ArrowUp, LogOut } from "lucide-react";
import { useAuth } from "./useAuth";
import { pullAll, pushAll, pullSacrament, pushSacrament, testConnection } from "./sheets";
import config from "./config";

// ─── Constants ────────────────────────────────────────────────────────────────
const APPT_STAGES      = ["Need to Schedule","Contacted","Scheduled","Completed"];
const CALLING_STAGES   = ["Discuss","Approved to Call","Accepted","Sustained","Set Apart","Completed"];
const RELEASING_STAGES = ["Discuss","Approved to Release","Releasing Acknowledged","Thanked at Pulpit","Released","Completed"];
const LEADERS          = ["Bishop","1st Counsellor","2nd Counsellor"];
const ROSTER_POSITIONS = [
  {role:"Bishop",           group:"bishopric"},
  {role:"1st Counsellor",   group:"bishopric"},
  {role:"2nd Counsellor",   group:"bishopric"},
  {role:"Executive Secretary",group:"bishopric"},
  {role:"Clerk",            group:"bishopric"},
  {role:"Assigned HC",      group:"bishopric"},
  {role:"EQ President",     group:"ward_council"},
  {role:"SS President",     group:"ward_council"},
  {role:"Primary President",group:"ward_council"},
  {role:"YW President",     group:"ward_council"},
  {role:"RS President",     group:"ward_council"},
];
// Returns name for a role from roster array, or falls back to role label
function rosterName(roster, role) {
  const entry = roster.find(r => r.role === role);
  return (entry && entry.name) ? entry.name : role;
}
// Returns initials from a display name
function nameInitials(name) {
  if (!name) return "?";
  const parts = name.trim().split(/\s+/);
  if (parts.length === 1) return parts[0][0].toUpperCase();
  return (parts[0][0] + parts[parts.length-1][0]).toUpperCase();
}
const PURPOSE_OPTIONS  = ["Annual Youth","Semi-Annual Youth","Calling","Temple Recommend","Mission","Patriarchal Blessing","Ecclesiastical Endorsement","Follow Up","General","Releasing","Set Apart"];
const AUTO_PULL_MS = 60000;
const AUTO_PUSH_MS = 10000;

// ─── Church Style Guide §2.5 Palette ─────────────────────────────────────────
const C = {
  pageBg:"#F5F3EE", surfaceWhite:"#FFFFFF", surfaceWarm:"#EFEFE7",
  border:"#D5CFBE", borderLight:"#E9E6E0",
  blue40:"#003057", blue35:"#005581", blue25:"#007DA5", blue15:"#49CCE6", blue5:"#B0EEFC",
  gold20:"#F68D2E", goldDeep:"#C1A01E", gold10:"#DBBF6B",
  green25:"#50A83E", green35:"#206B3F", green15:"#93C742",
  red15:"#E10F5A", red10:"#FC4E6D",
  purple:"#7B5EA7",
  textPrimary:"#35383A", textSecond:"#53575B", textMuted:"#878A8C", textLight:"#BDC0C0",
};

const APPT_STAGE_STYLE = {
  "Need to Schedule":{ bg:"#FEF3E2",border:"#F68D2E",text:"#974A07",dot:"#F68D2E"},
  "Contacted":       { bg:"#E6F4FA",border:"#007DA5",text:"#005581",dot:"#007DA5"},
  "Scheduled":       { bg:"#EAF0F6",border:"#005581",text:"#003057",dot:"#005581"},
  "Completed":       { bg:"#EAF4EA",border:"#50A83E",text:"#206B3F",dot:"#50A83E"},
};
const PIPELINE_STAGE_STYLE = {
  "Discuss":                {bg:"#FEF3E2",border:"#F68D2E",text:"#974A07",dot:"#F68D2E"},
  "Approved to Call":       {bg:"#E6F4FA",border:"#007DA5",text:"#005581",dot:"#007DA5"},
  "Accepted":               {bg:"#EAF0F6",border:"#005581",text:"#003057",dot:"#005581"},
  "Sustained":              {bg:"#F3EDF8",border:"#7B5EA7",text:"#4A2A7A",dot:"#7B5EA7"},
  "Set Apart":              {bg:"#E8F4FD",border:"#49CCE6",text:"#00558F",dot:"#49CCE6"},
  "Approved to Release":    {bg:"#E6F4FA",border:"#007DA5",text:"#005581",dot:"#007DA5"},
  "Releasing Acknowledged": {bg:"#EAF0F6",border:"#005581",text:"#003057",dot:"#005581"},
  "Thanked at Pulpit":      {bg:"#F3EDF8",border:"#7B5EA7",text:"#4A2A7A",dot:"#7B5EA7"},
  "Released":               {bg:"#E8F4FD",border:"#49CCE6",text:"#00558F",dot:"#49CCE6"},
  "Completed":              {bg:"#EAF4EA",border:"#50A83E",text:"#206B3F",dot:"#50A83E"},
};
const OWNER_STYLE = {
  "Bishop":         {bg:"#FEF9E6",border:"#C1A01E",text:"#974A07",initials:"B"},
  "1st Counsellor": {bg:"#EAF0F6",border:"#007DA5",text:"#005581",initials:"1C"},
  "2nd Counsellor": {bg:"#EDF4F0",border:"#50A83E",text:"#206B3F",initials:"2C"},
};
const PURPOSE_ICON_MAP = {
  "Temple Recommend": Landmark,
  "Annual Youth": User,
  "Semi-Annual Youth": User,
  "Calling": ClipboardList,
  "Mission": Plane,
  "Patriarchal Blessing": BookOpen,
  "Ecclesiastical Endorsement": FileText,
  "Follow Up": RefreshCw,
  "General": MessageSquare,
  "Releasing": Tag,
  "Set Apart": Star,
};
function PurposeIcon({purpose,size=13,color}) {
  const Ic = PURPOSE_ICON_MAP[purpose];
  if(!Ic) return null;
  return <Ic size={size} color={color||"currentColor"} style={{flexShrink:0}}/>;
}

// ─── Global CSS ───────────────────────────────────────────────────────────────
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Crimson+Pro:ital,wght@0,300;0,400;0,600;1,300;1,400&display=swap');
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
body{background:#F5F3EE;font-family:Georgia,serif;}
::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-track{background:#EFEFE7;}
::-webkit-scrollbar-thumb{background:#D5CFBE;border-radius:3px;}
.tab-btn{background:none;border:none;cursor:pointer;font-family:inherit;transition:all .2s;}
@keyframes fadeUp{from{opacity:0;transform:translateY(10px);}to{opacity:1;transform:translateY(0);}}
@keyframes pulse{0%,100%{opacity:1;}50%{opacity:.3;}}
@keyframes slideIn{from{opacity:0;transform:translateX(100%);}to{opacity:1;transform:translateX(0);}}
.animate-in{animation:fadeUp .28s ease forwards;}
.toast-in{animation:slideIn .3s ease forwards;}
tr:hover td{background:rgba(0,85,129,.025);}
input[type="text"],input[type="date"],input[type="search"],input:not([type]),select,textarea{
  font-family:'Helvetica Neue',Arial,sans-serif;background:#EFEFE7;
  border:1.5px solid #D5CFBE;border-radius:7px;color:#35383A;
  padding:9px 13px;width:100%;font-size:14px;outline:none;transition:border .18s;
}
input[type="text"]:focus,input[type="search"]:focus,input:not([type]):focus,select:focus,textarea:focus{border-color:#007DA5;background:#fff;}
select option{background:#fff;color:#35383A;}
label{font-family:'Helvetica Neue',Arial,sans-serif;font-size:10px;font-weight:700;
  letter-spacing:.1em;text-transform:uppercase;color:#878A8C;display:block;margin-bottom:5px;}
.chip{display:inline-flex;align-items:center;gap:5px;padding:3px 10px;
  border-radius:20px;font-size:11px;font-weight:700;letter-spacing:.03em;
  border-width:1px;border-style:solid;white-space:nowrap;cursor:pointer;
  font-family:'Helvetica Neue',Arial,sans-serif;}
.chip-select{background:transparent;border:none;outline:none;cursor:pointer;
  font-weight:inherit;font-size:inherit;font-family:inherit;color:inherit;padding:0;}
.btn-primary{background:#005581;color:#fff;border:none;border-radius:8px;
  font-size:13px;font-weight:700;padding:9px 20px;cursor:pointer;
  transition:background .15s;display:inline-flex;align-items:center;gap:7px;
  font-family:'Helvetica Neue',Arial,sans-serif;letter-spacing:.01em;}
.btn-primary:hover{background:#007DA5;}
.btn-primary:disabled{opacity:.5;cursor:not-allowed;}
.btn-secondary{background:#fff;color:#53575B;border:1.5px solid #D5CFBE;border-radius:8px;
  font-size:12px;padding:7px 14px;cursor:pointer;transition:all .15s;
  display:inline-flex;align-items:center;gap:6px;
  font-family:'Helvetica Neue',Arial,sans-serif;}
.btn-secondary:hover{border-color:#007DA5;color:#007DA5;}
.btn-secondary:disabled{opacity:.5;cursor:not-allowed;}
.btn-del{background:transparent;border:1px solid #E9E6E0;border-radius:6px;
  color:#BDC0C0;width:28px;height:28px;cursor:pointer;font-size:11px;
  display:flex;align-items:center;justify-content:center;transition:all .15s;}
.btn-del:hover{border-color:#BD0057;color:#BD0057;background:#FDF0F4;}
.kanban-col{background:#EFEFE7;border:1.5px solid #D5CFBE;border-radius:14px;
  min-width:200px;flex:1;padding:15px 12px;}
.kanban-card{background:#fff;border:1.5px solid #E9E6E0;border-radius:10px;
  padding:14px;margin-bottom:9px;cursor:pointer;transition:all .15s;}
.kanban-card:hover{border-color:#007DA5;box-shadow:0 3px 14px rgba(0,85,129,.1);transform:translateY(-1px);}
.kanban-card[draggable]:active{cursor:grabbing;}
.kanban-card[draggable="true"]{user-select:none;}
.modal-overlay{position:fixed;inset:0;background:rgba(0,48,87,.45);
  backdrop-filter:blur(6px);z-index:200;
  display:flex;align-items:center;justify-content:center;padding:20px;}
.modal{background:#fff;border-radius:16px;width:100%;max-width:560px;
  overflow:hidden;box-shadow:0 24px 64px rgba(0,48,87,.18);}
.modal-lg{max-width:760px;}
.modal-hd{background:#005581;padding:22px 28px 20px;position:relative;overflow:hidden;}
.modal-hd::before{content:'';position:absolute;top:-20%;right:-5%;width:60%;height:160%;
  background:linear-gradient(240deg,rgba(255,255,255,.13) 0%,transparent 55%);pointer-events:none;}
.modal-hd::after{content:'';position:absolute;top:25%;right:10%;width:45%;height:130%;
  background:linear-gradient(240deg,rgba(255,255,255,.06) 0%,transparent 60%);pointer-events:none;}
.modal-body{padding:24px 28px 28px;max-height:70vh;overflow-y:auto;}
.wm-table{width:100%;border-collapse:collapse;}
.wm-table thead th{font-size:10px;font-weight:700;letter-spacing:.1em;text-transform:uppercase;
  color:#878A8C;padding:10px 16px;text-align:left;border-bottom:2px solid #D5CFBE;
  white-space:nowrap;font-family:'Helvetica Neue',Arial,sans-serif;}
.wm-table tbody tr{border-bottom:1px solid #E9E6E0;transition:background .1s;}
.wm-table tbody tr:hover{background:#F0EDE6;}
.wm-table tbody td{padding:9px 16px;vertical-align:middle;font-size:14px;color:#35383A;}
tr.row-done td{opacity:.45;}
.cell-text{cursor:text;display:block;padding:3px 6px;border-radius:4px;
  font-size:14px;color:#35383A;transition:background .1s;min-width:90px;}
.cell-text:hover{background:#EAE5DC;}
.cell-input{padding:4px 8px;font-size:14px;color:#35383A;background:#fff;
  border:1.5px solid #007DA5;border-radius:4px;outline:none;min-width:130px;}
/* Toast notifications */
.toast-stack{position:fixed;bottom:24px;right:24px;z-index:9999;display:flex;flex-direction:column;gap:8px;pointer-events:none;}
.toast{pointer-events:all;display:flex;align-items:center;gap:10px;padding:12px 16px;
  border-radius:10px;font-family:'Helvetica Neue',Arial,sans-serif;font-size:13px;
  box-shadow:0 4px 20px rgba(0,0,0,.15);min-width:260px;max-width:400px;
  border-left:4px solid transparent;}
.toast-success{background:#EAF4EA;border-color:#50A83E;color:#206B3F;}
.toast-error  {background:#FEF0F4;border-color:#E10F5A;color:#BD0057;}
.toast-info   {background:#E6F4FA;border-color:#007DA5;color:#005581;}
.toast-warn   {background:#FEF3E2;border-color:#F68D2E;color:#974A07;}
/* Connection badge */
.conn-badge{display:inline-flex;align-items:center;gap:5px;padding:3px 10px;
  border-radius:20px;font-size:11px;font-weight:700;font-family:'Helvetica Neue',Arial,sans-serif;}
/* Members search */
.member-row{display:flex;align-items:center;gap:12px;padding:10px 14px;border-radius:10px;
  border:1.5px solid #E9E6E0;background:#fff;cursor:pointer;transition:all .15s;margin-bottom:6px;}
.member-row:hover{border-color:#007DA5;background:#F0F7FA;}
.member-avatar{width:36px;height:36px;border-radius:50%;background:#EAF0F6;border:2px solid #D5CFBE;
  display:flex;align-items:center;justify-content:center;font-size:14px;font-weight:700;
  color:#005581;flex-shrink:0;font-family:'Helvetica Neue',Arial,sans-serif;}
/* Agenda items */
.agenda-item{background:#fff;border:1.5px solid #E9E6E0;border-radius:10px;padding:14px 16px;margin-bottom:8px;transition:border .15s;}
.agenda-item:hover{border-color:#D5CFBE;}
.agenda-item.done{opacity:.5;}
.agenda-item-drag{cursor:grab;}
`;

// ─── Toast System ─────────────────────────────────────────────────────────────
let _toastSet = null;
export function useToasts() {
  const [toasts, setToasts] = useState([]);
  _toastSet = setToasts;
  const dismiss = id => setToasts(t => t.filter(x => x.id !== id));
  return { toasts, dismiss };
}
function toast(type, msg, duration = 4000) {
  if (!_toastSet) return;
  const id = Date.now();
  _toastSet(t => [...t, { id, type, msg }]);
  setTimeout(() => _toastSet(t => t.filter(x => x.id !== id)), duration);
}
export const notify = {
  success: (m, d) => toast("success", m, d),
  error:   (m, d) => toast("error",   m, d),
  info:    (m, d) => toast("info",    m, d),
  warn:    (m, d) => toast("warn",    m, d),
};

function ToastStack() {
  const { toasts, dismiss } = useToasts();
  const icons = { success:<Check size={14}/>, error:<XCircle size={14}/>, info:<Info size={14}/>, warn:<AlertTriangle size={14}/> };
  return (
    <div className="toast-stack">
      {toasts.map(t => (
        <div key={t.id} className={`toast toast-${t.type} toast-in`}>
          <span style={{fontWeight:700,fontSize:16,flexShrink:0}}>{icons[t.type]}</span>
          <span style={{flex:1,lineHeight:1.5}}>{t.msg}</span>
          <button onClick={()=>dismiss(t.id)} style={{background:"none",border:"none",cursor:"pointer",color:"inherit",opacity:.6,padding:"0 0 0 6px",flexShrink:0,display:"flex",alignItems:"center"}}><X size={13}/></button>
        </div>
      ))}
    </div>
  );
}

// ─── Shared UI Primitives ─────────────────────────────────────────────────────
function HeroBanner({ title, sub, children }) {
  return (
    <div style={{background:`linear-gradient(130deg,${C.blue40} 0%,${C.blue35} 50%,${C.blue25} 100%)`,borderRadius:14,padding:"20px 26px 18px",marginBottom:22,position:"relative",overflow:"hidden"}}>
      <div style={{position:"absolute",inset:0,pointerEvents:"none"}}>
        <div style={{position:"absolute",top:0,right:0,width:"60%",height:"100%",background:"linear-gradient(250deg,rgba(255,255,255,.13) 0%,transparent 50%)"}}/>
        <div style={{position:"absolute",top:"30%",right:"-5%",width:"50%",height:"130%",background:"linear-gradient(250deg,rgba(255,255,255,.07) 0%,transparent 55%)"}}/>
      </div>
      <div style={{position:"relative",display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:12}}>
        <div>
          <div style={{fontFamily:"Georgia,serif",fontSize:22,fontWeight:400,color:"#fff",marginBottom:3}}>{title}</div>
          {sub && <div style={{fontFamily:"'Helvetica Neue',Arial,sans-serif",fontSize:12,color:"rgba(255,255,255,.7)",letterSpacing:".03em"}}>{sub}</div>}
        </div>
        {children && <div style={{display:"flex",gap:8,flexWrap:"wrap"}}>{children}</div>}
      </div>
    </div>
  );
}

function ModalShell({ onClose, title, subtitle, size, children }) {
  return (
    <div className="modal-overlay" onClick={e=>e.target===e.currentTarget&&onClose()}>
      <div className={`modal ${size==="lg"?"modal-lg":""} animate-in`}>
        <div className="modal-hd">
          <div style={{position:"relative",display:"flex",justifyContent:"space-between",alignItems:"flex-start"}}>
            <div>
              {subtitle && <div style={{fontFamily:"'Helvetica Neue',Arial,sans-serif",fontSize:10,fontWeight:700,letterSpacing:".12em",textTransform:"uppercase",color:"rgba(255,255,255,.6)",marginBottom:5}}>{subtitle}</div>}
              <div style={{fontFamily:"Georgia,serif",fontSize:22,color:"#fff"}}>{title}</div>
            </div>
            <button onClick={onClose} style={{background:"rgba(255,255,255,.15)",border:"1px solid rgba(255,255,255,.25)",borderRadius:8,color:"#fff",width:32,height:32,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0}}><XIcon/></button>
          </div>
        </div>
        <div className="modal-body">{children}</div>
      </div>
    </div>
  );
}

function FormRow({ label, children, full }) {
  return <div style={{gridColumn:full?"1 / -1":"auto"}}><label>{label}</label>{children}</div>;
}
function ModalFooter({ onClose, onSave, saveLabel="Save", saving }) {
  return (
    <div style={{display:"flex",justifyContent:"flex-end",gap:10,marginTop:22}}>
      <button onClick={onClose} className="btn-secondary">Cancel</button>
      <button onClick={onSave} className="btn-primary" disabled={saving}>{saving?"Saving…":saveLabel}</button>
    </div>
  );
}

// ─── Slack Preview Modal ─────────────────────────────────────────────────────
// Used by all three Slack send flows: Bishopric, Ward Council, Alerts.
// Shows editable first line + full message preview before sending.
function SlackPreviewModal({ firstLine, setFirstLine, bodyLines, channelName, sending, onConfirm, onClose }) {
  const previewText = [firstLine, ...bodyLines].join("\n");
  return (
    <ModalShell onClose={onClose} title="Preview Message" subtitle={`Send to #${channelName}`} size="lg">
      <div style={{display:"flex",flexDirection:"column",gap:20}}>

        {/* Editable first line */}
        <div>
          <label style={{display:"block",fontSize:11,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",
            color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginBottom:8}}>
            First Line
          </label>
          <input
            type="text"
            value={firstLine}
            onChange={e => setFirstLine(e.target.value)}
            style={{width:"100%",padding:"9px 14px",fontSize:14,fontFamily:"'Helvetica Neue',Arial,sans-serif",
              border:`1.5px solid ${C.blue25}`,borderRadius:8,boxSizing:"border-box",
              background:C.surfaceWarm,color:C.textPrimary,outline:"none"}}
          />
          <div style={{fontSize:11,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginTop:5}}>
            This is the bold header line that appears at the top of the Slack message.
          </div>
        </div>

        {/* Full message preview */}
        <div>
          <label style={{display:"block",fontSize:11,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",
            color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginBottom:8}}>
            Message Preview
          </label>
          <div style={{background:"#1a1d21",borderRadius:10,padding:"16px 18px",maxHeight:340,overflowY:"auto",
            border:`1px solid ${C.border}`}}>
            <pre style={{margin:0,whiteSpace:"pre-wrap",wordBreak:"break-word",
              fontFamily:"'SF Mono','Fira Code','Consolas',monospace",fontSize:12,
              color:"#d1d2d3",lineHeight:1.6}}>
              {previewText}
            </pre>
          </div>
          <div style={{fontSize:11,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginTop:5}}>
            Slack markdown (*bold*, _italic_) will render in the channel.
          </div>
        </div>

        {/* Actions */}
        <div style={{display:"flex",justifyContent:"flex-end",gap:10,paddingTop:4,borderTop:`1px solid ${C.borderLight}`}}>
          <button onClick={onClose} className="btn-secondary" disabled={sending}>Cancel</button>
          <button onClick={onConfirm} className="btn-primary" disabled={sending}
            style={{display:"flex",alignItems:"center",gap:7}}>
            <Bell size={13}/> {sending ? "Sending…" : `Send to #${channelName}`}
          </button>
        </div>

      </div>
    </ModalShell>
  );
}

function InlineText({ value, onChange, serif, placeholder }) {
  const [ed,setEd]=useState(false);const[draft,setDraft]=useState(value);
  const commit=()=>{setEd(false);if(draft!==value)onChange(draft);};
  if(ed) return <input autoFocus value={draft} className="cell-input"
    style={{fontFamily:serif?"Georgia,serif":"'Helvetica Neue',Arial,sans-serif"}}
    onChange={e=>setDraft(e.target.value)} onBlur={commit}
    onKeyDown={e=>{if(e.key==="Enter")commit();if(e.key==="Escape"){setDraft(value);setEd(false);}}}/>;
  return <span className="cell-text" style={{fontFamily:serif?"Georgia,serif":"'Helvetica Neue',Arial,sans-serif"}}
    onClick={()=>{setDraft(value);setEd(true);}} title="Click to edit">
    {value||<span style={{color:C.textLight,fontStyle:"italic"}}>{placeholder||"—"}</span>}
  </span>;
}

function StageChip({ status, onChange, stageList, styleMap }) {
  const map=styleMap||APPT_STAGE_STYLE;const list=stageList||APPT_STAGES;
  const st=map[status]||{bg:C.surfaceWarm,border:C.border,text:C.textSecond,dot:C.textMuted};
  return <span className="chip" style={{background:st.bg,borderColor:st.border,color:st.text}}>
    <span style={{width:6,height:6,borderRadius:"50%",background:st.dot,flexShrink:0}}/>
    <select value={status} onChange={e=>onChange(e.target.value)} className="chip-select">
      {list.map(s=><option key={s} value={s}>{s}</option>)}
    </select>
  </span>;
}

function OwnerChip({ owner, onChange, compact, roster=[] }) {
  const os=OWNER_STYLE[owner]||{bg:C.surfaceWarm,border:C.border,text:C.textSecond};
  const displayName = rosterName(roster, owner);
  const initials = nameInitials(displayName);
  return <span style={{display:"inline-flex",alignItems:"center",gap:compact?4:6,padding:compact?"2px 8px 2px 3px":"3px 10px 3px 4px",
    borderRadius:7,fontSize:compact?10:11,fontWeight:600,border:`1px solid ${os.border}`,
    background:os.bg,color:os.text,fontFamily:"'Helvetica Neue',Arial,sans-serif",cursor:onChange?"pointer":"default",whiteSpace:"nowrap"}}>
    <span style={{width:compact?16:20,height:compact?16:20,borderRadius:"50%",background:os.border,fontSize:compact?8:9,fontWeight:700,
      display:"flex",alignItems:"center",justifyContent:"center",color:"#fff",flexShrink:0}}>{initials}</span>
    {onChange
      ? <select value={owner} onChange={e=>onChange(e.target.value)} className="chip-select">
          {LEADERS.map(l=><option key={l} value={l}>{rosterName(roster,l)}</option>)}
        </select>
      : displayName}
  </span>;
}

// ─── NavGroup — grouped dropdown nav item ────────────────────────────────────
function NavGroup({ group, activeTab, isActiveGroup, onSelect }) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);

  // Close on outside click
  useEffect(() => {
    const handler = e => { if (ref.current && !ref.current.contains(e.target)) setOpen(false); };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, []);

  // Total badge count across children
  const totalBadge = group.children.reduce((sum, c) => sum + (c.badge || 0), 0);

  return (
    <div ref={ref} style={{ position: "relative" }}>
      <button
        className="tab-btn"
        onClick={() => setOpen(o => !o)}
        style={{
          padding: "5px 14px", borderRadius: 6, fontSize: 13,
          fontFamily: "'Helvetica Neue',Arial,sans-serif",
          color: isActiveGroup ? C.blue35 : C.textMuted,
          background: isActiveGroup ? "rgba(0,85,129,.07)" : "transparent",
          borderBottom: `2.5px solid ${isActiveGroup ? C.blue35 : "transparent"}`,
          fontWeight: isActiveGroup ? 700 : 400,
          display: "flex", alignItems: "center", gap: 5,
          position: "relative",
        }}>
        {group.label}
        <svg width="9" height="6" viewBox="0 0 9 6" fill="none" style={{
          transform: open ? "rotate(180deg)" : "none", transition: "transform .15s",
          opacity: .6, marginTop: 1
        }}>
          <path d="M1 1L4.5 5L8 1" stroke="currentColor" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/>
        </svg>
        {totalBadge > 0 && (
          <span style={{ position: "absolute", top: -5, right: -7, minWidth: 16, height: 16,
            borderRadius: 8, background: C.gold20, color: "#fff", fontSize: 9, fontWeight: 700,
            display: "flex", alignItems: "center", justifyContent: "center", padding: "0 3px",
            pointerEvents: "none", boxShadow: "0 1px 3px rgba(0,0,0,.2)" }}>
            {totalBadge}
          </span>
        )}
      </button>

      {open && (
        <div style={{
          position: "absolute", top: "calc(100% + 6px)", left: 0,
          background: "#fff", border: `1.5px solid ${C.border}`,
          borderRadius: 10, boxShadow: "0 8px 24px rgba(0,0,0,.12)",
          minWidth: 180, zIndex: 200, overflow: "hidden", padding: "4px 0",
        }}>
          {group.children.map(child => (
            <button
              key={child.id}
              onClick={() => { onSelect(child.id); setOpen(false); }}
              style={{
                display: "flex", alignItems: "center", justifyContent: "space-between",
                width: "100%", padding: "9px 16px",
                background: activeTab === child.id ? "rgba(0,85,129,.07)" : "transparent",
                border: "none", cursor: "pointer", textAlign: "left",
                fontFamily: "'Helvetica Neue',Arial,sans-serif", fontSize: 13,
                color: activeTab === child.id ? C.blue35 : C.textPrimary,
                fontWeight: activeTab === child.id ? 700 : 400,
              }}>
              <span>{child.label}</span>
              {child.badge > 0 && (
                <span style={{ minWidth: 18, height: 18, borderRadius: 9,
                  background: C.gold20, color: "#fff", fontSize: 10, fontWeight: 700,
                  display: "flex", alignItems: "center", justifyContent: "center", padding: "0 4px" }}>
                  {child.badge}
                </span>
              )}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}

// ─── Root ──────────────────────────────────────────────────────────────────────
export default function App() {
  const { user, token, error, loading, signIn, signOut } = useAuth();
  if (!user) return <LoginScreen onSignIn={signIn} error={error} loading={loading}/>;
  return <MainApp user={user} token={token} onSignOut={signOut}/>;
}

function LoginScreen({ onSignIn, error, loading }) {
  return (
    <div style={{minHeight:"100vh",background:C.pageBg,display:"flex",alignItems:"center",justifyContent:"center",fontFamily:"Georgia,serif"}}>
      <style>{CSS}</style>
      <div className="animate-in" style={{textAlign:"center",maxWidth:400,padding:32}}>
        <div style={{width:72,height:72,borderRadius:20,margin:"0 auto 24px",
          background:`linear-gradient(135deg,${C.blue35},${C.blue25})`,
          display:"flex",alignItems:"center",justifyContent:"center",
          position:"relative",overflow:"hidden",boxShadow:"0 8px 32px rgba(0,85,129,.25)"}}>
          <div style={{position:"absolute",top:"-20%",right:"-10%",width:"70%",height:"160%",background:"linear-gradient(240deg,rgba(255,255,255,.18) 0%,transparent 55%)",pointerEvents:"none"}}/>
          <Building2 size={32} style={{color:"#fff",position:"relative"}}/>
        </div>
        <h1 style={{fontSize:32,fontWeight:300,color:C.textPrimary,marginBottom:4}}>
          Ward <span style={{fontStyle:"italic",color:C.blue35}}>Manager</span>
        </h1>
        <p style={{fontSize:13,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginBottom:6}}>{config.WARD_NAME}</p>
        <p style={{fontSize:13,color:C.textSecond,marginBottom:32,lineHeight:1.6}}>Sign in with your authorised Google account to continue.</p>
        {error && <div style={{background:"#FEF0F4",border:"1px solid #FC4E6D",borderRadius:8,padding:"12px 16px",marginBottom:20,fontSize:13,color:"#BD0057",fontFamily:"'Helvetica Neue',Arial,sans-serif",lineHeight:1.6}}>{error}</div>}
        <button onClick={onSignIn} disabled={loading} className="btn-primary"
          style={{width:"100%",justifyContent:"center",fontSize:15,padding:"13px 24px",borderRadius:10}}>
          {loading?"Signing in…":"Sign in with Google"}
        </button>
      </div>
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────
function MainApp({ user, token, onSignOut }) {
  const [tab,setTab]               = useState("appointments");
  const [appointments,setAppts]    = useState([]);
  const [callings,setCallings]     = useState([]);
  const [releasings,setReleasings] = useState([]);
  const [bishopricMeeting,setBishopricMeeting] = useState([]);
  const [wardCouncilMeeting,setWardCouncilMeeting] = useState([]);
  const [members,setMembers]       = useState([]);
  const [sacramentProgram,setSacrament] = useState([]);
  const [calendar,setCalendar]     = useState([]);
  const [bishopricLinks,setBishopricLinks] = useState([]);
  const [wcLinks,setWcLinks]               = useState([]);
  const [hasPulled,setHasPulled]   = useState(false);
  const [syncStatus,setSyncStatus] = useState("idle");
  const [lastSyncedAt,setLastSync] = useState(null);
  const [syncError,setSyncError]   = useState(null);
  const [connStatus,setConnStatus] = useState("unknown"); // unknown | ok | error
  const [connMsg,setConnMsg]       = useState("");
  const [showSettings,setShowSettings] = useState(false);
  const [roster,setRoster]             = useState(ROSTER_POSITIONS.map(p=>({role:p.role,name:""})));

  // Compute role early so doPull can branch before isAdmin is declared below
  const userEmail = user?.email?.toLowerCase()||"";
  const isAdminEarly  = (config.ALLOWED_EMAILS||[]).map(e=>e.toLowerCase()).includes(userEmail);
  const isAdminRef = useRef(isAdminEarly);
  useEffect(()=>{ isAdminRef.current = isAdminEarly; },[isAdminEarly]);

  const tokenRef=useRef(token),apptRef=useRef(appointments),callRef=useRef(callings),
    relRef=useRef(releasings),bmRef=useRef(bishopricMeeting),membRef=useRef(members),rosterRef=useRef(roster),
    wcmRef=useRef(wardCouncilMeeting),
    sacrRef=useRef(sacramentProgram),tabRef=useRef(tab),calendarRef=useRef(calendar),
    pulledRef=useRef(hasPulled),pushTimer=useRef(null),pushing=useRef(false);

  useEffect(()=>{tokenRef.current=token;},[token]);
  useEffect(()=>{apptRef.current=appointments;},[appointments]);
  useEffect(()=>{callRef.current=callings;},[callings]);
  useEffect(()=>{relRef.current=releasings;},[releasings]);
  useEffect(()=>{bmRef.current=bishopricMeeting;},[bishopricMeeting]);
  useEffect(()=>{membRef.current=members;},[members]);
  useEffect(()=>{rosterRef.current=roster;},[roster]);
  useEffect(()=>{tabRef.current=tab;},[tab]);
  useEffect(()=>{sacrRef.current=sacramentProgram;},[sacramentProgram]);
  useEffect(()=>{wcmRef.current=wardCouncilMeeting;},[wardCouncilMeeting]);
  useEffect(()=>{calendarRef.current=calendar;},[calendar]);
  useEffect(()=>{pulledRef.current=hasPulled;},[hasPulled]);

  const doTestConnection = useCallback(async() => {
    setConnStatus("testing");
    try {
      const result = await testConnection(tokenRef.current);
      if (result.missing.length > 0) {
        setConnStatus("warn");
        setConnMsg(`Missing tabs: ${result.missing.join(", ")}`);
        notify.warn(`Sheet connected but missing tabs: ${result.missing.join(", ")}`);
      } else {
        setConnStatus("ok");
        setConnMsg(`Connected · ${result.tabs.length} tabs found`);
        notify.success("Google Sheets connected successfully");
      }
    } catch(e) {
      setConnStatus("error");
      setConnMsg(e.message);
      notify.error("Connection failed: " + e.message, 8000);
    }
  },[]);

  const doPull = useCallback(async(silent=false) => {
    if(!silent) setSyncStatus("pulling");
    try {
      if(isAdminRef.current) {
        // ── Admin: pull everything from the Bishopric sheet ──
        const d = await pullAll(tokenRef.current);
        setAppts(d.appointments); setCallings(d.callings);
        setReleasings(d.releasings); setBishopricMeeting(d.bishopricMeeting||[]);
        if(d.members) setMembers(d.members);
        if(d.roster) setRoster(d.roster);
        // Sacrament (non-blocking)
        try {
          const sp = await pullSacrament(tokenRef.current);
          if(sp.sacramentProgram){ setSacrament(sp.sacramentProgram); sacrRef.current=sp.sacramentProgram; }
        } catch(e) { /* sacrament sheet not yet configured */ }
      }
      // ── Both roles: pull Ward Council sheet, Calendar, and Roster ──
      try {
        const { pullCalendar, pullWardCouncilMeeting, pullRosterFromWardCouncil } = await import("./sheets");
        const [cp, wc, ros] = await Promise.all([
          pullCalendar(tokenRef.current),
          pullWardCouncilMeeting(tokenRef.current),
          pullRosterFromWardCouncil(tokenRef.current),
        ]);
        if(cp.calendar){ setCalendar(cp.calendar); calendarRef.current=cp.calendar; }
        if(wc.wardCouncilMeeting){ setWardCouncilMeeting(wc.wardCouncilMeeting); wcmRef.current=wc.wardCouncilMeeting; }
        if(ros.roster?.length){ setRoster(ros.roster); rosterRef.current=ros.roster; }
        // Pull links for both sheets
        try {
          const { pullBishopricLinks, pullWardCouncilLinks } = await import("./sheets");
          if(isAdminRef.current) {
            const bl = await pullBishopricLinks(tokenRef.current);
            if(bl.links) setBishopricLinks(bl.links);
          }
          const wl = await pullWardCouncilLinks(tokenRef.current);
          if(wl.links) setWcLinks(wl.links);
        } catch(_) {}
      } catch(e) { /* ward council sheet not yet configured */ }
      setHasPulled(true); setLastSync(new Date()); setSyncError(null);
      if(!silent){ setSyncStatus("idle"); notify.success("Data pulled from Google Sheets"); }
      if(connStatus==="unknown") setConnStatus("ok");
    } catch(e) {
      setSyncError(e.message); setSyncStatus("error");
      notify.error("Pull failed: " + e.message, 8000);
      setTimeout(()=>setSyncStatus("idle"),5000);
    }
  },[connStatus]);

  const doPush = useCallback(async() => {
    if(!pulledRef.current||pushing.current||!isAdminRef.current) return;
    pushing.current=true; setSyncStatus("pushing");
    try {
      await pushAll(tokenRef.current,{
        appointments:apptRef.current, callings:callRef.current,
        releasings:relRef.current, members:membRef.current,
      });
      setLastSync(new Date()); setSyncError(null); setSyncStatus("idle");
    } catch(e) {
      setSyncError(e.message); setSyncStatus("error");
      notify.error("Save failed: " + e.message, 8000);
      setTimeout(()=>setSyncStatus("idle"),5000);
    }
    pushing.current=false;
  },[]);

  // Sacrament save — called by SacramentTab's useMeetingSync hook
  const doSacramentSave = useCallback(async(data) => {
    if(!pulledRef.current) return;
    await pushSacrament(tokenRef.current, data ?? sacrRef.current);
  },[]);

  // Sacrament pull — used by useMeetingSync for pending-changes detection
  const doSacramentPull = useCallback(async() => {
    const sp = await pullSacrament(tokenRef.current);
    return sp.sacramentProgram || [];
  },[]);

  const pull = useCallback(()=>doPull(false),[doPull]);
  const push = useCallback(()=>{clearTimeout(pushTimer.current);doPush();},[doPush]);

  // Auto-pull on mount
  useEffect(()=>{ doPull(false); },[]);// eslint-disable-line

  useEffect(()=>{
    let id=null;
    const start=()=>{if(id)return;id=setInterval(()=>{
      // Skip background pull while on Sacrament or Bishopric tabs — avoids clobbering unsaved work
      const t=tabRef.current;
      if(pulledRef.current&&t!=="sacrament"&&t!=="bishopric"&&t!=="ward-council")doPull(true);
    },AUTO_PULL_MS);};
    const stop=()=>{clearInterval(id);id=null;};
    const onVis=()=>{const t=tabRef.current;if(document.visibilityState==="visible"){if(pulledRef.current&&t!=="sacrament"&&t!=="bishopric"&&t!=="ward-council")doPull(true);start();}else stop();};
    if(document.visibilityState==="visible")start();
    document.addEventListener("visibilitychange",onVis);
    return()=>{stop();document.removeEventListener("visibilitychange",onVis);};
  },[doPull]);

  // When leaving the Sacrament or Bishopric tab, resync that tab's data from the sheet
  const prevTabRef=useRef(tab);
  useEffect(()=>{
    const prev=prevTabRef.current;
    // Leaving Sacrament — pull fresh sacrament data
    if(prev==="sacrament"&&tab!=="sacrament"&&pulledRef.current){
      pullSacrament(tokenRef.current).then(sp=>{
        if(sp.sacramentProgram){setSacrament(sp.sacramentProgram);sacrRef.current=sp.sacramentProgram;}
      }).catch(()=>{});
    }
    // Leaving Bishopric — pull fresh bishopric meeting data (admin only)
    if(prev==="bishopric"&&tab!=="bishopric"&&pulledRef.current&&isAdminRef.current){
      pullAll(tokenRef.current).then(d=>{
        if(d.bishopricMeeting){setBishopricMeeting(d.bishopricMeeting);bmRef.current=d.bishopricMeeting;}
      }).catch(()=>{});
    }
    prevTabRef.current=tab;
  },[tab]); // eslint-disable-line

  useEffect(()=>{
    if(!hasPulled)return;
    clearTimeout(pushTimer.current);
    pushTimer.current=setTimeout(()=>doPush(),AUTO_PUSH_MS);
    return()=>clearTimeout(pushTimer.current);
  },[appointments,callings,releasings,hasPulled,doPush]);

  // Stale-data alert: warn if data hasn't been pulled in > 5 minutes
  useEffect(()=>{
    const id=setInterval(()=>{
      if(!lastSyncedAt) return;
      const mins=(Date.now()-lastSyncedAt)/60000;
      if(mins>5 && document.visibilityState==="visible"){
        notify.warn("Data may be stale — consider pulling latest changes");
      }
    },5*60*1000);
    return()=>clearInterval(id);
  },[lastSyncedAt]);

  // ── Auto-create appointments when a calling/releasing reaches Approved stage ──
  useEffect(()=>{
    if(!hasPulled) return;

    const hasAppt=(name,purpose,calling)=>appointments.some(a=>
      a.name.toLowerCase()===name.toLowerCase() &&
      a.purpose===purpose &&
      a.notes===calling &&
      a.status!=="Completed"
    );

    const toCreate=[];

    callings.forEach(c=>{
      if(c.stage!=="Approved to Call"||!c.name) return;
      if(!hasAppt(c.name,"Calling",c.calling))
        toCreate.push({id:`a_${Date.now()}_${Math.random().toString(36).slice(2,7)}`,
          name:c.name,status:"Need to Schedule",owner:"Bishop",
          purpose:"Calling",apptDate:"",notes:c.calling});
    });

    releasings.forEach(r=>{
      if(r.stage!=="Approved to Release"||!r.name) return;
      if(!hasAppt(r.name,"Releasing",r.calling))
        toCreate.push({id:`a_${Date.now()}_${Math.random().toString(36).slice(2,7)}`,
          name:r.name,status:"Need to Schedule",owner:"Bishop",
          purpose:"Releasing",apptDate:"",notes:r.calling});
    });

    if(toCreate.length===0) return;

    // Fire toasts outside of setAppts to avoid React Strict Mode double-invoke
    toCreate.forEach(nc=>notify.info(`Appointment created for ${nc.name} (${nc.purpose})`));
    setAppts(prev=>[...prev,...toCreate]);

  },[callings, releasings, appointments, hasPulled]);

  // ── Role: admin = in ALLOWED_EMAILS, limited = in WARD_COUNCIL_EMAILS ──
  // userEmail already declared above for isAdminRef
  const isAdmin = isAdminEarly;
  const isWardCouncil = !isAdmin && (config.WARD_COUNCIL_EMAILS||[]).map(e=>e.toLowerCase()).includes(userEmail);

  // Default tab for WC-only users is ward-council, not appointments
  useEffect(() => {
    if (isWardCouncil) setTab("ward-council");
  }, [isWardCouncil]); // eslint-disable-line

  const isBusy=syncStatus==="pulling"||syncStatus==="pushing";
  const isErr=syncStatus==="error";
  const syncLabel=(()=>{
    if(syncStatus==="pulling")return<><ArrowDown size={10}/> pulling…</>;
    if(syncStatus==="pushing")return<><ArrowUp size={10}/> saving…</>;
    if(isErr)return<><AlertTriangle size={10}/> {syncError||"error"}</>;
    if(!lastSyncedAt)return null;
    const s=Math.round((Date.now()-lastSyncedAt)/1000);
    if(s<10)return"synced just now";if(s<60)return`synced ${s}s ago`;
    return`synced ${Math.round(s/60)}m ago`;
  })();

  const NAV_GROUPS = isAdmin ? [
    { id:"scheduling", label:"Scheduling", children:[
      {id:"appointments", label:"Appointments", badge:appointments.filter(a=>a.status!=="Completed").length||null},
      {id:"callings",     label:"Callings",     badge:callings.filter(c=>c.stage!=="Completed").length||null},
      {id:"releasings",   label:"Releasings",   badge:releasings.filter(r=>r.stage!=="Completed").length||null},
    ]},
    { id:"meetings", label:"Meetings", children:[
      {id:"bishopric",    label:"Bishopric Council", badge:null},
      {id:"sacrament",    label:"Sacrament",         badge:null},
      {id:"ward-council", label:"Ward Council",       badge:null},
    ]},
    { id:"calendar", label:"Calendar", flat:true },
    { id:"links",    label:"Links",    flat:true },
    { id:"admin", label:"Admin", children:[
      {id:"alerts",  label:"Alerts",  badge:null},
      {id:"members", label:"Members", badge:null},
    ]},
  ] : [
    { id:"ward-council", label:"Ward Council", flat:true },
    { id:"calendar",     label:"Calendar",     flat:true },
    { id:"links",        label:"Links",        flat:true },
  ];
  // Which group contains the active tab
  const activeGroup = NAV_GROUPS.find(g => g.flat ? g.id === tab : g.children?.some(c => c.id === tab))?.id || null;

  const connColors = {ok:"#50A83E",error:"#E10F5A",warn:"#F68D2E",testing:"#007DA5",unknown:C.textLight};

  return (
    <div style={{fontFamily:"Georgia,serif",minHeight:"100vh",background:C.pageBg,color:C.textPrimary}}>
      <style>{CSS}</style>
      <ToastStack/>

      {/* ── Header ── */}
      <header style={{borderBottom:`1.5px solid ${C.border}`,padding:"0 28px",display:"flex",alignItems:"center",justifyContent:"space-between",height:58,position:"sticky",top:0,zIndex:50,background:C.surfaceWhite,boxShadow:"0 1px 0 rgba(0,0,0,.04)"}}>
        <div style={{display:"flex",alignItems:"center",gap:24}}>
          <div style={{display:"flex",alignItems:"baseline",gap:6}}>
            <span style={{fontSize:18,fontWeight:600,color:C.blue35}}>Ward</span>
            <span style={{fontSize:18,fontWeight:400,fontStyle:"italic",color:C.blue25}}>Manager</span>
            <span style={{fontSize:10,fontFamily:"'Helvetica Neue',Arial,sans-serif",fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",color:C.textLight,marginLeft:4}}>{config.WARD_NAME}</span>
          </div>
          <nav style={{display:"flex",gap:2}}>
            {NAV_GROUPS.map(group => (
              group.flat
                ? <button key={group.id} className="tab-btn" onClick={() => setTab(group.id)}
                    style={{
                      padding:"5px 14px", borderRadius:6, fontSize:13,
                      fontFamily:"'Helvetica Neue',Arial,sans-serif",
                      color: tab === group.id ? C.blue35 : C.textMuted,
                      background: tab === group.id ? "rgba(0,85,129,.07)" : "transparent",
                      borderBottom: `2.5px solid ${tab === group.id ? C.blue35 : "transparent"}`,
                      fontWeight: tab === group.id ? 700 : 400,
                    }}>
                    {group.label}
                  </button>
                : <NavGroup key={group.id}
                    group={group}
                    activeTab={tab}
                    isActiveGroup={activeGroup===group.id}
                    onSelect={id=>setTab(id)}
                  />
            ))}
          </nav>
        </div>

        <div style={{display:"flex",alignItems:"center",gap:8}}>
          {/* Connection status */}
          <div title={connMsg||"Click to test connection"} onClick={doTestConnection} style={{display:"flex",alignItems:"center",gap:5,padding:"4px 10px",borderRadius:20,background:C.surfaceWarm,border:`1px solid ${C.borderLight}`,cursor:"pointer",transition:"border .2s"}}>
            <span style={{width:7,height:7,borderRadius:"50%",background:connColors[connStatus]||C.textLight,animation:connStatus==="testing"?"pulse 1s infinite":"none"}}/>
            <span style={{fontSize:10,fontFamily:"'Helvetica Neue',Arial,sans-serif",color:C.textMuted}}>
              {connStatus==="unknown"?"Test"  :connStatus==="testing"?"Testing…":connStatus==="ok"?"Connected":connStatus==="warn"?"Check config":"Error"}
            </span>
          </div>

          {/* Sync pill */}
          <div style={{display:"flex",alignItems:"center",gap:6,padding:"4px 10px",borderRadius:20,background:C.surfaceWarm,border:`1px solid ${isErr?C.red10:C.borderLight}`}}>
            <span style={{width:7,height:7,borderRadius:"50%",flexShrink:0,background:isBusy?C.blue25:isErr?C.red15:C.green25,boxShadow:isBusy?"0 0 0 3px rgba(0,125,165,.2)":"none",animation:isBusy?"pulse 1s infinite":"none"}}/>
            {syncLabel&&<span style={{fontSize:10,fontFamily:"'Helvetica Neue',Arial,sans-serif",color:isErr?C.red15:C.textMuted,whiteSpace:"nowrap"}}>{syncLabel}</span>}
          </div>

          <button className="btn-secondary" onClick={pull} disabled={isBusy} style={{opacity:isBusy?.5:1,fontSize:11,padding:"6px 14px",display:"flex",alignItems:"center",gap:4}}><PullIcon/> Pull</button>
          {isAdmin&&<button className="btn-primary" onClick={push} disabled={isBusy} style={{opacity:isBusy?.5:1,fontSize:11,padding:"6px 14px",display:"flex",alignItems:"center",gap:4}}><PushIcon/> Push</button>}

          {/* Settings — admin only */}
          {isAdmin&&<button onClick={()=>setShowSettings(true)} title="Settings" style={{background:"none",border:`1.5px solid ${C.border}`,borderRadius:8,width:34,height:34,cursor:"pointer",display:"flex",alignItems:"center",justifyContent:"center",color:C.textMuted,fontSize:16}}><Settings2 size={15}/></button>}

          {/* User chip */}
          <div style={{display:"flex",alignItems:"center",gap:8,padding:"4px 12px",background:C.surfaceWarm,border:`1px solid ${C.border}`,borderRadius:20}}>
            {user.picture?<img src={user.picture} alt="" style={{width:22,height:22,borderRadius:"50%"}}/>:<User size={16} color={C.textMuted}/>}
            <span style={{fontSize:12,fontFamily:"'Helvetica Neue',Arial,sans-serif",color:C.textMuted}}>{user.name?.split(" ")[0]}</span>
            <button onClick={onSignOut} title="Sign out" style={{background:"none",border:"none",color:C.textLight,cursor:"pointer",display:"flex",alignItems:"center"}}><LogOut size={13}/></button>
          </div>
        </div>
      </header>



      <main style={{padding:"24px 28px",maxWidth:1440,margin:"0 auto"}}>
        {isAdmin&&tab==="appointments"&&<AppointmentsTab data={appointments} setData={setAppts} roster={roster}/>}
        {isAdmin&&tab==="callings"    &&<PipelineTab title="Callings"   stages={CALLING_STAGES}   data={callings}   setData={setCallings}/>}
        {isAdmin&&tab==="releasings"  &&<PipelineTab title="Releasings" stages={RELEASING_STAGES} data={releasings} setData={setReleasings}/>}
        {isAdmin&&tab==="bishopric"   &&<BishopricCouncilTab bishopricMeeting={bishopricMeeting} setBishopricMeeting={setBishopricMeeting} callings={callings} releasings={releasings} sacramentProgram={sacramentProgram} calendar={calendar} roster={roster} token={token} onNavigate={setTab}/>}
        {isAdmin&&tab==="members"     &&<MembersTab  data={members}  setData={setMembers} callings={callings} releasings={releasings}/>}
        {isAdmin&&tab==="alerts"      &&<AlertsTab appointments={appointments} callings={callings} releasings={releasings}/>}
        {isAdmin&&tab==="sacrament"&&<SacramentTab data={sacramentProgram} setData={setSacrament} saveFn={doSacramentSave} pullFn={doSacramentPull} tokenRef={tokenRef}/>}
        {(isAdmin||isWardCouncil)&&tab==="calendar"&&<CalendarTab calendar={calendar} setCalendar={setCalendar} token={token}/>}
        {(isAdmin||isWardCouncil)&&tab==="ward-council"&&<WardCouncilTab wardCouncilMeeting={wardCouncilMeeting} setWardCouncilMeeting={setWardCouncilMeeting} calendar={calendar} roster={roster} token={token} onNavigate={setTab} isAdmin={isAdmin}/>}
        {(isAdmin||isWardCouncil)&&tab==="links"&&<LinksTab bishopricLinks={isAdmin?bishopricLinks:null} setBishopricLinks={setBishopricLinks} wcLinks={wcLinks} setWcLinks={setWcLinks} token={token} isAdmin={isAdmin}/>}
        {!isAdmin&&!isWardCouncil&&<UnauthorizedView/>}
      </main>

      {showSettings&&<SettingsModal token={token} onTestConn={doTestConnection} connStatus={connStatus} connMsg={connMsg} roster={roster} setRoster={setRoster} onClose={()=>setShowSettings(false)}/>}
    </div>
  );
}

// ─── Appointments Tab ─────────────────────────────────────────────────────────
function AppointmentsTab({data,setData,roster=[]}){
  const[stageFilters,setSF]=useState(APPT_STAGES.filter(s=>s!=="Completed"));
  const[ownerFilter,setOF]=useState("All");
  const[view,setView]=useState("table");
  const[showForm,setSF2]=useState(false);
  const[editing,setEditing]=useState(null);

  const toggle=s=>setSF(p=>p.includes(s)?p.filter(x=>x!==s):[...p,s]);
  const showingAll=stageFilters.length===0||stageFilters.length===APPT_STAGES.length;
  const filtered=data.filter(a=>(showingAll||stageFilters.includes(a.status))&&(ownerFilter==="All"||a.owner===ownerFilter));
  const save=item=>{
    if(item.id){setData(d=>d.map(x=>x.id===item.id?item:x));notify.success("Appointment updated");}
    else{setData(d=>[...d,{...item,id:`a_${Date.now()}`}]);notify.success("Appointment created");}
    setSF2(false);setEditing(null);
  };
  const upd=(id,f,v)=>setData(d=>d.map(x=>x.id===id?{...x,[f]:v}:x));
  const del=id=>{setData(d=>d.filter(x=>x.id!==id));notify.info("Appointment removed");};
  const dragStatus=(id,status)=>{setData(d=>d.map(x=>x.id===id?{...x,status}:x));notify.success("Status updated");};
  const open=item=>{setEditing(item);setSF2(true);};
  const counts=APPT_STAGES.reduce((a,s)=>({...a,[s]:data.filter(x=>x.status===s).length}),{});
  const active=data.filter(a=>a.status!=="Completed").length;

  return(
    <div className="animate-in">
      <HeroBanner title="Appointments" sub={`${active} active · ${counts.Completed||0} completed`}>
        <button className="btn-secondary" style={{background:"rgba(255,255,255,.14)",border:"1.5px solid rgba(255,255,255,.3)",color:"#fff"}} onClick={()=>setView(v=>v==="table"?"kanban":"table")}>
          {view==="table"?<><KanbanIcon/> Kanban</>:<><TableIcon/> Table</>}
        </button>
        <button className="btn-primary" style={{background:"rgba(255,255,255,.18)",border:"1.5px solid rgba(255,255,255,.4)"}} onClick={()=>{setEditing(null);setSF2(true);}}>
          <PlusIcon/> New Appointment
        </button>
      </HeroBanner>

      <div style={{display:"flex",gap:10,marginBottom:16,flexWrap:"wrap"}}>
        {APPT_STAGES.map(s=>{
          const a=stageFilters.includes(s);const st=APPT_STAGE_STYLE[s];
          return<div key={s} onClick={()=>toggle(s)} style={{borderRadius:12,padding:"14px 18px",cursor:"pointer",flex:1,minWidth:140,border:`1.5px solid ${a?st.border:C.border}`,background:a?st.bg:C.surfaceWhite,opacity:!showingAll&&!a?.42:1,boxShadow:a?"0 2px 8px rgba(0,0,0,.06)":"none",transition:"all .18s",position:"relative",overflow:"hidden"}}>
            {a&&<div style={{position:"absolute",top:0,right:0,width:"55%",height:"100%",background:"linear-gradient(270deg,rgba(255,255,255,.2) 0%,transparent 100%)",pointerEvents:"none"}}/>}
            <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:8}}>
              <span style={{width:8,height:8,borderRadius:"50%",background:st.dot,flexShrink:0}}/>
              <span style={{fontSize:10,fontWeight:700,letterSpacing:".09em",textTransform:"uppercase",color:a?st.text:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>{s}</span>
            </div>
            <div style={{fontFamily:"Georgia,serif",fontSize:28,fontWeight:400,color:a?st.text:C.textPrimary,lineHeight:1}}>{counts[s]||0}</div>
          </div>;
        })}
      </div>

      <div style={{display:"flex",alignItems:"center",gap:8,marginBottom:16,flexWrap:"wrap"}}>
        <span style={{fontSize:10,fontWeight:700,letterSpacing:".1em",textTransform:"uppercase",color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>Owner</span>
        {["All",...LEADERS].map(o=>{const os=OWNER_STYLE[o];const act=ownerFilter===o;return(
          <button key={o} onClick={()=>setOF(o)} style={{borderRadius:7,fontSize:12,padding:"5px 13px",cursor:"pointer",transition:"all .15s",border:`1.5px solid ${act?(os?os.border:C.blue35):C.border}`,color:act?(os?os.text:C.blue35):C.textSecond,background:act?(os?os.bg:C.surfaceWarm):"transparent",fontWeight:act?700:400,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>{o}</button>
        );})}
      </div>

      {view==="table"
        ?<div style={{background:C.surfaceWhite,border:`1.5px solid ${C.border}`,borderRadius:12,overflow:"hidden"}}><ApptTable data={filtered} onUpd={upd} onDel={del} roster={roster}/></div>
        :<ApptKanban data={filtered} onEdit={open} onStatusChange={dragStatus} roster={roster}/>
      }
      {showForm&&<AppointmentModal item={editing} onSave={save} onClose={()=>{setSF2(false);setEditing(null);}} roster={roster}/>}
    </div>
  );
}

function PurposeCell({value,onChange}){
  return<div style={{display:"flex",alignItems:"center",gap:6}}>
    {value&&<span style={{display:"flex",alignItems:"center",color:C.textMuted}}><PurposeIcon purpose={value} size={14}/></span>}
    <select value={value||""} onChange={e=>onChange(e.target.value)} style={{background:"transparent",border:"none",outline:"none",cursor:"pointer",fontSize:13,fontFamily:"Georgia,serif",color:value?C.textPrimary:C.textLight,padding:0,maxWidth:170}}>
      <option value="">— Select —</option>{PURPOSE_OPTIONS.map(p=><option key={p} value={p}>{p}</option>)}
    </select>
  </div>;
}

function ApptTable({data,onUpd,onDel,roster=[]}){
  return<div style={{overflowX:"auto"}}>
    <table className="wm-table">
      <thead><tr>{["Member","Status","Owner","Purpose","Date","Notes",""].map(h=><th key={h}>{h}</th>)}</tr></thead>
      <tbody>
        {!data.length&&<tr><td colSpan={7}><div style={{textAlign:"center",padding:"60px 24px"}}><div style={{marginBottom:12,opacity:.3,color:C.textMuted,display:"flex",justifyContent:"center"}}><Calendar size={38}/></div><div style={{fontFamily:"Georgia,serif",fontSize:17,color:C.textSecond,marginBottom:5}}>No appointments</div><div style={{fontSize:12,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>Add one or adjust filters</div></div></td></tr>}
        {data.map(a=>(
          <tr key={a.id} className={a.status==="Completed"?"row-done":""}>
            <td style={{minWidth:150}}><InlineText serif value={a.name} onChange={v=>onUpd(a.id,"name",v)}/></td>
            <td style={{minWidth:158}}><StageChip status={a.status} onChange={v=>onUpd(a.id,"status",v)}/></td>
            <td style={{minWidth:152}}><OwnerChip owner={a.owner} onChange={v=>onUpd(a.id,"owner",v)} roster={roster}/></td>
            <td style={{minWidth:185}}><PurposeCell value={a.purpose} onChange={v=>onUpd(a.id,"purpose",v)}/></td>
            <td style={{minWidth:126}}><input type="date" value={a.apptDate} onChange={e=>onUpd(a.id,"apptDate",e.target.value)} style={{background:"transparent",border:"none",outline:"none",cursor:"pointer",fontSize:12,fontFamily:"'Helvetica Neue',Arial,sans-serif",color:a.apptDate?C.textSecond:C.textLight,padding:0,width:"auto"}}/></td>
            <td style={{minWidth:165}}><InlineText serif value={a.notes} onChange={v=>onUpd(a.id,"notes",v)}/></td>
            <td style={{width:44,textAlign:"center"}}><button className="btn-del" onClick={()=>onDel(a.id)}><XIcon/></button></td>
          </tr>
        ))}
      </tbody>
    </table>
  </div>;
}

function ApptKanban({data,onEdit,onStatusChange,roster=[]}){
  const[dragging,setDragging]=useState(null);
  const[dragOver,setDragOver]=useState(null); // {col, index}

  const onDragStart=(e,item)=>{
    setDragging(item);
    e.dataTransfer.effectAllowed="move";
    e.dataTransfer.setData("text/plain",item.id);
  };
  const onDragEnd=()=>{setDragging(null);setDragOver(null);};
  const onDragOverCol=(e,col)=>{e.preventDefault();e.dataTransfer.dropEffect="move";setDragOver({col});};
  const onDrop=(e,targetStage)=>{
    e.preventDefault();
    if(dragging&&dragging.status!==targetStage) onStatusChange(dragging.id,targetStage);
    setDragging(null);setDragOver(null);
  };

  return<div style={{display:"flex",gap:12,overflowX:"auto",paddingBottom:8}}>
    {APPT_STAGES.map(s=>{
      const items=data.filter(a=>a.status===s);
      const st=APPT_STAGE_STYLE[s];
      const isOver=dragOver?.col===s;
      const isDragSource=dragging?.status===s;
      return(
        <div key={s} className="kanban-col"
          onDragOver={e=>onDragOverCol(e,s)}
          onDragLeave={()=>setDragOver(null)}
          onDrop={e=>onDrop(e,s)}
          style={{transition:"background .15s, box-shadow .15s",
            background:isOver?(isDragSource?undefined:`color-mix(in srgb, ${st.bg} 60%, white)`):undefined,
            boxShadow:isOver&&!isDragSource?`inset 0 0 0 2px ${st.border}`:undefined,
            outline:isOver&&!isDragSource?`2px dashed ${st.border}`:undefined,
            outlineOffset:-2}}>
          <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:13,paddingBottom:10,borderBottom:`2px solid ${st.border}`}}>
            <span style={{width:8,height:8,borderRadius:"50%",background:st.dot,flexShrink:0}}/>
            <span style={{fontSize:10,fontWeight:700,letterSpacing:".1em",textTransform:"uppercase",color:st.text,flex:1,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>{s}</span>
            <span style={{fontFamily:"Georgia,serif",fontSize:15,color:st.text,background:st.bg,border:`1px solid ${st.border}`,borderRadius:20,padding:"1px 8px"}}>{items.length}</span>
          </div>
          {items.map(a=>(
            <div key={a.id} className="kanban-card"
              draggable
              onDragStart={e=>onDragStart(e,a)}
              onDragEnd={onDragEnd}
              onClick={()=>!dragging&&onEdit(a)}
              style={{opacity:dragging?.id===a.id?.35:1,cursor:"grab",transition:"opacity .15s, transform .15s",
                transform:dragging?.id===a.id?"scale(.97)":undefined}}>
              <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:6,marginBottom:5}}>
                <div style={{fontFamily:"Georgia,serif",fontSize:15,color:C.textPrimary}}>{a.name||<span style={{color:C.textLight,fontStyle:"italic"}}>Unnamed</span>}</div>
                <span style={{fontSize:12,color:C.textLight,cursor:"grab",flexShrink:0,marginTop:1}} title="Drag to move">⠿</span>
              </div>
              {a.purpose&&<div style={{display:"flex",alignItems:"center",gap:5,marginBottom:8}}><PurposeIcon purpose={a.purpose} size={13} color={C.textMuted}/><span style={{fontSize:11,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>{a.purpose}</span></div>}
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginTop:8,paddingTop:8,borderTop:`1px solid ${C.borderLight}`}}>
                <OwnerChip owner={a.owner} onChange={null} compact roster={roster}/>
                {a.apptDate&&<span style={{fontSize:10,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>{a.apptDate}</span>}
              </div>
            </div>
          ))}
          {!items.length&&(
            <div style={{textAlign:"center",padding:"22px 12px",color:isOver&&!isDragSource?st.text:C.textLight,
              fontFamily:"Georgia,serif",fontStyle:"italic",fontSize:13,borderRadius:8,transition:"all .15s",
              border:isOver&&!isDragSource?`1.5px dashed ${st.border}`:"1.5px dashed transparent",
              background:isOver&&!isDragSource?st.bg:undefined}}>
              {isOver&&!isDragSource?"Drop here":"Empty"}
            </div>
          )}
        </div>
      );
    })}
  </div>;
}

function AppointmentModal({item,onSave,onClose,roster=[]}){
  const[f,setF]=useState(item||{name:"",status:"Need to Schedule",owner:"Bishop",purpose:"",apptDate:"",notes:""});
  const s=(k,v)=>setF(x=>({...x,[k]:v}));const st=APPT_STAGE_STYLE[f.status]||{};
  return<ModalShell onClose={onClose} title={f.name||(item?"Appointment":"New Appointment")} subtitle={`${item?"Edit":"New"} Appointment`}>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
      <FormRow label="Member Name" full><input value={f.name} onChange={e=>s("name",e.target.value)} placeholder="Full name"/></FormRow>
      <FormRow label="Status"><select value={f.status} onChange={e=>s("status",e.target.value)}>{APPT_STAGES.map(x=><option key={x}>{x}</option>)}</select></FormRow>
      <FormRow label="Owner"><select value={f.owner} onChange={e=>s("owner",e.target.value)}>{LEADERS.map(x=><option key={x} value={x}>{rosterName(roster,x)}</option>)}</select></FormRow>
      <FormRow label="Purpose"><select value={f.purpose} onChange={e=>s("purpose",e.target.value)}><option value="">— Select —</option>{PURPOSE_OPTIONS.map(x=><option key={x}>{x}</option>)}</select></FormRow>
      <FormRow label="Date"><input type="date" value={f.apptDate} onChange={e=>s("apptDate",e.target.value)}/></FormRow>
      <FormRow label="Notes" full><textarea rows={3} value={f.notes} onChange={e=>s("notes",e.target.value)} style={{resize:"vertical"}}/></FormRow>
    </div>
    <ModalFooter onClose={onClose} onSave={()=>onSave(f)} saveLabel="Save Appointment"/>
  </ModalShell>;
}

// ─── Pipeline Tab (Callings & Releasings) ────────────────────────────────────
function PipelineTab({title,stages,data,setData}){
  const[view,setView]=useState("kanban");const[search,setSrch]=useState("");
  const[showForm,setSF]=useState(false);const[editing,setEditing]=useState(null);
  const filtered=data.filter(c=>c.name.toLowerCase().includes(search.toLowerCase())||c.calling.toLowerCase().includes(search.toLowerCase()));
  const save=item=>{
    if(item.id){setData(d=>d.map(x=>x.id===item.id?item:x));notify.success(`${title.slice(0,-1)} updated`);}
    else{setData(d=>[...d,{...item,id:`p_${Date.now()}`}]);notify.success(`${title.slice(0,-1)} added`);}
    setSF(false);setEditing(null);
  };
  const del=id=>{setData(d=>d.filter(x=>x.id!==id));notify.info("Record removed");};
  const open=item=>{setEditing(item);setSF(true);};
  const setStage=(id,stage)=>setData(d=>d.map(x=>x.id===id?{...x,stage}:x));
  const dragStage=(id,stage)=>{setData(d=>d.map(x=>x.id===id?{...x,stage}:x));notify.success("Stage updated");};
  const move=(id,dir)=>setData(d=>d.map(x=>{if(x.id!==id)return x;const i=stages.indexOf(x.stage);return stages[i+dir]?{...x,stage:stages[i+dir]}:x;}));
  const counts=stages.reduce((a,s)=>({...a,[s]:data.filter(x=>x.stage===s).length}),{});
  const active=data.filter(x=>x.stage!=="Completed").length;

  return(
    <div className="animate-in">
      <HeroBanner title={title} sub={`${active} active · ${counts.Completed||0} completed`}>
        <button className="btn-secondary" style={{background:"rgba(255,255,255,.14)",border:"1.5px solid rgba(255,255,255,.3)",color:"#fff"}} onClick={()=>setView(v=>v==="kanban"?"table":"kanban")}>
          {view==="kanban"?<><TableIcon/> Table</>:<><KanbanIcon/> Kanban</>}
        </button>
        <button className="btn-primary" style={{background:"rgba(255,255,255,.18)",border:"1.5px solid rgba(255,255,255,.4)"}} onClick={()=>{setEditing(null);setSF(true);}}>
          <PlusIcon/> New {title.slice(0,-1)}
        </button>
      </HeroBanner>

      <div style={{display:"flex",gap:3,height:6,borderRadius:4,overflow:"hidden",marginBottom:12}}>
        {stages.map(s=>{const st=PIPELINE_STAGE_STYLE[s]||{};return<div key={s} style={{flex:Math.max(counts[s]||0,.3),background:st.border||C.border,transition:"flex .5s",minWidth:4}}/>;})}</div>

      <div style={{display:"flex",gap:12,marginBottom:16,flexWrap:"wrap"}}>
        {stages.map(s=>{const st=PIPELINE_STAGE_STYLE[s]||{};return(
          <div key={s} style={{display:"flex",alignItems:"center",gap:5}}>
            <span style={{width:7,height:7,borderRadius:"50%",background:st.dot||C.textMuted,flexShrink:0}}/>
            <span style={{fontSize:11,fontFamily:"'Helvetica Neue',Arial,sans-serif",color:C.textSecond}}>{s}</span>
            <span style={{fontSize:11,fontFamily:"'Helvetica Neue',Arial,sans-serif",color:C.textMuted}}>({counts[s]||0})</span>
          </div>
        );})}
      </div>

      <div style={{marginBottom:16}}>
        <input value={search} onChange={e=>setSrch(e.target.value)} placeholder="Search name or calling…" style={{maxWidth:300}}/>
      </div>

      {view==="kanban"?<PipelineKanban data={filtered} stages={stages} onEdit={open} onMove={move} onStageChange={dragStage}/>
        :<PipelineTable data={filtered} stages={stages} onEdit={open} onDel={del} onStage={setStage}/>}

      {showForm&&<PipelineModal item={editing} stages={stages} title={title.slice(0,-1)} onSave={save} onClose={()=>{setSF(false);setEditing(null);}}/>}
    </div>
  );
}

function PipelineKanban({data,stages,onEdit,onMove,onStageChange}){
  const[dragging,setDragging]=useState(null);
  const[dragOver,setDragOver]=useState(null);
  const onDragStart=(e,item)=>{setDragging(item);e.dataTransfer.effectAllowed="move";e.dataTransfer.setData("text/plain",item.id);};
  const onDragEnd=()=>{setDragging(null);setDragOver(null);};
  const onDragOverCol=(e,col)=>{e.preventDefault();e.dataTransfer.dropEffect="move";setDragOver({col});};
  const onDrop=(e,targetStage)=>{e.preventDefault();if(dragging&&dragging.stage!==targetStage)onStageChange(dragging.id,targetStage);setDragging(null);setDragOver(null);};
  return<div style={{display:"flex",gap:10,overflowX:"auto",paddingBottom:8}}>
    {stages.map(s=>{
      const items=data.filter(c=>c.stage===s);const st=PIPELINE_STAGE_STYLE[s]||{};
      const isOver=dragOver?.col===s;const isDragSource=dragging?.stage===s;
      return(
        <div key={s} className="kanban-col" style={{minWidth:200,transition:"background .15s",
          background:isOver&&!isDragSource?`color-mix(in srgb, ${st.bg||C.surfaceWarm} 70%, white)`:undefined,
          outline:isOver&&!isDragSource?`2px dashed ${st.border||C.border}`:undefined,outlineOffset:-2}}
          onDragOver={e=>onDragOverCol(e,s)} onDragLeave={()=>setDragOver(null)} onDrop={e=>onDrop(e,s)}>
          <div style={{display:"flex",alignItems:"center",gap:7,marginBottom:13,paddingBottom:10,borderBottom:`2px solid ${st.border||C.border}`}}>
            <span style={{width:8,height:8,borderRadius:"50%",background:st.dot||C.textMuted,flexShrink:0}}/>
            <span style={{fontSize:10,fontWeight:700,letterSpacing:".09em",textTransform:"uppercase",color:st.text||C.textMuted,flex:1,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>{s}</span>
            <span style={{fontFamily:"Georgia,serif",fontSize:14,color:st.text||C.textMuted,background:st.bg||C.surfaceWarm,border:`1px solid ${st.border||C.border}`,borderRadius:20,padding:"1px 8px"}}>{items.length}</span>
          </div>
          {items.map(c=>(
            <div key={c.id} className="kanban-card" draggable
              onDragStart={e=>onDragStart(e,c)} onDragEnd={onDragEnd}
              onClick={()=>!dragging&&onEdit(c)}
              style={{opacity:dragging?.id===c.id?.35:1,cursor:"grab",transition:"opacity .15s, transform .15s",
                transform:dragging?.id===c.id?"scale(.97)":undefined}}>
              <div style={{display:"flex",alignItems:"flex-start",justifyContent:"space-between",gap:6,marginBottom:4}}>
                <div style={{fontFamily:"Georgia,serif",fontSize:15,color:C.textPrimary}}>{c.name||<span style={{color:C.textLight,fontStyle:"italic"}}>No name</span>}</div>
                <span style={{fontSize:12,color:C.textLight,cursor:"grab",flexShrink:0}} title="Drag to move">⠿</span>
              </div>
              <div style={{fontSize:12,color:C.blue25,fontStyle:"italic",fontFamily:"Georgia,serif",marginBottom:c.notes?8:4}}>{c.calling}</div>
              {c.notes&&<div style={{fontSize:11,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",paddingTop:6,borderTop:`1px solid ${C.borderLight}`,lineHeight:1.5,marginBottom:6}}>{c.notes}</div>}
              <div style={{display:"flex",gap:4,justifyContent:"flex-end",marginTop:4}}>
                <button className="btn-secondary" onClick={e=>{e.stopPropagation();onMove(c.id,-1);}} style={{padding:"3px 9px",display:"flex",alignItems:"center"}} title="Move left"><ChevLeftIcon/></button>
                <button className="btn-secondary" onClick={e=>{e.stopPropagation();onMove(c.id, 1);}} style={{padding:"3px 9px",display:"flex",alignItems:"center"}} title="Move right"><ChevRightIcon/></button>
              </div>
            </div>
          ))}
          {!items.length&&(
            <div style={{textAlign:"center",padding:"22px 12px",color:isOver&&!isDragSource?st.text||C.textMuted:C.textLight,
              fontFamily:"Georgia,serif",fontStyle:"italic",fontSize:13,borderRadius:8,transition:"all .15s",
              border:isOver&&!isDragSource?`1.5px dashed ${st.border||C.border}`:"1.5px dashed transparent",
              background:isOver&&!isDragSource?st.bg:undefined}}>
              {isOver&&!isDragSource?"Drop here":"Empty"}
            </div>
          )}
        </div>
      );
    })}
  </div>;
}

function PipelineTable({data,stages,onEdit,onDel,onStage}){
  return<div style={{background:C.surfaceWhite,border:`1.5px solid ${C.border}`,borderRadius:12,overflow:"hidden"}}>
    <div style={{overflowX:"auto"}}>
      <table className="wm-table">
        <thead><tr>{["Calling / Position","Member","Stage","Notes",""].map(h=><th key={h}>{h}</th>)}</tr></thead>
        <tbody>
          {!data.length&&<tr><td colSpan={5}><div style={{textAlign:"center",padding:"60px 24px"}}><div style={{marginBottom:12,opacity:.3,color:C.textMuted}}><ClipboardIcon/></div><div style={{fontFamily:"Georgia,serif",fontSize:17,color:C.textSecond,marginBottom:5}}>Nothing here yet</div><div style={{fontSize:12,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>Add an entry or pull from Sheets</div></div></td></tr>}
          {data.map(c=>{const st=PIPELINE_STAGE_STYLE[c.stage]||{};return(
            <tr key={c.id}>
              <td style={{minWidth:180,fontStyle:"italic",color:C.blue35,fontFamily:"Georgia,serif"}}>{c.calling}</td>
              <td style={{minWidth:160,fontFamily:"Georgia,serif"}}>{c.name}</td>
              <td style={{minWidth:200}}><span className="chip" style={{background:st.bg,borderColor:st.border,color:st.text}}><span style={{width:6,height:6,borderRadius:"50%",background:st.dot,flexShrink:0}}/><select value={c.stage} onChange={e=>onStage(c.id,e.target.value)} className="chip-select">{stages.map(s=><option key={s} value={s}>{s}</option>)}</select></span></td>
              <td style={{minWidth:200,fontSize:13,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",maxWidth:260,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.notes||"—"}</td>
              <td style={{width:88}}><div style={{display:"flex",gap:5}}><button className="btn-secondary" onClick={()=>onEdit(c)} style={{padding:"3px 10px",fontSize:11}}>Edit</button><button className="btn-del" onClick={()=>onDel(c.id)}><XIcon/></button></div></td>
            </tr>
          );})}
        </tbody>
      </table>
    </div>
  </div>;
}

function PipelineModal({item,stages,title,onSave,onClose}){
  const[f,setF]=useState(item||{calling:"",name:"",stage:stages[0],notes:""});
  const s=(k,v)=>setF(x=>({...x,[k]:v}));
  return<ModalShell onClose={onClose} title={f.name||f.calling||`New ${title}`} subtitle={`${item?"Edit":"New"} ${title}`}>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
      <FormRow label="Calling / Position" full><input value={f.calling} onChange={e=>s("calling",e.target.value)} placeholder="e.g. Primary President"/></FormRow>
      <FormRow label="Member Name" full><input value={f.name} onChange={e=>s("name",e.target.value)} placeholder="Full name"/></FormRow>
      <FormRow label="Stage" full><select value={f.stage} onChange={e=>s("stage",e.target.value)}>{stages.map(x=><option key={x}>{x}</option>)}</select></FormRow>
      <FormRow label="Notes" full><textarea rows={3} value={f.notes} onChange={e=>s("notes",e.target.value)} style={{resize:"vertical"}}/></FormRow>
    </div>
    <ModalFooter onClose={onClose} onSave={()=>onSave(f)} saveLabel={`Save ${title}`}/>
  </ModalShell>;
}


// ─── useMeetingSync — shared auto-push + pending-changes detection ────────────
// Provides:
//   markDirty()        — call after any local edit; triggers debounced auto-save
//   syncStatus         — "idle" | "pushing" | "error"
//   saveStatus         — "saved" | "saving" | "unsaved" | "error"
//   pendingCount       — number of remote rows that differ from local snapshot
//   applyPending()     — merge remote snapshot into local state
//   doSave()           — immediate manual save
//
// Parameters:
//   getData()          — returns current data array to save
//   saveFn(token,data) — async function that writes to Sheets
//   pullFn(token)      — async function that returns { data: [] } from Sheets
//   diffFn(a,b)        — returns count of meaningful differences between two data arrays
//   token              — Google OAuth token ref
//   onApply(data)      — called with fresh remote data when user applies pending changes
//   enabled            — false to skip all sync (e.g. ward council user on BM tab)

function useMeetingSync({ getData, saveFn, pullFn, diffFn, tokenRef, onApply, enabled = true }) {
  const [saveStatus,   setSaveStatus]   = useState("saved");   // saved|saving|unsaved|error
  const [pendingCount, setPendingCount] = useState(0);
  const [pendingData,  setPendingData]  = useState(null);      // buffered remote snapshot

  const saveTimer    = useRef(null);
  const pollTimer    = useRef(null);
  const lastEditRef  = useRef(0);
  const isSaving     = useRef(false);
  const dirtyKeysRef = useRef(new Set()); // date|itemKey pairs edited since last save

  // ── Immediate save function ──
  const doSave = useCallback(async () => {
    if (!enabled || isSaving.current) return;
    isSaving.current = true;
    setSaveStatus("saving");
    clearTimeout(saveTimer.current);
    const dirtySnapshot = new Set(dirtyKeysRef.current);
    dirtyKeysRef.current.clear();
    try {
      await saveFn(tokenRef.current, getData());
      setSaveStatus("saved");
    } catch (e) {
      // Restore dirty keys so the next save attempt retries them
      dirtySnapshot.forEach(k => dirtyKeysRef.current.add(k));
      setSaveStatus("error");
      notify.error("Save failed: " + e.message, 6000);
      setTimeout(() => setSaveStatus(prev => prev === "error" ? "unsaved" : prev), 4000);
    } finally {
      isSaving.current = false;
    }
  }, [enabled, saveFn, getData, tokenRef]);

  // ── Mark dirty — track which keys changed, debounce auto-save 2.5s ──
  const markDirty = useCallback((dateItemKey) => {
    if (!enabled) return;
    if (dateItemKey) dirtyKeysRef.current.add(dateItemKey);
    lastEditRef.current = Date.now();
    setSaveStatus("unsaved");
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => { doSave(); }, 2500);
  }, [enabled, doSave]);

  // ── Background poll every 30s ──
  useEffect(() => {
    if (!enabled || !pullFn) return;
    const poll = async () => {
      if (Date.now() - lastEditRef.current < 5000) return;
      try {
        const remote = await pullFn(tokenRef.current);
        if (!remote) return;
        const diff = diffFn(getData(), remote);
        if (diff > 0) {
          setPendingData(remote);
          setPendingCount(diff);
        } else {
          setPendingCount(0);
          setPendingData(null);
        }
      } catch (_) { /* silent */ }
    };
    pollTimer.current = setInterval(poll, 30000);
    return () => clearInterval(pollTimer.current);
  }, [enabled, pullFn, diffFn, getData, tokenRef]);

  // ── Apply pending — merge remote into local, preserving any unsaved local edits ──
  const applyPending = useCallback(() => {
    if (!pendingData) return;
    const dirty = dirtyKeysRef.current;
    if (dirty.size === 0) {
      // No local unsaved edits — safe to replace entirely
      onApply(pendingData);
    } else {
      // Merge: keep locally-dirty rows, take remote for everything else
      const local = getData();
      const localMap = new Map(local.map(r => [`${r.date}|${r.itemKey}`, r]));
      const merged = pendingData.map(remoteRow => {
        const key = `${remoteRow.date}|${remoteRow.itemKey}`;
        return dirty.has(key) ? (localMap.get(key) || remoteRow) : remoteRow;
      });
      // Add any locally-dirty rows that don't exist in remote yet (new rows)
      local.forEach(r => {
        const key = `${r.date}|${r.itemKey}`;
        if (dirty.has(key) && !pendingData.some(pr => `${pr.date}|${pr.itemKey}` === key)) {
          merged.push(r);
        }
      });
      onApply(merged);
    }
    setPendingCount(0);
    setPendingData(null);
    setSaveStatus("saved");
  }, [pendingData, onApply, getData]);

  // ── Cleanup on unmount ──
  useEffect(() => () => {
    clearTimeout(saveTimer.current);
    clearInterval(pollTimer.current);
  }, []);

  return { markDirty, doSave, saveStatus, pendingCount, applyPending };
}

// ─── MeetingSyncBar — replaces the static "Auto-sync paused" banner ───────────
// Shows: save status + pending-changes badge
function MeetingSyncBar({ saveStatus, pendingCount, onApply, onSave }) {
  const statusText = {
    saved:   "Saved",
    saving:  "Saving…",
    unsaved: "Unsaved changes",
    error:   "Save error",
  }[saveStatus] || "Saved";

  const statusColor = {
    saved:   C.green25,
    saving:  C.blue25,
    unsaved: C.gold20,
    error:   C.red15,
  }[saveStatus] || C.green25;

  return (
    <div style={{ display: "flex", alignItems: "center", gap: 10, marginBottom: 16, flexWrap: "wrap" }}>
      {/* Save status pill */}
      <div style={{ display: "flex", alignItems: "center", gap: 6, padding: "5px 12px",
        background: C.surfaceWarm, border: `1px solid ${C.borderLight}`, borderRadius: 20,
        fontSize: 11, fontFamily: "'Helvetica Neue',Arial,sans-serif", color: C.textMuted }}>
        <span style={{ width: 7, height: 7, borderRadius: "50%", background: statusColor,
          flexShrink: 0, animation: saveStatus === "saving" ? "pulse 1s infinite" : "none" }}/>
        {statusText}
        {saveStatus === "unsaved" && (
          <button onClick={onSave} style={{ marginLeft: 4, background: "none", border: "none",
            color: C.blue25, cursor: "pointer", fontSize: 11, fontFamily: "'Helvetica Neue',Arial,sans-serif",
            fontWeight: 700, padding: 0 }}>
            Save now
          </button>
        )}
      </div>

      {/* Pending changes badge */}
      {pendingCount > 0 && (
        <button onClick={onApply} style={{ display: "flex", alignItems: "center", gap: 6,
          padding: "5px 12px", background: "#EAF0F6", border: `1.5px solid ${C.blue25}`,
          borderRadius: 20, fontSize: 11, fontFamily: "'Helvetica Neue',Arial,sans-serif",
          color: C.blue35, fontWeight: 700, cursor: "pointer" }}>
          <ArrowDown size={11}/>
          {pendingCount} update{pendingCount !== 1 ? "s" : ""} available — click to apply
        </button>
      )}
    </div>
  );
}

// ─── Bishopric Council Tab ────────────────────────────────────────────────────

// Get ISO week number from a date string YYYY-MM-DD
function isoWeek(dateStr) {
  const d = new Date(dateStr);
  const jan4 = new Date(d.getFullYear(), 0, 4);
  const startOfWeek1 = new Date(jan4);
  startOfWeek1.setDate(jan4.getDate() - (jan4.getDay() || 7) + 1);
  return Math.floor((d - startOfWeek1) / 604800000) + 1;
}

// Sundays surrounding a date (prev 2, next 4)
function localDateStr(d) {
  // Format a Date as YYYY-MM-DD using LOCAL date parts — never toISOString()
  // which would shift to UTC and produce the wrong day in negative-offset timezones.
  const pad = n => String(n).padStart(2, '0');
  return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
}

function getUpcomingSunday() {
  // Returns YYYY-MM-DD of the upcoming Sunday (returns today if today IS Sunday).
  const now = new Date();
  const daysUntilSun = (7 - now.getDay()) % 7; // 0 if today is Sunday
  const s = new Date(now.getFullYear(), now.getMonth(), now.getDate() + daysUntilSun);
  return localDateStr(s);
}

function getSurroundingSundays() {
  // Always anchored to the real upcoming Sunday — never the selected date.
  // Returns exactly 5 Sundays: [−2 weeks, −1 week, upcoming, +1 week, +2 weeks].
  const upcomingStr = getUpcomingSunday();
  const [y, m, day] = upcomingStr.split("-").map(Number);
  const anchor = new Date(y, m - 1, day); // local midnight — no UTC drift
  const sundays = [];
  for (let i = -2; i <= 2; i++) {
    const s = new Date(anchor.getFullYear(), anchor.getMonth(), anchor.getDate() + i * 7);
    sundays.push(localDateStr(s)); // local date parts — never toISOString()
  }
  return sundays; // chronological order, exactly 5 entries
}

function toDisplayDate(dateStr) {
  if (!dateStr) return "";
  const [y, m, d] = dateStr.split("-").map(Number);
  return new Date(y, m - 1, d).toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
}

// Template items in order — itemKey is stable identifier used in sheet
const BM_TEMPLATE = [
  { itemKey: "prayer_list",      section: "Opening",   label: "Review Prayer List",      hasAssignee: false, isStatic: true,  isPrayerList: true },
  { itemKey: "opening_song",     section: "Opening",   label: "Opening Song",            hasAssignee: true,  isStatic: false, hasSongField: true },
  { itemKey: "opening_prayer",   section: "Opening",   label: "Opening Prayer",          hasAssignee: true,  isStatic: false },
  { itemKey: "spirit_thought",   section: "Opening",   label: "Spiritual Thought",       hasAssignee: true,  isStatic: false, alternates: true },
  { itemKey: "sacrament_prep",   section: "Business",  label: "Sacrament Meeting Prep",  hasAssignee: false, isStatic: true,  isLive: "sacrament" },
  { itemKey: "outstanding_tasks",section: "Business",  label: "Review Outstanding Tasks",hasAssignee: false, isStatic: true,  isLive: "tasks" },
  { itemKey: "discussion",       section: "Business",  label: "Discussion Topics",       hasAssignee: false, isStatic: false, isList: true },
  { itemKey: "calling_pipeline", section: "Business",  label: "Calling Pipeline",        hasAssignee: false, isStatic: true,  isLive: "callings" },
  { itemKey: "review_calendar",  section: "Business",  label: "Review Calendar",          hasAssignee: false, isStatic: true,  isLive: "calendar" },
  { itemKey: "closing_prayer",   section: "Closing",   label: "Closing Prayer",          hasAssignee: true,  isStatic: false },
];

const BM_SECTIONS = ["Opening", "Business", "Closing"];
const BM_SECTION_STYLE = {
  "Opening":  { color: C.blue35,  bg: "#EAF0F6", border: C.blue25 },
  "Business": { color: C.textPrimary, bg: C.surfaceWarm, border: C.border },
  "Closing":  { color: C.purple,  bg: "#F3EDF8", border: "#7B5EA7" },
};

function BishopricCouncilTab({ bishopricMeeting, setBishopricMeeting, callings, releasings, sacramentProgram, calendar=[], roster=[], token="", onNavigate }) {
  const today = localDateStr(new Date());
  // Find next Sunday
  const nextSundayDate = getUpcomingSunday();

  const [selectedDate, setSelectedDate] = useState(nextSundayDate);
  const [bmData, setBmData] = useState(bishopricMeeting);
  const [syncStatus, setSyncStatus] = useState("idle");
  const [sending, setSending] = useState(false);
  const [showOtherDate, setShowOtherDate] = useState(false);
  const [slackDraft, setSlackDraft] = useState(null); // {firstLine, bodyLines, relayURL, channelKey, channelName}
  const [showPrayerList, setShowPrayerList] = useState(false);
  const [prayerListData, setPrayerListData] = useState(null);
  const [prayerListLoading, setPrayerListLoading] = useState(false);
  const tokenRef = useRef(token);
  useEffect(() => { tokenRef.current = token; }, [token]);

  // Keep local bmData in sync with parent (on initial pull)
  useEffect(() => { setBmData(bishopricMeeting); }, [bishopricMeeting]);

  // Pause auto-pull while on this tab (same pattern as Sacrament)
  // (handled in parent via tabRef)

  const sundays = getSurroundingSundays(); // always anchored to upcoming Sunday
  const week = isoWeek(selectedDate);
  const isEvenWeek = week % 2 === 0;

  // Items for this date from sheet data
  const dateItems = bmData.filter(r => r.date === selectedDate);
  const hasData   = dateItems.length > 0;

  // ── Rotation helper — balanced auto-assignment across bishopric ──
  // ASSIGNABLE_KEYS: the 4 slots that get rotated
  const ASSIGNABLE_KEYS = ["opening_prayer", "opening_song", "spirit_thought", "closing_prayer"];

  // Returns the suggested assignee name for a given itemKey.
  // Picks the bishopric member with the fewest past assignments to that slot,
  // with ties broken by who was assigned least recently.
  // exclude: set of names already assigned to OTHER slots in this same meeting —
  //          one person cannot hold more than one assignment per meeting.
  const suggestAssignee = (itemKey, exclude = new Set()) => {
    const fullPool = ROSTER_POSITIONS
      .filter(p => p.group === "bishopric")
      .map(p => rosterName(roster, p.role))
      .filter(Boolean);
    if (!fullPool.length) return "";

    // Remove anyone already used elsewhere in this meeting
    const pool = fullPool.filter(n => !exclude.has(n));
    if (!pool.length) return ""; // all members already assigned (edge case: tiny group)

    // Gather all past rows for this itemKey (across all dates, not just selected)
    const pastRows = bmData
      .filter(r => r.itemKey === itemKey && r.assignee && r.date !== selectedDate)
      .sort((a, b) => (a.date > b.date ? -1 : 1)); // most recent first

    // Count assignments per person
    const counts = {};
    pool.forEach(name => { counts[name] = 0; });
    pastRows.forEach(r => { if (counts[r.assignee] !== undefined) counts[r.assignee]++; });

    // Find minimum count
    const minCount = Math.min(...pool.map(n => counts[n]));
    const tied = pool.filter(n => counts[n] === minCount);

    // Among tied, pick whoever hasn't been assigned most recently
    // (i.e., the one whose last assignment was furthest in the past, or never)
    const lastAssigned = {};
    tied.forEach(name => {
      const lastRow = pastRows.find(r => r.assignee === name);
      lastAssigned[name] = lastRow ? lastRow.date : "0000-00-00";
    });
    tied.sort((a, b) => (lastAssigned[a] < lastAssigned[b] ? -1 : 1));
    return tied[0];
  };

  // Auto-assign all 4 unassigned slots at once.
  // usedThisMeeting grows as each slot is filled so no one is double-booked.
  const doAutoAssign = () => {
    // Seed with anyone already saved to a slot for this date
    const usedThisMeeting = new Set(
      ASSIGNABLE_KEYS.map(k => getItem(k)?.assignee).filter(Boolean)
    );
    ASSIGNABLE_KEYS.forEach(key => {
      if (!getItem(key)?.assignee) {
        const suggestion = suggestAssignee(key, usedThisMeeting);
        if (suggestion) {
          updateItem(key, "assignee", suggestion);
          usedThisMeeting.add(suggestion); // block from all subsequent slots in this pass
        }
      }
    });
    notify.success("Auto-assigned open slots based on rotation balance");
  };

  // Auto-assign a single slot — excludes anyone already assigned elsewhere this meeting.
  const doAutoAssignOne = (itemKey) => {
    // Collect names already committed to OTHER assignable slots for this date
    const usedElsewhere = new Set(
      ASSIGNABLE_KEYS
        .filter(k => k !== itemKey)
        .map(k => getItem(k)?.assignee)
        .filter(Boolean)
    );
    const suggestion = suggestAssignee(itemKey, usedElsewhere);
    if (suggestion) updateItem(itemKey, "assignee", suggestion);
  };

  // Get persisted value for an itemKey
  const getItem = (key) => dateItems.find(r => r.itemKey === key) || null;

  // For spirit_thought alternating label
  const spiritLabel = (() => {
    const override = getItem("spirit_thought")?.spiritToggle;
    if (override) return override;
    return isEvenWeek ? "handbook_review" : "spiritual_thought";
  })();

  // Open prayer list modal — fetch from separate sheet
  const openPrayerList = async () => {
    setShowPrayerList(true);
    if (prayerListData !== null) return; // already loaded
    const plid = config.PRAYER_LIST_SHEET_ID;
    if (!plid || plid.includes("YOUR_")) {
      setPrayerListData([]);
      return;
    }
    setPrayerListLoading(true);
    try {
      const { pullPrayerList } = await import("./sheets");
      const data = await pullPrayerList(tokenRef.current);
      setPrayerListData(data);
    } catch(e) {
      setPrayerListData([]);
      notify.error("Could not load prayer list: " + e.message);
    } finally {
      setPrayerListLoading(false);
    }
  };

  // Create program from template — carries over unchecked discussion topics
  // and incomplete tasks from the most recent prior week.
  const createProgram = () => {
    // Find the most recent prior date that has agenda data
    const priorDates = [...new Set(bmData.map(r => r.date))]
      .filter(d => d < selectedDate)
      .sort()
      .reverse();
    const priorDate = priorDates[0] || null;

    // Carry over unchecked discussion topics (each is now its own row)
    const priorTopicRows = priorDate
      ? bmData.filter(r => r.date === priorDate && r.itemKey.startsWith("topic_") && !r.done)
      : [];
    const carryTopicRows = priorTopicRows.map(r => ({
      ...r, id: `bm_new_topic_carry_${Date.now()}_${r.id}`,
      date: selectedDate,
    }));

    // Carry over incomplete tasks (each is now its own row)
    const priorTaskRows = priorDate
      ? bmData.filter(r => r.date === priorDate && r.itemKey.startsWith("task_") && !r.done)
      : [];
    const carryTaskRows = priorTaskRows.map(r => ({
      ...r, id: `bm_new_task_carry_${Date.now()}_${r.id}`,
      date: selectedDate,
    }));

    const rows = BM_TEMPLATE.map(t => ({
      id: `bm_new_${t.itemKey}`,
      date: selectedDate,
      itemKey: t.itemKey,
      assignee: "",
      done: false,
      notes: "",
      customLabel: "",
      spiritToggle: t.itemKey === "spirit_thought" ? (isEvenWeek ? "handbook_review" : "spiritual_thought") : "",
    }));
    const newData = [...bmData.filter(r => r.date !== selectedDate), ...rows, ...carryTopicRows, ...carryTaskRows];
    setBmData(newData);
    setBishopricMeeting(newData);
    if (carryTopicRows.length || carryTaskRows.length) {
      notify.info(`Carried over ${carryTopicRows.length} topic${carryTopicRows.length!==1?"s":""} and ${carryTaskRows.length} task${carryTaskRows.length!==1?"s":""} from last week`);
    }
  };

  // Update a field for an itemKey on this date
  const updateItem = (itemKey, field, value) => {
    setBmData(prev => {
      const exists = prev.find(r => r.date === selectedDate && r.itemKey === itemKey);
      let next;
      if (exists) {
        next = prev.map(r => r.date === selectedDate && r.itemKey === itemKey ? { ...r, [field]: value } : r);
      } else {
        next = [...prev, { id: `bm_${itemKey}_${Date.now()}`, date: selectedDate, itemKey, assignee: "", done: false, notes: "", customLabel: "", spiritToggle: "", [field]: value }];
      }
      setBishopricMeeting(next);
      return next;
    });
    bmMarkDirty(`${selectedDate}|${itemKey}`);
  };

  // ── Discussion topics ──
  // Stored as JSON [{id, text, done}] in discussion row's notes.
  // Falls back gracefully from old newline-separated plain text.
  // Discussion topics — each is its own row with itemKey "topic_<id>"
  const discussionTopics = dateItems
    .filter(r => r.itemKey.startsWith("topic_"))
    .sort((a, b) => a.id < b.id ? -1 : 1);

  const addTopic = (text) => {
    if (!text.trim()) return;
    const id = `topic_${Date.now()}`;
    const newRow = { id: `bm_${id}`, date: selectedDate, itemKey: id, assignee: "", done: false, notes: text.trim(), customLabel: "", spiritToggle: "" };
    const newData = [...bmData, newRow];
    setBmData(newData); setBishopricMeeting(newData);
    bmMarkDirty(`${selectedDate}|${id}`);
  };

  const toggleTopic = (id) => {
    const key = id; // id IS the itemKey for topic rows
    updateItem(key, "done", !dateItems.find(r => r.itemKey === key)?.done);
  };

  const removeTopic = (id) => {
    const newData = bmData.filter(r => !(r.date === selectedDate && r.itemKey === id));
    setBmData(newData); setBishopricMeeting(newData);
    bmMarkDirty(`${selectedDate}|${id}`);
  };

  // ── Outstanding tasks ──
  // Stored as JSON [{id, text, assignee, done}] in outstanding_tasks row's notes.
  // Outstanding tasks — each is its own row with itemKey "task_<id>"
  const outstandingTasks = dateItems
    .filter(r => r.itemKey.startsWith("task_"))
    .sort((a, b) => a.id < b.id ? -1 : 1);

  const addTask = (text, assignee) => {
    if (!text.trim()) return;
    const id = `task_${Date.now()}`;
    const newRow = { id: `bm_${id}`, date: selectedDate, itemKey: id, assignee: assignee || "", done: false, notes: text.trim(), customLabel: "", spiritToggle: "" };
    const newData = [...bmData, newRow];
    setBmData(newData); setBishopricMeeting(newData);
    bmMarkDirty(`${selectedDate}|${id}`);
  };

  const toggleTask = (id) => {
    updateItem(id, "done", !dateItems.find(r => r.itemKey === id)?.done);
  };

  const removeTask = (id) => {
    const newData = bmData.filter(r => !(r.date === selectedDate && r.itemKey === id));
    setBmData(newData); setBishopricMeeting(newData);
    bmMarkDirty(`${selectedDate}|${id}`);
  };

  const updateTaskField = (itemKey, field, value) => {
    // "text" maps to the notes column; all other fields map directly
    const col = field === "text" ? "notes" : field;
    updateItem(itemKey, col, value);
  };

  // Save to sheet
  // ── Meeting sync ──
  const bmDiff = useCallback((local, remote) => {
    if (!remote || !Array.isArray(remote)) return 0;
    // Count rows that differ by assignee, done, notes, customLabel, or spiritToggle
    // for the currently selected date — ignore other dates to reduce noise
    const localDate  = local.filter(r => r.date === selectedDate);
    const remoteDate = remote.filter(r => r.date === selectedDate);
    if (localDate.length !== remoteDate.length) return Math.abs(localDate.length - remoteDate.length);
    return localDate.filter((r, i) => {
      const rem = remoteDate[i];
      return !rem || r.assignee !== rem.assignee || r.done !== rem.done ||
             r.notes !== rem.notes || r.customLabel !== rem.customLabel;
    }).length;
  }, [selectedDate]);

  const bmDataRef = useRef(bmData);
  useEffect(() => { bmDataRef.current = bmData; }, [bmData]);

  const { markDirty: bmMarkDirty, doSave, saveStatus: bmSaveStatus,
          pendingCount: bmPendingCount, applyPending: bmApplyPending } = useMeetingSync({
    getData:  () => bmDataRef.current,
    saveFn:   async (tok, data) => {
      if (!tok) return; // skip if token not yet available
      const { pushBishopricMeeting } = await import("./sheets");
      await pushBishopricMeeting(tok, data);
    },
    pullFn:   async (tok) => {
      const { pullAll } = await import("./sheets");
      const d = await pullAll(tok);
      return d.bishopricMeeting || [];
    },
    diffFn:   bmDiff,
    tokenRef,
    onApply:  (remote) => { setBmData(remote); setBishopricMeeting(remote); },
    enabled:  true,
  });

  // Send to Slack
  // Build message lines and open preview modal
  const doSendSlack = () => {
    const relayURL = config.SLACK_RELAY_URL || "";
    const webhookURL = config.SLACK_WEBHOOK_BISHOPRIC || "";
    const channelName = config.SLACK_WEBHOOK_BISHOPRIC_NAME || "bishopric";
    if (!relayURL || relayURL.includes("YOUR_")) { notify.error("No relay URL configured in config.js"); return; }
    if (!webhookURL || webhookURL.includes("YOUR/")) { notify.error("No bishopric Slack webhook configured in config.js"); return; }

    const approvedCallings  = callings.filter(c => c.stage === "Approved to Call").length;
    const approvedReleasings= releasings.filter(r => r.stage === "Approved to Release").length;
    const sacrItems = sacramentProgram.filter(s => s.date === selectedDate);
    const speakers  = sacrItems.filter(s => s.section === "Speaker").map(s => s.value).filter(Boolean);
    const organist  = sacrItems.find(s => s.label === "Organist")?.value || "—";
    const chorister = sacrItems.find(s => s.label === "Chorister")?.value || "—";
    const sLabel = spiritLabel === "handbook_review" ? "Handbook Review" : "Spiritual Thought";
    const songRow = getItem("opening_song");
    const songNum = songRow?.notes || "";
    const songTitle = songRow?.customLabel || "";
    const songLine = [songNum, songTitle].filter(Boolean).join(" — ") || "_not set_";

    // Assignments only — the four people-facing roles for the meeting
    const assign = (label, value) => (value && value !== "_unassigned_") ? `${label}: ${value}` : null;
    const assignments = [
      assign("Opening Song",  songRow?.assignee),
      assign("Opening Prayer", getItem("opening_prayer")?.assignee),
      assign(sLabel,           getItem("spirit_thought")?.assignee),
      assign("Closing Prayer", getItem("closing_prayer")?.assignee),
    ].filter(Boolean);

    const divider = "─────────────────────────────";
    const bodyLines = [
      divider,
      assignments.length ? assignments.join("\n") : "_No assignments set_",
      divider,
      `_Ward Manager · ${config.WARD_NAME}_`,
    ];

    setSlackDraft({ firstLine: `*Bishopric Meeting — ${toDisplayDate(selectedDate)}*`, bodyLines, relayURL, channelKey: webhookURL, channelName });
  };

  // Actually send after user confirms in modal
  const doConfirmSlack = async (firstLine) => {
    if (!slackDraft) return;
    const { bodyLines, relayURL, channelKey, channelName } = slackDraft;
    const text = [firstLine, ...bodyLines].join("\n");
    setSending(true);
    try {
      const res = await fetch(relayURL, {
        method: "POST", headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({ channel: channelKey, text }),
      });
      const data = await res.json().catch(() => ({ status: res.status }));
      if (res.ok && data.status === 200) {
        notify.success(`Agenda sent to #${channelName}`);
        setSlackDraft(null);
      } else {
        notify.error(`Relay error: ${data.message || res.status}`);
      }
    } catch (e) {
      notify.error("Failed to reach relay: " + e.message);
    } finally { setSending(false); }
  };

  // Live badge data
  const approvedCallings  = callings.filter(c => c.stage === "Approved to Call").length;
  const approvedReleasings= releasings.filter(r => r.stage === "Approved to Release").length;
  const sacrItems = sacramentProgram.filter(s => s.date === selectedDate);
  const speakers  = sacrItems.filter(s => s.section === "Speaker").map(s => s.value).filter(Boolean);
  const organist  = sacrItems.find(s => s.label === "Organist")?.value;
  const chorister = sacrItems.find(s => s.label === "Chorister")?.value;

  // Songs and prayers pulled from sacrament program for the live badge
  const sacrSongs = [
    { label: "Opening Hymn",   value: sacrItems.find(s => s.label === "Opening Hymn")?.value   || "" },
    { label: "Sacrament Hymn", value: sacrItems.find(s => s.label === "Sacrament Hymn")?.value || "" },
    { label: "Closing Hymn",   value: sacrItems.find(s => s.label === "Closing Hymn")?.value   || "" },
  ];
  const sacrPrayers = [
    { label: "Opening Prayer",  value: sacrItems.find(s => s.label === "Opening Prayer")?.value  || "" },
    { label: "Closing Prayer",  value: sacrItems.find(s => s.label === "Closing Prayer")?.value  || "" },
  ];

  // Calling pipeline counts by stage (excluding Completed)
  const callingStageActive  = CALLING_STAGES.filter(s => s !== "Completed");
  const releasingStageActive = RELEASING_STAGES.filter(s => s !== "Completed");
  const callingStageCounts  = callingStageActive.map(s  => ({ stage: s,  count: callings.filter(c  => c.stage  === s).length })).filter(x => x.count > 0);
  const releasingStageCounts = releasingStageActive.map(s => ({ stage: s, count: releasings.filter(r => r.stage === s).length })).filter(x => x.count > 0);

  // Upcoming calendar events: selectedDate + 1 through selectedDate + 7
  const upcomingCalendarEvents = (() => {
    if (!selectedDate) return [];
    const base = new Date(selectedDate + "T12:00:00");
    const pad = n => String(n).padStart(2, "0");
    const dayAfter = new Date(base); dayAfter.setDate(base.getDate() + 1);
    const weekOut  = new Date(base); weekOut.setDate(base.getDate() + 7);
    const startStr = `${dayAfter.getFullYear()}-${pad(dayAfter.getMonth()+1)}-${pad(dayAfter.getDate())}`;
    const endStr   = `${weekOut.getFullYear()}-${pad(weekOut.getMonth()+1)}-${pad(weekOut.getDate())}`;
    return calendar
      .filter(e => e.date >= startStr && e.date <= endStr)
      .sort((a, b) => a.date > b.date ? 1 : a.date < b.date ? -1 : (a.time > b.time ? 1 : -1));
  })();

  return (
    <div className="animate-in">
      {/* ── Hero ── */}
      <div style={{ background: `linear-gradient(130deg, ${C.blue40} 0%, ${C.blue35} 50%, ${C.blue25} 100%)`, borderRadius: 14, padding: "22px 28px", marginBottom: 24, color: "#fff", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
        <div>
          <div style={{ fontSize: 22, fontWeight: 600, letterSpacing: "-.01em", marginBottom: 2 }}>Bishopric Council</div>
          <div style={{ fontSize: 13, opacity: .75, fontFamily: "'Helvetica Neue',Arial,sans-serif" }}>{toDisplayDate(selectedDate)}</div>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button onClick={doSendSlack} disabled={sending || !hasData}
            style={{ background: "rgba(255,255,255,.14)", border: "1.5px solid rgba(255,255,255,.3)", color: "#fff", borderRadius: 8, padding: "7px 14px", fontSize: 12, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600, cursor: sending || !hasData ? "not-allowed" : "pointer", opacity: sending || !hasData ? .5 : 1, display: "flex", alignItems: "center", gap: 6 }}>
            <Bell size={13}/> {sending ? "Sending…" : "Send to Slack"}
          </button>
          <button onClick={doAutoAssign} disabled={!hasData}
            title="Fill unassigned slots using balanced rotation"
            style={{ background: "rgba(255,255,255,.14)", border: "1.5px solid rgba(255,255,255,.3)", color: "#fff", borderRadius: 8, padding: "7px 14px", fontSize: 12, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600, cursor: !hasData ? "not-allowed" : "pointer", opacity: !hasData ? .5 : 1, display: "flex", alignItems: "center", gap: 6 }}>
            <RotateCcw size={12}/> Auto-assign
          </button>

          <button onClick={doSave} disabled={syncStatus === "pushing" || !hasData}
            style={{ background: "rgba(255,255,255,.18)", border: "1.5px solid rgba(255,255,255,.4)", color: "#fff", borderRadius: 8, padding: "7px 14px", fontSize: 12, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600, cursor: syncStatus === "pushing" || !hasData ? "not-allowed" : "pointer", opacity: syncStatus === "pushing" || !hasData ? .5 : 1, display: "flex", alignItems: "center", gap: 6 }}>
            <Save size={13}/> {syncStatus === "pushing" ? "Saving…" : "Save"}
          </button>
        </div>
      </div>

      {/* Sync status bar */}
      <MeetingSyncBar saveStatus={bmSaveStatus} pendingCount={bmPendingCount}
        onApply={bmApplyPending} onSave={doSave}/>

      {/* ── Date nav ── */}
      <div style={{ display: "flex", gap: 6, marginBottom: 24, flexWrap: "wrap", alignItems: "center" }}>
        {sundays.map(s => {
          const isNext = s === nextSundayDate;
          const isSel  = s === selectedDate;
          return (
            <button key={s} onClick={() => setSelectedDate(s)}
              style={{ padding: "6px 14px", borderRadius: 20, border: `1.5px solid ${isSel ? C.blue35 : C.border}`, background: isSel ? C.blue35 : C.surfaceWhite, color: isSel ? "#fff" : C.textMuted, fontSize: 12, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: isSel ? 700 : 400, cursor: "pointer", display: "flex", alignItems: "center", gap: 6, position: "relative" }}>
              {toDisplayDate(s)}
              {isNext && <span style={{ fontSize: 9, fontWeight: 700, letterSpacing: ".06em", background: C.gold20, color: "#fff", borderRadius: 8, padding: "1px 6px" }}>NEXT</span>}
            </button>
          );
        })}
        <button onClick={() => setShowOtherDate(v => !v)}
          style={{ padding: "6px 14px", borderRadius: 20, border: `1.5px solid ${C.border}`, background: C.surfaceWhite, color: C.textMuted, fontSize: 12, fontFamily: "'Helvetica Neue',Arial,sans-serif", cursor: "pointer" }}>
          Other…
        </button>
        {showOtherDate && (
          <input type="date" value={selectedDate} onChange={e => { setSelectedDate(e.target.value); setShowOtherDate(false); }}
            style={{ width: 160, padding: "6px 10px", fontSize: 12 }}/>
        )}
        <div style={{ fontSize: 11, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", marginLeft: 6 }}>
          Week {week} · {spiritLabel === "handbook_review" ? "Handbook Review week" : "Spiritual Thought week"}
        </div>
      </div>

      {/* ── Empty state ── */}
      {!hasData && (
        <div style={{ textAlign: "center", padding: "60px 24px", background: C.surfaceWhite, borderRadius: 12, border: `1.5px solid ${C.border}` }}>
          <div style={{ marginBottom: 12, opacity: .3, color: C.textMuted, display: "flex", justifyContent: "center" }}><ClipboardList size={42}/></div>
          <div style={{ fontFamily: "Georgia,serif", fontSize: 17, color: C.textSecond, marginBottom: 8 }}>No agenda for this Sunday</div>
          <div style={{ fontSize: 12, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", marginBottom: 20 }}>Create the agenda from the standard template to get started</div>
          <button className="btn-primary" onClick={createProgram}><Plus size={13}/> Create Agenda</button>
        </div>
      )}

      {/* ── Agenda sections ── */}
      {hasData && BM_SECTIONS.map(section => {
        const sStyle = BM_SECTION_STYLE[section];
        const items  = BM_TEMPLATE.filter(t => t.section === section);
        return (
          <div key={section} style={{ marginBottom: 20, background: C.surfaceWhite, border: `1.5px solid ${C.border}`, borderRadius: 12, overflow: "hidden" }}>
            {/* Section header */}
            <div style={{ background: sStyle.bg, borderBottom: `1.5px solid ${sStyle.border}`, padding: "10px 20px", display: "flex", alignItems: "center", gap: 8 }}>
              <div style={{ width: 8, height: 8, borderRadius: "50%", background: sStyle.color, flexShrink: 0 }}/>
              <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: sStyle.color, fontFamily: "'Helvetica Neue',Arial,sans-serif" }}>{section}</span>
            </div>

            {/* Items */}
            <div style={{ display: "flex", flexDirection: "column" }}>
              {items.map((tmpl, idx) => {
                const row = getItem(tmpl.itemKey);
                const isLast = idx === items.length - 1;

                // Special: alternating label
                const displayLabel = tmpl.alternates
                  ? (spiritLabel === "handbook_review" ? "Handbook Review" : "Spiritual Thought")
                  : tmpl.label;

                return (
                  <div key={tmpl.itemKey} style={{ padding: "14px 20px", borderBottom: isLast ? "none" : `1px solid ${C.borderLight}`, display: "flex", alignItems: "flex-start", gap: 16 }}>
                    {/* Done checkbox */}
                    <button onClick={() => updateItem(tmpl.itemKey, "done", !(row?.done))}
                      style={{ width: 20, height: 20, borderRadius: 4, border: `2px solid ${row?.done ? C.green25 : C.border}`, background: row?.done ? C.green25 : "transparent", cursor: "pointer", flexShrink: 0, marginTop: 2, display: "flex", alignItems: "center", justifyContent: "center" }}>
                      {row?.done && <svg width="10" height="10" viewBox="0 0 10 10" fill="none"><path d="M1.5 5L4 7.5L8.5 2.5" stroke="white" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/></svg>}
                    </button>

                    <div style={{ flex: 1, minWidth: 0 }}>
                      {/* Label row */}
                      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                        <span style={{ fontFamily: "Georgia,serif", fontSize: 15, color: row?.done ? C.textMuted : C.textPrimary, textDecoration: row?.done ? "line-through" : "none", minWidth: 200 }}>
                          {displayLabel}
                        </span>

                        {/* Assignee field + per-slot rotation button */}
                        {tmpl.hasAssignee && (
                          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                            <select
                              value={row?.assignee || ""}
                              onChange={e => updateItem(tmpl.itemKey, "assignee", e.target.value)}
                              style={{ width: 200, padding: "5px 10px", fontSize: 13, background: row?.assignee ? C.surfaceWarm : "#fff", border: `1.5px solid ${row?.assignee ? C.border : C.borderLight}` }}>
                              <option value="">— Assign to… —</option>
                              {ROSTER_POSITIONS.filter(p => p.group === "bishopric").map(p => (
                                <option key={p.role} value={rosterName(roster, p.role)}>
                                  {rosterName(roster, p.role)}
                                </option>
                              ))}
                            </select>
                            {ASSIGNABLE_KEYS.includes(tmpl.itemKey) && (
                              <button
                                onClick={() => doAutoAssignOne(tmpl.itemKey)}
                                title={row?.assignee ? "Re-suggest based on rotation" : "Suggest based on rotation"}
                                style={{ width: 28, height: 28, borderRadius: 6, border: `1.5px solid ${C.borderLight}`, background: C.surfaceWarm, color: C.textMuted, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                                <RotateCcw size={13}/>
                              </button>
                            )}
                          </div>
                        )}

                        {/* Song number + title (Opening Song) */}
                        {tmpl.hasSongField && (
                          <>
                            <input
                              type="text"
                              placeholder="Hymn #"
                              value={row?.notes || ""}
                              onChange={e => updateItem(tmpl.itemKey, "notes", e.target.value)}
                              style={{ width: 72, padding: "5px 10px", fontSize: 13,
                                background: row?.notes ? C.surfaceWarm : "#fff",
                                border: `1.5px solid ${row?.notes ? C.border : C.borderLight}` }}
                            />
                            <input
                              type="text"
                              placeholder="Song title…"
                              value={row?.customLabel || ""}
                              onChange={e => updateItem(tmpl.itemKey, "customLabel", e.target.value)}
                              style={{ width: 200, padding: "5px 10px", fontSize: 13,
                                background: row?.customLabel ? C.surfaceWarm : "#fff",
                                border: `1.5px solid ${row?.customLabel ? C.border : C.borderLight}` }}
                            />
                          </>
                        )}

                        {/* Alternates toggle */}
                        {tmpl.alternates && (
                          <button onClick={() => updateItem(tmpl.itemKey, "spiritToggle", spiritLabel === "handbook_review" ? "spiritual_thought" : "handbook_review")}
                            style={{ fontSize: 11, padding: "3px 10px", borderRadius: 12, border: `1px solid ${C.border}`, background: C.surfaceWarm, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", cursor: "pointer" }}>
                            ⇄ Switch
                          </button>
                        )}
                      </div>

                      {/* Prayer List button */}
                      {tmpl.isPrayerList && (
                        <div style={{ marginTop: 8 }}>
                          <button
                            onClick={openPrayerList}
                            className="btn-secondary"
                            style={{ fontSize: 12, padding: "5px 14px", display: "inline-flex", alignItems: "center", gap: 6 }}>
                            <BookOpen size={13}/> View Prayer List
                          </button>
                        </div>
                      )}

                      {/* Live badge: Sacrament Prep */}
                      {tmpl.isLive === "sacrament" && (() => {
                        const hasSacrData = organist || chorister || speakers.length || sacrSongs.some(s => s.value) || sacrPrayers.some(p => p.value);
                        return (
                          <div style={{ marginTop: 10 }}>
                            {!hasSacrData ? (
                              <span style={{ fontSize: 11, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontStyle: "italic" }}>No sacrament program found for this Sunday</span>
                            ) : (
                              <>
                                {/* Accompaniment row */}
                                {(organist || chorister) && (
                                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 6 }}>
                                    {[organist && `Organist: ${organist}`, chorister && `Chorister: ${chorister}`].filter(Boolean).map((badge, i) => (
                                      <span key={i} style={{ fontSize: 11, padding: "3px 10px", borderRadius: 10, background: "#EAF4EA", border: `1px solid ${C.green25}`, color: C.green35, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600 }}>{badge}</span>
                                    ))}
                                  </div>
                                )}
                                {/* Songs row */}
                                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 6 }}>
                                  {sacrSongs.map(s => (
                                    <span key={s.label} style={{ fontSize: 11, padding: "3px 10px", borderRadius: 10, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600,
                                      background: s.value ? "#EAF0F6" : "#F5F3EE",
                                      border: `1px solid ${s.value ? C.blue25 : C.borderLight}`,
                                      color: s.value ? C.blue35 : C.textMuted }}>
                                      <Music2 size={11}/> {s.label}{s.value ? `: ${s.value}` : " —"}
                                    </span>
                                  ))}
                                </div>
                                {/* Prayers row */}
                                <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 6 }}>
                                  {sacrPrayers.map(p => (
                                    <span key={p.label} style={{ fontSize: 11, padding: "3px 10px", borderRadius: 10, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600,
                                      background: p.value ? "#FEF8ED" : "#F5F3EE",
                                      border: `1px solid ${p.value ? C.goldDeep : C.borderLight}`,
                                      color: p.value ? "#7A5200" : C.textMuted }}>
                                      <HandMetal size={11}/> {p.label}{p.value ? `: ${p.value}` : " —"}
                                    </span>
                                  ))}
                                </div>
                                {/* Speakers row */}
                                {speakers.length > 0 && (
                                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap", marginBottom: 6 }}>
                                    <span style={{ fontSize: 11, padding: "3px 10px", borderRadius: 10, background: "#EAF4EA", border: `1px solid ${C.green25}`, color: C.green35, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600 }}>
                                      Speakers: {speakers.join(", ")}
                                    </span>
                                  </div>
                                )}
                              </>
                            )}
                            {/* Nav link */}
                            {onNavigate && (
                              <button onClick={() => onNavigate("sacrament")}
                                style={{ marginTop: 4, background: "none", border: "none", padding: 0, cursor: "pointer", fontSize: 11, color: C.blue25, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600, display: "flex", alignItems: "center", gap: 4, textDecoration: "underline" }}>
                                Open Sacrament Meeting →
                              </button>
                            )}
                          </div>
                        );
                      })()}

                      {/* Outstanding tasks list */}
                      {tmpl.isLive === "tasks" && (
                        <div style={{ marginTop: 8 }}>
                          {outstandingTasks.length === 0 && (
                            <div style={{ fontSize: 12, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontStyle: "italic", marginBottom: 8 }}>No tasks yet</div>
                          )}
                          {outstandingTasks.map(task => (
                            <div key={task.itemKey} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8, background: task.done ? C.surfaceWarm : C.surfaceWhite, border: `1.5px solid ${task.done ? C.borderLight : C.border}`, borderRadius: 8, padding: "8px 12px" }}>
                              <button onClick={() => toggleTask(task.itemKey)}
                                style={{ width: 18, height: 18, borderRadius: 4, border: `2px solid ${task.done ? C.green25 : C.border}`, background: task.done ? C.green25 : "transparent", cursor: "pointer", flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center" }}>
                                {task.done && <svg width="9" height="9" viewBox="0 0 10 10" fill="none"><path d="M1.5 5L4 7.5L8.5 2.5" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></svg>}
                              </button>
                              <input
                                value={task.notes}
                                onChange={e => updateTaskField(task.itemKey, "text", e.target.value)}
                                placeholder="Task description…"
                                style={{ flex: 2, padding: "4px 8px", fontSize: 13, fontFamily: "'Helvetica Neue',Arial,sans-serif", color: task.done ? C.textMuted : C.textPrimary, textDecoration: task.done ? "line-through" : "none", background: "transparent", border: `1px solid ${C.borderLight}`, borderRadius: 5 }}
                              />
                              <select
                                value={task.assignee || ""}
                                onChange={e => updateTaskField(task.itemKey, "assignee", e.target.value)}
                                style={{ flex: 1, minWidth: 120, padding: "4px 8px", fontSize: 12, fontFamily: "'Helvetica Neue',Arial,sans-serif", background: "transparent", border: `1px solid ${C.borderLight}`, borderRadius: 5, color: task.assignee ? C.textPrimary : C.textMuted }}>
                                <option value="">— Assign —</option>
                                {ROSTER_POSITIONS.filter(p => p.group === "bishopric").map(p => (
                                  <option key={p.role} value={rosterName(roster, p.role)}>{rosterName(roster, p.role)}</option>
                                ))}
                              </select>
                              <button onClick={() => removeTask(task.itemKey)} style={{ background: "none", border: "none", cursor: "pointer", color: C.textLight, display: "flex", alignItems: "center", flexShrink: 0 }}><X size={13}/></button>
                            </div>
                          ))}
                          <AddTaskRow onAdd={addTask} roster={roster}/>
                        </div>
                      )}

                      {/* Live badge: Calling Pipeline — per-stage counts */}
                      {tmpl.isLive === "callings" && (
                        <div style={{ marginTop: 10 }}>
                          {callingStageCounts.length === 0 && releasingStageCounts.length === 0 ? (
                            <span style={{ fontSize: 11, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontStyle: "italic" }}>No active callings or releasings</span>
                          ) : (
                            <>
                              {/* Callings by stage */}
                              {callingStageCounts.length > 0 && (
                                <div style={{ marginBottom: 6 }}>
                                  <span style={{ fontSize: 10, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 700, letterSpacing: ".06em", textTransform: "uppercase", color: C.textMuted, marginRight: 6 }}>Callings</span>
                                  <div style={{ display: "inline-flex", gap: 5, flexWrap: "wrap" }}>
                                    {callingStageCounts.map(({ stage, count }) => (
                                      <span key={stage} style={{ fontSize: 11, padding: "2px 9px", borderRadius: 10, background: "#E6F4FA", border: `1px solid ${C.blue25}`, color: C.blue35, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600 }}>
                                        {count} {stage}
                                      </span>
                                    ))}
                                  </div>
                                </div>
                              )}
                              {/* Releasings by stage */}
                              {releasingStageCounts.length > 0 && (
                                <div style={{ marginBottom: 6 }}>
                                  <span style={{ fontSize: 10, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 700, letterSpacing: ".06em", textTransform: "uppercase", color: C.textMuted, marginRight: 6 }}>Releasings</span>
                                  <div style={{ display: "inline-flex", gap: 5, flexWrap: "wrap" }}>
                                    {releasingStageCounts.map(({ stage, count }) => (
                                      <span key={stage} style={{ fontSize: 11, padding: "2px 9px", borderRadius: 10, background: "#F3EDF8", border: `1px solid ${C.purple}`, color: "#4A2A7A", fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600 }}>
                                        {count} {stage}
                                      </span>
                                    ))}
                                  </div>
                                </div>
                              )}
                            </>
                          )}
                          {/* Nav links */}
                          {onNavigate && (
                            <div style={{ display: "flex", gap: 12, marginTop: 4 }}>
                              <button onClick={() => onNavigate("callings")}
                                style={{ background: "none", border: "none", padding: 0, cursor: "pointer", fontSize: 11, color: C.blue25, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600, display: "flex", alignItems: "center", gap: 4, textDecoration: "underline" }}>
                                Open Callings →
                              </button>
                              <button onClick={() => onNavigate("releasings")}
                                style={{ background: "none", border: "none", padding: 0, cursor: "pointer", fontSize: 11, color: C.purple, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600, display: "flex", alignItems: "center", gap: 4, textDecoration: "underline" }}>
                                Open Releasings →
                              </button>
                            </div>
                          )}
                        </div>
                      )}

                      {/* Live badge: Review Calendar */}
                      {tmpl.isLive === "calendar" && (
                        <div style={{ marginTop: 10 }}>
                          {upcomingCalendarEvents.length === 0 ? (
                            <span style={{ fontSize: 11, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontStyle: "italic" }}>No events in the next 7 days</span>
                          ) : (
                            <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                              {upcomingCalendarEvents.map(ev => {
                                const evDate = new Date(ev.date + "T12:00:00");
                                const dayLabel = evDate.toLocaleDateString("default", { weekday: "short", month: "short", day: "numeric" });
                                return (
                                  <div key={ev.id} style={{ display: "flex", alignItems: "baseline", gap: 8 }}>
                                    <span style={{ fontSize: 11, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 700, color: C.blue25, flexShrink: 0, minWidth: 90 }}>{dayLabel}</span>
                                    {ev.time && <span style={{ fontSize: 11, fontFamily: "'Helvetica Neue',Arial,sans-serif", color: C.textMuted, flexShrink: 0 }}>{ev.time}</span>}
                                    <span style={{ fontSize: 13, fontFamily: "Georgia,serif", color: C.textPrimary }}>{ev.event}</span>
                                  </div>
                                );
                              })}
                            </div>
                          )}
                          {onNavigate && (
                            <button onClick={() => onNavigate("calendar")}
                              style={{ marginTop: 8, background: "none", border: "none", padding: 0, cursor: "pointer", fontSize: 11, color: C.blue25, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600, textDecoration: "underline" }}>
                              Open Calendar →
                            </button>
                          )}
                        </div>
                      )}

                      {/* Discussion topics checklist */}
                      {tmpl.isList && (
                        <div style={{ marginTop: 8 }}>
                          {discussionTopics.length === 0 && (
                            <div style={{ fontSize: 12, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontStyle: "italic", marginBottom: 8 }}>No topics yet</div>
                          )}
                          {discussionTopics.map(topic => (
                            <div key={topic.itemKey} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                              <button onClick={() => toggleTopic(topic.itemKey)}
                                style={{ width: 18, height: 18, borderRadius: 4, border: `2px solid ${topic.done ? C.green25 : C.border}`, background: topic.done ? C.green25 : "transparent", cursor: "pointer", flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center" }}>
                                {topic.done && <svg width="9" height="9" viewBox="0 0 10 10" fill="none"><path d="M1.5 5L4 7.5L8.5 2.5" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></svg>}
                              </button>
                              <span style={{ flex: 1, fontSize: 13, fontFamily: "'Helvetica Neue',Arial,sans-serif", color: topic.done ? C.textMuted : C.textPrimary, textDecoration: topic.done ? "line-through" : "none" }}>{topic.notes}</span>
                              <button onClick={() => removeTopic(topic.itemKey)} style={{ background: "none", border: "none", cursor: "pointer", color: C.textLight, display: "flex", alignItems: "center" }}><X size={13}/></button>
                            </div>
                          ))}
                          <AddTopicRow onAdd={addTopic}/>
                        </div>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}

      {/* Prayer List Modal */}
      {showPrayerList && (
        <PrayerListModal
          data={prayerListData}
          loading={prayerListLoading}
          sheetUrl={config.PRAYER_LIST_SHEET_URL}
          onClose={() => setShowPrayerList(false)}
          onRefresh={() => { setPrayerListData(null); openPrayerList(); }}
        />
      )}

      {/* Slack Preview Modal */}
      {slackDraft && (() => {
        const fl = slackDraft.firstLine;
        const setFl = v => setSlackDraft(d => ({...d, firstLine: v}));
        return (
          <SlackPreviewModal
            firstLine={fl}
            setFirstLine={setFl}
            bodyLines={slackDraft.bodyLines}
            channelName={slackDraft.channelName}
            sending={sending}
            onConfirm={() => doConfirmSlack(fl)}
            onClose={() => setSlackDraft(null)}
          />
        );
      })()}
    </div>
  );
}

function PrayerListModal({ data, loading, sheetUrl, onClose, onRefresh }) {
  // Group entries by category
  const grouped = (data || []).reduce((acc, row) => {
    const cat = row.category || "General";
    if (!acc[cat]) acc[cat] = [];
    acc[cat].push(row.name);
    return acc;
  }, {});
  const categories = Object.keys(grouped);
  const totalCount = (data || []).length;
  const isNotConfigured = !config.PRAYER_LIST_SHEET_ID || config.PRAYER_LIST_SHEET_ID.includes("YOUR_");

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.45)", zIndex: 1000,
      display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}
      onClick={e => e.target === e.currentTarget && onClose()}>
      <div style={{ background: "#fff", borderRadius: 16, width: "100%", maxWidth: 520,
        maxHeight: "80vh", display: "flex", flexDirection: "column",
        boxShadow: "0 20px 60px rgba(0,0,0,.2)" }}>

        {/* Header */}
        <div style={{ padding: "20px 24px 16px", borderBottom: `1px solid ${C.borderLight}`,
          display: "flex", alignItems: "center", justifyContent: "space-between", flexShrink: 0 }}>
          <div>
            <div style={{ fontFamily: "Georgia,serif", fontSize: 18, color: C.textPrimary, marginBottom: 2 }}>Prayer List</div>
            {!loading && data && !isNotConfigured && (
              <div style={{ fontSize: 11, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif" }}>
                {totalCount} {totalCount === 1 ? "name" : "names"} · read-only
              </div>
            )}
          </div>
          <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
            {!isNotConfigured && sheetUrl && (
              <a href={sheetUrl} target="_blank" rel="noopener noreferrer"
                className="btn-secondary"
                style={{ fontSize: 12, padding: "6px 14px", display: "inline-flex", alignItems: "center", gap: 6, textDecoration: "none" }}>
                <ExternalLink size={12}/> Edit in Sheets
              </a>
            )}
            <button onClick={onClose} style={{ background: "none", border: "none", cursor: "pointer",
              color: C.textMuted, display: "flex", alignItems: "center", padding: 4 }}>
              <X size={18}/>
            </button>
          </div>
        </div>

        {/* Body */}
        <div style={{ overflowY: "auto", padding: "16px 24px 24px", flex: 1 }}>

          {/* Not configured */}
          {isNotConfigured && (
            <div style={{ textAlign: "center", padding: "32px 16px" }}>
              <div style={{ fontSize: 13, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", lineHeight: 1.6 }}>
                Prayer list sheet not configured.<br/>
                Add <code style={{ background: C.surfaceWarm, padding: "1px 5px", borderRadius: 4, fontSize: 12 }}>PRAYER_LIST_SHEET_ID</code> to <code style={{ background: C.surfaceWarm, padding: "1px 5px", borderRadius: 4, fontSize: 12 }}>config.js</code>.
              </div>
            </div>
          )}

          {/* Loading */}
          {!isNotConfigured && loading && (
            <div style={{ textAlign: "center", padding: "40px 16px", color: C.textMuted,
              fontFamily: "'Helvetica Neue',Arial,sans-serif", fontSize: 13 }}>
              Loading…
            </div>
          )}

          {/* Empty */}
          {!isNotConfigured && !loading && data && totalCount === 0 && (
            <div style={{ textAlign: "center", padding: "40px 16px", color: C.textMuted,
              fontFamily: "'Helvetica Neue',Arial,sans-serif", fontSize: 13, fontStyle: "italic" }}>
              No names on the prayer list
            </div>
          )}

          {/* Grouped list */}
          {!isNotConfigured && !loading && data && totalCount > 0 && (
            <div style={{ display: "flex", flexDirection: "column", gap: 16 }}>
              {categories.map(cat => (
                <div key={cat}>
                  <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: ".1em",
                    textTransform: "uppercase", color: C.blue35,
                    fontFamily: "'Helvetica Neue',Arial,sans-serif",
                    marginBottom: 8, paddingBottom: 6,
                    borderBottom: `1px solid ${C.borderLight}` }}>{cat}</div>
                  <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                    {grouped[cat].map((name, i) => (
                      <div key={i} style={{ display: "flex", alignItems: "center", gap: 10,
                        padding: "6px 12px", background: C.surfaceWarm, borderRadius: 7 }}>
                        <div style={{ width: 5, height: 5, borderRadius: "50%",
                          background: C.blue25, flexShrink: 0 }}/>
                        <span style={{ fontSize: 14, fontFamily: "Georgia,serif",
                          color: C.textPrimary }}>{name}</span>
                      </div>
                    ))}
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>

        {/* Footer */}
        {!isNotConfigured && !loading && data && (
          <div style={{ padding: "12px 24px", borderTop: `1px solid ${C.borderLight}`,
            display: "flex", justifyContent: "space-between", alignItems: "center", flexShrink: 0 }}>
            <button onClick={onRefresh} className="btn-secondary"
              style={{ fontSize: 12, padding: "5px 14px", display: "inline-flex", alignItems: "center", gap: 6 }}>
              <RefreshCw size={12}/> Refresh
            </button>
            <span style={{ fontSize: 11, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif" }}>Changes must be made in Google Sheets</span>
          </div>
        )}
      </div>
    </div>
  );
}

function AddTopicRow({ onAdd }) {
  const [val, setVal] = useState("");
  const submit = () => { if (val.trim()) { onAdd(val.trim()); setVal(""); } };
  return (
    <div style={{ display: "flex", gap: 8, marginTop: 6 }}>
      <input type="text" placeholder="Add discussion topic…" value={val}
        onChange={e => setVal(e.target.value)}
        onKeyDown={e => e.key === "Enter" && submit()}
        style={{ flex: 1, padding: "6px 10px", fontSize: 13 }}/>
      <button onClick={submit} className="btn-secondary" style={{ padding: "6px 14px", fontSize: 12, display: "flex", alignItems: "center", gap: 4 }}>
        <Plus size={12}/> Add
      </button>
    </div>
  );
}

function AddTaskRow({ onAdd, roster = [], allRoles = false }) {
  const [text, setText] = useState("");
  const [assignee, setAssignee] = useState("");
  const submit = () => {
    if (text.trim()) { onAdd(text.trim(), assignee); setText(""); setAssignee(""); }
  };
  const positions = allRoles ? ROSTER_POSITIONS : ROSTER_POSITIONS.filter(p => p.group === "bishopric");
  return (
    <div style={{ display: "flex", gap: 8, marginTop: 8, alignItems: "center" }}>
      <input type="text" placeholder="New task…" value={text}
        onChange={e => setText(e.target.value)}
        onKeyDown={e => e.key === "Enter" && submit()}
        style={{ flex: 2, padding: "6px 10px", fontSize: 13, borderRadius: 6, border: `1.5px solid ${C.borderLight}` }}/>
      <select value={assignee} onChange={e => setAssignee(e.target.value)}
        style={{ flex: 1, minWidth: 120, padding: "6px 10px", fontSize: 12, borderRadius: 6, border: `1.5px solid ${C.borderLight}`, color: assignee ? C.textPrimary : C.textMuted, background: C.surfaceWhite }}>
        <option value="">— Assign —</option>
        {positions.map(p => (
          <option key={p.role} value={rosterName(roster, p.role)}>{rosterName(roster, p.role) || p.role}</option>
        ))}
      </select>
      <button onClick={submit} className="btn-secondary" style={{ padding: "6px 14px", fontSize: 12, display: "flex", alignItems: "center", gap: 4, flexShrink: 0 }}>
        <Plus size={12}/> Add Task
      </button>
    </div>
  );
}


// ─── Calendar Tab ──────────────────────────────────────────────────────────────
function CalendarTab({ calendar, setCalendar, token }) {
  const today = new Date();
  const [viewYear,  setViewYear]  = useState(today.getFullYear());
  const [viewMonth, setViewMonth] = useState(today.getMonth()); // 0-indexed
  const [selected,  setSelected]  = useState(null);   // "YYYY-MM-DD" or null
  const [showForm,  setShowForm]  = useState(false);
  const [editing,   setEditing]   = useState(null);   // event obj or null (new)
  const [saving,    setSaving]    = useState(false);
  const [syncing,   setSyncing]   = useState(false);
  const isConfigured = config.WARD_COUNCIL_SHEET_ID && !config.WARD_COUNCIL_SHEET_ID.includes("YOUR_");

  const pad = n => String(n).padStart(2, "0");
  const ymd = d => `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
  const todayStr = ymd(today);

  // ── Month grid helpers ──
  const monthName = new Date(viewYear, viewMonth, 1).toLocaleString("default", { month: "long" });
  const firstDow  = new Date(viewYear, viewMonth, 1).getDay();   // 0=Sun
  const daysInMonth = new Date(viewYear, viewMonth + 1, 0).getDate();

  const prevMonth = () => {
    if (viewMonth === 0) { setViewYear(y => y - 1); setViewMonth(11); }
    else setViewMonth(m => m - 1);
  };
  const nextMonth = () => {
    if (viewMonth === 11) { setViewYear(y => y + 1); setViewMonth(0); }
    else setViewMonth(m => m + 1);
  };

  // ── Events for this month ──
  const monthPrefix = `${viewYear}-${pad(viewMonth + 1)}`;
  const eventsForDate = (dateStr) => calendar.filter(e => e.date === dateStr);
  const monthEvents = calendar
    .filter(e => e.date.startsWith(monthPrefix))
    .sort((a, b) => (a.date > b.date ? 1 : a.date < b.date ? -1 : (a.time > b.time ? 1 : -1)));

  // ── Save to sheet ──
  const doSave = async (data) => {
    setSaving(true);
    try {
      const { pushCalendar } = await import("./sheets");
      await pushCalendar(token, data);
      notify.success("Calendar saved");
    } catch(e) {
      notify.error("Save failed: " + e.message);
    }
    setSaving(false);
  };

  const addEvent = async (ev) => {
    const newCal = [...calendar, { id: `cal_${Date.now()}`, ...ev }]
      .sort((a,b) => a.date > b.date ? 1 : a.date < b.date ? -1 : (a.time > b.time ? 1 : -1));
    setCalendar(newCal);
    await doSave(newCal);
  };

  const updateEvent = async (id, ev) => {
    const newCal = calendar.map(e => e.id === id ? { ...e, ...ev } : e);
    setCalendar(newCal);
    await doSave(newCal);
  };

  const deleteEvent = async (id) => {
    const newCal = calendar.filter(e => e.id !== id);
    setCalendar(newCal);
    await doSave(newCal);
  };

  const doRefresh = async () => {
    if (!isConfigured) return;
    setSyncing(true);
    try {
      const { pullCalendar } = await import("./sheets");
      const cp = await pullCalendar(token);
      if (cp.calendar) setCalendar(cp.calendar);
      notify.success("Calendar refreshed");
    } catch(e) {
      notify.error("Refresh failed: " + e.message);
    }
    setSyncing(false);
  };

  // Build grid cells: blanks + day numbers
  const cells = [];
  for (let i = 0; i < firstDow; i++) cells.push({ blank: true, key: `b${i}` });
  for (let d = 1; d <= daysInMonth; d++) {
    const dateStr = `${viewYear}-${pad(viewMonth+1)}-${pad(d)}`;
    cells.push({ blank: false, day: d, dateStr, key: dateStr });
  }

  const selectedEvents = selected ? eventsForDate(selected) : [];

  return (
    <div className="animate-in">
      {/* Hero */}
      <div style={{ background: `linear-gradient(130deg, ${C.blue40} 0%, ${C.blue35} 50%, ${C.blue25} 100%)`, borderRadius: 14, padding: "22px 28px", marginBottom: 24, color: "#fff", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
        <div>
          <div style={{ fontSize: 22, fontWeight: 600, letterSpacing: "-.01em", marginBottom: 2 }}>Calendar</div>
          <div style={{ fontSize: 13, opacity: .75, fontFamily: "'Helvetica Neue',Arial,sans-serif" }}>Ward events and schedule</div>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", alignItems: "center" }}>
          {isConfigured && config.WARD_COUNCIL_SHEET_URL && (
            <a href={config.WARD_COUNCIL_SHEET_URL} target="_blank" rel="noopener noreferrer"
              style={{ background: "rgba(255,255,255,.14)", border: "1.5px solid rgba(255,255,255,.3)", color: "#fff", borderRadius: 8, padding: "7px 14px", fontSize: 12, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600, textDecoration: "none", display: "flex", alignItems: "center", gap: 6 }}>
              ↗ Open Sheet
            </a>
          )}
          <button onClick={doRefresh} disabled={syncing || !isConfigured}
            style={{ background: "rgba(255,255,255,.14)", border: "1.5px solid rgba(255,255,255,.3)", color: "#fff", borderRadius: 8, padding: "7px 14px", fontSize: 12, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600, cursor: syncing || !isConfigured ? "not-allowed" : "pointer", opacity: syncing || !isConfigured ? .5 : 1, display: "flex", alignItems: "center", gap: 6 }}>
            {syncing ? <><RotateCcw size={12} style={{animation:"spin 1s linear infinite"}}/> Refreshing…</> : <><RotateCcw size={12}/> Refresh</>}
          </button>
          <button onClick={() => { setEditing(null); setShowForm(true); }}
            style={{ background: "#fff", border: "none", color: C.blue40, borderRadius: 8, padding: "7px 14px", fontSize: 12, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}>
            <Plus size={13}/> Add Event
          </button>
        </div>
      </div>

      {!isConfigured && (
        <div style={{ background: "#FEF3E2", border: `1px solid ${C.goldDeep}`, borderRadius: 10, padding: "12px 16px", marginBottom: 20, fontSize: 13, fontFamily: "'Helvetica Neue',Arial,sans-serif", color: "#7A5200" }}>
          <strong>Setup needed:</strong> Add <code>WARD_COUNCIL_SHEET_ID</code> to your <code>config.js</code> to enable calendar sync.
          Create a new Google Sheet with a tab named <strong>Calendar</strong> and columns: <code>Date | Time | Event</code>.
        </div>
      )}

      <div style={{ display: "grid", gridTemplateColumns: "1fr 320px", gap: 20, alignItems: "start" }}>
        {/* ── Month Grid ── */}
        <div style={{ background: C.surfaceWhite, borderRadius: 12, border: `1.5px solid ${C.border}`, overflow: "hidden" }}>
          {/* Month nav */}
          <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", padding: "16px 20px", borderBottom: `1px solid ${C.borderLight}` }}>
            <button onClick={prevMonth} style={{ background: "none", border: `1px solid ${C.border}`, borderRadius: 6, width: 32, height: 32, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", color: C.textSecond }}>‹</button>
            <div style={{ fontFamily: "Georgia,serif", fontSize: 17, fontWeight: 600, color: C.textPrimary }}>
              {monthName} {viewYear}
            </div>
            <button onClick={nextMonth} style={{ background: "none", border: `1px solid ${C.border}`, borderRadius: 6, width: 32, height: 32, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", color: C.textSecond }}>›</button>
          </div>

          {/* Day-of-week headers */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", borderBottom: `1px solid ${C.borderLight}` }}>
            {["Sun","Mon","Tue","Wed","Thu","Fri","Sat"].map(d => (
              <div key={d} style={{ padding: "8px 0", textAlign: "center", fontSize: 11, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 700, letterSpacing: ".05em", textTransform: "uppercase", color: C.textMuted }}>
                {d}
              </div>
            ))}
          </div>

          {/* Calendar cells */}
          <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)" }}>
            {cells.map(cell => {
              if (cell.blank) return <div key={cell.key} style={{ minHeight: 80, borderRight: `1px solid ${C.borderLight}`, borderBottom: `1px solid ${C.borderLight}`, background: C.pageBg }}/>;
              const evs = eventsForDate(cell.dateStr);
              const isToday = cell.dateStr === todayStr;
              const isSel   = cell.dateStr === selected;
              return (
                <div key={cell.key} onClick={() => setSelected(isSel ? null : cell.dateStr)}
                  style={{ minHeight: 80, padding: "6px 8px", borderRight: `1px solid ${C.borderLight}`, borderBottom: `1px solid ${C.borderLight}`, cursor: "pointer", background: isSel ? "#EAF0F6" : "transparent", transition: "background .15s" }}>
                  <div style={{ width: 24, height: 24, borderRadius: "50%", display: "flex", alignItems: "center", justifyContent: "center", marginBottom: 4,
                    background: isToday ? C.blue35 : "transparent",
                    color: isToday ? "#fff" : C.textPrimary,
                    fontSize: 13, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: isToday ? 700 : 400 }}>
                    {cell.day}
                  </div>
                  {evs.slice(0, 3).map((ev, i) => (
                    <div key={ev.id} style={{ fontSize: 10, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600, background: C.blue25, color: "#fff", borderRadius: 3, padding: "1px 5px", marginBottom: 2, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>
                      {ev.time ? `${ev.time} ` : ""}{ev.event}
                    </div>
                  ))}
                  {evs.length > 3 && (
                    <div style={{ fontSize: 10, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif" }}>+{evs.length - 3} more</div>
                  )}
                </div>
              );
            })}
          </div>
        </div>

        {/* ── Right panel: selected day or month list ── */}
        <div>
          {selected ? (
            <div style={{ background: C.surfaceWhite, borderRadius: 12, border: `1.5px solid ${C.border}`, overflow: "hidden" }}>
              <div style={{ padding: "14px 16px", borderBottom: `1px solid ${C.borderLight}`, display: "flex", alignItems: "center", justifyContent: "space-between" }}>
                <div style={{ fontFamily: "Georgia,serif", fontSize: 15, fontWeight: 600, color: C.textPrimary }}>
                  {new Date(selected + "T12:00:00").toLocaleDateString("default", { weekday: "long", month: "long", day: "numeric" })}
                </div>
                <button onClick={() => { setEditing({ date: selected, time: "", event: "" }); setShowForm(true); }}
                  style={{ background: C.blue35, border: "none", color: "#fff", borderRadius: 6, padding: "5px 12px", fontSize: 12, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600, cursor: "pointer", display: "flex", alignItems: "center", gap: 4 }}>
                  <Plus size={11}/> Add
                </button>
              </div>
              {selectedEvents.length === 0 ? (
                <div style={{ padding: "20px 16px", fontSize: 13, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontStyle: "italic", textAlign: "center" }}>No events this day</div>
              ) : (
                selectedEvents.map(ev => (
                  <div key={ev.id} style={{ padding: "12px 16px", borderBottom: `1px solid ${C.borderLight}`, display: "flex", alignItems: "flex-start", gap: 10 }}>
                    <div style={{ flex: 1 }}>
                      {ev.time && <div style={{ fontSize: 11, fontFamily: "'Helvetica Neue',Arial,sans-serif", color: C.blue25, fontWeight: 700, marginBottom: 2 }}>{ev.time}</div>}
                      <div style={{ fontSize: 14, fontFamily: "Georgia,serif", color: C.textPrimary }}>{ev.event}</div>
                    </div>
                    <div style={{ display: "flex", gap: 4, flexShrink: 0 }}>
                      <button onClick={() => { setEditing(ev); setShowForm(true); }}
                        style={{ background: "none", border: `1px solid ${C.borderLight}`, borderRadius: 5, padding: "3px 8px", fontSize: 11, color: C.textSecond, cursor: "pointer", fontFamily: "'Helvetica Neue',Arial,sans-serif" }}>Edit</button>
                      <button onClick={() => deleteEvent(ev.id)}
                        style={{ background: "none", border: `1px solid ${C.borderLight}`, borderRadius: 5, padding: "3px 8px", fontSize: 11, color: C.red15, cursor: "pointer", fontFamily: "'Helvetica Neue',Arial,sans-serif" }}style={{display:"flex",alignItems:"center"}}><X size={11}/></button>
                    </div>
                  </div>
                ))
              )}
            </div>
          ) : (
            /* Month event list */
            <div style={{ background: C.surfaceWhite, borderRadius: 12, border: `1.5px solid ${C.border}`, overflow: "hidden" }}>
              <div style={{ padding: "14px 16px", borderBottom: `1px solid ${C.borderLight}` }}>
                <div style={{ fontFamily: "Georgia,serif", fontSize: 15, fontWeight: 600, color: C.textPrimary }}>
                  {monthName} Events
                </div>
                <div style={{ fontSize: 12, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", marginTop: 2 }}>
                  {monthEvents.length} event{monthEvents.length !== 1 ? "s" : ""}
                </div>
              </div>
              {monthEvents.length === 0 ? (
                <div style={{ padding: "24px 16px", fontSize: 13, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontStyle: "italic", textAlign: "center" }}>No events this month</div>
              ) : (
                monthEvents.map(ev => (
                  <div key={ev.id} onClick={() => setSelected(ev.date)}
                    style={{ padding: "10px 16px", borderBottom: `1px solid ${C.borderLight}`, cursor: "pointer", display: "flex", alignItems: "flex-start", gap: 10 }}>
                    <div style={{ minWidth: 36, textAlign: "center" }}>
                      <div style={{ fontSize: 18, fontWeight: 700, fontFamily: "Georgia,serif", color: C.blue35, lineHeight: 1 }}>
                        {parseInt(ev.date.slice(8))}
                      </div>
                      <div style={{ fontSize: 10, fontFamily: "'Helvetica Neue',Arial,sans-serif", color: C.textMuted, textTransform: "uppercase" }}>
                        {new Date(ev.date + "T12:00:00").toLocaleDateString("default", { weekday: "short" })}
                      </div>
                    </div>
                    <div style={{ flex: 1, paddingTop: 2 }}>
                      {ev.time && <div style={{ fontSize: 11, fontFamily: "'Helvetica Neue',Arial,sans-serif", color: C.blue25, fontWeight: 700, marginBottom: 1 }}>{ev.time}</div>}
                      <div style={{ fontSize: 13, fontFamily: "Georgia,serif", color: C.textPrimary }}>{ev.event}</div>
                    </div>
                  </div>
                ))
              )}
            </div>
          )}
        </div>
      </div>

      {/* Add/Edit Modal */}
      {showForm && (
        <CalendarEventModal
          event={editing}
          saving={saving}
          onSave={async (ev) => {
            if (editing && editing.id) await updateEvent(editing.id, ev);
            else await addEvent(ev);
            setShowForm(false); setEditing(null);
          }}
          onClose={() => { setShowForm(false); setEditing(null); }}
        />
      )}
    </div>
  );
}

function CalendarEventModal({ event, saving, onSave, onClose }) {
  const isNew = !event?.id;
  const pad = n => String(n).padStart(2, "0");
  const todayStr = (() => { const d = new Date(); return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`; })();
  const [f, setF] = useState({
    date:  event?.date  || todayStr,
    time:  event?.time  || "",
    event: event?.event || "",
  });
  const s = (k, v) => setF(p => ({ ...p, [k]: v }));

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.45)", zIndex: 200, display: "flex", alignItems: "center", justifyContent: "center", padding: 24 }}>
      <div style={{ background: "#fff", borderRadius: 14, padding: "28px 32px", width: "100%", maxWidth: 420, boxShadow: "0 12px 48px rgba(0,0,0,.22)" }}>
        <div style={{ fontFamily: "Georgia,serif", fontSize: 18, fontWeight: 600, marginBottom: 20, color: C.textPrimary }}>
          {isNew ? "Add Event" : "Edit Event"}
        </div>
        <div style={{ marginBottom: 14 }}>
          <label style={{ display: "block", fontSize: 11, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 700, letterSpacing: ".05em", textTransform: "uppercase", color: C.textMuted, marginBottom: 5 }}>Date</label>
          <input type="date" value={f.date} onChange={e => s("date", e.target.value)}
            style={{ width: "100%", padding: "8px 12px", fontSize: 14, borderRadius: 7, border: `1.5px solid ${C.border}`, fontFamily: "'Helvetica Neue',Arial,sans-serif", boxSizing: "border-box" }}/>
        </div>
        <div style={{ marginBottom: 14 }}>
          <label style={{ display: "block", fontSize: 11, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 700, letterSpacing: ".05em", textTransform: "uppercase", color: C.textMuted, marginBottom: 5 }}>Time <span style={{ fontWeight: 400, textTransform: "none" }}>(optional)</span></label>
          <input type="time" value={f.time} onChange={e => s("time", e.target.value)}
            style={{ width: "100%", padding: "8px 12px", fontSize: 14, borderRadius: 7, border: `1.5px solid ${C.border}`, fontFamily: "'Helvetica Neue',Arial,sans-serif", boxSizing: "border-box" }}/>
        </div>
        <div style={{ marginBottom: 22 }}>
          <label style={{ display: "block", fontSize: 11, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 700, letterSpacing: ".05em", textTransform: "uppercase", color: C.textMuted, marginBottom: 5 }}>Event</label>
          <input type="text" value={f.event} onChange={e => s("event", e.target.value)}
            placeholder="e.g. Ward Temple Night"
            autoFocus
            style={{ width: "100%", padding: "8px 12px", fontSize: 14, borderRadius: 7, border: `1.5px solid ${C.border}`, fontFamily: "'Helvetica Neue',Arial,sans-serif", boxSizing: "border-box" }}/>
        </div>
        <div style={{ display: "flex", gap: 10, justifyContent: "flex-end" }}>
          <button onClick={onClose} className="btn-secondary" style={{ padding: "8px 18px", fontSize: 13 }}>Cancel</button>
          <button onClick={() => { if (f.event.trim() && f.date) onSave(f); }}
            disabled={saving || !f.event.trim() || !f.date}
            className="btn-primary" style={{ padding: "8px 18px", fontSize: 13 }}>
            {saving ? "Saving…" : isNew ? "Add Event" : "Save Changes"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ─── Ward Council View (placeholder — future tab) ─────────────────────────────
const WC_TEMPLATE = [
  { itemKey: "prayer_list",      section: "Opening",   label: "Review Prayer List",      hasAssignee: false, isStatic: true,  isPrayerList: true },
  { itemKey: "opening_song",     section: "Opening",   label: "Opening Song",            hasAssignee: true,  isStatic: false, hasSongField: true },
  { itemKey: "opening_prayer",   section: "Opening",   label: "Opening Prayer",          hasAssignee: true,  isStatic: true },
  { itemKey: "spirit_thought",   section: "Opening",   label: "Spiritual Thought",       hasAssignee: true,  isStatic: true,  alternates: true },
  { itemKey: "outstanding_tasks",section: "Business",  label: "Review Outstanding Tasks",hasAssignee: false, isStatic: true,  isLive: "tasks" },
  { itemKey: "discussion",       section: "Business",  label: "Discussion Topics",       hasAssignee: false, isStatic: false, isList: true },
  { itemKey: "review_calendar",  section: "Business",  label: "Review Calendar",          hasAssignee: false, isStatic: true,  isLive: "calendar" },
  { itemKey: "closing_prayer",   section: "Closing",   label: "Closing Prayer",          hasAssignee: true,  isStatic: true },
];

function WardCouncilTab({ wardCouncilMeeting, setWardCouncilMeeting, calendar=[], roster=[], token="", onNavigate, isAdmin=false }) {
  const today = localDateStr(new Date());
  const nextSundayDate = getUpcomingSunday();

  const [selectedDate, setSelectedDate] = useState(nextSundayDate);
  const [wcData, setWcData] = useState(wardCouncilMeeting);
  const [showOtherDate, setShowOtherDate] = useState(false);
  const [slackDraft, setSlackDraft] = useState(null); // {firstLine, bodyLines, relayURL, channelKey, channelName}
  const [sending, setSending] = useState(false);
  const [showPrayerList, setShowPrayerList] = useState(false);
  const [prayerListData, setPrayerListData] = useState(null);
  const [prayerListLoading, setPrayerListLoading] = useState(false);
  const tokenRef = useRef(token);
  useEffect(() => { tokenRef.current = token; }, [token]);
  useEffect(() => { setWcData(wardCouncilMeeting); }, [wardCouncilMeeting]);

  const sundays = getSurroundingSundays();
  const week = isoWeek(selectedDate);
  const isEvenWeek = week % 2 === 0;

  const dateItems = wcData.filter(r => r.date === selectedDate);
  const hasData   = dateItems.length > 0;

  // Upcoming calendar events for review_calendar live badge
  const upcomingCalendarEvents = (() => {
    if (!selectedDate || !calendar.length) return [];
    const base = new Date(selectedDate + "T12:00:00");
    const limit = new Date(base); limit.setDate(limit.getDate() + 7);
    return calendar
      .filter(ev => { const d = new Date(ev.date + "T12:00:00"); return d > base && d <= limit; })
      .sort((a, b) => a.date > b.date ? 1 : a.date < b.date ? -1 : 0);
  })();

  // All roles pool (bishopric + ward_council) for assignment
  const ALL_ASSIGNABLE_KEYS = ["opening_prayer", "opening_song", "spirit_thought", "closing_prayer"];
  const fullRosterPool = ROSTER_POSITIONS.map(p => rosterName(roster, p.role)).filter(Boolean);

  const suggestAssignee = (itemKey, exclude = new Set()) => {
    if (!fullRosterPool.length) return "";
    const pool = fullRosterPool.filter(n => !exclude.has(n));
    if (!pool.length) return "";
    const pastRows = wcData
      .filter(r => r.itemKey === itemKey && r.assignee && r.date !== selectedDate)
      .sort((a, b) => (a.date > b.date ? -1 : 1));
    const counts = {};
    pool.forEach(name => { counts[name] = 0; });
    pastRows.forEach(r => { if (counts[r.assignee] !== undefined) counts[r.assignee]++; });
    const minCount = Math.min(...pool.map(n => counts[n]));
    const tied = pool.filter(n => counts[n] === minCount);
    const lastAssigned = {};
    tied.forEach(name => {
      const lastRow = pastRows.find(r => r.assignee === name);
      lastAssigned[name] = lastRow ? lastRow.date : "0000-00-00";
    });
    tied.sort((a, b) => (lastAssigned[a] < lastAssigned[b] ? -1 : 1));
    return tied[0];
  };

  const doAutoAssign = () => {
    const usedThisMeeting = new Set(
      ALL_ASSIGNABLE_KEYS.map(k => getItem(k)?.assignee).filter(Boolean)
    );
    ALL_ASSIGNABLE_KEYS.forEach(key => {
      if (!getItem(key)?.assignee) {
        const suggestion = suggestAssignee(key, usedThisMeeting);
        if (suggestion) {
          updateItem(key, "assignee", suggestion);
          usedThisMeeting.add(suggestion);
        }
      }
    });
    notify.success("Auto-assigned open slots based on rotation balance");
  };

  const doAutoAssignOne = (itemKey) => {
    const usedElsewhere = new Set(
      ALL_ASSIGNABLE_KEYS
        .filter(k => k !== itemKey)
        .map(k => getItem(k)?.assignee)
        .filter(Boolean)
    );
    const suggestion = suggestAssignee(itemKey, usedElsewhere);
    if (suggestion) updateItem(itemKey, "assignee", suggestion);
  };

  const getItem = (key) => dateItems.find(r => r.itemKey === key) || null;

  const spiritLabel = (() => {
    const override = getItem("spirit_thought")?.spiritToggle;
    if (override) return override;
    return isEvenWeek ? "handbook_review" : "spiritual_thought";
  })();

  const openPrayerList = async () => {
    setShowPrayerList(true);
    if (prayerListData !== null) return;
    const plid = config.PRAYER_LIST_SHEET_ID;
    if (!plid || plid.includes("YOUR_")) { setPrayerListData([]); return; }
    setPrayerListLoading(true);
    try {
      const { pullPrayerList } = await import("./sheets");
      const data = await pullPrayerList(tokenRef.current);
      setPrayerListData(data);
    } catch(e) {
      setPrayerListData([]);
      notify.error("Could not load prayer list: " + e.message);
    } finally { setPrayerListLoading(false); }
  };

  // Data ref — prevents stale closures in useMeetingSync
  const wcDataRef = useRef(wcData);
  useEffect(() => { wcDataRef.current = wcData; }, [wcData]);

  // ── Meeting sync ──
  const wcDiff = useCallback((local, remote) => {
    if (!remote || !Array.isArray(remote)) return 0;
    const localDate  = local.filter(r => r.date === selectedDate);
    const remoteDate = remote.filter(r => r.date === selectedDate);
    if (localDate.length !== remoteDate.length) return Math.abs(localDate.length - remoteDate.length);
    return localDate.filter((r, i) => {
      const rem = remoteDate[i];
      return !rem || r.assignee !== rem.assignee || r.done !== rem.done || r.notes !== rem.notes;
    }).length;
  }, [selectedDate]);

  const { markDirty: wcMarkDirty, doSave, saveStatus: wcSaveStatus,
          pendingCount: wcPendingCount, applyPending: wcApplyPending } = useMeetingSync({
    getData:  () => wcDataRef.current,
    saveFn:   async (tok, data) => {
      if (!tok) return; // skip if token not yet available
      const { pushWardCouncilMeeting } = await import("./sheets");
      await pushWardCouncilMeeting(tok, data);
    },
    pullFn:   async (tok) => {
      const { pullWardCouncilMeeting } = await import("./sheets");
      const d = await pullWardCouncilMeeting(tok);
      return d.wardCouncilMeeting || [];
    },
    diffFn:   wcDiff,
    tokenRef,
    onApply:  (remote) => { setWcData(remote); setWardCouncilMeeting(remote); },
    enabled:  true,
  });

  const createProgram = () => {
    const priorDates = [...new Set(wcData.map(r => r.date))]
      .filter(d => d < selectedDate).sort().reverse();
    const priorDate = priorDates[0] || null;

    const priorTopicRows = priorDate
      ? wcData.filter(r => r.date === priorDate && r.itemKey.startsWith("topic_") && !r.done)
      : [];
    const carryTopicRows = priorTopicRows.map(r => ({
      ...r, id: `wc_new_topic_carry_${Date.now()}_${r.id}`, date: selectedDate,
    }));

    const priorTaskRows = priorDate
      ? wcData.filter(r => r.date === priorDate && r.itemKey.startsWith("task_") && !r.done)
      : [];
    const carryTaskRows = priorTaskRows.map(r => ({
      ...r, id: `wc_new_task_carry_${Date.now()}_${r.id}`, date: selectedDate,
    }));

    const rows = WC_TEMPLATE.map(t => ({
      id: `wc_new_${t.itemKey}`,
      date: selectedDate,
      itemKey: t.itemKey,
      assignee: "",
      done: false,
      notes: "",
      customLabel: "",
      spiritToggle: t.itemKey === "spirit_thought" ? (isEvenWeek ? "handbook_review" : "spiritual_thought") : "",
    }));
    const newData = [...wcData.filter(r => r.date !== selectedDate), ...rows, ...carryTopicRows, ...carryTaskRows];
    setWcData(newData);
    setWardCouncilMeeting(newData);
    if (carryTopicRows.length || carryTaskRows.length) {
      notify.info(`Carried over ${carryTopicRows.length} topic${carryTopicRows.length!==1?"s":""} and ${carryTaskRows.length} task${carryTaskRows.length!==1?"s":""} from last week`);
    }
  };

  const updateItem = (itemKey, field, value) => {
    setWcData(prev => {
      const exists = prev.find(r => r.date === selectedDate && r.itemKey === itemKey);
      let next;
      if (exists) {
        next = prev.map(r => r.date === selectedDate && r.itemKey === itemKey ? { ...r, [field]: value } : r);
      } else {
        next = [...prev, { id: `wc_${itemKey}_${Date.now()}`, date: selectedDate, itemKey, assignee: "", done: false, notes: "", customLabel: "", spiritToggle: "", [field]: value }];
      }
      setWardCouncilMeeting(next);
      return next;
    });
    wcMarkDirty(`${selectedDate}|${itemKey}`);
  };

  // Discussion topics
  // Discussion topics — each is its own row with itemKey "topic_<id>"
  const discussionTopics = dateItems
    .filter(r => r.itemKey.startsWith("topic_"))
    .sort((a, b) => a.id < b.id ? -1 : 1);

  const addTopic = (text) => {
    if (!text.trim()) return;
    const id = `topic_${Date.now()}`;
    const newRow = { id: `wc_${id}`, date: selectedDate, itemKey: id, assignee: "", done: false, notes: text.trim(), customLabel: "", spiritToggle: "" };
    const newData = [...wcData, newRow];
    setWcData(newData); setWardCouncilMeeting(newData);
    wcMarkDirty(`${selectedDate}|${id}`);
  };
  const toggleTopic = (id) => {
    updateItem(id, "done", !dateItems.find(r => r.itemKey === id)?.done);
  };

  const removeTopic = (id) => {
    const newData = wcData.filter(r => !(r.date === selectedDate && r.itemKey === id));
    setWcData(newData); setWardCouncilMeeting(newData);
    wcMarkDirty(`${selectedDate}|${id}`);
  };

  const toggleTask = (id) => {
    updateItem(id, "done", !dateItems.find(r => r.itemKey === id)?.done);
  };

  const removeTask = (id) => {
    const newData = wcData.filter(r => !(r.date === selectedDate && r.itemKey === id));
    setWcData(newData); setWardCouncilMeeting(newData);
    wcMarkDirty(`${selectedDate}|${id}`);
  };

  const updateTaskField = (itemKey, field, value) => {
    const col = field === "text" ? "notes" : field;
    updateItem(itemKey, col, value);
  };

  // Outstanding tasks — each is its own row with itemKey "task_<id>"
  const outstandingTasks = dateItems
    .filter(r => r.itemKey.startsWith("task_"))
    .sort((a, b) => a.id < b.id ? -1 : 1);

  const addTask = (text, assignee) => {
    if (!text.trim()) return;
    const id = `task_${Date.now()}`;
    const newRow = { id: `wc_${id}`, date: selectedDate, itemKey: id, assignee: assignee || "", done: false, notes: text.trim(), customLabel: "", spiritToggle: "" };
    const newData = [...wcData, newRow];
    setWcData(newData); setWardCouncilMeeting(newData);
    wcMarkDirty(`${selectedDate}|${id}`);
  };

  // Build and open Slack preview modal
  const doSendSlack = () => {
    const relayURL = config.SLACK_RELAY_URL || "";
    const webhookURL = config.SLACK_WEBHOOK_WARD_COUNCIL || "";
    const channelName = config.SLACK_WEBHOOK_WARD_COUNCIL_NAME || "ward-council";
    if (!relayURL || relayURL.includes("YOUR_")) { notify.error("No relay URL configured in config.js"); return; }
    if (!webhookURL || webhookURL.includes("YOUR/")) { notify.error("No Ward Council Slack webhook configured in config.js"); return; }

    const sLabel = spiritLabel === "handbook_review" ? "Handbook Review" : "Spiritual Thought";
    const songRow = getItem("opening_song");
    const assign = (label, value) => (value && value !== "_unassigned_") ? `${label}: ${value}` : null;
    const assignments = [
      assign("Opening Song",  songRow?.assignee),
      assign("Opening Prayer", getItem("opening_prayer")?.assignee),
      assign(sLabel,           getItem("spirit_thought")?.assignee),
      assign("Closing Prayer", getItem("closing_prayer")?.assignee),
    ].filter(Boolean);

    const divider = "─────────────────────────────";
    const bodyLines = [
      divider,
      assignments.length ? assignments.join("\n") : "_No assignments set_",
      divider,
      `_Ward Manager · ${config.WARD_NAME}_`,
    ];
    setSlackDraft({ firstLine: `*Ward Council — ${toDisplayDate(selectedDate)}*`, bodyLines, relayURL, channelKey: webhookURL, channelName });
  };

  const doConfirmSlack = async (firstLine) => {
    if (!slackDraft) return;
    const { bodyLines, relayURL, channelKey, channelName } = slackDraft;
    const text = [firstLine, ...bodyLines].join("\n");
    setSending(true);
    try {
      const res = await fetch(relayURL, {
        method: "POST", headers: { "Content-Type": "text/plain" },
        body: JSON.stringify({ channel: channelKey, text }),
      });
      const data = await res.json().catch(() => ({ status: res.status }));
      if (res.ok && data.status === 200) {
        notify.success(`Agenda sent to #${channelName}`);
        setSlackDraft(null);
      } else {
        notify.error(`Relay error: ${data.message || res.status}`);
      }
    } catch (e) {
      notify.error("Failed to reach relay: " + e.message);
    } finally { setSending(false); }
  };

  return (
    <div className="animate-in">
      {/* Hero */}
      <div style={{ background: `linear-gradient(130deg, ${C.blue40} 0%, ${C.blue35} 50%, ${C.blue25} 100%)`, borderRadius: 14, padding: "22px 28px", marginBottom: 24, color: "#fff", display: "flex", alignItems: "center", justifyContent: "space-between", flexWrap: "wrap", gap: 12 }}>
        <div>
          <div style={{ fontSize: 22, fontWeight: 600, letterSpacing: "-.01em", marginBottom: 2 }}>Ward Council</div>
          <div style={{ fontSize: 13, opacity: .75, fontFamily: "'Helvetica Neue',Arial,sans-serif" }}>{toDisplayDate(selectedDate)}</div>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          {isAdmin && (<>
          <button onClick={doSendSlack} disabled={slackDraft !== null || !hasData}
            style={{ background: "rgba(255,255,255,.14)", border: "1.5px solid rgba(255,255,255,.3)", color: "#fff", borderRadius: 8, padding: "7px 14px", fontSize: 12, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600, cursor: !hasData ? "not-allowed" : "pointer", opacity: !hasData ? .5 : 1, display: "flex", alignItems: "center", gap: 6 }}>
            <Bell size={13}/> Send to Slack
          </button>
          <button onClick={doAutoAssign} disabled={!hasData}
            title="Fill unassigned slots using balanced rotation"
            style={{ background: "rgba(255,255,255,.14)", border: "1.5px solid rgba(255,255,255,.3)", color: "#fff", borderRadius: 8, padding: "7px 14px", fontSize: 12, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600, cursor: !hasData ? "not-allowed" : "pointer", opacity: !hasData ? .5 : 1, display: "flex", alignItems: "center", gap: 6 }}>
            <RotateCcw size={12}/> Auto-assign
          </button>
          </>)}
          <button onClick={doSave} disabled={wcSaveStatus === "saving" || !hasData}
            style={{ background: "rgba(255,255,255,.18)", border: "1.5px solid rgba(255,255,255,.4)", color: "#fff", borderRadius: 8, padding: "7px 14px", fontSize: 12, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600, cursor: wcSaveStatus === "saving" || !hasData ? "not-allowed" : "pointer", opacity: wcSaveStatus === "saving" || !hasData ? .5 : 1, display: "flex", alignItems: "center", gap: 6 }}>
            <Save size={13}/> {wcSaveStatus === "saving" ? "Saving…" : "Save"}
          </button>
        </div>
      </div>

      {/* Sync status bar */}
      <MeetingSyncBar saveStatus={wcSaveStatus} pendingCount={wcPendingCount}
        onApply={wcApplyPending} onSave={doSave}/>

      {/* Date nav */}
      <div style={{ display: "flex", gap: 6, marginBottom: 24, flexWrap: "wrap", alignItems: "center" }}>
        {sundays.map(s => {
          const isNext = s === nextSundayDate;
          const isSel  = s === selectedDate;
          return (
            <button key={s} onClick={() => setSelectedDate(s)}
              style={{ padding: "6px 14px", borderRadius: 20, border: `1.5px solid ${isSel ? C.blue35 : C.border}`, background: isSel ? C.blue35 : C.surfaceWhite, color: isSel ? "#fff" : C.textMuted, fontSize: 12, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: isSel ? 700 : 400, cursor: "pointer", display: "flex", alignItems: "center", gap: 6 }}>
              {toDisplayDate(s)}
              {isNext && <span style={{ fontSize: 9, fontWeight: 700, letterSpacing: ".06em", background: C.gold20, color: "#fff", borderRadius: 8, padding: "1px 6px" }}>NEXT</span>}
            </button>
          );
        })}
        <button onClick={() => setShowOtherDate(v => !v)}
          style={{ padding: "6px 14px", borderRadius: 20, border: `1.5px solid ${C.border}`, background: C.surfaceWhite, color: C.textMuted, fontSize: 12, fontFamily: "'Helvetica Neue',Arial,sans-serif", cursor: "pointer" }}>
          Other…
        </button>
        {showOtherDate && (
          <input type="date" value={selectedDate} onChange={e => { setSelectedDate(e.target.value); setShowOtherDate(false); }}
            style={{ width: 160, padding: "6px 10px", fontSize: 12 }}/>
        )}
        <div style={{ fontSize: 11, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", marginLeft: 6 }}>
          Week {week} · {spiritLabel === "handbook_review" ? "Handbook Review week" : "Spiritual Thought week"}
        </div>
      </div>

      {/* Empty state */}
      {!hasData && (
        <div style={{ textAlign: "center", padding: "60px 24px", background: C.surfaceWhite, borderRadius: 12, border: `1.5px solid ${C.border}` }}>
          <div style={{ marginBottom: 12, opacity: .3, color: C.textMuted, display: "flex", justifyContent: "center" }}><ClipboardList size={42}/></div>
          <div style={{ fontFamily: "Georgia,serif", fontSize: 17, color: C.textSecond, marginBottom: 8 }}>No agenda for this Sunday</div>
          <div style={{ fontSize: 12, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", marginBottom: 20 }}>Create the agenda from the standard template to get started</div>
          <button className="btn-primary" onClick={createProgram}><Plus size={13}/> Create Agenda</button>
        </div>
      )}

      {/* Agenda sections */}
      {hasData && BM_SECTIONS.map(section => {
        const sStyle = BM_SECTION_STYLE[section];
        const items  = WC_TEMPLATE.filter(t => t.section === section);
        if (!items.length) return null;
        return (
          <div key={section} style={{ marginBottom: 20, background: C.surfaceWhite, border: `1.5px solid ${C.border}`, borderRadius: 12, overflow: "hidden" }}>
            <div style={{ background: sStyle.bg, borderBottom: `1.5px solid ${sStyle.border}`, padding: "10px 20px", display: "flex", alignItems: "center", gap: 8 }}>
              <div style={{ width: 8, height: 8, borderRadius: "50%", background: sStyle.color, flexShrink: 0 }}/>
              <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: sStyle.color, fontFamily: "'Helvetica Neue',Arial,sans-serif" }}>{section}</span>
            </div>
            <div style={{ display: "flex", flexDirection: "column" }}>
              {items.map((tmpl, idx) => {
                const row = getItem(tmpl.itemKey);
                const isLast = idx === items.length - 1;
                const displayLabel = tmpl.alternates
                  ? (spiritLabel === "handbook_review" ? "Handbook Review" : "Spiritual Thought")
                  : tmpl.label;
                return (
                  <div key={tmpl.itemKey} style={{ padding: "14px 20px", borderBottom: isLast ? "none" : `1px solid ${C.borderLight}`, display: "flex", alignItems: "flex-start", gap: 16 }}>
                    <button onClick={() => updateItem(tmpl.itemKey, "done", !(row?.done))}
                      style={{ width: 20, height: 20, borderRadius: 4, border: `2px solid ${row?.done ? C.green25 : C.border}`, background: row?.done ? C.green25 : "transparent", cursor: "pointer", flexShrink: 0, marginTop: 2, display: "flex", alignItems: "center", justifyContent: "center" }}>
                      {row?.done && <svg width="10" height="10" viewBox="0 0 10 10" fill="none"><path d="M1.5 5L4 7.5L8.5 2.5" stroke="white" strokeWidth="1.5" strokeLinecap="round" strokeLinejoin="round"/></svg>}
                    </button>
                    <div style={{ flex: 1, minWidth: 0 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                        <span style={{ fontFamily: "Georgia,serif", fontSize: 15, color: row?.done ? C.textMuted : C.textPrimary, textDecoration: row?.done ? "line-through" : "none", minWidth: 200 }}>
                          {displayLabel}
                        </span>
                        {/* Assignee — editable for admin, read-only for ward council */}
                        {tmpl.hasAssignee && (
                          isAdmin ? (
                            <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
                              <select value={row?.assignee || ""} onChange={e => updateItem(tmpl.itemKey, "assignee", e.target.value)}
                                style={{ width: 200, padding: "5px 10px", fontSize: 13, background: row?.assignee ? C.surfaceWarm : "#fff", border: `1.5px solid ${row?.assignee ? C.border : C.borderLight}` }}>
                                <option value="">— Assign to… —</option>
                                {ROSTER_POSITIONS.map(p => (
                                  <option key={p.role} value={rosterName(roster, p.role)}>
                                    {rosterName(roster, p.role) || p.role}
                                  </option>
                                ))}
                              </select>
                              {ALL_ASSIGNABLE_KEYS.includes(tmpl.itemKey) && (
                                <button onClick={() => doAutoAssignOne(tmpl.itemKey)}
                                  title={row?.assignee ? "Re-suggest based on rotation" : "Suggest based on rotation"}
                                  style={{ width: 28, height: 28, borderRadius: 6, border: `1.5px solid ${C.borderLight}`, background: C.surfaceWarm, color: C.textMuted, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                                  <RotateCcw size={13}/>
                                </button>
                              )}
                            </div>
                          ) : (
                            row?.assignee
                              ? <span style={{ fontSize: 13, fontFamily: "'Helvetica Neue',Arial,sans-serif", color: C.textSecond, padding: "5px 0" }}>{row.assignee}</span>
                              : <span style={{ fontSize: 13, fontFamily: "'Helvetica Neue',Arial,sans-serif", color: C.textLight, fontStyle: "italic" }}>Unassigned</span>
                          )
                        )}
                        {/* Song field */}
                        {tmpl.hasSongField && (
                          <>
                            <input type="text" placeholder="Hymn #" value={row?.notes || ""} onChange={e => updateItem(tmpl.itemKey, "notes", e.target.value)}
                              style={{ width: 72, padding: "5px 10px", fontSize: 13, background: row?.notes ? C.surfaceWarm : "#fff", border: `1.5px solid ${row?.notes ? C.border : C.borderLight}` }}/>
                            <input type="text" placeholder="Song title…" value={row?.customLabel || ""} onChange={e => updateItem(tmpl.itemKey, "customLabel", e.target.value)}
                              style={{ width: 200, padding: "5px 10px", fontSize: 13, background: row?.customLabel ? C.surfaceWarm : "#fff", border: `1.5px solid ${row?.customLabel ? C.border : C.borderLight}` }}/>
                          </>
                        )}
                        {/* Alternates toggle — admin only */}
                        {tmpl.alternates && isAdmin && (
                          <button onClick={() => updateItem(tmpl.itemKey, "spiritToggle", spiritLabel === "handbook_review" ? "spiritual_thought" : "handbook_review")}
                            style={{ fontSize: 11, padding: "3px 10px", borderRadius: 12, border: `1px solid ${C.border}`, background: C.surfaceWarm, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", cursor: "pointer" }}>
                            ⇄ Switch
                          </button>
                        )}
                      </div>

                      {/* Prayer List */}
                      {tmpl.isPrayerList && (
                        <div style={{ marginTop: 8 }}>
                          <button onClick={openPrayerList} className="btn-secondary"
                            style={{ fontSize: 12, padding: "5px 14px", display: "inline-flex", alignItems: "center", gap: 6 }}>
                            <BookOpen size={13}/> View Prayer List
                          </button>
                        </div>
                      )}

                      {/* Live badge: Outstanding Tasks */}
                      {tmpl.isLive === "tasks" && (
                        <div style={{ marginTop: 10 }}>
                          {outstandingTasks.length === 0 ? (
                            <span style={{ fontSize: 11, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontStyle: "italic" }}>No outstanding tasks</span>
                          ) : (
                            <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                              {outstandingTasks.map(task => (
                                <div key={task.itemKey} style={{ display: "flex", alignItems: "center", gap: 8 }}>
                                  <button onClick={() => toggleTask(task.itemKey)}
                                    style={{ width: 16, height: 16, borderRadius: 3, border: `2px solid ${task.done ? C.green25 : C.border}`, background: task.done ? C.green25 : "transparent", cursor: "pointer", flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center" }}>
                                    {task.done && <svg width="8" height="8" viewBox="0 0 10 10" fill="none"><path d="M1.5 5L4 7.5L8.5 2.5" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></svg>}
                                  </button>
                                  <input value={task.notes} onChange={e => updateTaskField(task.itemKey, "text", e.target.value)}
                                    style={{ flex: 1, border: "none", outline: "none", fontSize: 13, fontFamily: "Georgia,serif", background: "transparent", color: task.done ? C.textMuted : C.textPrimary, textDecoration: task.done ? "line-through" : "none", minWidth: 0 }}/>
                                  <select value={task.assignee || ""} onChange={e => updateTaskField(task.itemKey, "assignee", e.target.value)}
                                    style={{ width: 160, fontSize: 11, padding: "3px 8px", border: `1px solid ${C.borderLight}`, background: C.surfaceWarm, borderRadius: 6, fontFamily: "'Helvetica Neue',Arial,sans-serif" }}>
                                    <option value="">Unassigned</option>
                                    {ROSTER_POSITIONS.map(p => (
                                      <option key={p.role} value={rosterName(roster, p.role)}>
                                        {rosterName(roster, p.role) || p.role}
                                      </option>
                                    ))}
                                  </select>
                                  <button onClick={() => removeTask(task.itemKey)} style={{ background: "none", border: "none", cursor: "pointer", color: C.textLight, display: "flex", alignItems: "center" }}><X size={13}/></button>
                                </div>
                              ))}
                            </div>
                          )}
                          <AddTaskRow onAdd={addTask} roster={roster} allRoles={true}/>
                        </div>
                      )}

                      {/* Live badge: Review Calendar */}
                      {tmpl.isLive === "calendar" && (
                        <div style={{ marginTop: 10 }}>
                          {upcomingCalendarEvents.length === 0 ? (
                            <span style={{ fontSize: 11, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontStyle: "italic" }}>No events in the next 7 days</span>
                          ) : (
                            <div style={{ display: "flex", flexDirection: "column", gap: 5 }}>
                              {upcomingCalendarEvents.map(ev => {
                                const evDate = new Date(ev.date + "T12:00:00");
                                const dayLabel = evDate.toLocaleDateString("default", { weekday: "short", month: "short", day: "numeric" });
                                return (
                                  <div key={ev.id} style={{ display: "flex", alignItems: "baseline", gap: 8 }}>
                                    <span style={{ fontSize: 11, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 700, color: C.blue25, flexShrink: 0, minWidth: 90 }}>{dayLabel}</span>
                                    {ev.time && <span style={{ fontSize: 11, fontFamily: "'Helvetica Neue',Arial,sans-serif", color: C.textMuted, flexShrink: 0 }}>{ev.time}</span>}
                                    <span style={{ fontSize: 13, fontFamily: "Georgia,serif", color: C.textPrimary }}>{ev.event}</span>
                                  </div>
                                );
                              })}
                            </div>
                          )}
                          {onNavigate && (
                            <button onClick={() => onNavigate("calendar")}
                              style={{ marginTop: 8, background: "none", border: "none", padding: 0, cursor: "pointer", fontSize: 11, color: C.blue25, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontWeight: 600, textDecoration: "underline" }}>
                              Open Calendar →
                            </button>
                          )}
                        </div>
                      )}

                      {/* Discussion topics */}
                      {tmpl.isList && (
                        <div style={{ marginTop: 8 }}>
                          {discussionTopics.length === 0 && (
                            <div style={{ fontSize: 12, color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontStyle: "italic", marginBottom: 8 }}>No topics yet</div>
                          )}
                          {discussionTopics.map(topic => (
                            <div key={topic.itemKey} style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 6 }}>
                              <button onClick={() => toggleTopic(topic.itemKey)}
                                style={{ width: 18, height: 18, borderRadius: 4, border: `2px solid ${topic.done ? C.green25 : C.border}`, background: topic.done ? C.green25 : "transparent", cursor: "pointer", flexShrink: 0, display: "flex", alignItems: "center", justifyContent: "center" }}>
                                {topic.done && <svg width="9" height="9" viewBox="0 0 10 10" fill="none"><path d="M1.5 5L4 7.5L8.5 2.5" stroke="white" strokeWidth="1.8" strokeLinecap="round" strokeLinejoin="round"/></svg>}
                              </button>
                              <span style={{ flex: 1, fontSize: 13, fontFamily: "'Helvetica Neue',Arial,sans-serif", color: topic.done ? C.textMuted : C.textPrimary, textDecoration: topic.done ? "line-through" : "none" }}>{topic.notes}</span>
                              <button onClick={() => removeTopic(topic.itemKey)} style={{ background: "none", border: "none", cursor: "pointer", color: C.textLight, display: "flex", alignItems: "center" }}><X size={13}/></button>
                            </div>
                          ))}
                          <div style={{ display: "flex", gap: 8, marginTop: 4 }}>
                            <AddTopicRow onAdd={addTopic}/>
                          </div>
                        </div>
                      )}

                      {/* Notes field for non-special items */}
                      {!tmpl.isStatic && !tmpl.isList && !tmpl.hasSongField && !tmpl.isPrayerList && (
                        <textarea placeholder="Notes…" value={row?.notes || ""} onChange={e => updateItem(tmpl.itemKey, "notes", e.target.value)}
                          style={{ width: "100%", marginTop: 8, padding: "7px 10px", fontSize: 12, fontFamily: "'Helvetica Neue',Arial,sans-serif", border: `1px solid ${C.borderLight}`, borderRadius: 6, resize: "vertical", minHeight: 50, background: row?.notes ? C.surfaceWarm : "#fff" }}/>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        );
      })}

      {/* Prayer List Modal */}
      {showPrayerList && (
        <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,.45)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}
          onClick={e => { if (e.target === e.currentTarget) setShowPrayerList(false); }}>
          <div style={{ background: "#fff", borderRadius: 14, padding: "28px 32px", maxWidth: 480, width: "100%", maxHeight: "80vh", display: "flex", flexDirection: "column", gap: 16 }}>
            <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between" }}>
              <span style={{ fontFamily: "Georgia,serif", fontSize: 18, color: C.textPrimary }}>Prayer List</span>
              <button onClick={() => setShowPrayerList(false)} style={{ background: "none", border: "none", cursor: "pointer", color: C.textMuted }}><X size={18}/></button>
            </div>
            {prayerListLoading ? (
              <div style={{ textAlign: "center", padding: "24px 0", color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontSize: 13 }}>Loading…</div>
            ) : !prayerListData || prayerListData.length === 0 ? (
              <div style={{ textAlign: "center", padding: "24px 0", color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", fontSize: 13 }}>No prayer list data found</div>
            ) : (
              <div style={{ overflowY: "auto" }}>
                {Object.entries(prayerListData.reduce((acc, item) => { (acc[item.category] = acc[item.category] || []).push(item); return acc; }, {})).map(([cat, items]) => (
                  <div key={cat} style={{ marginBottom: 16 }}>
                    <div style={{ fontSize: 10, fontWeight: 700, letterSpacing: ".1em", textTransform: "uppercase", color: C.textMuted, fontFamily: "'Helvetica Neue',Arial,sans-serif", marginBottom: 6 }}>{cat}</div>
                    {items.map((item, i) => (
                      <div key={i} style={{ fontSize: 14, fontFamily: "Georgia,serif", color: C.textPrimary, paddingLeft: 8, marginBottom: 4 }}>{item.name}</div>
                    ))}
                  </div>
                ))}
              </div>
            )}
          </div>
        </div>
      )}

      {/* Slack Preview Modal */}
      {slackDraft && (() => {
        const fl = slackDraft.firstLine;
        const setFl = v => setSlackDraft(d => ({...d, firstLine: v}));
        return (
          <SlackPreviewModal
            firstLine={fl}
            setFirstLine={setFl}
            bodyLines={slackDraft.bodyLines}
            channelName={slackDraft.channelName}
            sending={sending}
            onConfirm={() => doConfirmSlack(fl)}
            onClose={() => setSlackDraft(null)}
          />
        );
      })()}
    </div>
  );
}


// ─── Unauthorized View ────────────────────────────────────────────────────────
function UnauthorizedView(){
  return(
    <div style={{display:"flex",flexDirection:"column",alignItems:"center",justifyContent:"center",minHeight:"60vh",gap:16}}>
      <div style={{marginBottom:12,opacity:.3,color:C.textMuted}}><LockIcon/></div>
      <div style={{fontFamily:"Georgia,serif",fontSize:20,color:C.textSecond}}>Access Restricted</div>
      <div style={{fontSize:13,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",textAlign:"center",maxWidth:340,lineHeight:1.6}}>
        Your account does not have access to this application. Please contact your administrator to be granted access.
      </div>
    </div>
  );
}


// ─── Sacrament Meeting Tab ────────────────────────────────────────────────────

const SACRAMENT_SECTIONS = ["Presiding","Prayer","Music","Accompaniment","Speaker","Ordinance","Announcement"];

const SECTION_ICONS = {
  "Presiding": User, "Prayer": HandMetal, "Music": Music2,
  "Accompaniment": Piano, "Speaker": Mic2,
  "Ordinance": Sparkles, "Announcement": Megaphone,
};
const SECTION_META = {
  "Presiding":      { color:"#003057", bg:"#EAF0F6", border:"#005581" },
  "Prayer":         { color:"#4A2A7A", bg:"#F3EDF8", border:"#7B5EA7" },
  "Music":          { color:"#206B3F", bg:"#EAF4EA", border:"#50A83E" },
  "Accompaniment":  { color:"#974A07", bg:"#FEF3E2", border:"#F68D2E" },
  "Speaker":        { color:"#005581", bg:"#E6F4FA", border:"#007DA5" },
  "Ordinance":      { color:"#00558F", bg:"#E8F4FD", border:"#49CCE6" },
  "Announcement":   { color:"#53575B", bg:"#F5F3EE", border:"#D5CFBE" },
};
function MembersTab({data,setData,callings,releasings}){
  const[search,setSearch]=useState("");
  const[showForm,setSF]=useState(false);
  const[editing,setEditing]=useState(null);
  const[filter,setFilter]=useState("all"); // all | called | pipeline

  const lc=search.toLowerCase();
  let filtered=data.filter(m=>
    m.name.toLowerCase().includes(lc)||
    (m.calling||"").toLowerCase().includes(lc)||
    (m.notes||"").toLowerCase().includes(lc)
  );

  // Enrich members with their active pipeline items
  const enriched=filtered.map(m=>{
    const activeCalling  =callings.find(c=>c.name.toLowerCase()===m.name.toLowerCase()&&c.stage!=="Completed");
    const activeReleasing=releasings.find(r=>r.name.toLowerCase()===m.name.toLowerCase()&&r.stage!=="Completed");
    return{...m,activeCalling,activeReleasing};
  });

  const shown=enriched.filter(m=>{
    if(filter==="called")   return m.calling;
    if(filter==="pipeline") return m.activeCalling||m.activeReleasing;
    return true;
  });

  const save=item=>{
    if(item.id){setData(d=>d.map(x=>x.id===item.id?item:x));notify.success("Member updated");}
    else{setData(d=>[...d,{...item,id:`mb_${Date.now()}`}]);notify.success("Member added");}
    setSF(false);setEditing(null);
  };
  const del=id=>{setData(d=>d.filter(x=>x.id!==id));notify.info("Member removed");};
  const open=item=>{setEditing(item);setSF(true);};


  return(
    <div className="animate-in">
      <HeroBanner title="Members" sub={`${data.length} in pool · ${data.filter(m=>m.calling).length} currently called`}>
        <button className="btn-primary" style={{background:"rgba(255,255,255,.18)",border:"1.5px solid rgba(255,255,255,.4)"}} onClick={()=>{setEditing(null);setSF(true);}}>
          <PlusIcon/> Add Member
        </button>
      </HeroBanner>

      <div style={{display:"flex",gap:12,marginBottom:16,flexWrap:"wrap",alignItems:"center"}}>
        <input type="search" value={search} onChange={e=>setSearch(e.target.value)} placeholder="Search name, calling, or notes…" style={{maxWidth:320}}/>
        <div style={{display:"flex",gap:6}}>
          {[["all","All"],["called","Currently Called"],["pipeline","In Pipeline"]].map(([v,l])=>(
            <button key={v} onClick={()=>setFilter(v)} style={{borderRadius:7,fontSize:12,padding:"5px 13px",cursor:"pointer",border:`1.5px solid ${filter===v?C.blue35:C.border}`,color:filter===v?C.blue35:C.textSecond,background:filter===v?C.surfaceWarm:"transparent",fontWeight:filter===v?700:400,fontFamily:"'Helvetica Neue',Arial,sans-serif",transition:"all .15s"}}>{l}</button>
          ))}
        </div>
        <span style={{fontSize:12,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginLeft:"auto"}}>{shown.length} shown</span>
      </div>

      {shown.length===0&&(
        <div style={{textAlign:"center",padding:"60px 24px"}}>
          <div style={{marginBottom:12,opacity:.3,color:C.textMuted,display:"flex",justifyContent:"center"}}><User size={38}/></div>
          <div style={{fontFamily:"Georgia,serif",fontSize:17,color:C.textSecond,marginBottom:5}}>{data.length===0?"No members yet":"No results"}</div>
          <div style={{fontSize:12,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>{data.length===0?"Add members or pull from Sheets":"Try a different search or filter"}</div>
        </div>
      )}

      <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(340px,1fr))",gap:10}}>
        {shown.map(m=>(
          <div key={m.id} onClick={()=>open(m)} className="member-row" style={{cursor:"pointer",flexDirection:"column",alignItems:"stretch",gap:0,padding:"14px 16px"}}>
            <div style={{display:"flex",alignItems:"center",gap:12}}>
              <div className="member-avatar">{nameInitials(m.name||"?")}</div>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontFamily:"Georgia,serif",fontSize:16,color:C.textPrimary,marginBottom:2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.name}</div>
                {m.calling&&<div style={{fontSize:12,color:C.blue25,fontStyle:"italic",fontFamily:"Georgia,serif",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.calling}</div>}
                {m.notes&&<div style={{fontSize:11,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginTop:2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{m.notes}</div>}
              </div>
              <button className="btn-del" onClick={e=>{e.stopPropagation();del(m.id);}} style={{flexShrink:0}}><XIcon/></button>
            </div>
            {(m.activeCalling||m.activeReleasing)&&(
              <div style={{display:"flex",gap:6,marginTop:10,paddingTop:8,borderTop:`1px solid ${C.borderLight}`,flexWrap:"wrap"}}>
                {m.activeCalling&&(()=>{const st=PIPELINE_STAGE_STYLE[m.activeCalling.stage]||{};return<span className="chip" style={{background:st.bg,borderColor:st.border,color:st.text,fontSize:10}}><span style={{width:5,height:5,borderRadius:"50%",background:st.dot}}/>{m.activeCalling.stage} (calling)</span>;})()}
                {m.activeReleasing&&(()=>{const st=PIPELINE_STAGE_STYLE[m.activeReleasing.stage]||{};return<span className="chip" style={{background:st.bg,borderColor:st.border,color:st.text,fontSize:10}}><span style={{width:5,height:5,borderRadius:"50%",background:st.dot}}/>{m.activeReleasing.stage} (releasing)</span>;})()}
              </div>
            )}
          </div>
        ))}
      </div>

      {showForm&&<MemberModal item={editing} onSave={save} onClose={()=>{setSF(false);setEditing(null);}}/>}
    </div>
  );
}

function MemberModal({item,onSave,onClose}){
  const[f,setF]=useState(item||{name:"",calling:"",phone:"",email:"",notes:""});
  const s=(k,v)=>setF(x=>({...x,[k]:v}));
  return<ModalShell onClose={onClose} title={f.name||"Member"} subtitle={item?"Edit Member":"New Member"}>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:16}}>
      <FormRow label="Full Name" full><input value={f.name} onChange={e=>s("name",e.target.value)} placeholder="First Last"/></FormRow>
      <FormRow label="Current Calling (if any)" full><input value={f.calling||""} onChange={e=>s("calling",e.target.value)} placeholder="e.g. Primary Teacher"/></FormRow>
      <FormRow label="Phone"><input value={f.phone||""} onChange={e=>s("phone",e.target.value)} placeholder="+1 555 000 0000"/></FormRow>
      <FormRow label="Email"><input value={f.email||""} onChange={e=>s("email",e.target.value)} placeholder="member@email.com"/></FormRow>
      <FormRow label="Notes" full><textarea rows={3} value={f.notes||""} onChange={e=>s("notes",e.target.value)} placeholder="Additional notes…" style={{resize:"vertical"}}/></FormRow>
    </div>
    <ModalFooter onClose={onClose} onSave={()=>onSave(f)} saveLabel="Save Member"/>
  </ModalShell>;
}


// ─── Links Tab ────────────────────────────────────────────────────────────────
function LinksTab({ bishopricLinks, setBishopricLinks, wcLinks, setWcLinks, token, isAdmin }) {
  // Admin sees two sections: their private Bishopric links + shared WC links
  // WC-only users see just the WC links
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing]   = useState(null);  // {link, source: "bishopric"|"wc"}
  const [saving, setSaving]     = useState(false);

  const allSections = isAdmin
    ? [
        { key: "bishopric", label: "Bishopric Links", data: bishopricLinks || [], setData: setBishopricLinks,
          color: C.blue35, bg: "#EAF0F6", border: C.blue25,
          pushFn: "pushBishopricLinks", desc: "Visible to Bishopric only" },
        { key: "wc",        label: "Ward Council Links", data: wcLinks, setData: setWcLinks,
          color: C.green35, bg: "#EAF4EA", border: C.green25,
          pushFn: "pushWardCouncilLinks", desc: "Visible to everyone" },
      ]
    : [
        { key: "wc", label: "Important Links", data: wcLinks, setData: setWcLinks,
          color: C.blue35, bg: "#EAF0F6", border: C.blue25,
          pushFn: "pushWardCouncilLinks", desc: "" },
      ];

  const openAdd = (sectionKey) => {
    setEditing({ link: { id: `link_${Date.now()}`, name: "", url: "", description: "" }, source: sectionKey });
    setShowForm(true);
  };

  const openEdit = (link, sectionKey) => {
    setEditing({ link: { ...link }, source: sectionKey });
    setShowForm(true);
  };

  const handleSave = async (updated) => {
    const section = allSections.find(s => s.key === editing.source);
    if (!section) return;
    setSaving(true);
    const isNew = !section.data.some(l => l.id === updated.id);
    const newData = isNew
      ? [...section.data, updated]
      : section.data.map(l => l.id === updated.id ? updated : l);
    section.setData(newData);
    try {
      const mod = await import("./sheets");
      await mod[section.pushFn](token, newData);
      notify.success(isNew ? "Link added" : "Link updated");
    } catch(e) {
      notify.error("Save failed: " + e.message);
    }
    setSaving(false);
    setShowForm(false);
    setEditing(null);
  };

  const handleDelete = async (id, sectionKey) => {
    const section = allSections.find(s => s.key === sectionKey);
    if (!section) return;
    const newData = section.data.filter(l => l.id !== id);
    section.setData(newData);
    try {
      const mod = await import("./sheets");
      await mod[section.pushFn](token, newData);
      notify.info("Link removed");
    } catch(e) {
      notify.error("Save failed: " + e.message);
    }
  };

  return (
    <div className="animate-in">
      <HeroBanner title="Links" sub="Quick access to important resources">
        {isAdmin && (
          <>
            <button onClick={() => openAdd("bishopric")}
              style={{ background:"rgba(255,255,255,.14)", border:"1.5px solid rgba(255,255,255,.3)", color:"#fff", borderRadius:8, padding:"7px 14px", fontSize:12, fontFamily:"'Helvetica Neue',Arial,sans-serif", fontWeight:600, cursor:"pointer", display:"flex", alignItems:"center", gap:6 }}>
              <Plus size={13}/> Add Bishopric Link
            </button>
            <button onClick={() => openAdd("wc")}
              style={{ background:"rgba(255,255,255,.18)", border:"1.5px solid rgba(255,255,255,.4)", color:"#fff", borderRadius:8, padding:"7px 14px", fontSize:12, fontFamily:"'Helvetica Neue',Arial,sans-serif", fontWeight:700, cursor:"pointer", display:"flex", alignItems:"center", gap:6 }}>
              <Plus size={13}/> Add Ward Council Link
            </button>
          </>
        )}
        {!isAdmin && (
          <button onClick={() => openAdd("wc")}
            style={{ background:"rgba(255,255,255,.18)", border:"1.5px solid rgba(255,255,255,.4)", color:"#fff", borderRadius:8, padding:"7px 14px", fontSize:12, fontFamily:"'Helvetica Neue',Arial,sans-serif", fontWeight:700, cursor:"pointer", display:"flex", alignItems:"center", gap:6 }}>
            <Plus size={13}/> Add Link
          </button>
        )}
      </HeroBanner>

      {allSections.map(section => (
        <div key={section.key} style={{ marginBottom: 28 }}>
          {/* Section header — only shown for admin (two sections) */}
          {isAdmin && (
            <div style={{ display:"flex", alignItems:"center", gap:10, marginBottom:14 }}>
              <div style={{ width:10, height:10, borderRadius:"50%", background:section.color, flexShrink:0 }}/>
              <span style={{ fontFamily:"Georgia,serif", fontSize:16, color:section.color }}>{section.label}</span>
              {section.desc && <span style={{ fontSize:11, color:C.textMuted, fontFamily:"'Helvetica Neue',Arial,sans-serif", padding:"2px 8px", background:section.bg, border:`1px solid ${section.border}`, borderRadius:10 }}>{section.desc}</span>}
              <button onClick={() => openAdd(section.key)}
                style={{ marginLeft:"auto", background:"none", border:`1.5px solid ${section.border}`, color:section.color, borderRadius:7, padding:"4px 12px", fontSize:11, fontFamily:"'Helvetica Neue',Arial,sans-serif", fontWeight:700, cursor:"pointer", display:"flex", alignItems:"center", gap:5 }}>
                <Plus size={11}/> Add
              </button>
            </div>
          )}

          {section.data.length === 0 ? (
            <div style={{ textAlign:"center", padding:"40px 24px", background:C.surfaceWhite, borderRadius:12, border:`1.5px dashed ${C.border}` }}>
              <div style={{ fontSize:13, color:C.textMuted, fontFamily:"'Helvetica Neue',Arial,sans-serif", fontStyle:"italic" }}>
                No links yet — add one to get started
              </div>
            </div>
          ) : (
            <div style={{ display:"grid", gridTemplateColumns:"repeat(auto-fill, minmax(280px, 1fr))", gap:12 }}>
              {section.data.map(link => (
                <div key={link.id} style={{ background:C.surfaceWhite, border:`1.5px solid ${C.border}`, borderRadius:12, padding:"18px 20px", display:"flex", flexDirection:"column", gap:10, transition:"border .15s, box-shadow .15s" }}
                  onMouseEnter={e => { e.currentTarget.style.borderColor=section.border; e.currentTarget.style.boxShadow=`0 3px 14px rgba(0,0,0,.07)`; }}
                  onMouseLeave={e => { e.currentTarget.style.borderColor=C.border; e.currentTarget.style.boxShadow="none"; }}>
                  {/* Name + action buttons */}
                  <div style={{ display:"flex", alignItems:"flex-start", justifyContent:"space-between", gap:8 }}>
                    <div style={{ fontFamily:"Georgia,serif", fontSize:16, color:C.textPrimary, lineHeight:1.3, flex:1 }}>{link.name || <span style={{color:C.textLight,fontStyle:"italic"}}>Untitled</span>}</div>
                    <div style={{ display:"flex", gap:4, flexShrink:0 }}>
                      <button onClick={() => openEdit(link, section.key)}
                        style={{ background:"none", border:`1px solid ${C.borderLight}`, borderRadius:5, padding:"3px 8px", fontSize:11, color:C.textSecond, cursor:"pointer", fontFamily:"'Helvetica Neue',Arial,sans-serif" }}>
                        Edit
                      </button>
                      <button onClick={() => handleDelete(link.id, section.key)}
                        style={{ background:"none", border:`1px solid ${C.borderLight}`, borderRadius:5, width:26, height:26, cursor:"pointer", color:C.textLight, display:"flex", alignItems:"center", justifyContent:"center" }}
                        onMouseEnter={e=>{e.currentTarget.style.borderColor=C.red15;e.currentTarget.style.color=C.red15;}}
                        onMouseLeave={e=>{e.currentTarget.style.borderColor=C.borderLight;e.currentTarget.style.color=C.textLight;}}>
                        <X size={12}/>
                      </button>
                    </div>
                  </div>

                  {/* Description */}
                  {link.description && (
                    <div style={{ fontSize:13, color:C.textMuted, fontFamily:"'Helvetica Neue',Arial,sans-serif", lineHeight:1.5 }}>
                      {link.description}
                    </div>
                  )}

                  {/* URL + open button */}
                  <div style={{ marginTop:"auto", paddingTop:10, borderTop:`1px solid ${C.borderLight}`, display:"flex", alignItems:"center", justifyContent:"space-between", gap:8 }}>
                    <span style={{ fontSize:11, color:C.textMuted, fontFamily:"'Helvetica Neue',Arial,sans-serif", overflow:"hidden", textOverflow:"ellipsis", whiteSpace:"nowrap", flex:1 }}>
                      {link.url || <span style={{fontStyle:"italic"}}>No URL set</span>}
                    </span>
                    <a href={link.url} target="_blank" rel="noopener noreferrer"
                      onClick={e => { if (!link.url) e.preventDefault(); }}
                      style={{ display:"inline-flex", alignItems:"center", gap:5, padding:"6px 14px", borderRadius:7, fontSize:12, fontWeight:700, fontFamily:"'Helvetica Neue',Arial,sans-serif", textDecoration:"none", flexShrink:0, cursor: link.url ? "pointer" : "not-allowed", opacity: link.url ? 1 : .4, background:section.color, color:"#fff" }}>
                      <ExternalLink size={12}/> Open
                    </a>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
      ))}

      {/* Add/Edit Modal */}
      {showForm && editing && (
        <LinkFormModal
          link={editing.link}
          source={editing.source}
          isAdmin={isAdmin}
          saving={saving}
          onSave={handleSave}
          onClose={() => { setShowForm(false); setEditing(null); }}
        />
      )}
    </div>
  );
}

function LinkFormModal({ link, source, isAdmin, saving, onSave, onClose }) {
  const [f, setF] = useState({ ...link });
  const s = (k, v) => setF(p => ({ ...p, [k]: v }));
  const isNew = !link.name && !link.url;
  const sectionLabel = source === "bishopric" ? "Bishopric Link" : isAdmin ? "Ward Council Link" : "Link";

  return (
    <ModalShell onClose={onClose} title={isNew ? `New ${sectionLabel}` : (f.name || "Edit Link")} subtitle={isNew ? "Add Link" : "Edit Link"}>
      <div style={{ display:"flex", flexDirection:"column", gap:16 }}>
        <FormRow label="Name">
          <input value={f.name} onChange={e=>s("name",e.target.value)} placeholder="e.g. Church Handbook"/>
        </FormRow>
        <FormRow label="URL">
          <input value={f.url} onChange={e=>s("url",e.target.value)} placeholder="https://…"/>
        </FormRow>
        <FormRow label="Description (optional)">
          <textarea rows={3} value={f.description||""} onChange={e=>s("description",e.target.value)}
            placeholder="Brief description of what this link is for…" style={{resize:"vertical"}}/>
        </FormRow>
      </div>
      <ModalFooter onClose={onClose} onSave={() => { if(f.name.trim()||f.url.trim()) onSave(f); }} saving={saving} saveLabel="Save Link"/>
    </ModalShell>
  );
}

// ─── Settings Modal ───────────────────────────────────────────────────────────
function SettingsModal({token,onTestConn,connStatus,connMsg,roster,setRoster,onClose}){
  const connColors={ok:C.green25,error:C.red15,warn:C.gold20,testing:C.blue25,unknown:C.textLight};
  const connLabels={ok:"Connected",error:"Error",warn:"Warning",testing:"Testing…",unknown:"Not tested"};
  return(
    <ModalShell onClose={onClose} title="Settings & Configuration" subtitle="Ward Manager">
      <div style={{display:"flex",flexDirection:"column",gap:20}}>

        {/* Connection */}
        <section>
          <div style={{fontSize:12,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginBottom:12,paddingBottom:8,borderBottom:`1px solid ${C.borderLight}`}}>Google Sheets Connection</div>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
            <div style={{display:"flex",alignItems:"center",gap:8}}>
              <span style={{width:10,height:10,borderRadius:"50%",background:connColors[connStatus]||C.textLight,animation:connStatus==="testing"?"pulse 1s infinite":"none"}}/>
              <span style={{fontFamily:"'Helvetica Neue',Arial,sans-serif",fontSize:14,fontWeight:600,color:connColors[connStatus]||C.textMuted}}>{connLabels[connStatus]||connStatus}</span>
            </div>
            <button className="btn-primary" onClick={onTestConn} style={{fontSize:12,padding:"7px 16px"}}>Test Connection</button>
          </div>
          {connMsg&&<div style={{background:connStatus==="ok"?"#EAF4EA":connStatus==="error"?"#FEF0F4":"#FEF3E2",border:`1px solid ${connColors[connStatus]}`,borderRadius:8,padding:"10px 14px",fontSize:12,color:connStatus==="ok"?C.green35:connStatus==="error"?"#BD0057":"#974A07",fontFamily:"'Helvetica Neue',Arial,sans-serif",lineHeight:1.6}}>{connMsg}</div>}
        </section>

        {/* Config reference */}
        <section>
          <div style={{fontSize:12,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginBottom:12,paddingBottom:8,borderBottom:`1px solid ${C.borderLight}`}}>Current Configuration</div>
          <div style={{display:"flex",flexDirection:"column",gap:8}}>
            {(() => {
              const nc = (v, placeholder) => !v || v.includes("YOUR_") || v.includes("YOUR/") ? "Not configured" : placeholder || "Configured";
              const rows = [
                {label:"Ward Name",         value: config.WARD_NAME || "Not set"},
                {label:"Google Client ID",  value: nc(config.GOOGLE_CLIENT_ID)},
                {label:"Bishopric Sheet",   value: nc(config.SPREADSHEET_ID, config.SPREADSHEET_ID?.slice(0,22)+"…")},
                {label:"Sacrament Sheet",   value: nc(config.SACRAMENT_SHEET_ID)},
                {label:"Prayer List Sheet", value: nc(config.PRAYER_LIST_SHEET_ID)},
                {label:"Ward Council Sheet",value: nc(config.WARD_COUNCIL_SHEET_ID)},
                {label:"Slack Relay",       value: nc(config.SLACK_RELAY_URL)},
                {label:"Slack — Appointments",  value: nc(config.SLACK_WEBHOOK_APPOINTMENTS,  `#${config.SLACK_WEBHOOK_APPOINTMENTS_NAME||"?"}`)},
                {label:"Slack — Callings",       value: nc(config.SLACK_WEBHOOK_CALLINGS,      `#${config.SLACK_WEBHOOK_CALLINGS_NAME||"?"}`)},
                {label:"Slack — Bishopric",      value: nc(config.SLACK_WEBHOOK_BISHOPRIC,     `#${config.SLACK_WEBHOOK_BISHOPRIC_NAME||"?"}`)},
                {label:"Slack — Ward Council",   value: nc(config.SLACK_WEBHOOK_WARD_COUNCIL,  `#${config.SLACK_WEBHOOK_WARD_COUNCIL_NAME||"?"}`)},
                {label:"Admin users",        value: `${config.ALLOWED_EMAILS?.length||0} email${config.ALLOWED_EMAILS?.length!==1?"s":""}`},
                {label:"Ward Council users", value: `${config.WARD_COUNCIL_EMAILS?.length||0} email${config.WARD_COUNCIL_EMAILS?.length!==1?"s":""}`},
              ];
              return rows.map(({label,value})=>(
                <div key={label} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 12px",background:C.surfaceWarm,borderRadius:7}}>
                  <span style={{fontSize:12,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>{label}</span>
                  <span style={{fontSize:12,fontWeight:600,color:!value || value === "Not configured" || value === "Not set"?C.gold20:C.green35,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>{value}</span>
                </div>
              ));
            })()}
          </div>
          <p style={{fontSize:11,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginTop:12,lineHeight:1.6}}>To change settings, edit <code style={{background:C.surfaceWarm,padding:"1px 5px",borderRadius:4,fontSize:11}}>src/config.js</code> in your repository and redeploy.</p>
        </section>

        {/* Sheet structure */}
        <section>
          <div style={{fontSize:12,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginBottom:12,paddingBottom:8,borderBottom:`1px solid ${C.borderLight}`}}>Required Google Sheet Tabs</div>
          {[
            {
              sheet: "Bishopric Sheet", color: C.blue35,
              tabs: [
                {tab:"Appointments",     cols:"Name | Status | Owner | Purpose | Appt Date | Notes"},
                {tab:"Callings",         cols:"Calling | Name | Stage | Notes"},
                {tab:"Releasings",       cols:"Calling | Name | Stage | Notes"},
                {tab:"BishopricMeeting", cols:"Date | ItemKey | Assignee | Done | Notes | CustomLabel | SpiritualToggle"},
                {tab:"Members",          cols:"Name | Calling | Phone | Email | Notes"},
                {tab:"Roster",           cols:"Role | Name"},
                {tab:"Links",            cols:"ID | Name | URL | Description"},
              ]
            },
            {
              sheet: "Sacrament Sheet", color: C.purple,
              tabs: [
                {tab:"SacramentProgram", cols:"Date | Section | Order | Label | Value | Notes"},
              ]
            },
            {
              sheet: "Ward Council Sheet", color: C.green35,
              tabs: [
                {tab:"WardCouncilMeeting", cols:"Date | ItemKey | Assignee | Done | Notes | CustomLabel | SpiritualToggle"},
                {tab:"Calendar",           cols:"Date | Time | Event"},
                {tab:"Roster",             cols:"Role | Name"},
                {tab:"Links",              cols:"ID | Name | URL | Description"},
              ]
            },
            {
              sheet: "Prayer List Sheet", color: C.gold20,
              tabs: [
                {tab:"Sheet1", cols:"Category | Name"},
              ]
            },
          ].map(({sheet,color,tabs})=>(
            <div key={sheet} style={{marginBottom:12}}>
              <div style={{fontSize:10,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",color,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginBottom:6}}>{sheet}</div>
              <div style={{display:"flex",flexDirection:"column",gap:4}}>
                {tabs.map(({tab,cols})=>(
                  <div key={tab} style={{padding:"7px 12px",background:C.surfaceWarm,borderRadius:7,borderLeft:`3px solid ${color}`}}>
                    <div style={{fontSize:12,fontWeight:700,color:C.textPrimary,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginBottom:2}}>{tab}</div>
                    <div style={{fontSize:11,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>{cols}</div>
                  </div>
                ))}
              </div>
            </div>
          ))}
        </section>

        {/* ── Leadership Roster ── */}
        <section>
          <div style={{fontSize:12,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginBottom:6,paddingBottom:8,borderBottom:`1px solid ${C.borderLight}`}}>Leadership Roster</div>
          <div style={{fontSize:11,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginBottom:12,lineHeight:1.5}}>Names entered here appear throughout the app in assignee dropdowns and owner fields. Leave blank to show role titles only.</div>
          <RosterEditor roster={roster} setRoster={setRoster} token={token}/>
        </section>
      </div>
    </ModalShell>
  );
}

function RosterEditor({roster,setRoster,token}) {
  // Build a complete local copy keyed by all 11 positions.
  // We do NOT sync back from the roster prop after mount — doing so
  // would overwrite in-progress edits every render cycle.
  const buildLocal = (src) =>
    ROSTER_POSITIONS.map(p => {
      const found = src.find(r => r.role === p.role);
      return { role: p.role, name: found ? (found.name || "") : "" };
    });

  const [localRoster, setLocal] = useState(() => buildLocal(roster));
  const [saving, setSaving] = useState(false);

  // Only re-initialise when the parent roster changes AND we haven't
  // made any local edits yet (i.e. all names are still empty — this
  // handles the async-load case where the real names arrive after mount).
  const hasLocalEdits = localRoster.some(r => r.name !== "");
  useEffect(() => {
    if (!hasLocalEdits) setLocal(buildLocal(roster));
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [roster]);

  const updateName = (role, name) => {
    setLocal(prev => prev.map(r => r.role === role ? { ...r, name } : r));
  };

  const doSave = async () => {
    setSaving(true);
    try {
      const { pushRoster } = await import("./sheets");
      await pushRoster(token, localRoster);
      setRoster(localRoster);
      notify.success("Roster saved");
    } catch(e) {
      notify.error("Save failed: " + e.message);
    } finally {
      setSaving(false);
    }
  };

  const groups = [
    { label: "Bishopric Council", positions: ROSTER_POSITIONS.filter(p=>p.group==="bishopric") },
    { label: "Ward Council (additional)", positions: ROSTER_POSITIONS.filter(p=>p.group==="ward_council") },
  ];

  return (
    <div>
      {groups.map(g => (
        <div key={g.label} style={{marginBottom:16}}>
          <div style={{fontSize:10,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",color:C.blue35,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginBottom:8}}>{g.label}</div>
          <div style={{display:"flex",flexDirection:"column",gap:6}}>
            {g.positions.map(p => {
              const entry = localRoster.find(r => r.role === p.role) || {role:p.role,name:""};
              return (
                <div key={p.role} style={{display:"flex",alignItems:"center",gap:10,padding:"6px 10px",background:C.surfaceWarm,borderRadius:7}}>
                  <span style={{fontSize:11,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",width:160,flexShrink:0}}>{p.role}</span>
                  <input
                    type="text"
                    value={entry.name}
                    onChange={e => updateName(p.role, e.target.value)}
                    placeholder={`Name for ${p.role}…`}
                    style={{flex:1,padding:"5px 10px",fontSize:13,margin:0}}
                  />
                </div>
              );
            })}
          </div>
        </div>
      ))}
      <button onClick={doSave} disabled={saving} className="btn-primary" style={{marginTop:4,fontSize:12,padding:"7px 18px",display:"flex",alignItems:"center",gap:6}}>
        <Save size={12}/> {saving ? "Saving…" : "Save Roster"}
      </button>
    </div>
  );
}


// ─── Alerts Tab ───────────────────────────────────────────────────────────────
function AlertsTab({appointments, callings, releasings}){
  const[source,setSource]=useState("Appointments");
  const[status,setStatus]=useState("");
  const[sending1,setSending1]=useState(false);
  const[sending2,setSending2]=useState(false);
  const[slackDraft,setSlackDraft]=useState(null); // {firstLine, bodyLines, channelKey, channelName, setSending}

  const ch1Name = config.SLACK_WEBHOOK_APPOINTMENTS_NAME||"Appointments";
  const ch2Name = config.SLACK_WEBHOOK_CALLINGS_NAME||"Callings";
  const ch1URL  = config.SLACK_WEBHOOK_APPOINTMENTS||"";
  const ch2URL  = config.SLACK_WEBHOOK_CALLINGS||"";

  // Build stage options per source
  const stageOptions = source==="Appointments" ? APPT_STAGES
    : source==="Callings" ? CALLING_STAGES
    : RELEASING_STAGES;

  // Reset status when source changes
  useEffect(()=>{ setStatus(stageOptions[0]||""); },[source]);

  // Filter matching records
  const preview = (() => {
    if(!status) return [];
    if(source==="Appointments")
      return appointments.filter(a=>a.status===status);
    if(source==="Callings")
      return callings.filter(c=>c.stage===status);
    return releasings.filter(r=>r.stage===status);
  })();

  // Build Slack message body lines from current preview data
  const buildBodyLines = () => {
    const itemLines = preview.map(item => {
      if(source==="Appointments"){
        const date = item.apptDate ? `  •  ${item.apptDate}` : "";
        const purpose = item.purpose ? `  •  ${item.purpose}` : "";
        return `• *${item.name||"—"}*${purpose}  •  ${item.status}${date}`;
      } else {
        const notes = item.notes ? `  •  _${item.notes}_` : "";
        return `• *${item.name||"—"}*  •  ${item.calling||"—"}  •  ${item.stage}${notes}`;
      }
    });
    const divider = "─────────────────────────────";
    const footer  = `_${preview.length} record${preview.length===1?"":"s"} · ${source} · Ward Manager_`;
    const body    = itemLines.length > 0 ? itemLines.join("\n") : "_No records found_";
    return [divider, body, divider, footer];
  };

  // Open preview modal — saves which channel was chosen
  const openSlackPreview = (channelKey, channelName, setSending) => {
    const relayURL = config.SLACK_RELAY_URL||"";
    if(!relayURL || relayURL.includes("YOUR_SCRIPT_ID")){
      notify.error("No relay URL configured. Add SLACK_RELAY_URL to config.js");
      return;
    }
    const firstLine = `*Ward Manager Alert — ${source}: "${status}"*`;
    setSlackDraft({ firstLine, bodyLines: buildBodyLines(), channelKey, channelName, setSending, relayURL });
  };

  // Actually send after user confirms
  const doConfirmSlack = async (firstLine) => {
    if (!slackDraft) return;
    const { bodyLines, channelKey, channelName, setSending, relayURL } = slackDraft;
    const text = [firstLine, ...bodyLines].join("\n");
    // Rebuild blocks with edited firstLine
    const blocks = [
      {type:"section", text:{type:"mrkdwn", text: firstLine + "\n" + bodyLines[0]}},
      {type:"section", text:{type:"mrkdwn", text: bodyLines[1] || "_No records found_"}},
      {type:"context", elements:[{type:"mrkdwn", text: bodyLines[bodyLines.length-1]}]},
    ];
    setSending(true);
    try {
      const res = await fetch(relayURL, {
        method:"POST", headers:{"Content-Type":"text/plain"},
        body: JSON.stringify({channel: channelKey, text, blocks}),
      });
      const data = await res.json().catch(()=>({status:res.status}));
      if(res.ok && data.status===200){
        notify.success(`Alert sent to #${channelName} (${preview.length} record${preview.length===1?"":"s"})`);
        setSlackDraft(null);
      } else {
        notify.error(`Relay error: ${data.message||res.status}. Check your Apps Script and webhook URLs.`);
      }
    } catch(e){
      notify.error(`Failed to reach relay: ${e.message}`);
    } finally {
      setSending(false);
    }
  };

  const stageStyle = source==="Appointments"
    ? APPT_STAGE_STYLE[status]||{}
    : PIPELINE_STAGE_STYLE[status]||{};

  return(
    <div className="animate-in">
      <HeroBanner title="Alerts" sub="Send status reports to Slack channels"/>

      <div style={{display:"grid",gridTemplateColumns:"340px 1fr",gap:24,marginTop:4}}>

        {/* ── Left panel: controls ── */}
        <div style={{display:"flex",flexDirection:"column",gap:16}}>

          {/* Source picker */}
          <div style={{background:C.surfaceWhite,border:`1.5px solid ${C.border}`,borderRadius:10,padding:20}}>
            <div style={{fontSize:11,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginBottom:12}}>Data Source</div>
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {["Appointments","Callings","Releasings"].map(s=>(
                <button key={s} onClick={()=>setSource(s)} style={{
                  display:"flex",alignItems:"center",justifyContent:"space-between",
                  padding:"10px 14px",borderRadius:8,border:`1.5px solid ${source===s?C.blue35:C.borderLight}`,
                  background:source===s?"rgba(0,85,129,.06)":C.surfaceWarm,
                  cursor:"pointer",transition:"all .15s",
                }}>
                  <span style={{fontSize:13,fontWeight:source===s?700:400,color:source===s?C.blue35:C.textPrimary,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>
                    {s==="Appointments"?<Calendar size={12}/>:s==="Callings"?<ClipboardList size={12}/>:<Tag size={12}/>} {s}
                  </span>
                  {source===s&&<span style={{width:8,height:8,borderRadius:"50%",background:C.blue35}}/>}
                </button>
              ))}
            </div>
          </div>

          {/* Status picker */}
          <div style={{background:C.surfaceWhite,border:`1.5px solid ${C.border}`,borderRadius:10,padding:20}}>
            <div style={{fontSize:11,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginBottom:12}}>Filter by Status / Stage</div>
            <div style={{display:"flex",flexDirection:"column",gap:6}}>
              {stageOptions.map(s=>{
                const st = source==="Appointments" ? APPT_STAGE_STYLE[s]||{} : PIPELINE_STAGE_STYLE[s]||{};
                const count = source==="Appointments"
                  ? appointments.filter(a=>a.status===s).length
                  : source==="Callings"
                  ? callings.filter(c=>c.stage===s).length
                  : releasings.filter(r=>r.stage===s).length;
                return(
                  <button key={s} onClick={()=>setStatus(s)} style={{
                    display:"flex",alignItems:"center",justifyContent:"space-between",
                    padding:"8px 12px",borderRadius:7,
                    border:`1.5px solid ${status===s?(st.border||C.blue35):C.borderLight}`,
                    background:status===s?(st.bg||"rgba(0,85,129,.06)"):C.surfaceWarm,
                    cursor:"pointer",transition:"all .15s",
                  }}>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      <span style={{width:7,height:7,borderRadius:"50%",background:st.dot||C.textLight,flexShrink:0}}/>
                      <span style={{fontSize:12,fontWeight:status===s?600:400,color:status===s?(st.text||C.textPrimary):C.textPrimary,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>{s}</span>
                    </div>
                    <span style={{fontSize:11,fontWeight:700,color:count>0?C.blue25:C.textLight,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>{count}</span>
                  </button>
                );
              })}
            </div>
          </div>

          {/* Send buttons */}
          <div style={{background:C.surfaceWhite,border:`1.5px solid ${C.border}`,borderRadius:10,padding:20}}>
            <div style={{fontSize:11,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginBottom:12}}>Send Alert</div>
            {(!ch1URL && !ch2URL) && (
              <div style={{padding:"10px 12px",background:"#FEF3E2",border:`1px solid ${C.gold20}`,borderRadius:7,marginBottom:12}}>
                <p style={{fontSize:12,color:"#974A07",fontFamily:"'Helvetica Neue',Arial,sans-serif",lineHeight:1.5}}>
                  <AlertTriangle size={12} style={{flexShrink:0}}/> No Slack webhooks configured.<br/>Add <code style={{fontSize:11}}>SLACK_WEBHOOK_APPOINTMENTS</code> and <code style={{fontSize:11}}>SLACK_WEBHOOK_CALLINGS</code> to <code style={{fontSize:11}}>config.js</code>.
                </p>
              </div>
            )}
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              <button
                onClick={()=>openSlackPreview(ch1URL,ch1Name,setSending1)}
                disabled={sending1||preview.length===0}
                style={{
                  padding:"10px 16px",borderRadius:8,border:"none",
                  background:preview.length>0?C.blue35:"#ccc",
                  color:"#fff",fontWeight:600,fontSize:13,
                  fontFamily:"'Helvetica Neue',Arial,sans-serif",
                  cursor:preview.length>0?"pointer":"not-allowed",
                  opacity:sending1?.7:1,transition:"all .15s",
                  display:"flex",alignItems:"center",justifyContent:"center",gap:8,
                }}>
                {sending1?"Sending…":<><Bell size={13}/> Send to #{ch1Name} <span style={{opacity:.7,fontWeight:400,fontSize:11}}>({preview.length})</span></>}
              </button>
              <button
                onClick={()=>openSlackPreview(ch2URL,ch2Name,setSending2)}
                disabled={sending2||preview.length===0}
                style={{
                  padding:"10px 16px",borderRadius:8,border:`1.5px solid ${C.blue35}`,
                  background:"transparent",
                  color:preview.length>0?C.blue35:"#ccc",fontWeight:600,fontSize:13,
                  fontFamily:"'Helvetica Neue',Arial,sans-serif",
                  cursor:preview.length>0?"pointer":"not-allowed",
                  opacity:sending2?.7:1,transition:"all .15s",
                  display:"flex",alignItems:"center",justifyContent:"center",gap:8,
                }}>
                {sending2?"Sending…":<><Bell size={13}/> Send to #{ch2Name} <span style={{opacity:.7,fontWeight:400,fontSize:11}}>({preview.length})</span></>}
              </button>
            </div>
          </div>
        </div>

        {/* ── Right panel: preview ── */}
        <div style={{background:C.surfaceWhite,border:`1.5px solid ${C.border}`,borderRadius:10,padding:24}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:16}}>
            <div>
              <div style={{fontSize:11,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>Alert Preview</div>
              <div style={{fontSize:13,color:C.textPrimary,marginTop:4,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>
                {source} · <span style={{padding:"2px 8px",borderRadius:10,fontSize:12,background:stageStyle.bg||C.surfaceWarm,color:stageStyle.text||C.textPrimary,border:`1px solid ${stageStyle.border||C.border}`}}>{status}</span>
              </div>
            </div>
            <div style={{fontSize:24,fontWeight:700,color:preview.length>0?C.blue35:C.textLight}}>{preview.length}</div>
          </div>

          {preview.length===0?(
            <div style={{textAlign:"center",padding:"48px 24px",color:C.textLight}}>
              <div style={{marginBottom:12,opacity:.3,color:C.textMuted,display:"flex",justifyContent:"center"}}><Bell size={32}/></div>
              <div style={{fontSize:14,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>No records match this filter</div>
            </div>
          ):(
            <div style={{display:"flex",flexDirection:"column",gap:8}}>
              {preview.map((item,i)=>{
                const name   = item.name||"—";
                const detail = source==="Appointments" ? item.purpose||"—" : item.calling||"—";
                const stage  = source==="Appointments" ? item.status : item.stage;
                const date   = source==="Appointments" ? item.apptDate : item.notes;
                const st     = source==="Appointments" ? APPT_STAGE_STYLE[stage]||{} : PIPELINE_STAGE_STYLE[stage]||{};
                return(
                  <div key={item.id||i} style={{display:"grid",gridTemplateColumns:"1fr 160px 160px",gap:12,padding:"12px 14px",background:C.surfaceWarm,borderRadius:8,alignItems:"center"}}>
                    <div>
                      <div style={{fontSize:14,fontWeight:600,color:C.textPrimary,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>{name}</div>
                      <div style={{fontSize:12,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginTop:2}}>{detail}</div>
                    </div>
                    <div>
                      <span style={{padding:"3px 9px",borderRadius:10,fontSize:11,fontWeight:600,background:st.bg||C.surfaceWarm,color:st.text||C.textPrimary,border:`1px solid ${st.border||C.border}`,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>{stage}</span>
                    </div>
                    <div style={{fontSize:12,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>{date||<span style={{color:C.textLight,fontStyle:"italic"}}>no date</span>}</div>
                  </div>
                );
              })}
            </div>
          )}

          {/* Slack message preview */}
          {preview.length>0&&(
            <div style={{marginTop:20,padding:16,background:"#1a1d21",borderRadius:8}}>
              <div style={{fontSize:10,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",color:"#868686",fontFamily:"'Helvetica Neue',Arial,sans-serif",marginBottom:10}}>Slack Message Preview</div>
              <pre style={{fontSize:12,color:"#d1d2d3",fontFamily:"'Slack-Lato',monospace",lineHeight:1.7,whiteSpace:"pre-wrap",margin:0}}>
{`Ward Manager Alert — ${source}: "${status}"
─────────────────────────────
${preview.map(item=>{
  if(source==="Appointments"){
    const d=item.apptDate?`  •  ${item.apptDate}`:"";
    return `• ${item.name||"—"}  •  ${item.purpose||"—"}  •  ${item.status}${d}`;
  } else {
    const n=item.notes?`  •  ${item.notes}`:"";
    return `• ${item.name||"—"}  •  ${item.calling||"—"}  •  ${item.stage}${n}`;
  }
}).join("\n")}
─────────────────────────────
${preview.length} record${preview.length===1?"":"s"} · ${source} · Ward Manager`}
              </pre>
            </div>
          )}
        </div>
      </div>

      {/* Slack Preview Modal */}
      {slackDraft && (() => {
        const fl = slackDraft.firstLine;
        const setFl = v => setSlackDraft(d => ({...d, firstLine: v}));
        return (
          <SlackPreviewModal
            firstLine={fl}
            setFirstLine={setFl}
            bodyLines={slackDraft.bodyLines}
            channelName={slackDraft.channelName}
            sending={sending1||sending2}
            onConfirm={() => doConfirmSlack(fl)}
            onClose={() => setSlackDraft(null)}
          />
        );
      })()}
    </div>
  );
}


function SacramentTab({ data, setData, saveFn, pullFn, tokenRef }) {
  const allDates = [...new Set(data.map(r=>r.date).filter(Boolean))].sort().reverse();
  // Always default to the upcoming Sunday (matches Bishopric tab behaviour)
  const [activeDate, setActiveDate] = useState(nextSunday);
  const [editingItem, setEditingItem] = useState(null);
  const [showPrint, setShowPrint] = useState(false);
  const [showOtherDate, setShowOtherDate] = useState(false);

  // ── Meeting sync ──
  const sacrDiff = useCallback((local, remote) => {
    if (!remote || !Array.isArray(remote)) return 0;
    const localDate  = local.filter(r => r.date === activeDate);
    const remoteDate = remote.filter(r => r.date === activeDate);
    if (localDate.length !== remoteDate.length) return Math.abs(localDate.length - remoteDate.length);
    return localDate.filter((r, i) => {
      const rem = remoteDate[i];
      return !rem || r.value !== rem.value || r.label !== rem.label || r.globalOrder !== rem.globalOrder;
    }).length;
  }, [activeDate]);

  const sacrTokenRef = tokenRef || { current: "" };

  const sacrDataRef = useRef(data);
  useEffect(() => { sacrDataRef.current = data; }, [data]);

  const { markDirty: sacrMarkDirty, doSave, saveStatus: sacrSaveStatus,
          pendingCount: sacrPendingCount, applyPending: sacrApplyPending } = useMeetingSync({
    getData:  () => sacrDataRef.current,
    saveFn:   async (tok, d) => { if (saveFn) await saveFn(d); },
    pullFn:   pullFn || null,
    diffFn:   sacrDiff,
    tokenRef: sacrTokenRef,
    onApply:  (remote) => setData(remote),
    enabled:  true,
  });

  const programForDate = data.filter(r => r.date === activeDate);
  const hasProgram = programForDate.length > 0;

  // Sort by globalOrder — the Order column in the sheet is the single source of truth
  const sortedProgram = [...programForDate].sort((a, b) => (a.globalOrder ?? 0) - (b.globalOrder ?? 0));

  const createProgram = () => {
    const newEntries = buildDefaultProgram(activeDate);
    setData(prev => [...prev.filter(r => r.date !== activeDate), ...newEntries]);
    notify.success("Program created for " + formatSundayLabel(activeDate));
  };

  const updateItem = (id, field, value) => {
    setData(prev => prev.map(r => r.id === id ? {...r, [field]: value} : r));
    sacrMarkDirty(id);
  };

  const deleteItem = (id) => {
    setData(prev => prev.filter(r => r.id !== id));
    notify.info("Item removed");
  };

  const addItem = (section, label) => {
    const maxGlobal = programForDate.length > 0 ? Math.max(...programForDate.map(r => r.globalOrder ?? 0)) : -1;
    const newItem = {
      id: `sp_${Date.now()}`,
      date: activeDate,
      section,
      globalOrder: maxGlobal + 1,
      label: label || "",
      value: "",
      notes: "",
    };
    setData(prev => [...prev, newItem]);
  };

  const handleReorder = (reorderedItems) => {
    setData(prev => {
      const otherDates = prev.filter(r => r.date !== activeDate);
      return [...otherDates, ...reorderedItems];
    });
  };

  const addDateProgram = (dateStr) => {
    if (!data.some(r => r.date === dateStr)) {
      const newEntries = buildDefaultProgram(dateStr);
      setData(prev => [...prev, ...newEntries]);
    }
    setActiveDate(dateStr);
  };

  // Use the shared global helper — same 5-Sunday window as Bishopric tab
  const sundayOptions = getSurroundingSundays();
  const sectionsInProgram = [...new Set(sortedProgram.map(r=>r.section))];
  const sectionsNotInProgram = SACRAMENT_SECTIONS.filter(s => !sectionsInProgram.includes(s));

  return (
    <div className="animate-in">
      <HeroBanner title="Sacrament Meeting" sub={hasProgram ? `Program for ${formatSundayLabel(activeDate)}` : "Select or create a program"}>
        {hasProgram && (<>
          <button className="btn-secondary" style={{background:"rgba(255,255,255,.14)",border:"1.5px solid rgba(255,255,255,.3)",color:"#fff"}}
            onClick={()=>setShowPrint(true)}>
            <PrinterIcon/> Print Program
          </button>
          <button className="btn-primary"
            onClick={doSave}
            disabled={sacrSaveStatus==="saving"}
            style={{background:"rgba(255,255,255,.22)",border:"1.5px solid rgba(255,255,255,.5)",
              opacity:sacrSaveStatus==="saving"?.6:1,minWidth:90}}>
            {sacrSaveStatus==="saving" ? "Saving…" : <><SaveIcon/> Save</>}
          </button>
        </>)}
      </HeroBanner>

      {/* Date navigation — 5 Sunday pills matching Bishopric tab */}
      <div style={{display:"flex",gap:6,marginBottom:24,flexWrap:"wrap",alignItems:"center"}}>
        {sundayOptions.map(d => {
          const hasProg = data.some(r=>r.date===d);
          const isActive = d === activeDate;
          const isNext = d === nextSunday();
          return (
            <button key={d} onClick={()=>setActiveDate(d)} style={{
              padding:"6px 14px",borderRadius:20,fontSize:12,
              fontFamily:"'Helvetica Neue',Arial,sans-serif",fontWeight:isActive?700:400,
              border:`1.5px solid ${isActive?C.blue35:hasProg?C.green25:C.border}`,
              background:isActive?C.blue35:hasProg?"rgba(80,168,62,.08)":C.surfaceWarm,
              color:isActive?"#fff":hasProg?C.green35:C.textMuted,
              cursor:"pointer",position:"relative",display:"flex",alignItems:"center",gap:6,
            }}>
              {toDisplayDate(d)}
              {isNext&&!isActive&&<span style={{fontSize:9,fontWeight:700,letterSpacing:".06em",background:C.gold20,color:"#fff",borderRadius:8,padding:"1px 6px"}}>NEXT</span>}
              {isNext&&isActive&&<span style={{fontSize:9,fontWeight:700,letterSpacing:".06em",background:C.blue15,color:C.blue40,borderRadius:8,padding:"1px 6px"}}>NEXT</span>}
            </button>
          );
        })}
        <button onClick={()=>setShowOtherDate(v=>!v)}
          style={{padding:"6px 14px",borderRadius:20,border:`1.5px solid ${C.border}`,background:C.surfaceWhite,color:C.textMuted,fontSize:12,fontFamily:"'Helvetica Neue',Arial,sans-serif",cursor:"pointer"}}>
          Other…
        </button>
        {showOtherDate && (
          <input type="date" value={activeDate}
            onChange={e=>{if(e.target.value){addDateProgram(e.target.value);setShowOtherDate(false);}}}
            style={{width:160,padding:"6px 10px",fontSize:12,borderRadius:7,border:`1px solid ${C.border}`}}/>
        )}
      </div>

      {/* No program yet */}
      {!hasProgram && (
        <div style={{textAlign:"center",padding:"60px 24px",background:C.surfaceWhite,border:`1.5px solid ${C.border}`,borderRadius:12}}>
          <div style={{marginBottom:12,opacity:.3,color:C.textMuted}}><ChurchIcon/></div>
          <div style={{fontFamily:"Georgia,serif",fontSize:17,color:C.textSecond,marginBottom:8}}>No program for {formatSundayLabel(activeDate)}</div>
          <div style={{fontSize:12,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",marginBottom:20}}>Create a program with the standard template.</div>
          <button className="btn-primary" onClick={createProgram}><PlusIcon/> Create Program</button>
        </div>
      )}

      {/* Program — flat draggable list */}
      {hasProgram && (
        <div style={{display:"flex",flexDirection:"column",gap:12}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:4}}>
            <div style={{display:"flex",alignItems:"center",gap:16,flexWrap:"wrap"}}>
            <div style={{fontSize:11,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",fontWeight:600,letterSpacing:".06em",textTransform:"uppercase"}}>
              <GripVertical size={12} style={{display:"inline",verticalAlign:"middle",marginRight:4}}/> Drag rows to reorder
            </div>
            <MeetingSyncBar saveStatus={sacrSaveStatus} pendingCount={sacrPendingCount}
              onApply={sacrApplyPending} onSave={doSave}/>
          </div>
          </div>

          <div style={{background:C.surfaceWhite,border:`1.5px solid ${C.border}`,borderRadius:12,overflow:"hidden"}}>
            {/* Header row */}
            <div style={{display:"grid",gridTemplateColumns:"32px 150px 1fr 1fr 180px 32px",gap:0,padding:"8px 16px",background:C.surfaceWarm,borderBottom:`1px solid ${C.border}`}}>
              <div/>
              <div style={{fontSize:10,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>Section</div>
              <div style={{fontSize:10,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>Label</div>
              <div style={{fontSize:10,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>Name / Details</div>
              <div style={{fontSize:10,fontWeight:700,letterSpacing:".08em",textTransform:"uppercase",color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif"}}>Notes</div>
              <div/>
            </div>
            <DraggableProgramList items={sortedProgram} onReorder={handleReorder} onUpdate={updateItem} onDelete={deleteItem}/>
          </div>

          {/* Add item controls */}
          <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
            <span style={{fontSize:11,color:C.textMuted,fontFamily:"'Helvetica Neue',Arial,sans-serif",fontWeight:600,letterSpacing:".06em",textTransform:"uppercase"}}>Add</span>
            {SACRAMENT_SECTIONS.map(s => {
              const meta = SECTION_META[s]||{};
              return (
                <button key={s} onClick={()=>addItem(s, defaultLabelForSection(s))}
                  style={{padding:"6px 14px",borderRadius:20,fontSize:12,border:`1.5px solid ${meta.border||C.border}`,
                  background:meta.bg||C.surfaceWarm,color:meta.color||C.textMuted,cursor:"pointer",fontFamily:"'Helvetica Neue',Arial,sans-serif",fontWeight:500}}>
                  {meta.icon} {s}
                </button>
              );
            })}
          </div>
        </div>
      )}

      {editingItem && (
        <SacramentItemModal item={editingItem}
          onSave={(updated) => { updateItem(updated.id,"label",updated.label); updateItem(updated.id,"value",updated.value); updateItem(updated.id,"notes",updated.notes); setEditingItem(null); notify.success("Updated"); }}
          onClose={() => setEditingItem(null)}/>
      )}

      {showPrint && <SacramentPrintView program={sortedProgram} date={activeDate} onClose={()=>setShowPrint(false)}/>}
    </div>
  );
}

function DraggableProgramList({ items, onReorder, onUpdate, onDelete }) {
  const dragItem = useRef(null);
  const dragOver = useRef(null);
  const [dragIdx, setDragIdx] = useState(null);

  const handleDragStart = (e, idx) => {
    dragItem.current = idx;
    setDragIdx(idx);
    e.dataTransfer.effectAllowed = "move";
  };
  const handleDragEnd = () => {
    setDragIdx(null);
    if (dragItem.current === null || dragOver.current === null || dragItem.current === dragOver.current) {
      dragItem.current = null; dragOver.current = null; return;
    }
    const next = [...items];
    const [moved] = next.splice(dragItem.current, 1);
    next.splice(dragOver.current, 0, moved);
    onReorder(next.map((item, i) => ({ ...item, globalOrder: i })));
    dragItem.current = null; dragOver.current = null;
  };
  const handleDragOver = (e, idx) => {
    e.preventDefault();
    dragOver.current = idx;
  };

  return (
    <div>
      {items.map((item, idx) => {
        const meta = SECTION_META[item.section]||{icon:"•",color:C.textPrimary,bg:C.surfaceWarm,border:C.border};
        const isDragging = dragIdx === idx;
        return (
          <div key={item.id}
            draggable
            onDragStart={e=>handleDragStart(e,idx)}
            onDragEnd={handleDragEnd}
            onDragOver={e=>handleDragOver(e,idx)}
            style={{
              display:"grid",gridTemplateColumns:"32px 150px 1fr 1fr 180px 32px",
              gap:0,alignItems:"center",
              borderBottom:`1px solid ${C.borderLight}`,
              padding:"7px 16px",
              opacity: isDragging ? 0.4 : 1,
              background: isDragging ? C.surfaceWarm : "transparent",
              transition:"background .1s",
              cursor:"grab",
            }}
            onMouseEnter={e=>{ if(!isDragging) e.currentTarget.style.background=C.surfaceWarm; }}
            onMouseLeave={e=>{ if(!isDragging) e.currentTarget.style.background="transparent"; }}
          >
            {/* Drag handle */}
            <div style={{color:C.textLight,display:"flex",alignItems:"center",cursor:"grab",paddingRight:6}}>
              <DragHandleIcon/>
            </div>
            {/* Section badge */}
            <div style={{display:"flex",alignItems:"center",gap:5,paddingRight:8}}>
              <span style={{fontSize:13}}>{meta.icon}</span>
              <span style={{fontSize:11,fontWeight:600,color:meta.color,fontFamily:"'Helvetica Neue',Arial,sans-serif",
                background:meta.bg,border:`1px solid ${meta.border}`,borderRadius:10,padding:"1px 7px",whiteSpace:"nowrap"}}>
                {item.section}
              </span>
            </div>
            {/* Label inline edit */}
            <input
              value={item.label}
              onChange={e=>onUpdate(item.id,"label",e.target.value)}
              placeholder="Label"
              style={{border:"none",background:"transparent",fontSize:12,fontFamily:"'Helvetica Neue',Arial,sans-serif",
                fontWeight:600,color:meta.color,padding:"2px 4px",borderRadius:4,outline:"none",width:"100%"}}
              onFocus={e=>e.target.style.background=meta.bg}
              onBlur={e=>e.target.style.background="transparent"}
            />
            {/* Value inline edit */}
            <input
              value={item.value}
              onChange={e=>onUpdate(item.id,"value",e.target.value)}
              placeholder={placeholderForLabel(item.label)}
              style={{border:"none",background:"transparent",fontSize:13,fontFamily:"Georgia,serif",color:C.textPrimary,
                padding:"2px 4px",borderRadius:4,outline:"none",width:"100%"}}
              onFocus={e=>e.target.style.background=C.surfaceWarm}
              onBlur={e=>e.target.style.background="transparent"}
            />
            {/* Notes inline edit */}
            <input
              value={item.notes}
              onChange={e=>onUpdate(item.id,"notes",e.target.value)}
              placeholder="Notes..."
              style={{border:"none",background:"transparent",fontSize:12,fontFamily:"'Helvetica Neue',Arial,sans-serif",
                color:C.textMuted,padding:"2px 4px",borderRadius:4,outline:"none",width:"100%"}}
              onFocus={e=>e.target.style.background=C.surfaceWarm}
              onBlur={e=>e.target.style.background="transparent"}
            />
            {/* Delete */}
            <button onClick={()=>onDelete(item.id)} style={{background:"none",border:"none",cursor:"pointer",color:C.textLight,padding:"4px",borderRadius:4,display:"flex",alignItems:"center"}}
              onMouseEnter={e=>e.currentTarget.style.color=C.red15}
              onMouseLeave={e=>e.currentTarget.style.color=C.textLight}>
              <XIcon/>
            </button>
          </div>
        );
      })}
    </div>
  );
}

function defaultLabelForSection(section) {
  const defaults = {
    "Presiding":"Presiding","Prayer":"Prayer","Music":"Hymn",
    "Accompaniment":"Organist","Speaker":"Speaker","Ordinance":"Ordinance","Announcement":"Announcement",
  };
  return defaults[section]||"";
}

function placeholderForLabel(label) {
  const lc = label.toLowerCase();
  if (lc.includes("hymn")||lc.includes("music")) return "Hymn # — Title";
  if (lc.includes("prayer")) return "Name";
  if (lc.includes("speaker")) return "Name — Topic";
  if (lc.includes("presiding")||lc.includes("conducting")) return "Name";
  if (lc.includes("organist")||lc.includes("chorister")) return "Name";
  if (lc.includes("ordinance")||lc.includes("blessing")) return "Name — Type";
  return "Details...";
}

function SacramentItemModal({ item, onSave, onClose }) {
  const [f, setF] = useState({ ...item });
  const s = (k,v) => setF(x=>({...x,[k]:v}));
  return (
    <ModalShell onClose={onClose} title={item.label} subtitle={item.section}>
      <div style={{display:"flex",flexDirection:"column",gap:14}}>
        <FormRow label="Label"><input value={f.label} onChange={e=>s("label",e.target.value)} placeholder="Label"/></FormRow>
        <FormRow label="Value"><input value={f.value} onChange={e=>s("value",e.target.value)} placeholder={placeholderForLabel(f.label)}/></FormRow>
        <FormRow label="Notes"><textarea rows={3} value={f.notes} onChange={e=>s("notes",e.target.value)} style={{resize:"vertical"}}/></FormRow>
      </div>
      <ModalFooter onClose={onClose} onSave={()=>onSave(f)} saveLabel="Save"/>
    </ModalShell>
  );
}

function SacramentPrintView({ program, date, onClose }) {
  // Print respects the globalOrder that was set by dragging
  const sortedForPrint = [...program].sort((a,b) => (a.globalOrder??0) - (b.globalOrder??0));

  // Dynamically collect ALL Presiding and Accompaniment rows in order —
  // so adding a 3rd (or 4th) Presiding item automatically appears in the header.
  const presidingItems    = sortedForPrint.filter(r => r.section === "Presiding");
  const accompanimentItems = sortedForPrint.filter(r => r.section === "Accompaniment");

  // Everything else flows as the body of the program
  const mainItems = sortedForPrint.filter(r => r.section !== "Presiding" && r.section !== "Accompaniment");

  const PRINT_CSS = `
    @media print {
      body * { visibility: hidden; }
      #sacrament-print, #sacrament-print * { visibility: visible; }
      #sacrament-print { position: absolute; left: 0; top: 0; width: 100%; }
      .no-print { display: none !important; }
    }
    #sacrament-print {
      font-family: 'Crimson Pro', Georgia, serif;
      max-width: 680px;
      margin: 0 auto;
      padding: 32px;
      background: white;
      color: #222;
    }
    .sp-header { text-align: center; margin-bottom: 24px; border-bottom: 2px solid #003057; padding-bottom: 16px; }
    .sp-ward { font-size: 13px; font-family: sans-serif; letter-spacing: .1em; text-transform: uppercase; color: #666; margin-bottom: 6px; }
    .sp-title { font-size: 28px; font-weight: 600; color: #003057; margin-bottom: 4px; }
    .sp-date { font-size: 15px; color: #555; }
    .sp-meta { font-size: 12px; font-family: sans-serif; color: #666; margin-top: 8px; line-height: 1.8; }
    .sp-item { display: flex; gap: 16px; padding: 7px 0; align-items: baseline; border-bottom: 1px solid #F0EDE8; }
    .sp-section-badge { font-size: 10px; font-family: sans-serif; letter-spacing:.06em; text-transform:uppercase; color: #999; min-width: 120px; }
    .sp-label { font-size: 12px; font-family: sans-serif; color: #666; min-width: 130px; }
    .sp-value { font-size: 15px; color: #222; flex: 1; }
    .sp-notes { font-size: 12px; color: #888; font-style: italic; }
    .sp-footer { margin-top: 32px; text-align: center; font-size: 11px; font-family: sans-serif; color: #aaa; border-top: 1px solid #eee; padding-top: 12px; }
  `;

  return (
    <div style={{position:"fixed",inset:0,background:"rgba(0,0,0,.6)",zIndex:200,overflow:"auto",padding:"24px"}}>
      <div style={{maxWidth:740,margin:"0 auto"}}>
        <div className="no-print" style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <span style={{color:"#fff",fontFamily:"'Helvetica Neue',Arial,sans-serif",fontSize:13}}>Print Preview — {formatSundayLabel(date)}</span>
          <div style={{display:"flex",gap:8}}>
            <button className="btn-primary" onClick={()=>window.print()} style={{fontSize:13}}><PrinterIcon/> Print</button>
            <button className="btn-secondary" onClick={onClose} style={{fontSize:13,color:"#fff",border:"1.5px solid rgba(255,255,255,.3)",background:"rgba(255,255,255,.1)"}}>Close</button>
          </div>
        </div>

        <style>{PRINT_CSS}</style>
        <div id="sacrament-print" style={{background:"white",borderRadius:8,padding:"40px 48px",boxShadow:"0 8px 40px rgba(0,0,0,.3)"}}>
          <div className="sp-header">
            <div className="sp-ward">{config.WARD_NAME}</div>
            <div className="sp-title">Sacrament Meeting</div>
            <div className="sp-date">{formatSundayLabel(date)}</div>
            <div className="sp-meta">
              {presidingItems.map(r => r.value && (
                <div key={r.id}>{r.label}: {r.value}</div>
              ))}
              {accompanimentItems.map(r => r.value && (
                <div key={r.id}>{r.label}: {r.value}</div>
              ))}
            </div>
          </div>

          {mainItems.map(item => {
            const meta = SECTION_META[item.section]||{};
            return (
              <div key={item.id} className="sp-item">
                <div className="sp-label">{item.label}</div>
                <div style={{flex:1}}>
                  <div className="sp-value">{item.value||<span style={{color:"#ccc",fontStyle:"italic"}}>—</span>}</div>
                  {item.notes&&<div className="sp-notes">{item.notes}</div>}
                </div>
              </div>
            );
          })}

          <div className="sp-footer">The Church of Jesus Christ of Latter-day Saints</div>
        </div>
      </div>
    </div>
  );
}

// ─── SVG Icons ────────────────────────────────────────────────────────────────
function SectionIcon({section,size=13,color}) {
  const Ic = SECTION_ICONS[section];
  if(!Ic) return null;
  return <Ic size={size} color={color||"currentColor"} style={{flexShrink:0}}/>;
}

function buildDefaultProgram(date) {
  const entries = [
    { section:"Presiding",     label:"Presiding",       value:"", notes:"" },
    { section:"Presiding",     label:"Conducting",      value:"", notes:"" },
    { section:"Accompaniment", label:"Organist",        value:"", notes:"" },
    { section:"Accompaniment", label:"Chorister",       value:"", notes:"" },
    { section:"Music",         label:"Opening Hymn",    value:"", notes:"" },
    { section:"Prayer",        label:"Opening Prayer",  value:"", notes:"" },
    { section:"Music",         label:"Sacrament Hymn",  value:"", notes:"" },
    { section:"Speaker",       label:"Speaker",         value:"", notes:"" },
    { section:"Music",         label:"Closing Hymn",    value:"", notes:"" },
    { section:"Prayer",        label:"Closing Prayer",  value:"", notes:"" },
  ];
  return entries.map((e,i) => ({
    id: `sp_new_${date}_${i}`,
    date,
    globalOrder: i,
    ...e,
  }));
}

function nextSunday() {
  // Use getUpcomingSunday() so both tabs share identical Sunday logic
  return getUpcomingSunday();
}

function formatSundayLabel(dateStr) {
  if (!dateStr) return "";
  const d = new Date(dateStr + "T12:00:00");
  return d.toLocaleDateString("en-US", { month:"long", day:"numeric", year:"numeric" });
}

// ── Drag-and-drop hook for sacrament program reordering ──
function useDragReorder(items, onReorder) {
  const dragItem = useRef(null);
  const dragOver = useRef(null);

  const handleDragStart = (e, idx) => {
    dragItem.current = idx;
    e.dataTransfer.effectAllowed = "move";
    e.currentTarget.style.opacity = "0.5";
  };
  const handleDragEnd = (e) => {
    e.currentTarget.style.opacity = "1";
    if (dragItem.current === null || dragOver.current === null || dragItem.current === dragOver.current) {
      dragItem.current = null; dragOver.current = null; return;
    }
    const next = [...items];
    const [moved] = next.splice(dragItem.current, 1);
    next.splice(dragOver.current, 0, moved);
    onReorder(next.map((item, i) => ({ ...item, globalOrder: i })));
    dragItem.current = null; dragOver.current = null;
  };
  const handleDragOver = (e, idx) => {
    e.preventDefault();
    dragOver.current = idx;
  };
  const handleDragEnter = (e) => {
    e.currentTarget.style.borderTop = `2px solid ${C.blue35}`;
  };
  const handleDragLeave = (e) => {
    e.currentTarget.style.borderTop = "";
  };
  const handleDrop = (e) => {
    e.currentTarget.style.borderTop = "";
  };

  return { handleDragStart, handleDragEnd, handleDragOver, handleDragEnter, handleDragLeave, handleDrop };
}

function PlusIcon(){return<Plus size={12}/>;}
function DragHandleIcon(){return<GripVertical size={14}/>;}
function XIcon(){return<X size={12}/>;}
function TableIcon(){return<LayoutList size={13}/>;}
function KanbanIcon(){return<Columns2 size={13}/>;}
function PrinterIcon(){return<Printer size={13}/>;}
function SaveIcon(){return<Save size={13}/>;}
function LockIcon(){return<Lock size={36}/>;}
function ChurchIcon(){return<Building2 size={36}/>;}
function ClipboardIcon(){return<ClipboardList size={36}/>;}
function ChevLeftIcon(){return<ChevronLeft size={14}/>;}
function ChevRightIcon(){return<ChevronRight size={14}/>;}
function PullIcon(){return<RefreshCw size={11}/>;}
function PushIcon(){return<Save size={11}/>;}
