import React, { useState, useMemo, useEffect, useCallback } from "react";
import {
  FileText, Plus, Save, Printer, FolderOpen, Settings, ChevronDown,
  User, Wallet, Home, Briefcase, TrendingUp, Coins, ShieldCheck,
  Scale, Receipt, CheckCircle2, AlertCircle, Trash2, Copy, X, Search, LayoutGrid, GripVertical,
  FileSpreadsheet, FileDown
} from "lucide-react";
import * as XLSX from "xlsx";
import * as XLSXStyle from "xlsx-js-style";
import { supabase } from "./supabaseClient";

// Mock window.storage for browser environment using localStorage
if (typeof window !== "undefined" && !window.storage) {
  window.storage = {
    get: async (key) => {
      try {
        const val = localStorage.getItem(key);
        return val ? { value: val } : null;
      } catch {
        return null;
      }
    },
    set: async (key, value) => {
      try {
        localStorage.setItem(key, value);
      } catch (e) {
        console.error("LocalStorage write failed:", e);
      }
    }
  };
}

// ─── Theme ────────────────────────────────────────────────────────────────────
// G is updated dynamically by App theme selection
let G = {
  bg:"#060E0A", surf:"#0B1610", card:"#0F1D14", bdr:"#1A3020",
  green:"#22C55E", g2:"#16A34A", g3:"#4ADE80", gd:"#14532D",
  red:"#EF4444", amb:"#F59E0B", cyn:"#06B6D4", vio:"#8B5CF6", ind:"#6366F1",
  wh:"#F0FDF4", txt:"#DCFCE7", mut:"#86EFAC", bdr2:"#223829",
};

// ─── Passwords & dropdown defaults ───────────────────────────────────────────
const DEF_PW = { owner:"456", dev:"123" };
const DARK_THEME = {
  bg:"#060E0A", surf:"#0B1610", card:"#0F1D14", bdr:"#1A3020", border2:"#223829",
  green:"#22C55E", g2:"#16A34A", g3:"#4ADE80", gd:"#14532D",
  red:"#EF4444", amb:"#F59E0B", cyn:"#06B6D4", vio:"#8B5CF6", ind:"#6366F1",
  wh:"#F0FDF4", txt:"#DCFCE7", mut:"#86EFAC", bdr2:"#223829",
};
const LIGHT_THEME = {
  bg:"#F1F5F1",      // page background - soft grey-green
  surf:"#FFFFFF",    // card surface - white
  card:"#F7FBF7",    // inner card - very light
  bdr:"#C8E6C9",     // border - medium green
  border2:"#A5D6A7", // border accent
  green:"#2E7D32",   // primary green - dark readable
  g2:"#1B5E20",      // dark green
  g3:"#388E3C",      // mid green
  gd:"#E8F5E9",      // green tint bg
  red:"#C62828",     // error red - dark
  amb:"#E65100",     // amber - dark orange
  cyn:"#00695C",     // teal - dark
  vio:"#4527A0",     // violet - dark
  ind:"#283593",     // indigo - dark
  wh:"#1A1A1A",      // PRIMARY TEXT - near black (was dark green - invisible!)
  txt:"#212121",     // body text - near black
  mut:"#424242",     // muted text - dark grey (was green - hard to read!)
  bdr2:"#81C784",    // soft border
};

const DEF_DD = {
  services:["ITR Filing","GST Return Filing","GST Registration","TDS Return","Tax Audit","Balance Sheet & P&L","MSME Registration","ROC Filing","Accounting","Income Tax Notice","Other"],
  states:["Andhra Pradesh","Assam","Bihar","Chhattisgarh","Goa","Gujarat","Haryana","Himachal Pradesh","Jharkhand","Karnataka","Kerala","Madhya Pradesh","Maharashtra","Odisha","Punjab","Rajasthan","Tamil Nadu","Telangana","Uttar Pradesh","Uttarakhand","West Bengal","Delhi","Puducherry"],
  sources:["Referral","Walk-in","Online","Social Media","Advertisement","CA Reference","Other"],
  clientTypes:["Individual","Proprietorship","Company","LLP","Partnership","HUF","Trust"],
  fyOptions:["2025-26","2024-25","2023-24","2022-23"],
  staff:["Amit Singh","Neha Gupta","Ravi Patel","Sonal Mehta","Deepak Joshi"],
  billingItems:[
    "ITR Filing|1500",
    "GST Return Filing|1000",
    "GST Registration|2000",
    "TDS Return|1500",
    "Tax Audit|5000",
    "Balance Sheet & P&L|3000",
    "MSME Registration|2500",
    "ROC Filing|3000",
    "Accounting|2000",
    "Income Tax Notice|2000",
  ],
  pwTypes:["GST Portal","Income Tax e-Filing","TRACES","TDS / CPC","MCA / ROC","EPFO","Email","Net Banking","DSC / Digital Signature","Software License","Other"],
};
const PORTALS = [
  {key:"gst",   label:"GST Portal",  icon:"📦", col:"#22C55E"},
  {key:"it",    label:"Income Tax",  icon:"🏛",  col:"#6366F1"},
  {key:"traces",label:"TRACES",      icon:"📄",  col:"#06B6D4"},
  {key:"tds",   label:"TDS / CPC",   icon:"🔢",  col:"#F59E0B"},
  {key:"mca",   label:"MCA / ROC",   icon:"🏢",  col:"#8B5CF6"},
  {key:"epf",   label:"EPF / PF",    icon:"👷",  col:"#EF4444"},
];
const mkP = () => Object.fromEntries(PORTALS.map(p=>[p.key,{id:"",pw:"",on:false}]));

// ─── Seed ─────────────────────────────────────────────────────────────────────
const SC = [
  {pan:"ABCPK1234A",name:"Rajesh Kumar",biz:"Rajesh Enterprises",mob:"9876543210",email:"rajesh@email.com",gstin:"27ABCPK1234A1Z5",addr:"123 MG Road, Mumbai",state:"Maharashtra",type:"Proprietorship",src:"Referral",added:"2024-01-15",status:"Active",note:"Regular",aadhaar:"",dob:"",fatherName:"",gender:"",residentialStatus:"",pin:"",bankName:"",ifsc:"",accountNumber:"",itrType:"",extraPw:[],portals:{gst:{id:"raj_gst",pw:"G@123",on:true},it:{id:"ABCPK1234A",pw:"I@456",on:true},traces:{id:"raj_tr",pw:"T@789",on:true},tds:{id:"",pw:"",on:false},mca:{id:"",pw:"",on:false},epf:{id:"",pw:"",on:false}}},
  {pan:"BCDQL5678B",name:"Priya Sharma",biz:"",mob:"9812345678",email:"priya@gmail.com",gstin:"",addr:"Flat 4B, Delhi",state:"Delhi",type:"Individual",src:"Walk-in",added:"2024-02-20",status:"Active",note:"",aadhaar:"",dob:"",fatherName:"",gender:"",residentialStatus:"",pin:"",bankName:"",ifsc:"",accountNumber:"",itrType:"",extraPw:[],portals:{gst:{id:"",pw:"",on:false},it:{id:"BCDQL5678B",pw:"P@IT",on:true},traces:{id:"",pw:"",on:false},tds:{id:"",pw:"",on:false},mca:{id:"",pw:"",on:false},epf:{id:"",pw:"",on:false}}},
  {pan:"CDERM7890C",name:"Tech Solutions Pvt Ltd",biz:"Tech Solutions Pvt Ltd",mob:"9988776655",email:"info@tech.com",gstin:"29CDERM7890C1ZA",addr:"45 IT Park, Bengaluru",state:"Karnataka",type:"Company",src:"Online",added:"2024-03-10",status:"Active",note:"Priority",aadhaar:"",dob:"",fatherName:"",gender:"",residentialStatus:"",pin:"",bankName:"",ifsc:"",accountNumber:"",itrType:"",extraPw:[],portals:{gst:{id:"t_gst",pw:"TG@1",on:true},it:{id:"CDERM7890C",pw:"TI@2",on:true},traces:{id:"t_tr",pw:"TT@3",on:true},tds:{id:"t_tds",pw:"TD@4",on:true},mca:{id:"t_mca",pw:"TM@5",on:true},epf:{id:"",pw:"",on:false}}},
  {pan:"DEFNS2345D",name:"Mehta & Associates",biz:"Mehta & Associates LLP",mob:"8877665544",email:"mehta@assoc.com",gstin:"24DEFNS2345D1ZB",addr:"78 CG Road, Ahmedabad",state:"Gujarat",type:"LLP",src:"Referral",added:"2024-04-05",status:"Active",note:"",aadhaar:"",dob:"",fatherName:"",gender:"",residentialStatus:"",pin:"",bankName:"",ifsc:"",accountNumber:"",itrType:"",extraPw:[],portals:{gst:{id:"m_gst",pw:"MG@1",on:true},it:{id:"DEFNS2345D",pw:"MI@2",on:true},traces:{id:"",pw:"",on:false},tds:{id:"",pw:"",on:false},mca:{id:"m_mca",pw:"MM@3",on:true},epf:{id:"",pw:"",on:false}}},
  {pan:"EFGOT6789E",name:"Sunita Patel",biz:"Patel Traders",mob:"7766554433",email:"sunita@yahoo.com",gstin:"08EFGOT6789E1ZC",addr:"32 Linking Rd, Mumbai",state:"Maharashtra",type:"Partnership",src:"Walk-in",added:"2024-05-12",status:"Inactive",note:"On hold",aadhaar:"",dob:"",fatherName:"",gender:"",residentialStatus:"",pin:"",bankName:"",ifsc:"",accountNumber:"",itrType:"",extraPw:[],portals:{gst:{id:"p_gst",pw:"PG@1",on:true},it:{id:"EFGOT6789E",pw:"PI@2",on:true},traces:{id:"",pw:"",on:false},tds:{id:"",pw:"",on:false},mca:{id:"",pw:"",on:false},epf:{id:"",pw:"",on:false}}},
];
const SW = [
  {id:1,pan:"ABCPK1234A",cn:"Rajesh Kumar",svc:"ITR Filing",fy:"2024-25",due:"2024-07-31",staff:"Amit Singh",status:"Completed",fees:5000,comm:500,rcvd:5000,src:"Referral",note:"Done"},
  {id:2,pan:"BCDQL5678B",cn:"Priya Sharma",svc:"GST Return Filing",fy:"2024-25",due:"2024-06-20",staff:"Neha Gupta",status:"In Progress",fees:3000,comm:300,rcvd:0,src:"Walk-in",note:""},
  {id:3,pan:"CDERM7890C",cn:"Tech Solutions Pvt Ltd",svc:"Tax Audit",fy:"2024-25",due:"2024-09-30",staff:"Amit Singh",status:"Pending",fees:25000,comm:2500,rcvd:12500,src:"Online",note:"Priority"},
  {id:4,pan:"ABCPK1234A",cn:"Rajesh Kumar",svc:"TDS Return",fy:"2024-25",due:"2024-06-15",staff:"Neha Gupta",status:"Pending",fees:2000,comm:200,rcvd:0,src:"Referral",note:""},
  {id:5,pan:"DEFNS2345D",cn:"Mehta & Associates",svc:"GST Registration",fy:"2025-26",due:"2025-06-25",staff:"Ravi Patel",status:"Completed",fees:1500,comm:150,rcvd:1500,src:"Referral",note:""},
  {id:6,pan:"CDERM7890C",cn:"Tech Solutions Pvt Ltd",svc:"GST Return Filing",fy:"2025-26",due:"2025-07-20",staff:"Amit Singh",status:"In Progress",fees:4000,comm:400,rcvd:2000,src:"Online",note:""},
];

// ─── Utils ────────────────────────────────────────────────────────────────────
const td = () => new Date().toISOString().split("T")[0];
const fd = d => d ? new Date(d).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"}) : "-";
const isOD = w => w.status!=="Completed" && new Date(w.due)<new Date();
const getCurrentFY = () => {
  const now = new Date();
  const yr = now.getFullYear();
  return now.getMonth() >= 3 ? `${yr}-${String(yr + 1).slice(-2)}` : `${yr - 1}-${String(yr).slice(-2)}`;
};
// ─── Work ↔ Invoice ↔ Receipt linkage helpers ──────────────────────────────
// A work can have one or more invoices raised against it (invoice.workId === work.id).
// Each invoice can have one or more payment receipts (receipt.invId === invoice.id).
// "Received" for a work is ALWAYS the live rollup of actual receipts recorded against
// its linked invoice(s) — never a number typed in directly. This is the single source
// of truth used everywhere (Work Tracker, Finance Dashboard, Invoice status).
const linkedInvoices = (w, invoices) => (invoices||[]).filter(inv=>inv.workId!==undefined&&inv.workId!==""&&String(inv.workId)===String(w.id));
const invoiceReceived = (inv, receipts) => (receipts||[]).filter(r=>r.invId===inv.id).reduce((s,r)=>s+(r.received||0),0);
// Correct status for one invoice, derived purely from receipts vs its total — this is
// what stops an invoice being marked "Paid" from a single partial entry, and what
// keeps status correct after a receipt is edited/deleted.
const computeInvStatus = (inv, receipts) => {
  const rcvd = invoiceReceived(inv, receipts), total = inv.total||0;
  if(total<=0) return "Unpaid";
  if(rcvd<=0) return "Unpaid";
  if(rcvd>=total) return "Paid";
  return "Partial";
};
const workReceived = (w, invoices, receipts) => {
  const linked = linkedInvoices(w, invoices);
  if(!linked.length) return w.rcvd||0; // legacy fallback: work never invoiced yet
  return linked.reduce((s,inv)=>s+invoiceReceived(inv, receipts),0);
};
const workBilled = (w, invoices) => {
  const linked = linkedInvoices(w, invoices);
  return linked.length ? linked.reduce((s,inv)=>s+(inv.total||0),0) : (w.fees||0);
};
const workPayStatus = (w, invoices, receipts) => {
  const linked = linkedInvoices(w, invoices);
  if(!linked.length) return (w.rcvd||0)>0 ? "Unlinked" : "No Invoice";
  const billed=workBilled(w,invoices), rcvd=workReceived(w,invoices,receipts);
  if(rcvd<=0) return "Unpaid";
  if(rcvd>=billed) return "Paid";
  return "Partial";
};
const odDays = w => { if(w.status==="Completed") return 0; const d=Math.floor((Date.now()-new Date(w.due))/86400000); return d>0?d:0; };
const inr = n => "₹"+Number(n||0).toLocaleString("en-IN");
const IS = { background:G.card, border:"1.5px solid #1A3020", borderRadius:8, color:G.wh, fontSize:13, padding:"9px 12px", outline:"none", width:"100%", boxSizing:"border-box", fontFamily:"Segoe UI,sans-serif" };

// ─── Mini components ──────────────────────────────────────────────────────────
function Logo({sz=36}){
  // Fin-Tax Mitra branded logo - circular green design matching brand identity
  const r=sz/2;
  return <svg width={sz} height={sz} viewBox="0 0 100 100" fill="none">
    {/* Outer ring - brand green */}
    <circle cx="50" cy="50" r="48" fill="#1A3020" stroke="#22C55E" strokeWidth="4"/>
    <circle cx="50" cy="50" r="42" fill="#F0F9F0" stroke="#22C55E" strokeWidth="1.5"/>
    {/* Document/scroll icon */}
    <rect x="18" y="26" width="26" height="32" rx="4" fill="#22C55E"/>
    <rect x="22" y="32" width="14" height="2.5" rx="1" fill="white" opacity=".9"/>
    <rect x="22" y="37" width="17" height="2" rx="1" fill="white" opacity=".7"/>
    <rect x="22" y="42" width="12" height="2" rx="1" fill="white" opacity=".7"/>
    <circle cx="24" cy="53" r="4" fill="white" opacity=".8"/>
    <text x="21" y="56" fontSize="6" fill="#16A34A" fontWeight="bold">✕</text>
    {/* Vertical divider */}
    <rect x="46" y="24" width="2.5" height="52" rx="1.2" fill="#22C55E"/>
    {/* Calculator */}
    <rect x="51" y="30" width="26" height="32" rx="4" fill="#16A34A"/>
    <rect x="55" y="34" width="18" height="8" rx="2" fill="white" opacity=".9"/>
    {[0,1,2].map(c=>[0,1,2,3].map(r2=>(
      <rect key={c*10+r2} x={55+c*6.5} y={46+r2*5} width="5" height="3.5" rx="1" fill="white" opacity=".75"/>
    )))}
    {/* Fin-Tax text */}
    <text x="50" y="84" fontSize="10" fill="#1B5E20" fontWeight="900" textAnchor="middle" fontFamily="Arial,sans-serif">Fin-Tax</text>
    <text x="50" y="95" fontSize="8" fill="#22C55E" fontWeight="700" textAnchor="middle" fontFamily="Arial,sans-serif">mitra</text>
  </svg>;
}

function Bdg({label}){
  const m={Active:{b:"#14532D",f:"#4ADE80"},Inactive:{b:"#450A0A",f:"#FCA5A5"},Completed:{b:"#14532D",f:"#4ADE80"},"In Progress":{b:"#1E1B4B",f:"#A5B4FC"},Pending:{b:"#431407",f:"#FCD34D"},Overdue:{b:"#450A0A",f:"#FCA5A5"}};
  const s=m[label]||{b:G.bdr,f:G.mut};
  return <span style={{background:s.b,color:s.f,padding:"2px 9px",borderRadius:20,fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{label}</span>;
}

function Toasts({list}){
  return <div style={{position:"fixed",top:20,right:20,zIndex:9999,display:"flex",flexDirection:"column",gap:8}}>
    {list.map(t=><div key={t.id} style={{background:t.tp==="ok"?G.g2:t.tp==="err"?G.red:G.ind,color:"#fff",padding:"10px 18px",borderRadius:10,fontSize:13,fontWeight:600,boxShadow:"0 4px 20px #0006",animation:"sIn .25s ease"}}>
      {t.tp==="ok"?"✓":t.tp==="err"?"✕":"ℹ"} {t.msg}
    </div>)}
  </div>;
}

function F({label,req,children,w="100%"}){
  return <div style={{display:"flex",flexDirection:"column",gap:4,width:w,minWidth:0}}>
    <label style={{fontSize:11,fontWeight:700,color:G.mut,textTransform:"uppercase",letterSpacing:.7}}>{label}{req&&<span style={{color:G.red}}> *</span>}</label>
    {children}
  </div>;
}
function I({val,set,ph,type="text",sty={},mono,ro,kd,onFocus:ofn,onBlur:obn}){
  const[foc,setFoc]=useState(false);
  return <input type={type} value={val} onChange={e=>set&&set(e.target.value)} placeholder={ph} readOnly={ro} onKeyDown={kd}
    onFocus={e=>{setFoc(true);ofn&&ofn(e);}} onBlur={e=>{setFoc(false);obn&&obn(e);}}
    style={{...IS,borderColor:foc?G.green:G.bdr,background:ro?"#060E0A":G.card,fontFamily:mono?"Courier New,monospace":"Segoe UI,sans-serif",...sty}}/>;
}
function S({val,set,opts,ph}){
  return <select value={val} onChange={e=>set(e.target.value)} style={{...IS,cursor:"pointer",color:val?G.wh:G.mut}}>
    {ph&&<option value="">{ph}</option>}
    {opts.map(o=><option key={o} value={o} style={{background:"#0B1610"}}>{o}</option>)}
  </select>;
}
function R({children,gap=12}){return <div style={{display:"flex",gap,flexWrap:"wrap"}}>{children}</div>;}
function Crd({children,sty={}}){return <div style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:16,padding:20,...sty}}>{children}</div>;}
function SH({icon,title,acc=G.green}){
  return <div style={{display:"flex",alignItems:"center",gap:10,paddingBottom:11,borderBottom:`1px solid ${G.bdr}`,marginBottom:12}}>
    <span style={{background:acc+"22",color:acc,width:30,height:30,borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",fontSize:15,flexShrink:0}}>{icon}</span>
    <span style={{fontWeight:700,color:G.wh,fontSize:14}}>{title}</span>
  </div>;
}
function Btn({children,onClick,sty={},color=G.green,color2=G.g2}){
  return <button onClick={onClick} style={{padding:"10px 18px",borderRadius:10,border:"none",cursor:"pointer",background:`linear-gradient(135deg,${color2},${color})`,color:"#fff",fontWeight:700,fontSize:13,...sty}}>{children}</button>;
}

// ─── Auth Modal ───────────────────────────────────────────────────────────────
function Auth({title,hint,pws,onOk,onX}){
  const[pw,setPw]=useState(""),[err,setErr]=useState("");
  const go=()=>{if(Object.values(pws).includes(pw)){onOk();}else{setErr("Wrong password");setPw("");}};
  return <div style={{position:"fixed",inset:0,background:"#000D",zIndex:10000,display:"flex",alignItems:"center",justifyContent:"center"}}>
    <div style={{background:G.surf,border:`1px solid ${G.green}44`,borderRadius:20,padding:36,width:360,boxShadow:"0 32px 80px #000C"}}>
      <div style={{textAlign:"center",marginBottom:20}}><Logo sz={52}/><div style={{fontWeight:800,fontSize:17,color:G.wh,marginTop:12}}>{title}</div></div>
      <F label="Password" req>
        <I val={pw} set={v=>{setPw(v);setErr("");}} ph="Enter password..." type="password" kd={e=>e.key==="Enter"&&go()}/>
        {err&&<span style={{fontSize:11,color:G.red,marginTop:2,display:"block"}}>{err}</span>}
      </F>
      <div style={{display:"flex",gap:10,marginTop:14}}>
        <Btn onClick={go} sty={{flex:1,padding:12,fontSize:14}}>Unlock</Btn>
        <button onClick={onX} style={{padding:"12px 18px",borderRadius:10,border:`1px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer",fontWeight:600}}>Cancel</button>
      </div>
      {hint&&<div style={{marginTop:10,fontSize:11,color:"#334155",textAlign:"center"}}>{hint}</div>}
    </div>
  </div>;
}

// ─── PAN Picker ───────────────────────────────────────────────────────────────
function PanPick({clients,val,set}){
  const[open,setOpen]=useState(false),[q,setQ]=useState("");
  const sel=clients.find(c=>c.pan===val);
  const list=useMemo(()=>{
    const base=clients.filter(c=>c.status==="Active");
    if(!q)return base.slice(0,8);
    const lq=q.toLowerCase();
    return base.filter(c=>c.pan.toLowerCase().includes(lq)||c.name.toLowerCase().includes(lq)||c.mob.includes(q)).slice(0,8);
  },[q,clients]);
  return <div style={{position:"relative"}}>
    <div style={{display:"flex",gap:8}}>
      <I val={q} set={v=>{setQ(v);setOpen(true);if(val)set("");}} ph="Type PAN / Name / Mobile..." sty={{flex:1}} onFocus={()=>setOpen(true)}/>
      {sel&&<button onClick={()=>{setQ("");set("");}} style={{background:"#450A0A",border:`1.5px solid ${G.red}55`,color:G.red,borderRadius:8,padding:"0 12px",cursor:"pointer",fontSize:13}}>✕</button>}
    </div>
    {sel&&<div style={{marginTop:8,background:`${G.green}08`,border:`1.5px solid ${G.green}30`,borderRadius:11,padding:"11px 14px",display:"flex",gap:12,alignItems:"center"}}>
      <div style={{width:40,height:40,borderRadius:10,background:`linear-gradient(135deg,${G.g2},${G.green})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,fontWeight:800,color:"#fff",flexShrink:0}}>{sel.name[0]}</div>
      <div style={{flex:1,minWidth:0}}>
        <div style={{fontWeight:700,fontSize:13,color:G.wh}}>{sel.name}</div>
        {sel.biz&&<div style={{fontSize:11,color:G.mut}}>{sel.biz}</div>}
        <div style={{display:"flex",gap:8,marginTop:3,flexWrap:"wrap"}}>
          <span style={{fontSize:11,color:G.green,fontWeight:700,fontFamily:"monospace"}}>🆔 {sel.pan}</span>
          <span style={{fontSize:11,color:G.mut}}>📱 {sel.mob}</span>
          <span style={{fontSize:11,color:G.mut}}>📍 {sel.state}</span>
          <Bdg label={sel.status}/>
        </div>
      </div>
    </div>}
    {open&&!sel&&<>
      <div style={{position:"fixed",inset:0,zIndex:998}} onClick={()=>setOpen(false)}/>
      <div style={{position:"absolute",top:"calc(100% + 5px)",left:0,right:0,zIndex:999,background:"#0F1D14",border:`1.5px solid ${G.bdr}`,borderRadius:11,overflow:"hidden",boxShadow:"0 20px 60px #000A"}}>
        {list.length===0?<div style={{padding:16,color:G.mut,textAlign:"center",fontSize:13}}>No active clients found</div>
        :list.map(c=><div key={c.pan} onClick={()=>{set(c.pan);setQ(c.pan);setOpen(false);}}
          style={{padding:"9px 13px",cursor:"pointer",borderBottom:`1px solid ${G.bg}`,display:"flex",gap:11,alignItems:"center"}}
          onMouseEnter={e=>e.currentTarget.style.background=G.bg} onMouseLeave={e=>e.currentTarget.style.background="transparent"}>
          <div style={{width:34,height:34,borderRadius:8,background:`${G.green}20`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:14,fontWeight:800,color:G.green,flexShrink:0}}>{c.name[0]}</div>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontWeight:600,color:G.txt,fontSize:13}}>{c.name}</div>
            <div style={{fontSize:11,color:G.mut}}>{c.pan} · {c.mob} · {c.state}</div>
          </div>
          <Bdg label={c.status}/>
        </div>)}
      </div>
    </>}
  </div>;
}

// ─── Edit Client Modal ────────────────────────────────────────────────────────
function EditClient({c,onSave,onX,dd}){
  const[f,setF]=useState({...c,extraPw:c.extraPw||[]});
  const[spw,setSpw]=useState({});
  const sp=(k,fld,v)=>setF(p=>({...p,portals:{...p.portals,[k]:{...p.portals[k],[fld]:v}}}));
  const tp=(k)=>setF(p=>({...p,portals:{...p.portals,[k]:{...p.portals[k],on:!p.portals[k].on}}}));
  const addPw=()=>setF(p=>({...p,extraPw:[...(p.extraPw||[]),{type:dd.pwTypes?.[0]||"Other",label:"",id:"",pw:""}]}));
  const setPwAt=(i,patch)=>setF(p=>({...p,extraPw:p.extraPw.map((x,j)=>j===i?{...x,...patch}:x)}));
  const delPwAt=(i)=>setF(p=>({...p,extraPw:p.extraPw.filter((_,j)=>j!==i)}));
  return <div style={{position:"fixed",inset:0,background:"#000C",zIndex:5000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
    <div style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:18,width:"min(680px,98%)",maxHeight:"90vh",overflow:"auto",boxShadow:"0 32px 80px #000C"}}>
      <div style={{padding:"16px 20px",borderBottom:`1px solid ${G.bdr}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <span style={{fontWeight:800,fontSize:15,color:G.wh}}>✏️ Edit - {f.pan}</span>
        <button onClick={onX} style={{background:"transparent",border:`1px solid ${G.bdr}`,color:G.mut,borderRadius:8,padding:"4px 11px",cursor:"pointer"}}>✕</button>
      </div>
      <div style={{padding:20,display:"flex",flexDirection:"column",gap:11}}>
        <R><F label="Name" w="calc(50% - 6px)"><I val={f.name} set={v=>setF({...f,name:v})}/></F><F label="Business" w="calc(50% - 6px)"><I val={f.biz} set={v=>setF({...f,biz:v})}/></F></R>
        <R><F label="Mobile" w="calc(50% - 6px)"><I val={f.mob} set={v=>setF({...f,mob:v})}/></F><F label="Email" w="calc(50% - 6px)"><I val={f.email} set={v=>setF({...f,email:v})}/></F></R>
        <R><F label="GSTIN" w="calc(50% - 6px)"><I val={f.gstin} set={v=>setF({...f,gstin:v.toUpperCase()})} mono/></F><F label="Source" w="calc(50% - 6px)"><S val={f.src} set={v=>setF({...f,src:v})} opts={dd.sources}/></F></R>
        <R><F label="State" w="calc(50% - 6px)"><S val={f.state} set={v=>setF({...f,state:v})} opts={dd.states}/></F><F label="Status" w="calc(50% - 6px)"><S val={f.status} set={v=>setF({...f,status:v})} opts={["Active","Inactive"]}/></F></R>
        <F label="Address"><I val={f.addr} set={v=>setF({...f,addr:v})}/></F>
        <F label="Remarks"><textarea value={f.note} onChange={e=>setF({...f,note:e.target.value})} rows={2} style={{...IS,resize:"vertical"}}/></F>
        <SH icon="🧾" title="ITR Profile Details (optional)" acc={G.cyn}/>
        <R><F label="Father's Name" w="calc(50% - 6px)"><I val={f.fatherName||""} set={v=>setF({...f,fatherName:v})}/></F>
        <F label="Date of Birth" w="calc(50% - 6px)"><I val={f.dob||""} set={v=>setF({...f,dob:v})} type="date"/></F></R>
        <R><F label="Gender" w="calc(50% - 6px)"><S val={f.gender||""} set={v=>setF({...f,gender:v})} opts={["Male","Female","Other"]} ph="Select..."/></F>
        <F label="Residential Status" w="calc(50% - 6px)"><S val={f.residentialStatus||""} set={v=>setF({...f,residentialStatus:v})} opts={["Resident","Resident but Not Ordinarily Resident","Non-Resident"]} ph="Select..."/></F></R>
        <R><F label="Aadhaar No." w="calc(50% - 6px)"><I val={f.aadhaar||""} set={v=>setF({...f,aadhaar:v.replace(/\D/g,"").slice(0,12)})} mono ph="12-digit number"/></F>
        <F label="PIN Code" w="calc(50% - 6px)"><I val={f.pin||""} set={v=>setF({...f,pin:v.replace(/\D/g,"").slice(0,6)})} mono/></F></R>
        <R><F label="Bank Name" w="calc(50% - 6px)"><I val={f.bankName||""} set={v=>setF({...f,bankName:v})}/></F>
        <F label="Bank IFSC" w="calc(50% - 6px)"><I val={f.ifsc||""} set={v=>setF({...f,ifsc:v.toUpperCase()})} mono ph="SBIN0001234"/></F></R>
        <R><F label="Bank Account No." w="calc(50% - 6px)"><I val={f.accountNumber||""} set={v=>setF({...f,accountNumber:v})} mono/></F>
        <F label="ITR Type" w="calc(50% - 6px)"><S val={f.itrType||""} set={v=>setF({...f,itrType:v})} opts={["ITR-1","ITR-2","ITR-3","ITR-4"]} ph="Select..."/></F></R>
        <SH icon="🔑" title="Other Passwords (optional)" acc={G.cyn}/>
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {(f.extraPw||[]).map((pwItem,i)=><div key={i} style={{display:"flex",gap:8,alignItems:"flex-end",padding:"9px 10px",background:G.bg,border:`1px solid ${G.bdr}`,borderRadius:9}}>
            <F label="Type" w="22%"><S val={pwItem.type} set={v=>setPwAt(i,{type:v})} opts={dd.pwTypes||[]} ph="Select..."/></F>
            <F label="Label" w="26%"><I val={pwItem.label} set={v=>setPwAt(i,{label:v})} ph="e.g. Personal Email"/></F>
            <F label="User ID" w="24%"><I val={pwItem.id} set={v=>setPwAt(i,{id:v})} mono/></F>
            <F label="Password" w="20%"><I val={pwItem.pw} set={v=>setPwAt(i,{pw:v})} type="password" mono/></F>
            <button onClick={()=>delPwAt(i)} style={{background:"#450A0A",border:`1px solid ${G.red}44`,color:G.red,borderRadius:8,padding:"9px 10px",cursor:"pointer",fontSize:12,flexShrink:0}}>🗑</button>
          </div>)}
          <button onClick={addPw} style={{alignSelf:"flex-start",padding:"7px 14px",borderRadius:8,border:`1px dashed ${G.bdr}`,background:"transparent",color:G.cyn,cursor:"pointer",fontSize:12,fontWeight:600}}>➕ Add Password</button>
        </div>
        <SH icon="🔐" title="Portal Credentials" acc={G.vio}/>
        {PORTALS.map(p=><div key={p.key}>
          <label style={{display:"flex",alignItems:"center",gap:9,cursor:"pointer",marginBottom:f.portals[p.key].on?8:0}}>
            <div onClick={()=>tp(p.key)} style={{width:19,height:19,borderRadius:5,border:`2px solid ${f.portals[p.key].on?p.col:G.bdr}`,background:f.portals[p.key].on?p.col+"20":"transparent",display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",flexShrink:0}}>
              {f.portals[p.key].on&&<span style={{color:p.col,fontSize:11,fontWeight:900,lineHeight:1}}>✓</span>}
            </div>
            <span style={{fontSize:13,color:f.portals[p.key].on?p.col:G.mut,fontWeight:f.portals[p.key].on?700:400}}>{p.icon} {p.label}</span>
          </label>
          {f.portals[p.key].on&&<div style={{display:"flex",gap:10,paddingLeft:27}}>
            <F label="User ID" w="calc(50% - 5px)"><I val={f.portals[p.key].id} set={v=>sp(p.key,"id",v)} mono sty={{borderColor:p.col+"55"}}/></F>
            <F label="Password" w="calc(50% - 5px)"><div style={{position:"relative"}}>
              <I val={f.portals[p.key].pw} set={v=>sp(p.key,"pw",v)} type={spw[p.key]?"text":"password"} mono sty={{borderColor:p.col+"55",paddingRight:34}}/>
              <button onClick={()=>setSpw(s=>({...s,[p.key]:!s[p.key]}))} style={{position:"absolute",right:9,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",color:G.mut,fontSize:13}}>{spw[p.key]?"🙈":"👁"}</button>
            </div></F>
          </div>}
        </div>)}
      </div>
      <div style={{padding:"14px 20px",borderTop:`1px solid ${G.bdr}`,display:"flex",gap:10}}>
        <Btn onClick={()=>onSave(f)} sty={{flex:1,padding:12,fontSize:14}}>💾 Save Changes</Btn>
        <button onClick={onX} style={{padding:"12px 18px",borderRadius:10,border:`1px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer"}}>Cancel</button>
      </div>
    </div>
  </div>;
}

// ─── Edit Work Modal ──────────────────────────────────────────────────────────
function EditWork({w,onSave,onX,dd,rcvdDisplay,linkedCount}){
  const[f,setF]=useState({...w});
  return <div style={{position:"fixed",inset:0,background:"#000C",zIndex:5000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
    <div style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:18,width:"min(560px,98%)",maxHeight:"90vh",overflow:"auto",boxShadow:"0 32px 80px #000C"}}>
      <div style={{padding:"16px 20px",borderBottom:`1px solid ${G.bdr}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <span style={{fontWeight:800,fontSize:15,color:G.wh}}>✏️ Edit Work - {f.cn}</span>
        <button onClick={onX} style={{background:"transparent",border:`1px solid ${G.bdr}`,color:G.mut,borderRadius:8,padding:"4px 11px",cursor:"pointer"}}>✕</button>
      </div>
      <div style={{padding:20,display:"flex",flexDirection:"column",gap:11}}>
        <R><F label="Service" w="calc(50% - 6px)"><S val={f.svc} set={v=>setF({...f,svc:v})} opts={dd.services}/></F><F label="FY" w="calc(50% - 6px)"><S val={f.fy} set={v=>setF({...f,fy:v})} opts={dd.fyOptions}/></F></R>
        <R><F label="Staff" w="calc(50% - 6px)"><S val={f.staff} set={v=>setF({...f,staff:v})} opts={dd.staff}/></F><F label="Status" w="calc(50% - 6px)"><S val={f.status} set={v=>setF({...f,status:v})} opts={["Pending","In Progress","Completed"]}/></F></R>
        <R><F label="Due Date" w="calc(50% - 6px)"><I val={f.due} set={v=>setF({...f,due:v})} type="date"/></F><F label="Source" w="calc(50% - 6px)"><S val={f.src||""} set={v=>setF({...f,src:v})} opts={dd.sources} ph="Select..."/></F></R>
        <R>
          <F label="Fees (₹)" w="calc(33% - 8px)"><I val={String(f.fees||"")} set={v=>setF({...f,fees:Number(v.replace(/\D/g,""))||0})}/></F>
          <F label="Commission (₹)" w="calc(33% - 8px)"><I val={String(f.comm||"")} set={v=>setF({...f,comm:Number(v.replace(/\D/g,""))||0})}/></F>
          {linkedCount>0
          ?<F label="Received (₹)" w="calc(33% - 8px)">
            <div style={{...IS,display:"flex",alignItems:"center",color:G.g3,fontWeight:700,cursor:"not-allowed",opacity:.85}}>{inr(rcvdDisplay)}</div>
          </F>
          :<F label="Received (₹) — placeholder" w="calc(33% - 8px)"><I val={String(f.rcvd||"")} set={v=>setF({...f,rcvd:Number(v.replace(/\D/g,""))||0})}/></F>}
        </R>
        <div style={{fontSize:11,color:G.mut,marginTop:-4}}>{linkedCount>0
          ?`💡 ${linkedCount} invoice(s) linked — Received is the live total of actual payment receipts. Manage it from the 🔗 column in the Work Tracker.`
          :`💡 No invoice linked yet — this is a manual placeholder. Link an invoice from the 🔗 column in the Work Tracker to track real payments.`}</div>
        <F label="Remarks"><textarea value={f.note} onChange={e=>setF({...f,note:e.target.value})} rows={2} style={{...IS,resize:"vertical"}}/></F>
      </div>
      <div style={{padding:"14px 20px",borderTop:`1px solid ${G.bdr}`,display:"flex",gap:10}}>
        <Btn onClick={()=>onSave(f)} sty={{flex:1,padding:12,fontSize:14}}>💾 Save</Btn>
        <button onClick={onX} style={{padding:"12px 18px",borderRadius:10,border:`1px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer"}}>Cancel</button>
      </div>
    </div>
  </div>;
}

// ─── Work Dashboard ───────────────────────────────────────────────────────────
function WorkDash({works,clients,dd}){
  const[fy,setFy]=useState(getCurrentFY());
  const fw=fy==="All"?works:works.filter(w=>w.fy===fy);
  const kpis=[
    {l:"Total",v:fw.length,icon:"📋",col:G.green},
    {l:"Pending",v:fw.filter(w=>w.status==="Pending").length,icon:"⏳",col:G.amb},
    {l:"In Progress",v:fw.filter(w=>w.status==="In Progress").length,icon:"🔵",col:G.cyn},
    {l:"Completed",v:fw.filter(w=>w.status==="Completed").length,icon:"✅",col:G.g3},
    {l:"Overdue",v:fw.filter(isOD).length,icon:"🔴",col:G.red},
    {l:"Active Clients",v:clients.filter(c=>c.status==="Active").length,icon:"👥",col:G.vio},
  ];
  const byStaff=dd.staff.map(s=>({n:s,tot:fw.filter(w=>w.staff===s).length,done:fw.filter(w=>w.staff===s&&w.status==="Completed").length})).filter(s=>s.tot>0);
  const bySvc=[...new Set(fw.map(w=>w.svc))].map(s=>({n:s,c:fw.filter(w=>w.svc===s).length})).sort((a,b)=>b.c-a.c);
  const ods=fw.filter(isOD);
  return <div style={{display:"flex",flexDirection:"column",gap:16}}>
    <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
      <span style={{fontSize:12,color:G.mut,fontWeight:600}}>FY:</span>
      {["All",...dd.fyOptions].map(f=><button key={f} onClick={()=>setFy(f)} style={{padding:"5px 13px",borderRadius:20,border:`1.5px solid ${fy===f?G.green:G.bdr}`,cursor:"pointer",fontSize:12,fontWeight:600,background:fy===f?G.green+"18":"transparent",color:fy===f?G.green:G.mut}}>
        {f}{f!=="All"&&<span style={{opacity:.6}}> ({works.filter(w=>w.fy===f).length})</span>}
      </button>)}
    </div>
    <div className="kpi-grid-3" style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
      {kpis.map(k=><div key={k.l} style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:14,padding:"15px 17px",display:"flex",gap:11,alignItems:"center"}}>
        <div style={{width:42,height:42,borderRadius:11,background:k.col+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,flexShrink:0}}>{k.icon}</div>
        <div><div style={{fontWeight:800,fontSize:26,color:k.col,lineHeight:1}}>{k.v}</div><div style={{fontSize:11,color:G.mut,marginTop:3}}>{k.l}</div></div>
      </div>)}
    </div>
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
      <Crd><SH icon="👤" title="Staff Workload" acc={G.vio}/>
        {byStaff.length===0?<div style={{color:G.mut,fontSize:13}}>No data for this FY</div>
        :byStaff.map(s=><div key={s.n} style={{marginBottom:10}}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:3}}><span style={{fontSize:13,color:G.txt,fontWeight:600}}>{s.n}</span><span style={{fontSize:12,color:G.mut}}>{s.done}/{s.tot}</span></div>
          <div style={{height:6,background:G.bdr,borderRadius:10,overflow:"hidden"}}><div style={{height:"100%",width:`${s.tot?s.done/s.tot*100:0}%`,background:`linear-gradient(90deg,${G.g2},${G.g3})`,borderRadius:10}}/></div>
        </div>)}
      </Crd>
      <Crd><SH icon="📊" title="Service Breakdown" acc={G.cyn}/>
        {bySvc.slice(0,7).map((s,i)=>{const cls=[G.green,G.g3,G.cyn,G.amb,G.vio,G.red,G.mut];return <div key={s.n} style={{display:"flex",alignItems:"center",gap:9,marginBottom:7}}>
          <div style={{width:3,height:20,borderRadius:2,background:cls[i%cls.length],flexShrink:0}}/>
          <span style={{fontSize:12,color:G.txt,flex:1}}>{s.n}</span>
          <div style={{width:65,height:5,background:G.bdr,borderRadius:10,overflow:"hidden"}}><div style={{height:"100%",width:`${s.c/(bySvc[0]?.c||1)*100}%`,background:cls[i%cls.length],borderRadius:10}}/></div>
          <span style={{fontSize:12,color:G.mut,width:16,textAlign:"right"}}>{s.c}</span>
        </div>;})}
      </Crd>
    </div>
    {ods.length>0&&<Crd sty={{borderColor:G.red+"44"}}><SH icon="🔴" title={`Overdue (${ods.length})`} acc={G.red}/>
      {ods.map(w=><div key={w.id} style={{display:"flex",gap:11,alignItems:"center",padding:"9px 13px",background:"#450A0A18",border:`1px solid ${G.red}22`,borderRadius:9,marginBottom:7}}>
        <div style={{flex:1}}><div style={{fontWeight:700,fontSize:13,color:G.txt}}>{w.cn}</div><div style={{fontSize:11,color:G.mut}}>{w.svc} · {w.staff}</div></div>
        <div style={{textAlign:"right"}}><div style={{fontSize:11,color:G.red,fontWeight:700}}>{odDays(w)}d late</div><div style={{fontSize:11,color:G.mut}}>Due: {fd(w.due)}</div></div>
        <Bdg label="Overdue"/>
      </div>)}
    </Crd>}
    <Crd sty={{padding:0,overflow:"hidden"}}><div style={{padding:"13px 17px",borderBottom:`1px solid ${G.bdr}`,fontWeight:700,fontSize:14,color:G.wh}}>Recent Assignments {fy!=="All"&&`- ${fy}`}</div>
      <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
        <thead><tr style={{background:G.bg}}>{["Client","Service","FY","Due","Staff","Status"].map(h=><th key={h} style={{padding:"9px 12px",textAlign:"left",color:G.mut,fontWeight:600,fontSize:11,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
        <tbody>{fw.slice(0,8).map((w,i)=><tr key={w.id} style={{borderTop:`1px solid ${G.bdr}`,background:i%2?"#0F1D1408":"transparent"}}>
          <td style={{padding:"9px 12px",fontWeight:600,color:G.txt,whiteSpace:"nowrap"}}>{w.cn}</td>
          <td style={{padding:"9px 12px",color:G.mut}}>{w.svc}</td>
          <td style={{padding:"9px 12px",color:G.mut}}>{w.fy}</td>
          <td style={{padding:"9px 12px",color:isOD(w)?G.red:G.mut,whiteSpace:"nowrap"}}>{fd(w.due)}{isOD(w)&&<span style={{marginLeft:5,fontSize:10,background:"#450A0A",color:G.red,padding:"1px 6px",borderRadius:8}}>{odDays(w)}d</span>}</td>
          <td style={{padding:"9px 12px",color:G.mut}}>{w.staff}</td>
          <td style={{padding:"9px 12px"}}><Bdg label={isOD(w)?"Overdue":w.status}/></td>
        </tr>)}</tbody>
      </table></div>
    </Crd>
  </div>;
}

// ─── Finance Dashboard ────────────────────────────────────────────────────────
function FinDash({works,invoices,receipts,onLogout,dd,clients,setClients,toast}){
  const[fy,setFy]=useState(getCurrentFY()),[show,setShow]=useState(false);
  const[viewPan,setViewPan]=useState(null);
  const[ledgerQ,setLedgerQ]=useState("");
  const[hoveredKpi,setHoveredKpi]=useState(null);

  const fw = useMemo(() => {
    if (fy === "All") return works;
    return works.filter(w => {
      const linked = linkedInvoices(w, invoices);
      if (linked.length > 0) {
        return linked.some(inv => inv.fy === fy);
      }
      return w.fy === fy;
    });
  }, [works, invoices, fy]);

  // Filter invoices and receipts by selected FY
  const activeInvoices = useMemo(() => {
    return fy === "All" ? invoices : invoices.filter(inv => inv.fy === fy);
  }, [invoices, fy]);

  const activeReceipts = useMemo(() => {
    if (fy === "All") return receipts;
    return receipts.filter(r => {
      const inv = invoices.find(i => i.id === r.invId || i.id === r.invoiceId);
      return inv && inv.fy === fy;
    });
  }, [receipts, invoices, fy]);

  // 1. Billed (Invoiced) = TOTAL Amount FROM INV & BILLING
  const tf = activeInvoices.reduce((s, inv) => s + (inv.total || 0), 0);

  // 2. Received (Receipts) = TOTAL Received FROM Payment Receipts
  const tr = activeReceipts.reduce((s, r) => s + (r.received || 0), 0);

  // 3. Outstanding Balance = Billed - Received
  const to = tf - tr;

  // 4. Referral Commission = TOTAL COMMISSION FROM WORK TRACKER
  const tc = fw.reduce((s, w) => s + (w.comm || 0), 0);

  // 5. NET Revenue = Billed - Referral Commission
  const net = tf - tc;

  const mk=v=>show?v:"****";

  const bySrc = dd.sources.map(s => ({
    n: s,
    comm: fw.filter(w => w.src === s).reduce((a, w) => a + (w.comm || 0), 0),
    cnt: fw.filter(w => w.src === s).length
  })).filter(s => s.cnt > 0);

  const bySvc = [...new Set(fw.map(w => w.svc))].map(s => ({
    n: s,
    fees: fw.filter(w => w.svc === s).reduce((a, w) => a + workBilled(w, invoices), 0)
  })).sort((a, b) => b.fees - a.fees);

  const kpis = [
    { l: "Billed (Invoiced)", v: inr(tf), icon: "🧾", col: G.green },
    { l: "Received (Receipts)", v: inr(tr), icon: "💰", col: G.g3 },
    { l: "Outstanding Balance", v: inr(to), icon: "⚠️", col: G.red },
    { l: "Referral Commission", v: inr(tc), icon: "🔁", col: G.amb },
    { l: "Net Revenue", v: inr(net), icon: "📈", col: G.cyn },
    { l: "Works Tracked", v: String(fw.length) + " jobs", icon: "📋", col: G.vio }
  ];

  // Financial Ledger = IF BILL ISSUED THEN SHOW & DATA FETCH FROM WORK TRACKER
  const ledgerWorks = useMemo(() => {
    return fw.filter(w => linkedInvoices(w, invoices).length > 0);
  }, [fw, invoices]);

  // Filter ledger list locally by search term
  const ledgerFiltered = useMemo(() => {
    if (!ledgerQ) return ledgerWorks;
    const lq = ledgerQ.toLowerCase();
    return ledgerWorks.filter(w =>
      (w.cn || "").toLowerCase().includes(lq) ||
      (w.svc || "").toLowerCase().includes(lq) ||
      (w.src || "").toLowerCase().includes(lq) ||
      (w.status || "").toLowerCase().includes(lq)
    );
  }, [ledgerWorks, ledgerQ]);

  // Export ledger to standard CSV file format
  const exportLedgerCSV = () => {
    if (!ledgerWorks.length) {
      toast("No financial ledger rows to export", "err");
      return;
    }
    const headers = ["Client Name", "Service", "FY", "Lead Source", "Billed Fees", "Received", "Outstanding", "Commission", "Net Revenue", "Work Status"];
    const csvRows = ledgerWorks.map(w => {
      const billedVal = workBilled(w, invoices);
      const rcvdVal = workReceived(w, invoices, receipts);
      const os = billedVal - rcvdVal;
      const n = billedVal - (w.comm || 0);
      return [
        w.cn || "-",
        w.svc || "-",
        w.fy || "-",
        w.src || "-",
        billedVal,
        rcvdVal,
        os,
        w.comm || 0,
        n,
        w.status || "-"
      ];
    });

    const csvContent = [headers, ...csvRows]
      .map(row => row.map(cell => `"${String(cell).replace(/"/g, '""')}"`).join(","))
      .join("\n");

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = `FMT_Financial_Ledger_${fy}_${new Date().toISOString().split('T')[0]}.csv`;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    toast("Financial Ledger exported ✓");
  };

  return <div style={{display:"flex",flexDirection:"column",gap:16}}>
    {/* Header Security Actions */}
    <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",flexWrap:"wrap",gap:10}}>
      <div style={{
        display: "flex",
        alignItems: "center",
        gap: 10,
        padding: "8px 16px",
        background: "linear-gradient(135deg, rgba(67, 20, 7, 0.45), rgba(115, 30, 10, 0.35))",
        border: `1.5px solid ${G.amb}40`,
        borderRadius: 12,
        backdropFilter: "blur(10px)",
        boxShadow: "0 8px 32px 0 rgba(0, 0, 0, 0.3)"
      }}>
        <span style={{ animation: "pulse 2s infinite", display: "inline-block" }}>👑</span>
        <span style={{fontSize:13,color:G.amb,fontWeight:800,letterSpacing:0.25}}>Owner View — Private & Confidential</span>
      </div>
      <div style={{display:"flex",gap:7,alignItems:"center",flexWrap:"wrap"}}>
        {["All",...dd.fyOptions].map(f=><button key={f} onClick={()=>setFy(f)} style={{padding:"5px 12px",borderRadius:18,border:`1.5px solid ${fy===f?G.green:G.bdr}`,cursor:"pointer",fontSize:11,fontWeight:700,background:fy===f?G.green+"18":"transparent",color:fy===f?G.green:G.mut}}>{f}</button>)}
        {!show?<button onClick={()=>setShow(true)} style={{padding:"6px 14px",borderRadius:9,border:`1.5px solid ${G.amb}44`,background:"#43140730",color:G.amb,cursor:"pointer",fontWeight:700,fontSize:12}}>🔓 Show Figures</button>
        :<button onClick={()=>setShow(false)} style={{padding:"6px 14px",borderRadius:9,border:`1px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer",fontWeight:700,fontSize:12}}>🔒 Hide</button>}
        <button onClick={onLogout} style={{padding:"6px 14px",borderRadius:9,border:`1.5px solid ${G.red}44`,background:"#450A0A",color:G.red,cursor:"pointer",fontWeight:700,fontSize:12}}>🔐 Lock</button>
      </div>
    </div>

    {/* KPI Summary Cards Grid */}
    <div className="kpi-grid-3" style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
      {kpis.map((k, idx)=><div 
        key={k.l} 
        onMouseEnter={()=>setHoveredKpi(idx)}
        onMouseLeave={()=>setHoveredKpi(null)}
        style={{
          background: G.surf,
          border: `1.5px solid ${hoveredKpi === idx ? k.col + "66" : G.bdr}`,
          borderRadius: 16,
          padding: "16px 20px",
          boxShadow: hoveredKpi === idx ? `0 12px 30px ${k.col}22` : "0 4px 15px rgba(0, 0, 0, 0.15)",
          transform: hoveredKpi === idx ? "translateY(-4px)" : "none",
          transition: "all 0.25s cubic-bezier(0.4, 0, 0.2, 1)",
          cursor: "default"
        }}
      >
        <div style={{display:"flex",gap:12,alignItems:"center"}}>
          <div style={{width:40,height:40,borderRadius:12,background:k.col+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:19,flexShrink:0}}>{k.icon}</div>
          <div>
            <div style={{fontSize:10,color:G.mut,textTransform:"uppercase",letterSpacing:0.6,fontWeight:700}}>{k.l}</div>
            <div style={{fontWeight:800,fontSize:20,color:show?k.col:G.bdr,marginTop:3}}>{mk(k.v)}</div>
          </div>
        </div>
      </div>)}
    </div>

    {/* Subsections: Commissions & Revenue share */}
    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:14}}>
      <Crd><SH icon="🔁" title="Commission by Lead Source" acc={G.amb}/>
        <table style={{width:"100%",borderCollapse:"collapse",fontSize:12,marginTop:10}}>
          <thead><tr>{["Source","Works","Commission"].map(h=><th key={h} style={{padding:"7px 9px",textAlign:"left",color:G.mut,fontWeight:600,fontSize:11,borderBottom:`1px solid ${G.bdr}`}}>{h}</th>)}</tr></thead>
          <tbody>{bySrc.map((s,i)=><tr key={s.n} style={{borderTop:`1px solid ${G.bdr}`,background:i%2?"#0f1d1408":"transparent"}}>
            <td style={{padding:"8px 9px",fontWeight:600,color:G.txt}}>{s.n}</td>
            <td style={{padding:"8px 9px",color:G.mut}}>{s.cnt}</td>
            <td style={{padding:"8px 9px",color:show?G.amb:G.bdr,fontWeight:700}}>{mk(inr(s.comm))}</td>
          </tr>)}</tbody>
        </table>
      </Crd>
      <Crd><SH icon="📊" title="Revenue by Service Classification" acc={G.cyn}/>
        {bySvc.slice(0,7).map((s,i)=>{const cls=[G.green,G.g3,G.cyn,G.amb,G.vio,G.red,G.mut];return <div key={s.n} style={{marginBottom:9}}>
          <div style={{display:"flex",justifyContent:"space-between",marginBottom:4}}><span style={{fontSize:12,color:G.txt}}>{s.n}</span><span style={{fontSize:12,color:show?cls[i%cls.length]:G.bdr,fontWeight:700}}>{mk(inr(s.fees))}</span></div>
          <div style={{height:5,background:G.bdr,borderRadius:10,overflow:"hidden"}}><div style={{height:"100%",width:show?`${s.fees/(bySvc[0]?.fees||1)*100}%`:"0%",background:cls[i%cls.length],borderRadius:10,transition:"width .5s"}}/></div>
        </div>;})}
      </Crd>
    </div>

    {/* Financial Ledger Section */}
    <Crd sty={{padding:0,overflow:"hidden"}}>
      <div style={{padding:"12px 18px",borderBottom:`1px solid ${G.bdr}`,display:"flex",justifyContent:"space-between",alignItems:"center",flexWrap:"wrap",gap:10,background:G.surf}}>
        <div style={{fontWeight:800,fontSize:14,color:G.wh}}>Financial Ledger {fy!=="All"&&`- ${fy}`} ({ledgerFiltered.length} records)</div>
        <div style={{display:"flex",alignItems:"center",gap:8}}>
          <div style={{position:"relative",width:200}}>
            <span style={{position:"absolute",left:8,top:"50%",transform:"translateY(-50%)",fontSize:11,color:G.mut}}>🔍</span>
            <input 
              value={ledgerQ} 
              onChange={e=>setLedgerQ(e.target.value)} 
              placeholder="Search ledger..." 
              style={{...IS,padding:"5px 8px 5px 24px",fontSize:11,borderRadius:6,border:`1px solid ${G.bdr}`}}
            />
          </div>
          <button 
            onClick={exportLedgerCSV}
            style={{
              background: `${G.green}18`,
              border: `1.5px solid ${G.green}40`,
              color: G.green,
              borderRadius: 6,
              padding: "5px 12px",
              cursor: "pointer",
              fontSize: 11,
              fontWeight: 700,
              transition: "all 0.2s"
            }}
          >📥 Export CSV</button>
        </div>
      </div>
      <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
        <thead><tr style={{background:G.bg}}>{["Client","Service","FY","Source","Fees (Billed)","Received","Outstanding","Commission","Net Revenue","Status"].map(h=><th key={h} style={{padding:"9px 11px",textAlign:"left",color:G.mut,fontWeight:600,fontSize:11,whiteSpace:"nowrap"}}>{h}</th>)}</tr></thead>
        <tbody>{ledgerFiltered.map((w,i)=>{
          const billedVal = workBilled(w, invoices);
          const rcvdVal = workReceived(w, invoices, receipts);
          const os = billedVal - rcvdVal;
          const n = billedVal - (w.comm || 0);
          return <tr key={w.id} style={{borderTop:`1px solid ${G.bdr}`,background:i%2?"#0f1d1408":"transparent",transition:"background 0.2s"}}>
            <td style={{padding:"9px 11px",whiteSpace:"nowrap"}}>
              <button onClick={()=>setViewPan(w.pan)} style={{background:"none",border:"none",cursor:"pointer",fontWeight:700,color:G.green,fontSize:12,padding:0,textDecoration:"underline",textUnderlineOffset:3,textDecorationColor:G.green+"55",whiteSpace:"nowrap"} } title="View client details">{w.cn}</button>
            </td>
            <td style={{padding:"9px 11px",color:G.mut}}>{w.svc}</td>
            <td style={{padding:"9px 11px",color:G.mut}}>{w.fy}</td>
            <td style={{padding:"9px 11px",color:G.mut}}>{w.src||"-"}</td>
            <td style={{padding:"9px 11px",color:show?G.green:G.bdr,fontWeight:600}}>{mk(inr(billedVal))}</td>
            <td style={{padding:"9px 11px",color:show?G.g3:G.bdr,fontWeight:600}}>{mk(inr(rcvdVal))}</td>
            <td style={{
              padding: "9px 11px",
              color: show ? (os > 0 ? G.red : G.g3) : G.bdr,
              fontWeight: 600,
              background: show && os > 0 ? "rgba(239, 68, 68, 0.05)" : "transparent"
            }}>{mk(inr(os))}</td>
            <td style={{padding:"9px 11px",color:show?G.amb:G.bdr}}>{mk(inr(w.comm))}</td>
            <td style={{padding:"9px 11px",color:show?G.cyn:G.bdr,fontWeight:800}}>{mk(inr(n))}</td>
            <td style={{padding:"9px 11px"}}><Bdg label={w.status}/></td>
          </tr>;
        })}
        </tbody>
        <tfoot><tr style={{background:G.gd,borderTop:`2px solid ${G.green}44`}}>
          <td colSpan={4} style={{padding:"10px 11px",fontWeight:800,color:G.wh,fontSize:13}}>TOTALS</td>
          {[[tf,G.green],[tr,G.g3],[to,G.red],[tc,G.amb],[net,G.cyn]].map(([v,col],i)=><td key={i} style={{padding:"10px 11px",color:show?col:G.bdr,fontWeight:800}}>{mk(inr(v))}</td>)}
          <td/>
        </tr></tfoot>
      </table></div>
    </Crd>
    {viewPan&&((()=>{const c=clients&&clients.find(x=>x.pan===viewPan);return c?<EditClient c={c} dd={dd} onX={()=>setViewPan(null)} onSave={cf=>{setClients&&setClients(p=>p.map(x=>x.pan===cf.pan?cf:x));setViewPan(null);toast&&toast("Client updated!");}}/>:null;})())}
  </div>;
}


// ─── Add Client ───────────────────────────────────────────────────────────────
function AddClient({clients,setClients,dd,toast}){
  const bk={pan:"",name:"",biz:"",mob:"",email:"",gstin:"",addr:"",state:"",type:"",src:"",added:td(),status:"Active",note:"",aadhaar:"",dob:"",fatherName:"",gender:"",residentialStatus:"",pin:"",bankName:"",ifsc:"",accountNumber:"",itrType:"",extraPw:[],portals:mkP()};
  const[f,setF]=useState(bk),[err,setErr]=useState({}),[spw,setSpw]=useState({});
  const sp=(k,fld,v)=>setF(p=>({...p,portals:{...p.portals,[k]:{...p.portals[k],[fld]:v}}}));
  const tp=(k)=>setF(p=>({...p,portals:{...p.portals,[k]:{...p.portals[k],on:!p.portals[k].on}}}));
  const addPw=()=>setF(p=>({...p,extraPw:[...(p.extraPw||[]),{type:dd.pwTypes?.[0]||"Other",label:"",id:"",pw:""}]}));
  const setPwAt=(i,patch)=>setF(p=>({...p,extraPw:p.extraPw.map((x,j)=>j===i?{...x,...patch}:x)}));
  const delPwAt=(i)=>setF(p=>({...p,extraPw:p.extraPw.filter((_,j)=>j!==i)}));
  const save=()=>{
    const e={};
    if(!f.pan)e.pan="Required";else if(f.pan.length!==10)e.pan="10 chars";else if(clients.find(c=>c.pan===f.pan.toUpperCase()))e.pan="Already exists";
    if(!f.name)e.name="Required";if(!f.mob||f.mob.length<10)e.mob="Valid mobile";if(!f.type)e.type="Select";if(!f.src)e.src="Select";
    setErr(e);if(Object.keys(e).length){toast("Fix errors","err");return;}
    setClients(p=>[{...f,pan:f.pan.toUpperCase()},...p]);setF(bk);setErr({});toast(`${f.name} added!`);
  };
  return <div style={{display:"flex",gap:16}}>
    <div style={{flex:1,display:"flex",flexDirection:"column",gap:13}}>
      <Crd><SH icon="👤" title="Client Information"/>
        <div style={{display:"flex",flexDirection:"column",gap:11}}>
          <R><F label="PAN (Client ID)" req w="calc(50% - 6px)"><I val={f.pan} set={v=>setF({...f,pan:v.toUpperCase().slice(0,10)})} ph="ABCDE1234F" mono sty={{borderColor:err.pan?G.red:G.bdr}}/>{err.pan&&<span style={{fontSize:11,color:G.red}}>{err.pan}</span>}</F>
          <F label="Client Type" req w="calc(50% - 6px)"><S val={f.type} set={v=>setF({...f,type:v})} opts={dd.clientTypes} ph="Select..."/>{err.type&&<span style={{fontSize:11,color:G.red}}>{err.type}</span>}</F></R>
          <R><F label="Full Name" req w="calc(50% - 6px)"><I val={f.name} set={v=>setF({...f,name:v})} sty={{borderColor:err.name?G.red:G.bdr}}/>{err.name&&<span style={{fontSize:11,color:G.red}}>{err.name}</span>}</F>
          <F label="Business Name" w="calc(50% - 6px)"><I val={f.biz} set={v=>setF({...f,biz:v})}/></F></R>
          <R><F label="Mobile" req w="calc(50% - 6px)"><I val={f.mob} set={v=>setF({...f,mob:v.replace(/\D/g,"").slice(0,10)})} sty={{borderColor:err.mob?G.red:G.bdr}}/>{err.mob&&<span style={{fontSize:11,color:G.red}}>{err.mob}</span>}</F>
          <F label="Email" w="calc(50% - 6px)"><I val={f.email} set={v=>setF({...f,email:v})} type="email"/></F></R>
          <R><F label="GSTIN" w="calc(50% - 6px)"><I val={f.gstin} set={v=>setF({...f,gstin:v.toUpperCase()})} mono/></F>
          <F label="Source" req w="calc(50% - 6px)"><S val={f.src} set={v=>setF({...f,src:v})} opts={dd.sources} ph="Select..."/>{err.src&&<span style={{fontSize:11,color:G.red}}>{err.src}</span>}</F></R>
          <F label="Address"><I val={f.addr} set={v=>setF({...f,addr:v})}/></F>
          <R><F label="State" w="calc(50% - 6px)"><S val={f.state} set={v=>setF({...f,state:v})} opts={dd.states} ph="Select..."/></F>
          <F label="Status" w="calc(50% - 6px)"><S val={f.status} set={v=>setF({...f,status:v})} opts={["Active","Inactive"]}/></F></R>
          <F label="Remarks"><textarea value={f.note} onChange={e=>setF({...f,note:e.target.value})} rows={2} style={{...IS,resize:"vertical"}}/></F>
        </div>
      </Crd>
      <Crd><SH icon="🧾" title="ITR Profile Details - optional, used to auto-fill Tax Computations" acc={G.cyn}/>
        <div style={{display:"flex",flexDirection:"column",gap:11}}>
          <R><F label="Father's Name" w="calc(50% - 6px)"><I val={f.fatherName} set={v=>setF({...f,fatherName:v})}/></F>
          <F label="Date of Birth" w="calc(50% - 6px)"><I val={f.dob} set={v=>setF({...f,dob:v})} type="date"/></F></R>
          <R><F label="Gender" w="calc(50% - 6px)"><S val={f.gender} set={v=>setF({...f,gender:v})} opts={["Male","Female","Other"]} ph="Select..."/></F>
          <F label="Residential Status" w="calc(50% - 6px)"><S val={f.residentialStatus} set={v=>setF({...f,residentialStatus:v})} opts={["Resident","Resident but Not Ordinarily Resident","Non-Resident"]} ph="Select..."/></F></R>
          <R><F label="Aadhaar No." w="calc(50% - 6px)"><I val={f.aadhaar} set={v=>setF({...f,aadhaar:v.replace(/\D/g,"").slice(0,12)})} mono ph="12-digit number"/></F>
          <F label="PIN Code" w="calc(50% - 6px)"><I val={f.pin} set={v=>setF({...f,pin:v.replace(/\D/g,"").slice(0,6)})} mono/></F></R>
          <R><F label="Bank Name" w="calc(50% - 6px)"><I val={f.bankName} set={v=>setF({...f,bankName:v})}/></F>
          <F label="Bank IFSC" w="calc(50% - 6px)"><I val={f.ifsc} set={v=>setF({...f,ifsc:v.toUpperCase()})} mono ph="SBIN0001234"/></F></R>
          <R><F label="Bank Account No." w="calc(50% - 6px)"><I val={f.accountNumber} set={v=>setF({...f,accountNumber:v})} mono/></F>
          <F label="ITR Type" w="calc(50% - 6px)"><S val={f.itrType} set={v=>setF({...f,itrType:v})} opts={["ITR-1","ITR-2","ITR-3","ITR-4"]} ph="Select..."/></F></R>
        </div>
      </Crd>
      <Crd><SH icon="🔑" title="Other Passwords (optional)" acc={G.cyn}/>
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {(f.extraPw||[]).map((pwItem,i)=><div key={i} style={{display:"flex",gap:8,alignItems:"flex-end",padding:"9px 10px",background:G.bg,border:`1px solid ${G.bdr}`,borderRadius:9}}>
            <F label="Type" w="22%"><S val={pwItem.type} set={v=>setPwAt(i,{type:v})} opts={dd.pwTypes||[]} ph="Select..."/></F>
            <F label="Label" w="26%"><I val={pwItem.label} set={v=>setPwAt(i,{label:v})} ph="e.g. Personal Email"/></F>
            <F label="User ID" w="24%"><I val={pwItem.id} set={v=>setPwAt(i,{id:v})} mono/></F>
            <F label="Password" w="20%"><I val={pwItem.pw} set={v=>setPwAt(i,{pw:v})} type="password" mono/></F>
            <button onClick={()=>delPwAt(i)} style={{background:"#450A0A",border:`1px solid ${G.red}44`,color:G.red,borderRadius:8,padding:"9px 10px",cursor:"pointer",fontSize:12,flexShrink:0}}>🗑</button>
          </div>)}
          <button onClick={addPw} style={{alignSelf:"flex-start",padding:"7px 14px",borderRadius:8,border:`1px dashed ${G.bdr}`,background:"transparent",color:G.cyn,cursor:"pointer",fontSize:12,fontWeight:600}}>➕ Add Password</button>
        </div>
      </Crd>
      <Crd><SH icon="🔐" title="Portal Credentials - tick only portals this client uses" acc={G.vio}/>
        <div style={{display:"flex",flexDirection:"column",gap:11}}>
          {PORTALS.map(p=><div key={p.key}>
            <label style={{display:"flex",alignItems:"center",gap:9,cursor:"pointer",marginBottom:f.portals[p.key].on?8:0}}>
              <div onClick={()=>tp(p.key)} style={{width:19,height:19,borderRadius:5,border:`2px solid ${f.portals[p.key].on?p.col:G.bdr}`,background:f.portals[p.key].on?p.col+"20":"transparent",display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",flexShrink:0}}>
                {f.portals[p.key].on&&<span style={{color:p.col,fontSize:11,fontWeight:900,lineHeight:1}}>✓</span>}
              </div>
              <span style={{fontSize:13,color:f.portals[p.key].on?p.col:G.mut,fontWeight:f.portals[p.key].on?700:400}}>{p.icon} {p.label}</span>
            </label>
            {f.portals[p.key].on&&<div style={{display:"flex",gap:10,paddingLeft:27}}>
              <F label="User ID" w="calc(50% - 5px)"><I val={f.portals[p.key].id} set={v=>sp(p.key,"id",v)} mono sty={{borderColor:p.col+"55"}}/></F>
              <F label="Password" w="calc(50% - 5px)"><div style={{position:"relative"}}>
                <I val={f.portals[p.key].pw} set={v=>sp(p.key,"pw",v)} type={spw[p.key]?"text":"password"} mono sty={{borderColor:p.col+"55",paddingRight:34}}/>
                <button onClick={()=>setSpw(s=>({...s,[p.key]:!s[p.key]}))} style={{position:"absolute",right:9,top:"50%",transform:"translateY(-50%)",background:"none",border:"none",cursor:"pointer",color:G.mut,fontSize:13}}>{spw[p.key]?"🙈":"👁"}</button>
              </div></F>
            </div>}
          </div>)}
        </div>
      </Crd>
      <div style={{display:"flex",gap:10}}>
        <Btn onClick={save} sty={{flex:1,padding:13,fontSize:14}}>➕ Save Client</Btn>
        <button onClick={()=>{setF(bk);setErr({});}} style={{padding:"13px 17px",borderRadius:11,border:`1.5px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer",fontWeight:600}}>Clear</button>
      </div>
    </div>
    <div style={{width:230,flexShrink:0,display:"flex",flexDirection:"column",gap:12}}>
      <Crd><SH icon="ℹ️" title="Tips" acc={G.cyn}/>
        <div style={{marginTop:10,fontSize:12,color:G.mut,lineHeight:1.9}}>
          <div>* PAN = 10 chars, auto-uppercase</div><div>* Tick only portals client uses</div>
          <div>* Fees added when assigning work</div><div>* Credentials visible in Client list</div>
        </div>
      </Crd>
      <Crd><SH icon="📊" title="By Type" acc={G.green}/>
        <div style={{marginTop:10,display:"flex",flexDirection:"column",gap:6}}>
          {dd.clientTypes.map(t=>{const n=clients.filter(c=>c.type===t).length;return n>0&&<div key={t} style={{display:"flex",justifyContent:"space-between",fontSize:12}}><span style={{color:G.mut}}>{t}</span><span style={{color:G.green,fontWeight:700}}>{n}</span></div>;})}
        </div>
      </Crd>
    </div>
  </div>;
}

// ─── Assign Work ──────────────────────────────────────────────────────────────
function AssignWork({clients,works,setWorks,dd,toast}){
  const bk={pan:"",svc:"",fy:dd.fyOptions[0]||"2025-26",due:"",staff:"",fees:"",comm:"",rcvd:"",src:"",note:""};
  const[f,setF]=useState(bk),[err,setErr]=useState({}),[showP,setShowP]=useState(false);
  const sel=clients.find(c=>c.pan===f.pan);
  const hFees=v=>{const n=Number(v.replace(/\D/g,""));setF(p=>({...p,fees:v.replace(/\D/g,""),comm:p.comm||String(Math.round(n*0.1))}));};
  const save=()=>{
    const e={};if(!f.pan)e.pan="Select";if(!f.svc)e.svc="Select";if(!f.due)e.due="Required";if(!f.staff)e.staff="Required";if(!f.fees)e.fees="Enter";
    setErr(e);if(Object.keys(e).length){toast("Fix errors","err");return;}
    setWorks(p=>[{id:Date.now(),pan:f.pan,cn:sel?.name||"",svc:f.svc,fy:f.fy,due:f.due,staff:f.staff,status:"Pending",fees:Number(f.fees)||0,comm:Number(f.comm)||0,rcvd:Number(f.rcvd)||0,src:f.src||sel?.src||"",note:f.note},...p]);
    setF(bk);setErr({});setShowP(false);toast(`Assigned - ${f.svc} for ${sel?.name}`);
  };
  return <div style={{display:"grid",gridTemplateColumns:"1fr 270px",gap:16}}>
    <div style={{display:"flex",flexDirection:"column",gap:13}}>
      <Crd><SH icon="🔍" title="Select Client" acc={G.green}/>
        <PanPick clients={clients} val={f.pan} set={v=>{setF(p=>({...p,pan:v,src:clients.find(c=>c.pan===v)?.src||""}));setShowP(false);}}/>
        {err.pan&&<span style={{fontSize:11,color:G.red,marginTop:3,display:"block"}}>{err.pan}</span>}
        {sel&&<div style={{marginTop:14,display:"flex",flexDirection:"column",gap:10}}>
          {/* ── Full Client Details Card ── */}
          <div style={{background:G.bg,border:`1px solid ${G.green}30`,borderRadius:13,overflow:"hidden"}}>
            {/* Header */}
            <div style={{background:`linear-gradient(135deg,${G.gd},${G.g2})`,padding:"12px 16px",display:"flex",gap:12,alignItems:"center"}}>
              <div style={{width:44,height:44,borderRadius:11,background:`linear-gradient(135deg,${G.g2},${G.green})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,fontWeight:800,color:"#fff",flexShrink:0,border:"2px solid #4ADE8055"}}>{sel.name[0]}</div>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontWeight:800,fontSize:15,color:G.wh,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{sel.name}</div>
                {sel.biz&&<div style={{fontSize:12,color:G.g3,marginTop:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{sel.biz}</div>}
                <div style={{display:"flex",gap:8,marginTop:4,flexWrap:"wrap"}}>
                  <span style={{fontSize:11,color:G.green,fontWeight:700,fontFamily:"monospace",background:G.green+"15",padding:"1px 8px",borderRadius:8}}>🆔 {sel.pan}</span>
                  <span style={{fontSize:11,background:sel.status==="Active"?"#14532D":"#450A0A",color:sel.status==="Active"?"#4ADE80":"#FCA5A5",padding:"1px 8px",borderRadius:8,fontWeight:700}}>{sel.status}</span>
                  <span style={{fontSize:11,color:G.mut,background:G.bdr,padding:"1px 8px",borderRadius:8}}>{sel.type}</span>
                </div>
              </div>
            </div>
            {/* Details Grid */}
            <div style={{padding:"12px 16px",display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
              {[
                {l:"📱 Mobile",v:sel.mob},
                {l:"📍 State",v:sel.state},
                {l:"✉ Email",v:sel.email,full:true},
                {l:"📣 Source",v:sel.src},
                {l:"📦 GSTIN",v:sel.gstin||"Not registered",mono:true},
                {l:"📅 Client Since",v:sel.added?new Date(sel.added).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"}):"-"},
              ].map(d=><div key={d.l} style={{gridColumn:d.full?"1 / -1":"auto"}}>
                <div style={{fontSize:10,color:G.mut,fontWeight:700,marginBottom:2}}>{d.l}</div>
                <div style={{fontSize:12,color:d.mono?G.cyn:G.txt,fontFamily:d.mono?"monospace":"inherit",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{d.v||"-"}</div>
              </div>)}
            </div>
            {/* Remarks */}
            {sel.note&&<div style={{padding:"0 16px 12px"}}>
              <div style={{fontSize:10,color:G.mut,fontWeight:700,marginBottom:3}}>📝 REMARKS</div>
              <div style={{fontSize:12,color:G.txt,background:G.surf,padding:"7px 10px",borderRadius:8,border:`1px solid ${G.bdr}`}}>{sel.note}</div>
            </div>}
            {/* Past Works for this client */}
            {works.filter(w=>w.pan===sel.pan).length>0&&<div style={{padding:"0 16px 12px"}}>
              <div style={{fontSize:10,color:G.mut,fontWeight:700,marginBottom:6}}>📋 PAST ASSIGNMENTS ({works.filter(w=>w.pan===sel.pan).length})</div>
              <div style={{display:"flex",flexDirection:"column",gap:4}}>
                {works.filter(w=>w.pan===sel.pan).slice(0,4).map(w=><div key={w.id} style={{display:"flex",alignItems:"center",gap:8,padding:"5px 9px",background:G.surf,borderRadius:7,border:`1px solid ${G.bdr}`}}>
                  <span style={{fontSize:11,color:G.txt,flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{w.svc}</span>
                  <span style={{fontSize:10,color:G.mut}}>{w.fy}</span>
                  <span style={{fontSize:10,padding:"1px 7px",borderRadius:10,fontWeight:700,background:w.status==="Completed"?"#14532D":w.status==="In Progress"?"#1E1B4B":"#431407",color:w.status==="Completed"?"#4ADE80":w.status==="In Progress"?"#A5B4FC":"#FCD34D"}}>{w.status}</span>
                </div>)}
              </div>
            </div>}
          </div>
          {/* Portal Credentials */}
          <div style={{background:G.bg,border:`1px solid ${G.bdr}`,borderRadius:11,padding:"10px 14px"}}>
            <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
              <div style={{fontSize:11,color:G.mut,fontWeight:700}}>🔐 PORTAL CREDENTIALS ({PORTALS.filter(p=>sel.portals[p.key]?.on).length} active)</div>
              <button onClick={()=>setShowP(v=>!v)} style={{fontSize:11,padding:"3px 10px",borderRadius:7,border:`1px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer"}}>{showP?"🙈 Hide":"👁 Show"}</button>
            </div>
            <div style={{display:"flex",gap:5,flexWrap:"wrap",marginBottom:showP?8:0}}>
              {PORTALS.filter(p=>sel.portals[p.key]?.on).map(p=><span key={p.key} style={{fontSize:11,padding:"2px 8px",borderRadius:10,background:p.col+"18",color:p.col,fontWeight:600}}>{p.icon} {p.label}</span>)}
              {!PORTALS.some(p=>sel.portals[p.key]?.on)&&<span style={{fontSize:11,color:G.bdr}}>No portals configured</span>}
            </div>
            {showP&&<div style={{display:"flex",flexDirection:"column",gap:6,marginTop:4}}>
              {PORTALS.filter(p=>sel.portals[p.key]?.on).map(p=><div key={p.key} style={{display:"flex",gap:9,padding:"8px 11px",background:p.col+"0A",border:`1px solid ${p.col}22`,borderRadius:9}}>
                <span style={{fontSize:15,flexShrink:0}}>{p.icon}</span>
                <div style={{flex:1,minWidth:0}}>
                  <div style={{fontSize:10,color:p.col,fontWeight:700,marginBottom:2}}>{p.label}</div>
                  <div style={{fontSize:11,color:G.mut}}>ID: <span style={{color:G.wh,fontFamily:"monospace",fontWeight:600}}>{sel.portals[p.key].id||"-"}</span></div>
                  <div style={{fontSize:11,color:G.mut}}>PW: <span style={{color:G.wh,fontFamily:"monospace",fontWeight:600}}>{sel.portals[p.key].pw||"-"}</span></div>
                </div>
              </div>)}
              {!PORTALS.some(p=>sel.portals[p.key]?.on)&&<div style={{fontSize:12,color:G.mut}}>No portals configured for this client</div>}
            </div>}
          </div>
        </div>}
      </Crd>
      <Crd><SH icon="📋" title="Work Details" acc={G.ind}/>
        <div style={{display:"flex",flexDirection:"column",gap:11}}>
          <F label="Service Type" req><S val={f.svc} set={v=>setF({...f,svc:v})} opts={dd.services} ph="Choose..."/>{err.svc&&<span style={{fontSize:11,color:G.red}}>{err.svc}</span>}</F>
          <R><F label="Financial Year" w="calc(50% - 6px)"><S val={f.fy} set={v=>setF({...f,fy:v})} opts={dd.fyOptions}/></F>
          <F label="Assign To" req w="calc(50% - 6px)"><S val={f.staff} set={v=>setF({...f,staff:v})} opts={dd.staff} ph="Staff..."/>{err.staff&&<span style={{fontSize:11,color:G.red}}>{err.staff}</span>}</F></R>
          <R><F label="Due Date" req w="calc(50% - 6px)"><I val={f.due} set={v=>setF({...f,due:v})} type="date" sty={{borderColor:err.due?G.red:G.bdr}}/>{err.due&&<span style={{fontSize:11,color:G.red}}>{err.due}</span>}</F>
          <F label="Source" w="calc(50% - 6px)"><S val={f.src} set={v=>setF({...f,src:v})} opts={dd.sources} ph="Select..."/></F></R>
        </div>
      </Crd>
      <Crd><SH icon="💰" title="Fees & Commission" acc={G.amb}/>
        <div style={{display:"flex",flexDirection:"column",gap:11}}>
          <div style={{padding:"8px 12px",background:"#43140720",border:`1px solid ${G.amb}33`,borderRadius:9,fontSize:12,color:G.amb}}>💡 Commission auto 10% - editable. Tracked by Source.</div>
          <R>
            <F label="Fees (₹)" req w="calc(33% - 8px)"><I val={f.fees} set={hFees} ph="0" sty={{borderColor:err.fees?G.red:G.bdr}}/>{err.fees&&<span style={{fontSize:11,color:G.red}}>{err.fees}</span>}</F>
            <F label="Commission (₹)" w="calc(33% - 8px)"><I val={f.comm} set={v=>setF({...f,comm:v.replace(/\D/g,"")})} ph="Auto 10%"/></F>
            <F label="Received (₹)" w="calc(33% - 8px)"><I val={f.rcvd} set={v=>setF({...f,rcvd:v.replace(/\D/g,"")})} ph="0"/></F>
          </R>
          <div style={{fontSize:11,color:G.mut,marginTop:-4}}>💡 "Received" here is just a placeholder until you raise & link an invoice — once linked from the Work Tracker, it's replaced by real payments recorded in Receipts.</div>
          {f.fees&&<div style={{display:"flex",gap:12,padding:"8px 12px",background:G.green+"08",border:`1px solid ${G.green}20`,borderRadius:9,flexWrap:"wrap"}}>
            {[["Fees",inr(f.fees),G.green],["Comm.",inr(f.comm||0),G.amb],["O/S",inr((Number(f.fees)||0)-(Number(f.rcvd)||0)),G.red],["Net",inr((Number(f.rcvd)||0)-(Number(f.comm)||0)),G.cyn]].map(([l,v,col])=><div key={l} style={{fontSize:12}}><span style={{color:G.mut}}>{l}: </span><span style={{color:col,fontWeight:700}}>{v}</span></div>)}
          </div>}
          <F label="Remarks"><textarea value={f.note} onChange={e=>setF({...f,note:e.target.value})} rows={2} style={{...IS,resize:"vertical"}}/></F>
        </div>
      </Crd>
      <Btn onClick={save} sty={{width:"100%",padding:13,fontSize:14}}>📋 Add to Work Tracker</Btn>
    </div>
    <div style={{display:"flex",flexDirection:"column",gap:12}}>
      <Crd><SH icon="⚡" title="Quick Service" acc={G.green}/>
        <div style={{display:"flex",flexDirection:"column",gap:3}}>
          {dd.services.map(s=><div key={s} onClick={()=>setF(p=>({...p,svc:s}))}
            style={{padding:"7px 10px",borderRadius:7,cursor:"pointer",fontSize:12,background:f.svc===s?G.green+"18":"transparent",border:f.svc===s?`1px solid ${G.green}44`:"1px solid transparent",color:f.svc===s?G.green:G.mut,fontWeight:f.svc===s?700:400,display:"flex",alignItems:"center",gap:6}}>
            <span style={{fontSize:7,color:f.svc===s?G.green:G.bdr}}>●</span>{s}
          </div>)}
        </div>
      </Crd>
      <Crd><SH icon="🕐" title="Recent" acc={G.vio}/>
        <div style={{display:"flex",flexDirection:"column",gap:6}}>
          {works.slice(0,4).map(w=><div key={w.id} style={{padding:"8px 10px",background:G.bg,borderRadius:8,border:`1px solid ${G.bdr}`}}>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}><span style={{fontWeight:600,fontSize:12,color:G.txt}}>{w.cn}</span><Bdg label={isOD(w)?"Overdue":w.status}/></div>
            <div style={{fontSize:11,color:G.mut,marginTop:2}}>{w.svc} · {w.fy}</div>
            <div style={{fontSize:11,color:isOD(w)?G.red:G.mut,marginTop:1}}>Due: {fd(w.due)}</div>
          </div>)}
        </div>
      </Crd>
    </div>
  </div>;
}

// ─── Client List ──────────────────────────────────────────────────────────────
function ClientList({clients,setClients,dd,toast}){
  const[q,setQ]=useState(""),[view,setView]=useState("grid"),[opP,setOpP]=useState({}),[vpw,setVpw]=useState({}),[editC,setEditC]=useState(null);
  const[delC,setDelC]=useState(null);
  const[showImport,setShowImport]=useState(false),[importBusy,setImportBusy]=useState(false),[importResult,setImportResult]=useState(null);
  const list=useMemo(()=>{if(!q)return clients;const lq=q.toLowerCase();return clients.filter(c=>c.pan.toLowerCase().includes(lq)||c.name.toLowerCase().includes(lq)||c.mob.includes(q));},[q,clients]);
  const doDeleteClient=()=>{setClients(p=>p.filter(x=>x.pan!==delC.pan));toast(`${delC.name} deleted`,"err");setDelC(null);};

  // ─── Bulk CSV Import ────────────────────────────────────────────────────
  const TEMPLATE_HEADERS=["PAN*","Full Name*","Business Name","Mobile*","Email","GSTIN","Address","State","Client Type*","Source*","Status","Remarks","Aadhaar","DOB (YYYY-MM-DD)","Father's Name","Gender","Residential Status","PIN Code","Bank Name","IFSC","Bank Account No.","ITR Type"];
  const TEMPLATE_EXAMPLE=["ABCDE1234F","Rajesh Kumar","Kumar Traders","9876543210","rajesh@example.com","27ABCDE1234F1Z5","123 MG Road, Mumbai","Maharashtra",dd.clientTypes?.[0]||"Individual",dd.sources?.[0]||"Referral","Active","Sample row - replace with real data / delete before importing","123456789012","1990-05-15","Suresh Kumar","Male","Resident","400001","State Bank of India","SBIN0001234","123456789012","ITR-3"];
  const csvEscape=v=>{const s=String(v??"");return /[",\n]/.test(s)?'"'+s.replace(/"/g,'""')+'"':s;};
  const downloadTemplate=()=>{
    const csv=[TEMPLATE_HEADERS,TEMPLATE_EXAMPLE].map(row=>row.map(csvEscape).join(",")).join("\r\n");
    const blob=new Blob([csv],{type:"text/csv;charset=utf-8;"});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url;a.download="client_import_template.csv";
    document.body.appendChild(a);a.click();document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };
  const norm=h=>String(h||"").toLowerCase().replace(/[^a-z0-9]/g,"");
  const HEADER_MAP={
    pan:"pan",
    fullname:"name",name:"name",clientname:"name",
    businessname:"biz",business:"biz",firmname:"biz",
    mobile:"mob",mobileno:"mob",mobilenumber:"mob",phone:"mob",phonenumber:"mob",
    email:"email",emailid:"email",
    gstin:"gstin",gstno:"gstin",
    address:"addr",
    state:"state",
    clienttype:"type",type:"type",
    source:"src",
    status:"status",
    remarks:"note",notes:"note",note:"note",
    aadhaar:"aadhaar",aadhaarno:"aadhaar",aadhaarnumber:"aadhaar",
    dob:"dob",dateofbirth:"dob",dobyyyymmdd:"dob",
    fathersname:"fatherName",fathername:"fatherName",
    gender:"gender",
    residentialstatus:"residentialStatus",
    pincode:"pin",pin:"pin",
    bankname:"bankName",
    ifsc:"ifsc",ifsccode:"ifsc",
    bankaccountno:"accountNumber",accountnumber:"accountNumber",bankaccount:"accountNumber",bankaccountnumber:"accountNumber",
    itrtype:"itrType",
  };
  const handleImportFile=e=>{
    const file=e.target.files[0];
    e.target.value="";
    if(!file)return;
    setImportBusy(true);
    const reader=new FileReader();
    reader.onload=ev=>{
      try{
        const wb=XLSX.read(ev.target.result,{type:"string"});
        const sheet=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(sheet,{defval:"",raw:false});
        const seenPans=new Set(clients.map(c=>c.pan));
        const toAdd=[],skipped=[];
        rows.forEach((row,idx)=>{
          const rec={};
          Object.entries(row).forEach(([k,v])=>{const key=HEADER_MAP[norm(k)];if(key)rec[key]=typeof v==="string"?v.trim():v;});
          const rowNo=idx+2;
          const pan=(rec.pan||"").toString().toUpperCase().replace(/[^A-Z0-9]/g,"").slice(0,10);
          const name=(rec.name||"").toString().trim();
          const mob=(rec.mob||"").toString().replace(/\D/g,"").slice(0,10);
          const type=(rec.type||"").toString().trim();
          const src=(rec.src||"").toString().trim();
          if(!pan&&!name){return;} // fully blank row, ignore silently
          if(!pan||pan.length!==10){skipped.push({row:rowNo,pan:rec.pan||"-",reason:"PAN missing or not 10 characters"});return;}
          if(!name){skipped.push({row:rowNo,pan,reason:"Full Name missing"});return;}
          if(!mob||mob.length<10){skipped.push({row:rowNo,pan,reason:"Valid 10-digit Mobile missing"});return;}
          if(!type){skipped.push({row:rowNo,pan,reason:"Client Type missing"});return;}
          if(!src){skipped.push({row:rowNo,pan,reason:"Source missing"});return;}
          if(seenPans.has(pan)){skipped.push({row:rowNo,pan,reason:"PAN already exists (duplicate)"});return;}
          seenPans.add(pan);
          toAdd.push({
            pan,name,biz:rec.biz||"",mob,email:rec.email||"",gstin:(rec.gstin||"").toString().toUpperCase(),
            addr:rec.addr||"",state:rec.state||"",type,src,
            added:td(),status:(rec.status||"Active").toString().trim()||"Active",note:rec.note||"",
            aadhaar:(rec.aadhaar||"").toString().replace(/\D/g,"").slice(0,12),
            dob:rec.dob||"",fatherName:rec.fatherName||"",gender:rec.gender||"",
            residentialStatus:rec.residentialStatus||"",pin:(rec.pin||"").toString().replace(/\D/g,"").slice(0,6),
            bankName:rec.bankName||"",ifsc:(rec.ifsc||"").toString().toUpperCase(),accountNumber:rec.accountNumber||"",
            itrType:rec.itrType||"",extraPw:[],portals:mkP(),
          });
        });
        if(toAdd.length)setClients(p=>[...toAdd,...p]);
        setImportResult({total:rows.length,imported:toAdd.length,skipped});
        toast(`${toAdd.length} client(s) imported${skipped.length?`, ${skipped.length} skipped`:""}`,toAdd.length?"ok":"err");
      }catch(err){
        toast("Could not read this file - make sure it's a valid CSV","err");
      }
      setImportBusy(false);
    };
    reader.onerror=()=>{toast("Could not read file","err");setImportBusy(false);};
    reader.readAsText(file);
  };
  const PortBlock=({c})=><>{PORTALS.filter(p=>c.portals[p.key]?.on).map(p=><div key={p.key} style={{display:"flex",gap:8,padding:"6px 10px",background:p.col+"0A",border:`1px solid ${p.col}22`,borderRadius:8,marginBottom:4}}>
    <span style={{fontSize:14,flexShrink:0}}>{p.icon}</span>
    <div style={{flex:1,minWidth:0}}>
      <div style={{fontSize:10,color:p.col,fontWeight:700}}>{p.label}</div>
      <div style={{fontSize:11,color:G.mut}}>ID: <span style={{color:G.wh,fontFamily:"monospace"}}>{c.portals[p.key].id||"-"}</span></div>
      <div style={{fontSize:11,color:G.mut,display:"flex",alignItems:"center",gap:4}}>PW: <span style={{color:G.wh,fontFamily:"monospace"}}>{vpw[`${c.pan}_${p.key}`]?c.portals[p.key].pw||"-":"********"}</span>
        <button onClick={()=>setVpw(s=>({...s,[`${c.pan}_${p.key}`]:!s[`${c.pan}_${p.key}`]}))} style={{background:"none",border:"none",cursor:"pointer",color:G.mut,fontSize:12,padding:0}}>{vpw[`${c.pan}_${p.key}`]?"🙈":"👁"}</button>
      </div>
    </div>
  </div>)}
  {!PORTALS.some(p=>c.portals[p.key]?.on)&&<div style={{fontSize:12,color:G.bdr}}>No portals configured</div>}
  {(c.extraPw||[]).map((pwItem,i)=><div key={"x"+i} style={{display:"flex",gap:8,padding:"6px 10px",background:G.cyn+"0A",border:`1px solid ${G.cyn}22`,borderRadius:8,marginBottom:4}}>
    <span style={{fontSize:14,flexShrink:0}}>🔑</span>
    <div style={{flex:1,minWidth:0}}>
      <div style={{fontSize:10,color:G.cyn,fontWeight:700}}>{pwItem.type}{pwItem.label?" · "+pwItem.label:""}</div>
      <div style={{fontSize:11,color:G.mut}}>ID: <span style={{color:G.wh,fontFamily:"monospace"}}>{pwItem.id||"-"}</span></div>
      <div style={{fontSize:11,color:G.mut,display:"flex",alignItems:"center",gap:4}}>PW: <span style={{color:G.wh,fontFamily:"monospace"}}>{vpw[`${c.pan}_x${i}`]?pwItem.pw||"-":"********"}</span>
        <button onClick={()=>setVpw(s=>({...s,[`${c.pan}_x${i}`]:!s[`${c.pan}_x${i}`]}))} style={{background:"none",border:"none",cursor:"pointer",color:G.mut,fontSize:12,padding:0}}>{vpw[`${c.pan}_x${i}`]?"🙈":"👁"}</button>
      </div>
    </div>
  </div>)}
  </>;
  return <div>
    {editC&&<EditClient c={editC} onSave={cf=>{setClients(p=>p.map(c=>c.pan===cf.pan?cf:c));setEditC(null);toast("Client updated!");}} onX={()=>setEditC(null)} dd={dd}/>}
    {showImport&&<div style={{position:"fixed",inset:0,background:"#000C",zIndex:6000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:18,width:"min(560px,98%)",maxHeight:"90vh",overflow:"auto",boxShadow:"0 32px 80px #000C"}}>
        <div style={{padding:"16px 20px",borderBottom:`1px solid ${G.bdr}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
          <span style={{fontWeight:800,fontSize:15,color:G.wh}}>⬆️ Bulk Import Clients (CSV)</span>
          <button onClick={()=>{setShowImport(false);setImportResult(null);}} style={{background:"transparent",border:`1px solid ${G.bdr}`,color:G.mut,borderRadius:8,padding:"4px 11px",cursor:"pointer"}}>✕</button>
        </div>
        <div style={{padding:20,display:"flex",flexDirection:"column",gap:14}}>
          {!importResult?<>
            <div style={{fontSize:12,color:G.mut,lineHeight:1.6}}>
              1. Download the template CSV below.<br/>
              2. Fill one row per client — columns marked <b style={{color:G.amb}}>*</b> are required.<br/>
              3. Save as CSV and upload it here. Existing PANs are skipped automatically, so it's safe to re-upload.
            </div>
            <div style={{fontSize:11,color:G.mut,background:G.bg,border:`1px solid ${G.bdr}`,borderRadius:9,padding:"9px 11px"}}>
              <div><b style={{color:G.txt}}>Valid Client Types:</b> {dd.clientTypes?.join(", ")||"—"}</div>
              <div style={{marginTop:4}}><b style={{color:G.txt}}>Valid Sources:</b> {dd.sources?.join(", ")||"—"}</div>
            </div>
            <button onClick={downloadTemplate} style={{padding:"10px 16px",borderRadius:10,border:`1px solid ${G.cyn}55`,background:G.cyn+"12",color:G.cyn,fontWeight:700,fontSize:13,cursor:"pointer"}}>⬇️ Download Template CSV</button>
            <label style={{padding:"12px 16px",borderRadius:10,border:`2px dashed ${G.green}55`,background:G.green+"0A",color:G.green,fontWeight:700,fontSize:13,cursor:"pointer",textAlign:"center"}}>
              {importBusy?"⏳ Reading file...":"📤 Choose CSV File to Import"}
              <input type="file" accept=".csv,text/csv" onChange={handleImportFile} disabled={importBusy} style={{display:"none"}}/>
            </label>
          </>:<>
            <div style={{display:"flex",gap:10}}>
              <div style={{flex:1,background:G.green+"0F",border:`1px solid ${G.green}33`,borderRadius:10,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:20,fontWeight:800,color:G.green}}>{importResult.imported}</div><div style={{fontSize:11,color:G.mut}}>Imported</div></div>
              <div style={{flex:1,background:G.red+"0F",border:`1px solid ${G.red}33`,borderRadius:10,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:20,fontWeight:800,color:G.red}}>{importResult.skipped.length}</div><div style={{fontSize:11,color:G.mut}}>Skipped</div></div>
              <div style={{flex:1,background:G.card,border:`1px solid ${G.bdr}`,borderRadius:10,padding:"10px 12px",textAlign:"center"}}><div style={{fontSize:20,fontWeight:800,color:G.txt}}>{importResult.total}</div><div style={{fontSize:11,color:G.mut}}>Rows Read</div></div>
            </div>
            {importResult.skipped.length>0&&<div>
              <div style={{fontSize:12,fontWeight:700,color:G.mut,marginBottom:6}}>Skipped rows:</div>
              <div style={{maxHeight:180,overflowY:"auto",display:"flex",flexDirection:"column",gap:4}}>
                {importResult.skipped.map((s,i)=><div key={i} style={{display:"flex",justifyContent:"space-between",gap:8,fontSize:11,background:G.bg,border:`1px solid ${G.bdr}`,borderRadius:7,padding:"5px 9px"}}>
                  <span style={{color:G.mut,whiteSpace:"nowrap"}}>Row {s.row} · {s.pan}</span><span style={{color:G.red,textAlign:"right"}}>{s.reason}</span>
                </div>)}
              </div>
            </div>}
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>setImportResult(null)} style={{flex:1,padding:"10px",borderRadius:10,border:`1px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer",fontWeight:600}}>Import Another File</button>
              <button onClick={()=>{setShowImport(false);setImportResult(null);}} style={{flex:1,padding:"10px",borderRadius:10,border:"none",background:`linear-gradient(135deg,${G.g2},${G.green})`,color:"#fff",fontWeight:700,cursor:"pointer"}}>Done</button>
            </div>
          </>}
        </div>
      </div>
    </div>}
    <div style={{display:"flex",gap:10,marginBottom:14,alignItems:"center",flexWrap:"wrap"}}>
      <div style={{flex:1,position:"relative",minWidth:200}}>
        <span style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)",color:G.mut}}>🔍</span>
        <input value={q} onChange={e=>setQ(e.target.value)} placeholder="Search name, PAN, mobile..." style={{...IS,paddingLeft:34}}/>
      </div>
      <button onClick={()=>setShowImport(true)} style={{padding:"8px 14px",borderRadius:9,border:`1px solid ${G.green}55`,background:G.green+"12",color:G.green,fontWeight:700,fontSize:12,cursor:"pointer",whiteSpace:"nowrap"}}>⬆️ Import CSV</button>
      <div style={{display:"flex",gap:3,background:G.card,border:`1px solid ${G.bdr}`,borderRadius:9,padding:3}}>
        {[{id:"grid",icon:"⊞"},{id:"list",icon:"☰"}].map(v=><button key={v.id} onClick={()=>setView(v.id)} style={{padding:"6px 12px",borderRadius:7,border:"none",cursor:"pointer",background:view===v.id?`linear-gradient(135deg,${G.g2},${G.green})`:"transparent",color:view===v.id?"#fff":G.mut,fontSize:16,fontWeight:700}}>{v.icon}</button>)}
      </div>
      <span style={{fontSize:13,color:G.mut,whiteSpace:"nowrap"}}>{list.length} clients</span>
    </div>
    {view==="grid"
    ?<div className="auto-grid" style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(310px,1fr))",gap:13}}>
      {list.map(c=><div key={c.pan} style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:15,overflow:"hidden",transition:"border .2s"}} onMouseEnter={e=>e.currentTarget.style.borderColor=G.green+"55"} onMouseLeave={e=>e.currentTarget.style.borderColor=G.bdr}>
        <div style={{padding:"13px 15px",borderBottom:`1px solid ${G.bdr}`,display:"flex",gap:11,alignItems:"center"}}>
          <div style={{width:40,height:40,borderRadius:10,background:`linear-gradient(135deg,${G.g2},${G.green})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:16,fontWeight:800,color:"#fff",flexShrink:0}}>{c.name[0]}</div>
          <div style={{flex:1,minWidth:0}}><div style={{fontWeight:700,fontSize:13,color:G.wh,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.name}</div>{c.biz&&<div style={{fontSize:11,color:G.mut,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{c.biz}</div>}</div>
          <div style={{display:"flex",gap:5,alignItems:"center",flexShrink:0}}><Bdg label={c.status}/><button onClick={()=>setEditC(c)} style={{background:"transparent",border:`1px solid ${G.bdr}`,color:G.cyn,borderRadius:6,padding:"3px 7px",cursor:"pointer",fontSize:12}}>✏️</button><button onClick={()=>setClients(p=>p.map(x=>x.pan===c.pan?{...x,status:x.status==="Active"?"Inactive":"Active"}:x))} title="Toggle Status" style={{background:"transparent",border:`1px solid ${G.bdr}`,color:G.amb,borderRadius:6,padding:"3px 7px",cursor:"pointer",fontSize:11}}>{c.status==="Active"?"⏸":"▶"}</button><button onClick={()=>setDelC(c)} title="Delete Client" style={{background:"#450A0A",border:`1px solid ${G.red}44`,color:G.red,borderRadius:6,padding:"3px 7px",cursor:"pointer",fontSize:12}}>🗑</button></div>
        </div>
        <div style={{padding:"11px 15px",display:"grid",gridTemplateColumns:"1fr 1fr",gap:8}}>
          {[{l:"PAN",v:c.pan,mono:true,col:G.g3},{l:"Type",v:c.type},{l:"Mobile",v:c.mob},{l:"State",v:c.state},{l:"Source",v:c.src},{l:"Added",v:fd(c.added)}].map(f=><div key={f.l}><div style={{fontSize:10,color:G.mut,textTransform:"uppercase",letterSpacing:.5,fontWeight:600}}>{f.l}</div><div style={{fontSize:12,color:f.col||G.mut,marginTop:1,fontFamily:f.mono?"monospace":"inherit",fontWeight:f.mono?700:400}}>{f.v||"-"}</div></div>)}
        </div>
        {c.email&&<div style={{padding:"0 15px 9px",fontSize:11,color:G.mut,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>✉ {c.email}</div>}
        <div style={{padding:"9px 15px",borderTop:`1px solid ${G.bdr}`,display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div style={{display:"flex",gap:4,flexWrap:"wrap"}}>{PORTALS.filter(p=>c.portals[p.key]?.on).map(p=><span key={p.key} style={{fontSize:10,padding:"2px 7px",borderRadius:10,background:p.col+"18",color:p.col,fontWeight:600}}>{p.icon}</span>)}{!PORTALS.some(p=>c.portals[p.key]?.on)&&<span style={{fontSize:11,color:G.bdr}}>No portals</span>}</div>
          <button onClick={()=>setOpP(s=>({...s,[c.pan]:!s[c.pan]}))} style={{fontSize:11,padding:"3px 9px",borderRadius:7,border:`1px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer",fontWeight:600,whiteSpace:"nowrap"}}>{opP[c.pan]?"🙈 Hide":"🔑 View"}</button>
        </div>
        {opP[c.pan]&&<div style={{padding:"9px 15px",borderTop:`1px solid ${G.bdr}`,background:G.bg}}><PortBlock c={c}/></div>}
      </div>)}
    </div>
    :<Crd sty={{padding:0,overflow:"hidden"}}><div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
      <thead><tr style={{background:G.bg}}>{["Name","PAN","Type","Mobile","State","Source","Status","Portals",""].map(h=><th key={h} style={{padding:"9px 12px",textAlign:"left",color:G.mut,fontWeight:600,fontSize:11,whiteSpace:"nowrap",borderBottom:`1px solid ${G.bdr}`}}>{h}</th>)}</tr></thead>
      <tbody>{list.map((c,i)=><tr key={c.pan} style={{borderTop:`1px solid ${G.bdr}`,background:i%2?"#0F1D1408":"transparent"}}>
        <td style={{padding:"9px 12px"}}><div style={{fontWeight:700,color:G.txt,whiteSpace:"nowrap"}}>{c.name}</div>{c.biz&&<div style={{fontSize:11,color:G.mut}}>{c.biz}</div>}</td>
        <td style={{padding:"9px 12px",color:G.g3,fontFamily:"monospace",fontSize:11,fontWeight:700}}>{c.pan}</td>
        <td style={{padding:"9px 12px",color:G.mut,whiteSpace:"nowrap"}}>{c.type}</td>
        <td style={{padding:"9px 12px",color:G.mut}}>{c.mob}</td>
        <td style={{padding:"9px 12px",color:G.mut,whiteSpace:"nowrap"}}>{c.state}</td>
        <td style={{padding:"9px 12px",color:G.mut}}>{c.src}</td>
        <td style={{padding:"9px 12px"}}><Bdg label={c.status}/></td>
        <td style={{padding:"9px 12px"}}><div style={{display:"flex",gap:3}}>{PORTALS.filter(p=>c.portals[p.key]?.on).map(p=><span key={p.key} style={{fontSize:11,padding:"1px 6px",borderRadius:8,background:p.col+"18",color:p.col,fontWeight:600}}>{p.icon}</span>)}</div></td>
        <td style={{padding:"9px 12px"}}>
          <div style={{display:"flex",gap:5}}>
            <button onClick={()=>setEditC(c)} style={{fontSize:11,padding:"3px 9px",borderRadius:7,border:`1px solid ${G.bdr}`,background:"transparent",color:G.cyn,cursor:"pointer",fontWeight:600}}>✏️ Edit</button><button onClick={()=>setClients(p=>p.map(x=>x.pan===c.pan?{...x,status:x.status==="Active"?"Inactive":"Active"}:x))} style={{fontSize:11,padding:"3px 9px",borderRadius:7,border:`1px solid ${G.bdr}`,background:"transparent",color:G.amb,cursor:"pointer"}}>{c.status==="Active"?"⏸ Deactivate":"▶ Activate"}</button>
            <button onClick={()=>setOpP(s=>({...s,[c.pan]:!s[c.pan]}))} style={{fontSize:11,padding:"3px 9px",borderRadius:7,border:`1px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer"}}>🔑</button>
            <button onClick={()=>setDelC(c)} style={{fontSize:11,padding:"3px 9px",borderRadius:7,border:`1px solid ${G.red}44`,background:"#450A0A",color:G.red,cursor:"pointer",fontWeight:600}}>🗑 Delete</button>
          </div>
          {opP[c.pan]&&<div style={{marginTop:7}}><PortBlock c={c}/></div>}
        </td>
      </tr>)}</tbody>
    </table></div></Crd>}
    {delC&&<div style={{position:"fixed",inset:0,background:"#000C",zIndex:6000,display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:16,padding:28,width:360,boxShadow:"0 20px 60px #000C",textAlign:"center"}}>
        <div style={{fontSize:36,marginBottom:12}}>🗑</div>
        <div style={{fontWeight:700,fontSize:16,color:G.wh,marginBottom:8}}>Delete this client?</div>
        <div style={{fontSize:13,color:G.mut,marginBottom:8,padding:"8px 14px",background:G.red+"10",border:`1px solid ${G.red}33`,borderRadius:9}}>{delC.name} · {delC.pan}</div>
        <div style={{fontSize:11,color:G.amb,marginBottom:20}}>This removes the client record and their saved credentials. Related work items and invoices are kept.</div>
        <div style={{display:"flex",gap:10}}>
          <button onClick={doDeleteClient} style={{flex:1,padding:"11px",borderRadius:10,border:"none",cursor:"pointer",background:G.red,color:"#fff",fontWeight:700,fontSize:14}}>Yes, Delete</button>
          <button onClick={()=>setDelC(null)} style={{flex:1,padding:"11px",borderRadius:10,border:`1px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer",fontWeight:600}}>Cancel</button>
        </div>
      </div>
    </div>}
  </div>;
}

// ─── Work Tracker ─────────────────────────────────────────────────────────────


// ─── Invoice seed data ────────────────────────────────────────────────────────
const SEED_INVOICES = [
  {id:"INV-2025-001",date:"2025-04-15",pan:"ABCPK1234A",clientName:"Rajesh Kumar",workId:5,service:"GST Registration",fy:"2025-26",amount:1500,gst:18,gstAmt:270,total:1770,status:"Paid",note:""},
  {id:"INV-2025-002",date:"2025-04-20",pan:"CDERM7890C",clientName:"Tech Solutions Pvt Ltd",workId:6,service:"GST Return Filing",fy:"2025-26",amount:4000,gst:18,gstAmt:720,total:4720,status:"Unpaid",note:"Reminder sent"},
  {id:"INV-2024-001",date:"2024-06-30",pan:"ABCPK1234A",clientName:"Rajesh Kumar",workId:1,service:"ITR Filing",fy:"2024-25",amount:5000,gst:18,gstAmt:900,total:5900,status:"Paid",note:""},
  {id:"INV-2024-002",date:"2024-07-05",pan:"BCDQL5678B",clientName:"Priya Sharma",workId:2,service:"GST Return Filing",fy:"2024-25",amount:3000,gst:18,gstAmt:540,total:3540,status:"Partial",note:"Balance pending"},
  {id:"INV-2024-003",date:"2024-07-10",pan:"CDERM7890C",clientName:"Tech Solutions Pvt Ltd",workId:3,service:"Tax Audit",fy:"2024-25",amount:25000,gst:18,gstAmt:4500,total:29500,status:"Unpaid",note:""},
];
const SEED_RECEIPTS = [
  {id:"RCP-001",date:"2025-04-17",invId:"INV-2025-001",pan:"ABCPK1234A",clientName:"Rajesh Kumar",invAmt:1770,received:1770,mode:"UPI",ref:"UPI123456",note:"Full payment"},
  {id:"RCP-002",date:"2024-07-01",invId:"INV-2024-001",pan:"ABCPK1234A",clientName:"Rajesh Kumar",invAmt:5900,received:5900,mode:"Bank Transfer",ref:"NEFT789012",note:""},
  {id:"RCP-003",date:"2024-07-12",invId:"INV-2024-002",pan:"BCDQL5678B",clientName:"Priya Sharma",invAmt:3540,received:2000,mode:"Cash",ref:"",note:"Partial - balance 1540"},
];

// ─── Invoice Module ───────────────────────────────────────────────────────────
function InvoiceModule({invoices,setInvoices,receipts,setReceipts,clients,works,dd,toast,firmSettings,setFirmSettings}){
  const[fy,setFy]=useState(getCurrentFY());
  const[status,setStatus]=useState("All");
  const[q,setQ]=useState("");
  const[showForm,setShowForm]=useState(false);
  const[showPrint,setShowPrint]=useState(null);
  const[editInv,setEditInv]=useState(null);

  const filtered=useMemo(()=>{
    let list=invoices;
    if(fy!=="All") list=list.filter(i=>i.fy===fy);
    if(status!=="All") list=list.filter(i=>i.status===status);
    if(q){const lq=q.toLowerCase();list=list.filter(i=>i.clientName.toLowerCase().includes(lq)||i.id.toLowerCase().includes(lq)||i.pan.toLowerCase().includes(lq));}
    return list.sort((a,b)=>new Date(b.date)-new Date(a.date));
  },[invoices,fy,status,q]);

  const totBilled=filtered.reduce((s,i)=>s+i.total,0);
  const totPaid=filtered.filter(i=>i.status==="Paid").reduce((s,i)=>s+i.total,0);
  const totUnpaid=filtered.filter(i=>i.status!=="Paid").reduce((s,i)=>s+i.total,0);

  const genInvId=()=>{
    const getFySegment = () => {
      const now = new Date();
      const yr = now.getFullYear();
      const month = now.getMonth();
      if (month >= 3) {
        return `${yr}-${String(yr + 1).slice(-2)}`;
      } else {
        return `${yr - 1}-${String(yr).slice(-2)}`;
      }
    };
    if (!invoices || !invoices.length) {
      return `FTM/${getFySegment()}/10001`;
    }
    const ftmInvoices = invoices.filter(inv => inv.id && inv.id.startsWith("FTM/"));
    if (!ftmInvoices.length) {
      return `FTM/${getFySegment()}/10001`;
    }
    ftmInvoices.sort((a, b) => b.id.localeCompare(a.id));
    const lastId = ftmInvoices[0].id;
    const parts = lastId.split('/');
    if (parts.length >= 3) {
      const prefix = parts[0];
      const fySeg = parts[1];
      const lastNumStr = parts[parts.length - 1];
      const nextNum = parseInt(lastNumStr, 10) + 1;
      const nextNumStr = String(nextNum).padStart(lastNumStr.length, '0');
      return `${prefix}/${fySeg}/${nextNumStr}`;
    }
    return `FTM/${getFySegment()}/10001`;
  }

  const deleteInv=id=>{
    setInvoices(p=>p.filter(i=>i.id!==id));
    toast("Invoice deleted");
  };

  const markPaid=id=>{
    setInvoices(p=>p.map(i=>i.id===id?{...i,status:"Paid"}:i));
    toast("Marked as Paid ✓");
  };

  return <div style={{display:"flex",flexDirection:"column",gap:16,animation:"fadeUp .3s ease"}}>
    {showForm&&<InvoiceForm invoices={invoices} setInvoices={setInvoices} clients={clients} works={works} dd={dd} toast={toast} onClose={()=>setShowForm(false)} genId={genInvId} editInv={editInv} setEditInv={setEditInv} receipts={receipts}/>}
    {showPrint&&<InvoicePrint inv={showPrint} clients={clients} firmSettings={firmSettings} onClose={()=>setShowPrint(null)} toast={toast}/>}

    {/* Summary KPIs */}
    <div className="kpi-grid-4" style={{display:"grid",gridTemplateColumns:"repeat(4,1fr)",gap:12}}>
      {[
        {l:"Total Invoices",v:filtered.length,icon:"🧾",col:G.green},
        {l:"Total Billed",v:inr(totBilled),icon:"💳",col:G.cyn},
        {l:"Collected",v:inr(totPaid),icon:"✅",col:G.g3},
        {l:"Outstanding",v:inr(totUnpaid),icon:"⚠️",col:G.red},
      ].map(k=><div key={k.l} style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"13px 16px",display:"flex",gap:10,alignItems:"center"}}>
        <div style={{width:38,height:38,borderRadius:10,background:k.col+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>{k.icon}</div>
        <div><div style={{fontWeight:800,fontSize:16,color:k.col}}>{k.v}</div><div style={{fontSize:10,color:G.mut,marginTop:2}}>{k.l}</div></div>
      </div>)}
    </div>

    {/* Filters + New button */}
    <div style={{display:"flex",gap:10,flexWrap:"wrap",alignItems:"center"}}>
      <div style={{display:"flex",gap:5,padding:"3px 8px",background:G.card,border:`1px solid ${G.bdr}`,borderRadius:20,alignItems:"center"}}>
        <span style={{fontSize:11,color:G.mut,fontWeight:600}}>📅 FY:</span>
        {["All",...dd.fyOptions].map(f=><button key={f} onClick={()=>setFy(f)} style={{padding:"4px 10px",borderRadius:14,border:"none",cursor:"pointer",fontSize:12,fontWeight:fy===f?700:500,background:fy===f?`linear-gradient(135deg,${G.g2},${G.green})`:"transparent",color:fy===f?"#fff":G.mut}}>{f}</button>)}
      </div>
      <div style={{display:"flex",gap:5}}>
        {["All","Paid","Unpaid","Partial"].map(s=>{
          const cols={Paid:G.g3,Unpaid:G.red,Partial:G.amb,All:G.green};
          return <button key={s} onClick={()=>setStatus(s)} style={{padding:"5px 12px",borderRadius:18,border:`1.5px solid ${status===s?cols[s]:G.bdr}`,cursor:"pointer",fontSize:12,fontWeight:600,background:status===s?cols[s]+"18":"transparent",color:status===s?cols[s]:G.mut}}>{s}</button>;
        })}
      </div>
      <div style={{position:"relative",flex:1,minWidth:160}}>
        <span style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:G.mut}}>🔍</span>
        <input value={q} onChange={e=>setQ(e.target.value)} placeholder="Search client, invoice no..." style={{...IS,paddingLeft:30,fontSize:12}}/>
      </div>
      <button onClick={()=>{setEditInv(null);setShowForm(true);}} style={{padding:"8px 18px",borderRadius:10,border:"none",cursor:"pointer",background:`linear-gradient(135deg,${G.g2},${G.green})`,color:"#fff",fontWeight:700,fontSize:13,whiteSpace:"nowrap"}}>➕ New Invoice</button>
    </div>

    {/* Invoice Table */}
    <div style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:14,overflow:"hidden"}}>
      <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
        <thead><tr style={{background:G.bg}}>
          {["Invoice No.","Date","Client","Service","FY","Amount","GST","Total","Status",""].map(h=>(
            <th key={h} style={{padding:"10px 12px",textAlign:"left",color:G.mut,fontWeight:600,fontSize:11,whiteSpace:"nowrap",borderBottom:`1px solid ${G.bdr}`}}>{h}</th>
          ))}
        </tr></thead>
        <tbody>
          {filtered.length===0?<tr><td colSpan={10} style={{padding:32,textAlign:"center",color:G.bdr}}>No invoices found</td></tr>
          :filtered.map((inv,i)=>{
            const sc={Paid:{bg:"#14532D",fg:"#4ADE80"},Unpaid:{bg:"#450A0A",fg:"#FCA5A5"},Partial:{bg:"#431407",fg:"#FCD34D"}};
            const s=sc[inv.status]||{bg:G.bdr,fg:G.mut};
            return <tr key={inv.id} style={{borderTop:`1px solid ${G.bdr}`,background:i%2?G.card+"80":"transparent"}}>
              <td style={{padding:"10px 12px",fontWeight:700,color:G.cyn,fontFamily:"monospace",fontSize:11}}>{inv.id}</td>
              <td style={{padding:"10px 12px",color:G.mut,whiteSpace:"nowrap"}}>{inv.date?new Date(inv.date).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"}):"-"}</td>
              <td style={{padding:"10px 12px"}}>
                <div style={{fontWeight:600,color:G.txt,whiteSpace:"nowrap"}}>{inv.clientName}</div>
                <div style={{fontSize:10,color:G.mut,fontFamily:"monospace"}}>{inv.pan}</div>
              </td>
              <td style={{padding:"10px 12px",color:G.mut,maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{inv.service}</td>
              <td style={{padding:"10px 12px",color:G.mut}}>{inv.fy}</td>
              <td style={{padding:"10px 12px",color:G.txt,fontWeight:600}}>{inr(inv.amount)}</td>
              <td style={{padding:"10px 12px",color:G.mut}}>{inv.gst}%</td>
              <td style={{padding:"10px 12px",color:G.green,fontWeight:700,whiteSpace:"nowrap"}}>{inr(inv.total)}</td>
              <td style={{padding:"10px 12px"}}>
                <span style={{background:s.bg,color:s.fg,padding:"2px 9px",borderRadius:20,fontSize:11,fontWeight:700,whiteSpace:"nowrap"}}>{inv.status}</span>
              </td>
              <td style={{padding:"10px 12px"}}>
                <div style={{display:"flex",gap:5,flexWrap:"nowrap"}}>
                  <button onClick={()=>setShowPrint(inv)} title="Print Invoice" style={{background:"transparent",border:`1px solid ${G.bdr}`,color:G.cyn,borderRadius:6,padding:"3px 8px",cursor:"pointer",fontSize:12}}>🖨</button>
                  <button onClick={()=>{setEditInv(inv);setShowForm(true);}} title="Edit" style={{background:"transparent",border:`1px solid ${G.bdr}`,color:G.amb,borderRadius:6,padding:"3px 8px",cursor:"pointer",fontSize:12}}>✏️</button>
                  {inv.status!=="Paid"&&<button onClick={()=>markPaid(inv.id)} title="Mark Paid" style={{background:"transparent",border:`1px solid ${G.g3}44`,color:G.g3,borderRadius:6,padding:"3px 8px",cursor:"pointer",fontSize:11,fontWeight:700}}>✓ Paid</button>}
                  <button onClick={()=>deleteInv(inv.id)} title="Delete" style={{background:"#450A0A",border:`1px solid ${G.red}44`,color:G.red,borderRadius:6,padding:"3px 7px",cursor:"pointer",fontSize:11,fontWeight:700}}>🗑</button>
                </div>
              </td>
            </tr>;
          })}
        </tbody>
        {filtered.length>0&&<tfoot>
          <tr style={{background:G.bg,borderTop:`2px solid ${G.green}33`}}>
            <td colSpan={5} style={{padding:"9px 12px",fontWeight:800,color:G.wh,fontSize:12}}>TOTALS ({filtered.length} invoices)</td>
            <td style={{padding:"9px 12px",color:G.txt,fontWeight:700}}>{inr(filtered.reduce((s,i)=>s+i.amount,0))}</td>
            <td/>
            <td style={{padding:"9px 12px",color:G.green,fontWeight:800}}>{inr(totBilled)}</td>
            <td colSpan={2}/>
          </tr>
        </tfoot>}
      </table></div>
    </div>
  </div>;
}

// ─── Invoice Form (Create / Edit) ─────────────────────────────────────────────
function InvoiceForm({invoices,setInvoices,clients,works,dd,toast,onClose,genId,editInv,setEditInv,receipts}){
  const isEdit=!!editInv;
  const blank={id:"",date:td(),pan:"",clientName:"",workId:"",service:"",fy:dd.fyOptions[0]||"2025-26",items:[{service:"",desc:"",qty:"1",rate:""}],discount:"",gst:18,status:"Unpaid",note:""};
  const[f,setF]=useState(isEdit?{...editInv,
    items:(editInv.items&&editInv.items.length?editInv.items.map(it=>({service:it.service,desc:it.desc||"",qty:String(it.qty??1),rate:String(it.rate??"")})):[{service:editInv.service||"",desc:"",qty:"1",rate:String(editInv.amount||"")}]),
    discount:String(editInv.discount||""),gst:editInv.gst}:blank);
  const[err,setErr]=useState({});
  const[manualId,setManualId]=useState(isEdit);
  const lastInv = useMemo(() => {
    if (!invoices || !invoices.length) return "None";
    const ftm = [...invoices].filter(i => i.id && i.id.startsWith("FTM/"));
    ftm.sort((a, b) => b.id.localeCompare(a.id));
    return ftm.length ? ftm[0].id : "None";
  }, [invoices]);

  const selClient=clients.find(c=>c.pan===f.pan);
  const clientWorks=works.filter(w=>w.pan===f.pan);
  const subTotal=(f.items||[]).reduce((s,it)=>{
    const amt=it.type==="reimbursement"?(Number(it.rate)||0):(Number(it.qty)||0)*(Number(it.rate)||0);
    return s+amt;
  },0);
  const reimbTotal=(f.items||[]).filter(it=>it.type==="reimbursement").reduce((s,it)=>s+(Number(it.rate)||0),0);
  const discount=Math.min(Number(f.discount)||0,subTotal);
  const taxable=(subTotal-reimbTotal)-discount;  // GST only on non-reimbursement
  const gstAmt=Math.round(taxable*(Number(f.gst)||0)/100);
  const total=taxable+gstAmt+reimbTotal;
  const setItem=(idx,patch)=>setF(p=>({...p,items:p.items.map((it,i)=>i===idx?{...it,...patch}:it)}));
  const addItem=()=>setF(p=>({...p,items:[...(p.items||[]),{service:"",desc:"",qty:"1",rate:"",type:"service"}]}));
  const addReimb=()=>setF(p=>({...p,items:[...(p.items||[]),{service:"Reimbursement",desc:"",qty:"1",rate:"",type:"reimbursement"}]}));
  const delItem=(idx)=>setF(p=>({...p,items:p.items.length>1?p.items.filter((_,i)=>i!==idx):p.items}));

  const save=()=>{
    const e={};
    if(!f.pan)e.pan="Select client";
    const items=(f.items||[]).filter(it=>(it.service&&it.service.trim())||it.rate);
    if(!items.length)e.service="Add at least one service item";
    else if(items.some(it=>!it.service||!it.service.trim()))e.service="Every item needs a service name";
    else if(items.some(it=>!it.rate||isNaN(Number(it.rate))))e.service="Every item needs a valid rate";
    if(!f.date)e.date="Select date";
    if(manualId&&f.id){
      const dup=invoices.find(i=>i.id===f.id.trim()&&(!isEdit||i.id!==editInv?.id));
      if(dup)e.id="Invoice No. already exists! Use a different number.";
    }
    setErr(e);
    if(Object.keys(e).length){toast("Fix errors first","err");return;}
    const finalId=manualId&&f.id.trim()?f.id.trim():genId();
    const finalInvId=isEdit?f.id:finalId;
    // Status is never taken from a free-typed field — it's recalculated from whatever
    // receipts already exist against this invoice, so editing an invoice (e.g. fixing
    // a discount) can never silently flip it to "Paid" without money actually received.
    const computedStatus=computeInvStatus({id:finalInvId,total},receipts);
    const inv={
      id:finalInvId,
      date:f.date, pan:f.pan, clientName:selClient?.name||f.clientName,
      workId:f.workId,
      items:items.map(it=>({service:it.service.trim(),desc:(it.desc||"").trim(),qty:Number(it.qty)||1,rate:Number(it.rate)||0,amount:(Number(it.qty)||1)*(Number(it.rate)||0)})),
      service:items.map(it=>it.service.trim()).join(", "), fy:f.fy,
      subTotal, discount, amount:taxable, gst:Number(f.gst),
      gstAmt, total, status:computedStatus, note:f.note,
    };
    if(isEdit){setInvoices(p=>p.map(i=>i.id===inv.id?inv:i));}
    else{setInvoices(p=>[inv,...p]);}
    toast(isEdit?"Invoice updated!":"Invoice "+inv.id+" created!");
    onClose();
  };

  return <div style={{position:"fixed",inset:0,background:"#000C",zIndex:5000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
    <div style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:18,width:"min(620px,98%)",maxHeight:"92vh",overflow:"auto",boxShadow:"0 32px 80px #000C"}}>
      <div style={{padding:"16px 20px",borderBottom:`1px solid ${G.bdr}`,display:"flex",justifyContent:"space-between",alignItems:"center",position:"sticky",top:0,background:G.surf,zIndex:1}}>
        <span style={{fontWeight:800,fontSize:15,color:G.wh}}>{isEdit?"✏️ Edit Invoice":"🧾 New Invoice - "+genId()}</span>
        <button onClick={onClose} style={{background:"transparent",border:`1px solid ${G.bdr}`,color:G.mut,borderRadius:8,padding:"4px 11px",cursor:"pointer"}}>✕</button>
      </div>
      <div style={{padding:20,display:"flex",flexDirection:"column",gap:12}}>
        {/* Invoice Number */}
        <div style={{background:G.card,border:`1px solid ${G.green}33`,borderRadius:10,padding:"10px 14px"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:8}}>
            <div style={{display:"flex",alignItems:"center",gap:6}}>
              <span style={{fontSize:12,fontWeight:700,color:G.txt}}>Invoice Number</span>
              <span style={{fontSize:11,color:G.mut}}>| Last: <span style={{color:G.wh,fontWeight:700,fontFamily:"monospace"}}>{lastInv}</span></span>
            </div>
            <div style={{display:"flex",gap:6,alignItems:"center"}}>
              <span style={{fontSize:11,color:G.mut}}>Auto-generate</span>
              <div onClick={()=>{setManualId(p=>!p);if(manualId)setF(pf=>({...pf,id:""}));}}
                style={{width:36,height:20,borderRadius:10,background:manualId?G.bdr:G.green,cursor:"pointer",position:"relative",transition:"background .2s",flexShrink:0}}>
                <div style={{position:"absolute",top:2,width:16,height:16,borderRadius:"50%",background:"#fff",transition:"left .2s",left:manualId?"2px":"18px"}}/>
              </div>
              <span style={{fontSize:11,color:G.mut}}>Manual</span>
            </div>
          </div>
          {manualId
            ?<div>
              <div style={{display:"flex",gap:6,alignItems:"center",marginBottom:6}}>
                <div style={{flex:1}}>
                  <I val={f.id||""} set={v=>setF({...f,id:v.toUpperCase()})} ph="e.g. FTM/2025-26/10032" mono sty={{borderColor:err.id?G.red:G.bdr}}/>
                </div>
                <button 
                  onClick={(e)=>{e.preventDefault();setF(pf=>({...pf,id:genId()}));}} 
                  style={{
                    background: `${G.green}18`,
                    border: `1.5px solid ${G.green}40`,
                    color: G.green,
                    borderRadius: 6,
                    padding: "6px 12px",
                    cursor: "pointer",
                    fontSize: 11,
                    fontWeight: 700,
                    whiteSpace: "nowrap",
                    transition: "all 0.2s"
                  }}
                  title="Auto-fill the next invoice serial number"
                >📋 Next No.</button>
              </div>
              {err.id&&<span style={{fontSize:11,color:G.red,marginTop:3,display:"block"}}>⚠️ {err.id}</span>}
              <div style={{fontSize:10,color:G.mut,marginTop:4}}>Enter a unique invoice number, or click "Next No." to auto-fill.</div>
            </div>
            :<div style={{padding:"6px 10px",background:G.surf,borderRadius:8,border:`1px solid ${G.bdr}`,fontFamily:"monospace",fontSize:13,color:G.cyn,fontWeight:700}}>{genId()}</div>}
        </div>

        {/* Client selector */}
        <F label="Select Client" req>
          <PanPick clients={clients} val={f.pan} set={v=>{
            const c=clients.find(x=>x.pan===v);
            setF(p=>({...p,pan:v,clientName:c?.name||""}));
          }}/>
          {err.pan&&<span style={{fontSize:11,color:G.red}}>{err.pan}</span>}
        </F>
        {/* Work assignment link */}
        {clientWorks.length>0&&<F label="Link to Work Assignment (optional)">
          <select value={f.workId} onChange={e=>{
            const w=works.find(x=>x.id===Number(e.target.value));
            setF(p=>{const its=[...(p.items||[])];if(w?.svc){if(!its.length)its.push({service:w.svc,desc:"",qty:"1",rate:""});else its[0]={...its[0],service:w.svc};}return {...p,workId:e.target.value?Number(e.target.value):"",service:w?.svc||p.service,fy:w?.fy||p.fy,items:its};});
          }} style={{...IS,cursor:"pointer"}}>
            <option value="">- Select work (auto-fills service & FY) -</option>
            {clientWorks.map(w=><option key={w.id} value={w.id} style={{background:"#0B1610"}}>{w.svc} - {w.fy} - {w.status}</option>)}
          </select>
        </F>}
        <div style={{display:"flex",gap:12}}>
          <F label="Invoice Date" req w="calc(50% - 6px)">
            <I val={f.date} set={v=>setF({...f,date:v})} type="date" sty={{borderColor:err.date?G.red:G.bdr}}/>
          </F>
          <F label="Financial Year" w="calc(50% - 6px)">
            <S val={f.fy} set={v=>setF({...f,fy:v})} opts={dd.fyOptions}/>
          </F>
        </div>
        <F label="Service Items (add multiple)" req>
          {(f.items||[]).map((it,idx)=>{
            const isReimb=it.type==="reimbursement";
            return <div key={idx} style={{border:`1px solid ${isReimb?G.amb:G.bdr}`,borderRadius:10,padding:"10px 12px",marginBottom:8,background:isReimb?G.amb+"08":G.card}}>
              <div style={{display:"flex",gap:8,alignItems:"center",marginBottom:8}}>
                <span style={{fontSize:11,fontWeight:800,color:isReimb?G.amb:G.mut,flexShrink:0}}>{isReimb?"REIMB":"SVC"} #{idx+1}</span>
                {isReimb
                  ?<div style={{flex:1,fontSize:12,fontWeight:700,color:G.amb}}>Reimbursement (No GST, Qty fixed = 1)</div>
                  :<div style={{flex:1}}>
                    <select
                      value={it.service && dd.billingItems && dd.billingItems.some(bi => bi.split("|")[0] === it.service) 
                        ? dd.billingItems.find(bi => bi.split("|")[0] === it.service) 
                        : it.service}
                      onChange={e => {
                        const val = e.target.value;
                        if (val.includes("|")) {
                          const [name, price] = val.split("|");
                          setItem(idx, { service: name, rate: price });
                        } else {
                          setItem(idx, { service: val });
                        }
                      }}
                      style={{...IS, cursor:"pointer", color:it.service?G.wh:G.mut}}
                    >
                      <option value="">Select billing item...</option>
                      {(dd.billingItems || []).map(o => {
                        const [name, price] = o.split("|");
                        return <option key={o} value={o} style={{background:"#0B1610"}}>
                          {name} {price ? `(₹${price})` : ""}
                        </option>;
                      })}
                    </select>
                  </div>}
                {f.items.length>1&&<button onClick={()=>delItem(idx)} style={{background:"transparent",border:`1px solid ${G.red}55`,color:G.red,borderRadius:8,padding:"6px 10px",cursor:"pointer",flexShrink:0}}>🗑</button>}
              </div>
              <div style={{marginBottom:8}}>
                <I val={it.desc} set={v=>setItem(idx,{desc:v})} ph={isReimb?"Reimbursement description (e.g. Travel Expenses, Courier Charges)":"Optional description - shows in ( brackets ) on bill"}/>
              </div>
              <div style={{display:"flex",gap:8,alignItems:"center"}}>
                {!isReimb&&<div style={{flex:1}}>
                  <div style={{fontSize:10,color:G.mut,fontWeight:700,marginBottom:3}}>QTY / NOS.</div>
                  <I val={it.qty} set={v=>setItem(idx,{qty:v.replace(/[^0-9.]/g,"")})} ph="1"/>
                </div>}
                <div style={{flex:2}}>
                  <div style={{fontSize:10,color:G.mut,fontWeight:700,marginBottom:3}}>{isReimb?"REIMBURSEMENT AMOUNT (₹)":"RATE (₹)"}</div>
                  <I val={it.rate} set={v=>setItem(idx,{rate:v.replace(/[^0-9.]/g,"")})} ph="0" sty={{borderColor:isReimb?G.amb+"99":G.bdr}}/>
                </div>
                <div style={{flex:1,textAlign:"right"}}>
                  <div style={{fontSize:10,color:G.mut,fontWeight:700,marginBottom:3}}>AMOUNT</div>
                  <div style={{fontSize:14,fontWeight:800,color:isReimb?G.amb:G.txt,padding:"8px 0"}}>{inr(isReimb?(Number(it.rate)||0):(Number(it.qty)||0)*(Number(it.rate)||0))}</div>
                </div>
              </div>
              {isReimb&&<div style={{fontSize:10,color:G.amb,marginTop:4,fontStyle:"italic"}}>* No GST applied on reimbursement amounts</div>}
            </div>;
          })}
          <div style={{display:"flex",gap:8}}>
            <button onClick={addItem} style={{flex:1,padding:"9px",borderRadius:10,border:`1px dashed ${G.green}66`,background:G.green+"0A",color:G.green,cursor:"pointer",fontWeight:700,fontSize:12}}>＋ Add Service / Item</button>
            <button onClick={addReimb} style={{flex:1,padding:"9px",borderRadius:10,border:`1px dashed ${G.amb}66`,background:G.amb+"0A",color:G.amb,cursor:"pointer",fontWeight:700,fontSize:12}}>＋ Add Reimbursement</button>
          </div>
          {err.service&&<span style={{fontSize:11,color:G.red}}>{err.service}</span>}
        </F>
        <div style={{display:"flex",gap:12}}>
          <F label="Discount (₹)" w="calc(40% - 6px)">
            <I val={f.discount} set={v=>setF({...f,discount:v.replace(/[^0-9.]/g,"")})} ph="0"/>
          </F>
          <F label="GST %" w="calc(30% - 6px)">
            <S val={String(f.gst)} set={v=>setF({...f,gst:Number(v)})} opts={["0","5","12","18","28"]}/>
          </F>
          <F label="Status" w="calc(30% - 6px)">
            {(()=>{const st=isEdit?computeInvStatus({id:editInv.id,total},receipts):"Unpaid";const c={Paid:{bg:"#14532D",fg:"#4ADE80"},Partial:{bg:"#431407",fg:"#FCD34D"},Unpaid:{bg:"#450A0A",fg:"#FCA5A5"}}[st];return <div style={{...IS,display:"flex",alignItems:"center",cursor:"not-allowed"}}>
              <span style={{background:c.bg,color:c.fg,padding:"2px 10px",borderRadius:20,fontSize:12,fontWeight:700}}>{st}</span>
              <span style={{fontSize:10,color:G.mut,marginLeft:8}}>auto, from Receipts</span>
            </div>;})()}
          </F>
        </div>
        {isEdit&&<div style={{fontSize:11,color:G.mut,marginTop:-4}}>💡 Status is calculated from actual payment receipts recorded against this invoice — record or edit payments from the Receipts tab, not here.</div>}
        {/* Live total preview */}
        {subTotal>0&&<div style={{display:"flex",gap:16,padding:"10px 14px",background:G.green+"0A",border:`1px solid ${G.green}22`,borderRadius:10,flexWrap:"wrap"}}>
          {[["Sub Total",inr(subTotal),G.txt],...(discount>0?[["Discount","- "+inr(discount),G.red],["Taxable",inr(taxable),G.txt]]:[]),[`GST (${f.gst}%)`,inr(gstAmt),G.amb],["Total Invoice",inr(total),G.green]].map(([l,v,col])=>(
            <div key={l} style={{display:"flex",flexDirection:"column",gap:2}}>
              <span style={{fontSize:10,color:G.mut,fontWeight:600,textTransform:"uppercase"}}>{l}</span>
              <span style={{fontSize:15,fontWeight:800,color:col}}>{v}</span>
            </div>
          ))}
        </div>}
        <F label="Notes">
          <textarea value={f.note} onChange={e=>setF({...f,note:e.target.value})} rows={2} style={{...IS,resize:"vertical"}}/>
        </F>
      </div>
      <div style={{padding:"14px 20px",borderTop:`1px solid ${G.bdr}`,display:"flex",gap:10}}>
        <button onClick={save} style={{flex:1,padding:"12px",borderRadius:10,border:"none",cursor:"pointer",background:`linear-gradient(135deg,${G.g2},${G.green})`,color:"#fff",fontWeight:700,fontSize:14}}>{isEdit?"💾 Update Invoice":"🧾 Create Invoice"}</button>
        <button onClick={onClose} style={{padding:"12px 18px",borderRadius:10,border:`1px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer"}}>Cancel</button>
      </div>
    </div>
  </div>;
}


// ─── Invoice Print View ───────────────────────────────────────────────────────
function InvoicePrint({inv,clients,firmSettings,onClose,toast}){
  const [printStatus, setPrintStatus] = useState(true);
  const cl=clients.find(c=>c.pan===inv.pan);
  const F=firmSettings||{};
  const isGST=(inv.gst||0)>0;
  const invType=isGST?(F.invoiceType||"TAX INVOICE"):"INVOICE";
  const copyLabel=F.copyLabel||"ORIGINAL FOR RECIPIENT";
  const emptyRows=Number(F.emptyRows??3);
  const clrP=F.clrPrimary||"#1B5E20";
  const clrA=F.clrAccent||"#F1F8E9";
  const clrB=F.clrBorder||"#C8E6C9";
  const clrT=F.clrTotal||"#F1F8E9";
  // Base sizes bumped ~25% larger than the original design, then the user's
  // "Invoice Scale" slider (Firm Settings) multiplies on top of that so they
  // can fine-tune further without editing raw font-size numbers.
  const invScale=(F.invoiceScale??100)/100;
  const bFsz=Math.round((F.bodyFontSz||15)*invScale);
  const tFsz=Math.round((F.titleFontSz||28)*invScale);
  const firmName=F.name||"Fin-Tax Mitra";
  const firmAddr=F.addr||"17, Sebadol Road, Belgharia, Kolkata, West Bengal - 700049";
  const firmPhone=F.phone||"7980718092";
  const firmEmail=F.email||"care.fintaxmitra@gmail.com";
  const bankName=F.bankName||"Kotak Mahindra Bank";
  const bankHolder=F.bankHolder||"BILTU DEY";
  const bankAcc=F.bankAcc||"1449547644";
  const bankIFSC=F.bankIFSC||"KKBK0000328";
  const upiId=F.upiId||"7319440039@kotak";
  const terms=F.terms||"Fees once paid are non-refundable. The client is responsible for ensuring data accuracy and timely submission. Fin-Tax Mitra is not liable for any penalties or losses resulting from client delays or incorrect data.";
  const totalAmt=inv.total||inv.amount||0;
  const items=(inv.items&&inv.items.length)?inv.items:[{service:inv.service,desc:"",qty:1,rate:inv.amount,amount:inv.amount}];
  const subTotal=inv.subTotal!==undefined?inv.subTotal:items.reduce((s,it)=>s+(it.amount!==undefined?it.amount:(it.qty||1)*(it.rate||0)),0);
  const discount=inv.discount||0;
  const fmt=n=>Number(n||0).toLocaleString("en-IN",{minimumFractionDigits:2});
  const invDate=inv.date?new Date(inv.date).toLocaleDateString("en-IN",{day:"2-digit",month:"2-digit",year:"numeric"}):"";

  const imgSty=(key,dw,dh)=>({
    width:F[key+"W"]||dw, height:F[key+"H"]||dh,
    objectFit:F[key+"Fit"]||"contain",
    filter:`brightness(${F[key+"Brightness"]??100}%) contrast(${F[key+"Contrast"]??100}%) saturate(${F[key+"Saturate"]??100}%)`,
    opacity:(F[key+"Opacity"]??100)/100, display:"block", flexShrink:0,
  });
  // Check if image has been repositioned in Performa Designer
  const hasPos=(key)=>F[key+"X"]!==undefined||F[key+"Y"]!==undefined;
  const absImgSty=(key,dw,dh,defX,defY)=>({
    position:"absolute",
    left:F[key+"X"]!==undefined?F[key+"X"]:defX,
    top:F[key+"Y"]!==undefined?F[key+"Y"]:defY,
    width:F[key+"W"]||dw, height:F[key+"H"]||dh,
    objectFit:F[key+"Fit"]||"contain",
    filter:`brightness(${F[key+"Brightness"]??100}%) contrast(${F[key+"Contrast"]??100}%) saturate(${F[key+"Saturate"]??100}%)`,
    opacity:(F[key+"Opacity"]??100)/100,
    zIndex:10+(F[key+"Z"]||1),
  });

  const ones=["","One","Two","Three","Four","Five","Six","Seven","Eight","Nine","Ten","Eleven","Twelve","Thirteen","Fourteen","Fifteen","Sixteen","Seventeen","Eighteen","Nineteen"];
  const tw=["","","Twenty","Thirty","Forty","Fifty","Sixty","Seventy","Eighty","Ninety"];
  const toW=n=>{
    if(!n)return"Zero"; n=Math.round(n);
    if(n>=10000000)return toW(Math.floor(n/10000000))+" Crore "+(n%10000000?toW(n%10000000):"");
    if(n>=100000)return toW(Math.floor(n/100000))+" Lakh "+(n%100000?toW(n%100000):"");
    if(n>=1000)return toW(Math.floor(n/1000))+" Thousand "+(n%1000?toW(n%1000):"");
    if(n>=100)return ones[Math.floor(n/100)]+" Hundred "+(n%100?toW(n%100):"");
    if(n>=20)return tw[Math.floor(n/10)]+(n%10?" "+ones[n%10]:"");
    return ones[n];
  };
  const totalWords=toW(Math.round(totalAmt)).trim()+" Rupees Only";

  // REDESIGN: #ftm-inv-area is now a fixed-height (not min-height) A4 flex
  // column. Header/title/customer/footer/terms never grow or shrink
  // (flexShrink:0); the service-table section is the single flexible region
  // (flex:1) and always has a spacer row with no fixed height, so the layout
  // naturally fills exactly one A4 page (794x1123px = 210x297mm @96dpi) on
  // screen AND on print — no scale-to-fit / scale-to-fill hacks required.
  const PAGE_W=794, PAGE_H=1123;
  const[pdfBusy,setPdfBusy]=useState(false);
  const[pdfMsg,setPdfMsg]=useState("");
  const downloadPDF=()=>{
    const el=document.getElementById("ftm-inv-area");
    if(!el||pdfBusy)return;
    setPdfBusy(true); setPdfMsg("");
    try{
      // No external PDF library available in this environment (html2canvas/
      // jsPDF are not supported) - use the browser's own print-to-PDF engine
      // instead. Opening a dedicated print window and calling window.print()
      // lets the user pick "Save as PDF" as the destination, which produces
      // an identical, dependency-free result.
      doPrint();
      setPdfMsg("✓ Choose \"Save as PDF\" in the print dialog");
      toast&&toast(`Invoice ${inv.id} sent to print/PDF`);
      setTimeout(()=>setPdfMsg(""),4500);
    }catch(e){
      setPdfMsg("✗ Download failed: "+e.message);
      setTimeout(()=>setPdfMsg(""),4000);
    }finally{
      setPdfBusy(false);
    }
  };
  const doPrint=()=>{
    const el=document.getElementById("ftm-inv-area");
    if(!el)return;
    const wrapper=el.parentElement||el;

    const printTitle = `${inv.id} ${inv.clientName}`;

    const htmlContent = `<!DOCTYPE html><html><head><meta charset="utf-8"><title>${printTitle}</title>
    <style>
      *{box-sizing:border-box;margin:0;padding:0;}
      html,body{width:${PAGE_W}px;background:#fff;color:#000;-webkit-print-color-adjust:exact;print-color-adjust:exact;}
      body{font-family:'Times New Roman',Times,serif;}
      table{width:100%;border-collapse:collapse;}
      td,th{word-break:break-word;}
      #ftm-print-page{width:${PAGE_W}px;height:${PAGE_H}px;overflow:hidden;position:relative;}
      @media print{@page{size:A4 portrait;margin:0;}body{margin:0;}}
    </style></head><body>
    <div id="ftm-print-page">${wrapper.outerHTML}</div>
    <script>
      window.onload = function() {
        document.title = "${printTitle}";
        setTimeout(function() {
          window.focus();
          window.print();
        }, 300);
      };
    </script>
    </body></html>`;

    const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const win = window.open(url, "_blank");

    // Clean up blob URL after print load triggers
    setTimeout(() => { URL.revokeObjectURL(url); }, 6000);
  };

  // shared cell styles
  const BDR=`1px solid ${clrB}`;
  const cs={fontFamily:"'Times New Roman',Times,serif",fontSize:bFsz,color:"#000"};

  return <div style={{position:"fixed",inset:0,background:"#000D",zIndex:6000,display:"flex",flexDirection:"column",overflow:"hidden"}}>
    {/* ── Toolbar ── */}
    <div style={{background:G.surf,borderBottom:`1px solid ${G.bdr}`,padding:"10px 20px",display:"flex",gap:10,alignItems:"center",justifyContent:"space-between",flexShrink:0}}>
      <div style={{display:"flex",alignItems:"center",gap:8,flexWrap:"wrap"}}>
        <span style={{fontWeight:700,fontSize:14,color:G.wh}}>🧾 {inv.id}</span>
        <span style={{fontSize:11,padding:"2px 9px",borderRadius:20,fontWeight:700,background:isGST?G.green+"20":G.amb+"20",color:isGST?G.green:G.amb,border:`1px solid ${isGST?G.green:G.amb}44`}}>{isGST?`With GST (${inv.gst}%)`:"Without GST"}</span>
        <span style={{fontSize:11,padding:"2px 9px",borderRadius:20,fontWeight:600,background:G.ind+"20",color:G.ind,border:`1px solid ${G.ind}44`}}>{invType}</span>
      </div>
      <div style={{display:"flex",gap:8,alignItems:"center"}}>
        <label style={{display:"flex",alignItems:"center",gap:6,color:G.wh,fontSize:12,cursor:"pointer",marginRight:12,userSelect:"none"}}>
          <input 
            type="checkbox" 
            checked={printStatus} 
            onChange={e=>setPrintStatus(e.target.checked)} 
            style={{cursor:"pointer",accentColor:G.green}}
          />
          Print Universal Stamp
        </label>
        {pdfMsg&&<span style={{fontSize:11,color:pdfMsg[0]==="✓"?G.green:G.red,fontWeight:600}}>{pdfMsg}</span>}
        <button onClick={downloadPDF} disabled={pdfBusy} style={{padding:"8px 20px",borderRadius:9,border:"none",cursor:pdfBusy?"default":"pointer",background:`linear-gradient(135deg,${G.g2},${G.green})`,color:"#fff",fontWeight:700,fontSize:13,opacity:pdfBusy?.7:1}}>{pdfBusy?"⏳ Opening...":"🖨 Print / Save as PDF"}</button>
        <button onClick={onClose} style={{padding:"8px 14px",borderRadius:9,border:`1px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer",fontWeight:600}}>✕ Close</button>
      </div>
    </div>

    {/* ── Paper area ── */}
    <div style={{flex:1,overflowY:"auto",overflowX:"auto",background:"#D1D5DB",padding:"12px",display:"flex",flexDirection:"column",alignItems:"center"}}>
      <div style={{width:PAGE_W,height:PAGE_H,position:"relative",flexShrink:0}}>
      {/* REDESIGN: fixed height (not min-height) + flex column. Every section
          below is flexShrink:0 (never squeezed) except the service-table
          section, which is the one flexible region (flex:1) that naturally
          absorbs/releases space so the invoice always fills exactly one A4
          page - on screen and identically on print, with no scaling tricks. */}
      <div id="ftm-inv-area" style={{width:"100%",height:PAGE_H,background:"#fff",fontFamily:"'Times New Roman',Times,serif",color:"#000",boxShadow:"0 4px 32px #0004",border:"1px solid #ccc",display:"flex",flexDirection:"column",overflow:"hidden"}}>

        {/* ── 1. FIRM HEADER ── */}
        <div style={{padding:"16px 20px",borderBottom:"4px solid transparent",borderImage:`linear-gradient(90deg,${clrP},${clrB}) 1`,display:"flex",gap:14,alignItems:"center",flexShrink:0}}>
          {/* Logo placeholder - actual logo is absolute overlay */}
          <div style={{width:F.logoW||72,height:F.logoH||72,flexShrink:0,
            border:`2px solid ${clrP}`,borderRadius:"50%",
            display:"flex",alignItems:"center",justifyContent:"center",
            background:clrA,visibility:F.logo?"hidden":"visible"}}>
            <div style={{textAlign:"center",lineHeight:1.2}}>
              <div style={{fontSize:11,fontWeight:900,color:clrP}}>Fin</div>
              <div style={{fontSize:10,fontWeight:900,color:clrP}}>Tax</div>
              <div style={{fontSize:9,fontWeight:700,color:clrP}}>mitra</div>
            </div>
          </div>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontSize:tFsz,fontWeight:900,fontStyle:"italic",color:clrP,letterSpacing:-.5,fontFamily:"Georgia,serif",lineHeight:1.15}}>{firmName}</div>
            <div style={{fontSize:bFsz-1,color:"#444",marginTop:5,lineHeight:1.8}}>{firmAddr}</div>
            <div style={{fontSize:bFsz-1,color:"#444",lineHeight:1.8}}>Ph: {firmPhone}  |  Email: {firmEmail}</div>
          </div>
        </div>

        {/* ── 2. INVOICE TITLE ── */}
        <div style={{background:clrA,padding:"6px 20px",display:"flex",justifyContent:"space-between",alignItems:"center",borderBottom:BDR}}>
          <div style={{fontSize:tFsz-2,fontWeight:900,fontStyle:"italic",letterSpacing:3,color:clrP,fontFamily:"Georgia,serif"}}>{invType}</div>
          <div style={{fontSize:10,color:clrP,fontStyle:"italic",fontWeight:700,border:`1px solid ${clrP}`,borderRadius:20,padding:"3px 12px",letterSpacing:.5,flexShrink:0,marginLeft:10,background:"#fff"}}>{copyLabel}</div>
        </div>

        {/* ── 3. CUSTOMER + INVOICE DETAILS - rounded soft boxes ── */}
        <div style={{display:"flex",gap:12,padding:"12px 16px",borderBottom:BDR,flexShrink:0}}>
          <div style={{flex:1,border:BDR,borderRadius:10,padding:"12px 16px",background:"#fff"}}>
            <div style={{fontWeight:900,fontSize:bFsz,fontStyle:"italic",color:clrP,textAlign:"center",borderBottom:BDR,paddingBottom:6,marginBottom:10}}>Customer Details</div>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <tbody>
                {[
                  ["Name",inv.clientName,true],
                  ["Address",cl?.addr||"-",false],
                  ["Phone No.",cl?.mob||"-",false],
                  ["PAN",inv.pan,true],
                ].map(([l,v,bold])=>(
                  <tr key={l}><td style={{...cs,fontWeight:700,width:84,paddingBottom:7,verticalAlign:"top",whiteSpace:"nowrap"}}>{l}</td><td style={{...cs,paddingBottom:7,fontWeight:bold?700:400}}>: {v}</td></tr>
                ))}
                {cl?.gstin&&<tr><td style={{...cs,fontWeight:700,paddingBottom:7}}>GSTIN</td><td style={{...cs,paddingBottom:7,fontFamily:"monospace",fontSize:bFsz-1}}>: {cl.gstin}</td></tr>}
              </tbody>
            </table>
          </div>
          <div style={{flex:1,border:BDR,borderRadius:10,padding:"12px 16px",background:"#fff"}}>
            <div style={{fontWeight:900,fontSize:bFsz,fontStyle:"italic",color:clrP,textAlign:"center",borderBottom:BDR,paddingBottom:6,marginBottom:10}}>Invoice Details</div>
            <table style={{width:"100%",borderCollapse:"collapse"}}>
              <tbody>
                <tr><td style={{...cs,fontWeight:700,width:104,paddingBottom:10}}>Invoice No.</td><td style={{...cs,fontWeight:700,fontStyle:"italic",paddingBottom:10}}>: {inv.id}</td></tr>
                <tr><td style={{...cs,fontWeight:700,paddingBottom:10}}>Invoice Date</td><td style={{...cs,fontWeight:700,paddingBottom:10}}>: {invDate}</td></tr>
                <tr><td style={{...cs,fontWeight:700}}>Financial Year</td><td style={{...cs}}>: {inv.fy}</td></tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* ── 4. SERVICE TABLE ── */}
        {/* REDESIGN: this is the ONE flexible region of the page. It grows or
            shrinks (flex:1, minHeight:0) to exactly fill whatever space is
            left between the fixed header/title/customer blocks above and the
            fixed footer/terms blocks below - so the invoice always occupies
            the full A4 page, never leaving a blank gap or overflowing,
            regardless of the Empty Rows setting or GST rows. */}
        <div style={{flex:"1 1 auto",minHeight:0,position:"relative",borderBottom:BDR}}>
        <div style={{position:"absolute",inset:0,overflow:"hidden"}}>
        {/* Faint diagonal firm-name watermark behind the table */}
        <div style={{position:"absolute",inset:0,display:"flex",alignItems:"center",justifyContent:"center",pointerEvents:"none",zIndex:0}}>
          <div style={{transform:"rotate(-28deg)",fontSize:52,fontWeight:900,fontStyle:"italic",fontFamily:"Georgia,serif",color:clrP,opacity:.06,whiteSpace:"nowrap",letterSpacing:2}}>{firmName}</div>
        </div>
        <table style={{width:"100%",height:"100%",borderCollapse:"collapse",position:"relative",zIndex:1}}>
          <colgroup>
            <col style={{width:"6%"}}/>
            <col style={{width:"54%"}}/>
            <col style={{width:"12%"}}/>
            <col style={{width:"14%"}}/>
            <col style={{width:"14%"}}/>
          </colgroup>
          <thead>
            <tr style={{background:clrA}}>
              {[["Sr. No.","center"],["Name of Service","center"],["Period / Nos.","center"],["Rate (Rs.)","right"],["Amount (Rs.)","right"]].map(([h,align],i)=>(
                <th key={i} style={{...cs,fontWeight:900,fontStyle:"italic",color:clrP,padding:"9px 10px",textAlign:align,borderBottom:BDR,borderRight:i<4?BDR:"none",whiteSpace:"pre-line",lineHeight:1.3}}>{h}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {items.map((it,i)=>(
            <tr key={i}>
              <td style={{...cs,padding:"10px",textAlign:"center",fontStyle:"italic",borderRight:BDR,verticalAlign:"top"}}>{i+1}.</td>
              <td style={{...cs,padding:"10px",fontStyle:"italic",borderRight:BDR,verticalAlign:"top"}}>
                {it.type==="reimbursement"
                  ?<span>{it.desc||"Reimbursement"} <em style={{fontSize:cs.fontSize-1,color:"#666"}}>(Reimbursement)</em></span>
                  :<span>{it.service}{it.desc?` ( ${it.desc} )`:""}</span>}
              </td>
              <td style={{...cs,padding:"10px",textAlign:"center",borderRight:BDR,verticalAlign:"top"}}>{it.type==="reimbursement"?"1":it.qty||1}</td>
              <td style={{...cs,padding:"10px",textAlign:"right",borderRight:BDR,verticalAlign:"top",fontVariantNumeric:"tabular-nums"}}>{it.type==="reimbursement"?"-":fmt(it.rate)}</td>
              <td style={{...cs,padding:"10px",textAlign:"right",verticalAlign:"top",fontVariantNumeric:"tabular-nums"}}>{fmt(it.type==="reimbursement"?(Number(it.rate)||0):(it.amount!==undefined?it.amount:(it.qty||1)*(it.rate||0)))}</td>
            </tr>
            ))}
            {Array.from({length:emptyRows}).map((_,i)=>(
              <tr key={i}>
                <td style={{height:28,borderRight:BDR}}/>
                <td style={{borderRight:BDR}}/>
                <td style={{borderRight:BDR}}/>
                <td style={{borderRight:BDR}}/>
                <td/>
              </tr>
            ))}
            {/* Guaranteed spacer row: always present (unlike the user-configurable
                empty rows above, which may be set to 0) so this row alone ever
                absorbs left-over page space. CGST/SGST/Total rows below all have
                an explicit fixed height, so they can never be stretched. */}
            <tr><td style={{borderRight:BDR}}/><td style={{borderRight:BDR}}/><td style={{borderRight:BDR}}/><td style={{borderRight:BDR}}/><td/></tr>
            {discount>0&&<>
              <tr style={{height:Math.round(30*invScale)}}>
                <td colSpan={3} style={{borderRight:BDR,borderTop:BDR}}/>
                <td style={{...cs,padding:"6px 10px",textAlign:"right",fontStyle:"italic",fontSize:bFsz-1,borderRight:BDR,borderTop:BDR}}>Sub Total</td>
                <td style={{...cs,padding:"6px 10px",textAlign:"right",fontSize:bFsz-1,borderTop:BDR,fontVariantNumeric:"tabular-nums"}}>{fmt(subTotal)}</td>
              </tr>
              <tr style={{height:Math.round(30*invScale)}}>
                <td colSpan={3} style={{borderRight:BDR}}/>
                <td style={{...cs,padding:"6px 10px",textAlign:"right",fontStyle:"italic",fontSize:bFsz-1,borderRight:BDR}}>Discount</td>
                <td style={{...cs,padding:"6px 10px",textAlign:"right",fontSize:bFsz-1,fontVariantNumeric:"tabular-nums"}}>(-) {fmt(discount)}</td>
              </tr>
            </>}
            {isGST&&<>
              <tr style={{height:Math.round(30*invScale)}}>
                <td colSpan={3} style={{borderRight:BDR,borderTop:BDR}}/>
                <td style={{...cs,padding:"6px 10px",textAlign:"right",fontStyle:"italic",fontSize:bFsz-1,borderRight:BDR,borderTop:BDR}}>CGST ({inv.gst/2}%)</td>
                <td style={{...cs,padding:"6px 10px",textAlign:"right",fontSize:bFsz-1,borderTop:BDR,fontVariantNumeric:"tabular-nums"}}>{fmt((inv.gstAmt||0)/2)}</td>
              </tr>
              <tr style={{height:Math.round(30*invScale)}}>
                <td colSpan={3} style={{borderRight:BDR}}/>
                <td style={{...cs,padding:"6px 10px",textAlign:"right",fontStyle:"italic",fontSize:bFsz-1,borderRight:BDR}}>SGST ({inv.gst/2}%)</td>
                <td style={{...cs,padding:"6px 10px",textAlign:"right",fontSize:bFsz-1,fontVariantNumeric:"tabular-nums"}}>{fmt((inv.gstAmt||0)/2)}</td>
              </tr>
            </>}
            <tr style={{background:clrT,height:Math.round(38*invScale)}}>
              <td colSpan={4} style={{...cs,padding:"10px",fontWeight:900,fontStyle:"italic",fontSize:bFsz+1,textAlign:"center",borderTop:BDR,borderRight:BDR}}>Total</td>
              <td style={{...cs,padding:"10px",fontWeight:900,fontSize:bFsz+1,textAlign:"right",borderTop:BDR,fontVariantNumeric:"tabular-nums"}}>{fmt(totalAmt)}</td>
            </tr>
          </tbody>
        </table>
        </div>
        </div>

        {/* ── 5. FOOTER: words + bank | cert + sig ── */}
        <table style={{width:"100%",borderCollapse:"collapse",borderBottom:BDR,flexShrink:0}}>
          <tbody>
            <tr>
              <td style={{width:"55%",verticalAlign:"top",borderRight:BDR}}>
                {/* Total in words row */}
                <table style={{width:"100%",borderCollapse:"collapse",borderBottom:BDR}}>
                  <tbody>
                    <tr>
                      <td style={{padding:"13px 18px",verticalAlign:"top",width:"50%"}}>
                        <div style={{...cs,fontWeight:900,fontStyle:"italic",marginBottom:6}}>Total in words</div>
                        <div style={{...cs,fontStyle:"italic",fontSize:bFsz-1,lineHeight:1.7,color:"#333"}}>{totalWords}</div>
                      </td>
                      <td style={{padding:"13px 18px",verticalAlign:"top",textAlign:"right"}}>
                        <div style={{...cs,fontWeight:900,fontStyle:"italic",marginBottom:6}}>Total Amount</div>
                        <div style={{display:"inline-block",background:clrA,color:clrP,border:`1.5px solid ${clrB}`,fontWeight:900,fontSize:bFsz+3,fontFamily:"'Times New Roman',serif",padding:"5px 16px",borderRadius:8,fontVariantNumeric:"tabular-nums"}}>₹ {fmt(totalAmt)}</div>
                        <div style={{fontSize:bFsz-2,color:"#777",fontStyle:"italic",marginTop:4}}>(E & O.E.)</div>
                      </td>
                    </tr>
                  </tbody>
                </table>
                {/* Bank details */}
                <div style={{padding:"13px 18px"}}>
                  <div style={{...cs,fontWeight:900,fontStyle:"italic",color:clrP,marginBottom:8,borderBottom:`1px solid ${clrB}66`,paddingBottom:4}}>Bank Details</div>
                  <div style={{display:"flex",gap:14,alignItems:"flex-start"}}>
                    <div style={{flex:1,minWidth:0}}>
                      <table style={{width:"100%",borderCollapse:"collapse"}}>
                        <tbody>
                          {[["Bank Name",bankName],["A/C Holder Name",bankHolder],["A/C Number",bankAcc],["IFSC",bankIFSC]].map(([l,v])=>(
                            <tr key={l}>
                              <td style={{...cs,fontWeight:700,paddingRight:6,paddingBottom:5,whiteSpace:"nowrap",verticalAlign:"top",width:110}}>{l} :</td>
                              <td style={{...cs,paddingBottom:5}}>{v}</td>
                            </tr>
                          ))}
                          <tr>
                            <td style={{...cs,fontWeight:700,paddingTop:3,whiteSpace:"nowrap",verticalAlign:"top"}}>UPI ID :</td>
                            <td style={{...cs,paddingTop:3}}>{upiId} </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                    {/* Fixed QR box - uploaded QR always renders here */}
                    <div style={{flexShrink:0,textAlign:"center"}}>
                      <div style={{width:88,height:88,border:`1.5px solid ${clrB}`,borderRadius:8,display:"flex",alignItems:"center",justifyContent:"center",background:"#fff",overflow:"hidden"}}>
                        {F.qrCode
                          ?<img src={F.qrCode} alt="UPI QR" style={{width:80,height:80,objectFit:"contain",display:"block"}}/>
                          :<span style={{fontSize:9,color:"#999",fontStyle:"italic",textAlign:"center",padding:4}}>QR Code</span>}
                      </div>
                      <div style={{...cs,fontSize:Math.max(bFsz-5,9),fontWeight:700,fontStyle:"italic",marginTop:3,whiteSpace:"nowrap"}}>Scan to pay using UPI</div>
                    </div>
                  </div>
                </div>
              </td>
              <td style={{width:"45%",verticalAlign:"top",padding:"12px 14px"}}>
                <div style={{...cs,fontStyle:"italic",lineHeight:1.6,color:"#444"}}>Certified that the particulars given above are true and correct.</div>
                <div style={{...cs,fontWeight:900,fontStyle:"italic",fontSize:bFsz+1,color:clrP,marginTop:8}}>For {firmName}</div>
                <div style={{textAlign:"center",marginTop:16}}>
                  <div style={{minHeight:80}}/>
                  {/* Stamp and signature rendered as absolute overlays */}
                  {!F.signature&&<div style={{width:140,borderBottom:"1px solid #888",margin:"36px auto 0"}}/>}
                  <div style={{...cs,fontWeight:900,fontStyle:"italic",marginTop:10}}>Authorised Signatory</div>
                </div>
              </td>
            </tr>
          </tbody>
        </table>

        {/* ── 6. TERMS ── */}
        <div style={{padding:"10px 16px 24px",background:clrA,flexShrink:0}}>
          <div style={{...cs,fontWeight:900,fontStyle:"italic",color:clrP,marginBottom:4}}>Terms and Conditions</div>
          <div style={{...cs,fontSize:bFsz-2,color:"#444",fontStyle:"italic",lineHeight:1.6}}>{terms}</div>
        </div>

      </div>
      {/* ── Image overlays - OUTSIDE paper, INSIDE position:relative wrapper ── */}
      {/* Same coordinate system as PerformaDesigner */}
      {[
        {key:"logo",        dw:72,  dh:72,  defX:16,  defY:16},
        {key:"stamp",       dw:80,  dh:80,  defX:490, defY:870},
        {key:"signature",   dw:140, dh:44,  defX:450, defY:930},
        {key:"statusStamp", dw:160, dh:160, defX:280, defY:380},
      ].filter(img=>!!F[img.key] && (img.key !== "statusStamp" || printStatus)).map(img=>{
        const x=F[img.key+"X"]!==undefined?F[img.key+"X"]:img.defX;
        const y=F[img.key+"Y"]!==undefined?F[img.key+"Y"]:img.defY;
        const w=F[img.key+"W"]||img.dw;
        const h=F[img.key+"H"]||img.dh;
        const z=F[img.key+"Z"]||1;
        const br=F[img.key+"Brightness"]!==undefined?F[img.key+"Brightness"]:100;
        const co=F[img.key+"Contrast"]!==undefined?F[img.key+"Contrast"]:100;
        const sa=F[img.key+"Saturate"]!==undefined?F[img.key+"Saturate"]:100;
        const op=(F[img.key+"Opacity"]!==undefined?F[img.key+"Opacity"]:100)/100;
        return <img key={img.key} src={F[img.key]} alt={img.key}
          style={{position:"absolute",left:x,top:y,width:w,height:h,
            objectFit:F[img.key+"Fit"]||"contain",
            filter:`brightness(${br}%) contrast(${co}%) saturate(${sa}%)`,
            opacity:op, zIndex:10+z, pointerEvents:"none"}}/>;
      })}
      </div>
    </div>
  </div>;
}




// ─── Payment Receipts Module ───────────────────────────────────────────────────
function ReceiptsModule({receipts,setReceipts,invoices,setInvoices,clients,dd,toast}){
  const[showForm,setShowForm]=useState(false);
  const[q,setQ]=useState("");
  const MODES=["Cash","UPI","Bank Transfer","Cheque","NEFT","RTGS","Other"];

  const filtered=useMemo(()=>{
    if(!q)return [...receipts].sort((a,b)=>new Date(b.date)-new Date(a.date));
    const lq=q.toLowerCase();
    return receipts.filter(r=>r.clientName.toLowerCase().includes(lq)||r.invId.toLowerCase().includes(lq)||r.id.toLowerCase().includes(lq)).sort((a,b)=>new Date(b.date)-new Date(a.date));
  },[receipts,q]);

  const totalReceived=filtered.reduce((s,r)=>s+(r.received||0),0);

  const genRcpId=()=>{
    if(!receipts || !receipts.length) return "RCP-001";
    const rcpIds = receipts
      .filter(r => r.id && /^RCP-\d+$/.test(r.id))
      .map(r => parseInt(r.id.split("-")[1], 10));
    if(!rcpIds.length) return "RCP-001";
    const maxId = Math.max(...rcpIds);
    return `RCP-${String(maxId + 1).padStart(3, "0")}`;
  };

  const deleteRcp=id=>{
    const rcp=receipts.find(r=>r.id===id);
    setReceipts(p=>{
      const remaining=p.filter(r=>r.id!==id);
      // Recompute the invoice's status from what's left — otherwise deleting a receipt
      // (e.g. entered by mistake) would leave the invoice permanently stuck as Paid/Partial.
      if(rcp) setInvoices(pi=>pi.map(inv=>inv.id===rcp.invId?{...inv,status:computeInvStatus(inv,remaining)}:inv));
      return remaining;
    });
    toast("Receipt deleted");
  };

  return <div style={{display:"flex",flexDirection:"column",gap:16,animation:"fadeUp .3s ease"}}>
    {showForm&&<ReceiptForm receipts={receipts} setReceipts={setReceipts} invoices={invoices} setInvoices={setInvoices} clients={clients} dd={dd} toast={toast} onClose={()=>setShowForm(false)} genId={genRcpId} MODES={MODES}/>}
    {/* KPIs */}
    <div className="kpi-grid-3" style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:12}}>
      {[
        {l:"Total Receipts",v:filtered.length,icon:"🧾",col:G.green},
        {l:"Total Received",v:inr(totalReceived),icon:"💰",col:G.g3},
        {l:"This Month",v:inr(receipts.filter(r=>new Date(r.date).getMonth()===new Date().getMonth()&&new Date(r.date).getFullYear()===new Date().getFullYear()).reduce((s,r)=>s+(r.received||0),0)),icon:"📅",col:G.cyn},
      ].map(k=><div key={k.l} style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"13px 16px",display:"flex",gap:10,alignItems:"center"}}>
        <div style={{width:38,height:38,borderRadius:10,background:k.col+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:18,flexShrink:0}}>{k.icon}</div>
        <div><div style={{fontWeight:800,fontSize:16,color:k.col}}>{k.v}</div><div style={{fontSize:10,color:G.mut,marginTop:2}}>{k.l}</div></div>
      </div>)}
    </div>
    {/* Search + New */}
    <div style={{display:"flex",gap:10,alignItems:"center"}}>
      <div style={{position:"relative",flex:1}}>
        <span style={{position:"absolute",left:10,top:"50%",transform:"translateY(-50%)",color:G.mut}}>🔍</span>
        <input value={q} onChange={e=>setQ(e.target.value)} placeholder="Search client, receipt or invoice no..." style={{...IS,paddingLeft:30,fontSize:12}}/>
      </div>
      <button onClick={()=>setShowForm(true)} style={{padding:"8px 18px",borderRadius:10,border:"none",cursor:"pointer",background:`linear-gradient(135deg,${G.g2},${G.green})`,color:"#fff",fontWeight:700,fontSize:13,whiteSpace:"nowrap"}}>➕ Record Payment</button>
    </div>
    {/* Receipts table */}
    <div style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:14,overflow:"hidden"}}>
      <div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
        <thead><tr style={{background:G.bg}}>
          {["Receipt No.","Date","Client","Invoice No.","Invoice Amt","Received","Balance","Mode","Ref No.",""].map(h=>(
            <th key={h} style={{padding:"10px 12px",textAlign:"left",color:G.mut,fontWeight:600,fontSize:11,whiteSpace:"nowrap",borderBottom:`1px solid ${G.bdr}`}}>{h}</th>
          ))}
        </tr></thead>
        <tbody>
          {filtered.length===0?<tr><td colSpan={10} style={{padding:32,textAlign:"center",color:G.bdr}}>No receipts yet</td></tr>
          :filtered.map((r,i)=>{
            const bal=(r.invAmt||0)-(r.received||0);
            return <tr key={r.id} style={{borderTop:`1px solid ${G.bdr}`,background:i%2?G.card+"80":"transparent"}}>
              <td style={{padding:"10px 12px",fontWeight:700,color:G.cyn,fontFamily:"monospace",fontSize:11}}>{r.id}</td>
              <td style={{padding:"10px 12px",color:G.mut,whiteSpace:"nowrap"}}>{r.date?new Date(r.date).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"}):"-"}</td>
              <td style={{padding:"10px 12px"}}>
                <div style={{fontWeight:600,color:G.txt,whiteSpace:"nowrap"}}>{r.clientName}</div>
                <div style={{fontSize:10,color:G.mut,fontFamily:"monospace"}}>{r.pan}</div>
              </td>
              <td style={{padding:"10px 12px",color:G.ind,fontFamily:"monospace",fontSize:11}}>{r.invId}</td>
              <td style={{padding:"10px 12px",color:G.txt}}>{inr(r.invAmt||0)}</td>
              <td style={{padding:"10px 12px",color:G.g3,fontWeight:700}}>{inr(r.received||0)}</td>
              <td style={{padding:"10px 12px",color:bal>0?G.red:G.g3,fontWeight:700}}>{inr(bal)}</td>
              <td style={{padding:"10px 12px"}}><span style={{background:G.card,border:`1px solid ${G.bdr}`,color:G.txt,padding:"2px 8px",borderRadius:8,fontSize:11,whiteSpace:"nowrap"}}>{r.mode}</span></td>
              <td style={{padding:"10px 12px",color:G.mut,fontFamily:"monospace",fontSize:11}}>{r.ref||"-"}</td>
              <td style={{padding:"10px 12px"}}>
                <button onClick={()=>deleteRcp(r.id)} style={{background:"#450A0A",border:`1px solid ${G.red}44`,color:G.red,borderRadius:6,padding:"3px 7px",cursor:"pointer",fontSize:11,fontWeight:700}}>🗑</button>
              </td>
            </tr>;
          })}
        </tbody>
        {filtered.length>0&&<tfoot>
          <tr style={{background:G.bg,borderTop:`2px solid ${G.green}33`}}>
            <td colSpan={4} style={{padding:"9px 12px",fontWeight:800,color:G.wh,fontSize:12}}>TOTALS</td>
            <td style={{padding:"9px 12px",color:G.txt,fontWeight:700}}>{inr(filtered.reduce((s,r)=>s+(r.invAmt||0),0))}</td>
            <td style={{padding:"9px 12px",color:G.g3,fontWeight:800}}>{inr(totalReceived)}</td>
            <td style={{padding:"9px 12px",color:G.red,fontWeight:800}}>{inr(filtered.reduce((s,r)=>s+((r.invAmt||0)-(r.received||0)),0))}</td>
            <td colSpan={3}/>
          </tr>
        </tfoot>}
      </table></div>
    </div>
  </div>;
}

// ─── Receipt Form ─────────────────────────────────────────────────────────────
function ReceiptForm({receipts,setReceipts,invoices,setInvoices,clients,dd,toast,onClose,genId,MODES}){
  const[f,setF]=useState({date:td(),invId:"",pan:"",clientName:"",invAmt:"",received:"",mode:"UPI",ref:"",note:""});
  const[err,setErr]=useState({});

  const unpaidInvoices=invoices.filter(i=>i.status!=="Paid");
  const selInv=invoices.find(i=>i.id===f.invId);

  const save=()=>{
    const e={};
    if(!f.invId)e.invId="Select invoice";
    if(!f.received||isNaN(Number(f.received)))e.received="Enter amount";
    if(!f.mode)e.mode="Select mode";
    setErr(e);
    if(Object.keys(e).length){toast("Fix errors","err");return;}
    const rcp={id:genId(),date:f.date,invId:f.invId,pan:f.pan,clientName:f.clientName,invAmt:Number(f.invAmt)||0,received:Number(f.received)||0,mode:f.mode,ref:f.ref,note:f.note};
    setReceipts(p=>[rcp,...p]);
    // Recompute invoice status from ALL receipts against it (existing + this new one) —
    // single shared rule, so it can never drift from what Receipts actually shows.
    if(selInv){
      const allReceiptsForInv=[rcp,...receipts.filter(r=>r.invId===f.invId)];
      const newStatus=computeInvStatus(selInv,allReceiptsForInv);
      setInvoices(p=>p.map(i=>i.id===f.invId?{...i,status:newStatus}:i));
    }
    toast("Payment recorded - "+genId());
    onClose();
  };

  return <div style={{position:"fixed",inset:0,background:"#000C",zIndex:5500,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
    <div style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:18,width:"min(560px,98%)",maxHeight:"90vh",overflow:"auto",boxShadow:"0 32px 80px #000C"}}>
      <div style={{padding:"16px 20px",borderBottom:`1px solid ${G.bdr}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
        <span style={{fontWeight:800,fontSize:15,color:G.wh}}>💳 Record Payment Receipt</span>
        <button onClick={onClose} style={{background:"transparent",border:`1px solid ${G.bdr}`,color:G.mut,borderRadius:8,padding:"4px 11px",cursor:"pointer"}}>✕</button>
      </div>
      <div style={{padding:20,display:"flex",flexDirection:"column",gap:12}}>
        <F label="Select Invoice (Unpaid / Partial)" req>
          <select value={f.invId} onChange={e=>{
            const inv=invoices.find(i=>i.id===e.target.value);
            setF(p=>({...p,invId:e.target.value,pan:inv?.pan||"",clientName:inv?.clientName||"",invAmt:String(inv?.total||"")}));
          }} style={{...IS,cursor:"pointer",color:f.invId?G.wh:G.mut}}>
            <option value="">- Select outstanding invoice -</option>
            {unpaidInvoices.map(i=><option key={i.id} value={i.id} style={{background:"#0B1610"}}>{i.id} | {i.clientName} | {inr(i.total)} | {i.status}</option>)}
          </select>
          {err.invId&&<span style={{fontSize:11,color:G.red}}>{err.invId}</span>}
        </F>
        {selInv&&<div style={{padding:"10px 14px",background:G.green+"08",border:`1px solid ${G.green}22`,borderRadius:10,fontSize:12}}>
          <div style={{display:"flex",gap:16,flexWrap:"wrap"}}>
            <span style={{color:G.mut}}>Invoice: <strong style={{color:G.txt}}>{selInv.id}</strong></span>
            <span style={{color:G.mut}}>Total: <strong style={{color:G.green}}>{inr(selInv.total)}</strong></span>
            <span style={{color:G.mut}}>Status: <strong style={{color:selInv.status==="Paid"?G.g3:G.red}}>{selInv.status}</strong></span>
          </div>
          <div style={{marginTop:4,color:G.mut}}>Client: <strong style={{color:G.txt}}>{selInv.clientName}</strong> · {selInv.service}</div>
        </div>}
        <div style={{display:"flex",gap:12}}>
          <F label="Payment Date" req w="calc(50% - 6px)">
            <I val={f.date} set={v=>setF({...f,date:v})} type="date"/>
          </F>
          <F label="Amount Received (₹)" req w="calc(50% - 6px)">
            <I val={f.received} set={v=>setF({...f,received:v.replace(/[^0-9.]/g,"")})} ph="0" sty={{borderColor:err.received?G.red:G.bdr}}/>
            {err.received&&<span style={{fontSize:11,color:G.red}}>{err.received}</span>}
          </F>
        </div>
        <div style={{display:"flex",gap:12}}>
          <F label="Payment Mode" req w="calc(50% - 6px)">
            <S val={f.mode} set={v=>setF({...f,mode:v})} opts={MODES}/>
          </F>
          <F label="Reference / Transaction No." w="calc(50% - 6px)">
            <I val={f.ref} set={v=>setF({...f,ref:v})} ph="UPI ref, cheque no..."/>
          </F>
        </div>
        <F label="Notes">
          <textarea value={f.note} onChange={e=>setF({...f,note:e.target.value})} rows={2} style={{...IS,resize:"vertical"}}/>
        </F>
      </div>
      <div style={{padding:"14px 20px",borderTop:`1px solid ${G.bdr}`,display:"flex",gap:10}}>
        <button onClick={save} style={{flex:1,padding:"12px",borderRadius:10,border:"none",cursor:"pointer",background:`linear-gradient(135deg,${G.g2},${G.green})`,color:"#fff",fontWeight:700,fontSize:14}}>💳 Record Payment</button>
        <button onClick={onClose} style={{padding:"12px 18px",borderRadius:10,border:`1px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer"}}>Cancel</button>
      </div>
    </div>
  </div>;
}

function PortalPw({pw,pid}){
  const[show,setShow]=useState(false);
  return <div style={{display:"flex",alignItems:"center",gap:5,flex:1}}>
    <span style={{fontSize:12,color:"#F0FDF4",fontFamily:"monospace",fontWeight:600,flex:1}}>{show?pw||"-":"********"}</span>
    <button onClick={()=>setShow(s=>!s)} style={{background:"none",border:"none",cursor:"pointer",color:"#86EFAC",fontSize:13,padding:"0 2px",flexShrink:0}}>{show?"🙈":"👁"}</button>
  </div>;
}

function WorkTracker({works,setWorks,clients,ownerOn,dd,pws,toast,invoices,setInvoices,receipts}){
  const[flt,setFlt]=useState(() => localStorage.getItem("fmt_tracker_flt") || "All");
  const[fy,setFy]=useState(() => localStorage.getItem("fmt_tracker_fy") || getCurrentFY());
  const[q,setQ]=useState(() => localStorage.getItem("fmt_tracker_q") || "");

  useEffect(() => {
    localStorage.setItem("fmt_tracker_flt", flt);
  }, [flt]);

  useEffect(() => {
    localStorage.setItem("fmt_tracker_fy", fy);
  }, [fy]);

  useEffect(() => {
    localStorage.setItem("fmt_tracker_q", q);
  }, [q]);

  const clearFilters = () => {
    setFlt("All");
    setFy(getCurrentFY());
    setQ("");
    localStorage.setItem("fmt_tracker_flt", "All");
    localStorage.setItem("fmt_tracker_fy", getCurrentFY());
    localStorage.setItem("fmt_tracker_q", "");
  };
  const[showAuth,setShowAuth]=useState(false),[feesOn,setFeesOn]=useState(ownerOn);
  const[editW,setEditW]=useState(null),[opP,setOpP]=useState({}),[vpw,setVpw]=useState({});
  const[selClient,setSelClient]=useState(null);
  const[delW,setDelW]=useState(null);
  const[linkPop,setLinkPop]=useState(null),[linkSel,setLinkSel]=useState("");
  const doDeleteWork=()=>{setWorks(p=>p.filter(x=>x.id!==delW.id));setDelW(null);};
  // Attaches an existing invoice to this work — from this point on, "Received" for
  // the work is the live total of actual receipts recorded against that invoice, so
  // multiple works for the same client each track their own payments correctly, no
  // matter how many installments come in or when.
  const linkInvoice=(workId,invId)=>{
    if(!invId)return;
    setInvoices(p=>p.map(inv=>inv.id===invId?{...inv,workId}:inv));
    setLinkSel("");
    toast&&toast("Invoice "+invId+" linked to this work");
  };
  const unlinkInvoice=invId=>{
    setInvoices(p=>p.map(inv=>inv.id===invId?{...inv,workId:""}:inv));
    toast&&toast("Invoice unlinked");
  };
  const fw=useMemo(()=>{let w=works;if(fy!=="All")w=w.filter(x=>x.fy===fy);if(flt==="Overdue")w=w.filter(isOD);else if(flt!=="All")w=w.filter(x=>x.status===flt);if(q){const lq=q.toLowerCase();w=w.filter(x=>x.cn.toLowerCase().includes(lq)||x.pan.toLowerCase().includes(lq)||x.svc.toLowerCase().includes(lq));}return w;},[works,flt,fy,q]);
  const cnt={All:works.length,Pending:works.filter(w=>w.status==="Pending").length,"In Progress":works.filter(w=>w.status==="In Progress").length,Completed:works.filter(w=>w.status==="Completed").length,Overdue:works.filter(isOD).length};
  return <div style={{display:"flex",flexDirection:"column",gap:13}}>
    {showAuth&&<Auth title="Unlock Fee Data" hint="Finance: 456 | Developer: 123" pws={pws} onOk={()=>{setFeesOn(true);setShowAuth(false);}} onX={()=>setShowAuth(false)}/>}
    {editW&&<EditWork w={editW} rcvdDisplay={workReceived(editW,invoices,receipts)} linkedCount={linkedInvoices(editW,invoices).length} onSave={wf=>{setWorks(p=>p.map(w=>w.id===wf.id?wf:w));setEditW(null);}} onX={()=>setEditW(null)} dd={dd}/>}
    {/* Client Detail Slide Panel */}
    {selClient&&<div style={{position:"fixed",inset:0,zIndex:4000,display:"flex"}} onClick={()=>setSelClient(null)}>
      <div style={{flex:1,background:"#000A"}}/>
      <div style={{width:"min(480px,95vw)",background:G.surf,borderLeft:`1px solid ${G.green}44`,height:"100%",overflow:"auto",boxShadow:"-20px 0 60px #000C",animation:"slidePanel .25s ease"}} onClick={e=>e.stopPropagation()}>
        {/* Panel Header */}
        <div style={{background:`linear-gradient(135deg,${G.gd},${G.g2})`,padding:"18px 20px",display:"flex",gap:14,alignItems:"center",position:"sticky",top:0,zIndex:1}}>
          <div style={{width:50,height:50,borderRadius:13,background:`linear-gradient(135deg,${G.g2},${G.green})`,display:"flex",alignItems:"center",justifyContent:"center",fontSize:20,fontWeight:800,color:"#fff",flexShrink:0,border:"2px solid #4ADE8055"}}>{selClient.name[0]}</div>
          <div style={{flex:1,minWidth:0}}>
            <div style={{fontWeight:800,fontSize:16,color:G.wh,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{selClient.name}</div>
            {selClient.biz&&<div style={{fontSize:12,color:G.g3,marginTop:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{selClient.biz}</div>}
            <div style={{display:"flex",gap:6,marginTop:5,flexWrap:"wrap"}}>
              <span style={{fontSize:11,color:G.green,fontWeight:700,fontFamily:"monospace",background:G.green+"18",padding:"2px 9px",borderRadius:8}}>🆔 {selClient.pan}</span>
              <span style={{fontSize:11,fontWeight:700,padding:"2px 9px",borderRadius:8,background:selClient.status==="Active"?"#14532D":"#450A0A",color:selClient.status==="Active"?"#4ADE80":"#FCA5A5"}}>{selClient.status}</span>
              <span style={{fontSize:11,color:G.mut,background:G.bdr,padding:"2px 9px",borderRadius:8}}>{selClient.type}</span>
            </div>
          </div>
          <button onClick={()=>setSelClient(null)} style={{background:"#ffffff15",border:"none",color:G.wh,borderRadius:8,width:32,height:32,cursor:"pointer",fontSize:16,flexShrink:0,display:"flex",alignItems:"center",justifyContent:"center"}}>✕</button>
        </div>
        <div style={{padding:"18px 20px",display:"flex",flexDirection:"column",gap:16}}>
          {/* Contact & Info */}
          <div style={{background:G.card,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"14px 16px"}}>
            <div style={{fontSize:11,color:G.mut,fontWeight:700,letterSpacing:.7,textTransform:"uppercase",marginBottom:10}}>Contact & Information</div>
            <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
              {[{l:"📱 Mobile",v:selClient.mob},{l:"📍 State",v:selClient.state},{l:"📣 Source",v:selClient.src},{l:"📅 Since",v:selClient.added?new Date(selClient.added).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"}):"-"}].map(d=><div key={d.l}>
                <div style={{fontSize:10,color:G.mut,fontWeight:600,marginBottom:2}}>{d.l}</div>
                <div style={{fontSize:13,color:G.txt,fontWeight:600}}>{d.v||"-"}</div>
              </div>)}
              {selClient.email&&<div style={{gridColumn:"1 / -1"}}>
                <div style={{fontSize:10,color:G.mut,fontWeight:600,marginBottom:2}}>✉ Email</div>
                <div style={{fontSize:13,color:G.txt}}>{selClient.email}</div>
              </div>}
              <div style={{gridColumn:"1 / -1"}}>
                <div style={{fontSize:10,color:G.mut,fontWeight:600,marginBottom:2}}>📍 Address</div>
                <div style={{fontSize:13,color:G.txt}}>{selClient.addr||"-"}</div>
              </div>
            </div>
          </div>
          {/* Tax IDs */}
          <div style={{background:G.card,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"14px 16px"}}>
            <div style={{fontSize:11,color:G.mut,fontWeight:700,letterSpacing:.7,textTransform:"uppercase",marginBottom:10}}>Tax Identifiers</div>
            {[{l:"PAN",v:selClient.pan,col:G.green},{l:"GSTIN",v:selClient.gstin||"Not Registered",col:selClient.gstin?G.cyn:G.bdr}].map(d=><div key={d.l} style={{display:"flex",justifyContent:"space-between",alignItems:"center",padding:"7px 0",borderBottom:`1px solid ${G.bdr}`}}>
              <span style={{fontSize:12,color:G.mut,fontWeight:600}}>{d.l}</span>
              <span style={{fontSize:12,color:d.col,fontFamily:"monospace",fontWeight:700}}>{d.v}</span>
            </div>)}
          </div>
          {/* Remarks */}
          {selClient.note&&<div style={{background:G.card,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"14px 16px"}}>
            <div style={{fontSize:11,color:G.mut,fontWeight:700,letterSpacing:.7,textTransform:"uppercase",marginBottom:8}}>📝 Remarks</div>
            <div style={{fontSize:13,color:G.txt,lineHeight:1.6}}>{selClient.note}</div>
          </div>}
          {/* All Assignments */}
          <div style={{background:G.card,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"14px 16px"}}>
            <div style={{fontSize:11,color:G.mut,fontWeight:700,letterSpacing:.7,textTransform:"uppercase",marginBottom:10}}>
              📋 All Assignments ({works.filter(w=>w.pan===selClient.pan).length})
            </div>
            {works.filter(w=>w.pan===selClient.pan).length===0
            ?<div style={{fontSize:12,color:G.bdr}}>No assignments yet</div>
            :works.filter(w=>w.pan===selClient.pan).map(w=><div key={w.id} style={{padding:"9px 12px",background:G.surf,borderRadius:9,border:`1px solid ${G.bdr}`,marginBottom:7}}>
              <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:4}}>
                <span style={{fontWeight:700,fontSize:13,color:G.txt}}>{w.svc}</span>
                <span style={{fontSize:11,padding:"2px 8px",borderRadius:10,fontWeight:700,background:w.status==="Completed"?"#14532D":w.status==="In Progress"?"#1E1B4B":"#431407",color:w.status==="Completed"?"#4ADE80":w.status==="In Progress"?"#A5B4FC":"#FCD34D"}}>{w.status}</span>
              </div>
              <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
                <span style={{fontSize:11,color:G.mut}}>📅 {w.fy}</span>
                <span style={{fontSize:11,color:G.mut}}>👤 {w.staff}</span>
                <span style={{fontSize:11,color:w.status==="Completed"?G.g3:(new Date(w.due)<new Date()?G.red:G.mut)}}>🗓 Due: {w.due?new Date(w.due).toLocaleDateString("en-IN",{day:"2-digit",month:"short",year:"numeric"}):"-"}</span>
              </div>
            </div>)}
          </div>
          {/* Portal Credentials */}
          <div style={{background:G.card,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"14px 16px"}}>
            <div style={{fontSize:11,color:G.mut,fontWeight:700,letterSpacing:.7,textTransform:"uppercase",marginBottom:10}}>
              🔐 Portal Credentials ({PORTALS.filter(p=>selClient.portals[p.key]?.on).length} Active)
            </div>
            {PORTALS.filter(p=>selClient.portals[p.key]?.on).length===0
            ?<div style={{fontSize:12,color:G.bdr}}>No portals configured</div>
            :PORTALS.filter(p=>selClient.portals[p.key]?.on).map(p=><div key={p.key} style={{display:"flex",gap:10,padding:"10px 12px",background:G.surf,border:`1px solid ${p.col}25`,borderRadius:9,marginBottom:7}}>
              <div style={{width:34,height:34,borderRadius:9,background:p.col+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,flexShrink:0}}>{p.icon}</div>
              <div style={{flex:1,minWidth:0}}>
                <div style={{fontSize:11,color:p.col,fontWeight:700,marginBottom:4}}>{p.label}</div>
                <div style={{display:"flex",gap:6,marginBottom:3}}>
                  <span style={{fontSize:10,color:G.mut,minWidth:22}}>ID</span>
                  <span style={{fontSize:12,color:G.wh,fontFamily:"monospace",fontWeight:600,flex:1,overflow:"hidden",textOverflow:"ellipsis"}}>{selClient.portals[p.key].id||"-"}</span>
                </div>
                <div style={{display:"flex",gap:6,alignItems:"center"}}>
                  <span style={{fontSize:10,color:G.mut,minWidth:22}}>PW</span>
                  <PortalPw pw={selClient.portals[p.key].pw} pid={selClient.pan+"_"+p.key}/>
                </div>
              </div>
            </div>)}
          </div>
        </div>
      </div>
    </div>}
    <div style={{display:"flex",gap:8,flexWrap:"wrap",alignItems:"center"}}>
      <div style={{display:"flex",gap:5,flexWrap:"wrap"}}>{["All","Pending","In Progress","Completed","Overdue"].map(f=><button key={f} onClick={()=>setFlt(f)} style={{padding:"5px 12px",borderRadius:18,border:`1.5px solid ${flt===f?G.green:G.bdr}`,cursor:"pointer",fontSize:12,fontWeight:600,background:flt===f?G.green+"18":"transparent",color:flt===f?G.green:G.mut}}>{f} <span style={{opacity:.6}}>({cnt[f]})</span></button>)}</div>
      <div style={{display:"flex",gap:5,flexWrap:"wrap",padding:"3px 8px",background:G.card,border:`1px solid ${G.bdr}`,borderRadius:20,alignItems:"center"}}>
        <span style={{fontSize:11,color:G.mut,fontWeight:600,whiteSpace:"nowrap"}}>📅 FY:</span>
        {["All",...dd.fyOptions].map(f=><button key={f} onClick={()=>setFy(f)} style={{padding:"4px 11px",borderRadius:14,border:"none",cursor:"pointer",fontSize:12,fontWeight:fy===f?700:500,background:fy===f?`linear-gradient(135deg,${G.g2},${G.green})`:"transparent",color:fy===f?"#fff":G.mut,transition:"all .15s",whiteSpace:"nowrap"}}>{f}{f!=="All"&&<span style={{fontSize:10,opacity:.7}}> ({works.filter(w=>w.fy===f).length})</span>}</button>)}
      </div>
      <div style={{position:"relative",flex:1,minWidth:150}}><span style={{position:"absolute",left:9,top:"50%",transform:"translateY(-50%)",color:G.mut}}>🔍</span><input value={q} onChange={e=>setQ(e.target.value)} placeholder="Search..." style={{...IS,paddingLeft:28,fontSize:12}}/></div>
      {(flt !== "All" || fy !== getCurrentFY() || q !== "") && (
        <button 
          onClick={clearFilters}
          style={{
            padding: "5px 14px",
            borderRadius: 18,
            border: `1.5px solid ${G.red}aa`,
            background: `${G.red}18`,
            color: G.red,
            cursor: "pointer",
            fontSize: 12,
            fontWeight: 700,
            whiteSpace: "nowrap",
            transition: "all 0.2s"
          }}
          title="Reset all search filters"
        >✕ Clear Filter</button>
      )}
      {!feesOn?<button onClick={()=>setShowAuth(true)} style={{padding:"6px 14px",borderRadius:9,border:`1px solid ${G.amb}44`,background:"#43140720",color:G.amb,cursor:"pointer",fontWeight:700,fontSize:12,whiteSpace:"nowrap"}}>🔒 Show Fees</button>
      :<button onClick={()=>setFeesOn(false)} style={{padding:"6px 14px",borderRadius:9,border:`1px solid ${G.red}44`,background:"#450A0A",color:G.red,cursor:"pointer",fontWeight:700,fontSize:12,whiteSpace:"nowrap"}}>🔓 Hide Fees</button>}
    </div>
    {/* Year-wise summary bar */}
    <div className="kpi-grid-5" style={{display:"grid",gridTemplateColumns:"repeat(5,1fr)",gap:10}}>
      {[
        {l:"Total Works",   v:fw.length,                                    icon:"📋", col:G.green},
        {l:"Pending",       v:fw.filter(w=>w.status==="Pending").length,    icon:"⏳", col:G.amb},
        {l:"In Progress",   v:fw.filter(w=>w.status==="In Progress").length,icon:"🔵", col:G.cyn},
        {l:"Completed",     v:fw.filter(w=>w.status==="Completed").length,  icon:"✅", col:G.g3},
        {l:"Overdue",       v:fw.filter(isOD).length,                       icon:"🔴", col:G.red},
      ].map(k=><div key={k.l} style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"11px 14px",display:"flex",gap:10,alignItems:"center"}}>
        <div style={{width:34,height:34,borderRadius:9,background:k.col+"18",display:"flex",alignItems:"center",justifyContent:"center",fontSize:17,flexShrink:0}}>{k.icon}</div>
        <div><div style={{fontWeight:800,fontSize:20,color:k.col,lineHeight:1}}>{k.v}</div><div style={{fontSize:10,color:G.mut,marginTop:2}}>{k.l}</div></div>
      </div>)}
    </div>
    <Crd sty={{padding:0,overflow:"hidden"}}><div style={{overflowX:"auto"}}><table style={{width:"100%",borderCollapse:"collapse",fontSize:12}}>
      <thead><tr style={{background:G.bg}}>{["Client","PAN","Service","FY","Due","Staff","Status",...(feesOn?["Fees","Comm.","Received","O/S"]:["💰"]),"O/D","Portals",""].map(h=><th key={h} style={{padding:"9px 11px",textAlign:"left",color:feesOn&&["Fees","Comm.","Received","O/S"].includes(h)?G.amb:G.mut,fontWeight:600,fontSize:11,whiteSpace:"nowrap",borderBottom:`1px solid ${G.bdr}`}}>{h}</th>)}</tr></thead>
      <tbody>{fw.length===0?<tr><td colSpan={15} style={{padding:28,textAlign:"center",color:G.bdr}}>No records</td></tr>
      :fw.map((w,i)=>{
        const linked=linkedInvoices(w,invoices),wRcvd=workReceived(w,invoices,receipts),wBilled=workBilled(w,invoices),os=wBilled-wRcvd,pStat=workPayStatus(w,invoices,receipts),od=odDays(w),cl=clients.find(c=>c.pan===w.pan),aP=PORTALS.filter(p=>cl?.portals[p.key]?.on);
        const linkableInvs=(invoices||[]).filter(inv=>inv.pan===w.pan&&(!inv.workId||inv.workId===""));
        return <tr key={w.id} style={{borderTop:`1px solid ${G.bdr}`,background:isOD(w)?"#450A0A08":i%2?"#0F1D1408":"transparent"}}>
          <td style={{padding:"9px 11px",whiteSpace:"nowrap"}}>
                  <button onClick={()=>setSelClient(clients.find(c=>c.pan===w.pan)||null)} style={{background:"none",border:"none",cursor:"pointer",fontWeight:700,color:G.green,fontSize:12,padding:0,textDecoration:"underline",textUnderlineOffset:3,textDecorationColor:G.green+"55",whiteSpace:"nowrap"}} title="View client details">{w.cn}</button>
                </td>
          <td style={{padding:"9px 11px",color:G.g3,fontFamily:"monospace",fontSize:11}}>{w.pan}</td>
          <td style={{padding:"9px 11px",color:G.mut,whiteSpace:"nowrap",maxWidth:120,overflow:"hidden",textOverflow:"ellipsis"}}>{w.svc}</td>
          <td style={{padding:"9px 11px",color:G.mut}}>{w.fy}</td>
          <td style={{padding:"9px 11px",whiteSpace:"nowrap",color:isOD(w)?G.red:G.mut}}>{fd(w.due)}</td>
          <td style={{padding:"9px 11px",color:G.mut,whiteSpace:"nowrap"}}>{w.staff}</td>
          <td style={{padding:"9px 11px"}}><select value={w.status} onChange={e=>setWorks(p=>p.map(x=>x.id===w.id?{...x,status:e.target.value}:x))} style={{background:"transparent",border:"none",cursor:"pointer",fontSize:12,color:w.status==="Completed"?G.green:w.status==="In Progress"?G.cyn:G.amb}}>{["Pending","In Progress","Completed"].map(s=><option key={s} value={s} style={{background:"#0B1610"}}>{s}</option>)}</select></td>
          {feesOn?<>
            <td style={{padding:"9px 11px",minWidth:78}}><input type="number" value={w.fees||0} onChange={e=>setWorks(p=>p.map(x=>x.id===w.id?{...x,fees:Number(e.target.value)||0}:x))} style={{...IS,padding:"3px 6px",fontSize:11,width:72,color:G.green}}/></td>
            <td style={{padding:"9px 11px",minWidth:78}}><input type="number" value={w.comm||0} onChange={e=>setWorks(p=>p.map(x=>x.id===w.id?{...x,comm:Number(e.target.value)||0}:x))} style={{...IS,padding:"3px 6px",fontSize:11,width:68,color:G.amb}}/></td>
            <td style={{padding:"9px 11px",minWidth:120,position:"relative"}}>
              <div style={{display:"flex",alignItems:"center",gap:5}}>
                <span style={{color:G.g3,fontWeight:700,fontSize:11,whiteSpace:"nowrap"}}>{inr(wRcvd)}</span>
                <button onClick={()=>setLinkPop(linkPop===w.id?null:w.id)} title="Link invoice / view payments" style={{background:linked.length?G.cyn+"18":G.amb+"18",border:`1px solid ${linked.length?G.cyn:G.amb}55`,color:linked.length?G.cyn:G.amb,borderRadius:5,padding:"1px 6px",cursor:"pointer",fontSize:10,fontWeight:800,flexShrink:0,whiteSpace:"nowrap"}}>🔗{linked.length>0?` ${linked.length}`:""}</button>
                <span style={{fontSize:9,fontWeight:800,padding:"1px 5px",borderRadius:6,whiteSpace:"nowrap",background:pStat==="Paid"?"#14532D":pStat==="Partial"?"#431407":pStat==="Unlinked"?"#1E1B4B":"#450A0A",color:pStat==="Paid"?"#4ADE80":pStat==="Partial"?"#FCD34D":pStat==="Unlinked"?"#A5B4FC":"#FCA5A5"}}>{pStat}</span>
              </div>
              {linkPop===w.id&&<>
                <div onClick={()=>{setLinkPop(null);setLinkSel("");}} style={{position:"fixed",inset:0,zIndex:900}}/>
                <div onClick={e=>e.stopPropagation()} style={{position:"absolute",top:"100%",left:0,marginTop:5,background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:11,padding:11,width:260,zIndex:901,boxShadow:"0 16px 40px #000B"}}>
                  <div style={{fontSize:11,color:G.mut,fontWeight:700,marginBottom:7}}>🔗 Linked Invoices — {w.svc}</div>
                  <div style={{maxHeight:150,overflowY:"auto",marginBottom:8,display:"flex",flexDirection:"column",gap:5}}>
                    {linked.length===0
                      ?<div style={{fontSize:11,color:G.bdr}}>No invoice linked yet{w.rcvd>0?` — showing legacy amount ${inr(w.rcvd)}`:""}</div>
                      :linked.map(inv=>{const ir=invoiceReceived(inv,receipts),ist=computeInvStatus(inv,receipts);return <div key={inv.id} style={{background:G.card,borderRadius:8,padding:"6px 8px"}}>
                        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
                          <span style={{fontSize:11,color:G.ind,fontFamily:"monospace",fontWeight:700}}>{inv.id}</span>
                          <button onClick={()=>unlinkInvoice(inv.id)} title="Unlink" style={{background:"none",border:"none",color:G.red,cursor:"pointer",fontSize:11,padding:0}}>✕</button>
                        </div>
                        <div style={{display:"flex",justifyContent:"space-between",fontSize:11,marginTop:2}}>
                          <span style={{color:G.mut}}>Total {inr(inv.total)}</span>
                          <span style={{color:G.g3,fontWeight:700}}>Rcvd {inr(ir)}</span>
                        </div>
                        <span style={{fontSize:9,fontWeight:800,padding:"1px 5px",borderRadius:6,background:ist==="Paid"?"#14532D":ist==="Partial"?"#431407":"#450A0A",color:ist==="Paid"?"#4ADE80":ist==="Partial"?"#FCD34D":"#FCA5A5"}}>{ist}</span>
                      </div>;})}
                  </div>
                  <div style={{fontSize:11,color:G.mut,marginBottom:6,paddingTop:6,borderTop:`1px solid ${G.bdr}`}}>Link an existing invoice for {w.cn}:</div>
                  {linkableInvs.length===0
                    ?<div style={{fontSize:11,color:G.bdr,marginBottom:4}}>No unlinked invoices for this client. Raise one from Billing & Invoicing, then link it here.</div>
                    :<div style={{display:"flex",gap:6}}>
                      <select value={linkSel} onChange={e=>setLinkSel(e.target.value)} style={{...IS,padding:"5px 6px",fontSize:11,flex:1}}>
                        <option value="">Select invoice...</option>
                        {linkableInvs.map(inv=><option key={inv.id} value={inv.id} style={{background:"#0B1610"}}>{inv.id} · {inr(inv.total)}</option>)}
                      </select>
                      <button onClick={()=>linkInvoice(w.id,linkSel)} disabled={!linkSel} style={{padding:"5px 11px",borderRadius:8,border:"none",background:linkSel?`linear-gradient(135deg,${G.g2},${G.green})`:G.bdr,color:"#fff",fontWeight:700,fontSize:11,cursor:linkSel?"pointer":"not-allowed",whiteSpace:"nowrap"}}>Link</button>
                    </div>}
                </div>
              </>}
            </td>
            <td style={{padding:"9px 11px",color:os>0?G.red:G.g3,fontWeight:700,whiteSpace:"nowrap"}}>{inr(os)}</td>
          </>:<td style={{padding:"9px 11px",color:G.bdr,fontSize:12}}>🔒 Locked</td>}
          <td style={{padding:"9px 11px"}}>{od>0?<span style={{background:"#450A0A",color:G.red,padding:"2px 7px",borderRadius:8,fontWeight:700}}>{od}d</span>:<span style={{color:G.bdr}}>-</span>}</td>
          <td style={{padding:"9px 11px"}}>
            <div style={{display:"flex",gap:3,flexWrap:"wrap"}}>{aP.slice(0,3).map(p=><span key={p.key} title={p.label} onClick={()=>setOpP(s=>({...s,[w.id]:!s[w.id]}))} style={{fontSize:13,cursor:"pointer"}}>{p.icon}</span>)}
              {aP.length>0&&<button onClick={()=>setOpP(s=>({...s,[w.id]:!s[w.id]}))} style={{background:"none",border:"none",cursor:"pointer",color:G.mut,fontSize:11,padding:"0 2px"}}>🔑</button>}
            </div>
            {opP[w.id]&&cl&&<div style={{marginTop:5,background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:8,padding:"8px 10px",minWidth:180,position:"relative",zIndex:10}}>
              {aP.map(p=><div key={p.key} style={{marginBottom:4}}>
                <div style={{fontSize:10,color:p.col,fontWeight:700}}>{p.icon} {p.label}</div>
                <div style={{fontSize:11,color:G.mut}}>ID: <span style={{color:G.wh,fontFamily:"monospace"}}>{cl.portals[p.key].id||"-"}</span></div>
                <div style={{fontSize:11,color:G.mut,display:"flex",alignItems:"center",gap:4}}>PW: <span style={{color:G.wh,fontFamily:"monospace"}}>{vpw[`${w.id}_${p.key}`]?cl.portals[p.key].pw||"-":"******"}</span>
                  <button onClick={()=>setVpw(s=>({...s,[`${w.id}_${p.key}`]:!s[`${w.id}_${p.key}`]}))} style={{background:"none",border:"none",cursor:"pointer",color:G.mut,fontSize:11,padding:0}}>{vpw[`${w.id}_${p.key}`]?"🙈":"👁"}</button>
                </div>
              </div>)}
              {aP.length===0&&<div style={{fontSize:12,color:G.mut}}>No portals</div>}
            </div>}
          </td>
          <td style={{padding:"9px 11px"}}><div style={{display:"flex",gap:4}}><button onClick={()=>setEditW(w)} style={{background:"transparent",border:`1px solid ${G.bdr}`,color:G.cyn,borderRadius:6,padding:"3px 7px",cursor:"pointer",fontSize:12}}>✏️</button><button onClick={()=>setDelW(w)} title="Delete Work" style={{background:"#450A0A",border:`1px solid ${G.red}44`,color:G.red,borderRadius:6,padding:"3px 7px",cursor:"pointer",fontSize:12}}>🗑</button></div></td>
        </tr>;
      })}</tbody>
    </table></div></Crd>
    {delW&&<div style={{position:"fixed",inset:0,background:"#000C",zIndex:6000,display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:16,padding:28,width:360,boxShadow:"0 20px 60px #000C",textAlign:"center"}}>
        <div style={{fontSize:36,marginBottom:12}}>🗑</div>
        <div style={{fontWeight:700,fontSize:16,color:G.wh,marginBottom:8}}>Delete this work item?</div>
        <div style={{fontSize:13,color:G.mut,marginBottom:20,padding:"8px 14px",background:G.red+"10",border:`1px solid ${G.red}33`,borderRadius:9}}>{delW.svc} — {delW.cn}</div>
        <div style={{display:"flex",gap:10}}>
          <button onClick={doDeleteWork} style={{flex:1,padding:"11px",borderRadius:10,border:"none",cursor:"pointer",background:G.red,color:"#fff",fontWeight:700,fontSize:14}}>Yes, Delete</button>
          <button onClick={()=>setDelW(null)} style={{flex:1,padding:"11px",borderRadius:10,border:`1px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer",fontWeight:600}}>Cancel</button>
        </div>
      </div>
    </div>}
  </div>;
}

// ─── Developer Tab ────────────────────────────────────────────────────────────

// ─── Performa Designer ────────────────────────────────────────────────────────
// Simple drag-anywhere image overlay designer on top of a static invoice preview
function PerformaDesigner({firmSettings, setFirmSettings}){
  const F = firmSettings || {};
  const [selImg, setSelImg] = useState(null); // which image is selected: 'logo'|'stamp'|'signature'|'qrCode'
  const [dragging, setDragging] = useState(false);
  const [dragStart, setDragStart] = useState({x:0,y:0});
  // FIX: real GST invoices render 2 extra table rows (CGST + SGST) that this mock
  // never showed, so any stamp/signature placed here landed lower than intended
  // on GST invoices (overlapping the CGST/SGST row instead of sitting near the
  // signature). Default to previewing WITH GST since that's the layout most
  // likely to be affected; let the user toggle to check both layouts.
  const [gstPreview, setGstPreview] = useState(true);
  const previewRef = useState(null);

  const IMGS = [
    {key:'logo',        label:'Logo',             defX:16,  defY:16,  defW:72,  defH:72},
    {key:'stamp',       label:'Stamp',            defX:490, defY:870, defW:80,  defH:80},
    {key:'signature',   label:'Signature',        defX:450, defY:930, defW:140, defH:44},
    {key:'statusStamp', label:'Universal Stamp',  defX:280, defY:380, defW:160, defH:160},
  ];

  const getPos = (key, def) => F[key+'X'] !== undefined ? F[key+'X'] : def;
  const getVal = (key, def) => F[key] !== undefined ? F[key] : def;

  const startDrag = (e, key) => {
    e.preventDefault();
    e.stopPropagation();
    setSelImg(key);
    setDragging(true);
    const img = IMGS.find(i=>i.key===key);
    setDragStart({
      mx: e.clientX, my: e.clientY,
      ox: getPos(key+'X', img.defX),
      oy: getPos(key+'Y', img.defY),
    });
  };

  const onMouseMove = (e) => {
    if(!dragging || !selImg) return;
    const dx = e.clientX - dragStart.mx;
    const dy = e.clientY - dragStart.my;
    setFirmSettings(p => ({...p,
      [selImg+'X']: dragStart.ox + dx,
      [selImg+'Y']: dragStart.oy + dy,
    }));
  };

  const onMouseUp = () => setDragging(false);

  const sel = selImg ? IMGS.find(i=>i.key===selImg) : null;

  // Get sorted images by zIndex for rendering order
  const sortedImgs = [...IMGS].sort((a,b) => (getVal(a.key+'Z',1)) - (getVal(b.key+'Z',1)));

  // Text style helper
  const txtSty = (el, defSz, defBold, defItalic, defCol) => ({
    fontSize: F2[el+'Sz'] !== undefined ? F2[el+'Sz'] : defSz,
    fontWeight: (F2[el+'Bold'] !== undefined ? F2[el+'Bold'] : defBold) ? 900 : 400,
    fontStyle: (F2[el+'Italic'] !== undefined ? F2[el+'Italic'] : defItalic) ? 'italic' : 'normal',
    color: F2[el+'Col'] || defCol,
  });

  const bringForward = (key) => {
    const curZ = getVal(key+'Z', 1);
    setFirmSettings(p => ({...p, [key+'Z']: curZ + 1}));
  };
  const sendBack = (key) => {
    const curZ = getVal(key+'Z', 1);
    setFirmSettings(p => ({...p, [key+'Z']: Math.max(0, curZ - 1)}));
  };

  const F2 = firmSettings || {};
  const clrP = F2.clrPrimary || '#1B5E20';
  const clrA = F2.clrAccent || '#F1F8E9';
  const clrB = F2.clrBorder || '#C8E6C9';
  // Mirrors the real invoice's font-size formula exactly (bFsz/tFsz there),
  // so image positions set here still line up on the actual printed page.
  const invScaleD = (F2.invoiceScale ?? 100)/100;
  const dBFsz = Math.round((F2.bodyFontSz||15)*invScaleD);
  const dTFsz = Math.round((F2.titleFontSz||28)*invScaleD);

  return (
    <div style={{display:'flex', gap:14, height:'calc(100vh - 160px)', overflow:'hidden'}}>

      {/* Left: Image controls */}
      <div style={{width:230, flexShrink:0, display:'flex', flexDirection:'column', gap:10, overflowY:'auto'}}>
        <div style={{background:G.card, border:`1px solid ${G.bdr}`, borderRadius:11, padding:'12px 14px'}}>
          <div style={{fontWeight:700, fontSize:13, color:G.wh, marginBottom:10}}>🖼 Images</div>
          <div style={{fontSize:11, color:G.mut, marginBottom:10, lineHeight:1.6}}>
            Click an image on the preview to select it. Drag to reposition freely.
          </div>
          {IMGS.map(img => {
            const hasImg = !!F2[img.key];
            const isSelected = selImg === img.key;
            return (
              <div key={img.key} onClick={() => setSelImg(img.key)}
                style={{padding:'8px 10px', borderRadius:9, marginBottom:6, cursor:'pointer',
                  border:`2px solid ${isSelected ? G.green : G.bdr}`,
                  background: isSelected ? G.green+'18' : G.card,
                  opacity: hasImg ? 1 : 0.45}}>
                <div style={{display:'flex', alignItems:'center', justifyContent:'space-between'}}>
                  <span style={{fontSize:12, fontWeight:700, color: isSelected ? G.green : G.mut}}>{img.label}</span>
                  {hasImg
                    ? <span style={{fontSize:10, background:G.g2+'30', color:G.g3, padding:'1px 7px', borderRadius:8, fontWeight:600}}>Uploaded</span>
                    : <span style={{fontSize:10, color:G.bdr}}>No image</span>}
                </div>
                {hasImg && isSelected && (
                  <div style={{display:'flex', gap:5, marginTop:6}}>
                    <button onClick={e=>{e.stopPropagation();sendBack(img.key);}} style={{flex:1, fontSize:11, padding:'3px', borderRadius:6, border:`1px solid ${G.bdr}`, background:'transparent', color:G.mut, cursor:'pointer'}}>Send Back</button>
                    <button onClick={e=>{e.stopPropagation();bringForward(img.key);}} style={{flex:1, fontSize:11, padding:'3px', borderRadius:6, border:`1px solid ${G.green}55`, background:G.green+'10', color:G.green, cursor:'pointer'}}>Bring Front</button>
                  </div>
                )}
              </div>
            );
          })}
        </div>

        {/* Single controller for selected image */}
        {selImg && F2[selImg] && (
          <div style={{background:G.card, border:`1px solid ${G.green}55`, borderRadius:11, padding:'12px 14px'}}>
            <div style={{fontWeight:700, fontSize:12, color:G.green, marginBottom:10}}>
              Adjust: {sel?.label}
            </div>
            {[
              {k:'W', label:'Width',      def:sel?.defW||80,  min:20, max:400, unit:'px'},
              {k:'H', label:'Height',     def:sel?.defH||80,  min:10, max:400, unit:'px'},
              {k:'Brightness', label:'Brightness', def:100, min:10, max:200, unit:'%'},
              {k:'Contrast',   label:'Contrast',   def:100, min:10, max:300, unit:'%'},
              {k:'Saturate',   label:'Saturate',   def:100, min:0,  max:300, unit:'%'},
              {k:'Opacity',    label:'Opacity',    def:100, min:10, max:100, unit:'%'},
            ].map(sl => {
              const val = F2[selImg+sl.k] !== undefined ? F2[selImg+sl.k] : sl.def;
              return (
                <div key={sl.k} style={{marginBottom:9}}>
                  <div style={{display:'flex', justifyContent:'space-between', fontSize:11, color:G.mut, marginBottom:3}}>
                    <span>{sl.label}</span>
                    <span style={{color:G.green, fontWeight:700}}>{val}{sl.unit}</span>
                  </div>
                  <input type="range" min={sl.min} max={sl.max} value={val}
                    onChange={e => setFirmSettings(p=>({...p, [selImg+sl.k]: Number(e.target.value)}))}
                    style={{width:'100%', accentColor:G.green, cursor:'pointer'}}/>
                </div>
              );
            })}
            <div style={{display:'flex', gap:6, marginTop:4}}>
              <div style={{flex:1, fontSize:11, color:G.mut, background:G.surf, padding:'4px 8px', borderRadius:7, border:`1px solid ${G.bdr}`}}>
                x:{Math.round(getPos(selImg+'X', sel?.defX||0))} y:{Math.round(getPos(selImg+'Y', sel?.defY||0))}
              </div>
              <button onClick={()=>setFirmSettings(p=>{
                const n={...p};
                ['X','Y','W','H','Brightness','Contrast','Saturate','Opacity','Z'].forEach(k=>delete n[selImg+k]);
                return n;
              })} style={{fontSize:11, padding:'4px 9px', borderRadius:7, border:`1px solid ${G.red}44`, background:'transparent', color:G.red, cursor:'pointer'}}>Reset</button>
            </div>
          </div>
        )}

        {/* Text Style Controller */}
        <div style={{background:G.card, border:`1px solid ${G.bdr}`, borderRadius:11, padding:'12px 14px'}}>
          <div style={{fontWeight:700, fontSize:12, color:G.txt, marginBottom:10}}>Text Style</div>
          {/* Element selector - like selecting text */}
          <div style={{marginBottom:10}}>
            <div style={{fontSize:10, color:G.mut, fontWeight:700, marginBottom:5, textTransform:'uppercase', letterSpacing:.6}}>Select Element</div>
            <select value={F2.selTextEl||'firmName'} onChange={e=>setFirmSettings(p=>({...p,selTextEl:e.target.value}))}
              style={{...Object.assign({},
                {background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:7,color:G.wh,fontSize:12,padding:'7px 10px',outline:'none',width:'100%',cursor:'pointer'}
              )}}>
              {[
                {v:'firmName',    l:'Firm Name'},
                {v:'invoiceType', l:'Invoice Title (TAX INVOICE)'},
                {v:'sectionHdr',  l:'Section Headings'},
                {v:'customerText',l:'Customer / Invoice Details'},
                {v:'tableHeader', l:'Table Header'},
                {v:'srNo',        l:'Sr. No. Cell'},
                {v:'totalLabel',  l:'Total Label'},
                {v:'totalAmt',    l:'Total Amount'},
                {v:'bankLabel',   l:'Bank Details Label'},
                {v:'bankText',    l:'Bank Details Text'},
                {v:'termsText',   l:'Terms Text'},
                {v:'signature',   l:'Authorised Signatory'},
              ].map(o=><option key={o.v} value={o.v} style={{background:'#0B1610'}}>{o.l}</option>)}
            </select>
          </div>
          {/* Single set of controls for selected element */}
          {(()=>{
            const el = F2.selTextEl || 'firmName';
            const defaults = {
              firmName:    {sz:22, bold:true,  italic:true,  col:'#1B5E20'},
              invoiceType: {sz:18, bold:true,  italic:true,  col:'#1B5E20'},
              sectionHdr:  {sz:12, bold:true,  italic:true,  col:'#1B5E20'},
              customerText:{sz:11, bold:false, italic:false, col:'#000000'},
              tableHeader: {sz:11, bold:true,  italic:true,  col:'#1B5E20'},
              srNo:        {sz:11, bold:false, italic:true,  col:'#000000'},
              totalLabel:  {sz:12, bold:true,  italic:true,  col:'#000000'},
              totalAmt:    {sz:14, bold:true,  italic:false, col:'#1B5E20'},
              bankLabel:   {sz:11, bold:true,  italic:true,  col:'#1B5E20'},
              bankText:    {sz:10, bold:false, italic:false, col:'#333333'},
              termsText:   {sz:9,  bold:false, italic:true,  col:'#555555'},
              signature:   {sz:11, bold:true,  italic:true,  col:'#000000'},
            };
            const def = defaults[el] || {sz:11, bold:false, italic:false, col:'#000000'};
            const sz    = F2[el+'Sz']    !== undefined ? F2[el+'Sz']    : def.sz;
            const bold  = F2[el+'Bold']  !== undefined ? F2[el+'Bold']  : def.bold;
            const italic= F2[el+'Italic']!== undefined ? F2[el+'Italic']: def.italic;
            const col   = F2[el+'Col']   !== undefined ? F2[el+'Col']   : def.col;
            return <>
              {/* Toolbar row like a text editor */}
              <div style={{display:'flex', gap:6, alignItems:'center', marginBottom:10, padding:'6px 8px', background:G.surf, borderRadius:8, border:`1px solid ${G.bdr}`, flexWrap:'wrap'}}>
                {/* Bold */}
                <button onClick={()=>setFirmSettings(p=>({...p,[el+'Bold']:!bold}))}
                  style={{width:30, height:28, borderRadius:6, border:`1px solid ${bold?G.green:G.bdr}`, background:bold?G.green+'20':'transparent', color:bold?G.green:G.mut, fontWeight:900, fontSize:14, cursor:'pointer'}}>B</button>
                {/* Italic */}
                <button onClick={()=>setFirmSettings(p=>({...p,[el+'Italic']:!italic}))}
                  style={{width:30, height:28, borderRadius:6, border:`1px solid ${italic?G.green:G.bdr}`, background:italic?G.green+'20':'transparent', color:italic?G.green:G.mut, fontWeight:700, fontSize:14, fontStyle:'italic', cursor:'pointer'}}>I</button>
                {/* Color picker */}
                <div style={{position:'relative', display:'flex', alignItems:'center', gap:4}}>
                  <div style={{width:22, height:22, borderRadius:5, background:col, border:`1px solid ${G.bdr}`, cursor:'pointer'}} onClick={()=>document.getElementById('txtCol_'+el).click()}/>
                  <input id={'txtCol_'+el} type="color" value={col}
                    onChange={e=>setFirmSettings(p=>({...p,[el+'Col']:e.target.value}))}
                    style={{position:'absolute', opacity:0, width:0, height:0, pointerEvents:'none'}}/>
                  <span style={{fontSize:10, color:G.mut}}>Color</span>
                </div>
                {/* Reset this element */}
                <button onClick={()=>setFirmSettings(p=>{const n={...p};[el+'Sz',el+'Bold',el+'Italic',el+'Col'].forEach(k=>delete n[k]);return n;})}
                  style={{marginLeft:'auto', fontSize:10, padding:'3px 8px', borderRadius:6, border:`1px solid ${G.bdr}`, background:'transparent', color:G.mut, cursor:'pointer'}}>Reset</button>
              </div>
              {/* Font size slider */}
              <div>
                <div style={{display:'flex', justifyContent:'space-between', fontSize:11, color:G.mut, marginBottom:3}}>
                  <span>Font Size</span><span style={{color:G.amb, fontWeight:700}}>{sz}px</span>
                </div>
                <input type="range" min={7} max={40} value={sz}
                  onChange={e=>setFirmSettings(p=>({...p,[el+'Sz']:Number(e.target.value)}))}
                  style={{width:'100%', accentColor:G.amb, cursor:'pointer'}}/>
              </div>
              {/* Live preview */}
              <div style={{marginTop:8, padding:'8px 10px', background:G.surf, borderRadius:8, border:`1px solid ${G.bdr}`, overflow:'hidden'}}>
                <span style={{fontSize:sz>20?16:sz, fontWeight:bold?900:400, fontStyle:italic?'italic':'normal', color:col, fontFamily:"'Times New Roman',Times,serif"}}>
                  Preview Text
                </span>
              </div>
            </>;
          })()}
        </div>
        {/* Color controls */}
        <div style={{background:G.card, border:`1px solid ${G.bdr}`, borderRadius:11, padding:'12px 14px'}}>
          <div style={{fontWeight:700, fontSize:12, color:G.txt, marginBottom:10}}>Colors</div>
          {[
            {k:'clrPrimary', label:'Primary', def:'#1B5E20'},
            {k:'clrAccent',  label:'Header BG', def:'#F1F8E9'},
            {k:'clrBorder',  label:'Border', def:'#C8E6C9'},
          ].map(c => (
            <div key={c.k} style={{display:'flex', alignItems:'center', justifyContent:'space-between', marginBottom:8}}>
              <span style={{fontSize:11, color:G.mut}}>{c.label}</span>
              <div style={{display:'flex', alignItems:'center', gap:6}}>
                <div style={{width:20, height:20, borderRadius:5, background:F2[c.k]||c.def, border:`1px solid ${G.bdr}`}}/>
                <input type="color" value={F2[c.k]||c.def}
                  onChange={e=>setFirmSettings(p=>({...p,[c.k]:e.target.value}))}
                  style={{width:30, height:24, border:'none', background:'transparent', cursor:'pointer', padding:0}}/>
                <button onClick={()=>setFirmSettings(p=>({...p,[c.k]:c.def}))}
                  style={{fontSize:10, padding:'2px 5px', borderRadius:5, border:`1px solid ${G.bdr}`, background:'transparent', color:G.mut, cursor:'pointer'}}>R</button>
              </div>
            </div>
          ))}
        </div>
      </div>

      {/* Right: Invoice preview with drag overlay */}
      <div style={{flex:1, overflow:'auto', background:'#9CA3AF', borderRadius:10, padding:'12px', display:'flex', flexDirection:'column', alignItems:'center'}}
        onMouseMove={onMouseMove} onMouseUp={onMouseUp} onMouseLeave={onMouseUp}>

        {/* Static invoice paper */}
        <div style={{position:'relative', width:794, height:1123, margin:'0 auto', flexShrink:0}}>
          {/* REDESIGN: mirrors the real invoice's fixed-height flex-column layout
              exactly (same header/title/customer/footer/terms as flexShrink:0,
              same single flexible service-table region) so the canvas the user
              positions images on is pixel-identical to the printed page. */}
          <div style={{background:'#fff', fontFamily:"'Times New Roman',Times,serif", color:'#000',
            boxShadow:'0 4px 24px #0004', border:'1px solid #ccc', userSelect:'none', height:1123, width:'100%',
            display:'flex', flexDirection:'column', overflow:'hidden'}}>

            {/* Header - mirrors real invoice: gradient accent bar under header */}
            <div style={{padding:'16px 20px', borderBottom:'4px solid transparent', borderImage:`linear-gradient(90deg,${clrP},${clrB}) 1`, display:'flex', gap:14, alignItems:'center', flexShrink:0}}>
              <div style={{width:F2.logoW||72, height:F2.logoH||72, flexShrink:0}}/>
              <div style={{flex:1, minWidth:0}}>
                <div style={{...txtSty('firmName',dTFsz,true,true,clrP), fontFamily:'Georgia,serif'}}>{F2.name||'Fin-Tax Mitra'}</div>
                <div style={{fontSize:dBFsz-1, color:'#444', marginTop:5, lineHeight:1.8}}>{F2.addr||'17, Sebadol Road, Belgharia, Kolkata, West Bengal - 700049'}</div>
                <div style={{fontSize:dBFsz-1, color:'#444', lineHeight:1.8}}>Ph: {F2.phone||'7980718092'}  |  Email: {F2.email||'care.fintaxmitra@gmail.com'}</div>
              </div>
            </div>

            {/* Invoice title - mirrors real: 6px 20px padding + pill copy label */}
            <div style={{background:clrA, padding:'6px 20px', display:'flex', justifyContent:'space-between', alignItems:'center', borderBottom:`1px solid ${clrB}`}}>
              <div style={{...txtSty('invoiceType',dTFsz-2,true,true,clrP), letterSpacing:3, fontFamily:'Georgia,serif'}}>{F2.invoiceType||'TAX INVOICE'}</div>
              <div style={{fontSize:10, color:clrP, fontStyle:'italic', fontWeight:700, border:`1px solid ${clrP}`, borderRadius:20, padding:'3px 12px', letterSpacing:.5, flexShrink:0, marginLeft:10, background:'#fff'}}>{F2.copyLabel||'ORIGINAL FOR RECIPIENT'}</div>
            </div>

            {/* Customer + Invoice - rounded soft boxes, mirrors real invoice */}
            <div style={{display:'flex', gap:12, padding:'12px 16px', borderBottom:`1px solid ${clrB}`, flexShrink:0}}>
              <div style={{flex:1, border:`1px solid ${clrB}`, borderRadius:10, padding:'12px 16px', background:'#fff'}}>
                <div style={{...txtSty('sectionHdr',dBFsz,true,true,clrP), textAlign:'center', borderBottom:`1px solid ${clrB}`, paddingBottom:6, marginBottom:10}}>Customer Details</div>
                {[['Name','TAPAS KUMAR JANA'],['Address','Chorepalia, Egra, Medinipur East'],['Phone No.','7477590942'],['PAN','AVOPJ4897R']].map(([l,v])=>(
                  <div key={l} style={{display:'flex', gap:6, fontSize:dBFsz, marginBottom:7}}>
                    <span style={{fontWeight:700, minWidth:84}}>{l}</span><span>: {v}</span>
                  </div>
                ))}
              </div>
              <div style={{flex:1, border:`1px solid ${clrB}`, borderRadius:10, padding:'12px 16px', background:'#fff'}}>
                <div style={{...txtSty('sectionHdr',dBFsz,true,true,clrP), textAlign:'center', borderBottom:`1px solid ${clrB}`, paddingBottom:6, marginBottom:10}}>Invoice Details</div>
                {[['Invoice No.',`${F2.invPrefix||'FTM'}/2025-26/10031`],['Invoice Date','12/02/2026'],['Financial Year','2025-26']].map(([l,v])=>(
                  <div key={l} style={{display:'flex', gap:6, fontSize:dBFsz, marginBottom:10}}>
                    <span style={{fontWeight:700, minWidth:104}}>{l}</span><span style={{fontStyle:'italic', fontWeight:700}}>: {v}</span>
                  </div>
                ))}
              </div>
            </div>

            {/* Service table - the ONE flexible region, same as the real invoice */}
            <div style={{flex:'1 1 auto', minHeight:0, position:'relative', borderBottom:`1px solid ${clrB}`}}>
            <div style={{position:'absolute', inset:0, overflow:'hidden'}}>
            {/* Faint diagonal watermark - mirrors real invoice */}
            <div style={{position:'absolute', inset:0, display:'flex', alignItems:'center', justifyContent:'center', pointerEvents:'none', zIndex:0}}>
              <div style={{transform:'rotate(-28deg)', fontSize:52, fontWeight:900, fontStyle:'italic', fontFamily:'Georgia,serif', color:clrP, opacity:.06, whiteSpace:'nowrap', letterSpacing:2}}>{F2.name||'Fin-Tax Mitra'}</div>
            </div>
            <table style={{width:'100%', height:'100%', borderCollapse:'collapse', position:'relative', zIndex:1}}>
              <colgroup>
                <col style={{width:'6%'}}/>
                <col style={{width:'54%'}}/>
                <col style={{width:'12%'}}/>
                <col style={{width:'14%'}}/>
                <col style={{width:'14%'}}/>
              </colgroup>
              <thead>
                <tr style={{background:clrA}}>
                  {[['Sr. No.','center'],['Name of Service','center'],['Period / Nos.','center'],['Rate (Rs.)','right'],['Amount (Rs.)','right']].map(([h,align],i)=>(
                    <th key={h} style={{...txtSty('tableHeader',dBFsz,true,true,clrP), padding:'9px 10px', textAlign:align, borderBottom:`1px solid ${clrB}`, borderRight:i<4?`1px solid ${clrB}`:'none', lineHeight:1.3}}>{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {[['1.','Income Tax Return Filling Fees ( A.Y 2025-26 )','1','1,000.00','1,000.00'],['2.','GST Return Filing ( Reimbursement ) ( A.Y 2025-26 )','1','500.00','500.00']].map(([sr,svc,qty,rate,amt])=>(
                <tr key={sr}>
                  <td style={{...txtSty('srNo',dBFsz,false,true,'#000'), padding:'10px', textAlign:'center', borderRight:`1px solid ${clrB}`, verticalAlign:'top'}}>{sr}</td>
                  <td style={{padding:'10px', fontSize:dBFsz, fontStyle:'italic', borderRight:`1px solid ${clrB}`, verticalAlign:'top'}}>{svc}</td>
                  <td style={{padding:'10px', fontSize:dBFsz, textAlign:'center', borderRight:`1px solid ${clrB}`, verticalAlign:'top'}}>{qty}</td>
                  <td style={{padding:'10px', fontSize:dBFsz, textAlign:'right', borderRight:`1px solid ${clrB}`, verticalAlign:'top', fontVariantNumeric:'tabular-nums'}}>{rate}</td>
                  <td style={{padding:'10px', fontSize:dBFsz, textAlign:'right', verticalAlign:'top', fontVariantNumeric:'tabular-nums'}}>{amt}</td>
                </tr>
                ))}
                {/* Empty rows mirror the user's "Empty Service Rows" setting + zebra striping */}
                {Array.from({length:Number(F2.emptyRows??3)}).map((_,i)=><tr key={i} style={{height:28}}><td style={{borderRight:`1px solid ${clrB}`}}/><td style={{borderRight:`1px solid ${clrB}`}}/><td style={{borderRight:`1px solid ${clrB}`}}/><td style={{borderRight:`1px solid ${clrB}`}}/><td/></tr>)}
                {/* Guaranteed spacer row - mirrors the real invoice's always-present
                    filler row, so CGST/SGST/Total below never get stretched. */}
                <tr><td style={{borderRight:`1px solid ${clrB}`}}/><td style={{borderRight:`1px solid ${clrB}`}}/><td style={{borderRight:`1px solid ${clrB}`}}/><td style={{borderRight:`1px solid ${clrB}`}}/><td/></tr>
                {/* FIX: mirror the CGST/SGST rows real GST invoices render, so stamp/signature
                    positions set here don't overlap this section on actual GST invoices */}
                {gstPreview&&<>
                  <tr style={{height:Math.round(30*invScaleD)}}>
                    <td colSpan={3} style={{borderRight:`1px solid ${clrB}`, borderTop:`1px solid ${clrB}`}}/>
                    <td style={{padding:'6px 10px', textAlign:'right', fontStyle:'italic', fontSize:dBFsz-1, borderRight:`1px solid ${clrB}`, borderTop:`1px solid ${clrB}`}}>CGST (9%)</td>
                    <td style={{padding:'6px 10px', textAlign:'right', fontSize:dBFsz-1, borderTop:`1px solid ${clrB}`, fontVariantNumeric:'tabular-nums'}}>67.50</td>
                  </tr>
                  <tr style={{height:Math.round(30*invScaleD)}}>
                    <td colSpan={3} style={{borderRight:`1px solid ${clrB}`}}/>
                    <td style={{padding:'6px 10px', textAlign:'right', fontStyle:'italic', fontSize:dBFsz-1, borderRight:`1px solid ${clrB}`}}>SGST (9%)</td>
                    <td style={{padding:'6px 10px', textAlign:'right', fontSize:dBFsz-1, fontVariantNumeric:'tabular-nums'}}>67.50</td>
                  </tr>
                </>}
                <tr style={{background:F2.clrTotal||clrA, height:Math.round(38*invScaleD)}}>
                  <td colSpan={4} style={{padding:'10px', fontWeight:900, fontStyle:'italic', fontSize:dBFsz+1, textAlign:'center', borderTop:`1px solid ${clrB}`, borderRight:`1px solid ${clrB}`}}>Total</td>
                  <td style={{padding:'10px', fontWeight:900, fontSize:dBFsz+1, textAlign:'right', borderTop:`1px solid ${clrB}`, fontVariantNumeric:'tabular-nums'}}>1,500.00</td>
                </tr>
              </tbody>
            </table>
            </div>
            </div>

            {/* Footer */}
            <div style={{display:'grid', gridTemplateColumns:'55% 45%', borderBottom:`1px solid ${clrB}`, flexShrink:0}}>
              <div style={{borderRight:`1px solid ${clrB}`}}>
                <div style={{padding:'13px 18px', borderBottom:`1px solid ${clrB}`}}>
                  <div style={{display:'flex', justifyContent:'space-between'}}>
                    <div><div style={{...txtSty('totalLabel',dBFsz,true,true,'#000'), marginBottom:6}}>Total in words</div><div style={{fontStyle:'italic', fontSize:dBFsz-1, color:'#333', lineHeight:1.7}}>One Thousand Five Hundred Rupees Only</div></div>
                    <div style={{textAlign:'right'}}><div style={{fontWeight:900, fontStyle:'italic', fontSize:dBFsz, marginBottom:6}}>Total Amount</div><div style={{display:'inline-block', background:clrA, color:clrP, border:`1.5px solid ${clrB}`, fontWeight:900, fontSize:dBFsz+3, padding:'5px 16px', borderRadius:8, fontVariantNumeric:'tabular-nums'}}>₹ 1,500.00</div><div style={{fontSize:dBFsz-2, color:'#777', fontStyle:'italic', marginTop:4}}>(E & O.E.)</div></div>
                  </div>
                </div>
                {/* AUDIT FIX: bank block now mirrors real invoice exactly - no minHeight,
                    same paddings/row spacing/label width, full label names, UPI hint */}
                <div style={{padding:'13px 18px'}}>
                  <div style={{...txtSty('bankLabel',dBFsz,true,true,clrP), marginBottom:8, borderBottom:`1px solid ${clrB}66`, paddingBottom:4}}>Bank Details</div>
                  <div style={{display:'flex', gap:14, alignItems:'flex-start'}}>
                    <div style={{flex:1, minWidth:0}}>
                      {[['Bank Name',F2.bankName||'Kotak Mahindra Bank'],['A/C Holder Name',F2.bankHolder||'BILTU DEY'],['A/C Number',F2.bankAcc||'1449547644'],['IFSC',F2.bankIFSC||'KKBK0000328']].map(([l,v])=>(
                        <div key={l} style={{display:'flex', gap:6, fontSize:dBFsz, marginBottom:5}}><span style={{fontWeight:700, minWidth:110}}>{l} :</span><span>{v}</span></div>
                      ))}
                      <div style={{display:'flex', gap:6, fontSize:dBFsz, marginTop:3}}><span style={{fontWeight:700, minWidth:110}}>UPI ID :</span><span>{F2.upiId||'7319440039@kotak'} <em style={{fontSize:dBFsz-1}}> - Pay using UPI</em></span></div>
                    </div>
                    {/* Fixed QR box - mirrors real invoice */}
                    <div style={{flexShrink:0, textAlign:'center'}}>
                      <div style={{width:88, height:88, border:`1.5px solid ${clrB}`, borderRadius:8, display:'flex', alignItems:'center', justifyContent:'center', background:'#fff', overflow:'hidden'}}>
                        {F2.qrCode
                          ?<img src={F2.qrCode} alt="UPI QR" style={{width:80, height:80, objectFit:'contain', display:'block'}} draggable={false}/>
                          :<span style={{fontSize:9, color:'#999', fontStyle:'italic', textAlign:'center', padding:4}}>QR Code</span>}
                      </div>
                      <div style={{fontSize:Math.max(dBFsz-5,9), fontWeight:700, fontStyle:'italic', marginTop:3, whiteSpace:'nowrap'}}>Scan to pay using UPI</div>
                    </div>
                  </div>
                </div>
              </div>
              {/* AUDIT FIX: right cell now mirrors real invoice exactly - same padding
                  (12px 14px), same lineHeight, same spacer/margins for signature area */}
              <div style={{padding:'12px 14px'}}>
                <div style={{fontSize:dBFsz, fontStyle:'italic', color:'#444', lineHeight:1.6}}>Certified that the particulars given above are true and correct.</div>
                <div style={{fontWeight:900, fontStyle:'italic', fontSize:dBFsz+1, color:clrP, marginTop:8}}>For {F2.name||'Fin-Tax Mitra'}</div>
                <div style={{textAlign:'center', marginTop:16}}>
                  <div style={{minHeight:80}}/>
                  {!F2.signature&&<div style={{width:140, borderBottom:'1px solid #888', margin:'36px auto 0'}}/>}
                  <div style={{...txtSty('signature',dBFsz,true,true,'#000'), marginTop:10}}>Authorised Signatory</div>
                </div>
              </div>
            </div>

            {/* Terms - mirrors real invoice: 8px 16px 12px padding, #444, lineHeight 1.6 */}
            <div style={{padding:'10px 16px 24px', background:clrA, flexShrink:0}}>
              <div style={{fontWeight:900, fontStyle:'italic', fontSize:dBFsz, color:clrP, marginBottom:4}}>Terms and Conditions</div>
              <div style={{...txtSty('termsText',dBFsz-2,false,true,'#444'), lineHeight:1.6}}>{F2.terms||'Fees once paid are non-refundable. The client is responsible for ensuring data accuracy and timely submission. Fin-Tax Mitra is not liable for any penalties or losses resulting from client delays or incorrect data.'}</div>
            </div>
          </div>

          {/* Drag overlay - images float above invoice */}
          <div style={{position:'absolute', inset:0, pointerEvents:'none'}}>
            {sortedImgs.map(img => {
              if(!F2[img.key]) return null;
              const x = F2[img.key+'X'] !== undefined ? F2[img.key+'X'] : img.defX;
              const y = F2[img.key+'Y'] !== undefined ? F2[img.key+'Y'] : img.defY;
              const w = F2[img.key+'W'] || img.defW;
              const h = F2[img.key+'H'] || img.defH;
              const z = F2[img.key+'Z'] || 1;
              const br = F2[img.key+'Brightness'] !== undefined ? F2[img.key+'Brightness'] : 100;
              const co = F2[img.key+'Contrast']   !== undefined ? F2[img.key+'Contrast']   : 100;
              const sa = F2[img.key+'Saturate']   !== undefined ? F2[img.key+'Saturate']   : 100;
              const op = (F2[img.key+'Opacity']   !== undefined ? F2[img.key+'Opacity']    : 100) / 100;
              const isSel = selImg === img.key;
              return (
                <div key={img.key}
                  style={{position:'absolute', left:x, top:y, width:w, height:h, zIndex:10+z,
                    pointerEvents:'all', cursor:'grab', userSelect:'none',
                    outline: isSel ? `2px dashed ${G.green}` : '2px dashed transparent',
                    outlineOffset:2}}
                  onMouseDown={e=>startDrag(e, img.key)}>
                  <img src={F2[img.key]} alt={img.label}
                    style={{width:'100%', height:'100%', objectFit:'contain', display:'block',
                      filter:`brightness(${br}%) contrast(${co}%) saturate(${sa}%)`,
                      opacity:op, draggable:false}}/>
                  {isSel && (
                    <div style={{position:'absolute', top:-18, left:0, background:G.green, color:'#fff',
                      fontSize:9, padding:'1px 6px', borderRadius:4, whiteSpace:'nowrap', fontWeight:700}}>
                      {img.label}
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>

        {/* Instructions + Sync button */}
        <div style={{textAlign:'center', marginTop:12, display:'flex', flexDirection:'column', gap:8, alignItems:'center'}}>
          <div style={{fontSize:11, color:'#555', fontWeight:500}}>
            Click image to select - Drag to move - Use left panel to resize & adjust
          </div>
          <button onClick={()=>setGstPreview(v=>!v)} style={{padding:'6px 14px',borderRadius:8,border:`1px solid ${G.bdr}`,background:gstPreview?G.green+'22':'transparent',color:gstPreview?G.green:G.mut,cursor:'pointer',fontWeight:700,fontSize:11}}>
            {gstPreview?'✓ Previewing WITH GST (CGST/SGST rows shown)':'Previewing WITHOUT GST — click to preview WITH GST'}
          </button>
          <div style={{display:'flex', gap:8, alignItems:'center'}}>
            <button onClick={()=>{
              // Sync: explicitly save current positions to firmSettings so InvoicePrint uses them
              const updates = {};
              IMGS.forEach(img=>{
                updates[img.key+'X'] = F2[img.key+'X'] !== undefined ? F2[img.key+'X'] : img.defX;
                updates[img.key+'Y'] = F2[img.key+'Y'] !== undefined ? F2[img.key+'Y'] : img.defY;
                updates[img.key+'W'] = F2[img.key+'W'] || img.defW;
                updates[img.key+'H'] = F2[img.key+'H'] || img.defH;
                if(F2[img.key+'Z'] !== undefined) updates[img.key+'Z'] = F2[img.key+'Z'];
                if(F2[img.key+'Brightness'] !== undefined) updates[img.key+'Brightness'] = F2[img.key+'Brightness'];
                if(F2[img.key+'Contrast'] !== undefined) updates[img.key+'Contrast'] = F2[img.key+'Contrast'];
                if(F2[img.key+'Saturate'] !== undefined) updates[img.key+'Saturate'] = F2[img.key+'Saturate'];
                if(F2[img.key+'Opacity'] !== undefined) updates[img.key+'Opacity'] = F2[img.key+'Opacity'];
                if(F2[img.key+'Fit'] !== undefined) updates[img.key+'Fit'] = F2[img.key+'Fit'];
              });
              setFirmSettings(p=>({...p,...updates,_synced:Date.now()}));
              // positions saved to firmSettings
            }} style={{padding:'8px 22px', borderRadius:10, border:'none', cursor:'pointer',
              background:`linear-gradient(135deg,${G.g2},${G.green})`, color:'#fff',
              fontWeight:700, fontSize:13}}>
              Sync to Invoice Print
            </button>
            <span style={{fontSize:11, color:F2._synced?G.green:G.bdr, fontWeight:600}}>
              {F2._synced ? 'Synced ' + new Date(F2._synced).toLocaleTimeString('en-IN',{hour:'2-digit',minute:'2-digit'}) : 'Not synced yet'}
            </span>
          </div>
        </div>
      </div>
    </div>
  );
}


function DevTab({dd,setDd,pws,setPws,darkMode,setDarkMode,profilePic,setProfilePic,firmSettings,setFirmSettings,clients,setClients,works,setWorks,invoices,setInvoices,receipts,setReceipts,toast,googleUser,linkGoogleDrive,disconnectGoogleDrive,uploadBackupToGoogleDrive}){
  const[sec,setSec]=useState("services"),[nv,setNv]=useState(""),[ei,setEi]=useState(null),[ev,setEv]=useState("");
  const[pf,setPf]=useState({owner:"",dev:""}),[pm,setPm]=useState("");
  const[confirmDel,setConfirmDel]=useState(null);
  const[confirmRestore,setConfirmRestore]=useState(null);
  const[bkMsg,setBkMsg]=useState("");
  const LISTS=[{k:"services",l:"Services",icon:"📋",cat:"Work"},{k:"staff",l:"Staff Members",icon:"👤",cat:"Work"},{k:"billingItems",l:"Billing Items",icon:"🧾",cat:"Work"},{k:"sources",l:"Lead Sources",icon:"📣",cat:"Client"},{k:"clientTypes",l:"Client Types",icon:"🏢",cat:"Client"},{k:"pwTypes",l:"Password Types",icon:"🔑",cat:"Client"},{k:"fyOptions",l:"Financial Years",icon:"📅",cat:"General"},{k:"states",l:"States",icon:"🗺",cat:"General"}];
  const cur=dd[sec]||[];
  const cm=LISTS.find(l=>l.k===sec);
  const add=()=>{const v=nv.trim();if(!v)return;if(cur.includes(v)){setConfirmDel({type:"dup",msg:v+" already exists in this list!"});return;}setDd(p=>({...p,[sec]:[...(p[sec]||[]),v]}));setNv("");};
  const del=i=>setConfirmDel({type:"del",idx:i,label:cur[i]});
  const doDelete=i=>{setDd(p=>({...p,[sec]:(p[sec]||[]).filter((_,j)=>j!==i)}));setConfirmDel(null);};
  const saveE=i=>{const v=ev.trim();if(!v)return;setDd(p=>({...p,[sec]:(p[sec]||[]).map((x,j)=>j===i?v:x)}));setEi(null);setEv("");};
  const mu=i=>{if(i===0)return;setDd(p=>{const a=[...(p[sec]||[])];[a[i-1],a[i]]=[a[i],a[i-1]];return{...p,[sec]:a};});};
  const md=i=>{if(i>=cur.length-1)return;setDd(p=>{const a=[...(p[sec]||[])];[a[i],a[i+1]]=[a[i+1],a[i]];return{...p,[sec]:a};});};
  const savePw=()=>{const np={...pws};if(pf.owner.trim())np.owner=pf.owner.trim();if(pf.dev.trim())np.dev=pf.dev.trim();setPws(np);setPf({owner:"",dev:""});setPm("Passwords updated!");setTimeout(()=>setPm(""),3000);};
  const[devSideOpen,setDevSideOpen]=useState(true);

  // ─── Backup & Restore ──────────────────────────────────────────────────
  const stamp=()=>{const d=new Date();const p=n=>String(n).padStart(2,"0");return `${d.getFullYear()}-${p(d.getMonth()+1)}-${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}`;};
  const buildBackup=()=>({
    meta:{app:"Fin-Tax Mitra",exportedAt:new Date().toISOString(),version:1},
    clients, works, invoices, receipts, dd, pws, firmSettings, darkMode,
  });
  const downloadBlob=(data,filename,type)=>{
    const blob=new Blob([data],{type});
    const url=URL.createObjectURL(blob);
    const a=document.createElement("a");
    a.href=url; a.download=filename; document.body.appendChild(a); a.click();
    document.body.removeChild(a); URL.revokeObjectURL(url);
  };
  const exportJSON=()=>{
    try{
      const data=JSON.stringify(buildBackup(),null,2);
      downloadBlob(data,`FinTaxMitra_Backup_${stamp()}.json`,"application/json");
      setBkMsg("✓ JSON backup downloaded!"); toast&&toast("Backup exported (JSON)");
    }catch(e){ setBkMsg("✗ Export failed: "+e.message); }
    setTimeout(()=>setBkMsg(""),3500);
  };
  const exportExcel=()=>{
    try{
      const wb=XLSX.utils.book_new();
      const PK=PORTALS.map(p=>p.key);

      // Clients sheet - portals flattened into columns
      const clientRows=clients.map(c=>{
        const row={pan:c.pan,name:c.name,biz:c.biz,mob:c.mob,email:c.email,gstin:c.gstin,addr:c.addr,state:c.state,type:c.type,src:c.src,added:c.added,status:c.status,note:c.note};
        PK.forEach(k=>{const p=(c.portals&&c.portals[k])||{}; row[`${k}_id`]=p.id||""; row[`${k}_pw`]=p.pw||""; row[`${k}_active`]=p.on?"Yes":"No";});
        return row;
      });
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(clientRows),"Clients");

      // Work Tracker sheet
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(works||[]),"WorkTracker");

      // Invoices sheet (line items flattened separately + json snapshot column)
      const invRows=(invoices||[]).map(inv=>({...inv,items:inv.items?JSON.stringify(inv.items):""}));
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(invRows),"Invoices");
      const invItemRows=[];
      (invoices||[]).forEach(inv=>(inv.items||[]).forEach((it,i)=>invItemRows.push({invId:inv.id,lineNo:i+1,service:it.service||"",desc:it.desc||"",type:it.type||"service",qty:it.qty||"",rate:it.rate||""})));
      if(invItemRows.length) XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(invItemRows),"InvoiceItems");

      // Receipts sheet
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(receipts||[]),"Receipts");

      // Dropdowns sheet (columns of unequal length, padded)
      const ddKeys=Object.keys(dd||{});
      const maxLen=Math.max(0,...ddKeys.map(k=>(dd[k]||[]).length));
      const ddRows=[]; for(let i=0;i<maxLen;i++){const row={}; ddKeys.forEach(k=>row[k]=(dd[k]||[])[i]||""); ddRows.push(row);}
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(ddRows),"Dropdowns");

      // Settings sheet - passwords + firm settings + theme
      const settingsRow=[{
        ownerPassword:pws.owner, devPassword:pws.dev, darkMode:darkMode?"Dark":"Light",
        firmName:firmSettings.name, firmAddress:firmSettings.addr, firmPhone:firmSettings.phone, firmEmail:firmSettings.email,
        bankName:firmSettings.bankName, bankHolder:firmSettings.bankHolder, bankAcc:firmSettings.bankAcc, bankIFSC:firmSettings.bankIFSC, upiId:firmSettings.upiId,
        terms:firmSettings.terms, exportedAt:new Date().toISOString(),
      }];
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(settingsRow),"Settings");

      XLSX.writeFile(wb,`FinTaxMitra_Backup_${stamp()}.xlsx`);
      setBkMsg("✓ Excel backup downloaded!"); toast&&toast("Backup exported (Excel)");
    }catch(e){ setBkMsg("✗ Export failed: "+e.message); }
    setTimeout(()=>setBkMsg(""),3500);
  };
  const handleImportFile=(e)=>{
    const file=e.target.files[0]; if(!file) return;
    const reader=new FileReader();
    reader.onload=ev=>{
      try{
        const data=JSON.parse(ev.target.result);
        if(!data||typeof data!=="object") throw new Error("Invalid backup file");
        setConfirmRestore(data);
      }catch(err){ setBkMsg("✗ Invalid JSON backup file: "+err.message); setTimeout(()=>setBkMsg(""),4000); }
    };
    reader.readAsText(file); e.target.value="";
  };
  const doRestore=()=>{
    const data=confirmRestore; if(!data) return;
    if(Array.isArray(data.clients)) setClients(data.clients);
    if(Array.isArray(data.works)) setWorks(data.works);
    if(Array.isArray(data.invoices)) setInvoices(data.invoices);
    if(Array.isArray(data.receipts)) setReceipts(data.receipts);
    if(data.dd&&typeof data.dd==="object") setDd(data.dd);
    if(data.pws&&typeof data.pws==="object") setPws(data.pws);
    if(data.firmSettings&&typeof data.firmSettings==="object") setFirmSettings(data.firmSettings);
    if(typeof data.darkMode==="boolean") setDarkMode(data.darkMode);
    setConfirmRestore(null);
    setBkMsg("✓ Backup restored successfully!"); toast&&toast("Backup restored");
    setTimeout(()=>setBkMsg(""),3500);
  };
  return <div style={{display:"grid",gridTemplateColumns:(devSideOpen?220:52)+"px 1fr",gap:14,transition:"grid-template-columns .2s"}}>
    <div style={{display:"flex",flexDirection:"column",gap:4,overflow:"hidden"}}>
      {/* Dev sidebar header with collapse toggle */}
      <div style={{display:"flex",alignItems:"center",gap:8,padding:"8px 10px",background:G.green+"0E",border:`1px solid ${G.green}30`,borderRadius:11,marginBottom:4,flexShrink:0}}>
        {devSideOpen&&<div style={{flex:1,minWidth:0}}>
          <div style={{fontWeight:700,fontSize:12,color:G.green,whiteSpace:"nowrap"}}>⚙️ Developer Settings</div>
          <div style={{color:G.mut,fontSize:10,marginTop:2}}>Manage dropdowns & passwords.</div>
        </div>}
        {!devSideOpen&&<span style={{fontSize:16,margin:"0 auto"}}>⚙️</span>}
        <button onClick={()=>setDevSideOpen(p=>!p)} title={devSideOpen?"Collapse":"Expand"}
          style={{background:"transparent",border:`1px solid ${G.green}44`,color:G.green,borderRadius:7,width:24,height:24,cursor:"pointer",fontSize:13,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,transition:"transform .2s",transform:devSideOpen?"none":"rotate(180deg)"}}>{String.fromCharCode(60)}</button>
      </div>
      {["Work","Client","General"].map(cat=><div key={cat}>
        {devSideOpen&&<div style={{fontSize:10,color:G.bdr,fontWeight:700,letterSpacing:.9,padding:"8px 7px 3px",textTransform:"uppercase"}}>{cat}</div>}
        {!devSideOpen&&<div style={{height:6}}/>}
        {LISTS.filter(l=>l.cat===cat).map(l=><div key={l.k} style={{position:"relative"}} title={!devSideOpen?l.l+` (${dd[l.k]?.length})`:""}>
          <button onClick={()=>{setSec(l.k);setEi(null);setNv("");}}
            style={{display:"flex",alignItems:"center",gap:devSideOpen?9:0,justifyContent:devSideOpen?"flex-start":"center",padding:devSideOpen?"7px 10px":"0",width:"100%",height:36,borderRadius:8,border:"none",cursor:"pointer",textAlign:"left",marginBottom:2,background:sec===l.k?`linear-gradient(90deg,${G.g2},${G.green})`:"transparent",color:sec===l.k?"#fff":G.mut,fontWeight:sec===l.k?700:500,fontSize:13,transition:"all .15s"}}
            onMouseEnter={e=>{if(sec!==l.k)e.currentTarget.style.background=G.bdr;}} onMouseLeave={e=>{if(sec!==l.k)e.currentTarget.style.background="transparent";}}>
            <span style={{fontSize:15,width:20,textAlign:"center",flexShrink:0}}>{l.icon}</span>
            {devSideOpen&&<><span style={{flex:1,fontSize:12,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{l.l}</span>
            <span style={{fontSize:10,opacity:.7,background:sec===l.k?"#fff3":G.bdr,padding:"1px 5px",borderRadius:8,flexShrink:0}}>{dd[l.k]?.length}</span></>}
          </button>
          {!devSideOpen&&sec===l.k&&<div style={{position:"absolute",left:0,top:0,bottom:0,width:3,background:G.green,borderRadius:3}}/>}
        </div>)}
      </div>)}
      {devSideOpen&&<div style={{fontSize:10,color:G.bdr,fontWeight:700,letterSpacing:.9,padding:"8px 7px 3px",textTransform:"uppercase"}}>Appearance</div>}
      {!devSideOpen&&<div style={{height:6}}/>}
      {[
        {id:"theme",  icon:darkMode?"🌙":"☀️", label:"Dark / Light Mode",      col:G.amb},
        {id:"profile",icon:"🖼",              label:"Profile Picture",         col:G.cyn},
        {id:"firm",   icon:"🏢",              label:"Firm / Invoice Settings", col:G.amb},
        {id:"performa",icon:"🎨",             label:"Performa Designer",       col:G.cyn},
      ].map(btn=><div key={btn.id} style={{position:"relative"}} title={!devSideOpen?btn.label:""}>
        <button onClick={()=>setSec(btn.id)}
          style={{display:"flex",alignItems:"center",gap:devSideOpen?9:0,justifyContent:devSideOpen?"flex-start":"center",padding:devSideOpen?"7px 10px":"0",width:"100%",height:36,borderRadius:8,border:"none",cursor:"pointer",textAlign:"left",marginBottom:2,background:sec===btn.id?`linear-gradient(90deg,${G.g2},${G.green})`:"transparent",color:sec===btn.id?"#fff":btn.col,fontWeight:sec===btn.id?700:600,fontSize:13,transition:"all .15s"}}
          onMouseEnter={e=>{if(sec!==btn.id)e.currentTarget.style.background=G.bdr;}} onMouseLeave={e=>{if(sec!==btn.id)e.currentTarget.style.background="transparent";}}>
          <span style={{fontSize:15,width:20,textAlign:"center",flexShrink:0}}>{btn.icon}</span>
          {devSideOpen&&<span style={{flex:1,fontSize:12,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{btn.label}</span>}
        </button>
        {!devSideOpen&&sec===btn.id&&<div style={{position:"absolute",left:0,top:0,bottom:0,width:3,background:G.green,borderRadius:3}}/>}
      </div>)}
      {devSideOpen&&<div style={{fontSize:10,color:G.bdr,fontWeight:700,letterSpacing:.9,padding:"8px 7px 3px",textTransform:"uppercase"}}>Security</div>}
      {!devSideOpen&&<div style={{height:6}}/>}
      <div style={{position:"relative"}} title={!devSideOpen?"Change Passwords":""}>
        <button onClick={()=>setSec("passwords")}
          style={{display:"flex",alignItems:"center",gap:devSideOpen?9:0,justifyContent:devSideOpen?"flex-start":"center",padding:devSideOpen?"7px 10px":"0",width:"100%",height:36,borderRadius:8,border:"none",cursor:"pointer",textAlign:"left",background:sec==="passwords"?`linear-gradient(90deg,${G.g2},${G.green})`:"transparent",color:sec==="passwords"?"#fff":G.cyn,fontWeight:sec==="passwords"?700:600,fontSize:13,transition:"all .15s"}}
          onMouseEnter={e=>{if(sec!=="passwords")e.currentTarget.style.background=G.bdr;}} onMouseLeave={e=>{if(sec!=="passwords")e.currentTarget.style.background="transparent";}}>
          <span style={{fontSize:15,width:20,textAlign:"center",flexShrink:0}}>🔑</span>
          {devSideOpen&&<span style={{flex:1,fontSize:12,whiteSpace:"nowrap"}}>Change Passwords</span>}
        </button>
        {!devSideOpen&&sec==="passwords"&&<div style={{position:"absolute",left:0,top:0,bottom:0,width:3,background:G.green,borderRadius:3}}/>}
      </div>
      {devSideOpen&&<div style={{fontSize:10,color:G.bdr,fontWeight:700,letterSpacing:.9,padding:"8px 7px 3px",textTransform:"uppercase"}}>Data</div>}
      {!devSideOpen&&<div style={{height:6}}/>}
      <div style={{position:"relative"}} title={!devSideOpen?"Backup & Restore":""}>
        <button onClick={()=>setSec("backup")}
          style={{display:"flex",alignItems:"center",gap:devSideOpen?9:0,justifyContent:devSideOpen?"flex-start":"center",padding:devSideOpen?"7px 10px":"0",width:"100%",height:36,borderRadius:8,border:"none",cursor:"pointer",textAlign:"left",background:sec==="backup"?`linear-gradient(90deg,${G.g2},${G.green})`:"transparent",color:sec==="backup"?"#fff":G.amb,fontWeight:sec==="backup"?700:600,fontSize:13,transition:"all .15s"}}
          onMouseEnter={e=>{if(sec!=="backup")e.currentTarget.style.background=G.bdr;}} onMouseLeave={e=>{if(sec!=="backup")e.currentTarget.style.background="transparent";}}>
          <span style={{fontSize:15,width:20,textAlign:"center",flexShrink:0}}>💾</span>
          {devSideOpen&&<span style={{flex:1,fontSize:12,whiteSpace:"nowrap"}}>Backup & Restore</span>}
        </button>
        {!devSideOpen&&sec==="backup"&&<div style={{position:"absolute",left:0,top:0,bottom:0,width:3,background:G.green,borderRadius:3}}/>}
      </div>
    </div>
    <div>
    {confirmRestore&&<div style={{position:"fixed",inset:0,background:"#000C",zIndex:6000,display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:16,padding:28,width:380,boxShadow:"0 20px 60px #000C",textAlign:"center"}}>
        <div style={{fontSize:36,marginBottom:12}}>⚠️</div>
        <div style={{fontWeight:700,fontSize:16,color:G.wh,marginBottom:8}}>Restore this backup?</div>
        <div style={{fontSize:13,color:G.mut,marginBottom:6,padding:"8px 14px",background:G.amb+"10",border:`1px solid ${G.amb}33`,borderRadius:9,textAlign:"left"}}>
          Clients: <b>{Array.isArray(confirmRestore.clients)?confirmRestore.clients.length:"—"}</b><br/>
          Work items: <b>{Array.isArray(confirmRestore.works)?confirmRestore.works.length:"—"}</b><br/>
          Invoices: <b>{Array.isArray(confirmRestore.invoices)?confirmRestore.invoices.length:"—"}</b><br/>
          Receipts: <b>{Array.isArray(confirmRestore.receipts)?confirmRestore.receipts.length:"—"}</b><br/>
          {confirmRestore.meta?.exportedAt&&<>Exported: <b>{new Date(confirmRestore.meta.exportedAt).toLocaleString("en-IN")}</b></>}
        </div>
        <div style={{fontSize:12,color:G.red,marginBottom:18}}>This will overwrite ALL current data in the app. This cannot be undone.</div>
        <div style={{display:"flex",gap:10}}>
          <button onClick={doRestore} style={{flex:1,padding:"11px",borderRadius:10,border:"none",cursor:"pointer",background:G.red,color:"#fff",fontWeight:700,fontSize:14}}>Yes, Restore</button>
          <button onClick={()=>setConfirmRestore(null)} style={{flex:1,padding:"11px",borderRadius:10,border:`1px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer",fontWeight:600}}>Cancel</button>
        </div>
      </div>
    </div>}
    {confirmDel&&<div style={{position:"fixed",inset:0,background:"#000C",zIndex:6000,display:"flex",alignItems:"center",justifyContent:"center"}}>
      <div style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:16,padding:28,width:360,boxShadow:"0 20px 60px #000C",textAlign:"center"}}>
        {confirmDel.type==="del"
          ?<><div style={{fontSize:36,marginBottom:12}}>🗑</div>
            <div style={{fontWeight:700,fontSize:16,color:G.wh,marginBottom:8}}>Remove this item?</div>
            <div style={{fontSize:13,color:G.mut,marginBottom:20,padding:"8px 14px",background:G.red+"10",border:`1px solid ${G.red}33`,borderRadius:9}}>"{confirmDel.label}"</div>
            <div style={{display:"flex",gap:10}}>
              <button onClick={()=>doDelete(confirmDel.idx)} style={{flex:1,padding:"11px",borderRadius:10,border:"none",cursor:"pointer",background:G.red,color:"#fff",fontWeight:700,fontSize:14}}>Yes, Delete</button>
              <button onClick={()=>setConfirmDel(null)} style={{flex:1,padding:"11px",borderRadius:10,border:`1px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer",fontWeight:600}}>Cancel</button>
            </div></>
          :<><div style={{fontSize:36,marginBottom:12}}>⚠️</div>
            <div style={{fontWeight:700,fontSize:15,color:G.wh,marginBottom:8}}>Duplicate Item</div>
            <div style={{fontSize:13,color:G.amb,marginBottom:20}}>{confirmDel.msg}</div>
            <button onClick={()=>setConfirmDel(null)} style={{padding:"10px 28px",borderRadius:10,border:"none",cursor:"pointer",background:`linear-gradient(135deg,${G.g2},${G.green})`,color:"#fff",fontWeight:700,fontSize:13}}>OK</button>
          </>}
      </div>
    </div>}
    {sec==="backup"
    ?<Crd><SH icon="💾" title="Backup & Restore" acc={G.amb}/>
      <div style={{marginTop:14,display:"flex",flexDirection:"column",gap:16}}>
        <div style={{padding:"10px 14px",background:G.green+"10",border:`1px solid ${G.green}33`,borderRadius:9,fontSize:12,color:G.mut,lineHeight:1.7}}>
          Export a full backup of all app data - clients, work tracker, invoices, receipts, dropdown lists, passwords and firm settings - exactly as stored. Keep backups somewhere safe; anyone with the file can see saved portal passwords.
        </div>

        {/* Auto Backup Toggle */}
        <div style={{background:G.card,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"14px 16px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
          <div style={{flex:1,marginRight:16}}>
            <div style={{fontSize:12,fontWeight:700,color:G.txt,display:"flex",alignItems:"center",gap:8}}><span style={{background:G.green+"20",color:G.green,width:24,height:24,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13}}>🔄</span> Daily Auto-Backup</div>
            <div style={{fontSize:11,color:G.mut,marginTop:4}}>Downloads a JSON backup file to your computer once every 24 hours automatically on page launch.</div>
          </div>
          <div onClick={()=>setFirmSettings(p=>({...p,autoBackup:!p.autoBackup}))}
            style={{width:44,height:24,borderRadius:12,background:firmSettings.autoBackup?G.green:G.bdr,cursor:"pointer",position:"relative",transition:"background .2s",flexShrink:0}}>
            <div style={{position:"absolute",top:2,width:20,height:20,borderRadius:"50%",background:"#fff",transition:"left .2s",left:firmSettings.autoBackup?"22px":"2px"}}/>
          </div>
        </div>

        {/* Google Drive Auto-Backup Card */}
        <div style={{background:G.card,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"14px 16px"}}>
          <div style={{display:"flex",alignItems:"center",justifyContent:"space-between",marginBottom:12}}>
            <div>
              <div style={{fontSize:12,fontWeight:700,color:G.txt,display:"flex",alignItems:"center",gap:8}}><span style={{background:"#4285F420",color:"#4285F4",width:24,height:24,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13}}>☁️</span> Google Drive Sync</div>
              <div style={{fontSize:11,color:G.mut,marginTop:4}}>Automatically uploads JSON backups to Google Drive.</div>
            </div>
            <div onClick={()=>setFirmSettings(p=>({...p,googleDriveEnabled:!p.googleDriveEnabled}))}
              style={{width:44,height:24,borderRadius:12,background:firmSettings.googleDriveEnabled?G.green:G.bdr,cursor:"pointer",position:"relative",transition:"background .2s",flexShrink:0}}>
              <div style={{position:"absolute",top:2,width:20,height:20,borderRadius:"50%",background:"#fff",transition:"left .2s",left:firmSettings.googleDriveEnabled?"22px":"2px"}}/>
            </div>
          </div>

          <div style={{display:"flex",flexDirection:"column",gap:10,marginTop:12,paddingTop:12,borderTop:`1px dashed ${G.bdr}`}}>
            <div>
              <label style={{fontSize:10,fontWeight:700,color:G.mut,textTransform:"uppercase",display:"block",marginBottom:4}}>Google OAuth Client ID</label>
              <input 
                type="text" 
                value={firmSettings.googleClientId || ""} 
                onChange={e=>setFirmSettings(p=>({...p,googleClientId:e.target.value}))} 
                placeholder="Enter Google Client ID" 
                style={{
                  background: G.surf, 
                  border: `1.5px solid ${G.bdr}`, 
                  borderRadius: 8, 
                  color: G.wh, 
                  fontSize: 12, 
                  padding: "8px 11px", 
                  outline: "none", 
                  width: "100%", 
                  boxSizing: "border-box", 
                  fontFamily: "monospace"
                }}
              />
            </div>

            <div>
              <label style={{fontSize:10,fontWeight:700,color:G.mut,textTransform:"uppercase",display:"block",marginBottom:4}}>Google OAuth Client Secret (Optional - Enable Permanent Silent Refresh)</label>
              <input 
                type="password" 
                value={firmSettings.googleClientSecret || ""} 
                onChange={e=>setFirmSettings(p=>({...p,googleClientSecret:e.target.value}))} 
                placeholder="Enter Client Secret for silent background backups" 
                style={{
                  background: G.surf, 
                  border: `1.5px solid ${G.bdr}`, 
                  borderRadius: 8, 
                  color: G.wh, 
                  fontSize: 12, 
                  padding: "8px 11px", 
                  outline: "none", 
                  width: "100%", 
                  boxSizing: "border-box", 
                  fontFamily: "monospace"
                }}
              />
            </div>

            <div style={{display:"flex",alignItems:"center",gap:12,flexWrap:"wrap",marginTop:4}}>
              {googleUser ? (
                <>
                  <div style={{flex:1,minWidth:180}}>
                    <span style={{fontSize:11,color:G.mut}}>Status: </span>
                    <span style={{fontSize:11,color:G.green,fontWeight:700}}>Connected</span>
                    <div style={{fontSize:10,color:G.txt,fontFamily:"monospace",marginTop:2,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{googleUser.email}</div>
                  </div>
                  <button onClick={disconnectGoogleDrive} style={{padding:"8px 12px",borderRadius:8,border:`1px solid ${G.red}33`,background:"transparent",color:G.red,fontSize:11,fontWeight:700,cursor:"pointer"}}>Disconnect</button>
                  <button onClick={()=>uploadBackupToGoogleDrive(true)} style={{padding:"8px 12px",borderRadius:8,border:"none",background:`linear-gradient(135deg,#3B82F6,#2563EB)`,color:"#fff",fontSize:11,fontWeight:700,cursor:"pointer"}}>Upload Backup Now</button>
                </>
              ) : (
                <>
                  <div style={{flex:1,minWidth:180}}>
                    <span style={{fontSize:11,color:G.mut}}>Status: </span>
                    <span style={{fontSize:11,color:G.red,fontWeight:700}}>Not Linked</span>
                  </div>
                  <button onClick={linkGoogleDrive} style={{padding:"8px 14px",borderRadius:8,border:"none",background:`linear-gradient(135deg,#3B82F6,#2563EB)`,color:"#fff",fontSize:11,fontWeight:700,cursor:"pointer"}}>🔗 Link Google Drive</button>
                </>
              )}
            </div>
            
            <div style={{fontSize:10,color:G.mut,lineHeight:1.5,marginTop:6,background:G.surf,padding:"8px 12px",borderRadius:8,border:`1px solid ${G.bdr}`}}>
              🔑 Note: Google OAuth only requests permission to write files created by this app (Drive scope: `drive.file`). Your other files are 100% private and protected.
            </div>
          </div>
        </div>

        {/* Export */}
        <div style={{background:G.card,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"14px 16px"}}>
          <div style={{fontSize:12,fontWeight:700,color:G.txt,marginBottom:12,display:"flex",alignItems:"center",gap:8}}><span style={{background:G.green+"20",color:G.green,width:24,height:24,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13}}>⬇️</span> Export Data</div>
          <div style={{display:"flex",gap:12,flexWrap:"wrap"}}>
            <Btn onClick={exportJSON} sty={{padding:"11px 20px",fontSize:13,flex:"1 1 200px"}}>📄 Export as JSON</Btn>
            <Btn onClick={exportExcel} sty={{padding:"11px 20px",fontSize:13,flex:"1 1 200px"}} color={G.cyn} color2={G.ind}>📊 Export as Excel</Btn>
          </div>
          <div style={{fontSize:11,color:G.bdr,marginTop:10,lineHeight:1.6}}>
            <div>* JSON - complete raw backup, best for restoring into this app later.</div>
            <div>* Excel - all records laid out in sheets (Clients, WorkTracker, Invoices, InvoiceItems, Receipts, Dropdowns, Settings) for viewing/editing in Excel/Sheets.</div>
          </div>
        </div>

        {/* Import */}
        <div style={{background:G.card,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"14px 16px"}}>
          <div style={{fontSize:12,fontWeight:700,color:G.txt,marginBottom:12,display:"flex",alignItems:"center",gap:8}}><span style={{background:G.red+"20",color:G.red,width:24,height:24,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13}}>⬆️</span> Restore from Backup</div>
          <input id="backup_restore_file" type="file" accept="application/json,.json" onChange={handleImportFile} style={{display:"none"}}/>
          <Btn onClick={()=>document.getElementById("backup_restore_file").click()} sty={{padding:"11px 20px",fontSize:13}} color={G.red} color2="#B91C1C">📤 Choose JSON Backup File</Btn>
          <div style={{fontSize:11,color:G.amb,marginTop:10,lineHeight:1.6}}>⚠️ Only JSON backups exported from this app can be restored. Restoring replaces all current data - export a fresh backup first if unsure.</div>
        </div>

        {bkMsg&&<div style={{padding:"9px 13px",background:(bkMsg[0]==="✓"?G.green:G.red)+"15",border:`1px solid ${(bkMsg[0]==="✓"?G.green:G.red)}44`,borderRadius:9,fontSize:13,color:bkMsg[0]==="✓"?G.green:G.red,fontWeight:600}}>{bkMsg}</div>}
      </div>
    </Crd>
    :sec==="performa"
    ?<PerformaDesigner firmSettings={firmSettings} setFirmSettings={setFirmSettings}/>
    :sec==="firm"
    ?<Crd><SH icon="🏢" title="Firm & Invoice Settings" acc={G.amb}/>
      <div style={{marginTop:14,display:"flex",flexDirection:"column",gap:14}}>
        {/* Upload helpers */}
        {(()=>{
          const uploadImg=(key,maxW,maxH)=>e=>{
            const file=e.target.files[0]; if(!file) return;
            if(file.size>5*1024*1024){alert("Max 5MB");return;}
            const reader=new FileReader();
            reader.onload=ev=>setFirmSettings(p=>({...p,[key]:ev.target.result}));
            reader.readAsDataURL(file); e.target.value="";
          };
          const FS=firmSettings;
          const setF=k=>v=>setFirmSettings(p=>({...p,[k]:v}));
          return <>
            {/* Firm Identity */}
            <div style={{background:G.card,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"14px 16px"}}>
              <div style={{fontSize:12,fontWeight:700,color:G.txt,marginBottom:12,display:"flex",alignItems:"center",gap:8}}><span style={{background:G.green+"20",color:G.green,width:24,height:24,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13}}>🏢</span> Firm Identity</div>
              <div style={{display:"flex",flexDirection:"column",gap:10}}>
                <F label="Firm / Company Name"><I val={FS.name||""} set={setF("name")} ph="Fin-Tax Mitra"/></F>
                <F label="Address"><textarea value={FS.addr||""} onChange={e=>setFirmSettings(p=>({...p,addr:e.target.value}))} rows={2} placeholder="Full address..." style={{...IS,resize:"vertical"}}/></F>
                <div style={{display:"flex",gap:12}}>
                  <F label="Phone" w="calc(50% - 6px)"><I val={FS.phone||""} set={setF("phone")} ph="Phone number"/></F>
                  <F label="Email" w="calc(50% - 6px)"><I val={FS.email||""} set={setF("email")} ph="Email address"/></F>
                </div>
              </div>
            </div>

            {/* Bank Details */}
            <div style={{background:G.card,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"14px 16px"}}>
              <div style={{fontSize:12,fontWeight:700,color:G.txt,marginBottom:12,display:"flex",alignItems:"center",gap:8}}><span style={{background:G.cyn+"20",color:G.cyn,width:24,height:24,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13}}>🏦</span> Bank Details</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                <F label="Bank Name"><I val={FS.bankName||""} set={setF("bankName")} ph="Bank name"/></F>
                <F label="A/C Holder Name"><I val={FS.bankHolder||""} set={setF("bankHolder")} ph="Account holder"/></F>
                <F label="Account Number"><I val={FS.bankAcc||""} set={setF("bankAcc")} ph="Account number" mono/></F>
                <F label="IFSC Code"><I val={FS.bankIFSC||""} set={setF("bankIFSC")} ph="IFSC" mono/></F>
                <F label="UPI ID" w="100%"><I val={FS.upiId||""} set={setF("upiId")} ph="yourname@upi" mono/></F>
              </div>
            </div>

            {/* Logo / Stamp / Signature uploads */}
            <div style={{background:G.card,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"14px 16px"}}>
              <div style={{fontSize:12,fontWeight:700,color:G.txt,marginBottom:12,display:"flex",alignItems:"center",gap:8}}><span style={{background:G.vio+"20",color:G.vio,width:24,height:24,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13}}>🖼</span> Invoice Images</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr 1fr",gap:12}}>
                {[
                  {key:"logo",        label:"Company Logo",      hint:"Top-left header",        icon:"🏢", defW:80,  defH:80},
                  {key:"stamp",       label:"Official Stamp",     hint:"Signature area",          icon:"🔏", defW:70,  defH:70},
                  {key:"signature",   label:"Signature",          hint:"Authorised signatory",    icon:"✍️", defW:140, defH:45},
                  {key:"qrCode",      label:"Payment QR Code",    hint:"UPI/Bank QR on invoice",  icon:"📱", defW:80,  defH:80},
                  {key:"statusStamp", label:"Universal Stamp",    hint:"PAID / UNPAID / custom",  icon:"🔖", defW:100, defH:100},
                ].map(item=>(
                  <div key={item.key} style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:11,padding:"12px 14px",display:"flex",flexDirection:"column",gap:8}}>
                    <div style={{fontSize:11,fontWeight:700,color:G.txt,display:"flex",alignItems:"center",gap:6}}><span>{item.icon}</span>{item.label}</div>
                    {/* Upload box */}
                    <div style={{height:100,border:`2px dashed ${FS[item.key]?G.green:G.bdr}`,borderRadius:9,display:"flex",alignItems:"center",justifyContent:"center",cursor:"pointer",background:FS[item.key]?G.green+"06":G.card,overflow:"hidden",transition:"border .2s"}}
                      onClick={()=>document.getElementById("firm_"+item.key).click()}>
                      {FS[item.key]
                        ?<img src={FS[item.key]} alt={item.label} style={{width:"100%",height:"100%",objectFit:"contain",padding:6}}/>
                        :<div style={{textAlign:"center",padding:8}}>
                          <div style={{fontSize:28,marginBottom:5}}>📁</div>
                          <div style={{fontSize:11,color:G.mut,fontWeight:600}}>Click to upload</div>
                          <div style={{fontSize:10,color:G.bdr,marginTop:2}}>PNG / JPG / SVG</div>
                        </div>}
                      <input id={"firm_"+item.key} type="file" accept="image/*" style={{display:"none"}} onChange={uploadImg(item.key)}/>
                    </div>
                    {/* Size controls */}
                    {FS[item.key]&&<>
                      <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:6}}>
                        <div>
                          <div style={{fontSize:10,color:G.mut,fontWeight:600,marginBottom:3}}>Width (px)</div>
                          <input type="number" value={FS[item.key+"W"]||item.defW} min={20} max={400}
                            onChange={e=>setFirmSettings(p=>({...p,[item.key+"W"]:Number(e.target.value)}))}
                            style={{...IS,padding:"4px 8px",fontSize:12}}/>
                        </div>
                        <div>
                          <div style={{fontSize:10,color:G.mut,fontWeight:600,marginBottom:3}}>Height (px)</div>
                          <input type="number" value={FS[item.key+"H"]||item.defH} min={20} max={400}
                            onChange={e=>setFirmSettings(p=>({...p,[item.key+"H"]:Number(e.target.value)}))}
                            style={{...IS,padding:"4px 8px",fontSize:12}}/>
                        </div>
                      </div>
                      <div>
                        <div style={{fontSize:10,color:G.mut,fontWeight:600,marginBottom:3}}>Fit</div>
                        <select value={FS[item.key+"Fit"]||"contain"} onChange={e=>setFirmSettings(p=>({...p,[item.key+"Fit"]:e.target.value}))} style={{...IS,padding:"4px 8px",fontSize:12,cursor:"pointer"}}>
                          {["contain","cover","fill","scale-down"].map(f=><option key={f} value={f} style={{background:"#0B1610"}}>{f}</option>)}
                        </select>
                      </div>
                      <div style={{display:"flex",gap:6}}>
                        <div style={{flex:1,background:G.card,border:`1px solid ${G.bdr}`,borderRadius:7,padding:"5px 8px",fontSize:10,color:G.mut,textAlign:"center"}}>
                          Preview: <span style={{color:G.green,fontWeight:700}}>{FS[item.key+"W"]||item.defW}x{FS[item.key+"H"]||item.defH}</span>
                        </div>
                        <button onClick={()=>setFirmSettings(p=>({...p,[item.key]:null,[item.key+"W"]:item.defW,[item.key+"H"]:item.defH,[item.key+"Fit"]:"contain"}))} style={{fontSize:11,padding:"4px 9px",borderRadius:7,border:`1px solid ${G.red}44`,background:"transparent",color:G.red,cursor:"pointer",fontWeight:700}}>🗑</button>
                      </div>
                    </>}
                    {!FS[item.key]&&<div style={{fontSize:10,color:G.bdr,textAlign:"center"}}>{item.hint}</div>}
                  </div>
                ))}
              </div>
            </div>

            {/* Invoice Format / Performa Settings */}
            <div style={{background:G.card,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"14px 16px"}}>
              <div style={{fontSize:12,fontWeight:700,color:G.txt,marginBottom:12,display:"flex",alignItems:"center",gap:8}}><span style={{background:G.ind+"20",color:G.ind,width:24,height:24,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13}}>📋</span> Invoice Format / Performa Settings</div>
              <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10}}>
                <div>
                  <div style={{fontSize:10,color:G.mut,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Invoice Type</div>
                  <select value={FS.invoiceType||"TAX INVOICE"} onChange={e=>setFirmSettings(p=>({...p,invoiceType:e.target.value}))} style={{...IS,padding:"7px 10px",fontSize:12,cursor:"pointer"}}>
                    {["TAX INVOICE","PROFORMA INVOICE","RECEIPT","CREDIT NOTE","DEBIT NOTE"].map(t=><option key={t} value={t} style={{background:"#0B1610"}}>{t}</option>)}
                  </select>
                </div>
                <div>
                  <div style={{fontSize:10,color:G.mut,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Invoice No. Prefix</div>
                  <I val={FS.invPrefix||"FTM"} set={v=>setFirmSettings(p=>({...p,invPrefix:v}))} ph="FTM"/>
                </div>
                <div>
                  <div style={{fontSize:10,color:G.mut,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Copy Label</div>
                  <select value={FS.copyLabel||"ORIGINAL FOR RECIPIENT"} onChange={e=>setFirmSettings(p=>({...p,copyLabel:e.target.value}))} style={{...IS,padding:"7px 10px",fontSize:12,cursor:"pointer"}}>
                    {["ORIGINAL FOR RECIPIENT","DUPLICATE FOR SUPPLIER","TRIPLICATE FOR TRANSPORTER","OFFICE COPY"].map(t=><option key={t} value={t} style={{background:"#0B1610"}}>{t}</option>)}
                  </select>
                </div>
                <div>
                  <div style={{fontSize:10,color:G.mut,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Empty Service Rows</div>
                  <select value={String(FS.emptyRows||3)} onChange={e=>setFirmSettings(p=>({...p,emptyRows:Number(e.target.value)}))} style={{...IS,padding:"7px 10px",fontSize:12,cursor:"pointer"}}>
                    {["0","1","2","3","4","5","6","7","8"].map(n=><option key={n} value={n} style={{background:"#0B1610"}}>{n} empty rows</option>)}
                  </select>
                </div>
                <div style={{gridColumn:"1 / -1"}}>
                  <div style={{fontSize:10,color:G.mut,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>Firm Tagline (below name)</div>
                  <I val={FS.tagline||""} set={v=>setFirmSettings(p=>({...p,tagline:v}))} ph="e.g. Chartered Accountants & Tax Consultants"/>
                </div>
                <div style={{gridColumn:"1 / -1"}}>
                  <div style={{display:"flex",justifyContent:"space-between",fontSize:10,color:G.mut,fontWeight:700,marginBottom:4,textTransform:"uppercase"}}>
                    <span>Invoice Scale (fonts &amp; spacing)</span><span style={{color:G.amb,fontWeight:700}}>{FS.invoiceScale||100}%</span>
                  </div>
                  <input type="range" min={80} max={130} step={5} value={FS.invoiceScale||100}
                    onChange={e=>setFirmSettings(p=>({...p,invoiceScale:Number(e.target.value)}))}
                    style={{width:"100%",accentColor:G.amb,cursor:"pointer"}}/>
                  <div style={{fontSize:10,color:G.mut,marginTop:3}}>Scales all invoice text and spacing up or down together. 100% is the default larger, easier-to-read size.</div>
                </div>
              </div>
              {/* Live performa preview */}
              <div style={{marginTop:12,padding:"10px 14px",background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:9,fontSize:11,color:G.mut,lineHeight:1.8}}>
                <div style={{fontWeight:700,color:G.txt,marginBottom:4}}>📋 Format Preview</div>
                <div>Invoice No: <span style={{color:G.cyn,fontFamily:"monospace"}}>{FS.invPrefix||"FTM"}/2025-26/10031</span></div>
                <div>Type: <span style={{color:G.green,fontWeight:700}}>{FS.invoiceType||"TAX INVOICE"}</span></div>
                <div>Copy: <span style={{color:G.mut,fontStyle:"italic"}}>{FS.copyLabel||"ORIGINAL FOR RECIPIENT"}</span></div>
                <div>Empty rows in table: <span style={{color:G.amb,fontWeight:700}}>{FS.emptyRows||3}</span></div>
                <div>Invoice scale: <span style={{color:G.amb,fontWeight:700}}>{FS.invoiceScale||100}%</span></div>
              </div>
            </div>

            {/* Terms */}
            <div style={{background:G.card,border:`1px solid ${G.bdr}`,borderRadius:12,padding:"14px 16px"}}>
              <div style={{fontSize:12,fontWeight:700,color:G.txt,marginBottom:10,display:"flex",alignItems:"center",gap:8}}><span style={{background:G.amb+"20",color:G.amb,width:24,height:24,borderRadius:6,display:"flex",alignItems:"center",justifyContent:"center",fontSize:13}}>📜</span> Terms & Conditions (on Invoice)</div>
              <textarea value={FS.terms||""} onChange={e=>setFirmSettings(p=>({...p,terms:e.target.value}))} rows={3} style={{...IS,resize:"vertical",fontSize:12}} placeholder="Terms and conditions text..."/>
            </div>

            <div style={{padding:"10px 14px",background:G.green+"0A",border:`1px solid ${G.green}22`,borderRadius:10,fontSize:12,color:G.mut}}>
              ✅ All settings here are automatically applied to every invoice you print. No need to re-enter each time.
            </div>
          </>;
        })()}
      </div>
    </Crd>
    :sec==="profile"
    ?<Crd><SH icon="🖼" title="Profile Picture / Logo" acc={G.cyn}/>
      <div style={{marginTop:16,display:"flex",flexDirection:"column",gap:18}}>
        {/* Current logo preview */}
        <div style={{display:"flex",gap:18,alignItems:"center",padding:"16px 20px",background:G.card,border:`1px solid ${G.bdr}`,borderRadius:14}}>
          <div style={{position:"relative",flexShrink:0}}>
            {profilePic
              ?<img src={profilePic} alt="Profile" style={{width:80,height:80,borderRadius:16,objectFit:"cover",border:`3px solid ${G.green}`,display:"block"}}/>
              :<div style={{width:80,height:80,borderRadius:16,border:`3px solid ${G.green}`,display:"flex",alignItems:"center",justifyContent:"center",background:G.surf}}><Logo sz={56}/></div>}
            {profilePic&&<button onClick={()=>setProfilePic(null)}
              style={{position:"absolute",top:-8,right:-8,width:22,height:22,borderRadius:"50%",background:G.red,border:"none",color:"#fff",cursor:"pointer",fontSize:13,display:"flex",alignItems:"center",justifyContent:"center",lineHeight:1}}>✕</button>}
          </div>
          <div>
            <div style={{fontWeight:700,fontSize:15,color:G.wh}}>Fin-Tax Mitra</div>
            <div style={{fontSize:12,color:G.mut,marginTop:3}}>Making Taxes Easy</div>
            <div style={{fontSize:11,color:G.mut,marginTop:8}}>{profilePic?"Custom logo active":"Default SVG logo active"}</div>
            {profilePic&&<button onClick={()=>setProfilePic(null)} style={{marginTop:8,fontSize:12,padding:"4px 12px",borderRadius:8,border:`1px solid ${G.red}44`,background:"transparent",color:G.red,cursor:"pointer",fontWeight:600}}>🗑 Remove - use default</button>}
          </div>
        </div>
        {/* Upload area */}
        <div style={{border:`2px dashed ${G.green}55`,borderRadius:14,padding:"28px 20px",textAlign:"center",background:G.card,cursor:"pointer",transition:"border .2s"}}
          onMouseEnter={e=>e.currentTarget.style.borderColor=G.green}
          onMouseLeave={e=>e.currentTarget.style.borderColor=G.green+"55"}
          onClick={()=>document.getElementById("picUpload").click()}>
          <div style={{fontSize:36,marginBottom:10}}>📁</div>
          <div style={{fontWeight:700,fontSize:14,color:G.wh,marginBottom:6}}>Click to Upload Image</div>
          <div style={{fontSize:12,color:G.mut}}>PNG, JPG, JPEG, SVG * Max 5MB</div>
          <div style={{fontSize:11,color:G.mut,marginTop:4}}>Recommended: Square image (1:1 ratio)</div>
          <input id="picUpload" type="file" accept="image/*" style={{display:"none"}}
            onChange={e=>{
              const file=e.target.files[0];
              if(!file)return;
              if(file.size>5*1024*1024){alert("File too large! Max 5MB.");return;}
              const reader=new FileReader();
              reader.onload=ev=>setProfilePic(ev.target.result);
              reader.readAsDataURL(file);
              e.target.value="";
            }}/>
        </div>
        {/* Tips */}
        <div style={{padding:"12px 16px",background:G.card,border:`1px solid ${G.bdr}`,borderRadius:10,fontSize:12,color:G.mut,lineHeight:1.8}}>
          <div style={{fontWeight:700,color:G.txt,marginBottom:4}}>📌 Tips</div>
          <div>* Square images look best in the sidebar (1:1 ratio)</div>
          <div>* The logo appears in the sidebar header and all auth modals</div>
          <div>* Click ✕ on the preview or "Remove" to restore the default SVG logo</div>
          <div>* Image is stored in browser memory - re-upload after page refresh</div>
        </div>
      </div>
    </Crd>
    :sec==="theme"
    ?<Crd><SH icon={darkMode?"🌙":"☀️"} title="App Appearance - Dark / Light Mode" acc={G.amb}/>
      <div style={{marginTop:16,display:"flex",flexDirection:"column",gap:16}}>
        <div style={{display:"flex",gap:14}}>
          <div onClick={()=>setDarkMode(true)} style={{flex:1,cursor:"pointer",border:`2px solid ${darkMode?G.green:G.bdr}`,borderRadius:14,overflow:"hidden",transition:"all .2s",boxShadow:darkMode?`0 0 20px ${G.green}44`:"none"}}>
            <div style={{background:"#060E0A",padding:"14px 12px",borderBottom:"1px solid #1A3020"}}>
              <div style={{width:"100%",height:7,background:"#0B1610",borderRadius:4,marginBottom:6}}/>
              <div style={{display:"flex",gap:5,marginBottom:6}}>
                <div style={{width:24,height:24,borderRadius:5,background:"#16A34A"}}/>
                <div style={{flex:1}}><div style={{width:"60%",height:4,background:"#1A3020",borderRadius:3,marginBottom:3}}/><div style={{width:"40%",height:3,background:"#1A3020",borderRadius:3}}/></div>
              </div>
              <div style={{width:"100%",height:28,background:"#0B1610",borderRadius:5,marginBottom:5}}/>
              <div style={{display:"flex",gap:3}}>{[1,2,3].map(i=><div key={i} style={{flex:1,height:16,background:"#0F1D14",borderRadius:4}}/>)}</div>
            </div>
            <div style={{background:"#060E0A",padding:"9px 12px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
              <span style={{fontSize:13,fontWeight:700,color:"#F0FDF4"}}>🌙 Dark</span>
              {darkMode&&<span style={{fontSize:11,background:"#14532D",color:"#4ADE80",padding:"2px 8px",borderRadius:10,fontWeight:700}}>✓ Active</span>}
            </div>
          </div>
          <div onClick={()=>setDarkMode(false)} style={{flex:1,cursor:"pointer",border:`2px solid ${!darkMode?"#16A34A":G.bdr}`,borderRadius:14,overflow:"hidden",transition:"all .2s",boxShadow:!darkMode?"0 0 20px #16A34A44":"none"}}>
            <div style={{background:"#F0FDF4",padding:"14px 12px",borderBottom:"1px solid #BBF7D0"}}>
              <div style={{width:"100%",height:7,background:"#FFFFFF",borderRadius:4,marginBottom:6}}/>
              <div style={{display:"flex",gap:5,marginBottom:6}}>
                <div style={{width:24,height:24,borderRadius:5,background:"#16A34A"}}/>
                <div style={{flex:1}}><div style={{width:"60%",height:4,background:"#BBF7D0",borderRadius:3,marginBottom:3}}/><div style={{width:"40%",height:3,background:"#BBF7D0",borderRadius:3}}/></div>
              </div>
              <div style={{width:"100%",height:28,background:"#FFFFFF",borderRadius:5,border:"1px solid #BBF7D0",marginBottom:5}}/>
              <div style={{display:"flex",gap:3}}>{[1,2,3].map(i=><div key={i} style={{flex:1,height:16,background:"#F8FFFE",borderRadius:4,border:"1px solid #BBF7D0"}}/>)}</div>
            </div>
            <div style={{background:"#F0FDF4",padding:"9px 12px",display:"flex",alignItems:"center",justifyContent:"space-between"}}>
              <span style={{fontSize:13,fontWeight:700,color:"#14532D"}}>☀️ Light</span>
              {!darkMode&&<span style={{fontSize:11,background:"#DCFCE7",color:"#15803D",padding:"2px 8px",borderRadius:10,fontWeight:700}}>✓ Active</span>}
            </div>
          </div>
        </div>
        <div style={{padding:"11px 14px",background:darkMode?"#0B1610":"#DCFCE7",border:`1px solid ${G.bdr}`,borderRadius:10,fontSize:12,color:G.mut,lineHeight:1.7}}>
          <span style={{fontWeight:700,color:G.txt}}>{darkMode?"🌙 Dark Mode Active - ":"☀️ Light Mode Active - "}</span>
          Click a card above to switch theme. Changes apply instantly across the entire app.
        </div>
      </div>
    </Crd>
    :sec==="passwords"
    ?<Crd><SH icon="🔑" title="Change System Passwords" acc={G.amb}/>
      <div style={{display:"flex",flexDirection:"column",gap:14,marginTop:4}}>
        <div style={{padding:"10px 14px",background:"#43140720",border:`1px solid ${G.amb}33`,borderRadius:9,fontSize:12,color:G.amb}}>⚠️ Leave blank to keep current. Changes apply immediately.</div>
        <F label="New Owner Password (Finance)"><I val={pf.owner} set={v=>setPf({...pf,owner:v})} type="password" ph="Leave blank to keep current"/></F>
        <F label="New Developer Password (Dev Tab)"><I val={pf.dev} set={v=>setPf({...pf,dev:v})} type="password" ph="Leave blank to keep current"/></F>
        <Btn onClick={savePw} sty={{padding:12,fontSize:14}}>💾 Update Passwords</Btn>
        {pm&&<div style={{padding:"9px 13px",background:G.green+"15",border:`1px solid ${G.green}44`,borderRadius:9,fontSize:13,color:G.green,fontWeight:600}}>✓ {pm}</div>}
      </div>
    </Crd>
    :<Crd><SH icon={cm?.icon||"📋"} title={`Edit - ${cm?.l||""}`} acc={G.green}/>
      <div style={{fontSize:12,color:G.mut,marginBottom:13}}>{cur.length} items · {cm?.cat}</div>
      <div style={{display:"flex",gap:8,marginBottom:14}}>
        <I val={nv} set={setNv} ph={`Add new ${cm?.l?.toLowerCase()}...`} kd={e=>e.key==="Enter"&&add()} sty={{flex:1}}/>
        <Btn onClick={add} sty={{padding:"9px 17px",fontSize:13,whiteSpace:"nowrap"}}>➕ Add</Btn>
      </div>
      <div style={{display:"flex",flexDirection:"column",gap:5}}>
        {cur.map((item,i)=><div key={i} style={{display:"flex",alignItems:"center",gap:7,padding:"8px 11px",background:G.bg,border:`1px solid ${G.bdr}`,borderRadius:8}}>
          <span style={{fontSize:11,color:G.bdr,minWidth:22,textAlign:"right",fontWeight:600}}>{i+1}.</span>
          {ei===i?<input value={ev} onChange={e=>setEv(e.target.value)} autoFocus onKeyDown={e=>e.key==="Enter"&&saveE(i)} style={{...IS,flex:1,padding:"4px 8px",fontSize:12}}/>
          :<span style={{flex:1,fontSize:13,color:G.txt}}>{item}</span>}
          <div style={{display:"flex",gap:3,flexShrink:0}}>
            <button onClick={()=>mu(i)} style={{background:"transparent",border:`1px solid ${G.bdr}`,color:G.mut,borderRadius:6,padding:"2px 7px",cursor:"pointer",fontSize:11}}>^</button>
            <button onClick={()=>md(i)} style={{background:"transparent",border:`1px solid ${G.bdr}`,color:G.mut,borderRadius:6,padding:"2px 7px",cursor:"pointer",fontSize:11}}>v</button>
            {ei===i
            ?<><button onClick={()=>saveE(i)} style={{background:G.g2,border:"none",color:"#fff",borderRadius:6,padding:"2px 8px",cursor:"pointer",fontSize:11,fontWeight:700}}>✓</button>
              <button onClick={()=>{setEi(null);setEv("");}} style={{background:"transparent",border:`1px solid ${G.bdr}`,color:G.mut,borderRadius:6,padding:"2px 7px",cursor:"pointer",fontSize:11}}>✕</button></>
            :<><button onClick={()=>{setEi(i);setEv(item);}} style={{background:"transparent",border:`1px solid ${G.bdr}`,color:G.cyn,borderRadius:6,padding:"2px 8px",cursor:"pointer",fontSize:11}}>✏️</button>
              <button onClick={()=>del(i)} style={{background:"#450A0A",border:`1px solid ${G.red}44`,color:G.red,borderRadius:6,padding:"2px 7px",cursor:"pointer",fontSize:11,fontWeight:700}}>🗑</button></>}
          </div>
        </div>)}
        {cur.length===0&&<div style={{padding:"20px",textAlign:"center",color:G.bdr,fontSize:13}}>No items. Add one above.</div>}
      </div>
    </Crd>}
    </div>
  </div>;
}

// ─── Root App ─────────────────────────────────────────────────────────────────

// ══════════════════════════════════════════════════════════════════════════
// ITR COMPUTATION MODULE (embedded Tax Computation tool)
// ══════════════════════════════════════════════════════════════════════════
/* ============================================================
   TAX RULE CONFIGURATION — AY 2026-27 (FY 2025-26)
   Kept as data, not hard-coded logic, so future years can be
   added by appending a new key here.
   ============================================================ */
const TAX_CONFIG = {
  "AY 2026-27": {
    fy: "2025-26",
    newRegime: {
      standardDeduction: 75000,
      slabs: [
        { upto: 400000, rate: 0 },
        { upto: 800000, rate: 5 },
        { upto: 1200000, rate: 10 },
        { upto: 1600000, rate: 15 },
        { upto: 2000000, rate: 20 },
        { upto: 2400000, rate: 25 },
        { upto: Infinity, rate: 30 },
      ],
      rebateLimit: 1200000,
      rebateMax: 60000,
      surcharge: [
        { above: 5000000, rate: 10 },
        { above: 10000000, rate: 15 },
        { above: 20000000, rate: 25 },
      ],
      surchargeCap: 25,
    },
    oldRegime: {
      standardDeduction: 50000,
      slabsByAge: {
        normal: [
          { upto: 250000, rate: 0 },
          { upto: 500000, rate: 5 },
          { upto: 1000000, rate: 20 },
          { upto: Infinity, rate: 30 },
        ],
        senior: [
          { upto: 300000, rate: 0 },
          { upto: 500000, rate: 5 },
          { upto: 1000000, rate: 20 },
          { upto: Infinity, rate: 30 },
        ],
        superSenior: [
          { upto: 500000, rate: 0 },
          { upto: 1000000, rate: 20 },
          { upto: Infinity, rate: 30 },
        ],
      },
      rebateLimit: 500000,
      rebateMax: 12500,
      surcharge: [
        { above: 5000000, rate: 10 },
        { above: 10000000, rate: 15 },
        { above: 20000000, rate: 25 },
        { above: 50000000, rate: 37 },
      ],
      surchargeCap: 37,
    },
    cess: 4,
    ltcgRate: 12.5,
    ltcgIndexedRate: 20,
    ltcgEquityExemption: 125000,
    stcgEquityRate: 20,
    hpLossSetOffCap: 200000,
    chapterVIA: {
      "80C": { label: "80C — LIC / PPF / ELSS / Tuition / NSC / Sukanya / EPF etc.", cap: 150000, regime: "old" },
      "80CCD1B": { label: "80CCD(1B) — NPS additional contribution", cap: 50000, regime: "old" },
      "80CCD2": { label: "80CCD(2) — Employer's NPS contribution", cap: null, regime: "both" },
      "80D": { label: "80D — Medical insurance premium", cap: null, regime: "old" },
      "80DD": { label: "80DD — Maintenance of disabled dependant", cap: 75000, capSevere: 125000, regime: "old" },
      "80DDB": { label: "80DDB — Medical treatment of specified disease", cap: 40000, capSenior: 100000, regime: "old" },
      "80E": { label: "80E — Interest on education loan", cap: null, regime: "old" },
      "80EE": { label: "80EE — Additional interest on housing loan", cap: 50000, regime: "old" },
      "80EEA": { label: "80EEA — Interest on loan for affordable housing", cap: 150000, regime: "old" },
      "80G": { label: "80G — Donations to approved funds", cap: null, regime: "old" },
      "80GGA": { label: "80GGA — Donations for scientific research / rural development", cap: null, regime: "old" },
      "80TTA": { label: "80TTA — Interest on savings account", cap: 10000, regime: "old" },
      "80TTB": { label: "80TTB — Interest income (senior citizens)", cap: 50000, regime: "old" },
      "80U": { label: "80U — Self disability", cap: 75000, capSevere: 125000, regime: "old" },
    },
  },
};

const fmt = (n) => {
  const v = Math.round(Number(n) || 0);
  const neg = v < 0;
  const s = Math.abs(v).toLocaleString("en-IN");
  return (neg ? "-₹" : "₹") + s;
};
const num = (v) => (isNaN(parseFloat(v)) ? 0 : parseFloat(v));
const uid = () => Math.random().toString(36).slice(2, 10);
/** Strips a leading "AY "/"FY " label from a value when displaying it next
 *  to a caption that already says "Assessment Year"/"Financial Year" —
 *  avoids showing e.g. "Assessment Year : AY 2026-27". */
const yearOnly = (s) => (s || "").replace(/^(AY|FY)\s*/i, "");
const moveItem = (arr, from, to) => {
  const next = [...arr];
  const [moved] = next.splice(from, 1);
  next.splice(to, 0, moved);
  return next;
};

/** Age as on the last day of the financial year (31 March of the AY's
 *  second year) — the standard rule for senior-citizen slab eligibility. */
function getFYEndDate(ayString) {
  const m = (ayString || "").match(/(\d{4})/);
  const endYear = m ? parseInt(m[1], 10) : new Date().getFullYear();
  return new Date(endYear, 2, 31);
}
function getAgeCategory(dobStr, ayString) {
  if (!dobStr) return "normal";
  const dob = new Date(dobStr);
  if (isNaN(dob.getTime())) return "normal";
  const fyEnd = getFYEndDate(ayString);
  let age = fyEnd.getFullYear() - dob.getFullYear();
  const hadBirthday = fyEnd.getMonth() > dob.getMonth() || (fyEnd.getMonth() === dob.getMonth() && fyEnd.getDate() >= dob.getDate());
  if (!hadBirthday) age -= 1;
  if (age >= 80) return "superSenior";
  if (age >= 60) return "senior";
  return "normal";
}
/** Whole months between two dates, counting any part-month as a full month
 *  (how 234A/B/C interest is actually charged). */
function monthsBetween(fromDate, toDate) {
  if (!(fromDate instanceof Date) || !(toDate instanceof Date) || isNaN(fromDate) || isNaN(toDate) || toDate <= fromDate) return 0;
  let months = (toDate.getFullYear() - fromDate.getFullYear()) * 12 + (toDate.getMonth() - fromDate.getMonth());
  if (toDate.getDate() > fromDate.getDate()) months += 1;
  return Math.max(months, 0);
}

/* ============================================================
   CALCULATION ENGINE
   ============================================================ */
function calculateSalary(s) {
  const manualTotal = (s.manual || []).reduce((sum, m) => sum + num(m.amount), 0);
  const gross = num(s.basic) + num(s.da) + num(s.hra) + num(s.bonus) + num(s.commission) + num(s.perquisites) + num(s.otherAllowances) + manualTotal;
  const preStdDeduction = gross - num(s.professionalTax) - num(s.entertainmentAllowance);
  return { gross, preStdDeduction, manualTotal };
}

function calculateHouseProperty(h) {
  const manualTotal = (h.manual || []).reduce((sum, m) => sum + num(m.amount), 0);
  if (h.type === "self") {
    const interest = Math.min(num(h.interest), 200000);
    return { income: -interest + manualTotal, interestAllowedOld: interest, interestAllowedNew: 0, gav: 0, nav: 0, stdDed: 0, manualTotal };
  }
  const gav = Math.max(num(h.municipalValue), num(h.rentReceived));
  const nav = gav - num(h.municipalTax);
  const stdDed = Math.max(nav, 0) * 0.3;
  const interest = num(h.interest);
  const income = nav - stdDed - interest + manualTotal;
  return { income, interestAllowedOld: interest, interestAllowedNew: interest, gav, nav, stdDed, manualTotal };
}

function calculateBusinessIncome(b) {
  const manualTotal = (b.manual || []).reduce((sum, m) => sum + num(m.amount), 0);
  if (b.type === "normal") {
    return { income: num(b.income) - num(b.expenses) - num(b.depreciation) + manualTotal, manualTotal };
  }
  if (b.type === "44AD") {
    const rate = b.digital ? 6 : 8;
    const computed = num(b.turnover) * (rate / 100);
    return { income: Math.max(computed, num(b.declaredIncome) || 0) + manualTotal, manualTotal };
  }
  if (b.type === "44ADA") {
    const computed = num(b.grossReceipts) * 0.5;
    return { income: Math.max(computed, num(b.declaredIncome) || 0) + manualTotal, manualTotal };
  }
  if (b.type === "44AE") {
    const perVehicleMonthly = 7500;
    const computed = num(b.vehicles) * num(b.months) * perVehicleMonthly;
    return { income: Math.max(computed, num(b.declaredIncome) || 0) + manualTotal, manualTotal };
  }
  return { income: manualTotal, manualTotal };
}

function calculateCapitalGain(cg, config) {
  const st = cg.st.rows.reduce((acc, r) => {
    const gain = num(r.sale) - num(r.cost) - num(r.improvement) - num(r.expenses);
    if (r.equitySTT) acc.special20 += gain; else acc.slab += gain;
    return acc;
  }, { special20: 0, slab: 0 });

  let ltEquity = 0;
  let ltOther = 0; // taxed at flat 12.5%, no indexation
  let ltGrandfatheredTax = 0; // pre-computed tax for rows that chose the 20%-with-indexation method
  let ltGrandfatheredGain = 0; // the gain amount used for those rows (for total-income display)
  const ltDetail = [];

  cg.lt.rows.forEach((r) => {
    const costPlain = num(r.cost) + num(r.improvement) + num(r.expenses);
    const gainPlain = num(r.sale) - costPlain;
    if (r.equitySTT) {
      ltEquity += gainPlain;
      ltDetail.push({ row: r, gain: gainPlain, method: "equity" });
      return;
    }
    if (r.grandfathered && num(r.indexedCost) > 0) {
      const costIndexed = num(r.indexedCost) + num(r.improvement) + num(r.expenses);
      const gainIndexed = num(r.sale) - costIndexed;
      const taxPlain = Math.max(gainPlain, 0) * ((config ? config.ltcgRate : 12.5) / 100);
      const taxIndexed = Math.max(gainIndexed, 0) * ((config ? config.ltcgIndexedRate : 20) / 100);
      if (taxIndexed < taxPlain) {
        ltGrandfatheredTax += taxIndexed;
        ltGrandfatheredGain += gainIndexed;
        ltDetail.push({ row: r, gain: gainIndexed, method: "20% with indexation", tax: taxIndexed });
      } else {
        ltOther += gainPlain;
        ltDetail.push({ row: r, gain: gainPlain, method: "12.5% without indexation", tax: taxPlain });
      }
    } else {
      ltOther += gainPlain;
      ltDetail.push({ row: r, gain: gainPlain, method: "12.5% without indexation" });
    }
  });

  return {
    st,
    lt: { equity: ltEquity, other: ltOther, grandfatheredTax: ltGrandfatheredTax, grandfatheredGain: ltGrandfatheredGain, detail: ltDetail },
    total: st.special20 + st.slab + ltEquity + ltOther + ltGrandfatheredGain,
  };
}

function calculateOtherIncome(rows) {
  const total = rows.reduce((s, r) => s + num(r.amount), 0);
  const savingsInterest = rows.filter((r) => r.type === "Interest on Savings Bank").reduce((s, r) => s + num(r.amount), 0);
  const fdInterest = rows.filter((r) => r.type === "Interest on FD").reduce((s, r) => s + num(r.amount), 0);
  return { total, savingsInterest, fdInterest };
}

function slabTax(income, slabs) {
  let tax = 0, prev = 0;
  for (const slab of slabs) {
    if (income > prev) {
      const taxable = Math.min(income, slab.upto) - prev;
      tax += taxable * (slab.rate / 100);
      prev = slab.upto;
    } else break;
  }
  return tax;
}

function surchargeFor(taxableIncome, taxBeforeSurcharge, surchargeSlabs, cap, slabsForRelief) {
  let rate = 0, threshold = 0;
  for (const s of surchargeSlabs) {
    if (taxableIncome > s.above) { rate = s.rate; threshold = s.above; }
  }
  rate = Math.min(rate, cap);
  if (rate === 0 || threshold === 0) return { rate, amount: taxBeforeSurcharge * (rate / 100) };

  let amount = taxBeforeSurcharge * (rate / 100);

  // Marginal relief: tax + surcharge just above a surcharge threshold must
  // never exceed (tax at the threshold) + (income beyond the threshold) —
  // applies at every surcharge slab boundary, not only the rebate limit.
  if (slabsForRelief) {
    const taxAtThreshold = slabTax(threshold, slabsForRelief);
    const excessIncome = taxableIncome - threshold;
    const cappedTotal = taxAtThreshold + excessIncome;
    if (taxBeforeSurcharge + amount > cappedTotal) {
      amount = Math.max(cappedTotal - taxBeforeSurcharge, 0);
    }
  }
  return { rate, amount };
}

function computeRegimeTax({ regimeKey, config, totalIncomeExclSpecial, stcgSpecial20, ltcgEquity, ltcgOther, ltcgGrandfatheredTax, ltcgGrandfatheredGain, isEligibleRebate, ageCategory }) {
  const regime = config[regimeKey];
  const slabs = regimeKey === "oldRegime" ? regime.slabsByAge[ageCategory || "normal"] : regime.slabs;
  
  // Section 288A: Round total taxable income to the nearest multiple of ten
  const ltcgEquityTaxable = Math.max(ltcgEquity - config.ltcgEquityExemption, 0);
  const rawTotalTaxableIncome = totalIncomeExclSpecial + stcgSpecial20 + ltcgEquityTaxable + Math.max(ltcgOther, 0) + Math.max(ltcgGrandfatheredGain || 0, 0);
  const totalTaxableIncome = Math.round(rawTotalTaxableIncome / 10) * 10;
  
  // Adjust totalIncomeExclSpecial based on rounded difference for slab calculations
  const diff = totalTaxableIncome - rawTotalTaxableIncome;
  const adjustedTotalIncomeExclSpecial = Math.max(totalIncomeExclSpecial + diff, 0);

  const baseTax = slabTax(Math.max(adjustedTotalIncomeExclSpecial, 0), slabs);
  const specialTax = stcgSpecial20 * (config.stcgEquityRate / 100) + ltcgEquityTaxable * (config.ltcgRate / 100) + Math.max(ltcgOther, 0) * (config.ltcgRate / 100) + (ltcgGrandfatheredTax || 0);
  let taxBeforeRebate = baseTax + specialTax;

  let rebate = 0;
  if (isEligibleRebate && totalTaxableIncome <= regime.rebateLimit) {
    rebate = Math.min(baseTax, regime.rebateMax);
  }
  let taxAfterRebate = Math.max(taxBeforeRebate - rebate, 0);

  // Marginal relief just above the rebate threshold (Section 87A)
  if (isEligibleRebate && totalTaxableIncome > regime.rebateLimit) {
    const excessIncome = totalTaxableIncome - regime.rebateLimit;
    if (taxAfterRebate > excessIncome) {
      taxAfterRebate = excessIncome;
    }
  }

  // Surcharge, with marginal relief applied at every surcharge slab boundary
  const surcharge = surchargeFor(totalTaxableIncome, taxAfterRebate, regime.surcharge, regime.surchargeCap, slabs);
  const cessAmount = (taxAfterRebate + surcharge.amount) * (config.cess / 100);
  const rawTotal = taxAfterRebate + surcharge.amount + cessAmount;
  
  // Section 288B: Round tax liability to the nearest multiple of ten
  const total = Math.round(rawTotal / 10) * 10;

  return {
    baseTax, specialTax, taxBeforeRebate, rebate, taxAfterRebate,
    surchargeRate: surcharge.rate, surchargeAmount: surcharge.amount,
    cessAmount, total, totalTaxableIncome, slabsUsed: slabs,
  };
}

/** Suggests 234A/B/C interest and the 234F late fee from the actual filing
 *  date and quarter-wise advance tax paid. This is a standard-case
 *  approximation (non-audit due date, no updated-return scenarios) —
 *  returned as suggestions the user can accept or override. */
function computeInterestSuggestions({ assessedTax, tdsAndTcs, advanceTaxByQuarter, selfAssessmentTax, filingDateStr, dueDateStr, totalIncome, ay }) {
  const zero = { interest234A: 0, interest234B: 0, interest234C: 0, fee234F: 0 };
  const netTaxLiability = Math.max(assessedTax - tdsAndTcs, 0); // the amount that should have come via advance tax
  if (netTaxLiability <= 0) return zero;

  const fyEndYear = getFYEndDate(ay).getFullYear();
  const q = advanceTaxByQuarter || {};
  const q1 = num(q.q1), q2 = num(q.q2), q3 = num(q.q3), q4 = num(q.q4);
  const totalAdvanceTax = q1 + q2 + q3 + q4;

  const filingDate = filingDateStr ? new Date(filingDateStr) : null;
  const dueDate = dueDateStr ? new Date(dueDateStr) : new Date(fyEndYear, 6, 31); // 31 July by default

  // 234C: quarterly shortfall, 1% for 3 months each (1 month for Q4), only
  // if advance tax obligation applies at all (small net-liability exemption
  // of ₹10,000 is handled by the netTaxLiability<=0 guard above being lenient;
  // real threshold is technically ₹10,000 — approximate check here).
  let interest234C = 0;
  if (netTaxLiability > 10000) {
    const req1 = netTaxLiability * 0.15, req2 = netTaxLiability * 0.45, req3 = netTaxLiability * 0.75, req4 = netTaxLiability;
    const short1 = Math.max(req1 - q1, 0);
    const short2 = Math.max(req2 - (q1 + q2), 0);
    const short3 = Math.max(req3 - (q1 + q2 + q3), 0);
    const short4 = Math.max(req4 - totalAdvanceTax, 0);
    interest234C = (short1 + short2 + short3) * 0.01 * 3 + short4 * 0.01 * 1;
  }

  // 234B: shortfall interest if advance tax paid < 90% of assessed tax,
  // 1% per month from 1 April of the AY to the filing date (or today).
  let interest234B = 0;
  if (totalAdvanceTax < netTaxLiability * 0.9) {
    const shortfall = netTaxLiability - totalAdvanceTax;
    const aprilFirst = new Date(fyEndYear, 3, 1);
    const endPoint = filingDate && !isNaN(filingDate) ? filingDate : new Date();
    const months = monthsBetween(aprilFirst, endPoint) || 1;
    interest234B = shortfall * 0.01 * months;
  }

  // 234A: interest for late filing, 1% per month from due date to filing
  // date, on tax outstanding after TDS/TCS/advance tax/self-assessment tax
  // paid on or before the due date.
  let interest234A = 0;
  if (filingDate && !isNaN(filingDate) && filingDate > dueDate) {
    const outstanding = Math.max(netTaxLiability - totalAdvanceTax - num(selfAssessmentTax), 0);
    const months = monthsBetween(dueDate, filingDate);
    interest234A = outstanding * 0.01 * months;
  }

  // 234F: flat late-filing fee — ₹5,000 normally, ₹1,000 if total income
  // is ≤ ₹5,00,000, ₹0 if filed on or before the due date.
  let fee234F = 0;
  if (filingDate && !isNaN(filingDate) && filingDate > dueDate) {
    fee234F = totalIncome <= 500000 ? 1000 : 5000;
  }

  return {
    interest234A: Math.round(interest234A),
    interest234B: Math.round(interest234B),
    interest234C: Math.round(interest234C),
    fee234F,
  };
}


/* ============================================================
   DEFAULT STATE SHAPES
   ============================================================ */
const defaultAssessee = {
  ay: "AY 2026-27", fy: "2025-26", itrType: "ITR-2", name: "", pan: "", aadhaar: "", fatherName: "",
  dob: "", gender: "Male", residentialStatus: "Resident", category: "Individual",
  address: "", state: "", pin: "", email: "", mobile: "",
  bankName: "", ifsc: "", accountNumber: "",
};

const defaultIncome = {
  salary: { enabled: false, employerName: "", basic: "", da: "", hra: "", bonus: "", commission: "", perquisites: "", otherAllowances: "", professionalTax: "", entertainmentAllowance: "", manual: [] },
  houseProperty: { enabled: false, type: "self", municipalValue: "", rentReceived: "", municipalTax: "", interest: "", manual: [] },
  business: { enabled: false, businessName: "", type: "normal", income: "", expenses: "", depreciation: "", turnover: "", grossReceipts: "", vehicles: "", months: "", declaredIncome: "", digital: false, manual: [] },
  capitalGains: {
    enabled: false,
    st: { rows: [] },
    lt: { rows: [] },
  },
  otherSources: { enabled: false, rows: [] },
};

const defaultDeductions = Object.fromEntries(
  Object.keys(TAX_CONFIG["AY 2026-27"].chapterVIA).map((k) => [k, { enabled: false, amount: "", severe: false, senior: false }])
);

const defaultManualDeductions = [];

const defaultTaxPaid = {
  tds: "", advanceTax: "", selfAssessmentTax: "", tcs: "",
  advanceTaxQ1: "", advanceTaxQ2: "", advanceTaxQ3: "", advanceTaxQ4: "",
  filingDate: "", dueDate: "",
  interest234A: "", interest234B: "", interest234C: "", fee234F: "", manual: [],
};

const defaultLetterhead = {
  logo: null, logoSize: 64, logoX: 6, logoY: 0,
  footerLogo: null, footerLogoSize: 48, footerLogoX: 0, footerLogoY: 0,
  banner: null, bannerWidthPct: 100, bannerHeight: 70,
  firmName: "", tagline: "", contact: "", footerAd: "",
};

/* Page Setup — Excel-style print configuration */
const PAPER_SIZES = {
  A4: [210, 297],
  Letter: [215.9, 279.4],
  Legal: [215.9, 355.6],
};
const defaultPageSetup = {
  paperSize: "A4",
  orientation: "portrait",
  marginTop: 14,
  marginBottom: 14,
  marginLeft: 14,
  marginRight: 14,
  scale: 100,
  piTwoColumn: false,
  piFields: {
    name: true, pan: true, aadhaar: true, fatherName: true, dob: true, address: true,
    email: true, mobile: true, bankName: true, ifsc: true, accountNumber: true, status: true,
  },
};
const PI_FIELD_DEFS = [
  { key: "name", label: "Name" },
  { key: "pan", label: "PAN" },
  { key: "aadhaar", label: "Aadhaar Card No." },
  { key: "fatherName", label: "Father's Name" },
  { key: "dob", label: "Date of Birth" },
  { key: "address", label: "Address" },
  { key: "email", label: "E-Mail ID" },
  { key: "mobile", label: "Mobile No." },
  { key: "bankName", label: "Primary Bank Name" },
  { key: "ifsc", label: "Primary Bank IFSC Code" },
  { key: "accountNumber", label: "Primary Bank Account No." },
  { key: "status", label: "Status" },
];

function buildPrintCSS(ps) {
  const base = PAPER_SIZES[ps.paperSize] || PAPER_SIZES.A4;
  const w = ps.orientation === "landscape" ? base[1] : base[0];
  const h = ps.orientation === "landscape" ? base[0] : base[1];
  const scale = Math.max(50, Math.min(150, num(ps.scale) || 100)) / 100;
  const printH = h - num(ps.marginTop) - num(ps.marginBottom);
  return `@media print {
    @page { size: ${w}mm ${h}mm; margin: ${num(ps.marginTop)}mm ${num(ps.marginRight)}mm ${num(ps.marginBottom)}mm ${num(ps.marginLeft)}mm; }
    .sheet {
      height: ${printH}mm !important;
      border: none !important;
      box-shadow: none !important;
      padding: 0 !important;
      margin: 0 !important;
      display: flex !important;
      flex-direction: column !important;
      box-sizing: border-box !important;
    }
    .sheet-page { zoom: ${scale}; }
    .sheet-footer {
      margin-top: auto !important;
    }
  }`;
}

/* ============================================================
   DUMMY / DEMO DATA - five realistic, fully-filled sample
   assessees so the app (and the saved-computations list) can
   be explored without typing in real data first.
   ============================================================ */
function fullDeductions(overrides) {
  const base = Object.fromEntries(
    Object.keys(TAX_CONFIG["AY 2026-27"].chapterVIA).map((k) => [k, { enabled: false, amount: "", severe: false, senior: false }])
  );
  Object.entries(overrides).forEach(([k, v]) => { base[k] = { enabled: true, severe: false, senior: false, ...v }; });
  return base;
}

function buildDummyEntries() {
  const entries = [];

  // 1. Salaried individual with a side business, rental property, and capital gains
  entries.push({
    assessee: {
      ay: "AY 2026-27", fy: "2025-26",
      name: "Ananya Sharma", pan: "ABCDE1234F", aadhaar: "234567891012",
      fatherName: "Rajesh Sharma", dob: "1990-06-14", gender: "Female",
      residentialStatus: "Resident", category: "Individual",
      address: "14/B, Lake Gardens", state: "West Bengal", pin: "700045",
      email: "ananya.sharma@example.com", mobile: "9830012345",
      bankName: "State Bank of India", ifsc: "SBIN0001234", accountNumber: "30456789123",
    },
    income: {
      salary: { enabled: true, employerName: "Meridian Software Solutions Pvt. Ltd.", basic: "720000", da: "36000", hra: "288000", bonus: "60000", commission: "18000", perquisites: "22000", otherAllowances: "45000", professionalTax: "2400", entertainmentAllowance: "3000", manual: [] },
      houseProperty: { enabled: true, type: "letout", municipalValue: "240000", rentReceived: "300000", municipalTax: "12000", interest: "180000", manual: [] },
      business: { enabled: true, businessName: "Sharma Trading Co.", type: "normal", income: "480000", expenses: "150000", depreciation: "30000", turnover: "1800000", grossReceipts: "950000", vehicles: "2", months: "10", declaredIncome: "160000", digital: true, manual: [] },
      capitalGains: {
        enabled: true,
        st: { rows: [{ id: uid(), description: "equity mutual fund units", sale: "180000", cost: "150000", improvement: "5000", expenses: "1000", equitySTT: true }] },
        lt: { rows: [{ id: uid(), description: "listed equity shares", sale: "500000", cost: "300000", improvement: "15000", expenses: "5000", equitySTT: true }] },
      },
      otherSources: { enabled: true, rows: [
        { id: uid(), type: "Interest on Savings Bank", amount: "8500" },
        { id: uid(), type: "Interest on FD", amount: "42000" },
        { id: uid(), type: "Dividend", amount: "15000" },
        { id: uid(), type: "Family Pension", amount: "36000" },
        { id: uid(), type: "Interest on Income Tax Refund", amount: "1200" },
      ] },
    },
    deductions: fullDeductions({
      "80C": { amount: "150000" }, "80CCD1B": { amount: "50000" }, "80CCD2": { amount: "60000" },
      "80D": { amount: "25000" }, "80DD": { amount: "125000", severe: true }, "80DDB": { amount: "100000", senior: true },
      "80E": { amount: "45000" }, "80EE": { amount: "50000" }, "80EEA": { amount: "150000" },
      "80G": { amount: "20000" }, "80GGA": { amount: "10000" }, "80TTA": { amount: "8500" }, "80U": { amount: "125000", severe: true },
    }),
    taxPaid: { tds: "95000", advanceTax: "40000", selfAssessmentTax: "5000", tcs: "3000", interest234A: "0", interest234B: "1800", interest234C: "1200", fee234F: "0", manual: [] },
  });

  // 2. Freelance professional (44ADA), self-occupied home, education loan
  entries.push({
    assessee: {
      ay: "AY 2026-27", fy: "2025-26",
      name: "Rohan Verma", pan: "BXKPV5678R", aadhaar: "345678912034",
      fatherName: "Suresh Verma", dob: "1988-11-02", gender: "Male",
      residentialStatus: "Resident", category: "Individual",
      address: "22, MG Road", state: "Karnataka", pin: "560001",
      email: "rohan.verma@example.com", mobile: "9845123456",
      bankName: "HDFC Bank", ifsc: "HDFC0000456", accountNumber: "50100234567",
    },
    income: {
      salary: { enabled: true, employerName: "Quantum Labs India Pvt. Ltd.", basic: "300000", da: "12000", hra: "90000", bonus: "20000", commission: "5000", perquisites: "8000", otherAllowances: "15000", professionalTax: "2400", entertainmentAllowance: "0", manual: [] },
      houseProperty: { enabled: true, type: "self", municipalValue: "180000", rentReceived: "0", municipalTax: "0", interest: "195000", manual: [] },
      business: { enabled: true, businessName: "Verma Design Consultancy", type: "44ADA", income: "0", expenses: "0", depreciation: "0", turnover: "0", grossReceipts: "1400000", vehicles: "0", months: "0", declaredIncome: "780000", digital: true, manual: [] },
      capitalGains: {
        enabled: true,
        st: { rows: [{ id: uid(), description: "equity mutual fund units", sale: "95000", cost: "70000", improvement: "0", expenses: "500", equitySTT: true }] },
        lt: { rows: [{ id: uid(), description: "residential plot", sale: "260000", cost: "150000", improvement: "8000", expenses: "2000", equitySTT: false }] },
      },
      otherSources: { enabled: true, rows: [
        { id: uid(), type: "Interest on Savings Bank", amount: "6200" },
        { id: uid(), type: "Interest on FD", amount: "18000" },
        { id: uid(), type: "Royalty", amount: "25000" },
      ] },
    },
    deductions: fullDeductions({
      "80C": { amount: "120000" }, "80CCD1B": { amount: "50000" }, "80CCD2": { amount: "36000" },
      "80D": { amount: "18000" }, "80E": { amount: "62000" }, "80EEA": { amount: "150000" },
      "80G": { amount: "12000" }, "80TTA": { amount: "6200" },
    }),
    taxPaid: { tds: "138000", advanceTax: "60000", selfAssessmentTax: "0", tcs: "2000", interest234A: "0", interest234B: "0", interest234C: "600", fee234F: "0", manual: [] },
  });

  // 3. Senior citizen with pension income and 80TTB
  entries.push({
    assessee: {
      ay: "AY 2026-27", fy: "2025-26",
      name: "Priya Mehta", pan: "CQRTM9012L", aadhaar: "456789120345",
      fatherName: "Late Devendra Mehta", dob: "1962-03-10", gender: "Female",
      residentialStatus: "Resident", category: "Individual",
      address: "7, Malabar Hill", state: "Maharashtra", pin: "400006",
      email: "priya.mehta@example.com", mobile: "9820098765",
      bankName: "ICICI Bank", ifsc: "ICIC0001789", accountNumber: "60723456891",
    },
    income: {
      salary: { enabled: true, employerName: "Retired — Pension Payer: LIC Jeevan Suraksha", basic: "480000", da: "0", hra: "0", bonus: "0", commission: "0", perquisites: "0", otherAllowances: "0", professionalTax: "0", entertainmentAllowance: "0", manual: [] },
      houseProperty: { enabled: true, type: "deemed", municipalValue: "210000", rentReceived: "195000", municipalTax: "9500", interest: "40000", manual: [] },
      business: { enabled: false, businessName: "", type: "normal", income: "", expenses: "", depreciation: "", turnover: "", grossReceipts: "", vehicles: "", months: "", declaredIncome: "", digital: false, manual: [] },
      capitalGains: {
        enabled: true,
        st: { rows: [] },
        lt: { rows: [{ id: uid(), description: "equity mutual fund units", sale: "620000", cost: "410000", improvement: "10000", expenses: "6000", equitySTT: true }] },
      },
      otherSources: { enabled: true, rows: [
        { id: uid(), type: "Interest on Savings Bank", amount: "14000" },
        { id: uid(), type: "Interest on FD", amount: "165000" },
        { id: uid(), type: "Dividend", amount: "48000" },
        { id: uid(), type: "Family Pension", amount: "96000" },
      ] },
    },
    deductions: fullDeductions({
      "80C": { amount: "90000" }, "80D": { amount: "50000" }, "80DDB": { amount: "100000", senior: true },
      "80G": { amount: "15000" }, "80TTB": { amount: "50000" },
    }),
    taxPaid: { tds: "42000", advanceTax: "15000", selfAssessmentTax: "8000", tcs: "0", interest234A: "0", interest234B: "0", interest234C: "0", fee234F: "0", manual: [] },
  });

  // 4. Small-business owner under 44AD (digital), let-out property, no salary
  entries.push({
    assessee: {
      ay: "AY 2026-27", fy: "2025-26",
      name: "Vikram Nair", pan: "DPLNV3456K", aadhaar: "567891203456",
      fatherName: "Gopalan Nair", dob: "1979-08-21", gender: "Male",
      residentialStatus: "Resident", category: "Individual",
      address: "3rd Cross, Anna Nagar", state: "Tamil Nadu", pin: "600040",
      email: "vikram.nair@example.com", mobile: "9884567123",
      bankName: "Axis Bank", ifsc: "UTIB0000234", accountNumber: "91802345671",
    },
    income: {
      salary: { enabled: false, employerName: "", basic: "", da: "", hra: "", bonus: "", commission: "", perquisites: "", otherAllowances: "", professionalTax: "", entertainmentAllowance: "", manual: [] },
      houseProperty: { enabled: true, type: "letout", municipalValue: "150000", rentReceived: "216000", municipalTax: "8000", interest: "95000", manual: [] },
      business: { enabled: true, businessName: "Nair General Store", type: "44AD", income: "0", expenses: "0", depreciation: "0", turnover: "3200000", grossReceipts: "0", vehicles: "0", months: "0", declaredIncome: "210000", digital: true, manual: [] },
      capitalGains: {
        enabled: true,
        st: { rows: [{ id: uid(), description: "gold ETF units", sale: "310000", cost: "240000", improvement: "4000", expenses: "1500", equitySTT: false }] },
        lt: { rows: [] },
      },
      otherSources: { enabled: true, rows: [
        { id: uid(), type: "Interest on Savings Bank", amount: "5100" },
        { id: uid(), type: "Interest on FD", amount: "22000" },
        { id: uid(), type: "Any Other", customLabel: "Freelance consulting fee", amount: "9000" },
      ] },
    },
    deductions: fullDeductions({
      "80C": { amount: "150000" }, "80CCD2": { amount: "0" }, "80D": { amount: "22000" },
      "80E": { amount: "0" }, "80EE": { amount: "0" }, "80G": { amount: "8000" }, "80TTA": { amount: "5100" },
    }),
    taxPaid: { tds: "18000", advanceTax: "55000", selfAssessmentTax: "12000", tcs: "4500", interest234A: "500", interest234B: "0", interest234C: "0", fee234F: "0", manual: [] },
  });

  // 5. HUF with business and rental income
  entries.push({
    assessee: {
      ay: "AY 2026-27", fy: "2025-26",
      name: "Sunita Iyer (HUF)", pan: "EHGFI7890M", aadhaar: "678912034567",
      fatherName: "N/A (HUF - Karta: Ganesh Iyer)", dob: "1975-01-15", gender: "Other",
      residentialStatus: "Resident", category: "HUF",
      address: "45, T Nagar", state: "Tamil Nadu", pin: "600017",
      email: "sunita.iyerhuf@example.com", mobile: "9789012345",
      bankName: "Canara Bank", ifsc: "CNRB0001122", accountNumber: "11023456789",
    },
    income: {
      salary: { enabled: false, employerName: "", basic: "", da: "", hra: "", bonus: "", commission: "", perquisites: "", otherAllowances: "", professionalTax: "", entertainmentAllowance: "", manual: [] },
      houseProperty: { enabled: true, type: "letout", municipalValue: "300000", rentReceived: "360000", municipalTax: "15000", interest: "220000", manual: [] },
      business: { enabled: true, businessName: "Iyer Family Enterprises (HUF)", type: "normal", income: "950000", expenses: "410000", depreciation: "60000", turnover: "", grossReceipts: "", vehicles: "", months: "", declaredIncome: "", digital: false, manual: [] },
      capitalGains: {
        enabled: true,
        st: { rows: [] },
        lt: { rows: [{ id: uid(), description: "commercial shop", sale: "890000", cost: "540000", improvement: "25000", expenses: "9000", equitySTT: false }] },
      },
      otherSources: { enabled: true, rows: [
        { id: uid(), type: "Interest on FD", amount: "58000" },
        { id: uid(), type: "Dividend", amount: "31000" },
        { id: uid(), type: "Agricultural Income", amount: "40000" },
      ] },
    },
    deductions: fullDeductions({
      "80C": { amount: "150000" }, "80D": { amount: "25000" }, "80G": { amount: "30000" }, "80GGA": { amount: "15000" },
    }),
    taxPaid: { tds: "72000", advanceTax: "90000", selfAssessmentTax: "10000", tcs: "6000", interest234A: "0", interest234B: "2400", interest234C: "900", fee234F: "0", manual: [] },
  });

  return entries;
}

const OTHER_INCOME_TYPES = [
  "Interest on Savings Bank", "Interest on FD", "Dividend", "Family Pension",
  "Lottery / Winnings", "Interest on Income Tax Refund", "Agricultural Income", "Royalty", "Gifts", "Any Other",
];

/* ============================================================
   SMALL UI PRIMITIVES
   ============================================================ */
function Card({ icon: Icon, title, subtitle, enabled, onToggle, children, defaultOpen = false, badge }) {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <div className="itc-card">
      <div className="itc-card-head" onClick={() => setOpen((o) => !o)}>
        <div className="itc-card-head-left">
          {onToggle && (
            <input
              type="checkbox"
              checked={enabled}
              onClick={(e) => e.stopPropagation()}
              onChange={(e) => onToggle(e.target.checked)}
              className="itc-check"
            />
          )}
          <div className="itc-card-icon"><Icon size={17} /></div>
          <div>
            <div className="itc-card-title">{title}</div>
            {subtitle && <div className="itc-card-subtitle">{subtitle}</div>}
          </div>
        </div>
        <div className="itc-card-head-right">
          {badge !== undefined && <span className="itc-badge">{badge}</span>}
          <ChevronDown size={16} className={"itc-chev" + (open ? " open" : "")} />
        </div>
      </div>
      {open && <div className="itc-card-body">{children}</div>}
    </div>
  );
}

function Field({ label, value, onChange, type = "number", suffix, prefix, placeholder, options }) {
  if (options) {
    return (
      <label className="itc-field">
        <span>{label}</span>
        <select value={value} onChange={(e) => onChange(e.target.value)}>
          {options.map((o) => <option key={o} value={o}>{yearOnly(o)}</option>)}
        </select>
      </label>
    );
  }
  const showPrefix = prefix !== undefined ? prefix : (type === "number" ? "₹" : null);
  return (
    <label className="itc-field">
      <span>{label}</span>
      <div className="itc-input-wrap">
        {showPrefix && <span className="itc-rupee">{showPrefix}</span>}
        <input
          type={type}
          value={value}
          placeholder={placeholder}
          onChange={(e) => onChange(e.target.value)}
          inputMode={type === "number" ? "decimal" : undefined}
        />
        {suffix && <span className="itc-suffix">{suffix}</span>}
      </div>
    </label>
  );
}

function Row({ children }) { return <div className="itc-row">{children}</div>; }

function LedgerLine({ label, value, bold, indent, positive, negative }) {
  return (
    <div className={"ledger-line" + (bold ? " bold" : "") + (indent ? " indent" : "")}>
      <span>{label}</span>
      <span className={"ledger-amt" + (positive ? " pos" : "") + (negative ? " neg" : "")}>{fmt(value)}</span>
    </div>
  );
}

/** Reusable "add your own line item" control — used across income heads,
 *  deductions, and tax-paid so any field the form doesn't anticipate can
 *  still be entered manually and will flow into totals and print. */
function ManualEntries({ items, onChange, label = "Add manual entry", showRegimeToggle }) {
  const rows = items || [];
  const dragIndex = React.useRef(null);
  const [overIndex, setOverIndex] = useState(null);
  const addRow = () => onChange([...rows, { id: uid(), label: "", amount: "", regime: "old" }]);
  const updateRow = (id, patch) => onChange(rows.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  const removeRow = (id) => onChange(rows.filter((r) => r.id !== id));

  const onDragStart = (idx) => (e) => { dragIndex.current = idx; e.dataTransfer.effectAllowed = "move"; };
  const onDragOver = (idx) => (e) => { e.preventDefault(); if (overIndex !== idx) setOverIndex(idx); };
  const onDrop = (idx) => (e) => {
    e.preventDefault();
    const from = dragIndex.current;
    setOverIndex(null);
    if (from === null || from === idx) return;
    onChange(moveItem(rows, from, idx));
    dragIndex.current = null;
  };
  const onDragEnd = () => { dragIndex.current = null; setOverIndex(null); };

  return (
    <div className="itc-subsection">
      <div className="itc-subsection-head">
        <span>Manual Entries</span>
        <button className="itc-add-row" onClick={addRow}><Plus size={13} /> {label}</button>
      </div>
      {rows.map((r, idx) => (
        <div
          key={r.id}
          className={"itc-manual-row" + (overIndex === idx ? " drag-over" : "")}
          draggable
          onDragStart={onDragStart(idx)}
          onDragOver={onDragOver(idx)}
          onDrop={onDrop(idx)}
          onDragEnd={onDragEnd}
        >
          <span className="itc-drag-handle" title="Drag to reorder"><GripVertical size={14} /></span>
          <input type="text" placeholder="Description (as it should appear on the print)" value={r.label} onChange={(e) => updateRow(r.id, { label: e.target.value })} />
          <div className="itc-input-wrap small"><span className="itc-rupee">₹</span><input type="number" value={r.amount} onChange={(e) => updateRow(r.id, { amount: e.target.value })} /></div>
          {showRegimeToggle && (
            <select value={r.regime || "old"} onChange={(e) => updateRow(r.id, { regime: e.target.value })}>
              <option value="old">Old regime only</option>
              <option value="both">Both regimes</option>
            </select>
          )}
          <button className="itc-icon-btn" onClick={() => removeRow(r.id)}><Trash2 size={14} /></button>
        </div>
      ))}
    </div>
  );
}

/* ============================================================
   MAIN APP
   ============================================================ */
function TaxComputationEditor({ initialRecord, clients, allComputations, onSave, onExit, onOpenClientEdit, onOpenRecord }) {
  const [nav, setNav] = useState("assessee");
  const [assessee, setAssessee] = useState(() => {
    const rawAssessee = initialRecord.assessee || defaultAssessee;
    if (rawAssessee.pan && clients) {
      const client = clients.find(c => c.pan === rawAssessee.pan);
      if (client) {
        return {
          ...rawAssessee,
          pan: client.pan || rawAssessee.pan,
          name: client.name || rawAssessee.name,
          address: client.addr || rawAssessee.address,
          state: client.state || rawAssessee.state,
          pin: client.pin || rawAssessee.pin,
          email: client.email || rawAssessee.email,
          mobile: client.mob || rawAssessee.mobile,
          aadhaar: client.aadhaar || rawAssessee.aadhaar,
          dob: client.dob || rawAssessee.dob,
          fatherName: client.fatherName || rawAssessee.fatherName,
          gender: client.gender || rawAssessee.gender,
          residentialStatus: client.residentialStatus || rawAssessee.residentialStatus,
          category: CLIENT_CATEGORY_MAP[client.type] || rawAssessee.category,
          bankName: client.bankName || rawAssessee.bankName,
          ifsc: client.ifsc || rawAssessee.ifsc,
          accountNumber: client.accountNumber || rawAssessee.accountNumber,
          itrType: client.itrType || rawAssessee.itrType,
        };
      }
    }
    return rawAssessee;
  });
  const [income, setIncome] = useState(initialRecord.income || defaultIncome);
  const [deductions, setDeductions] = useState(initialRecord.deductions || defaultDeductions);
  const [manualDeductions, setManualDeductions] = useState(initialRecord.manualDeductions || defaultManualDeductions);
  const [taxPaid, setTaxPaid] = useState(initialRecord.taxPaid || defaultTaxPaid);
  const [letterhead, setLetterhead] = useState(defaultLetterhead);
  const [pageSetup, setPageSetup] = useState(defaultPageSetup);
  const [printRegime, setPrintRegime] = useState(initialRecord.printRegime || "new");
  const [toast, setToast] = useState(null);
  const [errors, setErrors] = useState({});
  const [sideOpen, setSideOpen] = useState(true);
  const [openModal, setOpenModal] = useState(false);
  const [osearch, setOsearch] = useState("");

  // Real-time synchronization of assessee details when client profile changes
  useEffect(() => {
    if (assessee.pan && clients) {
      const client = clients.find(c => c.pan === assessee.pan);
      if (client) {
        setAssessee(prev => {
          const next = {
            ...prev,
            pan: client.pan || prev.pan,
            name: client.name || prev.name,
            address: client.addr || prev.address,
            state: client.state || prev.state,
            pin: client.pin || prev.pin,
            email: client.email || prev.email,
            mobile: client.mob || prev.mobile,
            aadhaar: client.aadhaar || prev.aadhaar,
            dob: client.dob || prev.dob,
            fatherName: client.fatherName || prev.fatherName,
            gender: client.gender || prev.gender,
            residentialStatus: client.residentialStatus || prev.residentialStatus,
            category: CLIENT_CATEGORY_MAP[client.type] || prev.category,
            bankName: client.bankName || prev.bankName,
            ifsc: client.ifsc || prev.ifsc,
            accountNumber: client.accountNumber || prev.accountNumber,
            itrType: client.itrType || prev.itrType,
          };
          const hasChanged = Object.keys(next).some(k => next[k] !== prev[k]);
          return hasChanged ? next : prev;
        });
      }
    }
  }, [clients, assessee.pan]);

  const config = TAX_CONFIG[assessee.ay] || TAX_CONFIG["AY 2026-27"];

  /* -------- letterhead (firm profile) persistence, shared across computations -------- */
  useEffect(() => {
    (async () => {
      try {
        const r = await window.storage.get("firmProfile", false);
        if (r && r.value) setLetterhead(JSON.parse(r.value));
      } catch { /* no saved profile yet */ }
    })();
  }, []);
  const saveLetterhead = async (patch) => {
    setLetterhead((prev) => {
      const next = { ...prev, ...patch };
      window.storage.set("firmProfile", JSON.stringify(next), false).catch(() => {});
      return next;
    });
  };
  const updateLetterheadLive = (patch) => {
    setLetterhead((prev) => ({ ...prev, ...patch }));
  };

  /* -------- page setup persistence, shared across computations like a printer preference -------- */
  useEffect(() => {
    (async () => {
      try {
        const r = await window.storage.get("pageSetup", false);
        if (r && r.value) {
          const saved = JSON.parse(r.value);
          setPageSetup({ ...defaultPageSetup, ...saved, piFields: { ...defaultPageSetup.piFields, ...(saved.piFields || {}) } });
        }
      } catch { /* no saved page setup yet */ }
    })();
  }, []);
  const savePageSetup = (patch) => {
    setPageSetup((prev) => {
      const next = { ...prev, ...patch };
      window.storage.set("pageSetup", JSON.stringify(next), false).catch(() => {});
      return next;
    });
  };

  const showToast = (msg, type = "ok") => {
    setToast({ msg, type });
    setTimeout(() => setToast(null), 2600);
  };

  /* -------- save back to the ITR Computation tab (per-client, per-FY list) -------- */
  const persistComputation = () => {
    const record = {
      id: initialRecord.id,
      fy: initialRecord.fy,
      pan: assessee.pan,
      name: assessee.name,
      assessee, income, deductions, manualDeductions, taxPaid, printRegime,
      savedAt: Date.now(),
    };
    onSave(record);
    showToast("Computation saved");
  };

  /* -------- calculation pipeline -------- */
  const calc = useMemo(() => {
    const sal = income.salary.enabled ? calculateSalary(income.salary) : { gross: 0, preStdDeduction: 0 };
    const hp = income.houseProperty.enabled ? calculateHouseProperty(income.houseProperty) : { income: 0 };
    const biz = income.business.enabled ? calculateBusinessIncome(income.business) : { income: 0 };
    const cg = income.capitalGains.enabled
      ? calculateCapitalGain(income.capitalGains, config)
      : { st: { special20: 0, slab: 0 }, lt: { equity: 0, other: 0, grandfatheredTax: 0, grandfatheredGain: 0, detail: [] }, total: 0 };
    const other = income.otherSources.enabled ? calculateOtherIncome(income.otherSources.rows) : { total: 0, savingsInterest: 0, fdInterest: 0 };

    const ageCategory = (assessee.category === "Individual" && assessee.residentialStatus === "Resident")
      ? getAgeCategory(assessee.dob, assessee.ay)
      : "normal";

    const buildForRegime = (regimeKey) => {
      const stdDed = config[regimeKey].standardDeduction;
      const salaryIncome = income.salary.enabled ? Math.max(sal.preStdDeduction - stdDed, 0) : 0;
      const hpIncomeRaw = income.houseProperty.enabled
        ? (income.houseProperty.type === "self" && regimeKey === "newRegime" ? 0 : hp.income)
        : 0;
      // Section 71(3A): house-property loss set off against other heads is
      // capped at ₹2,00,000 (regime-neutral); any excess carries forward.
      const hpLossCap = config.hpLossSetOffCap || 200000;
      const hpIncome = hpIncomeRaw < 0 ? Math.max(hpIncomeRaw, -hpLossCap) : hpIncomeRaw;
      const hpCarryForward = hpIncomeRaw < -hpLossCap ? -(hpIncomeRaw + hpLossCap) : 0;
      const slabIncomeCG = cg.st.slab;
      const totalIncomeExclSpecial = salaryIncome + hpIncome + biz.income + other.total + slabIncomeCG;

      // Chapter VI-A only for old regime (except 80CCD2)
      let chapterVIATotal = 0;
      const chapterVIABreakup = [];
      Object.entries(deductions).forEach(([code, d]) => {
        if (!d.enabled) return;
        const rule = config.chapterVIA[code];
        if (!rule) return;
        const appliesHere = rule.regime === "both" || (rule.regime === "old" && regimeKey === "oldRegime");
        if (!appliesHere) return;
        let cap = rule.cap;
        if (rule.capSevere && d.severe) cap = rule.capSevere;
        if (rule.capSenior && d.senior) cap = rule.capSenior;
        const amt = cap != null ? Math.min(num(d.amount), cap) : num(d.amount);
        chapterVIATotal += amt;
        chapterVIABreakup.push({ code, label: rule.label, amount: amt });
      });
      manualDeductions.forEach((m) => {
        const appliesHere = m.regime === "both" || regimeKey === "oldRegime";
        if (!appliesHere) return;
        const amt = num(m.amount);
        chapterVIATotal += amt;
        chapterVIABreakup.push({ code: "Other", label: m.label || "Other deduction", amount: amt });
      });
      const taxableExclSpecial = Math.max(totalIncomeExclSpecial - chapterVIATotal, 0);
      const isResidentIndividual = assessee.residentialStatus === "Resident" && assessee.category === "Individual";

      const result = computeRegimeTax({
        regimeKey,
        config,
        totalIncomeExclSpecial: taxableExclSpecial,
        stcgSpecial20: cg.st.special20,
        ltcgEquity: cg.lt.equity,
        ltcgOther: cg.lt.other,
        ltcgGrandfatheredTax: cg.lt.grandfatheredTax,
        ltcgGrandfatheredGain: cg.lt.grandfatheredGain,
        isEligibleRebate: isResidentIndividual,
        ageCategory,
      });

      return {
        salaryIncome, hpIncome, hpIncomeRaw, hpCarryForward, bizIncome: biz.income, otherIncome: other.total, slabIncomeCG,
        gti: totalIncomeExclSpecial + cg.st.special20 + cg.lt.equity + cg.lt.other + cg.lt.grandfatheredGain,
        chapterVIATotal, chapterVIABreakup, taxableExclSpecial, ageCategory, ...result,
      };
    };

    const oldR = buildForRegime("oldRegime");
    const newR = buildForRegime("newRegime");

    const qSum = num(taxPaid.advanceTaxQ1) + num(taxPaid.advanceTaxQ2) + num(taxPaid.advanceTaxQ3) + num(taxPaid.advanceTaxQ4);
    const usingQuarters = qSum > 0 || taxPaid.advanceTaxQ1 || taxPaid.advanceTaxQ2 || taxPaid.advanceTaxQ3 || taxPaid.advanceTaxQ4;
    const effectiveAdvanceTax = usingQuarters ? qSum : num(taxPaid.advanceTax);

    const manualPrepaidTotal = (taxPaid.manual || []).reduce((s, m) => s + num(m.amount), 0);
    const interestTotal = num(taxPaid.interest234A) + num(taxPaid.interest234B) + num(taxPaid.interest234C) + num(taxPaid.fee234F);
    const tdsTotal = num(taxPaid.tds) + effectiveAdvanceTax + num(taxPaid.selfAssessmentTax) + num(taxPaid.tcs) + manualPrepaidTotal;

    const withLiability = (r) => {
      const grossTaxLiability = r.total;
      const totalWithInterest = grossTaxLiability + interestTotal;
      // Section 288B: Net amount payable or refund due is rounded to the nearest multiple of ten
      const rawBalance = totalWithInterest - tdsTotal;
      const balance = Math.round(rawBalance / 10) * 10;
      const interestSuggestions = computeInterestSuggestions({
        assessedTax: grossTaxLiability,
        tdsAndTcs: num(taxPaid.tds) + num(taxPaid.tcs),
        advanceTaxByQuarter: { q1: taxPaid.advanceTaxQ1, q2: taxPaid.advanceTaxQ2, q3: taxPaid.advanceTaxQ3, q4: taxPaid.advanceTaxQ4 },
        selfAssessmentTax: taxPaid.selfAssessmentTax,
        filingDateStr: taxPaid.filingDate,
        dueDateStr: taxPaid.dueDate,
        totalIncome: r.totalTaxableIncome,
        ay: assessee.ay,
      });
      return {
        ...r, grossTaxLiability, interestTotal, totalWithInterest, effectiveAdvanceTax, interestSuggestions,
        taxPaidTotal: tdsTotal, refund: balance < 0 ? -balance : 0, payable: balance > 0 ? balance : 0,
      };
    };

    return { sal, hp, biz, cg, other, old: withLiability(oldR), new: withLiability(newR) };
  }, [income, deductions, manualDeductions, taxPaid, config, assessee.residentialStatus, assessee.category, assessee.dob, assessee.ay]);

  const betterRegime = calc.old.total <= calc.new.total ? "old" : "new";

  /* -------- validation -------- */
  useEffect(() => {
    const e = {};
    if (assessee.pan && !/^[A-Z]{5}[0-9]{4}[A-Z]$/.test(assessee.pan.toUpperCase())) e.pan = "PAN should look like ABCDE1234F";
    if (assessee.aadhaar && !/^\d{12}$/.test(assessee.aadhaar.replace(/\s/g, ""))) e.aadhaar = "Aadhaar should be 12 digits";
    if (assessee.ifsc && !/^[A-Z]{4}0[A-Z0-9]{6}$/.test(assessee.ifsc.toUpperCase())) e.ifsc = "IFSC looks invalid";
    if (assessee.email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(assessee.email)) e.email = "Enter a valid email";
    setErrors(e);
  }, [assessee.pan, assessee.aadhaar, assessee.ifsc, assessee.email]);

  const navItems = [
    { id: "assessee", label: "Assessee Details", icon: User },
    { id: "income", label: "Income Computation", icon: Wallet },
    { id: "deductions", label: "Deductions (Ch. VI-A)", icon: ShieldCheck },
    { id: "regime", label: "Tax Regime Comparison", icon: Scale },
    { id: "liability", label: "Tax Paid & Liability", icon: Receipt },
    { id: "sheet", label: "Computation Sheet", icon: FileText },
    { id: "letterhead", label: "Letterhead & Logo", icon: Settings },
  ];

  return (
    <div className="itc-app">
      <style>{CSS}</style>
      <style>{buildPrintCSS(pageSetup)}</style>

      {/* SIDEBAR */}
      <aside className="itc-sidebar no-print" style={sideOpen ? undefined : { width: 68, padding: "20px 8px" }}>
        <div className="itc-brand" style={{ justifyContent: sideOpen ? "flex-start" : "center" }}>
          <div className="itc-seal">IT</div>
          {sideOpen && (
            <div style={{ flex: 1, minWidth: 0 }}>
              <div className="itc-brand-title" style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{assessee.name || "New Assessee"}</div>
              <div className="itc-brand-sub" style={{ overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{assessee.pan || "PAN not set"} · FY {initialRecord.fy}</div>
            </div>
          )}
          <button onClick={() => setSideOpen((p) => !p)} title={sideOpen ? "Collapse sidebar" : "Expand sidebar"}
            style={{ background: "rgba(255,255,255,0.08)", border: "1px solid rgba(255,255,255,0.16)", color: "#E8EDF2", borderRadius: 6, width: 22, height: 22, cursor: "pointer", fontSize: 11, display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0, transition: "transform .2s", transform: sideOpen ? "none" : "rotate(180deg)" }}>◀</button>
        </div>

        {/* Actions moved to the top — Back / Open / Save / Print */}
        <div className="itc-sidebar-actions" style={{ gridTemplateColumns: sideOpen ? "1fr 1fr" : "1fr", marginTop: 0 }}>
          <button className="itc-sbtn" onClick={onExit} title="Back to List"><X size={14} />{sideOpen && " Back"}</button>
          <button className="itc-sbtn" onClick={() => setOpenModal(true)} title="Open a saved / draft computation"><FolderOpen size={14} />{sideOpen && " Open"}</button>
          <button className="itc-sbtn primary" onClick={persistComputation} title="Save"><Save size={14} />{sideOpen && " Save"}</button>
          <button className="itc-sbtn" onClick={() => {
            setNav("sheet");
            const oldTitle = document.title;
            document.title = `ITR COMPUTATION FY ${assessee.fy} ${assessee.name || "Computation"}`;
            setTimeout(() => {
              window.print();
              setTimeout(() => { document.title = oldTitle; }, 1000);
            }, 200);
          }} title="Print"><Printer size={14} />{sideOpen && " Print"}</button>
        </div>

        {sideOpen && <div className="itc-ay-badge">{yearOnly(assessee.ay)} <span>({yearOnly(config.fy)})</span></div>}

        <nav className="itc-nav">
          {navItems.map((it) => (
            <button key={it.id} className={"itc-nav-item" + (nav === it.id ? " active" : "")} onClick={() => setNav(it.id)} title={!sideOpen ? it.label : ""} style={{ justifyContent: sideOpen ? "flex-start" : "center" }}>
              <it.icon size={16} />{sideOpen && <span>{it.label}</span>}
            </button>
          ))}
        </nav>
      </aside>

      {/* MAIN */}
      <main className="itc-main">
        {nav === "assessee" && (
          <AssesseeView assessee={assessee} setAssessee={setAssessee} errors={errors} ayOptions={Object.keys(TAX_CONFIG)} clients={clients} onOpenClientEdit={onOpenClientEdit} />
        )}
        {nav === "income" && (
          <IncomeView income={income} setIncome={setIncome} calc={calc} />
        )}
        {nav === "deductions" && (
          <DeductionsView deductions={deductions} setDeductions={setDeductions} manualDeductions={manualDeductions} setManualDeductions={setManualDeductions} config={config} old={calc.old} />
        )}
        {nav === "regime" && (
          <RegimeView calc={calc} betterRegime={betterRegime} printRegime={printRegime} setPrintRegime={setPrintRegime} />
        )}
        {nav === "liability" && (
          <LiabilityView taxPaid={taxPaid} setTaxPaid={setTaxPaid} calc={calc} printRegime={printRegime} />
        )}
        {nav === "sheet" && (
          <SheetView assessee={assessee} income={income} deductions={deductions} manualDeductions={manualDeductions} config={config} calc={calc} printRegime={printRegime} setPrintRegime={setPrintRegime} taxPaid={taxPaid} letterhead={letterhead} saveLetterhead={saveLetterhead} updateLetterheadLive={updateLetterheadLive} pageSetup={pageSetup} savePageSetup={savePageSetup} onOpenClientEdit={onOpenClientEdit} />
        )}
        {nav === "letterhead" && (
          <LetterheadView letterhead={letterhead} saveLetterhead={saveLetterhead} updateLetterheadLive={updateLetterheadLive} />
        )}
      </main>

      {openModal && (() => {
        const items = (allComputations || []).filter((r) => {
          if (!osearch) return true;
          const lq = osearch.toLowerCase();
          return (r.name || r.assessee?.name || "").toLowerCase().includes(lq) || (r.pan || "").toLowerCase().includes(lq);
        }).sort((a, b) => (b.savedAt || 0) - (a.savedAt || 0));
        return (
          <div className="no-print" style={{ position: "fixed", inset: 0, background: "#000C", zIndex: 7000, display: "flex", alignItems: "center", justifyContent: "center", padding: 20 }}>
            <div style={{ background: G.surf, border: `1px solid ${G.bdr}`, borderRadius: 16, width: "min(560px,96%)", maxHeight: "82vh", display: "flex", flexDirection: "column", boxShadow: "0 24px 70px #000C" }}>
              <div style={{ padding: "16px 20px", borderBottom: `1px solid ${G.bdr}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                <span style={{ fontWeight: 800, fontSize: 15, color: G.wh }}>📂 Open Computation</span>
                <button onClick={() => { setOpenModal(false); setOsearch(""); }} style={{ background: "transparent", border: `1px solid ${G.bdr}`, color: G.mut, borderRadius: 8, padding: "4px 11px", cursor: "pointer" }}>✕</button>
              </div>
              <div style={{ padding: "14px 20px 0" }}>
                <div style={{ position: "relative" }}>
                  <span style={{ position: "absolute", left: 11, top: "50%", transform: "translateY(-50%)", color: G.mut }}>🔍</span>
                  <input value={osearch} onChange={(e) => setOsearch(e.target.value)} placeholder="Search by name or PAN..." style={{ ...IS, paddingLeft: 34 }} autoFocus />
                </div>
              </div>
              <div style={{ padding: 16, overflow: "auto", display: "flex", flexDirection: "column", gap: 8 }}>
                {items.length === 0 && <div style={{ textAlign: "center", color: G.mut, fontSize: 13, padding: 24 }}>No computations found.</div>}
                {items.map((r) => (
                  <div key={r.id} onClick={() => { onOpenRecord(r); setOpenModal(false); setOsearch(""); }}
                    style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, padding: "10px 13px", background: r.id === initialRecord.id ? G.green + "14" : G.bg, border: `1px solid ${r.id === initialRecord.id ? G.green + "55" : G.bdr}`, borderRadius: 10, cursor: "pointer" }}>
                    <div style={{ minWidth: 0 }}>
                      <div style={{ fontWeight: 700, fontSize: 13, color: G.wh, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{r.name || r.assessee?.name || "Untitled"}{r.id === initialRecord.id ? " (current)" : ""}</div>
                      <div style={{ fontSize: 11, color: G.cyn, fontFamily: "monospace", fontWeight: 700 }}>{r.pan || "No PAN"} <span style={{ color: G.mut, fontFamily: "inherit", fontWeight: 500 }}>· FY {r.fy} · {r.assessee?.ay}</span></div>
                    </div>
                    <div style={{ fontSize: 10, color: G.mut, whiteSpace: "nowrap", flexShrink: 0 }}>{r.savedAt ? new Date(r.savedAt).toLocaleDateString("en-IN", { dateStyle: "medium" }) : "Draft"}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        );
      })()}

      {toast && <div className={"itc-toast " + toast.type + " no-print"}>{toast.type === "ok" ? <CheckCircle2 size={15} /> : <AlertCircle size={15} />} {toast.msg}</div>}
    </div>
  );
}
function Modal({ children, onClose, title, wide }) {
  return (
    <div className="itc-modal-backdrop no-print" onClick={onClose}>
      <div className={"itc-modal" + (wide ? " wide" : "")} onClick={(e) => e.stopPropagation()}>
        <div className="itc-modal-head">
          <h3>{title}</h3>
          <button onClick={onClose}><X size={16} /></button>
        </div>
        {children}
      </div>
    </div>
  );
}

/* ============================================================
   ASSESSEE VIEW
   ============================================================ */
function AssesseeView({ assessee, setAssessee, errors, ayOptions, clients, onOpenClientEdit }) {
  const set = (k) => (v) => setAssessee((a) => ({ ...a, [k]: v }));
  const linkedClient = clients && assessee.pan ? clients.find((c) => c.pan === assessee.pan) : null;
  return (
    <div className="itc-page">
      <PageHeader title="Assessee Details" subtitle="Personal and contact information for this assessment" />
      <div className="itc-card static">
        <div className="itc-section-label">Assessment</div>
        <Row>
          <Field label="Assessment Year" value={assessee.ay} onChange={(v) => { set("ay")(v); }} options={ayOptions} />
          <Field label="Financial Year" value={assessee.fy} onChange={set("fy")} type="text" />
        </Row>
        <Row><Field label="ITR Type" value={assessee.itrType} onChange={set("itrType")} options={["ITR-1", "ITR-2", "ITR-3", "ITR-4"]} /></Row>
        <div className="itc-section-label">Personal Information</div>
        <Row>
          <Field label="Name" value={assessee.name} onChange={set("name")} type="text" placeholder="Full legal name" />
          <Field label="Father's Name" value={assessee.fatherName} onChange={set("fatherName")} type="text" />
        </Row>
        <Row>
          <Field label="PAN" value={assessee.pan} onChange={(v) => set("pan")(v.toUpperCase())} type="text" placeholder="ABCDE1234F" />
          <Field label="Aadhaar" value={assessee.aadhaar} onChange={set("aadhaar")} type="text" placeholder="12-digit number" />
        </Row>
        {linkedClient && onOpenClientEdit && (
          <div className="itc-linked-client-row">
            <span>Linked client record —</span>
            <a className="pan-hyperlink" onClick={() => onOpenClientEdit(assessee.pan)}>{assessee.pan} · view / edit client details ↗</a>
          </div>
        )}
        {(errors.pan || errors.aadhaar) && (
          <div className="itc-error-line">{errors.pan && <span><AlertCircle size={13} /> {errors.pan}</span>}{errors.aadhaar && <span><AlertCircle size={13} /> {errors.aadhaar}</span>}</div>
        )}
        <Row>
          <Field label="Date of Birth" value={assessee.dob} onChange={set("dob")} type="date" />
          <Field label="Gender" value={assessee.gender} onChange={set("gender")} options={["Male", "Female", "Other"]} />
        </Row>
        <Row>
          <Field label="Residential Status" value={assessee.residentialStatus} onChange={set("residentialStatus")} options={["Resident", "Resident but Not Ordinarily Resident", "Non-Resident"]} />
          <Field label="Category" value={assessee.category} onChange={set("category")} options={["Individual", "HUF", "Firm", "Company", "LLP", "Trust"]} />
        </Row>
        <div className="itc-section-label">Address</div>
        <Row><Field label="Address" value={assessee.address} onChange={set("address")} type="text" /></Row>
        <Row>
          <Field label="State" value={assessee.state} onChange={set("state")} type="text" />
          <Field label="PIN Code" value={assessee.pin} onChange={set("pin")} type="text" />
        </Row>
        <div className="itc-section-label">Contact</div>
        <Row>
          <Field label="Email" value={assessee.email} onChange={set("email")} type="email" />
          <Field label="Mobile" value={assessee.mobile} onChange={set("mobile")} type="tel" />
        </Row>
        {errors.email && <div className="itc-error-line"><span><AlertCircle size={13} /> {errors.email}</span></div>}
        <div className="itc-section-label">Bank Details (for refund)</div>
        <Row>
          <Field label="Bank Name" value={assessee.bankName} onChange={set("bankName")} type="text" />
          <Field label="IFSC" value={assessee.ifsc} onChange={(v) => set("ifsc")(v.toUpperCase())} type="text" placeholder="SBIN0001234" />
        </Row>
        {errors.ifsc && <div className="itc-error-line"><span><AlertCircle size={13} /> {errors.ifsc}</span></div>}
        <Row><Field label="Account Number" value={assessee.accountNumber} onChange={set("accountNumber")} type="text" /></Row>
      </div>
    </div>
  );
}
function PageHeader({ title, subtitle, right }) {
  return (
    <div className="itc-page-head">
      <div><h1>{title}</h1>{subtitle && <p>{subtitle}</p>}</div>
      {right}
    </div>
  );
}

function LogoEditor({ image, size, x, y, onLiveChange, onCommitChange, boxWidth = 200, boxHeight = 110 }) {
  const dragMode = React.useRef(null); // "move" | "resize"
  const start = React.useRef({ mx: 0, my: 0, size: 64, x: 0, y: 0 });
  const latest = React.useRef({ size, x, y });

  const onMouseMove = React.useCallback((e) => {
    if (!dragMode.current) return;
    const dx = e.clientX - start.current.mx;
    const dy = e.clientY - start.current.my;
    let next;
    if (dragMode.current === "resize") {
      next = { ...latest.current, size: Math.max(28, Math.min(140, start.current.size + dx)) };
    } else {
      next = {
        ...latest.current,
        x: Math.max(-70, Math.min(70, start.current.x + dx)),
        y: Math.max(-40, Math.min(40, start.current.y + dy)),
      };
    }
    latest.current = next;
    onLiveChange(next);
  }, [onLiveChange]);

  const onMouseUp = React.useCallback(() => {
    dragMode.current = null;
    window.removeEventListener("mousemove", onMouseMove);
    window.removeEventListener("mouseup", onMouseUp);
    onCommitChange(latest.current);
  }, [onMouseMove, onCommitChange]);

  const startDrag = (mode) => (e) => {
    e.preventDefault();
    e.stopPropagation();
    dragMode.current = mode;
    start.current = { mx: e.clientX, my: e.clientY, size, x, y };
    latest.current = { size, x, y };
    window.addEventListener("mousemove", onMouseMove);
    window.addEventListener("mouseup", onMouseUp);
  };

  return (
    <div className="itc-logo-editor-box" style={{ width: boxWidth, height: boxHeight }}>
      {image ? (
        <div
          className="itc-logo-editor-frame"
          style={{ width: size, height: size, transform: `translate(-50%, -50%) translate(${x}px, ${y}px)` }}
          onMouseDown={startDrag("move")}
          title="Drag to move"
        >
          <img src={image} alt="Logo" draggable={false} />
          <div className="itc-logo-resize-handle" onMouseDown={startDrag("resize")} title="Drag to resize" />
        </div>
      ) : (
        <span className="itc-logo-editor-empty">No logo</span>
      )}
    </div>
  );
}

function PageSetupPanel({ pageSetup, savePageSetup, open, setOpen }) {
  const set = (k) => (v) => savePageSetup({ [k]: v });
  const togglePI = (key) => (checked) => savePageSetup({ piFields: { ...pageSetup.piFields, [key]: checked } });
  return (
    <div className="itc-pagesetup-panel no-print">
      <button className="itc-pagesetup-toggle" onClick={() => setOpen(!open)}>
        <LayoutGrid size={14} /> Page Setup
        <ChevronDown size={14} className={"itc-chev" + (open ? " open" : "")} />
      </button>
      {open && (
        <div className="itc-pagesetup-body">
          <div className="itc-pagesetup-row">
            <Field label="Paper Size" value={pageSetup.paperSize} onChange={set("paperSize")} options={Object.keys(PAPER_SIZES)} />
            <Field label="Orientation" value={pageSetup.orientation} onChange={set("orientation")} options={["portrait", "landscape"]} />
            <Field label="Scale (%)" value={pageSetup.scale} onChange={set("scale")} prefix="" suffix="%" />
          </div>
          <div className="itc-pagesetup-row">
            <Field label="Margin Top (mm)" value={pageSetup.marginTop} onChange={set("marginTop")} prefix="" />
            <Field label="Margin Bottom (mm)" value={pageSetup.marginBottom} onChange={set("marginBottom")} prefix="" />
            <Field label="Margin Left (mm)" value={pageSetup.marginLeft} onChange={set("marginLeft")} prefix="" />
            <Field label="Margin Right (mm)" value={pageSetup.marginRight} onChange={set("marginRight")} prefix="" />
          </div>
          <div className="itc-note">Adjust and watch the page below update live — this is exactly what will print. Saved automatically as your printer preference for every computation.</div>

          <div className="itc-section-label">Personal Information — Fields to Print</div>
          <label className="itc-checkbox-row" style={{ marginBottom: 8 }}>
            <input type="checkbox" checked={!!pageSetup.piTwoColumn} onChange={(e) => savePageSetup({ piTwoColumn: e.target.checked })} />
            Split into two side-by-side columns to save space
          </label>
          <div className="itc-pi-checklist">
            {PI_FIELD_DEFS.map((f) => (
              <label key={f.key} className="itc-checkbox-row">
                <input type="checkbox" checked={pageSetup.piFields[f.key] !== false} onChange={(e) => togglePI(f.key)(e.target.checked)} />
                {f.label}
              </label>
            ))}
          </div>
        </div>
      )}
    </div>
  );
}

function BannerEditor({ image, widthPct, height, onLiveChange, onCommitChange }) {
  const dragging = React.useRef(false);
  const start = React.useRef({ mx: 0, my: 0, widthPct: 100, height: 70 });
  const latest = React.useRef({ widthPct, height });

  const onMouseMove = React.useCallback((e) => {
    if (!dragging.current) return;
    const dx = e.clientX - start.current.mx;
    const dy = e.clientY - start.current.my;
    const next = {
      widthPct: Math.max(20, Math.min(100, start.current.widthPct + dx / 3)),
      height: Math.max(24, Math.min(260, start.current.height + dy)),
    };
    latest.current = next;
    onLiveChange(next);
  }, [onLiveChange]);

  const onMouseUp = React.useCallback(() => {
    if (!dragging.current) return;
    dragging.current = false;
    window.removeEventListener("mousemove", onMouseMove);
    window.removeEventListener("mouseup", onMouseUp);
    onCommitChange(latest.current);
  }, [onMouseMove, onCommitChange]);

  const startDrag = (e) => {
    e.preventDefault();
    dragging.current = true;
    start.current = { mx: e.clientX, my: e.clientY, widthPct, height };
    latest.current = { widthPct, height };
    window.addEventListener("mousemove", onMouseMove);
    window.addEventListener("mouseup", onMouseUp);
  };

  return (
    <div className="itc-banner-editor-wrap">
      <div className="itc-banner-editor-box" style={{ height, width: `${widthPct}%` }}>
        {image ? <img src={image} alt="Banner" /> : <span className="itc-logo-editor-empty">No banner</span>}
        {image && <div className="itc-banner-resize-handle" onMouseDown={startDrag} title="Drag to resize" />}
      </div>
    </div>
  );
}

function LogoPanel({ letterhead, saveLetterhead, updateLetterheadLive, open, setOpen }) {
  const fileRef = React.useRef(null);
  const footerFileRef = React.useRef(null);
  const bannerFileRef = React.useRef(null);
  const onFile = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => saveLetterhead({ logo: reader.result });
    reader.readAsDataURL(file);
  };
  const onFooterFile = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => saveLetterhead({ footerLogo: reader.result });
    reader.readAsDataURL(file);
  };
  const onBannerFile = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => saveLetterhead({ banner: reader.result });
    reader.readAsDataURL(file);
  };
  return (
    <div className="itc-pagesetup-panel no-print">
      <button className="itc-pagesetup-toggle" onClick={() => setOpen(!open)}>
        <FileText size={14} /> Logo & Banner
        <ChevronDown size={14} className={"itc-chev" + (open ? " open" : "")} />
      </button>
      {open && (
        <div className="itc-pagesetup-body">
          <div className="itc-logo-panel-row">
            <div>
              <div className="itc-section-label" style={{ marginTop: 0 }}>Header Logo</div>
              <div className="itc-logo-upload">
                <LogoEditor
                  image={letterhead.logo}
                  size={letterhead.logoSize || 64}
                  x={letterhead.logoX || 0}
                  y={letterhead.logoY || 0}
                  boxWidth={160} boxHeight={90}
                  onLiveChange={(v) => updateLetterheadLive({ logoSize: v.size, logoX: v.x, logoY: v.y })}
                  onCommitChange={(v) => saveLetterhead({ logoSize: v.size, logoX: v.x, logoY: v.y })}
                />
                <div className="itc-logo-upload-actions">
                  <input ref={fileRef} type="file" accept="image/*" style={{ display: "none" }} onChange={onFile} />
                  <button className="itc-btn primary" onClick={() => fileRef.current.click()}>Upload</button>
                  {letterhead.logo && <button className="itc-btn ghost" onClick={() => saveLetterhead({ logo: null })}>Remove</button>}
                </div>
              </div>
            </div>
            <div>
              <div className="itc-section-label" style={{ marginTop: 0 }}>Footer Logo</div>
              <div className="itc-logo-upload">
                <LogoEditor
                  image={letterhead.footerLogo}
                  size={letterhead.footerLogoSize || 48}
                  x={letterhead.footerLogoX || 0}
                  y={letterhead.footerLogoY || 0}
                  boxWidth={160} boxHeight={90}
                  onLiveChange={(v) => updateLetterheadLive({ footerLogoSize: v.size, footerLogoX: v.x, footerLogoY: v.y })}
                  onCommitChange={(v) => saveLetterhead({ footerLogoSize: v.size, footerLogoX: v.x, footerLogoY: v.y })}
                />
                <div className="itc-logo-upload-actions">
                  <input ref={footerFileRef} type="file" accept="image/*" style={{ display: "none" }} onChange={onFooterFile} />
                  <button className="itc-btn primary" onClick={() => footerFileRef.current.click()}>Upload</button>
                  {letterhead.footerLogo && <button className="itc-btn ghost" onClick={() => saveLetterhead({ footerLogo: null })}>Remove</button>}
                </div>
              </div>
            </div>
          </div>
          <div className="itc-note">Drag a logo to reposition it, or its corner handle to resize — the logo's own box grows, text position never moves.</div>

          <div className="itc-section-label">Banner (Last Page)</div>
          <div className="itc-logo-upload">
            <div style={{ flex: 1 }}>
              <BannerEditor
                image={letterhead.banner}
                widthPct={letterhead.bannerWidthPct != null ? letterhead.bannerWidthPct : 100}
                height={letterhead.bannerHeight || 70}
                onLiveChange={(v) => updateLetterheadLive({ bannerWidthPct: v.widthPct, bannerHeight: v.height })}
                onCommitChange={(v) => saveLetterhead({ bannerWidthPct: v.widthPct, bannerHeight: v.height })}
              />
            </div>
            <div className="itc-logo-upload-actions">
              <input ref={bannerFileRef} type="file" accept="image/*" style={{ display: "none" }} onChange={onBannerFile} />
              <button className="itc-btn primary" onClick={() => bannerFileRef.current.click()}>Upload</button>
              {letterhead.banner && <button className="itc-btn ghost" onClick={() => saveLetterhead({ banner: null })}>Remove</button>}
            </div>
          </div>
          {letterhead.banner && <div className="itc-note">Drag the corner handle on the banner to resize it — appears at the bottom of the last (Tax Liability) page.</div>}
        </div>
      )}
    </div>
  );
}

function LetterheadView({ letterhead, saveLetterhead, updateLetterheadLive }) {
  const fileRef = React.useRef(null);
  const footerFileRef = React.useRef(null);
  const bannerFileRef = React.useRef(null);
  const onFile = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => saveLetterhead({ logo: reader.result });
    reader.readAsDataURL(file);
  };
  const onFooterFile = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => saveLetterhead({ footerLogo: reader.result });
    reader.readAsDataURL(file);
  };
  const onBannerFile = (e) => {
    const file = e.target.files && e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = () => saveLetterhead({ banner: reader.result });
    reader.readAsDataURL(file);
  };
  return (
    <div className="itc-page">
      <PageHeader title="Letterhead & Logo" subtitle="Branding used on the printed computation sheet — saved once and reused for every computation" />
      <div className="itc-card static">
        <div className="itc-section-label">Header Logo</div>
        <div className="itc-logo-upload">
          <LogoEditor
            image={letterhead.logo}
            size={letterhead.logoSize || 64}
            x={letterhead.logoX || 0}
            y={letterhead.logoY || 0}
            onLiveChange={(v) => updateLetterheadLive({ logoSize: v.size, logoX: v.x, logoY: v.y })}
            onCommitChange={(v) => saveLetterhead({ logoSize: v.size, logoX: v.x, logoY: v.y })}
          />
          <div className="itc-logo-upload-actions">
            <input ref={fileRef} type="file" accept="image/*" style={{ display: "none" }} onChange={onFile} />
            <button className="itc-btn primary" onClick={() => fileRef.current.click()}>Upload Logo</button>
            {letterhead.logo && <button className="itc-btn ghost" onClick={() => saveLetterhead({ logo: null })}>Remove</button>}
          </div>
        </div>
        {letterhead.logo && <div className="itc-note">Drag the logo itself to reposition it, or the small handle at its corner to resize — the print header updates to match, text position never moves.</div>}

        <div className="itc-section-label">Footer Logo</div>
        <div className="itc-logo-upload">
          <LogoEditor
            image={letterhead.footerLogo}
            size={letterhead.footerLogoSize || 48}
            x={letterhead.footerLogoX || 0}
            y={letterhead.footerLogoY || 0}
            onLiveChange={(v) => updateLetterheadLive({ footerLogoSize: v.size, footerLogoX: v.x, footerLogoY: v.y })}
            onCommitChange={(v) => saveLetterhead({ footerLogoSize: v.size, footerLogoX: v.x, footerLogoY: v.y })}
          />
          <div className="itc-logo-upload-actions">
            <input ref={footerFileRef} type="file" accept="image/*" style={{ display: "none" }} onChange={onFooterFile} />
            <button className="itc-btn primary" onClick={() => footerFileRef.current.click()}>Upload Footer Logo</button>
            {letterhead.footerLogo && <button className="itc-btn ghost" onClick={() => saveLetterhead({ footerLogo: null })}>Remove</button>}
          </div>
        </div>
        {letterhead.footerLogo && <div className="itc-note">Shown on the left of the footer, with your contact line and page details filling the rest of the space.</div>}

        <div className="itc-section-label">Banner (Last Page)</div>
        <div className="itc-logo-upload">
          <div style={{ flex: 1, maxWidth: 360 }}>
            <BannerEditor
              image={letterhead.banner}
              widthPct={letterhead.bannerWidthPct != null ? letterhead.bannerWidthPct : 100}
              height={letterhead.bannerHeight || 70}
              onLiveChange={(v) => updateLetterheadLive({ bannerWidthPct: v.widthPct, bannerHeight: v.height })}
              onCommitChange={(v) => saveLetterhead({ bannerWidthPct: v.widthPct, bannerHeight: v.height })}
            />
          </div>
          <div className="itc-logo-upload-actions">
            <input ref={bannerFileRef} type="file" accept="image/*" style={{ display: "none" }} onChange={onBannerFile} />
            <button className="itc-btn primary" onClick={() => bannerFileRef.current.click()}>Upload Banner</button>
            {letterhead.banner && <button className="itc-btn ghost" onClick={() => saveLetterhead({ banner: null })}>Remove</button>}
          </div>
        </div>
        {letterhead.banner && <div className="itc-note">Drag the corner handle on the banner to resize it — appears at the bottom of the last (Tax Liability) page.</div>}

        <div className="itc-section-label">Firm Details</div>
        <Row>
          <Field label="Firm / Practice Name" value={letterhead.firmName} onChange={(v) => saveLetterhead({ firmName: v })} type="text" placeholder="e.g. Fin-Tax Mitra" />
          <Field label="Tagline" value={letterhead.tagline} onChange={(v) => saveLetterhead({ tagline: v })} type="text" placeholder="e.g. Making Taxes Easy" />
        </Row>
        <Row><Field label="Contact Line (shown in footer)" value={letterhead.contact} onChange={(v) => saveLetterhead({ contact: v })} type="text" placeholder="e.g. care@firm.com | +91 XXXXXXXXXX" /></Row>
        <div className="itc-section-label">Footer Advertisement (last page only)</div>
        <Row><Field label="Services / Promotional Line" value={letterhead.footerAd} onChange={(v) => saveLetterhead({ footerAd: v })} type="text" placeholder="e.g. GST Registration & Filing | TDS Return | Balance Sheet | ITR | MSME" /></Row>
        <div className="itc-note">This is saved separately from your computations, so it carries over automatically the next time you print — no need to re-enter it for every assessee.</div>
      </div>
    </div>
  );
}

/* ============================================================
   INCOME VIEW
   ============================================================ */
function IncomeView({ income, setIncome, calc }) {
  const upd = (section) => (patch) => setIncome((i) => ({ ...i, [section]: { ...i[section], ...patch } }));
  const toggle = (section) => (v) => upd(section)({ enabled: v });

  return (
    <div className="itc-page">
      <PageHeader title="Income Computation" subtitle="Enable and fill in each applicable income head" />

      {/* SALARY */}
      <Card icon={Wallet} title="Income from Salary" subtitle="Basic, allowances, perquisites" enabled={income.salary.enabled} onToggle={toggle("salary")} defaultOpen badge={income.salary.enabled ? fmt(Math.max(calc.sal.preStdDeduction, 0)) : undefined}>
        <Row><Field label="Employer Name" value={income.salary.employerName} onChange={(v) => upd("salary")({ employerName: v })} type="text" placeholder="e.g. Randstad India Private Limited" /></Row>
        <Row>
          <Field label="Basic Salary" value={income.salary.basic} onChange={(v) => upd("salary")({ basic: v })} />
          <Field label="Dearness Allowance (DA)" value={income.salary.da} onChange={(v) => upd("salary")({ da: v })} />
        </Row>
        <Row>
          <Field label="HRA" value={income.salary.hra} onChange={(v) => upd("salary")({ hra: v })} />
          <Field label="Bonus" value={income.salary.bonus} onChange={(v) => upd("salary")({ bonus: v })} />
        </Row>
        <Row>
          <Field label="Commission" value={income.salary.commission} onChange={(v) => upd("salary")({ commission: v })} />
          <Field label="Perquisites" value={income.salary.perquisites} onChange={(v) => upd("salary")({ perquisites: v })} />
        </Row>
        <Row><Field label="Other Allowances" value={income.salary.otherAllowances} onChange={(v) => upd("salary")({ otherAllowances: v })} /></Row>
        <ManualEntries items={income.salary.manual} onChange={(rows) => upd("salary")({ manual: rows })} label="Add salary line item" />
        <div className="itc-section-label">Deductions from Salary</div>
        <Row>
          <Field label="Professional Tax" value={income.salary.professionalTax} onChange={(v) => upd("salary")({ professionalTax: v })} />
          <Field label="Entertainment Allowance (Govt. employees)" value={income.salary.entertainmentAllowance} onChange={(v) => upd("salary")({ entertainmentAllowance: v })} />
        </Row>
        <div className="itc-note">Standard deduction (₹75,000 new / ₹50,000 old regime) is applied automatically per regime.</div>
        <LedgerLine label="Salary income before standard deduction" value={calc.sal.preStdDeduction} bold />
      </Card>

      {/* HOUSE PROPERTY */}
      <Card icon={Home} title="Income from House Property" subtitle="Self-occupied, let-out, or deemed let-out" enabled={income.houseProperty.enabled} onToggle={toggle("houseProperty")} badge={income.houseProperty.enabled ? fmt(calc.hp.income) : undefined}>
        <Row><Field label="Property Type" value={income.houseProperty.type} onChange={(v) => upd("houseProperty")({ type: v })} options={["self", "letout", "deemed"]} /></Row>
        {income.houseProperty.type !== "self" && (
          <Row>
            <Field label="Municipal Value" value={income.houseProperty.municipalValue} onChange={(v) => upd("houseProperty")({ municipalValue: v })} />
            <Field label="Rent Received" value={income.houseProperty.rentReceived} onChange={(v) => upd("houseProperty")({ rentReceived: v })} />
          </Row>
        )}
        {income.houseProperty.type !== "self" && (
          <Row><Field label="Municipal Tax Paid" value={income.houseProperty.municipalTax} onChange={(v) => upd("houseProperty")({ municipalTax: v })} /></Row>
        )}
        <Row><Field label="Interest on Housing Loan" value={income.houseProperty.interest} onChange={(v) => upd("houseProperty")({ interest: v })} suffix={income.houseProperty.type === "self" ? "capped ₹2,00,000" : "no cap"} /></Row>
        <ManualEntries items={income.houseProperty.manual} onChange={(rows) => upd("houseProperty")({ manual: rows })} label="Add line item" />
        {income.houseProperty.type !== "self" && <div className="itc-note">Standard deduction of 30% of Net Annual Value is applied automatically.</div>}
        {income.houseProperty.type === "self" && <div className="itc-note">Interest deduction for self-occupied property applies under the old regime only.</div>}
        <LedgerLine label="Income from House Property" value={calc.hp.income} bold negative={calc.hp.income < 0} />
        {calc.old.hpCarryForward > 0 && (
          <div className="itc-note">
            Loss set off against other heads is capped at ₹2,00,000 (Section 71(3A)); the remaining {fmt(calc.old.hpCarryForward)} carries forward to future years and isn't reflected in this year's tax.
          </div>
        )}
      </Card>

      {/* BUSINESS */}
      <Card icon={Briefcase} title="Income from Business / Profession" subtitle="Normal, 44AD, 44ADA, 44AE" enabled={income.business.enabled} onToggle={toggle("business")} badge={income.business.enabled ? fmt(calc.biz.income) : undefined}>
        <Row>
          <Field label="Business Name" value={income.business.businessName} onChange={(v) => upd("business")({ businessName: v })} type="text" placeholder="e.g. Sharma Trading Co." />
          <Field label="Computation Method" value={income.business.type} onChange={(v) => upd("business")({ type: v })} options={["normal", "44AD", "44ADA", "44AE"]} />
        </Row>
        {income.business.type === "normal" && (
          <>
            <Row>
              <Field label="Business Income / Receipts" value={income.business.income} onChange={(v) => upd("business")({ income: v })} />
              <Field label="Expenses" value={income.business.expenses} onChange={(v) => upd("business")({ expenses: v })} />
            </Row>
            <Row><Field label="Depreciation" value={income.business.depreciation} onChange={(v) => upd("business")({ depreciation: v })} /></Row>
          </>
        )}
        {income.business.type === "44AD" && (
          <>
            <Row>
              <Field label="Turnover" value={income.business.turnover} onChange={(v) => upd("business")({ turnover: v })} />
              <Field label="Declared Income (if higher)" value={income.business.declaredIncome} onChange={(v) => upd("business")({ declaredIncome: v })} />
            </Row>
            <label className="itc-checkbox-row"><input type="checkbox" checked={income.business.digital} onChange={(e) => upd("business")({ digital: e.target.checked })} /> Receipts through digital modes (6% rate applies)</label>
          </>
        )}
        {income.business.type === "44ADA" && (
          <Row>
            <Field label="Gross Receipts" value={income.business.grossReceipts} onChange={(v) => upd("business")({ grossReceipts: v })} />
            <Field label="Declared Income (if higher)" value={income.business.declaredIncome} onChange={(v) => upd("business")({ declaredIncome: v })} />
          </Row>
        )}
        {income.business.type === "44AE" && (
          <Row>
            <Field label="Number of Vehicles" value={income.business.vehicles} onChange={(v) => upd("business")({ vehicles: v })} />
            <Field label="Months Owned (per vehicle)" value={income.business.months} onChange={(v) => upd("business")({ months: v })} />
          </Row>
        )}
        <ManualEntries items={income.business.manual} onChange={(rows) => upd("business")({ manual: rows })} label="Add line item" />
        <div className="itc-note">Presumptive schemes use standard rates (8%/6% for 44AD, 50% for 44ADA, ₹7,500/vehicle/month for 44AE).</div>
        <LedgerLine label="Taxable Business / Profession Income" value={calc.biz.income} bold />
      </Card>

      {/* CAPITAL GAINS */}
      <Card icon={TrendingUp} title="Capital Gains" subtitle="Short-term and long-term" enabled={income.capitalGains.enabled} onToggle={toggle("capitalGains")} badge={income.capitalGains.enabled ? fmt(calc.cg.total) : undefined}>
        <CGSection title="Short-Term Capital Gains" rows={income.capitalGains.st.rows} onChange={(rows) => upd("capitalGains")({ st: { rows } })} equityLabel="Equity (STT paid, taxed @ 20%)" />
        <CGSection title="Long-Term Capital Gains" rows={income.capitalGains.lt.rows} onChange={(rows) => upd("capitalGains")({ lt: { rows } })} equityLabel="Listed Equity / Equity MF (₹1,25,000 exemption, taxed @ 12.5%)" isLongTerm />
        <div className="itc-note">Simplified: LTCG taxed at flat 12.5% without indexation (Finance Act 2024 regime). Consult a CA for assets acquired before 23 July 2024 where the indexation option may apply.</div>
        <LedgerLine label="Total Capital Gains" value={calc.cg.total} bold />
      </Card>

      {/* OTHER SOURCES */}
      <Card icon={Coins} title="Income from Other Sources" subtitle="Interest, dividend, family pension, and more" enabled={income.otherSources.enabled} onToggle={toggle("otherSources")} badge={income.otherSources.enabled ? fmt(calc.other.total) : undefined}>
        <OtherSourcesSection rows={income.otherSources.rows} onChange={(rows) => upd("otherSources")({ rows })} />
        <LedgerLine label="Total Income from Other Sources" value={calc.other.total} bold />
      </Card>

      <div className="itc-gti-strip">
        <div>Gross Total Income (Old Regime)</div>
        <div>{fmt(calc.old.gti)}</div>
      </div>
      <div className="itc-gti-strip alt">
        <div>Gross Total Income (New Regime)</div>
        <div>{fmt(calc.new.gti)}</div>
      </div>
    </div>
  );
}

function CGSection({ title, rows, onChange, equityLabel, isLongTerm }) {
  const dragIndex = React.useRef(null);
  const [overIndex, setOverIndex] = useState(null);
  const addRow = () => onChange([...rows, { id: uid(), description: "", sale: "", cost: "", improvement: "", expenses: "", equitySTT: false, grandfathered: false, indexedCost: "" }]);
  const updateRow = (id, patch) => onChange(rows.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  const removeRow = (id) => onChange(rows.filter((r) => r.id !== id));

  const onDragStart = (idx) => (e) => { dragIndex.current = idx; e.dataTransfer.effectAllowed = "move"; };
  const onDragOver = (idx) => (e) => { e.preventDefault(); if (overIndex !== idx) setOverIndex(idx); };
  const onDrop = (idx) => (e) => {
    e.preventDefault();
    const from = dragIndex.current;
    setOverIndex(null);
    if (from === null || from === idx) return;
    onChange(moveItem(rows, from, idx));
    dragIndex.current = null;
  };
  const onDragEnd = () => { dragIndex.current = null; setOverIndex(null); };

  return (
    <div className="itc-subsection">
      <div className="itc-subsection-head">
        <span>{title}</span>
        <button className="itc-add-row" onClick={addRow}><Plus size={13} /> Add entry</button>
      </div>
      {rows.length === 0 && <div className="itc-empty small">No entries yet.</div>}
      {rows.map((r, idx) => (
        <div
          key={r.id}
          className={"itc-cg-row" + (overIndex === idx ? " drag-over" : "")}
          draggable
          onDragStart={onDragStart(idx)}
          onDragOver={onDragOver(idx)}
          onDrop={onDrop(idx)}
          onDragEnd={onDragEnd}
        >
          <div className="itc-drag-handle-row"><span className="itc-drag-handle" title="Drag to reorder"><GripVertical size={14} /></span></div>
          <Row><Field label="Description of Asset" value={r.description} onChange={(v) => updateRow(r.id, { description: v })} type="text" placeholder="e.g. Sale of mutual fund units" /></Row>
          <Row>
            <Field label="Sale Consideration" value={r.sale} onChange={(v) => updateRow(r.id, { sale: v })} />
            <Field label="Cost of Acquisition" value={r.cost} onChange={(v) => updateRow(r.id, { cost: v })} />
          </Row>
          <Row>
            <Field label="Improvement Cost" value={r.improvement} onChange={(v) => updateRow(r.id, { improvement: v })} />
            <Field label="Transfer Expenses" value={r.expenses} onChange={(v) => updateRow(r.id, { expenses: v })} />
          </Row>
          {isLongTerm && !r.equitySTT && (
            <>
              <label className="itc-checkbox-row" style={{ marginBottom: 8 }}>
                <input type="checkbox" checked={!!r.grandfathered} onChange={(e) => updateRow(r.id, { grandfathered: e.target.checked })} />
                Land/building acquired before 23 Jul 2024 (indexation option available)
              </label>
              {r.grandfathered && (
                <Row>
                  <Field label="Indexed Cost of Acquisition" value={r.indexedCost} onChange={(v) => updateRow(r.id, { indexedCost: v })} suffix="for 20% method" />
                </Row>
              )}
            </>
          )}
          <div className="itc-cg-row-foot">
            <label className="itc-checkbox-row"><input type="checkbox" checked={r.equitySTT} onChange={(e) => updateRow(r.id, { equitySTT: e.target.checked })} /> {equityLabel}</label>
            <button className="itc-icon-btn" onClick={() => removeRow(r.id)}><Trash2 size={14} /></button>
          </div>
        </div>
      ))}
    </div>
  );
}

function OtherSourcesSection({ rows, onChange }) {
  const dragIndex = React.useRef(null);
  const [overIndex, setOverIndex] = useState(null);
  const addRow = () => onChange([...rows, { id: uid(), type: OTHER_INCOME_TYPES[0], customLabel: "", amount: "" }]);
  const updateRow = (id, patch) => onChange(rows.map((r) => (r.id === id ? { ...r, ...patch } : r)));
  const removeRow = (id) => onChange(rows.filter((r) => r.id !== id));

  const onDragStart = (idx) => (e) => { dragIndex.current = idx; e.dataTransfer.effectAllowed = "move"; };
  const onDragOver = (idx) => (e) => { e.preventDefault(); if (overIndex !== idx) setOverIndex(idx); };
  const onDrop = (idx) => (e) => {
    e.preventDefault();
    const from = dragIndex.current;
    setOverIndex(null);
    if (from === null || from === idx) return;
    onChange(moveItem(rows, from, idx));
    dragIndex.current = null;
  };
  const onDragEnd = () => { dragIndex.current = null; setOverIndex(null); };

  return (
    <div className="itc-subsection">
      <div className="itc-subsection-head">
        <span>Entries</span>
        <button className="itc-add-row" onClick={addRow}><Plus size={13} /> Add entry</button>
      </div>
      {rows.length === 0 && <div className="itc-empty small">No entries yet — add savings interest, FD interest, dividend etc.</div>}
      {rows.map((r, idx) => (
        <div
          key={r.id}
          className={"itc-os-row" + (overIndex === idx ? " drag-over" : "")}
          draggable
          onDragStart={onDragStart(idx)}
          onDragOver={onDragOver(idx)}
          onDrop={onDrop(idx)}
          onDragEnd={onDragEnd}
        >
          <span className="itc-drag-handle" title="Drag to reorder"><GripVertical size={14} /></span>
          <select value={r.type} onChange={(e) => updateRow(r.id, { type: e.target.value })}>
            {OTHER_INCOME_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
          </select>
          <div className="itc-input-wrap"><span className="itc-rupee">₹</span><input type="number" value={r.amount} onChange={(e) => updateRow(r.id, { amount: e.target.value })} /></div>
          <button className="itc-icon-btn" onClick={() => removeRow(r.id)}><Trash2 size={14} /></button>
          {r.type === "Any Other" && (
            <input className="itc-os-custom" type="text" placeholder="Describe this income for print" value={r.customLabel || ""} onChange={(e) => updateRow(r.id, { customLabel: e.target.value })} />
          )}
        </div>
      ))}
    </div>
  );
}

/* ============================================================
   DEDUCTIONS VIEW
   ============================================================ */
function DeductionsView({ deductions, setDeductions, manualDeductions, setManualDeductions, config, old }) {
  const toggle = (code) => (v) => setDeductions((d) => ({ ...d, [code]: { ...d[code], enabled: v } }));
  const upd = (code) => (patch) => setDeductions((d) => ({ ...d, [code]: { ...d[code], ...patch } }));
  return (
    <div className="itc-page">
      <PageHeader title="Deductions — Chapter VI-A" subtitle="Only checked items appear on the printed computation sheet. Applicable only under the Old Regime (except 80CCD(2))." />
      <div className="itc-card static">
        {Object.entries(config.chapterVIA).map(([code, rule]) => (
          <div key={code} className="itc-ded-row">
            <label className="itc-checkbox-row wide">
              <input type="checkbox" checked={deductions[code].enabled} onChange={(e) => toggle(code)(e.target.checked)} />
              <span>
                <strong>{code}</strong> {rule.label.replace(code + " — ", "")}
                {rule.regime === "both" && <span className="itc-pill">Both regimes</span>}
                {rule.cap && <span className="itc-cap-note"> (max {fmt(rule.capSevere && deductions[code].severe ? rule.capSevere : rule.capSenior && deductions[code].senior ? rule.capSenior : rule.cap)})</span>}
              </span>
            </label>
            {deductions[code].enabled && (
              <div className="itc-ded-inputs">
                <div className="itc-input-wrap small"><span className="itc-rupee">₹</span><input type="number" value={deductions[code].amount} onChange={(e) => upd(code)({ amount: e.target.value })} /></div>
                {rule.capSevere && <label className="itc-checkbox-row"><input type="checkbox" checked={deductions[code].severe} onChange={(e) => upd(code)({ severe: e.target.checked })} /> Severe disability (≥80%)</label>}
                {rule.capSenior && <label className="itc-checkbox-row"><input type="checkbox" checked={deductions[code].senior} onChange={(e) => upd(code)({ senior: e.target.checked })} /> Senior citizen</label>}
              </div>
            )}
          </div>
        ))}
        <ManualEntries items={manualDeductions} onChange={setManualDeductions} label="Add other deduction" showRegimeToggle />
      </div>
      <div className="itc-gti-strip">
        <div>Total Chapter VI-A Deductions (Old Regime)</div>
        <div>{fmt(old.chapterVIATotal)}</div>
      </div>
    </div>
  );
}

/* ============================================================
   REGIME COMPARISON VIEW
   ============================================================ */
function RegimeView({ calc, betterRegime, printRegime, setPrintRegime }) {
  return (
    <div className="itc-page">
      <PageHeader title="Tax Regime Comparison" subtitle="Both regimes calculated simultaneously from the same income data" />
      <div className="itc-regime-grid">
        <RegimeCard label="Old Regime" data={calc.old} isBetter={betterRegime === "old"} />
        <RegimeCard label="New Regime" data={calc.new} isBetter={betterRegime === "new"} />
      </div>
      <div className="itc-card static">
        <div className="itc-section-label">Select Regime for Print / Filing</div>
        <div className="itc-radio-group">
          <label><input type="radio" checked={printRegime === "old"} onChange={() => setPrintRegime("old")} /> Print Old Regime</label>
          <label><input type="radio" checked={printRegime === "new"} onChange={() => setPrintRegime("new")} /> Print New Regime</label>
        </div>
      </div>
    </div>
  );
}

function RegimeCard({ label, data, isBetter }) {
  const ageLabel = label === "Old Regime" ? { normal: null, senior: "Senior citizen slabs (60–79 yrs)", superSenior: "Super senior citizen slabs (80+ yrs)" }[data.ageCategory] : null;
  return (
    <div className={"itc-regime-card" + (isBetter ? " better" : "")}>
      {isBetter && <div className="itc-better-tag"><CheckCircle2 size={13} /> Lower Tax Liability</div>}
      <div className="itc-regime-card-title">{label}</div>
      {ageLabel && <div className="itc-age-note">{ageLabel}</div>}
      <LedgerLine label="Total Taxable Income" value={data.totalTaxableIncome} />
      <LedgerLine label="Tax on Slab Income" value={data.baseTax} />
      {data.specialTax > 0 && <LedgerLine label="Tax on Special Rate Income (CG)" value={data.specialTax} />}
      <LedgerLine label="Less: Rebate u/s 87A" value={-data.rebate} />
      <LedgerLine label="Tax after Rebate" value={data.taxAfterRebate} indent />
      <LedgerLine label={`Surcharge (${data.surchargeRate}%)`} value={data.surchargeAmount} />
      <LedgerLine label="Health & Education Cess (4%)" value={data.cessAmount} />
      <LedgerLine label="Total Tax Liability" value={data.total} bold />
    </div>
  );
}

/* ============================================================
   LIABILITY VIEW
   ============================================================ */
function LiabilityView({ taxPaid, setTaxPaid, calc, printRegime }) {
  const d = calc[printRegime];
  const set = (k) => (v) => setTaxPaid((t) => ({ ...t, [k]: v }));
  const applySuggestions = () => {
    const s = d.interestSuggestions;
    setTaxPaid((t) => ({ ...t, interest234A: String(s.interest234A), interest234B: String(s.interest234B), interest234C: String(s.interest234C), fee234F: String(s.fee234F) }));
  };
  return (
    <div className="itc-page">
      <PageHeader title="Tax Paid & Final Liability" subtitle={`Based on the ${printRegime === "old" ? "Old" : "New"} Regime selected for print`} />
      <div className="itc-card static">
        <div className="itc-section-label">Advance Tax (by installment)</div>
        <Row>
          <Field label="By 15 Jun (Q1)" value={taxPaid.advanceTaxQ1} onChange={set("advanceTaxQ1")} />
          <Field label="By 15 Sep (Q2)" value={taxPaid.advanceTaxQ2} onChange={set("advanceTaxQ2")} />
        </Row>
        <Row>
          <Field label="By 15 Dec (Q3)" value={taxPaid.advanceTaxQ3} onChange={set("advanceTaxQ3")} />
          <Field label="By 15 Mar (Q4)" value={taxPaid.advanceTaxQ4} onChange={set("advanceTaxQ4")} />
        </Row>
        <div className="itc-note">Filling these in (instead of a single total) lets 234B/234C interest be calculated automatically below. If left blank, the single "Advance Tax" total is used instead.</div>

        <div className="itc-section-label">Taxes Already Paid</div>
        <Row>
          <Field label="TDS" value={taxPaid.tds} onChange={set("tds")} />
          <Field label="Advance Tax (total, if not entered by installment above)" value={taxPaid.advanceTax} onChange={set("advanceTax")} />
        </Row>
        <Row>
          <Field label="Self Assessment Tax" value={taxPaid.selfAssessmentTax} onChange={set("selfAssessmentTax")} />
          <Field label="TCS" value={taxPaid.tcs} onChange={set("tcs")} />
        </Row>
        <ManualEntries items={taxPaid.manual} onChange={(rows) => setTaxPaid((t) => ({ ...t, manual: rows }))} label="Add other prepaid tax / challan" />
      </div>
      <div className="itc-card static">
        <div className="itc-section-label">Filing Dates</div>
        <Row>
          <Field label="Actual Filing Date" value={taxPaid.filingDate} onChange={set("filingDate")} type="date" />
          <Field label="Due Date (leave blank for 31 Jul default)" value={taxPaid.dueDate} onChange={set("dueDate")} type="date" />
        </Row>
        <div className="itc-section-label">Interest & Fees</div>
        <Row>
          <Field label="Interest u/s 234A (delay in filing)" value={taxPaid.interest234A} onChange={set("interest234A")} />
          <Field label="Interest u/s 234B (advance tax shortfall)" value={taxPaid.interest234B} onChange={set("interest234B")} />
        </Row>
        <Row>
          <Field label="Interest u/s 234C (deferment of instalments)" value={taxPaid.interest234C} onChange={set("interest234C")} />
          <Field label="Fee u/s 234F (late filing fee)" value={taxPaid.fee234F} onChange={set("fee234F")} />
        </Row>
        <button className="itc-btn primary" onClick={applySuggestions}>
          Auto-fill from Filing Date &amp; Advance Tax (suggests {fmt(d.interestSuggestions.interest234A)} / {fmt(d.interestSuggestions.interest234B)} / {fmt(d.interestSuggestions.interest234C)} / {fmt(d.interestSuggestions.fee234F)})
        </button>
        <div className="itc-note">Suggestions use the standard non-audit due date (31 July) and standard formulas — a reasonable estimate, not a substitute for a CA's sign-off on edge cases (audit cases, updated returns, etc.). Values above remain fully editable after auto-filling.</div>
      </div>
      <div className="itc-card static">
        <LedgerLine label="Gross Tax Liability" value={d.grossTaxLiability} />
        <LedgerLine label="Add: Interest & Fees (234A/B/C/F)" value={d.interestTotal} />
        <LedgerLine label="Total Tax Liability (Including Interest)" value={d.totalWithInterest} bold />
        <LedgerLine label="Less: Total Taxes Paid" value={-d.taxPaidTotal} />
        <div className="itc-final-result">
          {d.refund > 0 ? (
            <>
              <span className="label">Refund Due</span>
              <span className="amount refund">{fmt(d.refund)}</span>
            </>
          ) : d.payable > 0 ? (
            <>
              <span className="label">Tax Payable</span>
              <span className="amount payable">{fmt(d.payable)}</span>
            </>
          ) : (
            <>
              <span className="label">Balance</span>
              <span className="amount">₹0 — Nil</span>
            </>
          )}
        </div>
      </div>
    </div>
  );
}

/* ============================================================
   COMPUTATION SHEET VIEW (printable)
   ============================================================ */
/** Formats a number the way the reference sheet does: plain digits with
 *  Indian comma grouping, negatives in parentheses, zero as a dash. */
function printFmt(n) {
  if (n === null || n === undefined) return "";
  const v = Math.round(n * 100) / 100;
  if (v === 0) return "-";
  const formatted = Math.abs(v).toLocaleString("en-IN", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  return v < 0 ? `(${formatted})` : formatted;
}

/** Two-amount-column ledger table — mirrors the reference computation
 *  sheet, where most rows use only the outer (right) column, and a
 *  block's very last line carries both the closing inner deduction and
 *  the outer subtotal on the same row. */
function LedgerTable({ rows }) {
  return (
    <table className="sheet-table">
      <thead>
        <tr><th>Description</th><th>Amount (Rs.)</th><th>Amount (Rs.)</th></tr>
      </thead>
      <tbody>
        {rows.map((r, i) => {
          if (r.kind === "head") return <tr key={i} className="sheet-table-head"><td colSpan={3}>{r.label}</td></tr>;
          if (r.kind === "sub") return <tr key={i}><td colSpan={3} className="sheet-table-sub">{r.label}</td></tr>;
          if (r.kind === "note") return <tr key={i}><td colSpan={3} className="sheet-table-note">{r.label}</td></tr>;
          const hasOuter = r.outer != null;
          return (
            <tr key={i} className={r.bold ? "bold" : ""}>
              <td className="desc">{r.label}</td>
              <td className={"amt" + (hasOuter ? " underline" : "")}>{r.inner != null ? printFmt(r.inner) : ""}</td>
              <td className={"amt" + (r.red ? " red" : "")}>{hasOuter ? printFmt(r.outer) : ""}</td>
            </tr>
          );
        })}
      </tbody>
    </table>
  );
}

function SheetView({ assessee, income, deductions, manualDeductions, config, calc, printRegime, setPrintRegime, taxPaid, letterhead, saveLetterhead, updateLetterheadLive, pageSetup, savePageSetup, onOpenClientEdit }) {
  const d = calc[printRegime];
  const regimeLabel = printRegime === "old" ? "Old Tax Regime" : "New Tax Regime";
  const stdDed = config[printRegime === "old" ? "oldRegime" : "newRegime"].standardDeduction;
  const [pageSetupOpen, setPageSetupOpen] = useState(false);
  const [logoOpen, setLogoOpen] = useState(false);
  const logoDrag = React.useRef(null);

  /* -------- live "how many pages will this print as?" estimate --------
     Measures the actual (unscaled) rendered height of each section and
     divides by how much vertical space one physical page gives it at the
     configured paper size / margins / print scale. The Income section
     (Personal Info + Computation of Income) can spread across 1-2 pages
     depending on how much data is in it; the Tax Liability section always
     starts on its own fresh page (see .sheet-page-liability page-break). */
  const incomeSheetRef = React.useRef(null);
  const liabilitySheetRef = React.useRef(null);
  const [pageCount, setPageCount] = useState({ income: 1, liability: 1 });

  useEffect(() => {
    const measure = () => {
      const MM_TO_PX_1 = 3.7795;
      const b = PAPER_SIZES[pageSetup.paperSize] || PAPER_SIZES.A4;
      const paperH = pageSetup.orientation === "landscape" ? b[0] : b[1];
      const marginsPx = (num(pageSetup.marginTop) + num(pageSetup.marginBottom)) * MM_TO_PX_1;
      const printScale = Math.max(50, Math.min(150, num(pageSetup.scale) || 100)) / 100;
      const usableHeightPerPage = Math.max(200, (paperH * MM_TO_PX_1 - marginsPx) / printScale);
      const incomeH = incomeSheetRef.current ? incomeSheetRef.current.scrollHeight : 0;
      const liabilityH = liabilitySheetRef.current ? liabilitySheetRef.current.scrollHeight : 0;
      setPageCount({
        income: Math.max(1, Math.ceil(incomeH / usableHeightPerPage)),
        liability: Math.max(1, Math.ceil(liabilityH / usableHeightPerPage)),
      });
    };
    const t = setTimeout(measure, 60);
    window.addEventListener("resize", measure);
    return () => { clearTimeout(t); window.removeEventListener("resize", measure); };
  }, [pageSetup, letterhead, printRegime, assessee, income, deductions, manualDeductions, taxPaid]);

  const totalPages = pageCount.income + pageCount.liability;

  const onLogoMouseMove = React.useCallback((e) => {
    const d2 = logoDrag.current;
    if (!d2) return;
    const dx = e.clientX - d2.mx;
    const dy = e.clientY - d2.my;
    if (d2.mode === "resize") {
      const next = Math.max(28, Math.min(120, d2.size + dx));
      d2.latest = { size: next, x: d2.x, y: d2.y };
      if (d2.target === "header") updateLetterheadLive({ logoSize: next });
      else updateLetterheadLive({ footerLogoSize: next });
    } else {
      const nx = Math.max(-70, Math.min(70, d2.x + dx));
      const ny = Math.max(-40, Math.min(40, d2.y + dy));
      d2.latest = { size: d2.size, x: nx, y: ny };
      if (d2.target === "header") updateLetterheadLive({ logoX: nx, logoY: ny });
      else updateLetterheadLive({ footerLogoX: nx, footerLogoY: ny });
    }
  }, [updateLetterheadLive]);

  const onLogoMouseUp = React.useCallback(() => {
    const d2 = logoDrag.current;
    if (!d2) return;
    logoDrag.current = null;
    window.removeEventListener("mousemove", onLogoMouseMove);
    window.removeEventListener("mouseup", onLogoMouseUp);
    if (d2.target === "header") saveLetterhead({ logoSize: d2.latest.size, logoX: d2.latest.x, logoY: d2.latest.y });
    else saveLetterhead({ footerLogoSize: d2.latest.size, footerLogoX: d2.latest.x, footerLogoY: d2.latest.y });
  }, [onLogoMouseMove, saveLetterhead]);

  const startLogoDrag = (target, mode) => (e) => {
    e.preventDefault();
    e.stopPropagation();
    const size = target === "header" ? (letterhead.logoSize || 64) : (letterhead.footerLogoSize || 48);
    const x = target === "header" ? (letterhead.logoX || 0) : (letterhead.footerLogoX || 0);
    const y = target === "header" ? (letterhead.logoY || 0) : (letterhead.footerLogoY || 0);
    logoDrag.current = { target, mode, mx: e.clientX, my: e.clientY, size, x, y, latest: { size, x, y } };
    window.addEventListener("mousemove", onLogoMouseMove);
    window.addEventListener("mouseup", onLogoMouseUp);
  };

  const bannerDrag = React.useRef(null);

  const onBannerMouseMove = React.useCallback((e) => {
    const b = bannerDrag.current;
    if (!b) return;
    const dx = e.clientX - b.mx;
    const dy = e.clientY - b.my;
    const next = {
      widthPct: Math.max(20, Math.min(100, b.widthPct + dx / 3)),
      height: Math.max(24, Math.min(260, b.height + dy)),
    };
    b.latest = next;
    updateLetterheadLive({ bannerWidthPct: next.widthPct, bannerHeight: next.height });
  }, [updateLetterheadLive]);

  const onBannerMouseUp = React.useCallback(() => {
    const b = bannerDrag.current;
    if (!b) return;
    bannerDrag.current = null;
    window.removeEventListener("mousemove", onBannerMouseMove);
    window.removeEventListener("mouseup", onBannerMouseUp);
    saveLetterhead({ bannerWidthPct: b.latest.widthPct, bannerHeight: b.latest.height });
  }, [onBannerMouseMove, saveLetterhead]);

  const startBannerDrag = (e) => {
    e.preventDefault();
    e.stopPropagation();
    const widthPct = letterhead.bannerWidthPct != null ? letterhead.bannerWidthPct : 100;
    const height = letterhead.bannerHeight || 70;
    bannerDrag.current = { mx: e.clientX, my: e.clientY, widthPct, height, latest: { widthPct, height } };
    window.addEventListener("mousemove", onBannerMouseMove);
    window.addEventListener("mouseup", onBannerMouseUp);
  };

  /* Live on-screen preview of the actual print page: real paper width at
     screen scale, with the configured margins as real padding, so what you
     see here is what prints. Content flows naturally across as many pages
     as it needs at the configured scale (100% by default) — no forced
     single-page shrinking. */
  const MM_TO_PX = 3.7795;
  const base = PAPER_SIZES[pageSetup.paperSize] || PAPER_SIZES.A4;
  const paperW = pageSetup.orientation === "landscape" ? base[1] : base[0];
  const paperH = pageSetup.orientation === "landscape" ? base[0] : base[1];
  const screenScale = 0.82; // fit comfortably in the main content column
  const scalePct = Math.max(50, Math.min(150, num(pageSetup.scale) || 100)) / 100;
  const pageStyle = {
    width: `${paperW * MM_TO_PX * screenScale}px`,
    minHeight: `${paperH * MM_TO_PX * screenScale}px`,
    paddingTop: `${num(pageSetup.marginTop) * MM_TO_PX * screenScale}px`,
    paddingBottom: `${num(pageSetup.marginBottom) * MM_TO_PX * screenScale}px`,
    paddingLeft: `${num(pageSetup.marginLeft) * MM_TO_PX * screenScale}px`,
    paddingRight: `${num(pageSetup.marginRight) * MM_TO_PX * screenScale}px`,
    transform: `scale(${scalePct})`,
    transformOrigin: "top center",
    maxWidth: "none",
    position: "relative",
    display: "flex",
    flexDirection: "column",
  };

  const Letterhead = ({ small }) => (
    <div className={"sheet-header" + (small ? " small" : "")}>
      <div className="sheet-logo-wrap">
        {letterhead.logo ? (
          <div
            className="sheet-logo-frame"
            style={{
              width: letterhead.logoSize || 64,
              height: letterhead.logoSize || 64,
              transform: `translate(${letterhead.logoX || 0}px, ${letterhead.logoY || 0}px)`,
            }}
            onMouseDown={startLogoDrag("header", "move")}
            title="Drag to move"
          >
            <img className="sheet-logo" src={letterhead.logo} alt="" draggable={false} />
            <div className="itc-inline-logo-handle no-print" onMouseDown={startLogoDrag("header", "resize")} title="Drag to resize" />
          </div>
        ) : <div className="sheet-logo-placeholder" />}
      </div>
      <div className="sheet-header-text">
        <div className="sheet-org">Computation of Income Tax</div>
        {(letterhead.firmName || letterhead.tagline) && (
          <div className="sheet-tagline">{[letterhead.firmName, letterhead.tagline].filter(Boolean).join(" — ")}</div>
        )}
      </div>
    </div>
  );

  const MetaBar = () => (
    <div className="sheet-meta-bar">
      <div><span>Financial Year :</span> {yearOnly(assessee.fy)}</div>
      <div><span>Assessment Year :</span> {yearOnly(assessee.ay)}</div>
      <div className="sheet-meta-itr">{assessee.itrType}</div>
    </div>
  );

  const Footer = ({ section }) => (
    <div className="sheet-footer">
      {section === "Tax Liability" && letterhead.footerAd && (
        <div className="sheet-footer-ad">{letterhead.footerAd}</div>
      )}
      <div className="sheet-footer-main">
        {letterhead.footerLogo && (
          <div className="sheet-footer-logo-wrap">
            <div
              className="sheet-footer-logo-frame"
              style={{
                width: letterhead.footerLogoSize || 48,
                height: letterhead.footerLogoSize || 48,
                transform: `translate(${letterhead.footerLogoX || 0}px, ${letterhead.footerLogoY || 0}px)`,
              }}
              onMouseDown={startLogoDrag("footer", "move")}
              title="Drag to move"
            >
              <img className="sheet-footer-logo" src={letterhead.footerLogo} alt="" draggable={false} />
              <div className="itc-inline-logo-handle no-print" onMouseDown={startLogoDrag("footer", "resize")} title="Drag to resize" />
            </div>
          </div>
        )}
        <div className="sheet-footer-grid">
          <div>{letterhead.contact || ""}</div>
        </div>
      </div>
    </div>
  );

  /* -------- build the Computation of Income ledger rows -------- */
  const incomeRows = [];

  /** Drops zero-amount rows (per user request: no clutter for nil figures),
   *  but always keeps at least one line so the block's outer subtotal has
   *  somewhere to attach to. */
  const finalizeLines = (lines, outerValue) => {
    let filtered = lines.filter((l) => l.inner === null || l.inner === undefined || Math.round(l.inner * 100) !== 0);
    if (filtered.length === 0) filtered = [{ label: lines.length ? lines[lines.length - 1].label : "Net Income", inner: null }];
    filtered[filtered.length - 1] = { ...filtered[filtered.length - 1], outer: outerValue };
    return filtered.map((l) => ({ kind: "line", ...l }));
  };

  if (income.salary.enabled) {
    incomeRows.push({ kind: "head", label: "Income from Salary" });
    const lines = [];
    if (income.salary.employerName) {
      lines.push({ label: `Employer: ${income.salary.employerName}`, inner: null });
    }
    if (num(income.salary.basic) > 0) {
      lines.push({ label: "Basic Salary (U/s 17(1))", inner: num(income.salary.basic) });
    }
    if (num(income.salary.da) > 0) {
      lines.push({ label: "Dearness Allowance (U/s 17(1))", inner: num(income.salary.da) });
    }
    if (num(income.salary.hra) > 0) {
      lines.push({ label: "House Rent Allowance (U/s 17(1))", inner: num(income.salary.hra) });
    }
    if (num(income.salary.bonus) > 0) {
      lines.push({ label: "Bonus / Ex-gratia (U/s 17(1)(iv))", inner: num(income.salary.bonus) });
    }
    if (num(income.salary.commission) > 0) {
      lines.push({ label: "Commission (U/s 17(1)(iv))", inner: num(income.salary.commission) });
    }
    if (num(income.salary.perquisites) > 0) {
      lines.push({ label: "Perquisites (U/s 17(2))", inner: num(income.salary.perquisites) });
    }
    if (num(income.salary.otherAllowances) > 0) {
      lines.push({ label: "Other Allowances (U/s 17(1))", inner: num(income.salary.otherAllowances) });
    }

    (income.salary.manual || []).forEach((m) => lines.push({ label: m.label || "Additional salary item", inner: num(m.amount) }));
    lines.push({ label: "Less: Standard Deduction U/s 16(ia)", inner: -stdDed });
    if (num(income.salary.professionalTax) > 0) lines.push({ label: "Less: Professional Tax U/s 16(iii)", inner: -num(income.salary.professionalTax) });
    if (num(income.salary.entertainmentAllowance) > 0) lines.push({ label: "Less: Entertainment Allowance U/s 16(ii)", inner: -num(income.salary.entertainmentAllowance) });
    incomeRows.push(...finalizeLines(lines, d.salaryIncome));
  }

  if (income.houseProperty.enabled) {
    incomeRows.push({ kind: "head", label: "Income From House Property" });
    const lines = [];
    if (income.houseProperty.type === "self") {
      lines.push({ label: "Self-Occupied — Less: Interest on Housing Loan U/s 24(b)", inner: -calc.hp.interestAllowedOld });
    } else {
      lines.push({ label: "Gross Annual Value", inner: calc.hp.gav });
      lines.push({ label: "Less: Municipal Tax Paid", inner: -num(income.houseProperty.municipalTax) });
      lines.push({ label: "Net Annual Value", inner: calc.hp.nav });
      lines.push({ label: "Less: Standard Deduction @ 30%", inner: -calc.hp.stdDed });
      lines.push({ label: "Less: Interest on Housing Loan U/s 24(b)", inner: -num(income.houseProperty.interest) });
    }
    (income.houseProperty.manual || []).forEach((m) => lines.push({ label: m.label || "Additional item", inner: num(m.amount) }));
    incomeRows.push(...finalizeLines(lines, d.hpIncome));
  }

  if (income.business.enabled) {
    incomeRows.push({ kind: "head", label: "Income From Business / Profession" });
    const methodLabel = { normal: "Normal computation", "44AD": "Presumptive u/s 44AD", "44ADA": "Presumptive u/s 44ADA", "44AE": "Presumptive u/s 44AE" }[income.business.type];
    const noteParts = [];
    if (income.business.businessName) noteParts.push(`Business Name: ${income.business.businessName}`);
    noteParts.push(`Method: ${methodLabel}`);
    incomeRows.push({ kind: "note", label: noteParts.join("   |   ") });
    const lines = [];
    if (income.business.type === "normal") {
      lines.push({ label: "Business Income / Receipts (T/O)", inner: num(income.business.income) });
      lines.push({ label: "Less: Expenses", inner: -num(income.business.expenses) });
      lines.push({ label: "Less: Depreciation", inner: -num(income.business.depreciation) });
    } else {
      if (income.business.type === "44AD") lines.push({ label: "Turnover (T/O)", inner: num(income.business.turnover) });
      if (income.business.type === "44ADA") lines.push({ label: "Gross Receipts (T/O)", inner: num(income.business.grossReceipts) });
      if (income.business.type === "44AE") lines.push({ label: `Vehicles: ${income.business.vehicles || 0}  ×  Months: ${income.business.months || 0}`, inner: null });
      lines.push({ label: "Presumptive Income Computed", inner: calc.biz.income - calc.biz.manualTotal });
    }
    (income.business.manual || []).forEach((m) => lines.push({ label: m.label || "Additional item", inner: num(m.amount) }));
    incomeRows.push(...finalizeLines(lines, d.bizIncome));
  }

  if (income.capitalGains.enabled) {
    incomeRows.push({ kind: "head", label: "Income From Capital Gain" });
    const ltDetailById = new Map((calc.cg.lt.detail || []).map((d2) => [d2.row.id, d2]));
    const buildBlock = (rowsIn, label, isLT) => {
      if (!rowsIn.length) return;
      incomeRows.push({ kind: "sub", label: `${label}:` });
      let blockTotal = 0;
      rowsIn.forEach((r) => {
        const detail = isLT ? ltDetailById.get(r.id) : null;
        const usedIndexed = detail && detail.method === "20% with indexation";
        const costTotal = usedIndexed ? (num(r.indexedCost) + num(r.improvement) + num(r.expenses)) : (num(r.cost) + num(r.improvement) + num(r.expenses));
        const gain = detail ? detail.gain : num(r.sale) - costTotal;
        blockTotal += gain;
        incomeRows.push({ kind: "line", label: `Sales Consideration of ${r.description || "asset"}`, inner: num(r.sale) });
        incomeRows.push({
          kind: "line",
          label: usedIndexed ? "Less: Indexed Cost of Acquisition, Improvement & Expenses" : "Less: Cost of Acquisition, Improvement & Expenses",
          inner: -costTotal,
          outer: gain,
        });
        if (isLT && r.grandfathered && num(r.indexedCost) > 0) {
          incomeRows.push({ kind: "note", label: `Taxed at ${usedIndexed ? "20% with indexation" : "12.5% without indexation"} (lower of the two methods)` });
        }
      });
      if (rowsIn.length > 1) incomeRows.push({ kind: "line", label: `Net ${label}`, outer: blockTotal, bold: true });
    };
    buildBlock(income.capitalGains.st.rows, "Short Term Capital Gain", false);
    buildBlock(income.capitalGains.lt.rows, "Long Term Capital Gain", true);
  }

  if (income.otherSources.enabled) {
    incomeRows.push({ kind: "head", label: "Income From Other Sources" });
    const lines = income.otherSources.rows.map((r) => ({ label: r.type === "Any Other" && r.customLabel ? r.customLabel : r.type, inner: num(r.amount) }));
    if (lines.length) incomeRows.push(...finalizeLines(lines, d.otherIncome));
  }

  incomeRows.push({ kind: "line", label: "Gross Total Income", outer: d.gti, bold: true });
  if (d.chapterVIABreakup.length > 0) {
    incomeRows.push({ kind: "head", label: "Less: Deduction (Chapter VI-A)" });
    d.chapterVIABreakup.forEach((b) => {
      incomeRows.push({ kind: "line", label: b.code === "Other" ? b.label : `${b.code} — ${(b.label.split(" — ")[1] || b.label)}`, inner: -b.amount });
    });
    incomeRows.push({ kind: "line", label: "Total Deduction", outer: -d.chapterVIATotal });
  } else {
    incomeRows.push({ kind: "line", label: "Less: Deduction", outer: 0 });
  }
  incomeRows.push({
    kind: "line",
    label: <>Total Income <span style={{ fontSize: "11px", fontWeight: "normal", color: "#666", marginLeft: "6px" }}>[after Rounded off as per U/s 288A]</span></>,
    excelLabel: "Total Income [after Rounded off as per U/s 288A]",
    outer: d.totalTaxableIncome,
    bold: true,
    red: true
  });

  /* -------- tax liability table (page 2) — liability + taxes paid in one table -------- */
  const liabilityRowsRaw = [
    {
      kind: "line",
      label: <>Taxable Total Income <span style={{ fontSize: "11px", fontWeight: "normal", color: "#666", marginLeft: "6px" }}>[after Rounded off as per U/s 288A]</span></>,
      excelLabel: "Taxable Total Income [after Rounded off as per U/s 288A]",
      outer: d.totalTaxableIncome,
      bold: true
    },
    { kind: "line", label: "Tax Payable on Total Income", outer: d.baseTax + d.specialTax },
    { kind: "line", label: "Less: Rebate U/s 87A", outer: -d.rebate },
    { kind: "line", label: "Tax Payable After Rebate", outer: d.taxAfterRebate, bold: true },
    { kind: "line", label: `Add: Surcharge (${d.surchargeRate}%)`, outer: d.surchargeAmount },
    { kind: "line", label: "Add: Health & Education Cess (4%)", outer: d.cessAmount },
    { kind: "line", label: "Gross Tax Liability", outer: d.grossTaxLiability, bold: true },
    { kind: "line", label: "Add: Interest u/s 234A", inner: num(taxPaid.interest234A) },
    { kind: "line", label: "Add: Interest u/s 234B", inner: num(taxPaid.interest234B) },
    { kind: "line", label: "Add: Interest u/s 234C", inner: num(taxPaid.interest234C) },
    { kind: "line", label: "Add: Fee u/s 234F", inner: num(taxPaid.fee234F), outer: d.interestTotal },
    { kind: "line", label: "Total Tax Liability (Including Interest)", outer: d.totalWithInterest, bold: true },
    { kind: "head", label: "Less: Taxes Paid" },
    { kind: "line", label: "Total Advance Tax Paid", inner: -d.effectiveAdvanceTax },
    { kind: "line", label: "Total Self Assessment Tax Paid", inner: -num(taxPaid.selfAssessmentTax) },
    { kind: "line", label: "Total TDS Claimed", inner: -num(taxPaid.tds) },
    { kind: "line", label: "Tax Collection at Source", inner: -num(taxPaid.tcs) },
    ...(taxPaid.manual || []).map((m) => ({ kind: "line", label: m.label || "Other prepaid tax", inner: -num(m.amount) })),
    { kind: "line", label: "Total Taxes Paid", outer: d.taxPaidTotal, bold: true },
  ];

  /** The Tax Liability page always prints every field (Rebate, Surcharge,
   *  each interest section, etc.) even when the value is nil — shown as
   *  "-" via printFmt — so the layout is fixed and predictable. */
  const liabilityRows = liabilityRowsRaw;

  /* -------- personal information: values + checkbox-controlled visibility -------- */
  const piValues = {
    name: assessee.name || "—",
    pan: assessee.pan || "—",
    aadhaar: assessee.aadhaar || "—",
    fatherName: assessee.fatherName || "—",
    dob: assessee.dob || "—",
    address: [assessee.address, assessee.state, assessee.pin].filter(Boolean).join(", ") || "—",
    email: assessee.email || "—",
    mobile: assessee.mobile || "—",
    bankName: assessee.bankName || "—",
    ifsc: assessee.ifsc || "—",
    accountNumber: assessee.accountNumber || "—",
    status: `${assessee.category} · ${assessee.residentialStatus}`,
  };
  const piRows = PI_FIELD_DEFS.filter((f) => pageSetup.piFields[f.key] !== false).map((f) => ({ key: f.key, label: f.label, value: piValues[f.key] }));
  const renderPIValue = (r) => (
    r.key === "pan" && assessee.pan && onOpenClientEdit
      ? <a className="pan-hyperlink" onClick={() => onOpenClientEdit(assessee.pan)}>{r.value}</a>
      : r.value
  );
  const piMid = Math.ceil(piRows.length / 2);
  const piLeft = pageSetup.piTwoColumn ? piRows.slice(0, piMid) : piRows;
  const piRight = pageSetup.piTwoColumn ? piRows.slice(piMid) : [];

  const fileBaseName = `ITR COMPUTATION FY ${assessee.fy} ${assessee.name || "Computation"}`;

  /* Coloured / formatted Excel export — uses xlsx-js-style (a drop-in fork
     of the xlsx package that can actually write cell fills/fonts/borders;
     the plain "xlsx" community build silently ignores style props). */
  const exportExcel = () => {
    const NAVY = "0E2338", GOLD = "A9791B", HEAD_BG = "1F3A5C", SUBTLE_BG = "EAF0F6",
          TOTAL_BG = "FCE9CB", FINAL_BG = "0E2338", WHITE = "FFFFFF", RED = "B3372A", GREEN = "1E7A4C";
    const thin = { style: "thin", color: { rgb: "D9DEE3" } };
    const border = { top: thin, bottom: thin, left: thin, right: thin };

    const styleHeaderCell = (v) => ({ v, t: "s", s: {
      font: { bold: true, color: { rgb: WHITE }, sz: 11 },
      fill: { fgColor: { rgb: HEAD_BG } },
      alignment: { vertical: "center", horizontal: "center" },
      border,
    }});
    const styleLabelCell = (v, bold) => ({ v, t: "s", s: {
      font: { bold: !!bold, sz: 10.5, color: { rgb: "1C232C" } },
      alignment: { vertical: "center", wrapText: true },
      border,
    }});
    const styleAmountCell = (v, opts = {}) => ({
      v: v === "" ? "" : v, t: v === "" ? "s" : "n",
      z: v === "" ? undefined : "#,##0.00;[Red](#,##0.00)",
      s: {
        font: { bold: !!opts.bold, sz: 10.5, color: { rgb: opts.red ? RED : "1C232C" } },
        alignment: { vertical: "center", horizontal: "right" },
        fill: opts.fill ? { fgColor: { rgb: opts.fill } } : undefined,
        border,
      },
    });
    const rowFill = (kind) => kind === "head" ? SUBTLE_BG : (kind === "sub" || kind === "note") ? "F5F3EC" : undefined;

    const wb = XLSXStyle.utils.book_new();

    // ---- Personal Info sheet ----
    const wsPIData = [
      [styleHeaderCell("Field"), styleHeaderCell("Value")],
      ...piRows.map((r) => [styleLabelCell(r.label, true), styleLabelCell(String(r.value))]),
    ];
    const wsPI = XLSXStyle.utils.aoa_to_sheet(wsPIData);
    wsPI["!cols"] = [{ wch: 26 }, { wch: 44 }];
    wsPI["!rows"] = [{ hpt: 20 }];
    XLSXStyle.utils.book_append_sheet(wb, wsPI, "Personal Info");

    // ---- Computation sheet builder (shared by Income & Liability) ----
    const buildLedgerSheet = (rows, finalRow) => {
      const aoa = [[styleHeaderCell("Description"), styleHeaderCell("Amount (Rs.)"), styleHeaderCell("Amount (Rs.)")]];
      rows.forEach((r) => {
        const isTotal = !!r.bold;
        const fill = rowFill(r.kind) || (isTotal ? TOTAL_BG : undefined);
        if (r.kind === "head" || r.kind === "sub" || r.kind === "note") {
          const labelText = r.excelLabel || r.label;
          aoa.push([
            { ...styleLabelCell(labelText, r.kind === "head"), s: { ...styleLabelCell(labelText, r.kind === "head").s, fill: fill ? { fgColor: { rgb: fill } } : undefined } },
            { v: "", t: "s", s: { fill: fill ? { fgColor: { rgb: fill } } : undefined, border } },
            { v: "", t: "s", s: { fill: fill ? { fgColor: { rgb: fill } } : undefined, border } },
          ]);
        } else {
          const labelText = r.excelLabel || r.label;
          aoa.push([
            { ...styleLabelCell(labelText, isTotal), s: { ...styleLabelCell(labelText, isTotal).s, fill: fill ? { fgColor: { rgb: fill } } : undefined } },
            styleAmountCell(r.inner != null ? Math.round(r.inner * 100) / 100 : "", { bold: isTotal, fill }),
            styleAmountCell(r.outer != null ? Math.round(r.outer * 100) / 100 : "", { bold: isTotal, fill, red: !!r.red }),
          ]);
        }
      });
      if (finalRow) {
        const labelText = finalRow.excelLabel || finalRow.label;
        aoa.push([
          { v: labelText, t: "s", s: { font: { bold: true, sz: 12, color: { rgb: WHITE } }, fill: { fgColor: { rgb: FINAL_BG } }, alignment: { vertical: "center" }, border } },
          { v: "", t: "s", s: { fill: { fgColor: { rgb: FINAL_BG } }, border } },
          { v: finalRow.value, t: "n", z: "#,##0.00;[Red](#,##0.00)", s: { font: { bold: true, sz: 12, color: { rgb: finalRow.value < 0 ? "FF6B6B" : "FFD874" } }, fill: { fgColor: { rgb: FINAL_BG } }, alignment: { horizontal: "right" }, border } },
        ]);
      }
      const ws = XLSXStyle.utils.aoa_to_sheet(aoa);
      ws["!cols"] = [{ wch: 46 }, { wch: 17 }, { wch: 17 }];
      return ws;
    };

    const wsIncome = buildLedgerSheet(incomeRows, null);
    XLSXStyle.utils.book_append_sheet(wb, wsIncome, "Computation of Income");

    const refundAmt = Math.round((d.refund > 0 ? -d.refund : d.payable) * 100) / 100;
    const wsLiability = buildLedgerSheet(liabilityRows, {
      label: <>Tax Paid / (Refund) <span style={{ fontSize: "11px", fontWeight: "normal", color: "#666", marginLeft: "6px" }}>[after Rounded off as per U/s 288B]</span></>,
      excelLabel: "Tax Paid / (Refund) [after Rounded off as per U/s 288B]",
      value: refundAmt
    });
    XLSXStyle.utils.book_append_sheet(wb, wsLiability, "Tax Liability");

    XLSXStyle.writeFile(wb, `${fileBaseName}.xlsx`);
  };

  const exportPDF = () => {
    // No PDF-generation library is bundled here — the browser's own
    // print-to-PDF does the same job with zero dependencies and matches
    // the on-screen preview exactly, since it uses the same @page rules.
    const oldTitle = document.title;
    document.title = `ITR COMPUTATION FY ${assessee.fy} ${assessee.name || "Computation"}`;
    window.print();
    setTimeout(() => { document.title = oldTitle; }, 1000);
  }

  return (
    <div className="itc-page">
      <PageHeader
        title="Computation Sheet"
        subtitle="Pages flow naturally with your data, at 100% scale by default"
        right={
          <div className="itc-sheet-header-actions no-print">
            <div className="itc-page-count-badge" title={`Computation of Income: ${pageCount.income} page${pageCount.income > 1 ? "s" : ""} · Tax Liability: ${pageCount.liability} page${pageCount.liability > 1 ? "s" : ""}`}>
              📄 Preview: {totalPages} page{totalPages > 1 ? "s" : ""}
            </div>
            <div className="itc-radio-group inline">
              <label><input type="radio" checked={printRegime === "old"} onChange={() => setPrintRegime("old")} /> Old</label>
              <label><input type="radio" checked={printRegime === "new"} onChange={() => setPrintRegime("new")} /> New</label>
            </div>
            <button className="itc-export-btn" onClick={exportExcel}><FileSpreadsheet size={14} /> Export Excel</button>
            <button className="itc-export-btn" onClick={exportPDF}><FileDown size={14} /> Export PDF</button>
          </div>
        }
      />

      <PageSetupPanel pageSetup={pageSetup} savePageSetup={savePageSetup} open={pageSetupOpen} setOpen={setPageSetupOpen} />
      <LogoPanel letterhead={letterhead} saveLetterhead={saveLetterhead} updateLetterheadLive={updateLetterheadLive} open={logoOpen} setOpen={setLogoOpen} />

      {/* ---------------- INCOME SECTION — flows across 1–2 pages depending on data ---------------- */}
      <div className="sheet sheet-page" style={pageStyle} ref={incomeSheetRef}>
        <Letterhead />
        <MetaBar />

        <div className="sheet-section-title">Personal Information</div>
        {pageSetup.piTwoColumn ? (
          <div className="sheet-kv-two-col">
            <table className="sheet-kv-table">
              <tbody>
                {piLeft.map((r, i) => <tr key={i}><td>{r.label}</td><td>{renderPIValue(r)}</td></tr>)}
              </tbody>
            </table>
            <table className="sheet-kv-table">
              <tbody>
                {piRight.map((r, i) => <tr key={i}><td>{r.label}</td><td>{renderPIValue(r)}</td></tr>)}
              </tbody>
            </table>
          </div>
        ) : (
          <table className="sheet-kv-table">
            <tbody>
              {piRows.map((r, i) => <tr key={i}><td>{r.label}</td><td>{renderPIValue(r)}</td></tr>)}
            </tbody>
          </table>
        )}

        <div className="sheet-section-title">Computation of Income</div>
        <LedgerTable rows={incomeRows} />

        <Footer section="Computation of Income" />
      </div>

      {/* ---------------- TAX LIABILITY — always starts on a fresh page ---------------- */}
      <div className="sheet sheet-page sheet-page-liability" style={pageStyle} ref={liabilitySheetRef}>
        <Letterhead small />
        <div className="sheet-section-title">Computation of Income Tax Liability</div>
        <div className="sheet-regime-line">Tax Regime : {regimeLabel}</div>

        <LedgerTable rows={liabilityRows} />

                <div className="sheet-final" style={d.refund > 0 ? { background: "linear-gradient(135deg, #F0FDF4, #DCFCE7)" } : undefined}>
          <span>Tax Paid / (Refund) <span style={{ fontSize: "11.5px", fontWeight: "normal", color: "#666", marginLeft: "6px" }}>[after Rounded off as per U/s 288B]</span></span>
          <b style={d.refund > 0 ? { color: "#1E7A4C" } : undefined}>{printFmt(d.refund > 0 ? -d.refund : d.payable)}</b>
        </div>

        {letterhead.banner && (
          <div
            className="sheet-banner"
            style={{ height: letterhead.bannerHeight || 70, width: `${letterhead.bannerWidthPct != null ? letterhead.bannerWidthPct : 100}%` }}
          >
            <img src={letterhead.banner} alt="" />
            <div className="itc-inline-banner-handle no-print" onMouseDown={startBannerDrag} title="Drag to resize" />
          </div>
        )}

        <Footer section="Tax Liability" />
      </div>
    </div>
  );
}

/* ============================================================
   STYLES
   ============================================================ */
const CSS = `
@import url('https://fonts.googleapis.com/css2?family=Source+Serif+4:opsz,wght@8..60,500;8..60,600;8..60,700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');

:root {
  --navy: #16324F;
  --navy-deep: #0E2338;
  --blue: #2A5C8A;
  --paper: #FAF9F6;
  --line: #D9DEE3;
  --ink: #1C232C;
  --muted: #5E6B78;
  --gold: #A9791B;
  --green: #1E7A4C;
  --red: #B3372A;
  --card: #FFFFFF;
}

* { box-sizing: border-box; }
.itc-app {
  display: flex;
  min-height: 100vh;
  background: var(--paper);
  color: var(--ink);
  font-family: 'Inter', sans-serif;
  font-variant-numeric: tabular-nums;
}

/* SIDEBAR */
.itc-sidebar {
  width: 248px;
  flex-shrink: 0;
  background: var(--navy-deep);
  color: #E8EDF2;
  display: flex;
  flex-direction: column;
  padding: 20px 14px;
  gap: 18px;
  position: sticky;
  top: 0;
  align-self: flex-start;
  height: 100vh;
  overflow-y: auto;
}
.itc-brand { display: flex; align-items: center; gap: 10px; padding: 0 4px; }
.itc-seal {
  width: 34px; height: 34px; border-radius: 50%;
  border: 1.5px solid var(--gold);
  display: flex; align-items: center; justify-content: center;
  font-family: 'Source Serif 4', serif; font-weight: 700; font-size: 13px;
  color: var(--gold); letter-spacing: 0.5px;
}
.itc-brand-title { font-family: 'Source Serif 4', serif; font-size: 14.5px; font-weight: 600; line-height: 1.2; }
.itc-brand-sub { font-size: 10.5px; color: #8FA1B3; margin-top: 1px; }
.itc-ay-badge {
  background: rgba(255,255,255,0.06); border: 1px solid rgba(255,255,255,0.1);
  border-radius: 6px; padding: 8px 10px; font-size: 12.5px; font-weight: 600;
  font-family: 'IBM Plex Mono', monospace;
}
.itc-ay-badge span { color: #8FA1B3; font-weight: 500; }
.itc-nav { display: flex; flex-direction: column; gap: 2px; margin-top: 4px; }
.itc-nav-item {
  display: flex; align-items: center; gap: 10px;
  background: none; border: none; color: #C6D2DD;
  padding: 9px 10px; border-radius: 6px; font-size: 13px; text-align: left; cursor: pointer;
  font-family: 'Inter', sans-serif; transition: background 0.12s;
}
.itc-nav-item:hover { background: rgba(255,255,255,0.06); }
.itc-nav-item.active { background: var(--blue); color: #fff; font-weight: 600; }
.itc-sidebar-actions { margin-top: auto; display: grid; grid-template-columns: 1fr 1fr; gap: 6px; }
.itc-sbtn {
  display: flex; align-items: center; justify-content: center; gap: 6px;
  background: rgba(255,255,255,0.07); border: 1px solid rgba(255,255,255,0.12);
  color: #E8EDF2; padding: 8px 6px; border-radius: 6px; font-size: 12px; cursor: pointer;
  font-family: 'Inter', sans-serif; font-weight: 500;
}
.itc-sbtn:hover { background: rgba(255,255,255,0.13); }
.itc-sbtn.primary { background: var(--gold); border-color: var(--gold); color: #241A05; }

/* MAIN */
.itc-main { flex: 1; padding: 32px 40px 60px; max-width: 900px; overflow-x: auto; }
.itc-page-head { display: flex; align-items: flex-end; justify-content: space-between; margin-bottom: 20px; }
.itc-page-head h1 { font-family: 'Source Serif 4', serif; font-size: 24px; font-weight: 600; margin: 0 0 4px; color: var(--navy-deep); }
.itc-page-head p { margin: 0; font-size: 13px; color: var(--muted); }

/* CARD */
.itc-card { background: var(--card); border: 1px solid var(--line); border-radius: 10px; margin-bottom: 14px; overflow: hidden; }
.itc-card.static { padding: 20px 22px; }
.itc-card-head { display: flex; align-items: center; justify-content: space-between; padding: 14px 18px; cursor: pointer; }
.itc-card-head-left { display: flex; align-items: center; gap: 12px; }
.itc-card-icon { width: 32px; height: 32px; border-radius: 7px; background: #EEF3F7; color: var(--blue); display: flex; align-items: center; justify-content: center; }
.itc-card-title { font-weight: 600; font-size: 14px; color: var(--navy-deep); }
.itc-card-subtitle { font-size: 11.5px; color: var(--muted); margin-top: 1px; }
.itc-card-head-right { display: flex; align-items: center; gap: 10px; }
.itc-badge { font-family: 'IBM Plex Mono', monospace; font-size: 12.5px; font-weight: 600; color: var(--navy); background: #EEF3F7; padding: 3px 9px; border-radius: 5px; }
.itc-chev { transition: transform 0.15s; color: var(--muted); }
.itc-chev.open { transform: rotate(180deg); }
.itc-card-body { padding: 4px 18px 18px; border-top: 1px solid var(--line); }
.itc-check { width: 16px; height: 16px; accent-color: var(--blue); }

.itc-section-label { font-size: 11px; font-weight: 700; text-transform: uppercase; letter-spacing: 0.06em; color: var(--gold); margin: 16px 0 8px; }
.itc-section-label:first-child { margin-top: 4px; }

.itc-row { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin-bottom: 12px; }
.itc-field { display: flex; flex-direction: column; gap: 5px; font-size: 12.5px; color: var(--muted); font-weight: 500; }
.itc-field select, .itc-field input[type=date], .itc-field input[type=text], .itc-field input[type=email], .itc-field input[type=tel] {
  border: 1px solid var(--line); border-radius: 6px; padding: 8px 10px; font-size: 13.5px; font-family: 'Inter', sans-serif; color: var(--ink); background: #fff;
}
.itc-input-wrap { display: flex; align-items: center; border: 1px solid var(--line); border-radius: 6px; background: #fff; overflow: hidden; }
.itc-input-wrap.small { max-width: 160px; }
.itc-input-wrap input { border: none; padding: 8px 8px 8px 2px; font-size: 13.5px; width: 100%; font-family: 'IBM Plex Mono', monospace; }
.itc-input-wrap input:focus, .itc-field select:focus { outline: 2px solid var(--blue); outline-offset: -1px; }
.itc-rupee { padding-left: 10px; color: var(--muted); font-size: 13px; }
.itc-suffix { padding-right: 10px; color: var(--muted); font-size: 11px; white-space: nowrap; }

.itc-note { font-size: 11.5px; color: var(--muted); background: #F3F1E9; border-left: 3px solid var(--gold); padding: 8px 10px; border-radius: 4px; margin: 10px 0; line-height: 1.5; }
.itc-error-line { display: flex; gap: 14px; margin: -6px 0 10px; }
.itc-error-line span { display: flex; align-items: center; gap: 4px; color: var(--red); font-size: 11.5px; }

.ledger-line { display: flex; justify-content: space-between; padding: 5px 0; font-size: 13px; border-bottom: 1px dashed var(--line); }
.ledger-line.bold { font-weight: 700; border-bottom: 1.5px solid var(--ink); border-top: 1px solid var(--ink); margin-top: 4px; padding-top: 7px; }
.ledger-line.indent { padding-left: 14px; color: var(--muted); }
.ledger-amt { font-family: 'IBM Plex Mono', monospace; }
.ledger-amt.pos { color: var(--green); }
.ledger-amt.neg { color: var(--red); }

.itc-gti-strip { display: flex; justify-content: space-between; align-items: center; background: var(--navy-deep); color: #fff; padding: 12px 18px; border-radius: 8px; font-weight: 700; font-size: 14px; margin-top: 6px; font-family: 'IBM Plex Mono', monospace; }
.itc-gti-strip.alt { background: var(--blue); margin-top: 8px; }

.itc-subsection { margin: 12px 0; }
.itc-subsection-head { display: flex; justify-content: space-between; align-items: center; font-size: 12.5px; font-weight: 600; color: var(--navy); margin-bottom: 8px; }
.itc-add-row { display: flex; align-items: center; gap: 4px; background: #EEF3F7; border: none; padding: 5px 10px; border-radius: 5px; font-size: 12px; color: var(--blue); cursor: pointer; font-weight: 600; }
.itc-cg-row { border: 1px solid var(--line); border-radius: 8px; padding: 10px 12px 12px; margin-bottom: 10px; background: #FCFCFB; transition: background 0.12s, border-color 0.12s; }
.itc-cg-row.drag-over { border-color: var(--blue); background: #EEF3F7; }
.itc-cg-row-foot { display: flex; justify-content: space-between; align-items: center; margin-top: 6px; }
.itc-checkbox-row { display: flex; align-items: center; gap: 7px; font-size: 12px; color: var(--muted); cursor: pointer; }
.itc-checkbox-row.wide { flex: 1; font-size: 13px; color: var(--ink); }
.itc-icon-btn { background: none; border: none; color: var(--red); cursor: pointer; padding: 4px; display: flex; }
.itc-os-row { display: grid; grid-template-columns: auto 1.4fr 1fr auto; gap: 10px; align-items: center; margin-bottom: 8px; transition: background 0.12s, border-color 0.12s; border-radius: 6px; }
.itc-os-row.drag-over { background: #EEF3F7; }
.itc-os-row select { border: 1px solid var(--line); border-radius: 6px; padding: 8px 10px; font-size: 13px; }
.itc-empty { text-align: center; color: var(--muted); padding: 30px; font-size: 13px; }
.itc-empty.small { padding: 10px; font-size: 12px; }
.itc-drag-handle { display: flex; align-items: center; justify-content: center; color: var(--muted); cursor: grab; flex-shrink: 0; touch-action: none; }
.itc-drag-handle:active { cursor: grabbing; }
.itc-drag-handle-row { display: flex; justify-content: center; margin-bottom: 4px; cursor: grab; }
.itc-drag-handle-row:active { cursor: grabbing; }

.itc-ded-row { padding: 12px 0; border-bottom: 1px solid var(--line); }
.itc-ded-row:last-child { border-bottom: none; }
.itc-pill { background: #E7F2EC; color: var(--green); font-size: 10px; font-weight: 700; padding: 2px 7px; border-radius: 10px; margin-left: 8px; }
.itc-cap-note { color: var(--muted); font-size: 12px; font-weight: 400; }
.itc-ded-inputs { display: flex; align-items: center; gap: 16px; margin: 10px 0 0 26px; }

.itc-regime-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin-bottom: 16px; }
.itc-regime-card { background: var(--card); border: 1.5px solid var(--line); border-radius: 10px; padding: 18px 20px; position: relative; }
.itc-regime-card.better { border-color: var(--green); box-shadow: 0 0 0 3px rgba(30,122,76,0.08); }
.itc-better-tag { position: absolute; top: -11px; right: 16px; background: var(--green); color: #fff; font-size: 10.5px; font-weight: 700; padding: 3px 10px; border-radius: 12px; display: flex; align-items: center; gap: 4px; }
.itc-regime-card-title { font-family: 'Source Serif 4', serif; font-size: 16px; font-weight: 600; margin-bottom: 10px; color: var(--navy-deep); }
.itc-age-note { display: inline-block; background: #FBF1E3; color: #A9791B; font-size: 11px; font-weight: 600; padding: 3px 9px; border-radius: 10px; margin: -6px 0 10px; }
.itc-radio-group { display: flex; gap: 22px; }
.itc-radio-group label { display: flex; align-items: center; gap: 7px; font-size: 13.5px; cursor: pointer; }
.itc-radio-group.inline { gap: 12px; }

.itc-final-result { display: flex; justify-content: space-between; align-items: center; margin-top: 14px; padding-top: 14px; border-top: 2px solid var(--ink); }
.itc-final-result .label { font-weight: 700; font-size: 14px; }
.itc-final-result .amount { font-family: 'IBM Plex Mono', monospace; font-size: 20px; font-weight: 700; }
.itc-final-result .amount.refund { color: var(--green); }
.itc-final-result .amount.payable { color: var(--red); }

/* TOAST */
.itc-toast { position: fixed; bottom: 24px; left: 50%; transform: translateX(-50%); background: var(--navy-deep); color: #fff; padding: 10px 18px; border-radius: 8px; font-size: 13px; display: flex; align-items: center; gap: 8px; box-shadow: 0 8px 24px rgba(0,0,0,0.2); z-index: 100; }
.itc-toast.err { background: var(--red); }

/* MODAL */
.itc-modal-backdrop { position: fixed; inset: 0; background: rgba(14,35,56,0.5); display: flex; align-items: center; justify-content: center; z-index: 200; }
.itc-modal { background: #fff; border-radius: 12px; width: 380px; padding: 20px 22px; }
.itc-modal.wide { width: 480px; }
.itc-modal-head { display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px; }
.itc-modal-head h3 { font-family: 'Source Serif 4', serif; font-size: 16px; margin: 0; }
.itc-modal-head button { background: none; border: none; cursor: pointer; color: var(--muted); }
.itc-modal-text { font-size: 13px; color: var(--muted); margin-bottom: 16px; }
.itc-modal-actions { display: flex; justify-content: flex-end; gap: 8px; }
.itc-btn { padding: 8px 16px; border-radius: 6px; font-size: 13px; font-weight: 600; cursor: pointer; border: none; }
.itc-btn.ghost { background: #F0F1F2; color: var(--ink); }
.itc-btn.primary { background: var(--blue); color: #fff; }
.itc-search { display: flex; align-items: center; gap: 8px; border: 1px solid var(--line); border-radius: 7px; padding: 8px 10px; margin-bottom: 12px; color: var(--muted); }
.itc-search input { border: none; outline: none; font-size: 13px; width: 100%; }
.itc-saved-list { max-height: 340px; overflow-y: auto; display: flex; flex-direction: column; gap: 6px; }
.itc-saved-row { display: flex; justify-content: space-between; align-items: center; padding: 10px 12px; border: 1px solid var(--line); border-radius: 8px; }
.itc-saved-info { cursor: pointer; flex: 1; }
.itc-saved-name { font-weight: 600; font-size: 13px; }
.itc-saved-meta { font-size: 11px; color: var(--muted); margin-top: 2px; }
.itc-saved-actions { display: flex; gap: 8px; }
.itc-saved-actions button { background: none; border: none; color: var(--muted); cursor: pointer; padding: 4px; }
.itc-saved-actions button:hover { color: var(--navy); }

/* COMPUTATION SHEET — styled to match the reference CA-office layout */
.sheet {
  background: #fff; border: 1px solid var(--line); border-radius: 4px; padding: 18px 26px; margin-bottom: 24px;
  font-family: 'Book Antiqua', 'Palatino Linotype', Palatino, 'URW Palladio L', Georgia, serif;
  color: #1a1a1a;
  display: flex;
  flex-direction: column;
  box-sizing: border-box;
}
.sheet-page { max-width: 780px; }
.sheet-header { position: relative; display: flex; align-items: center; background: linear-gradient(135deg, #F2F2ED, #EAEAE3); border-radius: 7px; padding: 9px 15px 9px 92px; margin-bottom: 11px; min-height: 46px; }
.sheet-header.small { padding: 7px 15px 7px 92px; margin-bottom: 9px; min-height: 40px; }
.sheet-header-text { display: block; }
.sheet-logo-wrap { position: absolute; left: 14px; top: 50%; transform: translateY(-50%); }
.sheet-logo-frame { position: relative; cursor: move; }
.sheet-logo-frame img { width: 100%; height: 100%; object-fit: contain; display: block; }
.sheet-logo-placeholder { position: absolute; left: 14px; top: 50%; transform: translateY(-50%); height: 40px; width: 40px; border-radius: 6px; border: 2px dashed #B7C4CE; }
.sheet-org { font-size: 18px; font-weight: 700; color: #A6272C; letter-spacing: 0.01em; }
.sheet-tagline { font-size: 10.5px; color: #666; margin-top: 1px; }
.sheet-meta-bar { display: grid; grid-template-columns: 1fr 1fr 0.6fr; gap: 8px; margin-bottom: 12px; font-size: 12px; }
.sheet-meta-bar div { padding: 6px 10px; background: #F7F7F4; border-radius: 5px; }
.sheet-meta-bar div:last-child { text-align: center; font-weight: 700; color: #A6272C; background: #FBF1F1; }
.sheet-meta-bar span { color: #777; margin-right: 4px; }
.sheet-regime-line { font-size: 12.5px; color: #2E5C8A; font-weight: 700; margin: -4px 0 10px; }
.sheet-section-title {
  display: table; margin: 12px auto 8px; font-size: 13.5px; font-weight: 700; color: #1F5C99; text-align: center; position: relative;
  padding-bottom: 4px;
}
.sheet-section-title::after {
  content: ""; position: absolute; left: 0; right: 0; bottom: 0;
  height: 2px; background: #1F5C99; border-radius: 1px;
}
.sheet-section-title.small { font-size: 12px; margin-top: 14px; }
.sheet-final { display: flex; justify-content: space-between; align-items: center; background: linear-gradient(135deg, #FBF1F1, #F6E9E9); padding: 10px 16px; border-radius: 7px; margin: 12px 0; }
.sheet-final span { font-weight: 700; font-size: 13px; color: #444; }
.sheet-final b { font-size: 18px; color: #A6272C; }
.sheet-footer { border-top: 1px solid #EEE; padding-top: 8px; font-size: 10.5px; color: var(--muted); margin-top: auto; }
.sheet-footer-ad { background: linear-gradient(135deg, #EFF4F8, #E8F0F6); border-radius: 7px; padding: 8px 12px; margin-bottom: 8px; text-align: center; font-size: 11px; font-weight: 700; color: #2E5C8A; letter-spacing: 0.01em; }
.sheet-footer-main { position: relative; padding-left: 60px; min-height: 22px; display: flex; align-items: center; }
.sheet-footer-logo-wrap { position: absolute; left: 0; top: 50%; transform: translateY(-50%); }
.sheet-footer-logo-frame { position: relative; cursor: move; }
.sheet-footer-logo-frame img { width: 100%; height: 100%; object-fit: contain; display: block; }
.sheet-footer-grid { flex: 1; display: flex; justify-content: space-between; }

/* Personal Information key/value table */
.sheet-kv-table { width: 100%; border-collapse: separate; border-spacing: 0; font-size: 12.5px; margin-bottom: 2px; }
.sheet-kv-table tr:nth-child(even) td { background: #FAFAF8; }
.sheet-kv-table td { padding: 5px 8px; text-align: left; }
.sheet-kv-table td:first-child { width: 200px; color: #666; font-weight: 700; border-radius: 4px 0 0 4px; }
.sheet-kv-table td:last-child { color: #111; border-radius: 0 4px 4px 0; }
.sheet-kv-two-col { display: grid; grid-template-columns: 1fr 1fr; gap: 0 14px; margin-bottom: 2px; }
.sheet-kv-two-col .sheet-kv-table td:first-child { width: 140px; }

/* Two-amount-column ledger table (Computation of Income / Tax Liability) — clean, no cell borders */
.sheet-table { width: 100%; border-collapse: separate; border-spacing: 0; font-size: 12.5px; margin-bottom: 6px; }
.sheet-table thead th { background: #EAF0F6; border-bottom: 2px solid #1F5C99; padding: 7px 6px; font-size: 11px; text-transform: uppercase; letter-spacing: 0.04em; color: #1F5C99; font-weight: 700; }
.sheet-table thead th:first-child { text-align: left; border-radius: 4px 0 0 4px; }
.sheet-table thead th:not(:first-child) { text-align: right; width: 120px; }
.sheet-table thead th:last-child { border-radius: 0 4px 4px 0; }
.sheet-table tbody tr.line-row:nth-child(even) td { background: #FAFAF8; }
.sheet-table td { padding: 4px 6px; }
.sheet-table td.desc { text-align: left; }
.sheet-table td.amt { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; width: 120px; }
.sheet-table td.amt.red { color: #A6272C; font-weight: 700; }
.sheet-table tr.bold td { font-weight: 700; background: #FBF8F0 !important; }
.sheet-table tr.bold td:first-child { border-radius: 4px 0 0 4px; }
.sheet-table tr.bold td:last-child { border-radius: 0 4px 4px 0; }
.sheet-table-head td { font-weight: 800; font-size: 12.5px; color: #1F5C99; background: #F1F4F7; padding: 6px 8px; border-radius: 3px; }
.sheet-table-sub td { font-weight: 800; font-style: italic; font-size: 12px; padding: 4px 6px 4px 14px; color: #444; }
.sheet-table-note td { font-style: italic; color: #777; font-size: 11.5px; padding: 3px 6px 3px 14px; }

/* MANUAL ENTRIES */
.itc-manual-row { display: grid; grid-template-columns: auto 1.6fr 1fr auto auto; gap: 10px; align-items: center; margin-bottom: 8px; transition: background 0.12s; border-radius: 6px; }
.itc-manual-row.drag-over { background: #EEF3F7; }
.itc-manual-row input[type=text] { border: 1px solid var(--line); border-radius: 6px; padding: 8px 10px; font-size: 13px; font-family: 'Inter', sans-serif; }
.itc-manual-row select { border: 1px solid var(--line); border-radius: 6px; padding: 8px 8px; font-size: 12px; }
.itc-os-custom { grid-column: 1 / -1; border: 1px solid var(--line); border-radius: 6px; padding: 7px 10px; font-size: 12.5px; margin-top: -2px; }

/* LOGO UPLOAD, MOVE & RESIZE */
.itc-logo-upload { display: flex; align-items: center; gap: 18px; margin-bottom: 8px; }
.itc-logo-editor-box { position: relative; border: 1.5px dashed var(--line); border-radius: 10px; overflow: hidden; background: #FAFAF9; }
.itc-logo-editor-empty { position: absolute; left: 50%; top: 50%; transform: translate(-50%, -50%); font-size: 11px; color: var(--muted); }
.itc-logo-editor-frame { position: absolute; left: 50%; top: 50%; cursor: move; }
.itc-logo-editor-frame img { width: 100%; height: 100%; object-fit: contain; border-radius: 8px; user-select: none; }
.itc-logo-resize-handle { position: absolute; right: -6px; bottom: -6px; width: 14px; height: 14px; border-radius: 50%; background: var(--blue); border: 2px solid #fff; box-shadow: 0 0 0 1px var(--blue); cursor: nwse-resize; }

/* Same corner-drag handle shown directly on the sheet's logos (on-screen only) */
.itc-inline-logo-handle {
  position: absolute; right: -6px; bottom: -6px; width: 12px; height: 12px; border-radius: 50%;
  background: var(--blue); border: 2px solid #fff; box-shadow: 0 0 0 1px var(--blue); cursor: nwse-resize; z-index: 5;
}

/* BANNER UPLOAD & RESIZE (settings editor) */
.itc-banner-editor-wrap { width: 100%; }
.itc-banner-editor-box { position: relative; border: 1.5px dashed var(--line); border-radius: 10px; overflow: hidden; background: #FAFAF9; }
.itc-banner-editor-box img { width: 100%; height: 100%; object-fit: contain; display: block; }
.itc-banner-resize-handle { position: absolute; right: -6px; bottom: -6px; width: 14px; height: 14px; border-radius: 50%; background: var(--blue); border: 2px solid #fff; box-shadow: 0 0 0 1px var(--blue); cursor: nwse-resize; }

/* Banner as it appears on the actual printed sheet (last page) */
.sheet-banner { position: relative; margin-top: 12px; border-radius: 6px; overflow: hidden; }
.sheet-banner img { width: 100%; height: 100%; object-fit: contain; display: block; border-radius: 6px; }
.itc-inline-banner-handle {
  position: absolute; right: -6px; bottom: -6px; width: 12px; height: 12px; border-radius: 50%;
  background: var(--blue); border: 2px solid #fff; box-shadow: 0 0 0 1px var(--blue); cursor: nwse-resize; z-index: 5;
}
.itc-logo-upload-actions { display: flex; flex-direction: column; gap: 8px; }

/* PAGE SETUP PANEL — embedded inline atop the Computation Sheet */
.itc-pagesetup-panel { border: 1px solid var(--line); border-radius: 8px; margin-bottom: 16px; background: #fff; overflow: hidden; }
.itc-pagesetup-toggle {
  width: 100%; display: flex; align-items: center; gap: 8px; background: #F7F7F4; border: none;
  padding: 10px 16px; font-size: 13px; font-weight: 600; color: var(--navy-deep); cursor: pointer;
}
.itc-pagesetup-toggle .itc-chev { margin-left: auto; }
.itc-pagesetup-body { padding: 14px 16px 16px; border-top: 1px solid var(--line); }
.itc-pagesetup-row { display: grid; grid-template-columns: repeat(auto-fit, minmax(120px, 1fr)); gap: 12px; margin-bottom: 12px; }
.itc-pi-checklist { display: grid; grid-template-columns: repeat(auto-fit, minmax(180px, 1fr)); gap: 6px 14px; }
.itc-sheet-header-actions { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; justify-content: flex-end; }
.itc-export-btn {
  display: flex; align-items: center; gap: 6px; background: #EEF3F7; color: var(--blue);
  border: 1px solid #D7E3EC; padding: 8px 12px; border-radius: 6px; font-size: 12.5px; font-weight: 600; cursor: pointer;
}
.itc-export-btn:hover { background: #E3ECF3; }
.itc-page-count-badge {
  display: flex; align-items: center; gap: 6px; background: #F3F0E6; color: var(--gold);
  border: 1px solid #E4DBC0; padding: 8px 12px; border-radius: 6px; font-size: 12.5px; font-weight: 700;
  cursor: default; white-space: nowrap;
}
.itc-logo-panel-row { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin-bottom: 10px; }

/* Sheet page frame centered under its live-adjusted width/scale */
.sheet-page { margin-left: auto; margin-right: auto; }

@media print {
  .no-print { display: none !important; }
  .itc-app { display: block; background: #fff; }
  .itc-main { padding: 0; max-width: 100%; }
  .itc-page-head { display: none; }
  .itc-page > *:not(.sheet-page) { display: none; }
  .sheet-page {
    border: none !important; margin: 0 !important; max-width: 100% !important; width: auto !important;
    padding: 0 !important; transform: none !important;
  }
  /* The income section (personal info + computation of income) flows
     naturally across as many pages as its content needs at 100% scale.
     The tax liability section always starts on its own fresh page after
     that, whether the income section took one page or two. */
  .sheet-page-liability { page-break-before: always; break-before: page; }
}

@media (max-width: 860px) {
  .itc-app { flex-direction: column; }
  .itc-sidebar { width: 100%; flex-direction: row; flex-wrap: wrap; position: static; height: auto; overflow-y: visible; }
  .itc-nav { flex-direction: row; flex-wrap: wrap; }
  .itc-main { padding: 20px; }
  .itc-row, .itc-regime-grid, .sheet-grid { grid-template-columns: 1fr; }
  .itc-manual-row { grid-template-columns: 1fr; }
}

/* PAN hyperlink -> jumps back to the client record in Fin-Tax Mitra (screen only; print stays identical to a normal cell) */
.pan-hyperlink { color: #1F5C99; text-decoration: underline; cursor: pointer; font-weight: 700; }
.pan-hyperlink:hover { color: #163f6e; }
@media print {
  .pan-hyperlink { color: #111 !important; text-decoration: none !important; cursor: default !important; font-weight: 400 !important; }
}
.itc-linked-client-row { display: flex; align-items: center; gap: 6px; font-size: 12px; color: var(--muted); margin: -4px 0 10px; }
.itc-linked-client-row .pan-hyperlink { font-size: 12px; }
`;

// ─── ITR Computation — helpers ────────────────────────────────────────────────
const CLIENT_CATEGORY_MAP = { Individual:"Individual", Proprietorship:"Individual", Company:"Company", LLP:"LLP", Partnership:"Firm", HUF:"HUF", Trust:"Trust" };
function ayFromFY(fy){
  const m=(fy||"").match(/(\d{4})-(\d{2})/);
  if(!m) return null;
  const y1=parseInt(m[1],10);
  return `AY ${y1+1}-${String(y1+2).slice(-2)}`;
}
/** Builds a fresh assessee object for the Tax Computation tool, auto-populated
 *  from an existing client record. Every extra ITR-only field (Aadhaar, DOB,
 *  Father's Name, Gender, Residential Status, PIN, Bank details, ITR Type)
 *  is optional on the client record — if the client hasn't filled it in on
 *  Add/Edit Client, it's simply left blank here and can be typed directly
 *  into the computation instead. */
function clientToAssessee(client, fy){
  const ay = ayFromFY(fy) && TAX_CONFIG[ayFromFY(fy)] ? ayFromFY(fy) : defaultAssessee.ay;
  if(!client) return { ...defaultAssessee, ay, fy };
  return {
    ...defaultAssessee,
    ay, fy,
    pan: client.pan||"",
    name: client.name||"",
    address: client.addr||"",
    state: client.state||"",
    pin: client.pin||"",
    email: client.email||"",
    mobile: client.mob||"",
    aadhaar: client.aadhaar||"",
    dob: client.dob||"",
    fatherName: client.fatherName||"",
    gender: client.gender||defaultAssessee.gender,
    residentialStatus: client.residentialStatus||defaultAssessee.residentialStatus,
    category: CLIENT_CATEGORY_MAP[client.type]||defaultAssessee.category,
    bankName: client.bankName||"",
    ifsc: client.ifsc||"",
    accountNumber: client.accountNumber||"",
    itrType: client.itrType||defaultAssessee.itrType,
  };
}
const genCompId = () => "c_"+Date.now().toString(36)+Math.random().toString(36).slice(2,8);

// ─── ITR Computation Tab ───────────────────────────────────────────────────────
function ITRComputationTab({clients,setClients,computations,setComputations,dd,toast}){
  const configuredFYs = useMemo(()=>[...new Set(Object.values(TAX_CONFIG).map(c=>c.fy))],[]);
  const[fy,setFy]=useState(configuredFYs[0]||dd.fyOptions[0]);
  const[mode,setMode]=useState("list");
  const[draft,setDraft]=useState(null);
  const[showPicker,setShowPicker]=useState(false);
  const[pickPan,setPickPan]=useState("");
  const[newFy,setNewFy]=useState("");
  const[editClientPan,setEditClientPan]=useState(null);
  const[q,setQ]=useState("");

  const list=useMemo(()=>{
    const base=computations.filter(c=>c.fy===fy);
    if(!q) return base;
    const lq=q.toLowerCase();
    return base.filter(c=>(c.name||"").toLowerCase().includes(lq)||(c.pan||"").toLowerCase().includes(lq));
  },[computations,fy,q]);

  const openNew=()=>{setPickPan("");setNewFy(fy);setShowPicker(true);};
  const beginWithClient=(pan,useFy)=>{
    const client=clients.find(c=>c.pan===pan);
    const id=genCompId();
    const draftObj={id,fy:useFy,pan,name:client?.name||"",assessee:clientToAssessee(client,useFy),income:defaultIncome,deductions:defaultDeductions,manualDeductions:defaultManualDeductions,taxPaid:defaultTaxPaid,printRegime:"new",savedAt:null};
    setComputations(p=>[draftObj,...p]);
    setDraft(draftObj);
    setShowPicker(false);
    setFy(useFy);
    setMode("editor");
  };
  const openExisting=(rec)=>{setFy(rec.fy);setDraft(rec);setMode("editor");};
  const handleSave=(rec)=>{
    setComputations(p=>{
      const exists=p.some(c=>c.id===rec.id);
      return exists?p.map(c=>c.id===rec.id?rec:c):[rec,...p];
    });
    setDraft(rec);
  };
  const handleDelete=(id)=>{setComputations(p=>p.filter(c=>c.id!==id));toast("Computation deleted","err");};
  const handleDuplicate=(rec)=>{
    const id=genCompId();
    const clone={...rec,id,name:(rec.name||rec.assessee?.name||"")+" (Copy)",assessee:{...rec.assessee,name:(rec.assessee?.name||"")+" (Copy)"},savedAt:Date.now()};
    setComputations(p=>[clone,...p]);
    toast("Computation duplicated");
  };

  const smallBtn={fontSize:11,padding:"5px 10px",borderRadius:7,border:`1px solid ${G.bdr}`,background:"transparent",cursor:"pointer",fontWeight:600};

  if(mode==="editor"&&draft){
    return <>
      <TaxComputationEditor
        key={draft.id}
        initialRecord={draft}
        clients={clients}
        allComputations={computations}
        onSave={handleSave}
        onExit={()=>{setMode("list");setDraft(null);}}
        onOpenClientEdit={(pan)=>setEditClientPan(pan)}
        onOpenRecord={(rec)=>{setFy(rec.fy);setDraft(rec);}}
      />
      {editClientPan&&(()=>{const c=clients.find(x=>x.pan===editClientPan);return c?<EditClient c={c} dd={dd} onX={()=>setEditClientPan(null)} onSave={cf=>{setClients(p=>p.map(x=>x.pan===cf.pan?cf:x));setEditClientPan(null);toast("Client updated!");}}/>:null;})()}
    </>;
  }

  return <div style={{display:"flex",flexDirection:"column",gap:16}}>
    <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap",justifyContent:"space-between"}}>
      <div style={{display:"flex",gap:6,flexWrap:"wrap",alignItems:"center"}}>
        <span style={{fontSize:12,color:G.mut,fontWeight:600}}>Financial Year:</span>
        {configuredFYs.map(f=><button key={f} onClick={()=>setFy(f)} style={{padding:"6px 14px",borderRadius:20,border:`1.5px solid ${fy===f?G.green:G.bdr}`,cursor:"pointer",fontSize:12,fontWeight:700,background:fy===f?G.green+"18":"transparent",color:fy===f?G.green:G.mut}}>
          FY {f} <span style={{opacity:.6}}>({computations.filter(c=>c.fy===f).length})</span>
        </button>)}
        {dd.fyOptions.filter(f=>!configuredFYs.includes(f)).map(f=><span key={f} title="Tax rules for this year aren't configured yet — add an entry to TAX_CONFIG to enable it" style={{padding:"6px 14px",borderRadius:20,border:`1.5px dashed ${G.bdr}`,fontSize:12,color:G.bdr}}>FY {f} 🔒</span>)}
      </div>
      <Btn onClick={openNew}>➕ New Computation</Btn>
    </div>

    <div style={{position:"relative",maxWidth:340}}>
      <span style={{position:"absolute",left:11,top:"50%",transform:"translateY(-50%)",color:G.mut}}>🔍</span>
      <input value={q} onChange={e=>setQ(e.target.value)} placeholder="Search name or PAN..." style={{...IS,paddingLeft:34}}/>
    </div>

    {list.length===0
      ? <Crd sty={{textAlign:"center",padding:44,color:G.mut}}>
          <div style={{fontSize:34,marginBottom:10}}>🧮</div>
          No ITR computations yet for FY {fy}.<br/>Click "New Computation" to start one for an existing client.
        </Crd>
      : <div className="auto-grid" style={{display:"grid",gridTemplateColumns:"repeat(auto-fill,minmax(300px,1fr))",gap:13}}>
          {list.map(rec=><div key={rec.id} style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:15,overflow:"hidden"}}>
            <div style={{padding:"13px 15px",borderBottom:`1px solid ${G.bdr}`}}>
              <div style={{fontWeight:700,fontSize:13,color:G.wh,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{rec.assessee?.name||rec.name||"Untitled"}</div>
              <div style={{display:"flex",gap:8,marginTop:5,alignItems:"center",flexWrap:"wrap"}}>
                <span onClick={()=>setEditClientPan(rec.pan)} title="View / edit client details" style={{fontSize:11,color:G.cyn,fontWeight:700,fontFamily:"monospace",cursor:"pointer",textDecoration:"underline"}}>{rec.pan||"No PAN"}</span>
                <span style={{fontSize:11,color:G.mut}}>{rec.assessee?.ay}</span>
              </div>
            </div>
            <div style={{padding:"10px 15px",fontSize:11,color:G.mut}}>Saved: {rec.savedAt?new Date(rec.savedAt).toLocaleDateString("en-IN",{dateStyle:"medium"}):"Not saved yet"}</div>
            <div style={{padding:"9px 15px",borderTop:`1px solid ${G.bdr}`,display:"flex",gap:6,flexWrap:"wrap"}}>
              <button onClick={()=>openExisting(rec)} style={{...smallBtn,color:G.green}}>📂 Open</button>
              <button onClick={()=>handleDuplicate(rec)} style={{...smallBtn,color:G.amb}}>⧉ Duplicate</button>
              <button onClick={()=>handleDelete(rec.id)} style={{...smallBtn,color:G.red}}>🗑 Delete</button>
            </div>
          </div>)}
        </div>}

    {showPicker&&<div style={{position:"fixed",inset:0,background:"#000C",zIndex:5000,display:"flex",alignItems:"center",justifyContent:"center",padding:20}}>
      <div style={{background:G.surf,border:`1px solid ${G.bdr}`,borderRadius:18,width:"min(480px,96%)",padding:20}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:14}}>
          <span style={{fontWeight:800,fontSize:15,color:G.wh}}>🧮 New ITR Computation</span>
          <button onClick={()=>setShowPicker(false)} style={{background:"transparent",border:`1px solid ${G.bdr}`,color:G.mut,borderRadius:8,padding:"4px 11px",cursor:"pointer"}}>✕</button>
        </div>
        <F label="Financial Year">
          <S val={newFy} set={setNewFy} opts={dd.fyOptions} ph="Select Financial Year..."/>
        </F>
        {!newFy
          ? <div style={{fontSize:12,color:G.mut,marginTop:8}}>Choose the Financial Year first — every year has its own slabs, so the computation is built for that year's rules.</div>
        : !configuredFYs.includes(newFy)
          ? <div style={{marginTop:10,padding:"12px 14px",background:G.amb+"14",border:`1px solid ${G.amb}44`,borderRadius:10,fontSize:12,color:G.amb}}>
              🔒 Tax rules for FY {newFy} haven't been added yet. FY 2025-26 is available now — other years will be added later.
            </div>
          : <>
              <F label="Select Client">
                <PanPick clients={clients} val={pickPan} set={setPickPan}/>
              </F>
              <div style={{fontSize:11,color:G.mut,marginTop:8}}>PAN, name, address, contact & any ITR details already on the client's profile will auto-fill. Anything missing can be typed directly into the computation.</div>
            </>}
        <div style={{display:"flex",gap:10,marginTop:16}}>
          <Btn onClick={()=>newFy&&configuredFYs.includes(newFy)&&pickPan&&beginWithClient(pickPan,newFy)} sty={{flex:1,padding:12,fontSize:14,opacity:(newFy&&configuredFYs.includes(newFy)&&pickPan)?1:.5,cursor:(newFy&&configuredFYs.includes(newFy)&&pickPan)?"pointer":"not-allowed"}}>Start Computation</Btn>
          <button onClick={()=>setShowPicker(false)} style={{padding:"12px 18px",borderRadius:10,border:`1px solid ${G.bdr}`,background:"transparent",color:G.mut,cursor:"pointer"}}>Cancel</button>
        </div>
      </div>
    </div>}

    {editClientPan&&(()=>{const c=clients.find(x=>x.pan===editClientPan);return c?<EditClient c={c} dd={dd} onX={()=>setEditClientPan(null)} onSave={cf=>{setClients(p=>p.map(x=>x.pan===cf.pan?cf:x));setEditClientPan(null);toast("Client updated!");}}/>:null;})()}
  </div>;
}

function AdminLogin({ setSession }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  const handleLogin = async (e) => {
    e.preventDefault();
    setLoading(true);
    setError("");
    const { data, error } = await supabase.auth.signInWithPassword({
      email,
      password,
    });
    if (error) {
      setError(error.message);
    } else {
      setSession(data.session);
    }
    setLoading(false);
  };

  return (
    <div style={{
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      minHeight: "100vh",
      background: "#060E0A",
      fontFamily: "Segoe UI, sans-serif",
      color: "#F0FDF4",
      padding: "20px"
    }}>
      <div style={{
        background: "#0B1610",
        border: "1px solid #1A3020",
        borderRadius: "20px",
        padding: "40px",
        width: "100%",
        maxWidth: "400px",
        boxShadow: "0 20px 50px rgba(0,0,0,0.8)",
        textAlign: "center"
      }}>
        <div style={{ display: "flex", justifyContent: "center", marginBottom: "20px" }}>
          <Logo sz={64} />
        </div>
        <h2 style={{ fontSize: "22px", fontWeight: "800", color: "#F0FDF4", marginBottom: "6px" }}>Fintax Mitra</h2>
        <p style={{ fontSize: "13px", color: "#86EFAC", marginBottom: "30px" }}>Admin Portal Login</p>
        
        {error && (
          <div style={{
            background: "#450A0A",
            border: "1px solid #EF4444",
            color: "#FCA5A5",
            borderRadius: "8px",
            padding: "10px 14px",
            fontSize: "12px",
            marginBottom: "20px",
            textAlign: "left"
          }}>
            ✕ {error}
          </div>
        )}
        
        <form onSubmit={handleLogin} style={{ display: "flex", flexDirection: "column", gap: "16px", textAlign: "left" }}>
          <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
            <label style={{ fontSize: "11px", fontWeight: "700", color: "#86EFAC", textTransform: "uppercase" }}>Email Address</label>
            <input 
              type="email" 
              required 
              value={email} 
              onChange={e => setEmail(e.target.value)} 
              placeholder="admin@fintaxmitra.com"
              style={{
                background: "#0F1D14",
                border: "1.5px solid #1A3020",
                borderRadius: "8px",
                color: "#F0FDF4",
                fontSize: "13px",
                padding: "10px 14px",
                outline: "none"
              }}
            />
          </div>
          
          <div style={{ display: "flex", flexDirection: "column", gap: "6px" }}>
            <label style={{ fontSize: "11px", fontWeight: "700", color: "#86EFAC", textTransform: "uppercase" }}>Password</label>
            <input 
              type="password" 
              required 
              value={password} 
              onChange={e => setPassword(e.target.value)} 
              placeholder="••••••••"
              style={{
                background: "#0F1D14",
                border: "1.5px solid #1A3020",
                borderRadius: "8px",
                color: "#F0FDF4",
                fontSize: "13px",
                padding: "10px 14px",
                outline: "none"
              }}
            />
          </div>

          <button 
            type="submit" 
            disabled={loading}
            style={{
              padding: "12px",
              borderRadius: "10px",
              border: "none",
              background: "linear-gradient(135deg, #16A34A, #22C55E)",
              color: "#fff",
              fontWeight: "700",
              fontSize: "14px",
              cursor: "pointer",
              marginTop: "10px",
              transition: "opacity 0.2s"
            }}
          >
            {loading ? "Authenticating..." : "Sign In"}
          </button>
        </form>
      </div>
    </div>
  );
}

export default function App(){
  const [session, setSession] = useState(null);
  const [authLoading, setAuthLoading] = useState(true);
  const [dbLoading, setDbLoading] = useState(true);
  const [syncEnabled, setSyncEnabled] = useState(false);

  const[clients, _setClients]=useState([]);
  const[works, _setWorks]=useState([]);
  const[invoices, _setInvoices]=useState([]);
  const[receipts, _setReceipts]=useState([]);
  const[firmSettings, _setFirmSettings]=useState({
    name:"Fin-Tax Mitra",
    addr:"17, Sebadol Road, Belgharia, Kolkata, West Bengal - 700049",
    phone:"7980718092",
    email:"care.fintaxmitra@gmail.com",
    bankName:"Kotak Mahindra Bank",
    bankHolder:"BILTU DEY",
    bankAcc:"1449547644",
    bankIFSC:"KKBK0000328",
    upiId:"7319440039@kotak",
    terms:"Fees once paid are non-refundable. The client is responsible for ensuring data accuracy and timely submission. Fin-Tax Mitra is not liable for any penalties or losses resulting from client delays or incorrect data.",
    logo:null, stamp:null, signature:null, qrCode:null, statusStamp:null,
    autoBackup:true,
    googleClientId:"738596578042-qd24uv0mkqe5j9bvjm8d4blpntg3vm7b.apps.googleusercontent.com",
    googleClientSecret:"",
    googleBackupEmail:"",

    googleDriveEnabled:false,
  });
  const[dd, _setDd]=useState(DEF_DD);
  const[pws, _setPws]=useState(DEF_PW);
  const[computations, _setComputations]=useState([]);

  // Database Sync functions
  const syncClientsToSupabase = async (prev, next) => {
    try {
      const deleted = prev.filter(p => !next.some(n => n.pan === p.pan));
      for (const c of deleted) {
        const { error } = await supabase.from('clients').delete().eq('pan', c.pan);
        if (error) throw error;
      }
      const addedOrUpdated = next.filter(n => {
        const p = prev.find(x => x.pan === n.pan);
        return !p || JSON.stringify(p) !== JSON.stringify(n);
      });
      if (addedOrUpdated.length > 0) {
        const { error } = await supabase.from('clients').upsert(addedOrUpdated);
        if (error) throw error;
      }
    } catch (err) {
      console.error("Error syncing clients:", err);
      toast("Error syncing clients: " + err.message, "err");
    }
  };

  const syncWorksToSupabase = async (prev, next) => {
    try {
      const deleted = prev.filter(p => !next.some(n => n.id === p.id));
      for (const w of deleted) {
        const { error } = await supabase.from('works').delete().eq('id', w.id);
        if (error) throw error;
      }
      const addedOrUpdated = next.filter(n => {
        const p = prev.find(x => x.id === n.id);
        return !p || JSON.stringify(p) !== JSON.stringify(n);
      });
      if (addedOrUpdated.length > 0) {
        const { error } = await supabase.from('works').upsert(addedOrUpdated);
        if (error) throw error;
      }
    } catch (err) {
      console.error("Error syncing works:", err);
      toast("Error syncing works: " + err.message, "err");
    }
  };

  const syncInvoicesToSupabase = async (prev, next) => {
    try {
      const deleted = prev.filter(p => !next.some(n => n.id === p.id));
      for (const inv of deleted) {
        const { error } = await supabase.from('invoices').delete().eq('id', inv.id);
        if (error) throw error;
      }
      const addedOrUpdated = next.filter(n => {
        const p = prev.find(x => x.id === n.id);
        return !p || JSON.stringify(p) !== JSON.stringify(n);
      });
      if (addedOrUpdated.length > 0) {
        const mapped = addedOrUpdated.map(inv => ({
          ...inv,
          workId: inv.workId === "" || inv.workId == null ? null : Number(inv.workId)
        }));
        const { error } = await supabase.from('invoices').upsert(mapped);
        if (error) throw error;
      }
    } catch (err) {
      console.error("Error syncing invoices:", err);
      toast("Error syncing invoices: " + err.message, "err");
    }
  };

  const syncReceiptsToSupabase = async (prev, next) => {
    try {
      const deleted = prev.filter(p => !next.some(n => n.id === p.id));
      for (const r of deleted) {
        const { error } = await supabase.from('receipts').delete().eq('id', r.id);
        if (error) throw error;
      }
      const addedOrUpdated = next.filter(n => {
        const p = prev.find(x => x.id === n.id);
        return !p || JSON.stringify(p) !== JSON.stringify(n);
      });
      if (addedOrUpdated.length > 0) {
        const mapped = addedOrUpdated.map(r => ({
          id: r.id,
          date: r.date,
          invId: r.invId,
          pan: r.pan,
          clientName: r.clientName,
          invAmt: Number(r.invAmt) || 0,
          received: Number(r.received) || 0,
          mode: r.mode,
          ref: r.ref || null,
          note: r.note || null
        }));
        const { error } = await supabase.from('receipts').upsert(mapped);
        if (error) throw error;
      }
    } catch (err) {
      console.error("Error syncing receipts:", err);
      toast("Error syncing receipts: " + err.message, "err");
    }
  };

  const syncComputationsToSupabase = async (prev, next) => {
    try {
      const deleted = prev.filter(p => !next.some(n => n.id === p.id));
      for (const comp of deleted) {
        const { error } = await supabase.from('computations').delete().eq('id', comp.id);
        if (error) throw error;
      }
      const addedOrUpdated = next.filter(n => {
        const p = prev.find(x => x.id === n.id);
        return !p || JSON.stringify(p) !== JSON.stringify(n);
      });
      if (addedOrUpdated.length > 0) {
        const { error } = await supabase.from('computations').upsert(addedOrUpdated);
        if (error) throw error;
      }
    } catch (err) {
      console.error("Error syncing computations:", err);
      toast("Error syncing computations: " + err.message, "err");
    }
  };

  const syncFirmSettingsToSupabase = async (settings) => {
    try {
      // Strip heavy base64 images to avoid "Failed to fetch" payload size errors
      const lightSettings = { ...settings };
      const imageKeys = ['logo', 'stamp', 'signature', 'qrCode', 'statusStamp'];
      imageKeys.forEach(k => {
        if (lightSettings[k] && lightSettings[k].startsWith('data:image')) {
          lightSettings[k] = null;
        }
      });
      const { error } = await supabase.from('firm_settings').upsert({ id: 1, settings: lightSettings });
      if (error) throw error;
    } catch (err) {
      console.error("Error syncing firm settings:", err);
      toast("Error syncing firm settings: " + err.message, "err");
    }
  };

  const syncDevSettingsToSupabase = async (dropdown_defaults, passwords) => {
    try {
      const { error } = await supabase.from('developer_settings').upsert({ id: 1, dropdown_defaults, passwords });
      if (error) throw error;
    } catch (err) {
      console.error("Error syncing developer settings:", err);
      toast("Error syncing developer settings: " + err.message, "err");
    }
  };

  // State wrappers to automatically sync changes
  const setClients = (valOrFn) => {
    _setClients(prev => {
      const next = typeof valOrFn === 'function' ? valOrFn(prev) : valOrFn;
      if (syncEnabled) setTimeout(() => syncClientsToSupabase(prev, next), 0);
      return next;
    });
  };

  const setWorks = (valOrFn) => {
    _setWorks(prev => {
      const next = typeof valOrFn === 'function' ? valOrFn(prev) : valOrFn;
      if (syncEnabled) setTimeout(() => syncWorksToSupabase(prev, next), 0);
      return next;
    });
  };

  const setInvoices = (valOrFn) => {
    _setInvoices(prev => {
      const next = typeof valOrFn === 'function' ? valOrFn(prev) : valOrFn;
      if (syncEnabled) setTimeout(() => syncInvoicesToSupabase(prev, next), 0);
      return next;
    });
  };

  const setReceipts = (valOrFn) => {
    _setReceipts(prev => {
      const next = typeof valOrFn === 'function' ? valOrFn(prev) : valOrFn;
      if (syncEnabled) setTimeout(() => syncReceiptsToSupabase(prev, next), 0);
      return next;
    });
  };

  const setComputations = (valOrFn) => {
    _setComputations(prev => {
      const next = typeof valOrFn === 'function' ? valOrFn(prev) : valOrFn;
      if (syncEnabled) setTimeout(() => syncComputationsToSupabase(prev, next), 0);
      return next;
    });
  };

  const setFirmSettings = (valOrFn) => {
    _setFirmSettings(prev => {
      const next = typeof valOrFn === 'function' ? valOrFn(prev) : valOrFn;
      if (syncEnabled) setTimeout(() => syncFirmSettingsToSupabase(next), 0);
      return next;
    });
  };

  const setDd = (valOrFn) => {
    _setDd(prev => {
      const next = typeof valOrFn === 'function' ? valOrFn(prev) : valOrFn;
      if (syncEnabled) setTimeout(() => syncDevSettingsToSupabase(next, pws), 0);
      return next;
    });
  };

  const setPws = (valOrFn) => {
    _setPws(prev => {
      const next = typeof valOrFn === 'function' ? valOrFn(prev) : valOrFn;
      if (syncEnabled) setTimeout(() => syncDevSettingsToSupabase(dd, next), 0);
      return next;
    });
  };

  // Auth session listener
  useEffect(() => {
    supabase.auth.getSession().then(({ data: { session } }) => {
      setSession(session);
      setAuthLoading(false);
    });
    const { data: { subscription } } = supabase.auth.onAuthStateChange((_event, newSession) => {
      setSession(prev => {
        if (prev?.user?.id === newSession?.user?.id) {
          return prev;
        }
        return newSession;
      });
    });
    return () => subscription.unsubscribe();
  }, []);

  // Persist heavy firm images locally to avoid database payload limits
  useEffect(() => {
    const imageKeys = ['logo', 'stamp', 'signature', 'qrCode', 'statusStamp'];
    const images = {};
    imageKeys.forEach(k => {
      if (firmSettings[k]) {
        images[k] = firmSettings[k];
      }
    });
    if (Object.keys(images).length > 0) {
      try {
        localStorage.setItem("ftm_firm_images", JSON.stringify(images));
      } catch (e) {
        console.error("Error saving images locally:", e);
      }
    }
  }, [firmSettings]);

  // Database fetch & seeding
  useEffect(() => {
    if (!session) return;
    const loadData = async () => {
      if (!syncEnabled) setDbLoading(true);
      try {
        let { data: dbClients, error: errClients } = await supabase.from('clients').select('*');
        if (errClients) throw errClients;
        let { data: dbWorks, error: errWorks } = await supabase.from('works').select('*');
        if (errWorks) throw errWorks;
        let { data: dbInvoices, error: errInvoices } = await supabase.from('invoices').select('*');
        if (errInvoices) throw errInvoices;
        if (dbInvoices) {
          dbInvoices = dbInvoices.map(inv => ({
            ...inv,
            workId: inv.workId == null ? "" : String(inv.workId)
          }));
        }
        let { data: dbReceipts, error: errReceipts } = await supabase.from('receipts').select('*');
        if (errReceipts) throw errReceipts;
        if (dbInvoices && dbReceipts) {
          dbInvoices = dbInvoices.map(inv => ({
            ...inv,
            status: computeInvStatus(inv, dbReceipts)
          }));
        }
        let { data: dbComputations, error: errComputations } = await supabase.from('computations').select('*');
        if (errComputations) throw errComputations;
        let { data: dbFirm, error: errFirm } = await supabase.from('firm_settings').select('*').maybeSingle();
        if (errFirm) throw errFirm;
        let { data: dbDev, error: errDev } = await supabase.from('developer_settings').select('*').maybeSingle();
        if (errDev) throw errDev;

        // Cleanup dummy data from Supabase DB using active session
        const dummyPans = ["ABCPK1234A", "BCDQL5678B", "CDERM7890C", "DEFNS2345D", "EFGOT6789E"];
        supabase.from('receipts').delete().in('pan', dummyPans).then();
        supabase.from('invoices').delete().in('pan', dummyPans).then();
        supabase.from('computations').delete().in('pan', dummyPans).then();
        supabase.from('works').delete().in('pan', dummyPans).then();
        supabase.from('clients').delete().in('pan', dummyPans).then();

        // Local filter to immediately strip dummy data from UI loads
        dbClients = (dbClients || []).filter(c => !dummyPans.includes(c.pan));
        dbWorks = (dbWorks || []).filter(w => !dummyPans.includes(w.pan));
        dbInvoices = (dbInvoices || []).filter(i => !dummyPans.includes(i.pan));
        dbReceipts = (dbReceipts || []).filter(r => !dummyPans.includes(r.pan));
        if (dbComputations) {
          dbComputations = dbComputations.filter(c => !dummyPans.includes(c.pan));
        }

        if (!dbFirm) {
          const defaultFirm = { id: 1, settings: {
            name:"Fin-Tax Mitra",
            addr:"17, Sebadol Road, Belgharia, Kolkata, West Bengal - 700049",
            phone:"7980718092",
            email:"care.fintaxmitra@gmail.com",
            bankName:"Kotak Mahindra Bank",
            bankHolder:"BILTU DEY",
            bankAcc:"1449547644",
            bankIFSC:"KKBK0000328",
            upiId:"7319440039@kotak",
            terms:"Fees once paid are non-refundable. The client is responsible for ensuring data accuracy and timely submission. Fin-Tax Mitra is not liable for any penalties or losses resulting from client delays or incorrect data.",
            logo:null, stamp:null, signature:null, qrCode:null, statusStamp:null,
            autoBackup:true,
            googleClientId:"738596578042-qd24uv0mkqe5j9bvjm8d4blpntg3vm7b.apps.googleusercontent.com",
            googleDriveEnabled:false,
          }};
          const { error } = await supabase.from('firm_settings').insert(defaultFirm);
          if (error) throw new Error("Seeding firm settings failed: " + error.message);
          dbFirm = defaultFirm;
        }
        if (!dbDev) {
          const defaultDev = { id: 1, dropdown_defaults: DEF_DD, passwords: DEF_PW };
          const { error } = await supabase.from('developer_settings').insert(defaultDev);
          if (error) throw new Error("Seeding dev settings failed: " + error.message);
          dbDev = defaultDev;
        }

        _setClients(dbClients);
        _setWorks(dbWorks);
        _setInvoices(dbInvoices);
        _setReceipts(dbReceipts);
        _setComputations(dbComputations || []);
        let localImages = {};
        try {
          const saved = localStorage.getItem("ftm_firm_images");
          if (saved) localImages = JSON.parse(saved);
        } catch (e) {
          console.error("Error loading local images:", e);
        }
        _setFirmSettings({ ...dbFirm.settings, ...localImages });
        _setDd(dbDev.dropdown_defaults);
        _setPws(dbDev.passwords);
        setSyncEnabled(true);
      } catch (err) {
        console.error("Error loading data from Supabase:", err);
        toast("Database load failed: " + err.message, "err");
      } finally {
        setDbLoading(false);
      }
    };
    loadData();
  }, [session]);
  const[tab,setTab]=useState("wdash"),[sub,setSub]=useState("client");
  const[toasts,setToasts]=useState([]);
  const[showOA,setShowOA]=useState(false),[showDA,setShowDA]=useState(false);
  const[ownerOn,setOwnerOn]=useState(false),[devOn,setDevOn]=useState(false);
  const[pendingProt,setPendingProt]=useState("fin");
  const[sideOpen,setSideOpen]=useState(typeof window==="undefined"||window.innerWidth>768);
  const[isMobile,setIsMobile]=useState(typeof window!=="undefined"&&window.innerWidth<=768);
  useEffect(()=>{
    if(typeof document==="undefined")return;
    if(!document.querySelector('meta[name="viewport"]')){
      const m=document.createElement("meta");
      m.name="viewport";
      m.content="width=device-width, initial-scale=1, maximum-scale=1";
      document.head.appendChild(m);
    }
  },[]);
  useEffect(()=>{
    const onResize=()=>{
      const m=window.innerWidth<=768;
      setIsMobile(prev=>{if(m!==prev){setSideOpen(!m);} return m;});
    };
    window.addEventListener("resize",onResize);
    return ()=>window.removeEventListener("resize",onResize);
  },[]);

  // Dynamically load Google Identity Services SDK
  useEffect(() => {
    if (typeof window === "undefined" || document.getElementById("google-gsi-client")) return;
    const script = document.createElement("script");
    script.id = "google-gsi-client";
    script.src = "https://accounts.google.com/gsi/client";
    script.async = true;
    script.defer = true;
    document.body.appendChild(script);
  }, []);

  const[darkMode,setDarkMode]=useState(true);
  const[profilePic,setProfilePic]=useState(null);
  // Apply theme globally
  Object.assign(G, darkMode ? DARK_THEME : LIGHT_THEME);
  const toast=(msg,tp="ok")=>{const id=Date.now();setToasts(t=>[...t,{id,msg,tp}]);setTimeout(()=>setToasts(t=>t.filter(x=>x.id!==id)),3200);};

  const [googleUser, setGoogleUser] = useState(() => {
    try {
      const saved = localStorage.getItem("ftm_google_user");
      return saved ? JSON.parse(saved) : null;
    } catch (e) {
      return null;
    }
  });

  const getBackupChecksum = () => {
    try {
      const payload = { clients, works, invoices, receipts, dd, pws, firmSettings };
      const str = JSON.stringify(payload);
      let hash = 0;
      for (let i = 0; i < str.length; i++) {
        const char = str.charCodeAt(i);
        hash = (hash << 5) - hash + char;
        hash = hash & hash;
      }
      return String(hash) + "_" + str.length;
    } catch (e) {
      return "";
    }
  };

  const linkGoogleDrive = () => {
    if (typeof window === "undefined" || !window.google) {
      toast("Google Identity Services not loaded yet, please try again in a few seconds.", "err");
      return;
    }
    const clientId = firmSettings.googleClientId?.trim();
    const clientSecret = firmSettings.googleClientSecret?.trim();
    if (!clientId) {
      toast("Please enter a valid Google Client ID first.", "err");
      return;
    }
    try {
      if (clientSecret) {
        const codeClient = window.google.accounts.oauth2.initCodeClient({
          client_id: clientId,
          scope: "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/userinfo.email",
          ux_mode: "popup",
          callback: async (response) => {
            if (response.error) {
              toast("Google link failed: " + response.error, "err");
              return;
            }
            if (response.code) {
              try {
                const tokenRes = await fetch("https://oauth2.googleapis.com/token", {
                  method: "POST",
                  headers: { "Content-Type": "application/x-www-form-urlencoded" },
                  body: new URLSearchParams({
                    code: response.code,
                    client_id: clientId,
                    client_secret: clientSecret,
                    redirect_uri: window.location.origin,
                    grant_type: "authorization_code"
                  })
                });
                const tokenData = await tokenRes.json();
                if (tokenData.access_token) {
                  let userEmail = "Connected Account";
                  try {
                    const infoRes = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
                      headers: { Authorization: `Bearer ${tokenData.access_token}` }
                    });
                    const info = await infoRes.json();
                    if (info.email) userEmail = info.email;
                  } catch (e) {}

                  const userObj = {
                    email: userEmail,
                    accessToken: tokenData.access_token,
                    refreshToken: tokenData.refresh_token || null,
                    expiresAt: Date.now() + (Number(tokenData.expires_in) || 3600) * 1000
                  };
                  setGoogleUser(userObj);
                  localStorage.setItem("ftm_google_user", JSON.stringify(userObj));
                  toast("Google Drive linked with silent auto-backup!", "ok");
                  return;
                }
              } catch (err) {
                console.error("Code exchange error:", err);
              }
            }
          }
        });
        codeClient.requestCode();
        return;
      }

      const client = window.google.accounts.oauth2.initTokenClient({
        client_id: clientId,
        scope: "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/userinfo.email",
        callback: async (response) => {
          if (response.error) {
            toast("Google link failed: " + response.error, "err");
            return;
          }
          if (response.access_token) {
            try {
              const res = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
                headers: { Authorization: `Bearer ${response.access_token}` }
              });
              const info = await res.json();
              const userObj = {
                email: info.email || "Connected Account",
                accessToken: response.access_token,
                refreshToken: null,
                expiresAt: Date.now() + (Number(response.expires_in) || 3600) * 1000
              };
              setGoogleUser(userObj);
              localStorage.setItem("ftm_google_user", JSON.stringify(userObj));
              toast("Google Drive linked successfully!", "ok");
            } catch (e) {
              const userObj = {
                email: "Connected Account",
                accessToken: response.access_token,
                refreshToken: null,
                expiresAt: Date.now() + (Number(response.expires_in) || 3600) * 1000
              };
              setGoogleUser(userObj);
              localStorage.setItem("ftm_google_user", JSON.stringify(userObj));
              toast("Google Drive linked!", "ok");
            }
          }
        }
      });
      client.requestAccessToken({ prompt: "consent" });
    } catch (err) {
      toast("Google Auth initialization error: " + err.message, "err");
    }
  };

  const disconnectGoogleDrive = () => {
    setGoogleUser(null);
    localStorage.removeItem("ftm_google_user");
    toast("Google Drive disconnected", "ok");
  };

  const refreshGoogleTokenSilently = async (refreshToken) => {
    const clientId = firmSettings.googleClientId?.trim();
    const clientSecret = firmSettings.googleClientSecret?.trim();
    if (!clientId || !refreshToken) return null;

    try {
      const params = new URLSearchParams({
        client_id: clientId,
        grant_type: "refresh_token",
        refresh_token: refreshToken
      });
      if (clientSecret) {
        params.append("client_secret", clientSecret);
      }
      const res = await fetch("https://oauth2.googleapis.com/token", {
        method: "POST",
        headers: { "Content-Type": "application/x-www-form-urlencoded" },
        body: params
      });
      const data = await res.json();
      if (data.access_token) {
        const userObj = {
          ...googleUser,
          accessToken: data.access_token,
          expiresAt: Date.now() + (Number(data.expires_in) || 3600) * 1000,
          refreshToken: refreshToken
        };
        setGoogleUser(userObj);
        localStorage.setItem("ftm_google_user", JSON.stringify(userObj));
        return data.access_token;
      }
    } catch (e) {
      console.error("Silent token refresh failed:", e);
    }
    return null;
  };

  const getGoogleAccessToken = async (interactive = true) => {
    if (!googleUser) {
      throw new Error("Google account not linked.");
    }
    if (googleUser.accessToken && googleUser.expiresAt > Date.now() + 300 * 1000) {
      return googleUser.accessToken;
    }

    if (googleUser.refreshToken) {
      const freshToken = await refreshGoogleTokenSilently(googleUser.refreshToken);
      if (freshToken) return freshToken;
    }

    if (!interactive) {
      throw new Error("Google access token expired.");
    }

    return new Promise((resolve, reject) => {
      if (typeof window === "undefined" || !window.google) {
        reject(new Error("Google library not loaded"));
        return;
      }
      const clientId = firmSettings.googleClientId?.trim();
      if (!clientId) {
        reject(new Error("Google Client ID is missing"));
        return;
      }
      try {
        const client = window.google.accounts.oauth2.initTokenClient({
          client_id: clientId,
          scope: "https://www.googleapis.com/auth/drive.file https://www.googleapis.com/auth/userinfo.email",
          callback: async (response) => {
            if (response.error) {
              reject(new Error("Token refresh failed: " + response.error));
              return;
            }
            if (response.access_token) {
              try {
                const res = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
                  headers: { Authorization: `Bearer ${response.access_token}` }
                });
                const info = await res.json();
                const userObj = {
                  email: info.email || googleUser.email,
                  accessToken: response.access_token,
                  refreshToken: googleUser.refreshToken || null,
                  expiresAt: Date.now() + (Number(response.expires_in) || 3600) * 1000
                };
                setGoogleUser(userObj);
                localStorage.setItem("ftm_google_user", JSON.stringify(userObj));
                resolve(response.access_token);
              } catch (e) {
                const userObj = {
                  email: googleUser.email,
                  accessToken: response.access_token,
                  refreshToken: googleUser.refreshToken || null,
                  expiresAt: Date.now() + (Number(response.expires_in) || 3600) * 1000
                };
                setGoogleUser(userObj);
                localStorage.setItem("ftm_google_user", JSON.stringify(userObj));
                resolve(response.access_token);
              }
            } else {
              reject(new Error("No token returned"));
            }
          }
        });
        client.requestAccessToken({ prompt: "" });
      } catch (err) {
        reject(err);
      }
    });
  };

  const uploadBackupToGoogleDrive = async (isManual = false) => {
    try {
      const token = await getGoogleAccessToken(isManual);
      if (isManual) toast("Uploading backup to Google Drive...", "ok");
      
      let folderId = "";
      const searchRes = await fetch(
        "https://www.googleapis.com/drive/v3/files?q=name='FinTaxMitra Backups' and mimeType='application/vnd.google-apps.folder' and trashed=false",
        { headers: { Authorization: `Bearer ${token}` } }
      );
      const searchData = await searchRes.json();
      if (searchData.files && searchData.files.length > 0) {
        folderId = searchData.files[0].id;
      } else {
        const createRes = await fetch("https://www.googleapis.com/drive/v3/files", {
          method: "POST",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": "application/json"
          },
          body: JSON.stringify({
            name: "FinTaxMitra Backups",
            mimeType: "application/vnd.google-apps.folder"
          })
        });
        const folder = await createRes.json();
        folderId = folder.id;
      }
      
      if (!folderId) throw new Error("Could not find or create Google Drive folder");
      
      const payload = {
        meta: { app: "Fin-Tax Mitra", exportedAt: new Date().toISOString(), version: 1, type: "google_auto" },
        clients, works, invoices, receipts, dd, pws, firmSettings, darkMode
      };
      const jsonContent = JSON.stringify(payload, null, 2);
      
      const d = new Date();
      const p = n => String(n).padStart(2, "0");
      const filename = `FinTaxMitra_AutoBackup_${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}.json`;
      
      const boundary = "-------314159265358979323846";
      const delimiter = "\r\n--" + boundary + "\r\n";
      const closeDelim = "\r\n--" + boundary + "--";
      const metadata = { name: filename, mimeType: "application/json", parents: [folderId] };
      
      const multipartBody =
        delimiter +
        "Content-Type: application/json; charset=UTF-8\r\n\r\n" +
        JSON.stringify(metadata) +
        delimiter +
        "Content-Type: application/json\r\n\r\n" +
        jsonContent +
        closeDelim;
        
      const uploadRes = await fetch(
        "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart",
        {
          method: "POST",
          headers: {
            Authorization: `Bearer ${token}`,
            "Content-Type": `multipart/related; boundary=${boundary}`
          },
          body: multipartBody
        }
      );
      
      if (!uploadRes.ok) {
        const errorData = await uploadRes.json();
        throw new Error(errorData.error?.message || "Upload failed");
      }
      
      const currentChecksum = getBackupChecksum();
      localStorage.setItem("ftm_last_google_backup_checksum", currentChecksum);
      localStorage.setItem("ftm_last_google_backup", new Date().toDateString());
      toast("Backup uploaded to Google Drive!", "ok");
    } catch (err) {
      console.error("Google Drive upload error:", err);
      if (isManual) {
        toast("Google Drive upload failed: " + err.message, "err");
      } else {
        console.log("Auto-backup to Google Drive skipped:", err.message);
      }
    }
  };

  // Auto-backup handler (generates daily backups automatically on launch if data changed)
  useEffect(() => {
    if (!syncEnabled || !clients.length) return;
    
    const today = new Date().toDateString();
    const currentChecksum = getBackupChecksum();
    
    const runLocalBackup = () => {
      if (!firmSettings.autoBackup) return;
      const lastLocalBackup = localStorage.getItem("ftm_last_autobackup");
      const lastLocalChecksum = localStorage.getItem("ftm_last_local_backup_checksum");
      
      if (lastLocalBackup !== today) {
        if (currentChecksum === lastLocalChecksum) {
          console.log("Local auto-backup skipped: No changes since last backup");
          localStorage.setItem("ftm_last_autobackup", today);
          return;
        }
        
        try {
          const data = JSON.stringify({
            meta: { app: "Fin-Tax Mitra", exportedAt: new Date().toISOString(), version: 1, type: "auto" },
            clients, works, invoices, receipts, dd, pws, firmSettings, darkMode,
          }, null, 2);
          
          const d = new Date();
          const p = n => String(n).padStart(2, "0");
          const filename = `FinTaxMitra_AutoBackup_${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}.json`;
          
          const blob = new Blob([data], { type: "application/json" });
          const url = URL.createObjectURL(blob);
          const a = document.createElement("a");
          a.href = url; a.download = filename; document.body.appendChild(a); a.click();
          document.body.removeChild(a); URL.revokeObjectURL(url);
          
          localStorage.setItem("ftm_last_autobackup", today);
          localStorage.setItem("ftm_last_local_backup_checksum", currentChecksum);
          toast("Daily auto-backup downloaded!", "ok");
        } catch (e) {
          console.error("Local auto backup failed:", e);
        }
      }
    };

    const runGoogleBackup = () => {
      if (!firmSettings.googleDriveEnabled || !googleUser) return;
      const lastGoogleBackup = localStorage.getItem("ftm_last_google_backup");
      const lastGoogleChecksum = localStorage.getItem("ftm_last_google_backup_checksum");
      
      if (lastGoogleBackup !== today) {
        if (currentChecksum === lastGoogleChecksum) {
          console.log("Google Drive auto-backup skipped: No changes since last backup");
          localStorage.setItem("ftm_last_google_backup", today);
          return;
        }
        setTimeout(() => {
          uploadBackupToGoogleDrive(false);
        }, 3000);
      }
    };

    const timer = setTimeout(() => {
      runLocalBackup();
      runGoogleBackup();
    }, 5000);
    
    return () => clearTimeout(timer);
  }, [syncEnabled, firmSettings.autoBackup, firmSettings.googleDriveEnabled, googleUser, clients, works, invoices, receipts, dd, pws, darkMode]);
  const NAV=[
    {id:"wdash",icon:"⬡",label:"Work Dashboard"},
    {id:"fin",icon:"💰",label:"Finance",prot:true,go:()=>ownerOn?setTab("fin"):(setPendingProt("fin"),setShowOA(true))},
    null,
    {id:"add",icon:"＋",label:"Add / Assign"},
    {id:"clients",icon:"◉",label:"Clients"},
    {id:"itr",icon:"🧮",label:"ITR Computation"},
    {id:"tracker",icon:"▦",label:"Work Tracker"},
    {id:"invoice",icon:"🧾",label:"Invoices & Billing",prot:true,go:()=>ownerOn?setTab("invoice"):(setPendingProt("invoice"),setShowOA(true))},
    {id:"receipts",icon:"💳",label:"Payment Receipts"},
    null,
    {id:"dev",icon:"⚙️",label:"Developer",dev:true,go:()=>devOn?setTab("dev"):setShowDA(true)},
  ];
  const titles={wdash:"Work Dashboard",fin:"Finance Dashboard",add:"Add Client & Assign Work",clients:"Client Master",itr:"ITR Computation",tracker:"Work Tracker",invoice:"Invoices & Billing",receipts:"Payment Receipts",dev:"Developer Settings"};
  const navExpanded=isMobile?true:sideOpen;

  if (authLoading) {
    return (
      <div style={{
        background: "#060E0A",
        height: "100vh",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        color: "#22C55E",
        fontWeight: "bold",
        fontFamily: "Segoe UI, sans-serif"
      }}>
        Loading Admin Session...
      </div>
    );
  }
  if (!session) {
    return <AdminLogin setSession={setSession} />;
  }
  if (dbLoading) {
    return (
      <div style={{
        background: "#060E0A",
        height: "100vh",
        display: "flex",
        flexDirection: "column",
        gap: 15,
        alignItems: "center",
        justifyContent: "center",
        color: "#22C55E",
        fontWeight: "bold",
        fontFamily: "Segoe UI, sans-serif"
      }}>
        <Logo sz={64} />
        <span>Loading Fintax Mitra Database...</span>
      </div>
    );
  }

  return <div className="app-root" style={{display:"flex",height:"100vh",background:G.bg,fontFamily:"Segoe UI,system-ui,sans-serif",color:G.wh,overflow:"hidden",transition:"background .3s,color .3s"}}>
    <style>{`*{box-sizing:border-box;margin:0;padding:0;}::-webkit-scrollbar{width:5px;}::-webkit-scrollbar-track{background:${G.bg};}::-webkit-scrollbar-thumb{background:${G.bdr};border-radius:6px;}@keyframes sIn{from{transform:translateX(60px);opacity:0}to{transform:none;opacity:1}}@keyframes slidePanel{from{transform:translateX(100%)}to{transform:translateX(0)}}input::placeholder,textarea::placeholder{color:#223829;}input:focus,select:focus,textarea:focus{outline:none;}select option{background:#0B1610;}textarea{font-family:Segoe UI,sans-serif;}button:active{transform:scale(.97);}div:hover>.sidetip{opacity:1 !important;}
      .mobile-hamburger{display:none;}
      .app-header{padding:0 22px;}
      .app-content{padding:18px 22px;}
      @media (max-width:768px){
        .mobile-hamburger{display:flex !important;}
        .app-header{padding:0 12px !important;}
        .app-content{padding:12px !important;}
        .kpi-grid-3,.kpi-grid-4,.kpi-grid-5{grid-template-columns:repeat(2,1fr) !important;}
        table{font-size:11px !important;}
      }
      @media (max-width:430px){
        .kpi-grid-3,.kpi-grid-4,.kpi-grid-5{grid-template-columns:1fr !important;}
        .auto-grid{grid-template-columns:1fr !important;}
        .app-header .app-header-date{display:none !important;}
      }
      .overflow-x-touch{-webkit-overflow-scrolling:touch;}
      /* ── Print fix: when the ITR Computation sheet is printed, the outer
         app chrome (branded sidebar + black top header) must not appear,
         and the 100vh/overflow:hidden shell must not clip pages 2, 3, ... */
      @media print {
        .no-print { display: none !important; }
        html, body { height: auto !important; overflow: visible !important; }
        .app-root { display: block !important; height: auto !important; overflow: visible !important; background: #fff !important; }
        .app-content { overflow: visible !important; height: auto !important; padding: 0 !important; }
      }
    `}</style>
    {showOA&&<Auth title="Owner Access" hint="Finance Dashboard & Invoices/Billing password: 456" pws={{o:pws.owner}} onOk={()=>{setOwnerOn(true);setShowOA(false);setTab(pendingProt);toast("Owner access granted");}} onX={()=>setShowOA(false)}/>}
    {showDA&&<Auth title="Developer Access" hint="Developer Settings password: 123" pws={{d:pws.dev}} onOk={()=>{setDevOn(true);setShowDA(false);setTab("dev");toast("Developer access granted");}} onX={()=>setShowDA(false)}/>}
    {/* Mobile drawer backdrop */}
    {isMobile&&sideOpen&&<div onClick={()=>setSideOpen(false)} style={{position:"fixed",inset:0,background:"#000A",zIndex:9400}}/>}
    {/* Sidebar */}
    <div className="no-print" style={isMobile?{position:"fixed",top:0,left:sideOpen?0:-240,height:"100%",width:230,background:"#070E09",borderRight:`1px solid ${G.bdr}`,display:"flex",flexDirection:"column",zIndex:9500,transition:"left .25s cubic-bezier(.4,0,.2,1)",overflow:"hidden",boxShadow:sideOpen?"8px 0 30px #000A":"none"}:{width:sideOpen?218:58,background:"#070E09",borderRight:`1px solid ${G.bdr}`,display:"flex",flexDirection:"column",flexShrink:0,transition:"width .22s cubic-bezier(.4,0,.2,1)",overflow:"hidden",position:"relative"}}>
      {/* Logo + Toggle */}
      <div style={{padding:"12px 10px",borderBottom:`1px solid ${G.bdr}`,display:"flex",alignItems:"center",gap:10,minHeight:58,flexShrink:0}}>
        {navExpanded&&<>{profilePic
          ?<img src={profilePic} alt="Logo" style={{width:34,height:34,borderRadius:9,objectFit:"cover",border:`2px solid ${G.green}`,flexShrink:0}}/>
          :<Logo sz={34}/>}
        <div style={{flex:1,minWidth:0,overflow:"hidden"}}>
          <div style={{fontWeight:800,fontSize:13,color:G.wh,letterSpacing:-.3,whiteSpace:"nowrap"}}>Fin-Tax Mitra</div>
          <div style={{fontSize:10,color:G.mut,marginTop:1,whiteSpace:"nowrap"}}>Making Taxes Easy</div>
        </div></>}
        {!navExpanded&&<div style={{width:38,display:"flex",justifyContent:"center"}}>{profilePic?<img src={profilePic} alt="Logo" style={{width:32,height:32,borderRadius:8,objectFit:"cover",border:`2px solid ${G.green}`}}/>:<Logo sz={32}/>}</div>}
        <button onClick={()=>setSideOpen(p=>!p)} title={isMobile?"Close menu":sideOpen?"Collapse sidebar":"Expand sidebar"}
          style={{background:"transparent",border:`1px solid ${G.bdr}`,color:G.mut,borderRadius:7,width:26,height:26,cursor:"pointer",fontSize:13,display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,transition:"transform .2s",transform:isMobile?"none":sideOpen?"none":"rotate(180deg)"}}>{isMobile?"✕":String.fromCharCode(60)}</button>
      </div>

      {/* Nav items */}
      <nav style={{padding:"8px 6px",flex:1,display:"flex",flexDirection:"column",gap:1,overflowY:"auto",overflowX:"hidden"}}>
        {NAV.map((n,i)=>{
          if(n===null)return <div key={i} style={{height:1,background:G.bdr,margin:"5px 4px",flexShrink:0}}/>;
          const active=tab===n.id;
          const col=active?"#fff":n.dev?G.cyn:G.mut;
          return <div key={n.id} style={{position:"relative"}} title={!navExpanded?n.label:""}>
            <button onClick={()=>{n.go?n.go():setTab(n.id);if(isMobile)setSideOpen(false);}}
              style={{display:"flex",alignItems:"center",gap:navExpanded?9:0,padding:navExpanded?"8px 10px":"0",justifyContent:navExpanded?"flex-start":"center",width:"100%",height:38,borderRadius:9,border:"none",cursor:"pointer",transition:"all .15s",background:active?`linear-gradient(90deg,${G.g2},${G.green})`:"transparent",color:col,fontWeight:active?700:500,fontSize:13,flexShrink:0,overflow:"hidden"}}
              onMouseEnter={e=>{if(!active)e.currentTarget.style.background=G.bdr;}}
              onMouseLeave={e=>{if(!active)e.currentTarget.style.background="transparent";}}>
              <span style={{fontSize:16,flexShrink:0,width:20,textAlign:"center"}}>{n.icon}</span>
              {navExpanded&&<span style={{flex:1,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",fontSize:12}}>{n.label}</span>}
              {navExpanded&&n.prot&&!ownerOn&&<span style={{fontSize:10,opacity:.5,flexShrink:0}}>🔐</span>}
              {navExpanded&&n.dev&&!devOn&&<span style={{fontSize:9,background:`${G.cyn}20`,color:G.cyn,padding:"1px 4px",borderRadius:4,flexShrink:0}}>DEV</span>}
            </button>
            {/* Tooltip when collapsed */}
            {!navExpanded&&<div className="sidetip" style={{position:"absolute",left:52,top:"50%",transform:"translateY(-50%)",background:"#1A2820",color:G.wh,fontSize:12,fontWeight:600,padding:"5px 11px",borderRadius:8,whiteSpace:"nowrap",pointerEvents:"none",opacity:0,zIndex:999,border:`1px solid ${G.bdr}`,boxShadow:"0 4px 16px #000A",transition:"opacity .15s"}}
              onMouseEnter={e=>e.currentTarget.style.opacity=1} onMouseLeave={e=>e.currentTarget.style.opacity=0}>
              {n.label}
            </div>}
          </div>;
        })}
      </nav>

      {/* Bottom status */}
      <div style={{padding:"10px 10px",borderTop:`1px solid ${G.bdr}`,flexShrink:0}}>
        {navExpanded
        ?<><div style={{fontSize:11,color:G.bdr,lineHeight:1.9}}>
          <div style={{color:G.mut,fontWeight:700,marginBottom:2}}>Live Stats</div>
          <div>Clients: <span style={{color:G.green,fontWeight:700}}>{clients.length}</span></div>
          <div>Active: <span style={{color:G.g3,fontWeight:700}}>{clients.filter(c=>c.status==="Active").length}</span></div>
          <div>Overdue: <span style={{color:G.red,fontWeight:700}}>{works.filter(isOD).length}</span></div>
        </div>
        {ownerOn&&<div style={{marginTop:6,fontSize:11,color:G.amb,fontWeight:700,display:"flex",alignItems:"center",gap:4}}>👑 Owner<button onClick={()=>{setOwnerOn(false);if(tab==="fin")setTab("wdash");}} style={{marginLeft:"auto",background:"transparent",border:`1px solid ${G.red}55`,color:G.red,fontSize:10,borderRadius:5,padding:"1px 5px",cursor:"pointer"}}>Lock</button></div>}
        {devOn&&<div style={{marginTop:4,fontSize:11,color:G.cyn,fontWeight:700,display:"flex",alignItems:"center",gap:4}}>⚙️ Dev<button onClick={()=>{setDevOn(false);if(tab==="dev")setTab("wdash");}} style={{marginLeft:"auto",background:"transparent",border:`1px solid ${G.bdr}`,color:G.mut,fontSize:10,borderRadius:5,padding:"1px 5px",cursor:"pointer"}}>Lock</button></div>}
        </>
        :<div style={{display:"flex",flexDirection:"column",gap:5,alignItems:"center"}}>
          {ownerOn&&<span title="Owner unlocked" style={{fontSize:14,cursor:"pointer"}} onClick={()=>{setOwnerOn(false);if(tab==="fin")setTab("wdash");}}>👑</span>}
          {devOn&&<span title="Dev unlocked" style={{fontSize:14,cursor:"pointer"}} onClick={()=>{setDevOn(false);if(tab==="dev")setTab("wdash");}}>⚙️</span>}
          <div style={{fontSize:10,color:G.red,fontWeight:700,textAlign:"center"}}>{works.filter(isOD).length>0&&works.filter(isOD).length}</div>
        </div>}
      </div>
    </div>
    {/* Main */}
    <div style={{flex:1,display:"flex",flexDirection:"column",overflow:"hidden",minWidth:0}}>
      <div className="app-header no-print" style={{background:"#070E09",borderBottom:`1px solid ${G.bdr}`,height:50,display:"flex",alignItems:"center",justifyContent:"space-between",flexShrink:0,gap:8}}>
        <div style={{display:"flex",alignItems:"center",gap:10,minWidth:0}}>
          <button className="mobile-hamburger" onClick={()=>setSideOpen(true)} style={{display:"none",background:"transparent",border:`1px solid ${G.bdr}`,color:G.txt,borderRadius:8,width:32,height:32,alignItems:"center",justifyContent:"center",fontSize:16,cursor:"pointer",flexShrink:0}}>☰</button>
          <div style={{fontWeight:700,fontSize:15,color:G.txt,whiteSpace:"nowrap",overflow:"hidden",textOverflow:"ellipsis"}}>{titles[tab]||""}</div>
        </div>
        <div style={{display:"flex",gap:10,alignItems:"center",flexShrink:0}}>
          {tab==="add"&&<div style={{display:"flex",gap:3,background:G.card,border:`1px solid ${G.bdr}`,borderRadius:9,padding:3}}>
            {[{id:"client",l:"➕ Add Client"},{id:"work",l:"📋 Assign Work"}].map(s=><button key={s.id} onClick={()=>setSub(s.id)} style={{padding:"5px 14px",borderRadius:7,border:"none",cursor:"pointer",fontSize:12,fontWeight:700,background:sub===s.id?`linear-gradient(135deg,${G.g2},${G.green})`:"transparent",color:sub===s.id?"#fff":G.mut,transition:"all .2s"}}>{s.l}</button>)}
          </div>}
          <div className="app-header-date" style={{fontSize:11,color:G.bdr,whiteSpace:"nowrap"}}>{new Date().toLocaleDateString("en-IN",{dateStyle:"medium"})}</div>
          <div style={{width:28,height:28,borderRadius:7,background:`linear-gradient(135deg,${G.g2},${G.green})`,display:"flex",alignItems:"center",justifyContent:"center",fontWeight:800,fontSize:13,flexShrink:0}}>F</div>
          <button onClick={() => supabase.auth.signOut()} title="Logout" style={{background:"transparent", border:`1px solid ${G.bdr}`, color:G.red, borderRadius:8, padding:"5px 10px", cursor:"pointer", fontSize:11, fontWeight:700, display:"flex", alignItems:"center", gap: 3}}>🚪 Logout</button>
        </div>
      </div>
      <div className="app-content" style={{flex:1,overflow:"auto"}}>
        {tab==="wdash"&&<WorkDash works={works} clients={clients} dd={dd}/>}
        {tab==="fin"&&(ownerOn?<FinDash works={works} invoices={invoices} receipts={receipts} onLogout={()=>{setOwnerOn(false);setTab("wdash");}} dd={dd} clients={clients} setClients={setClients} toast={toast}/>
          :<div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"60vh",flexDirection:"column",gap:14}}>
            <Logo sz={64}/><div style={{fontWeight:700,fontSize:17,marginTop:8}}>Owner Access Required</div>
            <button onClick={()=>{setPendingProt("fin");setShowOA(true);}} style={{padding:"11px 26px",borderRadius:11,border:"none",cursor:"pointer",background:`linear-gradient(135deg,${G.g2},${G.green})`,color:"#fff",fontWeight:700,fontSize:14}}>🔐 Unlock</button>
          </div>)}
        {tab==="add"&&sub==="client"&&<AddClient clients={clients} setClients={setClients} dd={dd} toast={toast}/>}
        {tab==="add"&&sub==="work"&&<AssignWork clients={clients} works={works} setWorks={setWorks} dd={dd} toast={toast}/>}
        {tab==="clients"&&<ClientList clients={clients} setClients={setClients} dd={dd} toast={toast}/>}
        {tab==="itr"&&<ITRComputationTab clients={clients} setClients={setClients} computations={computations} setComputations={setComputations} dd={dd} toast={toast}/>}
        {tab==="tracker"&&<WorkTracker works={works} setWorks={setWorks} clients={clients} ownerOn={ownerOn} dd={dd} pws={pws} toast={toast} invoices={invoices} setInvoices={setInvoices} receipts={receipts}/>}
        {tab==="invoice"&&(ownerOn?<InvoiceModule invoices={invoices} setInvoices={setInvoices} receipts={receipts} setReceipts={setReceipts} clients={clients} works={works} dd={dd} toast={toast} firmSettings={firmSettings} setFirmSettings={setFirmSettings}/>
          :<div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"60vh",flexDirection:"column",gap:14}}>
            <Logo sz={64}/><div style={{fontWeight:700,fontSize:17,marginTop:8}}>Owner Access Required</div>
            <button onClick={()=>{setPendingProt("invoice");setShowOA(true);}} style={{padding:"11px 26px",borderRadius:11,border:"none",cursor:"pointer",background:`linear-gradient(135deg,${G.g2},${G.green})`,color:"#fff",fontWeight:700,fontSize:14}}>🔐 Unlock</button>
          </div>)}
        {tab==="receipts"&&<ReceiptsModule receipts={receipts} setReceipts={setReceipts} invoices={invoices} setInvoices={setInvoices} clients={clients} dd={dd} toast={toast}/>}
        {tab==="dev"&&(devOn?<DevTab dd={dd} setDd={setDd} pws={pws} setPws={setPws} darkMode={darkMode} setDarkMode={setDarkMode} profilePic={profilePic} setProfilePic={setProfilePic} firmSettings={firmSettings} setFirmSettings={setFirmSettings} clients={clients} setClients={setClients} works={works} setWorks={setWorks} invoices={invoices} setInvoices={setInvoices} receipts={receipts} setReceipts={setReceipts} toast={toast} googleUser={googleUser} linkGoogleDrive={linkGoogleDrive} disconnectGoogleDrive={disconnectGoogleDrive} uploadBackupToGoogleDrive={uploadBackupToGoogleDrive}/>
          :<div style={{display:"flex",alignItems:"center",justifyContent:"center",height:"60vh",flexDirection:"column",gap:14}}>
            <span style={{fontSize:48}}>⚙️</span><div style={{fontWeight:700,fontSize:17}}>Developer Access Required</div>
            <button onClick={()=>setShowDA(true)} style={{padding:"11px 26px",borderRadius:11,border:"none",cursor:"pointer",background:`linear-gradient(135deg,${G.g2},${G.green})`,color:"#fff",fontWeight:700,fontSize:14}}>🔐 Unlock Dev</button>
          </div>)}
      </div>
    </div>
    <Toasts list={toasts}/>
  </div>;
}
