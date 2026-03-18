import { useState, useMemo, useCallback } from "react";
import * as XLSX from "xlsx";

const TODAY = new Date();
const fmt = n => n.toLocaleString("ko-KR");
const fmtW = n => "₩" + fmt(n);
const parseD = s => new Date(s + "T00:00:00");
const diffD = (a, b) => Math.ceil((b - a) / 86400000);
const fmtDt = d => `${d.getMonth() + 1}/${d.getDate()}`;

const GSHEET_CSV = "https://docs.google.com/spreadsheets/d/e/2PACX-1vQQ9NLchQXAa7utcOOT7B_fu8IEsLDxZp7wTqEXOAiGxJS8OeDE_PYpT6G84RZiDSOquQK7ro98tMYu/pub?gid=915218543&single=true&output=csv";
const GSHEET_EDIT = "https://docs.google.com/spreadsheets/d/1-R7zk7L32Q35aKPhberMouz1rwPhMH3A6T4FWsb8zas/edit?gid=915218543";

const INIT_S = [
  {id:1,cat:"설계/감리",sub:"실시설계",task:"실시설계 보완 수정",start:"2026-03-09",end:"2026-04-30",status:"진행중",progress:40,memo:""},
  {id:2,cat:"설계/감리",sub:"감리",task:"배관배선 검수",start:"2026-03-09",end:"2026-03-15",status:"완료",progress:100,memo:""},
  {id:3,cat:"설계/감리",sub:"감리",task:"인테리어 1차 검수",start:"2026-03-16",end:"2026-03-22",status:"완료",progress:100,memo:""},
  {id:4,cat:"설계/감리",sub:"감리",task:"인테리어 2차 검수",start:"2026-04-06",end:"2026-04-12",status:"예정",progress:0,memo:""},
  {id:5,cat:"설계/감리",sub:"감리",task:"인테리어 최종 검수",start:"2026-05-18",end:"2026-05-24",status:"예정",progress:0,memo:""},
  {id:6,cat:"인테리어",sub:"도장공사",task:"1층 노출천정 도장",start:"2026-03-10",end:"2026-03-23",status:"지연",progress:70,memo:"자재 수급 지연"},
  {id:7,cat:"인테리어",sub:"도장공사",task:"2층 퍼티",start:"2026-03-10",end:"2026-03-15",status:"완료",progress:100,memo:""},
  {id:8,cat:"인테리어",sub:"경량공사",task:"2층 천정 골조 및 석고보드",start:"2026-03-17",end:"2026-03-28",status:"진행중",progress:60,memo:""},
  {id:9,cat:"인테리어",sub:"수장공사",task:"1층 디럭스 타일 시공",start:"2026-03-17",end:"2026-03-28",status:"진행중",progress:45,memo:""},
  {id:10,cat:"인테리어",sub:"타일공사",task:"타일공사",start:"2026-03-30",end:"2026-04-05",status:"예정",progress:0,memo:""},
  {id:11,cat:"인테리어",sub:"유리공사",task:"도어 유리 및 거울 설치",start:"2026-04-13",end:"2026-04-19",status:"예정",progress:0,memo:""},
  {id:12,cat:"인테리어",sub:"사인공사",task:"사인 설치",start:"2026-05-04",end:"2026-05-10",status:"예정",progress:0,memo:""},
  {id:13,cat:"인테리어",sub:"가구공사",task:"가구 설치",start:"2026-05-11",end:"2026-05-17",status:"예정",progress:0,memo:""},
  {id:14,cat:"인테리어",sub:"준공청소",task:"준공청소",start:"2026-05-25",end:"2026-05-31",status:"예정",progress:0,memo:""},
  {id:15,cat:"놀이기구",sub:"1층",task:"볼스윙 - 철물작업",start:"2026-03-17",end:"2026-03-19",status:"진행중",progress:80,memo:""},
  {id:16,cat:"놀이기구",sub:"1층",task:"볼스윙 - 목공작업",start:"2026-03-20",end:"2026-03-28",status:"예정",progress:0,memo:""},
  {id:17,cat:"놀이기구",sub:"1층",task:"고공 챌린지 - 파이프공사",start:"2026-03-19",end:"2026-03-30",status:"진행중",progress:30,memo:""},
  {id:18,cat:"놀이기구",sub:"안전인증",task:"제품안전인증검사",start:"2026-04-15",end:"2026-05-03",status:"예정",progress:0,memo:""},
  {id:19,cat:"하드웨어",sub:"발주",task:"계약·발주",start:"2026-03-09",end:"2026-03-22",status:"완료",progress:100,memo:""},
  {id:20,cat:"하드웨어",sub:"발주",task:"입고",start:"2026-03-23",end:"2026-03-31",status:"진행중",progress:50,memo:""},
  {id:21,cat:"하드웨어",sub:"설치",task:"프로젝터 설치",start:"2026-03-30",end:"2026-04-05",status:"예정",progress:0,memo:""},
  {id:22,cat:"하드웨어",sub:"설치",task:"스피커 설치",start:"2026-04-06",end:"2026-04-14",status:"예정",progress:0,memo:""},
  {id:23,cat:"하드웨어",sub:"설치",task:"테스트 및 최종세팅",start:"2026-05-11",end:"2026-05-17",status:"예정",progress:0,memo:""},
  {id:24,cat:"브랜딩",sub:"사이니지",task:"사이니지 기획/제작",start:"2026-03-09",end:"2026-03-29",status:"진행중",progress:65,memo:""},
  {id:25,cat:"브랜딩",sub:"홍보물",task:"홍보물 제작",start:"2026-03-30",end:"2026-04-14",status:"예정",progress:0,memo:""},
  {id:26,cat:"콘텐츠",sub:"스타점핑",task:"테스트 후 디벨롭",start:"2026-03-09",end:"2026-03-15",status:"완료",progress:100,memo:""},
  {id:27,cat:"콘텐츠",sub:"스타점핑",task:"1차 완료 테스트",start:"2026-03-23",end:"2026-03-29",status:"진행중",progress:40,memo:""},
  {id:28,cat:"콘텐츠",sub:"스카이라이딩",task:"자전거 RPM/핸들 센서 연동",start:"2026-03-23",end:"2026-03-29",status:"진행중",progress:50,memo:""},
  {id:29,cat:"콘텐츠",sub:"머치테크",task:"ESP32·스캐너 통합",start:"2026-04-27",end:"2026-05-03",status:"예정",progress:0,memo:""},
  {id:30,cat:"콘텐츠",sub:"키오스크",task:"결제용 GUI 기본구조",start:"2026-03-30",end:"2026-04-12",status:"예정",progress:0,memo:""},
];
const INIT_P = [
  {id:1,code:"DS250111",name:"별내 아르떼 키즈파크 H/W",status:"진행중",scope:"영상·음향 H/W 공급 및 설치",company:"디비인솔",mgr:"최상욱",terms:"선/중/잔",amt:848000000,paid:0,monthly:{},note:"변경 계약 필요"},
  {id:2,code:"DS25R001",name:"롯데월드 부산 로리캐슬",status:"진행중",scope:"음향 장비 공급",company:"디비인솔",mgr:"권가윤",terms:"선/중/잔",amt:180000000,paid:144000000,monthly:{},note:""},
  {id:3,code:"DS250111",name:"별내 아르떼 키즈파크 H/W",status:"진행예정",scope:"RFID 시스템",company:"머치테크",mgr:"최상욱",terms:"선/잔",amt:98400000,paid:0,monthly:{},note:""},
  {id:4,code:"DS250111",name:"별내 아르떼 키즈파크 H/W",status:"진행예정",scope:"LG전자 DID 공급",company:"에이텍정보통신",mgr:"최상욱",terms:"일시금",amt:14500000,paid:0,monthly:{},note:""},
  {id:5,code:"DS250111",name:"별내 아르떼 키즈파크 H/W",status:"진행예정",scope:"미디어서버 라이센스",company:"스타네트웍스",mgr:"최상욱",terms:"일시금",amt:9600000,paid:0,monthly:{},note:"판도라박스"},
  {id:6,code:"DS250111",name:"별내 아르떼 키즈파크 H/W",status:"진행예정",scope:"미디어서버 시스템",company:"랩에이해쉬",mgr:"최상욱",terms:"선/중/잔",amt:17325000,paid:0,monthly:{},note:"서버 1대로 변경"},
  {id:7,code:"DS25R018",name:"별내 아르떼키즈파크 설계 및 감리",status:"진행예정",scope:"인테리어 실시설계 보완",company:"바로아이앤디",mgr:"권가윤",terms:"일시금",amt:6000000,paid:0,monthly:{},note:""},
];

const C={bg:"#f5f5f5",card:"#fff",cardA:"#fafafa",bdr:"#e8e8e8",t1:"#111",t2:"#555",t3:"#999",t4:"#ccc",ok:"#22c55e",warn:"#f59e0b",err:"#ef4444",barBg:"#e8e8e8",barF:"#333"};
const stCl={"완료":{bg:"#f0fdf4",tx:"#166534",dt:C.ok},"진행중":{bg:"#f5f5f5",tx:C.t1,dt:C.t1},"지연":{bg:"#fef2f2",tx:"#991b1b",dt:C.err},"예정":{bg:"#fafafa",tx:C.t3,dt:"#d4d4d4"},"진행예정":{bg:"#fffbeb",tx:"#92400e",dt:C.warn}};
const catT={"설계/감리":"#111","인테리어":"#333","놀이기구":"#555","하드웨어":"#777","브랜딩":"#999","콘텐츠":"#444"};

function Badge({status}){const c=stCl[status]||stCl["예정"];return(<span style={{display:"inline-flex",alignItems:"center",gap:5,padding:"3px 12px",borderRadius:20,fontSize:11,fontWeight:600,background:c.bg,color:c.tx,border:`1px solid ${c.dt}30`,letterSpacing:".02em"}}><span style={{width:6,height:6,borderRadius:"50%",background:c.dt}}/>{status}</span>);}
function Bar({value,h=6,color}){const f=value===100?C.ok:(color||C.barF);return(<div style={{width:"100%",background:C.barBg,borderRadius:99,height:h,overflow:"hidden"}}><div style={{width:`${value}%`,height:"100%",background:f,borderRadius:99,transition:"width .5s cubic-bezier(.4,0,.2,1)"}}/></div>);}
function Tab({active,label,onClick,count}){return(<button onClick={onClick} style={{padding:"10px 20px",border:"none",borderBottom:active?"2px solid #111":"2px solid transparent",background:"none",color:active?"#111":"#999",fontWeight:active?700:400,cursor:"pointer",fontSize:13,display:"flex",alignItems:"center",gap:8,whiteSpace:"nowrap"}}>{label}{count!==undefined&&<span style={{background:active?"#111":"#e8e8e8",color:active?"#fff":"#999",borderRadius:20,padding:"2px 8px",fontSize:10,fontWeight:700}}>{count}</span>}</button>);}
function Chip({active,label,onClick}){return(<button onClick={onClick} style={{padding:"6px 16px",borderRadius:20,border:active?"none":`1px solid ${C.bdr}`,background:active?"#111":"#fff",color:active?"#fff":C.t2,fontSize:12,fontWeight:600,cursor:"pointer"}}>{label}</button>);}
function Card({children,style,accent}){return(<div style={{background:C.card,borderRadius:16,padding:24,border:`1px solid ${C.bdr}`,...(accent?{borderLeft:`3px solid ${accent}`}:{}),...style}}>{children}</div>);}
function StatCard({label,value,sub,accent}){return(<Card style={{textAlign:"center"}} accent={accent}><div style={{fontSize:11,color:C.t3,marginBottom:6,textTransform:"uppercase",letterSpacing:".08em",fontWeight:600}}>{label}</div><div style={{fontSize:32,fontWeight:800,color:C.t1,lineHeight:1.1}}>{value}</div>{sub&&<div style={{fontSize:11,color:C.t3,marginTop:6}}>{sub}</div>}</Card>);}

function getWR(d,off){const day=d.getDay();const m=new Date(d);m.setDate(m.getDate()-(day===0?6:day-1)+off*7);const s=new Date(m);s.setDate(s.getDate()+6);return[m,s];}
function inRange(t,s,e){const ts=parseD(t.start),te=parseD(t.end);return ts<=e&&te>=s;}

async function aiMap(rows,type,memo){
  const dr=rows.slice(0,200);const ds=dr.map((r,i)=>`[R${i}] ${r.map((c,j)=>`C${j}:${c}`).join("|")}`).join("\n");
  const sys=type==="schedule"?`Expert Korean project schedule analyst. Any format. Return ONLY JSON array: {"cat":"","sub":"","task":"","start":"YYYY-MM-DD","end":"YYYY-MM-DD","status":"완료|진행중|지연|예정","progress":0-100,"memo":""}. Today=${TODAY.toISOString().slice(0,10)}.${memo?` PM notes: ${memo}`:""}`:`Expert Korean payment analyst. Return ONLY JSON array: {"code":"","name":"","status":"","scope":"","company":"","mgr":"","terms":"","amt":0,"paid":0,"monthly":{},"note":""}`;
  try{const r=await fetch("https://api.anthropic.com/v1/messages",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({model:"claude-sonnet-4-20250514",max_tokens:8000,system:sys,messages:[{role:"user",content:`Data (${dr.length}r):\n${ds}`}]})});const d=await r.json();const t=d.content?.map(b=>b.type==="text"?b.text:"").join("")||"";const m=t.replace(/```json|```/g,"").trim().match(/\[[\s\S]*\]/);return m?JSON.parse(m[0]):JSON.parse(t);}catch(e){throw new Error("AI 분석 실패: "+e.message);}
}

function Gantt({tasks,catFilter}){
  const fl=catFilter==="전체"?tasks:tasks.filter(t=>t.cat===catFilter);
  if(!fl.length) return(<p style={{color:C.t3,fontSize:13}}>표시할 작업이 없습니다.</p>);
  const allD=fl.flatMap(t=>[parseD(t.start),parseD(t.end)]);
  let minD=new Date(Math.min(...allD)),maxD=new Date(Math.max(...allD));
  minD.setDate(minD.getDate()-2);maxD.setDate(maxD.getDate()+2);
  const totD=diffD(minD,maxD),rH=28,hH=44,pL=240,px=8;
  const months=[];let cur=new Date(minD);
  while(cur<=maxD){const ms=new Date(cur.getFullYear(),cur.getMonth(),1);const me=new Date(cur.getFullYear(),cur.getMonth()+1,0);const s=Math.max(0,diffD(minD,ms<minD?minD:ms));const e=Math.min(totD,diffD(minD,me>maxD?maxD:me));if(e>s)months.push({l:`${cur.getMonth()+1}월`,s,e});cur=new Date(cur.getFullYear(),cur.getMonth()+1,1);}
  const todayX=diffD(minD,TODAY);const grp={};fl.forEach(t=>{if(!grp[t.cat])grp[t.cat]=[];grp[t.cat].push(t);});
  let rows=[];Object.entries(grp).forEach(([cat,items])=>{rows.push({type:"h",cat});items.forEach(t=>rows.push({type:"t",...t}));});
  const svgH=hH+rows.length*rH+10,svgW=pL+totD*px+20;
  return(
    <div style={{overflowX:"auto",overflowY:"auto",maxHeight:500,borderRadius:12,border:`1px solid ${C.bdr}`,background:"#fff"}}>
      <svg width={svgW} height={svgH} style={{display:"block"}}>
        {months.map((m,i)=>(<g key={i}><rect x={pL+m.s*px} y={0} width={(m.e-m.s)*px} height={hH} fill={i%2===0?"#fafafa":"#f5f5f5"}/><text x={pL+(m.s+m.e)/2*px} y={26} textAnchor="middle" fontSize={12} fontWeight={600} fill={C.t2}>{m.l}</text></g>))}
        {Array.from({length:totD}).map((_,i)=>(<line key={i} x1={pL+i*px} y1={hH} x2={pL+i*px} y2={svgH} stroke={C.bdr} strokeWidth={.3}/>))}
        {todayX>=0&&todayX<=totD&&(<g><line x1={pL+todayX*px} y1={hH-6} x2={pL+todayX*px} y2={svgH} stroke={C.err} strokeWidth={1.5} strokeDasharray="3,3"/><text x={pL+todayX*px} y={hH-1} textAnchor="middle" fontSize={8} fontWeight={700} fill={C.err}>TODAY</text></g>)}
        {rows.map((r,i)=>{const y=hH+i*rH;if(r.type==="h") return(<g key={`h${i}`}><rect x={0} y={y} width={svgW} height={rH} fill="#f0f0f0"/><text x={12} y={y+18} fontSize={11} fontWeight={700} fill={catT[r.cat]||C.t1}>{r.cat}</text></g>);
        const ts=parseD(r.start),te=parseD(r.end),x1=pL+diffD(minD,ts)*px,w=Math.max(diffD(ts,te)*px,4);const tone=catT[r.cat]||C.t1;const isDel=r.status==="지연"||(r.status!=="완료"&&te<TODAY);
        return(<g key={`t${i}`}><rect x={0} y={y} width={svgW} height={rH} fill={i%2===0?"#fff":"#fcfcfc"}/><text x={12} y={y+18} fontSize={10} fill={C.t3}><tspan>{r.sub} · </tspan><tspan fontWeight={600} fill={C.t1}>{r.task?.length>24?r.task.slice(0,24)+"…":r.task}</tspan></text><rect x={x1} y={y+7} width={w} height={14} rx={4} fill={isDel?C.err+"15":tone+"15"} stroke={isDel?C.err:"none"} strokeWidth={isDel?.8:0}/><rect x={x1} y={y+7} width={w*r.progress/100} height={14} rx={4} fill={isDel?C.err:r.progress===100?C.ok:tone} opacity={.85}/>{w>32&&<text x={x1+w/2} y={y+18} textAnchor="middle" fontSize={8} fontWeight={700} fill="#fff">{r.progress}%</text>}</g>);})}
      </svg>
    </div>
  );
}

export default function App(){
  const [schedule,setSchedule]=useState(INIT_S);
  const [payments,setPayments]=useState(INIT_P);
  const [tab,setTab]=useState("overview");
  const [catF,setCatF]=useState("전체");
  const [uploadMode,setUploadMode]=useState(null);
  const [memoText,setMemoText]=useState("");
  const [showMemo,setShowMemo]=useState(false);
  const [projectCode,setProjectCode]=useState("DS250111");
  const [activeProject,setActiveProject]=useState(null);
  const [projectName,setProjectName]=useState("아르떼키즈파크 별내");
  const [aiLoading,setAiLoading]=useState(false);
  const [aiLog,setAiLog]=useState("");
  const [sheetNames,setSheetNames]=useState([]);
  const [selSheet,setSelSheet]=useState("");
  const [wbRef,setWbRef]=useState(null);
  const [manualMode,setManualMode]=useState(false);
  const [rawRows,setRawRows]=useState([]);
  const [rawType,setRawType]=useState("schedule");
  const [gLoading,setGLoading]=useState(false);
  const [gStatus,setGStatus]=useState("");

  const [lastW,thisW,nextW]=useMemo(()=>[getWR(TODAY,-1),getWR(TODAY,0),getWR(TODAY,1)],[]);
  const lastT=useMemo(()=>schedule.filter(t=>inRange(t,lastW[0],lastW[1])),[schedule,lastW]);
  const thisT=useMemo(()=>schedule.filter(t=>inRange(t,thisW[0],thisW[1])),[schedule,thisW]);
  const nextT=useMemo(()=>schedule.filter(t=>inRange(t,nextW[0],nextW[1])),[schedule,nextW]);
  const urgent=useMemo(()=>schedule.filter(t=>{if(t.status==="완료")return false;const e=parseD(t.end);const d=diffD(TODAY,e);return d<=7&&d>=0;}).sort((a,b)=>parseD(a.end)-parseD(b.end)),[schedule]);
  const delayed=useMemo(()=>schedule.filter(t=>t.status==="지연"||(t.status!=="완료"&&parseD(t.end)<TODAY)),[schedule]);
  const overallProg=useMemo(()=>{if(!schedule.length)return 0;return Math.round(schedule.reduce((s,t)=>s+t.progress,0)/schedule.length);},[schedule]);
  const catProg=useMemo(()=>{const m={};schedule.forEach(t=>{if(!m[t.cat])m[t.cat]={s:0,c:0};m[t.cat].s+=t.progress;m[t.cat].c++;});return Object.entries(m).map(([k,v])=>({cat:k,prog:Math.round(v.s/v.c),cnt:v.c}));},[schedule]);
  const projectPayments=useMemo(()=>payments.filter(p=>p.code===projectCode),[payments,projectCode]);
  const projectCodes=useMemo(()=>[...new Set(payments.map(p=>p.code).filter(Boolean))],[payments]);
  const totalAmt=useMemo(()=>projectPayments.reduce((s,p)=>s+p.amt,0),[projectPayments]);
  const totalPaid=useMemo(()=>projectPayments.reduce((s,p)=>s+p.paid,0),[projectPayments]);
  const allTotalAmt=useMemo(()=>payments.reduce((s,p)=>s+p.amt,0),[payments]);
  const allTotalPaid=useMemo(()=>payments.reduce((s,p)=>s+p.paid,0),[payments]);
  const categories=useMemo(()=>["전체",...new Set(schedule.map(t=>t.cat))],[schedule]);

  const fetchGSheet=useCallback(async()=>{
    setGLoading(true);setGStatus("연동 중...");
    try{
      const resp=await fetch(GSHEET_CSV);
      if(!resp.ok) throw new Error("시트를 불러올 수 없습니다.");
      const csv=await resp.text();
      const rows=csv.split("\n").map(line=>{const res=[];let cur="",inQ=false;for(let i=0;i<line.length;i++){const ch=line[i];if(ch==='"')inQ=!inQ;else if(ch===','&&!inQ){res.push(cur.trim());cur="";}else cur+=ch;}res.push(cur.trim());return res;}).filter(r=>r.some(c=>c));
      if(rows.length<2){setGStatus("데이터 부족");setGLoading(false);return;}
      setGStatus(`${rows.length}행 수신 — AI 분석 중...`);
      const result=await aiMap(rows,"payment","");
      if(!Array.isArray(result)||!result.length){setGStatus("AI 추출 실패");setGLoading(false);return;}
      const data=result.map((r,i)=>({id:i+1,code:r.code||"",name:r.name||"",status:r.status||"",scope:r.scope||"",company:r.company||"",mgr:r.mgr||"",terms:r.terms||"",amt:Number(r.amt)||0,paid:Number(r.paid)||0,monthly:r.monthly||{},note:r.note||""}));
      setPayments(data);if(data[0]?.code)setProjectCode(data[0].code);
      setGStatus(`✓ ${data.length}건 연동 완료`);
    }catch(err){setGStatus("오류: "+err.message);}
    setGLoading(false);
  },[]);

  const autoMatch=useCallback((fn,sh)=>{
    const src=[fn||"",sh||""].join(" ").replace(/\.(xlsx?|csv)$/i,"").replace(/[_\-\.]/g," ");
    const sw=src.split(/\s+/).filter(w=>w.length>=2);if(!sw.length||!payments.length)return null;
    const g={};payments.forEach(p=>{if(!p.code)return;if(!g[p.code])g[p.code]={code:p.code,names:new Set()};if(p.name)g[p.code].names.add(p.name);});
    let best=null,bs=0;Object.values(g).forEach(gr=>{const nw=[...gr.names].join(" ").split(/\s+/).filter(w=>w.length>=2);let sc=0;sw.forEach(s=>{nw.forEach(n=>{if(s.includes(n)||n.includes(s))sc+=Math.min(s.length,n.length);});});if(sc>bs){bs=sc;best={code:gr.code,name:[...gr.names][0]||gr.code};}});
    return bs>=4?best:null;
  },[payments]);

  const processSheet=useCallback(async(wb,sn,type)=>{
    const ws=wb.Sheets[sn];const rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
    if(rows.length<2){alert("데이터 부족");return;}
    setRawRows(rows);setRawType(type);setManualMode(false);setAiLoading(true);setAiLog(`분석 중 — "${sn}" (${rows.length}행)\n`);
    try{
      const res=await aiMap(rows,type,type==="schedule"?memoText:"");
      if(!Array.isArray(res)||!res.length){setAiLog(p=>p+"추출 실패\n");setAiLoading(false);return;}
      setAiLog(p=>p+`${res.length}건 완료\n`);
      if(type==="schedule"){setSchedule(res.map((r,i)=>({id:i+1,cat:r.cat||"기타",sub:r.sub||"",task:r.task||"",start:r.start||"",end:r.end||"",status:r.status||"예정",progress:typeof r.progress==="number"?r.progress:0,memo:r.memo||""})));}
      else{const d=res.map((r,i)=>({id:i+1,code:r.code||"",name:r.name||"",status:r.status||"",scope:r.scope||"",company:r.company||"",mgr:r.mgr||"",terms:r.terms||"",amt:Number(r.amt)||0,paid:Number(r.paid)||0,monthly:r.monthly||{},note:r.note||""}));setPayments(d);if(d[0]?.code)setProjectCode(d[0].code);}
      setUploadMode(null);setSheetNames([]);setMemoText("");setShowMemo(false);
    }catch(e){setAiLog(p=>p+`오류: ${e.message}\n`);}
    setAiLoading(false);
  },[memoText]);

  const handleFile=useCallback((e,type)=>{
    const file=e.target.files[0];if(!file)return;
    const fn=file.name,isX=file.name.match(/\.xlsx?$/i);
    const reader=new FileReader();
    reader.onload=(evt)=>{try{let wb,names;if(isX){wb=XLSX.read(new Uint8Array(evt.target.result),{type:"array"});names=wb.SheetNames;}else{const txt=typeof evt.target.result==="string"?evt.target.result:new TextDecoder().decode(evt.target.result);const rows=txt.split("\n").map(l=>l.split(",").map(c=>c.trim().replace(/^"|"$/g,"")));wb={Sheets:{"Sheet1":XLSX.utils.aoa_to_sheet(rows)},SheetNames:["Sheet1"]};names=["Sheet1"];}setWbRef({wb,type,fn});if(type==="schedule"){const m=autoMatch(fn,names[0]);if(m){setActiveProject(m);setProjectCode(m.code);}setProjectName(fn.replace(/\.(xlsx?|csv)$/i,"").replace(/[_\-]/g," "));}if(names.length>1){setSheetNames(names);setSelSheet(names[0]);}else processSheet(wb,names[0],type);}catch(err){alert("파일 오류: "+err.message);}};
    if(isX)reader.readAsArrayBuffer(file);else reader.readAsText(file);
  },[autoMatch,processSheet]);

  const handleSheetConfirm=useCallback(()=>{if(wbRef){if(wbRef.type==="schedule"){const m=autoMatch(wbRef.fn,selSheet);if(m){setActiveProject(m);setProjectCode(m.code);}}processSheet(wbRef.wb,selSheet,wbRef.type);}},[wbRef,selSheet,processSheet,autoMatch]);

  const wkLabel=(s,e)=>`${fmtDt(s)} ~ ${fmtDt(e)}`;
  const bp={padding:"10px 20px",background:"#111",color:"#fff",border:"none",borderRadius:10,fontSize:13,fontWeight:600,cursor:"pointer"};
  const bs={padding:"10px 20px",background:C.cardA,border:`1px solid ${C.bdr}`,borderRadius:10,fontSize:13,cursor:"pointer",color:C.t2};

  function WeekSec({title,tasks,wr,icon}){
    return(
      <Card style={{marginBottom:16}}>
        <div style={{display:"flex",justifyContent:"space-between",alignItems:"center",marginBottom:16}}>
          <div style={{display:"flex",alignItems:"center",gap:10}}><span style={{fontSize:16}}>{icon}</span><span style={{fontSize:16,fontWeight:700}}>{title}</span><span style={{fontSize:12,color:C.t3}}>{wkLabel(wr[0],wr[1])}</span></div>
          <span style={{background:"#f0f0f0",padding:"4px 12px",borderRadius:20,fontSize:12,fontWeight:700}}>{tasks.length}</span>
        </div>
        {!tasks.length?(<p style={{color:C.t3,fontSize:13,margin:0}}>작업 없음</p>):
        <div style={{display:"flex",flexDirection:"column",gap:8}}>
          {tasks.map(t=>(<div key={t.id} style={{display:"flex",alignItems:"center",gap:12,padding:"10px 16px",background:C.cardA,borderRadius:12,borderLeft:`3px solid ${catT[t.cat]||"#888"}`}}><div style={{flex:1,minWidth:0}}><div style={{fontSize:13,fontWeight:600,color:C.t1}}>{t.task}</div><div style={{fontSize:11,color:C.t3,marginTop:2}}>{t.cat} · {t.sub}{t.memo&&<span style={{color:C.warn}}> — {t.memo}</span>}</div></div><div style={{width:80}}><Bar value={t.progress}/></div><span style={{fontSize:12,fontWeight:700,color:C.t2,width:36,textAlign:"right"}}>{t.progress}%</span><Badge status={t.status}/></div>))}
        </div>}
      </Card>
    );
  }

  return(
    <div style={{fontFamily:"'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif",background:C.bg,minHeight:"100vh",padding:"24px 20px"}}>
      {/* Header */}
      <div style={{display:"flex",justifyContent:"space-between",alignItems:"flex-start",marginBottom:28}}>
        <div>
          <p style={{margin:0,fontSize:11,color:C.t3,fontWeight:600,letterSpacing:".06em"}}>DX본부 PM팀</p>
          <h1 style={{margin:"2px 0 0",fontSize:22,fontWeight:800,color:C.t1,letterSpacing:"-.02em"}}>프로젝트 대시보드</h1>
          <p style={{margin:"4px 0 0",fontSize:14,fontWeight:700,color:C.t1}}>{projectName}</p>
          <p style={{margin:"2px 0 0",fontSize:12,color:C.t3}}>{TODAY.toLocaleDateString("ko-KR",{year:"numeric",month:"long",day:"numeric",weekday:"long"})}</p>
        </div>
        <div style={{display:"flex",flexDirection:"column",alignItems:"flex-end",gap:10}}>
          <img src="https://seoul.designfestival.co.kr/wp-content/uploads/2022/12/회사로고_dstrict_CI_BLACK.png" alt="d'strict" style={{height:22}} />
          <button onClick={()=>{setUploadMode("schedule");setShowMemo(false);setAiLog("");setSheetNames([]);}} style={{...bp,fontSize:12,padding:"8px 16px",display:"flex",alignItems:"center",gap:6}}>+ 공정표 업로드</button>
        </div>
      </div>

      {/* Upload Panel */}
      {uploadMode==="schedule"&&(
        <Card style={{marginBottom:20,background:"#fafafa"}}>
          <div style={{display:"flex",alignItems:"center",gap:10,marginBottom:12}}><h3 style={{margin:0,fontSize:15,fontWeight:700}}>공정표 업로드</h3><span style={{background:"#111",color:"#fff",padding:"3px 10px",borderRadius:20,fontSize:10,fontWeight:700}}>AI 자동 매핑</span></div>
          <p style={{fontSize:12,color:C.t3,margin:"0 0 12px"}}>CSV, Excel 모두 지원. 어떤 양식이든 AI가 자동 분석합니다.</p>
          <div style={{display:"flex",gap:10,alignItems:"center",flexWrap:"wrap"}}>
            <input type="file" accept=".csv,.xlsx,.xls" onChange={e=>handleFile(e,"schedule")} style={{fontSize:13}} disabled={aiLoading}/>
            <button onClick={()=>{setUploadMode(null);setAiLog("");}} style={bs}>취소</button>
          </div>
          <div style={{marginTop:12}}>
            <button onClick={()=>setShowMemo(!showMemo)} style={{...bs,fontSize:11,padding:"6px 14px"}}>{showMemo?"메모 접기":"＋ 현재 상황 메모 추가"}</button>
            {showMemo&&(<textarea value={memoText} onChange={e=>setMemoText(e.target.value)} placeholder="예: 도장공사 지연 70%" style={{display:"block",width:"100%",minHeight:60,padding:10,borderRadius:10,border:`1px solid ${C.bdr}`,fontSize:12,fontFamily:"inherit",marginTop:8,resize:"vertical",boxSizing:"border-box",background:C.cardA}}/>)}
          </div>
          {sheetNames.length>1&&(<div style={{marginTop:12}}><p style={{fontSize:12,fontWeight:600,margin:"0 0 8px"}}>시트 선택:</p><div style={{display:"flex",gap:8,flexWrap:"wrap",marginBottom:10}}>{sheetNames.map(n=>(<Chip key={n} active={selSheet===n} label={n} onClick={()=>setSelSheet(n)}/>))}</div><button onClick={handleSheetConfirm} disabled={aiLoading} style={bp}>{aiLoading?"분석 중...":"AI 분석 시작"}</button></div>)}
          {aiLog&&(<div style={{marginTop:12,padding:14,background:"#111",borderRadius:10,maxHeight:160,overflowY:"auto"}}><pre style={{margin:0,fontSize:11,color:"#aaa",fontFamily:"monospace",whiteSpace:"pre-wrap"}}>{aiLog}</pre>{aiLoading&&<span style={{display:"inline-block",width:10,height:10,border:"2px solid #666",borderTopColor:"#aaa",borderRadius:"50%",animation:"sp .8s linear infinite",marginTop:6}}/>}<style>{`@keyframes sp{to{transform:rotate(360deg)}}`}</style></div>)}
          {!aiLoading&&rawRows.length>0&&!manualMode&&(<button onClick={()=>setManualMode(true)} style={{...bs,marginTop:10,fontSize:11}}>수동 매핑으로 전환</button>)}
        </Card>
      )}

      {/* Tabs */}
      <div style={{display:"flex",gap:0,borderBottom:`1px solid ${C.bdr}`,marginBottom:24,overflowX:"auto"}}>
        <Tab active={tab==="overview"} label="전체 현황" onClick={()=>setTab("overview")}/>
        <Tab active={tab==="weekly"} label="주간 작업" onClick={()=>setTab("weekly")}/>
        <Tab active={tab==="gantt"} label="간트차트" onClick={()=>setTab("gantt")}/>
        <Tab active={tab==="schedule"} label="전체 공정" onClick={()=>setTab("schedule")} count={schedule.length}/>
        <Tab active={tab==="payment"} label="대금집행" onClick={()=>setTab("payment")} count={payments.length}/>
      </div>

      {/* OVERVIEW */}
      {tab==="overview"&&(
        <div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(auto-fit,minmax(180px,1fr))",gap:16,marginBottom:24}}>
            <StatCard label="전체 공정률" value={`${overallProg}%`} sub={<Bar value={overallProg} h={8}/>}/>
            <StatCard label="금주 작업" value={thisT.length} sub={`${thisT.filter(t=>t.status==="완료").length}건 완료 · ${thisT.filter(t=>t.status==="진행중").length}건 진행중`}/>
            <StatCard label="마감 임박" value={urgent.length} sub="7일 이내" accent={urgent.length>0?C.warn:undefined}/>
            <StatCard label="지연 작업" value={delayed.length} sub="확인 필요" accent={delayed.length>0?C.err:undefined}/>
          </div>
          <Card style={{marginBottom:20}}>
            <h3 style={{margin:"0 0 20px",fontSize:15,fontWeight:700}}>분야별 공정률</h3>
            {catProg.map(c=>(<div key={c.cat} style={{display:"flex",alignItems:"center",gap:14,marginBottom:14}}><span style={{width:80,fontSize:13,fontWeight:600,color:catT[c.cat]||"#333"}}>{c.cat}</span><div style={{flex:1}}><Bar value={c.prog} h={10} color={catT[c.cat]}/></div><span style={{fontSize:14,fontWeight:800,width:44,textAlign:"right"}}>{c.prog}%</span><span style={{fontSize:11,color:C.t3,width:30}}>{c.cnt}건</span></div>))}
          </Card>
          {(urgent.length>0||delayed.length>0)&&(
            <div style={{display:"grid",gridTemplateColumns:delayed.length>0&&urgent.length>0?"1fr 1fr":"1fr",gap:16,marginBottom:20}}>
              {urgent.length>0&&(<Card accent={C.warn}><h3 style={{margin:"0 0 14px",fontSize:14,fontWeight:700,color:C.warn}}>마감 임박</h3>{urgent.map(t=>(<div key={t.id} style={{padding:"8px 0",borderBottom:`1px solid ${C.bdr}`}}><div style={{fontSize:13,fontWeight:600}}>{t.task}</div><div style={{fontSize:11,color:C.t3}}>D-{diffD(TODAY,parseD(t.end))} · {t.end}</div></div>))}</Card>)}
              {delayed.length>0&&(<Card accent={C.err}><h3 style={{margin:"0 0 14px",fontSize:14,fontWeight:700,color:C.err}}>지연 작업</h3>{delayed.map(t=>(<div key={t.id} style={{padding:"8px 0",borderBottom:`1px solid ${C.bdr}`}}><div style={{fontSize:13,fontWeight:600}}>{t.task}</div><div style={{fontSize:11,color:C.err}}>{t.cat} · {t.progress}%{t.memo&&` — ${t.memo}`}</div></div>))}</Card>)}
            </div>
          )}
          <Card>
            <h3 style={{margin:"0 0 16px",fontSize:15,fontWeight:700}}>대금집행 요약</h3>
            {activeProject&&(<div style={{display:"inline-flex",alignItems:"center",gap:6,marginBottom:12,padding:"6px 14px",background:"#f0f0f0",borderRadius:20,fontSize:12}}><span style={{width:6,height:6,borderRadius:"50%",background:C.ok}}/>{activeProject.code} — {activeProject.name}</div>)}
            <div style={{display:"flex",gap:8,marginBottom:16,flexWrap:"wrap"}}>{projectCodes.map(c=>(<Chip key={c} active={projectCode===c} label={c} onClick={()=>setProjectCode(c)}/>))}</div>
            <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16}}>
              <div style={{padding:16,background:C.cardA,borderRadius:12,textAlign:"center"}}><div style={{fontSize:10,color:C.t3,textTransform:"uppercase",letterSpacing:".08em",fontWeight:600}}>계약금액</div><div style={{fontSize:20,fontWeight:800,marginTop:4}}>{fmtW(totalAmt)}</div></div>
              <div style={{padding:16,background:"#f0fdf4",borderRadius:12,textAlign:"center"}}><div style={{fontSize:10,color:"#166534",textTransform:"uppercase",letterSpacing:".08em",fontWeight:600}}>지급완료</div><div style={{fontSize:20,fontWeight:800,color:"#166534",marginTop:4}}>{fmtW(totalPaid)}</div></div>
              <div style={{padding:16,background:"#fef2f2",borderRadius:12,textAlign:"center"}}><div style={{fontSize:10,color:"#991b1b",textTransform:"uppercase",letterSpacing:".08em",fontWeight:600}}>잔여금액</div><div style={{fontSize:20,fontWeight:800,color:"#991b1b",marginTop:4}}>{fmtW(totalAmt-totalPaid)}</div></div>
            </div>
            <div style={{marginTop:14}}><Bar value={totalAmt>0?Math.round(totalPaid/totalAmt*100):0} h={10} color={C.ok}/></div>
          </Card>
        </div>
      )}

      {tab==="weekly"&&(<div><WeekSec title="지난주" tasks={lastT} wr={lastW} icon="←"/><WeekSec title="금주" tasks={thisT} wr={thisW} icon="●"/><WeekSec title="차주" tasks={nextT} wr={nextW} icon="→"/></div>)}

      {tab==="gantt"&&(
        <div>
          <div style={{display:"flex",gap:8,marginBottom:20,flexWrap:"wrap"}}>{categories.map(c=>(<Chip key={c} active={catF===c} label={c} onClick={()=>setCatF(c)}/>))}</div>
          <Gantt tasks={schedule} catFilter={catF}/>
        </div>
      )}

      {tab==="schedule"&&(
        <div>
          <div style={{display:"flex",gap:8,marginBottom:20,flexWrap:"wrap"}}>{categories.map(c=>(<Chip key={c} active={catF===c} label={c} onClick={()=>setCatF(c)}/>))}</div>
          <Card style={{padding:0,overflow:"hidden"}}>
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr style={{borderBottom:`2px solid ${C.bdr}`}}>{["분류","중분류","작업명","시작","마감","상태","진행률","메모"].map(h=>(<th key={h} style={{padding:"14px 12px",textAlign:"left",fontSize:11,color:C.t3,fontWeight:600,textTransform:"uppercase",letterSpacing:".06em"}}>{h}</th>))}</tr></thead>
                <tbody>
                  {(catF==="전체"?schedule:schedule.filter(t=>t.cat===catF)).map(t=>(
                    <tr key={t.id} style={{borderBottom:`1px solid ${C.bdr}`}}>
                      <td style={{padding:12,fontWeight:600,color:catT[t.cat],fontSize:12}}>{t.cat}</td>
                      <td style={{padding:12,color:C.t3,fontSize:12}}>{t.sub}</td>
                      <td style={{padding:12,fontWeight:600}}>{t.task}</td>
                      <td style={{padding:12,fontSize:12,color:C.t3}}>{t.start}</td>
                      <td style={{padding:12,fontSize:12,color:C.t3}}>{t.end}</td>
                      <td style={{padding:12}}><Badge status={t.status}/></td>
                      <td style={{padding:12,width:110}}><div style={{display:"flex",alignItems:"center",gap:6}}><Bar value={t.progress}/><span style={{fontSize:11,fontWeight:700}}>{t.progress}%</span></div></td>
                      <td style={{padding:12,fontSize:11,color:C.warn,maxWidth:140,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{t.memo||"—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
        </div>
      )}

      {tab==="payment"&&(
        <div>
          <div style={{display:"flex",gap:8,marginBottom:20,flexWrap:"wrap",alignItems:"center"}}>
            {projectCodes.map(c=>(<Chip key={c} active={projectCode===c} label={c} onClick={()=>setProjectCode(c)}/>))}
            <Chip active={projectCode==="ALL"} label="전체" onClick={()=>setProjectCode("ALL")}/>
            {projectCode!=="ALL"&&<span style={{fontSize:12,color:C.t3,marginLeft:8}}>{projectPayments[0]?.name} — {projectPayments.length}건</span>}
          </div>
          <div style={{display:"grid",gridTemplateColumns:"repeat(3,1fr)",gap:16,marginBottom:20}}>
            <StatCard label="계약금액" value={fmtW(projectCode==="ALL"?allTotalAmt:totalAmt)}/>
            <StatCard label="지급완료" value={fmtW(projectCode==="ALL"?allTotalPaid:totalPaid)} accent={C.ok}/>
            <StatCard label="잔여금액" value={fmtW(projectCode==="ALL"?(allTotalAmt-allTotalPaid):(totalAmt-totalPaid))} accent={C.err}/>
          </div>
          <Card style={{padding:0,overflow:"hidden",marginBottom:20}}>
            <div style={{padding:"16px 20px",borderBottom:`1px solid ${C.bdr}`,display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <h3 style={{margin:0,fontSize:15,fontWeight:700}}>대금집행 상세</h3>
              <div style={{display:"flex",gap:8}}>
                <button onClick={fetchGSheet} disabled={gLoading} style={{...bp,fontSize:11,padding:"6px 14px",opacity:gLoading?.6:1}}>{gLoading?"연동 중...":"↻ 구글시트 동기화"}</button>
                <label style={{...bs,fontSize:11,padding:"6px 14px",cursor:"pointer",borderRadius:8}}>CSV 업로드<input type="file" accept=".csv,.xlsx,.xls" onChange={e=>handleFile(e,"payment")} style={{display:"none"}}/></label>
              </div>
            </div>
            {gStatus&&<div style={{padding:"8px 20px",fontSize:12,color:gStatus.startsWith("✓")?C.ok:gStatus.startsWith("오류")?C.err:C.t2,fontWeight:600,background:C.cardA}}>{gStatus}</div>}
            <div style={{overflowX:"auto"}}>
              <table style={{width:"100%",borderCollapse:"collapse",fontSize:13}}>
                <thead><tr style={{borderBottom:`2px solid ${C.bdr}`}}>{["상태","업체","스콥","계약금액","지급완료","잔여","조건","담당","비고"].map(h=>(<th key={h} style={{padding:"14px 12px",textAlign:"left",fontSize:11,color:C.t3,fontWeight:600,textTransform:"uppercase",letterSpacing:".06em"}}>{h}</th>))}</tr></thead>
                <tbody>
                  {(projectCode==="ALL"?payments:projectPayments).map(p=>(
                    <tr key={p.id} style={{borderBottom:`1px solid ${C.bdr}`}}>
                      <td style={{padding:12}}><Badge status={p.status}/></td>
                      <td style={{padding:12,fontWeight:600}}>{p.company}</td>
                      <td style={{padding:12,maxWidth:180,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap",color:C.t2}}>{p.scope}</td>
                      <td style={{padding:12,fontWeight:700,whiteSpace:"nowrap"}}>{fmtW(p.amt)}</td>
                      <td style={{padding:12,color:"#166534",fontWeight:600,whiteSpace:"nowrap"}}>{fmtW(p.paid)}</td>
                      <td style={{padding:12,color:"#991b1b",fontWeight:600,whiteSpace:"nowrap"}}>{fmtW(p.amt-p.paid)}</td>
                      <td style={{padding:12,color:C.t3}}>{p.terms}</td>
                      <td style={{padding:12,color:C.t3}}>{p.mgr}</td>
                      <td style={{padding:12,color:C.t3,maxWidth:100,overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{p.note||"—"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </Card>
          <Card>
            <div style={{display:"flex",justifyContent:"space-between",alignItems:"center"}}>
              <div>
                <h3 style={{margin:0,fontSize:15,fontWeight:700}}>구글 스프레드시트 원본</h3>
                <p style={{margin:"4px 0 0",fontSize:11,color:C.t3}}>시트에서 수정 후 "구글시트 동기화" 버튼으로 대시보드에 반영</p>
              </div>
              <a href={GSHEET_EDIT} target="_blank" rel="noopener noreferrer" style={{...bs,fontSize:12,padding:"10px 16px",textDecoration:"none",display:"inline-block",flexShrink:0}}>시트 열기 ↗</a>
            </div>
          </Card>
        </div>
      )}

      {/* Footer */}
      <div style={{marginTop:40,paddingTop:20,borderTop:`1px solid ${C.bdr}`,textAlign:"center"}}>
        <p style={{margin:0,fontSize:10,color:C.t4,letterSpacing:".04em",opacity:.5}}>© d'strict DX DIV. PM TEAM — Sanguk Choi</p>
      </div>
    </div>
  );
}
