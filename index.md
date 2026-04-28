[沐蜜財務系統.html](https://github.com/user-attachments/files/27158903/default.html)
<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0">
<title>🌸 沐蜜財務系統</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Noto+Serif+TC:wght@400;600;700;900&family=Noto+Sans+TC:wght@300;400;500;700&display=swap" rel="stylesheet">
<style>
:root{
  --rose:#C2747A;--roseDk:#9E5A60;--roseL:#F5E8EA;--roseMid:#E8C4C8;
  --gold:#C4A46B;--goldL:#F9F0E0;--goldDk:#8C6B3A;
  --grn:#5A8C6B;--grnL:#E8F5ED;--red:#C25A5A;--redL:#FAEAEA;
  --blu:#5A7AC2;--bluL:#E8EEFA;--pur:#8C5AC2;--purL:#F0E8FA;
  --brown:#6B4C3A;--txt:#2D1F1A;--mut:#9E8880;--bdr:#EAE0D8;
  --bg:#FDFAF8;--bg2:#F7F0EC;--bg3:#F0E8E2;--card:#FFFFFF;
  --shadow:0 2px 16px rgba(180,100,100,0.08);
  --shadow2:0 4px 32px rgba(180,100,100,0.14);
}
*{box-sizing:border-box;margin:0;padding:0;}
body{font-family:'Noto Sans TC',sans-serif;background:var(--bg);color:var(--txt);font-size:14px;min-height:100vh;}
h1,h2,h3{font-family:'Noto Serif TC',serif;}

/* ── LAYOUT ── */
.app{display:flex;min-height:100vh;}
.sidebar{width:200px;background:linear-gradient(180deg,var(--roseDk) 0%,var(--brown) 100%);padding:0;flex-shrink:0;position:fixed;top:0;left:0;height:100vh;overflow-y:auto;z-index:100;}
.sb-logo{padding:20px 16px 12px;border-bottom:1px solid rgba(255,255,255,0.12);}
.sb-logo .name{font-family:'Noto Serif TC',serif;font-size:18px;font-weight:900;color:#fff;letter-spacing:2px;}
.sb-logo .sub{font-size:10px;color:rgba(255,255,255,0.55);margin-top:2px;letter-spacing:1px;}
.sb-item{display:flex;align-items:center;gap:10px;padding:11px 16px;cursor:pointer;transition:.2s;color:rgba(255,255,255,0.7);font-size:13px;border-left:3px solid transparent;}
.sb-item:hover{background:rgba(255,255,255,0.1);color:#fff;}
.sb-item.active{background:rgba(255,255,255,0.15);color:#fff;border-left-color:var(--gold);font-weight:700;}
.sb-item .ico{font-size:16px;width:22px;text-align:center;}
.sb-section{font-size:9px;font-weight:700;color:rgba(255,255,255,0.35);padding:14px 16px 4px;letter-spacing:2px;text-transform:uppercase;}

.main{margin-left:200px;flex:1;display:flex;flex-direction:column;}
.topbar{background:var(--card);border-bottom:1px solid var(--bdr);padding:0 24px;height:52px;display:flex;align-items:center;justify-content:space-between;position:sticky;top:0;z-index:50;box-shadow:var(--shadow);}
.tb-ttl{font-family:'Noto Serif TC',serif;font-size:15px;font-weight:700;color:var(--roseDk);}
.tb-right{display:flex;align-items:center;gap:10px;}
.sync-dot{width:8px;height:8px;border-radius:50%;background:var(--grn);}
.sync-dot.syncing{background:var(--gold);animation:pulse 1s infinite;}
.sync-dot.error{background:var(--red);}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
.btn-out{background:none;border:1.5px solid var(--bdr);border-radius:8px;padding:5px 12px;font-size:12px;cursor:pointer;color:var(--mut);transition:.2s;}
.btn-out:hover{border-color:var(--rose);color:var(--roseDk);}

.content{padding:24px;flex:1;}
.pg-ttl{font-family:'Noto Serif TC',serif;font-size:22px;font-weight:900;color:var(--roseDk);margin-bottom:4px;}
.pg-sub{font-size:12px;color:var(--mut);margin-bottom:20px;}

/* ── CARDS ── */
.card{background:var(--card);border-radius:14px;box-shadow:var(--shadow);margin-bottom:16px;overflow:hidden;}
.card-hd{display:flex;align-items:center;justify-content:space-between;padding:14px 18px;border-bottom:1px solid var(--bdr);}
.card-ttl{font-family:'Noto Serif TC',serif;font-size:14px;font-weight:700;color:var(--brown);}
.card-bd{padding:16px 18px;}

/* ── STAT GRID ── */
.sg{display:grid;gap:12px;margin-bottom:16px;}
.sg2{grid-template-columns:repeat(2,1fr);}
.sg3{grid-template-columns:repeat(3,1fr);}
.sg4{grid-template-columns:repeat(4,1fr);}
.sc{background:var(--card);border-radius:12px;padding:14px 16px;box-shadow:var(--shadow);border-top:3px solid var(--bdr);}
.sc.rose{border-top-color:var(--rose);}
.sc.grn{border-top-color:var(--grn);}
.sc.red{border-top-color:var(--red);}
.sc.gold{border-top-color:var(--gold);}
.sc.blu{border-top-color:var(--blu);}
.sc.pur{border-top-color:var(--pur);}
.sl{font-size:11px;color:var(--mut);margin-bottom:6px;font-weight:500;}
.sv{font-family:'Noto Serif TC',serif;font-size:22px;font-weight:700;color:var(--txt);}
.sv.sm{font-size:18px;}
.sv.rose{color:var(--rose);}
.sv.grn{color:var(--grn);}
.sv.red{color:var(--red);}
.sv.gold{color:var(--gold);}
.sv.blu{color:var(--blu);}
.ss{font-size:11px;color:var(--mut);margin-top:4px;}

/* ── TABLES ── */
.tw{overflow-x:auto;}
table{width:100%;border-collapse:collapse;font-size:13px;}
th{background:var(--bg2);padding:9px 12px;text-align:left;font-size:11px;font-weight:700;color:var(--mut);border-bottom:1px solid var(--bdr);white-space:nowrap;}
td{padding:10px 12px;border-bottom:1px solid var(--bdr);vertical-align:middle;}
tr:last-child td{border-bottom:none;}
tr:hover td{background:var(--bg);}

/* ── BADGES ── */
.bdg{display:inline-block;padding:2px 8px;border-radius:20px;font-size:11px;font-weight:600;background:var(--bg2);color:var(--mut);}
.bdg.rose{background:var(--roseL);color:var(--roseDk);}
.bdg.grn{background:var(--grnL);color:var(--grn);}
.bdg.red{background:var(--redL);color:var(--red);}
.bdg.gold{background:var(--goldL);color:var(--goldDk);}
.bdg.blu{background:var(--bluL);color:var(--blu);}
.bdg.pur{background:var(--purL);color:var(--pur);}
.pill{display:inline-block;padding:2px 10px;border-radius:20px;font-size:11px;font-weight:700;}
.pill.grn{background:var(--grnL);color:var(--grn);}
.pill.red{background:var(--redL);color:var(--red);}
.pill.gold{background:var(--goldL);color:var(--goldDk);}

/* ── BUTTONS ── */
.btn{border:none;border-radius:8px;padding:8px 16px;font-size:13px;cursor:pointer;font-family:'Noto Sans TC',sans-serif;font-weight:600;transition:.2s;}
.btn-p{background:linear-gradient(135deg,var(--rose),var(--roseDk));color:#fff;}
.btn-p:hover{opacity:.9;transform:translateY(-1px);}
.btn-s{background:var(--bg2);color:var(--brown);border:1.5px solid var(--bdr);}
.btn-s:hover{border-color:var(--rose);color:var(--roseDk);}
.btn-g{background:linear-gradient(135deg,var(--grn),#3a6b4c);color:#fff;}
.btn-sm{padding:5px 12px;font-size:12px;}
.bdel{background:none;border:none;color:var(--mut);cursor:pointer;padding:3px 6px;border-radius:6px;font-size:14px;transition:.2s;}
.bdel:hover{background:var(--redL);color:var(--red);}

/* ── PROGRESS ── */
.pb{background:var(--bg2);border-radius:10px;overflow:hidden;}
.pf{border-radius:10px;transition:width .4s ease;background:linear-gradient(90deg,var(--rose),var(--gold));}
.plbl{display:flex;justify-content:space-between;font-size:12px;color:var(--mut);margin-bottom:6px;}

/* ── FORMS ── */
.fg{display:grid;gap:14px;}
.fg2{grid-template-columns:1fr 1fr;}
.fgrp{display:flex;flex-direction:column;gap:5px;}
.fgrp label{font-size:12px;font-weight:600;color:var(--brown);}
.fgrp input,.fgrp select,.fgrp textarea{
  border:1.5px solid var(--bdr);border-radius:8px;padding:9px 12px;font-size:13px;
  font-family:'Noto Sans TC',sans-serif;background:var(--bg);color:var(--txt);transition:.2s;width:100%;}
.fgrp input:focus,.fgrp select:focus,.fgrp textarea:focus{outline:none;border-color:var(--rose);background:#fff;}
.yinp{border:1.5px solid var(--goldL);border-radius:8px;background:var(--goldL);}

/* ── MODAL ── */
.modal-bg{display:none;position:fixed;inset:0;background:rgba(45,31,26,0.5);z-index:1000;backdrop-filter:blur(4px);}
.modal-bg.open{display:flex;align-items:center;justify-content:center;}
.modal{background:var(--card);border-radius:18px;padding:28px;width:90%;max-width:540px;max-height:90vh;overflow-y:auto;box-shadow:var(--shadow2);}
.mttl{font-family:'Noto Serif TC',serif;font-size:17px;font-weight:700;color:var(--roseDk);margin-bottom:20px;}
.mbtns{display:flex;gap:10px;justify-content:flex-end;margin-top:20px;}

/* ── MISC ── */
.fb{display:flex;align-items:center;}
.fbb{display:flex;align-items:center;justify-content:space-between;}
.g8{gap:8px;}
.tm{color:var(--mut);font-size:12px;}
.ap{color:var(--grn);font-weight:600;}
.an{color:var(--red);font-weight:600;}
.ag{color:var(--gold);font-weight:600;}
.sec-hd{display:flex;align-items:center;justify-content:space-between;margin-bottom:12px;}
.sec-ttl{font-family:'Noto Serif TC',serif;font-size:15px;font-weight:700;color:var(--brown);}
.mo-scroll{display:flex;gap:6px;overflow-x:auto;padding-bottom:8px;margin-bottom:16px;scrollbar-width:none;}
.mo-scroll::-webkit-scrollbar{display:none;}
.mo-btn{padding:6px 14px;border-radius:20px;border:1.5px solid var(--bdr);background:var(--card);font-size:12px;cursor:pointer;white-space:nowrap;transition:.2s;color:var(--mut);}
.mo-btn.active{background:var(--rose);border-color:var(--rose);color:#fff;font-weight:700;}
.tab-bar{display:flex;gap:6px;margin-bottom:16px;}
.tab-p{padding:7px 16px;border-radius:20px;border:1.5px solid var(--bdr);background:var(--card);font-size:12px;cursor:pointer;transition:.2s;color:var(--mut);font-weight:500;}
.tab-p.active{background:var(--rose);border-color:var(--rose);color:#fff;font-weight:700;}
.empty{text-align:center;padding:40px 20px;color:var(--mut);}
.empty-ico{font-size:40px;margin-bottom:12px;}
.tag{display:inline-block;padding:1px 8px;border-radius:12px;font-size:10px;font-weight:700;}
.tag-low{background:var(--redL);color:var(--red);}
.tag-ok{background:var(--grnL);color:var(--grn);}
.tag-warn{background:var(--goldL);color:var(--goldDk);}
.run-bal{font-weight:700;font-size:12px;}
.notice{background:var(--roseL);border-left:4px solid var(--rose);border-radius:8px;padding:10px 14px;font-size:12px;color:var(--roseDk);margin-bottom:16px;}
@media(max-width:768px){
  .sidebar{width:60px;}
  .sb-logo .name,.sb-logo .sub,.sb-item span:last-child,.sb-section{display:none;}
  .sb-item{padding:12px;justify-content:center;}
  .main{margin-left:60px;}
  .sg4{grid-template-columns:repeat(2,1fr);}
  .fg2{grid-template-columns:1fr;}
}
.toast{position:fixed;bottom:24px;right:24px;background:var(--brown);color:#fff;padding:12px 20px;border-radius:12px;font-size:13px;font-weight:600;z-index:9999;opacity:0;transform:translateY(10px);transition:.3s;pointer-events:none;}
.toast.show{opacity:1;transform:translateY(0);}
.pw-err{color:var(--red);font-size:12px;margin-top:8px;text-align:center;min-height:16px;}
</style>
</head>
<body>
<!-- Firebase -->
<script type="module">
import{initializeApp}from"https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import{getFirestore,doc,collection,onSnapshot,setDoc,addDoc,deleteDoc,getDocs,query,where,orderBy}from"https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";
import{getAuth,signInWithEmailAndPassword,signOut,onAuthStateChanged}from"https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";

const FCFG={apiKey:"AIzaSyC5lPGRmFjkpBa4dIeGSIHPyMVJVxSNp3k",authDomain:"family-finance-97e44.firebaseapp.com",projectId:"family-finance-97e44"};
const app=initializeApp(FCFG);
const db=getFirestore(app);
const auth=getAuth(app);
const FID='mumei_studio'; // 沐蜜工作室獨立帳本

// ── 系統密碼設定 ──────────────────────────────────────────
const SYSTEM_PASSWORD='815391';
// 後端 Firebase 固定帳號（自動連線，使用者無需知道）
const BACKEND_EMAIL='luoxi@family.com';
const BACKEND_PASSWORD='family123';

// ── CONSTANTS ──────────────────────────────────────────────
const USERS=['珞熙','助理1','助理2','助理3'];
const SERVICE_CATS=['做臉護膚','美體按摩','霧眉','霧唇','熱蠟除毛','彩妝','其他服務'];
const PRODUCT_CATS=['保養品','彩妝品','耗材','工具','其他'];
const COURSE_CATS=['美容課程','彩妝課程','技術課程','其他課程'];
const EXP_CATS=['房租','水電','材料耗材','薪資','行銷廣告','平台費用','設備維修','進貨成本','雜支','其他'];
const PAY=['現金','信用卡','轉帳','Line Pay','其他'];
const MOS=['一月','二月','三月','四月','五月','六月','七月','八月','九月','十月','十一月','十二月'];

// ── STATE ──────────────────────────────────────────────────
let S={
  auth:false,page:'overview',month:new Date().getMonth()+1,
  subTab:0,sync:'idle',modal:null,
  data:{
    sales:[],expenses:[],purchases:[],inventory:[],
    courses:[],courseStudents:[],
    services:[],products:[],suppliers:[],staff:[],
    cashFlow:[],assets:[],settings:{}
  }
};

// ── UTILS ──────────────────────────────────────────────────
const $=id=>document.getElementById(id);
const g=id=>$('m_'+id)?.value?.trim();
const n=id=>parseFloat($('m_'+id)?.value||0)||0;
const uid=()=>Math.random().toString(36).slice(2)+Date.now().toString(36);
const fmt=v=>{const n=parseFloat(String(v||0).replace(/,/g,''))||0;return Math.abs(n)>=1000?n.toLocaleString('zh-TW'):String(Math.round(n));};
const toN=v=>parseFloat(String(v||0).replace(/,/g,''))||0;
const today=()=>new Date().toISOString().slice(0,10);
const yr=()=>{const m=S.month;return m>=8?new Date().getFullYear():new Date().getFullYear();};
const mk=()=>`${yr()}-${String(S.month).padStart(2,'0')}`;
const toast=msg=>{const t=document.createElement('div');t.className='toast';t.textContent=msg;document.body.appendChild(t);setTimeout(()=>t.classList.add('show'),10);setTimeout(()=>{t.classList.remove('show');setTimeout(()=>t.remove(),300)},2800);};
const setSyncing=s=>{S.sync=s;const d=$('sync-dot');if(d)d.className='sync-dot'+( s==='syncing'?' syncing':s==='error'?' error':'');};

// ── FIREBASE ───────────────────────────────────────────────
function startSync(){
  setSyncing('syncing');
  const mainRef=doc(db,'studios',FID);
  onSnapshot(mainRef,snap=>{
    const d=snap.data()||{};
    S.data.services=d.services||[];
    S.data.products=d.products||[];
    S.data.suppliers=d.suppliers||[];
    S.data.staff=d.staff||[];
    S.data.inventory=d.inventory||[];
    S.data.settings=d.settings||{};
    S.data.assets=d.assets||[];
    setSyncing('synced');render();
  },()=>setSyncing('error'));

  ['sales','expenses','purchases','courses','courseStudents','cashFlow'].forEach(col=>{
    onSnapshot(collection(db,'studios',FID,col),snap=>{
      S.data[col]=snap.docs.map(d=>({id:d.id,...d.data()}));
      render();
    });
  });
}
async function saveMain(data){setSyncing('syncing');try{await setDoc(doc(db,'studios',FID),data,{merge:true});setSyncing('synced');}catch{setSyncing('error');toast('儲存失敗');}}
async function addSub(col,data){setSyncing('syncing');try{const r=await addDoc(collection(db,'studios',FID,col),{...data,_at:Date.now()});setSyncing('synced');return r;}catch{setSyncing('error');toast('儲存失敗');}}
async function setSub(col,id,data){setSyncing('syncing');try{await setDoc(doc(db,'studios',FID,col,id),{...data,_at:Date.now()});setSyncing('synced');}catch{setSyncing('error');toast('儲存失敗');}}
async function delSub(col,id){setSyncing('syncing');try{await deleteDoc(doc(db,'studios',FID,col,id));setSyncing('synced');}catch{setSyncing('error');toast('刪除失敗');}}

// ── AUTH ───────────────────────────────────────────────────
// 監聽 Firebase 登入狀態（僅用於後端資料連線）
onAuthStateChanged(auth,user=>{
  if(user){
    // 後端已連線，若使用者通過密碼驗證才進入主畫面
    if(sessionStorage.getItem('mumei_unlock')==='1'){
      S.auth=true;
      startSync();
    }
  }else{
    // 後端未連線，自動以固定帳號登入
    signInWithEmailAndPassword(auth,BACKEND_EMAIL,BACKEND_PASSWORD).catch(()=>{
      console.error('後端連線失敗');
    });
  }
});

// 檢查是否已解鎖
if(sessionStorage.getItem('mumei_unlock')==='1'){
  S.auth=true;
}

// ── RENDER ─────────────────────────────────────────────────
function render(){
  $('app').innerHTML=S.auth?appHTML():loginHTML();
  bindEvents();
  // 若已解鎖但還沒開始同步，啟動同步
  if(S.auth&&S.sync==='idle'&&auth.currentUser){
    startSync();
  }
}

function loginHTML(){return`
<div style="min-height:100vh;display:flex;align-items:center;justify-content:center;background:linear-gradient(135deg,var(--roseL) 0%,var(--goldL) 100%);">
  <div style="background:var(--card);border-radius:20px;padding:40px;width:360px;box-shadow:var(--shadow2);text-align:center;">
    <div style="font-size:48px;margin-bottom:8px;">🌸</div>
    <div style="font-family:'Noto Serif TC',serif;font-size:24px;font-weight:900;color:var(--roseDk);margin-bottom:4px;">沐蜜</div>
    <div style="font-size:11px;color:var(--mut);margin-bottom:28px;letter-spacing:2px;">MUMEI STUDIO FINANCE</div>
    <div class="fgrp" style="margin-bottom:8px;text-align:left;">
      <label>請輸入密碼</label>
      <input type="password" id="m_pw" placeholder="••••••" autofocus onkeydown="if(event.key==='Enter')doLogin()">
    </div>
    <div class="pw-err" id="pw_err"></div>
    <button class="btn btn-p" style="width:100%;padding:12px;margin-top:8px;" onclick="doLogin()">進入系統</button>
  </div>
</div>`;}

function appHTML(){
  const PAGES=[
    {id:'overview',ico:'📊',lbl:'總覽'},
    {id:'sales',ico:'💰',lbl:'收入記帳'},
    {id:'order',ico:'🧴',lbl:'銷售開單'},
    {id:'expenses',ico:'💸',lbl:'支出記帳'},
    {id:'purchase',ico:'📦',lbl:'進貨管理'},
    {id:'inventory',ico:'🗂️',lbl:'庫存盤點'},
    {id:'courses',ico:'📋',lbl:'課程管理'},
    {id:'staff',ico:'👩‍💼',lbl:'員工業績'},
    {id:'cashflow',ico:'🏦',lbl:'現金流水帳'},
    {id:'reports',ico:'📈',lbl:'月報年報'},
    {id:'settings',ico:'⚙️',lbl:'基本設定'},
  ];
  const curPage=PAGES.find(p=>p.id===S.page)||PAGES[0];
  return`
<div class="app">
  <nav class="sidebar">
    <div class="sb-logo">
      <div style="font-size:28px;text-align:center;">🌸</div>
      <div class="name">沐蜜</div>
      <div class="sub">STUDIO FINANCE</div>
    </div>
    ${PAGES.map(p=>`<div class="sb-item${S.page===p.id?' active':''}" onclick="setP('${p.id}')">
      <span class="ico">${p.ico}</span><span>${p.lbl}</span>
    </div>`).join('')}
  </nav>
  <div class="main">
    <div class="topbar">
      <span class="tb-ttl">${curPage.ico} ${curPage.lbl}</span>
      <div class="tb-right">
        <div id="sync-dot" class="sync-dot"></div>
        <button class="btn-out" onclick="exportXLSX()">⬇ 匯出</button>
        <button class="btn-out" onclick="doOut()">登出</button>
      </div>
    </div>
    <div class="content" id="pg-content">${renderPage()}</div>
  </div>
</div>
<div class="modal-bg${S.modal?'.open':''}" id="modal-bg" onclick="closeMod(event)">
  <div class="modal" id="modal-inner" onclick="e=>e.stopPropagation()">${renderModal()}</div>
</div>
<div id="toast-area"></div>`;
}

function renderPage(){
  switch(S.page){
    case'overview':  return pgOverview();
    case'sales':     return pgSales();
    case'order':     return pgOrder();
    case'expenses':  return pgExpenses();
    case'purchase':  return pgPurchase();
    case'inventory': return pgInventory();
    case'courses':   return pgCourses();
    case'staff':     return pgStaff();
    case'cashflow':  return pgCashflow();
    case'reports':   return pgReports();
    case'settings':  return pgSettings();
    default:         return pgOverview();
  }
}

// ── MONTH SELECTOR ─────────────────────────────────────────
function moSel(){
  return`<div class="mo-scroll">${MOS.map((m,i)=>`<button class="mo-btn${S.month===i+1?' active':''}" onclick="setMo(${i+1})">${m}</button>`).join('')}</div>`;
}

// ── DATA HELPERS ───────────────────────────────────────────
function getSales(mo){
  const m=mo||S.month;
  return S.data.sales.filter(s=>s.month===m&&s.year===yr());
}
function getExp(mo){
  const m=mo||S.month;
  return S.data.expenses.filter(e=>e.month===m&&e.year===yr());
}
function getInv(pid){
  return S.data.inventory.find(i=>i.productId===pid)||{qty:0,minQty:5};
}
function getSvcTotal(mo){return getSales(mo).filter(s=>s.type==='service').reduce((s,r)=>s+toN(r.amount),0);}
function getPrdTotal(mo){return getSales(mo).filter(s=>s.type==='product').reduce((s,r)=>s+toN(r.amount),0);}
function getCrsTotal(mo){return getSales(mo).filter(s=>s.type==='course').reduce((s,r)=>s+toN(r.amount),0);}
function getExpTotal(mo){return getExp(mo).reduce((s,e)=>s+toN(e.amount),0);}
function getRevTotal(mo){return getSales(mo).reduce((s,r)=>s+toN(r.amount),0);}

// ─── OVERVIEW ──────────────────────────────────────────────
function pgOverview(){
  const rev=getRevTotal();
  const exp=getExpTotal();
  const profit=rev-exp;
  const svc=getSvcTotal(),prd=getPrdTotal(),crs=getCrsTotal();
  const lowStock=S.data.products.filter(p=>{const inv=getInv(p.id);return toN(inv.qty)<toN(inv.minQty||5);});
  const staffSales={};
  getSales().forEach(s=>{const u=s.staff||'未分配';staffSales[u]=(staffSales[u]||0)+toN(s.amount);});
  const pending=(S.data.courseStudents||[]).filter(s=>!s.paid&&s.month===S.month&&s.year===yr());
  const now=new Date();
  const week7=new Date(now-7*86400000).toISOString().slice(0,10);
  const recent=S.data.sales.filter(s=>s.date>=week7).reduce((s,r)=>s+toN(r.amount),0);

  return`
  <div class="pg-ttl">📊 工作室總覽</div>
  <div class="pg-sub">${yr()}年 ${MOS[S.month-1]} — 沐蜜財務一目了然</div>
  ${moSel()}

  ${lowStock.length>0?`<div class="notice">⚠️ 低庫存警示：${lowStock.map(p=>p.name).join('、')} 數量不足，請盡快補貨</div>`:''}

  <div class="sg sg4">
    <div class="sc rose">
      <div class="sl">💰 本月總營收</div>
      <div class="sv rose">$${fmt(rev)}</div>
      <div class="ss">近7天 $${fmt(recent)}</div>
    </div>
    <div class="sc red">
      <div class="sl">💸 本月總支出</div>
      <div class="sv red">$${fmt(exp)}</div>
    </div>
    <div class="sc ${profit>=0?'grn':'red'}">
      <div class="sl">📈 本月淨利</div>
      <div class="sv ${profit>=0?'grn':'red'}">$${fmt(profit)}</div>
      <div class="ss">毛利率 ${rev>0?Math.round(profit/rev*100):0}%</div>
    </div>
    <div class="sc gold">
      <div class="sl">🗂️ 庫存品項</div>
      <div class="sv gold">${S.data.products.length}</div>
      <div class="ss">低庫存 ${lowStock.length} 項</div>
    </div>
  </div>

  <div class="sg sg3">
    <div class="card">
      <div class="card-hd"><span class="card-ttl">💰 收入結構</span><button class="btn btn-p btn-sm" onclick="setP('order')">開立銷售單</button></div>
      <div class="card-bd">
        ${rev===0?'<div class="empty" style="padding:16px;"><p>本月尚無收入</p></div>':[
          {lbl:'🌸 服務',val:svc,color:'var(--rose)'},
          {lbl:'🧴 產品',val:prd,color:'var(--gold)'},
          {lbl:'📚 課程',val:crs,color:'var(--pur)'},
        ].map(({lbl,val,color})=>val>0?`
          <div style="margin-bottom:12px;">
            <div class="plbl"><span>${lbl}</span><span style="font-weight:700;color:${color};">$${fmt(val)}（${rev>0?Math.round(val/rev*100):0}%）</span></div>
            <div class="pb" style="height:10px;"><div class="pf" style="width:${rev>0?Math.round(val/rev*100):0}%;height:10px;background:${color};"></div></div>
          </div>`:'').join('')}
      </div>
    </div>

    <div class="card">
      <div class="card-hd"><span class="card-ttl">👩‍💼 本月業績</span><button class="btn btn-s btn-sm" onclick="setP('staff')">詳細</button></div>
      <div class="card-bd">
        ${Object.keys(staffSales).length===0?'<div class="empty" style="padding:16px;"><p>本月尚無業績</p></div>':
          Object.entries(staffSales).sort((a,b)=>b[1]-a[1]).map(([u,v])=>`
          <div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--bdr);">
            <span style="font-size:13px;">${u}</span>
            <span class="ap">$${fmt(v)}</span>
          </div>`).join('')}
      </div>
    </div>

    <div class="card">
      <div class="card-hd"><span class="card-ttl">📋 課程待收款</span><button class="btn btn-s btn-sm" onclick="setP('courses')">管理</button></div>
      <div class="card-bd">
        ${pending.length===0?'<div class="empty" style="padding:16px;"><p>本月無待收款</p></div>':
          pending.slice(0,5).map(s=>`
          <div style="display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid var(--bdr);">
            <div><div style="font-size:13px;">${s.studentName}</div><div class="tm">${s.courseName}</div></div>
            <div style="text-align:right;"><div class="an">$${fmt(s.amount)}</div>
              <button style="font-size:10px;margin-top:2px;" class="btn btn-g btn-sm" onclick="markPaid('${s.id}')">收款</button>
            </div>
          </div>`).join('')}
        ${pending.length>5?`<div class="tm" style="text-align:center;margin-top:8px;">還有 ${pending.length-5} 筆...</div>`:''}
      </div>
    </div>
  </div>

  <div class="sg sg2">
    <div class="card">
      <div class="card-hd"><span class="card-ttl">💰 錢從哪裡來</span></div>
      <div class="card-bd">
        <div class="sg sg2" style="margin-bottom:0;">
          <div><div class="sl">服務收入</div><div class="sv sm rose">$${fmt(svc)}</div></div>
          <div><div class="sl">產品銷售</div><div class="sv sm gold">$${fmt(prd)}</div></div>
          <div><div class="sl">課程收入</div><div class="sv sm" style="color:var(--pur);">$${fmt(crs)}</div></div>
          <div><div class="sl">總計</div><div class="sv sm rose">$${fmt(rev)}</div></div>
        </div>
      </div>
    </div>
    <div class="card">
      <div class="card-hd"><span class="card-ttl">💸 錢從哪裡去</span></div>
      <div class="card-bd">
        ${(()=>{
          const byCat={};getExp().forEach(e=>{byCat[e.category]=(byCat[e.category]||0)+toN(e.amount);});
          const entries=Object.entries(byCat).sort((a,b)=>b[1]-a[1]).slice(0,4);
          return entries.length===0?'<div class="empty" style="padding:8px;"><p>本月尚無支出</p></div>':
            entries.map(([c,v])=>`<div style="display:flex;justify-content:space-between;padding:5px 0;border-bottom:1px solid var(--bdr);"><span class="tm">${c}</span><span class="an">$${fmt(v)}</span></div>`).join('');
        })()}
        <div style="display:flex;justify-content:space-between;padding:6px 0;margin-top:4px;"><b>合計</b><b class="an">$${fmt(exp)}</b></div>
      </div>
    </div>
  </div>`;
}

// ─── SALES ─────────────────────────────────────────────────
function pgSales(){
  const sales=getSales().sort((a,b)=>(b.date||'').localeCompare(a.date||''));
  const rev=getRevTotal();
  const tabs=['全部','服務','產品','課程'];
  const filtered=S.subTab===0?sales:sales.filter(s=>['service','product','course'][S.subTab-1]===s.type);
  return`
  <div class="pg-ttl">💰 收入記帳</div>
  <div class="pg-sub">服務、產品、課程三類收入逐筆記錄</div>
  ${moSel()}
  <div class="sg sg3">
    <div class="sc rose"><div class="sl">服務收入</div><div class="sv rose">$${fmt(getSvcTotal())}</div></div>
    <div class="sc gold"><div class="sl">產品銷售</div><div class="sv gold">$${fmt(getPrdTotal())}</div></div>
    <div class="sc" style="border-top-color:var(--pur);"><div class="sl">課程收入</div><div class="sv" style="color:var(--pur);">$${fmt(getCrsTotal())}</div></div>
  </div>
  <div class="card">
    <div class="card-hd">
      <div class="tab-bar" style="margin:0;">${tabs.map((t,i)=>`<button class="tab-p${S.subTab===i?' active':''}" onclick="setT(${i})">${t}</button>`).join('')}</div>
      <button class="btn btn-p btn-sm" onclick="openM('addSale')">＋ 新增收入</button>
    </div>
    <div class="tw"><table>
      <thead><tr><th>日期</th><th>類型</th><th>項目</th><th>金額</th><th>付款</th><th>員工</th><th>客戶</th><th>備註</th><th></th></tr></thead>
      <tbody>${filtered.length===0?`<tr><td colspan="9"><div class="empty"><div class="empty-ico">💰</div><p>本月尚無記錄</p></div></td></tr>`:
        filtered.map(s=>`<tr>
          <td>${s.date||''}</td>
          <td><span class="bdg ${s.type==='service'?'rose':s.type==='product'?'gold':'pur'}">${s.type==='service'?'服務':s.type==='product'?'產品':'課程'}</span></td>
          <td><b>${s.item||''}</b></td>
          <td class="ap">$${fmt(s.amount)}</td>
          <td><span class="bdg">${s.payMethod||''}</span></td>
          <td>${s.staff||''}</td>
          <td class="tm">${s.customer||''}</td>
          <td class="tm">${s.note||''}</td>
          <td style="white-space:nowrap;">
            <button class="bdel" style="color:#aaa;font-size:13px;" onclick="openM('editSale','${s.id}')">✏️</button>
            <button class="bdel" onclick="delSale('${s.id}')">×</button>
          </td>
        </tr>`).join('')}
      </tbody>
    </table></div>
    <div class="card-bd" style="padding:12px 18px;border-top:1px solid var(--bdr);">
      <b>本月合計：</b><span class="ap">$${fmt(rev)}</span>
    </div>
  </div>`;
}

// ─── ORDER ─────────────────────────────────────────────────
function pgOrder(){
  const services=S.data.services||[];
  const products=S.data.products||[];
  return`
  <div class="pg-ttl">🧴 銷售開單</div>
  <div class="pg-sub">選擇服務或產品，自動記入收入並扣減庫存</div>
  <div class="sg sg2">
    <div class="card">
      <div class="card-hd"><span class="card-ttl">🌸 快速服務開單</span></div>
      <div class="card-bd">
        <div class="fg">
          <div class="fgrp"><label>日期</label><input type="date" id="m_od0" value="${today()}"></div>
          <div class="fgrp"><label>服務項目</label>
            <select id="m_od1">
              <option value="">-- 選擇服務 --</option>
              ${services.map(s=>`<option value="${s.id}" data-price="${s.price}">${s.name}（$${fmt(s.price)}）</option>`).join('')}
              <option value="custom">自訂項目</option>
            </select>
          </div>
          <div class="fgrp"><label>服務名稱（自訂）</label><input type="text" id="m_od2" placeholder="若選自訂請填寫"></div>
          <div class="fgrp"><label>金額</label><input type="number" id="m_od3" placeholder="0"></div>
          <div class="fgrp"><label>客戶姓名</label><input type="text" id="m_od4" placeholder="選填"></div>
          <div class="fgrp"><label>負責員工</label><select id="m_od5">${USERS.map(u=>`<option>${u}</option>`).join('')}</select></div>
          <div class="fgrp"><label>付款方式</label><select id="m_od6">${PAY.map(p=>`<option>${p}</option>`).join('')}</select></div>
          <div class="fgrp"><label>備註</label><input type="text" id="m_od7" placeholder="選填"></div>
        </div>
        <div style="margin-top:14px;"><button class="btn btn-p" style="width:100%;" onclick="saveOrder('service')">✓ 確認服務收入</button></div>
      </div>
    </div>

    <div class="card">
      <div class="card-hd"><span class="card-ttl">🧴 快速產品銷售</span></div>
      <div class="card-bd">
        <div class="fg">
          <div class="fgrp"><label>日期</label><input type="date" id="m_pd0" value="${today()}"></div>
          <div class="fgrp"><label>產品</label>
            <select id="m_pd1" onchange="fillPrice()">
              <option value="">-- 選擇產品 --</option>
              ${products.map(p=>{const inv=getInv(p.id);return`<option value="${p.id}" data-price="${p.price}" data-qty="${inv.qty}">${p.name}（庫存:${inv.qty}）$${fmt(p.price)}</option>`;}).join('')}
            </select>
          </div>
          <div class="fgrp"><label>數量</label><input type="number" id="m_pd2" placeholder="1" value="1" min="1" onchange="calcPrdTotal()"></div>
          <div class="fgrp"><label>單價</label><input type="number" id="m_pd3" placeholder="0" onchange="calcPrdTotal()"></div>
          <div class="fgrp"><label>小計</label><input type="number" id="m_pd4" placeholder="0" readonly style="background:var(--bg2);"></div>
          <div class="fgrp"><label>客戶</label><input type="text" id="m_pd5" placeholder="選填"></div>
          <div class="fgrp"><label>員工</label><select id="m_pd6">${USERS.map(u=>`<option>${u}</option>`).join('')}</select></div>
          <div class="fgrp"><label>付款方式</label><select id="m_pd7">${PAY.map(p=>`<option>${p}</option>`).join('')}</select></div>
        </div>
        <div style="margin-top:14px;"><button class="btn btn-p" style="width:100%;" onclick="saveOrder('product')">✓ 確認產品銷售</button></div>
      </div>
    </div>
  </div>

  <div class="card">
    <div class="card-hd"><span class="card-ttl">📋 今日銷售</span></div>
    <div class="tw"><table>
      <thead><tr><th>時間</th><th>類型</th><th>項目</th><th>金額</th><th>客戶</th><th>員工</th><th>付款</th></tr></thead>
      <tbody>${(()=>{
        const todays=S.data.sales.filter(s=>s.date===today()).sort((a,b)=>b._at-a._at);
        if(!todays.length)return`<tr><td colspan="7"><div class="empty"><div class="empty-ico">🧴</div><p>今日尚無銷售</p></div></td></tr>`;
        return todays.map(s=>`<tr>
          <td class="tm">${s._at?new Date(s._at).toLocaleTimeString('zh-TW',{hour:'2-digit',minute:'2-digit'}):''}</td>
          <td><span class="bdg ${s.type==='service'?'rose':s.type==='product'?'gold':'pur'}">${s.type==='service'?'服務':s.type==='product'?'產品':'課程'}</span></td>
          <td><b>${s.item}</b></td><td class="ap">$${fmt(s.amount)}</td>
          <td class="tm">${s.customer||'-'}</td><td>${s.staff||''}</td>
          <td><span class="bdg">${s.payMethod||''}</span></td>
        </tr>`).join('');
      })()}</tbody>
    </table></div>
    <div class="card-bd" style="border-top:1px solid var(--bdr);padding:10px 18px;">
      <b>今日合計：</b><span class="ap">$${fmt(S.data.sales.filter(s=>s.date===today()).reduce((s,r)=>s+toN(r.amount),0))}</span>
    </div>
  </div>`;
}

// ─── EXPENSES ──────────────────────────────────────────────
function pgExpenses(){
  const exps=getExp().sort((a,b)=>(b.date||'').localeCompare(a.date||''));
  const total=getExpTotal();
  const byCat={};exps.forEach(e=>{byCat[e.category]=(byCat[e.category]||0)+toN(e.amount);});
  return`
  <div class="pg-ttl">💸 支出記帳</div>
  <div class="pg-sub">工作室所有支出，含固定與變動費用</div>
  ${moSel()}
  <div class="sg sg4">
    ${Object.entries(byCat).sort((a,b)=>b[1]-a[1]).slice(0,4).map(([c,v])=>`<div class="sc red"><div class="sl">${c}</div><div class="sv red sm">$${fmt(v)}</div></div>`).join('')}
  </div>
  <div class="card">
    <div class="card-hd">
      <span class="card-ttl">支出明細</span>
      <button class="btn btn-p btn-sm" onclick="openM('addExp')">＋ 新增支出</button>
    </div>
    <div class="tw"><table>
      <thead><tr><th>日期</th><th>類別</th><th>金額</th><th>付款</th><th>供應商/備註</th><th></th></tr></thead>
      <tbody>${exps.length===0?`<tr><td colspan="6"><div class="empty"><div class="empty-ico">💸</div><p>本月尚無支出</p></div></td></tr>`:
        exps.map(e=>`<tr>
          <td>${e.date||''}</td>
          <td><span class="bdg">${e.category||''}</span></td>
          <td class="an">$${fmt(e.amount)}</td>
          <td><span class="bdg">${e.payMethod||''}</span></td>
          <td class="tm">${e.note||''}</td>
          <td><button class="bdel" style="color:#aaa;font-size:13px;" onclick="openM('editExp','${e.id}')">✏️</button><button class="bdel" onclick="delExp('${e.id}')">×</button></td>
        </tr>`).join('')}
      </tbody>
    </table></div>
    <div class="card-bd" style="border-top:1px solid var(--bdr);padding:10px 18px;"><b>本月合計：</b><span class="an">$${fmt(total)}</span></div>
  </div>`;
}

// ─── PURCHASE ──────────────────────────────────────────────
function pgPurchase(){
  const purchases=S.data.purchases.filter(p=>p.month===S.month&&p.year===yr()).sort((a,b)=>(b.date||'').localeCompare(a.date||''));
  const total=purchases.reduce((s,p)=>s+toN(p.totalCost),0);
  return`
  <div class="pg-ttl">📦 進貨管理</div>
  <div class="pg-sub">輸入進貨單，自動增加庫存數量</div>
  ${moSel()}
  <div class="sg sg3">
    <div class="sc gold"><div class="sl">本月進貨筆數</div><div class="sv gold">${purchases.length}</div></div>
    <div class="sc red"><div class="sl">本月進貨成本</div><div class="sv red">$${fmt(total)}</div></div>
    <div class="sc"><div class="sl">供應商數</div><div class="sv">${(S.data.suppliers||[]).length}</div></div>
  </div>
  <div class="sec-hd">
    <span class="sec-ttl">進貨記錄</span>
    <button class="btn btn-p btn-sm" onclick="openM('addPurchase')">＋ 新增進貨</button>
  </div>
  <div class="card">
    <div class="tw"><table>
      <thead><tr><th>日期</th><th>產品</th><th>數量</th><th>單價</th><th>總成本</th><th>供應商</th><th>備註</th><th></th></tr></thead>
      <tbody>${purchases.length===0?`<tr><td colspan="8"><div class="empty"><div class="empty-ico">📦</div><p>本月尚無進貨記錄</p></div></td></tr>`:
        purchases.map(p=>`<tr>
          <td>${p.date||''}</td>
          <td><b>${p.productName||''}</b></td>
          <td>${p.qty}</td>
          <td class="tm">$${fmt(p.unitCost)}</td>
          <td class="an">$${fmt(p.totalCost)}</td>
          <td class="tm">${p.supplier||''}</td>
          <td class="tm">${p.note||''}</td>
          <td><button class="bdel" onclick="delPurchase('${p.id}')">×</button></td>
        </tr>`).join('')}
      </tbody>
    </table></div>
  </div>`;
}

// ─── INVENTORY ─────────────────────────────────────────────
function pgInventory(){
  const prods=S.data.products||[];
  return`
  <div class="pg-ttl">🗂️ 庫存盤點</div>
  <div class="pg-sub">即時庫存數量，低量自動警示</div>
  <div class="sg sg4">
    <div class="sc"><div class="sl">產品總數</div><div class="sv">${prods.length}</div></div>
    <div class="sc red"><div class="sl">低庫存警示</div><div class="sv red">${prods.filter(p=>{const i=getInv(p.id);return toN(i.qty)<toN(i.minQty||5);}).length}</div></div>
    <div class="sc gold"><div class="sl">庫存總價值</div><div class="sv gold sm">$${fmt(prods.reduce((s,p)=>{const i=getInv(p.id);return s+toN(i.qty)*toN(p.cost||0);},0))}</div></div>
    <div class="sc grn"><div class="sl">正常庫存</div><div class="sv grn">${prods.filter(p=>{const i=getInv(p.id);return toN(i.qty)>=toN(i.minQty||5);}).length}</div></div>
  </div>
  <div class="card">
    <div class="card-hd"><span class="card-ttl">庫存清單</span><button class="btn btn-p btn-sm" onclick="openM('stockCheck')">📋 手動盤點</button></div>
    <div class="tw"><table>
      <thead><tr><th>產品名稱</th><th>類別</th><th>現有庫存</th><th>最低庫存</th><th>狀態</th><th>成本單價</th><th>庫存價值</th><th>快速調整</th></tr></thead>
      <tbody>${prods.length===0?`<tr><td colspan="8"><div class="empty"><div class="empty-ico">📦</div><p>尚未建立產品，請先到「基本設定」新增產品</p></div></td></tr>`:
        prods.map(p=>{
          const inv=getInv(p.id);
          const qty=toN(inv.qty);
          const minQty=toN(inv.minQty||5);
          const val=qty*toN(p.cost||0);
          const status=qty<=0?'<span class="tag tag-low">缺貨</span>':qty<minQty?'<span class="tag tag-warn">偏低</span>':'<span class="tag tag-ok">正常</span>';
          return`<tr style="${qty<minQty?'background:var(--redL);':''}">
            <td><b>${p.name}</b></td>
            <td><span class="bdg">${p.category||''}</span></td>
            <td style="font-size:16px;font-weight:700;color:${qty<minQty?'var(--red)':'var(--grn)'};">${qty}</td>
            <td class="tm">${minQty}</td>
            <td>${status}</td>
            <td class="tm">$${fmt(p.cost||0)}</td>
            <td class="ag">$${fmt(val)}</td>
            <td>
              <div style="display:flex;align-items:center;gap:6px;">
                <button class="btn btn-s btn-sm" onclick="adjInv('${p.id}',-1)">－</button>
                <span style="font-weight:700;min-width:28px;text-align:center;">${qty}</span>
                <button class="btn btn-p btn-sm" onclick="adjInv('${p.id}',1)">＋</button>
                <button class="btn btn-s btn-sm" onclick="openM('setInv','${p.id}')">設定</button>
              </div>
            </td>
          </tr>`;
        }).join('')}
      </tbody>
    </table></div>
  </div>`;
}

// ─── COURSES ───────────────────────────────────────────────
function pgCourses(){
  const courses=S.data.courses||[];
  const students=S.data.courseStudents||[];
  const moCourses=courses.filter(c=>c.month===S.month&&c.year===yr());
  const moStudents=students.filter(s=>s.month===S.month&&s.year===yr());
  const totalRev=moStudents.filter(s=>s.paid).reduce((s,r)=>s+toN(r.amount),0);
  const pending=moStudents.filter(s=>!s.paid).reduce((s,r)=>s+toN(r.amount),0);
  return`
  <div class="pg-ttl">📋 課程管理</div>
  <div class="pg-sub">課程場次管理、學員名單、收款追蹤</div>
  ${moSel()}
  <div class="sg sg3">
    <div class="sc pur"><div class="sl">本月課程場次</div><div class="sv" style="color:var(--pur);">${moCourses.length}</div></div>
    <div class="sc grn"><div class="sl">已收款</div><div class="sv grn">$${fmt(totalRev)}</div></div>
    <div class="sc red"><div class="sl">待收款</div><div class="sv red">$${fmt(pending)}</div></div>
  </div>
  <div class="sg sg2">
    <div class="card">
      <div class="card-hd"><span class="card-ttl">📚 課程場次</span><button class="btn btn-p btn-sm" onclick="openM('addCourse')">＋ 新增課程</button></div>
      <div class="card-bd">
        ${moCourses.length===0?'<div class="empty" style="padding:16px;"><p>本月尚無課程</p></div>':
          moCourses.map(c=>`
          <div style="border:1.5px solid var(--bdr);border-radius:10px;padding:12px;margin-bottom:8px;">
            <div class="fbb">
              <div><b>${c.name}</b><span class="bdg pur" style="margin-left:8px;">${c.category||''}</span></div>
              <button class="bdel" onclick="delCourse('${c.id}')">×</button>
            </div>
            <div class="tm" style="margin-top:4px;">📅 ${c.date} ｜ 👩‍🏫 ${c.instructor||''} ｜ 💰 $${fmt(c.price)}/人</div>
            <button class="btn btn-s btn-sm" style="margin-top:8px;" onclick="openM('addStudent','${c.id}')">＋ 新增學員</button>
          </div>`).join('')}
      </div>
    </div>

    <div class="card">
      <div class="card-hd"><span class="card-ttl">👥 學員收款</span></div>
      <div class="tw"><table>
        <thead><tr><th>學員</th><th>課程</th><th>金額</th><th>狀態</th><th></th></tr></thead>
        <tbody>${moStudents.length===0?`<tr><td colspan="5"><div class="empty"><p>尚無學員</p></div></td></tr>`:
          moStudents.sort((a,b)=>a.paid-b.paid).map(s=>`<tr>
            <td><b>${s.studentName}</b></td>
            <td class="tm">${s.courseName||''}</td>
            <td class="${s.paid?'ap':'an'}">$${fmt(s.amount)}</td>
            <td>${s.paid?'<span class="pill grn">已收款</span>':'<span class="pill red">待收款</span>'}</td>
            <td>${s.paid?'':` <button class="btn btn-g btn-sm" onclick="markPaid('${s.id}')">收款</button>`}
              <button class="bdel" onclick="delStudent('${s.id}')">×</button></td>
          </tr>`).join('')}
        </tbody>
      </table></div>
    </div>
  </div>`;
}

// ─── STAFF ─────────────────────────────────────────────────
function pgStaff(){
  const staff=S.data.staff||[];
  const sales=getSales();
  const byStaff={};
  sales.forEach(s=>{
    const u=s.staff||'未分配';
    if(!byStaff[u])byStaff[u]={total:0,cnt:0,svc:0,prd:0};
    byStaff[u].total+=toN(s.amount);byStaff[u].cnt++;
    if(s.type==='service')byStaff[u].svc+=toN(s.amount);
    if(s.type==='product')byStaff[u].prd+=toN(s.amount);
  });
  return`
  <div class="pg-ttl">👩‍💼 員工業績</div>
  <div class="pg-sub">本月各員工服務筆數、金額、抽成計算</div>
  ${moSel()}
  <div class="sg ${staff.length>0?'sg'+Math.min(4,staff.length+1):'sg2'}">
    ${Object.entries(byStaff).sort((a,b)=>b[1].total-a[1].total).map(([name,d])=>`
    <div class="sc rose">
      <div class="sl">${name}</div>
      <div class="sv rose sm">$${fmt(d.total)}</div>
      <div class="ss">${d.cnt} 筆 ｜ 服務 $${fmt(d.svc)} / 產品 $${fmt(d.prd)}</div>
    </div>`).join('')}
  </div>
  <div class="card">
    <div class="card-hd"><span class="card-ttl">👥 員工管理</span><button class="btn btn-p btn-sm" onclick="openM('addStaff')">＋ 新增員工</button></div>
    <div class="tw"><table>
      <thead><tr><th>姓名</th><th>職稱</th><th>抽成比例</th><th>本月業績</th><th>本月抽成</th><th></th></tr></thead>
      <tbody>${staff.length===0?`<tr><td colspan="6"><div class="empty"><p>尚未建立員工資料</p></div></td></tr>`:
        staff.map(u=>{
          const d=byStaff[u.name]||{total:0,cnt:0};
          const comm=Math.round(d.total*(toN(u.commRate||0)/100));
          return`<tr>
            <td><b>${u.name}</b></td>
            <td><span class="bdg">${u.title||''}</span></td>
            <td>${u.commRate||0}%</td>
            <td class="ap">$${fmt(d.total)}</td>
            <td class="ag">$${fmt(comm)}</td>
            <td><button class="bdel" onclick="delStaff('${u.id}')">×</button></td>
          </tr>`;
        }).join('')}
      </tbody>
    </table></div>
  </div>
  <div class="card">
    <div class="card-hd"><span class="card-ttl">📋 本月銷售明細（依員工）</span></div>
    <div class="tw"><table>
      <thead><tr><th>日期</th><th>員工</th><th>項目</th><th>類型</th><th>金額</th><th>客戶</th></tr></thead>
      <tbody>${sales.sort((a,b)=>(b.date||'').localeCompare(a.date||'')).map(s=>`<tr>
        <td>${s.date||''}</td>
        <td><b>${s.staff||'未分配'}</b></td>
        <td>${s.item||''}</td>
        <td><span class="bdg ${s.type==='service'?'rose':s.type==='product'?'gold':'pur'}">${s.type==='service'?'服務':s.type==='product'?'產品':'課程'}</span></td>
        <td class="ap">$${fmt(s.amount)}</td>
        <td class="tm">${s.customer||''}</td>
      </tr>`).join('')}</tbody>
    </table></div>
  </div>`;
}

// ─── CASHFLOW ──────────────────────────────────────────────
function pgCashflow(){
  const entries=S.data.cashFlow.filter(e=>e.month===S.month&&e.year===yr()).sort((a,b)=>(a.date||'').localeCompare(b.date||''));
  const allEntries=S.data.cashFlow.sort((a,b)=>(a.date||'').localeCompare(b.date||''));
  const curY=yr(),curM=S.month;
  const prevBal=allEntries.filter(e=>e.year<curY||(e.year===curY&&e.month<curM)).reduce((s,e)=>s+toN(e.amount),0);
  let running=prevBal;
  const rows=entries.map(e=>{running+=toN(e.amount);return{...e,run:running};});
  const moBal=running;
  const cumTotal=allEntries.reduce((s,e)=>s+toN(e.amount),0);
  const mIn=entries.filter(e=>toN(e.amount)>0).reduce((s,e)=>s+toN(e.amount),0);
  const mOut=entries.filter(e=>toN(e.amount)<0).reduce((s,e)=>s+Math.abs(toN(e.amount)),0);
  const CF_CATS=['存入','服務收入','產品收入','課程收入','現金提款','費用支出','進貨','薪資發放','其他'];
  return`
  <div class="pg-ttl">🏦 現金流水帳</div>
  <div class="pg-sub">工作室現金帳，隨時核對實際餘額</div>
  ${moSel()}
  <div class="sg sg4">
    <div class="sc"><div class="sl">上月結餘</div><div class="sv sm gold">$${fmt(prevBal)}</div></div>
    <div class="sc grn"><div class="sl">本月存入</div><div class="sv grn sm">$${fmt(mIn)}</div></div>
    <div class="sc red"><div class="sl">本月支出</div><div class="sv red sm">$${fmt(mOut)}</div></div>
    <div class="sc ${moBal>=0?'grn':'red'}">
      <div class="sl">本月結餘</div>
      <div class="sv sm ${moBal>=0?'grn':'red'}">$${fmt(moBal)}</div>
    </div>
  </div>
  <div class="card">
    <div class="card-hd">
      <span class="card-ttl">🏦 現金流明細</span>
      <button class="btn btn-p btn-sm" onclick="openM('addCF')">＋ 新增</button>
    </div>
    <div class="tw"><table>
      <thead><tr><th>日期</th><th>金額</th><th>類別</th><th>說明</th><th>累計餘額</th><th></th></tr></thead>
      <tbody>
        <tr style="background:var(--bg2);font-weight:700;">
          <td colspan="2" style="font-size:12px;color:var(--mut);">📅 上月結餘</td>
          <td colspan="2"></td>
          <td style="font-weight:700;color:${prevBal>=0?'var(--grn)':'var(--red)'};">$${fmt(prevBal)}</td>
          <td></td>
        </tr>
        ${rows.length===0?`<tr><td colspan="6"><div class="empty"><div class="empty-ico">🏦</div><p>本月尚無記錄</p></div></td></tr>`:
          rows.map(r=>`<tr>
            <td>${r.date||''}</td>
            <td class="${toN(r.amount)>=0?'ap':'an'}">${toN(r.amount)>=0?'+':''}$${fmt(r.amount)}</td>
            <td><span class="bdg">${r.cat||''}</span></td>
            <td class="tm">${r.note||''}</td>
            <td class="run-bal" style="color:${r.run>=0?'var(--grn)':'var(--red)'};">$${fmt(r.run)}</td>
            <td><button class="bdel" onclick="delCF('${r.id}')">×</button></td>
          </tr>`).join('')}
        <tr style="background:${moBal>=0?'var(--grnL)':'var(--redL)'};font-weight:700;border-top:2px solid var(--bdr);">
          <td colspan="2" style="font-size:12px;color:var(--mut);">💰 本月結餘</td>
          <td style="font-size:11px;color:var(--mut);">存入$${fmt(mIn)} / 支出$${fmt(mOut)}</td>
          <td></td>
          <td style="font-size:16px;font-weight:700;color:${moBal>=0?'var(--grn)':'var(--red)'};">$${fmt(moBal)}</td>
          <td></td>
        </tr>
      </tbody>
    </table></div>
  </div>`;
}

// ─── REPORTS ───────────────────────────────────────────────
function pgReports(){
  const mos=Array.from({length:12},(_,i)=>i+1);
  const moData=mos.map(m=>({
    m,rev:getRevTotal(m),exp:getExpTotal(m),
    svc:getSvcTotal(m),prd:getPrdTotal(m),crs:getCrsTotal(m)
  }));
  const tRev=moData.reduce((s,m)=>s+m.rev,0);
  const tExp=moData.reduce((s,m)=>s+m.exp,0);
  const maxV=Math.max(...moData.map(m=>Math.max(m.rev,m.exp)),1);
  const tSvc=moData.reduce((s,m)=>s+m.svc,0);
  const tPrd=moData.reduce((s,m)=>s+m.prd,0);
  const tCrs=moData.reduce((s,m)=>s+m.crs,0);
  return`
  <div class="pg-ttl">📈 月報年報</div>
  <div class="pg-sub">${yr()}年 全年財務分析</div>
  <div class="sg sg4">
    <div class="sc rose"><div class="sl">年度總營收</div><div class="sv rose sm">$${fmt(tRev)}</div></div>
    <div class="sc red"><div class="sl">年度總支出</div><div class="sv red sm">$${fmt(tExp)}</div></div>
    <div class="sc ${tRev-tExp>=0?'grn':'red'}"><div class="sl">年度淨利</div><div class="sv sm ${tRev-tExp>=0?'grn':'red'}">$${fmt(tRev-tExp)}</div></div>
    <div class="sc gold"><div class="sl">年度毛利率</div><div class="sv gold sm">${tRev>0?Math.round((tRev-tExp)/tRev*100):0}%</div></div>
  </div>
  <div class="card">
    <div class="card-hd"><span class="card-ttl">📊 月度營收走勢</span></div>
    <div class="card-bd" style="overflow-x:auto;">
      <div style="min-width:560px;">
        ${moData.map(m=>{
          const rp=Math.round(m.rev/maxV*100),ep=Math.round(m.exp/maxV*100);
          return`<div style="display:flex;align-items:center;gap:8px;margin-bottom:10px;">
            <div style="width:28px;font-size:11px;color:var(--mut);text-align:right;">${m.m}月</div>
            <div style="flex:1;display:flex;flex-direction:column;gap:3px;">
              <div class="pb" style="height:11px;"><div class="pf" style="width:${rp}%;height:11px;background:var(--rose);"></div></div>
              <div class="pb" style="height:11px;"><div class="pf" style="width:${ep}%;height:11px;background:var(--red);"></div></div>
            </div>
            <div style="width:100px;font-size:11px;"><div class="ap">$${fmt(m.rev)}</div><div class="an">$${fmt(m.exp)}</div></div>
            <div style="width:80px;font-size:11px;font-weight:700;${m.rev-m.exp>=0?'color:var(--grn)':'color:var(--red)'};">$${fmt(m.rev-m.exp)}</div>
          </div>`;
        }).join('')}
        <div style="display:flex;gap:16px;margin-top:8px;font-size:11px;color:var(--mut);"><span>🌸 營收</span><span>🔴 支出</span><span>右側＝淨利</span></div>
      </div>
    </div>
  </div>
  <div class="sg sg2">
    <div class="card">
      <div class="card-hd"><span class="card-ttl">💰 年度收入結構</span></div>
      <div class="card-bd">
        ${tRev===0?'<div class="empty"><p>尚無收入記錄</p></div>':[
          {lbl:'🌸 服務收入',val:tSvc,color:'var(--rose)'},
          {lbl:'🧴 產品銷售',val:tPrd,color:'var(--gold)'},
          {lbl:'📚 課程收入',val:tCrs,color:'var(--pur)'},
        ].map(({lbl,val,color})=>`
          <div style="margin-bottom:14px;">
            <div class="plbl"><span>${lbl}</span><span style="font-weight:700;color:${color};">$${fmt(val)}（${tRev>0?Math.round(val/tRev*100):0}%）</span></div>
            <div class="pb" style="height:12px;"><div class="pf" style="width:${tRev>0?Math.round(val/tRev*100):0}%;height:12px;background:${color};"></div></div>
          </div>`).join('')}
      </div>
    </div>
    <div class="card">
      <div class="card-hd"><span class="card-ttl">📋 月度明細表</span></div>
      <div class="tw"><table>
        <thead><tr><th>月份</th><th>營收</th><th>支出</th><th>淨利</th><th>毛利率</th></tr></thead>
        <tbody>${moData.map(m=>`<tr>
          <td><b>${m.m}月</b></td>
          <td class="ap">$${fmt(m.rev)}</td>
          <td class="an">$${fmt(m.exp)}</td>
          <td class="${m.rev-m.exp>=0?'ap':'an'}">$${fmt(m.rev-m.exp)}</td>
          <td class="${m.rev>0&&(m.rev-m.exp)/m.rev>0?'ap':'an'}">${m.rev>0?Math.round((m.rev-m.exp)/m.rev*100):0}%</td>
        </tr>`).join('')}</tbody>
      </table></div>
    </div>
  </div>`;
}

// ─── SETTINGS ──────────────────────────────────────────────
function pgSettings(){
  const svcs=S.data.services||[];
  const prods=S.data.products||[];
  const sups=S.data.suppliers||[];
  return`
  <div class="pg-ttl">⚙️ 基本設定</div>
  <div class="pg-sub">服務項目、產品清單、供應商管理</div>
  <div class="sg sg3">
    <div class="card">
      <div class="card-hd"><span class="card-ttl">🌸 服務項目</span><button class="btn btn-p btn-sm" onclick="openM('addService')">＋ 新增</button></div>
      <div class="card-bd" style="padding:0;">
        ${svcs.length===0?'<div class="empty" style="padding:20px;"><p>尚未建立服務項目</p></div>':
          svcs.map(s=>`<div style="display:flex;align-items:center;justify-content:space-between;padding:10px 16px;border-bottom:1px solid var(--bdr);">
            <div><b>${s.name}</b><span class="bdg rose" style="margin-left:6px;">${s.category||''}</span></div>
            <div class="fb g8"><span class="ap">$${fmt(s.price)}</span><button class="bdel" onclick="delService('${s.id}')">×</button></div>
          </div>`).join('')}
      </div>
    </div>
    <div class="card">
      <div class="card-hd"><span class="card-ttl">🧴 產品清單</span><button class="btn btn-p btn-sm" onclick="openM('addProduct')">＋ 新增</button></div>
      <div class="card-bd" style="padding:0;">
        ${prods.length===0?'<div class="empty" style="padding:20px;"><p>尚未建立產品</p></div>':
          prods.map(p=>`<div style="display:flex;align-items:center;justify-content:space-between;padding:10px 16px;border-bottom:1px solid var(--bdr);">
            <div><b>${p.name}</b><span class="bdg gold" style="margin-left:6px;">${p.category||''}</span></div>
            <div class="fb g8 tm">售$${fmt(p.price)} / 成$${fmt(p.cost||0)}<button class="bdel" onclick="delProduct('${p.id}')">×</button></div>
          </div>`).join('')}
      </div>
    </div>
    <div class="card">
      <div class="card-hd"><span class="card-ttl">🏭 供應商</span><button class="btn btn-p btn-sm" onclick="openM('addSupplier')">＋ 新增</button></div>
      <div class="card-bd" style="padding:0;">
        ${sups.length===0?'<div class="empty" style="padding:20px;"><p>尚未建立供應商</p></div>':
          sups.map(s=>`<div style="display:flex;align-items:center;justify-content:space-between;padding:10px 16px;border-bottom:1px solid var(--bdr);">
            <div><b>${s.name}</b><div class="tm">${s.contact||''} ${s.phone||''}</div></div>
            <button class="bdel" onclick="delSupplier('${s.id}')">×</button>
          </div>`).join('')}
      </div>
    </div>
  </div>`;
}

// ─── MODAL ─────────────────────────────────────────────────
function renderModal(){
  if(!S.modal)return'';
  const m=S.modal;let c='';
  const cats=c=>c.map(v=>`<option>${v}</option>`).join('');

  if(m.t==='addSale'){c=`<div class="mttl">💰 新增收入</div>
    <div class="fg fg2">
      <div class="fgrp"><label>日期</label><input type="date" id="m_s0" value="${today()}"></div>
      <div class="fgrp"><label>類型</label><select id="m_s1"><option value="service">🌸 服務</option><option value="product">🧴 產品</option><option value="course">📚 課程</option></select></div>
      <div class="fgrp"><label>項目名稱</label><input type="text" id="m_s2" placeholder="服務/產品名稱"></div>
      <div class="fgrp"><label>金額</label><input type="number" id="m_s3" placeholder="0"></div>
      <div class="fgrp"><label>付款方式</label><select id="m_s4">${PAY.map(p=>`<option>${p}</option>`).join('')}</select></div>
      <div class="fgrp"><label>員工</label><select id="m_s5">${USERS.map(u=>`<option>${u}</option>`).join('')}</select></div>
      <div class="fgrp"><label>客戶</label><input type="text" id="m_s6" placeholder="選填"></div>
      <div class="fgrp"><label>備註</label><input type="text" id="m_s7" placeholder="選填"></div>
    </div>
    <div class="mbtns"><button class="btn btn-s" onclick="closeMod()">取消</button><button class="btn btn-p" onclick="saveSale()">✓ 確認</button></div>`;}

  if(m.t==='editSale'){const s=(S.data.sales||[]).find(x=>x.id===m.x)||{};c=`<div class="mttl">✏️ 編輯收入</div>
    <div class="fg fg2">
      <div class="fgrp"><label>日期</label><input type="date" id="m_s0" value="${s.date||today()}"></div>
      <div class="fgrp"><label>類型</label><select id="m_s1"><option value="service"${s.type==='service'?' selected':''}>🌸 服務</option><option value="product"${s.type==='product'?' selected':''}>🧴 產品</option><option value="course"${s.type==='course'?' selected':''}>📚 課程</option></select></div>
      <div class="fgrp"><label>項目名稱</label><input type="text" id="m_s2" value="${s.item||''}"></div>
      <div class="fgrp"><label>金額</label><input type="number" id="m_s3" value="${s.amount||0}"></div>
      <div class="fgrp"><label>付款方式</label><select id="m_s4">${PAY.map(p=>`<option${p===s.payMethod?' selected':''}>${p}</option>`).join('')}</select></div>
      <div class="fgrp"><label>員工</label><select id="m_s5">${USERS.map(u=>`<option${u===s.staff?' selected':''}>${u}</option>`).join('')}</select></div>
      <div class="fgrp"><label>客戶</label><input type="text" id="m_s6" value="${s.customer||''}"></div>
      <div class="fgrp"><label>備註</label><input type="text" id="m_s7" value="${s.note||''}"></div>
    </div>
    <div class="mbtns"><button class="btn btn-s" onclick="closeMod()">取消</button><button class="btn btn-p" onclick="updSale('${s.id}')">✓ 更新</button></div>`;}

  if(m.t==='addExp'){c=`<div class="mttl">💸 新增支出</div>
    <div class="fg fg2">
      <div class="fgrp"><label>日期</label><input type="date" id="m_e0" value="${today()}"></div>
      <div class="fgrp"><label>類別</label><select id="m_e1">${EXP_CATS.map(c=>`<option>${c}</option>`).join('')}</select></div>
      <div class="fgrp"><label>金額</label><input type="number" id="m_e2" placeholder="0"></div>
      <div class="fgrp"><label>付款方式</label><select id="m_e3">${PAY.map(p=>`<option>${p}</option>`).join('')}</select></div>
      <div class="fgrp" style="grid-column:1/-1"><label>備註/供應商</label><input type="text" id="m_e4" placeholder="選填"></div>
    </div>
    <div class="mbtns"><button class="btn btn-s" onclick="closeMod()">取消</button><button class="btn btn-p" onclick="saveExp()">✓ 確認</button></div>`;}

  if(m.t==='editExp'){const e=(S.data.expenses||[]).find(x=>x.id===m.x)||{};c=`<div class="mttl">✏️ 編輯支出</div>
    <div class="fg fg2">
      <div class="fgrp"><label>日期</label><input type="date" id="m_e0" value="${e.date||today()}"></div>
      <div class="fgrp"><label>類別</label><select id="m_e1">${EXP_CATS.map(c=>`<option${c===e.category?' selected':''}>${c}</option>`).join('')}</select></div>
      <div class="fgrp"><label>金額</label><input type="number" id="m_e2" value="${e.amount||0}"></div>
      <div class="fgrp"><label>付款方式</label><select id="m_e3">${PAY.map(p=>`<option${p===e.payMethod?' selected':''}>${p}</option>`).join('')}</select></div>
      <div class="fgrp" style="grid-column:1/-1"><label>備註</label><input type="text" id="m_e4" value="${e.note||''}"></div>
    </div>
    <div class="mbtns"><button class="btn btn-s" onclick="closeMod()">取消</button><button class="btn btn-p" onclick="updExp('${e.id}')">✓ 更新</button></div>`;}

  if(m.t==='addPurchase'){c=`<div class="mttl">📦 新增進貨</div>
    <div class="fg fg2">
      <div class="fgrp"><label>日期</label><input type="date" id="m_p0" value="${today()}"></div>
      <div class="fgrp"><label>產品</label><select id="m_p1">
        ${(S.data.products||[]).map(p=>`<option value="${p.id}">${p.name}</option>`).join('')}
        <option value="custom">其他（手動輸入）</option>
      </select></div>
      <div class="fgrp"><label>產品名稱（手動）</label><input type="text" id="m_p2" placeholder="若選其他請填寫"></div>
      <div class="fgrp"><label>進貨數量</label><input type="number" id="m_p3" placeholder="0" min="1"></div>
      <div class="fgrp"><label>進貨單價</label><input type="number" id="m_p4" placeholder="0" onchange="calcPTotal()"></div>
      <div class="fgrp"><label>進貨總成本</label><input type="number" id="m_p5" placeholder="0"></div>
      <div class="fgrp"><label>供應商</label><select id="m_p6">
        <option value="">無</option>
        ${(S.data.suppliers||[]).map(s=>`<option>${s.name}</option>`).join('')}
      </select></div>
      <div class="fgrp"><label>備註</label><input type="text" id="m_p7" placeholder="選填"></div>
    </div>
    <div class="mbtns"><button class="btn btn-s" onclick="closeMod()">取消</button><button class="btn btn-p" onclick="savePurchase()">✓ 確認進貨</button></div>`;}

  if(m.t==='setInv'){const p=(S.data.products||[]).find(x=>x.id===m.x)||{};const inv=getInv(m.x);c=`<div class="mttl">🗂️ 設定庫存：${p.name}</div>
    <div class="fg">
      <div class="fgrp"><label>現有庫存數量</label><input type="number" id="m_i0" value="${inv.qty||0}"></div>
      <div class="fgrp"><label>最低庫存警示</label><input type="number" id="m_i1" value="${inv.minQty||5}"></div>
    </div>
    <div class="mbtns"><button class="btn btn-s" onclick="closeMod()">取消</button><button class="btn btn-p" onclick="saveInv('${m.x}')">✓ 儲存</button></div>`;}

  if(m.t==='stockCheck'){c=`<div class="mttl">📋 手動盤點</div>
    <div style="max-height:400px;overflow-y:auto;">
    ${(S.data.products||[]).map(p=>{const inv=getInv(p.id);return`
      <div style="display:flex;align-items:center;gap:12px;padding:10px 0;border-bottom:1px solid var(--bdr);">
        <div style="flex:1;"><b>${p.name}</b><div class="tm">系統庫存：${inv.qty}</div></div>
        <div style="width:100px;"><input type="number" id="ck_${p.id}" placeholder="${inv.qty}" style="width:100%;padding:6px 8px;border:1.5px solid var(--bdr);border-radius:8px;font-size:13px;"></div>
      </div>`}).join('')}
    </div>
    <div class="mbtns"><button class="btn btn-s" onclick="closeMod()">取消</button><button class="btn btn-p" onclick="saveStockCheck()">✓ 確認盤點</button></div>`;}

  if(m.t==='addCourse'){c=`<div class="mttl">📋 新增課程</div>
    <div class="fg fg2">
      <div class="fgrp"><label>課程名稱</label><input type="text" id="m_c0" placeholder="課程名稱"></div>
      <div class="fgrp"><label>類別</label><select id="m_c1">${COURSE_CATS.map(c=>`<option>${c}</option>`).join('')}</select></div>
      <div class="fgrp"><label>日期</label><input type="date" id="m_c2" value="${today()}"></div>
      <div class="fgrp"><label>講師</label><select id="m_c3">${USERS.map(u=>`<option>${u}</option>`).join('')}</select></div>
      <div class="fgrp"><label>學費/人</label><input type="number" id="m_c4" placeholder="0"></div>
      <div class="fgrp"><label>備註</label><input type="text" id="m_c5" placeholder="選填"></div>
    </div>
    <div class="mbtns"><button class="btn btn-s" onclick="closeMod()">取消</button><button class="btn btn-p" onclick="saveCourse()">✓ 新增</button></div>`;}

  if(m.t==='addStudent'){const crs=(S.data.courses||[]).find(x=>x.id===m.x)||{};c=`<div class="mttl">👥 新增學員：${crs.name||''}</div>
    <div class="fg">
      <div class="fgrp"><label>學員姓名</label><input type="text" id="m_st0" placeholder="姓名"></div>
      <div class="fgrp"><label>學費金額</label><input type="number" id="m_st1" value="${crs.price||0}"></div>
      <div class="fgrp"><label>付款狀態</label><select id="m_st2"><option value="0">未付款</option><option value="1">已付款</option></select></div>
      <div class="fgrp"><label>付款方式</label><select id="m_st3">${PAY.map(p=>`<option>${p}</option>`).join('')}</select></div>
    </div>
    <div class="mbtns"><button class="btn btn-s" onclick="closeMod()">取消</button><button class="btn btn-p" onclick="saveStudent('${m.x}','${crs.name||''}')">✓ 新增</button></div>`;}

  if(m.t==='addCF'){const CF_CATS=['存入','服務收入','產品收入','課程收入','現金提款','費用支出','進貨','薪資發放','其他'];c=`<div class="mttl">🏦 新增現金流記錄</div>
    <div class="fg fg2">
      <div class="fgrp"><label>日期</label><input type="date" id="m_cf0" value="${today()}"></div>
      <div class="fgrp"><label>金額（正=收入 / 負=支出）</label><input type="number" id="m_cf1" placeholder="0"></div>
      <div class="fgrp"><label>類別</label><select id="m_cf2">${CF_CATS.map(c=>`<option>${c}</option>`).join('')}</select></div>
      <div class="fgrp"><label>說明</label><input type="text" id="m_cf3" placeholder="選填"></div>
    </div>
    <div class="mbtns"><button class="btn btn-s" onclick="closeMod()">取消</button><button class="btn btn-p" onclick="saveCF()">✓ 新增</button></div>`;}

  if(m.t==='addService'){c=`<div class="mttl">🌸 新增服務項目</div>
    <div class="fg fg2">
      <div class="fgrp"><label>服務名稱</label><input type="text" id="m_sv0" placeholder="如：基礎做臉"></div>
      <div class="fgrp"><label>類別</label><select id="m_sv1">${SERVICE_CATS.map(c=>`<option>${c}</option>`).join('')}</select></div>
      <div class="fgrp"><label>售價</label><input type="number" id="m_sv2" placeholder="0"></div>
      <div class="fgrp"><label>時長（分鐘）</label><input type="number" id="m_sv3" placeholder="60"></div>
    </div>
    <div class="mbtns"><button class="btn btn-s" onclick="closeMod()">取消</button><button class="btn btn-p" onclick="saveService()">✓ 新增</button></div>`;}

  if(m.t==='addProduct'){c=`<div class="mttl">🧴 新增產品</div>
    <div class="fg fg2">
      <div class="fgrp"><label>產品名稱</label><input type="text" id="m_pr0" placeholder="產品名稱"></div>
      <div class="fgrp"><label>類別</label><select id="m_pr1">${PRODUCT_CATS.map(c=>`<option>${c}</option>`).join('')}</select></div>
      <div class="fgrp"><label>售價</label><input type="number" id="m_pr2" placeholder="0"></div>
      <div class="fgrp"><label>成本</label><input type="number" id="m_pr3" placeholder="0"></div>
      <div class="fgrp"><label>初始庫存</label><input type="number" id="m_pr4" placeholder="0"></div>
      <div class="fgrp"><label>最低庫存警示</label><input type="number" id="m_pr5" placeholder="5"></div>
    </div>
    <div class="mbtns"><button class="btn btn-s" onclick="closeMod()">取消</button><button class="btn btn-p" onclick="saveProduct()">✓ 新增</button></div>`;}

  if(m.t==='addSupplier'){c=`<div class="mttl">🏭 新增供應商</div>
    <div class="fg fg2">
      <div class="fgrp"><label>供應商名稱</label><input type="text" id="m_su0" placeholder="廠商名稱"></div>
      <div class="fgrp"><label>聯絡人</label><input type="text" id="m_su1" placeholder="選填"></div>
      <div class="fgrp"><label>電話</label><input type="text" id="m_su2" placeholder="選填"></div>
      <div class="fgrp"><label>備註</label><input type="text" id="m_su3" placeholder="選填"></div>
    </div>
    <div class="mbtns"><button class="btn btn-s" onclick="closeMod()">取消</button><button class="btn btn-p" onclick="saveSupplier()">✓ 新增</button></div>`;}

  if(m.t==='addStaff'){c=`<div class="mttl">👩‍💼 新增員工</div>
    <div class="fg fg2">
      <div class="fgrp"><label>姓名</label><input type="text" id="m_sf0" placeholder="姓名"></div>
      <div class="fgrp"><label>職稱</label><input type="text" id="m_sf1" placeholder="如：美容師"></div>
      <div class="fgrp"><label>抽成比例（%）</label><input type="number" id="m_sf2" placeholder="30" min="0" max="100"></div>
      <div class="fgrp"><label>備註</label><input type="text" id="m_sf3" placeholder="選填"></div>
    </div>
    <div class="mbtns"><button class="btn btn-s" onclick="closeMod()">取消</button><button class="btn btn-p" onclick="saveStaff()">✓ 新增</button></div>`;}

  const bg=document.getElementById('modal-bg');
  if(bg)bg.className='modal-bg open';
  return c;
}

// ─── ACTIONS ───────────────────────────────────────────────
function bindEvents(){
  const mbg=$('modal-bg');
  if(mbg)mbg.onclick=e=>{if(e.target===mbg)closeMod();};
  const pd1=$('m_pd1');
  if(pd1)pd1.onchange=()=>{
    const opt=pd1.options[pd1.selectedIndex];
    const price=opt.dataset.price||0;
    if($('m_pd3'))$('m_pd3').value=price;
    calcPrdTotal();
  };
  const od1=$('m_od1');
  if(od1)od1.onchange=()=>{
    const opt=od1.options[od1.selectedIndex];
    const price=opt.dataset.price||0;
    if($('m_od3'))$('m_od3').value=price;
  };
}

window.openM=(t,x)=>{S.modal={t,x};const bg=$('modal-bg');const inn=$('modal-inner');if(bg&&inn){inn.innerHTML=renderModal();bg.className='modal-bg open';bindEvents();}};
window.closeMod=()=>{S.modal=null;const bg=$('modal-bg');if(bg)bg.className='modal-bg';};
window.setP=p=>{S.page=p;S.subTab=0;render();};
window.setT=i=>{S.subTab=i;const c=$('pg-content');if(c)c.innerHTML=renderPage();};
window.setMo=m=>{S.month=m;const c=$('pg-content');if(c)c.innerHTML=renderPage();};

window.calcPrdTotal=()=>{const q=parseFloat($('m_pd2')?.value||1)||1,p=parseFloat($('m_pd3')?.value||0)||0;if($('m_pd4'))$('m_pd4').value=q*p;};
window.calcPTotal=()=>{const q=parseFloat($('m_p3')?.value||1)||1,p=parseFloat($('m_p4')?.value||0)||0;if($('m_p5'))$('m_p5').value=q*p;};

// Sales
window.saveSale=async()=>{
  const date=$('m_s0')?.value,amount=parseFloat($('m_s3')?.value)||0;
  if(!date||!amount){alert('請填日期和金額');return;}
  const d=new Date(date);
  await addSub('sales',{date,type:$('m_s1')?.value,item:$('m_s2')?.value,amount,payMethod:$('m_s4')?.value,staff:$('m_s5')?.value,customer:$('m_s6')?.value,note:$('m_s7')?.value,month:d.getMonth()+1,year:d.getFullYear()});
  toast('✅ 收入已記帳');closeMod();
};
window.updSale=async(id)=>{
  const date=$('m_s0')?.value,amount=parseFloat($('m_s3')?.value)||0;
  if(!date||!amount){alert('請填日期和金額');return;}
  const d=new Date(date);
  await setSub('sales',id,{date,type:$('m_s1')?.value,item:$('m_s2')?.value,amount,payMethod:$('m_s4')?.value,staff:$('m_s5')?.value,customer:$('m_s6')?.value,note:$('m_s7')?.value,month:d.getMonth()+1,year:d.getFullYear()});
  toast('✅ 已更新');closeMod();
};
window.delSale=async id=>{if(!confirm('確認刪除？'))return;await delSub('sales',id);toast('已刪除');};

// Order
window.saveOrder=async(type)=>{
  if(type==='service'){
    const date=$('m_od0')?.value,amount=parseFloat($('m_od3')?.value)||0;
    if(!date||!amount){alert('請填日期和金額');return;}
    const svcSel=$('m_od1'),isCustom=svcSel?.value==='custom';
    const item=isCustom?$('m_od2')?.value:(svcSel?.options[svcSel.selectedIndex]?.text.split('（')[0]||'');
    const d=new Date(date);
    await addSub('sales',{date,type:'service',item,amount,payMethod:$('m_od6')?.value,staff:$('m_od5')?.value,customer:$('m_od4')?.value,note:$('m_od7')?.value,month:d.getMonth()+1,year:d.getFullYear()});
    toast(`✅ 服務收入 $${fmt(amount)} 已記帳`);
  }else{
    const date=$('m_pd0')?.value,qty=parseFloat($('m_pd2')?.value)||1,price=parseFloat($('m_pd3')?.value)||0,total=qty*price;
    if(!date||!total){alert('請填日期和金額');return;}
    const prdSel=$('m_pd1'),pid=prdSel?.value;
    const item=prdSel?.options[prdSel.selectedIndex]?.text.split('（')[0]||'';
    const d=new Date(date);
    await addSub('sales',{date,type:'product',item,amount:total,qty,unitPrice:price,payMethod:$('m_pd7')?.value,staff:$('m_pd6')?.value,customer:$('m_pd5')?.value,month:d.getMonth()+1,year:d.getFullYear()});
    if(pid&&pid!=='custom'){
      const inv=getInv(pid);
      const newQty=Math.max(0,toN(inv.qty)-qty);
      const newInv=[...S.data.inventory.filter(i=>i.productId!==pid),{productId:pid,qty:newQty,minQty:inv.minQty||5}];
      await saveMain({inventory:newInv});
    }
    toast(`✅ 產品銷售 $${fmt(total)} 已記帳，庫存已扣減`);
  }
  const c=$('pg-content');if(c)c.innerHTML=renderPage();
};

// Expenses
window.saveExp=async()=>{
  const date=$('m_e0')?.value,amount=parseFloat($('m_e2')?.value)||0;
  if(!date||!amount){alert('請填日期和金額');return;}
  const d=new Date(date);
  await addSub('expenses',{date,category:$('m_e1')?.value,amount,payMethod:$('m_e3')?.value,note:$('m_e4')?.value,month:d.getMonth()+1,year:d.getFullYear()});
  toast('✅ 支出已記帳');closeMod();
};
window.updExp=async(id)=>{
  const date=$('m_e0')?.value,amount=parseFloat($('m_e2')?.value)||0;
  const d=new Date(date);
  await setSub('expenses',id,{date,category:$('m_e1')?.value,amount,payMethod:$('m_e3')?.value,note:$('m_e4')?.value,month:d.getMonth()+1,year:d.getFullYear()});
  toast('✅ 已更新');closeMod();
};
window.delExp=async id=>{if(!confirm('確認刪除？'))return;await delSub('expenses',id);toast('已刪除');};

// Purchase
window.savePurchase=async()=>{
  const date=$('m_p0')?.value,qty=parseFloat($('m_p3')?.value)||0,unitCost=parseFloat($('m_p4')?.value)||0;
  let totalCost=parseFloat($('m_p5')?.value)||qty*unitCost;
  if(!date||!qty){alert('請填日期和數量');return;}
  const prdSel=$('m_p1'),pid=prdSel?.value;
  const isCustom=pid==='custom';
  const productName=isCustom?$('m_p2')?.value:(S.data.products||[]).find(p=>p.id===pid)?.name||'';
  const d=new Date(date);
  await addSub('purchases',{date,productId:isCustom?'':pid,productName,qty,unitCost,totalCost,supplier:$('m_p6')?.value,note:$('m_p7')?.value,month:d.getMonth()+1,year:d.getFullYear()});
  if(!isCustom&&pid){
    const inv=getInv(pid);
    const newQty=toN(inv.qty)+qty;
    const newInv=[...S.data.inventory.filter(i=>i.productId!==pid),{productId:pid,qty:newQty,minQty:inv.minQty||5}];
    await saveMain({inventory:newInv});
  }
  toast(`✅ 進貨完成，庫存已增加 ${qty} 個`);closeMod();
};
window.delPurchase=async id=>{if(!confirm('確認刪除？'))return;await delSub('purchases',id);toast('已刪除');};

// Inventory
window.adjInv=async(pid,delta)=>{
  const inv=getInv(pid);
  const newQty=Math.max(0,toN(inv.qty)+delta);
  const newInv=[...S.data.inventory.filter(i=>i.productId!==pid),{productId:pid,qty:newQty,minQty:inv.minQty||5}];
  await saveMain({inventory:newInv});
};
window.saveInv=async(pid)=>{
  const qty=parseFloat($('m_i0')?.value)||0,minQty=parseFloat($('m_i1')?.value)||5;
  const newInv=[...S.data.inventory.filter(i=>i.productId!==pid),{productId:pid,qty,minQty}];
  await saveMain({inventory:newInv});toast('✅ 庫存已更新');closeMod();
};
window.saveStockCheck=async()=>{
  const prods=S.data.products||[];
  const newInv=prods.map(p=>{
    const el=$('ck_'+p.id);
    const inv=getInv(p.id);
    const qty=el&&el.value!==''?parseFloat(el.value):inv.qty;
    return{productId:p.id,qty,minQty:inv.minQty||5};
  });
  await saveMain({inventory:newInv});toast('✅ 盤點完成');closeMod();
};

// Courses
window.saveCourse=async()=>{
  const name=$('m_c0')?.value;if(!name){alert('請填課程名稱');return;}
  const d=new Date($('m_c2')?.value||today());
  await addSub('courses',{name,category:$('m_c1')?.value,date:$('m_c2')?.value,instructor:$('m_c3')?.value,price:parseFloat($('m_c4')?.value)||0,note:$('m_c5')?.value,month:d.getMonth()+1,year:d.getFullYear()});
  toast('✅ 課程已新增');closeMod();
};
window.delCourse=async id=>{if(!confirm('確認刪除？'))return;await delSub('courses',id);toast('已刪除');};
window.saveStudent=async(courseId,courseName)=>{
  const name=$('m_st0')?.value;if(!name){alert('請填學員姓名');return;}
  const paid=$('m_st2')?.value==='1';
  const amount=parseFloat($('m_st1')?.value)||0;
  await addSub('courseStudents',{studentName:name,courseId,courseName,amount,paid,payMethod:paid?$('m_st3')?.value:'',month:S.month,year:yr()});
  if(paid){const d=new Date();await addSub('sales',{date:today(),type:'course',item:courseName,amount,payMethod:$('m_st3')?.value,staff:'',customer:name,month:d.getMonth()+1,year:d.getFullYear()});}
  toast('✅ 學員已新增');closeMod();
};
window.markPaid=async(id)=>{
  const s=(S.data.courseStudents||[]).find(x=>x.id===id);if(!s)return;
  await setSub('courseStudents',id,{...s,paid:true,payMethod:'現金'});
  await addSub('sales',{date:today(),type:'course',item:s.courseName,amount:s.amount,payMethod:'現金',staff:'',customer:s.studentName,month:s.month,year:s.year});
  toast('✅ 已收款並記入收入');
};
window.delStudent=async id=>{if(!confirm('確認刪除？'))return;await delSub('courseStudents',id);toast('已刪除');};

// CashFlow
window.saveCF=async()=>{
  const date=$('m_cf0')?.value,amount=parseFloat($('m_cf1')?.value)||0;
  if(!date||!amount){alert('請填日期和金額');return;}
  const d=new Date(date);
  await addSub('cashFlow',{date,amount,cat:$('m_cf2')?.value,note:$('m_cf3')?.value,month:d.getMonth()+1,year:d.getFullYear()});
  toast('✅ 已新增');closeMod();
};
window.delCF=async id=>{if(!confirm('確認刪除？'))return;await delSub('cashFlow',id);toast('已刪除');};

// Settings
window.saveService=async()=>{
  const name=$('m_sv0')?.value;if(!name){alert('請填服務名稱');return;}
  const svcs=[...(S.data.services||[]),{id:uid(),name,category:$('m_sv1')?.value,price:parseFloat($('m_sv2')?.value)||0,duration:parseFloat($('m_sv3')?.value)||60}];
  await saveMain({services:svcs});toast('✅ 已新增');closeMod();
};
window.delService=async id=>{if(!confirm('確認刪除？'))return;await saveMain({services:(S.data.services||[]).filter(s=>s.id!==id)});toast('已刪除');};

window.saveProduct=async()=>{
  const name=$('m_pr0')?.value;if(!name){alert('請填產品名稱');return;}
  const pid=uid();
  const qty=parseFloat($('m_pr4')?.value)||0;
  const minQty=parseFloat($('m_pr5')?.value)||5;
  const prods=[...(S.data.products||[]),{id:pid,name,category:$('m_pr1')?.value,price:parseFloat($('m_pr2')?.value)||0,cost:parseFloat($('m_pr3')?.value)||0}];
  const newInv=[...S.data.inventory.filter(i=>i.productId!==pid),{productId:pid,qty,minQty}];
  await saveMain({products:prods,inventory:newInv});toast('✅ 產品已新增');closeMod();
};
window.delProduct=async id=>{
  if(!confirm('確認刪除？'))return;
  await saveMain({products:(S.data.products||[]).filter(p=>p.id!==id),inventory:S.data.inventory.filter(i=>i.productId!==id)});toast('已刪除');
};

window.saveSupplier=async()=>{
  const name=$('m_su0')?.value;if(!name){alert('請填供應商名稱');return;}
  const sups=[...(S.data.suppliers||[]),{id:uid(),name,contact:$('m_su1')?.value,phone:$('m_su2')?.value,note:$('m_su3')?.value}];
  await saveMain({suppliers:sups});toast('✅ 供應商已新增');closeMod();
};
window.delSupplier=async id=>{if(!confirm('確認刪除？'))return;await saveMain({suppliers:(S.data.suppliers||[]).filter(s=>s.id!==id)});toast('已刪除');};

window.saveStaff=async()=>{
  const name=$('m_sf0')?.value;if(!name){alert('請填姓名');return;}
  const staff=[...(S.data.staff||[]),{id:uid(),name,title:$('m_sf1')?.value,commRate:parseFloat($('m_sf2')?.value)||0,note:$('m_sf3')?.value}];
  await saveMain({staff});toast('✅ 員工已新增');closeMod();
};
window.delStaff=async id=>{if(!confirm('確認刪除？'))return;await saveMain({staff:(S.data.staff||[]).filter(s=>s.id!==id)});toast('已刪除');};

// ── AUTH（密碼登入）─────────────────────────────────────
window.doLogin=async()=>{
  const pw=$('m_pw')?.value||'';
  const errEl=$('pw_err');
  if(pw!==SYSTEM_PASSWORD){
    if(errEl)errEl.textContent='❌ 密碼錯誤，請重新輸入';
    if($('m_pw')){$('m_pw').value='';$('m_pw').focus();}
    return;
  }
  // 密碼正確 → 標記已解鎖
  sessionStorage.setItem('mumei_unlock','1');
  S.auth=true;
  // 確保 Firebase 後端已連線
  if(!auth.currentUser){
    try{
      await signInWithEmailAndPassword(auth,BACKEND_EMAIL,BACKEND_PASSWORD);
    }catch(e){
      console.error('後端連線失敗',e);
    }
  }
  // 啟動同步並重新渲染
  if(S.sync==='idle')startSync();
  render();
  toast('✅ 歡迎進入沐蜜系統');
};

window.doOut=async()=>{
  sessionStorage.removeItem('mumei_unlock');
  S.auth=false;
  S.sync='idle';
  render();
};

// Export
window.exportXLSX=()=>{
  if(typeof XLSX==='undefined'){toast('⏳ 載入中請稍候');return;}
  const wb=XLSX.utils.book_new();
  const ds=new Date().toISOString().slice(0,10);
  const salesRows=[['日期','類型','項目','金額','付款','員工','客戶','備註'],...S.data.sales.map(s=>[s.date,s.type==='service'?'服務':s.type==='product'?'產品':'課程',s.item,s.amount,s.payMethod,s.staff,s.customer,s.note])];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(salesRows),'銷售記錄');
  const expRows=[['日期','類別','金額','付款','備註'],...S.data.expenses.map(e=>[e.date,e.category,e.amount,e.payMethod,e.note])];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(expRows),'支出記錄');
  const invRows=[['產品名稱','類別','現有庫存','最低庫存','售價','成本','庫存價值'],...(S.data.products||[]).map(p=>{const i=getInv(p.id);return[p.name,p.category,i.qty,i.minQty||5,p.price,p.cost||0,toN(i.qty)*toN(p.cost||0)];})];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(invRows),'庫存清單');
  const purRows=[['日期','產品','數量','單價','總成本','供應商'],...S.data.purchases.map(p=>[p.date,p.productName,p.qty,p.unitCost,p.totalCost,p.supplier])];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(purRows),'進貨記錄');
  const cfRows=[['日期','金額','類別','說明'],...S.data.cashFlow.map(c=>[c.date,c.amount,c.cat,c.note])];
  XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(cfRows),'現金流水帳');
  XLSX.writeFile(wb,`沐蜜財務_${ds}.xlsx`);
  toast('✅ Excel 已下載');
};

render();
</script>
<div id="app"></div>
</body>
</html>
