import { useState, useCallback, useMemo, useRef } from "react";
import * as XLSX from "xlsx";

/* ═══════════════════════════════════════════════════════════════════
   CONSTANTS
═══════════════════════════════════════════════════════════════════ */
const NUM_ROOMS = 12;
const ROOMS = Array.from({length:NUM_ROOMS},(_,i)=>`Room ${i+1}`);
const HOURS = [9,10,11,12,13,14,15,16,17];
const DAYS  = ["apr14","apr15"];
const DAY_LONG  = { apr14:"Tuesday, April 14th 2026",   apr15:"Wednesday, April 15th 2026" };
const DAY_SHORT = { apr14:"Tue Apr 14",                 apr15:"Wed Apr 15" };
const ALL_SLOTS = DAYS.flatMap(d=>HOURS.map(h=>`${d}-${h}`));
const DINNER_HOUR_LABEL = "7:00 PM";
const DINNERS_INIT = [
  {id:"din-apr14", day:"apr14", name:"Conference Dinner", restaurant:"", address:""},
  {id:"din-apr15", day:"apr15", name:"Conference Dinner", restaurant:"", address:""},
];
const slotDay  = id => id.split("-")[0];
const slotHour = id => parseInt(id.split("-")[1]);
const hourLabel = h => h===12?"12:00 PM":h>12?`${h-12}:00 PM`:`${h}:00 AM`;
const slotLabel = id => hourLabel(slotHour(id));

function parseAvail(raw){
  if(!raw) return [];
  const ids=new Set();
  for(const p of raw.toLowerCase().split(";").map(s=>s.trim()).filter(Boolean)){
    if(p.includes("all")){ALL_SLOTS.forEach(s=>ids.add(s));continue;}
    const day=p.includes("apr - 14")||p.includes("apr 14")?"apr14":p.includes("apr - 15")||p.includes("apr 15")?"apr15":null;
    const period=p.includes("morning")?"morning":p.includes("afternoon")?"afternoon":null;
    if(!day) continue;
    HOURS.forEach(h=>{
      const m=h<=12;
      if(!period||(period==="morning"&&m)||(period==="afternoon"&&!m)) ids.add(`${day}-${h}`);
    });
  }
  return ALL_SLOTS.filter(s=>ids.has(s));
}

/* ═══════════════════════════════════════════════════════════════════
   COMPANIES MASTER LIST
═══════════════════════════════════════════════════════════════════ */
const COMPANIES_INIT = [
  {id:"BMA",  name:"Banco Macro",             ticker:"BMA",   sector:"Financials"},
  {id:"BBAR", name:"BBVA Argentina",           ticker:"BBAR",  sector:"Financials"},
  {id:"GGAL", name:"Grupo Fin. Galicia",       ticker:"GGAL",  sector:"Financials"},
  {id:"SUPV", name:"Grupo Supervielle",        ticker:"SUPV",  sector:"Financials"},
  {id:"BYMA", name:"BYMA",                     ticker:"BYMA",  sector:"Financials"},
  {id:"A3",   name:"A3 Mercados",              ticker:"A3",    sector:"Financials"},
  {id:"PAM",  name:"Pampa Energía",            ticker:"PAM",   sector:"Energy"},
  {id:"YPF",  name:"YPF",                      ticker:"YPF",   sector:"Energy"},
  {id:"YPFL", name:"YPF Luz",                  ticker:"YPFL",  sector:"Energy"},
  {id:"VIST", name:"Vista Energy",             ticker:"VIST",  sector:"Energy"},
  {id:"CEPU", name:"Central Puerto",           ticker:"CEPU",  sector:"Energy"},
  {id:"TGS",  name:"TGS",                      ticker:"TGS",   sector:"Energy"},
  {id:"GNNEIA",name:"Genneia",                 ticker:"GNNEIA",sector:"Energy"},
  {id:"MSU",  name:"MSU Energy",               ticker:"MSU",   sector:"Energy"},
  {id:"CAAP", name:"Corporación América",      ticker:"CAAP",  sector:"Infra"},
  {id:"IRS",  name:"IRSA / Cresud",            ticker:"IRS",   sector:"Real Estate"},
  {id:"LOMA", name:"Loma Negra",               ticker:"LOMA",  sector:"Infra"},
  {id:"TEO",  name:"Telecom Argentina",        ticker:"TEO",   sector:"TMT"},
];
const CO_MAP = {
  "banco macro (bma)":"BMA","banco macro":"BMA",
  "bbva argentina (bbar)":"BBAR","bbva argentina":"BBAR",
  "grupo financiero galicia (ggal)":"GGAL","grupo financiero galicia":"GGAL",
  "grupo supervielle (supv)":"SUPV","grupo supervielle":"SUPV",
  "byma (bolsas y mercados argentinos)":"BYMA","byma":"BYMA",
  "a3 mercados":"A3","a3":"A3",
  "pampa energía (pam)":"PAM","pampa energia (pam)":"PAM","pampa energía":"PAM","pampa energia":"PAM",
  "ypf":"YPF","ypf luz":"YPFL",
  "vista (vist)":"VIST","vista energy (vist)":"VIST","vista":"VIST",
  "central puerto (cepu)":"CEPU","central puerto":"CEPU",
  "transportadora de gas del sur (tgs)":"TGS","transportadora de gas del sur":"TGS","tgs":"TGS",
  "genneia (gnneia)":"GNNEIA","genneia":"GNNEIA",
  "msu energy":"MSU","msu":"MSU",
  "corporación américa (caap)":"CAAP","corporacion america (caap)":"CAAP","corporación america (caap)":"CAAP",
  "irsa (irs) - cresud (cresy)":"IRS","irsa (irs)":"IRS","cresud (cresy)":"IRS","irsa":"IRS",
  "loma negra (loma)":"LOMA","loma negra":"LOMA",
  "telecom argentina (teo)":"TEO","telecom argentina":"TEO",
};
const resolveCo = raw => CO_MAP[raw.trim().toLowerCase()]||null;
const SEC_CLR = {Financials:"#4a8fd4",Energy:"#d4854a",Infra:"#4aaf7a","Real Estate":"#c9a84c",TMT:"#9b6fd4"};

/* ═══════════════════════════════════════════════════════════════════
   ZIP BUILDER (pure JS, no deps)
═══════════════════════════════════════════════════════════════════ */
const CRC_TBL=(()=>{const t=new Uint32Array(256);for(let i=0;i<256;i++){let c=i;for(let j=0;j<8;j++)c=(c&1)?0xEDB88320^(c>>>1):c>>>1;t[i]=c;}return t;})();
function crc32(b){let c=0xFFFFFFFF;for(let i=0;i<b.length;i++)c=(c>>>8)^CRC_TBL[(c^b[i])&0xFF];return(c^0xFFFFFFFF)>>>0;}
function u16(n){return[n&0xFF,(n>>8)&0xFF];}
function u32(n){return[n&0xFF,(n>>8)&0xFF,(n>>16)&0xFF,(n>>24)&0xFF];}
function cat(...arrs){const total=arrs.reduce((s,a)=>s+a.length,0);const out=new Uint8Array(total);let i=0;for(const a of arrs){out.set(a,i);i+=a.length;}return out;}

function buildZip(files){
  // files: [{name:string, data:string|Uint8Array}]
  const enc=new TextEncoder();
  const parts=[];const cdirs=[];let offset=0;
  for(const f of files){
    const name=enc.encode(f.name);
    const data=f.data instanceof Uint8Array?f.data:enc.encode(f.data);
    const crc=crc32(data);const sz=data.length;
    const local=new Uint8Array([0x50,0x4B,0x03,0x04,20,0,0,0,0,0,0,0,0,0,...u32(crc),...u32(sz),...u32(sz),...u16(name.length),0,0,...name,...data]);
    const cdir=new Uint8Array([0x50,0x4B,0x01,0x02,20,0,20,0,0,0,0,0,0,0,0,0,...u32(crc),...u32(sz),...u32(sz),...u16(name.length),0,0,0,0,0,0,0,0,0,0,0,0,...u32(offset),...name]);
    parts.push(local);cdirs.push(cdir);offset+=local.length;
  }
  const cdOff=offset;const cdData=cat(...cdirs);
  const eocd=new Uint8Array([0x50,0x4B,0x05,0x06,0,0,0,0,...u16(files.length),...u16(files.length),...u32(cdData.length),...u32(cdOff),0,0]);
  return cat(...parts,cdData,eocd).buffer;
}

function downloadBlob(name,content,type){
  const blob=new Blob([content],{type});
  const url=URL.createObjectURL(blob);
  const a=document.createElement("a");a.href=url;a.download=name;a.click();
  setTimeout(()=>URL.revokeObjectURL(url),5000);
}

/* ═══════════════════════════════════════════════════════════════════
   WORD HTML GENERATOR  (Latin Securities style)
═══════════════════════════════════════════════════════════════════ */
const esc=s=>String(s||"").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");

function buildWordHTML(entityName, entitySub, sections, footerNote=""){
  // sections: [{dayLabel, rows: [{time, col1, col1b, col2, col3}], headerCols:[...]}]
  // One table per day with page breaks between them
  const tableSections = sections.map((sec,si)=>{
    const nCols=sec.headerCols.length;
    const rows=sec.rows.map((r,i)=>{
      const bg=i%2===0?"#f3f5fb":"#fff";
      const room=r.col4!==undefined?r.col4:r.col3;
      const typeColor=r.col3==="1x1"?"#2a7a4a":r.col3==="Group"?"#7b52a8":r.col3==="Dinner"?"#b45309":"#444";
      const typeCell=r.col4!==undefined
        ?`<td style="padding:8px 10px;vertical-align:top;font-size:9pt;color:${typeColor};font-weight:700;white-space:nowrap">${esc(r.col3)}</td>`
        :"";
      const col1display=r.isGroup
        ?r.col1.split("\n").map(n=>`<strong style="font-size:11pt">${esc(n)}</strong>`).join("<br/>")
        :`<strong style="font-size:11pt">${esc(r.col1)}</strong>${r.col1b?`<br/><span style="font-size:9pt;color:#666">${esc(r.col1b)}</span>`:""}${r.col1c?`<br/><span style="font-size:9pt;color:#555;font-style:italic">${esc(r.col1c)}</span>`:""}`;
      return `<tr style="background:${bg}">
        <td style="font-weight:bold;color:#1e3f87;font-size:11pt;white-space:nowrap;padding:8px 10px;vertical-align:top">${esc(r.time)}</td>
        <td style="padding:8px 10px;vertical-align:top">${col1display}</td>
        <td style="padding:8px 10px;vertical-align:top;font-size:10pt">${esc(r.col2)}</td>
        ${typeCell}
        <td style="padding:8px 10px;vertical-align:top;font-style:italic;font-size:10pt">${esc(room)}</td>
      </tr>`;
    }).join("");
    const pageBreak=si>0?`<p style="page-break-before:always;margin:0;font-size:1pt">&nbsp;</p>`:"";
    return `${pageBreak}
    <table style="width:100%;border-collapse:collapse;margin-bottom:20px;border:1px solid #c8cdd8">
    <tr><td colspan="${nCols}" style="background:#1e3f87;color:#fff;font-weight:bold;padding:7px 12px;font-size:11pt;letter-spacing:0.04em">${esc(sec.dayLabel)}</td></tr>
    <tr style="background:#2d5cb8">${sec.headerCols.map(h=>`<th style="color:#fff;padding:7px 10px;text-align:left;font-size:9.5pt;letter-spacing:0.05em;text-transform:uppercase">${esc(h)}</th>`).join("")}</tr>
    ${rows}
    </table>`;
  }).join("");

  return `<!DOCTYPE html>
<html xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="utf-8">
<meta name="ProgId" content="Word.Document">
<meta name="Originator" content="Latin Securities">
<title>${esc(entityName)} — Argentina in New York 2026</title>
<!--[if gte mso 9]><xml><w:WordDocument><w:View>Print</w:View><w:Zoom>90</w:Zoom><w:DoNotPromoteQF/></w:WordDocument></xml><![endif]-->
<style>
@page { size: 8.5in 11in; margin: 1in 1in; mso-page-orientation: portrait; }
body { font-family: Calibri, Arial, sans-serif; font-size: 11pt; color: #1a1a1a; margin: 0; padding: 0; }
.ls-header { display: table; width: 100%; border-bottom: 3pt solid #1e3f87; padding-bottom: 10px; margin-bottom: 18px; }
.ls-logo-cell { display: table-cell; vertical-align: middle; }
.ls-logo { background: #1e3f87; color: #fff; font-weight: 700; font-size: 13pt; letter-spacing: 1px; padding: 8px 14px; border-radius: 3px; line-height: 1.3; }
.ls-title-cell { display: table-cell; vertical-align: middle; text-align: right; padding-left: 20px; }
.ls-event-title { font-size: 13pt; font-weight: 700; color: #1e3f87; }
.ls-event-sub { font-size: 9pt; color: #666; margin-top: 2px; }
h1 { font-size: 18pt; font-weight: 700; color: #1e3f87; margin: 0 0 4px; }
h2 { font-size: 10.5pt; font-weight: 400; color: #666; margin: 0 0 18px; border-bottom: 1px solid #dde0ea; padding-bottom: 8px; }
table { width: 100%; border-collapse: collapse; margin-bottom: 20px; border: 1px solid #c8cdd8; }
th,td { padding: 0; }
.footer-note { font-size: 9pt; color: #888; margin-top: 12px; border-top: 1px solid #dde0ea; padding-top: 8px; }
</style>
</head>
<body>
<div class="ls-header">
  <div class="ls-logo-cell"><div class="ls-logo">LATIN<br>SECURITIES</div></div>
  <div class="ls-title-cell">
    <div class="ls-event-title">Argentina in New York 2026</div>
    <div class="ls-event-sub">Investment Conference &middot; April 14&ndash;15, 2026</div>
  </div>
</div>
<h1>${esc(entityName)}</h1>
<h2>${esc(entitySub)}</h2>
${tableSections}
${footerNote?`<div class="footer-note">${esc(footerNote)}</div>`:""}
</body>
</html>`;
}

function buildPrintHTML(entities, options={}){
  // One page per entity-day: header + entity title + day table, repeated for each day
  function renderRow(r,i){
    const bg=i%2===0?"#f3f5fb":"#fff";
    const typeColor=r.col3==="1x1"?"#2a7a4a":r.col3==="Group"?"#7b52a8":r.col3==="Dinner"?"#b45309":"#444";
    const typeWeight=r.col3==="1x1"||r.col3==="Group"||r.col3==="Dinner"?"700":"400";
    const col1html=r.isGroup
      ?("<strong>"+r.col1.split("\n").map(n=>esc(n)).join("</strong><br/><strong>")+"</strong>")
      :("<strong>"+esc(r.col1)+"</strong>"+(r.col1b?"<br/><small>"+esc(r.col1b)+"</small>":"")+(r.col1c?"<br/><em>"+esc(r.col1c)+"</em>":""));
    const room=r.col4!==undefined?r.col4:r.col3;
    const typeCell=r.col4!==undefined
      ?`<td style="font-size:9pt;white-space:nowrap;color:${typeColor};font-weight:${typeWeight}">${esc(r.col3)}</td>`
      :"";
    return `<tr style="background:${bg}">
      <td class="t-time">${esc(r.time)}</td>
      <td>${col1html}</td>
      <td>${esc(r.col2)}</td>
      ${typeCell}
      <td class="t-room">${esc(room)}</td>
    </tr>`;
  }

  function renderHeader(){
    return `<div class="ls-hdr">
      <div class="ls-logo">LATIN<br>SECURITIES</div>
      <div class="ev-info">
        <div style="font-size:13pt;font-weight:700;color:#1e3f87">Argentina in New York 2026</div>
        <div style="font-size:9pt;color:#666">Investment Conference &middot; April 14&ndash;15, 2026</div>
      </div>
    </div>`;
  }

  // Flatten: one page per entity × section-day
  const pages = [];
  entities.forEach(e=>{
    e.sections.forEach((sec,si)=>{
      const isLast = si===e.sections.length-1;
      const attendeesHtml = isLast && e.attendees?.length
        ? `<div class="attendees"><strong>Company Representatives:</strong> ${e.attendees.map(a=>`${esc(a.name)}${a.title?` (${esc(a.title)})`:""}`).join(" &bull; ")}</div>`
        : "";
      const nCols = sec.headerCols.length;
      pages.push(`<div class="page">
        ${renderHeader()}
        <h1>${esc(e.name)}</h1>
        <h2>${esc(e.sub)}</h2>
        <table>
          <tr><td colspan="${nCols}" class="day-hdr">${esc(sec.dayLabel)}</td></tr>
          <tr class="tbl-hdr">${sec.headerCols.map(h=>`<th>${esc(h)}</th>`).join("")}</tr>
          ${sec.rows.map((r,i)=>renderRow(r,i)).join("")}
        </table>
        ${attendeesHtml}
      </div>`);
    });
  });

  return `<!DOCTYPE html><html><head><meta charset="utf-8">
<title>Argentina in New York 2026 — Schedule</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#1a1a1a;background:#fff;padding:20px 28px}
.page{max-width:820px;margin:0 auto;padding-bottom:24px}
.ls-hdr{display:flex;align-items:center;justify-content:space-between;border-bottom:3px solid #1e3f87;padding-bottom:10px;margin-bottom:16px}
.ls-logo{background:#1e3f87;color:#fff;font-weight:700;font-size:13pt;padding:7px 13px;border-radius:3px;line-height:1.3;letter-spacing:.03em}
h1{font-size:18pt;font-weight:700;color:#1e3f87;margin:0 0 4px}
h2{font-size:10.5pt;color:#666;margin:0 0 16px;border-bottom:1px solid #dde;padding-bottom:8px}
table{width:100%;border-collapse:collapse;margin-bottom:16px}
.day-hdr{background:#1e3f87;color:#fff;font-weight:700;padding:6px 12px;font-size:10.5pt;letter-spacing:.04em}
.tbl-hdr th{background:#2d5cb8;color:#fff;padding:6px 10px;text-align:left;font-size:9.5pt;letter-spacing:.05em;text-transform:uppercase}
td{padding:7px 10px;border-bottom:1px solid #dde;vertical-align:top}
.t-time{font-weight:700;color:#1e3f87;white-space:nowrap;width:72px}
.t-room{font-style:italic;width:80px}
small{font-size:9pt;color:#666}em{font-size:9pt;color:#555}
.ev-info{text-align:right}
.attendees{font-size:9.5pt;color:#555;margin-top:8px;padding-top:8px;border-top:1px dashed #dde}
@media print{
  body{padding:0}
  .page{page-break-before:always;padding:12px 20px}
  .page:first-child{page-break-before:avoid}
  .day-hdr{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  .tbl-hdr th{-webkit-print-color-adjust:exact;print-color-adjust:exact}
  tr:nth-child(even) td{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
</style></head><body>${pages.join("")}</body></html>`;
}

/* ═══════════════════════════════════════════════════════════════════
   ENTITY → EXPORT DATA  helpers
═══════════════════════════════════════════════════════════════════ */
function companyToExportData(co, meetings, investors){
  const cms = meetings.filter(m=>m.coId===co.id)
    .sort((a,b)=>ALL_SLOTS.indexOf(a.slotId)-ALL_SLOTS.indexOf(b.slotId));
  if(!cms.length) return null;
  const dayGroups = {};
  cms.forEach(m=>{const d=slotDay(m.slotId);if(!dayGroups[d])dayGroups[d]=[];dayGroups[d].push(m);});
  const totalMeetings = cms.length;
  const sections = DAYS.filter(d=>dayGroups[d]).map(d=>({
    dayLabel: DAY_LONG[d],
    headerCols:["Time","Investor","Fund / Firm","Type","Room"],
    rows: dayGroups[d].map(m=>{
      const invs = (m.invIds||[]).map(id=>investors.find(i=>i.id===id)).filter(Boolean);
      const isGroup = invs.length > 1;
      const allSameFund = isGroup && invs.every(i=>i.fund && i.fund===invs[0].fund);
      const mType = !isGroup || allSameFund ? "1x1" : "Group";
      // For group: show each name + position stacked; for 1x1: single name
      const col1 = isGroup
        ? invs.map(i=>i.name).join("\n")
        : (invs[0]?.name||"");
      const col1b = isGroup ? null : (invs[0]?.position||null);
      return {
        time: hourLabel(slotHour(m.slotId)),
        col1, col1b, col1c: null,
        col2: [...new Set(invs.map(i=>i.fund).filter(Boolean))].join(", "),
        col3: mType,
        col4: m.room,
        isGroup,
      };
    })
  }));
  return {
    name: `${co.name} (${co.ticker})`,
    sub: `${co.sector} · ${totalMeetings} meeting${totalMeetings!==1?"s":""}`,
    sections,
    attendees: co.attendees||[],
  };
}

function investorToExportData(inv, meetings, companies, dinners=[]){
  const cms = meetings.filter(m=>(m.invIds||[]).includes(inv.id))
    .sort((a,b)=>ALL_SLOTS.indexOf(a.slotId)-ALL_SLOTS.indexOf(b.slotId));
  const dayGroups = {};
  cms.forEach(m=>{const d=slotDay(m.slotId);if(!dayGroups[d])dayGroups[d]=[];dayGroups[d].push(m);});
  // Include days that have dinners even if no meetings
  const activeDays = DAYS.filter(d=>dayGroups[d]||dinners.some(din=>din.day===d));
  if(!activeDays.length) return null;
  const sections = activeDays.map(d=>({
    dayLabel: DAY_LONG[d],
    headerCols:["Time","Company","Sector","Type","Room"],
    rows: [
      ...(dayGroups[d]||[]).map(m=>{
        const co=companies.find(c=>c.id===m.coId);
        const mType=(m.invIds||[]).length>1?"Group":"1x1";
        return {time:hourLabel(slotHour(m.slotId)),col1:co?.name||m.coId,col1b:co?.ticker,col1c:null,
          col2:co?.sector||"",col3:mType,col4:m.room};
      }),
      ...dinners.filter(din=>din.day===d).map(din=>({
        time:DINNER_HOUR_LABEL,
        col1:din.name||"Conference Dinner",
        col1b:din.restaurant||null,
        col1c:null,
        col2:"",
        col3:"Dinner",
        col4:din.address||"",
        isDinner:true,
      })),
    ]
  }));
  return {
    name: inv.name,
    sub: [inv.position,inv.fund,inv.aum].filter(Boolean).join(" · "),
    sections,
  };
}

/* ═══════════════════════════════════════════════════════════════════
   SCHEDULING ALGORITHM
   • invIds[] array supports group meetings
   • Respects inv.blockedSlots
   • Same-fund investors → grouped if fundGrouping[fund]===true
═══════════════════════════════════════════════════════════════════ */
function buildRoomMap(investors){
  const demand={};COMPANIES_INIT.forEach(c=>{demand[c.id]=0;});
  investors.forEach(inv=>(inv.companies||[]).forEach(cid=>{demand[cid]=(demand[cid]||0)+1;}));
  const sorted=[...COMPANIES_INIT].sort((a,b)=>demand[b.id]-demand[a.id]);
  const map={};
  sorted.slice(0,NUM_ROOMS).forEach((c,i)=>{map[c.id]=ROOMS[i];});
  return map;
}

function effectiveSlots(inv){
  return (inv.slots||[]).filter(s=>!(inv.blockedSlots||[]).includes(s));
}

function runSchedule(investors, fundGrouping){
  const fixedRoom=buildRoomMap(investors);

  // Build grouped request list
  // For same-fund investors requesting same co → one combined request
  const fundMap={};// fund→[invId]
  investors.forEach(inv=>{if(inv.fund){if(!fundMap[inv.fund])fundMap[inv.fund]=[];fundMap[inv.fund].push(inv.id);}});

  // Requests: each is {invIds:[], coId}
  const processed=new Set();
  const reqs=[];
  investors.forEach(inv=>{
    (inv.companies||[]).forEach(coId=>{
      const key=`${inv.id}::${coId}`;
      if(processed.has(key)) return;
      processed.add(key);
      const fundmates=(fundMap[inv.fund]||[]).filter(id=>id!==inv.id&&investors.find(i=>i.id===id)?.companies?.includes(coId));
      const grouped=inv.fund&&fundmates.length>0&&(fundGrouping[inv.fund]!==false);
      if(grouped){
        // Mark all fundmates as processed
        fundmates.forEach(id=>processed.add(`${id}::${coId}`));
        reqs.push({invIds:[inv.id,...fundmates],coId});
      } else {
        reqs.push({invIds:[inv.id],coId});
      }
    });
  });

  // Sort: most-constrained first (fewest shared available slots)
  reqs.sort((a,b)=>{
    const slotsA=a.invIds.reduce((s,id)=>{const inv=investors.find(i=>i.id===id);return s.filter(sl=>effectiveSlots(inv).includes(sl));},ALL_SLOTS);
    const slotsB=b.invIds.reduce((s,id)=>{const inv=investors.find(i=>i.id===id);return s.filter(sl=>effectiveSlots(inv).includes(sl));},ALL_SLOTS);
    return slotsA.length-slotsB.length;
  });

  const invBusy={};investors.forEach(i=>{invBusy[i.id]=new Set();});
  const coBusy={};COMPANIES_INIT.forEach(c=>{coBusy[c.id]=new Set();});
  const roomBusy={};
  const coLastRoom={};

  const meetings=[],unscheduled=[];

  for(const req of reqs){
    // Available slots = intersection of all investors' effective slots
    let shared=ALL_SLOTS;
    for(const id of req.invIds){
      const inv=investors.find(i=>i.id===id);
      shared=shared.filter(s=>effectiveSlots(inv).includes(s)&&!invBusy[id].has(s));
    }
    shared=shared.filter(s=>!coBusy[req.coId].has(s));

    let placed=false;
    for(const slotId of shared){
      const preferred=fixedRoom[req.coId]||coLastRoom[req.coId];
      let room=null;
      if(preferred&&!roomBusy[`${preferred}::${slotId}`]) room=preferred;
      else room=ROOMS.find(r=>!roomBusy[`${r}::${slotId}`])||null;
      if(room){
        const id=`m-${Date.now()}-${Math.random().toString(36).slice(2,5)}`;
        meetings.push({id,invIds:req.invIds,coId:req.coId,slotId,room});
        req.invIds.forEach(invId=>invBusy[invId].add(slotId));
        coBusy[req.coId].add(slotId);
        roomBusy[`${room}::${slotId}`]=true;
        coLastRoom[req.coId]=room;
        placed=true;break;
      }
    }
    if(!placed) unscheduled.push(req);
  }
  return{meetings,unscheduled,fixedRoom};
}

/* ═══════════════════════════════════════════════════════════════════
   CSS
═══════════════════════════════════════════════════════════════════ */
const CSS=`
@import url('https://fonts.googleapis.com/css2?family=Playfair+Display:wght@400;700&family=IBM+Plex+Mono:wght@400;500&family=Lora:wght@400;600&display=swap');
*{box-sizing:border-box;margin:0;padding:0}
:root{--ink:#06101f;--ink2:#0c1c34;--ink3:#12253d;--gold:#c9a84c;--gold2:#e5c76a;--cream:#ede7d4;--txt:#c0b59e;--dim:#566778;--red:#d64444;--grn:#4aaf7a;--blu:#4a8fd4;--pur:#9b6fd4}
html,body{background:var(--ink)}
.app{min-height:100vh;background:var(--ink);color:var(--txt);font-family:'Lora',Georgia,serif}
.hdr{background:var(--ink2);border-bottom:1px solid rgba(201,168,76,.14);padding:0 26px;display:flex;align-items:center;position:sticky;top:0;z-index:300;box-shadow:0 2px 24px rgba(0,0,0,.5)}
.brand{padding:12px 0;margin-right:auto}
.brand h1{font-family:'Playfair Display',serif;font-size:15.5px;color:var(--gold);letter-spacing:.03em}
.brand p{font-size:8.5px;color:var(--dim);letter-spacing:.14em;text-transform:uppercase;margin-top:2px}
.nav{display:flex}
.ntab{padding:0 15px;height:56px;display:flex;align-items:center;font-size:10px;letter-spacing:.07em;color:var(--dim);cursor:pointer;border:none;border-bottom:2px solid transparent;background:none;font-family:'IBM Plex Mono',monospace;text-transform:uppercase;transition:all .15s;gap:5px;white-space:nowrap}
.ntab:hover{color:var(--txt)}.ntab.on{color:var(--gold);border-bottom-color:var(--gold);background:rgba(201,168,76,.04)}
.body{padding:24px 26px;max-width:1700px;margin:0 auto}
.pg-h{font-family:'Playfair Display',serif;font-size:21px;color:var(--cream);margin-bottom:3px}
.pg-s{color:var(--dim);font-size:13px;margin-bottom:20px}
.card{background:var(--ink2);border:1px solid rgba(201,168,76,.1);border-radius:8px;padding:17px 21px;margin-bottom:13px}
.card-t{font-family:'Playfair Display',serif;font-size:13px;color:var(--gold);margin-bottom:11px;display:flex;align-items:center;gap:7px}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:13px}
.g3{display:grid;grid-template-columns:1fr 1fr 1fr;gap:13px}
.inp{background:var(--ink3);border:1px solid rgba(201,168,76,.14);border-radius:6px;padding:7px 11px;color:var(--txt);font-size:12.5px;width:100%;font-family:'Lora',serif;transition:border-color .15s}
.inp:focus{outline:none;border-color:var(--gold)}
.sel{background:var(--ink3);border:1px solid rgba(201,168,76,.14);border-radius:6px;padding:7px 11px;color:var(--txt);font-size:12.5px;width:100%;font-family:'Lora',serif;cursor:pointer}
.btn{padding:7px 15px;border-radius:6px;font-size:10.5px;cursor:pointer;font-family:'IBM Plex Mono',monospace;letter-spacing:.04em;transition:all .15s;border:none;display:inline-flex;align-items:center;gap:5px}
.bg{background:var(--gold);color:var(--ink);font-weight:700}.bg:hover{background:var(--gold2)}
.bo{background:transparent;color:var(--gold);border:1px solid rgba(201,168,76,.28)}.bo:hover{border-color:var(--gold);background:rgba(201,168,76,.06)}
.bd{background:rgba(214,68,68,.1);color:var(--red);border:1px solid rgba(214,68,68,.24)}.bd:hover{background:rgba(214,68,68,.2)}
.bs{padding:4px 10px;font-size:10px}
.tbl{width:100%;border-collapse:collapse}
.tbl th{background:rgba(201,168,76,.07);color:var(--gold);font-size:9px;letter-spacing:.08em;text-transform:uppercase;padding:7px 10px;text-align:left;font-family:'IBM Plex Mono',monospace;border-bottom:1px solid rgba(201,168,76,.11)}
.tbl td{padding:7px 10px;border-bottom:1px solid rgba(255,255,255,.04);font-size:12px;vertical-align:middle}
.tbl tr:hover td{background:rgba(201,168,76,.03)}
.bdg{display:inline-block;padding:1px 6px;border-radius:3px;font-size:10px;font-family:'IBM Plex Mono',monospace}
.bg-g{background:rgba(201,168,76,.12);color:var(--gold)}.bg-r{background:rgba(214,68,68,.12);color:var(--red)}.bg-b{background:rgba(74,143,212,.12);color:var(--blu)}.bg-grn{background:rgba(74,175,122,.12);color:var(--grn)}
.stats{display:flex;gap:10px;margin-bottom:18px;flex-wrap:wrap}
.stat{background:var(--ink2);border:1px solid rgba(201,168,76,.1);border-radius:7px;padding:11px 15px;flex:1;min-width:90px}
.sv{font-family:'Playfair Display',serif;font-size:26px;color:var(--gold);line-height:1}
.sl{font-size:9px;color:var(--dim);text-transform:uppercase;letter-spacing:.09em;margin-top:3px;font-family:'IBM Plex Mono',monospace}
.upz{border:2px dashed rgba(201,168,76,.19);border-radius:8px;padding:38px 20px;text-align:center;cursor:pointer;transition:all .2s}
.upz:hover{border-color:var(--gold);background:rgba(201,168,76,.03)}
.alert{padding:9px 12px;border-radius:6px;font-size:12px;margin-bottom:10px}
.aw{background:rgba(214,68,68,.07);border:1px solid rgba(214,68,68,.2);color:#e8a0a0}
.ai{background:rgba(74,143,212,.07);border:1px solid rgba(74,143,212,.2);color:#a0c4e8}
.ag{background:rgba(74,175,122,.07);border:1px solid rgba(74,175,122,.2);color:#96d4b4}
.tag{display:inline-flex;padding:2px 6px;border-radius:12px;font-size:10px;background:rgba(201,168,76,.07);color:var(--gold2);border:1px solid rgba(201,168,76,.12);margin:2px 2px 0 0}
.flex{display:flex;gap:8px;align-items:center;flex-wrap:wrap}
.lbl{font-size:9px;color:var(--dim);text-transform:uppercase;letter-spacing:.08em;font-family:'IBM Plex Mono',monospace;margin-bottom:3px}
/* entity rows */
.ent-row{background:var(--ink2);border:1px solid rgba(201,168,76,.07);border-radius:7px;padding:11px 14px;margin-bottom:5px;display:flex;align-items:flex-start;gap:10px;cursor:pointer;transition:all .15s}
.ent-row:hover{border-color:rgba(201,168,76,.22);background:rgba(201,168,76,.03)}
/* slot grid */
.slot-grid{display:grid;gap:3px;margin-top:4px}
.slot-cell{padding:3px 2px;text-align:center;border-radius:3px;cursor:pointer;font-size:9.5px;font-family:'IBM Plex Mono',monospace;transition:all .12s;user-select:none}
.slot-avail{background:rgba(74,175,122,.13);color:var(--grn);border:1px solid rgba(74,175,122,.2)}
.slot-avail:hover{background:rgba(74,175,122,.22)}
.slot-blocked{background:rgba(214,68,68,.13);color:var(--red);border:1px solid rgba(214,68,68,.2);text-decoration:line-through}
.slot-blocked:hover{background:rgba(214,68,68,.22)}
.slot-na{background:rgba(255,255,255,.03);color:rgba(255,255,255,.13);border:1px solid transparent;cursor:default}
/* grid schedule */
.grid-wrap{overflow-x:auto}
.grid-tbl{border-collapse:collapse;table-layout:fixed}
.grid-tbl .th-time{width:72px;background:rgba(201,168,76,.07);font-family:'IBM Plex Mono',monospace;font-size:9px;color:var(--gold);padding:7px 8px;border-bottom:1px solid rgba(201,168,76,.11);text-align:right;text-transform:uppercase;letter-spacing:.06em;position:sticky;left:0;z-index:10}
.grid-tbl .th-sect{font-size:7.5px;letter-spacing:.08em;text-transform:uppercase;padding:3px 6px;text-align:center}
.grid-tbl .th-co{background:var(--ink2);font-family:'IBM Plex Mono',monospace;font-size:9px;color:var(--txt);padding:5px 7px;border-bottom:2px solid;text-align:center;min-width:115px;white-space:nowrap}
.grid-tbl .td-time{background:rgba(201,168,76,.05);font-family:'IBM Plex Mono',monospace;font-size:10px;color:var(--gold);padding:4px 8px;border-right:2px solid rgba(201,168,76,.14);border-bottom:1px solid rgba(255,255,255,.04);text-align:right;white-space:nowrap;font-weight:600;position:sticky;left:0;z-index:9;vertical-align:middle}
.grid-tbl .td-c{padding:3px 4px;border-bottom:1px solid rgba(255,255,255,.04);border-right:1px solid rgba(255,255,255,.04);vertical-align:top;height:50px;cursor:pointer;transition:background .1s}
.grid-tbl .td-c:hover{background:rgba(201,168,76,.07)}
.m-pill{border-radius:4px;padding:3px 5px;height:44px;display:flex;flex-direction:column;justify-content:center;border-left:2px solid}
.mp-n{font-size:10.5px;color:var(--cream);font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;line-height:1.3}
.mp-f{font-size:9px;color:var(--dim);white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
.mp-r{font-size:8.5px;font-family:'IBM Plex Mono',monospace;color:var(--gold);margin-top:1px}
.mp-group{font-size:8.5px;color:var(--pur);font-style:italic}
.add-ic{color:rgba(255,255,255,.09);font-size:13px;text-align:center;line-height:50px;width:100%;display:block}
/* overlay */
.overlay{position:fixed;inset:0;background:rgba(0,0,0,.72);z-index:500;display:flex;align-items:flex-start;justify-content:center;padding:30px 16px;backdrop-filter:blur(5px);overflow-y:auto}
.modal{background:var(--ink2);border:1px solid rgba(201,168,76,.18);border-radius:10px;width:100%;box-shadow:0 24px 64px rgba(0,0,0,.6);position:relative}
.modal-hdr{padding:22px 24px 0;border-bottom:1px solid rgba(201,168,76,.1);padding-bottom:16px}
.modal-title{font-family:'Playfair Display',serif;font-size:18px;color:var(--gold)}
.modal-sub{font-size:12px;color:var(--dim);margin-top:3px}
.modal-body{padding:20px 24px}
.modal-footer{padding:14px 24px 20px;display:flex;gap:8px;justify-content:flex-end;border-top:1px solid rgba(255,255,255,.05)}
.modal-tabs{display:flex;border-bottom:1px solid rgba(255,255,255,.07);margin-bottom:18px;gap:0}
.mtab{padding:8px 16px;font-size:10.5px;cursor:pointer;color:var(--dim);border:none;background:none;font-family:'IBM Plex Mono',monospace;text-transform:uppercase;letter-spacing:.06em;border-bottom:2px solid transparent;transition:all .15s}
.mtab.on{color:var(--gold);border-bottom-color:var(--gold)}
/* export card */
.ex-card{background:var(--ink3);border:1px solid rgba(201,168,76,.12);border-radius:8px;padding:16px 18px;cursor:pointer;transition:all .15s;display:flex;flex-direction:column;gap:8px}
.ex-card:hover{border-color:rgba(201,168,76,.3);background:rgba(201,168,76,.04)}
.ex-card-ico{font-size:28px}
.ex-card-t{font-family:'Playfair Display',serif;font-size:14px;color:var(--cream)}
.ex-card-s{font-size:11.5px;color:var(--dim);line-height:1.6}
/* day tab btn */
.day-btn{padding:6px 14px;border-radius:6px;font-size:10.5px;cursor:pointer;font-family:'IBM Plex Mono',monospace;letter-spacing:.05em;text-transform:uppercase;transition:all .15s;border:1px solid}
.doff{background:transparent;color:var(--dim);border-color:rgba(255,255,255,.07)}.doff:hover{color:var(--txt)}
.d14on{background:rgba(74,143,212,.13);color:var(--blu);border-color:rgba(74,143,212,.28)}
.d15on{background:rgba(74,175,122,.13);color:var(--grn);border-color:rgba(74,175,122,.28)}
/* fund group toggle */
.fund-group{background:var(--ink3);border:1px solid rgba(201,168,76,.12);border-radius:7px;padding:10px 14px;margin-bottom:6px;display:flex;align-items:center;gap:10px}
.toggle{position:relative;display:inline-block;width:38px;height:20px;flex-shrink:0}
.toggle input{opacity:0;width:0;height:0;position:absolute}
.toggle-track{position:absolute;inset:0;border-radius:20px;background:rgba(255,255,255,.1);transition:.2s;cursor:pointer}
.toggle input:checked+.toggle-track{background:var(--gold)}
.toggle-thumb{position:absolute;width:16px;height:16px;border-radius:50%;background:#fff;top:2px;left:2px;transition:.2s;pointer-events:none}
.toggle input:checked~.toggle-thumb{left:20px}
/* attendees */
.attendee-row{display:flex;gap:8px;align-items:center;padding:6px 0;border-bottom:1px solid rgba(255,255,255,.04)}
/* search */
.srch{position:relative}
.srch-ic{position:absolute;left:9px;top:50%;transform:translateY(-50%);color:var(--dim);pointer-events:none;font-size:12px}
.srch .inp{padding-left:28px}
/* dbar */
.dbar{height:2px;border-radius:2px;margin-top:3px;background:rgba(255,255,255,.05)}
.dfill{height:2px;border-radius:2px}
/* section header */
.sec-hdr{font-family:'IBM Plex Mono',monospace;font-size:8.5px;letter-spacing:.12em;text-transform:uppercase;color:var(--dim);padding:10px 0 5px;border-bottom:1px solid rgba(255,255,255,.05);margin-bottom:6px}
`;

/* ═══════════════════════════════════════════════════════════════════
   INVESTOR PROFILE MODAL
═══════════════════════════════════════════════════════════════════ */
function InvestorModal({inv, investors, meetings, companies, fundGrouping, onUpdateInv, onToggleFundGroup, onExport, onClose}){
  const [activeTab, setActiveTab]=useState("profile");
  const [editField, setEditField]=useState({});

  const invMeetings=meetings.filter(m=>(m.invIds||[]).includes(inv.id))
    .sort((a,b)=>ALL_SLOTS.indexOf(a.slotId)-ALL_SLOTS.indexOf(b.slotId));

  // Fundmates
  const fundmates=investors.filter(i=>i.id!==inv.id&&i.fund===inv.fund&&inv.fund);
  const isGrouped=inv.fund?(fundGrouping[inv.fund]!==false):false;

  function toggleSlot(slotId){
    const base=inv.slots||[];
    if(!base.includes(slotId)){
      // Not in base availability — can't unblock something not available. But allow adding to base.
      onUpdateInv({...inv,slots:[...base,slotId].sort((a,b)=>ALL_SLOTS.indexOf(a)-ALL_SLOTS.indexOf(b))});
    } else {
      const blocked=inv.blockedSlots||[];
      if(blocked.includes(slotId)){
        onUpdateInv({...inv,blockedSlots:blocked.filter(s=>s!==slotId)});
      } else {
        onUpdateInv({...inv,blockedSlots:[...blocked,slotId]});
      }
    }
  }

  function toggleCo(coId){
    const cos=inv.companies||[];
    if(cos.includes(coId)) onUpdateInv({...inv,companies:cos.filter(c=>c!==coId)});
    else onUpdateInv({...inv,companies:[...cos,coId]});
  }

  const eff=effectiveSlots(inv);
  const totalSlots=ALL_SLOTS.length;

  return(
    <div className="overlay" onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
      <div className="modal" style={{maxWidth:680}}>
        <div className="modal-hdr">
          <div className="modal-title">{inv.name}</div>
          <div className="modal-sub">{[inv.position,inv.fund].filter(Boolean).join(" · ")}</div>
          <div className="modal-tabs" style={{marginTop:14}}>
            {["profile","restrictions","companies","meetings"].map(t=>(
              <button key={t} className={`mtab${activeTab===t?" on":""}`} onClick={()=>setActiveTab(t)}>
                {{profile:"👤 Perfil",restrictions:"🕐 Horarios",companies:"🏢 Compañías",meetings:"📅 Reuniones"}[t]}
              </button>
            ))}
          </div>
        </div>

        <div className="modal-body">

          {/* PROFILE */}
          {activeTab==="profile"&&(
            <div>
              <div className="g2" style={{gap:12,marginBottom:14}}>
                {[["name","Nombre completo","text"],["fund","Fondo / Firma","text"],["position","Cargo","text"],["email","Email","email"],["phone","Teléfono","text"],["aum","AUM","text"]].map(([f,label,type])=>(
                  <div key={f}>
                    <div className="lbl">{label}</div>
                    <input className="inp" type={type} value={editField[f]!==undefined?editField[f]:(inv[f]||"")}
                      onChange={e=>setEditField(p=>({...p,[f]:e.target.value}))}
                      onBlur={()=>{if(editField[f]!==undefined){onUpdateInv({...inv,...editField});setEditField({});}}}/>
                  </div>
                ))}
              </div>
              {fundmates.length>0&&(
                <div className="fund-group" style={{marginTop:4}}>
                  <div style={{flex:1}}>
                    <div style={{fontSize:12.5,color:"var(--cream)"}}>Agrupar con colegas del mismo fondo</div>
                    <div style={{fontSize:11,color:"var(--dim)",marginTop:2}}>
                      {fundmates.map(f=>f.name).join(", ")} — misma reunión por defecto
                    </div>
                  </div>
                  <label className="toggle">
                    <input type="checkbox" checked={isGrouped} onChange={()=>onToggleFundGroup(inv.fund,!isGrouped)}/>
                    <div className="toggle-track"/>
                    <div className="toggle-thumb"/>
                  </label>
                </div>
              )}
            </div>
          )}

          {/* TIME RESTRICTIONS */}
          {activeTab==="restrictions"&&(
            <div>
              <p style={{fontSize:12.5,color:"var(--dim)",marginBottom:14,lineHeight:1.7}}>
                Verde = disponible · Rojo = bloqueado · Hacé clic para bloquear / desbloquear un slot.
                Slots con fondo gris = fuera de su disponibilidad declarada en el formulario.
              </p>
              <div style={{fontSize:11.5,color:"var(--txt)",marginBottom:12}}>
                <span className="bdg bg-grn">{eff.length}</span> slots efectivos de {totalSlots} totales
              </div>
              {DAYS.map(d=>(
                <div key={d} style={{marginBottom:16}}>
                  <div style={{fontFamily:"IBM Plex Mono,monospace",fontSize:10,color:d==="apr14"?"var(--blu)":"var(--grn)",marginBottom:6,letterSpacing:".06em",textTransform:"uppercase"}}>◆ {DAY_SHORT[d]}</div>
                  <div style={{display:"grid",gridTemplateColumns:`repeat(${HOURS.length},1fr)`,gap:3}}>
                    {HOURS.map(h=>{
                      const sid=`${d}-${h}`;
                      const inBase=(inv.slots||[]).includes(sid);
                      const isBlocked=(inv.blockedSlots||[]).includes(sid);
                      const cls=!inBase?"slot-na":isBlocked?"slot-blocked":"slot-avail";
                      return(
                        <div key={h} className={`slot-cell ${cls}`} onClick={()=>inBase&&toggleSlot(sid)}>
                          {hourLabel(h)}
                        </div>
                      );
                    })}
                  </div>
                </div>
              ))}
            </div>
          )}

          {/* COMPANIES */}
          {activeTab==="companies"&&(
            <div>
              <p style={{fontSize:12.5,color:"var(--dim)",marginBottom:14}}>Seleccioná las compañías que este inversor quiere reunirse:</p>
              {["Financials","Energy","Infra","Real Estate","TMT"].map(sector=>{
                const scos=companies.filter(c=>c.sector===sector);
                if(!scos.length) return null;
                return(
                  <div key={sector}>
                    <div className="sec-hdr">{sector}</div>
                    <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:5,marginBottom:8}}>
                      {scos.map(c=>{
                        const on=(inv.companies||[]).includes(c.id);
                        return(
                          <div key={c.id} onClick={()=>toggleCo(c.id)}
                            style={{display:"flex",alignItems:"center",gap:8,padding:"7px 10px",borderRadius:6,cursor:"pointer",
                              background:on?"rgba(201,168,76,.1)":"rgba(255,255,255,.03)",
                              border:`1px solid ${on?"rgba(201,168,76,.25)":"rgba(255,255,255,.06)"}`,transition:"all .12s"}}>
                            <div style={{width:14,height:14,borderRadius:3,background:on?"var(--gold)":"rgba(255,255,255,.1)",display:"flex",alignItems:"center",justifyContent:"center",flexShrink:0,fontSize:10,color:"var(--ink)",fontWeight:700}}>
                              {on?"✓":""}
                            </div>
                            <span style={{fontSize:12,color:on?"var(--cream)":"var(--dim)"}}>{c.name}</span>
                            <span className="bdg bg-g" style={{marginLeft:"auto",fontSize:9}}>{c.ticker}</span>
                          </div>
                        );
                      })}
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {/* MEETINGS */}
          {activeTab==="meetings"&&(
            <div>
              {invMeetings.length===0
                ?<div className="alert ai">Sin reuniones asignadas todavía.</div>
                :(
                  <table className="tbl">
                    <thead><tr><th>Día</th><th>Hora</th><th>Compañía</th><th>Sala</th></tr></thead>
                    <tbody>
                      {invMeetings.map(m=>{
                        const co=companies.find(c=>c.id===m.coId);
                        return(
                          <tr key={m.id}>
                            <td><span className={`bdg ${slotDay(m.slotId)==="apr14"?"bg-b":"bg-grn"}`}>{slotDay(m.slotId)==="apr14"?"Apr 14":"Apr 15"}</span></td>
                            <td style={{fontFamily:"IBM Plex Mono,monospace",fontWeight:600,fontSize:11}}>{slotLabel(m.slotId)}</td>
                            <td style={{color:"var(--cream)",fontWeight:600}}>{co?.name}<span className="bdg bg-g" style={{marginLeft:6}}>{co?.ticker}</span></td>
                            <td style={{fontFamily:"IBM Plex Mono,monospace",fontSize:11,color:"var(--gold)"}}>{m.room}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                )
              }
            </div>
          )}
        </div>

        <div className="modal-footer">
          <button className="btn bo bs" onClick={()=>onExport(inv,"pdf")}>📄 PDF</button>
          <button className="btn bo bs" onClick={()=>onExport(inv,"word")}>📝 Word</button>
          <button className="btn bg bs" style={{marginLeft:8}} onClick={onClose}>Cerrar</button>
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════════
   COMPANY PROFILE MODAL
═══════════════════════════════════════════════════════════════════ */
function CompanyModal({co, allCos, meetings, investors, onUpdateCo, onExport, onClose}){
  const [activeTab,setActiveTab]=useState("info");
  const [newName,setNewName]=useState("");
  const [newTitle,setNewTitle]=useState("");

  const coMeetings=meetings.filter(m=>m.coId===co.id)
    .sort((a,b)=>ALL_SLOTS.indexOf(a.slotId)-ALL_SLOTS.indexOf(b.slotId));

  return(
    <div className="overlay" onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
      <div className="modal" style={{maxWidth:620}}>
        <div className="modal-hdr">
          <div style={{display:"flex",alignItems:"baseline",gap:10}}>
            <div className="modal-title">{co.name}</div>
            <span className="bdg bg-g">{co.ticker}</span>
          </div>
          <div className="modal-sub" style={{color:SEC_CLR[co.sector]||"var(--dim)"}}>{co.sector}</div>
          <div className="modal-tabs" style={{marginTop:14}}>
            {["info","attendees","meetings"].map(t=>(
              <button key={t} className={`mtab${activeTab===t?" on":""}`} onClick={()=>setActiveTab(t)}>
                {{info:"🏢 Info",attendees:"👤 Asistentes",meetings:"📅 Reuniones"}[t]}
              </button>
            ))}
          </div>
        </div>

        <div className="modal-body">

          {activeTab==="info"&&(
            <div>
              <div className="g2" style={{gap:12}}>
                {[["name","Nombre"],["ticker","Ticker"],["sector","Sector"]].map(([f,label])=>(
                  <div key={f}>
                    <div className="lbl">{label}</div>
                    <input className="inp" value={co[f]||""} onChange={e=>onUpdateCo({...co,[f]:e.target.value})}/>
                  </div>
                ))}
              </div>
              <div style={{marginTop:14,padding:12,background:"var(--ink3)",borderRadius:7,fontSize:12,color:"var(--dim)"}}>
                <strong style={{color:"var(--txt)"}}>Reuniones asignadas:</strong> {coMeetings.length}<br/>
                <strong style={{color:"var(--txt)"}}>Inversores únicos:</strong> {new Set(coMeetings.flatMap(m=>m.invIds)).size}
              </div>
            </div>
          )}

          {activeTab==="attendees"&&(
            <div>
              <p style={{fontSize:12.5,color:"var(--dim)",marginBottom:14}}>Representantes de la compañía que asistirán al evento:</p>
              {(co.attendees||[]).map((a,i)=>(
                <div key={i} className="attendee-row">
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{fontSize:13,color:"var(--cream)"}}>{a.name}</div>
                    {a.title&&<div style={{fontSize:11,color:"var(--dim)"}}>{a.title}</div>}
                  </div>
                  <button className="btn bd bs" onClick={()=>onUpdateCo({...co,attendees:(co.attendees||[]).filter((_,j)=>j!==i)})}>✕</button>
                </div>
              ))}
              <div style={{display:"flex",gap:8,marginTop:12}}>
                <div style={{flex:1}}>
                  <div className="lbl">Nombre</div>
                  <input className="inp" placeholder="Juan García" value={newName} onChange={e=>setNewName(e.target.value)}/>
                </div>
                <div style={{flex:1}}>
                  <div className="lbl">Cargo</div>
                  <input className="inp" placeholder="CEO" value={newTitle} onChange={e=>setNewTitle(e.target.value)}
                    onKeyDown={e=>{if(e.key==="Enter"&&newName.trim()){onUpdateCo({...co,attendees:[...(co.attendees||[]),(({name:newName.trim(),title:newTitle.trim()}))]}); setNewName("");setNewTitle("");}}}/>
                </div>
                <button className="btn bg bs" style={{alignSelf:"flex-end"}} onClick={()=>{if(newName.trim()){onUpdateCo({...co,attendees:[...(co.attendees||[]),{name:newName.trim(),title:newTitle.trim()}]});setNewName("");setNewTitle("");}}}> + </button>
              </div>
            </div>
          )}

          {activeTab==="meetings"&&(
            <div>
              {coMeetings.length===0
                ?<div className="alert ai">Sin reuniones asignadas todavía.</div>
                :(
                  <table className="tbl">
                    <thead><tr><th>Día</th><th>Hora</th><th>Inversor(es)</th><th>Sala</th></tr></thead>
                    <tbody>
                      {coMeetings.map(m=>{
                        const invs=(m.invIds||[]).map(id=>investors.find(i=>i.id===id)).filter(Boolean);
                        return(
                          <tr key={m.id}>
                            <td><span className={`bdg ${slotDay(m.slotId)==="apr14"?"bg-b":"bg-grn"}`}>{slotDay(m.slotId)==="apr14"?"Apr 14":"Apr 15"}</span></td>
                            <td style={{fontFamily:"IBM Plex Mono,monospace",fontWeight:600,fontSize:11}}>{slotLabel(m.slotId)}</td>
                            <td>
                              {invs.map(inv=>(
                                <div key={inv.id} style={{fontSize:12,color:"var(--cream)"}}>{inv.name}<span style={{color:"var(--dim)",fontSize:10.5}}> — {inv.fund}</span></div>
                              ))}
                            </td>
                            <td style={{fontFamily:"IBM Plex Mono,monospace",fontSize:11,color:"var(--gold)"}}>{m.room}</td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                )
              }
            </div>
          )}
        </div>

        <div className="modal-footer">
          <button className="btn bo bs" onClick={()=>onExport(co,"pdf")}>📄 PDF</button>
          <button className="btn bo bs" onClick={()=>onExport(co,"word")}>📝 Word</button>
          <button className="btn bg bs" style={{marginLeft:8}} onClick={onClose}>Cerrar</button>
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════════
   MEETING EDIT MODAL
═══════════════════════════════════════════════════════════════════ */
function MeetingModal({mode,meeting,investors,meetings,companies,onSave,onDelete,onClose}){
  const [invIds,setInvIds]=useState(meeting?.invIds||[]);
  const [coId,setCoId]=useState(meeting?.coId||"");
  const [slotId,setSlotId]=useState(meeting?.slotId||"");
  const [room,setRoom]=useState(meeting?.room||ROOMS[0]);

  const conflicts=useMemo(()=>{
    const c=[];
    if(!invIds.length||!coId||!slotId) return c;
    for(const invId of invIds){
      const clash=meetings.find(m=>m.invIds?.includes(invId)&&m.slotId===slotId&&m.id!==meeting?.id);
      if(clash) c.push(`${investors.find(i=>i.id===invId)?.name} ya tiene reunión a esa hora`);
    }
    const coCl=meetings.find(m=>m.coId===coId&&m.slotId===slotId&&m.id!==meeting?.id);
    if(coCl) c.push(`${companies.find(c2=>c2.id===coId)?.name} ya tiene reunión a esa hora`);
    const rmCl=meetings.find(m=>m.room===room&&m.slotId===slotId&&m.id!==meeting?.id);
    if(rmCl) c.push(`${room} ya está ocupada a esa hora`);
    return c;
  },[invIds,coId,slotId,room,meetings,meeting]);

  const toggleInv=id=>{if(invIds.includes(id)) setInvIds(invIds.filter(x=>x!==id));else setInvIds([...invIds,id]);};

  return(
    <div className="overlay" onClick={e=>{if(e.target===e.currentTarget)onClose();}}>
      <div className="modal" style={{maxWidth:500}}>
        <div className="modal-hdr">
          <div className="modal-title">{mode==="add"?"Nueva Reunión":"Editar Reunión"}</div>
        </div>
        <div className="modal-body">
          <div className="modal-field" style={{marginBottom:13}}>
            <div className="lbl">Inversor(es)</div>
            <div style={{maxHeight:160,overflowY:"auto",background:"var(--ink3)",borderRadius:6,border:"1px solid rgba(201,168,76,.14)",padding:"6px"}}>
              {investors.map(inv=>(
                <label key={inv.id} style={{display:"flex",alignItems:"center",gap:8,padding:"4px 6px",cursor:"pointer",borderRadius:4,background:invIds.includes(inv.id)?"rgba(201,168,76,.1)":"transparent"}}>
                  <input type="checkbox" checked={invIds.includes(inv.id)} onChange={()=>toggleInv(inv.id)} style={{accentColor:"var(--gold)"}}/>
                  <span style={{fontSize:12.5,color:"var(--txt)"}}>{inv.name}</span>
                  <span style={{fontSize:10.5,color:"var(--dim)",marginLeft:"auto"}}>{inv.fund}</span>
                </label>
              ))}
            </div>
          </div>

          <div className="g2" style={{gap:12,marginBottom:13}}>
            <div>
              <div className="lbl">Compañía</div>
              <select className="sel" value={coId} onChange={e=>setCoId(e.target.value)}>
                <option value="">-- seleccionar --</option>
                {companies.map(c=><option key={c.id} value={c.id}>{c.name} ({c.ticker})</option>)}
              </select>
            </div>
            <div>
              <div className="lbl">Sala</div>
              <select className="sel" value={room} onChange={e=>setRoom(e.target.value)}>
                {ROOMS.map(r=><option key={r} value={r}>{r}</option>)}
              </select>
            </div>
          </div>

          <div>
            <div className="lbl">Día y Hora</div>
            <select className="sel" value={slotId} onChange={e=>setSlotId(e.target.value)}>
              <option value="">-- seleccionar --</option>
              {DAYS.map(d=><optgroup key={d} label={DAY_SHORT[d]}>{HOURS.map(h=><option key={`${d}-${h}`} value={`${d}-${h}`}>{DAY_SHORT[d]} {hourLabel(h)}</option>)}</optgroup>)}
            </select>
          </div>

          {conflicts.length>0&&<div className="alert aw" style={{marginTop:12}}>⚠ {conflicts.join(" · ")}</div>}
        </div>

        <div className="modal-footer">
          {mode==="edit"&&<button className="btn bd bs" onClick={onDelete}>🗑 Eliminar</button>}
          <button className="btn bo bs" onClick={onClose}>Cancelar</button>
          <button className="btn bg bs" disabled={!invIds.length||!coId||!slotId}
            onClick={()=>onSave({invIds,coId,slotId,room})} style={{opacity:(!invIds.length||!coId||!slotId)?.5:1}}>
            {mode==="add"?"Agregar":"Guardar"}
          </button>
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════════
   MAIN APP
═══════════════════════════════════════════════════════════════════ */
export default function App(){
  const [tab,setTab]=useState("upload");
  const [investors,setInvestors]=useState([]);
  const [companies,setCompanies]=useState(COMPANIES_INIT.map(c=>({...c,attendees:[]})));
  const [meetings,setMeetings]=useState([]);
  const [unscheduled,setUnscheduled]=useState([]);
  const [fixedRoom,setFixedRoom]=useState({});
  const [fundGrouping,setFundGrouping]=useState({});// fund→bool (true=grouped)
  const [dinners,setDinners]=useState(DINNERS_INIT);
  const [activeDay,setActiveDay]=useState("apr14");
  const [search,setSearch]=useState("");
  const [fileName,setFileName]=useState("");
  const [modal,setModal]=useState(null);
  const [invProfile,setInvProfile]=useState(null);
  const [coProfile,setCoProfile]=useState(null);
  const fileRef=useRef();
  const scheduled=meetings.length>0;

  /* ── parse excel ── */
  const handleFile=useCallback(e=>{
    const file=e.target.files?.[0];if(!file)return;
    setFileName(file.name);
    const reader=new FileReader();
    reader.onload=ev=>{
      const wb=XLSX.read(ev.target.result,{type:"array"});
      const ws=wb.Sheets[wb.SheetNames[0]];
      const rows=XLSX.utils.sheet_to_json(ws,{header:1});
      if(rows.length<2)return;
      const hdrs=rows[0].map(String);
      const ci=pred=>hdrs.findIndex(h=>pred(h.toLowerCase().replace(/[\s\n]+/g," ").trim()));
      const fundIdx=ci(h=>h==="fund"),nameIdx=ci(h=>h==="name"),surIdx=ci(h=>h.startsWith("surname"));
      const posIdx=ci(h=>h.startsWith("position")),emailIdx=ci(h=>h==="email"),phoneIdx=ci(h=>h.includes("mobile")||h.includes("phone"));
      const aumIdx=ci(h=>h==="aum"),timeIdx=ci(h=>h.includes("preferred meeting date")),coIdx=ci(h=>h.includes("which meetings"));
      const g=(row,i)=>i>=0?String(row[i]??"").trim():"";
      const parsed=rows.slice(1).filter(row=>g(row,fundIdx)||g(row,nameIdx)).map((row,ri)=>{
        const name=[g(row,nameIdx),g(row,surIdx)].filter(Boolean).join(" ")||`Inversor ${ri+1}`;
        const comps=[...new Set(g(row,coIdx).split(";").map(s=>s.trim()).filter(Boolean).map(resolveCo).filter(Boolean))];
        return{id:`inv-${ri}`,name,fund:g(row,fundIdx),email:g(row,emailIdx),phone:g(row,phoneIdx),position:g(row,posIdx),aum:g(row,aumIdx),companies:comps,slots:parseAvail(g(row,timeIdx)),blockedSlots:[],notes:""};
      });
      setInvestors(parsed);
      // Auto-set fundGrouping to true for all multi-person funds
      const fg={};const fm={};
      parsed.forEach(inv=>{if(inv.fund){fm[inv.fund]=(fm[inv.fund]||0)+1;}});
      Object.entries(fm).forEach(([f,n])=>{if(n>1)fg[f]=true;});
      setFundGrouping(fg);
      setMeetings([]);setUnscheduled([]);setFixedRoom({});setTab("investors");
    };
    reader.readAsArrayBuffer(file);
  },[]);

  /* ── generate ── */
  const generate=()=>{
    const res=runSchedule(investors,fundGrouping);
    setMeetings(res.meetings);setUnscheduled(res.unscheduled);setFixedRoom(res.fixedRoom);setTab("schedule");
  };

  /* ── meeting edits ── */
  const handleMeetingSave=({invIds,coId,slotId,room})=>{
    const id=modal.mode==="edit"?modal.meeting.id:`m-${Date.now()}-${Math.random().toString(36).slice(2,5)}`;
    if(modal.mode==="edit") setMeetings(prev=>prev.map(m=>m.id===id?{...m,invIds,coId,slotId,room}:m));
    else setMeetings(prev=>[...prev,{id,invIds,coId,slotId,room}]);
    setModal(null);
  };

  /* ── export helpers ── */
  const openPrintWindow=html=>{
    const w=window.open("","_blank");
    w.document.write(html);w.document.close();
    setTimeout(()=>{w.focus();w.print();},700);
  };

  function exportInvestor(inv,format){
    const data=investorToExportData(inv,meetings,companies,dinners);
    if(!data){alert("Este inversor no tiene reuniones asignadas.");return;}
    const fname=`${inv.fund||inv.name}_${inv.name.replace(/\s+/g,"_")}`.replace(/[^a-zA-Z0-9_\-]/g,"");
    if(format==="word"){
      const html=buildWordHTML(data.name,data.sub,data.sections);
      downloadBlob(`${fname}_schedule.doc`,html,"application/msword");
    } else {
      openPrintWindow(buildPrintHTML([data]));
    }
  }

  function exportCompany(co,format){
    const data=companyToExportData(co,meetings,investors);
    if(!data){alert("Esta compañía no tiene reuniones asignadas.");return;}
    const fname=co.ticker;
    if(format==="word"){
      const html=buildWordHTML(data.name,data.sub,data.sections);
      downloadBlob(`${fname}_schedule.doc`,html,"application/msword");
    } else {
      openPrintWindow(buildPrintHTML([data],{attendees:co.attendees}));
    }
  }

  function exportAll(scope,format){
    if(!scheduled){alert("Generá la agenda primero.");return;}
    let entities=[];
    if(scope==="companies"){
      entities=companies.map(co=>companyToExportData(co,meetings,investors)).filter(Boolean);
    } else {
      entities=investors.map(inv=>investorToExportData(inv,meetings,companies,dinners)).filter(Boolean);
    }
    if(!entities.length){alert("No hay datos para exportar.");return;}

    if(format==="pdf_combined"){
      openPrintWindow(buildPrintHTML(entities));return;
    }

    // ZIP of individual files
    const files=entities.map(e=>{
      const safeName=e.name.replace(/[^a-zA-Z0-9\s\-_]/g,"").replace(/\s+/g,"_").slice(0,40);
      const ext=format==="word"?".doc":".html";
      const html=format==="word"?buildWordHTML(e.name,e.sub,e.sections):buildPrintHTML([e]);
      return{name:`${safeName}${ext}`,data:html};
    });
    const zipBuf=buildZip(files);
    const suffix=scope==="companies"?"Companies":"Investors";
    downloadBlob(`ArgentinaInNY2026_${suffix}_Schedules.zip`,zipBuf,"application/zip");
  }

  /* ── derived ── */
  const byCompany=useMemo(()=>{
    const map={};companies.forEach(c=>{map[c.id]=[];});
    meetings.forEach(m=>map[m.coId]?.push(m));
    Object.values(map).forEach(arr=>arr.sort((a,b)=>ALL_SLOTS.indexOf(a.slotId)-ALL_SLOTS.indexOf(b.slotId)));
    return map;
  },[meetings,companies]);

  const byInvestor=useMemo(()=>{
    const map={};investors.forEach(i=>{map[i.id]=[];});
    meetings.forEach(m=>(m.invIds||[]).forEach(id=>map[id]?.push(m)));
    Object.values(map).forEach(arr=>arr.sort((a,b)=>ALL_SLOTS.indexOf(a.slotId)-ALL_SLOTS.indexOf(b.slotId)));
    return map;
  },[meetings,investors]);

  const activeCos=useMemo(()=>companies.filter(c=>investors.some(i=>(i.companies||[]).includes(c.id))),[companies,investors]);
  const dayCos=activeCos.filter(c=>meetings.some(m=>slotDay(m.slotId)===activeDay&&m.coId===c.id)||investors.some(i=>(i.companies||[]).includes(c.id)));

  const gridMap=useMemo(()=>{
    const map={};
    meetings.filter(m=>slotDay(m.slotId)===activeDay).forEach(m=>{map[`${m.coId}::${slotHour(m.slotId)}`]=m;});
    return map;
  },[meetings,activeDay]);

  const filtered=useMemo(()=>{
    if(!search) return investors;
    const q=search.toLowerCase();
    return investors.filter(i=>i.name.toLowerCase().includes(q)||i.fund.toLowerCase().includes(q));
  },[investors,search]);

  // fund groups for schedule view
  const fundGroups=useMemo(()=>{
    const m={};investors.forEach(inv=>{if(inv.fund){if(!m[inv.fund])m[inv.fund]=[];m[inv.fund].push(inv.id);}});
    return Object.entries(m).filter(([,ids])=>ids.length>1);
  },[investors]);

  const TABS=[
    {id:"upload",label:"📥 Cargar"},
    {id:"investors",label:`👥 Inversores (${investors.length})`},
    {id:"companies",label:"🏢 Compañías"},
    {id:"schedule",label:"📅 Agenda"},
    {id:"export",label:"⬇ Exportar"},
  ];

  return(
    <div className="app">
      <style>{CSS}</style>

      {/* MODALS */}
      {invProfile&&<InvestorModal
        inv={invProfile} investors={investors} meetings={meetings} companies={companies}
        fundGrouping={fundGrouping}
        onUpdateInv={updated=>{setInvestors(prev=>prev.map(i=>i.id===updated.id?updated:i));setInvProfile(updated);}}
        onToggleFundGroup={(fund,val)=>setFundGrouping(p=>({...p,[fund]:val}))}
        onExport={exportInvestor}
        onClose={()=>setInvProfile(null)}
      />}

      {coProfile&&<CompanyModal
        co={coProfile} allCos={companies} meetings={meetings} investors={investors}
        onUpdateCo={updated=>{setCompanies(prev=>prev.map(c=>c.id===updated.id?updated:c));setCoProfile(updated);}}
        onExport={exportCompany}
        onClose={()=>setCoProfile(null)}
      />}

      {modal&&<MeetingModal
        mode={modal.mode} meeting={modal.meeting} investors={investors} meetings={meetings} companies={companies}
        onSave={handleMeetingSave}
        onDelete={()=>{setMeetings(prev=>prev.filter(m=>m.id!==modal.meeting.id));setModal(null);}}
        onClose={()=>setModal(null)}
      />}

      {/* HEADER */}
      <header className="hdr">
        <div className="brand">
          <h1>Argentina in New York 2026</h1>
          <p>Latin Securities · Roadshow Manager</p>
        </div>
        <nav className="nav">
          {TABS.map(t=>(
            <button key={t.id} className={`ntab${tab===t.id?" on":""}`} onClick={()=>setTab(t.id)}>{t.label}</button>
          ))}
        </nav>
      </header>

      <div className="body">

        {/* ════════════════════ UPLOAD ════════════════════ */}
        {tab==="upload"&&(
          <div>
            <h2 className="pg-h">Cargar Respuestas</h2>
            <p className="pg-s">Excel exportado de Microsoft Forms — procesamiento automático.</p>
            <div className="card">
              <div className="upz" onClick={()=>fileRef.current?.click()}>
                <div style={{fontSize:34,marginBottom:8}}>📊</div>
                <div style={{fontSize:15,color:"var(--cream)",marginBottom:4}}>{fileName||"Hacé clic para seleccionar el archivo"}</div>
                <div style={{fontSize:12,color:"var(--dim)"}}>
                  {fileName?<span style={{color:"var(--grn)"}}>✓ {investors.length} inversores procesados</span>:"Formato .xlsx · Microsoft Forms export"}
                </div>
                <input ref={fileRef} type="file" accept=".xlsx,.xls,.csv" style={{display:"none"}} onChange={handleFile}/>
              </div>
            </div>
            <div className="g2">
              <div className="card">
                <div className="card-t">⚙ Configuración</div>
                <div style={{fontSize:12.5,color:"var(--dim)",lineHeight:1.8}}>
                  <div>🚪 <strong style={{color:"var(--txt)"}}>12 salas</strong> fijas — Top 12 compañías por demanda</div>
                  <div>🕐 <strong style={{color:"var(--txt)"}}>9 slots por día</strong> × 2 días = 18 slots totales</div>
                  <div>👥 <strong style={{color:"var(--txt)"}}>Agrupación de fondos</strong> automática (configurable)</div>
                  <div>⚠ <strong style={{color:"var(--txt)"}}>Restricciones</strong> por inversor editables en perfil</div>
                </div>
              </div>
              <div className="card">
                <div className="card-t">📤 Exportación</div>
                <div style={{fontSize:12.5,color:"var(--dim)",lineHeight:1.8}}>
                  <div>📄 <strong style={{color:"var(--txt)"}}>PDF</strong> — ventana de impresión del navegador</div>
                  <div>📝 <strong style={{color:"var(--txt)"}}>Word (.doc)</strong> — descarga directa, abre en Word/Google Docs</div>
                  <div>🗜 <strong style={{color:"var(--txt)"}}>ZIP bulk</strong> — un archivo por entidad, todos juntos</div>
                  <div>🏢💼 Schedules <strong style={{color:"var(--txt)"}}>por compañía</strong> y <strong style={{color:"var(--txt)"}}>por inversor</strong></div>
                </div>
              </div>
            </div>
            {investors.length>0&&(
              <div className="flex" style={{marginTop:4}}>
                <button className="btn bg" onClick={generate}>🚀 Generar agenda</button>
                <button className="btn bo" onClick={()=>setTab("investors")}>Ver inversores →</button>
              </div>
            )}
          </div>
        )}

        {/* ════════════════════ INVESTORS ════════════════════ */}
        {tab==="investors"&&(
          <div>
            <h2 className="pg-h">Inversores / Fondos</h2>
            <p className="pg-s">Hacé clic en un inversor para ver su perfil completo, editar restricciones o exportar su schedule.</p>

            <div className="stats">
              <div className="stat"><div className="sv">{investors.length}</div><div className="sl">Inversores</div></div>
              <div className="stat"><div className="sv">{investors.reduce((s,i)=>s+(i.companies?.length||0),0)}</div><div className="sl">Solicitudes</div></div>
              <div className="stat"><div className="sv">{fundGroups.length}</div><div className="sl">Fondos agrupados</div></div>
              <div className="stat"><div className="sv">{investors.filter(i=>!(i.blockedSlots||[]).length).length}</div><div className="sl">Sin restricciones</div></div>
            </div>

            {/* Fund groups overview */}
            {fundGroups.length>0&&(
              <div className="card" style={{marginBottom:14}}>
                <div className="card-t">👥 Fondos con múltiples inversores</div>
                <div style={{display:"flex",flexWrap:"wrap",gap:8}}>
                  {fundGroups.map(([fund,ids])=>(
                    <div key={fund} style={{display:"flex",alignItems:"center",gap:8,padding:"7px 12px",background:"var(--ink3)",borderRadius:7,border:"1px solid rgba(201,168,76,.12)"}}>
                      <span style={{fontSize:12,color:"var(--txt)"}}>{fund}</span>
                      <span style={{fontSize:10,color:"var(--dim)"}}>{ids.length} personas</span>
                      <label className="toggle" style={{marginLeft:4}}>
                        <input type="checkbox" checked={fundGrouping[fund]!==false} onChange={()=>setFundGrouping(p=>({...p,[fund]:!(p[fund]!==false)}))}/>
                        <div className="toggle-track"/>
                        <div className="toggle-thumb"/>
                      </label>
                      <span style={{fontSize:9.5,color:fundGrouping[fund]!==false?"var(--gold)":"var(--dim)",fontFamily:"IBM Plex Mono,monospace"}}>
                        {fundGrouping[fund]!==false?"juntos":"separados"}
                      </span>
                    </div>
                  ))}
                </div>
              </div>
            )}

            <div className="flex" style={{marginBottom:13}}>
              <div className="srch" style={{flex:1,maxWidth:320}}>
                <span className="srch-ic">🔍</span>
                <input className="inp srch" placeholder="Buscar..." value={search} onChange={e=>setSearch(e.target.value)}/>
              </div>
              <button className="btn bg" style={{marginLeft:"auto"}} onClick={generate}>🚀 Generar agenda</button>
            </div>

            <div style={{maxHeight:560,overflowY:"auto"}}>
              {filtered.map(inv=>(
                <div key={inv.id} className="ent-row" onClick={()=>setInvProfile(inv)}>
                  <div style={{flex:1,minWidth:0}}>
                    <div style={{display:"flex",alignItems:"center",gap:8}}>
                      <span style={{fontFamily:"Playfair Display,serif",fontSize:14,color:"var(--cream)"}}>{inv.name}</span>
                      {(inv.blockedSlots||[]).length>0&&<span className="bdg bg-r">{inv.blockedSlots.length} bloqueados</span>}
                    </div>
                    <div style={{fontSize:11.5,color:"var(--dim)",marginTop:2}}>
                      {inv.fund&&<strong style={{color:"var(--txt)"}}>{inv.fund}</strong>}
                      {inv.position&&<> · {inv.position}</>}
                      {inv.aum&&<span className="bdg bg-g" style={{marginLeft:6}}>{inv.aum}</span>}
                    </div>
                    <div style={{marginTop:5,display:"flex",flexWrap:"wrap",gap:3}}>
                      {(inv.companies||[]).map(cid=>{const c=companies.find(x=>x.id===cid);return<span key={cid} className="tag" style={{borderColor:`${SEC_CLR[c?.sector]||"var(--gold)"}44`,color:SEC_CLR[c?.sector]||"var(--gold2)"}}>{c?.ticker||cid}</span>;})}
                    </div>
                  </div>
                  <div style={{fontSize:10,color:"var(--dim)",textAlign:"right",flexShrink:0}}>
                    <div>{(inv.companies||[]).length} co.</div>
                    <div>{effectiveSlots(inv).length} slots</div>
                    {scheduled&&<div className="bdg bg-grn" style={{marginTop:4}}>{(byInvestor[inv.id]||[]).length} mtgs</div>}
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        {/* ════════════════════ COMPANIES ════════════════════ */}
        {tab==="companies"&&(
          <div>
            <h2 className="pg-h">Compañías</h2>
            <p className="pg-s">Hacé clic en una compañía para gestionar asistentes, ver reuniones o exportar su schedule.</p>

            {["Financials","Energy","Infra","Real Estate","TMT"].map(sector=>{
              const scos=companies.filter(c=>c.sector===sector);
              if(!scos.length) return null;
              return(
                <div key={sector}>
                  <div className="sec-hdr">{{"Financials":"🏦 Financials","Energy":"⚡ Energy","Infra":"🏛 Infrastructure","Real Estate":"🏛 Real Estate","TMT":"📳 TMT"}[sector]||sector}</div>
                  <div className="g3" style={{marginBottom:10}}>
                    {scos.map(co=>{
                      const cms=byCompany[co.id]||[];
                      const demandInvs=new Set(investors.flatMap(i=>(i.companies||[]).includes(co.id)?[i.id]:[])).size;
                      const isFixed=fixedRoom[co.id];
                      return(
                        <div key={co.id} className="ent-row" onClick={()=>setCoProfile(co)} style={{flexDirection:"column",gap:6}}>
                          <div style={{display:"flex",alignItems:"baseline",gap:7}}>
                            <span style={{fontFamily:"Playfair Display,serif",fontSize:14,color:"var(--cream)"}}>{co.name}</span>
                            <span className="bdg bg-g">{co.ticker}</span>
                            {isFixed&&<span className="bdg bg-b" style={{fontSize:9}}>{fixedRoom[co.id]}</span>}
                          </div>
                          <div style={{fontSize:11,color:"var(--dim)"}}>
                            Demanda: <strong style={{color:"var(--txt)"}}>{demandInvs}</strong> inversores
                            {scheduled&&<> · <strong style={{color:"var(--grn)"}}>{cms.length}</strong> reuniones</>}
                          </div>
                          {(co.attendees||[]).length>0&&(
                            <div style={{fontSize:10.5,color:"var(--dim)"}}>
                              👤 {(co.attendees||[]).map(a=>a.name).join(", ")}
                            </div>
                          )}
                          <div className="dbar"><div className="dfill" style={{width:`${Math.min(100,(demandInvs/25)*100)}%`,background:SEC_CLR[co.sector]||"var(--gold)"}}/></div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* ════════════════════ SCHEDULE ════════════════════ */}
        {tab==="schedule"&&(
          <div>
            <h2 className="pg-h">Agenda</h2>
            <p className="pg-s">Compañías fijas · Los inversores se mueven · Clic en celda para editar</p>

            {!scheduled&&investors.length===0&&<div className="alert aw">Cargá el archivo Excel primero.</div>}
            {!scheduled&&investors.length>0&&<div className="alert ai">
              {investors.length} inversores listos.
              <button className="btn bg bs" style={{marginLeft:10}} onClick={generate}>🚀 Generar</button>
            </div>}

            {scheduled&&(
              <>
                <div className="stats">
                  <div className="stat"><div className="sv">{meetings.length}</div><div className="sl">Reuniones</div></div>
                  <div className="stat"><div className="sv" style={{color:unscheduled.length?"var(--red)":undefined}}>{unscheduled.length}</div><div className="sl" style={{color:unscheduled.length?"var(--red)":undefined}}>Sin asignar</div></div>
                  <div className="stat"><div className="sv">{meetings.filter(m=>slotDay(m.slotId)==="apr14").length}</div><div className="sl" style={{color:"var(--blu)"}}>Apr 14</div></div>
                  <div className="stat"><div className="sv">{meetings.filter(m=>slotDay(m.slotId)==="apr15").length}</div><div className="sl" style={{color:"var(--grn)"}}>Apr 15</div></div>
                  <div className="stat"><div className="sv">{meetings.filter(m=>(m.invIds||[]).length>1).length}</div><div className="sl">Grupales</div></div>
                  <div className="stat"><div className="sv">{new Set(meetings.map(m=>m.room)).size}</div><div className="sl">Salas usadas</div></div>
                </div>

                {unscheduled.length>0&&(
                  <div className="alert aw" style={{marginBottom:12}}>
                    ⚠ {unscheduled.length} reunión(es) sin asignar.
                    <button className="btn bd bs" style={{marginLeft:8}} onClick={()=>setTab("investors")}>Revisar perfiles</button>
                  </div>
                )}

                <div className="flex" style={{marginBottom:12}}>
                  {DAYS.map(d=>(
                    <button key={d} className={`day-btn ${activeDay===d?(d==="apr14"?"d14on":"d15on"):"doff"}`} onClick={()=>setActiveDay(d)}>
                      {d==="apr14"?"📅 Martes 14 Abril":"📅 Miércoles 15 Abril"}
                      <span style={{opacity:.7,marginLeft:4}}>({meetings.filter(m=>slotDay(m.slotId)===d).length})</span>
                    </button>
                  ))}
                  <button className="btn bo bs" style={{marginLeft:"auto"}} onClick={()=>setModal({mode:"add"})}>＋ Agregar</button>
                  <button className="btn bo bs" onClick={generate}>↺ Re-generar</button>
                  <button className="btn bg bs" onClick={()=>setTab("export")}>⬇ Exportar →</button>
                </div>

                <div className="card" style={{padding:"10px 4px"}}>
                  <div className="grid-wrap">
                    <table className="grid-tbl">
                      <colgroup><col style={{width:72}}/>{dayCos.map(c=><col key={c.id} style={{minWidth:115}}/>)}</colgroup>
                      <thead>
                        <tr>
                          <th className="th-time" style={{borderBottom:"1px solid rgba(201,168,76,.07)"}}/>
                          {dayCos.map(c=>(
                            <th key={c.id} className="th-sect" style={{background:`${SEC_CLR[c.sector]}13`,color:SEC_CLR[c.sector],borderBottom:`2px solid ${SEC_CLR[c.sector]}48`}}>
                              {c.sector}
                            </th>
                          ))}
                        </tr>
                        <tr>
                          <th className="th-time">Hora</th>
                          {dayCos.map(c=>{
                            const hr=fixedRoom[c.id];
                            return(
                              <th key={c.id} className="th-co" style={{borderBottom:`2px solid ${SEC_CLR[c.sector]}48`}}>
                                <div style={{color:SEC_CLR[c.sector],fontFamily:"Lora,serif",fontWeight:600,fontSize:11}}>{c.name}</div>
                                <div style={{fontSize:8.5,color:"var(--dim)",marginTop:1,fontFamily:"IBM Plex Mono,monospace"}}>{c.ticker}{hr?` · ${hr}`:""}</div>
                                <div className="dbar"><div className="dfill" style={{width:`${Math.min(100,((byCompany[c.id]||[]).length/9)*100)}%`,background:SEC_CLR[c.sector]}}/></div>
                              </th>
                            );
                          })}
                        </tr>
                      </thead>
                      <tbody>
                        {HOURS.map(h=>(
                          <tr key={h}>
                            <td className="td-time">{hourLabel(h)}</td>
                            {dayCos.map(c=>{
                              const m=gridMap[`${c.id}::${h}`];
                              if(m){
                                const invs=(m.invIds||[]).map(id=>investors.find(i=>i.id===id)).filter(Boolean);
                                const sclr=SEC_CLR[c.sector]||"var(--gold)";
                                const isGroup=invs.length>1;
                                return(
                                  <td key={c.id} className="td-c" onClick={()=>setModal({mode:"edit",meeting:m})}>
                                    <div className="m-pill" style={{background:`${sclr}11`,borderLeftColor:sclr}}>
                                      <div className="mp-n">{isGroup?invs.map(i=>i.name.split(" ")[0]).join(" + "):invs[0]?.name}</div>
                                      <div className="mp-f">{isGroup?`${invs[0]?.fund} (${invs.length})`:invs[0]?.fund}</div>
                                      <div className="mp-r">{m.room}{isGroup&&<span className="mp-group"> · grupo</span>}</div>
                                    </div>
                                  </td>
                                );
                              }
                              return <td key={c.id} className="td-c" onClick={()=>setModal({mode:"add",prefCoId:c.id,prefSlotId:`${activeDay}-${h}`})}><span className="add-ic">+</span></td>;
                            })}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                {unscheduled.length>0&&(
                  <div className="card" style={{marginTop:12}}>
                    <div className="card-t" style={{color:"var(--red)"}}>⚠ Sin asignar</div>
                    <table className="tbl">
                      <thead><tr><th>Inversor(es)</th><th>Compañía</th><th>Acción</th></tr></thead>
                      <tbody>
                        {unscheduled.map((u,i)=>(
                          <tr key={i}>
                            <td>{(u.invIds||[]).map(id=>investors.find(x=>x.id===id)?.name).join(", ")}</td>
                            <td>{companies.find(c=>c.id===u.coId)?.name}</td>
                            <td><button className="btn bo bs" onClick={()=>setModal({mode:"add",prefInvIds:u.invIds,prefCoId:u.coId})}>Asignar →</button></td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                )}
              </>
            )}
          </div>
        )}

        {/* ════════════════════ EXPORT ════════════════════ */}
        {tab==="export"&&(
          <div>
            <h2 className="pg-h">Exportar Schedules</h2>
            <p className="pg-s">Genera schedules individuales o en masa — formato Word (.doc) o PDF.</p>

            {/* ── Dinner Config ── */}
            <div style={{marginBottom:20}} className="card">
              <div className="card-t">🍽 Cenas / Dinners</div>
              <p style={{fontSize:12,color:"var(--dim)",marginBottom:14}}>Configurá los datos de las cenas para que aparezcan en los schedules de los inversores.</p>
              {dinners.map(din=>(
                <div key={din.id} style={{marginBottom:16,padding:"12px 14px",background:"var(--ink3)",borderRadius:8,border:"1px solid rgba(255,255,255,.07)"}}>
                  <div style={{fontSize:11,color:"var(--gold)",fontWeight:700,textTransform:"uppercase",letterSpacing:".08em",marginBottom:10}}>
                    {DAY_LONG[din.day]}
                  </div>
                  <div style={{display:"grid",gridTemplateColumns:"1fr 1fr",gap:10,marginBottom:10}}>
                    <label style={{display:"flex",flexDirection:"column",gap:4}}>
                      <span style={{fontSize:11,color:"var(--dim)"}}>Nombre del evento</span>
                      <input className="inp" value={din.name} onChange={e=>setDinners(prev=>prev.map(d=>d.id===din.id?{...d,name:e.target.value}:d))} placeholder="Conference Dinner"/>
                    </label>
                    <label style={{display:"flex",flexDirection:"column",gap:4}}>
                      <span style={{fontSize:11,color:"var(--dim)"}}>Restaurante</span>
                      <input className="inp" value={din.restaurant} onChange={e=>setDinners(prev=>prev.map(d=>d.id===din.id?{...d,restaurant:e.target.value}:d))} placeholder="The Lambs Club"/>
                    </label>
                  </div>
                  <label style={{display:"flex",flexDirection:"column",gap:4}}>
                    <span style={{fontSize:11,color:"var(--dim)"}}>Dirección</span>
                    <input className="inp" value={din.address} onChange={e=>setDinners(prev=>prev.map(d=>d.id===din.id?{...d,address:e.target.value}:d))} placeholder="132 W 44th St, New York, NY"/>
                  </label>
                </div>
              ))}
            </div>

            {!scheduled&&<div className="alert aw">Generá la agenda primero para poder exportar.</div>}

            {scheduled&&(
              <>
                <div className="card" style={{marginBottom:20}}>
                  <div className="card-t">📊 Resumen de contenido a exportar</div>
                  <div className="g3">
                    <div style={{padding:"10px 0",borderRight:"1px solid rgba(255,255,255,.06)"}}>
                      <div style={{fontSize:26,fontFamily:"Playfair Display,serif",color:"var(--gold)"}}>{companies.filter(c=>meetings.some(m=>m.coId===c.id)).length}</div>
                      <div style={{fontSize:10,color:"var(--dim)",textTransform:"uppercase",letterSpacing:".08em",fontFamily:"IBM Plex Mono,monospace",marginTop:3}}>Compañías con reuniones</div>
                    </div>
                    <div style={{padding:"10px 12px",borderRight:"1px solid rgba(255,255,255,.06)"}}>
                      <div style={{fontSize:26,fontFamily:"Playfair Display,serif",color:"var(--gold)"}}>{investors.filter(inv=>meetings.some(m=>(m.invIds||[]).includes(inv.id))).length}</div>
                      <div style={{fontSize:10,color:"var(--dim)",textTransform:"uppercase",letterSpacing:".08em",fontFamily:"IBM Plex Mono,monospace",marginTop:3}}>Inversores con reuniones</div>
                    </div>
                    <div style={{padding:"10px 12px"}}>
                      <div style={{fontSize:26,fontFamily:"Playfair Display,serif",color:"var(--gold)"}}>{meetings.length}</div>
                      <div style={{fontSize:10,color:"var(--dim)",textTransform:"uppercase",letterSpacing:".08em",fontFamily:"IBM Plex Mono,monospace",marginTop:3}}>Reuniones totales</div>
                    </div>
                  </div>
                </div>

                {/* POR COMPAÑIA */}
                <div style={{marginBottom:6}} className="sec-hdr">🏢 Por Compañía</div>
                <div className="g2" style={{marginBottom:22}}>
                  <div className="ex-card" onClick={()=>exportAll("companies","word")}>
                    <div className="ex-card-ico">📝🗜</div>
                    <div className="ex-card-t">Todas las compañías — Word ZIP</div>
                    <div className="ex-card-s">Un archivo .doc por compañía, todo en un ZIP. Abre directamente en Word o Google Docs.</div>
                  </div>
                  <div className="ex-card" onClick={()=>exportAll("companies","pdf_combined")}>
                    <div className="ex-card-ico">📄🗜</div>
                    <div className="ex-card-t">Todas las compañías — PDF combinado</div>
                    <div className="ex-card-s">Todas en un solo PDF, con salto de página entre compañías. Ideal para revisión interna.</div>
                  </div>
                </div>

                {/* POR INVERSOR */}
                <div style={{marginBottom:6}} className="sec-hdr">💼 Por Inversor / Fondo</div>
                <div className="g2" style={{marginBottom:22}}>
                  <div className="ex-card" onClick={()=>exportAll("investors","word")}>
                    <div className="ex-card-ico">📝🗜</div>
                    <div className="ex-card-t">Todos los inversores — Word ZIP</div>
                    <div className="ex-card-s">Un .doc por inversor con sus reuniones. Se envía individualmente a cada contacto.</div>
                  </div>
                  <div className="ex-card" onClick={()=>exportAll("investors","pdf_combined")}>
                    <div className="ex-card-ico">📄🗜</div>
                    <div className="ex-card-t">Todos los inversores — PDF combinado</div>
                    <div className="ex-card-s">Todos en un solo PDF. Útil para imprimir el booklet completo del evento.</div>
                  </div>
                </div>

                {/* INDIVIDUAL */}
                <div style={{marginBottom:6}} className="sec-hdr">🎯 Exportación individual</div>
                <div className="g2" style={{gap:13}}>
                  <div className="card">
                    <div className="card-t">Compañías individuales</div>
                    <div style={{maxHeight:300,overflowY:"auto",display:"flex",flexDirection:"column",gap:4}}>
                      {companies.filter(c=>meetings.some(m=>m.coId===c.id)).map(co=>(
                        <div key={co.id} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 8px",background:"var(--ink3)",borderRadius:6,border:"1px solid rgba(255,255,255,.05)"}}>
                          <span style={{flex:1,fontSize:12.5,color:"var(--txt)"}}>{co.name}</span>
                          <span className="bdg bg-g">{(byCompany[co.id]||[]).length} mtgs</span>
                          <button className="btn bo bs" onClick={()=>exportCompany(co,"pdf")}>PDF</button>
                          <button className="btn bo bs" onClick={()=>exportCompany(co,"word")}>Word</button>
                        </div>
                      ))}
                    </div>
                  </div>
                  <div className="card">
                    <div className="card-t">Inversores individuales</div>
                    <div style={{maxHeight:300,overflowY:"auto",display:"flex",flexDirection:"column",gap:4}}>
                      {investors.filter(inv=>meetings.some(m=>(m.invIds||[]).includes(inv.id))).map(inv=>(
                        <div key={inv.id} style={{display:"flex",alignItems:"center",gap:8,padding:"6px 8px",background:"var(--ink3)",borderRadius:6,border:"1px solid rgba(255,255,255,.05)"}}>
                          <div style={{flex:1,minWidth:0}}>
                            <div style={{fontSize:12.5,color:"var(--txt)",overflow:"hidden",textOverflow:"ellipsis",whiteSpace:"nowrap"}}>{inv.name}</div>
                            <div style={{fontSize:10,color:"var(--dim)"}}>{inv.fund}</div>
                          </div>
                          <span className="bdg bg-g">{(byInvestor[inv.id]||[]).length} mtgs</span>
                          <button className="btn bo bs" onClick={()=>exportInvestor(inv,"pdf")}>PDF</button>
                          <button className="btn bo bs" onClick={()=>exportInvestor(inv,"word")}>Word</button>
                        </div>
                      ))}
                    </div>
                  </div>
                </div>
              </>
            )}
          </div>
        )}

      </div>
    </div>
  );
}
