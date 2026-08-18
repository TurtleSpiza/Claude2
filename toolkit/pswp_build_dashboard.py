#!/usr/bin/env python3
"""
pswp_build_dashboard.py - build the standalone one-page LCC Parks expenditure map from the register.

    python3 pswp_build_dashboard.py PS_WP_Transaction_Register_3FY_v45_HANDOVER.xlsx \
        --gazetteer logan_park_gazetteer.csv \
        --divisions logan_divisions.geojson \
        --out LCC_Parks_Expenditure_Map_v45.html

Reads only: Register (cols DY/DZ site + basis, T amount, E FY, Z theme, P PK, L contractor,
AG status, D doc date), Site_Allocation, Config. Writes nothing back; the register is opened
read-only via python-calamine. Requires: pandas, python-calamine, shapely.

Positions come from Config (rule 19.6). The HTML is self-contained apart from the map tiles
and two CDN libraries (Leaflet, Chart.js), which need internet at view time.
"""
import argparse, json, re, sys, collections
import pandas as pd
from python_calamine import CalamineWorkbook
from shapely.geometry import shape, mapping

SHELL = r"""<!DOCTYPE html>
<html lang="en-AU">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>LCC Parks | Expenditure map {VER}, Park Services and Water Parks</title>
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.min.css">
<style>
:root{
--bg:#ECEEF2;--panel:#FFFFFF;--panel-alt:#F1F5F9;--border:#CBD5E1;
--header-bg:#1F3864;--header-fg:#FFFFFF;--txt:#0F172A;--muted:#475569;
--accent-teal:#0F766E;--accent-gold:#B45309;
--status-green:#047857;--status-amber:#C2410C;--status-red:#B91C1C;
--font-sans:'Segoe UI',Tahoma,Geneva,sans-serif;--font-mono:Consolas,'Courier New',monospace;
}
body.dark{
--bg:#0A0E1A;--panel:#121826;--panel-alt:#1A2236;--border:#243049;
--header-bg:#0E1422;--header-fg:#5EE0C1;--txt:#E8EDF7;--muted:#94A3B8;
--accent-teal:#5EE0C1;--accent-gold:#FFB547;
--status-green:#6FB088;--status-amber:#E8A34D;--status-red:#D96A6A;
}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--txt);font-family:var(--font-sans);font-size:14px;line-height:1.5}
a{color:var(--accent-teal)}
.wrap{max-width:1560px;margin:0 auto;padding:0 16px 48px}
header.bnr{background:var(--header-bg);color:var(--header-fg);padding:16px 24px;border-bottom:3px solid var(--accent-teal)}
header.bnr .inner{max-width:1560px;margin:0 auto;display:flex;justify-content:space-between;align-items:flex-end;gap:16px;flex-wrap:wrap}
h1{font-size:20px;margin:0;font-weight:600;letter-spacing:.2px}
.sub{font-size:12.5px;opacity:.8;margin-top:3px}
.stamp{font-size:11.5px;opacity:.75;text-align:right;font-family:var(--font-mono);line-height:1.6}
button,select,input[type=search]{font-family:inherit;font-size:12.5px;color:var(--txt);background:var(--panel);border:1px solid var(--border);border-radius:4px;padding:6px 10px;cursor:pointer}
button:hover,select:hover{border-color:var(--accent-teal)}
button.on{background:var(--accent-teal);color:#fff;border-color:var(--accent-teal)}
.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(140px,1fr));gap:1px;background:var(--border);border:1px solid var(--border);margin:16px 0 0}
.kpi{background:var(--panel);padding:11px 14px}
.kpi .lbl{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.4px}
.kpi .val{font-size:21px;font-weight:600;margin-top:2px;font-variant-numeric:tabular-nums}
.kpi .sm{font-size:11px;color:var(--muted);margin-top:1px}
.bar{background:var(--panel);border:1px solid var(--border);border-top:0;padding:10px 14px;display:flex;gap:9px;flex-wrap:wrap;align-items:center}
.bar label{font-size:11px;color:var(--muted);margin-right:-4px}
.eyebrow{font-size:11px;letter-spacing:1.2px;text-transform:uppercase;color:var(--accent-teal);font-weight:600;margin:26px 0 7px}
.grid{display:grid;grid-template-columns:minmax(0,2.05fr) minmax(0,1fr);gap:14px}
@media(max-width:1080px){.grid{grid-template-columns:1fr}}
.card{background:var(--panel);border:1px solid var(--border);border-radius:5px;padding:14px 16px}
.card h3{margin:0 0 10px;font-size:13.5px;font-weight:600}
#map{height:620px;border:1px solid var(--border);border-radius:5px;background:var(--panel-alt)}
.leaflet-popup-content-wrapper{border-radius:4px}
.lgd{display:flex;gap:14px;flex-wrap:wrap;font-size:11.5px;color:var(--muted);margin-top:8px;align-items:center}
.sw{width:11px;height:11px;border-radius:2px;display:inline-block;margin-right:5px;vertical-align:-1px}
table{width:100%;border-collapse:collapse;font-size:12.5px}
th{text-align:left;font-weight:600;color:var(--muted);font-size:11px;text-transform:uppercase;letter-spacing:.4px;border-bottom:1px solid var(--border);padding:6px 6px}
td{padding:5px 6px;border-bottom:1px solid var(--border)}
td.n{text-align:right;font-variant-numeric:tabular-nums}
tr.clk:hover{background:var(--panel-alt);cursor:pointer}
.chartbox{position:relative;height:230px}
.det dt{font-size:10.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.4px;margin-top:8px}
.det dd{margin:1px 0 0;font-size:13px}
.pill{display:inline-block;background:var(--panel-alt);border:1px solid var(--border);border-radius:11px;padding:1px 9px;font-size:11px;margin:2px 3px 0 0}
.tier1{color:var(--status-green);font-weight:600}
.tier2{color:var(--status-amber);font-weight:600}
.tier3{color:var(--muted);font-weight:600}
.note{font-size:11.5px;color:var(--muted);line-height:1.6}
.warn{border-left:3px solid var(--status-amber);background:var(--panel);padding:11px 14px;font-size:12.5px;margin-top:12px}
footer{margin-top:32px;padding-top:14px;border-top:1px solid var(--border);font-size:11.5px;color:var(--muted);line-height:1.7}
.dark .leaflet-tile{filter:invert(1) hue-rotate(180deg) brightness(.92) contrast(.92)}
</style>
</head>
<body>
<header class="bnr"><div class="inner">
<div>
<h1>Parks expenditure map</h1>
<div class="sub">Park Services 4090240 and Water Parks 4090260 &middot; FY2023/24 to FY2026/27 &middot; site read from Register columns DY/DZ ({VER}: printed-line site allocation + electricity address crosswalk)</div>
</div>
<div style="display:flex;gap:9px;align-items:flex-end">
<div class="stamp" id="stamp"></div>
<button id="thm" title="Toggle dark mode">Dark</button>
</div>
</div></header>

<div class="wrap">
<div class="kpis" id="kpis"></div>

<div class="bar">
<label>FY</label><select id="fFy"><option value="">All years</option></select>
<label>Theme</label><select id="fTh"><option value="">All themes</option></select>
<label>Division</label><select id="fDv"><option value="">All divisions</option></select>
<label>Hierarchy</label><select id="fHi"><option value="">All</option></select>
<label>Function</label><select id="fFn"><option value="">All</option></select>
<label>Basis</label><select id="fBs"><option value="">All tiers</option><option value="1">Tier 1 only</option><option value="2">Tier 1 and 2</option></select>
<label>Min $</label><input type="range" id="fMin" min="0" max="100000" step="1000" value="0" style="width:110px">
<span id="minOut" style="font-size:11.5px;color:var(--muted);min-width:56px">$0</span>
<input type="search" id="fQ" placeholder="Search park..." style="min-width:150px">
<button id="reset">Reset</button>
<button id="csv">Export CSV</button>
</div>

<div class="eyebrow">01 &nbsp;Where the money went</div>
<div class="grid">
<div>
<div id="map"></div>
<div class="lgd" id="legend"></div>
<div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:9px">
<button id="mColour" class="on">Colour: theme</button>
<button id="mChoro">Division choropleth</button>
<button id="mLabels">Labels</button>
<span class="note" style="align-self:center" id="mapCount"></span>
</div>
</div>
<div>
<div class="card"><h3 id="detTitle">Select a park</h3><div id="detail"><p class="note">Click a bubble on the map, or a row in the table below, to see that park&rsquo;s spend, evidence basis, facilities and strategy attributes.</p></div></div>
<div class="card" style="margin-top:14px"><h3>Composition of the current selection</h3><div class="chartbox" style="height:210px"><canvas id="cTheme"></canvas></div></div>
</div>
</div>

<div class="eyebrow">02 &nbsp;Trend and division</div>
<div class="grid">
<div class="card"><h3>Sited expenditure by month</h3><div class="chartbox"><canvas id="cMon"></canvas></div>
<p class="note" style="margin-bottom:0">Only expenditure attributed to a named park is shown. Network services such as waste collection carry no site and are excluded from this chart by definition, not by filter.</p></div>
<div class="card"><h3>By division</h3><div class="chartbox"><canvas id="cDiv"></canvas></div>
<p class="note" style="margin-bottom:0">Division from the Parks Strategy Layer, joined on park name. Not normalised for park count or area.</p></div>
</div>

<div class="eyebrow">03 &nbsp;Park league table</div>
<div class="card">
<div style="max-height:430px;overflow:auto">
<table id="tbl"><thead><tr>
<th>Park</th><th>Div</th><th>Suburb</th><th>Hierarchy</th><th class="n">$ ex GST</th><th class="n">Lines</th><th class="n">Confirmed</th><th>Basis</th>
</tr></thead><tbody></tbody></table>
</div>
</div>

<div class="warn" id="honesty"></div>

<footer>
<strong>Sources and basis.</strong> Expenditure from the PS/WP Transaction Register (workbook version shown above), control total fixed and unchanged by this view. Site attribution is invoice-first: printed site details, printed line descriptions and printed works narratives are Tier 1; GL narration is Tier 2; parsed utility supply points are Tier 3. Park names, coordinates and facilities from the Logan City Council Parks Directory; division, suburb, hierarchy, function, embellishment and area from the Parks Strategy Layer (5-Nov-2025); division boundaries from LPS2015 Divisions. Coordinates are directory points where available, otherwise planning scheme polygon centroids.<br>
<strong>Limitations.</strong> A name match is a matching aid, not evidence, and does not upgrade an evidence tier. Parks absent from the directory and the planning scheme layer cannot be plotted. Multi-site invoices are attributed to the first named park pending line-level splitting. Cost per hectare is deliberately not shown: the three area columns in the strategy layer disagree by a factor of roughly 2.5 and the authoritative column is unresolved.<br>
<strong>Not for external circulation.</strong> Internal Parks Branch working document.
</footer>
</div>

<script src="https://cdnjs.cloudflare.com/ajax/libs/leaflet/1.9.4/leaflet.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<script>const PAYLOAD=__PAYLOAD__;</script>
<script>
const THEME_COL={T1:'#2a78d6',T10:'#eb6834',T2:'#1baf7a',T11:'#eda100',T6:'#e87ba4',T4:'#008300',T3:'#4a3aa7',T5:'#e34948',T12:'#8a6f3d',T7:'#6b7280',T8:'#9d5cb5',T13:'#5f5e5a',T9:'#9ca3af',Other:'#9ca3af'};
const THEME_NM={T1:'Grounds & asset maintenance',T2:'Playground safety & renewal',T3:'Electrical, lighting & data',T4:'Utilities',T5:'Cleaning & amenities',T6:'Internal charges',T7:'Fleet & travel',T8:'Materials & equipment',T9:'Accounting events',T10:'Waste collection',T11:'Plumbing & water assets',T12:'Security & access',T13:'Recoveries & offsets',Other:'Other'};
const $=id=>document.getElementById(id);
const fm=v=>'$'+Math.round(v).toLocaleString('en-AU');
const fmk=v=>Math.abs(v)>=1e6?'$'+(v/1e6).toFixed(2)+'M':Math.abs(v)>=1000?'$'+Math.round(v/1000)+'k':'$'+Math.round(v);
const P=PAYLOAD.parks, META=PAYLOAD.meta;

$('stamp').innerHTML='Register '+META.ver+' &middot; '+META.asat+'<br>'+META.lines.toLocaleString('en-AU')+' lines &middot; '+META.parks+' parks plotted';

// ---- populate filters
const uniq=(f)=>[...new Set(P.map(f).filter(x=>x!==null&&x!==undefined&&x!==''))].sort();
const allFy=[...new Set(P.flatMap(p=>Object.keys(p.fy)))].filter(x=>x&&x!=='nan'&&x!=='Unstated').sort();
allFy.forEach(v=>$('fFy').add(new Option(v,v)));
const allTh=[...new Set(P.flatMap(p=>Object.keys(p.th)))].sort((a,b)=>(+a.slice(1)||99)-(+b.slice(1)||99));
allTh.forEach(v=>$('fTh').add(new Option(v+' '+(THEME_NM[v]||''),v)));
uniq(p=>p.dv).sort((a,b)=>a-b).forEach(v=>$('fDv').add(new Option('Division '+v,v)));
uniq(p=>p.hi).forEach(v=>$('fHi').add(new Option(v,v)));
uniq(p=>p.fn).forEach(v=>$('fFn').add(new Option(v,v)));

// ---- filter state
const F={fy:'',th:'',dv:'',hi:'',fn:'',bs:'',min:0,q:''};
function val(p){
  let v=p.t;
  if(F.fy) v=p.fy[F.fy]||0;
  if(F.th) v=F.fy?0:(p.th[F.th]||0);
  if(F.fy&&F.th) v=0;
  return v;
}
function tier(p){ const b=p.bs||''; return /^(Sighted invoice|Rates parcel)/.test(b)?1:/^(GL narration|Enquiry)/.test(b)?2:3; }
function pass(p){
  if(F.dv&&String(p.dv)!==F.dv) return false;
  if(F.hi&&p.hi!==F.hi) return false;
  if(F.fn&&p.fn!==F.fn) return false;
  if(F.bs==='1'&&tier(p)!==1) return false;
  if(F.bs==='2'&&tier(p)>2) return false;
  if(F.q&&p.n.toLowerCase().indexOf(F.q.toLowerCase())<0) return false;
  if(F.fy&&!(p.fy[F.fy]>0)) return false;
  if(F.th&&!(p.th[F.th]>0)) return false;
  if(val(p)<F.min) return false;
  return true;
}
function sel(){ return P.filter(pass); }
function topTheme(p){ let b='Other',m=-1e12; for(const k in p.th){ if(k==='T9')continue; if(p.th[k]>m){m=p.th[k];b=k;} } return b; }

// ---- map
const map=L.map('map',{scrollWheelZoom:true}).setView([-27.72,153.07],11);
L.tileLayer('https://{s}.basemaps.cartocdn.com/light_all/{z}/{x}/{y}{r}.png',{
  attribution:'&copy; OpenStreetMap contributors &copy; CARTO',subdomains:'abcd',maxZoom:19}).addTo(map);
const bubbleLayer=L.layerGroup().addTo(map);
const choroLayer=L.geoJSON(null,{style:f=>choroStyle(f)});
let showChoro=false, colourMode='theme', showLabels=false;

const divTot=()=>{const o={};sel().forEach(p=>{if(p.dv!=null)o[p.dv]=(o[p.dv]||0)+val(p);});return o;};
function choroStyle(f){
  const t=divTot(), mx=Math.max(1,...Object.values(t)), v=t[f.properties.d]||0;
  const a=0.08+0.62*(v/mx);
  return {color:'#1F3864',weight:1.2,fillColor:'#0F766E',fillOpacity:showChoro?a:0,opacity:showChoro?0.85:0.35};
}
choroLayer.addData(PAYLOAD.divgeo);
choroLayer.addTo(map);
choroLayer.eachLayer(l=>{
  l.on('click',()=>{ $('fDv').value=String(l.feature.properties.d); F.dv=String(l.feature.properties.d); render(); });
  l.bindTooltip(()=>{const t=divTot();return '<b>Division '+l.feature.properties.d+'</b><br>'+fm(t[l.feature.properties.d]||0);},{sticky:true});
});

function radius(v,mx){ return Math.max(4,Math.min(30,4+26*Math.sqrt(Math.max(v,0)/Math.max(mx,1)))); }
function colourOf(p){
  if(colourMode==='theme') return THEME_COL[topTheme(p)]||'#9ca3af';
  const r=p.t>0?p.cf/p.t:0;
  return r>=0.9?'#047857':r>=0.6?'#C2410C':'#B91C1C';
}
function drawBubbles(){
  bubbleLayer.clearLayers();
  const S=sel(), mx=Math.max(1,...S.map(val));
  S.forEach(p=>{
    const v=val(p); if(v<=0) return;
    const m=L.circleMarker([p.lat,p.lon],{radius:radius(v,mx),color:colourOf(p),weight:1.4,
      fillColor:colourOf(p),fillOpacity:0.45,opacity:0.95});
    m.on('click',()=>showDetail(p));
    m.bindTooltip('<b>'+p.n+'</b><br>'+fm(v)+' &middot; '+p.l+' lines',{sticky:true});
    if(showLabels&&v>mx*0.05) m.bindTooltip(p.n,{permanent:true,direction:'right',className:'lbl'});
    bubbleLayer.addLayer(m);
  });
  $('mapCount').textContent=S.length+' parks shown, '+fm(S.reduce((a,p)=>a+val(p),0))+' attributed';
}

// ---- detail panel
function showDetail(p){
  $('detTitle').textContent=p.n;
  const th=Object.entries(p.th).sort((a,b)=>b[1]-a[1]);
  const fy=Object.entries(p.fy).sort();
  const t=tier(p), tc=t===1?'tier1':t===2?'tier2':'tier3';
  let h='<dl class="det" style="margin:0">';
  h+='<dt>Total attributed</dt><dd style="font-size:19px;font-weight:600">'+fm(p.t)+' <span style="font-size:12px;font-weight:400;color:var(--muted)">across '+p.l+' lines</span></dd>';
  h+='<dt>Evidence basis</dt><dd class="'+tc+'">'+p.bs+'</dd>';
  h+='<dt>Confirmed share</dt><dd>'+fm(p.cf)+' ('+(p.t>0?Math.round(p.cf/p.t*100):0)+'% of this park&rsquo;s dollars)</dd>';
  h+='<dt>Themes</dt><dd>'+th.map(([k,v])=>'<span class="pill" style="border-color:'+(THEME_COL[k]||'#999')+'">'+k+' '+fmk(v)+'</span>').join('')+'</dd>';
  h+='<dt>By year</dt><dd>'+fy.map(([k,v])=>'<span class="pill">'+k+' '+fmk(v)+'</span>').join('')+'</dd>';
  h+='<dt>Top vendors</dt><dd>'+(p.vd.length?p.vd.map(v=>v[0]+' &mdash; '+fmk(v[1])).join('<br>'):'&mdash;')+'</dd>';
  h+='<dt>Strategy attributes</dt><dd>'+[p.dv!=null?'Division '+p.dv:null,p.sb,p.hi,p.fn,p.em!=null?'Embellishment '+p.em:null,p.ar?(p.ar/10000).toFixed(2)+' ha':null].filter(Boolean).join(' &middot; ')+'</dd>';
  h+='<dt>Facilities listed</dt><dd>'+(p.fc?p.fc.split(',').map(f=>'<span class="pill">'+f.replace(/_/g,' ')+'</span>').join(''):'<span class="note">Not in the public parks directory</span>')+'</dd>';
  h+='<dt>Coordinate source</dt><dd class="note">'+(p.cs==='point'?'Parks directory point':'Planning scheme polygon centroid')+'</dd>';
  h+='</dl>';
  $('detail').innerHTML=h;
  drawTheme(p);
  map.setView([p.lat,p.lon],Math.max(map.getZoom(),14));
}

// ---- charts
let chTheme,chMon,chDiv;
function drawTheme(p){
  const src=p?p.th:sel().reduce((a,q)=>{for(const k in q.th)a[k]=(a[k]||0)+q.th[k];return a;},{});
  const e=Object.entries(src).filter(x=>x[1]>0).sort((a,b)=>b[1]-a[1]);
  const d={labels:e.map(x=>x[0]),datasets:[{data:e.map(x=>x[1]),backgroundColor:e.map(x=>THEME_COL[x[0]]||'#9ca3af'),borderWidth:0}]};
  if(chTheme){chTheme.data=d;chTheme.update();return;}
  chTheme=new Chart($('cTheme'),{type:'doughnut',data:d,options:{responsive:true,maintainAspectRatio:false,
    plugins:{legend:{position:'right',labels:{boxWidth:10,font:{size:11},color:cssv('--muted')}},
    tooltip:{callbacks:{label:c=>c.label+' '+(THEME_NM[c.label]||'')+': '+fm(c.parsed)}}}}});
}
const cssv=n=>getComputedStyle(document.body).getPropertyValue(n)||'#666';
function drawMon(){
  const m=PAYLOAD.mon, ks=Object.keys(m).sort();
  const d={labels:ks,datasets:[{label:'Sited spend',data:ks.map(k=>m[k]),borderColor:'#0F766E',backgroundColor:'rgba(15,118,110,.13)',fill:true,tension:.28,pointRadius:0,borderWidth:2}]};
  if(chMon){chMon.data=d;chMon.update();return;}
  chMon=new Chart($('cMon'),{type:'line',data:d,options:{responsive:true,maintainAspectRatio:false,
    plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fm(c.parsed.y)}}},
    scales:{x:{ticks:{maxTicksLimit:9,color:cssv('--muted'),font:{size:10}},grid:{display:false}},
            y:{ticks:{callback:v=>fmk(v),color:cssv('--muted'),font:{size:10}},grid:{color:'rgba(128,128,128,.16)'}}}}});
}
function drawDiv(){
  const t=divTot(), ks=Object.keys(t).sort((a,b)=>t[b]-t[a]);
  const d={labels:ks.map(k=>'Div '+k),datasets:[{data:ks.map(k=>t[k]),backgroundColor:'#1F3864',borderRadius:3,barThickness:15}]};
  if(chDiv){chDiv.data=d;chDiv.update();return;}
  chDiv=new Chart($('cDiv'),{type:'bar',data:d,options:{indexAxis:'y',responsive:true,maintainAspectRatio:false,
    plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>fm(c.parsed.x)}}},
    scales:{x:{ticks:{callback:v=>fmk(v),color:cssv('--muted'),font:{size:10}},grid:{color:'rgba(128,128,128,.16)'}},
            y:{ticks:{color:cssv('--muted'),font:{size:10}},grid:{display:false}}}}});
}

// ---- table + kpis
function drawTable(){
  const S=sel().slice().sort((a,b)=>val(b)-val(a));
  const tb=$('tbl').tBodies[0]; tb.innerHTML='';
  S.forEach(p=>{
    const t=tier(p), tc=t===1?'tier1':t===2?'tier2':'tier3';
    const tr=document.createElement('tr'); tr.className='clk';
    tr.innerHTML='<td>'+p.n+'</td><td class="n">'+(p.dv??'')+'</td><td>'+(p.sb||'')+'</td><td>'+(p.hi||'')+'</td>'+
      '<td class="n">'+fm(val(p))+'</td><td class="n">'+p.l+'</td><td class="n">'+(p.t>0?Math.round(p.cf/p.t*100):0)+'%</td>'+
      '<td class="'+tc+'" title="'+p.bs+'">'+p.bs.replace(/ \(printed\)/,'')+'</td>';
    tr.onclick=()=>showDetail(p); tb.appendChild(tr);
  });
}
function drawKpis(){
  const S=sel(), v=S.reduce((a,p)=>a+val(p),0), l=S.reduce((a,p)=>a+p.l,0);
  const cf=S.reduce((a,p)=>a+p.cf,0), t1=S.filter(p=>tier(p)===1).reduce((a,p)=>a+val(p),0);
  $('kpis').innerHTML=[
    ['Register control total',fmk(META.total),'fixed, never moved by this view'],
    ['Named park',fmk(META.named),META.pct+'% of register, '+META.parks_total+' parks (incl. '+fmk(META.alloc)+' allocated per printed line)'],
    ['Multi-site held',fmk(META.multi),'named several parks, never allocated'],
    ['Plotted on map',fmk(META.plotted),META.parks+' parks with coordinates'],
    ['Current selection',fmk(v),S.length+' parks, '+l.toLocaleString('en-AU')+' lines'],
    ['Confirmed status',fmk(cf),(v>0?Math.round(cf/Math.max(v,1)*100):0)+'% of selection']
  ].map(k=>'<div class="kpi"><div class="lbl">'+k[0]+'</div><div class="val">'+k[1]+'</div><div class="sm">'+k[2]+'</div></div>').join('');
}
function drawLegend(){
  if(colourMode==='theme'){
    $('legend').innerHTML='<span style="color:var(--muted)">Bubble colour = dominant theme, size = dollars:</span>'+
      allTh.filter(k=>k!=='T9').map(k=>'<span><span class="sw" style="background:'+(THEME_COL[k]||'#999')+'"></span>'+k+'</span>').join('');
  } else {
    $('legend').innerHTML='<span style="color:var(--muted)">Bubble colour = share of the park&rsquo;s dollars confirmed by a sighted invoice:</span>'+
      '<span><span class="sw" style="background:#047857"></span>90% and over</span>'+
      '<span><span class="sw" style="background:#C2410C"></span>60 to 89%</span>'+
      '<span><span class="sw" style="background:#B91C1C"></span>under 60%</span>';
  }
}
function drawHonesty(){
  let h='<strong>Read this before quoting any figure.</strong> Every register line carries a site basis (v43). '+fmk(META.named)+' ('+META.pct+'%) names exactly one park; '+fmk(META.multi)+' names several and is held unallocated; the rest is classed below. This map shows where spend <em>is known</em> to have landed, never where it did not.';
  h+='<table style="margin-top:10px;max-width:760px"><thead><tr><th>Site basis class (Register column DZ)</th><th class="n">Lines</th><th class="n">$ ex GST</th></tr></thead><tbody>';
  PAYLOAD.classes.forEach(c=>{h+='<tr><td>'+c.k+'</td><td class="n">'+c.l.toLocaleString('en-AU')+'</td><td class="n">'+fm(c.d)+'</td></tr>';});
  h+='</tbody></table>';
  if(PAYLOAD.unplaced.length){h+='<div style="margin-top:10px"><strong>Named parks with no coordinate yet ('+META.unplaced+'):</strong> <span class="note">'+PAYLOAD.unplaced.slice(0,25).map(u=>u[0]+' '+fmk(u[1])).join(' &middot; ')+(PAYLOAD.unplaced.length>25?' &middot; ...':'')+'</span></div>';}
  $('honesty').innerHTML=h;
}

function render(){ choroLayer.setStyle(choroStyle); drawBubbles(); drawKpis(); drawTable(); drawTheme(null); drawDiv(); drawLegend(); drawHonesty(); }

// ---- wiring
['fFy','fTh','fDv','fHi','fFn','fBs'].forEach(id=>$(id).onchange=e=>{F[id.slice(1).toLowerCase()]=e.target.value;render();});
$('fMin').oninput=e=>{F.min=+e.target.value;$('minOut').textContent=fmk(F.min);render();};
$('fQ').oninput=e=>{F.q=e.target.value;render();};
$('reset').onclick=()=>{Object.assign(F,{fy:'',th:'',dv:'',hi:'',fn:'',bs:'',min:0,q:''});
  ['fFy','fTh','fDv','fHi','fFn','fBs','fQ'].forEach(i=>$(i).value='');$('fMin').value=0;$('minOut').textContent='$0';
  map.setView([-27.72,153.07],11);render();};
$('mColour').onclick=e=>{colourMode=colourMode==='theme'?'conf':'theme';e.target.textContent='Colour: '+(colourMode==='theme'?'theme':'confidence');drawBubbles();drawLegend();};
$('mChoro').onclick=e=>{showChoro=!showChoro;e.target.classList.toggle('on',showChoro);choroLayer.setStyle(choroStyle);};
$('mLabels').onclick=e=>{showLabels=!showLabels;e.target.classList.toggle('on',showLabels);drawBubbles();};
$('csv').onclick=()=>{
  const rows=[['Park','Division','Suburb','Hierarchy','Function','Dollars ex GST','Lines','Confirmed dollars','Evidence basis','Latitude','Longitude']];
  sel().sort((a,b)=>val(b)-val(a)).forEach(p=>rows.push([p.n,p.dv??'',p.sb||'',p.hi||'',p.fn||'',val(p).toFixed(2),p.l,p.cf.toFixed(2),p.bs,p.lat,p.lon]));
  const csv=rows.map(r=>r.map(c=>'"'+String(c).replace(/"/g,'""')+'"').join(',')).join('\n');
  const a=document.createElement('a');a.href=URL.createObjectURL(new Blob([csv],{type:'text/csv'}));
  a.download='parks_expenditure_selection.csv';a.click();};
$('thm').onclick=e=>{document.body.classList.toggle('dark');
  e.target.textContent=document.body.classList.contains('dark')?'Light':'Dark';
  [chTheme,chMon,chDiv].forEach(c=>{if(!c)return;
    if(c.options.scales){for(const k in c.options.scales)c.options.scales[k].ticks.color=cssv('--muted');}
    if(c.options.plugins.legend.labels)c.options.plugins.legend.labels.color=cssv('--muted');c.update();});};

drawMon(); render();
</script>
</body>
</html>
"""

# ----------------------------------------------------------------- CLI
ap = argparse.ArgumentParser()
ap.add_argument('workbook')
ap.add_argument('--gazetteer', default='logan_park_gazetteer.csv')
ap.add_argument('--divisions', default='logan_divisions.geojson')
ap.add_argument('--out', default=None)
ap.add_argument('--simplify', type=float, default=0.0008, help='division polygon simplify tolerance (degrees)')
A = ap.parse_args()

# ----------------------------------------------------------------- 1. positions from Config
cw = CalamineWorkbook.from_path(A.workbook)
cfg = {str(r[0]): r[1] for r in cw.get_sheet_by_name('Config').to_python() if r and r[0]}
VER = str(cfg.get('WORKBOOK_VERSION', 'v?'))
OUT = A.out or f'LCC_Parks_Expenditure_Map_{VER}.html'
CT  = float(cfg['CONTROL_TOTAL'])
SITE_COL  = int(str(cfg['SITE_COL']).split()[0])          # e.g. "129 (DY)"
BASIS_COL = int(str(cfg['SITE_BASIS_COL']).split()[0])
PA0, PA1  = [int(x) for x in str(cfg['SITES_PANELA']).split(':')]
GRP0, GRP1 = [int(x) for x in str(cfg['SITES_GROUP_ROWS']).split(':')]

# the closed-list DZ class labels are the Sites group rows (schema section 44)
sites = cw.get_sheet_by_name('Sites').to_python()
CLASS_LABELS = [str(r[0]) for r in sites[GRP0-1:GRP1] if r and r[0]]
MULTI_ALLOC = next((c for c in CLASS_LABELS if 'allocated per printed line' in c), None)
# the allocated-multi basis is NOT a Sites group row (the residue row stands for it), so add it explicitly
# or the register row AND its Site_Allocation rows both count -> double count
RESIDUE     = next((c for c in CLASS_LABELS if c.startswith('(residue')), '(residue')
CL = set(CLASS_LABELS)

# ----------------------------------------------------------------- 2. register
df = pd.read_excel(A.workbook, sheet_name='Register', header=3, engine='calamine')
df = df[df.iloc[:, 0].notna()]
amt = pd.to_numeric(df.iloc[:, 19], errors='coerce').fillna(0)
o = pd.DataFrame(dict(
    lk     = df.iloc[:, 0].astype(str),
    site   = df.iloc[:, SITE_COL-1].fillna('').astype(str),
    basis  = df.iloc[:, BASIS_COL-1].fillna('').astype(str),
    amt    = amt,
    fy     = df.iloc[:, 4].astype(str),
    theme  = df.iloc[:, 25].astype(str),
    pk     = df.iloc[:, 16].astype(str),
    vendor = df.iloc[:, 11].fillna('Unidentified').astype(str),
    status = df.iloc[:, 32].astype(str),
    date   = pd.to_datetime(df.iloc[:, 3], errors='coerce')))
print(f'register {len(o):,} lines, ${o.amt.sum():,.2f}')
assert abs(o.amt.sum() - CT) < 0.02, 'register does not tie to Config CONTROL_TOTAL'

# single-park rows. Exclude every class label, the allocated-multi basis, and any joined site list.
if MULTI_ALLOC is None:
    MULTI_ALLOC = next((b for b in o.basis.unique() if 'allocated per printed line' in b), '~none~')
CL.add(MULTI_ALLOC)
s = o[(o.site != '') & (~o.basis.isin(CL)) & (~o.site.str.contains('; ', regex=False))].copy()
assert not s.site.str.contains('; ', regex=False).any()

# ----------------------------------------------------------------- 3. Site_Allocation (multi-site, per printed line)
alloc = pd.DataFrame(columns=['site', 'basis', 'amt', 'fy', 'theme', 'pk', 'vendor', 'status', 'date'])
if 'Site_Allocation' in cw.sheet_names:
    rows = [r for r in cw.get_sheet_by_name('Site_Allocation').to_python()[4:]
            if r and r[0] not in (None, '') and isinstance(r[1], (int, float)) and isinstance(r[7], (int, float))]
    rows = [r for r in rows if str(r[6]) != RESIDUE]
    if rows:
        idx = o.set_index('lk')
        idx = idx[~idx.index.duplicated()]
        keys = [str(r[0]) for r in rows]
        alloc = pd.DataFrame(dict(
            site   = [str(r[6]) for r in rows],
            basis  = 'Site_Allocation (printed line amounts)',
            amt    = [float(r[7]) for r in rows],
            fy     = [str(r[3]) for r in rows],
            theme  = idx.loc[keys, 'theme'].values,
            pk     = [str(r[4]) for r in rows],
            vendor = idx.loc[keys, 'vendor'].values,
            status = idx.loc[keys, 'status'].values,
            date   = idx.loc[keys, 'date'].values))
s = pd.concat([s.drop(columns=['lk']), alloc], ignore_index=True)
s['th']  = s.theme.str.split(' ').str[0]
s['mon'] = pd.to_datetime(s.date, errors='coerce').dt.to_period('M').astype(str)
print(f'named-park rows {len(s):,}, ${s.amt.sum():,.2f}')

# ----------------------------------------------------------------- 4. gazetteer join
g = pd.read_csv(A.gazetteer)
g['base'] = g.name.str.replace(r'\s*\(.*\)', '', regex=True).str.strip()
gg = g.drop_duplicates('base').set_index('base')
low = {b.lower(): b for b in gg.index}
def coord(name):
    for k in (str(name).lower(), re.sub(r'\s*\(.*\)', '', str(name)).strip().lower()):
        if k in low:
            return gg.loc[low[k]]
    return None

parks, unplaced = [], []
for name, d in s.groupby('site'):
    r = coord(name)
    if r is None or pd.isna(r.lat):
        unplaced.append((name, round(float(d.amt.sum())), len(d))); continue
    th   = d.groupby('th').amt.sum().round(0)
    fy   = d.groupby('fy').amt.sum().round(0)
    vend = d.groupby('vendor').amt.sum().sort_values(ascending=False).head(4).round(0)
    bas  = d.basis.value_counts()
    parks.append(dict(
        n=name, lat=round(float(r.lat), 5), lon=round(float(r.lon), 5),
        t=round(float(d.amt.sum())), l=int(len(d)),
        cf=round(float(d[d.status == 'Confirmed'].amt.sum())),
        t1=round(float(d[d.basis.str.contains('Sighted invoice|Rates parcel|Site_Allocation|Printed line', na=False)].amt.sum())),
        th={k: float(v) for k, v in th.items()},
        fy={k: float(v) for k, v in fy.items()},
        vd=[[k, float(v)] for k, v in vend.items()],
        bs=bas.index[0], bsn=int(bas.iloc[0]),
        dv=None if pd.isna(r.division)     else int(r.division),
        sb=None if pd.isna(r.suburb)       else str(r.suburb).title(),
        hi=None if pd.isna(r.hierarchy)    else str(r.hierarchy),
        fn=None if pd.isna(r.function)     else str(r.function),
        em=None if pd.isna(r.embellishment) else str(r.embellishment),
        ar=None if pd.isna(r.area_sqm)     else round(float(r.area_sqm)),
        fc=None if pd.isna(r.facilities)   else str(r.facilities),
        cs='point' if r.coord_source == 'directory point' else 'centroid'))
parks.sort(key=lambda x: -x['t'])
print(f'plotted {len(parks)} parks, unplaced {len(unplaced)}')

# ----------------------------------------------------------------- 5. class table, divisions, months
cls = o.groupby(o.basis.where(o.basis.isin(CL) | o.basis.str.contains('USER CHECK', na=False), 'Named park')).amt.agg(['sum', 'size'])
classes = [dict(k=k, d=round(float(v['sum'])), l=int(v['size'])) for k, v in cls.sort_values('sum', ascending=False).iterrows()]
divmap = {p['n']: p['dv'] for p in parks}
div = {}
sd = s.assign(dv=s.site.map(divmap)).dropna(subset=['dv'])
for k, v in sd.groupby('dv'):
    div[str(int(k))] = dict(t=round(float(v.amt.sum())), l=int(len(v)), p=int(v.site.nunique()))
mon = {m: round(float(v)) for m, v in s.groupby('mon').amt.sum().items() if m not in ('NaT', 'nan')}

feats = []
for f in json.load(open(A.divisions))['features']:
    geo = shape(f['geometry']).simplify(A.simplify, preserve_topology=True)
    dn = next((int(v) for k, v in f['properties'].items() if 'div' in k.lower() and str(v).strip().isdigit()), None)
    feats.append(dict(type='Feature', properties=dict(d=dn), geometry=mapping(geo)))

named   = float(s[s.basis != 'Site_Allocation (printed line amounts)'].amt.sum()) + float(alloc.amt.sum() if len(alloc) else 0)
multi   = float(o[o.basis == 'Multiple sites named (not allocated)'].amt.sum())
plotted = sum(p['t'] for p in parks)
meta = dict(total=round(CT), named=round(float(s.amt.sum())), alloc=round(float(alloc.amt.sum()) if len(alloc) else 0),
            plotted=round(plotted), multi=round(multi), pct=round(float(s.amt.sum()) / CT * 100, 1),
            lines=int(len(o)), namedlines=int(len(s)), parks=len(parks), parks_total=int(s.site.nunique()),
            unplaced=len(unplaced), ver=VER, asat=pd.Timestamp.today().strftime('%-d-%b-%Y'))
payload = dict(meta=meta, parks=parks, div=div,
               divgeo=dict(type='FeatureCollection', features=feats), mon=mon, classes=classes,
               unplaced=sorted(unplaced, key=lambda x: -x[1])[:40])

# ----------------------------------------------------------------- 6. emit
html = SHELL.replace('{VER}', VER).replace('__PAYLOAD__', json.dumps(payload, separators=(',', ':')))
open(OUT, 'w').write(html)
print(f'wrote {OUT}  ({len(html)/1024:.0f} KB)  meta={meta}')
