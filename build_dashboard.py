#!/usr/bin/env python3
"""
Purchase Order HTML Dashboard Builder
Reads JSON from process_data.py and writes a standalone HTML file.
Usage: python3 build_dashboard.py --data dashboard_data.json --output dashboard.html
"""

import json
import argparse
from pathlib import Path

HTML_TEMPLATE = '''<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Purchase Order Dashboard</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.js"></script>
<link href="https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
*,*::before,*::after{box-sizing:border-box;margin:0;padding:0;}
:root{
  --bg:#060810;--s1:#0d1120;--s2:#111827;--s3:#1a2235;
  --border:rgba(255,255,255,0.06);--border2:rgba(255,255,255,0.12);
  --text:#e2e8f6;--muted:#5a6a8a;--muted2:#8896b0;
  --blue:#3d7eff;--blue2:#6fa3ff;--blue-bg:rgba(61,126,255,0.08);
  --teal:#00d4aa;--teal-bg:rgba(0,212,170,0.08);
  --orange:#ff7c3d;--purple:#a855f7;--red:#ff4d6d;
  --green:#22c55e;--amber:#f59e0b;
  --grid:rgba(255,255,255,0.03);
}
body{font-family:"Space Grotesk",sans-serif;background:var(--bg);color:var(--text);min-height:100vh;font-size:13px;}
nav{background:var(--s1);border-bottom:1px solid var(--border);padding:0 28px;display:flex;align-items:center;justify-content:space-between;height:52px;position:sticky;top:0;z-index:200;}
.nav-brand{display:flex;align-items:center;gap:10px;}
.nav-pulse{width:8px;height:8px;border-radius:50%;background:var(--blue);box-shadow:0 0 0 3px rgba(61,126,255,0.25),0 0 12px var(--blue);animation:pulse 2s infinite;}
@keyframes pulse{0%,100%{opacity:1;}50%{opacity:0.5;}}
.nav-title{font-size:14px;font-weight:600;letter-spacing:-0.3px;}
.nav-tabs{display:flex;gap:2px;}
.nav-tab{padding:6px 14px;border-radius:6px;font-size:12px;font-weight:500;cursor:pointer;color:var(--muted2);border:none;background:transparent;transition:all 0.15s;}
.nav-tab.active{background:var(--blue-bg);color:var(--blue);border:1px solid rgba(61,126,255,0.2);}
.nav-tab:hover:not(.active){color:var(--text);background:var(--s3);}
.nav-right{font-size:11px;color:var(--muted);font-family:"JetBrains Mono",monospace;display:flex;gap:10px;}
.nav-badge{background:var(--s3);border:1px solid var(--border2);border-radius:20px;padding:3px 10px;color:var(--muted2);}
.main{padding:24px 28px;max-width:1440px;margin:0 auto;}
.page{display:none;}.page.active{display:block;}
.kpi-row{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:20px;}
.kpi{background:var(--s1);border:1px solid var(--border);border-radius:10px;padding:16px 18px;position:relative;overflow:hidden;}
.kpi-accent{position:absolute;top:0;left:0;right:0;height:2px;}
.kpi-label{font-size:10px;font-weight:500;color:var(--muted2);text-transform:uppercase;letter-spacing:0.8px;margin-bottom:8px;}
.kpi-val{font-size:24px;font-weight:700;letter-spacing:-0.8px;line-height:1;}
.kpi-sub{font-size:10px;color:var(--muted);margin-top:5px;}
.section-hdr{display:flex;align-items:center;gap:10px;margin:20px 0 12px;}
.section-hdr h2{font-size:11px;font-weight:600;text-transform:uppercase;letter-spacing:1px;color:var(--muted2);}
.section-hdr::after{content:"";flex:1;height:1px;background:var(--border);}
.g2{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px;}
.g3{display:grid;grid-template-columns:2fr 1fr;gap:12px;margin-bottom:12px;}
.card{background:var(--s1);border:1px solid var(--border);border-radius:10px;padding:18px;}
.card-hdr{font-size:12px;font-weight:600;color:var(--text);margin-bottom:3px;}
.card-sub{font-size:10px;color:var(--muted);margin-bottom:14px;}
.chart-box{position:relative;width:100%;}
.ctrl-row{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:14px;}
.ctrl-input{background:var(--s2);border:1px solid var(--border2);color:var(--text);font-family:"Space Grotesk",sans-serif;font-size:12px;padding:7px 12px;border-radius:7px;outline:none;}
.ctrl-input:focus{border-color:var(--blue);}
.ctrl-select{background:var(--s2);border:1px solid var(--border2);color:var(--text);font-family:"Space Grotesk",sans-serif;font-size:12px;padding:7px 12px;border-radius:7px;outline:none;cursor:pointer;}
.tog-grp{display:flex;background:var(--s2);border-radius:7px;padding:3px;gap:2px;border:1px solid var(--border);}
.tog{padding:5px 12px;border-radius:5px;border:none;cursor:pointer;font-size:11px;font-weight:500;color:var(--muted2);background:transparent;font-family:"Space Grotesk",sans-serif;transition:all 0.15s;}
.tog.on{background:var(--blue);color:#fff;}
.btn{padding:7px 14px;border-radius:7px;border:none;cursor:pointer;font-family:"Space Grotesk",sans-serif;font-size:12px;font-weight:500;}
.btn-blue{background:var(--blue);color:#fff;}
.btn-ghost{background:var(--s3);color:var(--muted2);border:1px solid var(--border2);}
.insight{background:rgba(61,126,255,0.06);border:1px solid rgba(61,126,255,0.18);border-radius:8px;padding:10px 14px;font-size:11px;color:var(--muted2);margin-bottom:14px;line-height:1.6;}
.insight strong{color:var(--text);}
.mini-kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:8px;margin-bottom:14px;}
.mini-kpi{background:var(--s2);border:1px solid var(--border);border-radius:8px;padding:10px 12px;}
.mini-kpi-lbl{font-size:9px;color:var(--muted);text-transform:uppercase;letter-spacing:0.6px;margin-bottom:3px;}
.mini-kpi-val{font-size:16px;font-weight:700;letter-spacing:-0.5px;}
.plan-table-wrap{overflow:auto;max-height:520px;border-radius:8px;border:1px solid var(--border);}
table{width:100%;border-collapse:collapse;font-size:11px;}
thead{position:sticky;top:0;z-index:10;}
th{background:var(--s2);color:var(--muted2);font-weight:600;font-size:10px;text-transform:uppercase;letter-spacing:0.5px;padding:10px 12px;text-align:left;border-bottom:1px solid var(--border);}
td{padding:9px 12px;border-bottom:1px solid var(--border);color:var(--text);}
tr:last-child td{border-bottom:none;}
tr:hover td{background:rgba(255,255,255,0.02);}
.editable{background:var(--s2);border:1px solid var(--border2);border-radius:5px;color:var(--text);font-family:"JetBrains Mono",monospace;font-size:11px;padding:3px 7px;width:90px;outline:none;text-align:right;}
.editable:focus{border-color:var(--blue);}
.badge{display:inline-block;padding:2px 7px;border-radius:12px;font-size:9px;font-weight:600;}
.badge-up{background:rgba(34,197,94,0.15);color:var(--green);}
.badge-down{background:rgba(255,77,109,0.15);color:var(--red);}
.badge-stable{background:rgba(90,106,138,0.18);color:var(--muted2);}
.total-row td{background:var(--s3);font-weight:600;color:var(--blue);}
.legend{display:flex;flex-wrap:wrap;gap:10px;margin-bottom:10px;}
.leg{display:flex;align-items:center;gap:5px;font-size:10px;color:var(--muted2);}
.leg-dot{width:8px;height:8px;border-radius:2px;flex-shrink:0;}
::-webkit-scrollbar{width:5px;height:5px;}
::-webkit-scrollbar-thumb{background:var(--s3);border-radius:3px;}
footer{text-align:center;padding:20px;color:var(--muted);font-family:"JetBrains Mono",monospace;font-size:10px;border-top:1px solid var(--border);margin-top:24px;}
@media(max-width:900px){.kpi-row{grid-template-columns:repeat(2,1fr);}.g2,.g3{grid-template-columns:1fr;}.main{padding:14px;}}
</style>
</head>
<body>
<nav>
  <div class="nav-brand"><div class="nav-pulse"></div><span class="nav-title">Purchase Analytics</span></div>
  <div class="nav-tabs">
    <button class="nav-tab active" onclick="showPage(\'overview\',this)">Overview</button>
    <button class="nav-tab" onclick="showPage(\'products\',this)">All Products</button>
    <button class="nav-tab" onclick="showPage(\'usage\',this)">Usage Charts</button>
    <button class="nav-tab" onclick="showPage(\'planning\',this)">Order Planning</button>
  </div>
  <div class="nav-right">
    <span class="nav-badge">{{DATE_RANGE}}</span>
    <span class="nav-badge">{{UNIQUE_PRODUCTS}} products</span>
    <span class="nav-badge">{{UNIQUE_SUPPLIERS}} suppliers</span>
  </div>
</nav>
<div class="main">

<!-- OVERVIEW -->
<div id="page-overview" class="page active">
  <div class="kpi-row">
    <div class="kpi"><div class="kpi-accent" style="background:var(--blue)"></div><div class="kpi-label">Total Spend</div><div class="kpi-val" style="color:var(--blue)">{{TOTAL_SPEND_FMT}}</div><div class="kpi-sub">{{DATE_RANGE}}</div></div>
    <div class="kpi"><div class="kpi-accent" style="background:var(--teal)"></div><div class="kpi-label">Total Orders</div><div class="kpi-val" style="color:var(--teal)">{{TOTAL_ORDERS}}</div><div class="kpi-sub">Order lines</div></div>
    <div class="kpi"><div class="kpi-accent" style="background:var(--orange)"></div><div class="kpi-label">Avg Order</div><div class="kpi-val" style="color:var(--orange)">{{AVG_ORDER_FMT}}</div><div class="kpi-sub">Per order line</div></div>
    <div class="kpi"><div class="kpi-accent" style="background:var(--purple)"></div><div class="kpi-label">Products</div><div class="kpi-val" style="color:var(--purple)">{{UNIQUE_PRODUCTS}}</div><div class="kpi-sub">Unique SKUs</div></div>
    <div class="kpi"><div class="kpi-accent" style="background:var(--red)"></div><div class="kpi-label">Suppliers</div><div class="kpi-val" style="color:var(--red)">{{UNIQUE_SUPPLIERS}}</div><div class="kpi-sub">Active vendors</div></div>
  </div>
  <div class="section-hdr"><h2>Spend Overview</h2></div>
  <div class="g2">
    <div class="card"><div class="card-hdr">Daily spend trend</div><div class="card-sub">Purchase activity over time</div><div class="chart-box" style="height:200px;"><canvas id="c-daily"></canvas></div></div>
    <div class="card"><div class="card-hdr">Monthly spend</div><div class="card-sub">Total per month</div><div class="chart-box" style="height:200px;"><canvas id="c-monthly"></canvas></div></div>
  </div>
  <div class="section-hdr"><h2>Product Analysis</h2></div>
  <div class="card" style="margin-bottom:12px;"><div class="card-hdr">Top 20 products by spend</div><div class="card-sub">All products available in Products tab</div><div class="chart-box" style="height:320px;"><canvas id="c-topspend"></canvas></div></div>
  <div class="g2">
    <div class="card"><div class="card-hdr">Top 15 by quantity</div><div class="card-sub">Volume leaders</div><div class="chart-box" style="height:260px;"><canvas id="c-topqty"></canvas></div></div>
    <div class="card"><div class="card-hdr">Order frequency</div><div class="card-sub">Times reordered — set standing schedules for top items</div><div class="chart-box" style="height:260px;"><canvas id="c-freq"></canvas></div></div>
  </div>
  <div class="g2">
    <div class="card"><div class="card-hdr">Spend by category</div><div class="card-sub">By product code prefix</div><div class="chart-box" style="height:220px;"><canvas id="c-cat"></canvas></div></div>
    <div class="card"><div class="card-hdr">Supplier breakdown</div><div class="card-sub">Share of spend per supplier</div><div class="legend" id="sup-legend"></div><div class="chart-box" style="height:170px;"><canvas id="c-sup"></canvas></div></div>
  </div>
</div>

<!-- PRODUCTS -->
<div id="page-products" class="page">
  <div class="section-hdr"><h2>All Products — Monthly Comparison</h2></div>
  <div class="ctrl-row">
    <input type="text" class="ctrl-input" id="prod-search" placeholder="Search products..." oninput="filterProducts()" style="flex:1;min-width:200px;max-width:340px;">
    <select class="ctrl-select" id="prod-sort" onchange="filterProducts()">
      <option value="qty">Sort: Highest Qty</option>
      <option value="spend">Sort: Highest Spend</option>
      <option value="name">Sort: Name A–Z</option>
      <option value="trend_up">Trending Up ↑</option>
      <option value="trend_down">Trending Down ↓</option>
    </select>
    <div class="tog-grp">
      <button class="tog on" onclick="setProdMode(\'qty\',this)">Quantity</button>
      <button class="tog" onclick="setProdMode(\'spend\',this)">Spend</button>
    </div>
    <span style="font-size:11px;color:var(--muted);" id="prod-count"></span>
  </div>
  <div style="overflow:auto;max-height:480px;border:1px solid var(--border);border-radius:10px;margin-bottom:20px;">
    <table id="products-table">
      <thead><tr>
        <th>Product</th><th style="text-align:center">Code</th>
        {{MONTH_HEADERS}}
        <th style="text-align:center">Total</th>
        <th style="text-align:center">Avg/Month</th>
        <th style="text-align:center">Trend</th>
      </tr></thead>
      <tbody id="products-tbody"></tbody>
    </table>
  </div>
  <div class="section-hdr"><h2>Product Deep Dive</h2></div>
  <div class="card">
    <div class="ctrl-row">
      <select class="ctrl-select" id="deep-select" style="flex:1;max-width:400px;" onchange="updateDeep()"></select>
      <div class="tog-grp">
        <button class="tog on" onclick="setDeepMode(\'qty\',this)">Quantity</button>
        <button class="tog" onclick="setDeepMode(\'spend\',this)">Spend</button>
      </div>
    </div>
    <div class="mini-kpis">
      <div class="mini-kpi"><div class="mini-kpi-lbl">Total Qty</div><div class="mini-kpi-val" id="dk1">—</div></div>
      <div class="mini-kpi"><div class="mini-kpi-lbl">Total Spend</div><div class="mini-kpi-val" id="dk2">—</div></div>
      <div class="mini-kpi"><div class="mini-kpi-lbl">Avg Price</div><div class="mini-kpi-val" id="dk3">—</div></div>
      <div class="mini-kpi"><div class="mini-kpi-lbl">Trend</div><div class="mini-kpi-val" id="dk4">—</div></div>
    </div>
    <div class="insight" id="deep-insight">Select a product above.</div>
    <div class="g2">
      <div><div style="font-size:10px;color:var(--muted);margin-bottom:8px;" id="deep-lbl">Monthly quantity</div><div class="chart-box" style="height:200px;"><canvas id="c-deep-bar"></canvas></div></div>
      <div><div style="font-size:10px;color:var(--muted);margin-bottom:8px;">Qty vs Spend</div><div class="chart-box" style="height:200px;"><canvas id="c-deep-dual"></canvas></div></div>
    </div>
  </div>
</div>

<!-- USAGE -->
<div id="page-usage" class="page">
  <div class="section-hdr"><h2>Usage — Monthly Quantity by Product Group</h2></div>
  <div class="ctrl-row">
    <select class="ctrl-select" id="usage-category" onchange="renderUsage()">
      <option value="all">All categories</option>
      {{CATEGORY_OPTIONS}}
    </select>
    <select class="ctrl-select" id="usage-top" onchange="renderUsage()">
      <option value="10">Top 10</option>
      <option value="20">Top 20</option>
      <option value="30">Top 30</option>
      <option value="999">All</option>
    </select>
    <div class="tog-grp">
      <button class="tog on" onclick="setUsageMode(\'qty\',this)">Quantity</button>
      <button class="tog" onclick="setUsageMode(\'spend\',this)">Spend</button>
    </div>
  </div>
  <div class="card" style="margin-bottom:12px;"><div class="card-hdr" id="usage-title">Monthly quantity by product</div><div class="card-sub">Each product shows months side by side</div><div class="chart-box" style="height:420px;"><canvas id="c-usage-grouped"></canvas></div></div>
  <div class="card"><div class="card-hdr">Monthly stacked view</div><div class="card-sub">Total per month across selected products</div><div class="chart-box" style="height:280px;"><canvas id="c-usage-stacked"></canvas></div></div>
</div>

<!-- PLANNING -->
<div id="page-planning" class="page">
  <div class="section-hdr"><h2>Order Planning — Next Month</h2></div>
  <div class="insight">Planning based on <strong>{{N_MONTHS}}-month average</strong> with <strong>10% safety buffer</strong>. All quantities and prices are editable. Totals update live.</div>
  <div class="ctrl-row">
    <input type="text" class="ctrl-input" id="plan-search" placeholder="Filter products..." oninput="filterPlan()" style="flex:1;min-width:200px;max-width:300px;">
    <select class="ctrl-select" id="plan-filter" onchange="filterPlan()">
      <option value="all">All products</option>
      <option value="up">Trending up ↑</option>
      <option value="down">Trending down ↓</option>
      <option value="high_spend">High spend (>R50K)</option>
    </select>
    <button class="btn btn-blue" onclick="resetPlan()">↺ Reset</button>
    <button class="btn btn-ghost" onclick="exportPlan()">⬇ Export CSV</button>
  </div>
  <div class="mini-kpis">
    <div class="mini-kpi"><div class="mini-kpi-lbl">Products</div><div class="mini-kpi-val" id="plan-count" style="color:var(--blue)">—</div></div>
    <div class="mini-kpi"><div class="mini-kpi-lbl">Total qty</div><div class="mini-kpi-val" id="plan-total-qty" style="color:var(--teal)">—</div></div>
    <div class="mini-kpi"><div class="mini-kpi-lbl">Est. spend</div><div class="mini-kpi-val" id="plan-total-spend" style="color:var(--orange)">—</div></div>
    <div class="mini-kpi"><div class="mini-kpi-lbl">vs avg month</div><div class="mini-kpi-val" id="plan-vs-avg" style="color:var(--green)">—</div></div>
  </div>
  <div class="plan-table-wrap">
    <table>
      <thead><tr>
        <th>Product</th><th style="text-align:center">Cat</th>
        {{MONTH_HEADERS_PLAN}}
        <th style="text-align:center">Avg/Mo</th>
        <th style="text-align:center">Trend</th>
        <th style="text-align:center">Order Qty ✏️</th>
        <th style="text-align:right">Unit Price ✏️</th>
        <th style="text-align:right">Est. Spend</th>
      </tr></thead>
      <tbody id="plan-tbody"></tbody>
      <tfoot><tr class="total-row">
        <td colspan="{{PLAN_COLSPAN}}">TOTAL (visible)</td>
        <td style="text-align:center" id="foot-qty">—</td>
        <td></td>
        <td style="text-align:right" id="foot-spend">—</td>
      </tr></tfoot>
    </table>
  </div>
</div>

</div><!-- main -->
<footer>Purchase Order Dashboard · {{DATE_RANGE}} · {{TOTAL_ORDERS}} orders · {{UNIQUE_PRODUCTS}} products · Generated by Claude AI</footer>

<script>
const RAW = {{JSON_DATA}};
const MONTHS = RAW.months;
const PRODS = RAW.products;
const C = {blue:"#3d7eff",teal:"#00d4aa",orange:"#ff7c3d",purple:"#a855f7",red:"#ff4d6d",green:"#22c55e",muted:"#5a6a8a",grid:"rgba(255,255,255,0.04)"};
Chart.defaults.color = C.muted;
Chart.defaults.font.family = "'Space Grotesk', sans-serif";
Chart.defaults.font.size = 10;

const fR = v => v>=1e6?"R"+(v/1e6).toFixed(2)+"M":v>=1e3?"R"+(v/1e3).toFixed(1)+"K":"R"+Math.round(v);
const fQ = v => v>=1e3?(v/1e3).toFixed(1)+"K":String(Math.round(v));
const axX = () => ({ticks:{color:C.muted,font:{size:10}},grid:{display:false},border:{color:"transparent"}});
const axY = cb => ({ticks:{color:C.muted,font:{size:10},callback:cb},grid:{color:C.grid},border:{color:"transparent"}});
const noLeg = {plugins:{legend:{display:false}}};

let chartsInited={overview:false,products:false,usage:false,planning:false};
function showPage(id,btn){
  document.querySelectorAll(".page").forEach(p=>p.classList.remove("active"));
  document.querySelectorAll(".nav-tab").forEach(b=>b.classList.remove("active"));
  document.getElementById("page-"+id).classList.add("active");
  btn.classList.add("active");
  if(!chartsInited[id]){chartsInited[id]=true;initPage(id);}
}
function initPage(id){
  if(id==="overview")buildOverview();
  if(id==="products")buildProducts();
  if(id==="usage"){buildUsage();renderUsage();}
  if(id==="planning")buildPlanning();
}

const COLORS=["#3d7eff","#6fa3ff","#00d4aa","#ff7c3d","#a855f7","#ff4d6d","#22c55e","#f59e0b","#e879f9","#38bdf8","#f97316","#14b8a6","#8b5cf6","#ec4899"];
const mColor = i => COLORS[i%COLORS.length];

function buildOverview(){
  new Chart(document.getElementById("c-daily"),{type:"line",data:{labels:RAW.daily.labels,datasets:[{data:RAW.daily.vals,borderColor:C.blue,backgroundColor:"rgba(61,126,255,0.06)",borderWidth:1.5,pointRadius:2,fill:true,tension:0.35}]},options:{responsive:true,maintainAspectRatio:false,...noLeg,plugins:{...noLeg.plugins,tooltip:{callbacks:{label:c=>" R"+c.parsed.y.toLocaleString()}}},scales:{x:{ticks:{color:C.muted,font:{size:9},maxTicksLimit:12,maxRotation:35},grid:{display:false},border:{color:"transparent"}},y:axY(v=>v>=1e6?"R"+(v/1e6).toFixed(1)+"M":"R"+(v/1e3).toFixed(0)+"K")}}});
  new Chart(document.getElementById("c-monthly"),{type:"bar",data:{labels:RAW.monthly.labels,datasets:[{data:RAW.monthly.vals,backgroundColor:COLORS.slice(0,RAW.monthly.vals.length),borderRadius:6}]},options:{responsive:true,maintainAspectRatio:false,...noLeg,plugins:{...noLeg.plugins,tooltip:{callbacks:{label:c=>" R"+c.parsed.y.toLocaleString()}}},scales:{x:axX(),y:axY(v=>"R"+(v/1e6).toFixed(1)+"M")}}});
  const t20=[...PRODS].sort((a,b)=>b.total_spend-a.total_spend).slice(0,20);
  new Chart(document.getElementById("c-topspend"),{type:"bar",data:{labels:t20.map(p=>p.product.length>30?p.product.slice(0,28)+"…":p.product),datasets:[{data:t20.map(p=>p.total_spend),backgroundColor:C.blue,borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:"y",...noLeg,plugins:{...noLeg.plugins,tooltip:{callbacks:{label:c=>" R"+c.parsed.x.toLocaleString()}}},scales:{x:axY(v=>fR(v)),y:{ticks:{color:C.muted,font:{size:10}},grid:{display:false},border:{color:"transparent"}}}}});
  const tq=[...PRODS].sort((a,b)=>b.total_qty-a.total_qty).slice(0,15);
  new Chart(document.getElementById("c-topqty"),{type:"bar",data:{labels:tq.map(p=>p.product.length>28?p.product.slice(0,26)+"…":p.product),datasets:[{data:tq.map(p=>p.total_qty),backgroundColor:C.teal,borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:"y",...noLeg,plugins:{...noLeg.plugins,tooltip:{callbacks:{label:c=>" "+c.parsed.x.toLocaleString()+" units"}}},scales:{x:axY(v=>fQ(v)),y:{ticks:{color:C.muted,font:{size:10}},grid:{display:false},border:{color:"transparent"}}}}});
  const fr=[...PRODS].filter(p=>p.order_count>0).sort((a,b)=>b.order_count-a.order_count).slice(0,15);
  new Chart(document.getElementById("c-freq"),{type:"bar",data:{labels:fr.map(p=>p.product.length>28?p.product.slice(0,26)+"…":p.product),datasets:[{data:fr.map(p=>p.order_count),backgroundColor:C.orange,borderRadius:4}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:"y",...noLeg,plugins:{...noLeg.plugins,tooltip:{callbacks:{label:c=>" "+c.parsed.x+" orders"}}},scales:{x:axY(v=>v),y:{ticks:{color:C.muted,font:{size:10}},grid:{display:false},border:{color:"transparent"}}}}});
  new Chart(document.getElementById("c-cat"),{type:"bar",data:{labels:RAW.category.labels,datasets:[{data:RAW.category.vals,backgroundColor:COLORS,borderRadius:5}]},options:{responsive:true,maintainAspectRatio:false,...noLeg,plugins:{...noLeg.plugins,tooltip:{callbacks:{label:c=>" R"+c.parsed.y.toLocaleString()}}},scales:{x:axX(),y:axY(v=>fR(v))}}});
  const lg=document.getElementById("sup-legend");
  RAW.supplier.labels.forEach((l,i)=>{lg.innerHTML+=`<span class="leg"><span class="leg-dot" style="background:${mColor(i)}"></span>${l}</span>`;});
  new Chart(document.getElementById("c-sup"),{type:"doughnut",data:{labels:RAW.supplier.labels,datasets:[{data:RAW.supplier.vals,backgroundColor:COLORS,borderWidth:0,hoverOffset:6}]},options:{responsive:true,maintainAspectRatio:false,cutout:"65%",plugins:{legend:{display:false},tooltip:{callbacks:{label:c=>" R"+c.parsed.toLocaleString()}}}}});
}

let prodMode="qty",deepMode="qty",deepC1,deepC2;
function setProdMode(m,btn){prodMode=m;document.querySelectorAll("#page-products .tog").forEach(b=>b.classList.remove("on"));btn.classList.add("on");renderProdTable();}
function setDeepMode(m,btn){deepMode=m;document.querySelectorAll(".tog[onclick*=setDeepMode]").forEach(b=>b.classList.remove("on"));btn.classList.add("on");updateDeep();}

function buildProducts(){
  renderProdTable();
  const sel=document.getElementById("deep-select");
  PRODS.forEach(p=>{const o=document.createElement("option");o.value=p.product;o.textContent=p.product;sel.appendChild(o);});
  updateDeep();
}

function filterProducts(){
  const q=document.getElementById("prod-search").value.toLowerCase();
  const s=document.getElementById("prod-sort").value;
  let arr=PRODS.filter(p=>p.product.toLowerCase().includes(q));
  if(s==="qty")arr.sort((a,b)=>b.total_qty-a.total_qty);
  else if(s==="spend")arr.sort((a,b)=>b.total_spend-a.total_spend);
  else if(s==="name")arr.sort((a,b)=>a.product.localeCompare(b.product));
  else if(s==="trend_up")arr=[...arr.filter(p=>p.trend==="up"),...arr.filter(p=>p.trend!=="up")];
  else if(s==="trend_down")arr=[...arr.filter(p=>p.trend==="down"),...arr.filter(p=>p.trend!=="down")];
  document.getElementById("prod-count").textContent=arr.length+" products";
  renderProdTable(arr);
}

function renderProdTable(arr){
  arr=arr||PRODS;
  const isQ=prodMode==="qty";
  const tbody=document.getElementById("products-tbody");
  tbody.innerHTML="";
  arr.forEach(p=>{
    const mVals=MONTHS.map(m=>{const k=m.toLowerCase();return isQ?(p[k+"_qty"]||0):(p[k+"_spend"]||0);});
    const tot=isQ?p.total_qty:p.total_spend;
    const avg=isQ?p.avg_monthly_qty:(p.total_spend/MONTHS.length);
    const fmt=v=>isQ?(v?v.toLocaleString():"—"):(v?fR(v):"—");
    const badge=p.trend==="up"?"badge-up":p.trend==="down"?"badge-down":"badge-stable";
    const arrow=p.trend==="up"?"↑":p.trend==="down"?"↓":"→";
    const mCells=mVals.map(v=>`<td style="text-align:center;color:${v?"var(--text)":"var(--muted)"};">${fmt(v)}</td>`).join("");
    tbody.innerHTML+=`<tr><td style="max-width:220px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${p.product}</td><td style="text-align:center;color:var(--muted2);font-family:\'JetBrains Mono\',monospace;">${p.code}</td>${mCells}<td style="text-align:center;font-weight:600;color:var(--blue);">${fmt(tot)}</td><td style="text-align:center;color:var(--muted2);">${isQ?fQ(avg):fR(avg)}</td><td style="text-align:center;"><span class="badge ${badge}">${arrow} ${p.trend}</span></td></tr>`;
  });
}

function updateDeep(){
  const prod=document.getElementById("deep-select").value;
  const p=PRODS.find(x=>x.product===prod);
  if(!p)return;
  document.getElementById("dk1").textContent=p.total_qty.toLocaleString();
  document.getElementById("dk2").textContent=fR(p.total_spend);
  document.getElementById("dk3").textContent="R"+p.avg_price.toFixed(2);
  const tc=p.trend==="up"?C.green:p.trend==="down"?C.red:C.muted;
  document.getElementById("dk4").innerHTML=`<span style="color:${tc}">${p.trend==="up"?"↑ Growing":p.trend==="down"?"↓ Declining":"→ Stable"}</span>`;
  document.getElementById("deep-lbl").textContent=deepMode==="qty"?"Monthly quantity":"Monthly spend (R)";
  const isQ=deepMode==="qty";
  const vals=MONTHS.map(m=>{const k=m.toLowerCase();return isQ?(p[k+"_qty"]||0):(p[k+"_spend"]||0);});
  const peak=Math.max(...vals);
  const bgs=vals.map(v=>v===peak?C.blue:"rgba(61,126,255,0.25)");
  const missing=MONTHS.filter((_,i)=>vals[i]===0);
  let ins=`<strong>${p.product}</strong>`;
  if(p.supplier)ins+=` · Supplier: <strong>${p.supplier}</strong>`;
  ins+=` · Peak: <strong>${MONTHS[vals.indexOf(peak)]}</strong>`;
  if(missing.length)ins+=` · Not ordered in: <strong>${missing.join(", ")}</strong>`;
  ins+=` · Suggested next order: <strong>${p.suggested_next.toLocaleString()}</strong> units`;
  document.getElementById("deep-insight").innerHTML=ins;
  if(deepC1)deepC1.destroy();
  deepC1=new Chart(document.getElementById("c-deep-bar"),{type:"bar",data:{labels:MONTHS,datasets:[{data:vals,backgroundColor:bgs,borderRadius:6}]},options:{responsive:true,maintainAspectRatio:false,...noLeg,plugins:{...noLeg.plugins,tooltip:{callbacks:{label:c=>isQ?" "+c.parsed.y.toLocaleString()+" units":" R"+c.parsed.y.toLocaleString()}}},scales:{x:axX(),y:axY(v=>isQ?fQ(v):fR(v))}}});
  if(deepC2)deepC2.destroy();
  const qVals=MONTHS.map(m=>p[m.toLowerCase()+"_qty"]||0);
  const sVals=MONTHS.map(m=>p[m.toLowerCase()+"_spend"]||0);
  deepC2=new Chart(document.getElementById("c-deep-dual"),{type:"bar",data:{labels:MONTHS,datasets:[{label:"Qty",data:qVals,backgroundColor:"rgba(61,126,255,0.6)",borderRadius:5,yAxisID:"y",order:2},{label:"Spend",data:sVals,type:"line",borderColor:C.orange,backgroundColor:"rgba(255,124,61,0.07)",borderWidth:2,pointBackgroundColor:C.orange,pointRadius:4,fill:true,yAxisID:"y2",tension:0.35,order:1}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:true,position:"top",align:"end",labels:{boxWidth:8,color:C.muted,font:{size:10},padding:10}},tooltip:{mode:"index",intersect:false}},scales:{x:axX(),y:{ticks:{color:C.blue,font:{size:10},callback:v=>fQ(v)},grid:{color:C.grid},border:{color:"transparent"},position:"left",beginAtZero:true},y2:{ticks:{color:C.orange,font:{size:10},callback:v=>fR(v)},grid:{display:false},border:{color:"transparent"},position:"right",beginAtZero:true}}}});
}

let usageMode="qty",usageG,usageS;
function setUsageMode(m,btn){usageMode=m;document.querySelectorAll(".tog[onclick*=setUsageMode]").forEach(b=>b.classList.remove("on"));btn.classList.add("on");renderUsage();}
function buildUsage(){}

function renderUsage(){
  const cat=document.getElementById("usage-category").value;
  const topN=parseInt(document.getElementById("usage-top").value);
  const isQ=usageMode==="qty";
  let arr=cat==="all"?PRODS:PRODS.filter(p=>p.code.startsWith(cat));
  arr=[...arr].sort((a,b)=>isQ?b.total_qty-a.total_qty:b.total_spend-a.total_spend).slice(0,topN);
  document.getElementById("usage-title").textContent=`Monthly ${isQ?"quantity":"spend"} by product${cat!=="all"?" — "+cat:""} (top ${Math.min(topN,arr.length)})`;
  const labels=arr.map(p=>p.product.length>22?p.product.slice(0,20)+"…":p.product);
  const datasets=MONTHS.map((m,i)=>({label:m,data:arr.map(p=>isQ?(p[m.toLowerCase()+"_qty"]||0):(p[m.toLowerCase()+"_spend"]||0)),backgroundColor:`${[C.blue,C.teal,C.orange,C.purple,C.red][i%5]}BB`,borderRadius:3}));
  if(usageG)usageG.destroy();
  usageG=new Chart(document.getElementById("c-usage-grouped"),{type:"bar",data:{labels,datasets},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:true,position:"top",align:"end",labels:{boxWidth:10,color:C.muted,font:{size:10},padding:12}},tooltip:{mode:"index",intersect:false}},scales:{x:{ticks:{color:C.muted,font:{size:arr.length>20?8:10},maxRotation:45},grid:{display:false},border:{color:"transparent"}},y:axY(v=>isQ?fQ(v):fR(v))}}});
  if(usageS)usageS.destroy();
  usageS=new Chart(document.getElementById("c-usage-stacked"),{type:"bar",data:{labels:MONTHS,datasets:arr.map((p,i)=>({label:p.product.length>18?p.product.slice(0,16)+"…":p.product,data:MONTHS.map(m=>isQ?(p[m.toLowerCase()+"_qty"]||0):(p[m.toLowerCase()+"_spend"]||0)),backgroundColor:`hsl(${(i*137)%360},65%,55%)`,borderRadius:0}))},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:arr.length<=15,position:"right",labels:{boxWidth:8,color:C.muted,font:{size:9},padding:6}},tooltip:{mode:"index",intersect:false}},scales:{x:axX(),y:{stacked:true,...axY(v=>isQ?fQ(v):fR(v))}},datasets:{bar:{categoryPercentage:0.6}}}});
}

let planData=[];
function buildPlanning(){
  planData=PRODS.map(p=>({...p,order_qty:Math.max(p.suggested_next,0),unit_price:p.avg_price}));
  filterPlan();
}

function filterPlan(){
  const q=(document.getElementById("plan-search").value||"").toLowerCase();
  const f=document.getElementById("plan-filter").value;
  let arr=planData.filter(p=>{
    if(!p.product.toLowerCase().includes(q))return false;
    if(f==="up")return p.trend==="up";
    if(f==="down")return p.trend==="down";
    if(f==="high_spend")return p.total_spend>50000;
    return true;
  });
  renderPlanTable(arr);
}

function renderPlanTable(arr){
  const tbody=document.getElementById("plan-tbody");
  tbody.innerHTML="";
  let tq=0,ts=0;
  arr.forEach((p,i)=>{
    const est=Math.round(p.order_qty*p.unit_price);
    tq+=p.order_qty;ts+=est;
    const badge=p.trend==="up"?"badge-up":p.trend==="down"?"badge-down":"badge-stable";
    const mCells=MONTHS.map(m=>{const v=p[m.toLowerCase()+"_qty"]||0;return`<td style="text-align:center;color:${v?"var(--text)":"var(--muted)"};">${v?v.toLocaleString():"—"}</td>`;}).join("");
    tbody.innerHTML+=`<tr><td style="max-width:200px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${p.product}">${p.product}</td><td style="text-align:center;font-size:10px;color:var(--muted2);">${p.code.slice(0,6)}</td>${mCells}<td style="text-align:center;color:var(--muted2);">${fQ(p.avg_monthly_qty)}</td><td style="text-align:center;"><span class="badge ${badge}">${p.trend==="up"?"↑":p.trend==="down"?"↓":"→"}</span></td><td style="text-align:center;"><input class="editable" type="number" min="0" value="${p.order_qty}" oninput="updRow(this,${i},arr)"></td><td style="text-align:right;"><input class="editable" type="number" min="0" step="0.01" value="${p.unit_price.toFixed(2)}" oninput="updRow(this,${i},arr)"></td><td style="text-align:right;font-weight:600;color:var(--teal);" id="est-r${i}">${fR(est)}</td></tr>`;
  });
  document.getElementById("plan-count").textContent=arr.length;
  document.getElementById("plan-total-qty").textContent=tq.toLocaleString();
  document.getElementById("plan-total-spend").textContent=fR(ts);
  document.getElementById("foot-qty").textContent=tq.toLocaleString();
  document.getElementById("foot-spend").textContent=fR(ts);
  const avgMs=planData.reduce((s,p)=>s+p.total_spend/MONTHS.length,0);
  const pct=((ts/avgMs-1)*100).toFixed(0);
  const el=document.getElementById("plan-vs-avg");
  el.textContent=(pct>=0?"+":"")+pct+"%";
  el.style.color=pct>=0?"var(--green)":"var(--red)";
  window._planArr=arr;
}

function updRow(input,i){
  const arr=window._planArr||[];
  const p=arr[i];if(!p)return;
  const row=input.closest("tr");
  const inputs=row.querySelectorAll("input");
  p.order_qty=parseInt(inputs[0].value)||0;
  p.unit_price=parseFloat(inputs[1].value)||0;
  const est=Math.round(p.order_qty*p.unit_price);
  const cell=document.getElementById("est-r"+i);
  if(cell)cell.textContent=fR(est);
  let tq=0,ts=0;
  arr.forEach(r=>{tq+=r.order_qty;ts+=Math.round(r.order_qty*r.unit_price);});
  document.getElementById("plan-total-qty").textContent=tq.toLocaleString();
  document.getElementById("plan-total-spend").textContent=fR(ts);
  document.getElementById("foot-qty").textContent=tq.toLocaleString();
  document.getElementById("foot-spend").textContent=fR(ts);
}

function resetPlan(){planData.forEach(p=>{p.order_qty=p.suggested_next;p.unit_price=p.avg_price;});filterPlan();}
function exportPlan(){
  const arr=window._planArr||planData;
  const mHdrs=MONTHS.map(m=>m+" Qty").join(",");
  const hdr=["Product","Code",mHdrs,"Avg/Month","Trend","Order Qty","Unit Price","Est Spend"].join(",");
  const rows=arr.map(p=>{const mc=MONTHS.map(m=>p[m.toLowerCase()+"_qty"]||0).join(",");return[`"${p.product}"`,p.code,mc,Math.round(p.avg_monthly_qty),p.trend,p.order_qty,p.unit_price.toFixed(2),Math.round(p.order_qty*p.unit_price)].join(",");});
  const csv=[hdr,...rows].join("\\n");
  const a=document.createElement("a");
  a.href="data:text/csv;charset=utf-8,"+encodeURIComponent(csv);
  a.download="OrderPlan_NextMonth.csv";
  a.click();
}

chartsInited.overview=true;
buildOverview();
</script>
</body>
</html>'''


def fmt_spend(v):
    if v >= 1_000_000:
        return f"R{v/1_000_000:.1f}M"
    if v >= 1_000:
        return f"R{v/1_000:.0f}K"
    return f"R{v:.0f}"


def build_dashboard(data: dict, output_path: str):
    months = data['months']
    n = len(months)

    month_headers     = ''.join(f'<th style="text-align:center">{m}</th>' for m in months)
    month_headers_plan = ''.join(f'<th style="text-align:center">{m} Qty</th>' for m in months)
    plan_colspan = 2 + n + 2   # product + cat + months + avg + trend

    cats = set(p['code'][:3] for p in data['products'] if p['code'])
    cat_options = '\n'.join(
        f'<option value="{c}">{c}</option>' for c in sorted(cats)
    )

    total_spend = data['kpis']['total_spend']
    avg_order   = data['kpis']['avg_order']

    html = HTML_TEMPLATE \
        .replace('{{JSON_DATA}}',           json.dumps(data)) \
        .replace('{{DATE_RANGE}}',          data['kpis']['date_range']) \
        .replace('{{UNIQUE_PRODUCTS}}',     str(data['kpis']['unique_products'])) \
        .replace('{{UNIQUE_SUPPLIERS}}',    str(data['kpis']['unique_suppliers'])) \
        .replace('{{TOTAL_ORDERS}}',        str(data['kpis']['total_orders'])) \
        .replace('{{TOTAL_SPEND_FMT}}',     fmt_spend(total_spend)) \
        .replace('{{AVG_ORDER_FMT}}',       fmt_spend(avg_order)) \
        .replace('{{N_MONTHS}}',            str(n)) \
        .replace('{{MONTH_HEADERS}}',       month_headers) \
        .replace('{{MONTH_HEADERS_PLAN}}',  month_headers_plan) \
        .replace('{{PLAN_COLSPAN}}',        str(plan_colspan)) \
        .replace('{{CATEGORY_OPTIONS}}',    cat_options)

    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"Dashboard saved: {output_path}")


def main():
    parser = argparse.ArgumentParser(description='Build HTML dashboard from JSON data')
    parser.add_argument('--data',   required=True, help='Path to dashboard_data.json')
    parser.add_argument('--output', default='dashboard.html', help='Output HTML path')
    args = parser.parse_args()

    with open(args.data) as f:
        data = json.load(f)

    build_dashboard(data, args.output)


if __name__ == '__main__':
    main()
