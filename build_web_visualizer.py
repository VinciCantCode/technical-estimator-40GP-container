"""
Generate the standalone, zero-dependency Interactive Web Visualizer:
  File: Precise_Pipe_40GP_Visualizer.html
Embeds the exact ASME B36.10M / B36.19M dataset from pipe_data.py so all calculations
match Precise_Pipe_40GP_Container_Model_V2_Linked.xlsx with 100% precision.
"""

import json
from pipe_data import ASME_PIPE_SPECS, SCHEDULES

# Convert ASME_PIPE_SPECS into a clean JSON-serializable structure
specs_json_list = []
for nps_str, dn_val, od_in, sch_dict in ASME_PIPE_SPECS:
    specs_json_list.append({
        "nps": nps_str,
        "dn": dn_val,
        "od_in": od_in,
        "od_mm": round(od_in * 25.4, 2),
        "schedules": {k: {"wt_in": v, "wt_mm": round(v * 25.4, 2)} for k, v in sch_dict.items()}
    })

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />
  <title>Precise Pipe — 40GP Container Loading & Commercial Estimator</title>
  <style>
    :root {
      --bg-dark: #0b1120;
      --bg-panel: #111827;
      --bg-card: #1e293b;
      --bg-input: #0f172a;
      --border: #334155;
      --text-main: #f8fafc;
      --text-muted: #94a3b8;
      --accent-blue: #38bdf8;
      --accent-navy: #1d4ed8;
      --accent-emerald: #10b981;
      --accent-teal: #14b8a6;
      --accent-amber: #f59e0b;
      --accent-red: #ef4444;
    }
    * {
      box-sizing: border-box;
      margin: 0;
      padding: 0;
      font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Roboto, Helvetica, Arial, sans-serif;
    }
    body {
      background-color: var(--bg-dark);
      color: var(--text-main);
      min-height: 100vh;
      display: flex;
      flex-direction: column;
    }
    /* Top Header */
    header {
      background: linear-gradient(90deg, #0f2c59 0%, #1e3a8a 60%, #0f172a 100%);
      border-bottom: 1px solid var(--border);
      padding: 14px 24px;
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 12px;
    }
    .brand-title {
      display: flex;
      align-items: center;
      gap: 12px;
    }
    .brand-badge {
      background: var(--accent-amber);
      color: #0f172a;
      font-weight: 800;
      font-size: 13px;
      padding: 4px 10px;
      border-radius: 4px;
      letter-spacing: 0.8px;
    }
    header h1 {
      font-size: 18px;
      font-weight: 700;
      letter-spacing: 0.3px;
      color: #ffffff;
    }
    header p {
      font-size: 12px;
      color: #cbd5e1;
    }
    .header-meta {
      display: flex;
      gap: 16px;
      font-size: 12px;
      color: #e2e8f0;
      background: rgba(15, 23, 42, 0.55);
      padding: 6px 14px;
      border-radius: 6px;
      border: 1px solid rgba(255,255,255,0.12);
    }
    .header-meta span strong {
      color: var(--accent-blue);
    }

    /* Main Workspace Grid */
    .workspace {
      display: grid;
      grid-template-columns: 350px 1fr;
      gap: 16px;
      padding: 16px 20px;
      flex: 1;
    }
    @media (max-width: 1100px) {
      .workspace {
        grid-template-columns: 1fr;
      }
    }

    /* Left Sidebar Controls */
    .sidebar {
      background: var(--bg-panel);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 16px;
      display: flex;
      flex-direction: column;
      gap: 16px;
      height: fit-content;
    }
    .sec-title {
      font-size: 12px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.8px;
      color: var(--accent-blue);
      border-bottom: 1px solid var(--border);
      padding-bottom: 6px;
      margin-bottom: 10px;
      display: flex;
      justify-content: space-between;
      align-items: center;
    }
    .control-group {
      display: flex;
      flex-direction: column;
      gap: 6px;
      margin-bottom: 10px;
    }
    .control-label {
      font-size: 12px;
      font-weight: 600;
      color: #cbd5e1;
      display: flex;
      justify-content: space-between;
    }
    .control-label span.val-badge {
      color: var(--accent-amber);
      font-weight: 700;
    }
    select, input[type="number"] {
      width: 100%;
      background: var(--bg-input);
      color: var(--text-main);
      border: 1px solid var(--border);
      border-radius: 6px;
      padding: 8px 10px;
      font-size: 13px;
      font-weight: 600;
      outline: none;
      transition: border-color 0.15s;
    }
    select:focus, input[type="number"]:focus {
      border-color: var(--accent-blue);
    }
    input[type="range"] {
      width: 100%;
      accent-color: var(--accent-amber);
      cursor: pointer;
    }

    /* Material Segmented Buttons */
    .seg-group {
      display: grid;
      grid-template-columns: repeat(3, 1fr);
      gap: 6px;
    }
    .seg-group.two-col {
      grid-template-columns: repeat(2, 1fr);
    }
    .seg-group.four-col {
      grid-template-columns: repeat(4, 1fr);
    }
    .seg-btn {
      background: var(--bg-input);
      border: 1px solid var(--border);
      color: var(--text-muted);
      padding: 7px 6px;
      font-size: 11px;
      font-weight: 600;
      border-radius: 6px;
      cursor: pointer;
      text-align: center;
      transition: all 0.15s;
    }
    .seg-btn:hover:not(:disabled) {
      border-color: #64748b;
      color: var(--text-main);
    }
    .seg-btn.active {
      background: rgba(56, 189, 248, 0.15);
      border-color: var(--accent-blue);
      color: #ffffff;
      box-shadow: 0 0 8px rgba(56, 189, 248, 0.25);
    }
    .seg-btn.active-ss {
      background: rgba(20, 184, 166, 0.18);
      border-color: var(--accent-teal);
      color: #ffffff;
    }

    /* Schedule Pill Grid */
    .sch-grid {
      display: grid;
      grid-template-columns: repeat(6, 1fr);
      gap: 5px;
    }
    .sch-pill {
      background: var(--bg-input);
      border: 1px solid var(--border);
      color: var(--text-main);
      padding: 6px 2px;
      font-size: 11px;
      font-weight: 700;
      border-radius: 5px;
      cursor: pointer;
      text-align: center;
      transition: all 0.12s;
    }
    .sch-pill.ss-type {
      border-bottom: 2px solid var(--accent-teal);
    }
    .sch-pill:hover:not(:disabled) {
      border-color: var(--accent-amber);
    }
    .sch-pill.active {
      background: var(--accent-amber);
      border-color: var(--accent-amber);
      color: #0f172a;
    }
    .sch-pill:disabled {
      opacity: 0.25;
      cursor: not-allowed;
      text-decoration: line-through;
    }

    /* Nesting Toggle Box */
    .nesting-box {
      background: rgba(20, 184, 166, 0.08);
      border: 1px solid rgba(20, 184, 166, 0.35);
      border-radius: 8px;
      padding: 10px 12px;
      display: flex;
      flex-direction: column;
      gap: 6px;
    }
    .nesting-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      font-size: 12px;
      font-weight: 700;
      color: #2dd4bf;
    }
    .toggle-switch {
      display: flex;
      align-items: center;
      gap: 6px;
      cursor: pointer;
      font-size: 11px;
      color: #e2e8f0;
    }

    /* Right Content Area */
    .main-area {
      display: flex;
      flex-direction: column;
      gap: 16px;
    }

    /* Top KPI Cards Row */
    .kpi-row {
      display: grid;
      grid-template-columns: repeat(4, 1fr);
      gap: 12px;
    }
    @media (max-width: 900px) {
      .kpi-row {
        grid-template-columns: repeat(2, 1fr);
      }
    }
    .kpi-card {
      background: var(--bg-panel);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 14px 16px;
      position: relative;
      overflow: hidden;
    }
    .kpi-card.emerald {
      border-top: 3px solid var(--accent-emerald);
      background: linear-gradient(180deg, rgba(16,185,129,0.1) 0%, var(--bg-panel) 100%);
    }
    .kpi-card.blue {
      border-top: 3px solid var(--accent-blue);
      background: linear-gradient(180deg, rgba(56,189,248,0.1) 0%, var(--bg-panel) 100%);
    }
    .kpi-card.amber {
      border-top: 3px solid var(--accent-amber);
    }
    .kpi-card.purple {
      border-top: 3px solid #818cf8;
    }
    .kpi-label {
      font-size: 11px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.6px;
      color: var(--text-muted);
      margin-bottom: 4px;
    }
    .kpi-value {
      font-size: 28px;
      font-weight: 800;
      line-height: 1.15;
      margin-bottom: 4px;
    }
    .kpi-card.emerald .kpi-value { color: #34d399; }
    .kpi-card.blue .kpi-value { color: #38bdf8; }
    .kpi-card.amber .kpi-value { color: #fbbf24; font-size: 20px; }
    .kpi-card.purple .kpi-value { color: #a5b4fc; font-size: 22px; }
    .kpi-sub {
      font-size: 11px;
      color: #cbd5e1;
    }

    /* Visualizer Split Grid */
    .vis-grid {
      display: grid;
      grid-template-columns: 1.35fr 1fr;
      gap: 16px;
    }
    @media (max-width: 1200px) {
      .vis-grid {
        grid-template-columns: 1fr;
      }
    }
    .canvas-card {
      background: var(--bg-panel);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 14px;
      display: flex;
      flex-direction: column;
      gap: 10px;
    }
    .canvas-header {
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 8px;
    }
    .canvas-header h2 {
      font-size: 13px;
      font-weight: 700;
      text-transform: uppercase;
      letter-spacing: 0.6px;
      color: #e2e8f0;
    }
    .legend-bar {
      display: flex;
      gap: 12px;
      font-size: 11px;
      color: #cbd5e1;
      align-items: center;
    }
    .legend-dot {
      width: 10px;
      height: 10px;
      border-radius: 50%;
      display: inline-block;
      margin-right: 4px;
      vertical-align: middle;
    }
    .canvas-wrap {
      background: #070b14;
      border: 1px solid #1e293b;
      border-radius: 8px;
      display: flex;
      justify-content: center;
      align-items: center;
      padding: 8px;
      position: relative;
    }
    canvas {
      max-width: 100%;
      height: auto;
      display: block;
    }

    /* Right Side Specs & Elevation Card */
    .right-vis-col {
      display: flex;
      flex-direction: column;
      gap: 14px;
    }
    .spec-grid {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
    }
    .spec-item {
      background: var(--bg-card);
      border: 1px solid var(--border);
      border-radius: 6px;
      padding: 8px 10px;
      display: flex;
      flex-direction: column;
    }
    .spec-item .s-lbl {
      font-size: 10px;
      color: var(--text-muted);
      text-transform: uppercase;
    }
    .spec-item .s-val {
      font-size: 14px;
      font-weight: 700;
      color: #f8fafc;
      margin-top: 2px;
    }

    /* Utilization Progress Bars */
    .util-bar-wrap {
      display: flex;
      flex-direction: column;
      gap: 4px;
      margin-top: 4px;
    }
    .util-bar-top {
      display: flex;
      justify-content: space-between;
      font-size: 11px;
      font-weight: 600;
    }
    .progress-track {
      width: 100%;
      height: 8px;
      background: #0f172a;
      border-radius: 4px;
      overflow: hidden;
      border: 1px solid #334155;
    }
    .progress-fill {
      height: 100%;
      transition: width 0.25s ease;
    }

    /* Advisor Banner */
    .advisor-banner {
      background: rgba(20, 184, 166, 0.1);
      border-left: 4px solid var(--accent-teal);
      border-top: 1px solid rgba(20, 184, 166, 0.3);
      border-right: 1px solid rgba(20, 184, 166, 0.3);
      border-bottom: 1px solid rgba(20, 184, 166, 0.3);
      border-radius: 6px;
      padding: 10px 14px;
      font-size: 12px;
      line-height: 1.45;
      color: #e2e8f0;
    }
    .advisor-banner strong {
      color: #2dd4bf;
    }

    /* Bottom Tables Section */
    .table-card {
      background: var(--bg-panel);
      border: 1px solid var(--border);
      border-radius: 10px;
      padding: 14px;
      display: flex;
      flex-direction: column;
      gap: 12px;
    }
    .tab-bar {
      display: flex;
      justify-content: space-between;
      align-items: center;
      flex-wrap: wrap;
      gap: 10px;
      border-bottom: 1px solid var(--border);
      padding-bottom: 10px;
    }
    .tab-btns {
      display: flex;
      gap: 8px;
    }
    .tab-btn {
      background: var(--bg-card);
      border: 1px solid var(--border);
      color: var(--text-muted);
      padding: 7px 14px;
      border-radius: 6px;
      font-size: 12px;
      font-weight: 700;
      cursor: pointer;
    }
    .tab-btn.active {
      background: var(--accent-navy);
      border-color: var(--accent-blue);
      color: #ffffff;
    }
    .filter-bar {
      display: flex;
      gap: 8px;
      align-items: center;
      flex-wrap: wrap;
    }
    .filter-bar select, .filter-bar input {
      width: auto;
      min-width: 130px;
      padding: 6px 10px;
      font-size: 12px;
    }
    .table-scroll {
      max-height: 360px;
      overflow-y: auto;
      border: 1px solid var(--border);
      border-radius: 6px;
    }
    table {
      width: 100%;
      border-collapse: collapse;
      font-size: 12px;
    }
    thead th {
      position: sticky;
      top: 0;
      background: #0f2c59;
      color: #ffffff;
      font-weight: 700;
      padding: 9px 10px;
      text-align: right;
      border-bottom: 2px solid #38bdf8;
      z-index: 2;
      white-space: nowrap;
    }
    thead th:nth-child(1),
    thead th:nth-child(2),
    thead th:nth-child(3),
    thead th:nth-child(4) {
      text-align: center;
    }
    tbody tr {
      border-bottom: 1px solid #1e293b;
      cursor: pointer;
      transition: background 0.1s;
    }
    tbody tr:nth-child(even) {
      background: rgba(30, 41, 59, 0.4);
    }
    tbody tr:hover {
      background: rgba(56, 189, 248, 0.15);
    }
    tbody tr.selected-row {
      background: rgba(16, 185, 129, 0.22) !important;
      outline: 1px solid var(--accent-emerald);
    }
    tbody td {
      padding: 7px 10px;
      text-align: right;
      color: #e2e8f0;
      white-space: nowrap;
    }
    tbody td:nth-child(1),
    tbody td:nth-child(2),
    tbody td:nth-child(3),
    tbody td:nth-child(4) {
      text-align: center;
    }
    .badge-wt {
      background: rgba(245, 158, 11, 0.18);
      color: #fbbf24;
      padding: 2px 7px;
      border-radius: 4px;
      font-weight: 700;
      font-size: 11px;
    }
    .badge-vol {
      background: rgba(56, 189, 248, 0.18);
      color: #38bdf8;
      padding: 2px 7px;
      border-radius: 4px;
      font-weight: 700;
      font-size: 11px;
    }
    .col-out {
      color: #34d399;
      font-weight: 700;
      background: rgba(16, 185, 129, 0.06);
    }
  </style>
</head>
<body>

  <!-- Top Header -->
  <header>
    <div class="brand-title">
      <span class="brand-badge">PRECISE PIPE</span>
      <div>
        <h1>40GP Container Pipe Loading Visualizer &amp; Commercial Estimator</h1>
        <p>Interactive 2D Cross-Section Simulator | ASME B36.10M (Carbon Steel) &amp; ASME B36.19M (Stainless Steel)</p>
      </div>
    </div>
    <div class="header-meta">
      <span>Container: <strong>40' GP (12.032m &times; 2.352m &times; 2.393m)</strong></span>
      <span>Synced Dataset: <strong>36 NPS &times; 18 Schedules (1,176 Specs)</strong></span>
    </div>
  </header>

  <!-- Main Workspace -->
  <div class="workspace">

    <!-- Left Sidebar Controls -->
    <aside class="sidebar">
      <div>
        <div class="sec-title">
          <span>1. Pipe Specification</span>
          <span id="matchedRowBadge" style="font-size:11px; color:#34d399;">Tab 2 Row #139</span>
        </div>

        <!-- Material Grade -->
        <div class="control-group">
          <div class="control-label">
            <span>Material Grade</span>
            <span class="val-badge" id="densityBadge">7.85 g/cm&sup3;</span>
          </div>
          <div class="seg-group" id="matGroup">
            <button class="seg-btn active" data-mat="Carbon Steel">Carbon Steel</button>
            <button class="seg-btn" data-mat="Stainless Steel 304">SS 304</button>
            <button class="seg-btn" data-mat="Stainless Steel 316">SS 316</button>
          </div>
        </div>

        <!-- Nominal Pipe Size (NPS) -->
        <div class="control-group">
          <div class="control-label">
            <span>Nominal Pipe Size (NPS)</span>
            <span class="val-badge" id="odBadge">OD: 168.28 mm</span>
          </div>
          <select id="npsSelect"></select>
        </div>

        <!-- Pipe Schedule -->
        <div class="control-group">
          <div class="control-label">
            <span>Pipe Schedule (Wall Thickness)</span>
            <span class="val-badge" id="wtBadge">WT: 7.11 mm</span>
          </div>
          <div class="sch-grid" id="schGrid"></div>
        </div>

        <!-- Single Pipe Length -->
        <div class="control-group" style="margin-top:12px;">
          <div class="control-label">
            <span>Single Pipe Length (m)</span>
            <span class="val-badge" id="tierBadge">1 Lengthwise Tier</span>
          </div>
          <div class="seg-group four-col" id="lenGroup">
            <button class="seg-btn" data-len="5.80">5.80 m</button>
            <button class="seg-btn" data-len="6.00">6.00 m</button>
            <button class="seg-btn active" data-len="11.80">11.80 m</button>
            <button class="seg-btn" data-len="12.00">12.00 m</button>
          </div>
        </div>
      </div>

      <!-- Container & Packing Controls -->
      <div>
        <div class="sec-title">
          <span>2. 40GP Container &amp; Packing</span>
          <button id="resetBtn" style="background:none;border:none;color:#38bdf8;font-size:11px;cursor:pointer;">Reset Defaults</button>
        </div>

        <!-- Stacking Pattern -->
        <div class="control-group">
          <div class="control-label">
            <span>Cross-Section Stacking Mode</span>
          </div>
          <div class="seg-group two-col" id="stackGroup">
            <button class="seg-btn active" data-stack="Staggered (Hexagonal)">Staggered (Hexagonal)</button>
            <button class="seg-btn" data-stack="Square Grid">Square Grid</button>
          </div>
        </div>

        <!-- Max Payload Slider -->
        <div class="control-group">
          <div class="control-label">
            <span>Max Container Payload (MT)</span>
            <span class="val-badge" id="payloadValLabel">26.00 MT (26,000 kg)</span>
          </div>
          <input type="range" id="payloadSlider" min="18.0" max="28.0" step="0.2" value="26.0" />
        </div>

        <!-- Effective Width & Height -->
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:8px;">
          <div class="control-group">
            <div class="control-label"><span>Eff. Width (mm)</span></div>
            <input type="number" id="effWidthInput" value="2300" min="10" max="2352" step="10" />
          </div>
          <div class="control-group">
            <div class="control-label"><span>Eff. Height (mm)</span></div>
            <input type="number" id="effHeightInput" value="2280" min="10" max="2393" step="10" />
          </div>
        </div>

        <!-- Ocean Freight Input -->
        <div class="control-group">
          <div class="control-label">
            <span>Est. Ocean Freight / 40GP (USD)</span>
            <span class="val-badge" id="freightPerMeterBadge">$0.00 / m</span>
          </div>
          <input type="number" id="freightInput" value="3500" min="500" max="20000" step="100" />
        </div>

        <!-- Pipe Nesting (Telescoping) Simulator -->
        <div class="nesting-box">
          <div class="nesting-header">
            <span>Pipe Nesting (Telescoping)</span>
            <label class="toggle-switch">
              <input type="checkbox" id="nestingToggle" />
              <span>Visualize Inner Pipe</span>
            </label>
          </div>
          <p id="nestingSubtext" style="font-size:11px; color:#94a3b8;">
            Enable to preview telescoping smaller pipes inside large volume-limited pipes to capture unused weight capacity.
          </p>
        </div>
      </div>
    </aside>

    <!-- Right Main Content Area -->
    <main class="main-area">

      <!-- Top 4 KPI Output Cards -->
      <section class="kpi-row">
        <div class="kpi-card emerald">
          <div class="kpi-label">Output 1: 40GP Total Pieces</div>
          <div class="kpi-value" id="kpiPieces">78 pcs</div>
          <div class="kpi-sub" id="kpiLengthSub">Total Length: 920.4 m (3,020 ft)</div>
        </div>
        <div class="kpi-card blue">
          <div class="kpi-label">Output 2: 40GP Metric Tons</div>
          <div class="kpi-value" id="kpiTons">26.00 MT</div>
          <div class="kpi-sub" id="kpiWeightSub">Net Weight: 26,000 kg (57,320 lbs)</div>
        </div>
        <div class="kpi-card amber">
          <div class="kpi-label">Loading Bottleneck</div>
          <div class="kpi-value" id="kpiBottleneck">WEIGHT LIMITED</div>
          <div class="kpi-sub" id="kpiLimitSub">Wt Limit: 78 pcs | Vol Limit: 175 pcs</div>
        </div>
        <div class="kpi-card purple">
          <div class="kpi-label">Commercial Freight Cost</div>
          <div class="kpi-value" id="kpiFreightMT">$134.62 / MT</div>
          <div class="kpi-sub" id="kpiFreightPc">$44.87 / piece | $3.80 / m</div>
        </div>
      </section>

      <!-- Center Visualization Row: 2D Cross-Section Canvas + Side View & Engineering Breakdown -->
      <section class="vis-grid">

        <!-- Left: 2D Cross-Section Canvas -->
        <div class="canvas-card">
          <div class="canvas-header">
            <h2 id="canvasTitle">40GP Cross-Section End View (1 Tier)</h2>
            <div class="legend-bar">
              <span><i class="legend-dot" style="background:#38bdf8;"></i>Loaded Pipe</span>
              <span><i class="legend-dot" style="background:rgba(148,163,184,0.35); border:1px dashed #94a3b8;"></i>Unloaded Slot (Overweight)</span>
              <span id="nestedLegend" style="display:none;"><i class="legend-dot" style="background:#f59e0b;"></i>Nested Inner Pipe</span>
            </div>
          </div>
          <div class="canvas-wrap">
            <canvas id="crossCanvas" width="640" height="580"></canvas>
          </div>
        </div>

        <!-- Right: Lengthwise Side Elevation + Engineering Breakdown -->
        <div class="right-vis-col">
          <!-- Side Elevation Mini Canvas -->
          <div class="canvas-card">
            <div class="canvas-header">
              <h2>40GP Lengthwise Side Elevation (12.032 m)</h2>
              <span id="sideTierText" style="font-size:11px; color:#38bdf8; font-weight:700;">1 Tier &times; 11.80 m</span>
            </div>
            <div class="canvas-wrap" style="padding:6px;">
              <canvas id="sideCanvas" width="500" height="145"></canvas>
            </div>
          </div>

          <!-- Detailed Engineering & Geometry Specs Card -->
          <div class="canvas-card" style="flex:1;">
            <div class="canvas-header">
              <h2>Engineering &amp; Packing Breakdown (Synced with Excel Tab 2)</h2>
            </div>
            <div class="spec-grid">
              <div class="spec-item">
                <span class="s-lbl">Outer Diameter (OD)</span>
                <span class="s-val" id="specOD">168.28 mm (6.625")</span>
              </div>
              <div class="spec-item">
                <span class="s-lbl">Wall Thickness (WT)</span>
                <span class="s-val" id="specWT">7.11 mm (0.280")</span>
              </div>
              <div class="spec-item">
                <span class="s-lbl">Inner Diameter (ID)</span>
                <span class="s-val" id="specID">154.06 mm (6.065")</span>
              </div>
              <div class="spec-item">
                <span class="s-lbl">Unit Weight (Linear)</span>
                <span class="s-val" id="specUnitWt">28.262 kg/m (18.99 lb/ft)</span>
              </div>
              <div class="spec-item">
                <span class="s-lbl">Single Piece Weight</span>
                <span class="s-val" id="specPieceWt">333.49 kg / pc</span>
              </div>
              <div class="spec-item">
                <span class="s-lbl">Cross-Section Grid</span>
                <span class="s-val" id="specGrid">13 horiz &times; 14 layers</span>
              </div>
            </div>

            <!-- Utilization Bars -->
            <div class="util-bar-wrap">
              <div class="util-bar-top">
                <span>Container Payload Utilization (Weight)</span>
                <span id="wtUtilText" style="color:#34d399;">98.8% (25.68 / 26.00 MT)</span>
              </div>
              <div class="progress-track">
                <div class="progress-fill" id="wtUtilBar" style="width:98.8%; background:linear-gradient(90deg, #10b981, #34d399);"></div>
              </div>
            </div>

            <div class="util-bar-wrap">
              <div class="util-bar-top">
                <span>Container Cross-Section Slot Utilization (Space)</span>
                <span id="volUtilText" style="color:#38bdf8;">44.0% (77 / 175 slots)</span>
              </div>
              <div class="progress-track">
                <div class="progress-fill" id="volUtilBar" style="width:44.0%; background:linear-gradient(90deg, #0284c7, #38bdf8);"></div>
              </div>
            </div>

            <!-- Commercial Advisor Insight Box -->
            <div class="advisor-banner" id="advisorBanner">
              <strong>COMMERCIAL ADVISOR:</strong> Loading...
            </div>
          </div>
        </div>
      </section>

      <!-- Bottom Interactive Master Table Section -->
      <section class="table-card">
        <div class="tab-bar">
          <div class="tab-btns">
            <button class="tab-btn active" id="tabBtnCurrentNPS">All Schedules for Selected NPS (Tab 1 Matrix)</button>
            <button class="tab-btn" id="tabBtnMaster">Full Selectable Master Table — All 1,176 Specs (Tab 2 Engine)</button>
          </div>
          <div class="filter-bar" id="masterFilterBar" style="display:none;">
            <select id="filterMat">
              <option value="ALL">All Materials (3)</option>
              <option value="Carbon Steel">Carbon Steel</option>
              <option value="Stainless Steel 304">Stainless Steel 304</option>
              <option value="Stainless Steel 316">Stainless Steel 316</option>
            </select>
            <select id="filterNPS">
              <option value="ALL">All NPS Sizes (36)</option>
            </select>
            <select id="filterSch">
              <option value="ALL">All Schedules (18)</option>
            </select>
            <select id="filterLimit">
              <option value="ALL">All Bottlenecks</option>
              <option value="Weight Limited">Weight Limited</option>
              <option value="Volume Limited">Volume Limited</option>
            </select>
          </div>
        </div>

        <div class="table-scroll">
          <table>
            <thead>
              <tr>
                <th>Row ID</th>
                <th>Material</th>
                <th>NPS (in)</th>
                <th>Schedule</th>
                <th>OD (mm)</th>
                <th>WT (in)</th>
                <th>WT (mm)</th>
                <th>ID (mm)</th>
                <th>Unit Wt (kg/m)</th>
                <th>Piece Wt (kg)</th>
                <th>Horiz. Fit</th>
                <th>Vert. Layers</th>
                <th>Max Vol (pcs)</th>
                <th>Max Wt (pcs)</th>
                <th>40GP Pieces</th>
                <th>40GP Tons (MT)</th>
                <th>Payload Util</th>
                <th>Bottleneck</th>
              </tr>
            </thead>
            <tbody id="tableBody"></tbody>
          </table>
        </div>
      </section>

    </main>
  </div>

  <script>
    // Embedded ASME B36.10M & B36.19M Master Data (100% synced with Excel Model V2)
    const PIPE_SPECS = __PIPE_SPECS_JSON__;
    const SCHEDULES = __SCHEDULES_JSON__;
    const STAINLESS_SCH_SET = new Set(["5S", "10S", "40S", "80S"]);

    // Material Constants (Exact Excel PI()*10^-6*Density*1000 rounded to 7 decimals)
    const MATERIALS = {
      "Carbon Steel":        { density: 7.85, factorC: 0.0246615, tag: "CS",    colorOuter: "#3b82f6", colorInner: "#1d4ed8", stroke: "#93c5fd" },
      "Stainless Steel 304": { density: 7.93, factorC: 0.0249128, tag: "SS304", colorOuter: "#14b8a6", colorInner: "#0f766e", stroke: "#99f6e4" },
      "Stainless Steel 316": { density: 7.98, factorC: 0.0250699, tag: "SS316", colorOuter: "#06b6d4", colorInner: "#0e7490", stroke: "#a5f3fc" }
    };

    const CONTAINER_LEN_M = 12.032;
    const CONTAINER_INT_W = 2352;
    const CONTAINER_INT_H = 2393;

    // Application State
    const state = {
      material: "Carbon Steel",
      nps: "6''",
      schedule: "40",
      pipeLen: 11.80,
      stacking: "Staggered (Hexagonal)",
      maxPayloadMT: 26.00,
      effWidth: 2300,
      effHeight: 2280,
      freightUSD: 3500,
      nestingMargin: 25,
      showNesting: false,
      activeTableTab: "currentNPS" // "currentNPS" or "master"
    };

    // Exact rounding helper matching Excel ROUND(val, digits)
    function excelRound(val, digits) {
      const f = Math.pow(10, digits);
      return Math.round((val + Number.EPSILON) * f) / f;
    }

    // Core calculation function matching Sheet 2 (2_Selectable_Master_Table) formulas 100%
    function calculateSpecRow(material, npsObj, schCode, rowId) {
      const schData = npsObj.schedules[schCode];
      if (!schData) return null;

      const od_in = npsObj.od_in;
      const od_mm = npsObj.od_mm;
      const wt_in = schData.wt_in;
      const wt_mm = schData.wt_mm;
      const id_mm = excelRound(od_mm - 2 * wt_mm, 2);

      const factorC = MATERIALS[material].factorC;
      const unit_wt = excelRound((od_mm - wt_mm) * wt_mm * factorC, 3);
      const pc_wt = excelRound(unit_wt * state.pipeLen, 2);

      const tiers = Math.floor(CONTAINER_LEN_M / state.pipeLen);
      const horiz_fit = Math.floor(state.effWidth / od_mm);

      let vert_layers = 0;
      if (horiz_fit > 0) {
        if (state.stacking === "Square Grid" || horiz_fit === 1) {
          vert_layers = Math.floor(state.effHeight / od_mm);
        } else {
          if (state.effHeight >= od_mm) {
            vert_layers = 1 + Math.floor((state.effHeight - od_mm) / (0.8660254 * od_mm));
          }
        }
      }

      let cross_section_pcs = 0;
      let even_row_pcs = horiz_fit;
      const wideEnoughForFullEvenRow = (state.effWidth - horiz_fit * od_mm) >= (0.5 * od_mm);

      if (horiz_fit > 0 && vert_layers > 0) {
        if (state.stacking === "Square Grid" || horiz_fit === 1) {
          cross_section_pcs = horiz_fit * vert_layers;
          even_row_pcs = horiz_fit;
        } else {
          even_row_pcs = wideEnoughForFullEvenRow ? horiz_fit : (horiz_fit - 1);
          const odd_layers = Math.floor((vert_layers + 1) / 2);
          const even_layers = Math.floor(vert_layers / 2);
          cross_section_pcs = odd_layers * horiz_fit + even_layers * even_row_pcs;
        }
      }

      const max_vol = cross_section_pcs * tiers;
      const max_wt = pc_wt > 0 ? Math.floor((state.maxPayloadMT * 1000) / pc_wt) : 0;
      const act_pcs = Math.min(max_vol, max_wt);
      const act_mt = excelRound((act_pcs * pc_wt) / 1000, 2);
      const wt_util = state.maxPayloadMT > 0 ? (act_mt / state.maxPayloadMT) : 0;
      const vol_util = max_vol > 0 ? (act_pcs / max_vol) : 0;
      const bottleneck = act_pcs === 0 ? "Exceeds Limit" : (max_wt <= max_vol ? "Weight Limited" : "Volume Limited");

      return {
        rowId,
        material,
        nps: npsObj.nps,
        dn: npsObj.dn,
        schedule: schCode,
        od_in,
        od_mm,
        wt_in,
        wt_mm,
        id_mm,
        unit_wt,
        pc_wt,
        tiers,
        horiz_fit,
        vert_layers,
        even_row_pcs,
        wideEnoughForFullEvenRow,
        cross_section_pcs,
        max_vol,
        max_wt,
        act_pcs,
        act_mt,
        wt_util,
        vol_util,
        bottleneck
      };
    }

    // Build all 1,176 rows in exact Tab 2 order
    function buildAllMasterRows() {
      const rows = [];
      let idCounter = 1;
      for (const mat of ["Carbon Steel", "Stainless Steel 304", "Stainless Steel 316"]) {
        for (const npsObj of PIPE_SPECS) {
          for (const sch of SCHEDULES) {
            if (npsObj.schedules[sch]) {
              rows.push(calculateSpecRow(mat, npsObj, sch, idCounter++));
            }
          }
        }
      }
      return rows;
    }

    // Find best inner pipe for telescoping/nesting when Volume Limited
    function findBestNestedPipe(outerRow) {
      if (!outerRow || outerRow.bottleneck !== "Volume Limited" || outerRow.id_mm <= 80) return null;
      const maxInnerOD = outerRow.id_mm - state.nestingMargin;
      const unusedPayloadKg = (state.maxPayloadMT - outerRow.act_mt) * 1000;
      if (unusedPayloadKg <= 50) return null;

      let best = null;
      for (const npsObj of PIPE_SPECS) {
        if (npsObj.od_mm <= maxInnerOD) {
          const sch = npsObj.schedules["40"] ? "40" : (npsObj.schedules["STD"] ? "STD" : Object.keys(npsObj.schedules)[0]);
          const innerCalc = calculateSpecRow(outerRow.material, npsObj, sch, 0);
          if (innerCalc) {
            const possibleInnerPcs = Math.min(outerRow.act_pcs, Math.floor(unusedPayloadKg / innerCalc.pc_wt));
            if (possibleInnerPcs > 0) {
              const extraMT = excelRound((possibleInnerPcs * innerCalc.pc_wt) / 1000, 2);
              best = {
                nps: npsObj.nps,
                schedule: sch,
                od_mm: npsObj.od_mm,
                wt_mm: innerCalc.wt_mm,
                id_mm: innerCalc.id_mm,
                pieces: possibleInnerPcs,
                extraMT
              };
            }
          }
        }
      }
      return best;
    }

    // Initialize UI Controls
    function initControls() {
      const npsSelect = document.getElementById("npsSelect");
      const filterNPS = document.getElementById("filterNPS");
      PIPE_SPECS.forEach(spec => {
        const opt = document.createElement("option");
        opt.value = spec.nps;
        opt.textContent = `NPS ${spec.nps}  (DN ${spec.dn} | OD ${spec.od_mm} mm)`;
        if (spec.nps === state.nps) opt.selected = true;
        npsSelect.appendChild(opt);

        const fOpt = document.createElement("option");
        fOpt.value = spec.nps;
        fOpt.textContent = `NPS ${spec.nps} (DN ${spec.dn})`;
        filterNPS.appendChild(fOpt);
      });

      const filterSch = document.getElementById("filterSch");
      SCHEDULES.forEach(s => {
        const fOpt = document.createElement("option");
        fOpt.value = s;
        fOpt.textContent = `Sch ${s}`;
        filterSch.appendChild(fOpt);
      });

      // Material Buttons
      document.querySelectorAll("#matGroup .seg-btn").forEach(btn => {
        btn.addEventListener("click", () => {
          state.material = btn.dataset.mat;
          updateAll();
        });
      });

      // NPS Select
      npsSelect.addEventListener("change", () => {
        state.nps = npsSelect.value;
        ensureValidScheduleForNPS();
        updateAll();
      });

      // Length Buttons
      document.querySelectorAll("#lenGroup .seg-btn").forEach(btn => {
        btn.addEventListener("click", () => {
          state.pipeLen = parseFloat(btn.dataset.len);
          updateAll();
        });
      });

      // Stacking Buttons
      document.querySelectorAll("#stackGroup .seg-btn").forEach(btn => {
        btn.addEventListener("click", () => {
          state.stacking = btn.dataset.stack;
          updateAll();
        });
      });

      // Payload Slider
      const payloadSlider = document.getElementById("payloadSlider");
      payloadSlider.addEventListener("input", () => {
        state.maxPayloadMT = parseFloat(payloadSlider.value);
        updateAll();
      });

      // Width, Height, Freight Inputs
      document.getElementById("effWidthInput").addEventListener("input", (e) => {
        const v = parseFloat(e.target.value);
        if (v > 0 && v <= 5000) { state.effWidth = v; updateAll(); }
      });
      document.getElementById("effHeightInput").addEventListener("input", (e) => {
        const v = parseFloat(e.target.value);
        if (v > 0 && v <= 5000) { state.effHeight = v; updateAll(); }
      });
      document.getElementById("freightInput").addEventListener("input", (e) => {
        const v = parseFloat(e.target.value);
        if (v >= 0) { state.freightUSD = v; updateAll(); }
      });

      // Nesting Toggle
      document.getElementById("nestingToggle").addEventListener("change", (e) => {
        state.showNesting = e.target.checked;
        updateAll();
      });

      // Reset Button
      document.getElementById("resetBtn").addEventListener("click", () => {
        state.maxPayloadMT = 26.00;
        state.effWidth = 2300;
        state.effHeight = 2280;
        state.freightUSD = 3500;
        state.pipeLen = 11.80;
        state.stacking = "Staggered (Hexagonal)";
        document.getElementById("payloadSlider").value = "26.0";
        document.getElementById("effWidthInput").value = "2300";
        document.getElementById("effHeightInput").value = "2280";
        document.getElementById("freightInput").value = "3500";
        updateAll();
      });

      // Table Tabs
      document.getElementById("tabBtnCurrentNPS").addEventListener("click", () => {
        state.activeTableTab = "currentNPS";
        document.getElementById("tabBtnCurrentNPS").classList.add("active");
        document.getElementById("tabBtnMaster").classList.remove("active");
        document.getElementById("masterFilterBar").style.display = "none";
        renderTable();
      });
      document.getElementById("tabBtnMaster").addEventListener("click", () => {
        state.activeTableTab = "master";
        document.getElementById("tabBtnMaster").classList.add("active");
        document.getElementById("tabBtnCurrentNPS").classList.remove("active");
        document.getElementById("masterFilterBar").style.display = "flex";
        renderTable();
      });

      ["filterMat", "filterNPS", "filterSch", "filterLimit"].forEach(id => {
        document.getElementById(id).addEventListener("change", renderTable);
      });
    }

    function ensureValidScheduleForNPS() {
      const npsObj = PIPE_SPECS.find(s => s.nps === state.nps);
      if (!npsObj) return;
      if (!npsObj.schedules[state.schedule]) {
        const preferred = ["40", "STD", "40S", "10", "10S", "80", "XS"];
        for (const p of preferred) {
          if (npsObj.schedules[p]) {
            state.schedule = p;
            return;
          }
        }
        state.schedule = Object.keys(npsObj.schedules)[0];
      }
    }

    function renderSchedulePills() {
      const schGrid = document.getElementById("schGrid");
      schGrid.innerHTML = "";
      const npsObj = PIPE_SPECS.find(s => s.nps === state.nps);
      SCHEDULES.forEach(sch => {
        const btn = document.createElement("button");
        btn.className = "sch-pill" + (STAINLESS_SCH_SET.has(sch) ? " ss-type" : "");
        btn.textContent = sch;
        const exists = npsObj && npsObj.schedules[sch];
        if (!exists) {
          btn.disabled = true;
          btn.title = `Schedule ${sch} is not defined in ASME for NPS ${state.nps}`;
        } else {
          btn.title = `Select Schedule ${sch} (WT: ${npsObj.schedules[sch].wt_mm} mm)`;
          if (sch === state.schedule) btn.classList.add("active");
          btn.addEventListener("click", () => {
            state.schedule = sch;
            updateAll();
          });
        }
        schGrid.appendChild(btn);
      });
    }

    // Draw 2D Cross-Section Canvas
    function drawCrossSection(currentRow, nestedInfo) {
      const canvas = document.getElementById("crossCanvas");
      const ctx = canvas.getContext("2d");
      const W = canvas.width;
      const H = canvas.height;
      ctx.clearRect(0, 0, W, H);

      // Margins for dimension labels
      const padLeft = 58;
      const padRight = 32;
      const padTop = 36;
      const padBottom = 46;

      const availW = W - padLeft - padRight;
      const availH = H - padTop - padBottom;
      const scale = Math.min(availW / CONTAINER_INT_W, availH / CONTAINER_INT_H);

      const boxW = CONTAINER_INT_W * scale;
      const boxH = CONTAINER_INT_H * scale;
      const originX = padLeft + (availW - boxW) / 2;
      const originY = padTop + (availH - boxH) / 2;

      // 1. Draw Container Outer Frame (2352 x 2393 mm)
      ctx.fillStyle = "#0b1324";
      ctx.fillRect(originX, originY, boxW, boxH);

      ctx.strokeStyle = "#475569";
      ctx.lineWidth = 3;
      ctx.strokeRect(originX, originY, boxW, boxH);

      // Draw Wooden Skids / Dunnage at bottom floor
      const skidH = Math.max(4, (CONTAINER_INT_H - state.effHeight) * 0.35 * scale);
      ctx.fillStyle = "#78350f";
      const skidPositions = [0.12, 0.45, 0.78];
      skidPositions.forEach(pos => {
        ctx.fillRect(originX + boxW * pos, originY + boxH - skidH, boxW * 0.1, skidH);
      });

      // 2. Draw Effective Loading Zone (effWidth x effHeight)
      const effW_px = state.effWidth * scale;
      const effH_px = state.effHeight * scale;
      const effX = originX + (boxW - effW_px) / 2;
      const effY = originY + boxH - skidH - effH_px;

      ctx.save();
      ctx.strokeStyle = "rgba(56, 189, 248, 0.4)";
      ctx.lineWidth = 1.2;
      ctx.setLineDash([6, 4]);
      ctx.strokeRect(effX, effY, effW_px, effH_px);
      ctx.restore();

      // Dimension Annotations
      ctx.fillStyle = "#94a3b8";
      ctx.font = "600 11px 'Segoe UI', sans-serif";
      ctx.textAlign = "center";
      ctx.fillText(`40GP Internal Width: 2,352 mm (Effective Loading Width: ${state.effWidth} mm)`, originX + boxW / 2, originY + boxH + 26);

      ctx.save();
      ctx.translate(originX - 22, originY + boxH / 2);
      ctx.rotate(-Math.PI / 2);
      ctx.fillText(`Internal Height: 2,393 mm (Effective: ${state.effHeight} mm)`, 0, 0);
      ctx.restore();

      if (!currentRow || currentRow.horiz_fit === 0 || currentRow.vert_layers === 0) return;

      // 3. Compute Pipe Center Coordinates in Cross-Section (1 Tier)
      const od_px = currentRow.od_mm * scale;
      const r_outer = od_px / 2;
      const r_inner = Math.max(0, (currentRow.id_mm * scale) / 2);
      const loadedPerTier = Math.ceil(currentRow.act_pcs / currentRow.tiers);

      const matTheme = MATERIALS[state.material];
      const isStaggered = (state.stacking === "Staggered (Hexagonal)" && currentRow.horiz_fit >= 2);
      const dy = isStaggered ? (0.8660254 * od_px) : od_px;

      let slotIndex = 0;
      let cutoffY = null;

      for (let layer = 0; layer < currentRow.vert_layers; layer++) {
        const isEvenLayer = (layer % 2 === 1);
        const countInLayer = isStaggered ? (isEvenLayer ? currentRow.even_row_pcs : currentRow.horiz_fit) : currentRow.horiz_fit;
        const xOffset = (isStaggered && isEvenLayer) ? r_outer : 0;
        const cy = (effY + effH_px) - r_outer - layer * dy;

        for (let col = 0; col < countInLayer; col++) {
          const cx = effX + r_outer + xOffset + col * od_px;
          const isLoaded = (slotIndex < loadedPerTier);

          if (r_outer >= 2.2) {
            ctx.beginPath();
            ctx.arc(cx, cy, Math.max(1, r_outer - 0.4), 0, Math.PI * 2);

            if (isLoaded) {
              // Metallic Pipe Wall Ring
              ctx.fillStyle = matTheme.colorOuter;
              ctx.fill();
              if (r_outer >= 5) {
                ctx.lineWidth = Math.min(1.5, r_outer * 0.12);
                ctx.strokeStyle = matTheme.stroke;
                ctx.stroke();
              }
              // Hollow Pipe Inner Bore
              if (r_inner >= 1.2) {
                ctx.beginPath();
                ctx.arc(cx, cy, r_inner, 0, Math.PI * 2);
                ctx.fillStyle = "#070b14";
                ctx.fill();
              }
              // Optional Nested Inner Pipe
              if (state.showNesting && nestedInfo && slotIndex < Math.ceil(nestedInfo.pieces / currentRow.tiers)) {
                const r_nest_out = (nestedInfo.od_mm * scale) / 2;
                const r_nest_in = (nestedInfo.id_mm * scale) / 2;
                if (r_nest_out >= 1.5) {
                  ctx.beginPath();
                  ctx.arc(cx, cy + (r_inner - r_nest_out) * 0.65, r_nest_out, 0, Math.PI * 2);
                  ctx.fillStyle = "#f59e0b";
                  ctx.fill();
                  if (r_nest_in >= 1.0) {
                    ctx.beginPath();
                    ctx.arc(cx, cy + (r_inner - r_nest_out) * 0.65, r_nest_in, 0, Math.PI * 2);
                    ctx.fillStyle = "#070b14";
                    ctx.fill();
                  }
                }
              }
            } else {
              // Ghost Slot (Unloaded due to 26MT Weight Limit)
              ctx.save();
              ctx.strokeStyle = "rgba(148, 163, 184, 0.28)";
              ctx.lineWidth = 1;
              if (r_outer >= 6) ctx.setLineDash([3, 3]);
              ctx.stroke();
              ctx.restore();
              if (cutoffY === null) cutoffY = cy + r_outer;
            }
          } else {
            // Ultra-fast pixel block rendering for tiny pipes (e.g. 1/8", 1/4")
            ctx.fillStyle = isLoaded ? matTheme.colorOuter : "rgba(148, 163, 184, 0.18)";
            ctx.fillRect(cx - r_outer, cy - r_outer, od_px * 0.9, od_px * 0.9);
            if (!isLoaded && cutoffY === null) cutoffY = cy + r_outer;
          }

          slotIndex++;
        }
      }

      // 4. Draw Red/Amber Weight Limit Cutoff Line if Weight Limited
      if (currentRow.bottleneck === "Weight Limited" && cutoffY !== null && currentRow.act_pcs < currentRow.max_vol) {
        ctx.save();
        ctx.strokeStyle = "#f59e0b";
        ctx.lineWidth = 2;
        ctx.setLineDash([8, 4]);
        ctx.beginPath();
        ctx.moveTo(effX - 6, cutoffY);
        ctx.lineTo(effX + effW_px + 6, cutoffY);
        ctx.stroke();

        const tagText = `MAX PAYLOAD LIMIT REACHED (${state.maxPayloadMT.toFixed(2)} MT = ${currentRow.act_pcs.toLocaleString()} pcs total)`;
        ctx.font = "700 11px 'Segoe UI', sans-serif";
        const tw = ctx.measureText(tagText).width + 16;
        const tx = originX + boxW / 2 - tw / 2;
        const ty = Math.max(originY + 8, cutoffY - 22);
        ctx.fillStyle = "rgba(180, 83, 9, 0.92)";
        ctx.fillRect(tx, ty, tw, 20);
        ctx.strokeStyle = "#fbbf24";
        ctx.setLineDash([]);
        ctx.lineWidth = 1;
        ctx.strokeRect(tx, ty, tw, 20);
        ctx.fillStyle = "#ffffff";
        ctx.textAlign = "center";
        ctx.fillText(tagText, originX + boxW / 2, ty + 14);
        ctx.restore();
      }
    }

    // Draw Lengthwise Side Elevation View
    function drawSideElevation(currentRow) {
      const canvas = document.getElementById("sideCanvas");
      const ctx = canvas.getContext("2d");
      const W = canvas.width;
      const H = canvas.height;
      ctx.clearRect(0, 0, W, H);

      const padX = 28;
      const padY = 18;
      const boxW = W - padX * 2;
      const boxH = H - padY * 2 - 14;
      const scaleX = boxW / CONTAINER_LEN_M;

      // Container Side Frame
      ctx.fillStyle = "#0b1324";
      ctx.fillRect(padX, padY, boxW, boxH);
      ctx.strokeStyle = "#475569";
      ctx.lineWidth = 2;
      ctx.strokeRect(padX, padY, boxW, boxH);

      if (!currentRow || currentRow.act_pcs === 0) return;

      const matTheme = MATERIALS[state.material];
      const pipeLenPx = state.pipeLen * scaleX;
      const gapPx = 4;

      // Calculate filled height ratio from loaded layers
      const loadedPerTier = Math.ceil(currentRow.act_pcs / currentRow.tiers);
      const fillRatio = Math.min(1, currentRow.max_vol > 0 ? (loadedPerTier / currentRow.cross_section_pcs) : 0);
      const totalStackH = (state.effHeight / CONTAINER_INT_H) * boxH;
      const loadedStackH = Math.max(6, totalStackH * fillRatio);

      for (let t = 0; t < currentRow.tiers; t++) {
        const startX = padX + 4 + t * (pipeLenPx + gapPx);
        const startY = padY + boxH - loadedStackH - 3;

        // Ghost Volume Box
        ctx.strokeStyle = "rgba(148,163,184,0.25)";
        ctx.setLineDash([4, 3]);
        ctx.strokeRect(startX, padY + boxH - totalStackH - 3, pipeLenPx, totalStackH);
        ctx.setLineDash([]);

        // Loaded Stack Box
        ctx.fillStyle = matTheme.colorInner;
        ctx.fillRect(startX, startY, pipeLenPx, loadedStackH);
        ctx.strokeStyle = matTheme.stroke;
        ctx.lineWidth = 1;
        ctx.strokeRect(startX, startY, pipeLenPx, loadedStackH);

        // Horizontal layer stripes
        const stripeCount = Math.min(10, currentRow.vert_layers);
        ctx.strokeStyle = "rgba(255,255,255,0.18)";
        for (let s = 1; s < stripeCount; s++) {
          const sy = startY + (loadedStackH / stripeCount) * s;
          ctx.beginPath();
          ctx.moveTo(startX, sy);
          ctx.lineTo(startX + pipeLenPx, sy);
          ctx.stroke();
        }

        // Tier label
        ctx.fillStyle = "#ffffff";
        ctx.font = "700 11px 'Segoe UI', sans-serif";
        ctx.textAlign = "center";
        const pcsInThisTier = t === 0 ? Math.ceil(currentRow.act_pcs / currentRow.tiers) : Math.floor(currentRow.act_pcs / currentRow.tiers);
        ctx.fillText(`Tier ${t + 1}: ${state.pipeLen.toFixed(2)}m (${pcsInThisTier.toLocaleString()} pcs)`, startX + pipeLenPx / 2, startY + loadedStackH / 2 + 4);
      }

      // Bottom length annotation
      const usedLenM = (currentRow.tiers * state.pipeLen).toFixed(2);
      const clearanceM = (CONTAINER_LEN_M - currentRow.tiers * state.pipeLen).toFixed(2);
      ctx.fillStyle = "#94a3b8";
      ctx.font = "600 11px 'Segoe UI', sans-serif";
      ctx.textAlign = "center";
      ctx.fillText(`Used Length: ${usedLenM} m (${currentRow.tiers} Tier${currentRow.tiers > 1 ? 's' : ''})  |  Door Clearance: ${clearanceM} m`, padX + boxW / 2, H - 6);
    }

    // Render Bottom Data Table
    function renderTable() {
      const tbody = document.getElementById("tableBody");
      tbody.innerHTML = "";
      const allRows = buildAllMasterRows();

      let displayRows = [];
      if (state.activeTableTab === "currentNPS") {
        displayRows = allRows.filter(r => r.material === state.material && r.nps === state.nps);
      } else {
        const fMat = document.getElementById("filterMat").value;
        const fNPS = document.getElementById("filterNPS").value;
        const fSch = document.getElementById("filterSch").value;
        const fLim = document.getElementById("filterLimit").value;

        displayRows = allRows.filter(r => {
          if (fMat !== "ALL" && r.material !== fMat) return false;
          if (fNPS !== "ALL" && r.nps !== fNPS) return false;
          if (fSch !== "ALL" && r.schedule !== fSch) return false;
          if (fLim !== "ALL" && r.bottleneck !== fLim) return false;
          return true;
        });
      }

      displayRows.forEach(r => {
        const tr = document.createElement("tr");
        if (r.material === state.material && r.nps === state.nps && r.schedule === state.schedule) {
          tr.classList.add("selected-row");
        }
        tr.innerHTML = `
          <td>#${r.rowId}</td>
          <td style="font-weight:600; color:${r.material === 'Carbon Steel' ? '#93c5fd' : '#5eead4'}">${r.material}</td>
          <td style="font-weight:700;">${r.nps}</td>
          <td style="font-weight:700;">${r.schedule}</td>
          <td>${r.od_mm.toFixed(2)}</td>
          <td>${r.wt_in.toFixed(3)}</td>
          <td>${r.wt_mm.toFixed(2)}</td>
          <td>${r.id_mm.toFixed(2)}</td>
          <td>${r.unit_wt.toFixed(3)}</td>
          <td>${r.pc_wt.toFixed(2)}</td>
          <td>${r.horiz_fit}</td>
          <td>${r.vert_layers}</td>
          <td>${r.max_vol.toLocaleString()}</td>
          <td>${r.max_wt.toLocaleString()}</td>
          <td class="col-out">${r.act_pcs.toLocaleString()}</td>
          <td class="col-out">${r.act_mt.toFixed(2)}</td>
          <td>${(r.wt_util * 100).toFixed(1)}%</td>
          <td><span class="${r.bottleneck === 'Weight Limited' ? 'badge-wt' : 'badge-vol'}">${r.bottleneck}</span></td>
        `;
        tr.addEventListener("click", () => {
          state.material = r.material;
          state.nps = r.nps;
          state.schedule = r.schedule;
          document.getElementById("npsSelect").value = state.nps;
          updateAll();
        });
        tbody.appendChild(tr);
      });
    }

    // Master Update Function
    function updateAll() {
      // Sync Segmented Button Active States
      document.querySelectorAll("#matGroup .seg-btn").forEach(btn => {
        btn.className = "seg-btn";
        if (btn.dataset.mat === state.material) {
          btn.classList.add(state.material === "Carbon Steel" ? "active" : "active-ss");
        }
      });
      document.querySelectorAll("#lenGroup .seg-btn").forEach(btn => {
        btn.classList.toggle("active", parseFloat(btn.dataset.len) === state.pipeLen);
      });
      document.querySelectorAll("#stackGroup .seg-btn").forEach(btn => {
        btn.classList.toggle("active", btn.dataset.stack === state.stacking);
      });

      renderSchedulePills();

      const allRows = buildAllMasterRows();
      const currentRow = allRows.find(r => r.material === state.material && r.nps === state.nps && r.schedule === state.schedule);
      if (!currentRow) return;

      const nestedInfo = findBestNestedPipe(currentRow);

      // Update Sidebar Badges
      document.getElementById("matchedRowBadge").textContent = `Tab 2 Row #${currentRow.rowId}`;
      document.getElementById("densityBadge").textContent = `${MATERIALS[state.material].density.toFixed(2)} g/cm³`;
      document.getElementById("odBadge").textContent = `OD: ${currentRow.od_mm.toFixed(2)} mm`;
      document.getElementById("wtBadge").textContent = `WT: ${currentRow.wt_mm.toFixed(2)} mm`;
      document.getElementById("tierBadge").textContent = `${currentRow.tiers} Lengthwise Tier${currentRow.tiers > 1 ? 's' : ''}`;
      document.getElementById("payloadValLabel").textContent = `${state.maxPayloadMT.toFixed(2)} MT (${(state.maxPayloadMT * 1000).toLocaleString()} kg)`;

      // Update Top KPI Cards
      const totalLenM = currentRow.act_pcs * state.pipeLen;
      const totalLenFt = totalLenM * 3.28084;
      document.getElementById("kpiPieces").textContent = `${currentRow.act_pcs.toLocaleString()} pcs`;
      document.getElementById("kpiLengthSub").textContent = `Total Length: ${totalLenM.toLocaleString(undefined, {maximumFractionDigits:1})} m (${Math.round(totalLenFt).toLocaleString()} ft)`;

      document.getElementById("kpiTons").textContent = `${currentRow.act_mt.toFixed(2)} MT`;
      document.getElementById("kpiWeightSub").textContent = `Net Weight: ${Math.round(currentRow.act_mt * 1000).toLocaleString()} kg (${Math.round(currentRow.act_mt * 2204.62).toLocaleString()} lbs)`;

      const kpiBottleneck = document.getElementById("kpiBottleneck");
      kpiBottleneck.textContent = currentRow.bottleneck.toUpperCase();
      kpiBottleneck.style.color = currentRow.bottleneck === "Weight Limited" ? "#fbbf24" : "#38bdf8";
      document.getElementById("kpiLimitSub").textContent = `Wt Limit: ${currentRow.max_wt.toLocaleString()} pcs | Vol Limit: ${currentRow.max_vol.toLocaleString()} pcs`;

      const freightPerMT = currentRow.act_mt > 0 ? (state.freightUSD / currentRow.act_mt) : 0;
      const freightPerPc = currentRow.act_pcs > 0 ? (state.freightUSD / currentRow.act_pcs) : 0;
      const freightPerM = totalLenM > 0 ? (state.freightUSD / totalLenM) : 0;
      document.getElementById("kpiFreightMT").textContent = `$${freightPerMT.toFixed(2)} / MT`;
      document.getElementById("kpiFreightPc").textContent = `$${freightPerPc.toFixed(2)} / pc | $${freightPerM.toFixed(2)} / m`;
      document.getElementById("freightPerMeterBadge").textContent = `$${freightPerM.toFixed(2)} / m`;

      // Update Specs & Utilization Bars
      document.getElementById("specOD").textContent = `${currentRow.od_mm.toFixed(2)} mm (${currentRow.od_in.toFixed(3)}")`;
      document.getElementById("specWT").textContent = `${currentRow.wt_mm.toFixed(2)} mm (${currentRow.wt_in.toFixed(3)}")`;
      document.getElementById("specID").textContent = `${currentRow.id_mm.toFixed(2)} mm (${(currentRow.id_mm / 25.4).toFixed(3)}")`;
      document.getElementById("specUnitWt").textContent = `${currentRow.unit_wt.toFixed(3)} kg/m (${(currentRow.unit_wt * 0.671969).toFixed(2)} lb/ft)`;
      document.getElementById("specPieceWt").textContent = `${currentRow.pc_wt.toFixed(2)} kg / pc`;
      document.getElementById("specGrid").textContent = `${currentRow.horiz_fit} horiz × ${currentRow.vert_layers} layers (${currentRow.cross_section_pcs}/tier)`;

      const wtPct = Math.min(100, currentRow.wt_util * 100);
      const volPct = Math.min(100, currentRow.vol_util * 100);
      document.getElementById("wtUtilText").textContent = `${wtPct.toFixed(1)}% (${currentRow.act_mt.toFixed(2)} / ${state.maxPayloadMT.toFixed(2)} MT)`;
      document.getElementById("wtUtilBar").style.width = `${wtPct}%`;
      document.getElementById("volUtilText").textContent = `${volPct.toFixed(1)}% (${currentRow.act_pcs.toLocaleString()} / ${currentRow.max_vol.toLocaleString()} slots)`;
      document.getElementById("volUtilBar").style.width = `${volPct}%`;

      // Canvas Titles & Legend
      const loadedPerTier = Math.ceil(currentRow.act_pcs / currentRow.tiers);
      document.getElementById("canvasTitle").textContent = `40GP Cross-Section End View (${loadedPerTier.toLocaleString()} / ${currentRow.cross_section_pcs.toLocaleString()} pcs per tier)`;
      document.getElementById("sideTierText").textContent = `${currentRow.tiers} Tier${currentRow.tiers > 1 ? 's' : ''} × ${state.pipeLen.toFixed(2)} m = ${currentRow.act_pcs.toLocaleString()} pcs`;
      document.getElementById("nestedLegend").style.display = (state.showNesting && nestedInfo) ? "inline" : "none";

      // Commercial Advisor & Nesting Text
      const advisor = document.getElementById("advisorBanner");
      const nestSub = document.getElementById("nestingSubtext");
      if (currentRow.bottleneck === "Weight Limited") {
        advisor.innerHTML = `<strong>COMMERCIAL ADVISOR [WEIGHT LIMITED]:</strong> Container reaches max payload (<strong>${currentRow.act_mt.toFixed(2)} MT</strong>) at <strong>${volPct.toFixed(1)}%</strong> space occupancy (${(currentRow.max_vol - currentRow.act_pcs).toLocaleString()} empty pipe slots shown as dashed circles above the orange cutoff line).`;
        nestSub.textContent = "Current pipe is Weight Limited (container hits 26 MT before space fills up). Select a larger/thinner pipe (e.g. NPS 16'' or 24'') to test Pipe Nesting.";
      } else if (nestedInfo) {
        advisor.innerHTML = `<strong>COMMERCIAL ADVISOR [VOLUME LIMITED — TELESCOPING OPPORTUNITY]:</strong> Space is 100% full at <strong>${currentRow.act_mt.toFixed(2)} MT</strong> (leaving <strong>${(state.maxPayloadMT - currentRow.act_mt).toFixed(2)} MT</strong> unused free weight capacity!). Inner Diameter is <strong>${currentRow.id_mm.toFixed(1)} mm</strong> &rarr; You can telescope <strong>${nestedInfo.pieces} pcs of NPS ${nestedInfo.nps} Sch ${nestedInfo.schedule} (+${nestedInfo.extraMT.toFixed(2)} MT)</strong> inside for FREE ocean freight!`;
        nestSub.innerHTML = `Recommended Inner Pipe: <strong style="color:#fbbf24;">NPS ${nestedInfo.nps} Sch ${nestedInfo.schedule}</strong> (OD ${nestedInfo.od_mm} mm &le; Max ${ (currentRow.id_mm - state.nestingMargin).toFixed(0) } mm) &rarr; Adds <strong style="color:#34d399;">+${nestedInfo.pieces} pcs (+${nestedInfo.extraMT} MT)</strong>!`;
      } else {
        advisor.innerHTML = `<strong>COMMERCIAL ADVISOR [VOLUME LIMITED]:</strong> Container space is 100% utilized (<strong>${currentRow.act_pcs.toLocaleString()} pcs / ${currentRow.act_mt.toFixed(2)} MT</strong>).`;
        nestSub.textContent = "Container is Volume Limited.";
      }

      drawCrossSection(currentRow, nestedInfo);
      drawSideElevation(currentRow);
      renderTable();
    }

    // Boot application
    initControls();
    updateAll();
  </script>
</body>
</html>
"""

html_out = HTML_TEMPLATE.replace("__PIPE_SPECS_JSON__", json.dumps(specs_json_list))
html_out = html_out.replace("__SCHEDULES_JSON__", json.dumps(SCHEDULES))

output_path = "Precise_Pipe_40GP_Visualizer.html"
with open(output_path, "w", encoding="utf-8") as f:
    f.write(html_out)

print(f"Successfully generated {output_path} ({len(html_out):,} bytes)!")
