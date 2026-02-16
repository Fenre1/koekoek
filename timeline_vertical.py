#%%
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, time as dt_time
from pathlib import Path
import html
import json
import pandas as pd


@dataclass
class TimelineEvent:
    event_id: int
    description: str
    entities: list[str]
    date_label: str
    time_label: str
    start_dt: datetime | None
    end_dt: datetime | None
    certain: bool
    verified: bool


# -------------------------
# Parsing / formatting utils
# -------------------------
def _normalize_bool(value: object) -> bool:
    if value is None or pd.isna(value):
        return False
    return str(value).strip().lower() == "ja"


def _parse_date_time(date_value: object, time_value: object) -> datetime | None:
    if date_value is None or pd.isna(date_value):
        return None

    date_ts = pd.to_datetime(date_value, errors="coerce")
    if pd.isna(date_ts):
        return None
    date_ts = date_ts.normalize()

    if time_value is None or pd.isna(time_value):
        return date_ts.to_pydatetime()

    if isinstance(time_value, dt_time):
        return datetime.combine(date_ts.date(), time_value)

    if isinstance(time_value, (pd.Timestamp, datetime)):
        return datetime.combine(date_ts.date(), pd.to_datetime(time_value).time())

    if isinstance(time_value, (int, float)) and not pd.isna(time_value):
        t = pd.to_datetime(time_value, unit="D", origin="1899-12-30", errors="coerce")
        if pd.isna(t):
            return date_ts.to_pydatetime()
        return datetime.combine(date_ts.date(), t.time())

    t = pd.to_datetime(str(time_value).strip(), errors="coerce")
    if pd.isna(t):
        return date_ts.to_pydatetime()
    return datetime.combine(date_ts.date(), t.time())


def _format_date_label(date_value: object) -> str:
    if date_value is None or pd.isna(date_value):
        return "Onbekende datum"
    date = pd.to_datetime(date_value, errors="coerce")
    if pd.isna(date):
        return "Onbekende datum"
    return date.strftime("%Y-%m-%d")


def _format_time_label(start_dt: datetime | None, end_dt: datetime | None) -> str:
    if start_dt is None:
        return "Onbekende tijd"
    start_time = start_dt.strftime("%H:%M")
    if end_dt is None or end_dt == start_dt:
        return start_time
    return f"{start_time} - {end_dt.strftime('%H:%M')}"


def _split_entities(value: object) -> list[str]:
    if value is None or pd.isna(value):
        return ["Onbekend"]
    entities = [item.strip() for item in str(value).split("|") if item.strip()]
    return entities or ["Onbekend"]


def _is_range(e: TimelineEvent) -> bool:
    return e.start_dt is not None and e.end_dt is not None and e.end_dt != e.start_dt


def _entity_color(entity: str) -> str:
    h = 0
    for ch in entity:
        h = (h * 31 + ord(ch)) % 360
    return f"hsl({h}, 55%, 72%)"


def _truncate(s: str, n: int = 140) -> str:
    s = s.strip()
    return s if len(s) <= n else s[: n - 1] + "…"


# -------------------------
# Main
# -------------------------
def generate_vertical_timeline(excel_path: str | Path, output_path: str | Path) -> None:
    df = pd.read_excel(excel_path)

    required_columns = {
        "Datum",
        "Starttijd",
        "Eindtijd",
        "Zekerheid (ja/nee)",
        "Entiteit(en) (splits op met |)",
        "Gebeurtenis",
        "Geverifieerd",
    }
    missing = required_columns.difference(df.columns)
    if missing:
        raise ValueError(f"Missing columns in Excel file: {', '.join(sorted(missing))}")

    # Read events (base events may have multiple entities)
    base_events: list[TimelineEvent] = []
    for idx, row in df.iterrows():
        start_dt = _parse_date_time(row["Datum"], row["Starttijd"])
        end_dt = _parse_date_time(row["Datum"], row["Eindtijd"]) if not pd.isna(row["Eindtijd"]) else None
        description = str(row["Gebeurtenis"]) if not pd.isna(row["Gebeurtenis"]) else ""

        base_events.append(
            TimelineEvent(
                event_id=int(idx),
                description=description,
                entities=_split_entities(row["Entiteit(en) (splits op met |)"]),
                date_label=_format_date_label(row["Datum"]),
                time_label=_format_time_label(start_dt, end_dt),
                start_dt=start_dt,
                end_dt=end_dt if (start_dt and end_dt) else None,
                certain=_normalize_bool(row["Zekerheid (ja/nee)"]),
                verified=_normalize_bool(row["Geverifieerd"]),
            )
        )

    # Expand into per-entity "instances" (so filtering is clean & colouring is per card)
    # One event per row (keep multiple entities inside event)
    events_payload = []
    all_entities = set()

    for e in base_events:
        if e.start_dt is None:
            continue

        for ent in e.entities:
            all_entities.add(ent)

        events_payload.append(
            {
                "id": str(e.event_id),
                "entities": e.entities,
                "colors": {ent: _entity_color(ent) for ent in e.entities},
                "date_label": e.date_label,
                "time_label": e.time_label,
                "start_ms": int(pd.Timestamp(e.start_dt).value // 1_000_000),
                "end_ms": int(pd.Timestamp(e.end_dt).value // 1_000_000) if e.end_dt else int(pd.Timestamp(e.start_dt).value // 1_000_000),
                "is_range": _is_range(e),
                "certain": e.certain,
                "verified": e.verified,
                "desc_full": e.description,
                "desc_short": _truncate(e.description, 140),
            }
        )

    entities_sorted = sorted(all_entities, key=lambda s: s.lower())

    payload = {
        "entities": [
            {"name": ent, "color": _entity_color(ent)}
            for ent in entities_sorted
        ],
        "events": events_payload,
    }

    payload_json = json.dumps(payload, ensure_ascii=False)

    # HTML
    html_parts = [
        "<!DOCTYPE html>",
        "<html lang='nl'>",
        "<head>",
        "<meta charset='utf-8' />",
        "<meta name='viewport' content='width=device-width, initial-scale=1' />",
        "<title>Vertical Timeline</title>",
        "<style>",
        "  :root {",
        "    --bg: #ffffff;",
        "    --ink: #2f3e46;",
        "    --muted: rgba(47,62,70,0.55);",
        "    --panel: rgba(47,62,70,0.04);",
        "    --card: #f9fbfb;",
        "    --range: #eef6f7;",
        "    --sidebar-w: 320px;",
        "    --wrap-col-w: 240px;",
        "    --wrap-area: 0px; /* JS sets this to (#cols * colwidth + gaps) */",
        "    --gap-x: 20px;",
        "    --row-h: 98px;",
        "    --row-gap: 12px;",
        "    --indent: 16px;",
        "    --max-subcols: 4;",
        "    --subcol-w: 340px;",
        "  }",
        "  body { font-family: 'Segoe UI', sans-serif; margin: 0; color: var(--ink); background: var(--bg); }",
        "  .app { display: grid; grid-template-columns: var(--sidebar-w) 1fr; min-height: 100vh; }",
        "",
        "  /* Sidebar */",
        "  .sidebar { position: sticky; top: 0; align-self: start; height: 100vh; overflow: auto;",
        "    border-right: 2px solid rgba(47,62,70,0.12); background: #fff; }",
        "  .sidebar-inner { padding: 16px 14px 18px; }",
        "  .title { font-size: 16px; font-weight: 700; margin: 0 0 10px; }",
        "  .sub { font-size: 12px; color: var(--muted); margin: 0 0 14px; line-height: 1.35; }",
        "  .controls { display: flex; gap: 8px; margin-bottom: 12px; }",
        "  .btn { border: 2px solid rgba(47,62,70,0.18); background: #fff; color: var(--ink);",
        "    padding: 6px 10px; border-radius: 10px; cursor: pointer; font-weight: 600; font-size: 12px; }",
        "  .btn:hover { background: rgba(47,62,70,0.04); }",
        "  .entity-list { display: flex; flex-direction: column; gap: 6px; }",
        "  .entity-item { display: flex; align-items: center; gap: 10px; padding: 6px 8px; border-radius: 10px; }",
        "  .entity-item:hover { background: rgba(47,62,70,0.04); }",
        "  .swatch { width: 14px; height: 14px; border-radius: 5px; border: 2px solid rgba(47,62,70,0.25); }",
        "  .entity-name { font-size: 13px; font-weight: 600; overflow-wrap: anywhere; }",
        "  .count { margin-left: auto; font-size: 12px; color: var(--muted); }",
        "",
        "  /* Main */",
        "  .main { padding: 18px 18px 60px; }",
        "  .legend-top { display: flex; gap: 18px; align-items: center; margin-bottom: 14px; }",
        "  .chip { display: inline-flex; align-items: center; gap: 8px; font-size: 12px; }",
        "  .chip-box { width: 18px; height: 12px; border: 2px solid var(--ink); border-radius: 4px; background: #fff; }",
        "  .chip-box.uncertain { border-style: dashed; }",
        "  .chip-box.unverified { border-color: #9aa0a6; background: #f2f3f5; }",
        "",
        "  .viewport { position: relative; }",
        "  .timeline { position: relative; min-height: 400px; }",
        "",
        "  /* Range wrapper cards */",
        "  .range-wrap { position: absolute; left: 0; width: var(--wrap-col-w);",
        "    border: 2px solid var(--ink); border-left-width: 10px; border-radius: 18px;",
        "    background: var(--range); box-sizing: border-box; padding: 10px 12px; }",
        "  .range-wrap.uncertain { border-style: dashed; }",
        "  .range-wrap.unverified { border-color: #9aa0a6; background: #f2f3f5; }",
        "",
        "  /* Normal cards (and also range header inside wrapper) */",
        "  .card { position: absolute; border: 2px solid var(--ink); border-left-width: 10px; border-radius: 18px;",
        "    background: var(--card); box-sizing: border-box; padding: 10px 12px; overflow: hidden; }",
        "  .card.uncertain { border-style: dashed; }",
        "  .card.unverified { border-color: #9aa0a6; background: #f2f3f5; }",
        "",
        "  .card:hover, .range-wrap:hover { overflow: visible; z-index: 500; }",
        "",
        "  /* Stacked (same-time) card */",
        "  .stack-card {",
        "    position: absolute;",
        "    border: 2px solid var(--ink);",
        "    border-left-width: 10px;",
        "    border-left-color: #000;",
        "    border-radius: 18px;",
        "    background: var(--card);",
        "    box-sizing: border-box;",
        "    padding: 10px 12px;",
        "    overflow: hidden;",
        "  }",
        "  .stack-card:hover { overflow: visible; z-index: 500; }",
        "",
        "  .stack-hdr {",
        "    display: flex;",
        "    flex-wrap: wrap;",
        "    gap: 10px;",
        "    align-items: baseline;",
        "    margin-bottom: 8px;",
        "  }",
        "",
        "  .stack-items { display: flex; flex-direction: column; gap: 10px; }",
        "",
        "  .stack-item {",
        "    border: 2px solid var(--ink);",
        "    border-left-width: 10px;",
        "    border-radius: 16px;",
        "    background: #fff;",
        "    box-sizing: border-box;",
        "    padding: 8px 10px;",
        "    position: relative;",
        "    overflow: hidden;",
        "  }",
        "  .stack-item.uncertain { border-style: dashed; }",
        "  .stack-item.unverified { border-color: #9aa0a6; background: #f2f3f5; }",
        "",
        "  .stack-item:hover { overflow: visible; z-index: 600; }",
        "",
        "  .stack-item .body { margin-top: 6px; }",
        "  .hdr { display: flex; flex-wrap: wrap; gap: 10px; align-items: baseline; }",
        "  .ent { font-weight: 800; font-size: 13px; padding: 2px 8px; border-radius: 999px; background: rgba(255,255,255,0.7);",
        "    border: 2px solid rgba(47,62,70,0.15); }",
        "  .dt { font-weight: 800; font-size: 13px; }",
        "  .tm { font-weight: 700; font-size: 13px; color: rgba(47,62,70,0.85); }",
        "  .body { margin-top: 6px; font-size: 13px; line-height: 1.25; }",
        "",
        "  .tooltip { display: none; position: absolute; left: 0; top: 100%; margin-top: 10px;",
        "    max-width: 760px; padding: 10px 12px; border: 2px solid var(--ink); border-radius: 14px;",
        "    background: #fff; box-shadow: 0 10px 26px rgba(0,0,0,0.12); font-size: 13px; line-height: 1.3; }",
        "  .card:hover .tooltip, .range-wrap:hover .tooltip { display: block; }",
        "",
        "  .timeline {",
        "    position: relative;",
        "    min-height: 400px;",
        "    /* reserve real columns: wrapper + gap + event area */",
        "    padding-left: 0;",
        "  }",
        "",
        "  /* Nice divider line */",
        "  .spine {",
        "    position: absolute;",
        "    left: calc(var(--wrap-area) + (var(--gap-x) / 2));",
        "    top: 0;",
        "    bottom: 0;",
        "    width: 2px;",
        "    background: rgba(47,62,70,0.10);",
        "  }",
        "  /* Small helper */",
        "  .empty { color: var(--muted); font-size: 13px; padding: 40px 10px; }",
        "</style>",
        "</head>",
        "<body>",
        "<div class='app'>",
        "  <aside class='sidebar'>",
        "    <div class='sidebar-inner'>",
        "      <h1 class='title'>Entities</h1>",
        "      <p class='sub'>Tick/untick entities to filter. The timeline reflows instantly.</p>",
        "      <div class='controls'>",
        "        <button class='btn' id='btn-all' type='button'>Select all</button>",
        "        <button class='btn' id='btn-none' type='button'>Select none</button>",
        "      </div>",
        "      <div class='entity-list' id='entity-list'></div>",
        "    </div>",
        "  </aside>",
        "  <main class='main'>",
        "    <div class='legend-top'>",
        "      <div class='chip'><span class='chip-box'></span> Zeker</div>",
        "      <div class='chip'><span class='chip-box uncertain'></span> Onzeker</div>",
        "      <div class='chip'><span class='chip-box unverified'></span> Ongeverifieerd</div>",
        "    </div>",
        "    <div class='viewport'>",
        "      <div class='timeline' id='timeline'>",
        "        <div class='spine'></div>",
        "      </div>",
        "    </div>",
        "  </main>",
        "</div>",
        f"<script id='data' type='application/json'>{payload_json}</script>",
        "<script>",
        "(function(){",
        "  const data = JSON.parse(document.getElementById('data').textContent);",
        "  const timelineEl = document.getElementById('timeline');",
        "  const listEl = document.getElementById('entity-list');",
        "  const btnAll = document.getElementById('btn-all');",
        "  const btnNone = document.getElementById('btn-none');",
        "",
        "  const cfg = {",
        "    rowH: cssNum('--row-h', 98),",
        "    rowGap: cssNum('--row-gap', 12),",
        "    gapX: cssNum('--gap-x', 20),",
        "    wrapColW: cssNum('--wrap-col-w', 240),",
        "    subcolW: cssNum('--subcol-w', 340),",
        "    maxSubcols: Math.max(1, Math.floor(cssNum('--max-subcols', 4))),",
        "    indent: cssNum('--indent', 16),",
        "  };",
        "",
        "  function cssNum(varName, fallback){",
        "    const v = getComputedStyle(document.documentElement).getPropertyValue(varName).trim();",
        "    if (!v) return fallback;",
        "    const n = parseFloat(v.replace('px',''));",
        "    return Number.isFinite(n) ? n : fallback;",
        "  }",
        "",
        "  // State: all selected by default",
        "  const selected = new Set(data.entities.map(e => e.name));",
        "",
        "  // Sidebar UI",
        "  function renderSidebar(counts){",
        "    listEl.innerHTML = '';",
        "    for (const ent of data.entities) {",
        "      const id = 'chk_' + ent.name.replace(/[^a-z0-9]+/gi,'_');",
        "      const item = document.createElement('label');",
        "      item.className = 'entity-item';",
        "      item.setAttribute('for', id);",
        "",
        "      const chk = document.createElement('input');",
        "      chk.type = 'checkbox';",
        "      chk.id = id;",
        "      chk.checked = selected.has(ent.name);",
        "      chk.addEventListener('change', () => {",
        "        if (chk.checked) selected.add(ent.name); else selected.delete(ent.name);",
        "        render();",
        "      });",
        "",
        "      const sw = document.createElement('span');",
        "      sw.className = 'swatch';",
        "      sw.style.background = ent.color;",
        "",
        "      const nm = document.createElement('span');",
        "      nm.className = 'entity-name';",
        "      nm.textContent = ent.name;",
        "",
        "      const ct = document.createElement('span');",
        "      ct.className = 'count';",
        "      ct.textContent = String(counts.get(ent.name) || 0);",
        "",
        "      item.appendChild(chk);",
        "      item.appendChild(sw);",
        "      item.appendChild(nm);",
        "      item.appendChild(ct);",
        "      listEl.appendChild(item);",
        "    }",
        "  }",
        "",
        "  btnAll.addEventListener('click', () => {",
        "    selected.clear();",
        "    for (const ent of data.entities) selected.add(ent.name);",
        "    render();",
        "  });",
        "  btnNone.addEventListener('click', () => {",
        "    selected.clear();",
        "    render();",
        "  });",
        "",
        "  function escapeHtml(s){",
        "    return String(s).replace(/[&<>\"]+/g, (m) => ({'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;'}[m] || m));",
        "  }",
        "",
        "  function classFlags(ev){",
        "    let cls = '';",
        "    if (!ev.certain) cls += ' uncertain';",
        "    if (!ev.verified) cls += ' unverified';",
        "    return cls;",
        "  }",
        "",
        "  // Core layout: slot-based vertical packing (not linear time)",
        "  function layout(events){",
        "    // events already filtered; keep only those with start_ms",
        "    const evs = events.slice().sort((a,b) => {",
        "      if (a.start_ms !== b.start_ms) return a.start_ms - b.start_ms;",
        "      // range first for same start",
        "      if (a.is_range !== b.is_range) return a.is_range ? -1 : 1;",
        "      // stable",
        "      return a.id.localeCompare(b.id);",
        "    });",
        "",
        "    // slot by start time (millis) in sorted order (packed, not linear)",
        "    const slotOfId = new Map();",
        "    const uniqueStarts = [];",
        "    const byStart = new Map();",
        "    for (const e of evs) {",
        "      if (!byStart.has(e.start_ms)) { byStart.set(e.start_ms, []); uniqueStarts.push(e.start_ms); }",
        "      byStart.get(e.start_ms).push(e);",
        "    }",
        "    uniqueStarts.sort((a,b)=>a-b);",
        "    uniqueStarts.forEach((s, idx) => slotOfId.set(String(s), idx));",
        "",
        "    // slot index for each event (by start_ms)",
        "    for (const e of evs) e._slot = slotOfId.get(String(e.start_ms));",
        "",
        "    // y positions per slot",
        "",
        "    // For each range event: find max slot among events whose start is within [start,end]",
        "    const ranges = evs.filter(e => e.is_range);",
        "    for (const r of ranges) {",
        "      let maxSlot = r._slot;",
        "      for (const e of evs) {",
        "        if (e.start_ms >= r.start_ms && e.start_ms <= r.end_ms) {",
        "          if (e._slot > maxSlot) maxSlot = e._slot;",
        "        }",
        "      }",
        "      r._maxSlot = maxSlot;",
        "    }",
        "",
        "    // Assign wrapper columns to overlapping ranges (interval coloring, greedy)",
        "    const rangesSorted = ranges.slice().sort((a,b)=>{",
        "      if (a.start_ms !== b.start_ms) return a.start_ms - b.start_ms;",
        "      return a.end_ms - b.end_ms;",
        "    });",
        "    const colEnd = []; // end_ms per wrapper column",
        "    for (const r of rangesSorted) {",
        "      let col = -1;",
        "      for (let i = 0; i < colEnd.length; i++) {",
        "        // reuse column only if it ended strictly before this starts (inclusive overlap)",
        "        if (r.start_ms > colEnd[i]) { col = i; break; }",
        "      }",
        "      if (col === -1) {",
        "        col = colEnd.length;",
        "        colEnd.push(r.end_ms);",
        "      } else {",
        "        colEnd[col] = Math.max(colEnd[col], r.end_ms);",
        "      }",
        "      r._wrapCol = col;",
        "    }",
        "    const wrapCols = Math.max(1, colEnd.length);",
        "    const wrapArea = wrapCols * cfg.wrapColW + (wrapCols - 1) * cfg.gapX;",
        "",
        "    // Group NORMAL (non-range) events by start_ms => one stacked card per start time",
        "    const normalByStart = new Map();",
        "    for (const e of evs) {",
        "      if (e.is_range) continue;",
        "      if (!normalByStart.has(e.start_ms)) normalByStart.set(e.start_ms, []);",
        "      normalByStart.get(e.start_ms).push(e);",
        "    }",
        "    for (const [k, arr] of normalByStart.entries()) {",
        "      // stable order inside stack",
        "      arr.sort((a,b)=>a.id.localeCompare(b.id));",
        "    }",
        "",
        "    // Dynamic per-slot heights (because stacks can be taller than rowH)",
        "    const headerH = 34;      // approx stack header height",
        "    const itemH = 76;        // fixed subcard height target",
        "    const itemGap = 10;      // matches CSS gap",
        "    const padH = 20;         // stack card padding top+bottom approx",
        "",
        "    const slotStarts = uniqueStarts.slice();",
        "    const slotH = new Array(slotStarts.length).fill(cfg.rowH);",
        "",
        "    for (let i = 0; i < slotStarts.length; i++) {",
        "      const start = slotStarts[i];",
        "      const n = (normalByStart.get(start) || []).length;",
        "      if (n > 0) {",
        "        const stackH = headerH + padH + (n * itemH) + ((n - 1) * itemGap);",
        "        slotH[i] = Math.max(slotH[i], stackH);",
        "      }",
        "    }",
        "",
        "    // Build yStart/yEnd per slot (cumulative)",
        "    const yStart = new Array(slotStarts.length).fill(0);",
        "    const yEnd = new Array(slotStarts.length).fill(0);",
        "    let yCursor = 0;",
        "    for (let i = 0; i < slotStarts.length; i++) {",
        "      yStart[i] = yCursor;",
        "      yEnd[i] = yCursor + slotH[i];",
        "      yCursor = yEnd[i] + cfg.rowGap;",
        "    }",
        "",
        "    const yOfSlot = (slot) => yStart[slot] || 0;",
        "    // Subcolumns for same-slot events (excluding range wrappers since they live in wrapper column)",
        "    for (const start of uniqueStarts) {",
        "      const group = byStart.get(start) || [];",
        "      const normals = group.filter(e => !e.is_range);",
        "      normals.sort((a,b)=>a.id.localeCompare(b.id));",
        "      // if we overflow subrows, push them down within that same start time block",
        "      // (we do it by extra y offset)",
        "    }",
        "",
        "    // Compute final geometry",
        "    const positioned = [];",
        "",
        "    // Range wrappers: fixed left column, tall",
        "    for (const r of ranges) {",
        "      const y0 = yStart[r._slot];",
        "      const y1 = yEnd[r._maxSlot];",
        "      const y = y0;",
        "      const h = Math.max(cfg.rowH, (y1 - y0));",
        "      const x = (r._wrapCol || 0) * (cfg.wrapColW + cfg.gapX);",
        "      const w = cfg.wrapColW;",        
        "      positioned.push({ kind:'range', ev:r, x, y, w, h });",
        "    }",
        "",
        "    // Normal cards: right side, in subcolumns",
        "    // Stacked cards (one per start_ms), on the right side",
        "    for (const start of slotStarts) {",
        "      const arr = normalByStart.get(start) || [];",
        "      if (!arr.length) continue;",
        "      const slot = slotOfId.get(String(start));",
        "      const y = yStart[slot];",
        "      const x = wrapArea + cfg.gapX;",
        "      const w = (cfg.subcolW * cfg.maxSubcols) + (cfg.gapX * (cfg.maxSubcols - 1));",
        "      const h = slotH[slot];",
        "      positioned.push({ kind:'stack', start_ms: start, items: arr, x, y, w, h });",
        "    }",
        "",
        "    // Total height",
        "    let maxBottom = 0;",
        "    for (const p of positioned) {",
        "      maxBottom = Math.max(maxBottom, p.y + p.h);",
        "    }",
        "    const totalH = (yEnd.length ? (yEnd[yEnd.length - 1] + 40) : 400);",
        "    return { positioned, totalH, wrapArea };",
        "  }",
        "",
        "  function render(){",
        "    // Filter",
        "    const filtered = data.events.filter(e =>",
        "      e.entities.some(ent => selected.has(ent))",
        "    );",
        "",
        "    // Counts (for sidebar display)",
        "    const counts = new Map();",
        "    for (const e of data.events) {",
        "      for (const ent of e.entities) {",
        "        counts.set(ent, (counts.get(ent) || 0) + 1);",
        "      }",
        "    }",
        "    renderSidebar(counts);",
        "",
        "    // Clear timeline (keep spine)",
        "    const spine = timelineEl.querySelector('.spine');",
        "    timelineEl.innerHTML = '';",
        "    if (spine) timelineEl.appendChild(spine);",
        "",
        "    if (!filtered.length) {",
        "      const empty = document.createElement('div');",
        "      empty.className = 'empty';",
        "      empty.textContent = 'No events selected.';",
        "      timelineEl.appendChild(empty);",
        "      timelineEl.style.height = 'auto';",
        "      return;",
        "    }",
        "",
        "    const { positioned, totalH, wrapArea } = layout(filtered);",
        "    timelineEl.style.height = totalH + 'px';",
        "    timelineEl.style.setProperty('--wrap-area', wrapArea + 'px');",
        "",
        "    for (const p of positioned) {",
        "      if (p.kind === 'range') {",
        "        const e = p.ev;",
        "        const firstVisible = e.entities.find(ent => selected.has(ent));",
        "        const color = firstVisible ? e.colors[firstVisible] : '#2f3e46';",
        "        const cls = classFlags(e);",
        "",
        "        const div = document.createElement('div');",
        "        div.className = 'range-wrap' + cls;",
        "        div.style.top = p.y + 'px';",
        "        div.style.left = p.x + 'px';",
        "        div.style.width = p.w + 'px';",
        "        div.style.height = p.h + 'px';",
        "        div.style.borderLeftColor = color;",
        "",
        "        const ent = e.entities",
        "          .filter(entName => selected.has(entName))",
        "          .map(entName => {",
        "            const c = e.colors[entName];",
        "            return `<span class='ent' style='border-color: ${c};'>${escapeHtml(entName)}</span>`;",
        "          })",
        "          .join('');",
        "        const hdr = `<div class='hdr'>${ent}<span class='dt'>${escapeHtml(e.date_label)}</span><span class='tm'>• ${escapeHtml(e.time_label)}</span></div>`;",
        "        const body = `<div class='body'>${escapeHtml(e.desc_short)}</div>`;",
        "        const tip = `<div class='tooltip'>${escapeHtml(e.desc_full)}</div>`;",
        "        div.innerHTML = hdr + body + tip;",
        "        timelineEl.appendChild(div);",
        "        continue;",
        "      }",
        "",
        "      if (p.kind === 'stack') {",
        "        const div = document.createElement('div');",
        "        div.className = 'stack-card';",
        "        div.style.top = p.y + 'px';",
        "        div.style.left = p.x + 'px';",
        "        div.style.width = p.w + 'px';",
        "        div.style.height = p.h + 'px';",
        "",
        "        const first = p.items[0];",
        "        const hdr = `<div class='stack-hdr'><span class='dt'>${escapeHtml(first.date_label)}</span><span class='tm'>• ${escapeHtml(first.time_label)}</span></div>`;",
        "",
        "        const itemsHtml = p.items.map(ev => {",
        "          const cls2 = classFlags(ev);",
        "          const firstVisible2 = ev.entities.find(ent => selected.has(ent));",
        "          const bar = firstVisible2 ? ev.colors[firstVisible2] : '#2f3e46';",
        "          const ents = ev.entities",
        "            .filter(entName => selected.has(entName))",
        "            .map(entName => {",
        "              const c = ev.colors[entName];",
        "              return `<span class='ent' style='border-color: ${c};'>${escapeHtml(entName)}</span>`;",
        "            })",
        "            .join('');",
        "          const tip = `<div class='tooltip'>${escapeHtml(ev.desc_full)}</div>`;",
        "          return (",
        "            `<div class='stack-item${cls2}' style='border-left-color: ${bar};'>` +",
        "              `<div class='hdr'>${ents}</div>` +",
        "              `<div class='body'>${escapeHtml(ev.desc_short)}</div>` +",
        "              tip +",
        "            `</div>`",
        "          );",
        "        }).join('');",
        "",
        "        div.innerHTML = hdr + `<div class='stack-items'>${itemsHtml}</div>`;",
        "        timelineEl.appendChild(div);",
        "        continue;",
        "      }",
        "    }",
        "  }",
        "",
        "  // Initial",
        "  render();",
        "})();",
        "</script>",
        "</body>",
        "</html>",
    ]

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(html_parts), encoding="utf-8")
    print(f"Saved: {output_path.resolve()}")


#%%
excel_path = "20260115 Tijdlijn.xlsx"
output_path = "timeline_vertical_filterable.html"
generate_vertical_timeline(excel_path, output_path)
#%%
