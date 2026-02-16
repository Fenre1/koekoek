#%%
from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, time as dt_time
from pathlib import Path
from typing import Iterable
import html
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
    if value is None or (isinstance(value, float) and pd.isna(value)) or pd.isna(value):
        return False
    return str(value).strip().lower() == "ja"


def _parse_date_time(date_value: object, time_value: object) -> datetime | None:
    # Date missing/invalid
    if date_value is None or pd.isna(date_value):
        return None

    date_ts = pd.to_datetime(date_value, errors="coerce")
    if pd.isna(date_ts):
        return None
    date_ts = date_ts.normalize()

    # Time missing => date-only
    if time_value is None or pd.isna(time_value):
        return date_ts.to_pydatetime()

    # Already a time object
    if isinstance(time_value, dt_time):
        return datetime.combine(date_ts.date(), time_value)

    # Already datetime / timestamp
    if isinstance(time_value, (pd.Timestamp, datetime)):
        return datetime.combine(date_ts.date(), pd.to_datetime(time_value).time())

    # Excel fractional day (float)
    if isinstance(time_value, (int, float)) and not pd.isna(time_value):
        t = pd.to_datetime(time_value, unit="D", origin="1899-12-30", errors="coerce")
        if pd.isna(t):
            return date_ts.to_pydatetime()
        return datetime.combine(date_ts.date(), t.time())

    # Parse string
    t = pd.to_datetime(str(time_value).strip(), errors="coerce")
    if pd.isna(t):
        return date_ts.to_pydatetime()
    return datetime.combine(date_ts.date(), t.time())

def _effective_end_dt(e: TimelineEvent) -> datetime:
    # Treat instant events as having a tiny duration so they still “occupy” time.
    if e.start_dt is None:
        return datetime.max
    if e.end_dt is None:
        return e.start_dt
    return e.end_dt

def _overlaps(a_start: datetime, a_end: datetime, b_start: datetime, b_end: datetime) -> bool:
    # Inclusive overlap: sharing a boundary counts as overlap.
    return a_start <= b_end and b_start <= a_end

def _truncate(s: str, n: int = 120) -> str:
    s = s.strip()
    return s if len(s) <= n else s[: n - 1] + "…"

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


def _entity_color(entity: str) -> str:
    """
    Deterministic pastel-ish HSL colour from entity name.
    Looks distinct for dozens of entities and stays stable across runs.
    """
    h = 0
    for ch in entity:
        h = (h * 31 + ord(ch)) % 360
    # Pastel palette: fixed saturation/lightness, varying hue
    return f"hsl({h}, 55%, 72%)"

def _estimate_card_width(text: str, is_range: bool) -> int:
    base_width = 140 + int(len(text) * 5)
    min_width = 180
    max_width = 320
    width = max(min_width, min(base_width, max_width))
    # range cards will be expanded later, but keep a sensible minimum
    return width if not is_range else max(width, min_width)


def _font_size_for_text(text: str) -> int:
    length = max(len(text), 1)    
    return max(6, min(14, int(220 / (length ** 0.5))))


def _split_entities(value: object) -> list[str]:
    if value is None or pd.isna(value):
        return ["Onbekend"]
    entities = [item.strip() for item in str(value).split("|") if item.strip()]
    return entities or ["Onbekend"]


def _sorted_events(events: Iterable[TimelineEvent]) -> list[TimelineEvent]:
    def sort_key(event: TimelineEvent) -> tuple:
        # range first => False < True
        is_instant = not _is_range(event)
        return (event.start_dt is None, event.start_dt or datetime.max, is_instant, event.event_id)

    return sorted(events, key=sort_key)


def _is_range(e: TimelineEvent) -> bool:
    return (
        e.start_dt is not None
        and e.end_dt is not None
        and e.end_dt != e.start_dt
    )


# -------------------------
# Main
# -------------------------
def generate_timeline(excel_path: str | Path, output_path: str | Path) -> None:
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

    # 1) Read events
    events: list[TimelineEvent] = []
    for idx, row in df.iterrows():
        # Your current behaviour: keep date-only events (00:00). If you want to skip
        # "no Starttijd and no Eindtijd", uncomment next lines.
        #
        # if pd.isna(row["Starttijd"]) and pd.isna(row["Eindtijd"]):
        #     continue

        start_dt = _parse_date_time(row["Datum"], row["Starttijd"])
        end_dt = _parse_date_time(row["Datum"], row["Eindtijd"]) if not pd.isna(row["Eindtijd"]) else None

        description = str(row["Gebeurtenis"]) if not pd.isna(row["Gebeurtenis"]) else ""

        events.append(
            TimelineEvent(
                event_id=int(idx),
                description=description,
                entities=_split_entities(row["Entiteit(en) (splits op met |)"]),
                date_label=_format_date_label(row["Datum"]),
                time_label=_format_time_label(start_dt, end_dt),
                start_dt=start_dt,
                end_dt=end_dt if end_dt and start_dt else None,
                certain=_normalize_bool(row["Zekerheid (ja/nee)"]),
                verified=_normalize_bool(row["Geverifieerd"]),
            )
        )

    # Map entity -> events
    entity_events: dict[str, list[TimelineEvent]] = {}
    for event in events:
        for entity in event.entities:
            entity_events.setdefault(entity, []).append(event)

    # 2) Build GLOBAL packed x-axis over all entities (slot-based, not linear time)
    gap = 24

    timed_events = [e for e in events if e.start_dt is not None]
    timed_events.sort(key=lambda e: (e.start_dt, e.end_dt is None, e.event_id))

    if not timed_events:
        raise ValueError("No events with a valid start time were found.")

    # Assign slots globally
    slot_of: dict[int, int] = {}
    event_at_slot: list[TimelineEvent] = []
    for s, e in enumerate(timed_events):
        slot_of[e.event_id] = s
        event_at_slot.append(e)

    # Slot widths (initial), based on card text
    slot_width: list[int] = []
    for e in event_at_slot:
        txt = f"{e.date_label} {e.time_label} {e.description}"
        slot_width.append(_estimate_card_width(txt, _is_range(e)))

    # Compute packed x start for each slot
    x_start: list[int] = [0] * len(event_at_slot)
    x = 0
    for i in range(len(event_at_slot)):
        x_start[i] = x
        x += slot_width[i] + gap

    # For each range event, find the furthest slot whose event starts within the range
    range_max_slot: dict[int, int] = {}
    for r in (e for e in event_at_slot if _is_range(e)):
        r_slot = slot_of[r.event_id]
        max_slot = r_slot
        # Events are globally ordered by start_dt; we can scan and check containment
        for e in event_at_slot:
            if e.start_dt is None:
                continue
            if r.start_dt <= e.start_dt <= r.end_dt:  # type: ignore[operator]
                max_slot = max(max_slot, slot_of[e.event_id])
        range_max_slot[r.event_id] = max_slot

    # Final per-event x and width (range widths expanded to cover contained events)
    event_x: dict[int, int] = {}
    event_w: dict[int, int] = {}

    for e in event_at_slot:
        s = slot_of[e.event_id]
        event_x[e.event_id] = x_start[s]

        w = slot_width[s]
        if _is_range(e):
            last_slot = range_max_slot[e.event_id]
            right_edge = x_start[last_slot] + slot_width[last_slot]
            w = max(w, right_edge - x_start[s])
        event_w[e.event_id] = w

    # Timeline max width for container sizing
    max_width = (x_start[-1] + slot_width[-1]) if event_at_slot else 0

    # 3) Build per-entity layout (vertical stacking per entity only)
    layout: dict[str, list[list[dict]]] = {}

    for entity, entity_list in entity_events.items():
        subrows: list[list[dict]] = []

        for event in _sorted_events(entity_list):
            if event.start_dt is None:
                # If you want to include "unknown time" events, you could handle them separately.
                continue

            is_range = _is_range(event)
            card_text = f"{event.date_label} {event.time_label} {event.description}"
            width = event_w.get(event.event_id, _estimate_card_width(card_text, is_range))
            x = event_x.get(event.event_id, 0)

            placed = False
            for subrow in subrows:
                last = subrow[-1] if subrow else None
                # overlap check in time to decide vertical stacking inside this entity
                if last and event.start_dt and last["event"].start_dt:
                    last_start = last["event"].start_dt
                    last_end = last["end_dt"] or last_start
                    this_start = event.start_dt
                    this_end = event.end_dt or event.start_dt

                    if _overlaps(this_start, this_end, last_start, last_end):
                        continue

                subrow.append(
                    {
                        "event": event,
                        "x": x,
                        "width": width,
                        "is_range": is_range,
                        "end_dt": event.end_dt or event.start_dt,
                    }
                )
                placed = True
                break

            if not placed:
                subrows.append(
                    [
                        {
                            "event": event,
                            "x": x,
                            "width": width,
                            "is_range": is_range,
                            "end_dt": event.end_dt or event.start_dt,
                        }
                    ]
                )

        # Keep cards in each subrow ordered by x for nicer rendering
        for subrow in subrows:
            subrow.sort(key=lambda item: item["x"])

        layout[entity] = subrows

    # 4) Write HTML
    html_parts = [
        "<!DOCTYPE html>",
        "<html lang='nl'>",
        "<head>",
        "<meta charset='utf-8' />",
        "<title>Timeline</title>",
        "<style>",

        "body { font-family: 'Segoe UI', sans-serif; margin: 20px; }",

        "#minimap {",
        "  position: fixed;",
        "  right: 18px;",
        "  bottom: 18px;",
        "  width: 260px;",
        "  height: 180px;",
        "  background: rgba(255,255,255,0.92);",
        "  border: 2px solid #2f3e46;",
        "  border-radius: 12px;",
        "  box-shadow: 0 8px 24px rgba(0,0,0,0.12);",
        "  z-index: 999;",
        "  display: flex;",
        "  flex-direction: column;",
        "  overflow: hidden;",
        "}",
        "#minimap-title {",
        "  font-size: 12px;",
        "  font-weight: 600;",
        "  padding: 6px 10px;",
        "  border-bottom: 1px solid rgba(47,62,70,0.25);",
        "}",
        "#minimap-canvas {",
        "  width: 100%;",
        "  height: 100%;",
        "  display: block;",
        "  cursor: pointer;",
        "}",


        ":root { --label-w: 220px; --label-gap: 16px; }",

        ".timeline-scroller {",
        "  overflow-x: auto;",
        "  overflow-y: visible;",
        "  position: relative;",
        "  padding-bottom: 10px;",
        "}",

        ".timeline {",
        "  display: flex;",
        "  flex-direction: column;",
        "  gap: 5px;",
        "  min-width: fit-content;",
        "}",

        ".entity {",
        "  position: relative;",
        "  border-left: 6px solid var(--entity-color, #2f3e46);",
        "  padding-left: calc(var(--label-w) + var(--label-gap));",
        "}",

        ".entity-title {",
        "  position: sticky;",
        "  left: 0;",
        "  z-index: 100;",
        "  width: var(--label-w);",
        "  margin-left: calc(-1 * (var(--label-w) + var(--label-gap)));",
        "  font-weight: 600;",
        "  background: white;",
        "  padding: 4px 10px;",
        "  border-radius: 10px;",
        "  border: 2px solid var(--entity-color, #2f3e46);",
        "  white-space: nowrap;",
        "  margin-bottom: 12px;",
        "}",

        ".subrow {",
        "  position: relative;",
        "  height: 95px;",
        "  margin-bottom: 2px;",
        "}",

        ".card {",
        "  position: absolute;",
        "  height: 90px;",
        "  padding: 8px 10px;",
        "  border: 2px solid #2f3e46;",
        "  border-left: 8px solid var(--entity-color, #2f3e46);",
        "  border-radius: 10px;",
        "  background: #f9fbfb;",
        "  box-sizing: border-box;",
        "  overflow: hidden;",
        "}",

        ".card:hover {",
        "  overflow: visible;",
        "  z-index: 200;",
        "}",

        ".card-header {",
        "  font-weight: 600;",
        "  margin-bottom: 4px;",
        "  font-size: 12px;",
        "}",

        ".card-body {",
        "  line-height: 1.2;",
        "}",

        ".tooltip {",
        "  display: none;",
        "  position: absolute;",
        "  left: 0;",
        "  top: 100%;",
        "  margin-top: 8px;",
        "  max-width: 520px;",
        "  padding: 10px 12px;",
        "  border: 2px solid #2f3e46;",
        "  border-radius: 10px;",
        "  background: #ffffff;",
        "  box-shadow: 0 6px 18px rgba(0,0,0,0.12);",
        "  font-size: 13px;",
        "  line-height: 1.3;",
        "}",

        ".card:hover .tooltip {",
        "  display: block;",
        "}",

        ".legend {",
        "  margin-bottom: 20px;",
        "  display: flex;",
        "  gap: 20px;",
        "  font-size: 12px;",
        "}",

        ".legend-item {",
        "  display: flex;",
        "  align-items: center;",
        "  gap: 6px;",
        "}",

        ".legend-box {",
        "  width: 18px;",
        "  height: 12px;",
        "  border: 2px solid #2f3e46;",
        "  border-radius: 4px;",
        "}",

        ".legend-box.uncertain {",
        "  border-style: dashed;",
        "}",

        ".legend-box.unverified {",
        "  border-color: #9aa0a6;",
        "  background: #f2f3f5;",
        "}",

        "</style>",
        "</head>",




        "<body>",

        "<div class='legend'>",
        "  <div class='legend-item'><span class='legend-box'></span> Zeker</div>",
        "  <div class='legend-item'><span class='legend-box uncertain'></span> Onzeker</div>",
        "  <div class='legend-item'><span class='legend-box unverified'></span> Ongeverifieerd</div>",
        "</div>",

        "<div id='minimap'>",
        "  <div id='minimap-title'>Overview</div>",
        "  <canvas id='minimap-canvas'></canvas>",
        "</div>",


        "<div class='timeline-scroller'>",
        "  <div class='timeline'>",
    ]


    for entity, subrows in layout.items():
        color = _entity_color(entity)
        html_parts.append(
            f"<div class='entity' style='width: {max_width + 40}px; --entity-color: {color};'>"
        )
        html_parts.append(f"<div class='entity-title'>{entity}</div>")
        for subrow in subrows:
            html_parts.append("<div class='subrow'>")
            for card in subrow:
                event = card["event"]
                classes = ["card"]
                if card["is_range"]:
                    classes.append("range")
                if not event.certain:
                    classes.append("uncertain")
                if not event.verified:
                    classes.append("unverified")

                desc_full = html.escape(event.description)
                desc_short = html.escape(_truncate(event.description, 120))

                html_parts.append(
                    (
                        f"<div class='{ ' '.join(classes) }' "
                        f"data-event='1' "
                        f"style='left: {card['x']}px; width: {card['width']}px;'>"
                        f"<div class='card-header'>{event.date_label} • {event.time_label}</div>"
                        f"<div class='card-body'>{desc_short}</div>"
                        f"<div class='tooltip'>{desc_full}</div>"
                        "</div>"
                    )
                )
            html_parts.append("</div>")
        html_parts.append("</div>")
    html_parts.extend([
                "<script>",
                "(function(){",
                "  const scroller = document.querySelector('.timeline-scroller');",
                "  const timeline = document.querySelector('.timeline');",
                "  const canvas = document.getElementById('minimap-canvas');",
                "  const box = document.getElementById('minimap');",
                "  if (!scroller || !timeline || !canvas || !box) return;",
                "  const ctx = canvas.getContext('2d');",
                "  let dragging = false;",
                "  let needsRedraw = true;",
                "  function clamp(v, lo, hi) { return Math.max(lo, Math.min(hi, v)); }",
                "  function layoutCanvas(){",
                "    const dpr = window.devicePixelRatio || 1;",
                "    const rect = canvas.getBoundingClientRect();",
                "    const w = Math.max(1, Math.floor(rect.width * dpr));",
                "    const h = Math.max(1, Math.floor(rect.height * dpr));",
                "    if (canvas.width !== w || canvas.height !== h) {",
                "      canvas.width = w; canvas.height = h;",
                "      needsRedraw = true;",
                "    }",
                "  }",
                "  function getTimelineTop(){",
                "    const r = timeline.getBoundingClientRect();",
                "    return r.top + window.scrollY;",
                "  }",
                "  function draw(){",
                "    if (!needsRedraw) return;",
                "    needsRedraw = false;",
                "    layoutCanvas();",
                "    const dpr = window.devicePixelRatio || 1;",
                "    const W = canvas.width;",
                "    const H = canvas.height;",
                "    ctx.clearRect(0, 0, W, H);",
                "    const totalW = Math.max(1, scroller.scrollWidth);",
                "    const totalH = Math.max(1, timeline.scrollHeight);",
                "    const scrollerRect = scroller.getBoundingClientRect();",
                "    const timelineTop = getTimelineTop();",
                "    ctx.fillStyle = 'rgba(47,62,70,0.03)';",
                "    ctx.fillRect(0, 0, W, H);",
                "    const cards = timeline.querySelectorAll('.card[data-event=\"1\"]');",
                "    for (const card of cards) {",
                "      const r = card.getBoundingClientRect();",
                "      const xContent = (r.left - scrollerRect.left) + scroller.scrollLeft;",
                "      const yContent = (r.top + window.scrollY) - timelineTop;",
                "      const wContent = r.width;",
                "      const hContent = r.height;",
                "      const x = (xContent / totalW) * W;",
                "      const y = (yContent / totalH) * H;",
                "      const w = Math.max(1, (wContent / totalW) * W);",
                "      const h = Math.max(1, (hContent / totalH) * H);",
                "      const color = getComputedStyle(card).getPropertyValue('--entity-color').trim() || '#2f3e46';",
                "      ctx.fillStyle = color;",
                "      ctx.globalAlpha = 0.65;",
                "      ctx.fillRect(x, y, w, h);",
                "    }",
                "    ctx.globalAlpha = 1.0;",
                "    const viewX = (scroller.scrollLeft / totalW) * W;",
                "    const viewW = (scroller.clientWidth / totalW) * W;",
                "    const pageTopInTimeline = (window.scrollY - timelineTop);",
                "    const viewY = (pageTopInTimeline / totalH) * H;",
                "    const viewH = (window.innerHeight / totalH) * H;",
                "    const vx = clamp(viewX, 0, W);",
                "    const vy = clamp(viewY, 0, H);",
                "    const vw = clamp(viewW, 2, W);",
                "    const vh = clamp(viewH, 2, H);",
                "    ctx.strokeStyle = '#2f3e46';",
                "    ctx.lineWidth = Math.max(1, 2 * dpr);",
                "    ctx.strokeRect(vx, vy, vw, vh);",
                "    ctx.fillStyle = 'rgba(47,62,70,0.10)';",
                "    ctx.fillRect(vx, vy, vw, vh);",
                "  }",
                "  function requestRedraw(){",
                "    needsRedraw = true;",
                "    window.requestAnimationFrame(draw);",
                "  }",
                "  function goToCanvasPoint(clientX, clientY){",
                "    const rect = canvas.getBoundingClientRect();",
                "    const x = clamp(clientX - rect.left, 0, rect.width);",
                "    const y = clamp(clientY - rect.top, 0, rect.height);",
                "    const totalW = Math.max(1, scroller.scrollWidth);",
                "    const totalH = Math.max(1, timeline.scrollHeight);",
                "    const timelineTop = getTimelineTop();",
                "    const targetX = (x / rect.width) * totalW - scroller.clientWidth / 2;",
                "    scroller.scrollLeft = clamp(targetX, 0, totalW - scroller.clientWidth);",
                "    const targetYAbs = timelineTop + (y / rect.height) * totalH - window.innerHeight / 2;",
                "    const maxScrollY = document.documentElement.scrollHeight - window.innerHeight;",
                "    window.scrollTo({ top: clamp(targetYAbs, 0, maxScrollY), behavior: 'auto' });",
                "    requestRedraw();",
                "  }",
                "  scroller.addEventListener('scroll', requestRedraw, { passive: true });",
                "  window.addEventListener('scroll', requestRedraw, { passive: true });",
                "  window.addEventListener('resize', requestRedraw);",
                "  canvas.addEventListener('mousedown', (e) => { dragging = true; goToCanvasPoint(e.clientX, e.clientY); });",
                "  window.addEventListener('mousemove', (e) => { if (!dragging) return; goToCanvasPoint(e.clientX, e.clientY); });",
                "  window.addEventListener('mouseup', () => { dragging = false; });",
                "  canvas.addEventListener('touchstart', (e) => {",
                "    dragging = true;",
                "    const t = e.touches[0];",
                "    goToCanvasPoint(t.clientX, t.clientY);",
                "    e.preventDefault();",
                "  }, { passive: false });",
                "  window.addEventListener('touchmove', (e) => {",
                "    if (!dragging) return;",
                "    const t = e.touches[0];",
                "    goToCanvasPoint(t.clientX, t.clientY);",
                "    e.preventDefault();",
                "  }, { passive: false });",
                "  window.addEventListener('touchend', () => { dragging = false; });",
                "  window.setTimeout(requestRedraw, 50);",
                "})();",
                "</script>",
            ])

    html_parts.extend(["</div>", "</div>", "</body>", "</html>"])

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(html_parts), encoding="utf-8")
    print(f"Timeline saved to {output_path.resolve()}")


#%%
excel_path = "20260115 Tijdlijn.xlsx"
output_path = "timeline_v3_global_packed.html"
generate_timeline(excel_path, output_path)
#%%
