#%%
from __future__ import annotations
from datetime import datetime
from pathlib import Path
import html
from timeline_core import (
    TimelineEvent,
    build_entity_colors,
    estimate_card_width,
    group_events_by_entity,
    is_range,
    overlaps,
    read_events_from_excel,
    render_sources_html,
    sorted_events,
    truncate,
)

# -------------------------
# Main
# -------------------------
def generate_horizontal_timeline(excel_path: str | Path, output_path: str | Path) -> None:
    events = read_events_from_excel(excel_path)
    entity_events = group_events_by_entity(events)
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
        slot_width.append(estimate_card_width(txt, is_range(e)))

    # Compute packed x start for each slot
    x_start: list[int] = [0] * len(event_at_slot)
    x = 0
    for i in range(len(event_at_slot)):
        x_start[i] = x
        x += slot_width[i] + gap

    # For each range event, find the furthest slot whose event starts within the range
    range_max_slot: dict[int, int] = {}
    for r in (e for e in event_at_slot if is_range(e)):
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
        if is_range(e):
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

        for event in sorted_events(entity_list):
            if event.start_dt is None:
                # If you want to include "unknown time" events, you could handle them separately.
                continue

            is_range_event = is_range(event)
            card_text = f"{event.date_label} {event.time_label} {event.description}"
            width = event_w.get(event.event_id, estimate_card_width(card_text, is_range_event))
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

                    if overlaps(this_start, this_end, last_start, last_end):
                        continue

                subrow.append(
                    {
                        "event": event,
                        "x": x,
                        "width": width,
                        "is_range": is_range_event,
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
                            "is_range": is_range_event,
                            "end_dt": event.end_dt or event.start_dt,
                        }
                    ]
                )

        # Keep cards in each subrow ordered by x for nicer rendering
        for subrow in subrows:
            subrow.sort(key=lambda item: item["x"])

        layout[entity] = subrows

    # Build deterministic, dataset-local entity colours with strong hue separation.
    all_entities = sorted(entity_events.keys(), key=lambda s: s.lower())
    entity_colors = build_entity_colors(all_entities)

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
        "  padding: 8px 10px 34px 10px;",
        "  border: 2px solid #2f3e46;",
        "  border-left: 8px solid var(--entity-color, #2f3e46);",
        "  border-radius: 10px;",
        "  background: #f9fbfb;",
        "  box-sizing: border-box;",
        "  overflow: hidden;",
        "  cursor: pointer;",
        "}",

        ".card:hover {",
        "  overflow: visible;",
        "  z-index: 200;",
        "}",

        ".card.expanded {",
        "  height: auto !important;",
        "  min-height: 90px;",
        "  padding-bottom: 10px;",
        "  overflow: visible;",
        "  z-index: 1500;",
        "  box-shadow: 0 10px 26px rgba(0,0,0,0.16);",
        "}",

        ".card.expanded .card-body {",
        "  max-height: none;",
        "  overflow: visible;",
        "  padding-right: 0;",
        "}",

        ".card.expanded .card-body::after {",
        "  display: none;",
        "}",

        ".card.expanded .card-sources {",
        "  position: static;",
        "  margin-top: 8px;",
        "}",

        ".card-header {",
        "  font-weight: 600;",
        "  margin-bottom: 4px;",
        "  font-size: 12px;",
        "}",

        ".card-body {",
        "  line-height: 1.2;",
        "  margin-bottom: 0;",
        "  max-height: 42px;",
        "  overflow: hidden;",
        "  position: relative;",
        "  padding-right: 34px;",
        "}",

        ".card-body.has-more::after, .card-body.truncated::after {",
        "  content: '?';",
        "  position: absolute;",
        "  right: 0;",
        "  bottom: 24px;",
        "  width: 16px;",
        "  text-align: right;",
        "  font-size: 14px;",
        "  font-weight: 700;",
        "  color: rgba(47,62,70,0.78);",
        "  background: linear-gradient(90deg, rgba(249,251,251,0), rgba(249,251,251,0.97) 48%);",
        "}",

        ".card-sources {",
        "  position: absolute;",
        "  right: 8px;",
        "  bottom: 6px;",
        "  display: flex;",
        "  gap: 6px;",
        "  align-items: center;",
        "}",

        ".source-link, .source-copy {",
        "  width: 22px;",
        "  height: 22px;",
        "  border: 1px solid rgba(47,62,70,0.35);",
        "  border-radius: 6px;",
        "  background: #fff;",
        "  display: inline-flex;",
        "  align-items: center;",
        "  justify-content: center;",
        "  cursor: pointer;",
        "  padding: 0;",
        "  text-decoration: none;",
        "}",

        ".source-link:hover, .source-copy:hover {",
        "  background: rgba(47,62,70,0.08);",
        "}",

        ".source-icon {",
        "  width: 12px;",
        "  height: 12px;",
        "  position: relative;",
        "  display: inline-block;",
        "}",

        ".source-icon-link::before {",
        "  content: '';",
        "  position: absolute;",
        "  width: 8px;",
        "  height: 8px;",
        "  right: 1px;",
        "  top: 1px;",
        "  border-top: 2px solid #2f3e46;",
        "  border-right: 2px solid #2f3e46;",
        "}",

        ".source-icon-link::after {",
        "  content: '';",
        "  position: absolute;",
        "  width: 7px;",
        "  height: 2px;",
        "  left: 1px;",
        "  bottom: 2px;",
        "  background: #2f3e46;",
        "  transform: rotate(-45deg);",
        "  transform-origin: left center;",
        "}",

        ".source-icon-file::before {",
        "  content: '';",
        "  position: absolute;",
        "  left: 1px;",
        "  top: 1px;",
        "  width: 8px;",
        "  height: 10px;",
        "  border: 2px solid #2f3e46;",
        "  border-radius: 2px;",
        "}",

        ".source-icon-file::after {",
        "  content: '';",
        "  position: absolute;",
        "  right: 1px;",
        "  top: 2px;",
        "  width: 4px;",
        "  height: 4px;",
        "  border-top: 2px solid #2f3e46;",
        "  border-right: 2px solid #2f3e46;",
        "}",

        ".source-copy.copied {",
        "  background: #dfeff2;",
        "}",

        ".tooltip {",
        "  display: none;",
        "  position: fixed;",
        "  left: 0;",
        "  top: 0;",
        "  max-width: 520px;",
        "  max-height: min(60vh, 420px);",
        "  overflow: auto;",
        "  padding: 10px 12px;",
        "  border: 2px solid #2f3e46;",
        "  border-radius: 10px;",
        "  background: #ffffff;",
        "  box-shadow: 0 6px 18px rgba(0,0,0,0.12);",
        "  font-size: 13px;",
        "  line-height: 1.3;",
        "  pointer-events: none;",
        "  z-index: 2000;",
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
        "  border-style: dotted;",
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
        color = entity_colors[entity]
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
                desc_full_attr = html.escape(event.description, quote=True)
                short_plain = truncate(event.description, 120)
                desc_short = html.escape(short_plain)
                desc_short_attr = html.escape(short_plain, quote=True)
                has_more = event.description != short_plain
                sources_html = render_sources_html(event.sources)

                html_parts.append(
                    (
                        f"<div class='{ ' '.join(classes) }' "
                        f"data-event='1' "
                        f"style='left: {card['x']}px; width: {card['width']}px;'>"
                        f"<div class='card-header'>{event.date_label} &middot; {event.time_label}</div>"
                        f"<div class='card-body{' has-more' if has_more else ''}' "
                        f"data-short-text='{desc_short_attr}' data-full-text='{desc_full_attr}'>{desc_short}</div>"
                        f"{sources_html}"
                        f"<div class='tooltip'>{desc_full}</div>"
                        "</div>"
                    )
                )
            html_parts.append("</div>")
        html_parts.append("</div>")
    html_parts.extend([
                "<script>",
                "(function(){",
                "  document.addEventListener('click', async (e) => {",
                "    const btn = e.target.closest('.source-copy[data-copy-source]');",
                "    if (!btn) return;",
                "    const value = btn.getAttribute('data-copy-source') || '';",
                "    try {",
                "      await navigator.clipboard.writeText(value);",
                "      btn.classList.add('copied');",
                "      window.setTimeout(() => btn.classList.remove('copied'), 500);",
                "    } catch (_) {}",
                "  });",
                "  function updateBodyOverflowCues(){",
                "    const bodies = document.querySelectorAll('.card-body');",
                "    for (const body of bodies) {",
                "      const clipped = (body.scrollHeight - body.clientHeight) > 1;",
                "      body.classList.toggle('truncated', clipped);",
                "    }",
                "  }",
                "  function initCardExpansion(){",
                "    const cards = document.querySelectorAll('.card[data-event=\"1\"]');",
                "    let activeCard = null;",
                "    function setBodyText(card, useFull){",
                "      const body = card.querySelector('.card-body');",
                "      if (!body) return;",
                "      body.textContent = useFull ? (body.dataset.fullText || body.textContent || '') : (body.dataset.shortText || body.textContent || '');",
                "    }",
                "    function collapseActive(){",
                "      if (!activeCard) return;",
                "      setBodyText(activeCard, false);",
                "      activeCard.classList.remove('expanded');",
                "      activeCard = null;",
                "      updateBodyOverflowCues();",
                "      requestRedraw();",
                "    }",
                "    for (const card of cards) {",
                "      card.addEventListener('click', (e) => {",
                "        if (e.target.closest('.source-link, .source-copy')) return;",
                "        e.stopPropagation();",
                "        if (activeCard === card) { collapseActive(); return; }",
                "        collapseActive();",
                "        setBodyText(card, true);",
                "        card.classList.add('expanded');",
                "        activeCard = card;",
                "        updateBodyOverflowCues();",
                "        requestRedraw();",
                "      });",
                "    }",
                "    document.addEventListener('click', (e) => {",
                "      if (!activeCard) return;",
                "      if (e.target.closest('.card[data-event=\"1\"]')) return;",
                "      collapseActive();",
                "    });",
                "  }",
                "  const scroller = document.querySelector('.timeline-scroller');",
                "  const timeline = document.querySelector('.timeline');",
                "  const canvas = document.getElementById('minimap-canvas');",
                "  const box = document.getElementById('minimap');",
                "  if (!scroller || !timeline || !canvas || !box) return;",
                "  initCardExpansion();",
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
                "  window.addEventListener('resize', () => { updateBodyOverflowCues(); requestRedraw(); });",
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
                "  window.setTimeout(() => { updateBodyOverflowCues(); requestRedraw(); }, 50);",
                "})();",
                "</script>",
            ])

    html_parts.extend(["</div>", "</div>", "</body>", "</html>"])

    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(html_parts), encoding="utf-8")
    print(f"Timeline saved to {output_path.resolve()}")


# Backward-compatible name
generate_timeline = generate_horizontal_timeline


if __name__ == "__main__":
    excel_path = "20260115 Tijdlijn.xlsx"
    output_path = "timeline_v3_global_packed.html"
    generate_horizontal_timeline(excel_path, output_path)

