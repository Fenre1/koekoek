from __future__ import annotations

import csv
from dataclasses import dataclass
from datetime import datetime, timedelta, time as dt_time
from pathlib import Path
from typing import Iterable

import html
import pandas as pd


REQUIRED_COLUMNS = {
    "Datum",
    "Starttijd",
    "Eindtijd",
    "Zekerheid (ja/nee)",
    "Entiteit(en) (splits op met |)",
    "Gebeurtenis",
    "Geverifieerd",
}


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
    sources: list[str]


def normalize_bool(value: object) -> bool:
    if value is None or pd.isna(value):
        return False
    return str(value).strip().lower() == "ja"


def parse_date_time(date_value: object, time_value: object) -> datetime | None:
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


def format_date_label(date_value: object) -> str:
    if date_value is None or pd.isna(date_value):
        return "Onbekende datum"
    date = pd.to_datetime(date_value, errors="coerce")
    if pd.isna(date):
        return "Onbekende datum"
    return date.strftime("%Y-%m-%d")


def format_time_label(start_dt: datetime | None, end_dt: datetime | None) -> str:
    if start_dt is None:
        return "Onbekende tijd"
    start_time = start_dt.strftime("%H:%M")
    if end_dt is None or end_dt == start_dt:
        return start_time
    return f"{start_time} - {end_dt.strftime('%H:%M')}"


def normalize_end_dt(start_dt: datetime | None, end_dt: datetime | None) -> datetime | None:
    if start_dt is None or end_dt is None:
        return end_dt
    if end_dt < start_dt:
        return end_dt + timedelta(days=1)
    return end_dt


def split_entities(value: object) -> list[str]:
    if value is None or pd.isna(value):
        return ["Onbekend"]
    entities = [item.strip() for item in str(value).split("|") if item.strip()]
    return entities or ["Onbekend"]


def split_sources(value: object) -> list[str]:
    if value is None or pd.isna(value):
        return []
    raw = str(value).strip()
    if not raw:
        return []
    try:
        parts = next(csv.reader([raw], delimiter="|", quotechar='"', skipinitialspace=True))
    except Exception:
        parts = raw.split("|")
    cleaned: list[str] = []
    for part in parts:
        item = part.strip()
        if len(item) >= 2 and item.startswith('"') and item.endswith('"'):
            item = item[1:-1].strip()
        if item:
            cleaned.append(item)
    return cleaned


def is_web_link(source: str) -> bool:
    lower = source.strip().lower()
    return lower.startswith("http://") or lower.startswith("https://") or lower.startswith("www.")


def normalize_href(source: str) -> str:
    s = source.strip()
    if s.lower().startswith("www."):
        return f"https://{s}"
    return s


def render_sources_html(sources: list[str]) -> str:
    if not sources:
        return ""
    items: list[str] = []
    for src in sources:
        full = html.escape(src, quote=True)
        if is_web_link(src):
            href = html.escape(normalize_href(src), quote=True)
            items.append(
                "<a class='source-link' "
                f"href='{href}' target='_blank' rel='noopener noreferrer' title='{full}'>"
                "<span class='source-icon source-icon-link' aria-hidden='true'></span>"
                "</a>"
            )
            continue

        data_value = html.escape(src, quote=True)
        items.append(
            "<button class='source-copy' type='button' "
            f"data-copy-source='{data_value}' title='{full}'>"
            "<span class='source-icon source-icon-file' aria-hidden='true'></span>"
            "</button>"
        )
    return "<div class='card-sources'>" + "".join(items) + "</div>"


def is_range(event: TimelineEvent) -> bool:
    return event.start_dt is not None and event.end_dt is not None and event.end_dt != event.start_dt


def overlaps(a_start: datetime, a_end: datetime, b_start: datetime, b_end: datetime) -> bool:
    return a_start <= b_end and b_start <= a_end


def effective_end_dt(event: TimelineEvent) -> datetime:
    if event.start_dt is None:
        return datetime.max
    if event.end_dt is None:
        return event.start_dt
    return event.end_dt


def sorted_events(events: Iterable[TimelineEvent]) -> list[TimelineEvent]:
    def sort_key(event: TimelineEvent) -> tuple:
        is_instant = not is_range(event)
        return (event.start_dt is None, event.start_dt or datetime.max, is_instant, event.event_id)

    return sorted(events, key=sort_key)


def truncate(text: str, n: int = 120) -> str:
    text = text.strip()
    return text if len(text) <= n else text[: n - 1] + "..."


def estimate_card_width(text: str, is_range_event: bool) -> int:
    base_width = 140 + int(len(text) * 5)
    min_width = 180
    max_width = 320
    width = max(min_width, min(base_width, max_width))
    return width if not is_range_event else max(width, min_width)


def build_entity_colors(entities: list[str]) -> dict[str, str]:
    ordered: list[str] = []
    seen: set[str] = set()
    for entity in sorted(entities, key=lambda s: s.casefold()):
        if entity not in seen:
            seen.add(entity)
            ordered.append(entity)

    base_hue = 24.0
    golden_angle = 137.508
    colors: dict[str, str] = {}
    for index, entity in enumerate(ordered):
        hue = int((base_hue + (index * golden_angle)) % 360)
        colors[entity] = f"hsl({hue}, 60%, 68%)"
    return colors


def read_events_from_excel(excel_path: str | Path) -> list[TimelineEvent]:
    df = pd.read_excel(excel_path)
    missing = REQUIRED_COLUMNS.difference(df.columns)
    if missing:
        raise ValueError(f"Missing columns in Excel file: {', '.join(sorted(missing))}")

    events: list[TimelineEvent] = []
    for idx, row in df.iterrows():
        start_dt = parse_date_time(row["Datum"], row["Starttijd"])
        end_dt = parse_date_time(row["Datum"], row["Eindtijd"]) if not pd.isna(row["Eindtijd"]) else None
        end_dt = normalize_end_dt(start_dt, end_dt)
        description = str(row["Gebeurtenis"]) if not pd.isna(row["Gebeurtenis"]) else ""

        events.append(
            TimelineEvent(
                event_id=int(idx),
                description=description,
                entities=split_entities(row["Entiteit(en) (splits op met |)"]),
                date_label=format_date_label(row["Datum"]),
                time_label=format_time_label(start_dt, end_dt),
                start_dt=start_dt,
                end_dt=end_dt if (start_dt and end_dt) else None,
                certain=normalize_bool(row["Zekerheid (ja/nee)"]),
                verified=normalize_bool(row["Geverifieerd"]),
                sources=split_sources(row["Bron"]) if "Bron" in df.columns else [],
            )
        )
    return events


def group_events_by_entity(events: Iterable[TimelineEvent]) -> dict[str, list[TimelineEvent]]:
    entity_events: dict[str, list[TimelineEvent]] = {}
    for event in events:
        for entity in event.entities:
            entity_events.setdefault(entity, []).append(event)
    return entity_events
