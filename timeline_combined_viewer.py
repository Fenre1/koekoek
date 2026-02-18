from __future__ import annotations

from pathlib import Path
import json
import tempfile

from timeline_horizontal import generate_horizontal_timeline
from timeline_vertical_filterable import generate_vertical_timeline


def _build_combined_html(horizontal_html: str, vertical_html: str) -> str:
    payload = {
        "horizontal": horizontal_html,
        "vertical": vertical_html,
    }
    payload_json = json.dumps(payload, ensure_ascii=False).replace("</", "<\\/")

    return "\n".join(
        [
            "<!DOCTYPE html>",
            "<html lang='nl'>",
            "<head>",
            "<meta charset='utf-8' />",
            "<meta name='viewport' content='width=device-width, initial-scale=1' />",
            "<title>Timeline Viewer</title>",
            "<style>",
            "  :root { color-scheme: light; }",
            "  body { margin: 0; font-family: 'Segoe UI', sans-serif; color: #2f3e46; background: #f5f7f8; }",
            "  .app { min-height: 100vh; display: grid; grid-template-rows: auto 1fr; }",
            "  .toolbar { display: flex; flex-wrap: wrap; gap: 10px; align-items: center; padding: 14px 16px; border-bottom: 2px solid rgba(47,62,70,0.16); background: #fff; position: sticky; top: 0; z-index: 20; }",
            "  .toolbar-title { font-weight: 700; margin-right: 6px; }",
            "  .switch { border: 2px solid rgba(47,62,70,0.25); border-radius: 12px; overflow: hidden; display: inline-flex; background: #fff; }",
            "  .switch button { border: 0; background: transparent; padding: 8px 14px; font-weight: 700; cursor: pointer; color: #2f3e46; }",
            "  .switch button.active { background: #2f3e46; color: #fff; }",
            "  .hint { font-size: 12px; color: rgba(47,62,70,0.75); }",
            "  .viewer { height: calc(100vh - 66px); }",
            "  iframe { width: 100%; height: 100%; border: 0; background: #fff; display: none; }",
            "  iframe.active { display: block; }",
            "</style>",
            "</head>",
            "<body>",
            "  <div class='app'>",
            "    <div class='toolbar'>",
            "      <span class='toolbar-title'>Timeline style:</span>",
            "      <div class='switch'>",
            "        <button id='btn-horizontal' class='active' type='button'>Horizontal</button>",
            "        <button id='btn-vertical' type='button'>Vertical</button>",
            "      </div>",
            "      <span class='hint'>Switch instantly without regenerating data.</span>",
            "    </div>",
            "    <div class='viewer'>",
            "      <iframe id='frame-horizontal' class='active' title='Horizontal timeline'></iframe>",
            "      <iframe id='frame-vertical' title='Vertical timeline'></iframe>",
            "    </div>",
            "  </div>",
            f"<script id='payload' type='application/json'>{payload_json}</script>",
            "<script>",
            "(function(){",
            "  const payload = JSON.parse(document.getElementById('payload').textContent);",
            "  const hFrame = document.getElementById('frame-horizontal');",
            "  const vFrame = document.getElementById('frame-vertical');",
            "  const hBtn = document.getElementById('btn-horizontal');",
            "  const vBtn = document.getElementById('btn-vertical');",
            "  hFrame.srcdoc = payload.horizontal;",
            "  vFrame.srcdoc = payload.vertical;",
            "  function setMode(mode){",
            "    const isHorizontal = mode === 'horizontal';",
            "    hFrame.classList.toggle('active', isHorizontal);",
            "    vFrame.classList.toggle('active', !isHorizontal);",
            "    hBtn.classList.toggle('active', isHorizontal);",
            "    vBtn.classList.toggle('active', !isHorizontal);",
            "  }",
            "  hBtn.addEventListener('click', () => setMode('horizontal'));",
            "  vBtn.addEventListener('click', () => setMode('vertical'));",
            "})();",
            "</script>",
            "</body>",
            "</html>",
        ]
    )


def generate_combined_timeline(excel_path: str | Path, output_path: str | Path) -> None:
    excel_path = Path(excel_path)
    output_path = Path(output_path)

    with tempfile.TemporaryDirectory() as temp_dir:
        tmp_dir = Path(temp_dir)
        horizontal_path = tmp_dir / "horizontal.html"
        vertical_path = tmp_dir / "vertical.html"

        generate_horizontal_timeline(excel_path, horizontal_path)
        generate_vertical_timeline(excel_path, vertical_path)

        horizontal_html = horizontal_path.read_text(encoding="utf-8")
        vertical_html = vertical_path.read_text(encoding="utf-8")

    combined_html = _build_combined_html(horizontal_html, vertical_html)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(combined_html, encoding="utf-8")
    print(f"Combined timeline saved to {output_path.resolve()}")

