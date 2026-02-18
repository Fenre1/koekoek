#%%
from __future__ import annotations

from pathlib import Path

from timeline_combined_viewer import generate_combined_timeline
from timeline_horizontal import generate_horizontal_timeline
from timeline_vertical_filterable import generate_vertical_timeline


#%%
# User config
mode = "combined"  # "horizontal" | "vertical" | "combined"
excel_path = Path("dummytimeline.xlsx")
output_path = Path("timeline_combined2.html")


#%%
def generate_timeline(mode: str, excel_path: Path, output_path: Path) -> None:
    selected = mode.strip().lower()
    if selected == "horizontal":
        generate_horizontal_timeline(excel_path, output_path)
        return
    if selected == "vertical":
        generate_vertical_timeline(excel_path, output_path)
        return
    if selected == "combined":
        generate_combined_timeline(excel_path, output_path)
        return
    raise ValueError("mode must be one of: horizontal, vertical, combined")


#%%
# Run cell
generate_timeline(mode, excel_path, output_path)


# %%
