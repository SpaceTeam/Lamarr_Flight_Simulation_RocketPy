"""
This script processes a CSV file (drag.csv) containing rocket drag coefficient and mach number data and
generates two new CSV files: power_on_drag.csv and power_off_drag.csv.

Run:
    python prepare_drag_data.py drag.csv

Behavior:
- Detects header like: "# Drag coefficient (...),Mach number (...)"
    If "Drag coefficient" appears before "Mach number", it swaps the two values
    so outputs are always: Mach,DragCoefficient

- power_on_drag.csv:
    Extracts rows starting after the line starting with "# Event LAUNCH"
    Stops before a line starting with "# Event BURNOUT"
    Ignores comment lines (starting with "#")

- power_off_drag.csv:
    Extracts rows starting after the line starting with "# Event BURNOUT"
    Stops before a line starting with "# Event APOGEE"
    Ignores comment lines (starting with "#")

Notes:
- Outputs contain only data rows (no header).
- Handles comma as the delimiter.
"""

import sys
from pathlib import Path
import pandas as pd

MARK_LAUNCH = "# Event LAUNCH"
MARK_BURNOUT = "# Event BURNOUT"
MARK_APOGEE = "# Event APOGEE"

RED = "\033[91m"
RESET = "\033[0m"


def extract_section(lines, start_marker, end_marker):
    """
    Return all lines between start_marker and end_marker. Excluding lines starting with #.
    """
    capture = False
    section_lines = []
    
    for line in lines:
        if line.startswith(start_marker):
            capture = True
            continue
        if line.startswith(end_marker):
            break
        if capture and not line.startswith("#"):
            section_lines.append(line)
            
    return section_lines


def to_df(data_lines, drag_first):
    """
    Write lines to pandas dataframe in correct order.
    """
    df = pd.DataFrame([line.split(",") for line in data_lines], columns=["A", "B"])
    
    if drag_first:
        df = df.rename(columns={"A": "CD", "B": "Mach"})
    else:
        df = df.rename(columns={"A": "Mach", "B": "CD"})
        
    # Ensure correct order: Mach, DragCoefficient
    return df[["Mach", "CD"]]


def main():
    if len(sys.argv) < 2:
        print(f"{RED}Usage: python prepare_drag_data.py <input.csv>{RESET}")
        sys.exit(1)

    # Take the file path from sys.argv[1] and make it absolute. If it's relative (e.g., "../drag.csv" or "drag.csv"), create an absolute one by resolving it
    # against the current working directory.
    input_path = Path(sys.argv[1]).resolve()
    
    if not input_path.is_file():
        print(f"{RED}Error: file not found: {input_path}{RESET}")
        sys.exit(1)

    # Read all lines: create a list that contains every line as a entry
    lines = input_path.read_text().splitlines()

    # Detect if drag comes before mach
    header_line = next((l for l in lines if "drag coefficient" in l.lower() and "mach number" in l.lower()), None)
    if not header_line:
        print(f"{RED}Drag/Mach Header not found in the CSV file!{RESET}")
        sys.exit(1)
    
    drag_first = header_line.lower().find("drag") < header_line.lower().find("mach")

    # Extract both sections
    on_lines = extract_section(lines, MARK_LAUNCH, MARK_BURNOUT)
    off_lines = extract_section(lines, MARK_BURNOUT, MARK_APOGEE)

    # save lines to dataframe
    df_on = to_df(on_lines, drag_first)
    df_off = to_df(off_lines, drag_first)

    out_dir = input_path.parent
    on_outpath = out_dir / "power_on_drag.csv"
    off_outpath = out_dir / "power_off_drag.csv"

    df_on.to_csv(on_outpath, index=False, header=False)
    df_off.to_csv(off_outpath, index=False, header=False)

    print(f"Saved: {on_outpath}")
    print(f"Saved: {off_outpath}")


if __name__ == "__main__":
    main()