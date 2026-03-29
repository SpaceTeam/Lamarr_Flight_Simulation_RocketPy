import os
import re
import time
import json
import pathlib
import subprocess
import streamlit as st

from PIL import Image
from datetime import datetime
from simulation.utils import *
from streamlit_utils.streamlit_utils import *

constants = [] 
variables = []
sim_executed = False

projects = [
    p.name for p in pathlib.Path(".").iterdir()
    if (p / "config.txt").exists()
]
project = st.sidebar.selectbox("Project", projects)

if project:
    CONFIG_PATH = f"{project}/config.txt"
    PLOTS_PATH = f"{project}/plots/"

    if not os.path.exists(CONFIG_PATH):
        st.error(f"Config file not found: {CONFIG_PATH}")
        st.stop()
    else:
        lines, constants, variables = read_config_raw(CONFIG_PATH)
        if "project" not in st.session_state or st.session_state.project != project:
            st.session_state.project = project
            st.session_state.lines = lines.copy()

        lines = st.session_state.lines

st.title(project)

# ----------------------------
# Tabs
# ----------------------------

tab1, tab2, tab3 = st.tabs(["Config Editor", "Program Output", "PLOTS"])


with tab1:
    st.title("Config Editor")

    if st.button("Save Config", key="save_config_top"):

        errors = validate_params(lines)

        if errors:
            st.error(f"Syntax error in: {', '.join(errors)}")
        else:
            write_config_raw(CONFIG_PATH, lines)
            st.success("Config saved")

    st.header("CONSTANTS")

    for i,entry in enumerate(lines):

        if entry["type"] != "constant":
            continue
        
        col1, col2 = st.columns([4,1])

        with col1:
            text = st.text_input(
                entry["name"],
                entry["value"],
                key=entry["name"]
            )

            entry["value"] = text

            # preview parsing
            try:
                parsed = parse_value(text)
                st.caption(f"Parsed → {parsed}")

            except Exception:
                st.error("Invalid syntax")

        with col2:
            if st.button("❌", key=f"delete_{i}_const"):
                st.session_state.lines.pop(i)
                st.rerun()

    st.divider()



    updated_variables = {}

    st.header("VARIABLES")

    for i,entry in enumerate(lines):

        if entry["type"] != "variable":
            continue
        
        col1, col2 = st.columns([4,1])

        with col1:
            text = st.text_input(
                entry["name"],
                entry["value"],
                key=entry["name"]
            )

            entry["value"] = text

            # preview parsing
            try:
                parsed = parse_value(text)
                st.caption(f"Parsed → {parsed}")

            except Exception:
                st.error("Invalid syntax")

        with col2:
            if st.button("❌", key=f"delete_{i}_var"):
                st.session_state.lines.pop(i)
                st.rerun()

    if st.button("Save Config", key="save_config_bottom"):

        errors = validate_params(lines)

        if errors:
            st.error(f"Syntax error in: {', '.join(errors)}")
        else:
            write_config_raw(CONFIG_PATH, lines)
            st.success("Config saved")
    
    st.divider()

    st.subheader("Add Entry")

    new_name = st.text_input("Name")
    new_value = st.text_input("Value")
    try:
        parsed = parse_value(new_value)
        st.caption(f"Parsed → {parsed}")
    except Exception:
        st.error("Invalid syntax")
    inline_comment = st.text_input("Inline Comment", placeholder="(optional)", value = None)

    if inline_comment is not None and not inline_comment.strip().startswith("#"):
        inline_comment = "# " + inline_comment

    if st.button("Add Entry"):

        if any(e.get("name") == new_name for e in st.session_state.lines):
            st.error("Entry already exists")

        else:
            st.session_state.lines.append({"type": "blank"})
            st.session_state.lines.append({
                "type": "variable" if isinstance(parsed, list) else "constant",
                "name": new_name,
                "value": new_value,
                "inline_comment": inline_comment
            })

            st.rerun()

with tab2:

    if st.button("Run Program"):

        # Example: assume validate_params passes
        errors = []
        if errors:
            st.error("Cannot run. Syntax errors detected.")
        else:
            st.subheader("Program Output")

            # Start subprocess (unbuffered)
            proc = subprocess.Popen(
                ["python", "-u", "./main.py", f"{project}"],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True
            )

            # Live output container

            # Stream lines as they come
            while True:
                line = proc.stdout.readline()
                if not line:
                    break
                line = line.rstrip()
                if line.startswith("RESULT_JSON:"):
                    data = json.loads(line.replace("RESULT_JSON:",""))
                    constants = data["constants"]
                    variables = data["variables"]
                else:
                    # Render the latest output as Markdown or text
                    latest_line = line
                    if re.match(r"^\s*([#*>\-]|```)", latest_line):
                        st.markdown(latest_line)
                    else:
                        st.text("\n" + latest_line)

                # Yield control to Streamlit to force update
                time.sleep(0.01)

            proc.stdout.close()
            proc.wait()
            sim_executed = True

with tab3:

    heading_to_plot = st.number_input("Heading (0-360°)", 0.0, 360.0, value=None, placeholder="optional")
    inclination_to_plot = st.number_input("Inclination (0-90°)", 0.0, 90.0, value=None, placeholder="optional")
    overwrite = st.text_input("Overwrite", value="False", placeholder="optional")
    overwrite = overwrite == 'True'

    if st.button("Plot"):
        # Plot all flights for single heading and inclination
        
        h = "" if heading_to_plot is None else f"_heading_{int(heading_to_plot)}"
        i = "" if inclination_to_plot is None else f"_inclination_{int(inclination_to_plot)}"

        plot_path = f"{project}/plots/safe_flights{h}{i}.png"

        if os.path.exists(plot_path) and not overwrite:
            st.image(plot_path)
        else:
            if sim_executed:
                plot_safe_flights(project, lookup("exclusion_zones", constants, variables)[0], lookup("buffer_zones", constants, variables)[0], flight_groups={
                    "rocket_nominal": (lookup("safe_rocket_nominal", constants, variables)[0], "green"),
                    "rocket_no_main": (lookup("safe_rocket_no_main", constants, variables)[0], "orange"),
                    "rocket_ballistic": (lookup("safe_rocket_ballistic", constants, variables)[0], "red"),
                    "payload_nominal": (lookup("safe_payload_nominal", constants, variables)[0], "blue"),
                }, heading = heading_to_plot, inclination = inclination_to_plot)
                
                st.image(plot_path)
            else:
                st.info("First execute simulation")
            

    st.header("Existing Plots")
    plots_dir = pathlib.Path(PLOTS_PATH)
    if plots_dir is not None:
        images = list(plots_dir.glob("*.png")) + list(plots_dir.glob("*.jpg"))

        if not images:
            st.info("No images found.")
        else:
            # --- User selects sort order ---
            sort_option = st.selectbox(
                "Sort images by",
                ["Name (A-Z)", "Name (Z-A)", "Last Modified (Oldest First)", "Last Modified (Newest First)"]
            )

            # --- Sort based on selection ---
            if sort_option == "Name (A-Z)":
                images.sort(key=lambda p: p.stem.lower())
            elif sort_option == "Name (Z-A)":
                images.sort(key=lambda p: p.stem.lower(), reverse=True)
            elif sort_option == "Last Modified (Oldest First)":
                images.sort(key=lambda p: p.stat().st_mtime)
            elif sort_option == "Last Modified (Newest First)":
                images.sort(key=lambda p: p.stat().st_mtime, reverse=True)

            # --- Display images ---
            for img_path in images:
                name = re.sub(r"_+", " ", img_path.stem).title()
                modified = datetime.fromtimestamp(img_path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                st.subheader(f"{name} (Last modified: {modified})")
                img = Image.open(img_path)
                st.image(img, use_column_width=True)