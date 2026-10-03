"""
Streamlit app for the simulation: edit a project's config.toml and zones.toml, run the simulation and look at the results.

Start from the repo folder with: streamlit run streamlit_app/app.py

Simulation executes run_simulation.py
"""
import os
import re
import shutil
import subprocess
import sys
import tomllib
from datetime import datetime
from pathlib import Path

import streamlit as st
from streamlit.delta_generator import DeltaGenerator
from pydantic import BaseModel, ValidationError

REPO_DIR = Path(__file__).resolve().parent.parent
RUN_SIMULATION_SCRIPT = Path(__file__).resolve().parent / "run_simulation.py"

if str(REPO_DIR) not in sys.path:
    sys.path.insert(0, str(REPO_DIR))
# Streamlit only reloads changed modules from this file's folder or from PYTHONPATH; with the repo folder in PYTHONPATH,
# a saved change in simulation_core (e.g. a field description in config_schema.py) shows up on the next rerun.
python_paths = [path for path in os.environ.get("PYTHONPATH", "").split(os.pathsep) if path]
if str(REPO_DIR) not in python_paths:
    os.environ["PYTHONPATH"] = os.pathsep.join([str(REPO_DIR), *python_paths])

from streamlit_app.run_simulation import ERROR_PREFIX, FLIGHT_PROGRESS_PREFIX, PROGRESS_PREFIX, PROJECTS_DIR, REPORT_FILE_NAME
from simulation_core.config_schema import Config, ZonesConfig
from simulation_core.kml_zones import EXCLUSION_KEYWORD, LAUNCH_RAIL_NAME, format_zones, read_kml_zones
from streamlit_app.config_writer import read_comments, update_config_file, write_config_file
from streamlit_app.config_form import clear_form_state, render_model_form, to_toml_text

CONFIG_FILE_NAME = "config.toml"
ZONES_FILE_NAME = "zones.toml"
PLOTS_FOLDER_NAME = "plots"
TOP_ANCHOR = "top"      # for back to top button

EDIT_PAGE = "Edit project"
NEW_PAGE = "New project"

EMPTY_TEMPLATE = "(empty form)"

OUTPUT_NAMES = {REPORT_FILE_NAME, PLOTS_FOLDER_NAME, "trajectory_kml", "trajectory_csv"}

EMBED_HEIGHT_PX = 800           # Height of the embedded report.html
# Added to the report when it is displayed (not to the file). 
REPORT_REDRAW_SCRIPT = """
<script>
/* Draw all Plotly figures again once the report frame is visible, so their layout is measured at the real size. */
var redrawTimer;
window.addEventListener("resize", function () {
  clearTimeout(redrawTimer);
  redrawTimer = setTimeout(function () {
    if (!window.Plotly || window.innerWidth === 0 || window.innerHeight === 0) { return; }
    document.querySelectorAll(".plotly-graph-div").forEach(function (figure) {
      window.Plotly.newPlot(figure, figure.data, figure.layout, figure._context);
    });
  }, 200);
});
</script>
"""
# How many of the latest output lines the Run tab shows while the simulation runs.
# LOG_TAIL_LINES = 200
# ANSI colour codes like "\x1b[31m" that IPython puts into tracebacks; the output box would show them as raw text -> remove them
ANSI_COLOR_CODE_PATTERN = re.compile(r"\x1b\[[0-9;]*m")


# =============================================================================
# Helpers
# =============================================================================

def list_projects() -> list[str]:
    """Return the project folders, i.e. the folders in projects/ that contain a config.toml."""
    return sorted(folder.name for folder in PROJECTS_DIR.iterdir() if (folder / CONFIG_FILE_NAME).is_file())


def format_validation_error(error: ValidationError) -> str:
    """Turn a pydantic error into one line per problem with the TOML path, e.g. "flight.heading: Input should be a valid number"."""
    return "\n".join(f"- `{'.'.join(str(part) for part in problem['loc'])}`: {problem['msg']}" for problem in error.errors())


def show_project_builder(projects: list[str]) -> None:
    """Build a new project in a form, empty or filled with an existing project's values, and write its files with write_config_file.

    The new project gets config.toml, optionally zones.toml, and its input files (thrust and drag curves, weather CSVs,
    flight-computer data) from an upload or copied from the project it starts from. Run outputs (report, plots, trajectories) are never copied.
    """
    st.title("New project", anchor=TOP_ANCHOR)
    name = st.text_input("Project name", placeholder="e.g. LAMARR_2027", help="Also the name of the new folder in projects/.").strip()
    template_project = st.selectbox(
        "Start from", [EMPTY_TEMPLATE, *projects], help="An empty form, or a form filled with the values of an existing project."
    )
    template_dir = None if template_project == EMPTY_TEMPLATE else PROJECTS_DIR / template_project
    config_values = tomllib.loads((template_dir / CONFIG_FILE_NAME).read_text(encoding="utf-8")) if template_dir else {}
    # The field comments of the template (e.g. "measured") are offered too, so they carry over into the new files.
    config_comments = read_comments(template_dir / CONFIG_FILE_NAME) if template_dir else {}
    template_zones_path = template_dir / ZONES_FILE_NAME if template_dir else None
    has_template_zones = template_zones_path is not None and template_zones_path.is_file()
    zones_values = tomllib.loads(template_zones_path.read_text(encoding="utf-8")) if has_template_zones else {}
    zones_comments = read_comments(template_zones_path) if has_template_zones else {}
    # The widget keys depend on the start choice, so switching it shows that project's values instead of the typed ones.
    key_prefix = f"new_project.{template_project}"

    st.subheader(CONFIG_FILE_NAME)
    st.caption("Values are typed as TOML: 350, [0, 90, 180], \"84..88:2\", text in quotes. Empty fields are left out of the file.")
    # "project" is set from the project name above, so it is not asked twice.
    new_config, invalid_fields, new_config_comments = render_model_form(
        Config, config_values, f"{key_prefix}.config", hidden_fields=frozenset({"project"}), comments=config_comments
    )

    st.subheader(ZONES_FILE_NAME)
    include_zones = st.checkbox(f"Create a {ZONES_FILE_NAME}", value=bool(zones_values), key=f"{key_prefix}.include_zones")
    new_zones, new_zones_comments = {}, {}
    if include_zones:
        new_zones, invalid_zone_fields, new_zones_comments = render_model_form(
            ZonesConfig, zones_values, f"{key_prefix}.zones", comments=zones_comments
        )
        invalid_fields += invalid_zone_fields

    st.subheader("Input files")
    uploaded_files = st.file_uploader(
        "Thrust curve, drag curves and other files the config refers to", accept_multiple_files=True, key=f"{key_prefix}.uploads"
    )
    copy_template_files = template_dir is not None and st.checkbox(
        f"Copy the input files of {template_project}", value=True, key=f"{key_prefix}.copy_files",
        help="Everything in that project folder (thrust and drag curves, weather CSVs, flight-computer data) except config.toml, "
        "zones.toml and the run outputs (report, plots, trajectories). Uploaded files replace copied ones with the same name.",
    )

    if not st.button("Create project", type="primary"):
        return
    if not name or (PROJECTS_DIR / name).exists():
        st.error("Enter a project name that is not used by another folder in projects/ yet.")
        return
    if invalid_fields:
        st.error(f"Not created: fix the TOML in {', '.join(invalid_fields)}.")
        return
    new_config["project"] = name
    # Check both files before writing anything, so an invalid zones.toml cannot leave a half-created project behind.
    try:
        Config.model_validate(new_config)
        if include_zones:
            ZonesConfig.model_validate(new_zones)
    except ValidationError as error:
        st.error("Not created, the values are invalid:\n" + format_validation_error(error))
        return

    project_dir = PROJECTS_DIR / name
    write_config_file(Config, new_config, project_dir / CONFIG_FILE_NAME, new_config_comments)
    if include_zones:
        write_config_file(ZonesConfig, new_zones, project_dir / ZONES_FILE_NAME, new_zones_comments)
    if copy_template_files:
        # config.toml and zones.toml were just written from the form; everything else except run outputs is copied.
        for template_item in template_dir.iterdir():
            if template_item.name in {CONFIG_FILE_NAME, ZONES_FILE_NAME, *OUTPUT_NAMES}:
                continue
            if template_item.is_dir():
                shutil.copytree(template_item, project_dir / template_item.name)
            else:
                shutil.copy2(template_item, project_dir / template_item.name)
    for uploaded_file in uploaded_files or []:
        (project_dir / uploaded_file.name).write_bytes(uploaded_file.getvalue())

    st.success(f"Created projects/{name}. Switch to *{EDIT_PAGE}* in the sidebar and select it to run the simulation.")


def show_kml_import(zones_path: Path, key_prefix: str) -> None:
    """Let the user upload a KML file (e.g. from Google Earth) and fill the zone fields of the form below with its polygons.

    Nothing is written until Save is clicked; a missing zones.toml is first created empty, so the form to fill exists.
    """
    with st.expander("Import zones from a KML file", icon=":material/upload_file:"):
        st.caption(
            f'The KML file needs a point named "{LAUNCH_RAIL_NAME}" and one polygon per zone. '
            f'Polygons whose name contains "{EXCLUSION_KEYWORD}" become exclusion zones, all others buffer zones. '
            f"Importing replaces all zones in the form below; Save writes them to {zones_path.name}."
        )
        kml_file = st.file_uploader("KML file", type=["kml"], key=f"{key_prefix}.kml_file")
        if kml_file is None:
            return
        try:
            exclusion_zones, buffer_zones = read_kml_zones(kml_file.getvalue())
        except ValueError as error:
            st.error(f"Could not read {kml_file.name}: {error}")
            return
        st.code(format_zones(exclusion_zones, buffer_zones), language="toml")
        if not st.button("Fill the zone fields", key=f"{key_prefix}.kml_import", type="primary"):
            return

        if not zones_path.is_file():
            write_config_file(ZonesConfig, {}, zones_path)
        # The form is drawn after this function, so setting its widget state now shows the imported zones in this same run.
        st.session_state[f"{key_prefix}.exclusion_zones"] = to_toml_text(exclusion_zones)
        st.session_state[f"{key_prefix}.buffer_zones"] = to_toml_text(buffer_zones)
        st.success(
            f"Filled in {len(exclusion_zones)} exclusion and {len(buffer_zones)} buffer zones from {kml_file.name}. "
            f"Click Save {zones_path.name} to keep them."
        )


def show_file_editor(model_class: type[BaseModel], path: Path, key_prefix: str) -> None:
    """Show the form for one config file with Save and Discard buttons; Save validates the values first and keeps the file's comments."""
    values = tomllib.loads(path.read_text(encoding="utf-8"))
    new_values, invalid_fields, new_comments = render_model_form(model_class, values, key_prefix, comments=read_comments(path))

    save_column, discard_column = st.columns(2)
    if discard_column.button("Discard changes", key=f"{key_prefix}.discard", width="stretch"):
        clear_form_state(key_prefix)
        st.rerun()
    if save_column.button(f"Save {path.name}", key=f"{key_prefix}.save", type="primary", width="stretch"):
        if invalid_fields:
            st.error(f"Not saved: fix the TOML in {', '.join(invalid_fields)}.")
            return
        try:
            update_config_file(model_class, new_values, path, new_comments)
        except ValidationError as error:
            st.error("Not saved, the values are invalid:\n" + format_validation_error(error))
            return
        st.success(f"Saved {path.relative_to(REPO_DIR)}")


def run_simulation(project: str, button_slot: DeltaGenerator) -> None:
    """Run run_simulation.py for the project and show its progress and output live; the log stays visible after the run."""
    # Replace the Run button with Stop: clicking it reruns the script, which interrupts this function and the finally below kills the process.
    button_slot.button("Stop simulation", icon=":material/stop:", key="stop_run")
    progress_bar = st.progress(0.0, text="Starting the notebook kernel...")
    flight_progress_slot = st.empty()
    log_area = st.empty()
    output_lines = []
    error_message = None
    # "import streamlit" sets MPLBACKEND=Agg for this process; without removing it, the notebook kernel would inherit it,
    # and matplotlib plots (plt.show) would be missing from the report with a "FigureCanvasAgg is non-interactive" warning.
    run_environment = {name: value for name, value in os.environ.items() if name != "MPLBACKEND"}

    process = subprocess.Popen(
        [sys.executable, "-u", str(RUN_SIMULATION_SCRIPT), project],
        cwd=REPO_DIR,
        env=run_environment,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    try:
        for line in process.stdout:
            # Keep the traceback text, drop its color codes (the saved log is shown the same way after the run).
            line = ANSI_COLOR_CODE_PATTERN.sub("", line)
            if line.startswith((PROGRESS_PREFIX, FLIGHT_PROGRESS_PREFIX)):
                # "PROGRESS: 3/9 Environments Initialization" -> prefix "PROGRESS:", counter "3/9" and label; same for flights.
                prefix, _, counter_and_label = line.strip().partition(" ")
                counter, _, label = counter_and_label.partition(" ")
                done_count, total_count = (int(number) for number in counter.split("/"))
                if prefix == PROGRESS_PREFIX:
                    progress_bar.progress(done_count / total_count, text=f"Step {counter}: {label}")
                    # A new step starts, so the flight bar of the previous step is no longer relevant.
                    flight_progress_slot.empty()
                else:
                    flight_progress_slot.progress(done_count / total_count, text=label)
                continue
            if line.startswith(ERROR_PREFIX):
                # "ERROR: OpenMeteoRequestsError: ..." -> remembered for the failure message, logged without the prefix.
                line = line.removeprefix(ERROR_PREFIX).lstrip()
                error_message = line.strip()
            output_lines.append(line)
            # log_area.code("".join(output_lines[-LOG_TAIL_LINES:]), language=None)
            log_area.code("".join(output_lines), language=None)
        process.wait()
    finally:
        # Any navigation in the app restarts this script; stop the simulation then instead of leaving it running unseen.
        stopped = process.poll() is None
        if stopped:
            process.kill()
        # Saved here, not after the try, because a rerun skips everything after the finally; this keeps a stopped run's log visible.
        st.session_state[f"last_run.{project}"] = {
            "log": "".join(output_lines),
            "return_code": process.returncode,
            "error_message": error_message,
            "stopped": stopped,
        }
    # Redraw the page so the Run button replaces Stop again; show_last_run then shows the saved result and log.
    st.rerun()


def show_last_run(project: str) -> None:
    """Show the result and full output of the project's last run in this browser session."""
    last_run = st.session_state.get(f"last_run.{project}")
    if last_run is None:
        return
    if last_run.get("stopped"):
        st.warning("Simulation stopped before it finished. The output below shows how far it got.")
    elif last_run["return_code"] == 0:
        st.success("Simulation finished. See the Results tab.")
    elif last_run.get("error_message"):
        st.error(
            f"Simulation failed (exit code {last_run['return_code']}):\n\n`{last_run['error_message']}`\n\n"
            "Scroll down for the traceback."
        )
    else:
        st.error(f"Simulation failed (exit code {last_run['return_code']}).")
    # The same log box as during the run, so the output stays visible after the rerun that ends it.
    st.code(last_run["log"], language=None)

def show_results(project_dir: Path) -> None:
    """Show the project's report.html: the run's prints and plots in notebook order."""
    report_path = project_dir / REPORT_FILE_NAME
    if report_path.is_file():
        modified = datetime.fromtimestamp(report_path.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
        st.caption(f"Report of the last run ({modified})")
        # The redraw script goes right before </body>, after all figures have been created.
        report_html = report_path.read_text(encoding="utf-8")
        body_end = report_html.rfind("</body>")
        st.iframe(report_html[:body_end] + REPORT_REDRAW_SCRIPT + report_html[body_end:], height=EMBED_HEIGHT_PX)
    else:
        st.info("No report yet. Run the simulation first.")


# =============================================================================
# Page
# =============================================================================

st.set_page_config(page_title="Flight Simulation", layout="wide")

projects = list_projects()
sidebar_controls = st.sidebar.container()
st.sidebar.markdown(f"[:material/arrow_upward: Back to top](#{TOP_ANCHOR})")
# bind="query-params" keeps the choice in the page URL (?page=...&project=...), so reloading or sharing the link opens the same view.
page = sidebar_controls.radio("Page", [EDIT_PAGE, NEW_PAGE], key="page", bind="query-params")
if page == NEW_PAGE:
    show_project_builder(projects)
    st.stop()

project = sidebar_controls.selectbox("Project", projects, key="project", bind="query-params")
if project is None:
    st.info(f"No project folder with a {CONFIG_FILE_NAME} found in projects/.")
    st.stop()

project_dir = PROJECTS_DIR / project
st.title(project, anchor=TOP_ANCHOR)
config_tab, run_tab, results_tab = st.tabs(["Config", "Run", "Results"])

with config_tab:
    st.caption("Values are typed as TOML, like in the file: 350, [0, 90, 180], \"84..88:2\". Hover the ? of a field for its meaning and unit.")
    config_file_tab, zones_file_tab = st.tabs([CONFIG_FILE_NAME, ZONES_FILE_NAME])
    with config_file_tab:
        show_file_editor(Config, project_dir / CONFIG_FILE_NAME, key_prefix=f"{project}.config")
    with zones_file_tab:
        zones_path = project_dir / ZONES_FILE_NAME
        show_kml_import(zones_path, key_prefix=f"{project}.zones")
        if zones_path.is_file():
            show_file_editor(ZonesConfig, zones_path, key_prefix=f"{project}.zones")
        else:
            st.info(f"This project has no {ZONES_FILE_NAME}; the simulation then runs without landing zones.")

with run_tab:
    st.caption("Runs jupyternb/simulation_orchestration.ipynb for this project. Keep this tab open while it runs.")
    button_slot = st.empty()
    run_clicked = button_slot.button("Run simulation", type="primary")
    if run_clicked:
        # this runs in the foreground; could be changed to run in the background so navigation in the app does not stop
        # the simulation (more code)
        run_simulation(project, button_slot)
    # Only reached when no run is active: run_simulation always ends with a rerun.
    show_last_run(project)

with results_tab:
    show_results(project_dir)
