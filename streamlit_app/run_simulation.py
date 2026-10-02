"""
Simulation runner for the Streamlit app: app.py starts this script for each run and shows its output live.

Executes simulation_orchestration.ipynb for the given project (the notebook stays the only definition of the pipeline),
prints its text output while it runs and writes the rendered notebook to projects/<PROJECT>/report.html.

Only the Streamlit app uses it. If needed, it can also be run in the terminal from the repo folder:
    python streamlit_app/run_simulation.py ALBATROSS
"""
import os
import re
import sys
import warnings
from functools import partial
from pathlib import Path

import nbformat
from nbclient import NotebookClient
from nbclient.exceptions import CellExecutionError
from nbconvert import HTMLExporter

REPO_DIR = Path(__file__).resolve().parent.parent
NOTEBOOK_PATH = REPO_DIR / "jupyternb" / "simulation_orchestration.ipynb"
# Folder that holds one subfolder per project (config.toml, zones.toml, inputs and outputs).
PROJECTS_DIR = REPO_DIR / "projects"
REPORT_FILE_NAME = "report.html"
# The notebook line that selects the project, e.g. PROJECT = "ALBATROSS".
PROJECT_LINE_PATTERN = re.compile(r"^PROJECT = .*$", re.MULTILINE)
# Tells the notebook kernel it runs headless, so outputs.export_notebook_to_html skips its export (it would wait for an editor save).
HEADLESS_ENV_VAR = "SIMULATION_HEADLESS"
# Lines starting with this prefix report progress ("PROGRESS: 3/9 Environments Initialization"); app.py turns them into a progress bar.
PROGRESS_PREFIX = "PROGRESS:"
# Prefix of the short error line of a failing cell ("ERROR: OpenMeteoRequestsError: ..."); app.py shows it in its failure message.
ERROR_PREFIX = "ERROR:"


def print_progress(cell, cell_index: int, code_cell_indices: list[int], heading_by_cell_index: dict[int, str], **_) -> None:
    """Print which code cell starts, e.g. "PROGRESS: 3/9 Environments Initialization" (nbclient calls this before each cell)."""
    # nbclient calls this for markdown cells too; only code cells count as steps.
    if cell.cell_type != "code":
        return
    position = code_cell_indices.index(cell_index) + 1
    print(f"{PROGRESS_PREFIX} {position}/{len(code_cell_indices)} {heading_by_cell_index[cell_index]}", flush=True)


def print_cell_output(cell, **_) -> None:
    """Print a finished cell's text output: prints, printmd() markdown and errors; plots are only in the report."""
    for output in cell.get("outputs", []):
        if output.output_type == "stream":
            print(output.text, end="", flush=True)
        elif output.output_type == "display_data" and "text/markdown" in output.data:
            print(output.data["text/markdown"], flush=True)
        elif output.output_type == "error":
            print(f"{ERROR_PREFIX} {output.ename}: {output.evalue}", flush=True)


def run_project(project: str) -> Path:
    """Execute the orchestration notebook for one project and write the result to <project>/report.html; returns the report path."""
    notebook = nbformat.read(NOTEBOOK_PATH, as_version=4)

    # Point the notebook at the requested project, like editing its PROJECT line by hand.
    project_cells = [cell for cell in notebook.cells if cell.cell_type == "code" and PROJECT_LINE_PATTERN.search(cell.source)]
    if not project_cells:
        raise ValueError(f"No 'PROJECT = ...' line found in {NOTEBOOK_PATH.name}.")
    project_cells[0].source = PROJECT_LINE_PATTERN.sub(f'PROJECT = "{project}"', project_cells[0].source, count=1)

    # Remember the last markdown heading before each code cell, so progress lines can name the current step.
    heading_by_cell_index = {}
    current_heading = "Setup"
    for cell_index, cell in enumerate(notebook.cells):
        if cell.cell_type == "markdown" and cell.source.lstrip().startswith("#"):
            current_heading = cell.source.lstrip().splitlines()[0].lstrip("#").strip(" *")
        heading_by_cell_index[cell_index] = current_heading
    code_cell_indices = [cell_index for cell_index, cell in enumerate(notebook.cells) if cell.cell_type == "code"]

    # The kernel inherits this process's environment; this script writes the report itself from the executed notebook.
    os.environ[HEADLESS_ENV_VAR] = "1"
    client = NotebookClient(
        notebook, timeout=None, resources={"metadata": {"path": str(REPO_DIR)}}, extra_arguments=["--IPKernelApp.log_level=ERROR"]
    )
    client.on_cell_start = partial(print_progress, code_cell_indices=code_cell_indices, heading_by_cell_index=heading_by_cell_index)
    client.on_cell_executed = print_cell_output

    report_path = PROJECTS_DIR / project / REPORT_FILE_NAME
    try:
        client.execute()
    finally:
        # Write the report even after an error, so the output up to the failing cell can be inspected.
        html_body, _ = HTMLExporter(theme="dark").from_notebook_node(notebook)
        report_path.write_text(html_body, encoding="utf-8")
        print(f"Report written to {report_path}", flush=True)
    return report_path


if __name__ == "__main__":
    # A pipe to app.py would otherwise use the Windows code page, which cannot print characters like "°".
    sys.stdout.reconfigure(encoding="utf-8")
    # Harmless on Windows (the kernel connection still works), but it would be the first line of every run's output.
    warnings.filterwarnings("ignore", message="Proactor event loop does not implement add_reader")
    if len(sys.argv) != 2:
        sys.exit("Usage: python streamlit_app/run_simulation.py <PROJECT>")
    try:
        run_project(sys.argv[1])
    except CellExecutionError as error:
        # Only Open-Meteo errors (e.g. a date outside its forecast range) are shown without a traceback; all other errors keep it.
        # Its short "ERROR: ..." line is already printed by print_cell_output, so the run just ends with exit code 1.
        if error.ename != "OpenMeteoRequestsError":
            raise
        sys.exit(1)
