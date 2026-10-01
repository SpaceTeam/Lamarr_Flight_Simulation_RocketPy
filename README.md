# Lamarr Flight Simulation (RocketPy)

Flight simulation for the rockets of TU Wien Space Team, built on [RocketPy](https://github.com/RocketPy-Team/RocketPy).

---

## Project structure
```
Lamarr_Flight_Simulation_RocketPy/
├── simulation_core/                  shared code, used by the Jupyter notebook and the Streamlit app
│   ├── config_schema.py              pydantic models of config.toml and zones.toml, source of config_schemas/
│   ├── config_schemas/               generated JSON schemas for Tombi (never edit by hand)
│   ├── simulation.py                 runs the main simulation: builds environments, motor, rocket and flights from the config
│   ├── deployable_payload.py         simulates a separating payload as its own flight
│   ├── outputs.py                    prints, plots, landing safety classification, KML/CSV export
│   ├── reanalysis.py                 compares flight-computer data with the simulation after a flight
│   ├── export_weather_data.py        downloads atmosphere profiles for 'custom_atmosphere'
│   ├── kml_zones.py                  converts zones drawn in Google Earth (KML) to the zones.toml format
│   ├── custom_print_and_plot_functions.py   additional prints and plots
│   └── utils.py                      config and zone loading, variation handling, helpers
├── jupyternb/                        Jupyter notebook
│   └── simulation_orchestration.ipynb   the simulation pipeline: set PROJECT, then run all cells and writes report.html
├── streamlit_app/                    Streamlit app
│   ├── app.py                        pages: edit config.toml / zones.toml, run the simulation, view report and plots
│   ├── config_form.py                config form, generated from the pydantic models
│   ├── config_writer.py              creates and updates config.toml / zones.toml from the form values, keeping comments
│   └── run_simulation.py             runs simulation_orchestration.ipynb headless for one project
├── projects/                         data used by the notebook and the Streamlit app
│   └── ALBATROSS/, CANSAT/, ...      one folder per project: config.toml, zones.toml, thrust and drag curves, flight-computer data;
│                                     outputs (plots/, trajectory_kml/, trajectory_csv/, weather_csvs/, report.html) are written here
├── standalone_simulation_notebooks/  standalone Jupyter notebook versions (ALBATROSS, EuRoC 2025, LAMARR)
├── prepare_drag_data.py              splits drag.csv from OpenRocket into power_on_drag.csv and power_off_drag.csv for RocketPy
├── requirements.txt                  Python packages
├── .pre-commit-config.yaml           git hook that keeps simulation_core/config_schemas/ in sync with config_schema.py
└── .vscode/settings.json             Run on Save: regenerates simulation_core/config_schemas/ when config_schema.py is saved
```

---

## Initial setup (once after cloning from GitHub)
1. Use **Python 3.14**
2. **Create one virtual environment** for each project, so each one can have it's own RocketPy version and thus the simulation results do not change, if you run it again later.

3. **Install the packages:** `pip install -r requirements.txt` (adjust rocketpy version if needed)
4. **Install the git helpers** (run inside the repo folder, with the venv active):
   - `nbstripout --install`: removes notebook outputs before they are committed (configured in `.gitattributes`).
   - `pre-commit install`: before each commit, regenerates `simulation_core/config_schemas/`; if they were out of date the commit stops, and you `git add simulation_core/config_schemas` and commit again.
   - They save the path of that venv's Python, so re-run them if you delete or move the venv. It does not have to be repeated for every used venv in this repo. 
5. **VS Code extensions:**
   - **Tombi** (`tombi-toml.tombi`): autocomplete with `Ctrl+Space`, hover help and error checking in the TOML files.
   - **Run on Save** (`pucelle.run-on-save`): regenerates `simulation_core/config_schemas/` and restarts Tombi when `simulation_core/config_schema.py` is saved.
   - Disable **Even Better TOML** if it is installed, so the TOML files are not checked twice.
6. **Select the venv as Python interpreter** (`Python: Select Interpreter`) and as notebook kernel. Run on Save uses this interpreter too.
7. If you open the project through a `.code-workspace` file, copy `runOnSave.commands` from `.vscode/settings.json` into its `settings`:
   VS Code ignores this setting in folder settings when a workspace file is opened.
8. **Run a simulation**, one of:
   - Notebook: open `jupyternb/simulation_orchestration.ipynb`, set `PROJECT` (e.g. `"ALBATROSS"`) and run all cells.
   - App: see [Streamlit app](#streamlit-app).

---

## Streamlit app
### Start
1. Open a terminal in the repo folder and activate the venv.
2. Run `streamlit run streamlit_app\app.py`. The app opens in the browser at http://localhost:8501.
3. Stop it with `Ctrl+C` in the terminal.

### Use
- **Sidebar:** choose the page: *Edit project* (then pick the project) or *New project*.
- **New project page:** builds a new project folder in `projects/`. Start from an empty form or from the values of an existing project,
  fill in `config.toml` (and optionally `zones.toml`), upload the input files or copy them from the existing project, then press *Create project*.
  The values are checked before anything is written. The new files are written fresh, so comments of the existing project are not carried over.
- **Config tab:** forms for `config.toml` and `zones.toml`. Values are typed as TOML, like in the file: `350`, `[0, 90, 180]`, `"84..88:2"`, text in quotes.
  Hover the `?` of a field for its meaning and unit. *Save* checks the values first and keeps the file's comments; *Discard changes* reloads the file.
  *Import zones from a KML file* (in the `zones.toml` tab) fills the zone fields with the polygons of a KML file; *Save* writes them, see [Zones from Google Earth](#zones-from-google-earth).
- **Run tab:** runs `jupyternb/simulation_orchestration.ipynb` for the project and shows the progress and output live.
  While it runs, the *Run simulation* button becomes *Stop simulation*. Keep the tab open: any other click in the app also stops the run.
  After the run, its result and full output stay visible until the page is reloaded.
- **Results tab:** the project's `report.html` and all plots saved in `projects/<PROJECT>/plots/`.

---

## Zones from Google Earth
The landing zones in `zones.toml` can be drawn in Google Earth and imported from a KML file.

1. In Google Earth, add a **point** named `launch_rail` at the launch rail location.
2. Draw one **polygon** per zone. A polygon whose name contains `exclusion` becomes an exclusion zone, all others buffer zones.
   A `(exclusion)` or `exclusion` marker is removed from the name, e.g. `Cars (exclusion)` or `Cars exclusion` becomes the exclusion zone `"Cars"`.
3. Export the project as KML.
4. Import it, one of:
   - Streamlit app: `zones.toml` tab → *Import zones from a KML file* → *Fill the zone fields*, check them, then *Save zones.toml*.
     All zones are replaced; `exclusion_zone_safety_margin` and the comments stay. Without a `zones.toml`, the import first creates an empty one.
   - Terminal: `python simulation_core/kml_zones.py "projects/ALBATROSS_II/ALBATROSS II.kml"` prints the zones, ready to copy into `zones.toml`.

Each corner becomes `[distance_m, heading_deg]` (polar coordinates) from `launch_rail`, measured on the WGS84 ellipsoid (heading: 0 = North, 90 = East, clockwise).

---

## Config files
The config files are TOML files in the project folder: `projects/<PROJECT>/config.toml` and `projects/<PROJECT>/zones.toml`.
Everything about a single field (meaning, unit, allowed values, which fields are required together) is in the hover help:
hover a key or `[section]` header in VS Code (Tombi setup: see [Initial setup](#initial-setup-once-after-cloning-from-github)).

### How the pieces fit together
- **TOML:** the config file format (syntax): `key = value`, `[sections]`, lists.
- **Pydantic models** (`simulation_core/config_schema.py`): single source of truth for the definition of which keys and values are allowed.
- **JSON Schema** (`simulation_core/config_schemas/*.json`): the same rules, generated from the models in a format IDEs understand; never edit by hand.
- **Tombi** (VS Code extension): reads the JSON Schema and gives errors, hover help and autocomplete while typing in the TOML files.

The rules are used at two moments:
```
    While editing:     config_schema.py --(python config_schema.py)--> config.schema.json --> Tombi --> config.toml
    When simulating:   config.toml --(tomllib)--> dict --(Config.model_validate)--> checked Config object --> simulation
```
`tomllib` only checks the TOML syntax; `model_validate` checks the content (unknown keys such as typos, wrong types, range strings).
Pydantic is the check that counts; Tombi only gives the same feedback earlier.

### Changing the models
- Each field's description and examples are written in the `Attributes` section of its class docstring, which is also the Tombi hover text.
  How to write it: see the docstring of `_Base` in `config_schema.py`.
- Fields with a fixed set of values (e.g. `engine.type`, `envType`, `nosecone.kind`, `reanalysis.sources`) are defined as `Literal` types at the top of `config_schema.py`; hovering the key in a config file lists them.
- Without the VS Code extension `Run on Save`: after changing the models, run `python simulation_core/config_schema.py`, then `Tombi: Restart Language Server` from the Command Palette.
- Commit the regenerated schema files together with the model change (the pre-commit hook enforces this).
- An already open TOML file only shows new errors after the next keystroke in it.

### Syntax: general
| Type | Format | Example |
|------|--------|---------|
| Comment | `#` until the end of the line | `length = 659  # [mm] measured` |
| Disabled entry | comment the line out | `# inclination = "85..89:2"` |
| Section | `[section]` or `[section.subsection]` | `[environment]`, `[parachutes.main]` |
| Small table on one line | `key = { a = x }` | `reanalysis_csv = { icon_d2 = "file.csv" }` |
| Date and time | TOML local date-time, in launch-site time | `date = 2026-05-23T13:12:00` |

### Syntax: variations
Used to simulate varying flight scenarios, e.g. `flight.heading`, `flight.inclination`, `payload.mass`, `parachutes.payload.cd_s`.\
Fields that can be varied say *Variable* in their hover help (in `config_schema.py`: typed with `FLOAT_RANGE_EXPANSION` / `INT_RANGE_EXPANSION`). The simulation runs one flight per combination of all varied fields (full cartesian product).
Range strings are expanded to lists while the config is validated (see `FLOAT_RANGE_EXPANSION` in `config_schema.py`), so the rest of the pipeline only ever sees lists or scalar values.

| Type | Format | Example | Result |
|------|--------|---------|--------|
| List | `a = [x, y, z]` | `heading = [182, 190, 199]` | `[182, 190, 199]` |
| Range string | `a = "start..stop:step"` <br> end inclusive if the step does not overshoot; <br> if start > stop the range wraps around 360° (heading only) | `inclination = "86..90:2"`<br> `inclination = "160..200:15"`<br> `heading = "180..90:10"` | `[86, 88, 90]`<br> `[160, 175, 190]`<br> `[180, 190, ..., 350, 0, 10, ..., 90]`|

### Syntax: constants
Everything else is a constant.

| Type | Format | Example |
|------|--------|---------|
| String | `a = "abc"` | `date = "tomorrow_08_local"`|
| Scalar | `a = x` | `output_level = 2` |
| List | `a = [x, y, z]` | `envType = ["Windy", "custom_atmosphere"]` |
| Nested list | `a = [[v, w], [x, y]]` | `shape_points = [[0, 0], [0.250, -0.012]]` |
| Boolean | `a = false` | `enabled = false` |

### Notes
- TOML has no `null`: leave an optional key out instead.

---

## Simulation with variations
When any field is varied, the flight ascent is simulated once per combination and reused for the descent scenarios (this also happens without variations when a deployable payload is present). KML export is disabled when fields are varied.

**Scenarios produced per combination depend on what varies and the environment type:**

| What varies | Environment | Scenarios | Engine/rocket |
|---|---|---|---|
| nothing (single flight) | standard / forecast / custom | nominal + no-main + ballistic | pre-built once |
| nothing (single flight) | reanalysis | nominal only (+ matched flight from reanalysis section) | pre-built once |
| `payload.mass_total` and/or `parachutes.payload.*` and/or flight parameters only (`heading`, `inclination`, `rail_length`) | any | nominal + no-main + ballistic | pre-built once |
| any rocket component field | any | nominal only | rebuilt per combination |
| rocket components + flight parameters mixed | any | nominal only | rebuilt per combination |

### Rocket component sections
Varying any field here triggers a per-combo rebuild:
`motor`, `liquid_engine`, `rocket`, `nosecone`, `railbuttons`, `tailcone`, `fins`, `parachutes.main`, `parachutes.drogue`.

---

## Sharing results
You can share the `report.html` using `htmlpreview.github.io`.

Example:
https://htmlpreview.github.io/?https://github.com/SpaceTeam/Lamarr_Flight_Simulation_RocketPy/blob/config-variation-simulation/projects/ALBATROSS_II/report.html

Due to the file size it takes some time to load.