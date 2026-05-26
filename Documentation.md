### Config rules
The config is a JSON file placed in the project folder. e.g. `ALBATROSS/config.json`.

**General**
| Type | Format | Example | Result |
|------|--------|---------|--------|
| Comment / disabled entry | Any key starting with `_` is ignored | `"_inclination": "85..89:2"` | ignored in the loaded config |
| Nested object | `"a": { "b": x }` flattens keys with `_` | `"environment": { "latitude": 39.23 }` | `environment_latitude = 39.23` |


It is loaded into two dictionaries: **Variations** and **Constants**, with `{flat_key: value}`.

**Variations**\
Used to simulate varing flight scenarios. E.g. `flight_heading`, `flight_inclination`, `payload_mass`, `parachutes_payload_cd_s`.\
Only allowed for keys in `VARIATION_KEYS`.

| Type | Format | Example | Result |
|------|--------|---------|--------|
| List | `"a": [x, y, z]` | `"flight": {"heading": [182, 190, 199]}` | `flight_heading = [182, 190, 199]` |
| Range string | `"a": "start..stop:step"` <br> end inclusive if the step does not overshoot; <br> if start > stop the range wraps around 360° (heading only) | `"flight": {"inclination": "86..90:2"}`
<br> `"flight": {"inclination": "160..200:15"}`<br> `"flight": {"heading": "180..90:10"}` | `flight_inclination = [86, 88, 90]`<br> `flight_inclination = [175, 190]`<br> `flight_heading = [180, 190, ..., 350, 0, 10, ..., 90]`|

**Constants**\
Everything else is a constant.

| Type | Format | Example | Result |
|------|--------|---------|--------|
| Scalar | `"a": x` | `"latitude": 39.232465` | `latitude = 39.232465` |
| List | `"a": [x, y, z]` <br> JSON arrays become tuples | `"environment": { "envType": ["Windy", "custom_atmosphere"] }` | `environment_envType = ("Windy", "custom_atmosphere")` |
| Nested list | `"a": [[v,w], [x,y]]` | `"fins_shape_points": [[0,0], [0.250,-0.012]]` | `(((0, 0), (0.250, -0.012)))` |

**Notes**
- `Boolean values` work as constants and inside variation lists.

---

### Reserved keys and accepted values
The backend recognizes specific keys and only accepts certain values for them.

**Environment** (`environment_*`)
- `envType`: list of strings, each one of `"standard_atmosphere"`, `"Windy"`, `"custom_atmosphere"`, `"reanalysis"`, `"reanalysis_custom"`.
    - `"reanalysis"` and `"reanalysis_custom"` are exclusive. If used, they must be the **only** entry in the list.
    - `"Windy"` requires `environment_windy_weather_models` (list).
    - `"custom_atmosphere"` requires `environment_custom_weather_models` (list).
    - `"reanalysis"` requires `environment_reanalysis_file` and `environment_reanalysis_dictionary`.
    - `"reanalysis_custom"` requires `environment_reanalysis_csv` (list).
- `date`: either the literal string `"tomorrow_08_local"` (= tomorrow 08:00 local time) or a list `[YYYY, M, D, h, m, s]` (4-element lists are auto-padded with `0, 0`).
- Required: `envType`, `date`, `latitude`, `longitude`, `max_expected_height`, `elevation`, `timezone`.

**Engine** (`engine_*`)
- `type`: `"solid"` or `"liquid"`. Selects between `create_solid_engine` and `create_liquid_engine`.

**Motor** (`motor_*`, solid engine)
- Required: `name`, `total_mass`, `propellant_mass`, `diameter`, `length`, `burn_time`, `center_of_dry_mass_factor`, `thrust`.
- `center_of_dry_mass_factor`: fraction of motor length (0–1) used to place the CG from the nozzle.
- `thrust`: file name in the project folder (`.eng`, `.rse`, or `.csv`).

**Rocket** (`rocket_*`)
- Required: `diameter`, `length`, `total_mass_without_motor`, `total_CG_without_motor_from_tip`, `moment_of_intertia_XY`, `moment_of_intertia_Z`, `power_off_drag`, `power_on_drag`.
- `power_off_drag` / `power_on_drag`: file names in the project folder.

**Nosecone** (`nosecone_*`)
- Required: `length`, `cylindrical_section_length`, `kind`.
- `kind`: passed to RocketPy (e.g. `"vonKarman"`, `"conical"`, `"ogive"`, ...).

**Tailcone** (`tailcone_*`)
- Required: `bottom_radius`, `cylindrical_section_length`, `length`.

**Fins** (`fins_*`)
- Required: `amount`, `name`, `position`.
- If `fins_shape_points` is provided → `FreeFormFins` is used.
- Otherwise → `TrapezoidalFins`, which also requires `root_chord`, `tip_chord`, `span`, `sweep_length`.

**Parachutes** (`parachutes_main_*`, optional `parachutes_drogue_*` and `parachutes_payload_*`)
- Required for each parachute: `trigger`, `sampling_rate`.
- `cd_s` may be supplied directly OR computed from `cd + fabric_area` OR `cd + radius`.
- Optional: `lag`, `noise`; if `cd_s` is provided, `radius`, `cd` and `fabric_area` are optional.
- Drogue is only created if any `parachutes_drogue_*` key exists.
- Payload parachute is only created if any `parachutes_payload_*` key exists.

**Railbuttons** (`railbuttons_*`)
- Required: `upper_from_tip`, `lower_from_tip`.

**Flight** (`flight_*`)
- Required: `rail_length`, `inclination`, `heading`.
- `inclination`: degrees, 0 = horizontal, 90 = vertical.
- `heading`: degrees, 0 = north, 90 = east. A wrap-around range (start > stop) crosses 0°/360°, e.g. `"180..90:10"` → 180, 190, …, 350, 0, 10, …, 90.

**Payload** (`payload_*`)
- Only activated when `payload_mass_total > 0`.
- Required: `diameter`, `mass`, `length`, `moment_of_intertia_XY`, `moment_of_intertia_Z`.

---

### Unit conventions
- **Lengths, diameters, positions**: millimeters in the config, auto-converted to meters internally.
- **Masses**: grams in the config, auto-converted to kg internally.
- **Exceptions** (already in SI): `rail_length` (m), `elevation` (m ASL), `max_expected_height` (m AGL), `latitude`/`longitude` (°), `inclination`/`heading` (°), `burn_time` (s), `sampling_rate` (Hz).
- `*_from_tip` keys are measured from the **nose tip** and converted to RocketPy's bottom-origin coordinate system internally.

---

### Zones config file

**Exclusion zones** are hard no-land areas, e.g. crowd, houses, forrest.
They are shown in **red** on landing maps. Any simulated landing inside an exclusion zone means the configuration is unsafe.

**Buffer zones** are a safety margin surrounding exclusion zones, shown in **orange**. You can also add custom buffer 
zones in `zones.json`.
The safety classifier checks buffer zones only: a configuration is **safe** if and only if every simulated
scenario (nominal, drogue-only, ballistic), every environment and inclination lands outside every buffer zone.

Each polygon vertex is `[distance_m, heading_deg]` obtained from Google Earth:
- `distance_m` is measured from launch site to zone corner
- `heading_deg` is measured clockwise from north; meaning 0° = North, 90° = East.