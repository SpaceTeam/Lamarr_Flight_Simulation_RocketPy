# config.json structure

The config is a JSON file. Keys starting with `_` are treated as comments and ignored everywhere.

The top-level sections are:

```
project          – path to the project folder
environment      – launch site and atmosphere model
engine           – selects "solid" or "liquid" motor type
motor            – solid motor config (required when engine.type = "solid")
liquid_engine    – liquid motor config (required when engine.type = "liquid")
rocket           – rocket body geometry and mass
nosecone         – nosecone geometry
railbuttons      – rail button positions
tailcone         – tailcone geometry
fins             – fin set geometry
parachutes       – main, drogue, and payload parachutes
flight           – launch rail and heading/inclination
payload          – deployable payload (optional, disabled when mass_total = 0)
reanalysis       – post-flight comparison against flight-computer data (optional)
```

Field types and descriptions are defined in `config_schema.py`.

---

## Varying values / sweep mode

Any numeric config field accepts a **list** to sweep over multiple values. The simulation runs one flight per combination of all varied fields (full cartesian product).

```json
"heading": [0, 90, 180, 270]
```

Any numeric field also accepts a **range string** as a shorthand for a list:

```
"start..stop:step"
```

```json
"heading": "0..350:10"
```

For `heading` only, `start` > `stop` wraps around 360°:

```json
"heading": "350..10:5"   →  350, 355, 0, 5, 10
```

Range strings are expanded to lists by `load_config` before the config is validated, so the rest of the pipeline only ever sees lists or scalar values.

---

## Scan mode behaviour

When any field carries multiple values, the simulation enters **scan mode**: ascent is simulated once per combination and reused for the descent scenarios. KML export is disabled in scan mode.

**Scenarios produced per combination depend on what varies and the environment type:**

| What varies | Environment | Scenarios | Engine/rocket |
|---|---|---|---|
| nothing (single flight) | standard / forecast / custom | nominal + no-main + ballistic | pre-built once |
| nothing (single flight) | reanalysis | nominal only (+ matched flight from reanalysis section) | pre-built once |
| `payload.mass_total` and/or `parachutes.payload.*` and/or flight parameters only (`heading`, `inclination`, `rail_length`) | any | nominal + no-main + ballistic | pre-built once |
| any rocket component field | any | nominal only | rebuilt per combination |
| rocket components + flight parameters mixed | any | nominal only | rebuilt per combination |

**Rocket component sections** \
Varying any field here triggers a per-combo rebuild:
`motor`, `liquid_engine`, `rocket`, `nosecone`, `railbuttons`, `tailcone`, `fins`, `parachutes.main`, `parachutes.drogue`.
