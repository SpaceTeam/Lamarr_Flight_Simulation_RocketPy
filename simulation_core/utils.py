"""
Utility functions for the RocketPy simulation backend.
"""

import itertools
import math
import tomllib
from pathlib import Path
from IPython import get_ipython
from IPython.display import Markdown, display
from pydantic import BaseModel
from simulation_core.config_schema import Config, SimParams, ZonesConfig, accepts_variation


def printmd(string):
    """Print text as Markdown in a Jupyter notebook."""
    display(Markdown(string))


def _show_error_message_only(shell, exception_type, exception, traceback, tb_offset=None):
    """IPython error handler that shows only "Type: message" instead of the full traceback."""
    shell.showtraceback((exception_type, exception, traceback), exception_only=True)


def show_errors_without_traceback(*exception_types: type[BaseException]) -> None:
    """In a notebook, show the given errors as one line "Type: message"; the cell still fails, so the run stops there.

    Other errors keep their full traceback. Outside IPython (plain Python) nothing changes.
    """
    shell = get_ipython()
    if shell is not None:
        shell.set_custom_exc(exception_types, _show_error_message_only)

# =============================================================================
# Config combination helpers
# =============================================================================

def _is_numeric_list(value):
    """Return True when value is a non-empty list whose first element is a number.

    Used to distinguish variation lists ([0, 90, 180]) from structural lists
    (shape_points [[x,y], ...] or sources ["cats_vega", ...]).
    """
    return isinstance(value, list) and bool(value) and isinstance(value[0], (int, float))


def collect_nested_variations(config):
    """Collect all varied numeric fields, including fields nested one level inside sub-configs.

    Returns a dict keyed by path tuples:
    - (section, field) for top-level section fields, e.g. (flight, heading)
    - (section, subsection, field) for nested sub-config fields, e.g. (parachutes, payload, cd_s)
    Values are the list of variation values.

    Only variable fields (see accepts_variation in config_schema) are eligible — structural lists like
    fins.shape_points (list[list[float]]) are excluded.
    """
    varied = {}
    # model_fields is read from the class: reading it from an instance is deprecated since pydantic 2.11.
    for section_name in type(config).model_fields:
        section = getattr(config, section_name)
        if not isinstance(section, BaseModel):
            continue
        for field_name, field in type(section).model_fields.items():
            value = getattr(section, field_name)
            if _is_numeric_list(value) and accepts_variation(field):
                varied[(section_name, field_name)] = value
            elif isinstance(value, BaseModel):
                # Recurse one level into nested sub-configs (e.g. parachutes.main, parachutes.payload)
                for sub_field_name, sub_field in type(value).model_fields.items():
                    sub_value = getattr(value, sub_field_name)
                    if _is_numeric_list(sub_value) and accepts_variation(sub_field):
                        varied[(section_name, field_name, sub_field_name)] = sub_value
    return varied


def has_variations(config):
    """Return True when any config field carries multiple values (a list of scalars)."""
    return bool(collect_nested_variations(config))


def has_non_payload_variations(config):
    """Return True when any field other than payload.mass_total carries multiple values."""
    for path in collect_nested_variations(config):
        if path[:2] == ("payload", "mass_total"):
            continue
        return True
    return False


# Sections where variation does not require rebuilding engine/rocket.
_NON_ROCKET_SECTIONS = {"environment", "engine", "flight", "payload", "reanalysis"}


def has_rocket_component_variations(config):
    """Return True when any rocket-component field (motor, fins, nosecone, etc.) carries multiple values.

    parachutes.payload is excluded — it varies like a payload parameter, not a rocket component.
    """
    for path in collect_nested_variations(config):
        if path[0] in _NON_ROCKET_SECTIONS:
            continue
        # parachutes.payload behaves like a payload param; only main/drogue are rocket components
        if path[0] == "parachutes" and len(path) == 3 and path[1] == "payload":
            continue
        return True
    return False


def generate_config_combinations(config):
    """Yield one Config per combination of all varied fields across all config sections.

    Fields typed with FLOAT_RANGE_EXPANSION / INT_RANGE_EXPANSION accept a list of values to vary. Range strings are already expanded
    to lists by FLOAT_RANGE_EXPANSION / INT_RANGE_EXPANSION in config_schema. When nothing varies, yields
    the original config once.
    """
    varied = collect_nested_variations(config)

    if not varied:
        yield config
        return

    keys = list(varied.keys())
    for combo in itertools.product(*[varied[k] for k in keys]):
        values = dict(zip(keys, combo))

        # Group updates by section, separated into direct (depth-2) and nested (depth-3) updates
        by_section = {}
        for path, value in values.items():
            entry = by_section.setdefault(path[0], {"_direct": {}, "_nested": {}})
            if len(path) == 2:
                entry["_direct"][path[1]] = value
            else:
                entry["_nested"].setdefault(path[1], {})[path[2]] = value

        # Rebuild each affected section
        config_update = {}
        for section_name, updates in by_section.items():
            section = getattr(config, section_name)
            # Rebuild any varied sub-configs first
            for subsection_name, sub_updates in updates["_nested"].items():
                subsection = getattr(section, subsection_name)
                section = section.model_copy(update={subsection_name: subsection.model_copy(update=sub_updates)})
            # Then apply direct field updates on top
            if updates["_direct"]:
                section = section.model_copy(update=updates["_direct"])
            config_update[section_name] = section

        yield config.model_copy(update=config_update)


def count_config_combinations(config):
    """Count how many combinations generate_config_combinations will yield for this config."""
    count = 1
    for values in collect_nested_variations(config).values():
        count *= len(values)
    return count


def _diff_flat(base, combo):
    """Recursively collect leaf fields where base and combo differ, returning a flat dict."""
    result = {}
    for key, combo_val in combo.items():
        base_val = base.get(key)
        if isinstance(combo_val, dict) and isinstance(base_val, dict):
            result.update(_diff_flat(base_val, combo_val))
        elif base_val != combo_val:
            result[key] = combo_val
    return result


def build_variation_meta(base_config, combo_config):
    """Collect all leaf fields that changed between base config (with lists) and a combo config (single values)."""
    return _diff_flat(base_config.model_dump(), combo_config.model_dump())


# =============================================================================
# Parameter handling
# =============================================================================



def ensure_list(obj):
    """Normalize any value into a list so callers can iterate it without checking its type first.

    Lists pass through; tuples/sets are converted; scalars become [obj]; None becomes [].
    """
    if obj is None:
        return []
    if isinstance(obj, (list, tuple, set)):
        return list(obj)
    return [obj]


# =============================================================================
# Configuration loading
# =============================================================================

def load_config(config_path: Path):
    """Load and validate a TOML config file.

    Args:
        config_path: path to `config.toml` file.

    Returns:
        SimParams object with:
        - config: config.toml loaded into a typed `Config` class for structured parameter access
        - runtime: empty `RuntimeParams` class populated by create_* functions
        - project_path: the folder that contains config.toml (outputs like plots/ are written there)
        - export_kml: True when no variations are active
    """
    if not config_path.exists():
        raise FileNotFoundError(f"Config file not found: {config_path}")

    if config_path.suffix.lower() != ".toml":
        raise ValueError(f"Only TOML config files are supported: {config_path}")

    # tomllib requires the file opened in binary mode; range strings get expanded by the Config validators.
    with open(config_path, "rb") as config_file:
        config = Config.model_validate(tomllib.load(config_file))
    # The project folder is where config.toml lives, so projects can be stored anywhere (e.g. projects/ALBATROSS).
    project_path = config_path.parent
    export_kml = not has_variations(config)

    return SimParams(config=config, project_path=project_path, export_kml=export_kml)


# =============================================================================
# Zone loading
# =============================================================================

def scale_zones(zones: dict[str, list[tuple[float, float]]], scale_factor: float):
    """
    Scale polygon coordinates around each polygon centroid.
    
    - scale_factor > 1  -> enlarge
    - scale_factor = 1  -> unchanged
    - scale_factor < 1  -> shrink
    
    Args:
        zones: dict of {zone_name: [(x, y), ...]} in local Cartesian coordinates.
        scale_factor: how much to scale the zones (e.g. 1.1)
    """
    enlarged = {}

    for zone_name, coordinates in zones.items():
        # Calculate centroid (average of points)
        cx = sum(x for x, y in coordinates) / len(coordinates)
        cy = sum(y for x, y in coordinates) / len(coordinates)

        new_coords = []
        for x, y in coordinates:
            # Vector from center to point
            dx = x - cx
            dy = y - cy

            # Scale vector
            new_x = cx + dx * scale_factor
            new_y = cy + dy * scale_factor

            new_coords.append((new_x, new_y))

        enlarged[zone_name] = new_coords

    return enlarged


def _polar_to_cartesian(distance_m: float, heading_deg: float):
    """
    Convert a zone vertex from polar [distance, heading] to Cartesian (x, y).

    Args:
        distance_m: Distance from launch site to zone corner, in meters.
        heading_deg: Heading angle measured clockwise from north, in degrees; meaning 0° = North, 90° = East.
    """
    heading_rad = math.radians(heading_deg)
    return (distance_m * math.sin(heading_rad), distance_m * math.cos(heading_rad))


def load_zones(zones_path: Path):
    """
    Load polygon zones from a TOML file.

    The TOML file should contain the tables `exclusion_zones` and `buffer_zones`.
    Each zone is defined as polar points:

        `{name: [(distance_m, heading_deg), ...]}`

    - exclusion_zones: Zones where the rocket must not land.

    - buffer_zones: Safety margin zones around exclusion zones, or zones where the rocket should preferably not land,
        such as forests.

    - exclusion_zone_safety_margin: Factor used to enlarge exclusion zones when checking whether a flight is unsafe.

    We later mark a flight as unsafe if any trajectory (nominal, no_main, ballistic) enters a buffer zone.
    """
    if not zones_path.exists():
        return {}, {}, 1

    with open(zones_path, "rb") as zones_file:
        zones = ZonesConfig.model_validate(tomllib.load(zones_file))

    def convert(zone_dict):
        return {
            name: [_polar_to_cartesian(dist, heading) for dist, heading in points]
            for name, points in zone_dict.items()
        }

    exclusion_zones = convert(zones.exclusion_zones)
    buffer_zones = convert(zones.buffer_zones)
    return exclusion_zones, buffer_zones, zones.exclusion_zone_safety_margin


# =============================================================================
# Object metadata
# =============================================================================

def render_meta(meta, indent=0):
    """
    Recursively render a metadata dictionary as formatted text lines.

    Args:
        meta (dict): Metadata dictionary.
        indent (int): Number of leading spaces for indentation.

    Returns:
        list[str]: Formatted lines representing the metadata tree.
    """
    lines = []
    pad = " " * indent
    for k, v in meta.items():
        lines.append(f"{pad}|  {k}")
        if hasattr(v, "_meta"):
            lines.extend(render_meta(v._meta, indent + 3))
        else:
            lines[-1] += f" = {v}"
    return lines


def obj_to_pretty_label(obj, header):
    """Build a multi-line label for an object including its metadata."""
    if not hasattr(obj, "_meta"):
        return header
    return "\n".join([header] + render_meta(obj._meta))


def render_meta_flat(meta):
    """Render a metadata dictionary as a flat list of `key=value` strings for filenames."""
    parts = []
    for key, value in meta.items():
        if hasattr(value, "_meta"):
            # Nested object: include the key and recurse into its meta dict.
            parts.append(str(key))
            parts.extend(render_meta_flat(value._meta))
        else:
            parts.append(f"{key}={value}")
    return parts


def obj_to_filename_label(obj, header):
    """Build a flat, filesystem-safe identifier for a RocketPy object."""
    parts = [str(header)]
    if hasattr(obj, "_meta"):
        parts.extend(render_meta_flat(obj._meta))
    return "_".join(parts)


def set_labels(obj):
    """Set filename_label and name on a RocketPy object from its existing _meta (or no meta)."""
    try:
        original_name = obj.name if hasattr(obj, "name") else obj.__class__.__name__
        obj.filename_label = obj_to_filename_label(obj, header=original_name)
        obj.name = obj_to_pretty_label(obj, header=original_name)
    except Exception:
        pass


def tag_variation(obj, meta):
    """Stamp variation context (e.g. heading, inclination) onto a flight object, then set its labels."""
    obj._meta = meta
    set_labels(obj)


# =============================================================================
# File / path helpers
# =============================================================================

def get_project_file(params_or_path: SimParams | Path | str, file: str) -> str:
    """Resolve a file name to an absolute path, searching the project folder if needed.

    Accepts either a SimParams object or a plain Path / str as the first argument.
    """
    # Checks for the attribute instead of the class: after %autoreload the params object may be an instance of the old SimParams class.
    project_path = Path(getattr(params_or_path, "project_path", params_or_path))

    file_path = Path(file)
    if file_path.exists():
        return str(file_path)

    project_file_path = project_path / file
    if project_file_path.exists():
        return str(project_file_path)

    raise FileNotFoundError(f"Required project file not found: {file}")


def ensure_project_folders(project_path: Path):
    """Create the standard output folders used by the notebook and backend."""
    project_path.mkdir(parents=True, exist_ok=True)
    (project_path / "plots").mkdir(exist_ok=True)
    (project_path / "trajectory_kml").mkdir(exist_ok=True)
    (project_path / "trajectory_csv").mkdir(exist_ok=True)
    (project_path / "weather_csvs").mkdir(exist_ok=True)
    (project_path / "CATS_FLIGHT_DATA").mkdir(exist_ok=True)
    (project_path / "ALTIMAX_FLIGHT_DATA").mkdir(exist_ok=True)
    (project_path / "RCU_FLIGHT_DATA").mkdir(exist_ok=True)
