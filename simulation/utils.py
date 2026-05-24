"""
Utility functions for the RocketPy simulation backend.
"""

import itertools
import json
import math
from pathlib import Path
from IPython.display import Markdown, display


# keys in the json, that can be varied
VARIATION_KEYS_WITHOUT_FLIGHT_PLOTS = {"flight_heading", "flight_inclination", "payload_mass", "parachutes_payload_cd_s"}
# still show flight plots for the following variation keys
VARIATION_KEYS_WITH_FLIGHT_PLOTS = set()

VARIATION_KEYS = VARIATION_KEYS_WITHOUT_FLIGHT_PLOTS.union(VARIATION_KEYS_WITH_FLIGHT_PLOTS)

def printmd(string):
    """
    Print text as Markdown.
    """
    display(Markdown(string))


# =============================================================================
# Parameter handling
# =============================================================================

def generate_combinations(parameters: list[str], constants, variations):
    """
    Generate all possible value combinations for a given list of parameters.

    Each parameter name in `params` must exist in either the
    `constants` dictionary or the `variables` dictionary.

    The function yields dictionaries mapping each parameter name to
    one specific value combination.

    Example:
        If:
            variables = {
                "param0": [1, 2, 3, 4, 5],
                "param1": ["A", "B"]
            }
            constants = {"param2": 1}

        Then:
            generate_combinations(["param0", "param1", "param2"])

        Yields 10 dictionaries (5 x 2 x 1 combinations).

    Args:
        parameters (list[str]):
            List of parameter names to generate combinations for.

    Yields:
        dict:
            A dictionary mapping each parameter name to a selected value.

    Raises:
        KeyError:
            If any parameter is not found in either `constants` or `variables`.
    """
    for p in parameters:
        if p not in constants and p not in variations:
            raise KeyError(f"Parameter '{p}' not found in constants or variations.")

    var_parameters = [p for p in parameters if p in variations]

    if not var_parameters:
        # if all parameters are constants
        yield {p: constants[p] for p in parameters}
        return

    value_lists = [variations[p] for p in var_parameters]

    for combo in itertools.product(*value_lists):
        values = {p: constants.get(p) for p in parameters}
        values.update(dict(zip(var_parameters, combo)))
        yield values


def lookup(name, constants, variations):
    """
    Look up one parameter by name in constants or variations.

    Returns
    -------
    value : stored value
    location : str
        "constants" or "variations"

    Raises
    ------
    KeyError if name is not found.
    """
    if name in constants:
        return constants[name], "constants"

    if name in variations:
        return variations[name], "variations"

    raise KeyError(f"{name} not found in constants or variations")


def require_value(values, key, context):
    """
    Raise an error if a required key is missing from a dictionary.
    """
    if key not in values:
        raise KeyError(f"Missing required config key for {context}: {key}")


def fill_parameters_exact(parameter_list: list[str], names: str | list[str], constants, variations):
    """
    Searches both the global `constants` and `variations` 
    dictionaries and appends only exact parameter name matches.

    Args:
        parameter_list (list[str]):
            The list that matching parameter names will be added to.

        names (str | list[str]):
            A parameter name or list of parameter names to match exactly.
    """
    # Normalize to list
    if isinstance(names, str):
        names = [names]

    for name in names:
        if name in constants or name in variations:
            parameter_list.append(name)


def fill_parameters(parameter_list: list[str], prefixes: str | list[str], constants, variations):
    """
    Searches both the global `constants` and `variations` 
    dictionaries and appends all parameter names that start with
    that start with one of the supplied prefixes.

    Args:
        parameter_list (list[str]):
            The list that matching parameter names will be added to.

        prefixes (str | list[str]):
            A prefix string or a list of prefix strings.
            Any parameter whose name starts with one of these
            will be appended.
    """
    # Normalize to list
    if isinstance(prefixes, str):
        prefixes = [prefixes]

    for prefix in prefixes:
        parameter_list.extend(p for p in constants if p.startswith(prefix))
        parameter_list.extend(p for p in variations if p.startswith(prefix))


def check_required(parameter_list: list[str], required: list[str], kind: str | None = None):
    """
    Raise an error when required parameters are missing.

    If `kind` is given, it is used as a prefix for each required name
    (format: "{kind}_{name}").

    Args:
        parameter_list (list[str]): Existing parameter names.
        required (list[str]): Required parameter names (without prefix).
        kind (str | None): Optional prefix.

    Raises:
        ValueError: If any required parameter is missing.
    """
    if kind is not None:
        required = [f"{kind}_{item}" for item in required]

    missing = set(required) - set(parameter_list)

    if missing:
        raise ValueError(f"Missing required {kind} parameters: {missing}")


# def add_defaults(parameter_list, defaults, constants, variations, kind=None):
#     """
#     Add default constant parameters if they are not already defined.
#
#     If `kind` is given, it is used as a prefix for each default name
#      (format: "{kind}_{name}").
#     Args:
#         parameter_list (list[str]): Existing parameter names.
#         defaults (dict[str, Any]): Default constant names(key) and values.
#         kind (str | None): Optional prefix.

#     Updates `constants` and extends `parameter_list` when needed.
#     """
#     for key, value in defaults.items():
#         if kind is not None:
#             key = f"{kind}_{key}"
#
#         if key not in constants and key not in variations:
#             # if a default value is not already in constants or variations, we have to assume, that it also hasn't gotten a value yet
#             constants[key] = value
#             parameter_list.append(key)


def count_combinations(parameters, variations):
    """
    Count how many combinations a parameter list will generate.
    """
    variable_lengths = [len(variations[p]) for p in parameters if p in variations]

    if not variable_lengths:
        return 1

    return math.prod(variable_lengths)


def ensure_list(obj):
    """
    Normalize any value into a list so callers can iterate it without checking its type first.
    Lists pass through; tuples/sets are converted; scalars become `[obj]`; `None` becomes `[]`.
    """
    if obj is None:
        return []

    if isinstance(obj, (list, tuple, set)):
        return list(obj)

    return [obj]


# =============================================================================
# Configuration loading
# =============================================================================
def register(name, objects, constants, variations):
    """
    Updates either the global `constants` or `variations` dictionary.

    A parameter is stored in:
    - `constants` if it has exactly one possible value.
    - `variations` if it has multiple possible values.

    It removes the parameter of the other dict as safety, if it is called twice.
    
    Args:
        name (str): The parameter name.

        objects: A sequence of possible values for the parameter.
            - If length == 1 → stored as a constant.
            - If length > 1 → stored as a variable.
    """
    if isinstance(objects, (list, tuple)) and not isinstance(objects, str):
        if len(objects) == 1:
            constants[name] = objects[0]
            variations.pop(name, None)
        else:
            variations[name] = list(objects)
            constants.pop(name, None)
    else:
        constants[name] = objects
        variations.pop(name, None)

    return constants, variations


def convert_constant(value):
    """
    Recursively prepare a JSON value for storage as a constant: JSON arrays become tuples (immutable, RocketPy-friendly)
    and `_`-prefixed keys are stripped from any dict found inside.
    """
    if isinstance(value, list):
        # JSON array
        return tuple(convert_constant(item) for item in value)

    if isinstance(value, dict):
        return {
            key: convert_constant(item)
            for key, item in value.items()
            if not key.startswith("_")      # also skip comments in nested dicts
        }

    return value


def expand_string_range(value: str, flat_key):
    """
    Convert a range string start..stop:step to a list.

    If start <= stop: linear range, endpoint inclusive.
    Only for "flight_heading": if start > stop: wraps around 360.

    Examples:
        "86..90:2"   -> [86, 88, 90]
        "180..90:10" -> [180, 190, ..., 350, 0, 10, ..., 90]
    """
    if ".." not in value or ":" not in value:
        return None

    try:
        range_part, step = value.split(":", 1)
        start, stop = range_part.split("..", 1)
        start_value = float(start.strip())
        stop_value = float(stop.strip())
        step_value = float(step.strip())
    except ValueError as error:
        raise ValueError(f"Invalid range string '{value}'. Use the format start..stop:step.") from error

    if step_value <= 0:
        raise ValueError("Range step has to be positive.")

    wrapping = start_value > stop_value

    if wrapping and flat_key == "flight_heading":
        # total arc going clockwise past 360° back to stop
        total_span = 360 - start_value + stop_value
    elif wrapping and flat_key != "flight_heading":
        raise ValueError(f"The range of {flat_key} is not valid: {value}")
    else:
        total_span = stop_value - start_value

    # small 1e-9 so floor() doesn't drop the endpoint due to float rounding error
    num_steps = math.floor(total_span / step_value + 1e-9)

    values = []

    for i in range(num_steps + 1):
        current_value = round(start_value + i * step_value, 10)

        if wrapping:
            current_value = current_value % 360

        # collapse integers: 2.0000000000 -> 2
        if isinstance(current_value, float) and current_value.is_integer():
            current_value = int(current_value)

        values.append(current_value)

    return values


def walk_config(config_data, prefix=""):
    """
    Recursively flatten nested JSON keys into the existing backend key format.

    Example:

    ```json
    "parachutes": {
        "main": {
            "cd": 1.8
        }
    }
    ```

    becomes: `parachutes_main_cd`
    """
    constants = {}
    variations = {}

    for key, value in config_data.items():
        if key.startswith("_"):
            # ignore "commented out" entries
            continue

        flat_key = f"{prefix}_{key}" if prefix else key       

        if flat_key in VARIATION_KEYS and isinstance(value, list):
            # only treat lists as variation for set keys
            variations[flat_key] = value
            continue

        if flat_key in VARIATION_KEYS and isinstance(value, str):
            # we expect "start..stop:step"
            range_values = expand_string_range(value, flat_key)

            if range_values is not None:
                variations[flat_key] = range_values
                continue

        if isinstance(value, dict):
            # e.g. for parachutes
            child_constants, child_variations = walk_config(value, flat_key)
            constants.update(child_constants)
            variations.update(child_variations)
            continue

        constants[flat_key] = convert_constant(value)

    return constants, variations


def load_config(config_path: Path):
    """
    Load a JSON config file.
    
    Input: path to `config.json` file.

    The function fails if the config file does not exist.
    """

    if not config_path.exists():
        raise FileNotFoundError(f"Config file not found: {config_path}")

    if config_path.suffix.lower() != ".json":
        raise ValueError(f"Only JSON config files are supported: {config_path}")

    # load and parse config
    with open(config_path, encoding="utf-8") as config_file:
        data = json.load(config_file)

    constants, variations = walk_config(data)

    if variations:
        # do not create kml files if we vary flights
        export_kml = False
    else:
        export_kml = True
    
    return constants, variations, export_kml


# =============================================================================
# Zone loading
# =============================================================================

def scale_zones(zones, scale_factor):
    """
    Scale polygon coordinates around each polygon centroid.
    
    scale_factor > 1  -> enlarge
    scale_factor = 1  -> unchanged
    scale_factor < 1  -> shrink
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


def _polar_to_cartesian(distance_m, heading_deg):
    """
    Convert a zone vertex from polar [distance, heading] to Cartesian (x, y).

    Args:
        distance_m (float): Distance from launch site to zone corner, in meters.
        heading_deg (float): Heading angle measured clockwise from north, in degrees; meaning 0° = North, 90° = East.
    """
    heading_rad = math.radians(heading_deg)
    return (distance_m * math.sin(heading_rad), distance_m * math.cos(heading_rad))


def load_zones(zones_path: Path):
    """
    Load polygon zones from a JSON file with `exclusion_zones` and `buffer_zones` top-level keys into two dicts
    containing the name as key and the polygon points as values:
        {name: [(distance_m, heading_deg), ...]}

    - exclusion_zones: where we cannot land
    - buffer_zones: saftey margin around exclusion zones and zones where we rather should not land (e.g. forest)

    We later mark a flight as unsafe if any trajectory (nominal, no_main, ballistic) enters a buffer zone.
    """
    if not zones_path.exists():
        return {}, {}

    with open(zones_path) as f:
        data = json.load(f)

    def convert(zone_dict):
        return {
            name: [_polar_to_cartesian(dist, heading) for dist, heading in points]
            for name, points in zone_dict.items()
        }

    exclusion_zones = convert(data.get("exclusion_zones", {}))
    buffer_zones = convert(data.get("buffer_zones", {}))
    exclusion_zone_saftey_margin = data.get("exclusion_zone_saftey_margin", 1)
    return exclusion_zones, buffer_zones, exclusion_zone_saftey_margin


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
    """
    Build a multi-line label for an object including its metadata.
    """
    if not hasattr(obj, "_meta"):
        return header

    return "\n".join([header] + render_meta(obj._meta))


def render_meta_flat(meta):
    """
    Render a metadata dictionary as a flat list of `key=value` strings for filenames.
    """
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
    """
    Build a flat, filesystem-safe identifier for a RocketPy object.
    """
    parts = [str(header)]
    if hasattr(obj, "_meta"):
        parts.extend(render_meta_flat(obj._meta))

    return "_".join(parts)


def attach_meta(obj, meta):
    """
    Attach metadata to an object and update both its display name and its filename label.

    Args:
        obj: Target object.
        meta (dict): Metadata to attach.
    """
    obj._meta = meta

    try:
        original_name = obj.name if hasattr(obj, "name") else obj.__class__.__name__
        obj.filename_label = obj_to_filename_label(obj, header=original_name)
        obj.name = obj_to_pretty_label(obj, header=original_name)
    except Exception:
        pass


# =============================================================================
# File/path helpers
# =============================================================================
def get_project_file(constants, file: str):
    """
    Resolve a required file name relative to the configured project folder.
    
    Args:
        file:
            If it is a path we take this. If it is a file name we search in the project folder.
    """
    file_path = Path(file)

    if file_path.exists():
        return str(file_path)

    project_file_path = Path(constants["project"]) / file

    if project_file_path.exists():
        return str(project_file_path)

    raise FileNotFoundError(f"Required project file not found: {file}")


def ensure_project_folders(project_path: Path):
    """
    Create the standard output folders used by the notebook and backend.
    """
    project_path.mkdir(parents=True, exist_ok=True)
    (project_path / "plots").mkdir(exist_ok=True)
    (project_path / "trajectory_kml").mkdir(exist_ok=True)
    (project_path / "trajectory_csv").mkdir(exist_ok=True)
    (project_path / "weather_csvs").mkdir(exist_ok=True)
