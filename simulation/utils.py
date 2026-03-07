import sys
import itertools
import numpy as np
import matplotlib.pyplot as plt

from math import prod, pi
from matplotlib.path import Path
from rocketpy import CompareFlights
from matplotlib.patches import Polygon
from IPython.display import Markdown, display

# print in markdown
def printmd(string):
    display(Markdown(string))


# returns a range of floats
def float_range(start, stop, step=1):
    while start <= stop:
        yield start
        start += step

# function to parse different types of parameters in config
def parse_value(value_str):
    """Parse a value into constant or list of values."""
    value_str = value_str.strip()

    # Tuple of tuples
    # TODO accept tuples of tuples as variables
    if "((" in value_str and "))" in value_str:
        value_str = value_str.strip("()")
        items = value_str.split("),(")
        parsed_items = []
        for item in items:
            item = item.strip("()")
            sub_items = [v.strip() for v in item.split(",")]
            sub_parsed_items = []
            for sub_item in sub_items:
                try:
                    n = float(sub_item)
                    sub_parsed_items.append(int(n) if n.is_integer() else n)
                except ValueError:
                    sub_parsed_items.append(sub_item)
            parsed_items.append(tuple(sub_parsed_items))
        return tuple(parsed_items)
    
    # Tuple of values
    # TODO accept tuples as variables
    if "(" in value_str and ")" in value_str:
        value_str = value_str.strip("()")
        items = [v.strip() for v in value_str.split(",")]
        parsed_items = []
        for item in items:
            try:
                n = float(item)
                parsed_items.append(int(n) if n.is_integer() else n)
            except ValueError:
                parsed_items.append(item)
        return tuple(parsed_items)
    
    # Range: a..b:c    # TODO: add comma separated ranges
    if ".." in value_str:
        value_str, step = value_str.split(":")
        start, end = map(float, value_str.split(".."))
        return list(float_range(start, end, float(step)))

    # Comma-separated variable values
    if "," in value_str:
        items = [v.strip() for v in value_str.split(",")]
        parsed_items = []
        for item in items:
            try:
                n = float(item)
                parsed_items.append(int(n) if n.is_integer() else n)
            except ValueError:
                if value_str == "False":
                    parsed_items.append(False)
                elif value_str == "True":
                    parsed_items.append(True)
                else:
                    parsed_items.append(item)
        return parsed_items

    # Otherwise a constant (single value)
    try:
        # Convert to float or int where possible
        if value_str == "False":
            return False
        if value_str == "True":
            return True
        n = float(value_str)
        return int(n) if n.is_integer() else n
    except ValueError:
        return value_str  # string constant


# main function for parsing the config file
def parse_config(path):
    constants = {}
    variables = {}

    with open(path) as f:
        for line in f:
            line = line.split("#")[0].strip()  # remove comments
            if not line:
                continue
            try:
                name, value_str = map(str.strip, line.split("=", 1))
            except ValueError:
                # terminate program with error message
                print(f"Error parsing line: '{line}'. Missing '=' sign.")
                sys.exit(100)
                
            parsed = parse_value(value_str)

            # Variable = if parsed is a list
            if isinstance(parsed, list):
                variables[name] = parsed
            else:
                constants[name] = parsed

    return constants, variables

# generates "objects" -> a unique combination, containing each constant and one instance of a variable
# early concept, but never used in code
def generate_objects(constants, variables):
    names = list(variables.keys())
    value_lists = [variables[n] for n in names]

    for combo in itertools.product(*value_lists):
        obj = dict(constants)
        obj.update(zip(names, combo))
        yield obj



# TODO: ISSUE: if one param is 0, function returns 1 combination possible, regardless of input size
def generate_combinations(params, constants, variables):
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
        params (list[str]):
            List of parameter names to generate combinations for.

    Yields:
        dict:
            A dictionary mapping each parameter name to a selected value.

    Raises:
        SystemExit(102):
            If any parameter is not found in either `constants`
            or `variables`.
    """
    for p in params:
        if p not in constants and p not in variables:
            print(f"Error: Parameter '{p}' not found in constants or variables.")
            sys.exit(102)

    var_params = [p for p in params if p in variables]
    if not var_params:  # if all params are constants
        yield {p: constants[p] for p in params}
        return

    value_lists = [variables[p] for p in var_params]

    for combo in itertools.product(*value_lists):
        values = {p: constants.get(p) for p in params}
        values.update(dict(zip(var_params, combo)))
        yield values


def register(name, objects, constants, variables):
    """
     Updates either the global `constants` or `variables` dictionary.

    A parameter is stored in:
    - `constants` if it has exactly one possible value.
    - `variables` if it has multiple possible values.

    Args:
        name (str): The parameter name.

        objects: A sequence of possible values for the parameter.
            - If length == 1 → stored as a constant.
            - If length > 1 → stored as a variable.
    """
    if isinstance(objects, (list, tuple)) and not isinstance(objects, str):
        if len(objects) == 1:
            constants[name] = objects[0]
        else:
            variables[name] = objects
    else:
        constants[name] = objects

    return constants, variables

def lookup(name, constants, variables):
    """
    Look up a parameter by name in constants or variables.

    Returns
    -------
    value : stored value
    location : str
        "constants" or "variables"

    Raises
    ------
    KeyError if name is not found.
    """
    if name in constants:
        return constants[name], "constants"

    if name in variables:
        return variables[name], "variables"

    raise KeyError(f"{name} not found in constants or variables")

def fill_parameters_exact(parameter_list, names, constants, variables):
    """
    Searches both the global `constants` and `variables`
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
        if name in constants:
            parameter_list.append(name)
        if name in variables:
            parameter_list.append(name)

def fill_parameters(parameter_list, prefixes, constants, variables):
    """
    Searches both the global `constants` and `variables`
    dictionaries and appends all parameter names that start with
    the given prefix (or any prefix in a list).

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
        parameter_list.extend(p for p in variables if p.startswith(prefix))

def check_required(parameter_list, required, kind = None):
    """
    Ensure all required parameters are present.

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
        required = [kind + "_" + item for item in required]
    missing = set(required) - set(parameter_list)
    if missing:
        raise ValueError(f"Missing required {kind} parameters: {missing}")


def add_defaults(parameter_list, defaults,  constants, variables, kind = None):
    """
    Add default constant parameters if they are not already defined.

    If `kind` is given, it is used as a prefix for each default name
    (format: "{kind}_{name}").

    Args:
        parameter_list (list[str]): Existing parameter names.
        defaults (dict[str, Any]): Default constant names(key) and values.
        kind (str | None): Optional prefix.

    Updates `constants` and extends `parameter_list` when needed.
    """
    for k, v in defaults.items():
        if kind is not None:
            k = kind + "_" + k
        if k not in constants and k not in variables:
            # if a default value is not already in constants or variables, we have to assume, that it also hasn't gotten a value yet
            constants[k] = v
            parameter_list.append(k)
 
    # fill -> check -> add

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


def obj_to_pretty_label(obj):
    """
    Build a multi-line label for an object including its metadata.

    Uses `obj.name` if available, otherwise the class name.
    """
    header = obj.name if hasattr(obj, "name") else obj.__class__.__name__
    if not hasattr(obj, "_meta"):
        return header
    return "\n".join([header] + render_meta(obj._meta))

def attach_meta(obj, meta):
    """
    Attach metadata to an object and update its display name.

    Args:
        obj: Target object.
        meta (dict): Metadata to attach.
    """
    obj._meta = meta
    obj.name = obj_to_pretty_label(obj)


def count_combinations(params, constants, variables):
    return prod(
        len(variables[p]) for p in params if p in variables
    ) or 1






def scale_zones(zones, scale_factor):
    """
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

# ===============================
# Function to plot polygons
# ===============================
def plot_zones(ax, zones, color='red', alpha=0.3):
    for name, coords in zones.items():
        polygon = Polygon(coords, closed=True, facecolor=color, edgecolor=color, alpha=alpha)
        ax.add_patch(polygon)
    # Adjust limits
    all_x = [x for coords in zones.values() for x, y in coords]
    all_y = [y for coords in zones.values() for x, y in coords]
    ax.set_xlim(min(all_x) - 50, max(all_x) + 50)
    ax.set_ylim(min(all_y) - 50, max(all_y) + 50)
    ax.set_aspect('equal', adjustable='box')
    ax.grid(True)

# ===============================
# Function to plot flight impacts
# ===============================
def plot_flights(ax, flight_groups):
    """
    flight_groups = {
        "rocket_nominal": (flights, "green"),
        "rocket_no_main": (flights, "orange"),
        "rocket_ballistic": (flights, "red"),
        "payload_nominal": (flights, "blue"),
        "payload_no_chute": (flights, "purple"),
    }
    """
    for label, (flights, color) in flight_groups.items():
        xs = [f.x_impact for f in flights]
        ys = [f.y_impact for f in flights]
        ax.scatter(xs, ys, color=color, label=label)

    ax.scatter(0, 0, color="black", marker="x", label="launch rail")
    ax.legend()



def plot_safe_flights(project, exclusion_zones, buffer_zones,
                      flight_groups, heading=None, inclination=None):

    filtered_groups = {}

    for name, (flights, color) in flight_groups.items():
        filtered = flights

        if heading is not None:
            filtered = [f for f in filtered if f.heading == heading]

        if inclination is not None:
            filtered = [f for f in filtered if f.inclination == inclination]

        filtered_groups[name] = (filtered, color)

    plot_name = "safe_flights"
    if heading is not None:
        plot_name += f"_heading_{heading}"
    if inclination is not None:
        plot_name += f"_inclination_{inclination}"

    fig, ax = plt.subplots()

    if exclusion_zones:
        plot_zones(ax, exclusion_zones, color="red")
    if buffer_zones:
        plot_zones(ax, buffer_zones, color="orange")

    if flight_groups:
        plot_flights(ax, filtered_groups)
    
    fig.savefig(f"{project}/plots/{name}.png", dpi=300, bbox_inches="tight")
    #plt.show()

def is_in_exclusion_zone(coords, zones_dict):
    """
    Check which points are inside any of the exclusion zones.

    Parameters
    ----------
    coords : array-like of (x, y) tuples
        [(x1, y1), (x2, y2), ...]
    zones_dict : dict
        {"zone_name": [(x1,y1), (x2,y2), ...], ...}

    Returns
    -------
    inside_mask : np.ndarray (bool)
        True if point is inside ANY zone
    """
    points = np.asarray(coords)  # shape (N, 2)
    inside_mask = np.zeros(len(points), dtype=bool)

    for zone_coords in zones_dict.values():
        path = Path(zone_coords)
        inside_mask |= path.contains_points(points)

    return inside_mask


def ensure_list(obj):
    if obj is None:
        return []
    if isinstance(obj, (list, tuple, set)):
        return list(obj)
    return [obj]

def get_unsafe_headings(flights, zones):
    coords  = [(f.x_impact, f.y_impact) for f in flights]

    # ===============================
    # Check which impacts are in exclusion zones
    # ===============================
    impacts  = is_in_exclusion_zone(coords, zones)

    # ===============================
    # Find headings of rockets in exclusion zones
    # ===============================
    headings_in_zone = list({flights[i].heading 
                            for i, in_zone in enumerate(impacts) if in_zone})

    return headings_in_zone




def draw_initial_solutions(constants, variables):
    legend = False
    all_flights = lookup("all_flights", constants, variables)[0]

    #for flight in all_flights:
    #    print(flight.name)
        # print(flight._meta)
    comparison_normal = CompareFlights(all_flights)
    comparison_normal.trajectories_3d(legend=legend, filename = "3d.png")
    comparison_normal.trajectories_2d(legend=legend, filename = "2d_xy.png", plane = "xy")
    comparison_normal.trajectories_2d(legend=legend, filename = "2d_xz.png", plane = "xz")
    comparison_normal.trajectories_2d(legend=legend, filename = "2d_yz.png", plane = "yz")


