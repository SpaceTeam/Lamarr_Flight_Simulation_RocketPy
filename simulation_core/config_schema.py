"""
Pydantic models for the project config files config.toml and zones.toml: the definition of which keys and values are allowed.

- Each TOML table maps to a model class; unknown keys are rejected, so typos are causing an error.
- Fields typed with FLOAT_RANGE_EXPANSION / INT_RANGE_EXPANSION accept a single value, a list, or a range string ("start..stop:step");
  range strings are expanded to lists while validating, and their description gets VARIATION_NOTE (hover help).
- Field descriptions and examples are written once in the class docstrings (see _Base) and used for pydantic, 
  the JSON schemas and the app's help texts.

Run `python simulation_core/config_schema.py` after changing the models to regenerate simulation_core/config_schemas/ for Tombi.
See README.md, "Config files", for how the models, schemas and editor fit together.
"""

from __future__ import annotations
import dataclasses
import functools
import inspect
import json
import math
import textwrap
import tomllib
from pathlib import Path
from typing import Annotated, Union, Optional, Any, Literal, get_args
from pydantic import BaseModel, BeforeValidator, ConfigDict, Field, NaiveDatetime, ValidationInfo, WithJsonSchema, model_validator
from pydantic.fields import FieldInfo

SCHEMA_DIR = Path(__file__).resolve().parent / "config_schemas"
CONFIG_SCHEMA_PATH = SCHEMA_DIR / "config.schema.json"
ZONES_SCHEMA_PATH = SCHEMA_DIR / "zones.schema.json"
# Appended to the description (Tombi hover, app help text) of every field that accepts a variation.
VARIATION_NOTE = 'Variable: also accepts a list `[a, b, c]` or a range string `"start..stop:step"`; one flight per value.'
# Lines in a docstring "Attributes" section that hold an example as TOML value instead of description text, e.g. "- Example: -40".
EXAMPLE_LINE_PREFIX = "- Example:"
# Matches "start..stop:step", e.g. "84..88:2" or "0.05..0.08:0.01".
RANGE_STRING_PATTERN = r"^\s*-?\d+(\.\d+)?\s*\.\.\s*-?\d+(\.\d+)?\s*:\s*\d+(\.\d+)?\s*$"
# JSON schema type names of the Python values of the "- Example:" lines; an int example also fits a "number" alternative.
JSON_SCHEMA_TYPES_BY_PYTHON_TYPE = {
    str: {"string"},
    int: {"integer", "number"},
    float: {"number"},
    bool: {"boolean"},
    list: {"array"},
    dict: {"object"},
}
# Config fields that are varied pairwise and thus must have the same number of values.
PAIRED_VARIATIONS = [
    (("rocket", "total_mass_without_motor"), ("rocket", "total_CG_without_motor_from_tip")),
]


# =============================================================================
# Allowed string values
# =============================================================================
# Literal types show up as autocomplete choices in the TOML editor, and any other value is rejected when loading.
ScenarioName = Literal["nominal", "no_main", "ballistic"]
EngineType = Literal["solid", "liquid"]         # hybrid not supported by this codebase yet
EnvironmentType = Literal["standard_atmosphere", "standard_atmosphere_wind", "Windy", "custom_atmosphere", "reanalysis", "reanalysis_custom"]
WindyModel = Literal["ECMWF", "GFS", "ICON", "ICONEU"]
# RocketPy also knows "powerseries", but that needs a power parameter this config does not provide.
NoseconeKind = Literal["conical", "ogive", "tangent", "lvhaack", "vonkarman", "elliptical", "parabolic"]
FlightComputerSource = Literal["cats_vega", "altimax", "rcu"]
MatchParameter = Literal["heading", "inclination", "velocity"]


# =============================================================================
# Range strings
# =============================================================================

def expand_string_range(value: str, key: str = ""):
    """Convert a 'start..stop:step' range string to a list of values.

    Returns None when the string is not in range format. start > stop is only allowed
    for the 'heading' key, where it wraps around 360° (e.g. '350..10:5' → 350, 355, 0, 5, 10).
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

    # wrap around 360 only for heading; for other keys start must not exceed stop
    wrapping = start_value > stop_value
    if wrapping and key != "heading":
        raise ValueError(
            f"Range start ({start_value}) must not exceed stop ({stop_value}). "
            "Wrap-around is only supported for 'heading'."
        )

    total_span = (360 - start_value + stop_value) if wrapping else (stop_value - start_value)
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


def _expand_range_value(value: Any, info: ValidationInfo) -> Any:
    """Recursively convert a range string into a list of values before pydantic checks the number type of them; 
    other values pass through."""
    if isinstance(value, str):
        # The field name is passed along so only 'heading' may wrap around 360°.
        expanded = expand_string_range(value, info.field_name or "")
        if expanded is not None:
            return expanded
    return value


# Variable numeric fields: Annotated[float | list[float], FLOAT_RANGE_EXPANSION] (or the int variant) accepts a single value,
# a list, or a range string. 
_RangeString = Annotated[str, Field(pattern=RANGE_STRING_PATTERN)]
FLOAT_RANGE_EXPANSION = BeforeValidator(_expand_range_value, json_schema_input_type=Union[float, list[float], _RangeString])
INT_RANGE_EXPANSION = BeforeValidator(_expand_range_value, json_schema_input_type=Union[int, list[int], _RangeString])


# Define a local date-time type (timestamp without a timezone offset; e.g. 2026-05-23T13:12:00 instead of 
# 2026-05-23T13:12:00+02:00), because the timezone has its own config field.
_LocalDateTime = Annotated[NaiveDatetime, WithJsonSchema({"type": "string", "format": "date-time-local"})]


@dataclasses.dataclass
class AttributeDocumentation:
    """Field definition read by the class "Attributes" section."""
    description: str
    examples: list[Any] = dataclasses.field(default_factory=list)


def parse_toml_value(text: str, field_name: str) -> Any:
    """Read a value written in TOML syntax, e.g. '"K700W"', '-40' or '[0, 90]', as the matching Python value."""
    try:
        return tomllib.loads(f"value = {text}")["value"]
    except tomllib.TOMLDecodeError as error:
        raise ValueError(f"'{field_name}' has no valid TOML value in '{text}'") from error


def parse_attributes_section(docstring: str) -> dict[str, AttributeDocumentation]:
    """Parse the "Attributes" section of a class docstring and create an AttributeDocumentation object for field 
    containing the field's description text and its "- Example:". 

    Returns {field name: AttributeDocumentation}, or {} if the docstring has no such section (format: see _Base).
    """
    doctring_lines = inspect.cleandoc(docstring).splitlines()
    # The section starts below the "Attributes" heading and its "----------" underline.
    heading_index = next(
        (index for index in range(len(doctring_lines) - 1) 
         if doctring_lines[index] == "Attributes" and set(doctring_lines[index + 1]) == {"-"}), None
    )
    if heading_index is None:
        return {}

    description_lines_by_name: dict[str, list[str]] = {}
    for line in doctring_lines[heading_index + 2:]:
        if line and not line[0].isspace():
            # "name : type" -> "name"
            current_name = line.split(":")[0].strip()
            description_lines_by_name[current_name] = []
        elif description_lines_by_name:
            description_lines_by_name[current_name].append(line)
    documentation_by_name = {}
    for name, attribute_lines in description_lines_by_name.items():
        documentation = AttributeDocumentation(description="")
        description_lines = []
        # dedent removes the common indentation, so indented list lines inside a description keep their extra indent.
        for line in textwrap.dedent("\n".join(attribute_lines)).splitlines():
            if line.startswith(EXAMPLE_LINE_PREFIX):
                documentation.examples.append(parse_toml_value(line.removeprefix(EXAMPLE_LINE_PREFIX).strip(), name))
            else:
                description_lines.append(line)
        documentation.description = "\n".join(description_lines).strip()
        documentation_by_name[name] = documentation
    return documentation_by_name


def docstring_summary(docstring: str) -> str:
    """Return a class docstring without its "Attributes" section, for places that already show each field's description separately."""
    return inspect.cleandoc(docstring).split("\nAttributes\n")[0].rstrip()


def _schema_description_without_attributes(schema: dict[str, Any]) -> None:
    """Cut the "Attributes" section from a model's schema description, so Tombi's hover on a [section] shows only the summary."""
    if "description" in schema:
        schema["description"] = docstring_summary(schema["description"])


def accepts_variation(model_field: FieldInfo) -> bool:
    """Return True if the field is typed with FLOAT_RANGE_EXPANSION or INT_RANGE_EXPANSION, i.e. accepts a list or range string."""
    # Annotated[...] keeps its validator in metadata; Optional[Annotated[...]] keeps it nested inside the annotation.
    nested_metadata = [metadata for option in get_args(model_field.annotation) for metadata in getattr(option, "__metadata__", ())]
    return any(metadata in (FLOAT_RANGE_EXPANSION, INT_RANGE_EXPANSION) for metadata in [*model_field.metadata, *nested_metadata])

# =============================================================================
# Config classes - see _Base for how to document fields
# =============================================================================

class _Base(BaseModel):
    """Pydantic base class: rejects unknown keys (a typo in the TOML file raises an error) and reads the fields' documentation
    from the class docstring.

    Each field's description and examples are written in the "Attributes" section of the class docstring, so the IDE
    shows them all when hovering the class, and they end up in the schema (Tombi hover) and the Streamlit form (help text).
    Defaults are written in the code as usual (= value).

    How to write the "Attributes" section:
    - End the docstring with an "Attributes" heading, underlined with dashes. Everything below it belongs to the section.
    - Write each field name on its own line at the section's indentation ("name" or "name : type"; the type is ignored).
    - Indent the description lines below the name; they may span several lines and contain "- " bullet lists.
    - Add an example as an indented "- Example: <TOML value>" line, one line per example. Write strings with quotes.

    Example docstring:

        Launch site configuration.

        Attributes
        ----------
        timezone
            IANA timezone of the launch site.
            - Example: "Europe/Berlin"
        elevation : float
            Launch site elevation above sea level [m ASL].
            - Example: 0
            - Example: 1400.5

    Importing the module raises an error for a misspelled field name, an example that is not valid TOML,
    or examples written both in the docstring and in Field(examples=...).
    """
    model_config = ConfigDict(extra="forbid", json_schema_extra=_schema_description_without_attributes)

    @classmethod
    def __pydantic_init_subclass__(cls, **kwargs: Any) -> None:
        """Apply the description and examples of each field from the class docstring's "Attributes" section."""
        super().__pydantic_init_subclass__(**kwargs)
        try:
            documentation_by_name = parse_attributes_section(cls.__doc__ or "")
        except ValueError as error:
            raise TypeError(f"{cls.__name__} docstring: {error}") from error
        # A misspelled name in the section would otherwise silently leave its field without a description.
        unknown_names = set(documentation_by_name) - set(cls.model_fields)
        if unknown_names:
            raise TypeError(f"{cls.__name__} docstring describes unknown fields: {sorted(unknown_names)}")

        for name, model_field in cls.model_fields.items():
            documentation = documentation_by_name.get(name)
            if documentation:
                model_field.description = documentation.description
                # Examples must be defined in one place only, so both places at once is an error.
                if documentation.examples:
                    if model_field.examples:
                        raise TypeError(f"{cls.__name__}.{name} has examples both in the docstring and in Field(examples=...)")
                    model_field.examples = documentation.examples
            # The type already says whether a field can be varied, so the note is added here instead of in every docstring.
            if accepts_variation(model_field):
                model_field.description = f"{model_field.description or ''}\n\n{VARIATION_NOTE}".lstrip()
        # The schema was built when the class was created, so rebuild it to include the docstring values and variation notes.
        cls.model_rebuild(force=True)


# -----------------------------------------------------------------------------
# Environment
# -----------------------------------------------------------------------------

class EnvironmentConfig(_Base):
    """Launch site environment and atmosphere model configuration.

    Attributes
    ----------
    date
        Launch date and time in launch-site time, e.g. 2026-05-23T13:12:00. Or the magic string 'tomorrow_08_local' 
        for tomorrow at 08:00.
    latitude
        Launch site latitude [°].
    longitude
        Launch site longitude [°].
    max_expected_height
        Maximum expected height above ground level [m AGL].
    elevation
        Launch site elevation above sea level [m ASL].
    timezone
        IANA timezone of the launch site. Applied to `date`.
        - Example: "Europe/Berlin"
    envType
        Atmosphere model(s); a list simulates several environments. Supported values:
        - `standard_atmosphere`: RocketPy standard atmosphere (no wind).
        - `standard_atmosphere_wind`: RocketPy standard atmosphere with custom average wind speeds and wind direction 90° (wind coming from East). 
            Requires `standard_atmosphere_wind_speeds`.
        - `Windy`: RocketPy Windy forecast. Requires `windy_weather_models`.
        - `custom_atmosphere`: CSV files generated by the weather export script. Requires `custom_weather_models`.
        - `reanalysis`: RocketPy reanalysis file. Can be used as offline mode. Requires `reanalysis_file` and `reanalysis_dictionary`.
        - `reanalysis_custom`: existing historical CSV files. Can be used as offline mode. Requires `reanalysis_csv`.

        `reanalysis` and `reanalysis_custom` must be the only entry in the list.
    reanalysis_csv
        CSV filenames in the project folder keyed by weather model name, used with 'reanalysis_custom'.
        - Example: { icon_d2 = "rocketpy_atmosphere_icon_d2_reanalysis.csv" }
    standard_atmosphere_wind_speeds
        Average wind speeds for the standard_atmosphere_wind environment [m/s]; the wind comes from the East (90°)
        and varies over height by 10 % (OpenRocket's default turbulence intensity). One value is used for the whole altitude range.
        Multiple values simulate several environments, one per value.
        - Example: [0, 2, 4]
    windy_weather_models
        Windy weather model(s), used with 'Windy'.
    custom_weather_models
        Custom weather model name(s), used with 'custom_atmosphere'.
        - Example: ["icon_d2", "icon_eu"]
    reanalysis_file
        Reanalysis data file in the project folder, used with 'reanalysis'.
        - Example: "cansat_weather.nc"
    reanalysis_dictionary
        RocketPy reanalysis dictionary, used with 'reanalysis'.
        - Example: "ECMWF"
    """
    date: Union[Literal["tomorrow_08_local"], _LocalDateTime]
    latitude: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    longitude: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    max_expected_height: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    elevation: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    timezone: str
    envType: Union[EnvironmentType, list[EnvironmentType]]
    reanalysis_csv: Optional[dict[str, str]] = None
    standard_atmosphere_wind_speeds: Optional[float | list[float]] = None
    windy_weather_models: Optional[Union[WindyModel, list[WindyModel]]] = None
    custom_weather_models: Optional[Union[str, list[str]]] = None
    reanalysis_file: Optional[str] = None
    reanalysis_dictionary: Optional[Any] = None


# -----------------------------------------------------------------------------
# Engine
# -----------------------------------------------------------------------------

class EngineConfig(_Base):
    """Motor propulsion type selection.

    Attributes
    ----------
    type
        'solid' (uses the [motor] section) or 'liquid' (uses the [liquid_engine] section).
    """
    type: EngineType


class LiquidTankConfig(_Base):
    """Propellant (fuel or oxidizer) tank geometry, flow rate, and fluid properties.

    Attributes
    ----------
    length
        Tank length [mm].
    outer_diameter
        Tank outer diameter [mm].
    volume
        Propellant volume [L].
    massflow
        Propellant mass flow rate [kg/s].
    center_from_tip
        Distance of the tank's middle point to the nose tip [mm].
    substance
        CoolProp fluid name for this propellant.
        - Example: "ethanol"
        - Example: "oxygen"
    temperature
        Propellant temperature [K].
    pressure
        Tank pressure [Pa].
    """
    length: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    outer_diameter: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    volume: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    massflow: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    center_from_tip: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    substance: str
    temperature: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    pressure: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]


class PressureGasTankConfig(_Base):
    """Pressurant gas tank. Feeds both the fuel and oxidizer tanks from separate positions.

    Attributes
    ----------
    length
        Tank length [mm].
    outer_diameter
        Tank outer diameter [mm].
    volume
        Gas volume [L].
    massflow
        Pressurant mass flow rate [kg/s].
    fuel_center_from_tip
        Distance of the tank's middle point to the nose tip [mm]. This is the pressurant tank on the fuel side.
    oxidizer_center_from_tip
        Distance of the tank's middle point to the nose tip [mm]. This is the pressurant tank on the oxidizer side.
    substance
        CoolProp fluid name for the pressurant.
        - Example: "N2"
    temperature
        Gas temperature [K].
    pressure
        Tank pressure [Pa].
    """
    length: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    outer_diameter: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    volume: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    massflow: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    fuel_center_from_tip: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    oxidizer_center_from_tip: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    substance: str
    temperature: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    pressure: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]


class LiquidEngineConfig(_Base):
    """Liquid bipropellant motor configuration (pressure-fed, propellant-agnostic). Required when engine.type is 'liquid'.

    Attributes
    ----------
    holddown_time
        Time the rocket is held down after ignition before release [s].
    nozzle_diameter
        Nozzle exit diameter [mm].
    nozzle_position
        Nozzle exit position relative to the rocket's rear end [mm]. None = nozzle ends with bodytube; 
        Negative = nozzle sticks out below the bodytube. For a solid motor this moves the whole motor.
        - Example: -40
    thrust
        Thrust curve filename in the project folder (.eng, .rse, or .csv).
        - Example: "thrust.csv"
    pressure_gas_tank
        Pressurant gas tank configuration.
    fuel_tank
        Fuel tank configuration.
    oxidizer_tank
        Oxidizer tank configuration.
    """
    holddown_time: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    nozzle_diameter: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    nozzle_position: Annotated[float | list[float], FLOAT_RANGE_EXPANSION] = 0.0
    thrust: str
    pressure_gas_tank: PressureGasTankConfig
    fuel_tank: LiquidTankConfig
    oxidizer_tank: LiquidTankConfig


# -----------------------------------------------------------------------------
# Motor
# -----------------------------------------------------------------------------

class MotorConfig(_Base):
    """Solid motor physical properties and thrust curve. Required when engine.type is 'solid'.

    Attributes
    ----------
    name
        Motor name.
        - Example: "K700W"
    total_mass
        Total (wet) motor mass [g].
    propellant_mass
        Propellant mass [g].
    diameter
        Motor casing diameter [mm].
    length
        Motor casing length [mm].
    burn_time
        Nominal burn time [s].
    center_of_dry_mass_factor
        Fraction of the motor length (0-1) used to place the dry CG, measured from the nozzle.
    thrust
        Thrust curve filename in the project folder (.eng, .rse, or .csv).
        - Example: "AeroTech_K700W.eng"
    nozzle_position
        Nozzle exit position relative to the rocket's rear end [mm]. None = nozzle ends with bodytube; 
        Negative = nozzle sticks out below the bodytube. For a solid motor this moves the whole motor.
        - Example: -40
    """
    name: str
    total_mass: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    propellant_mass: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    diameter: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    length: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    burn_time: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    center_of_dry_mass_factor: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    thrust: str
    nozzle_position: Annotated[float | list[float], FLOAT_RANGE_EXPANSION] = 0.0


# -----------------------------------------------------------------------------
# Rocket body
# -----------------------------------------------------------------------------

class RocketConfig(_Base):
    """Rocket body geometry, mass properties, and drag curves.

    Attributes
    ----------
    total_mass_without_motor
        Total rocket mass without motor [g].
        Varies together with total_CG_without_motor_from_tip: as lists, both need the same length.
    length
        Rocket total length [mm].
    total_CG_without_motor_from_tip
        Center of gravity without motor, measured FROM THE NOSE TIP [mm].
        Varies together with total_mass_without_motor: as lists, both need the same length.
    diameter
        Rocket body diameter [mm].
    moment_of_intertia_Z
        Moment of inertia about the roll (longitudinal) axis without motor and propellant, calculated [kg·m²].
    moment_of_intertia_XY
        Moment of inertia about the pitch/yaw axes without motor and propellant, calculated [kg·m²].
    power_off_drag
        Power-off drag coefficient curve filename in the project folder.
        - Example: "power_off_drag.csv"
    power_on_drag
        Power-on drag coefficient curve filename in the project folder.
        - Example: "power_on_drag.csv"
    """
    total_mass_without_motor: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    length: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    total_CG_without_motor_from_tip: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    diameter: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    moment_of_intertia_Z: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    moment_of_intertia_XY: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    power_off_drag: str
    power_on_drag: str


class NoseconeConfig(_Base):
    """Nosecone geometry.

    Attributes
    ----------
    length
        Total nosecone length including the cylindrical section [mm], e.g. nosecone (585) + nosetip (74) = 659.
    cylindrical_section_length
        Cylindrical section length at the base of the nosecone [mm].
    kind
        RocketPy nosecone shape.
    """
    length: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    cylindrical_section_length: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    kind: NoseconeKind


class RailbuttonsConfig(_Base):
    """Rail button positions measured from the rocket tip.

    Attributes
    ----------
    upper_from_tip
        Upper rail button position, measured FROM THE NOSE TIP [mm].
    lower_from_tip
        Lower rail button position, measured FROM THE NOSE TIP [mm].
    """
    upper_from_tip: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    lower_from_tip: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]


class TailconeConfig(_Base):
    """Tailcone geometry.

    Attributes
    ----------
    bottom_radius
        Tailcone bottom (rear) radius [mm].
    cylindrical_section_length
        Cylindrical section length at the top of the tailcone [mm].
    length
        Total tailcone length [mm].
    """
    bottom_radius: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    cylindrical_section_length: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    length: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]


class FinsConfig(_Base):
    """Fin set geometry. Uses freeform fins if shape_points is given, otherwise trapezoidal fins.

    Attributes
    ----------
    amount
        Number of fins.
    position
        Distance of fins coming out of rocket to rear end of rocket [mm].
    name
        Fin set name.
        - Example: "Freeform"
    shape_points
        Freeform fin outline as [x, y] points in meters [m]. [0, 0] is the root leading edge; 
        the polygon is closed automatically.
    root_chord
        Root chord length for trapezoidal fins [mm].
    tip_chord
        Tip chord length for trapezoidal fins [mm].
    span
        Fin span for trapezoidal fins [mm].
    sweep_length
        Sweep length for trapezoidal fins [mm].
    """
    amount: Annotated[int | list[int], INT_RANGE_EXPANSION]
    position: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    name: str
    shape_points: Optional[list[list[float]]] = None
    root_chord: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    tip_chord: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    span: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    sweep_length: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None


# -----------------------------------------------------------------------------
# Parachutes
# -----------------------------------------------------------------------------

class ParachuteConfig(_Base):
    """Single parachute. Give cd_s directly, or cd together with radius or fabric_area.

    Attributes
    ----------
    trigger
        Deployment trigger: altitude above ground [m], or 'apogee'.
    sampling_rate
        Parachute sensor sampling rate [Hz].
    cd_s
        Drag area (cd × projected area) [m²]. Provide this, or cd + (radius or fabric_area).
    cd
        Drag coefficient. Combine with radius or fabric_area to derive cd_s.
    radius
        Projected parachute radius [m]. Used with cd to compute cd_s.
    fabric_area
        Parachute fabric area [m²]. Used with cd to compute cd_s when the projected area is close to the fabric area.
    lag
        Deployment lag after trigger [s].
    noise
        Pressure sensor noise model as [mean [Pa], std_dev [Pa], time_correlation].
    """
    trigger: Union[Annotated[float | list[float], FLOAT_RANGE_EXPANSION], Literal["apogee"]]
    sampling_rate: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    cd_s: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    cd: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    radius: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    fabric_area: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    lag: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    noise: Optional[tuple] = None


class ParachutesConfig(_Base):
    """All parachutes on the rocket.

    Attributes
    ----------
    main
        Main (recovery) parachute.
    drogue
        Drogue parachute, usually deployed at apogee. Leave out if there is none.
    payload
        Parachute of the separating payload. Leave out if there is none.
    """
    main: ParachuteConfig
    drogue: Optional[ParachuteConfig] = None
    payload: Optional[ParachuteConfig] = None


# -----------------------------------------------------------------------------
# Flight
# -----------------------------------------------------------------------------

class FlightConfig(_Base):
    """Launch rail configuration and variation parameters.

    Attributes
    ----------
    rail_length
        Launch rail length [m].
    inclination
        Launch rail inclination [°]. 0 = horizontal, 90 = vertical.
        - Example: 88
        - Example: [84, 86, 88]
        - Example: "84..88:2"
    heading
        Launch rail heading [°]. 0 = north, 90 = east. In a range string, start > stop wraps around 360°.
        - Example: 350
        - Example: [0, 90, 180, 270]
        - Example: "340..360:10"
        - Example: "350..10:5"
    """
    rail_length: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    # One example per accepted form: single value, list, range string.
    inclination: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]
    heading: Annotated[float | list[float], FLOAT_RANGE_EXPANSION]


# -----------------------------------------------------------------------------
# Payload
# -----------------------------------------------------------------------------

class PayloadConfig(_Base):
    """Deployable payload configuration. Only active when mass_total > 0.

    Attributes
    ----------
    mass_total
        Total deployable payload mass [g]. 0 = no payload separates.
    mass
        Mass of the separated payload body [g]. Required when mass_total > 0.
    diameter
        Payload body diameter [mm]. Required when mass_total > 0.
    length
        Payload total length [mm]. Required when mass_total > 0.
    moment_of_intertia_XY
        Payload moment of inertia about the pitch/yaw axes [kg·m²]. Required when mass_total > 0.
    moment_of_intertia_Z
        Payload moment of inertia about the roll axis [kg·m²]. Required when mass_total > 0.
    """
    mass_total: Annotated[float | list[float], FLOAT_RANGE_EXPANSION] = 0.0
    mass: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    diameter: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    length: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    moment_of_intertia_XY: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    moment_of_intertia_Z: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None


# -----------------------------------------------------------------------------
# Reanalysis
# -----------------------------------------------------------------------------

class MotorOverrideConfig(_Base):
    """Per-parameter motor overrides for the matched flight.

    Attributes
    ----------
    burn_time
        Override burn time [s].
    thrust
        Override thrust curve filename in the project folder.
        - Example: "AeroTech_K700W.eng"
    """
    burn_time: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    # TODO: implement multiple thrust curves for variation
    thrust: Optional[str] = None


class ParachuteOverrideConfig(_Base):
    """Per-parameter parachute overrides for the matched flight.

    Attributes
    ----------
    enabled
        Set to false to remove this parachute from the matched flight entirely.
    cd_s
        Override drag area (cd × projected area) [m²].
    cd
        Override drag coefficient.
    radius
        Override projected radius [m].
    fabric_area
        Override fabric area [m²].
    """
    enabled: bool = True
    cd_s: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    cd: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    radius: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    fabric_area: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None


class MatchedFlightConfig(_Base):
    """A flight that branches off the nominal at match_time with a sensor-matched initial state.

    Attributes
    ----------
    match_time
        Time [s] at which to branch off the nominal flight and apply the sensor-matched state.
    match_parameters
        Flight-computer parameters to match at match_time.
    motor
        Motor overrides applied to the matched flight.
    parachutes
        Per-parachute overrides keyed by parachute name, e.g. [reanalysis.matched_flight.parachutes.main].
    """
    match_time: Optional[Annotated[float | list[float], FLOAT_RANGE_EXPANSION]] = None
    match_parameters: list[MatchParameter] = Field(default_factory=list)
    motor: Optional[MotorOverrideConfig] = None
    parachutes: Optional[dict[str, ParachuteOverrideConfig]] = None


class ReanalysisConfig(_Base):
    """Flight-computer data sources and matched-flight configuration for reanalysis mode.

    Attributes
    ----------
    sources
        Flight computers to load data from.
    matched_flight
        Sensor-matched flight branch. Leave out to skip the matched-flight simulation.
    """
    sources: list[FlightComputerSource]
    matched_flight: Optional[MatchedFlightConfig] = None


# =============================================================================
# Top-level config
# =============================================================================

class Config(_Base):
    """Full project config loaded from config.toml. Lengths are in mm and masses in g unless stated otherwise.

    Attributes
    ----------
    project
        Project name, used for labeling outputs.
        - Example: "ALBATROSS"
    output_level
        Amount of output generated:
        - 0 (standard): most important prints + landing/safety plots
        - 1 (detailed): level 0 + per-flight custom plots + env prints
        - 2 (more detailed): level 1 + rocket/motor prints
        - 3 (debug): level 2 + more env plots
    scenarios
        Scenarios to simulate for every flight. Supported values:
        - "nominal": All parachutes deploy. Required, since the safety analysis and deployable payload build on it.
        - "no_main": Flight without main parachute (only simulated if a drogue parachute is configured)
        - "ballistic": No parachute flight scenario
        - Example: ["nominal", "ballistic"]
    """
    project: str
    output_level: int = 0
    scenarios: list[ScenarioName] = ["nominal", "no_main", "ballistic"]
    # The section fields below need no description: the hover shows the docstring of their model class.
    environment: EnvironmentConfig
    engine: EngineConfig
    motor: Optional[MotorConfig] = None
    liquid_engine: Optional[LiquidEngineConfig] = None
    rocket: RocketConfig
    nosecone: NoseconeConfig
    railbuttons: RailbuttonsConfig
    tailcone: TailconeConfig
    fins: FinsConfig
    parachutes: ParachutesConfig
    flight: FlightConfig
    payload: PayloadConfig = PayloadConfig()
    reanalysis: Optional[ReanalysisConfig] = None

    @model_validator(mode="after")
    def check_paired_variation_lengths(self) -> "Config":
        """Reject fields that vary together when their value lists have different lengths."""
        for group in PAIRED_VARIATIONS:
            values_by_name = {".".join(path): functools.reduce(getattr, path, self) for path in group}
            list_lengths = {name: len(values) for name, values in values_by_name.items() if isinstance(values, list)}
            if len(set(list_lengths.values())) > 1:
                raise ValueError(f"These fields vary together and need the same number of values: {list_lengths}")
        return self


# =============================================================================
# Zones
# =============================================================================

# One polygon corner as [distance_m, heading_deg] from the launch rail.
_PolarPoint = tuple[float, float]


class ZonesConfig(_Base):
    """Landing zones loaded from zones.toml. Each corner is [distance_m, heading_deg] from the launch rail, 
    measured in Google Earth (heading clockwise from north: 0° = North, 90° = East).

    Attributes
    ----------
    exclusion_zones
        Hard no-land areas (e.g. crowd, houses), shown in red. Keyed by zone name; quote names with spaces: 
        "My Zone" = [[distance_m, heading_deg], ...].
    buffer_zones
        Safety margin zones, or areas the rocket should preferably not land in (e.g. forest), shown in orange. 
        A configuration is safe only if every simulated landing is outside every buffer zone.
    exclusion_zone_safety_margin
        Factor that enlarges the exclusion zones into extra buffer zones, e.g. 1.3 = 30% larger.
    buffer_zones_are_suboptimal_but_safe
        If true, landings in buffer_zones mark a heading as "suboptimal" (still safe) instead of unsafe;
        shown in yellow. The enlarged exclusion zones stay unsafe either way.
    """
    exclusion_zones: dict[str, list[_PolarPoint]] = Field(default_factory=dict)
    buffer_zones: dict[str, list[_PolarPoint]] = Field(default_factory=dict)
    exclusion_zone_safety_margin: float = 1.0
    buffer_zones_are_suboptimal_but_safe: bool = False


# =============================================================================
# Runtime and top-level container
# =============================================================================

class RuntimeParams(BaseModel):
    """Runtime objects built incrementally during the simulation pipeline.

    Add new fields here when a create_* function produces an object other code needs.
    All fields are Optional so the object can be created before any simulation runs.

    Attributes:
    ----------
    environments: 
        Dict of {name: Environment} built by create_environment.
    motor: 
        Motor object built by create_engine.
    nosecone: 
        NoseCone object built by create_rocket.
    tailcone: 
        Tail object built by create_rocket.
    fin_set: 
        Fin set object built by create_rocket.
    parachutes: 
        Dict of {index: Parachute} built by create_rocket.
    rocket: 
        Rocket object built by create_rocket.
    flights_by_env: 
        Dict of {env_name: [scenario_set, ...]} built by create_flight.
    scenario_sets: 
        List of all scenario_set dicts across all environments, built by create_flight.
        Each scenario_set is a ``{scenario_name: Flight}`` dict for one heading/inclination/env combination.<br>
        Keys: 
        - ``"nominal"`` (always)
        - ``"no_main"`` (only with drogue and when selected in config.scenarios)
        - ``"ballistic"`` (only when selected in config.scenarios)
        - ``"matched"`` (only in reanalysis mode)
    ascent_flights_by_env: 
        Dict of {env_name: [Flight, ...]} when ascent flights are reused (variations or a deployable payload).
    payload_parachute: 
        Payload parachute dict built by deployable_payload.
    payload: 
        Payload rocket dict built by deployable_payload.
    flight_payload: 
        List of payload Flight objects.
    reanalysis_mode: True if any environment is of type reanalysis.
    gnss_3d_traces: 
        Dict of GNSS 3D traces from reanalysis for the trajectory plot.
    flight_computer_impacts: 
        List of flight-computer impact dicts from reanalysis.
    reanalysis_flight_computer_data: 
        Dict of {env_name: loaded flight-computer data}.
    safe_rocket_nominal: 
        Safe nominal rocket flights (safety analysis).
    safe_rocket_no_main: 
        Safe no-main rocket flights (safety analysis).
    safe_rocket_ballistic: 
        Safe ballistic rocket flights (safety analysis).
    safe_rocket_matched: 
        Safe matched rocket flights (safety analysis).
    safe_configurations: 
        List of (env, heading, inclination) tuples classified as safe.
    unsafe_rocket_nominal: 
        Unsafe nominal rocket flights (safety analysis).
    unsafe_rocket_no_main: 
        Unsafe no-main rocket flights (safety analysis).
    unsafe_rocket_ballistic: 
        Unsafe ballistic rocket flights (safety analysis).
    unsafe_rocket_matched: 
        Unsafe matched rocket flights (safety analysis).
    unsafe_configurations: 
        List of (env, heading, inclination) tuples classified as unsafe.
    unsafe_details:
        List of detail dicts for unsafe configurations.
    suboptimal_rocket_nominal, suboptimal_rocket_no_main, suboptimal_rocket_ballistic, suboptimal_rocket_matched:
        Flights of headings that are safe but land in a suboptimal zone (safety analysis).
    suboptimal_configurations:
        List of (env, heading, inclination) tuples classified as suboptimal.
    suboptimal_details:
        List of detail dicts for configurations landing in a suboptimal zone.
    safe_payload:
        Safe payload flights (safety analysis).
    suboptimal_payload:
        Suboptimal payload flights (safety analysis).
    unsafe_payload:
        Unsafe payload flights (safety analysis).
    """
    model_config = ConfigDict(arbitrary_types_allowed=True)

    environments: Optional[dict] = None
    motor: Optional[Any] = None
    nosecone: Optional[Any] = None
    tailcone: Optional[Any] = None
    fin_set: Optional[Any] = None
    parachutes: Optional[Any] = None
    rocket: Optional[Any] = None
    flights_by_env: Optional[dict] = None
    scenario_sets: list = Field(default_factory=list)
    ascent_flights_by_env: Optional[dict] = None
    payload_parachute: Optional[Any] = None
    payload: Optional[Any] = None
    flight_payload: Optional[list] = None
    reanalysis_mode: bool = False
    gnss_3d_traces: Optional[dict] = None
    flight_computer_impacts: Optional[list] = None
    reanalysis_flight_computer_data: Optional[dict] = None
    safe_rocket_nominal: Optional[list] = None
    safe_rocket_no_main: Optional[list] = None
    safe_rocket_ballistic: Optional[list] = None
    safe_rocket_matched: Optional[list] = None
    safe_configurations: Optional[list] = None
    unsafe_rocket_nominal: Optional[list] = None
    unsafe_rocket_no_main: Optional[list] = None
    unsafe_rocket_ballistic: Optional[list] = None
    unsafe_rocket_matched: Optional[list] = None
    unsafe_configurations: Optional[list] = None
    unsafe_details: Optional[list] = None
    suboptimal_rocket_nominal: Optional[list] = None
    suboptimal_rocket_no_main: Optional[list] = None
    suboptimal_rocket_ballistic: Optional[list] = None
    suboptimal_rocket_matched: Optional[list] = None
    suboptimal_configurations: Optional[list] = None
    suboptimal_details: Optional[list] = None
    safe_payload: Optional[list] = None
    suboptimal_payload: Optional[list] = None
    unsafe_payload: Optional[list] = None


class SimParams(BaseModel):
    """Top-level simulation container: typed config, runtime state, and project metadata.

    Attributes:
    ----------
    config (Config): 
        Typed configuration loaded from config.toml. Use for all parameter access
        instead of flat string-key dicts.
    runtime (RuntimeParams): 
        Runtime objects built during the simulation pipeline. Populated
        incrementally by create_environment, create_engine, create_rocket, etc.
    project_path (Path): 
        Absolute path to the project folder (e.g. ./ALBATROSS/).
    export_kml (bool): 
        True when no flight variations are active (single result → KML export).
    """
    model_config = ConfigDict(arbitrary_types_allowed=True)

    config: Config
    runtime: RuntimeParams = Field(default_factory=RuntimeParams)
    project_path: Path
    export_kml: bool


# =============================================================================
# Generate JSON schemas (python simulation_core/config_schema.py)
# =============================================================================

def adapt_schema_for_tombi(schema: dict) -> None:
    """Adjust the generated schema so Tombi's error checks and hover help match what pydantic accepts;
    pydantic's validation is unchanged.

    A field that accepts several types lists one alternative per type in "anyOf", e.g. heading (simplified):
        {"examples": [350, "84..88:2"], "anyOf": [{"type": "number"}, {"type": "array"}, {"type": "string"}]}

    1. Tombi reads "number" as float only, so heading = 350 would be an error. Thus an "integer" alternative is
        added by this function:
        "anyOf": [{"type": "integer"}, {"type": "number"}, ...]
    2. Hovering a value shows only the alternative of the value's type, e.g. {"type": "string"}. Add examples to the alternative
        so the hover shows it.
    3. Hovering an element inside a list, e.g. the 90 in heading = [0, 90, 180], does not show examples. So this function
        adds the list entries as examples, and the same "integer" alternative for the items.
    """

    models = [schema, *schema.get("$defs", {}).values()]
    for model in models:
        for field in model.get("properties", {}).values():
            if "anyOf" not in field:
                continue

            # 1. Put an "integer" alternative before each "number" one, so a value like heading = 350 matches an alternative.
            alternatives = []
            for alternative in field["anyOf"]:
                if alternative.get("type") == "number":
                    alternatives.append({"type": "integer"})
                alternatives.append(alternative)
            field["anyOf"] = alternatives

            # 2. Copy the field's examples into the alternatives of the same type, e.g. heading's "84..88:2" into the string one.
            for alternative in field["anyOf"]:
                # type(example) instead of isinstance, so True (a bool) is not treated as an int.
                matching_examples = [
                    example for example in field.get("examples", [])
                    if alternative.get("type") in JSON_SCHEMA_TYPES_BY_PYTHON_TYPE.get(type(example), set())
                ]
                if matching_examples:
                    alternative["examples"] = matching_examples
                # 3. Hovering an element inside a TOML array shows the array's item schema, so give it the list entries as examples.
                if alternative.get("type") == "array" and isinstance(alternative.get("items"), dict):
                    item_examples = [item for example in matching_examples for item in example]
                    if item_examples:
                        alternative["items"]["examples"] = item_examples
                    # Same integer problem as in 1., one level deeper: let a list element like the 90 in [0, 90, 180] match as integer too.
                    if alternative["items"].get("type") == "number":
                        number_items = alternative["items"]
                        integer_items = {"type": "integer"}
                        if "examples" in number_items:
                            integer_items["examples"] = number_items["examples"]
                        alternative["items"] = {"anyOf": [integer_items, number_items]}


def write_schema_files():
    """Write the JSON schemas of Config and ZonesConfig, so the Tombi editor extension can autocomplete and check the TOML files."""
    SCHEMA_DIR.mkdir(exist_ok=True)
    for model, schema_path in [(Config, CONFIG_SCHEMA_PATH), (ZonesConfig, ZONES_SCHEMA_PATH)]:
        # mode="validation" describes the input pydantic accepts, i.e. what you type into the TOML file.
        schema = model.model_json_schema(mode="validation")
        adapt_schema_for_tombi(schema)
        schema_path.write_text(json.dumps(schema, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")
        print(f"Wrote {schema_path}")


if __name__ == "__main__":
    write_schema_files()
