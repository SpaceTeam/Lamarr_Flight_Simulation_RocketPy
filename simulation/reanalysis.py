"""
Reanalysis: compare onboard flight-computer data to the RocketPy simulation post-flight.
"""

import copy
import inspect
import warnings

import numpy as np
import plotly.graph_objects as go
from rocketpy import Function, Flight, LiquidMotor, Motor, Parachute, SolidMotor
from rocketpy.simulation.flight_data_importer import FlightDataImporter
from pyproj import Geod
from scipy.spatial.transform import Rotation, Slerp
import pymap3d as pm
import pandas as pd

from simulation.utils import *
from simulation.custom_print_and_plot_functions import CustomPlots


# =============================================================================
# Flight-computer data locations
# =============================================================================
# TODO: maybe implement in config file instead?

# Subfolders inside the project folder that hold each flight-computer's CSV dumps
CATS_FOLDER = "CATS_FLIGHT_DATA"
ALTIMAX_FOLDER = "ALTIMAX_FLIGHT_DATA"
RCU_FOLDER = "RCU_FLIGHT_DATA"

CATS_FILES = [
    "filteredDataInfo.csv",
    "gnssInfo.csv",
    "flightInfo.csv",
    "imu.csv",
    "baro.csv",
]
ALTIMAX_FILES = ["altimax_export.csv"]
RCU_FILES = ["rcu_export.csv"]


CATS_COLUMNS_MAP = {
    "ts": "time",
    "filteredAltitudeAGL": "altitude",
    "filteredAcceleration": "az",
    "latitude": "latitude",
    "longitude": "longitude",
    "Ax": "acceleration_x",
    "Ay": "acceleration_y",
    "Az": "acceleration_z",
    "Gx": "gyro_x",
    "Gy": "gyro_y",
    "Gz": "gyro_z",
    "velocity": "speed",
    "P": "pressure",
}

ALTIMAX_COLUMNS_MAP = {
    "ZEIT": "time",
    "PRESS_FILTER": "pressure",
    "HEIGHT_FILTER": "altitude",
    "ACCEL": "acceleration",
    "SPEED": "speed",
}

RCU_COLUMNS_MAP = {
    "time": "time",
    "lora:gps_altitude:sensor": "altitude",
    "lora:gps_latitude:sensor": "latitude",
    "lora:gps_longitude:sensor": "longitude",
    "lora:gps_status:sensor": "gps_status",
    "lora:rcu_accel_x:sensor": "accel_x",
    "lora:rcu_accel_y:sensor": "accel_y",
    "lora:rcu_accel_z:sensor": "accel_z",
    "lora:rcu_gyro_x:sensor": "gyro_x",
    "lora:rcu_gyro_y:sensor": "gyro_y",
    "lora:rcu_gyro_z:sensor": "gyro_z",
    "lora:rcu_barometer:sensor": "pressure",
}


# =============================================================================
# Helpers
# =============================================================================

WGS84_GEOD = Geod(ellps="WGS84")


def pressure_to_altitude(data):
    """
    Convert barometric pressure (hPa) to altitude (m) using the standard barometric formula.
    Reference: https://www.weather.gov/media/epz/wxcalc/pressureAltitude.pdf
    """
    return 0.3048 * ((1 - (data / 1013.25) ** 0.190284) * 145366.45)


def _print_simulated_error_line(label, actual, simulated, unit, unit_error=None):
    """Print one simulated value with absolute and percentage error on the same line."""
    # Compute the error values against the actual measured value.
    error = abs(actual - simulated)
    percentage = error / actual * 100
    print(
        f"Simulated {label}: {simulated:.2f} {unit} | "
        f"Error: {error:.2f} {unit_error or unit} | "
        f"Percentage Error: {percentage:.2f}%"
    )


def cats_event_markers(project_path):
    """Load events markers from CATS vega."""
    event_file_path = project_path / CATS_FOLDER / "eventInfo.csv"

    event_markers = []

    if event_file_path.exists():
        event_data = pd.read_csv(event_file_path)

        known_event_labels = {
            2: "Liftoff",
            3: "Burnout",
            4: "Apogee",
            5: "Main deployment",
        }

        for _, row in event_data.iterrows():
            event_time = float(row["ts"])
            event = int(row["event"])
            label = known_event_labels.get(event)

            if label is None:
                continue

            event_markers.append((event_time, label, "gray"))

    return event_markers
            

# =============================================================================
# Flight-computer data loading
# =============================================================================

SUPPORTED_SOURCES = ("cats_vega", "altimax", "rcu")

def load_flight_computer_data(project_path, sources):
    """Load flight-computer CSVs for the selected sources. Returns a dict keyed by source name."""
    unknown = set(sources) - set(SUPPORTED_SOURCES)
    if unknown:
        raise ValueError(f"Unknown reanalysis source(s): {sorted(unknown)}. Choose from: {list(SUPPORTED_SOURCES)}.")

    data = {}

    if "cats_vega" in sources:
        data["cats_vega"] = FlightDataImporter(
            name="CATS Vega Flight Data",
            paths=[str(project_path / CATS_FOLDER / name) for name in CATS_FILES],
            columns_map=CATS_COLUMNS_MAP,
            units=None,
            interpolation="linear",
            extrapolation="zero",
            delimiter=",",
            encoding="utf-8",
        )

    if "altimax" in sources:
        data["altimax"] = FlightDataImporter(
            name="Altimax Flight Data",
            paths=[str(project_path / ALTIMAX_FOLDER / name) for name in ALTIMAX_FILES],
            columns_map=ALTIMAX_COLUMNS_MAP,
            units=None,
            interpolation="linear",
            extrapolation="zero",
            delimiter=",",
            encoding="utf-8",
        )

    if "rcu" in sources:
        data["rcu"] = FlightDataImporter(
            name="SRAD Flight Data",
            paths=[str(project_path / RCU_FOLDER / name) for name in RCU_FILES],
            columns_map=RCU_COLUMNS_MAP,
            units=None,
            interpolation="linear",
            extrapolation="constant",
            delimiter=",",
            encoding="utf-8",
        )

    return data



def _parachute_overrides_from_config(constants, variations):
    """Extract per-parachute override dicts from flattened reanalysis.matched_flight.parachutes.<name>.<param> config keys.

    Returns {parachute_name: {param: value}}. Empty dict when no overrides are configured.
    Assumes parachute names do not themselves contain underscores (e.g. main, drogue).
    """
    prefix = "reanalysis_matched_flight_parachutes_"
    overrides = {}
    for source in (constants, variations):
        for key, value in source.items():
            if not key.startswith(prefix):
                continue
            remainder = key[len(prefix):]
            parts = remainder.split("_", 1)
            if len(parts) != 2:
                continue
            parachute_name, param = parts
            overrides.setdefault(parachute_name, {})[param] = value
    return overrides


def _motor_overrides_from_config(constants, variations):
    """Extract matched-flight motor overrides from flattened reanalysis.matched_flight.motor.<param> config keys."""
    prefix = "reanalysis_matched_flight_motor_"
    overrides = {}
    for source in (constants, variations):
        for key, value in source.items():
            if key.startswith(prefix):
                overrides[key[len(prefix):]] = value
    return overrides

# TODO: maybe implement parachute and motor creation for the matched flight in simulation.py instead of here 
# for less code duplication.

def rebuild_motor_from_overrides(nominal_motor: Motor, overrides: dict, project_path=None):
    """Build a new RocketPy motor with only thrust source and burn time overrides applied."""
    if not overrides:
        return nominal_motor

    # Keep this intentionally narrow; broader motor changes should go through the normal motor config.
    unknown_keys = set(overrides) - {"thrust", "burn_time"}
    if unknown_keys:
        raise ValueError(
            f"Unknown matched-flight motor override(s): {sorted(unknown_keys)}. "
            "Supported keys: ['burn_time', 'thrust']"
        )

    thrust_source = overrides.get("thrust", nominal_motor.thrust_source)
    if isinstance(thrust_source, str) and project_path is not None:
        thrust_source = get_project_file({"project": project_path}, thrust_source)

    dry_inertia = (
        nominal_motor.dry_I_11,
        nominal_motor.dry_I_22,
        nominal_motor.dry_I_33,
        nominal_motor.dry_I_12,
        nominal_motor.dry_I_13,
        nominal_motor.dry_I_23,
    )
        
    base_options = {
        "thrust_source": thrust_source,
        "dry_mass": nominal_motor.dry_mass,
        "dry_inertia": dry_inertia,
        "nozzle_radius": nominal_motor.nozzle_radius,
        "center_of_dry_mass_position": nominal_motor.center_of_dry_mass_position,
        "nozzle_position": nominal_motor.nozzle_position,
        "burn_time": overrides.get("burn_time", nominal_motor.burn_time),
        "coordinate_system_orientation": nominal_motor.coordinate_system_orientation,
        "reference_pressure": nominal_motor.reference_pressure,
    }

    if isinstance(nominal_motor, SolidMotor):
        return SolidMotor(
            **base_options,
            grain_number=nominal_motor.grain_number,
            grain_density=nominal_motor.grain_density,
            grain_outer_radius=nominal_motor.grain_outer_radius,
            grain_initial_inner_radius=nominal_motor.grain_initial_inner_radius,
            grain_initial_height=nominal_motor.grain_initial_height,
            grain_separation=nominal_motor.grain_separation,
            grains_center_of_mass_position=nominal_motor.grains_center_of_mass_position,
            throat_radius=nominal_motor.throat_radius,
            only_radial_burn=nominal_motor.only_radial_burn,
        )

    if isinstance(nominal_motor, LiquidMotor):
        # TODO: implement
        raise TypeError("Matched-flight motor overrides currently support SolidMotor only.")

    raise TypeError(f"Matched-flight motor overrides do not support motor type {type(nominal_motor).__name__}.")



def rebuild_parachute_from_overrides(nominal_parachute: Parachute, overrides: dict):
    """Build a new Parachute from a nominal one with per-parameter overrides applied. Returns None when overrides set enabled=False."""
    # Drop the parachute entirely when the matched-flight config disables it.
    if overrides.get("enabled") is False:
        return None

    if "cd_s" in overrides:
        cd_s = overrides["cd_s"]

    elif "cd" in overrides and "fabric_area" in overrides:
        cd_s = overrides["cd"] * overrides["fabric_area"]

    elif "cd" in overrides and "radius" in overrides:
        cd_s = overrides["cd"] * np.pi * overrides["radius"] ** 2

    elif "cd" in overrides or "fabric_area" in overrides:
        raise ValueError(
            f"Parachute '{nominal_parachute.name}' override needs either 'cd_s', "
            f"'cd' with 'radius', or 'cd' with 'fabric_area'. Got: {overrides}"
        )

    else:
        cd_s = nominal_parachute.cd_s

    # Merge the user-provided settings over the nominal parachute settings.
    parachute_options = {
        "name": overrides.get("name", nominal_parachute.name),
        "cd_s": cd_s,
        "trigger": overrides.get("trigger", nominal_parachute.trigger),
        "sampling_rate": overrides.get("sampling_rate", nominal_parachute.sampling_rate),
        "lag": overrides.get("lag", nominal_parachute.lag),
        "noise": overrides.get("noise", nominal_parachute.noise),
        "radius": overrides.get("radius", nominal_parachute.radius),
        "height": overrides.get("height", nominal_parachute.height),
        "porosity": overrides.get("porosity", nominal_parachute.porosity),
        "drag_coefficient": overrides.get(
            "drag_coefficient",
            overrides.get("cd", nominal_parachute.drag_coefficient),
        ),
    }

    return Parachute(**parachute_options)


def cats_quaternion_at(project_path, t):
    """Load the CATS Vega attitude quaternion (xyzw, normalized) at time t by spherical interpolation."""
    orientation_file = project_path / CATS_FOLDER / "orientationInfo.csv"
    orientation_data = pd.read_csv(orientation_file)
    times = orientation_data["ts"].to_numpy()

    # Load the CATS quaternion columns and normalize them before interpolation.
    quaternions_xyzw = orientation_data[["q0_estimated", "q1_estimated", "q2_estimated", "q3_estimated"]].to_numpy()
    quaternion_norms = np.linalg.norm(quaternions_xyzw, axis=1, keepdims=True)
    quaternions_xyzw = quaternions_xyzw / quaternion_norms

    # Spherical linear interpolation follows the shortest path on the rotation sphere.
    rotations = Rotation.from_quat(quaternions_xyzw)
    spherical_interpolator = Slerp(times, rotations)
    return spherical_interpolator([t]).as_quat()[0]


def _bearing_between_samples(times, lats, lons, t_start, t_end, min_displacement_m=0.1):
    """Compass bearing (0=N, 90=E) of the ground track between t_start and t_end, using interpolation of pre-cleaned GNSS arrays. None when out of range or displacement too small."""
    if t_start < times[0] or t_end > times[-1]:
        return None

    lat_start = float(np.interp(t_start, times, lats))
    lat_end = float(np.interp(t_end, times, lats))
    lon_start = float(np.interp(t_start, times, lons))
    lon_end = float(np.interp(t_end, times, lons))

    # WGS84 ellipsoid inverse geodesic: returns forward azimuth, back azimuth, and distance.
    bearing_deg, _, distance_m = WGS84_GEOD.inv(lon_start, lat_start, lon_end, lat_end)

    if distance_m < min_displacement_m:
        return None

    return float(bearing_deg % 360.0)


def _valid_gnss_samples(times, lats, lons):
    """Return the subset of GNSS samples where lat and lon are both finite and non-zero (drops dropouts)."""
    valid = np.isfinite(lats) & np.isfinite(lons) & (lats != 0) & (lons != 0)
    return times[valid], lats[valid], lons[valid]


def compute_gnss_bearing(flight_computer, t_start, t_end):
    """Compute the compass bearing (degrees, 0=N, 90=E) of the GNSS ground track between t_start and t_end. Returns None when no usable samples cover the window."""
    if not hasattr(flight_computer, "latitude") or not hasattr(flight_computer, "longitude"):
        return None

    lat_source = flight_computer.latitude.source
    lon_source = flight_computer.longitude.source
    times, lats, lons = _valid_gnss_samples(lat_source[:, 0], lat_source[:, 1], lon_source[:, 1])
    if len(times) == 0:
        return None

    return _bearing_between_samples(times, lats, lons, t_start, t_end, min_displacement_m=1e-3)


def compute_cats_gnss_heading_at_times(project_path, times, window_half_width=0.5):
    """
    For each requested time, compute the CATS GNSS-derived heading averaged over [t - window, t + window]. 
    Returns NaN where unavailable.
    """
    gnss_file = project_path / CATS_FOLDER / "gnssInfo.csv"
    headings = np.full(len(times), np.nan)
    if not gnss_file.exists():
        return headings

    gnss_data = pd.read_csv(gnss_file)
    gnss_times, lats, lons = _valid_gnss_samples(
        gnss_data["ts"].to_numpy(),
        gnss_data["latitude"].to_numpy(),
        gnss_data["longitude"].to_numpy(),
    )
    if len(gnss_times) == 0:
        return headings

    for i, t in enumerate(times):
        bearing = _bearing_between_samples(gnss_times, lats, lons, t - window_half_width, t + window_half_width)
        if bearing is not None:
            headings[i] = bearing
    return headings


def create_matched_flight(nominal_flight: Flight, project_path, t_match, data=None, parachute_overrides=None, motor_overrides=None):
    """
    Create a Flight that branches off nominal at t_match with attitude+velocity rotated to match CATS pitch and GNSS heading at 
    that time, optionally rebuilding motor/parachutes from config overrides.
    """
    # Sample the 14-element RocketPy state vector [t, x, y, z, vx, vy, vz, e0..e3, w1..w3] from a Flight at time t.
    state = [
        t_match,
        float(nominal_flight.x(t_match)), float(nominal_flight.y(t_match)), float(nominal_flight.z(t_match)),
        float(nominal_flight.vx(t_match)), float(nominal_flight.vy(t_match)), float(nominal_flight.vz(t_match)),
        float(nominal_flight.e0(t_match)), float(nominal_flight.e1(t_match)), float(nominal_flight.e2(t_match)), float(nominal_flight.e3(t_match)),
        float(nominal_flight.w1(t_match)), float(nominal_flight.w2(t_match)), float(nominal_flight.w3(t_match)),
    ]

    # Sim rotation: RocketPy uses (e0=w, e1=x, e2=y, e3=z); scipy expects (x, y, z, w)
    sim_quat_wxyz = state[7:11]
    sim_rotation = Rotation.from_quat([sim_quat_wxyz[1], sim_quat_wxyz[2], sim_quat_wxyz[3], sim_quat_wxyz[0]])
    # RocketPy body +Z is the rocket longitudinal axis
    sim_rocket_axis_world = sim_rotation.apply([0.0, 0.0, 1.0])

    # CATS rocket axis in its own world frame — Z component is the pitch (earth-true); X/Y are bogus (sensor-fixed frame)
    cats_rotation = Rotation.from_quat(cats_quaternion_at(project_path, t_match))
    cats_rocket_axis_cats_world = cats_rotation.apply([0.0, 0.0, -1.0])    # body -Z is the rocket axis (same assumption as plot_cats_attitude_angle)
    pitch_from_vertical_rad = float(np.arccos(np.clip(cats_rocket_axis_cats_world[2], -1.0, 1.0)))

    # GNSS-derived heading at t_match (compass-true). Prefer CATS GNSS, fall back to RCU.
    heading_rad = None
    if data is not None:
        for source_key in ("cats_vega", "rcu"):
            if source_key not in data:
                continue
            bearing = compute_gnss_bearing(data[source_key], t_match - 0.5, t_match + 0.5)
            if bearing is not None:
                heading_rad = np.radians(bearing)
                break

    # Build the desired rocket axis in earth (ENU) frame from CATS pitch + GNSS heading
    if heading_rad is not None:
        desired_axis_world = np.array([
            np.sin(pitch_from_vertical_rad) * np.sin(heading_rad),    # east  (X)
            np.sin(pitch_from_vertical_rad) * np.cos(heading_rad),    # north (Y)
            np.cos(pitch_from_vertical_rad),                          # up    (Z)
        ])
    else:
        # GNSS unavailable — fall back to CATS quaternion only (bogus azimuth, but pitch is still right)
        print("[reanalysis] create_matched_flight: no GNSS bearing available at t_match, falling back to CATS-only attitude (azimuth will be wrong).")
        desired_axis_world = cats_rocket_axis_cats_world

    # Shortest-path correction that maps simulation rocket axis onto the desired axis
    vec_from = sim_rocket_axis_world / np.linalg.norm(sim_rocket_axis_world)
    vec_to = desired_axis_world / np.linalg.norm(desired_axis_world)
    dot = float(np.dot(vec_from, vec_to))

    # Already aligned: no rotation needed
    if dot > 1.0 - 1e-9:
        correction = Rotation.identity()

    # Opposite vectors: 180° around any perpendicular axis
    elif dot < -1.0 + 1e-9:
        perpendicular = np.array([1.0, 0.0, 0.0]) if abs(vec_from[0]) < 0.9 else np.array([0.0, 1.0, 0.0])
        axis = np.cross(vec_from, perpendicular)
        correction = Rotation.from_rotvec(np.pi * axis / np.linalg.norm(axis))

    else:
        axis = np.cross(vec_from, vec_to)
        angle = float(np.arccos(dot))
        correction = Rotation.from_rotvec(angle * axis / np.linalg.norm(axis))

    # Apply correction to both attitude and velocity so the splice has zero angle-of-attack induced
    new_rotation = correction * sim_rotation
    new_quat_xyzw = new_rotation.as_quat()
    new_quat_wxyz = [new_quat_xyzw[3], new_quat_xyzw[0], new_quat_xyzw[1], new_quat_xyzw[2]]
    new_velocity = correction.apply(state[4:7])

    new_state = [
        state[0],
        state[1], state[2], state[3],
        float(new_velocity[0]), float(new_velocity[1]), float(new_velocity[2]),
        *new_quat_wxyz,
        state[11], state[12], state[13],
    ]

    # Apply matched-flight overrides to a deep-copied rocket so the nominal rocket stays untouched.
    matched_rocket = nominal_flight.rocket
    if motor_overrides or parachute_overrides:
        matched_rocket = copy.deepcopy(nominal_flight.rocket)

    # Rebuild the motor when the matched-flight config changes thrust, mass, burn time, or motor geometry.
    if motor_overrides:
        rebuilt_motor = rebuild_motor_from_overrides(
            matched_rocket.motor,
            motor_overrides,
            project_path=project_path,
        )
        matched_rocket.add_motor(
            rebuilt_motor,
            position=matched_rocket.motor_position,
        )

    # Each matching parachute is rebuilt with merged params; parachutes with enabled=False are dropped.
    if parachute_overrides:
        rebuilt_parachutes = []
        matched_parachute_names = set()
        for parachute in matched_rocket.parachutes:
            if parachute.name in parachute_overrides:
                matched_parachute_names.add(parachute.name)
                rebuilt = rebuild_parachute_from_overrides(parachute, parachute_overrides[parachute.name])
                if rebuilt is None:
                    # drop parachute
                    continue
                rebuilt_parachutes.append(rebuilt)
            else:
                rebuilt_parachutes.append(parachute)
        unknown_parachute_names = set(parachute_overrides) - matched_parachute_names
        if unknown_parachute_names:
            available_names = [parachute.name for parachute in matched_rocket.parachutes]
            raise ValueError(
                f"Unknown matched-flight parachute override(s): {sorted(unknown_parachute_names)}. "
                f"Available parachutes: {available_names}"
            )
        matched_rocket.parachutes = rebuilt_parachutes

    matched_flight = Flight(
        rocket=matched_rocket,
        environment=nominal_flight.env,
        rail_length=nominal_flight.rail_length,
        inclination=nominal_flight.inclination,
        heading=nominal_flight.heading,
        initial_solution=new_state,
        terminate_on_apogee=False,
        name=f"{nominal_flight.env.name}_matched",
    )

    # attach_meta sets `filename_label` and `name`; downstream code (KML/CSV export) expects both
    attach_meta(matched_flight, getattr(nominal_flight, "_meta", {}))
    matched_flight.splice_time = t_match

    # Pre-warm matched's internals so the shallow-copied splice inherits valid caches.
    # `__evaluate_post_process` is a @cached_property that re-runs the simulation; if we let it
    # compute lazily on the spliced flight (where solution and t_initial have been overwritten),
    # it produces an empty array and ax/ay/az/acceleration break with an IndexError.
    _warm_matched_flight_cache(matched_flight)

    # Build the spliced flight for downstream per-flight plots/KML/trajectory_3d (these need nominal-until-t_match + matched-after).
    # The original matched_flight stays untouched and is attached as `_raw_matched` so the comparison plots can still
    # access the pure post-splice matched simulation.
    spliced_flight = copy.copy(matched_flight)
    splice_matched_flight_in_place(spliced_flight, nominal_flight, t_match)
    spliced_flight._raw_matched = matched_flight

    return spliced_flight


def _warm_matched_flight_cache(flight: Flight):
    """Touch the key derived attributes so RocketPy populates its lazy caches (post-process array + ax/ay/az/etc.) before we shallow-copy."""
    # Accessing .acceleration cascades through ax/ay/az which compute __evaluate_post_process
    for attr in ("acceleration", "alpha1", "alpha2", "alpha3", "R1", "R2", "R3", "M1", "M2", "M3", "net_thrust"):
        try:
            getattr(flight, attr)
        except Exception:
            pass



def splice_matched_flight_in_place(spliced_flight: Flight, nominal_flight: Flight, splice_time):
    """
    Mutate a (copy of a) matched Flight so its time-domain Function attributes, `solution`, and `time` array
    cover the spliced trajectory: nominal samples for t <= splice_time + matched samples for t >= splice_time.
    Also copies boundary scalars (t_initial, out_of_rail_*) from nominal so plots/exports see a coherent launch.
    """
    # Splice the solution array (the primary state RocketPy plots like trajectory_3d read from)
    nominal_solution = np.asarray(nominal_flight.solution)
    matched_solution = np.asarray(spliced_flight.solution)
    nominal_part = nominal_solution[nominal_solution[:, 0] <= splice_time]
    matched_part = matched_solution[matched_solution[:, 0] >= splice_time]
    combined_solution = np.vstack([nominal_part, matched_part])
    # Preserve whichever storage form RocketPy expects (list of state vectors vs ndarray)
    spliced_flight.solution = combined_solution.tolist() if isinstance(spliced_flight.solution, list) else combined_solution

    # Splice the time array (used by KML export and CustomPlots.get_time_samples)
    nominal_time = np.asarray(nominal_flight.time)
    matched_time = np.asarray(spliced_flight.time)
    spliced_flight.time = np.concatenate([nominal_time[nominal_time <= splice_time], matched_time[matched_time >= splice_time]])

    # Override boundary scalars; matched started mid-air, so its rail-exit/initial-time values are bogus
    spliced_flight.t_initial = nominal_flight.t_initial
    for attr in ("out_of_rail_time", "out_of_rail_velocity"):
        if hasattr(nominal_flight, attr):
            try:
                setattr(spliced_flight, attr, getattr(nominal_flight, attr))
            except Exception:
                pass

    # Splice every cached time-domain funcify_method Function attribute (Function objects RocketPy lazily builds from .solution).
    # Suppress RocketPy's "no rail phase" UserWarning that fires when matched (mid-air start) computes rail-button forces;
    # the splice still produces correct values (nominal during on-rail phase + zeros after rail exit).
    flight_cls = type(spliced_flight)
    with warnings.catch_warnings():
        warnings.simplefilter("ignore", category=UserWarning)
        for attr_name in dir(flight_cls):
            if attr_name.startswith("_"):
                continue
            try:
                cls_attr = inspect.getattr_static(flight_cls, attr_name)
            except AttributeError:
                continue
            if cls_attr.__class__.__name__ != "funcify_method_decorator":
                continue
            try:
                nominal_func = getattr(nominal_flight, attr_name)
                matched_func = getattr(spliced_flight, attr_name)
            except Exception:
                continue
            if not (isinstance(nominal_func, Function) and isinstance(matched_func, Function)):
                continue
            # Skip non-time-domain Functions (e.g. *_frequency_response which are frequency-domain)
            function_inputs = getattr(nominal_func, "__inputs__", []) or []
            if not function_inputs or "time" not in str(function_inputs[0]).lower():
                continue
            try:
                # Combine two Functions into one trace: nominal samples up to splice_time, matched samples from splice_time onwards.
                nominal_source = nominal_func.source
                matched_source = matched_func.source
                nominal_part = nominal_source[nominal_source[:, 0] <= splice_time]
                matched_part = matched_source[matched_source[:, 0] >= splice_time]
                spliced_flight.__dict__[attr_name] = Function(np.vstack([nominal_part, matched_part]))
            except Exception:
                pass


def latlon_to_local_xy(lat, lon, ref_lat, ref_lon):
    """Convert (lat, lon) to local Cartesian (x_east, y_north) meters around (ref_lat, ref_lon)."""
    x_east, y_north, _ = pm.geodetic2enu(lat, lon, 0.0, ref_lat, ref_lon, 0.0)
    return float(x_east), float(y_north)



def compute_flight_computer_impacts(data, env):
    """Compute the impact (lat, lon, x, y) of each GNSS-equipped flight computer relative to the launch site."""
    # Source-specific labels and colors; sources without GNSS (e.g. Altimax) are skipped automatically
    source_styles = {
        "cats_vega": ("CATS Vega impact", "magenta"),
        "rcu": ("SRAD impact", "darkgreen"),
    }

    impacts = []
    for source_key, (label, color) in source_styles.items():
        if source_key not in data:
            continue

        flight_computer = data[source_key]
        if not hasattr(flight_computer, "latitude") or not hasattr(flight_computer, "longitude"):
            print(f"[reanalysis] {label}: no valid GNSS sample, skipping impact marker.")
            continue

        lat_data = flight_computer.latitude.source
        lon_data = flight_computer.longitude.source

        # Both samples need to be finite (and non-zero — GNSS dropouts often log as 0)
        valid = np.isfinite(lat_data[:, 1]) & np.isfinite(lon_data[:, 1]) & (lat_data[:, 1] != 0) & (lon_data[:, 1] != 0)
        if not valid.any():
            print(f"[reanalysis] {label}: no valid GNSS sample, skipping impact marker.")
            continue

        last_index = int(np.where(valid)[0][-1])
        lat = float(lat_data[last_index, 1])
        lon = float(lon_data[last_index, 1])
        x, y = latlon_to_local_xy(lat, lon, env.latitude, env.longitude)
        impacts.append({"label": label, "lat": lat, "lon": lon, "x": x, "y": y, "color": color})

    return impacts


def compute_gnss_3d_traces(data, env):
    """Pre-compute 3D-plot-ready GNSS ground tracks (x_east, y_north, altitude_AGL) for each available flight computer. Returns a dict keyed by source name."""
    traces = {}

    for source_key, label, color in (
        ("cats_vega", "CATS Vega (GNSS)", "magenta"),
        ("rcu", "SRAD (GNSS)", "darkgreen"),
    ):
        if source_key not in data:
            continue

        flight_computer = data[source_key]
        lat_source = flight_computer.latitude.source
        lon_source = flight_computer.longitude.source
        times, lats, lons = _valid_gnss_samples(lat_source[:, 0], lat_source[:, 1], lon_source[:, 1])
        if len(times) == 0:
            continue

        x_east = np.empty(len(times))
        y_north = np.empty(len(times))
        for i, (lat, lon) in enumerate(zip(lats, lons)):
            x_east[i], y_north[i] = latlon_to_local_xy(lat, lon, env.latitude, env.longitude)

        # Altitude pulled from the same computer's barometric trace, interpolated at the GNSS timestamps
        altitudes = np.array([float(flight_computer.altitude(t)) for t in times])

        traces[source_key] = {
            "name": label,
            "color": color,
            "x": x_east.tolist(),
            "y": y_north.tolist(),
            "z": altitudes.tolist(),
        }

    return traces


def find_rail_exit_time(flight_computer, effective_1rl):
    """Find the timestamp when the rocket first cleared the rail in a flight-computer altitude trace, or None if never reached."""
    # FlightDataImporter.altitude.source is an Nx2 array of [time, altitude]
    altitude_data = flight_computer.altitude.source
    times = altitude_data[:, 0]
    altitudes = altitude_data[:, 1]

    # Use the altitude at t=0 (liftoff in the computer's clock) as the baseline
    baseline_altitude = float(flight_computer.altitude(0.0))
    target_altitude = baseline_altitude + effective_1rl
    # print(f"target_altitude={target_altitude}, baseline_altitude={baseline_altitude}")

    above_target = altitudes >= target_altitude
    if not above_target.any():
        return None

    # Pick the first sample at or above the target altitude (nearest one)
    return float(times[int(np.argmax(above_target))])


# =============================================================================
# Individual metric comparisons
# =============================================================================

def compare_altitude(flight: Flight, data, matched_flight: Flight = None):
    """Compare apogee and altitude-over-time between the simulation(s) and whichever flight computers are loaded."""
    traces = [(flight.altitude, "RocketPy (nominal)" if matched_flight is not None else "RocketPy")]
    if matched_flight is not None:
        traces.append((matched_flight.altitude, "RocketPy (matched)"))
    apogee_refs = []

    if "cats_vega" in data:
        traces.append((data["cats_vega"].altitude, "CATS Vega"))
        apogee_refs.append(data["cats_vega"].altitude.max)

    if "altimax" in data:
        traces.append((data["altimax"].altitude, "Altimax"))

    if "rcu" in data:
        # Drop NaN/inf rows from the SRAD pressure trace before deriving altitude
        pressure = data["rcu"].pressure
        altitude_srad = Function(pressure[np.isfinite(pressure[:, 0]) & np.isfinite(pressure[:, 1])])
        traces.append((pressure_to_altitude(altitude_srad), "SRAD"))
        # SRAD apogee from its minimum pressure reading
        apogee_refs.append(pressure_to_altitude(altitude_srad.min))

    # "Actual" apogee = average of available barometric references (CATS and/or SRAD-derived)
    if apogee_refs:
        apogee_actual = sum(apogee_refs) / len(apogee_refs)
        print(f"Actual apogee: {apogee_actual:.2f} m (AGL)")
        if matched_flight is not None:
            _print_simulated_error_line(
                "apogee (nominal)",
                apogee_actual,
                flight.apogee - flight.env.elevation,
                unit="m (AGL)",
                unit_error="m",
            )
            _print_simulated_error_line(
                "apogee (matched)",
                apogee_actual,
                matched_flight.apogee - matched_flight.env.elevation,
                unit="m (AGL)",
                unit_error="m",
            )
        else:
            _print_simulated_error_line(
                "apogee",
                apogee_actual,
                flight.apogee - flight.env.elevation,
                unit="m (AGL)",
                unit_error="m",
            )

    Function.compare_plots(
        traces,
        title="Altitude Comparison",
        xlabel="Time (s)",
        ylabel="Altitude (m)",
    )


def compare_gnss(flight, data, matched_flight=None):
    """Compare latitude and longitude tracks against whichever GNSS-equipped flight computers are loaded."""
    # Skip the whole section if no external GNSS source is loaded
    if "cats_vega" not in data and "rcu" not in data:
        return

    nominal_label = "RocketPy (nominal)" if matched_flight is not None else "RocketPy"
    lat_traces = [(flight.latitude, nominal_label)]
    lon_traces = [(flight.longitude, nominal_label)]

    if matched_flight is not None:
        lat_traces.append((matched_flight.latitude, "RocketPy (matched)"))
        lon_traces.append((matched_flight.longitude, "RocketPy (matched)"))

    if "rcu" in data:
        lat_traces.append((data["rcu"].latitude, "SRAD"))
        lon_traces.append((data["rcu"].longitude, "SRAD"))

    if "cats_vega" in data:
        lat_traces.append((data["cats_vega"].latitude, "CATS"))
        lon_traces.append((data["cats_vega"].longitude, "CATS"))

    Function.compare_plots(lat_traces, title="Latitude Comparison", xlabel="Time (s)", ylabel="Latitude (deg)")
    Function.compare_plots(lon_traces, title="Longitude Comparison", xlabel="Time (s)", ylabel="Longitude (deg)")


def plot_cats_attitude_angle(project_path, custom_plots, event_markers):
    """
    Plot the CATS Vega attitude angle from vertical over time.
    """
    orientation_file_path = project_path / CATS_FOLDER / "orientationInfo.csv"
    
    # -------------------------------------------------------------------------
    # Load and compute attitude angle
    # -------------------------------------------------------------------------
    
    orientation_data = pd.read_csv(orientation_file_path)
    time_samples = orientation_data["ts"].to_numpy()
    
    # Load the CATS quaternion columns.
    # Shape: one row per timestamp, four columns per quaternion.
    # CATS component order is: qx, qy, qz, qw
    quaternions_xyzw = orientation_data[["q0_estimated", "q1_estimated", "q2_estimated", "q3_estimated"]].to_numpy()

    # Normalize because the CATS values appear to be scaled by 10.
    quaternion_norms = np.linalg.norm(quaternions_xyzw, axis=1, keepdims=True)
    quaternions_xyzw = quaternions_xyzw / quaternion_norms

    # From the first samples, -Z is the likely rocket longitudinal axis in the CATS/body frame.
    rocket_axis_body = np.array([0.0, 0.0, -1.0])

    # World vertical direction.
    vertical_world = np.array([0.0, 0.0, 1.0])

    rotation = Rotation.from_quat(quaternions_xyzw)
    rocket_axis_world = rotation.apply(rocket_axis_body)

    cosine_attitude_angle = rocket_axis_world @ vertical_world
    cosine_attitude_angle = np.clip(cosine_attitude_angle, -1.0, 1.0)

    attitude_angle_from_vertical_deg = np.degrees(np.arccos(cosine_attitude_angle))
    attitude_angle_from_horizontal_deg = 90.0 - attitude_angle_from_vertical_deg

    # -------------------------------------------------------------------------
    # Plot
    # -------------------------------------------------------------------------

    time_start = 0
    time_end = float(time_samples[-1])

    # GNSS-derived compass heading sampled at the same timestamps, plotted on the secondary y-axis
    heading_deg = compute_cats_gnss_heading_at_times(project_path, time_samples, window_half_width=0.5)

    traces = [
        {
            "y": attitude_angle_from_horizontal_deg,
            "name": "CATS Vega attitude angle [°]",
            "hovertemplate": "Attitude angle: %{y:.2f} °<extra></extra>",
            "line": {"color": "firebrick"},
        },
        {
            "y": heading_deg,
            "name": "GNSS heading [°]",
            "hovertemplate": "Heading: %{y:.2f} °<extra></extra>",
            "line": {"color": "royalblue"},
            "yaxis": "y2",
        },
    ]

    custom_plots.create_plotly_plot(
        title="CATS Vega attitude angle from horizontal & GNSS heading",
        time_samples=time_samples,
        time_start=time_start,
        time_end=time_end,
        traces=traces,
        yaxis_title="Attitude angle from horizontal [°]",
        yaxis2_title="Heading (compass, °)",
        width=900,
        height=500,
        event_markers=event_markers,
    )
    

def compare_pressure(flight, data, matched_flight=None):
    """Compare the minimum (apogee) pressure and the full pressure trace across whichever sources are loaded."""
    # Prefer Altimax for the reference (pre-filtered), then CATS, then SRAD (converted from hPa)
    if "altimax" in data:
        actual_pressure = data["altimax"].pressure.min
    elif "cats_vega" in data:
        actual_pressure = data["cats_vega"].pressure.min
    elif "rcu" in data:
        actual_pressure = data["rcu"].pressure.min * 100
    else:
        actual_pressure = None

    if actual_pressure is not None:
        print(f"Actual min pressure: {actual_pressure:.2f} Pa")
        if matched_flight is not None:
            _print_simulated_error_line("min pressure (nominal)", actual_pressure, flight.pressure.min, "Pa")
            _print_simulated_error_line("min pressure (matched)", actual_pressure, matched_flight.pressure.min, "Pa")
        else:
            _print_simulated_error_line("min pressure", actual_pressure, flight.pressure.min, "Pa")

    traces = [(flight.pressure, "RocketPy (nominal)" if matched_flight is not None else "RocketPy")]
    if matched_flight is not None:
        traces.append((matched_flight.pressure, "RocketPy (matched)"))

    if "cats_vega" in data:
        traces.append((data["cats_vega"].pressure, "CATS vega"))

    if "altimax" in data:
        traces.append((data["altimax"].pressure, "Altimax"))

    if "rcu" in data:
        traces.append((data["rcu"].pressure * 100, "SRAD"))     # hPa to Pa

    Function.compare_plots(traces, title="Pressure Comparison", xlabel="Time (s)", ylabel="Pressure (Pa)")


def compare_acceleration(flight: Flight, data, motor: Motor, matched_flight: Flight = None):
    """Compare peak burn-phase acceleration and the vertical acceleration trace from whichever sources are loaded."""

    if "altimax" in data:
        acceleration_actual = data["altimax"].acceleration.crop([(0, motor.burn_out_time + 10)]).max
        acceleration_simulated = flight.acceleration.crop([(0, motor.burn_out_time + 10)]).max
        print(f"Actual max acceleration during burn: {acceleration_actual:.2f} m/s2")
        if matched_flight is not None:
            matched_acceleration_simulated = matched_flight.acceleration.crop([(0, motor.burn_out_time + 10)]).max
            _print_simulated_error_line(
                "max acceleration during burn (nominal)",
                acceleration_actual,
                acceleration_simulated,
                "m/s2",
            )
            _print_simulated_error_line(
                "max acceleration during burn (matched)",
                acceleration_actual,
                matched_acceleration_simulated,
                "m/s2",
            )
        else:
            _print_simulated_error_line(
                "max acceleration during burn",
                acceleration_actual,
                acceleration_simulated,
                "m/s2",
            )

    traces = [(flight.ay.crop([(0, flight.t_final)]), "RocketPy (nominal)" if matched_flight is not None else "RocketPy")]
    if matched_flight is not None:
        traces.append((matched_flight.ay.crop([(0, matched_flight.t_final)]), "RocketPy (matched)"))

    if "cats_vega" in data:
        # CATS reports in 0.1g units with inverted sign, so divide by 10 and flip
        time_end = float(np.asarray(data["cats_vega"].time)[-1])
        traces.append((data["cats_vega"].az.crop([(0, time_end)]) / 10 * (-1), "CATS Vega"))

    if "altimax" in data:
        # Altimax reports in 0.1g units with inverted sign as well
        time_end = float(np.asarray(data["altimax"].time)[-1])
        traces.append((data["altimax"].acceleration.crop([(0, time_end)]) / 10 * (-1), "Altimax"))

    if "rcu" in data:
        time_end = float(np.asarray(data["rcu"].time)[-1])
        traces.append((data["rcu"].accel_y.crop([(0, time_end)]), "SRAD Acceleration Y"))

    Function.compare_plots(
        traces,
        title="Acceleration Comparison",
        xlabel="Time (s)",
        ylabel="Vertical acceleration (m/s^2)",
    )


def compare_speed(flight, data, matched_flight=None):
    """Compare peak speed and the speed trace against whichever speed-reporting sources are loaded."""
    # Prefer Altimax for the reference (pre-filtered), fall back to CATS
    if "altimax" in data:
        speed_actual = data["altimax"].speed.max
    elif "cats_vega" in data:
        speed_actual = data["cats_vega"].speed.max
    else:
        speed_actual = None

    if speed_actual is not None:
        print(f"Actual max speed: {speed_actual:.2f} m/s")
        if matched_flight is not None:
            _print_simulated_error_line("max speed (nominal)", speed_actual, flight.speed.max, "m/s")
            _print_simulated_error_line("max speed (matched)", speed_actual, matched_flight.speed.max, "m/s")
        else:
            _print_simulated_error_line("max speed", speed_actual, flight.speed.max, "m/s")

    traces = [(flight.vz, "RocketPy (nominal)" if matched_flight is not None else "RocketPy")]
    if matched_flight is not None:
        traces.append((matched_flight.vz, "RocketPy (matched)"))

    if "cats_vega" in data:
        traces.append((data["cats_vega"].speed, "CATS Vega"))

    if "altimax" in data:
        traces.append((data["altimax"].speed, "Altimax"))

    Function.compare_plots(traces, title="Speed Comparison", xlabel="Time (s)", ylabel="Speed (m/s)")


def plot_vertical_motion_per_source(flight: Flight, data, motor: Motor, rocket, rocket_config, cats_markers=None):
    """Plot altitude / vertical velocity / vertical acceleration over time for each compatible flight computer."""
    # SRAD/RCU has no speed trace
    sources_in_order = []
    if "cats_vega" in data:
        sources_in_order.append(("cats_vega", "CATS VEGA"))
    if "altimax" in data:
        sources_in_order.append(("altimax", "ALTIMAX"))

    for source_key, name in sources_in_order:
        forecast = data[source_key]

        # Build per-source event markers: CATS gets its native event log + rail exit, Altimax gets only rail exit
        markers = []
        if source_key == "cats_vega" and cats_markers:
            markers.extend(cats_markers)
            print("Event markers from CATS Vega. 'Out Of Rail' derived.")
        rail_exit_time = find_rail_exit_time(forecast, flight.effective_1rl)
        if rail_exit_time is not None:
            markers.append((rail_exit_time, "Out Of Rail", "red"))

        custom_plots = CustomPlots(
            flight_forecast=forecast,
            motor=motor,
            plot_title=name,
            rocket=rocket,
            rocket_config=rocket_config,
        )
        time_end = float(np.asarray(forecast.time)[-1])
        custom_plots.plot_vertical_motion(time_interval=(0, time_end), event_markers=markers or None)


def compare_rail_exit_velocity(flight: Flight, data):
    """Compare the nominal sim rail exit velocity against CATS / Altimax observations."""
    print(f"Effective rail length: {flight.effective_1rl:.3f} m")

    for source_key, label in (("cats_vega", "CATS Vega"), ("altimax", "Altimax")):
        if source_key not in data:
            continue

        crossing_time = find_rail_exit_time(data[source_key], flight.effective_1rl)
        if crossing_time is None:
            print(f"  {label}: rail exit altitude never reached, skipping.")
            continue

        actual_velocity = float(data[source_key].speed(crossing_time))

        print(f"Actual rail exit velocity ({label}, at {crossing_time:.3f} s): {actual_velocity:.2f} m/s")
        _print_simulated_error_line("rail exit velocity (nominal)", actual_velocity, flight.out_of_rail_velocity, "m/s")
    print("\n")


# =============================================================================
# Entry
# =============================================================================

def _resolve_reanalysis_setup(constants, variations):
    """Shared preamble: returns (flights_by_env, reanalysis_envs, sources, project_path) or None when reanalysis should be skipped."""
    if any(key in variations for key in VARIATION_KEYS):
        print("No reanalysis for variations.")
        return None

    flights_by_env, _ = lookup("flights_by_env", constants, variations)

    reanalysis_envs = [name for name in flights_by_env if name.startswith("Reanalysis")]
    if not reanalysis_envs:
        return None

    try:
        sources, _ = lookup("reanalysis_sources", constants, variations)
    except KeyError:
        print("[reanalysis] No 'reanalysis.sources' configured, skipping.")
        return None

    if not sources:
        print("[reanalysis] 'reanalysis.sources' is empty, skipping.")
        return None

    return flights_by_env, reanalysis_envs, sources, constants["project_path"]


def build_reanalysis_artifacts(constants, variations):
    """
    Build the matched flight (if configured) and compute flight-computer impact markers, registering both so the display
    stage picks them up. Must run BEFORE outputs.run_notebook_display_mode. Silent except for status lines.
    """
    setup = _resolve_reanalysis_setup(constants, variations)
    if setup is None:
        return constants, variations
    flights_by_env, reanalysis_envs, sources, project_path = setup

    # Reanalysis.matched_flight.match_time controls whether a CATS-attitude-matched flight is also simulated; missing key = skip
    try:
        match_time, _ = lookup("reanalysis_matched_flight_match_time", constants, variations)
    except KeyError:
        match_time = None

    # Per-parachute overrides for the matched flight (e.g. main didn't deploy, drogue had different cd). Empty when none configured.
    parachute_overrides = _parachute_overrides_from_config(constants, variations)
    motor_overrides = _motor_overrides_from_config(constants, variations)

    flight_computer_impacts = []
    flight_computer_data_by_env = {}

    for env_name in reanalysis_envs:
        scenario_sets = flights_by_env[env_name]
        nominal_flight = scenario_sets[0]["nominal"]

        # Load once here, then register so the comparison stage can reuse the same imported data.
        data = load_flight_computer_data(project_path, sources)
        flight_computer_data_by_env[env_name] = data
        flight_computer_impacts.extend(compute_flight_computer_impacts(data, nominal_flight.env))

        if match_time is not None and "cats_vega" in sources:
            print(f"[reanalysis] Creating matched flight at t={match_time:.3f} s using CATS Vega attitude angle, velocity vector and GNSS bearing...")
            if motor_overrides:
                print(f"[reanalysis] applying motor overrides: {motor_overrides}")
            if parachute_overrides:
                print(f"[reanalysis] applying parachute overrides: {parachute_overrides}")
            matched_flight = create_matched_flight(
                nominal_flight,
                project_path,
                match_time,
                data=data,
                parachute_overrides=parachute_overrides,
                motor_overrides=motor_overrides,
            )
            scenario_sets[0]["matched"] = matched_flight
        elif match_time is not None:
            print("[reanalysis] 'matched_flight' requested but 'cats_vega' not in sources, skipping matched flight.")

        # Pre-compute GNSS 3D tracks so compare_trajectories can overlay them on the trajectory plot
        gnss_3d_traces = compute_gnss_3d_traces(data, nominal_flight.env)
        if gnss_3d_traces:
            constants, variations = register("gnss_3d_traces", gnss_3d_traces, constants, variations)

    constants, variations = register("flight_computer_impacts", flight_computer_impacts, constants, variations)
    constants, variations = register("reanalysis_flight_computer_data", flight_computer_data_by_env, constants, variations)
    return constants, variations


def run_reanalysis_comparison(constants, variations):
    """
    Compares nominal (and matched, if built) simulated flights per env against onboard flight-computer data.
    Should run AFTER build_reanalysis_artifacts (and typically AFTER outputs.run_notebook_display_mode).
    """
    setup = _resolve_reanalysis_setup(constants, variations)
    if setup is None:
        return constants, variations
    flights_by_env, reanalysis_envs, sources, project_path = setup

    try:
        flight_computer_data_by_env = lookup("reanalysis_flight_computer_data", constants, variations)[0]
    except KeyError:
        raise RuntimeError("run_reanalysis_comparison requires build_reanalysis_artifacts to run first.")

    event_markers = None

    rocket = lookup("rocket", constants, variations)[0]
    motor = lookup("motor", constants, variations)[0]
    rocket_config = {"total_length": lookup("rocket_length", constants, variations)[0]}

    for env_name in reanalysis_envs:
        scenario_sets = flights_by_env[env_name]
        nominal_flight = scenario_sets[0]["nominal"]
        # The scenario_set stores the spliced flight (for per-flight plots/KML); the comparison plots want the
        # raw matched-only simulation so the divergence is visible as a separate trace starting at t_match.
        spliced_flight = scenario_sets[0].get("matched")
        matched_flight = getattr(spliced_flight, "_raw_matched", None) if spliced_flight else None

        printmd(f"## Reanalysis comparison: {env_name}")
        data = flight_computer_data_by_env[env_name]

        compare_altitude(nominal_flight, data, matched_flight=matched_flight)
        compare_gnss(nominal_flight, data, matched_flight=matched_flight)

        if "cats_vega" in sources:
            # Rail-exit time is still needed for the attitude-plot marker below;
            # the CATS-derived initial heading itself is registered in build_reanalysis_artifacts.
            cats_rail_exit_time = find_rail_exit_time(data["cats_vega"], nominal_flight.effective_1rl)

            event_markers = cats_event_markers(project_path)
            custom_plots = CustomPlots(
                flight_forecast=nominal_flight,
                motor=motor,
                plot_title="CATS Vega",
                rocket=rocket,
                rocket_config=rocket_config,
            )
            attitude_markers = list(event_markers)
            if cats_rail_exit_time is not None:
                attitude_markers.append((cats_rail_exit_time, "Out Of Rail", "red"))
            plot_cats_attitude_angle(project_path, custom_plots, attitude_markers)

        compare_pressure(nominal_flight, data, matched_flight=matched_flight)
        compare_acceleration(nominal_flight, data, motor, matched_flight=matched_flight)
        compare_speed(nominal_flight, data, matched_flight=matched_flight)
        plot_vertical_motion_per_source(nominal_flight, data, motor, rocket, rocket_config, cats_markers=event_markers)
        compare_rail_exit_velocity(nominal_flight, data)

    return constants, variations
