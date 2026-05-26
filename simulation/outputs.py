"""
Output and post-processing helpers for the RocketPy simulation backend.
"""

import csv
from pathlib import Path

import nbformat
from nbconvert import HTMLExporter

from matplotlib.path import Path as MatplotlibPath
import numpy as np
import plotly.graph_objects as go
from rocketpy import CompareFlights
from rocketpy.simulation import FlightDataExporter
from rocketpy import Flight, Environment, Motor, Fins, Rocket

import simulation.utils as utils
from simulation.custom_print_and_plot_functions import CustomPlots, CustomPrints

SCENARIO_COLORS = {"nominal": "green", "no_main": "orange", "ballistic": "red", "matched": "purple", "payload": "blue"}

# =============================================================================
# Generic flight collection helpers
# =============================================================================

def collect_scenario_sets(constants, variations):
    """
    Return every `scenario_set` dict produced by the simulation as one flat list.

    Each `scenario_set` is a dict of `{scenario_name: Flight}` for one heading/inclination/env
    combination, with keys:
      - "nominal":   always present; the main flight with all parachutes.
      - "no_main":   only present when a drogue parachute is configured.
      - "ballistic": always present; the descent without any parachutes.

    Returns an empty list when no flights have been registered yet.
    """
    try:
        flights_by_env = utils.lookup("flights_by_env", constants, variations)[0]
    except KeyError:
        return []

    return [
        scenario_set
        for scenario_sets in flights_by_env.values()
        for scenario_set in scenario_sets
    ]


# =============================================================================
# Printing
# =============================================================================

def print_one_environment(env: Environment, title):
    utils.printmd(f"## {title}")  
    # env.prints.gravity_details()
    env.prints.launch_site_details()
    env.prints.atmospheric_model_details()
    env.prints.atmospheric_conditions()
    # env.prints.print_earth_details()


def print_one_motor(motor: Motor, type, inertia=None):
    if inertia:
        IxIy, Iz, = inertia
        print("Inertia of motor without propellant:")
        print(f"Ix={IxIy}, Iy={IxIy}, Iz={Iz}\n")
        
    motor.prints.nozzle_details()
    motor.prints.motor_details()
    
    if type == "solid":
        motor.prints.grain_details()


def print_one_fin_set(fin_set: Fins):
    # fin_set.prints.identity()
    fin_set.prints.geometry()
    # fin_set.prints.lift()
    # fin_set.plots.airfoil()
    # fin_set.plots.roll()
    # fin_set.plots.lift()


def print_one_rocket(rocket: Rocket, rocket_length_m):
    print(f"Rocket center of wet mass from tip: {(rocket_length_m - rocket.center_of_mass(0)) * 1000} mm")
    rocket.prints.inertia_details()
    # rocket.prints.rocket_geometrical_parameters()
    # rocket.prints.rocket_aerodynamics_quantities()
    rocket.prints.parachute_data()


def print_one_flight_with_custom_prints(flight: Flight):
    """
    Print flight summary.
    """
    custom_prints = CustomPrints(flight)

    flight.prints.launch_rail_conditions()
    flight.prints.out_of_rail_conditions()
    print(f"OpenRocket Rail Departure Velocity: {flight.speed(flight.out_of_rail_time + 0.01)} m/s")      # OpenRocket flags that event 0.01 s later
    print(f"Effective rail length: {flight.effective_1rl} m")
    custom_prints.apogee_conditions()
    # flight.prints.apogee_conditions()
    custom_prints.parachute_events()
    # flight.prints.events_registered()
    flight.prints.impact_conditions()
    custom_prints.impact_coordinates()
    # flight.prints.maximum_values()


def print_flight_section_heading(flight, scenario_name):
    """
    Print a readable section heading for one environment and scenario.
    """
    environment_name = flight.env.name if hasattr(flight.env, "name") else flight.name
    utils.printmd(f"## {environment_name} | {scenario_name}")


# =============================================================================
# Custom plots
# =============================================================================

def plot_one_motor(motor: Motor):
    motor.draw()
    motor.plots.thrust()
    # motor.plots.mass_flow_rate()
    # motor.plots.exhaust_velocity()
    motor.plots.total_mass()
    # motor.plots.propellant_mass()
    motor.plots.center_of_mass()
    # motor.plots.grain_inner_radius()
    # motor.plots.grain_height()
    # motor.plots.burn_rate()
    # motor.plots.burn_area()
    # motor.plots.Kn()
    motor.plots.inertia_tensor()


def plot_parachute_models(constants, variations):
    """
    Plot the parachute models with the same helper used in the Albatross notebook.
    """
    parachutes = utils.lookup("parachutes", constants, variations)[0]

    for parachute in parachutes.values():
        CustomPlots.plot_parachute_model(parachute)


def plot_one_rocket(rocket: Rocket):
    rocket.plots.draw()
    rocket.plots.total_mass()
    # rocket.plots.reduced_mass()
    rocket.plots.drag_curves()
    # rocket.plots.static_margin()
    # rocket.plots.stability_margin()
    rocket.plots.thrust_to_weight()


def plot_one_flight_with_custom_plots(constants, variations, flight, scenario_name=None):
    """
    Plot the same custom flight plots used in the Albatross notebook.
    """
    rocket = flight.rocket
    motor = flight.rocket.motor
    environment_name = flight.env.name if hasattr(flight.env, "name") else flight.name
    plot_label = environment_name

    if scenario_name is not None:
        plot_label = f"{environment_name} | {scenario_name}"

    custom_plots = CustomPlots(
        flight_forecast=flight,
        motor=motor,
        plot_title=plot_label,
        rocket=rocket,
        rocket_config={"total_length": utils.lookup("rocket_length", constants, variations)[0]},
    )

    custom_plots.plot_stability_and_cg_cp_position()
    # flight.plots.stability_and_control_data()
    custom_plots.plot_angle_of_attack_and_attitude_angle()
    custom_plots.plot_angular_velocity(transform_openrocket=False)
    custom_plots.plot_vertical_motion()


# =============================================================================
# Comparison plots
# =============================================================================

def add_plotly_buttons(
    figure: go.Figure,
    mode_trace_indices: dict[str, list[int]],
    always_visible_indices: list[int],
    zone_legend_flags: list[bool] | None = None,
):
    """
    Build Plotly update-menu buttons that toggle trace visibility by mode, and apply them to the figure.

    Args:
        figure: The Plotly figure to which the buttons will be added.
        mode_trace_indices: {mode_name: [list of trace indices]} - which traces to show per button
        always_visible_indices: trace indices visible and in the legend for every mode (launch rail, etc.)
        zone_legend_flags: one bool per zone trace - True means show in legend. Zone groups add one
            trace per polygon but only show the first; this list restores those original flags when a button fires.
    """
    buttons = []
    total_traces = len(figure.data)

    for mode_name, indices in mode_trace_indices.items():
        visibility = [False] * total_traces
        showlegend = [False] * total_traces

        for idx in always_visible_indices:
            visibility[idx] = True
            showlegend[idx] = True

        # Zone traces may suppress showlegend for duplicate polygon entries; restore original flags.
        if zone_legend_flags:
            for zone_idx, flag in enumerate(zone_legend_flags):
                showlegend[zone_idx] = flag

        for idx in indices:
            visibility[idx] = True
            showlegend[idx] = True

        buttons.append(dict(
            label=mode_name,
            method="update",
            args=[{"visible": visibility, "showlegend": showlegend}],
        ))

    figure.update_layout(
        updatemenus=[dict(
            type="buttons",
            direction="right",
            buttons=buttons,
            x=0.01,
            xanchor="left",
            y=1.02,
            yanchor="bottom",
            showactive=True,
            bgcolor="white",
            bordercolor="lightgray",
        )]
    )


def compare_trajectories(constants, variations, flights: list[Flight]):
    """
    Interactive 3D trajectory plot for the given flights, with optional GNSS overlays from reanalysis.
    When multiple environments are present, buttons group traces by environment.
    """
    figure = go.Figure()

    # Build environment groups before adding any traces so we can track per-environment indices.
    environment_flight_groups = {}
    env_names_ordered = []
    for flight in flights:
        env_name = flight.env.name if hasattr(flight.env, "name") else flight.name
        if env_name not in environment_flight_groups:
            environment_flight_groups[env_name] = {"environment": flight.env, "flights": []}
            env_names_ordered.append(env_name)
        environment_flight_groups[env_name]["flights"].append(flight)

    multiple_envs = len(env_names_ordered) > 1
    env_trace_indices = {name: [] for name in env_names_ordered}

    # -------------------------------------------------------------------------
    # Simulated flight traces: x, y in local meters from launch; altitude is AGL
    # -------------------------------------------------------------------------
    for flight in flights:
        env_name = flight.env.name if hasattr(flight.env, "name") else flight.name
        times = np.asarray(flight.time)
        # Match the scenario keyword inside the name so SCENARIO_COLORS picks the right color.
        color = next((c for scenario, c in SCENARIO_COLORS.items() if scenario in (getattr(flight, "name", "") or "")), None)
        idx = len(figure.data)
        figure.add_trace(go.Scatter3d(
            x=np.array([flight.x(t) for t in times]),
            y=np.array([flight.y(t) for t in times]),
            z=np.array([flight.altitude(t) for t in times]),
            mode="lines",
            name=flight.name,
            line=dict(color=color, width=3) if color else dict(width=3),
        ))
        env_trace_indices[env_name].append(idx)

    # -------------------------------------------------------------------------
    # Wind heading arrows, one cone column per environment
    # -------------------------------------------------------------------------
    for env_name, environment_group in environment_flight_groups.items():
        environment = environment_group["environment"]
        environment_flights = environment_group["flights"]

        max_trajectory_altitude = max(
            float(np.nanmax([flight.altitude(t) for t in np.asarray(flight.time)]))
            for flight in environment_flights
        )
        wind_altitude_samples = np.linspace(0.0, max_trajectory_altitude, 20)

        # RocketPy wind functions expect altitude ASL, so add environment.elevation.
        wind_u = np.array([environment.wind_velocity_x(z + environment.elevation) for z in wind_altitude_samples], dtype=float)
        wind_v = np.array([environment.wind_velocity_y(z + environment.elevation) for z in wind_altitude_samples], dtype=float)

        wind_speed = np.array([environment.wind_speed(z + environment.elevation) for z in wind_altitude_samples], dtype=float)
        wind_heading = np.array([environment.wind_heading(z + environment.elevation) for z in wind_altitude_samples], dtype=float)

        idx = len(figure.data)
        figure.add_trace(go.Cone(
            x=np.zeros_like(wind_altitude_samples),
            y=np.zeros_like(wind_altitude_samples),
            z=wind_altitude_samples,
            u=wind_u,
            v=wind_v,
            w=np.zeros_like(wind_altitude_samples),
            name=f"Wind heading: {env_name}",
            colorscale=[[0, "blue"], [1, "blue"]],
            showscale=False,
            showlegend=True,
            sizemode="scaled",
            sizeref=2.0,
            anchor="tail",
            customdata=np.column_stack([wind_speed, wind_heading]),
            hovertemplate=(
                "<b>Wind</b><br>"
                "altitude: %{z:.0f} m AGL<br>"
                "speed: %{customdata[0]:.1f} m/s<br>"
                "heading: %{customdata[1]:.0f}°"
                "<extra></extra>"
            ),
        ))
        env_trace_indices[env_name].append(idx)

    # -------------------------------------------------------------------------
    # GNSS trajectories: follow the environment button when trace carries an "env" key
    # -------------------------------------------------------------------------
    try:
        gnss_traces = utils.lookup("gnss_3d_traces", constants, variations)[0]
    except KeyError:
        gnss_traces = None

    always_visible_indices = []

    if gnss_traces:
        for trace in gnss_traces.values():
            idx = len(figure.data)
            figure.add_trace(go.Scatter3d(
                x=trace["x"], y=trace["y"], z=trace["z"],
                mode="lines+markers",
                name=trace["name"],
                line=dict(color=trace["color"], width=3),
                marker=dict(size=2),
            ))
            gnss_env = trace.get("env")
            if gnss_env and gnss_env in env_trace_indices:
                env_trace_indices[gnss_env].append(idx)
            else:
                always_visible_indices.append(idx)

    # -------------------------------------------------------------------------
    # Launch-rail reference marker (always visible)
    # -------------------------------------------------------------------------
    launch_idx = len(figure.data)
    figure.add_trace(go.Scatter3d(
        x=[0], y=[0], z=[0],
        mode="markers",
        marker=dict(color="black", symbol="x", size=4),
        name="Launch",
    ))
    always_visible_indices.append(launch_idx)

    # -------------------------------------------------------------------------
    # Environment-grouping buttons (only when multiple environments are present)
    # -------------------------------------------------------------------------
    if multiple_envs:
        mode_trace_indices = {}
        for env_name in env_names_ordered:
            mode_trace_indices[env_name] = env_trace_indices[env_name]

        # The first button is active by default; hide every trace that doesn't belong to it.
        first_mode_indices = set(next(iter(mode_trace_indices.values())))
        all_dynamic_indices = {idx for indices in env_trace_indices.values() for idx in indices}
        for idx in all_dynamic_indices - first_mode_indices:
            figure.data[idx].visible = False

        add_plotly_buttons(figure, mode_trace_indices, always_visible_indices)

    figure.update_layout(
        title="3D Trajectory Comparison",
        scene=dict(
            xaxis_title="East / West [m]",
            yaxis_title="North / South [m]",
            zaxis_title="Altitude AGL [m]",
            aspectmode="cube",
            camera=dict(
                # x rotates camera east with pos values
                # y rotates camera north with pos values
                # z controls elevation: bigger = looking down at a steeper angle
                eye=dict(x=0.1, y=-2.3, z=0.2),
            ),
        ),
        width=900,
        height=700,
    )

    figure.show(renderer="notebook")


# =============================================================================
# Exports
# =============================================================================

def export_all_kml(constants, variations):
    """
    Export KML files for all flights. 
    Only executed if we not vary flights or only vary those from `VARIATION_KEYS_WITH_FLIGHT_PLOTS`.
    """
    exported_files = []
    project_path = constants["project_path"]
    utils.ensure_project_folders(project_path)

    for i, scenario_set in enumerate(collect_scenario_sets(constants, variations), start=1):
        for flight in scenario_set.values():
            file_name = Path(f"flight_{i}_{flight.filename_label}.kml")
            file_path = project_path / "trajectory_kml" / file_name
            FlightDataExporter(flight).export_kml(file_name=file_path, altitude_mode="relativetoground")
            exported_files.append(file_path)

    return exported_files


def export_all_trajectory_csv(constants, variations):
    """
    Export one CSV per flight with columns latitude, longitude, altitude (m AGL) sampled at every ODE time step.
    By request for WARR, maybe useful, otherwise remove later.
    """
    exported_files = []
    project_path = constants["project_path"]
    utils.ensure_project_folders(project_path)

    for i, scenario_set in enumerate(collect_scenario_sets(constants, variations), start=1):
        for scenario_name, flight in scenario_set.items():
            times = flight.latitude.source[:, 0]

            file_name = Path(f"flight_{i}_{scenario_name}_{flight.filename_label}.csv")
            file_path = project_path / "trajectory_csv" / file_name

            with open(file_path, "w", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(["time", "latitude", "longitude", "altitude"])

                elevation = flight.env.elevation

                for t in times:
                    writer.writerow([t, flight.latitude(t), flight.longitude(t), flight.z(t) - elevation])  # altitude above ground level

            exported_files.append(file_path)

    return exported_files


def export_notebook_to_html(notebook_path, output_path=None):
    """
    Export a Jupyter notebook with its current outputs to a self-contained HTML file.
    Save the notebook first so the file on disk reflects the current state.
    """
    notebook_path = Path(notebook_path)

    if output_path is None:
        output_path = notebook_path.with_suffix(".html")

    output_path = Path(output_path)

    nb = nbformat.read(notebook_path, as_version=4)
    html_body, _ = HTMLExporter().from_notebook_node(nb)
    output_path.write_text(html_body, encoding="utf-8")

    print(f"Notebook exported to: {output_path}")
    return output_path


# =============================================================================
# Safety mode helpers
# =============================================================================

def should_create_flight_plots(variations):
    """
    Return `True` when variations is empty or no VARIATION_KEYS_WITHOUT_FLIGHT_PLOTS are present in variations.
    This prevents too many flight plots being created.
    """
    if not variations:
        return True

    for var_key in utils.VARIATION_KEYS_WITHOUT_FLIGHT_PLOTS:
        if variations.get(var_key):
            return False

    return True


def print_scan_values(constants, variations):
    """
    Print the heading and inclination values that are used in scan mode.
    """
    headings = utils.ensure_list(utils.lookup("flight_heading", constants, variations)[0])
    inclinations = utils.ensure_list(utils.lookup("flight_inclination", constants, variations)[0])

    utils.printmd("## Configured flight scan")
    print(f"Headings: {headings}")
    print(f"Inclinations: {inclinations}")


# =============================================================================
# Safety printing
# =============================================================================

def print_configurations(title, configurations):
    """
    Print configurations grouped by heading, then environment.
    """
    utils.printmd(f"## {title}")

    if not configurations:
        print("None")
        return

    grouped = {}

    for configuration in configurations:
        if len(configuration) == 3:
            environment, heading, inclination = configuration
        else:
            heading, inclination = configuration
            environment = "all environments"

        grouped.setdefault(heading, {}).setdefault(environment, set()).add(inclination)

    for heading, environments in sorted(grouped.items()):
        print(f"\nHeading {heading}:")

        for environment, inclinations in sorted(environments.items()):
            print(f"- {environment}: inclinations {sorted(inclinations)}")


def print_safe_flight_details(constants, variations):
    """
    Print lat/lon landing coordinates for every safe flight.
    """
    utils.printmd("## Safe flight details")

    # Collect all safe flights with their scenario type label.
    type_keys = [
        ("nominal",  "safe_rocket_nominal"),
        ("no_main",  "safe_rocket_no_main"),
        ("ballistic","safe_rocket_ballistic"),
        ("matched",  "safe_rocket_matched"),
        ("payload",  "safe_payload"),
    ]

    # Build: {heading: {inclination: {type: [(env_name, lat, lon)]}}}
    grouped = {}

    for type_label, key in type_keys:
        flights = get_registered_flights(constants, variations, key)

        for flight in flights:
            heading = flight.heading
            inclination = flight.inclination
            env_name = flight.env.name if hasattr(flight.env, "name") else flight.name
            lat = flight.latitude(flight.t_final)
            lon = flight.longitude(flight.t_final)
            grouped.setdefault(heading, {}).setdefault(inclination, {}).setdefault(type_label, []).append((env_name, lat, lon))

    if not grouped:
        print("None")
        return

    for heading in sorted(grouped):
        print(f"\nHeading {heading}°:")

        for inclination in sorted(grouped[heading]):
            print(f"  Inclination {inclination}°:")

            for type_label in ("nominal", "no_main", "ballistic", "payload"):
                entries = grouped[heading][inclination].get(type_label)

                if not entries:
                    continue

                print(f"    {type_label}:")

                for env_name, lat, lon in sorted(entries, key=lambda item: item[0]):
                    print(f"      {env_name}: lat={lat}°, lon={lon}°")


def print_unsafe_details(unsafe_details):
    """
    Print grouped details for every unsafe heading/inclination/environment combination.
    """
    utils.printmd("## Unsafe flight details")

    if not unsafe_details:
        print("None")
        return

    grouped = {}

    for detail in unsafe_details:
        heading = detail["heading"]
        environment = detail["environment"]
        grouped.setdefault(heading, {}).setdefault(environment, []).append(detail)

    for heading, environments in sorted(grouped.items()):
        print(f"\nHeading {heading}:")

        for environment, details in sorted(environments.items()):
            print(f"- {environment}:")

            for detail in sorted(details, key=lambda item: item["inclination"]):
                scenarios = ", ".join(detail["unsafe_scenarios"])
                print(f"  inclination {detail['inclination']}: {scenarios}")


# =============================================================================
# Safety calculations
# =============================================================================

def get_registered_flights(constants, variations, key):
    """
    Return a registered flight list or an empty list when the key was registered empty.
    """
    try:
        return utils.ensure_list(utils.lookup(key, constants, variations)[0])
    except KeyError:
        return []


def scenario_config_key(flight):
    """
    Identifier for the configuration a flight belongs to: (env_name, heading, inclination).
    """
    env_name = flight.env.name if hasattr(flight.env, "name") else flight.name
    return (env_name, flight.heading, flight.inclination)


def flight_lands_in_zone(flight, zones):
    """
    Return True when the flight's impact point lies inside any of the supplied zone polygons.
    """
    return any(
        MatplotlibPath(zone_coordinates).contains_point((flight.x_impact, flight.y_impact))
        for zone_coordinates in zones.values()
    )


def split_flights_by_scenario(scenario_sets):
    """
    Split scenario-set dictionaries into nominal, no-main, ballistic, and matched flight lists.
    """
    nominal_flights = []
    no_main_flights = []
    ballistic_flights = []
    matched_flights = []

    for scenario_set in scenario_sets:
        if "nominal" in scenario_set:
            nominal_flights.append(scenario_set["nominal"])

        if "no_main" in scenario_set:
            no_main_flights.append(scenario_set["no_main"])

        if "ballistic" in scenario_set:
            ballistic_flights.append(scenario_set["ballistic"])

        if "matched" in scenario_set:
            matched_flights.append(scenario_set["matched"])

    return nominal_flights, no_main_flights, ballistic_flights, matched_flights


def register_safety_results(constants, variations, prefix, scenario_sets):
    """
    Register nominal, no-main, ballistic, matched, and configuration lists for a safety result group.
    """
    nominal_flights, no_main_flights, ballistic_flights, matched_flights = split_flights_by_scenario(scenario_sets)
    configurations = sorted({scenario_config_key(flight) for flight in nominal_flights})

    constants, variations = utils.register(f"{prefix}_rocket_nominal", nominal_flights, constants, variations)
    constants, variations = utils.register(f"{prefix}_rocket_no_main", no_main_flights, constants, variations)
    constants, variations = utils.register(f"{prefix}_rocket_ballistic", ballistic_flights, constants, variations)
    constants, variations = utils.register(f"{prefix}_rocket_matched", matched_flights, constants, variations)
    variations[f"{prefix}_configurations"] = configurations

    return constants, variations


def calculate_safe_flights(constants, variations, buffer_zones):
    """
    Calculate safe and unsafe flights, classified by heading.

    A heading is unsafe if any of the following caused a landing in a buffer zone:
        - scenario (nominal / no_main / ballistic / optional payload)
        - inclination
        - environment
    """
    payload_flights = get_registered_flights(constants, variations, "flight_payload")

    # Group payload flights by their (env, heading, inclination) for the diagnostic detail list.
    payloads_by_config = {}
    for payload_flight in payload_flights:
        payloads_by_config.setdefault(scenario_config_key(payload_flight), []).append(payload_flight)

    scenario_sets = collect_scenario_sets(constants, variations)

    # First pass: find which headings have any unsafe scenario, and record diagnostic details.
    unsafe_headings = set()
    unsafe_details = []

    for scenario_set in scenario_sets:
        nominal_flight = scenario_set["nominal"]
        unsafe_scenarios = []

        for scenario_name, flight in scenario_set.items():
            if flight_lands_in_zone(flight, buffer_zones):
                unsafe_scenarios.append(scenario_name)

        # Include payload flights for the same configuration; a single unsafe payload is enough.
        for payload_flight in payloads_by_config.get(scenario_config_key(nominal_flight), []):
            if flight_lands_in_zone(payload_flight, buffer_zones):
                unsafe_scenarios.append("payload")
                break

        if unsafe_scenarios:
            unsafe_headings.add(nominal_flight.heading)
            unsafe_details.append({
                "heading": nominal_flight.heading,
                "inclination": nominal_flight.inclination,
                "environment": nominal_flight.env.name if hasattr(nominal_flight.env, "name") else nominal_flight.name,
                "unsafe_scenarios": unsafe_scenarios,
            })

    # Second pass: every scenario_set with an unsafe heading is unsafe, even if its specific
    # (env, inclination) flights all landed outside the zones.
    safe_scenario_sets = [s for s in scenario_sets if s["nominal"].heading not in unsafe_headings]
    unsafe_scenario_sets = [s for s in scenario_sets if s["nominal"].heading in unsafe_headings]

    constants, variations = register_safety_results(constants, variations, "safe", safe_scenario_sets)
    constants, variations = register_safety_results(constants, variations, "unsafe", unsafe_scenario_sets)
    constants, variations = utils.register("unsafe_details", unsafe_details, constants, variations)

    # Payload flights inherit the heading-level classification.
    safe_payload_flights = [payload for payload in payload_flights if payload.heading not in unsafe_headings]
    unsafe_payload_flights = [payload for payload in payload_flights if payload.heading in unsafe_headings]
    constants, variations = utils.register("safe_payload", safe_payload_flights, constants, variations)
    constants, variations = utils.register("unsafe_payload", unsafe_payload_flights, constants, variations)

    print_configurations("Safe Configurations by Heading", utils.lookup("safe_configurations", constants, variations)[0])
    print_safe_flight_details(constants, variations)
    print_configurations("Unsafe Configurations by Heading", utils.lookup("unsafe_configurations", constants, variations)[0])
    print_unsafe_details(unsafe_details)

    return constants, variations


# =============================================================================
# Safety plots
# =============================================================================
def build_safe_unsafe_flight_groups(constants, variations):
    """
    Build the safe and unsafe flight-group dicts from the registered safety lists.
    """
    safe_flight_groups = {
        "rocket_nominal": (get_registered_flights(constants, variations, "safe_rocket_nominal"), SCENARIO_COLORS["nominal"]),
        "rocket_no_main": (get_registered_flights(constants, variations, "safe_rocket_no_main"), SCENARIO_COLORS["no_main"]),
        "rocket_ballistic": (get_registered_flights(constants, variations, "safe_rocket_ballistic"), SCENARIO_COLORS["ballistic"]),
        "rocket_matched": (get_registered_flights(constants, variations, "safe_rocket_matched"), SCENARIO_COLORS["matched"]),
    }
    unsafe_flight_groups = {
        "rocket_nominal": (get_registered_flights(constants, variations, "unsafe_rocket_nominal"), SCENARIO_COLORS["nominal"]),
        "rocket_no_main": (get_registered_flights(constants, variations, "unsafe_rocket_no_main"), SCENARIO_COLORS["no_main"]),
        "rocket_ballistic": (get_registered_flights(constants, variations, "unsafe_rocket_ballistic"), SCENARIO_COLORS["ballistic"]),
        "rocket_matched": (get_registered_flights(constants, variations, "unsafe_rocket_matched"), SCENARIO_COLORS["matched"]),
    }

    safe_payload_flights = get_registered_flights(constants, variations, "safe_payload")
    unsafe_payload_flights = get_registered_flights(constants, variations, "unsafe_payload")
    if safe_payload_flights or unsafe_payload_flights:
        safe_flight_groups["payload_nominal"] = (safe_payload_flights, SCENARIO_COLORS["payload"])
        unsafe_flight_groups["payload_nominal"] = (unsafe_payload_flights, SCENARIO_COLORS["payload"])

    return safe_flight_groups, unsafe_flight_groups


def build_safety_by_config(constants, variations):
    """
    Build a {(environment, heading, inclination): "safe"|"unsafe"} map from the registered classification.
    """
    safety_by_config = {}

    for configuration in utils.lookup("safe_configurations", constants, variations)[0]:
        safety_by_config[configuration] = "safe"

    for configuration in utils.lookup("unsafe_configurations", constants, variations)[0]:
        safety_by_config[configuration] = "unsafe"

    return safety_by_config


# =============================================================================
# Landing position plots
# =============================================================================

def plot_zones(figure, zones, label, color, alpha=0.3):
    """
    Plot exclusion or buffer-zone polygons.
    """
    if not zones:
        return

    for index, (name, coordinates) in enumerate(zones.items()):
        x_values = [x for x, y in coordinates] + [coordinates[0][0]]
        y_values = [y for x, y in coordinates] + [coordinates[0][1]]

        # Only the first polygon in the group carries the legend entry so the legend stays compact.
        show_legend = index == 0

        figure.add_trace(go.Scatter(
            x=x_values,
            y=y_values,
            mode="lines",
            fill="toself",
            fillcolor=color,
            opacity=alpha,
            line=dict(color=color),
            name=label,
            legendgroup=label,
            showlegend=show_legend,
            hoverinfo="skip",
        ))

        # Label each zone at the centroid (mean of its vertices).
        centroid_x = sum(x for x, y in coordinates) / len(coordinates)
        centroid_y = sum(y for x, y in coordinates) / len(coordinates)
        figure.add_annotation(
            x=centroid_x,
            y=centroid_y,
            text=name,
            showarrow=False,
            font=dict(color=color, size=10),
        )


def add_compass_labels(figure):
    """
    Add bold N / E / S / W compass labels at the plot-area edges, aligned to the launch rail.
    """
    compass_points = [
        ("N", dict(x=0, y=0.99, xref="x", yref="y domain", xanchor="center", yanchor="top")),
        ("S", dict(x=0, y=0.01, xref="x", yref="y domain", xanchor="center", yanchor="bottom")),
        ("E", dict(x=0.99, y=0, xref="x domain", yref="y", xanchor="right", yanchor="middle")),
        ("W", dict(x=0.01, y=0, xref="x domain", yref="y", xanchor="left", yanchor="middle")),
    ]

    for text, position in compass_points:
        figure.add_annotation(
            text=text,
            showarrow=False,
            font=dict(size=14, color="black", family="Arial Black"),
            **position,
        )


def add_mode_flight_traces(figure, mode_name, flight_groups, visible, safety_by_config):
    """
    Add one Scatter trace per flight group for the given mode and return the trace index range.
    """
    start_index = len(figure.data)

    for label, (flights, color) in flight_groups.items():
        flight_list = utils.ensure_list(flights)

        if not flight_list:
            continue

        x_values = [flight.x_impact for flight in flight_list]
        y_values = [flight.y_impact for flight in flight_list]

        customdata = [
            [
                flight.heading,
                flight.inclination,
                flight.env.name if hasattr(flight.env, "name") else flight.name,
                safety_by_config.get(scenario_config_key(flight), "unknown"),
                flight.latitude(flight.t_final),
                flight.longitude(flight.t_final),
            ]
            for flight in flight_list
        ]

        figure.add_trace(go.Scatter(
            x=x_values,
            y=y_values,
            mode="markers",
            marker=dict(color=color, size=8),
            name=label,
            customdata=customdata,
            hovertemplate=(
                f"<b>{label}</b><br>"
                "safety: %{customdata[3]}<br>"
                "heading: %{customdata[0]}°<br>"
                "inclination: %{customdata[1]}°<br>"
                "env: %{customdata[2]}<br>"
                "impact: (%{x:.1f}, %{y:.1f}) m<br>"
                "lat: %{customdata[4]:.5f}°; lon: %{customdata[5]:.5f}°"
                "<extra></extra>"
            ),
            visible=visible,
            showlegend=visible,
        ))

    return start_index, len(figure.data)


def plot_landing_positions_with_modes(
    constants,
    exclusion_zones,
    buffer_zones,
    plot_name,
    mode_flight_groups=None,
    safety_by_config=None,
    zones_only=False,
    save_format="html",
    flight_computer_impacts=None,
):
    """
    Build a landing-position plot of the buffer/exclusion zones, with optional flight-mode buttons.

    When `zones_only` is True only the zones are drawn. Otherwise one button per entry of
    `mode_flight_groups` toggles which mode's flight markers are visible. The zones, the
    launch-rail marker, and any flight-computer impact markers stay visible across every mode.
    """
    project_path = constants["project_path"]
    utils.ensure_project_folders(project_path)
    figure = go.Figure()

    # Plot buffer zones first so they do not visually cover the red exclusion zones.
    plot_zones(figure, buffer_zones, label="Buffer zone", color="orange")
    plot_zones(figure, exclusion_zones, label="Exclusion zone", color="red")

    if not zones_only:
        zone_trace_count = len(figure.data)
        zone_legend_flags = [trace.showlegend is not False for trace in figure.data]

        # Add the flight traces grouped per mode and remember each mode's trace index range.
        mode_flight_groups = mode_flight_groups or {}
        mode_names = list(mode_flight_groups.keys())
        default_mode = mode_names[0] if mode_names else None
        mode_trace_ranges = {}

        for mode_name, flight_groups in mode_flight_groups.items():
            mode_trace_ranges[mode_name] = add_mode_flight_traces(
                figure,
                mode_name,
                flight_groups,
                visible=(mode_name == default_mode),
                safety_by_config=safety_by_config or {},
            )

        # Launch rail reference marker stays at the origin and stays visible in every mode.
        figure.add_trace(go.Scatter(
            x=[0],
            y=[0],
            mode="markers",
            marker=dict(color="black", symbol="x", size=10),
            name="launch rail",
            hoverinfo="skip",
        ))
        launch_trace_index = len(figure.data) - 1

        # Flight-computer impact markers (CATS/RCU); always visible across modes
        persistent_indices = [launch_trace_index]
        for impact in (flight_computer_impacts or []):
            figure.add_trace(go.Scatter(
                x=[impact["x"]],
                y=[impact["y"]],
                mode="markers",
                marker=dict(color=impact["color"], symbol="star", size=12),
                name=impact["label"],
                customdata=[[impact["lat"], impact["lon"]]],
                hovertemplate=(
                    f"<b>{impact['label']}</b><br>"
                    "impact: (%{x:.1f}, %{y:.1f}) m<br>"
                    "lat: %{customdata[0]:.5f}°; lon: %{customdata[1]:.5f}°"
                    "<extra></extra>"
                ),
            ))
            persistent_indices.append(len(figure.data) - 1)

        # Build one update button per mode; each sets a visibility and a showlegend array for all traces.
        always_visible_for_modes = list(range(zone_trace_count)) + persistent_indices
        mode_trace_indices = {
            mode_name: list(range(start, end))
            for mode_name, (start, end) in mode_trace_ranges.items()
        }
        add_plotly_buttons(figure, mode_trace_indices, always_visible_for_modes, zone_legend_flags)

    layout_kwargs = dict(
        title=dict(
            text="Landing Positions",
            x=0.5,
            xanchor="center",
        ),
        xaxis=dict(
            title="East / West distance from launch rail [m]",
            scaleanchor="y",
            scaleratio=1,
            showgrid=True,
            gridcolor="lightgray",
            zeroline=True,
            zerolinecolor="lightgray",
        ),
        yaxis=dict(
            title="North / South distance from launch rail [m]",
            showgrid=True,
            gridcolor="lightgray",
            zeroline=True,
            zerolinecolor="lightgray",
        ),
        width=900,
        height=700,
        hovermode="closest",
        paper_bgcolor="white",
        plot_bgcolor="white",
    )
    figure.update_layout(**layout_kwargs)

    add_compass_labels(figure)

    # Save in the requested format.
    if save_format == "png":
        figure.write_image(str(project_path / "plots" / f"{plot_name}.png"))
    else:
        figure.write_html(str(project_path / "plots" / f"{plot_name}.html"))

    figure.show(renderer="notebook")


# =============================================================================
# Notebook display mode
# =============================================================================

def run_notebook_display_mode(constants, variations, exclusion_zones, buffer_zones, OUTPUT_LEVEL=0):
    """
    Run the correct notebook display mode for the configured heading/inclination values.

    Both modes share the same All/Safe/Unsafe landing-positions plot. Single-flight mode adds a
    per-flight safety summary, trajectory comparison, and detailed per-scenario prints/plots.
    Scan mode adds the configured heading/inclination value list.
    """
    scenario_sets = collect_scenario_sets(constants, variations)
    payload_flights = get_registered_flights(constants, variations, "flight_payload")

    nominal_flights, no_main_flights, ballistic_flights, matched_flights = split_flights_by_scenario(scenario_sets)
    all_flight_groups = {
        "rocket_nominal": (nominal_flights, SCENARIO_COLORS["nominal"]),
        "rocket_no_main": (no_main_flights, SCENARIO_COLORS["no_main"]),
        "rocket_ballistic": (ballistic_flights, SCENARIO_COLORS["ballistic"]),
    }
    if matched_flights:
        all_flight_groups["rocket_matched"] = (matched_flights, SCENARIO_COLORS["matched"])
    if payload_flights:
        all_flight_groups["payload_nominal"] = (payload_flights, SCENARIO_COLORS["payload"])

    constants, variations = calculate_safe_flights(constants, variations, buffer_zones)

    safe_flight_groups, unsafe_flight_groups = build_safe_unsafe_flight_groups(constants, variations)
    safety_by_config = build_safety_by_config(constants, variations)

    # Only display the Safe/Unsafe buttons when those groups actually contain flights.
    mode_flight_groups = {"All": all_flight_groups}
    if any(flights for flights, _color in safe_flight_groups.values()):
        mode_flight_groups["Safe"] = safe_flight_groups
    if any(flights for flights, _color in unsafe_flight_groups.values()):
        mode_flight_groups["Unsafe"] = unsafe_flight_groups

    # Optional flight-computer impact markers (CATS/RCU), registered by reanalysis.run_reanalysis_comparison
    try:
        flight_computer_impacts = utils.ensure_list(utils.lookup("flight_computer_impacts", constants, variations)[0])
    except KeyError:
        flight_computer_impacts = None

    plot_landing_positions_with_modes(
        constants,
        exclusion_zones,
        buffer_zones,
        plot_name="landing_positions",
        mode_flight_groups=mode_flight_groups,
        safety_by_config=safety_by_config,
        flight_computer_impacts=flight_computer_impacts,
    )
    
    if should_create_flight_plots(variations):
        all_flights = [flight for scenario_set in scenario_sets for flight in scenario_set.values()]
        all_flights.extend(payload_flights)
        compare_trajectories(constants, variations, all_flights)

        for scenario_set in scenario_sets:
            for scenario_name, flight in scenario_set.items():
                print_flight_section_heading(flight, scenario_name)
                print_one_flight_with_custom_prints(flight)

                # Create the detailed notebook plots for this scenario and environment.
                if OUTPUT_LEVEL >= 1:
                    plot_one_flight_with_custom_plots(
                        constants,
                        variations,
                        flight,
                        scenario_name=scenario_name,
                    )

    return constants, variations
