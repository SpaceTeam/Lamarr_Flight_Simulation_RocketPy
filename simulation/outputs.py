"""
Output and post-processing helpers for the RocketPy simulation backend.
"""

from pathlib import Path

from matplotlib.path import Path as MatplotlibPath
import numpy as np
import plotly.graph_objects as go
from rocketpy import CompareFlights
from rocketpy.simulation import FlightDataExporter
from rocketpy import Flight, Environment, Motor, Fins, Rocket

import simulation.utils as utils
from simulation.custom_print_and_plot_functions import CustomPlots, CustomPrints

from IPython.display import display


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
    rocket = utils.lookup("rocket", constants, variations)[0]
    motor = utils.lookup("motor", constants, variations)[0]
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
    custom_plots.plot_angle_of_attack()
    custom_plots.plot_vertical_motion()
    flight.plots.trajectory_3d()


# =============================================================================
# Comparison plots
# =============================================================================

def compare_trajectories(constants, flights):
    """
    Draw 3D and 2D trajectory comparison plots for the given flights.
    """
    project_path = constants["project_path"]
    utils.ensure_project_folders(project_path)

    comparison = CompareFlights(flights)
    # save to disk
    comparison.trajectories_3d(figsize=(9,7), legend=True, filename=str(project_path / "plots" / "trajectories_3d_comparison.png"))
    # display plot inline
    comparison.trajectories_3d(figsize=(9,7), legend=True)
    # comparison.trajectories_2d(legend=legend, filename=str(project_path / "plots" / "2d_xy.png"), plane="xy")
    # comparison.trajectories_2d(legend=legend, filename=str(project_path / "plots" / "2d_xz.png"), plane="xz")
    # comparison.trajectories_2d(legend=legend, filename=str(project_path / "plots" / "2d_yz.png"), plane="yz")


# =============================================================================
# Export KML
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
    Split scenario-set dictionaries into nominal, no-main, and ballistic flight lists.
    """
    nominal_flights = []
    no_main_flights = []
    ballistic_flights = []

    for scenario_set in scenario_sets:
        if "nominal" in scenario_set:
            nominal_flights.append(scenario_set["nominal"])

        if "no_main" in scenario_set:
            no_main_flights.append(scenario_set["no_main"])

        if "ballistic" in scenario_set:
            ballistic_flights.append(scenario_set["ballistic"])

    return nominal_flights, no_main_flights, ballistic_flights


def register_safety_results(constants, variations, prefix, scenario_sets):
    """
    Register nominal, no-main, ballistic, and configuration lists for a safety result group.
    """
    nominal_flights, no_main_flights, ballistic_flights = split_flights_by_scenario(scenario_sets)
    configurations = sorted({scenario_config_key(flight) for flight in nominal_flights})

    constants, variations = utils.register(f"{prefix}_rocket_nominal", nominal_flights, constants, variations)
    constants, variations = utils.register(f"{prefix}_rocket_no_main", no_main_flights, constants, variations)
    constants, variations = utils.register(f"{prefix}_rocket_ballistic", ballistic_flights, constants, variations)
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
        "rocket_nominal": (get_registered_flights(constants, variations, "safe_rocket_nominal"), "green"),
        "rocket_no_main": (get_registered_flights(constants, variations, "safe_rocket_no_main"), "orange"),
        "rocket_ballistic": (get_registered_flights(constants, variations, "safe_rocket_ballistic"), "red"),
    }
    unsafe_flight_groups = {
        "rocket_nominal": (get_registered_flights(constants, variations, "unsafe_rocket_nominal"), "green"),
        "rocket_no_main": (get_registered_flights(constants, variations, "unsafe_rocket_no_main"), "orange"),
        "rocket_ballistic": (get_registered_flights(constants, variations, "unsafe_rocket_ballistic"), "red"),
    }

    safe_payload_flights = get_registered_flights(constants, variations, "safe_payload")
    unsafe_payload_flights = get_registered_flights(constants, variations, "unsafe_payload")
    if safe_payload_flights or unsafe_payload_flights:
        safe_flight_groups["payload_nominal"] = (safe_payload_flights, "blue")
        unsafe_flight_groups["payload_nominal"] = (unsafe_payload_flights, "blue")

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
                "impact: (%{x:.1f}, %{y:.1f}) m"
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
):
    """
    Build a landing-position plot of the buffer/exclusion zones, with optional flight-mode buttons.

    When `zones_only` is True only the zones are drawn. Otherwise one button per entry of
    `mode_flight_groups` toggles which mode's flight markers are visible. The zones and the
    launch-rail marker stay visible across every mode.
    """
    project_path = constants["project_path"]
    utils.ensure_project_folders(project_path)
    figure = go.Figure()

    # Plot buffer zones first so they do not visually cover the red exclusion zones.
    plot_zones(figure, buffer_zones, label="Buffer zone", color="orange")
    plot_zones(figure, exclusion_zones, label="Exclusion zone", color="red")

    buttons = []

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

        # Build one update button per mode; each sets a visibility and a showlegend array for all traces.
        total_traces = len(figure.data)
        for mode_name, (start, end) in mode_trace_ranges.items():
            visibility = [False] * total_traces
            # Zones and the launch rail stay visible across modes.
            for trace_index in range(zone_trace_count):
                visibility[trace_index] = True
            visibility[launch_trace_index] = True
            # Show this mode's flight traces only.
            for trace_index in range(start, end):
                visibility[trace_index] = True

            showlegend = [False] * total_traces
            for trace_index, legend_flag in enumerate(zone_legend_flags):
                showlegend[trace_index] = legend_flag
            showlegend[launch_trace_index] = True
            for trace_index in range(start, end):
                showlegend[trace_index] = True

            buttons.append(dict(
                label=mode_name,
                method="update",
                args=[{"visible": visibility, "showlegend": showlegend}],
            ))

    plot_titles = {
        "landing_positions": "Landing Positions",
    }
    layout_kwargs = dict(
        title=dict(
            text=plot_titles.get(plot_name, plot_name.replace("_", " ").title()),
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
    if buttons:
        layout_kwargs["updatemenus"] = [dict(
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
    figure.update_layout(**layout_kwargs)

    add_compass_labels(figure)

    # Save in the requested format.
    if save_format == "png":
        figure.write_image(str(project_path / "plots" / f"{plot_name}.png"))
    else:
        figure.write_html(str(project_path / "plots" / f"{plot_name}.html"))

    display(figure)


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

    nominal_flights, no_main_flights, ballistic_flights = split_flights_by_scenario(scenario_sets)
    all_flight_groups = {
        "rocket_nominal": (nominal_flights, "green"),
        "rocket_no_main": (no_main_flights, "orange"),
        "rocket_ballistic": (ballistic_flights, "red"),
    }
    if payload_flights:
        all_flight_groups["payload_nominal"] = (payload_flights, "blue")

    print_scan_values(constants, variations)
    constants, variations = calculate_safe_flights(constants, variations, buffer_zones)

    safe_flight_groups, unsafe_flight_groups = build_safe_unsafe_flight_groups(constants, variations)
    safety_by_config = build_safety_by_config(constants, variations)

    # Only display the Safe/Unsafe buttons when those groups actually contain flights.
    mode_flight_groups = {"All": all_flight_groups}
    if any(flights for flights, _color in safe_flight_groups.values()):
        mode_flight_groups["Safe"] = safe_flight_groups
    if any(flights for flights, _color in unsafe_flight_groups.values()):
        mode_flight_groups["Unsafe"] = unsafe_flight_groups

    plot_landing_positions_with_modes(
        constants,
        exclusion_zones,
        buffer_zones,
        plot_name="landing_positions",
        mode_flight_groups=mode_flight_groups,
        safety_by_config=safety_by_config,
    )
    
    if should_create_flight_plots(variations):
        all_flights = [flight for scenario_set in scenario_sets for flight in scenario_set.values()]
        all_flights.extend(payload_flights)
        compare_trajectories(constants, all_flights)

        for scenario_set in scenario_sets:
            for scenario_name, flight in scenario_set.items():
                print_flight_section_heading(flight, scenario_name)
                print_one_flight_with_custom_prints(flight)

                # Create the detailed notebook plots for this scenario and environment.
                if OUTPUT_LEVEL >= 1:
                    plot_one_flight_with_custom_plots(constants, variations, flight, scenario_name=scenario_name)

    return constants, variations
