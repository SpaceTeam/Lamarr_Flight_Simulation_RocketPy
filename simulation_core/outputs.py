"""
Output and post-processing helpers for the RocketPy simulation backend.
"""

import csv
import os
import time
from pathlib import Path

import nbformat
from tabulate import tabulate
from nbconvert import HTMLExporter

from matplotlib.path import Path as MatplotlibPath
import numpy as np
import pandas as pd
import plotly.graph_objects as go
from rocketpy.simulation import FlightDataExporter
from rocketpy import Flight, Environment, Motor, Fins, Rocket

from simulation_core.utils import *
from simulation_core.custom_print_and_plot_functions import CustomPlots, CustomPrints, get_shock_at_parachute_deployment, get_speed_at_parachute_deployment
from simulation_core.config_schema import SimParams, RocketConfig
from simulation_core.reanalysis import stitch_ascent_and_descent
SCENARIO_COLORS = {"nominal": "green", "no_main": "orange", "ballistic": "red", "matched": "purple", "payload": "blue", "ascent": "gray"}
SAFETY_LABELS = ("safe", "suboptimal", "unsafe")
NOTEBOOK_SAVE_TIMEOUT_S = 30
NOTEBOOK_SAVE_POLL_INTERVAL_S = 0.5
HEADLESS_ENV_VAR = "SIMULATION_HEADLESS"        # set by run_simulation.py
MAX_PLOTLY_BUTTONS = 5
SUPERSONIC_MACH = 1.2


# =============================================================================
# Printing
# =============================================================================

def print_one_environment(env: Environment, title):
    printmd(f"## {title}")
    # env.prints.gravity_details()
    env.prints.launch_site_details()
    env.prints.atmospheric_model_details()
    env.prints.atmospheric_conditions()
    # env.prints.print_earth_details()


def print_one_motor(motor: Motor, type: str, inertia=None):
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


def print_one_rocket(rocket: Rocket, rocket_length_m: float):
    print(f"Rocket center of wet mass from tip: {(rocket_length_m - rocket.center_of_mass(0)) * 1000} mm")
    print(f"Rocket center of mass without motor from tip: {(rocket_length_m - rocket.center_of_mass_without_motor) * 1000} mm")
    rocket.prints.inertia_details()
    # rocket.prints.rocket_geometrical_parameters()
    # rocket.prints.rocket_aerodynamics_quantities()
    rocket.prints.parachute_data()


def print_one_flight_with_custom_prints(flight: Flight, rocket_cfg: RocketConfig):
    """Print flight summary using custom print helpers."""
    custom_prints = CustomPrints(flight)

    # Flights started from a shared ascent's apogee have no rail phase, so print rail conditions from the ascent
    ascent = getattr(flight, "ascent_flight", None) or flight
    ascent.prints.launch_rail_conditions()
    ascent.prints.out_of_rail_conditions()
    print(f"OpenRocket Rail Departure Velocity: {ascent.speed(ascent.out_of_rail_time + 0.01)} m/s")      # OpenRocket flags that event 0.01 s later
    print(f"Effective rail length: {ascent.effective_1rl} m")
    fineness_ratio = round(rocket_cfg.length / rocket_cfg.diameter, 2)
    print(f"Fineness ratio: {fineness_ratio}:1 -> min stability {fineness_ratio * 0.1}")
    custom_prints.apogee_conditions()
    # flight.prints.apogee_conditions()
    custom_prints.parachute_events()
    # flight.prints.events_registered()
    flight.prints.impact_conditions()
    custom_prints.impact_coordinates()
    # flight.prints.maximum_values()


def print_flight_section_heading(flight: Flight, scenario_name: str):
    """Print a readable section heading for one environment and scenario."""
    environment_name = flight.env.name if hasattr(flight.env, "name") else flight.name
    printmd(f"## {environment_name} | {scenario_name}")


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


def plot_parachute_models(params: SimParams):
    """Plot the parachute models for all configured parachutes."""
    if params.runtime.parachutes is None:
        return

    for parachute in params.runtime.parachutes.values():
        CustomPlots.plot_parachute_model(parachute, save_dir=params.project_path / "plots")


def plot_one_rocket(rocket: Rocket):
    rocket.plots.draw()
    rocket.plots.total_mass()
    # rocket.plots.reduced_mass()
    rocket.plots.drag_curves()
    # rocket.plots.static_margin()
    # rocket.plots.stability_margin()
    rocket.plots.thrust_to_weight()


def plot_one_flight_with_custom_plots(params: SimParams, flight: Flight, scenario_name: str | None = None):
    """Plot the custom flight plots."""
    rocket = flight.rocket
    motor = flight.rocket.motor
    environment_name = flight.env.name if hasattr(flight.env, "name") else flight.name
    plot_label = environment_name

    if scenario_name is not None:
        plot_label = f"{environment_name} | {scenario_name}"

    custom_plots = CustomPlots(
        flight_forecast=[flight],
        motor=[motor],
        plot_title=[plot_label],
        rocket=[rocket],
        rocket_config=[{"total_length": params.config.rocket.length / 1000}],
        save_dir=params.project_path / "plots",
    )

    custom_plots.plot_stability_and_cg_cp_position()
    custom_plots.plot_angle_of_attack_and_attitude_angle()
    custom_plots.plot_angular_velocity(transform_openrocket=False)
    
    if not params.runtime.reanalysis_mode:
        custom_plots.plot_motion_over_time()


_CUSTOM_PLOTS_FLIGHT_LIMIT = 13


def plot_all_flights_with_custom_plots(params: SimParams, scenario_sets):
    """Plot all scenario flights together in grouped CustomPlots with one button per flight."""
    flights = []
    motors = []
    plot_titles = []
    rockets = []
    rocket_configs = []

    for scenario_set in scenario_sets:
        for scenario_name, scenario_flight in scenario_set.items():
            # flights from a shared ascent need that ascent stitched in front
            flight = full_trajectory_flight(scenario_flight)
            env_name = flight.env.name if hasattr(flight.env, "name") else flight.name

            # When variation metadata is present, include the varied settings in the label.
            if hasattr(flight, "_meta") and flight._meta:
                meta_str = ", ".join(render_meta_flat(flight._meta))
                label = f"{meta_str} | {scenario_name}"
            else:
                label = f"{env_name} | {scenario_name}"

            # CG and stability are only plotted up to apogee, so use the ascent's rocket (with payload mass) when there is one
            ascent_flight = getattr(scenario_flight, "ascent_flight", None)
            ascent_rocket = ascent_flight.rocket if ascent_flight is not None else flight.rocket

            flights.append(flight)
            motors.append(ascent_rocket.motor)
            plot_titles.append(label)
            rockets.append(ascent_rocket)
            rocket_configs.append({"total_length": params.config.rocket.length})

    if len(flights) > _CUSTOM_PLOTS_FLIGHT_LIMIT:
        print(
            f"CustomPlots skipped: {len(flights)} flights exceed the {_CUSTOM_PLOTS_FLIGHT_LIMIT}-flight "
            f"limit."
        )
        return

    custom_plots = CustomPlots(
        flight_forecast=flights,
        motor=motors,
        plot_title=plot_titles,
        rocket=rockets,
        rocket_config=rocket_configs,
        save_dir=params.project_path / "plots",
    )

    custom_plots.plot_stability_and_cg_cp_position()
    custom_plots.plot_angle_of_attack_and_attitude_angle()
    custom_plots.plot_angular_velocity(transform_openrocket=False)
    if not params.runtime.reanalysis_mode:
        custom_plots.plot_motion_over_time()


# =============================================================================
# Comparison plots
# =============================================================================

def add_plotly_buttons(
    figure: go.Figure,
    mode_trace_indices: dict[str, list[int]],
    always_visible_indices: list[int],
    zone_legend_flags: list[bool] | None = None,
):
    """Build Plotly update-menu buttons that toggle trace visibility by mode.

    Args:
        figure: The Plotly figure to add the buttons to.
        mode_trace_indices: {mode_name: [trace indices]} - which traces to show per button.
        always_visible_indices: trace indices visible in every mode (launch rail, etc.).
        zone_legend_flags: per-zone showlegend flags; restores original values when a button fires.
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

    use_dropdown = len(buttons) > MAX_PLOTLY_BUTTONS
    figure.update_layout(
        updatemenus=[dict(
            type="dropdown" if use_dropdown else "buttons",
            direction="down" if use_dropdown else "right",
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


def compare_trajectories(params: SimParams, flights: list[Flight]):
    """Interactive 3D trajectory plot with optional GNSS overlays from reanalysis.

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
    all_x_values = []
    all_y_values = []

    # -------------------------------------------------------------------------
    # Simulated flight traces: x, y in local meters from launch; altitude is AGL
    # -------------------------------------------------------------------------
    shared_ascents = list(dict.fromkeys(flight.ascent_flight for flight in flights if getattr(flight, "ascent_flight", None) is not None))
    for flight in [*shared_ascents, *flights]:
        env_name = flight.env.name if hasattr(flight.env, "name") else flight.name
        scenario = next((scenario for scenario in SCENARIO_COLORS if scenario in (getattr(flight, "name", "") or "")), "other")
        times = np.asarray(flight.time)
        # Match the scenario keyword inside the name so SCENARIO_COLORS picks the right color.
        color = next((c for scenario, c in SCENARIO_COLORS.items() if scenario in (getattr(flight, "name", "") or "")), None)
        # Use flat metadata label when available — flight.name may contain newlines from obj_to_pretty_label.
        if hasattr(flight, "_meta") and flight._meta:
            display_name = f"{', '.join(render_meta_flat(flight._meta))} | {scenario}"
        else:
            display_name = f"{env_name} | {scenario}"
        x_values = np.array([flight.x(t) for t in times])
        y_values = np.array([flight.y(t) for t in times])
        all_x_values.append(x_values)
        all_y_values.append(y_values)
        idx = len(figure.data)
        figure.add_trace(go.Scatter3d(
            x=x_values,
            y=y_values,
            z=np.array([flight.altitude(t) for t in times]),
            mode="lines",
            name=display_name,
            line=dict(color=color, width=3) if color else dict(width=3),
            hoverlabel=dict(namelength=-1),
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

    always_visible_indices = []
    gnss_traces = params.runtime.gnss_3d_traces

    if gnss_traces:
        for trace in gnss_traces.values():
            all_x_values.append(np.asarray(trace["x"], dtype=float))
            all_y_values.append(np.asarray(trace["y"], dtype=float))
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
        mode_trace_indices = {name: env_trace_indices[name] for name in env_names_ordered}
        # The first button is active by default; hide every trace that doesn't belong to it.
        first_mode_indices = set(next(iter(mode_trace_indices.values())))
        all_dynamic_indices = {idx for indices in env_trace_indices.values() for idx in indices}
        for idx in all_dynamic_indices - first_mode_indices:
            figure.data[idx].visible = False
        add_plotly_buttons(figure, mode_trace_indices, always_visible_indices)

    # give all plots the same horizontal scale for better comparison, with a 5% margin around the largest range
    all_x = np.concatenate(all_x_values)
    all_y = np.concatenate(all_y_values)
    half_span = 1.05 * max(np.ptp(all_x), np.ptp(all_y)) / 2
    x_center = (all_x.max() + all_x.min()) / 2
    y_center = (all_y.max() + all_y.min()) / 2

    figure.update_layout(
        title="3D Trajectory Comparison",
        scene=dict(
            xaxis=dict(title="East / West [m]", range=[x_center - half_span, x_center + half_span]),
            yaxis=dict(title="North / South [m]", range=[y_center - half_span, y_center + half_span]),
            zaxis_title="Altitude AGL [m]",
            aspectmode="cube",
            camera=dict(
                # x rotates camera east with pos values
                # y rotates camera north with pos values
                # z controls elevation: bigger = looking down at a steeper angle
                eye=dict(x=0.1, y=-2.3, z=0.2),
            ),
        ),
        hoverlabel=dict(namelength=-1),
        width=900,
        height=700,
    )

    figure.write_html(str(params.project_path / "plots" / "compare_trajectories.html"))
    figure.show(renderer="notebook")


# =============================================================================
# Exports
# =============================================================================

def full_trajectory_flight(flight: Flight) -> Flight:
    """Return the flight from launch to landing; a flight that started at a shared ascent's apogee gets that ascent stitched in front."""
    ascent_flight = getattr(flight, "ascent_flight", None)
    return stitch_ascent_and_descent(ascent_flight, flight) if ascent_flight is not None else flight


def export_all_kml(params: SimParams):
    """Export KML files for all flights, each from launch to landing."""
    exported_files = []
    project_path = params.project_path
    ensure_project_folders(project_path)

    for i, scenario_set in enumerate(params.runtime.scenario_sets, start=1):
        for scenario_flight in scenario_set.values():
            flight = full_trajectory_flight(scenario_flight)
            file_name = Path(f"flight_{i}_{flight.filename_label}.kml")
            file_path = project_path / "trajectory_kml" / file_name
            FlightDataExporter(flight).export_kml(file_name=file_path, altitude_mode="relativetoground")
            exported_files.append(file_path)

    return exported_files


def export_all_trajectory_csv(params: SimParams):
    """Export one CSV per flight from launch to landing with columns latitude, longitude, altitude (m AGL) at every ODE time step.
    By request for WARR, maybe useful, otherwise remove later."""
    exported_files = []
    project_path = params.project_path
    ensure_project_folders(project_path)

    for i, scenario_set in enumerate(params.runtime.scenario_sets, start=1):
        for scenario_name, scenario_flight in scenario_set.items():
            flight = full_trajectory_flight(scenario_flight)
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
    """Export a Jupyter notebook with its current outputs to a self-contained HTML file.
    The notebook must be saved first! Either by autosave or by pressing Ctrl+S.
    Needs to be called in the last cell of the notebook and have its own cell."""
    # Headless runs (run_simulation.py) have no editor that saves the file.
    if os.environ.get(HEADLESS_ENV_VAR):
        print("Headless run: report.html is written by run_simulation.py.")
        return None

    notebook_path = Path(notebook_path)

    if output_path is None:
        output_path = notebook_path.with_suffix(".html")

    output_path = Path(output_path)

    # Wait for save: all earlier cells finished before this call, so any save after this moment contains their outputs.
    call_time = time.time()
    while notebook_path.stat().st_mtime <= call_time:
        if time.time() - call_time > NOTEBOOK_SAVE_TIMEOUT_S:
            raise TimeoutError(f"{notebook_path.name} was not saved within {NOTEBOOK_SAVE_TIMEOUT_S} s; turn on autosave or press Ctrl+S. "
                               "report.html not written!")
        time.sleep(NOTEBOOK_SAVE_POLL_INTERVAL_S)

    nb = nbformat.read(notebook_path, as_version=4)

    # The export cell's own saved output is still from the previous run, so leave it out of the report.
    for cell in nb.cells:
        if cell.cell_type == "code" and "export_notebook_to_html(" in cell.source:
            cell.outputs = []

    html_body, _ = HTMLExporter(theme="dark").from_notebook_node(nb)
    output_path.write_text(html_body, encoding="utf-8")

    print(f"Notebook exported to: {output_path}")
    return output_path


# =============================================================================
# Safety printing
# =============================================================================

def print_configurations(title, configurations):
    """Print configurations grouped by heading, then environment."""
    printmd(f"## {title}")

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


def print_flight_details(params: SimParams, safety_label: str, buffer_zones: dict, suboptimal_zones: dict):
    """Print one table with landing coordinates, distance and landing zone type for every flight with the given safety label.

    The unsafe table only lists flights that landed in a buffer zone; heading/inclination/scenario cells that repeat the row above are left blank.
    """
    printmd(f"## {safety_label.capitalize()} flight details")

    # Scenario name -> params.runtime attribute holding this safety label's flights (also defines the table's scenario order)
    scenario_attributes = {
        "nominal":   f"{safety_label}_rocket_nominal",
        "no_main":   f"{safety_label}_rocket_no_main",
        "ballistic": f"{safety_label}_rocket_ballistic",
        "payload":   f"{safety_label}_payload",
    }
    rows = []

    for scenario, attr_name in scenario_attributes.items():
        for flight in get_flights(params, attr_name):
            if flight_lands_in_zone(flight, buffer_zones):
                zone_type = "unsafe"
            elif flight_lands_in_zone(flight, suboptimal_zones):
                zone_type = "subopt."
            else:
                zone_type = "safe"

            if safety_label == "unsafe" and zone_type != "unsafe":
                continue

            env_name = flight.env.name if hasattr(flight.env, "name") else flight.name
            coordinates = f"{flight.latitude(flight.t_final):.6f}, {flight.longitude(flight.t_final):.6f}"
            rows.append([flight.heading, flight.inclination, scenario, env_name, coordinates, round(landing_distance(flight)), zone_type])

    if not rows:
        print("None")
        return

    # Sort by heading, inclination, scenario (in the order above) and environment
    scenario_order = list(scenario_attributes)
    rows.sort(key=lambda row: (row[0], row[1], scenario_order.index(row[2]), row[3]))

    # Blank the leading heading/inclination/scenario cells that match the row above, so each group value is shown once
    group_column_count = 3
    merged_rows = []
    previous_row = [None] * group_column_count
    for row in rows:
        repeated_cells = 0
        while repeated_cells < group_column_count and row[repeated_cells] == previous_row[repeated_cells]:
            repeated_cells += 1
        merged_rows.append([""] * repeated_cells + row[repeated_cells:])
        previous_row = row

    headers = ["Heading [°]", "Incl. [°]", "Scenario", "Environment", "Coordinates (lat, lon)", "Distance [m]", "Type"]
    column_alignment = ("right", "right", "left", "left", "left", "right", "center")
    printmd(tabulate(merged_rows, headers=headers, tablefmt="pipe", colalign=column_alignment))


def print_landing_distance_summary(params: SimParams):
    """Print the min and max landing distance from the launch rail per heading and inclination, one column per scenario."""
    # Pair every flight (rocket scenarios and payload) with its scenario name.
    flights_with_scenario = [(name, flight) for scenario_set in params.runtime.scenario_sets for name, flight in scenario_set.items()]
    flights_with_scenario += [("payload", flight) for flight in get_flights(params, "flight_payload")]

    # Build: {(heading, inclination): {scenario: [distance, ...]}}; each list holds one distance per environment.
    distances_by_config = {}
    for scenario_name, flight in flights_with_scenario:
        config_distances = distances_by_config.setdefault((flight.heading, flight.inclination), {})
        config_distances.setdefault(scenario_name, []).append(landing_distance(flight))

    # SCENARIO_COLORS lists every scenario in display order; only keep the ones that were simulated.
    scenario_names = [name for name in SCENARIO_COLORS if any(name in by_scenario for by_scenario in distances_by_config.values())]
    # Safety is decided per heading, so any environment/inclination of that heading gives the same label.
    safety_by_heading = {heading: label for (_env, heading, _inclination), label in build_safety_by_config(params).items()}

    # Pool every heading and inclination per scenario for the closing "Total" row.
    total_by_scenario = {}
    for by_scenario in distances_by_config.values():
        for scenario_name, distances in by_scenario.items():
            total_by_scenario.setdefault(scenario_name, []).extend(distances)

    # Each row is (leading cells, {scenario: distances}); the Total row comes last.
    row_sources = [
        ([heading, inclination, safety_by_heading.get(heading, "unknown")], by_scenario)
        for (heading, inclination), by_scenario in sorted(distances_by_config.items())
    ]
    row_sources.append((["**Total**", "", ""], total_by_scenario))

    rows = []
    for leading_cells, by_scenario in row_sources:
        row = list(leading_cells)
        for scenario_name in scenario_names:
            distances = by_scenario.get(scenario_name)
            row.append(f"{min(distances):.0f} – {max(distances):.0f}" if distances else "–")
        rows.append(row)

    printmd("## Landing distance from launch rail")
    printmd("Min - max distance in meters across all environments, per heading and inclination.")
    column_alignment = ("right", "right", "left") + ("left",) * len(scenario_names)
    printmd(tabulate(rows, headers=["Heading [°]", "Inclination [°]", "Safety", *scenario_names], tablefmt="pipe", colalign=column_alignment))


# =============================================================================
# Safety calculations
# =============================================================================

def get_flights(params: SimParams, attr_name: str) -> list[Flight]:
    """Return a flight list from params.runtime by attribute name, or empty list when not set."""
    value = getattr(params.runtime, attr_name, None)
    if value is None:
        return []
    return ensure_list(value)


def scenario_config_key(flight: Flight):
    """Identifier for the configuration a flight belongs to: (env_name, heading, inclination)."""
    env_name = flight.env.name if hasattr(flight.env, "name") else flight.name
    return (env_name, flight.heading, flight.inclination)


def landing_distance(flight: Flight) -> float:
    """Return the horizontal distance in meters between the launch rail and the flight's impact point."""
    return float(np.hypot(flight.x_impact, flight.y_impact))


def flight_lands_in_zone(flight: Flight, zones: dict):
    """Return True when the flight's impact point lies inside any of the supplied zone polygons."""
    return any(
        MatplotlibPath(zone_coordinates).contains_point((flight.x_impact, flight.y_impact))
        for zone_coordinates in zones.values()
    )


def split_flights_by_scenario(scenario_sets: list[dict]) -> tuple[list[Flight], list[Flight], list[Flight], list[Flight]]:
    """Split scenario-set dicts into nominal, no-main, ballistic, and matched flight lists."""
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


def store_safety_results(params: SimParams, prefix: str, scenario_sets: list[dict]):
    """Store nominal, no-main, ballistic, matched, and configuration lists in params.runtime."""
    nominal_flights, no_main_flights, ballistic_flights, matched_flights = split_flights_by_scenario(scenario_sets)
    configurations = sorted({scenario_config_key(flight) for flight in nominal_flights})

    setattr(params.runtime, f"{prefix}_rocket_nominal", nominal_flights)
    setattr(params.runtime, f"{prefix}_rocket_no_main", no_main_flights)
    setattr(params.runtime, f"{prefix}_rocket_ballistic", ballistic_flights)
    setattr(params.runtime, f"{prefix}_rocket_matched", matched_flights)
    setattr(params.runtime, f"{prefix}_configurations", configurations)


def print_flight_stats(params: SimParams):
    """Print the stats tables for nominal, no_main, ballistic and matched flights and export them to one Excel file."""
    stats_by_scenario = {scenario: print_scenario_stats(params, scenario) for scenario in ("nominal", "no_main", "ballistic", "matched")}
    # Skip scenarios that were not simulated
    stats_by_scenario = {scenario: table for scenario, table in stats_by_scenario.items() if table is not None}
    if not stats_by_scenario:
        return

    # Export all tables to excel
    excel_path = params.project_path / "variation_stats.xlsx"
    with pd.ExcelWriter(excel_path) as excel_writer:
        for scenario, table in stats_by_scenario.items():
            table.to_excel(excel_writer, sheet_name=scenario, index=False)
    print(f"\nExported to: {excel_path}")


def print_scenario_stats(params: SimParams, scenario: str) -> pd.DataFrame | None:
    """Print a table of per-flight stats for one scenario (one row per environment/variation combination) and return it, 
    or None if there are no such flights."""
    # Launch, apogee and parachute stats repeat the nominal values for no_main/ballistic, so only nominal and matched 
    # (spliced launch-to-landing flight) show them
    show_full_stats = scenario in ("nominal", "matched")
    headers = ["Environment", "Config"]
    if show_full_stats:
        headers.extend(["Thrust to weight\nratio @ rail exit", "rail exit\nvelocity [m/s]",
                        "stability\n@ rail exit [cal]",
                        "max instability in\nsupersonic region [cal]",
                        "apogee\nAGL [m]"])
    headers.extend(["Impact\nspeed [m/s]", "Landing distance\nfrom launch [m]"])
    add_main_speed = False
    parachute_names = []
    rows = []

    scenario_flights = [scenario_set[scenario] for scenario_set in params.runtime.scenario_sets if scenario in scenario_set]
    # Sort by the variation values
    scenario_flights.sort(key=lambda flight: (tuple(getattr(flight, "_meta", {}).values()), flight.env.name))

    for flight in scenario_flights:
        meta = getattr(flight, "_meta", {})
        config_str = ",\n".join(f"{k}={v}" for k, v in meta.items()) if meta else "-"
        row_values = [flight.env.name, config_str]

        if show_full_stats:
            # Flights started from a shared ascent's apogee have no rail/burn phase, so read those stats from the ascent
            ascent = getattr(flight, "ascent_flight", None) or flight
            t_rail = ascent.out_of_rail_time
            t_w = ascent.rocket.motor.thrust(t_rail) / (ascent.rocket.total_mass(t_rail) * 9.81)
            v_rail = ascent.out_of_rail_velocity
            stability = ascent.stability_margin(t_rail)

            # max instability in supersonic region
            if ascent.max_mach_number > SUPERSONIC_MACH:
                stability_at_max_speed = round(ascent.stability_margin(ascent.max_speed_time), 2)
            else:
                stability_at_max_speed = "-"
            apogee_agl = flight.apogee - flight.env.elevation
            row_values.extend([round(t_w, 2), round(v_rail, 1), round(stability, 2), stability_at_max_speed, round(apogee_agl)])

        row_values.extend([round(abs(flight.impact_velocity), 2), round(landing_distance(flight))])

        if show_full_stats:
            opening_times = {parachute.name: trigger_time + parachute.lag for trigger_time, parachute in flight.parachute_events}

            # Shock around deployment for each parachute (flights of one scenario all deploy the same parachutes)
            for deployment_time in opening_times.values():
                row_values.append(round(get_shock_at_parachute_deployment(flight, deployment_time), 2))
            parachute_names = list(opening_times)

            # speed @ main deployment
            main_deployment_time = opening_times.get("main")
            if main_deployment_time:
                row_values.append(round(get_speed_at_parachute_deployment(flight, main_deployment_time), 2))
                add_main_speed = True

        rows.append(row_values)

    if not rows:
        return None

    headers.extend(f"{name} shock\n@ deployment [g]" for name in parachute_names)
    
    if add_main_speed:
        headers.append("speed @\nmain deployment [m/s]")

    # Markdown tables can't hold line breaks: headers get spaces, multi-line cells (config) get HTML <br> breaks
    single_line_headers = [header.replace("\n", " ") for header in headers]
    markdown_rows = [[value.replace("\n", "<br>") if isinstance(value, str) else value for value in row] for row in rows]

    printmd(f"## Stats for {scenario} flights")
    column_alignment = ("left", "left") + ("right",) * (len(headers) - 2)
    printmd(tabulate(markdown_rows, headers=single_line_headers, tablefmt="pipe", colalign=column_alignment))

    return pd.DataFrame(rows, columns=single_line_headers)


def calculate_safe_flights(params: SimParams, buffer_zones: dict, suboptimal_zones: dict):
    """Classify all flights as safe, suboptimal or unsafe based on landing zone membership.

    A heading is unsafe if any scenario (nominal/no_main/ballistic/payload), inclination,
    or environment caused a landing inside a buffer zone; otherwise it is suboptimal if any landed inside a suboptimal zone.
    """
    payload_flights = get_flights(params, "flight_payload")
    scenario_sets = params.runtime.scenario_sets
    all_flights = [flight for scenario_set in scenario_sets for flight in scenario_set.values()] + payload_flights

    # First pass: find which headings have any flight landing in each zone type.
    unsafe_headings = {flight.heading for flight in all_flights if flight_lands_in_zone(flight, buffer_zones)}
    # A heading that is unsafe stays unsafe even if it also lands in a suboptimal zone.
    suboptimal_headings = {flight.heading for flight in all_flights if flight_lands_in_zone(flight, suboptimal_zones)} - unsafe_headings
    label_by_heading = {heading: "suboptimal" for heading in suboptimal_headings} | {heading: "unsafe" for heading in unsafe_headings}

    # Second pass: every scenario_set and payload flight takes its heading's label, even if its specific
    # (env, inclination) flights all landed outside the zones.
    for safety_label in SAFETY_LABELS:
        store_safety_results(params, safety_label, [s for s in scenario_sets if label_by_heading.get(s["nominal"].heading, "safe") == safety_label])
        setattr(params.runtime, f"{safety_label}_payload", [p for p in payload_flights if label_by_heading.get(p.heading, "safe") == safety_label])

    print_flight_stats(params)
    print_landing_distance_summary(params)
    print_configurations("Safe Configurations by Heading", params.runtime.safe_configurations or [])
    print_flight_details(params, "safe", buffer_zones, suboptimal_zones)
    if suboptimal_zones:
        print_configurations("Suboptimal (but safe) Configurations by Heading", params.runtime.suboptimal_configurations or [])
        print_flight_details(params, "suboptimal", buffer_zones, suboptimal_zones)
    print_configurations("Unsafe Configurations by Heading", params.runtime.unsafe_configurations or [])
    print_flight_details(params, "unsafe", buffer_zones, suboptimal_zones)


# =============================================================================
# Safety plots
# =============================================================================

def build_safety_flight_groups(params: SimParams, safety_label: str):
    """Build the flight-group dict for one safety label ("safe", "suboptimal" or "unsafe") from params.runtime."""
    flight_groups = {
        f"rocket_{scenario}": (get_flights(params, f"{safety_label}_rocket_{scenario}"), SCENARIO_COLORS[scenario])
        for scenario in ("nominal", "no_main", "ballistic", "matched")
    }
    flight_groups["payload_nominal"] = (get_flights(params, f"{safety_label}_payload"), SCENARIO_COLORS["payload"])
    return flight_groups


def build_safety_by_config(params: SimParams):
    """Build a {(environment, heading, inclination): 'safe'|'suboptimal'|'unsafe'} map from params.runtime."""
    return {
        configuration: safety_label
        for safety_label in SAFETY_LABELS
        for configuration in (getattr(params.runtime, f"{safety_label}_configurations") or [])
    }


# =============================================================================
# Landing position plots
# =============================================================================

def plot_zones(figure: go.Figure, zones: dict[str, list[tuple[float, float]]], label: str, color: str, alpha=0.3):
    """Plot exclusion or buffer-zone polygons."""
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


def add_compass_labels(figure: go.Figure):
    """Add bold N / E / S / W compass labels at the plot-area edges."""
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


def add_mode_flight_traces(
    figure: go.Figure,
    flight_groups: dict[str, tuple[list, str]],
    visible: bool,
    safety_by_config: dict[tuple, str],
):
    """Add one Scatter trace per flight group for the given mode; return the trace index range."""
    start_index = len(figure.data)

    for label, (flights, color) in flight_groups.items():
        flight_list = ensure_list(flights)

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
                landing_distance(flight),
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
                "distance from launch: %{customdata[6]:.0f} m<br>"
                "lat: %{customdata[4]:.5f}°; lon: %{customdata[5]:.5f}°"
                "<extra></extra>"
            ),
            visible=visible,
            showlegend=visible,
        ))

    return start_index, len(figure.data)


def plot_landing_positions_with_modes(
    params: SimParams,
    exclusion_zones,
    buffer_zones,
    suboptimal_zones,
    plot_name,
    mode_flight_groups=None,
    safety_by_config=None,
    zones_only=False,
    flight_computer_impacts=None,
):
    """Build an interactive landing-position plot with optional flight-mode buttons.

    When zones_only is True only the zones are drawn. Otherwise one button per entry of
    mode_flight_groups toggles which mode's markers are visible.
    """
    project_path = params.project_path
    ensure_project_folders(project_path)
    figure = go.Figure()

    # Plot from least to most critical so the red exclusion zones stay on top.
    plot_zones(figure, suboptimal_zones, label="Suboptimal zone", color="gold")
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
                    f"distance from launch rail: {np.hypot(impact['x'], impact['y']):.0f} m<br>"
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

    figure.update_layout(
        title=dict(text="Landing Positions", x=0.5, xanchor="center"),
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

    add_compass_labels(figure)

    figure.write_html(str(project_path / "plots" / f"{plot_name}.html"))

    figure.show(renderer="notebook")


# =============================================================================
# Notebook display mode
# =============================================================================

def run_notebook_display_mode(params: SimParams, exclusion_zones, buffer_zones, suboptimal_zones):
    """Run the notebook display mode that fits the config: a single flight or variations.

    Both modes share the same All/Safe/Suboptimal/Unsafe landing-positions plot. Single-flight mode adds a
    per-flight safety summary, trajectory comparison, and detailed per-scenario prints/plots.
    With variations, the trajectory comparison is only drawn up to _CUSTOM_PLOTS_FLIGHT_LIMIT flights.
    """
    scenario_sets = params.runtime.scenario_sets
    payload_flights = get_flights(params, "flight_payload")

    nominal_flights, no_main_flights, ballistic_flights, matched_flights = split_flights_by_scenario(scenario_sets)
    all_flight_groups = {
        "rocket_nominal":  (nominal_flights, SCENARIO_COLORS["nominal"]),
        "rocket_no_main":  (no_main_flights, SCENARIO_COLORS["no_main"]),
        "rocket_ballistic": (ballistic_flights, SCENARIO_COLORS["ballistic"]),
    }
    if matched_flights:
        all_flight_groups["rocket_matched"] = (matched_flights, SCENARIO_COLORS["matched"])
    if payload_flights:
        all_flight_groups["payload_nominal"] = (payload_flights, SCENARIO_COLORS["payload"])

    calculate_safe_flights(params, buffer_zones, suboptimal_zones)
    safety_by_config = build_safety_by_config(params)

    # Only display a Safe/Suboptimal/Unsafe button when that group actually contains flights.
    mode_flight_groups = {"All": all_flight_groups}
    for safety_label in SAFETY_LABELS:
        flight_groups = build_safety_flight_groups(params, safety_label)
        if any(flights for flights, _color in flight_groups.values()):
            mode_flight_groups[safety_label.capitalize()] = flight_groups

    flight_computer_impacts = ensure_list(params.runtime.flight_computer_impacts or [])

    plot_landing_positions_with_modes(
        params,
        exclusion_zones,
        buffer_zones,
        suboptimal_zones,
        plot_name="landing_positions",
        mode_flight_groups=mode_flight_groups,
        safety_by_config=safety_by_config,
        flight_computer_impacts=flight_computer_impacts,
    )

    if not has_variations(params.config):
        all_flights = [flight for scenario_set in scenario_sets for flight in scenario_set.values()]
        all_flights.extend(payload_flights)
        compare_trajectories(params, all_flights)

        for scenario_set in scenario_sets:
            for scenario_name, flight in scenario_set.items():
                print_flight_section_heading(flight, scenario_name)
                print_one_flight_with_custom_prints(flight, params.config.rocket)

        if params.config.output_level >= 1:
            plot_all_flights_with_custom_plots(params, scenario_sets)

    else:
        all_flights = [flight for scenario_set in scenario_sets for flight in scenario_set.values()]
        all_flights.extend(payload_flights)
        if len(all_flights) <= _CUSTOM_PLOTS_FLIGHT_LIMIT:
            compare_trajectories(params, all_flights)
        else:
            print(
                f"compare_trajectories skipped: {len(all_flights)} flights exceed the "
                f"{_CUSTOM_PLOTS_FLIGHT_LIMIT}-flight limit."
            )

        if params.config.output_level >= 1:
            plot_all_flights_with_custom_plots(params, scenario_sets)
