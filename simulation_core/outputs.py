"""
Output and post-processing helpers for the RocketPy simulation backend.
"""

import base64
import csv
import io
import os
import time
from dataclasses import dataclass
from pathlib import Path

import nbformat
from tabulate import tabulate
from nbconvert import HTMLExporter

from matplotlib.path import Path as MatplotlibPath
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import rasterio
from rasterio.warp import transform_bounds
from PIL import Image
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
SATELLITE_JPEG_QUALITY = 85                     # JPEG keeps the embedded image (and thus the HTML) far smaller than PNG
# Stats table columns shared by RocketPy and OpenRocket: metric key -> (header, decimal places). The keys are also the metric
# names of the OpenRocket sweep notebook; parachute shocks are named after the parachute, e.g. "drogue_shock" (see stats_column).
STATS_COLUMNS = {
    "thrust_to_weight_at_rail_exit": ("Thrust to weight\nratio @ rail exit", 2),
    "velocity_at_rail_exit": ("rail exit\nvelocity [m/s]", 1),
    "stability_at_rail_exit": ("stability\n@ rail exit [cal]", 2),
    "max_instability_in_supersonic_region": ("max instability in\nsupersonic region [cal]", 2),
    "apogee": ("apogee\nAGL [m]", 0),
    "impact_speed": ("Impact\nspeed [m/s]", 2),
    "landing_distance": ("Landing distance\nfrom launch [m]", 0),
    "speed_at_main_deployment": ("speed @\nmain deployment [m/s]", 2),
}
STATS_LABEL_COLUMNS = ("Environment", "Config")


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
    
    if params.config.reanalysis is None:
        # when a reanalysis section is present reanalysis.py will plot plot_motion_per_source comparing simulation to flight computer data
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

def configuration_label(point: "LandingPoint") -> str:
    """Describe a landing point's configuration in one line."""
    return ", ".join(f"{name}={value}" for name, value in point.configuration.items())


def print_configurations(title: str, nominal_points: list["LandingPoint"]):
    """Print the configurations of the given nominal landing points grouped by heading, then configuration."""
    printmd(f"## {title}")

    if not nominal_points:
        print("None")
        return

    # {heading: {configuration: {inclination, ...}}}
    grouped = {}
    for point in nominal_points:
        grouped.setdefault(point.heading, {}).setdefault(configuration_label(point), set()).add(point.inclination)

    for heading, configurations in sorted(grouped.items()):
        print(f"\nHeading {heading}:")

        for configuration, inclinations in sorted(configurations.items()):
            print(f"- {configuration}: inclinations {sorted(inclinations)}")


def print_flight_details(landing_points: list["LandingPoint"], label_by_heading: dict[float, str], safety_label: str,
                         buffer_zones: dict, suboptimal_zones: dict):
    """Print one table with landing coordinates, distance and landing zone type for every landing point with the given safety label.

    The unsafe table only lists points that landed in a buffer zone; heading/inclination/scenario cells that repeat the row above are left blank.
    """
    printmd(f"## {safety_label.capitalize()} flight details")

    # Scenarios shown in the table, in table order
    scenario_order = ["nominal", "no_main", "ballistic", "payload"]
    rows = []

    for point in landing_points:
        if point.scenario not in scenario_order or label_by_heading[point.heading] != safety_label:
            continue

        if lands_in_zone(point.x_east_m, point.y_north_m, buffer_zones):
            zone_type = "unsafe"
        elif lands_in_zone(point.x_east_m, point.y_north_m, suboptimal_zones):
            zone_type = "subopt."
        else:
            zone_type = "safe"

        if safety_label == "unsafe" and zone_type != "unsafe":
            continue

        coordinates = f"{point.latitude:.6f}, {point.longitude:.6f}"
        distance = round(float(np.hypot(point.x_east_m, point.y_north_m)))
        rows.append([point.heading, point.inclination, point.scenario, *point.configuration.values(), coordinates, distance, zone_type])

    if not rows:
        print("None")
        return

    # Sort by heading, inclination, scenario (in the order above) and the configuration columns
    configuration_names = list(landing_points[0].configuration)
    rows.sort(key=lambda row: (row[0], row[1], scenario_order.index(row[2]), *row[3:3 + len(configuration_names)]))

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

    headers = ["Heading [°]", "Incl. [°]", "Scenario", *configuration_names, "Coordinates (lat, lon)", "Distance [m]", "Type"]
    column_alignment = ("right", "right", "left") + ("left",) * len(configuration_names) + ("left", "right", "center")
    printmd(tabulate(merged_rows, headers=headers, tablefmt="pipe", colalign=column_alignment))


def print_landing_distance_summary(landing_points: list["LandingPoint"], label_by_heading: dict[float, str]):
    """Print the min and max landing distance from the launch rail per heading and inclination, one column per scenario."""
    # Build: {(heading, inclination): {scenario: [distance, ...]}}; each list holds one distance per configuration.
    distances_by_config = {}
    for point in landing_points:
        config_distances = distances_by_config.setdefault((point.heading, point.inclination), {})
        config_distances.setdefault(point.scenario, []).append(float(np.hypot(point.x_east_m, point.y_north_m)))

    # SCENARIO_COLORS lists every scenario in display order; only keep the ones that were simulated.
    scenario_names = [name for name in SCENARIO_COLORS if any(name in by_scenario for by_scenario in distances_by_config.values())]
    # Pool every heading and inclination per scenario for the closing "Total" row.
    total_by_scenario = {}
    for by_scenario in distances_by_config.values():
        for scenario_name, distances in by_scenario.items():
            total_by_scenario.setdefault(scenario_name, []).extend(distances)

    # Each row is (leading cells, {scenario: distances}); the Total row comes last.
    row_sources = [
        ([heading, inclination, label_by_heading[heading]], by_scenario)
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
    configuration_names = ", ".join(landing_points[0].configuration)
    printmd(f"Min - max distance in meters across all configurations ({configuration_names}), per heading and inclination.")
    column_alignment = ("right", "right", "left") + ("left",) * len(scenario_names)
    printmd(
        tabulate(
            rows,
            headers=["Heading [°]", "Inclination [°]", "Safety", *scenario_names],
            tablefmt="pipe",
            colalign=column_alignment,
        )
    )


def print_safety_summary(
    landing_points: list["LandingPoint"], label_by_heading: dict[float, str], buffer_zones: dict, suboptimal_zones: dict
):
    """Print the landing distance summary, then the configurations and flight details per safety label 
    (shared by RocketPy and OpenRocket)."""
    print_landing_distance_summary(landing_points, label_by_heading)
    titles = {
        "safe": "Safe Configurations by Heading",
        "suboptimal": "Suboptimal (but safe) Configurations by Heading",
        "unsafe": "Unsafe Configurations by Heading",
    }
    for safety_label, title in titles.items():
        # Without suboptimal zones no heading can be suboptimal, so its section is left out.
        if safety_label == "suboptimal" and not suboptimal_zones:
            continue
        nominal_points = [
            point
            for point in landing_points
            if point.scenario == "nominal" and label_by_heading[point.heading] == safety_label
        ]
        print_configurations(title, nominal_points)
        print_flight_details(landing_points, label_by_heading, safety_label, buffer_zones, suboptimal_zones)


# =============================================================================
# Safety calculations
# =============================================================================

@dataclass
class LandingPoint:
    """One simulated landing position, shared by RocketPy and OpenRocket."""
    scenario: str                   # key of SCENARIO_COLORS, e.g. "nominal", "no_main", "ballistic" or "payload"
    heading: float                  # launch heading [°]; the safety classification groups by it
    inclination: float              # launch inclination [°] from horizontal
    x_east_m: float
    y_north_m: float
    latitude: float
    longitude: float
    configuration: dict[str, str]   # what else distinguishes the runs, e.g. {"Environment": "icon_d2"}; table columns and hover lines


def landing_points_from_runtime(params: SimParams) -> list[LandingPoint]:
    """Save every simulated RocketPy flight (all scenarios and payload flights) into a LandingPoint."""

    flights_with_scenario = [
        (scenario, flight) for scenario_set in params.runtime.scenario_sets for scenario, flight in scenario_set.items()
    ]
    flights_with_scenario += [("payload", flight) for flight in get_flights(params, "flight_payload")]
    variation_names = list(
        dict.fromkeys(
            name
            for _scenario, flight in flights_with_scenario
            for name in getattr(flight, "_meta", {})
            if name not in ("heading", "inclination")
        )
    )
    return [
        LandingPoint(
            scenario=scenario,
            heading=flight.heading,
            inclination=flight.inclination,
            x_east_m=flight.x_impact,
            y_north_m=flight.y_impact,
            latitude=flight.latitude(flight.t_final),
            longitude=flight.longitude(flight.t_final),
            configuration={
                "Environment": flight.env.name if hasattr(flight.env, "name") else flight.name,
                # "-" for flights without that variation value, e.g. a payload flight
                **{name: str(getattr(flight, "_meta", {}).get(name, "-")) for name in variation_names},
            },
        )
        for scenario, flight in flights_with_scenario
    ]


def classify_headings(landing_points: list[LandingPoint], buffer_zones: dict, suboptimal_zones: dict) -> dict[float, str]:
    """Return {heading: "safe" | "suboptimal" | "unsafe"}: a heading is unsafe if any of its points lands in a buffer zone,
    otherwise suboptimal if any lands in a suboptimal zone."""
    unsafe_headings = {point.heading for point in landing_points if lands_in_zone(point.x_east_m, point.y_north_m, buffer_zones)}
    # A heading that is unsafe stays unsafe even if it also lands in a suboptimal zone.
    suboptimal_headings = {point.heading for point in landing_points if lands_in_zone(point.x_east_m, point.y_north_m, suboptimal_zones)} - unsafe_headings
    return {
        point.heading: "unsafe" if point.heading in unsafe_headings else "suboptimal" if point.heading in suboptimal_headings else "safe"
        for point in landing_points
    }


def get_flights(params: SimParams, attr_name: str) -> list[Flight]:
    """Return a flight list from params.runtime by attribute name, or empty list when not set."""
    value = getattr(params.runtime, attr_name, None)
    if value is None:
        return []
    return ensure_list(value)


def landing_distance(flight: Flight) -> float:
    """Return the horizontal distance in meters between the launch rail and the flight's impact point."""
    return float(np.hypot(flight.x_impact, flight.y_impact))


def lands_in_zone(x_east_m: float, y_north_m: float, zones: dict) -> bool:
    """Return True when the landing point lies inside any of the supplied zone polygons."""
    return any(MatplotlibPath(zone_coordinates).contains_point((x_east_m, y_north_m)) for zone_coordinates in zones.values())


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
    rows = []

    scenario_flights = [scenario_set[scenario] for scenario_set in params.runtime.scenario_sets if scenario in scenario_set]
    # Sort by the variation values
    scenario_flights.sort(key=lambda flight: (tuple(getattr(flight, "_meta", {}).values()), flight.env.name))

    for flight in scenario_flights:
        meta = getattr(flight, "_meta", {})
        row = {"Environment": flight.env.name, "Config": ",\n".join(f"{k}={v}" for k, v in meta.items()) if meta else "-"}

        if show_full_stats:
            # Flights started from a shared ascent's apogee have no rail/burn phase, so read those stats from the ascent
            ascent = getattr(flight, "ascent_flight", None) or flight
            t_rail = ascent.out_of_rail_time
            row["thrust_to_weight_at_rail_exit"] = ascent.rocket.motor.thrust(t_rail) / (ascent.rocket.total_mass(t_rail) * 9.81)
            row["velocity_at_rail_exit"] = ascent.out_of_rail_velocity
            row["stability_at_rail_exit"] = ascent.stability_margin(t_rail)

            # max instability in supersonic region
            supersonic = ascent.max_mach_number > SUPERSONIC_MACH
            row["max_instability_in_supersonic_region"] = ascent.stability_margin(ascent.max_speed_time) if supersonic else "-"
            row["apogee"] = flight.apogee - flight.env.elevation

        row["impact_speed"] = abs(flight.impact_velocity)
        row["landing_distance"] = landing_distance(flight)

        if show_full_stats:
            opening_times = {parachute.name: trigger_time + parachute.lag for trigger_time, parachute in flight.parachute_events}

            # Shock around deployment for each parachute (flights of one scenario all deploy the same parachutes)
            for parachute_name, deployment_time in opening_times.items():
                row[f"{parachute_name}_shock"] = get_shock_at_parachute_deployment(flight, deployment_time)

            if "main" in opening_times:
                row["speed_at_main_deployment"] = get_speed_at_parachute_deployment(flight, opening_times["main"])

        rows.append(row)

    if not rows:
        return None
    return print_stats_table(scenario, rows)


def stats_column(metric_key: str) -> tuple[str, int]:
    """Return the header and the decimal places of a stats value; parachute shocks are named after the parachute, e.g. "drogue_shock"."""
    if metric_key.endswith("_shock"):
        return f"{metric_key.removesuffix('_shock')} shock\n@ deployment [g]", 2
    return STATS_COLUMNS[metric_key]


def print_stats_table(scenario: str, rows: list[dict]) -> pd.DataFrame:
    """Print the "Stats for <scenario> flights" table (shared by RocketPy and OpenRocket) and return it as a DataFrame.

    Each row is {"Environment": ..., "Config": ..., <metric key>: raw value, ...}; headers and rounding come from
    STATS_COLUMNS, the columns keep the order of the keys, and a value a row does not have shows as "-".
    """
    metric_keys = list(dict.fromkeys(key for row in rows for key in row if key not in STATS_LABEL_COLUMNS))
    headers = [*STATS_LABEL_COLUMNS, *(stats_column(key)[0] for key in metric_keys)]
    table_rows = []
    for row in rows:
        values = [row.get(key, "-") for key in metric_keys]
        # Round numbers to the column's decimal places
        rounded = [round(value, stats_column(key)[1] or None) if isinstance(value, float) else value for key, value in zip(metric_keys, values)]
        table_rows.append([row[label] for label in STATS_LABEL_COLUMNS] + rounded)

    # Markdown tables can't hold line breaks: headers get spaces, multi-line cells (config) get HTML <br> breaks
    single_line_headers = [header.replace("\n", " ") for header in headers]
    markdown_rows = [[value.replace("\n", "<br>") if isinstance(value, str) else value for value in row] for row in table_rows]

    printmd(f"## Stats for {scenario} flights")
    column_alignment = ("left",) * len(STATS_LABEL_COLUMNS) + ("right",) * len(metric_keys)
    printmd(tabulate(markdown_rows, headers=single_line_headers, tablefmt="pipe", colalign=column_alignment))

    return pd.DataFrame(table_rows, columns=single_line_headers)


def calculate_safe_flights(params: SimParams, landing_points: list[LandingPoint], buffer_zones: dict, suboptimal_zones: dict) -> dict[float, str]:
    """Classify all flights as safe, suboptimal or unsafe based on landing zone membership, print the results
    and return {heading: safety label}.

    A heading is unsafe if any scenario (nominal/no_main/ballistic/payload), inclination,
    or environment caused a landing inside a buffer zone; otherwise it is suboptimal if any landed inside a suboptimal zone.
    """
    label_by_heading = classify_headings(landing_points, buffer_zones, suboptimal_zones)
    print_flight_stats(params)
    print_safety_summary(landing_points, label_by_heading, buffer_zones, suboptimal_zones)
    return label_by_heading


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
            bgcolor="rgba(255, 255, 255, 0.5)",
            width=20,
            height=20,
            borderpad=0,
            **position,
        )


def add_mode_point_traces(
    figure: go.Figure,
    landing_points: list[LandingPoint],
    label_by_heading: dict[float, str],
    visible: bool,
) -> tuple[int, int]:
    """Add one Scatter trace per scenario (colored like SCENARIO_COLORS) for one button mode; return the trace index range."""
    start_index = len(figure.data)

    # SCENARIO_COLORS also fixes the legend order: nominal, no_main, ballistic, matched, payload.
    for scenario, color in SCENARIO_COLORS.items():
        scenario_points = [point for point in landing_points if point.scenario == scenario]
        if not scenario_points:
            continue

        label = "payload_nominal" if scenario == "payload" else f"rocket_{scenario}"
        # Each point gets its own hover text, because RocketPy and OpenRocket points carry different details.
        hover_texts = [
            f"<b>{label}</b><br>"
            f"safety: {label_by_heading[point.heading]}<br>"
            f"heading: {point.heading:g}°<br>"
            f"inclination: {point.inclination:g}°<br>"
            + "".join(f"{name}: {value}<br>" for name, value in point.configuration.items())
            + f"impact: ({point.x_east_m:.1f}, {point.y_north_m:.1f}) m<br>"
            f"distance from launch: {np.hypot(point.x_east_m, point.y_north_m):.0f} m<br>"
            f"lat: {point.latitude:.5f}°; lon: {point.longitude:.5f}°"
            for point in scenario_points
        ]

        figure.add_trace(go.Scatter(
            x=[point.x_east_m for point in scenario_points],
            y=[point.y_north_m for point in scenario_points],
            mode="markers",
            marker=dict(color=color, size=8),
            name=label,
            hovertext=hover_texts,
            hovertemplate="%{hovertext}<extra></extra>",
            visible=visible,
            showlegend=visible,
        ))

    return start_index, len(figure.data)


def add_satellite_background(figure: go.Figure, image_path: Path, rail_latitude: float, rail_longitude: float) -> None:
    """Draw a GeoTIFF satellite image behind the plot, placed in meters east/north of the launch rail like the zones."""
    with rasterio.open(image_path) as image_file:
        # Convert satellite image edges from EPSG:3857 (meters) to EPSG:4326 (lat/lon degrees)
        west, south, east, north = transform_bounds(image_file.crs, "EPSG:4326", *image_file.bounds)
        rgb_pixels = image_file.read()

    # Convert the satellite image edges from lat/lon to the plot's local meters.
    west_m, north_m = latlon_to_local_xy(north, west, rail_latitude, rail_longitude)
    east_m, south_m = latlon_to_local_xy(south, east, rail_latitude, rail_longitude)

    # rasterio returns (band, row, column); PIL needs (row, column, band).
    image = Image.fromarray(np.moveaxis(rgb_pixels, 0, -1))
    jpeg_buffer = io.BytesIO()
    image.save(jpeg_buffer, format="JPEG", quality=SATELLITE_JPEG_QUALITY)
    # embed image into the HTML file
    image_source = "data:image/jpeg;base64," + base64.b64encode(jpeg_buffer.getvalue()).decode("ascii")

    # Plotly anchors a layout image at its top-left corner
    figure.add_layout_image(
        source=image_source,
        xref="x",
        yref="y",
        x=west_m,
        y=north_m,
        sizex=east_m - west_m,
        sizey=north_m - south_m,
        sizing="stretch",
        layer="below",
    )


def plot_landing_positions_with_modes(
    params: SimParams,
    exclusion_zones,
    buffer_zones,
    suboptimal_zones,
    plot_name,
    landing_points: list[LandingPoint] | None = None,
    label_by_heading: dict[float, str] | None = None,
    zones_only=False,
    flight_computer_impacts=None,
    satellite_image: Path | None = None,
    source: str = "",
):
    """Build an interactive landing-position plot with All/Safe/Suboptimal/Unsafe buttons.

    When zones_only is True only the zones are drawn. Otherwise the buttons switch between all landing_points and the
    points of each safety label in label_by_heading (see classify_headings). satellite_image is an optional GeoTIFF background.
    """
    project_path = params.project_path
    ensure_project_folders(project_path)
    figure = go.Figure()

    if satellite_image:
        environment = params.config.environment
        add_satellite_background(figure, satellite_image, environment.latitude, environment.longitude)

    # Plot from least to most critical so the red exclusion zones stay on top.
    plot_zones(figure, suboptimal_zones, label="Suboptimal zone", color="gold")
    plot_zones(figure, buffer_zones, label="Buffer zone", color="orange")
    plot_zones(figure, exclusion_zones, label="Exclusion zone", color="red")

    if not zones_only:
        zone_trace_count = len(figure.data)
        zone_legend_flags = [trace.showlegend is not False for trace in figure.data]

        # One button per mode; a Safe/Suboptimal/Unsafe button only appears when that label has landing points.
        landing_points = landing_points or []
        points_by_mode = {"All": landing_points}
        for safety_label in SAFETY_LABELS:
            labeled_points = [point for point in landing_points if label_by_heading[point.heading] == safety_label]
            if labeled_points:
                points_by_mode[safety_label.capitalize()] = labeled_points

        # Add the point traces per mode ("All" visible first) and remember each mode's trace index range.
        mode_trace_ranges = {
            mode_name: add_mode_point_traces(figure, mode_points, label_by_heading, visible=(mode_name == "All"))
            for mode_name, mode_points in points_by_mode.items()
        }

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

    if satellite_image:
        gridcolor = "gray"
    else:
        gridcolor = "lightgray"
    
    figure.update_layout(
        title=dict(text=f"Landing Positions ({source})", x=0.5, xanchor="center"),
        xaxis=dict(
            title="East / West distance from launch rail [m]",
            scaleanchor="y",
            scaleratio=1,
            showgrid=True,
            gridcolor=gridcolor,
            zeroline=True,
            zerolinecolor=gridcolor,
        ),
        yaxis=dict(
            title="North / South distance from launch rail [m]",
            showgrid=True,
            gridcolor=gridcolor,
            zeroline=True,
            zerolinecolor=gridcolor,
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

def run_notebook_display_mode(params: SimParams, exclusion_zones, buffer_zones, suboptimal_zones, satellite_image: Path | None = None):
    """Run the notebook display mode that fits the config: a single flight or variations.

    Both modes share the same All/Safe/Suboptimal/Unsafe landing-positions plot. Single-flight mode adds a
    per-flight safety summary, trajectory comparison, and detailed per-scenario prints/plots.
    With variations, the trajectory comparison is only drawn up to _CUSTOM_PLOTS_FLIGHT_LIMIT flights.
    """
    scenario_sets = params.runtime.scenario_sets
    payload_flights = get_flights(params, "flight_payload")

    landing_points = landing_points_from_runtime(params)
    label_by_heading = calculate_safe_flights(params, landing_points, buffer_zones, suboptimal_zones)

    flight_computer_impacts = ensure_list(params.runtime.flight_computer_impacts or [])

    plot_landing_positions_with_modes(
        params,
        exclusion_zones,
        buffer_zones,
        suboptimal_zones,
        plot_name="landing_positions",
        landing_points=landing_points,
        label_by_heading=label_by_heading,
        flight_computer_impacts=flight_computer_impacts,
        satellite_image=satellite_image,
        source="RocketPy"
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
