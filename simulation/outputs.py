"""
Output and post-processing helpers for the RocketPy simulation backend.
"""

import csv
from pathlib import Path

import nbformat
from tabulate import tabulate
from nbconvert import HTMLExporter

from matplotlib.path import Path as MatplotlibPath
import numpy as np
import plotly.graph_objects as go
from rocketpy.simulation import FlightDataExporter
from rocketpy import Flight, Environment, Motor, Fins, Rocket

from simulation.utils import *
from simulation.custom_print_and_plot_functions import CustomPlots, CustomPrints
from simulation.config_schema import SimParams
SCENARIO_COLORS = {"nominal": "green", "no_main": "orange", "ballistic": "red", "matched": "purple", "payload": "blue"}


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
    rocket.prints.inertia_details()
    # rocket.prints.rocket_geometrical_parameters()
    # rocket.prints.rocket_aerodynamics_quantities()
    rocket.prints.parachute_data()


def print_one_flight_with_custom_prints(flight: Flight):
    """Print flight summary using custom print helpers."""
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
    
    if params.runtime.mode_type != "Reanalysis":
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
        for scenario_name, flight in scenario_set.items():
            env_name = flight.env.name if hasattr(flight.env, "name") else flight.name

            # When variation metadata is present, include the varied settings in the label.
            if hasattr(flight, "_meta") and flight._meta:
                meta_str = ", ".join(render_meta_flat(flight._meta))
                label = f"{meta_str} | {scenario_name}"
            else:
                label = f"{env_name} | {scenario_name}"

            flights.append(flight)
            motors.append(flight.rocket.motor)
            plot_titles.append(label)
            rockets.append(flight.rocket)
            rocket_configs.append({"total_length": params.config.rocket.length / 1000})

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
    if params.runtime.mode_type != "Reanalysis":
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

    # -------------------------------------------------------------------------
    # Simulated flight traces: x, y in local meters from launch; altitude is AGL
    # -------------------------------------------------------------------------
    for flight in flights:
        env_name = flight.env.name if hasattr(flight.env, "name") else flight.name
        scenario = next((scenario for scenario in SCENARIO_COLORS if scenario in (getattr(flight, "name", "") or "")), "other")
        times = np.asarray(flight.time)
        # Match the scenario keyword inside the name so SCENARIO_COLORS picks the right color.
        color = next((c for scenario, c in SCENARIO_COLORS.items() if scenario in (getattr(flight, "name", "") or "")), None)
        # Use flat metadata label when available — flight.name may contain newlines from obj_to_pretty_label.
        if hasattr(flight, "_meta") and flight._meta:
            display_name = ", ".join(render_meta_flat(flight._meta))
        else:
            display_name = f"{env_name} | {scenario}"
        idx = len(figure.data)
        figure.add_trace(go.Scatter3d(
            x=np.array([flight.x(t) for t in times]),
            y=np.array([flight.y(t) for t in times]),
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
        hoverlabel=dict(namelength=-1),
        width=900,
        height=700,
    )

    figure.show(renderer="notebook")


# =============================================================================
# Exports
# =============================================================================

def export_all_kml(params: SimParams):
    """Export KML files for all flights."""
    exported_files = []
    project_path = params.project_path
    ensure_project_folders(project_path)

    for i, scenario_set in enumerate(params.runtime.scenario_sets, start=1):
        for flight in scenario_set.values():
            file_name = Path(f"flight_{i}_{flight.filename_label}.kml")
            file_path = project_path / "trajectory_kml" / file_name
            FlightDataExporter(flight).export_kml(file_name=file_path, altitude_mode="relativetoground")
            exported_files.append(file_path)

    return exported_files


def export_all_trajectory_csv(params: SimParams):
    """Export one CSV per flight with columns latitude, longitude, altitude (m AGL) at every ODE time step.
    By request for WARR, maybe useful, otherwise remove later."""
    exported_files = []
    project_path = params.project_path
    ensure_project_folders(project_path)

    for i, scenario_set in enumerate(params.runtime.scenario_sets, start=1):
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
    """Export a Jupyter notebook with its current outputs to a self-contained HTML file.
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

def print_scan_values(params: SimParams):
    """Print the heading and inclination values configured for scan mode."""
    headings = ensure_list(params.config.flight.heading)
    inclinations = ensure_list(params.config.flight.inclination)

    printmd("## Configured flight scan")
    print(f"Headings: {headings}")
    print(f"Inclinations: {inclinations}")


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


def print_safe_flight_details(params: SimParams):
    """Print lat/lon landing coordinates for every safe flight."""
    printmd("## Safe flight details")

    # Collect all safe flights with their scenario type label.
    type_keys = [
        ("nominal",   "safe_rocket_nominal"),
        ("no_main",   "safe_rocket_no_main"),
        ("ballistic", "safe_rocket_ballistic"),
        ("matched",   "safe_rocket_matched"),
        ("payload",   "safe_payload"),
    ]

    # Build: {heading: {inclination: {type: [(env_name, lat, lon)]}}}
    grouped = {}

    for type_label, attr_name in type_keys:
        for flight in get_flights(params, attr_name):
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
    """Print grouped details for every unsafe heading/inclination/environment combination."""
    printmd("## Unsafe flight details")

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


def print_variation_stats(params: SimParams):
    """Print a table of per-combination flight stats for rocket-component variations.
    Only runs when rocket components are being varied."""
    if not has_rocket_component_variations(params.config):
        return

    headers = ["Config", "Thrust to weight ratio\n@ rail exit", "rail exit velocity\n[m/s]", "stability @ rail exit\n[cal]", "apogee AGL\n[m]"]
    rows = []

    for scenario_set in params.runtime.scenario_sets:
        flight = scenario_set.get("nominal")
        if flight is None:
            continue

        meta = getattr(flight, "_meta", {})
        t_rail = flight.out_of_rail_time
        v_rail = flight.out_of_rail_velocity
        stability = flight.stability_margin(t_rail)
        t_w = flight.rocket.motor.thrust(t_rail) / (flight.rocket.total_mass(t_rail) * 9.81)
        apogee_agl = flight.apogee - flight.env.elevation

        config_str = ", ".join(f"{k}={v}" for k, v in meta.items()) if meta else "—"
        rows.append([config_str, f"{t_w:.2f}", f"{v_rail:.1f}", f"{stability:.2f}", f"{apogee_agl:.0f}"])

    if not rows:
        return

    printmd("## Variation Stats")
    print(tabulate(rows, headers=headers, tablefmt="simple", colalign=("left", "right", "right", "right", "right")))


def calculate_safe_flights(params: SimParams, buffer_zones: dict):
    """Classify all flights as safe or unsafe based on landing zone membership.

    A heading is unsafe if any scenario (nominal/no_main/ballistic/payload), inclination,
    or environment caused a landing inside a buffer zone.
    """
    payload_flights = get_flights(params, "flight_payload")
    
    # Group payload flights by their (env, heading, inclination) for the diagnostic detail list.
    payloads_by_config = {}
    for payload_flight in payload_flights:
        payloads_by_config.setdefault(scenario_config_key(payload_flight), []).append(payload_flight)

    scenario_sets = params.runtime.scenario_sets

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

    store_safety_results(params, "safe", safe_scenario_sets)
    store_safety_results(params, "unsafe", unsafe_scenario_sets)
    params.runtime.unsafe_details = unsafe_details

    # Payload flights inherit the heading-level classification.
    safe_payload_flights = [payload for payload in payload_flights if payload.heading not in unsafe_headings]
    unsafe_payload_flights = [payload for payload in payload_flights if payload.heading in unsafe_headings]
    params.runtime.safe_payload = safe_payload_flights
    params.runtime.unsafe_payload = unsafe_payload_flights

    print_variation_stats(params)
    print_configurations("Safe Configurations by Heading", params.runtime.safe_configurations or [])
    print_safe_flight_details(params)
    print_configurations("Unsafe Configurations by Heading", params.runtime.unsafe_configurations or [])
    print_unsafe_details(unsafe_details)


# =============================================================================
# Safety plots
# =============================================================================

def build_safe_unsafe_flight_groups(params: SimParams):
    """Build safe and unsafe flight-group dicts from params.runtime safety lists."""
    safe_flight_groups = {
        "rocket_nominal":  (get_flights(params, "safe_rocket_nominal"),  SCENARIO_COLORS["nominal"]),
        "rocket_no_main":  (get_flights(params, "safe_rocket_no_main"),  SCENARIO_COLORS["no_main"]),
        "rocket_ballistic":(get_flights(params, "safe_rocket_ballistic"), SCENARIO_COLORS["ballistic"]),
        "rocket_matched":  (get_flights(params, "safe_rocket_matched"),  SCENARIO_COLORS["matched"]),
    }
    unsafe_flight_groups = {
        "rocket_nominal":  (get_flights(params, "unsafe_rocket_nominal"),  SCENARIO_COLORS["nominal"]),
        "rocket_no_main":  (get_flights(params, "unsafe_rocket_no_main"),  SCENARIO_COLORS["no_main"]),
        "rocket_ballistic":(get_flights(params, "unsafe_rocket_ballistic"), SCENARIO_COLORS["ballistic"]),
        "rocket_matched":  (get_flights(params, "unsafe_rocket_matched"),  SCENARIO_COLORS["matched"]),
    }

    safe_payload_flights = get_flights(params, "safe_payload")
    unsafe_payload_flights = get_flights(params, "unsafe_payload")
    if safe_payload_flights or unsafe_payload_flights:
        safe_flight_groups["payload_nominal"] = (safe_payload_flights, SCENARIO_COLORS["payload"])
        unsafe_flight_groups["payload_nominal"] = (unsafe_payload_flights, SCENARIO_COLORS["payload"])

    return safe_flight_groups, unsafe_flight_groups


def build_safety_by_config(params: SimParams):
    """Build a {(environment, heading, inclination): 'safe'|'unsafe'} map from params.runtime."""
    safety_by_config = {}

    for configuration in (params.runtime.safe_configurations or []):
        safety_by_config[configuration] = "safe"

    for configuration in (params.runtime.unsafe_configurations or []):
        safety_by_config[configuration] = "unsafe"

    return safety_by_config


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
    params: SimParams,
    exclusion_zones,
    buffer_zones,
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

def run_notebook_display_mode(params, exclusion_zones, buffer_zones):
    """Run the correct notebook display mode for the configured heading/inclination values.

    Both modes share the same All/Safe/Unsafe landing-positions plot. Single-flight mode adds a
    per-flight safety summary, trajectory comparison, and detailed per-scenario prints/plots.
    Scan mode adds the configured heading/inclination value list.
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

    calculate_safe_flights(params, buffer_zones)

    safe_flight_groups, unsafe_flight_groups = build_safe_unsafe_flight_groups(params)
    safety_by_config = build_safety_by_config(params)

    # Only display the Safe/Unsafe buttons when those groups actually contain flights.
    mode_flight_groups = {"All": all_flight_groups}
    if any(flights for flights, _color in safe_flight_groups.values()):
        mode_flight_groups["Safe"] = safe_flight_groups
    if any(flights for flights, _color in unsafe_flight_groups.values()):
        mode_flight_groups["Unsafe"] = unsafe_flight_groups

    flight_computer_impacts = ensure_list(params.runtime.flight_computer_impacts or [])

    plot_landing_positions_with_modes(
        params,
        exclusion_zones,
        buffer_zones,
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
                print_one_flight_with_custom_prints(flight)

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
