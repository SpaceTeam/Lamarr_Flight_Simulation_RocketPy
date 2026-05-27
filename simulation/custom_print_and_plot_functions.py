import re
import numpy as np
from colorama import Fore, Style                                  # https://github.com/tartley/colorama
import math
from pathlib import Path
from scipy.spatial.transform import Rotation
from rocketpy import SolidMotor, LiquidMotor, HybridMotor, Rocket, Flight, Parachute, Environment
import plotly.graph_objects as go
import matplotlib.pyplot as plt

class CustomPrints:
    """
    Store rocket simulation objects and provide reusable custom printing methods.

    Attributes
    ----------
    flight_forecast : Flight
        Flight forecast to print information from.
    """

    def __init__(
        self,
        flight_forecast: Flight,
    ):
        self.flight_forecast = flight_forecast


    def get_altitude_asl(self, time: float) -> float:
        """
        Return the rocket altitude above sea level at the given time.
        """
        return self.flight_forecast.z(time)


    def get_altitude_agl(self, time: float) -> float:
        """
        Return the rocket altitude above ground level at the given time.
        """
        altitude_asl = self.get_altitude_asl(time)
        launch_site_elevation = self.flight_forecast.env.elevation
        return altitude_asl - launch_site_elevation


    def apogee_conditions(self):
        """
        Print selected flight conditions at apogee.
        """
        if not hasattr(self.flight_forecast, "apogee_time"):
            return
        
        apogee_time = self.flight_forecast.apogee_time
        altitude_asl = self.get_altitude_asl(apogee_time)
        altitude_agl = self.get_altitude_agl(apogee_time)
        
        # horizontal speed
        velocity_x = self.flight_forecast.vx(apogee_time)
        velocity_y = self.flight_forecast.vy(apogee_time)
        horizontal_speed = math.sqrt(velocity_x**2 + velocity_y**2)

        print("\nApogee Conditions\n")
        print(f"Time: {apogee_time:.3f} s")
        print(f"Altitude: {altitude_asl:.3f} m (ASL) | {altitude_agl:.3f} m (AGL)")
        print(f"Horizontal speed: {horizontal_speed:.3f} m/s")


    def parachute_events(self):
        """
        Print selected parachute event information.
        """
        print("\nParachute Events\n")
        
        if not hasattr(self.flight_forecast, "parachute_events") or not self.flight_forecast.parachute_events:
            print("    No Parachute Events Were Triggered.")
            return

        for ejection_time, parachute in self.flight_forecast.parachute_events:
            inflation_time = ejection_time + parachute.lag

            altitude_asl = self.get_altitude_asl(inflation_time)
            altitude_agl = self.get_altitude_agl(inflation_time)

            time_samples = np.asarray(self.flight_forecast.time, dtype=float)
            time_mask = ((inflation_time - 0.25 <= time_samples) & (time_samples <= inflation_time + 0.25))
            inflation_time_samples = time_samples[time_mask]

            acceleration_values = np.array([self.flight_forecast.acceleration(time) for time in inflation_time_samples], dtype=float)
            az_values = np.array([self.flight_forecast.az(time) for time in inflation_time_samples], dtype=float)

            max_acceleration = np.max(acceleration_values)
            max_vertical_acceleration = np.max(az_values)

            gravity = 9.81
            shock_g = max_acceleration / gravity

            print(f"Parachute: {parachute.name}")
            print(f"    Altitude at inflation: {altitude_asl:.3f} m (ASL) | {altitude_agl:.3f} m (AGL)")
            print(f"    Ejection time: {ejection_time:.3f} s")
            print(f"    Inflation time: {inflation_time:.3f} s")
            print(f"    Max acceleration: {max_acceleration:.3f} m/s², max in ±0.25 s around inflation")
            print(f"    Max vertical acceleration: {max_vertical_acceleration:.3f} m/s², max in ±0.25 s around inflation")
            print(f"    Shock: {shock_g:.3f} g, max in ±0.25 s around inflation\n")
                

    def impact_coordinates(self):
        """
        Print impact coordinates in degrees, minutes and seconds.
        https://www.latlong.net/lat-long-dms.html
        https://tnp.uservoice.com/knowledgebase/articles/172110-latitude-longitude-formats-and-conversion
        """
        coordinates = {
            "latitude": self.flight_forecast.latitude(self.flight_forecast.t_final),
            "longitude": self.flight_forecast.longitude(self.flight_forecast.t_final),
        }

        for coordinate_name, decimal_degrees in coordinates.items():
            absolute_degrees = abs(decimal_degrees)
            degrees = int(absolute_degrees)
            remaining_minutes = (absolute_degrees - degrees) * 60
            minutes = int(remaining_minutes)
            seconds = (remaining_minutes - minutes) * 60

            if coordinate_name == "latitude":
                direction = "N" if decimal_degrees >= 0 else "S"
            else:
                direction = "E" if decimal_degrees >= 0 else "W"

            print(f"Impact {coordinate_name}: {degrees}° {minutes}' {seconds:.3f}\" {direction}")


class CustomPlots:
    """
    Store rocket simulation objects and provide reusable custom plotting methods.
    
    Attributes
    ----------
    CustomPlots.flight_forecasts : list[Flight]
        Flight forecasts to plot.
    CustomPlots.motors : list[SolidMotor | LiquidMotor | HybridMotor]
        Used motors for the flight forecasts.
    CustomPlots.plot_titles: list[str]
        Used labels for the flight forecasts.
    CustomPlots.rockets : list[Rocket]
        Used rockets for the flight forecasts.
    CustomPlots.rocket_configs : list[dict]
        Used rocket config dicts for the flight forecasts.
    """
    def __init__(
        self,
        flight_forecast: list[Flight],
        motor: list[SolidMotor | LiquidMotor | HybridMotor],
        plot_title: list[str],
        rocket: list[Rocket],
        rocket_config: list[dict],
        save_dir: Path | None = None,
    ):
        list_lengths = [len(flight_forecast), len(motor), len(plot_title), len(rocket), len(rocket_config)]

        if any(list_length != list_lengths[0] for list_length in list_lengths):
            raise ValueError("CustomPlots input lists must all have the same length.")

        self.flight_forecasts = flight_forecast
        self.motors = motor
        self.plot_titles = plot_title
        self.rockets = rocket
        self.rocket_configs = rocket_config
        # When set, every create_grouped_plotly_plot call saves an HTML file here.
        self.save_dir = save_dir


    def get_time_samples_for_flight(self, flight: Flight, time_start: float, time_end: float) -> np.ndarray:
        """
        Get time samples for the selected flight between the given start and end time.
        """
        time_samples = np.asarray(flight.time, dtype=float)
        time_mask = (time_start <= time_samples) & (time_samples <= time_end)
        time_samples = time_samples[time_mask]
        return time_samples
        
        
    def find_threshold_interval(
        self,
        values: np.ndarray,
        time_samples: np.ndarray,
        threshold: float,
    ) -> tuple[float | None, float | None]:
        """
        Find the first time interval where the given values are above a threshold.

        Returns the start and end time of the interval. If the values never cross the
        threshold, both values are None.
        """
        start_time = None
        end_time = None

        if values is None:
            return start_time, end_time
        
        for index, value in enumerate(values):
            if value >= threshold and start_time is None:
                start_time = time_samples[index]

            elif value < threshold and start_time is not None and end_time is None:
                end_time = time_samples[index]
                break

        return start_time, end_time
    

    def add_event_marker_lines(
        self,
        figure: go.Figure,
        time_samples: np.ndarray,
        event_markers: list[tuple[float, str, str]],
    ):
        """
        Filter, sort, and draw vertical event marker lines on a Plotly figure.
        """
        shapes, annotations = self.get_event_marker_layout_parts(time_samples, event_markers)
        existing_shapes = list(figure.layout.shapes) if figure.layout.shapes else []
        existing_annotations = list(figure.layout.annotations) if figure.layout.annotations else []
        figure.update_layout(
            shapes=existing_shapes + shapes,
            annotations=existing_annotations + annotations,
        )


    def get_event_marker_layout_parts(
        self,
        time_samples: np.ndarray,
        event_markers: list[tuple[float, str, str]],
    ) -> tuple[list[dict], list[dict]]:
        """
        Build Plotly shapes and annotations for event markers in the plotted time range.
        """
        # Keep only markers that have a time and fall inside the plotted time range.
        plot_start_time = float(time_samples[0])
        plot_end_time = float(time_samples[-1])
        event_markers = [
            event_marker
            for event_marker in event_markers
            if event_marker[0] is not None and plot_start_time <= event_marker[0] <= plot_end_time
        ]

        # Draw markers in time order so close labels can alternate sides predictably.
        event_markers.sort(key=lambda event_marker: event_marker[0])

        previous_event_time = None
        close_event_threshold = 2.5  # seconds
        close_event_side = "left"
        shapes = []
        annotations = []

        for event_time, label, color in event_markers:
            annotation_side = "right"

            # Alternate labels when event markers are close enough to overlap.
            if previous_event_time is not None:
                time_difference = event_time - previous_event_time

                if time_difference <= close_event_threshold:
                    annotation_side = close_event_side
                    close_event_side = "right" if close_event_side == "left" else "left"
                else:
                    close_event_side = "left"

            if annotation_side == "left":
                xshift = -5
                xanchor = "right"
            else:
                xshift = 4
                xanchor = "left"

            shapes.append(
                {
                    "type": "line",
                    "x0": event_time,
                    "x1": event_time,
                    "y0": 0,
                    "y1": 1,
                    "xref": "x",
                    "yref": "paper",
                    "line": {
                        "color": color,
                        "width": 1.5,
                        "dash": "dash",
                    },
                }
            )
            annotations.append(
                {
                    "x": event_time,
                    "y": 0.5,
                    "xref": "x",
                    "yref": "paper",
                    "text": label,
                    "showarrow": False,
                    "textangle": -90,
                    "font": {"color": color},
                    "xanchor": xanchor,
                    "yanchor": "middle",
                    "align": "center",
                    "xshift": xshift,
                }
            )

            previous_event_time = event_time

        return shapes, annotations


    def get_standard_event_markers_for_flight(
        self,
        flight: Flight,
        time_samples: np.ndarray,
        motor=None,
    ) -> list[tuple[float, str, str]]:
        """
        Build the standard event marker list for one flight. For RocketPy flights, motor is taken from
        `flight.rocket.motor` so matched/spliced flights get their own (overridden) burnout. For flight-computer
        data (no .rocket attribute), pass a `motor` explicitly to still draw the burnout marker.
        """
        if hasattr(flight, "out_of_rail_time"):
            out_of_rail_time = float(flight.out_of_rail_time)
        else:
            out_of_rail_time = None

        # RocketPy flights carry their own motor; flight-computer data needs an explicit motor fallback.
        if hasattr(flight, "rocket"):
            burn_out_time = float(flight.rocket.motor.burn_out_time)
        elif motor is not None:
            burn_out_time = float(motor.burn_out_time)
        else:
            burn_out_time = None

        if hasattr(flight, "apogee_time"):
            apogee_time = float(flight.apogee_time)
        else:
            apogee_time = None

        if hasattr(flight, "t_final"):
            ground_hit_time = float(flight.t_final)
        else:
            ground_hit_time = None

        if hasattr(flight, "mach_number"):
            mach_number = np.array([flight.mach_number(time) for time in time_samples], dtype=float)
        else:
            mach_number = None

        transonic_start_time, transonic_end_time = self.find_threshold_interval(
            mach_number,
            time_samples,
            threshold=0.8,
        )
        supersonic_start_time, supersonic_end_time = self.find_threshold_interval(
            mach_number,
            time_samples,
            threshold=1.2,
        )

        event_markers = [
            (out_of_rail_time, "Out Of Rail", "red"),
            (burn_out_time, "Motor Burn Out", "green"),
            (apogee_time, "Apogee", "magenta"),
            (ground_hit_time, "Ground Hit", "black"),
            (transonic_start_time, "Transonic Entry", "deepskyblue"),
            (transonic_end_time, "Transonic Exit", "deepskyblue"),
            (supersonic_start_time, "Supersonic Entry", "blue"),
            (supersonic_end_time, "Supersonic Exit", "blue"),
        ]

        if hasattr(flight, "parachute_events"):
            for ejection_time, parachute in flight.parachute_events:
                inflation_time = ejection_time + parachute.lag
                event_markers.append((float(ejection_time), f"{parachute.name} parachute ejected", "black"))
                event_markers.append((float(inflation_time), f"{parachute.name} parachute inflated", "black"))

        return event_markers


    def create_grouped_plotly_plot(
        self,
        title: str,
        flight_groups: list[dict],
        yaxis_title: str,
        yaxis2_title: str | None = None,
        yaxis_dtick: float | int | None = None,
        yaxis2_dtick: float | int | None = None,
        width: int = 1100,
        height: int = 500,
        legend_x_pos: float = 1.12,
        buttons_x_pos=1.1,
    ):
        """
        Create a Plotly plot where buttons switch between one flight group at a time.
        """
        figure = go.Figure()
        visibility = []
        marker_shapes = []
        marker_annotations = []

        is_single_group = len(flight_groups) == 1

        for group_index, flight_group in enumerate(flight_groups):
            shapes, annotations = self.get_event_marker_layout_parts(
                time_samples=flight_group["time_samples"],
                event_markers=flight_group["event_markers"],
            )
            marker_shapes.append(shapes)
            marker_annotations.append(annotations)

            # In single-group plots, drop the label prefix so trace names stay clean (e.g. "Altitude [m]"
            # instead of "CATS Vega - Altitude [m]"). In multi-group plots the prefix disambiguates groups.
            label_prefix = "" if is_single_group else f"{flight_group['label']} - "
            hover_prefix = "" if is_single_group else f"{flight_group['label']}<br>Time: %{{x:.1f}} s, "

            for trace in flight_group["traces"]:
                scatter_options = {
                    "x": flight_group["time_samples"],
                    "y": trace["y"],
                    "mode": "lines",
                    "name": f"{label_prefix}{trace['name']}",
                    "hovertemplate": f"{hover_prefix}{trace['hovertemplate']}",
                    "line": trace["line"],
                    "visible": True,
                    "showlegend": True,
                }

                # optional right y axis
                if trace.get("yaxis") == "y2":
                    scatter_options["yaxis"] = "y2"

                figure.add_trace(go.Scatter(**scatter_options))
                visibility.append(True)

        yaxis_settings = {
            "title": yaxis_title,
            "showgrid": True,
        }

        if yaxis_dtick is not None:
            yaxis_settings["dtick"] = yaxis_dtick

        all_marker_shapes = [shape for shapes in marker_shapes for shape in shapes]
        all_marker_annotations = [annotation for annotations in marker_annotations for annotation in annotations]
        all_time_start = min(flight_group["time_start"] for flight_group in flight_groups)
        all_time_end = max(flight_group["time_end"] for flight_group in flight_groups)

        layout_settings = {
            "title": {
                "text": title,
                "x": 0.5,
                "xanchor": "center",
            },
            "width": width,
            "height": height,
            "xaxis": {
                "title": "Time [s]",
                "range": [all_time_start, all_time_end + 1],
                "showgrid": True,
                "hoverformat": ".3f",
            },
            "yaxis": yaxis_settings,
            "legend": {
                "orientation": "v",
                "yanchor": "top",
                "y": 1,
                "xanchor": "left",
                "x": legend_x_pos,
            },
            "shapes": all_marker_shapes,
            "annotations": all_marker_annotations,
            "template": "plotly_white",
        }

        # Single-group plots get a unified hover line (all traces share one crosshair tooltip).
        if is_single_group:
            layout_settings["hovermode"] = "x unified"
            layout_settings["xaxis"]["unifiedhovertitle"] = {"text": "Time: %{x:.3f} s"}

        if yaxis2_title is not None:
            yaxis2_settings = {
                "title": yaxis2_title,
                "overlaying": "y",
                "side": "right",
                "showgrid": False,
            }

            if yaxis2_dtick is not None:
                yaxis2_settings["dtick"] = yaxis2_dtick

            layout_settings["yaxis2"] = yaxis2_settings

        if len(flight_groups) > 1:
            buttons = [
                {
                    "label": "All",
                    "method": "update",
                    "args": [
                        {"visible": [True] * len(visibility), "showlegend": [True] * len(visibility)},
                        {
                            "shapes": all_marker_shapes,
                            "annotations": all_marker_annotations,
                            "xaxis.range": [all_time_start, all_time_end + 1],
                        },
                    ],
                }
            ]
            trace_start = 0
            trace_ranges = []

            for flight_group in flight_groups:
                trace_end = trace_start + len(flight_group["traces"])
                trace_ranges.append((trace_start, trace_end))
                trace_start = trace_end

            for group_index, flight_group in enumerate(flight_groups):
                group_visibility = [False] * len(visibility)
                group_showlegend = [False] * len(visibility)
                trace_start, trace_end = trace_ranges[group_index]

                for trace_index in range(trace_start, trace_end):
                    group_visibility[trace_index] = True
                    group_showlegend[trace_index] = True

                buttons.append(
                    {
                        "label": flight_group["label"],
                        "method": "update",
                        "args": [
                            {"visible": group_visibility, "showlegend": group_showlegend},
                            {
                                "shapes": marker_shapes[group_index],
                                "annotations": marker_annotations[group_index],
                                "xaxis.range": [flight_group["time_start"], flight_group["time_end"] + 1],
                            },
                        ],
                    }
                )

            layout_settings["updatemenus"] = [
                {
                    "buttons": buttons,
                    "direction": "up",
                    "showactive": True,
                    "x": buttons_x_pos,
                    "xanchor": "right",
                    "y": -0.05,
                    "yanchor": "top",
                }
            ]
            layout_settings["margin"] = {"b": 80}

        figure.update_layout(**layout_settings)

        if self.save_dir is not None:
            # Build a unique filename from the first flight label + plot title, with filesystem-safe chars.
            # first_label = re.sub(r"[^\w\-]", "_", self.plot_titles[0])
            plot_slug  = re.sub(r"[^\w\-]", "_", title)
            # if is_single_group:
            #     filename = f"{plot_slug}.html"                
            # else:
            #     filename = f"{first_label}_{plot_slug}.html"
            filename = f"{plot_slug}.html"    
            figure.write_html(str(self.save_dir / filename))

        figure.show(renderer="notebook")


    def plot_stability_and_cg_cp_position(self, use_openrocket_coordinates=True):
        """
        Plot rocket CG (center of mass) and CP (center of pressure) positions over time and also Mach number + stability margin.
        - CG is evaluated as a function of time.
        - CP is evaluated as a function of Mach, using Mach(t) from the flight.
        - Rocket length must be given in mm.
            
        Parameters
        ----------
        use_openrocket_coordinates : boolean, optional
            Whether to use RocketPy (tail to nose) or OpenRocket (nose to tail) coordinate system.
            tail to nose: 0 = tail, top = nose; nose to tail: 0 = nose, top = tail.
            Default is True.
                    
        Returns
        -------
        None    
        """
        if use_openrocket_coordinates:
            position_axis_title = "Position along rocket axis [mm] (0 = nose)"
            title_suffix = "OpenRocket CG/CP orientation"
        else:
            position_axis_title = "Position along rocket axis [mm] (0 = tail)"
            title_suffix = "RocketPy CG/CP orientation"

        flight_groups = []
        for flight, motor, plot_title, rocket, rocket_config in zip(
            self.flight_forecasts,
            self.motors,
            self.plot_titles,
            self.rockets,
            self.rocket_configs,
        ):
            time_start = 0.0
            time_end = float(flight.apogee_time)
            time_samples = self.get_time_samples_for_flight(flight, time_start, time_end)

            # --- Flight data ---
            # CG(t): center of mass position along rocket axis (in rocket reference system)
            cg_position = np.array([rocket.center_of_mass(time) for time in time_samples], dtype=float) * 1000

            # Mach(t) from the flight, then CP(Mach(t))
            mach_number = np.array([flight.mach_number(time) for time in time_samples], dtype=float)
            cp_position = np.array([rocket.cp_position(mach) for mach in mach_number], dtype=float) * 1000
            stability_margin = np.array([flight.stability_margin(time) for time in time_samples], dtype=float)

            if use_openrocket_coordinates:
                rocket_length = rocket_config["total_length"]
                cg_position = rocket_length - cg_position
                cp_position = rocket_length - cp_position

            traces = [
                {
                    "y": mach_number,
                    "name": "Mach number",
                    "hovertemplate": "Mach: %{y:.2f}<extra></extra>",
                    "line": {"color": "darkturquoise"},
                },
                {
                    "y": stability_margin,
                    "name": "Stability margin [c]",
                    "hovertemplate": "Stability: %{y:.2f} c<extra></extra>",
                    "line": {"color": "royalblue", "dash": "dash"},
                },
                {
                    "y": cg_position,
                    "name": "CG position [mm]",
                    "hovertemplate": "CG: %{y:.1f} mm<extra></extra>",
                    "line": {"color": "gold"},
                    "yaxis": "y2",
                },
                {
                    "y": cp_position,
                    "name": "CP position [mm]",
                    "hovertemplate": "CP: %{y:.1f} mm<extra></extra>",
                    "line": {"color": "firebrick"},
                    "yaxis": "y2",
                },
            ]

            flight_groups.append(
                {
                    "label": plot_title,
                    "time_samples": time_samples,
                    "time_start": time_start,
                    "time_end": time_end,
                    "traces": traces,
                    "event_markers": self.get_standard_event_markers_for_flight(flight, time_samples, motor=motor),
                }
            )

        self.create_grouped_plotly_plot(
            title=f"CG/CP position, Mach, Stability ({title_suffix})",
            flight_groups=flight_groups,
            yaxis_title="Mach / Stability margin [c]",
            yaxis2_title=position_axis_title,
            yaxis_dtick=0.5,
            yaxis2_dtick=100,
            width=1100,
            height=550,
        )

    def plot_angle_of_attack_and_attitude_angle(self):
        """
        Plot angle of attack and attitude angle until apogee.
            - angle of attack(t): angle between the rocket's velocity vector and its longitudinal axis.
            - attitude angle(t): angle between the rocket's longitudinal axis and the local horizontal plane th earth's surface.
                In OpenRocket the attitude angle is called "vertical orientation (zenith)".
            - heading(t): angle from north to velocity vector in xy plane (0=north, 90=east)
        """
        flight_groups = []
        for flight, motor, plot_title in zip(self.flight_forecasts, self.motors, self.plot_titles):
            time_start = 0.0

            inflation_times = []
            if hasattr(flight, "parachute_events"):
                for ejection_time, parachute in flight.parachute_events:
                    inflation_times.append(ejection_time + parachute.lag)
            first_inflation_time = min(inflation_times) if inflation_times else None

            time_end = first_inflation_time if first_inflation_time is not None else float(flight.apogee_time)
            time_samples = self.get_time_samples_for_flight(flight, time_start, time_end)

            angle_of_attack = np.array([flight.angle_of_attack(time) for time in time_samples], dtype=float)
            attitude_angle = np.array([flight.attitude_angle(time) for time in time_samples], dtype=float)

            # Heading (compass bearing of the velocity vector):
            vx = np.array([flight.vx(time) for time in time_samples], dtype=float)
            vy = np.array([flight.vy(time) for time in time_samples], dtype=float)
            heading = np.degrees(np.arctan2(vx, vy)) % 360.0    # wrapped to [0, 360)

            traces = [
                {
                    "y": angle_of_attack,
                    "name": "Angle of attack [°]",
                    "hovertemplate": "Angle of attack: %{y:.2f} °<extra></extra>",
                    "line": {"color": "royalblue"},
                },
                {
                    "y": attitude_angle,
                    "name": "Attitude angle [°]",
                    "hovertemplate": "Attitude angle: %{y:.2f} °<extra></extra>",
                    "line": {"color": "firebrick"},
                },
                {
                    "y": heading,
                    "name": "Heading [°]",
                    "hovertemplate": "Heading: %{y:.2f} °<extra></extra>",
                    "line": {"color": "darkgreen"},
                    "yaxis": "y2",
                },
            ]

            flight_groups.append(
                {
                    "label": plot_title,
                    "time_samples": time_samples,
                    "time_start": time_start,
                    "time_end": time_end,
                    "traces": traces,
                    "event_markers": self.get_standard_event_markers_for_flight(flight, time_samples, motor=motor),
                }
            )

        self.create_grouped_plotly_plot(
            title="Angle of attack, attitude angle, and heading over time",
            flight_groups=flight_groups,
            yaxis_title="Angle [°]",
            yaxis2_title="Heading [°]",
            width=1100,
            height=500,
        )


    def plot_angular_velocity(self, transform_openrocket=False):
        """
        Plot angular velocity until first parachute inflation or apogee.

        transform_openrocket=False: RocketPy body reference frame is used.
            - angular_rate_1 (w1): pitch rate (angular velocity around the rocket's lateral axis / x axis)
            - angular_rate_2 (w2): yaw rate (angular velocity around the rocket's vertical axis / y axis)
            - angular_rate_3 (w3): roll rate (angular velocity around the rocket's longitudinal axis / z axis)

        https://www1.grc.nasa.gov/beginners-guide-to-aeronautics/rocket-rotations/

        transform_openrocket=True: OpenRocket flow-aligned reference frame (same as body frame when angle of attack α and
        sideslip β are zero).
            The body frame is rotated around the longitudinal z-axis by
                theta = atan2(airspeed y component in body frame, airspeed x component in body frame)
            where airspeed = rocket velocity - wind direction, expressed in the body frame (matches OpenRocket).

            Then:
            - angular_rate_1: pitch rate (rotation around flow-aligned reference frame y-axis; changes with angle of attack)
            - angular_rate_2: yaw rate (rotation around flow-aligned reference frame x-axis / airspeed lateral direction; changes with sideslip angle)
            - angular_rate_3: roll rate (same as body frame)

        See: https://www.researchgate.net/figure/Rocket-orientation-angles-in-Earth-frame-and-angle-of-attack-a-and-sideslip-angle-b_fig4_360423890
        with X_b = OpenRocket z, Y_b = OpenRocket y, Z_b = OpenRocket x.
        """
        # TODO: check this! (the result matches OpenRocket, which is good, but I still want to double check the math and the axis mapping)
        frame_label = "OpenRocket flow-aligned reference frame" if transform_openrocket else "RocketPy body reference frame"
        flight_groups = []

        for flight, motor, plot_title in zip(self.flight_forecasts, self.motors, self.plot_titles):
            time_start = 0.0

            inflation_times = []
            if hasattr(flight, "parachute_events"):
                for ejection_time, parachute in flight.parachute_events:
                    inflation_times.append(ejection_time + parachute.lag)
            first_inflation_time = min(inflation_times) if inflation_times else None

            time_end = first_inflation_time if first_inflation_time is not None else float(flight.apogee_time)
            time_samples = self.get_time_samples_for_flight(flight, time_start, time_end)

            if transform_openrocket:
                # Transform body-frame angular velocity to the flow-aligned reference frame.
                # pitch_rate = flow-aligned reference frame y-component = -sin(theta)*w1 + cos(theta)*w2
                # yaw_rate   = flow-aligned reference frame x-component =  cos(theta)*w1 + sin(theta)*w2
                # Matches OpenRocket AbstractSimulationStepper: pitchRate=rot.getY(), yawRate=rot.getX().

                # Angular velocity (in body frame)
                w1 = np.array([flight.w1(t) for t in time_samples])
                w2 = np.array([flight.w2(t) for t in time_samples])

                # Rocket orientation quaternions
                quaternions = np.column_stack([
                    [flight.e1(t) for t in time_samples],
                    [flight.e2(t) for t in time_samples],
                    [flight.e3(t) for t in time_samples],
                    [flight.e0(t) for t in time_samples],
                ])

                # Airspeed in inertial frame: rocket velocity - wind (matches OpenRocket convention)
                vx = np.array([flight.vx(t) for t in time_samples])
                vy = np.array([flight.vy(t) for t in time_samples])
                vz = np.array([flight.vz(t) for t in time_samples])
                z_alts = np.array([flight.z(t) for t in time_samples])
                wind_vx = np.array([flight.env.wind_velocity_x(z) for z in z_alts])
                wind_vy = np.array([flight.env.wind_velocity_y(z) for z in z_alts])
                airspeed_inertial = np.column_stack([vx - wind_vx, vy - wind_vy, vz])

                # Rotate airspeed to body frame (inverse of rocket orientation)
                airspeed_body = Rotation.from_quat(quaternions).inv().apply(airspeed_inertial)

                # Theta = azimuthal angle of airspeed lateral component in body frame
                len_xy = np.hypot(airspeed_body[:, 0], airspeed_body[:, 1])
                valid = len_xy > 0.001
                safe_len = np.where(valid, len_xy, 1.0)   # avoid division by zero
                cos_theta = np.where(valid, airspeed_body[:, 0] / safe_len, 1.0)
                sin_theta = np.where(valid, airspeed_body[:, 1] / safe_len, 0.0)

                # invRotateZ by theta: rotate body frame around z by -theta
                angular_rate_1 = np.degrees(-sin_theta * w1 + cos_theta * w2)   # pitch (flow-aligned reference frame y)
                angular_rate_2 = np.degrees( cos_theta * w1 + sin_theta * w2)   # yaw   (flow-aligned reference frame x)
            else:
                angular_rate_1 = np.degrees(np.array([flight.w1(time) for time in time_samples], dtype=float))
                angular_rate_2 = np.degrees(np.array([flight.w2(time) for time in time_samples], dtype=float))

            angular_rate_3 = np.degrees(np.array([flight.w3(time) for time in time_samples], dtype=float))

            traces = [
                {
                    "y": angular_rate_1,
                    "name": "pitch rate [°/s]",
                    "hovertemplate": "pitch rate: %{y:.3f} °/s<extra></extra>",
                    "line": {"color": "royalblue"},
                },
                {
                    "y": angular_rate_2,
                    "name": "yaw rate [°/s]",
                    "hovertemplate": "yaw rate: %{y:.3f} °/s<extra></extra>",
                    "line": {"color": "firebrick"},
                },
                {
                    "y": angular_rate_3,
                    "name": "roll rate [°/s]",
                    "hovertemplate": "roll rate: %{y:.3f} °/s<extra></extra>",
                    "line": {"color": "gold"},
                },
            ]

            flight_groups.append(
                {
                    "label": plot_title,
                    "time_samples": time_samples,
                    "time_start": time_start,
                    "time_end": time_end,
                    "traces": traces,
                    "event_markers": self.get_standard_event_markers_for_flight(flight, time_samples, motor=motor),
                }
            )

        self.create_grouped_plotly_plot(
            title=f"Angular velocity ({frame_label})",
            flight_groups=flight_groups,
            yaxis_title="Angular rate [°/s]",
            width=1100,
            height=500,
            legend_x_pos=1.02,
            buttons_x_pos=1,
        )


    def plot_motion_over_time(self, time_interval=None, event_markers=None):
        """
        Plot altitude, vertical velocity, horizontal velocity, and vertical acceleration over time.
        """
        custom_event_markers = None
        if event_markers is not None:
            custom_event_markers = event_markers
            if len(self.flight_forecasts) == 1 and (not event_markers or isinstance(event_markers[0], tuple)):
                custom_event_markers = [event_markers]

            if len(custom_event_markers) != len(self.flight_forecasts):
                raise ValueError("event_markers must contain one marker list for each flight.")

        flight_groups = []
        for flight_index, (flight, motor, plot_title) in enumerate(zip(self.flight_forecasts, self.motors, self.plot_titles)):
            if time_interval:
                time_start = time_interval[0]
                time_end = time_interval[1]
            else:
                time_start = 0.0
                time_end = float(flight.t_final)

            time_samples = self.get_time_samples_for_flight(flight, time_start, time_end)
            altitude = np.array([flight.altitude(time) for time in time_samples], dtype=float)

            if hasattr(flight, "vz"):
                vertical_velocity = np.array([flight.vz(time) for time in time_samples], dtype=float)
                name_vertical = "Vertical velocity [m/s]"
                hover_vertical = "Vertical velocity: %{y:.1f} m/s<extra></extra>"
            else:
                vertical_velocity = np.array([flight.speed(time) for time in time_samples], dtype=float)
                name_vertical = "Speed [m/s]"
                hover_vertical = "Speed: %{y:.1f} m/s<extra></extra>"

            # Horizontal velocity = sqrt(vx² + vy²): magnitude of the horizontal velocity vector.
            # OpenRocket uses the same definition and calls it "lateral velocity".
            if hasattr(flight, "vx") and hasattr(flight, "vy"):
                horizontal_velocity = np.array(
                    [math.sqrt(flight.vx(t) ** 2 + flight.vy(t) ** 2) for t in time_samples], dtype=float
                )
            else:
                horizontal_velocity = np.full(len(time_samples), np.nan)

            if hasattr(flight, "az"):
                acceleration = np.array([flight.az(time) for time in time_samples], dtype=float)
                name_acceleration = "Vertical acceleration [m/s^2]"
                name_acceleration_hover = "Vertical acceleration: %{y:.1f} m/s^2<extra></extra>"
            else:
                acceleration = np.array([flight.acceleration(time) for time in time_samples], dtype=float)
                name_acceleration = "Total acceleration [m/s^2]"
                name_acceleration_hover = "Total acceleration: %{y:.1f} m/s^2<extra></extra>"

            traces = [
                {
                    "y": altitude,
                    "name": "Altitude [m]",
                    "hovertemplate": "Altitude: %{y:.1f} m<extra></extra>",
                    "line": {"color": "royalblue"},
                },
                {
                    "y": vertical_velocity,
                    "name": name_vertical,
                    "hovertemplate": hover_vertical,
                    "line": {"color": "firebrick"},
                },
                {
                    "y": horizontal_velocity,
                    "name": "Horizontal velocity [m/s]",
                    "hovertemplate": "Horizontal velocity: %{y:.1f} m/s<extra></extra>",
                    "line": {"color": "mediumseagreen"},
                },
                {
                    "y": acceleration,
                    "name": name_acceleration,
                    "hovertemplate": name_acceleration_hover,
                    "line": {"color": "gold"},
                    "yaxis": "y2",
                },
            ]

            if custom_event_markers is None:
                flight_event_markers = self.get_standard_event_markers_for_flight(flight, time_samples, motor=motor)
            else:
                flight_event_markers = custom_event_markers[flight_index]

            flight_groups.append(
                {
                    "label": plot_title,
                    "time_samples": time_samples,
                    "time_start": time_start,
                    "time_end": time_end,
                    "traces": traces,
                    "event_markers": flight_event_markers,
                }
            )

        self.create_grouped_plotly_plot(
            title="Motion over time",
            flight_groups=flight_groups,
            yaxis_title="Altitude [m] / Velocity [m/s]",
            yaxis2_title="Acceleration [m/s^2]",
            width=1500,
            height=800,
            legend_x_pos=1.05,
            buttons_x_pos=1,
        )

    @staticmethod
    def plot_parachute_model(parachute: Parachute, save_dir: Path | None = None):
        """
        Plot the inflated parachute geometry as modeled by RocketPy (hemispheroid/semi-ellipsoid without hole).
        
        https://en.wikipedia.org/wiki/Ellipsoid 
        
        - with a=b=radius
        - c=height
        
        Parameterization semi-ellipsoid surface:
        - x = a * sinθ * cosφ
        - y = a * sinθ * sinφ
        - z = c * cosθ
        
        with 
        - 0 ≤ θ ≤ π/2​ to only get the upper half
        - 0 ≤ φ ≤ 2π
        """

        radius = parachute.radius
        height = parachute.height

        theta_values = np.linspace(0, np.pi / 2, 50)
        phi_values = np.linspace(0, 2 * np.pi, 100)

        theta_grid, phi_grid = np.meshgrid(theta_values, phi_values)

        x_values = radius * np.sin(theta_grid) * np.cos(phi_grid)
        y_values = radius * np.sin(theta_grid) * np.sin(phi_grid)
        z_values = height * np.cos(theta_grid)
        
        # --- plot ---
        number_of_color_panels = 12
        
        # Convert phi angle to panel index.
        # This creates alternating red/white slices around the parachute.
        panel_index = np.floor(number_of_color_panels * phi_grid / (2 * np.pi))

        surface_color_values = panel_index % 2

        figure = go.Figure()

        figure.add_trace(
            go.Surface(
                x=x_values,
                y=y_values,
                z=z_values,
                surfacecolor=surface_color_values,
                colorscale=[
                    [0.0, "white"],
                    [0.499, "white"],
                    [0.5, "red"],
                    [1.0, "red"],
                ],
                cmin=0,
                cmax=1,
                showscale=False,
                name=f"{parachute.name} canopy",
                opacity=1,
                hovertemplate=(
                    "x: %{x:.3f} m<br>"
                    "y: %{y:.3f} m<br>"
                    "z: %{z:.3f} m"
                    "<extra></extra>"
                ),
            )
        )

        figure.update_layout(
            title={
                "text": (
                    f"{parachute.name.title()} parachute model<br>"
                    f"radius = {radius:.3f} m, height = {height:.3f} m, "
                    f"cd_s = {parachute.cd_s:.3f} m²"
                ),
                "x": 0.5,
                "xanchor": "center",
            },
            width=550,
            height=550,
            scene={
                "xaxis": {
                    "title": "x [m]",
                },
                "yaxis": {
                    "title": "y [m]",
                },
                "zaxis": {
                    "title": "z [m]",
                },
                "aspectmode": "data",
            },
            template="plotly_white",
        )

        if save_dir is not None:
            filename = re.sub(r"[^\w\-]", "_", parachute.name)
            figure.write_html(str(save_dir / f"parachute_model_{filename}.html"))

        figure.show(renderer="notebook")


def plot_wind_speed_and_heading(environment: Environment, max_expected_height_asl, save_dir: Path | None = None):
    """
    Plot wind speed and heading.

    Wind heading describes the physical wind velocity vector and points where the wind is blowing:
        0°   = wind blows north, comes from south
        90°  = wind blows east, comes from west
        
    Its components are:
        - wind_u: eastward component of the physical wind velocity vector
        - wind_v: northward component of the physical wind velocity vector

    Wind_speed: magnitude of the wind velocity vector.
    
    max_expected_height_asl needs to be passed to the function instead of using Environment.max_expected_height,
    since RocketPy ignores it for some environments and uses the default 80.000 km.
    
    How Windy is Too Windy for a Launch? https://www.apogeerockets.com/Peak-of-Flight/Newsletter498
    """
    start_height = 0.0
    end_height = max_expected_height_asl

    height_samples = np.linspace(start_height, end_height, 100)

    wind_speed = np.array([environment.wind_speed(z) for z in height_samples], dtype=float)
    wind_u = np.array([environment.wind_velocity_x(z) for z in height_samples], dtype=float)
    wind_v = np.array([environment.wind_velocity_y(z) for z in height_samples], dtype=float)

    fig, axes = plt.subplots(
        nrows=1,
        ncols=2,
        figsize=(6, 4.5),
        sharey=True,
        constrained_layout=True,
    )

    speed_axis = axes[0]
    heading_axis = axes[1]

    # -------------------------------------------------------------------------
    # Left plot: wind speed
    # -------------------------------------------------------------------------
    speed_axis.plot(wind_speed, height_samples)

    speed_axis.set_xlabel("Wind speed [m/s]")
    speed_axis.set_ylabel("Height Above Sea Level [m]")
    speed_axis.grid(True)

    # -------------------------------------------------------------------------
    # Right plot: wind heading as arrows
    # -------------------------------------------------------------------------
    arrow_x_positions = np.zeros_like(height_samples)

    # Plot fewer arrows to avoid visual clutter.
    arrow_step = max(1, len(height_samples) // 25)

    heading_axis.quiver(
        arrow_x_positions[::arrow_step],
        height_samples[::arrow_step],
        wind_u[::arrow_step],
        wind_v[::arrow_step],
        angles="uv",
        scale_units="width",
        scale=20,           # length of arrows (smaller values = longer arrows)
        width=0.006,
    )

    heading_axis.set_xlabel("Wind heading")
    heading_axis.grid(True)
    heading_axis.set_xlim(-1.0, 1.0)

    # Compass direction labels.
    heading_axis.text(
        0.5,
        0.99,
        "N",
        transform=heading_axis.transAxes,
        ha="center",
        va="top",
    )

    heading_axis.text(
        0.98,
        0.5,
        "E",
        transform=heading_axis.transAxes,
        ha="right",
        va="center",
    )
    heading_axis.text(
        0.5,
        0.01,
        "S",
        transform=heading_axis.transAxes,
        ha="center",
        va="bottom",
    )

    heading_axis.text(
        0.02,
        0.5,
        "W",
        transform=heading_axis.transAxes,
        ha="left",
        va="center",
    )

    # -------------------------------------------------------------------------
    # Launch site elevation reference line
    # -------------------------------------------------------------------------
    speed_axis.axhline(
        environment.elevation,
        linestyle="--",
        linewidth=1.5,
    )

    speed_axis.annotate(
        f"Elevation Launch Site: {environment.elevation:.1f} m",
        xy=(0.98, environment.elevation),
        xycoords=speed_axis.get_yaxis_transform(),
        xytext=(0, -4),
        textcoords="offset points",
        va="top",
        ha="right",
    )

    if save_dir is not None:
        env_name = re.sub(r"[^\w\-]", "_", getattr(environment, "name", "environment"))
        fig.savefig(str(save_dir / f"{env_name}_wind_speed_and_heading.png"), dpi=150, bbox_inches="tight")

    plt.show()
