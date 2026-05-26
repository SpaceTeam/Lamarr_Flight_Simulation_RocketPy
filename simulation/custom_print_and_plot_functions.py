import numpy as np
from colorama import Fore, Style                                  # https://github.com/tartley/colorama
import math
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
    CustomPlots.flight_forecasts : Flight
        Flight forecast to plot.
    CustomPlots.motor : SolidMotor | LiquidMotor | HybridMotor
        Used motor for this flight forecast.
    CustomPlots.plot_title: str
        Used configuration for this flight forecast.
    CustomPlots.rocket : Rocket
        Used rocket for this flight forecast.
    CustomPlots.rocket_config : dict
        Used rocket config dict for this flight forecast.
    """
    def __init__(
        self,
        flight_forecast: Flight,
        motor: SolidMotor | LiquidMotor | HybridMotor,
        plot_title: str,
        rocket: Rocket,
        rocket_config: dict,
    ):
        self.flight_forecast = flight_forecast
        self.motor = motor
        self.plot_title = plot_title
        self.rocket = rocket
        self.rocket_config = rocket_config


    def get_time_samples(self, time_start: float, time_end: float) -> np.ndarray:
        """
        Get time samples between the given start and end time.
        """
        time_samples = np.asarray(self.flight_forecast.time, dtype=float)
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
    

    def add_one_vertical_event_line(
        self,
        figure: go.Figure,
        event_time: float | None,
        label: str,
        color: str,
        annotation_side: str = "right",
    ):
        """
        Add a vertical event line with a label to a Plotly figure.
        The annotation can be placed either on the left or on the right side of the line.
        """
        if event_time is None:
            return

        figure.add_shape(
            type="line",
            x0=event_time,
            x1=event_time,
            y0=0,
            y1=1,
            xref="x",
            yref="paper",
            line={
                "color": color,
                "width": 1.5,
                "dash": "dash",
            },
        )

        if annotation_side == "left":
            xshift = -5
            xanchor = "right"
        else:
            xshift = 4
            xanchor = "left"

        figure.add_annotation(
            x=event_time,
            y=0.5,
            xref="x",
            yref="paper",
            text=label,
            showarrow=False,
            textangle=-90,
            font={"color": color},
            xanchor=xanchor,
            yanchor="middle",
            align="center",
            xshift=xshift,
        )


    def add_event_marker_lines(
        self,
        figure: go.Figure,
        time_samples: np.ndarray,
        event_markers: list[tuple[float, str, str]],
    ):
        """
        Filter, sort, and draw vertical event marker lines on a Plotly figure.
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
        close_event_threshold = 1.5
        close_event_side = "left"

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

            self.add_one_vertical_event_line(
                figure=figure,
                event_time=event_time,
                label=label,
                color=color,
                annotation_side=annotation_side,
            )

            previous_event_time = event_time


    def add_vertical_event_markers(self, time_samples: np.ndarray, figure: go.Figure):
        """
        Add standard flight event markers to a Plotly figure.
        """
        if hasattr(self.flight_forecast, "out_of_rail_time"):
            out_of_rail_time = float(self.flight_forecast.out_of_rail_time)
        else:
            out_of_rail_time = None
            
        burn_out_time = float(self.motor.burn_out_time)
        
        if hasattr(self.flight_forecast, "apogee_time"):
            apogee_time = float(self.flight_forecast.apogee_time)
        else:
            apogee_time = None
        
        if hasattr(self.flight_forecast, "t_final"):
            ground_hit_time = float(self.flight_forecast.t_final)
        else: 
            ground_hit_time = None
        
        if hasattr(self.flight_forecast, "mach_number"):
            mach_number = np.array([self.flight_forecast.mach_number(time) for time in time_samples], dtype=float)
        else:
            mach_number = None
        

        # --- Transonic and supersonic intervals ---
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

        if hasattr(self.flight_forecast, "parachute_events"):
            for ejection_time, parachute in self.flight_forecast.parachute_events:
                inflation_time = ejection_time + parachute.lag
                event_markers.append((float(ejection_time),f"{parachute.name} parachute ejected; cd_s {parachute.cd_s:.2f} m²", "black"))
                event_markers.append((float(inflation_time),f"{parachute.name} parachute inflated", "black"))
                
        self.add_event_marker_lines(figure, time_samples, event_markers)


    def create_plotly_plot(
        self,
        title: str,
        time_samples: np.ndarray,
        time_start: float,
        time_end: float,
        traces: list[dict],
        yaxis_title: str,
        yaxis2_title: str | None = None,
        yaxis_dtick: float | int | None = None,
        yaxis2_dtick: float | int | None = None,
        width: int = 800,
        height: int = 500,
        event_markers=None,
    ):
        """
        Create and display a Plotly plot with a shared x-axis, optional second y-axis,
        unified hover display, and standard event markers.

        Each trace is passed as a dictionary containing the plotted data and formatting.
        """

        figure = go.Figure()

        for trace in traces:
            scatter_options = {
                "x": time_samples,
                "y": trace["y"],
                "mode": "lines",
                "name": trace["name"],
                "hovertemplate": trace["hovertemplate"],
                "line": trace["line"],
            }

            # optional right y axis
            if trace.get("yaxis") == "y2":
                scatter_options["yaxis"] = "y2"

            figure.add_trace(go.Scatter(**scatter_options))

        if event_markers is None:
            self.add_vertical_event_markers(time_samples, figure)
        else:
            self.add_event_marker_lines(figure, time_samples, event_markers)

        yaxis_settings = {
            "title": yaxis_title,
            "showgrid": True,
        }

        if yaxis_dtick is not None:
            yaxis_settings["dtick"] = yaxis_dtick

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
                "range": [time_start, time_end + 1],
                # "dtick": 4,
                "showgrid": True,
                "hoverformat": ".3f",
                "unifiedhovertitle": {
                    "text": "Time: %{x:.3f} s",
                },
            },
            "yaxis": yaxis_settings,
            "hovermode": "x unified",
            "legend": {
                "orientation": "h",
                "yanchor": "bottom",
                "y": 1.02,
                "xanchor": "center",
                "x": 0.5,
            },
            "template": "plotly_white",
        }

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

        figure.update_layout(**layout_settings)

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
        time_start = 0.0
        time_end = float(self.flight_forecast.apogee_time)
        time_samples = self.get_time_samples(time_start=time_start, time_end=time_end)
        
        # --- Flight data ---
        # CG(t): center of mass position along rocket axis (in rocket reference system)
        cg_position = np.array([self.rocket.center_of_mass(time) for time in time_samples], dtype=float) * 1000     # m to mm
        
        # Mach(t) from the flight, then CP(Mach(t))
        mach_number = np.array([self.flight_forecast.mach_number(time) for time in time_samples], dtype=float)
        cp_position = np.array([self.rocket.cp_position(mach) for mach in mach_number], dtype=float) * 1000     # m to mm
        
        # stability(t)
        stability_margin = np.array([self.flight_forecast.stability_margin(time) for time in time_samples], dtype=float)

        # convert
        if use_openrocket_coordinates:
            rocket_length = self.rocket_config["total_length"]          # mm
            cg_position = rocket_length - cg_position
            cp_position = rocket_length - cp_position

        # --- Plot ---
        if use_openrocket_coordinates:
            position_axis_title = "Position along rocket axis [mm] (0 = nose)"
            title_suffix = "OpenRocket CG/CP orientation"
        else:
            position_axis_title = "Position along rocket axis [mm] (0 = tail)"
            title_suffix = "RocketPy CG/CP orientation"
            
        traces = [
            # left y axis
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
            # right y axis
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

        self.create_plotly_plot(
            title=f"[{self.plot_title}] CG/CP position, Mach, Stability ({title_suffix})",
            time_samples=time_samples,
            time_start=time_start,
            time_end=time_end,
            traces=traces,
            yaxis_title="Mach / Stability margin [c]",
            yaxis2_title=position_axis_title,
            yaxis_dtick=0.5,
            yaxis2_dtick=100,
            width=900,
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
        time_start = 0.0
        
        inflation_times = []
        for ejection_time, parachute in self.flight_forecast.parachute_events:
            inflation_times.append(ejection_time + parachute.lag)
        first_inflation_time = min(inflation_times) if inflation_times else None
        
        time_end = first_inflation_time if first_inflation_time is not None else float(self.flight_forecast.apogee_time)
        time_samples = self.get_time_samples(time_start=time_start, time_end=time_end)
        
        angle_of_attack = np.array([self.flight_forecast.angle_of_attack(time) for time in time_samples], dtype=float)
        attitude_angle = np.array([self.flight_forecast.attitude_angle(time) for time in time_samples], dtype=float)

        # Heading (compass bearing of the velocity vector):
        vx = np.array([self.flight_forecast.vx(time) for time in time_samples], dtype=float)
        vy = np.array([self.flight_forecast.vy(time) for time in time_samples], dtype=float)
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

        self.create_plotly_plot(
            title=f"[{self.plot_title}] Angle of attack, attitude angle, and heading over time",
            time_samples=time_samples,
            time_start=time_start,
            time_end=time_end,
            traces=traces,
            yaxis_title="Angle [°]",
            yaxis2_title="Heading [°]",
            width=900,
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
        time_start = 0.0

        inflation_times = []
        for ejection_time, parachute in self.flight_forecast.parachute_events:
            inflation_times.append(ejection_time + parachute.lag)
        first_inflation_time = min(inflation_times) if inflation_times else None

        time_end = first_inflation_time if first_inflation_time is not None else float(self.flight_forecast.apogee_time)
        time_samples = self.get_time_samples(time_start=time_start, time_end=time_end)

        if transform_openrocket:
            # Transform body-frame angular velocity to the flow-aligned reference frame.
            # pitch_rate = flow-aligned reference frame y-component = -sin(theta)*w1 + cos(theta)*w2
            # yaw_rate   = flow-aligned reference frame x-component =  cos(theta)*w1 + sin(theta)*w2
            # Matches OpenRocket AbstractSimulationStepper: pitchRate=rot.getY(), yawRate=rot.getX().

            # Angular velocity (in body frame)
            w1 = np.array([self.flight_forecast.w1(t) for t in time_samples])
            w2 = np.array([self.flight_forecast.w2(t) for t in time_samples])

            # Rocket orientation quaternions
            quaternions = np.column_stack([
                [self.flight_forecast.e1(t) for t in time_samples],
                [self.flight_forecast.e2(t) for t in time_samples],
                [self.flight_forecast.e3(t) for t in time_samples],
                [self.flight_forecast.e0(t) for t in time_samples],
            ])

            # Airspeed in inertial frame: rocket velocity - wind (matches OpenRocket convention)
            vx = np.array([self.flight_forecast.vx(t) for t in time_samples])
            vy = np.array([self.flight_forecast.vy(t) for t in time_samples])
            vz = np.array([self.flight_forecast.vz(t) for t in time_samples])
            z_alts = np.array([self.flight_forecast.z(t) for t in time_samples])
            wind_vx = np.array([self.flight_forecast.env.wind_velocity_x(z) for z in z_alts])
            wind_vy = np.array([self.flight_forecast.env.wind_velocity_y(z) for z in z_alts])
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
            frame_label = "OpenRocket flow-aligned reference frame"
        else:
            angular_rate_1 = np.degrees(np.array([self.flight_forecast.w1(time) for time in time_samples], dtype=float))
            angular_rate_2 = np.degrees(np.array([self.flight_forecast.w2(time) for time in time_samples], dtype=float))
            frame_label = "RocketPy body reference frame"

        angular_rate_3 = np.degrees(np.array([self.flight_forecast.w3(time) for time in time_samples], dtype=float))
        
        traces = [
            {
                "y": angular_rate_1,
                "name": f"pitch rate [°/s]",
                "hovertemplate": "pitch rate: %{y:.3f} °/s<extra></extra>",
                "line": {"color": "royalblue"},
            },
            {
                "y": angular_rate_2,
                "name": f"yaw rate [°/s]",
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

        self.create_plotly_plot(
            title=f"[{self.plot_title}] Angular velocity ({frame_label})",
            time_samples=time_samples,
            time_start=time_start,
            time_end=time_end,
            traces=traces,
            yaxis_title="Angular rate [°/s]",
            width=900,
            height=500,
        )
        

        
    def plot_vertical_motion(self, time_interval=None, event_markers=None):
        """
        Plot altitude, vertical velocity, and vertical acceleration over time.
        """
        if time_interval:
            time_start = time_interval[0]
            time_end = time_interval[1]
        else:
            time_start = 0.0
            time_end = float(self.flight_forecast.t_final)
        time_samples = self.get_time_samples(time_start=0.0, time_end=time_end)
        
        # --- Flight data ---
        # altitude(t)
        altitude = np.array([self.flight_forecast.altitude(time) for time in time_samples], dtype=float)
        
        # vertical_velocity(t)
        if hasattr(self.flight_forecast, "vz"):
            motion = np.array([self.flight_forecast.vz(time) for time in time_samples], dtype=float)
            name_motion = "Vertical velocity [m/s]"
            name_motion_hover = "Vertical velocity: %{y:.1f} m/s<extra></extra>"
            title_yaxis = "Altitude [m] / Vertical velocity [m/s]"
        else:
            motion = np.array([self.flight_forecast.speed(time) for time in time_samples], dtype=float)
            name_motion = "Speed [m/s]"
            name_motion_hover = "Speed: %{y:.1f} m/s<extra></extra>"
            title_yaxis = "Altitude [m] / Speed [m/s]"
            
        # vertical_acceleration(t)
        if hasattr(self.flight_forecast, "az"):
            acceleration = np.array([self.flight_forecast.az(time) for time in time_samples], dtype=float)
            name_acceleration = "Vertical acceleration [m/s²]"
            name_acceleration_hover = "Vertical acceleration: %{y:.1f} m/s²<extra></extra>"
            title_yaxis2 = "Vertical acceleration [m/s²]"
        else:
            acceleration = np.array([self.flight_forecast.acceleration(time) for time in time_samples], dtype=float)
            name_acceleration = "Total acceleration [m/s²]"
            name_acceleration_hover = "Total acceleration: %{y:.1f} m/s²<extra></extra>"
            title_yaxis2 = "Total acceleration [m/s²]"

        # --- Plot ---
        traces = [
            # left y axis
            {
                "y": altitude,
                "name": "Altitude [m]",
                "hovertemplate": "Altitude: %{y:.1f} m<extra></extra>",
                "line": {"color": "royalblue"},
            },
            {
                "y": motion,
                "name": name_motion,
                "hovertemplate": name_motion_hover,
                "line": {"color": "firebrick"},
            },
            # right y axis
            {
                "y": acceleration,
                "name": name_acceleration,
                "hovertemplate": name_acceleration_hover,
                "line": {"color": "gold"},
                "yaxis": "y2",
            },
        ]

        self.create_plotly_plot(
            title=f"[{self.plot_title}] Vertical motion over time",
            time_samples=time_samples,
            time_start=time_start,
            time_end=time_end,
            traces=traces,
            yaxis_title=title_yaxis,
            yaxis2_title=title_yaxis2,
            width=1300,
            height=800,
            event_markers=event_markers,
        )
        
    @staticmethod
    def plot_parachute_model(parachute: Parachute):
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

        figure.show(renderer="notebook")


def plot_wind_speed_and_heading(environment: Environment, max_expected_height_asl):
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

    plt.show()