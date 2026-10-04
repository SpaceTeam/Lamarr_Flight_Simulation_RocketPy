"""
RocketPy simulation backend.
"""

import copy
import datetime
import warnings
from math import pi
from pathlib import Path

import CoolProp.CoolProp as CP
import numpy as np
import pandas as pd
from IPython.display import display
from openmeteo_requests import OpenMeteoRequestsError

from rocketpy import (
    Environment,
    SolidMotor,
    Rocket,
    Flight,
    Fluid,
    LiquidMotor,
    CylindricalTank,
    MassFlowRateBasedTank,
    TrapezoidalFins,
    FreeFormFins,
    NoseCone,
    Tail,
    Parachute,
    Function,
)

from simulation_core.utils import *
from simulation_core.config_schema import SimParams, ParachuteConfig
from simulation_core.export_weather_data import export_weather_data
from simulation_core.outputs import *
from simulation_core.custom_print_and_plot_functions import CustomPlots, plot_wind_speed_and_heading
from simulation_core import deployable_payload

DEBUG = False
# Standard Environment with wind
WIND_PROFILE_STEP_M = 200       # height spacing of wind levels
WIND_PROFILE_SEED = 42          # fixed seed so every run gets the same wind profiles and results stay comparable
WIND_TURBULENCE_INTENSITY = 0.1     # standard deviation of varied wind speed, as a fraction of the wind speed
# Display metadata key of the flight counter; streamlit_app/run_simulation.py reads it to show a live flight progress bar.
FLIGHT_PROGRESS_METADATA_KEY = "flight_progress"


# =============================================================================
# Small conversion helpers
# =============================================================================

def millimeters_to_meters(value):
    """Convert a value from millimeters to meters."""
    return value / 1000


def grams_to_kilograms(value):
    """Convert a value from grams to kilograms."""
    return value / 1000


def get_position_from_tip(value_from_tip_mm, rocket_length_m):
    """Convert a nose-tip-origin position to RocketPy's bottom-origin coordinate system."""
    return rocket_length_m - millimeters_to_meters(value_from_tip_mm)


# =============================================================================
# Environment
# =============================================================================

def create_environment(params: SimParams):
    """Create all configured RocketPy environments and store them in params.runtime.environments."""
    print("Creating environment...")

    env_config = params.config.environment
    # REMARK: defaults might create errors, rather deactivate
    # defaults = {
    #     "envType": "Forecast",
    #     "latitude": 39.232465,
    #     "longitude": -8.172108,
    #     "timezone": "Europe/Rome",
    #     "max_expected_height": 9000,
    #     "elevation": "Open-Elevation",
    #     "date": (datetime.datetime.now() + datetime.timedelta(days=1)),
    # }
    # add_defaults(parameter_env, defaults, constants, variations, "environment")
    
    environment_types = ensure_list(env_config.envType)

    reanalysis_types = {"reanalysis", "reanalysis_custom"}
    if reanalysis_types.intersection(environment_types) and len(environment_types) > 1:
        raise ValueError("environment_envType 'reanalysis' and 'reanalysis_custom' must be selected alone.")

    latitude = env_config.latitude
    longitude = env_config.longitude
    timezone = env_config.timezone
    elevation = env_config.elevation
    max_expected_height_agl = env_config.max_expected_height
    max_expected_height_asl = max_expected_height_agl + elevation
    date = env_config.date

    # date is either a naive local datetime from the TOML file or the magic string below; set_date applies the timezone.
    if date == "tomorrow_08_local":
        tomorrow = datetime.date.today() + datetime.timedelta(days=1)
        date = datetime.datetime(tomorrow.year, tomorrow.month, tomorrow.day, 8, 0, 0)

    environments = {}

    if "standard_atmosphere" in environment_types:
        standard_env = Environment(max_expected_height=max_expected_height_asl)
        standard_env.set_location(latitude=latitude, longitude=longitude)
        standard_env.set_elevation(elevation)
        standard_env.set_date(date, timezone=timezone)
        standard_env.set_atmospheric_model(type="standard_atmosphere")
        standard_env.name = "Standardized"
        set_labels(standard_env)
        environments[standard_env.name] = standard_env

    if "Windy" in environment_types:
        if not env_config.windy_weather_models:
            raise ValueError("environment.windy_weather_models is required for Windy envType.")
        for weather_model in ensure_list(env_config.windy_weather_models):
            forecast_env = Environment(max_expected_height=max_expected_height_asl)
            forecast_env.set_location(latitude=latitude, longitude=longitude)
            forecast_env.set_elevation(elevation)
            forecast_env.set_date(date, timezone=timezone)
            forecast_env.set_atmospheric_model(type="Windy", file=weather_model)
            forecast_env.name = f"{weather_model}_WINDY"
            set_labels(forecast_env)
            environments[forecast_env.name] = forecast_env

    if "standard_atmosphere_wind" in environment_types:
        if not env_config.standard_atmosphere_wind_speeds:
            raise ValueError("environment.standard_atmosphere_wind_speeds is required for standard_atmosphere_wind envType.")
        for wind_speed in ensure_list(env_config.standard_atmosphere_wind_speeds):
            standard_env_wind = Environment(max_expected_height=max_expected_height_asl)
            standard_env_wind.set_location(latitude=latitude, longitude=longitude)
            standard_env_wind.set_elevation(elevation)
            standard_env_wind.set_date(date, timezone=timezone)

            # Seeding with the speed gives each speed its own fixed profile, independent of the other list entries.
            rng = np.random.default_rng([WIND_PROFILE_SEED, round(wind_speed * 1000)])
            # One random speed per height level from the ground to the top of the model
            heights_asl = np.arange(elevation, max_expected_height_asl + WIND_PROFILE_STEP_M, WIND_PROFILE_STEP_M)
            level_speeds = rng.normal(wind_speed, wind_speed * WIND_TURBULENCE_INTENSITY, size=len(heights_asl))
            # Match the standard OpenRocket env: wind comes from 90° (East), so it blows West and wind_u is negative.
            wind_u_profile = np.column_stack([heights_asl, -level_speeds])
            standard_env_wind.set_atmospheric_model(type="custom_atmosphere", wind_u=wind_u_profile, wind_v=0)
            standard_env_wind.name = f"{wind_speed:g}_m_s_STANDARD_WIND"
            set_labels(standard_env_wind)
            environments[standard_env_wind.name] = standard_env_wind

    if "custom_atmosphere" in environment_types:
        if not env_config.custom_weather_models:
            raise ValueError("environment.custom_weather_models is required for custom_atmosphere envType.")
        output_folder = params.project_path / "weather_csvs"
        launch_time = date.strftime("%Y-%m-%dT%H:%M:%S")

        for weather_model in ensure_list(env_config.custom_weather_models):
            print(f"\nCreating weather CSV files for {weather_model}...")
            try:
                export_weather_data(
                    latitude=latitude,
                    longitude=longitude,
                    max_expected_height_agl_m=max_expected_height_agl,
                    launch_time=launch_time,
                    timezone=timezone,
                    weather_model=weather_model,
                    output_folder=output_folder,
                )
            except OpenMeteoRequestsError as error:
                # A rejected request (e.g. a date outside the forecast range) carries Open-Meteo's JSON reply with a "reason".
                reply = error.__cause__.args[0] if error.__cause__ is not None and error.__cause__.args else None
                reason = reply["reason"] if isinstance(reply, dict) and "reason" in reply else str(error)
                # Stop with only that reason; "from None" hides the long traceback through the openmeteo_requests library.
                raise OpenMeteoRequestsError(f"Open-Meteo ({weather_model}): {reason}") from None

            # Load the .csv file into the environment 
            # https://docs.rocketpy.org/en/latest/user/environment/3-further/data_csv.html#load-the-csv-file
            csv_path = output_folder / f"rocketpy_atmosphere_{weather_model}.csv"
            if not csv_path.exists():
                raise FileNotFoundError(f"Custom environment CSV was not created: {csv_path}")

            dataframe = pd.read_csv(csv_path)
            
            # Create Function objects to represent the profiles
            pressure_function = Function(np.column_stack([dataframe["height"], dataframe["pressure"]]))
            temperature_function = Function(np.column_stack([dataframe["height"], dataframe["temperature"]]))
            wind_u_function = Function(np.column_stack([dataframe["height"], dataframe["wind_u"]]))
            wind_v_function = Function(np.column_stack([dataframe["height"], dataframe["wind_v"]]))

            custom_env = Environment(max_expected_height=max_expected_height_asl)
            custom_env.set_location(latitude=latitude, longitude=longitude)
            custom_env.set_elevation(elevation)
            custom_env.set_date(date, timezone=timezone)
            custom_env.set_atmospheric_model(
                type="custom_atmosphere",
                pressure=pressure_function,
                temperature=temperature_function,
                wind_u=wind_u_function,
                wind_v=wind_v_function,
            )
            custom_env.name = f"Custom_{weather_model}"
            set_labels(custom_env)
            environments[custom_env.name] = custom_env

    if "reanalysis" in environment_types:
        if not env_config.reanalysis_file or not env_config.reanalysis_dictionary:
            raise ValueError("environment.reanalysis_file and reanalysis_dictionary are required for reanalysis envType.")
        reanalysis_env = Environment(max_expected_height=max_expected_height_asl)
        reanalysis_env.set_location(latitude=latitude, longitude=longitude)
        reanalysis_env.set_elevation(elevation)
        reanalysis_env.set_date(date, timezone=timezone)
        reanalysis_env.set_atmospheric_model(
            type="Reanalysis",
            file=get_project_file(params, env_config.reanalysis_file),
            dictionary=env_config.reanalysis_dictionary,
        )
        params.runtime.reanalysis_mode = True
        reanalysis_env.name = f"Reanalysis_{Path(env_config.reanalysis_file).stem}"
        set_labels(reanalysis_env)
        environments[reanalysis_env.name] = reanalysis_env

    if "reanalysis_custom" in environment_types:
        if not env_config.reanalysis_csv:
            raise KeyError("Missing required config key for reanalysis environment: reanalysis_csv")

        for model_name, csv_file in env_config.reanalysis_csv.items():
            csv_path = Path(csv_file)
            if not csv_path.exists():
                csv_path = params.project_path / csv_file
            if not csv_path.exists():
                csv_path = params.project_path / "weather_csvs" / csv_file
            if not csv_path.exists():
                raise FileNotFoundError(f"Reanalysis CSV file not found: {csv_file}")

            dataframe = pd.read_csv(csv_path)
            pressure_function = Function(np.column_stack([dataframe["height"], dataframe["pressure"]]))
            temperature_function = Function(np.column_stack([dataframe["height"], dataframe["temperature"]]))
            wind_u_function = Function(np.column_stack([dataframe["height"], dataframe["wind_u"]]))
            wind_v_function = Function(np.column_stack([dataframe["height"], dataframe["wind_v"]]))

            reanalysis_env = Environment(max_expected_height=max_expected_height_asl)
            reanalysis_env.set_location(latitude=latitude, longitude=longitude)
            reanalysis_env.set_elevation(elevation)
            reanalysis_env.set_date(date, timezone=timezone)
            reanalysis_env.set_atmospheric_model(
                type="custom_atmosphere",
                pressure=pressure_function,
                temperature=temperature_function,
                wind_u=wind_u_function,
                wind_v=wind_v_function,
            )
            params.runtime.reanalysis_mode = True
            reanalysis_env.name = f"Reanalysis_Custom_{model_name}"
            set_labels(reanalysis_env)
            environments[reanalysis_env.name] = reanalysis_env

    if params.config.output_level > 1:
        for name, env in environments.items():
            print_one_environment(env, name)
            plot_wind_speed_and_heading(env, max_expected_height_asl, save_dir=params.project_path / "plots")
        if params.config.output_level > 2:
            env.plots.atmospheric_model()

    params.runtime.environments = environments


# =============================================================================
# Engine
# =============================================================================

def create_engine(params: SimParams):
    """Create the configured motor type and store it in params.runtime.motor."""
    if has_rocket_component_variations(params.config):
        printmd("## Engine")
        print("Rocket components vary. Engine will be built per combination in create_flight.")
        return
    if params.config.output_level >= 0:
        printmd("## Engine")
    # engine.type is validated to "solid" or "liquid" when the config is loaded.
    if params.config.engine.type == "liquid":
        create_liquid_engine(params)
    else:
        create_solid_engine(params)


def create_liquid_engine(params: SimParams):
    """Create a LiquidMotor from params.config.liquid_engine and store it in params.runtime.motor."""
    if params.config.output_level >= 0:
        print("Creating liquid engine...")

    liquid_engine_config = params.config.liquid_engine
    if liquid_engine_config is None:
        raise ValueError("engine.type is 'liquid' but no 'liquid_engine' section found in config.")

    press_tank_cfg = liquid_engine_config.pressure_gas_tank
    fuel_tank_cfg  = liquid_engine_config.fuel_tank
    ox_tank_cfg    = liquid_engine_config.oxidizer_tank

    # defaults = {
    #     "temperature_nitrogen": 298.15,                # K       # preset (= 25°C)
    #     "temperature_ethanol": 298.15,                # K       # preset (= 25°C)
    #     "temperature_lox": 93.15,                    # K       # preset (= -180°C)
    #     "pressure_nitrogen": 30000000,              # Pa     # preset
    #     "pressure_ethanol": 3000000,              # Pa     # preset
    #     "pressure_lox": 3000000,             # Pa     # preset
    # }
    # add_defaults(parameter_motor, defaults, constants, variations)
    
    # Fluid densities derived from temperature and pressure via CoolProp
    # Propellants

    # define density
    rho_pressure_gas = CP.PropsSI("D", "T", press_tank_cfg.temperature, "P|gas", press_tank_cfg.pressure, press_tank_cfg.substance)
    rho_fuel         = CP.PropsSI("D", "T|liquid", fuel_tank_cfg.temperature, "P", fuel_tank_cfg.pressure, fuel_tank_cfg.substance)
    rho_oxidizer     = CP.PropsSI("D", "T|liquid", ox_tank_cfg.temperature, "P", ox_tank_cfg.pressure, ox_tank_cfg.substance)

    # define fluids
    pressure_gas_fluid = Fluid(name=press_tank_cfg.substance, density=rho_pressure_gas)
    fuel_fluid         = Fluid(name=fuel_tank_cfg.substance,  density=rho_fuel)
    oxidizer_fluid     = Fluid(name=ox_tank_cfg.substance,    density=rho_oxidizer)

    # define tanks geometry
    press_tank_cfg_shape = CylindricalTank(
        radius=millimeters_to_meters(press_tank_cfg.outer_diameter) / 2,
        height=millimeters_to_meters(press_tank_cfg.length),
        spherical_caps=True,
    )
    fuel_tank_shape = CylindricalTank(
        radius=millimeters_to_meters(fuel_tank_cfg.outer_diameter) / 2,
        height=millimeters_to_meters(fuel_tank_cfg.length),
        spherical_caps=True,
    )
    ox_tank_shape = CylindricalTank(
        radius=millimeters_to_meters(ox_tank_cfg.outer_diameter) / 2,
        height=millimeters_to_meters(ox_tank_cfg.length),
        spherical_caps=True,
    )

    # Burn time: shortest of the two propellant burn durations minus hold-down
    mass_fuel_init     = fuel_tank_cfg.volume / 1000 * rho_fuel                         # kg
    mass_oxidizer_init = ox_tank_cfg.volume / 1000 * rho_oxidizer                       # kg
    burn_time_fuel = mass_fuel_init / fuel_tank_cfg.massflow                            # s
    burn_time_ox = mass_oxidizer_init / ox_tank_cfg.massflow                            # s
    # account for holddown
    burn_time = min(burn_time_fuel, burn_time_ox) - liquid_engine_config.holddown_time  # s

    # Actual in-tank masses used during the burn (tiny offset avoids zero-mass issues in RocketPy)
    mass_press_gas = press_tank_cfg.massflow * burn_time + 0.0001
    mass_fuel      = fuel_tank_cfg.massflow  * burn_time + 0.0001
    mass_oxidizer  = ox_tank_cfg.massflow    * burn_time + 0.0001

    # define tanks
    press_tank = MassFlowRateBasedTank(
        name="pressure gas tank",
        geometry=press_tank_cfg_shape,
        flux_time=burn_time,                                                # s
        initial_liquid_mass=0,                                              # kg
        initial_gas_mass=mass_press_gas,                                    # kg
        liquid_mass_flow_rate_in=0,                                         # kg/s
        liquid_mass_flow_rate_out=0,                                        # kg/s
        gas_mass_flow_rate_in=0,                                            # kg/s
        gas_mass_flow_rate_out=lambda _: press_tank_cfg.massflow,           # kg/s
        liquid=Fluid(name="liquid", density=0.0001),                        # ignore
        gas=pressure_gas_fluid,
    )
    fuel_tank = MassFlowRateBasedTank(
        name="fuel tank",
        geometry=fuel_tank_shape,
        flux_time=burn_time,                                                # s
        initial_liquid_mass=mass_fuel,                                      # kg
        initial_gas_mass=0,                                                 # kg
        liquid_mass_flow_rate_in=0,                                         # kg/s
        liquid_mass_flow_rate_out=lambda _: fuel_tank_cfg.massflow,         # kg/s
        gas_mass_flow_rate_in=lambda _: press_tank_cfg.massflow,            # kg/s
        gas_mass_flow_rate_out=0,                                           # kg/s
        liquid=fuel_fluid,
        gas=pressure_gas_fluid,
    )
    ox_tank = MassFlowRateBasedTank(
        name="oxidizer tank",
        geometry=ox_tank_shape,
        flux_time=burn_time,                                                # s
        initial_liquid_mass=mass_oxidizer,                                  # kg
        initial_gas_mass=0,                                                 # kg
        liquid_mass_flow_rate_in=0,                                         # kg/s
        liquid_mass_flow_rate_out=lambda _: ox_tank_cfg.massflow,           # kg/s      
        gas_mass_flow_rate_in=lambda _: press_tank_cfg.massflow,            # kg/s
        gas_mass_flow_rate_out=0,                                           # kg/s
        liquid=oxidizer_fluid,
        gas=pressure_gas_fluid,
    )

    motor = LiquidMotor(
        dry_mass=0.0001,                                                                    # kg
        dry_inertia=(0, 0, 0),                                                              # kg*m^2
        center_of_dry_mass_position=0,                                                      # m
        nozzle_radius=millimeters_to_meters(liquid_engine_config.nozzle_diameter) / 2,      # m
        nozzle_position=0,                                                                  # m
        thrust_source=get_project_file(params, liquid_engine_config.thrust),                # N
        burn_time=burn_time,                                                                # s
        coordinate_system_orientation="nozzle_to_combustion_chamber",
    )

    rocket_length_m = millimeters_to_meters(params.config.rocket.length)
    nozzle_position_m = millimeters_to_meters(liquid_engine_config.nozzle_position)
    tanks_with_center_from_tip = [
        (press_tank, press_tank_cfg.fuel_center_from_tip),
        (press_tank, press_tank_cfg.oxidizer_center_from_tip),
        (fuel_tank, fuel_tank_cfg.center_from_tip),
        (ox_tank, ox_tank_cfg.center_from_tip),
    ]
    for tank, center_from_tip_mm in tanks_with_center_from_tip:
        # RocketPy wants distance from nozzle exit
        center_from_nozzle_m = get_position_from_tip(center_from_tip_mm, rocket_length_m) - nozzle_position_m
        motor.add_tank(tank=tank, position=center_from_nozzle_m)

    set_labels(motor)
    params.runtime.motor = motor

    if params.config.output_level > 1:
        print_one_motor(motor, type="liquid")
        plot_one_motor(motor)


def create_solid_engine(params: SimParams):
    """Create a SolidMotor from params.config.motor and store it in params.runtime.motor."""
    if params.config.output_level >= 0:
        print("Creating solid engine...")

    motor_cfg = params.config.motor

    motor_total_mass = grams_to_kilograms(motor_cfg.total_mass)
    motor_propellant_mass = grams_to_kilograms(motor_cfg.propellant_mass)
    motor_dry_mass = motor_total_mass - motor_propellant_mass
    motor_radius = millimeters_to_meters(motor_cfg.diameter) / 2
    motor_length = millimeters_to_meters(motor_cfg.length)
    motor_volume = pi * motor_radius**2 * motor_length                                          # m³
    motor_grain_density = motor_propellant_mass / motor_volume                                  # kg/m³

    # inertia of motor without propellant (dry mass) using formula for thin cylindrical shell with open ends
    dry_inertia_x_y = 1 / 12 * motor_dry_mass * (6 * motor_radius**2 + motor_length**2)         # kg*m³
    dry_inertia_z = motor_dry_mass * motor_radius**2                                            # kg/m³

    motor = SolidMotor(
        thrust_source=get_project_file(params, motor_cfg.thrust),
        dry_mass=motor_dry_mass,
        dry_inertia=(dry_inertia_x_y, dry_inertia_x_y, dry_inertia_z),                          # kg*m²
        nozzle_radius=motor_radius,
        grain_number=1,
        grain_density=motor_grain_density,
        grain_outer_radius=motor_radius,
        grain_initial_inner_radius=0,
        grain_initial_height=motor_length,
        grain_separation=0,
        grains_center_of_mass_position=motor_length / 2,
        center_of_dry_mass_position=motor_length * motor_cfg.center_of_dry_mass_factor,
        nozzle_position=0,
        burn_time=motor_cfg.burn_time,
        throat_radius=motor_radius / 2,
        coordinate_system_orientation="nozzle_to_combustion_chamber",
    )

    set_labels(motor)
    params.runtime.motor = motor

    if params.config.output_level > 1:
        print_one_motor(motor, type="solid", inertia=(dry_inertia_x_y, dry_inertia_z))
        plot_one_motor(motor)


# =============================================================================
# Rocket parts
# =============================================================================

def create_nosecone(params: SimParams):
    """Create a NoseCone from config and store it in params.runtime.nosecone."""
    if params.config.output_level >= 0:
        printmd("## NoseCone")
        print("Creating nosecone...")

    nosecone_config = params.config.nosecone
    rocket_config   = params.config.rocket

    # defaults = {"nosecone_kind": "von karman"}
    # add_defaults(parameter_nosecone, defaults, constants, variations)
    
    aerodynamic_length = nosecone_config.length - nosecone_config.cylindrical_section_length
    nosecone = NoseCone(
        length=millimeters_to_meters(aerodynamic_length),
        base_radius=millimeters_to_meters(rocket_config.diameter) / 2,
        kind=nosecone_config.kind,
    )
    set_labels(nosecone)
    params.runtime.nosecone = nosecone


def create_tailcone(params: SimParams):
    """Create a Tail from config and store it in params.runtime.tailcone."""
    if params.config.output_level >= 0:
        printmd("## TailCone")
        print("Creating tailcone...")

    tailcone_config = params.config.tailcone
    rocket_config   = params.config.rocket
    # defaults = {"tailcone_cylindrical_section_length": 0}
    # add_defaults(parameter_tailcone, defaults, constants, variations)
    
    aerodynamic_length = tailcone_config.length - tailcone_config.cylindrical_section_length
    tailcone = Tail(
        top_radius=millimeters_to_meters(rocket_config.diameter) / 2,
        bottom_radius=millimeters_to_meters(tailcone_config.bottom_radius),
        length=millimeters_to_meters(aerodynamic_length),
        rocket_radius=millimeters_to_meters(rocket_config.diameter) / 2,
    )
    set_labels(tailcone)
    params.runtime.tailcone = tailcone


def create_fins(params: SimParams):
    """Create a FreeFormFins or TrapezoidalFins from config and store in params.runtime.fin_set."""
    if params.config.output_level >= 0:
        printmd("## Fins")
        print("Creating fins...")

    # defaults = {"fins_amount": 4, "fins_name": "fins"}
    # add_defaults(parameter_fins, defaults, constants, variations)
    
    fins_cfg     = params.config.fins
    rocket_cfg   = params.config.rocket
    tailcone_cfg = params.config.tailcone

    if fins_cfg.shape_points is not None:
        fin_set = FreeFormFins(
            n=fins_cfg.amount,
            shape_points=fins_cfg.shape_points,                                                                     # m
            rocket_radius=millimeters_to_meters(tailcone_cfg.bottom_radius),                                        # m
            name=fins_cfg.name,
        )
    else:
        if any(v is None for v in (fins_cfg.root_chord, fins_cfg.tip_chord, fins_cfg.span, fins_cfg.sweep_length)):
            raise ValueError("Trapezoidal fins require root_chord, tip_chord, span, and sweep_length in config.")
        fin_set = TrapezoidalFins(
            n=fins_cfg.amount,
            root_chord=millimeters_to_meters(fins_cfg.root_chord),                                                  # m
            tip_chord=millimeters_to_meters(fins_cfg.tip_chord),                                                    # m
            span=millimeters_to_meters(fins_cfg.span),                                                              # m
            sweep_length=millimeters_to_meters(fins_cfg.sweep_length),                                              # m
            name=fins_cfg.name,
            rocket_radius=millimeters_to_meters(rocket_cfg.diameter) / 2,                                           # m
        )

    set_labels(fin_set)
    params.runtime.fin_set = fin_set

    if params.config.output_level > 1:
        print_one_fin_set(fin_set)
        fin_set.draw()


def _get_parachute_cd_s(para_config: ParachuteConfig):
    """Compute cd_s from a ParachuteConfig: directly, or from cd + radius/fabric_area."""
    if para_config.cd_s is not None:
        return para_config.cd_s
    if para_config.cd is not None and para_config.fabric_area is not None:
        return para_config.cd * para_config.fabric_area
    if para_config.cd is not None and para_config.radius is not None:
        return para_config.cd * pi * para_config.radius ** 2
    raise ValueError("Parachute needs either cd_s, cd + radius, or cd + fabric_area in config.")


def _build_parachute(para_config: ParachuteConfig, name: str) -> Parachute:
    """Build a single Parachute."""
    options = {
        "name": name,
        "cd_s": _get_parachute_cd_s(para_config),
        "trigger": para_config.trigger,
        "sampling_rate": para_config.sampling_rate,
    }
    if para_config.lag is not None:
        options["lag"] = para_config.lag
    if para_config.noise is not None:
        options["noise"] = para_config.noise
    if para_config.radius is not None:
        options["radius"] = para_config.radius
    if para_config.cd is not None:
        options["drag_coefficient"] = para_config.cd
    return Parachute(**options)


def create_parachutes(params: SimParams, OUTPUT_LEVEL=0):
    """Create main and optional drogue parachutes and store them in params.runtime.parachutes."""
    if params.config.output_level >= 0:
        printmd("## Parachutes")
        print("Creating parachutes...")

    parachute_cfg = params.config.parachutes
        # defaults = {
    #     "parachutes_main_sampling_rate": 100,          # hz              # preset
    #     "parachutes_main_lag": 4,                       # s               # measured
    # }
    # if has_drogue:
    #     defaults.update({
    #         "parachutes_drogue_sampling_rate": 100,
    #         "parachutes_drogue_lag": 1,
    #     })
    # add_defaults(parameter_parachutes, defaults, constants, variations)

    parachute_set = {}

    main_parachute = _build_parachute(parachute_cfg.main, "main")
    set_labels(main_parachute)
    parachute_set[0] = main_parachute

    if parachute_cfg.drogue is not None:
        drogue_parachute = _build_parachute(parachute_cfg.drogue, "drogue")
        set_labels(drogue_parachute)
        parachute_set[1] = drogue_parachute

    params.runtime.parachutes = parachute_set

    if params.config.output_level > 1:
        for parachute in parachute_set.values():
            CustomPlots.plot_parachute_model(parachute, save_dir=params.project_path / "plots")


# =============================================================================
# Rocket and flight
# =============================================================================

def create_rocket(params: SimParams):
    """Build nosecone, tailcone, fins, parachutes, and the full Rocket; store in params.runtime.rocket."""
    if has_rocket_component_variations(params.config):
        printmd("## Rocket")
        print("Rocket components vary. Rocket will be built per combination in create_flight.")
        return
    if params.config.output_level >= 0:
        printmd("## Rocket")
        print("Creating rocket...")

    create_nosecone(params)
    create_tailcone(params)
    create_fins(params)
    create_parachutes(params)

    rocket_cfg     = params.config.rocket
    tailcone_cfg   = params.config.tailcone
    fins_cfg       = params.config.fins
    railbuttons_cfg = params.config.railbuttons

    # defaults = {
    #     "nozzle_position": 0,
    #     "rocket_moment_of_intertia_XY": 0.1,
    #     "rocket_moment_of_intertia_Z": 1,
    # }
    # add_defaults(parameter_rockets, defaults, constants, variations)
    
    rocket_length_m = millimeters_to_meters(rocket_cfg.length)
    center_of_mass_from_bottom = get_position_from_tip(rocket_cfg.total_CG_without_motor_from_tip, rocket_length_m)
    upper_railbutton_from_bottom = get_position_from_tip(railbuttons_cfg.upper_from_tip, rocket_length_m)
    lower_railbutton_from_bottom = get_position_from_tip(railbuttons_cfg.lower_from_tip, rocket_length_m)
    fin_position_from_bottom = millimeters_to_meters(fins_cfg.position)
    tail_position_from_bottom = millimeters_to_meters(tailcone_cfg.length - tailcone_cfg.cylindrical_section_length)

    rocket = Rocket(
        radius=millimeters_to_meters(rocket_cfg.diameter) / 2,
        mass=grams_to_kilograms(rocket_cfg.total_mass_without_motor),
        inertia=(rocket_cfg.moment_of_intertia_XY, rocket_cfg.moment_of_intertia_XY, rocket_cfg.moment_of_intertia_Z),
        power_off_drag=get_project_file(params, rocket_cfg.power_off_drag),
        power_on_drag=get_project_file(params, rocket_cfg.power_on_drag),
        center_of_mass_without_motor=center_of_mass_from_bottom,
        coordinate_system_orientation="tail_to_nose",
    )
    # The motor's origin is its nozzle exit; a negative nozzle_position puts it below the rocket's rear end.
    engine_config = params.config.liquid_engine if params.config.engine.type == "liquid" else params.config.motor
    rocket.add_motor(params.runtime.motor, position=millimeters_to_meters(engine_config.nozzle_position))
    rocket.set_rail_buttons(
        upper_button_position=upper_railbutton_from_bottom,
        lower_button_position=lower_railbutton_from_bottom,
    )
    rocket.add_surfaces(
        surfaces=[params.runtime.nosecone, params.runtime.fin_set, params.runtime.tailcone],
        positions=[rocket_length_m, fin_position_from_bottom, tail_position_from_bottom],
    )
    rocket.parachutes = list(params.runtime.parachutes.values())

    set_labels(rocket)
    params.runtime.rocket = rocket

    if params.config.output_level > 1:
        print_one_rocket(rocket, rocket_length_m)
        plot_one_rocket(rocket)


# =============================================================================
# Flights
# =============================================================================

def create_flight(params: SimParams):
    """Create one flight per scenario for every environment and config combination; assumes environment, engine and rocket exist.

    - A rocket component is varied (fins, motor, nosecone, main/drogue parachute, ...): one "nominal" flight per 
        rocket component combination and per environment.
    - A deployable payload is present: one "nominal" flight per environment and mass variation, plus one flight per scenario.
        Ascent is simulated once per environment and mass, and reused by that mass's scenario flights.
        The payload flight starts at that apogee.
    - A flight parameter is varied (rail length, heading, inclination, ...): one flight per scenario, per environment and variation. 
    - Reanalysis environment: selected scenarios for each reanalysis file; the reanalysis later adds an optional matched flight to the nominal one.

    User can select which scenarios to simulate in params.config.scenarios ("nominal" is required); "no_main" only if a drogue is configured.
    Results are saved in Flight objects and go to params.runtime.flights_by_env.
    """
    printmd("## Flights")
    print("Creating flights...")

    varied = collect_nested_variations(params.config)
    if varied:
        print("Varied fields:")
        for path, values in varied.items():
            print(f"  {'.'.join(path)}: {values}")

    selected_scenarios = params.config.scenarios
    if "nominal" not in selected_scenarios:
        raise ValueError('config.scenarios must contain "nominal": the safety analysis and deployable payload build on it.')

    has_payload = deployable_payload.has_deployable_payload(params)
    # When rocket components vary, only the nominal scenario is produced per combination.
    # Payload and flight-parameter variations still produce all selected scenarios.
    only_nominal_variation = has_rocket_component_variations(params.config)
    rebuild_rocket = only_nominal_variation

    scenarios_to_build = ["nominal"] if only_nominal_variation else selected_scenarios
    has_drogue_base = params.config.parachutes.drogue is not None
    # no_main is only simulated when there is a drogue to fall back on.
    scenarios_per_combo = len([name for name in scenarios_to_build if name != "no_main" or has_drogue_base])
    # Ascent is simulated separately only when several descent scenarios reuse it, or when a deployable payload needs it.
    reuse_ascent_flights = (has_variations(params.config) and scenarios_per_combo > 1) or has_payload

    environments = params.runtime.environments
    flights_by_env = {env_name: [] for env_name in environments}
    ascent_flights_by_env = {env_name: [] for env_name in environments}

    combinations = list(generate_config_combinations(params.config))
    total = len(combinations)

    # Each combination produces one scenario set per environment, plus an ascent per environment in reuse mode.
    flights_per_environment = scenarios_per_combo + 1 if reuse_ascent_flights else scenarios_per_combo
    total_flights = total * len(environments) * flights_per_environment

    # The metadata [finished, total] lets the Streamlit runner draw a progress bar without parsing the text.
    progress = display(
        {"text/plain": f"Flight 0/{total_flights}"},
        metadata={FLIGHT_PROGRESS_METADATA_KEY: [0, total_flights]},
        raw=True,
        display_id=True,
    )
    finished_flights = 0

    # create one flight per scenario per environment; reuse ascent flight when configured
    for combo_config in combinations:
        # Build a params view with this combo's config; runtime is shared with params
        combo_params = params.model_copy(update={"config": combo_config})

        # Rebuild engine and rocket only when a rocket component is being swept.
        # Otherwise, the pre-built engine/rocket from the separate create_engine/create_rocket
        # cells is reused (already stored in params.runtime).
        if rebuild_rocket:
            build_params = combo_params.model_copy(update={"config": combo_params.config.model_copy(update={"output_level": -1})})
            create_engine(build_params)
            create_rocket(build_params)

        # Parachute presence may differ if parachutes are the varying component
        has_drogue = combo_config.parachutes.drogue is not None

        flight_config = combo_config.flight
        payload_mass_total = combo_config.payload.mass_total if isinstance(combo_config.payload.mass_total, (int, float)) else 0
        scenario_rockets = build_scenario_rockets(combo_params.runtime.rocket, has_drogue, scenarios_to_build, payload_mass_total)
        # RocketPy adds parachute pressure data to the rocket object during flight. A fresh copy for each combination flight avoids 
        # accumulating parachute data from previous flights (memory leak).
        ascent_rocket = copy.deepcopy(combo_params.runtime.rocket) if reuse_ascent_flights else None

        meta = build_variation_meta(params.config, combo_config)
        meta_suffix = " - " + ", ".join(f"{k}={v}" for k, v in meta.items()) if meta else ""

        for env_name, environment in environments.items():
            ascent_flight = None
            scenario_set = {}

            if reuse_ascent_flights:
                # First simulate the full-mass rocket only up to apogee
                ascent_flight = Flight(
                    rocket=ascent_rocket,
                    environment=environment,
                    rail_length=flight_config.rail_length,
                    inclination=flight_config.inclination,
                    heading=flight_config.heading,
                    terminate_on_apogee=True,
                    name=f"{env_name}_ascent",
                )
                tag_variation(ascent_flight, meta)
                ascent_flights_by_env[env_name].append(ascent_flight)
                finished_flights += 1
                progress.update(
                    {"text/plain": f"Flight {finished_flights}/{total_flights}: {env_name} | ascent{meta_suffix}"},
                    metadata={FLIGHT_PROGRESS_METADATA_KEY: [finished_flights, total_flights]},
                    raw=True,
                )

            for scenario_name, rocket in scenario_rockets.items():
                # print(f"scenario_rockets={scenario_rockets}")
                flight_options = {
                    "rocket": rocket,
                    "environment": environment,
                    "rail_length": flight_config.rail_length,
                    "inclination": flight_config.inclination,
                    "heading": flight_config.heading,
                    "terminate_on_apogee": False,
                    "max_time": FLIGHT_MAX_TIME_S,
                    "name": f"{env_name}_{scenario_name}",
                }
                if ascent_flight is not None:
                    flight_options["initial_solution"] = ascent_flight

                flight = Flight(**flight_options)
                # A flight still in the air at t_final has no valid impact point, so landing/safety results would be wrong
                if not flight_reached_ground(flight):
                    warnings.warn(
                        f"{flight.name}{meta_suffix} did not reach the ground within max_time={FLIGHT_MAX_TIME_S} s "
                        f"(altitude {flight.altitude(flight.t_final):.0f} m AGL at t={flight.t_final:.0f} s); "
                        f"impact point, landing distance and safety results for this flight are invalid."
                    )
                tag_variation(flight, meta)
                scenario_set[scenario_name] = flight
                finished_flights += 1
                progress.update(
                    {"text/plain": f"Flight {finished_flights}/{total_flights}: {env_name} | {scenario_name}{meta_suffix}"},
                    metadata={FLIGHT_PROGRESS_METADATA_KEY: [finished_flights, total_flights]},
                    raw=True,
                )

            flights_by_env[env_name].append(scenario_set)

    progress.update(
        {"text/plain": f"{finished_flights}/{total_flights} flights simulated"},
        metadata={FLIGHT_PROGRESS_METADATA_KEY: [finished_flights, total_flights]},
        raw=True,
    )

    params.runtime.flights_by_env = flights_by_env
    params.runtime.scenario_sets = [s for sets in flights_by_env.values() for s in sets]

    if reuse_ascent_flights:
        params.runtime.ascent_flights_by_env = ascent_flights_by_env

    if has_payload:
        deployable_payload.create_payload(params)
        deployable_payload.create_payload_flight(params)


def build_scenario_rockets(base_rocket, has_drogue, scenario_names, payload_mass_total=0):
    """Build rocket variants for the selected descent scenarios by deep-copying and trimming the parachute list.

    Returns {scenario_name: rocket}; no_main is skipped without a drogue. When payload_mass_total > 0, the payload mass is
    removed from all scenario rockets before descent is simulated.
    """
    scenario_rockets = {}
    for scenario_name in scenario_names:
        if scenario_name == "no_main" and not has_drogue:
            continue
        # Each scenario gets its own deep copy so RocketPy can hold scenario-specific parachute lists.
        rocket = copy.deepcopy(base_rocket)
        if scenario_name == "no_main":
            rocket.parachutes = [parachute for parachute in rocket.parachutes if parachute.name != "main"]
        elif scenario_name == "ballistic":
            rocket.parachutes = []
        scenario_rockets[scenario_name] = rocket

    payload_mass_kg = grams_to_kilograms(payload_mass_total)
    if payload_mass_kg > 0:
        for rocket in scenario_rockets.values():
            # remove payload mass from rocket after separation
            rocket.mass -= payload_mass_kg
            if rocket.mass <= 0:
                raise ValueError("payload_mass_total must be smaller than the rocket mass.")

    return scenario_rockets
