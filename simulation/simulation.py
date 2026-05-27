"""
RocketPy simulation backend.
"""

import copy
import datetime
from math import pi
from pathlib import Path

import CoolProp.CoolProp as CP
import numpy as np
import pandas as pd

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

from simulation.utils import *
from simulation.config_schema import SimParams, ParachuteConfig
from simulation.export_weather_data import export_weather_data
from simulation.outputs import *
from simulation.custom_print_and_plot_functions import CustomPlots, plot_wind_speed_and_heading
from simulation import deployable_payload

DEBUG = False


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

    if date == "tomorrow_08_local":
        tomorrow = datetime.date.today() + datetime.timedelta(days=1)
        date = (tomorrow.year, tomorrow.month, tomorrow.day, 8, 0, 0)

    if isinstance(date, (list, tuple)) and len(date) == 4:
        date = (date[0], date[1], date[2], date[3], 0, 0)

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

    if "custom_atmosphere" in environment_types:
        if not env_config.custom_weather_models:
            raise ValueError("environment.custom_weather_models is required for custom_atmosphere envType.")
        output_folder = params.project_path / "weather_csvs"
        launch_time = datetime.datetime(*date).strftime("%Y-%m-%dT%H:%M:%S")

        for weather_model in ensure_list(env_config.custom_weather_models):
            print(f"\nCreating weather CSV files for {weather_model}...")
            export_weather_data(
                latitude=latitude,
                longitude=longitude,
                max_expected_height_agl_m=max_expected_height_agl,
                launch_time=launch_time,
                timezone=timezone,
                weather_model=weather_model,
                output_folder=output_folder,
            )
            
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
    engine_type = params.config.engine.type

    if engine_type == "liquid":
        create_liquid_engine(params)
    elif engine_type == "solid":
        create_solid_engine(params)
    else:
        raise ValueError(f"Unknown engine_type '{engine_type}'. Use 'solid' or 'liquid'.")


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
    motor.add_tank(tank=press_tank, position=millimeters_to_meters(press_tank_cfg.position_fuel))
    motor.add_tank(tank=press_tank, position=millimeters_to_meters(press_tank_cfg.position_oxidizer))
    motor.add_tank(tank=fuel_tank,  position=millimeters_to_meters(fuel_tank_cfg.position))
    motor.add_tank(tank=ox_tank,    position=millimeters_to_meters(ox_tank_cfg.position))

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
    rocket.add_motor(params.runtime.motor, position=0)
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
    """Create scenario flights for every configured environment; store in params.runtime.flights_by_env.

    Assumes create_environment, create_engine, and create_rocket have already run.
    Rebuilds engine and rocket per combination only when a rocket component field is being swept.

    - Scan mode activates whenever any config field carries multiple values, or a deployable payload is present.
    - Rocket component variations (fins, motor, nosecone, etc.) produce one nominal flight per combination.
    - Payload mass variation produces all scenarios (nominal / no main / ballistic) per combination.
    - No variation produces all scenarios per combination.
    """
    printmd("## Flights")
    print("Creating flights...")

    varied = collect_nested_variations(params.config)
    if varied:
        print("Sweeping:")
        for path, values in varied.items():
            print(f"  {'.'.join(path)}: {values}")

    has_payload = deployable_payload.has_deployable_payload(params)
    # When rocket components vary, only the nominal scenario is produced per combination.
    # Payload and flight-parameter sweeps still produce all scenarios.
    only_nominal_sweep = has_rocket_component_variations(params.config)
    rebuild_rocket = only_nominal_sweep
    # Ascent is simulated separately only when multiple descent scenarios reuse it, or when a
    # payload needs it. Rocket-component sweeps produce one full nominal flight — no ascent split needed.
    reuse_ascent_flights = (has_variations(params.config) and not only_nominal_sweep) or has_payload

    environments = params.runtime.environments
    flights_by_env = {env_name: [] for env_name in environments}
    ascent_flights_by_env = {env_name: [] for env_name in environments}

    combinations = list(generate_config_combinations(params.config))
    total = len(combinations)

    has_drogue_base = params.config.parachutes.drogue is not None
    if only_nominal_sweep:
        scenarios_per_combo = 1
    else:
        scenarios_per_combo = 3 if has_drogue_base else 2
    flights_per_combo = scenarios_per_combo + 1 if reuse_ascent_flights else scenarios_per_combo
    # Each combination produces one scenario set per environment, plus an ascent in reuse mode.
    print(f"Amount: {total * flights_per_combo * len(environments)}")

    # create one flight per scenario per environment; reuse ascent flight when configured
    for i, combo_config in enumerate(combinations, start=1):
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
        scenario_rockets = build_scenario_rockets(combo_params.runtime.rocket, has_drogue, payload_mass_total)

        meta = build_variation_meta(params.config, combo_config)
        meta_str = (", ".join(f"{k}={v}" for k, v in meta.items())) if meta else ""
        print(f"{i / total:.0%}" + (f" - {meta_str}" if meta_str else ""))
        
        for env_name, environment in environments.items():
            # if i / total >= 0.77:
            #     print(f"env={env_name}, flight_values={flight_values}")
            ascent_flight = None
            scenario_set = {}

            if reuse_ascent_flights:
                # First simulate the full-mass rocket only up to apogee
                ascent_flight = Flight(
                    rocket=combo_params.runtime.rocket,
                    environment=environment,
                    rail_length=flight_config.rail_length,
                    inclination=flight_config.inclination,
                    heading=flight_config.heading,
                    terminate_on_apogee=True,
                    name=f"{env_name}_ascent",
                )
                tag_variation(ascent_flight, meta)
                ascent_flights_by_env[env_name].append(ascent_flight)

            # Reanalysis environments in single-flight mode: nominal only (matched flight is added separately).
            # Reanalysis environments in sweep mode, or any other env: all scenarios.
            reanalysis_single = env_name.startswith("Reanalysis") and not has_variations(params.config)
            scenarios_for_env = (
                {"nominal": scenario_rockets["nominal"]}
                if only_nominal_sweep or reanalysis_single
                else scenario_rockets
            )

            for scenario_name, rocket in scenarios_for_env.items():
                # print(f"scenario_rockets={scenario_rockets}")
                flight_options = {
                    "rocket": rocket,
                    "environment": environment,
                    "rail_length": flight_config.rail_length,
                    "inclination": flight_config.inclination,
                    "heading": flight_config.heading,
                    "terminate_on_apogee": False,
                    "name": f"{env_name}_{scenario_name}",
                }
                if ascent_flight is not None:
                    flight_options["initial_solution"] = ascent_flight

                flight = Flight(**flight_options)
                tag_variation(flight, meta)
                scenario_set[scenario_name] = flight

            flights_by_env[env_name].append(scenario_set)

    params.runtime.flights_by_env = flights_by_env
    params.runtime.scenario_sets = [s for sets in flights_by_env.values() for s in sets]

    if reuse_ascent_flights:
        params.runtime.ascent_flights_by_env = ascent_flights_by_env

    if has_payload:
        deployable_payload.create_payload(params)
        deployable_payload.create_payload_flight(params)


def build_scenario_rockets(base_rocket, has_drogue, payload_mass_total=0):
    """Build rocket variants for descent scenarios by deep-copying and trimming the parachute list.

    Returns {scenario_name: rocket}. When payload_mass_total > 0, the payload mass is
    removed from all scenario rockets before descent is simulated.
    """
    # Each scenario gets its own deep copy so RocketPy can hold scenario-specific parachute lists.
    nominal_rocket = copy.deepcopy(base_rocket)
    ballistic_rocket = copy.deepcopy(base_rocket)
    ballistic_rocket.parachutes = []

    if not has_drogue:
        scenario_rockets = {"nominal": nominal_rocket, "ballistic": ballistic_rocket}
    else:
        no_main_rocket = copy.deepcopy(base_rocket)
        no_main_rocket.parachutes = [parachute for parachute in no_main_rocket.parachutes if parachute.name != "main"]
        scenario_rockets = {"nominal": nominal_rocket, "no_main": no_main_rocket, "ballistic": ballistic_rocket}

    payload_mass_kg = grams_to_kilograms(payload_mass_total)
    if payload_mass_kg > 0:
        for rocket in scenario_rockets.values():
            # remove payload mass from rocket after separation
            rocket.mass -= payload_mass_kg
            if rocket.mass <= 0:
                raise ValueError("payload_mass_total must be smaller than the rocket mass.")

    return scenario_rockets
