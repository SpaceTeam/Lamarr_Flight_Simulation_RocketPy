import sys
import datetime
import CoolProp.CoolProp as CP

from datetime import datetime, timedelta
from rocketpy import Environment, SolidMotor, Rocket, Flight, Fluid, LiquidMotor, CylindricalTank, MassFlowRateBasedTank, TrapezoidalFins, FreeFormFins, RailButtons, NoseCone, Tail, Parachute, CompareFlights

from simulation.utils import *

DEBUG = False

def create_environment(constants, variables):
    print("Creating environment...")
    #Ponte de Sor: 39.12368, -8.03333
    #EUROC: 09.-15.10.2025
    #possible launch date: 11.10.2025

    parameter_env = []
    defaults = {
        "envType" : "Forecast",
        "latitude" : 39.232465,
        "longitude" : -8.172108,
        "timezone" : "Europe/Rome",
        "max_expected_height" : 9000,
        "elevation" : "Open-Elevation",
        "date" : (datetime.now() + timedelta(days=1))
    }
    fill_parameters(parameter_env, "environment", constants, variables)
    add_defaults(parameter_env, defaults, constants, variables, "environment")

    if DEBUG: print(parameter_env)

    environmentsForecast = []
    environmentsNormal = []
    environmentsCustom = []
    environmentsReanalysis = []

    enabled_env_types = []
    if "environment_envType" in constants:
        val = constants["environment_envType"]
        enabled_env_types = [val] if isinstance(val, str) else list(val)
    elif "environment_envType" in variables:
        val = variables["environment_envType"]
        enabled_env_types = val if isinstance(val, list) else [val]
    
    register("enabled_env_types", enabled_env_types, constants, variables)

    if DEBUG: print(enabled_env_types)

    if "Forecast" in enabled_env_types:
        for env_vals in generate_combinations(parameter_env, constants, variables):
            if "Forecast" in env_vals["environment_envType"]:
                # Environment based on forecast data
                envForecast = Environment(max_expected_height=env_vals["environment_max_expected_height"])
                envForecast.set_location(latitude=env_vals["environment_latitude"], longitude=env_vals["environment_longitude"])
                envForecast.set_date(env_vals["environment_date"], timezone=env_vals["environment_timezone"])
                envForecast.set_elevation(env_vals["environment_elevation"])
                envForecast.set_atmospheric_model(type="windy", file="ECMWF")
                # envForecast.info()
                attach_meta(envForecast, {
                    k: env_vals[k]
                    for k in variables
                    if k in env_vals
                })
                environmentsForecast.append(envForecast)
        register("envForecast", environmentsForecast, constants, variables)

    if "Normal" in enabled_env_types:
        for env_vals in generate_combinations(parameter_env, constants, variables):
            if "Normal" in env_vals["environment_envType"]:
                # Environment based on standard values
                envNormal = Environment(max_expected_height=env_vals["environment_max_expected_height"])
                envNormal.set_location(latitude=env_vals["environment_latitude"], longitude=env_vals["environment_longitude"])
                envNormal.set_date(env_vals["environment_date"], timezone=env_vals["environment_timezone"])
                envNormal.set_elevation(env_vals["environment_elevation"])
                envNormal.set_atmospheric_model(type = "standard_atmosphere")
                # envNormal.info()
                attach_meta(envNormal, {
                    k: env_vals[k]
                    for k in variables
                    if k in env_vals
                })
                environmentsNormal.append(envNormal)
        constants, variables = register("envNormal", environmentsNormal, constants, variables)

    if "Custom" in enabled_env_types:
        # TODO: add custom env variables to config
        for env_vals in generate_combinations(parameter_env, constants, variables):
            if "Custom" in env_vals["environment_envType"]:
                # Environment based on custom data
                envCustom = Environment(max_expected_height=env_vals["environment_max_expected_height"])
                envCustom.set_location(latitude=env_vals["environment_latitude"], longitude=env_vals["environment_longitude"])
                envCustom.set_atmospheric_model(type = "custom_atmosphere", wind_u = env_vals["environment_wind_u"], wind_v =  env_vals["environment_wind_v"])
                # envCustom.info()
                attach_meta(envCustom, {
                    k: env_vals[k]
                    for k in variables
                    if k in env_vals
                })
                environmentsCustom.append(envCustom)
        constants, variables = register("envCustom", environmentsCustom, constants, variables)

    if "Reanalysis" in enabled_env_types:
        for env_vals in generate_combinations(parameter_env, constants, variables):
            if "Reanalysis" in env_vals["environment_envType"]:
                # Environment based on historical data
                envReanalysis = Environment(max_expected_height=env_vals["environment_max_expected_height"])
                envReanalysis.set_location(latitude=env_vals["environment_latitude"], longitude=env_vals["environment_longitude"])
                envReanalysis.set_elevation(env_vals["environment_elevation"])
                envReanalysis.set_date((2025, 10, 12, 17, 00, 00), timezone="Europe/Lisbon")
                envReanalysis.set_atmospheric_model(
                    type="Reanalysis",
                    file="euroc_weather.nc",
                    dictionary="ECMWF",
                )
                # envReanalysis.info()
                attach_meta(envReanalysis, {
                    k: env_vals[k]
                    for k in variables
                    if k in env_vals
                })
                environmentsReanalysis.append(envReanalysis)
        constants, variables = register("envReanalysis", environmentsReanalysis, constants, variables)
    
    return constants, variables


def create_engine(constants, variables):
    print("Creating engine...")
    parameter_engine = []
    required = {"type"}
    fill_parameters(parameter_engine, "engine_", constants, variables)
    check_required(parameter_engine, required, "engine")

    engine_type, _ = lookup("engine_type", constants, variables)

    if engine_type == "liquid":
        return create_liquid_engine(constants, variables)
    elif engine_type == "solid":
        return create_solid_engine(constants, variables)
    else:
        print("ERROR") # TODO: better error handling
        sys.exit(104)

def create_liquid_engine(constants, variables):
  print("Creating liquid engine...")
  project, _ = lookup("project", constants, variables)

  parameter_motor = []

  fill_parameters_exact(parameter_motor, ["holddown_time", "nozzle_diameter", "thrust"], constants, variables)

  fill_parameters(parameter_motor, "nitrogen_tank_", constants, variables)
  fill_parameters(parameter_motor, "ethanol_tank_", constants, variables)
  fill_parameters(parameter_motor, "lox_tank_", constants, variables)
  fill_parameters(parameter_motor, "pressure_", constants, variables)
  fill_parameters(parameter_motor, "temperature_", constants, variables)

  required = [
    "length",
    "outer_diameter",
    "volume",
    "massflow"
  ]
  check_required(parameter_motor, required, "nitrogen_tank")
  check_required(parameter_motor, required, "ethanol_tank")
  check_required(parameter_motor, required, "lox_tank")

  required = [
    "nitrogen_tank_CG_ethanol",
    "nitrogen_tank_CG_lox",
  ]
  check_required(parameter_motor, required)

  defaults = {
    "temperature_nitrogen" : 298.15,                # K       # preset (= 25°C)
    "temperature_ethanol" : 298.15,                # K       # preset (= 25°C)
    "temperature_lox" : 93.15,                    # K       # preset (= -180°C)
    "pressure_nitrogen" : 30000000,              # Pa     # preset
    "pressure_ethanol" : 3000000,              # Pa     # preset
    "pressure_lox" : 3000000             # Pa     # preset
  }

  add_defaults(parameter_motor, defaults, constants, variables)

  if DEBUG: print(parameter_motor)


  motors = []

  for motor_vals in generate_combinations(parameter_motor, constants, variables):
    # Propellants

    # define density
    rho_nitrogen = CP.PropsSI("D","T",motor_vals["temperature_nitrogen"],"P|gas",motor_vals["pressure_nitrogen"],"N2")             # kg/m^3
    rho_ethanol = CP.PropsSI("D", "T|liquid", motor_vals["temperature_ethanol"], "P", motor_vals["pressure_ethanol"], "ethanol")   # kg/m^3
    rho_lox = CP.PropsSI("D", "T|liquid", motor_vals["temperature_lox"], "P", motor_vals["pressure_lox"], "oxygen")                # kg/m^3

    # define fluids
    nitrogen = Fluid(name = "N2", density = rho_nitrogen)
    ethanol = Fluid(name = "ethanol", density = rho_ethanol)
    lox = Fluid(name = "LOX", density = rho_lox)

    # define tanks geometry
    nitrogen_tank_shape = CylindricalTank(radius = motor_vals["nitrogen_tank_outer_diameter"] / 2 /1000, height = motor_vals["nitrogen_tank_length"] /1000, spherical_caps = True)
    ethanol_tank_shape = CylindricalTank(radius = motor_vals["ethanol_tank_outer_diameter"] / 2 /1000, height = motor_vals["ethanol_tank_length"] /1000, spherical_caps = True)
    lox_tank_shape = CylindricalTank(radius = motor_vals["lox_tank_outer_diameter"] / 2 /1000, height = motor_vals["lox_tank_length"] /1000, spherical_caps = True)


    m_nitrogen = motor_vals["nitrogen_tank_volume"] /1000 * rho_nitrogen                   # m
    m_ethanol = motor_vals["ethanol_tank_volume"] /1000 * rho_ethanol                      # m
    m_lox = motor_vals["lox_tank_volume"] /1000 * rho_lox                                  # m

    t_burn_nitrogen = (m_nitrogen / motor_vals["nitrogen_tank_massflow"])-0.001          # s


    t_burn_ethanol = (m_ethanol / motor_vals["ethanol_tank_massflow"])
    t_burn_lox = (m_lox / motor_vals["lox_tank_massflow"])

    if(t_burn_ethanol<t_burn_lox):
      t_burn = t_burn_ethanol
      print(f"using ethanol burn time ({t_burn_ethanol}). LOX burntime is {t_burn_lox}")
    else:
      t_burn = t_burn_lox
      print(f"using lox burn time ({t_burn_lox}). Ethanol burntime is {t_burn_ethanol}")


    #account for holddown
    t_burn -= motor_vals["holddown_time"]
    
    m_nitrogen = motor_vals["nitrogen_tank_massflow"] * t_burn + 0.0001
    m_ethanol = motor_vals["ethanol_tank_massflow"] * t_burn + 0.0001
    m_lox = motor_vals["lox_tank_massflow"] * t_burn + 0.0001

    # define tanks
    nitrogen_tank = MassFlowRateBasedTank(
        name = "nitrogen tank",
        geometry = nitrogen_tank_shape,
        flux_time = t_burn,                                 # s
        initial_liquid_mass = 0,                            # kg
        initial_gas_mass = m_nitrogen,                      # kg
        liquid_mass_flow_rate_in = 0,                       # kg/s
        liquid_mass_flow_rate_out = 0,                      # kg/s
        gas_mass_flow_rate_in = 0,                          # kg/s
        gas_mass_flow_rate_out = lambda t: motor_vals["nitrogen_tank_massflow"],   # ks/s
        liquid = Fluid(name = "liquid", density = 0.0001),  # ignore
        gas = nitrogen,
    )
    ethanol_tank = MassFlowRateBasedTank(
        name = "fuel tank",
        geometry = ethanol_tank_shape,
        flux_time = t_burn,                                 # s
        initial_liquid_mass = m_ethanol,                    # kg
        initial_gas_mass = 0,                               # kg
        liquid_mass_flow_rate_in = 0,                       # kg/s
        liquid_mass_flow_rate_out = lambda t: motor_vals["ethanol_tank_massflow"], # kg/s
        gas_mass_flow_rate_in = lambda t: motor_vals["nitrogen_tank_massflow"],    # kg/s
        gas_mass_flow_rate_out = 0,                         # kg/s
        liquid = ethanol,
        gas = nitrogen,
    )

    lox_tank = MassFlowRateBasedTank(
        name = "oxidizer tank",
        geometry = lox_tank_shape,
        flux_time = t_burn,                                 # s
        initial_liquid_mass = m_lox,                        # kg
        initial_gas_mass = 0,                               # kg
        liquid_mass_flow_rate_in = 0,                       # kg/s
        liquid_mass_flow_rate_out = lambda t: motor_vals["lox_tank_massflow"],     # kg/s
        gas_mass_flow_rate_in = lambda t: motor_vals["nitrogen_tank_massflow"],    # kg/s
        gas_mass_flow_rate_out = 0,                         # kg/s
        liquid = lox,
        gas = nitrogen,
    )

    if DEBUG: print(motor_vals["thrust"])
    skuld = LiquidMotor(
        dry_mass = 0.0001,                  # kg
        dry_inertia = (0,0,0),              # kg*m^2
        center_of_dry_mass_position = 0,    # m

        nozzle_radius = motor_vals["nozzle_diameter"] /2 /1000, # m
        nozzle_position = 0,  # m

        thrust_source = "./" + project + "/" + motor_vals["thrust"],   # N
        burn_time = t_burn,                 # s
        coordinate_system_orientation = "nozzle_to_combustion_chamber"
    )

    skuld.add_tank(tank = nitrogen_tank, position   = motor_vals["nitrogen_tank_CG_ethanol"]   / 1000)
    skuld.add_tank(tank = nitrogen_tank, position   = motor_vals["nitrogen_tank_CG_lox"]       / 1000)
    skuld.add_tank(tank = ethanol_tank, position    = motor_vals["ethanol_tank_CG"]            / 1000)
    skuld.add_tank(tank = lox_tank, position        = motor_vals["lox_tank_CG"]                / 1000)

    # skuld.all_info()
    attach_meta(skuld, {
        k: motor_vals[k]
        for k in variables
        if k in motor_vals
    })
    motors.append(skuld)
  if DEBUG: skuld.draw(filename="plots/liquid_engine.png")
  return register("motor", motors, constants, variables)



def create_solid_engine(constants, variables):
  print("Creating solid engine...")
  project, _ = lookup("project", constants, variables)
  parameter_motor = []

  fill_parameters(parameter_motor, "motor_", constants, variables)
  fill_parameters_exact(parameter_motor, "thrust", constants, variables)
  
  required = [
    "name",
    "total_mass",
    "propellant_mass",
    "diameter",
    "length",
    "burn_time"
  ]
  check_required(parameter_motor, required, "motor")
  
  defaults = {
    "motor_coordinate_system_orientation" : "nozzle_to_combustion_chamber",                # K       # preset (= 25°C)
    "motor_dry_inertia":(0.1, 0.1, 0.001),
    "holddown_time":0,

  }
  add_defaults(parameter_motor, defaults, constants, variables)
  
  if DEBUG: print(parameter_motor)


  motors = []

  for motor_vals in generate_combinations(parameter_motor, constants, variables):

    motor_volume            = pi * ((motor_vals["motor_diameter"] /1000 /2) ** 2) * motor_vals["motor_length"] /1000     #m³
    motor_grain_density     = (motor_vals["motor_propellant_mass"] /1000)/motor_volume          #kg/m³

    motor = SolidMotor(
      thrust_source                   = "./" + project + "/" + motor_vals["thrust"],
      dry_mass                        = motor_vals["motor_total_mass"] /1000 - motor_vals["motor_propellant_mass"] /1000,
      dry_inertia                     = motor_vals["motor_dry_inertia"],                        #kg*m²   #guess / not so relevant
      nozzle_radius                   = motor_vals["motor_diameter"] /1000 /2,
      grain_number                    = 1,
      grain_density                   = motor_grain_density,
      grain_outer_radius              = motor_vals["motor_diameter"] /1000 /2,
      grain_initial_inner_radius      = 0,
      grain_initial_height            = motor_vals["motor_length"] /1000,
      grain_separation                = 0,
      grains_center_of_mass_position  = motor_vals["motor_length"] /1000 / 2,
      center_of_dry_mass_position     = motor_vals["motor_length"] /1000 / 2,
      nozzle_position                 = 0,
      burn_time                       = motor_vals["motor_burn_time"],
      throat_radius                   = motor_vals["motor_diameter"] /1000 /2 / 2,
      coordinate_system_orientation   = motor_vals["motor_coordinate_system_orientation"],
    )
    # motor.all_info()

    attach_meta(motor, {
        k: motor_vals[k]
        for k in variables
        if k in motor_vals
    })
    motors.append(motor)
  if DEBUG: motor.draw(filename="plots/solid_engine.png")
  return register("motor", motors, constants, variables)


def create_nosecone(constants, variables):
    print("Creating nosecone...")
    parameter_nosecone = []
    required = ["length"]
    defaults = {"nosecone_kind":"von karman"}
    fill_parameters_exact(parameter_nosecone, "rocket_diameter", constants, variables)
    fill_parameters(parameter_nosecone, "nosecone_", constants, variables)
    check_required(parameter_nosecone, required, "nosecone")
    add_defaults(parameter_nosecone, defaults, constants, variables)

    if DEBUG: print(parameter_nosecone)

    nosecones = []

    for nosecone_vals in generate_combinations(parameter_nosecone, constants, variables):
        nosecone = NoseCone(
            length = nosecone_vals["nosecone_length"] /1000,
            base_radius = nosecone_vals["rocket_diameter"]/2 /1000,
            kind = nosecone_vals["nosecone_kind"]
        )
        attach_meta(nosecone, {
            k: nosecone_vals[k]
            for k in variables
            if k in nosecone_vals
        })
        nosecones.append(nosecone)
    if DEBUG: nosecone.draw(filename="plots/nosecone.png")
    return register("nosecone", nosecones, constants, variables)


def create_tailcone(constants, variables):
    print("Creating tailcone...")
    parameter_tailcone = []
    required = ["diameter", "length"]
    defaults = {"tailcone_cylindrical_length":0}
    fill_parameters(parameter_tailcone, "tailcone_", constants, variables)
    fill_parameters_exact(parameter_tailcone, "rocket_diameter", constants, variables)
    check_required(parameter_tailcone, required, "tailcone")
    add_defaults(parameter_tailcone, defaults, constants, variables)

    if DEBUG: print(parameter_tailcone)

    tailcones = []

    for tailcone_vals in generate_combinations(parameter_tailcone, constants, variables):
        tailcone = Tail(
            top_radius = tailcone_vals["rocket_diameter"] /2 /1000,
            bottom_radius = tailcone_vals["tailcone_diameter"] /2 /1000,
            length = tailcone_vals["tailcone_length"] /1000,
            rocket_radius = tailcone_vals["rocket_diameter"]/2 /1000
        )
        attach_meta(tailcone, {
            k: tailcone_vals[k]
            for k in variables
            if k in tailcone_vals
        })
        tailcones.append(tailcone)
    return register("tailcone", tailcones, constants, variables)

def create_fins(constants, variables):
    print("Creating fins...")
    parameter_fins = []

    defaults = {"fins_amount" : 4, "fins_name":"fins"}

    fill_parameters(parameter_fins, "fins_", constants, variables)
    fill_parameters_exact(parameter_fins, ["rocket_diameter","tailcone_diameter"], constants, variables)


    if not parameter_fins.__contains__("fins_shape_points"):
        required=[
            "root_chord",
            "tip_chord",
            "span",
            "sweep_length"
        ]
        check_required(parameter_fins, required, "fins")

    add_defaults(parameter_fins, defaults, constants, variables)

    if DEBUG: print(parameter_fins)


    fins = []
    trapezoidal_fins = []

    for fin_vals in generate_combinations(parameter_fins, constants, variables):
        if "fins_shape_points" in parameter_fins:
            fin_set = FreeFormFins(
                n = fin_vals["fins_amount"],
                shape_points = fin_vals["fins_shape_points"],         # m
                rocket_radius=fin_vals["tailcone_diameter"] /2 /1000,    # m
                name = "Freeform"
            )
            # fin_set.draw()
            attach_meta(fin_set, {
                k: fin_vals[k]
                for k in variables
                if k in fin_vals
            })
            fins.append(fin_set)
        else:
            trapezoidal_fin_set = TrapezoidalFins(
                n = fin_vals["fins_amount"],
                root_chord = fin_vals["fins_root_chord"] /1000,      # m
                tip_chord = fin_vals["fins_tip_chord"] /1000,        # m
                span = fin_vals["fins_span"] /1000,                  # m
                sweep_length = fin_vals["fins_sweep_length"] /1000,  # m
                name = fin_vals["fins_name"],
                rocket_radius = fin_vals["rocket_diameter"]/2 /1000 # m
            )
            # trapezoidal_fin_set.draw()
            attach_meta(trapezoidal_fin_set, {
                k: fin_vals[k]
                for k in variables
                if k in fin_vals
            })
            trapezoidal_fins.append(trapezoidal_fin_set)

    if "fins_shape_points" in parameter_fins:
        constants, variables = register("fin_set", fins, constants, variables)
        if DEBUG: fin_set.draw(filename="plots/fin_set.png")
        if DEBUG: print(fins)
    else:
        constants, variables = register("fin_set", trapezoidal_fins, constants, variables)
        if DEBUG: trapezoidal_fin_set.draw(filename="plots/trapezoidal_fin_set.png")
        if DEBUG: print(trapezoidal_fins)
    return constants, variables

def create_parachutes(constants, variables):
    print("Creating parachutes...")

    has_drogue = any(
        k.startswith("drogue_") for k in constants
    ) or any(
        k.startswith("drogue_") for k in variables
    )
    
    parameter_parachutes = []

    required=[
        "cd_s",
        "trigger"
    ]

    defaults = {
        "main_sampling_rate": 105,           # hz              # preset
        "main_lag": 4,                       # s               # measured
        "main_noise": (0, 8.3, 0.5)          # (pa, pa, pa)    # preset
    }
    if has_drogue:
        defaults.update({
            "drogue_sampling_rate": 105,         # hz              # preset
            "drogue_lag": 1,                     # s               # measured
            "drogue_noise": (0, 8.3, 0.5),       # (pa, pa, pa)    # preset
        })

    if has_drogue: fill_parameters(parameter_parachutes, "drogue_", constants, variables)
    fill_parameters(parameter_parachutes, "main_", constants, variables)

    if has_drogue: check_required(parameter_parachutes, required, "drogue")
    check_required(parameter_parachutes, required, "main")

    add_defaults(parameter_parachutes, defaults, constants, variables)

    if DEBUG: print(parameter_parachutes)


    parachutes = []

    for parachute_vals in generate_combinations(parameter_parachutes, constants, variables):
        parachute_list = {}
        parachute_list[0] = Parachute(
            name = "main",
            cd_s = parachute_vals["main_cd_s"],
            trigger = parachute_vals["main_trigger"],             # m
            sampling_rate = parachute_vals["main_sampling_rate"], # hz
            lag = parachute_vals["main_lag"],                     # s
            noise = parachute_vals["main_noise"],                 # (pa, pa, pa)
        )
        attach_meta(parachute_list[0], {
            k: parachute_vals[k]
            for k in variables
            if k in parachute_vals
        })

        if has_drogue:
            parachute_list[1] = Parachute(
                name = "drogue",
                cd_s = parachute_vals["drogue_cd_s"],
                trigger = parachute_vals["drogue_trigger"],             # m
                sampling_rate = parachute_vals["drogue_sampling_rate"], # hz
                lag = parachute_vals["drogue_lag"],                     # s
                noise = parachute_vals["drogue_noise"],                 # (pa, pa, pa)
            )
            attach_meta(parachute_list[1], {
                k: parachute_vals[k]
                for k in variables
                if k in parachute_vals
            })
        parachutes.append(parachute_list)
    return register("parachutes", parachutes, constants, variables)



def create_rocket_parts(constants, variables):
    constants, variables = create_nosecone(constants, variables)
    constants, variables = create_tailcone(constants, variables)
    constants, variables = create_fins(constants, variables)
    constants, variables = create_parachutes(constants, variables)
    return constants, variables

def create_rocket(constants, variables):
    print("Creating rocket...")
    constants, variables = create_rocket_parts(constants, variables)
    project, _ = lookup("project", constants, variables)
    # add MOI to defaults or required

    parameter_rockets = []

    required=[
        "diameter",
        "dry_mass",
        "CG",
        "length"
    ]
    defaults={
        "nozzle_position":0,
        "rocket_moment_of_intertia_XY":0.1,
        "rocket_moment_of_intertia_Z":1
    }

    fill_parameters(parameter_rockets, "rocket_", constants, variables)
    fill_parameters(parameter_rockets, "railbuttons_", constants, variables)
    fill_parameters_exact(parameter_rockets, [
        "nozzle_position",
        "fins_position",
        "tailcone_length",
        "tailcone_cylindrical_length",
        "motor",
        "nosecone",
        "tailcone",
        "fin_set",
        "parachutes"], constants, variables)
    check_required(parameter_rockets, required, "rocket")
    add_defaults(parameter_rockets, defaults, constants, variables)

    if DEBUG: print(parameter_rockets)

    rockets = []

    for rocket_vals in generate_combinations(parameter_rockets, constants, variables):
        rocket = Rocket(
            radius = rocket_vals["rocket_diameter"] /2 / 1000,                      # m
            mass = rocket_vals["rocket_dry_mass"] / 1000,                                # m
            inertia = (rocket_vals["rocket_moment_of_intertia_XY"], rocket_vals["rocket_moment_of_intertia_XY"], rocket_vals["rocket_moment_of_intertia_Z"]),  # kg * m^2
            power_off_drag = "./" + project + "/power_off_drag.csv",
            power_on_drag = "./" + project + "/power_on_drag.csv",
            center_of_mass_without_motor = rocket_vals["rocket_CG"] / 1000,                # m
            coordinate_system_orientation = "tail_to_nose"
        )
        rocket.add_motor(rocket_vals["motor"], position = rocket_vals["nozzle_position"] /1000)
        rocket.set_rail_buttons(upper_button_position= rocket_vals["railbuttons_upper"] / 1000, lower_button_position=rocket_vals["railbuttons_lower"] / 1000)
        rocket.add_surfaces(surfaces=[rocket_vals["nosecone"], rocket_vals["fin_set"], rocket_vals["tailcone"]], positions=[rocket_vals["rocket_length"] / 1000, rocket_vals["fins_position"] / 1000, rocket_vals["tailcone_length"] / 1000])

        rocket.parachutes = list(rocket_vals["parachutes"].values())

        #hedy.all_info()
        attach_meta(rocket, {
            k: rocket_vals[k]
            for k in variables
            if k in rocket_vals
        })
        rockets.append(rocket)

    if DEBUG:
        vis_args = {
            "background": "#EEEEEE",
            "tail": "black",
            "nose": "black",
            "body": "black",
            "fins": "black",
            "motor": "black",
            "buttons": "black",
            "line_width": 2.0,
        }
        rocket.draw(plane = 'xz', vis_args = vis_args, filename="plots/rocket.png")
    
    return register("rocket", rockets, constants, variables)


def create_flight(constants, variables):
    print("Creating flights...")
    enabled_env_types, _ = lookup("enabled_env_types", constants, variables)

    if isinstance(enabled_env_types, str):
        enabled_env_types = [enabled_env_types]
    else:
        enabled_env_types = list(enabled_env_types)

    parameter_flights = []

    required = [
        "rail_length",
        "inclination",
        "heading"
    ]

    defaults = {
        "flight_terminate_on_apogee" : True
    }

    fill_parameters(parameter_flights, "flight_", constants, variables)
    fill_parameters_exact(parameter_flights, "rocket", constants, variables)
    check_required(parameter_flights, required, "flight")
    add_defaults(parameter_flights, defaults, constants, variables)

    if DEBUG: print(parameter_flights)

    for type in enabled_env_types:
        parameter_flights.append("env" + type)


    flights_normal = []
    flights_forecast = []
    flights_custom = []
    flights_reanalysis = []

    total = count_combinations(parameter_flights, constants, variables)
    step = max(1, total // 10)
    print(total)

    for i, flight_vals in enumerate(generate_combinations(parameter_flights, constants, variables), start=1):

        if "Normal" in enabled_env_types:
            flight_normal = Flight(
                    rocket        = flight_vals["rocket"],
                    environment   = flight_vals["envNormal"],
                    rail_length   = flight_vals["flight_rail_length"],
                    inclination   = flight_vals["flight_inclination"],
                    heading       = flight_vals["flight_heading"],
                    terminate_on_apogee = flight_vals["flight_terminate_on_apogee"],
                    name          = "Normal"
            )
            #flight_normal.prints.out_of_rail_conditions()
            #flight_normal.prints.apogee_conditions()
            #flight_normal.prints.impact_conditions()
            #flight_normal.prints.maximum_values()
            #flight_normal.plots.trajectory_3d()
            #flight_normal.plots.stability_and_control_data()

            #flight_normal.plots.all()
            #flight_normal.prints.all()
            #flight_normal.all_info()
            attach_meta(flight_normal, {
                k: flight_vals[k]
                for k in variables
                if k in flight_vals
            })
            flights_normal.append(flight_normal)


        if "Forecast" in enabled_env_types:
            flight_forecast = Flight(
                    rocket        = flight_vals["rocket"],
                    environment   = flight_vals["envForecast"],
                    rail_length   = flight_vals["flight_rail_length"],
                    inclination   = flight_vals["flight_inclination"],
                    heading       = flight_vals["flight_heading"],
                    terminate_on_apogee = flight_vals["flight_terminate_on_apogee"],
                    name          = "Forecast"
            )
            #flight_forecast.prints.out_of_rail_conditions()
            #flight_forecast.prints.apogee_conditions()
            #flight_forecast.prints.impact_conditions()
            #flight_forecast.prints.maximum_values()
            #flight_forecast.plots.trajectory_3d()
            #flight_forecast.plots.stability_and_control_data()

            #flight_forecast.prints.all()
            #flight_forecast.plots.all()
            #flight_forecast.all_info()
            attach_meta(flight_forecast, {
                k: flight_vals[k]
                for k in variables
                if k in flight_vals
            })
            flights_forecast.append(flight_forecast)
        if "Custom" in enabled_env_types:
            flight_custom = Flight(
                    rocket        = flight_vals["rocket"],
                    environment   = flight_vals["envCustom"],
                    rail_length   = flight_vals["flight_rail_length"],
                    inclination   = flight_vals["flight_inclination"],
                    heading       = flight_vals["flight_heading"],
                    terminate_on_apogee = flight_vals["flight_terminate_on_apogee"],
                    name          = "Custom"
            )
            #flight_custom.prints.out_of_rail_conditions()
            #flight_custom.prints.apogee_conditions()
            #flight_custom.prints.impact_conditions()
            #flight_custom.prints.maximum_values()
            #flight_custom.plots.trajectory_3d()
            #flight_custom.plots.stability_and_control_data()

            #flight_custom.prints.all()
            #flight_custom.plots.all()
            #flight_custom.all_info()
            attach_meta(flight_custom, {
                k: flight_vals[k]
                for k in variables
                if k in flight_vals
            })
            flights_custom.append(flight_custom)
        
        if "Reanalysis" in enabled_env_types:
            flight_reanalysis = Flight(
                    rocket        = flight_vals["rocket"],
                    environment   = flight_vals["envReanalysis"],
                    rail_length   = flight_vals["flight_rail_length"],
                    inclination   = flight_vals["flight_inclination"],
                    heading       = flight_vals["flight_heading"],
                    terminate_on_apogee = flight_vals["flight_terminate_on_apogee"],
                    name          = "Reanalysis"
            )
            #flight_reanalysis.prints.out_of_rail_conditions()
            #flight_reanalysis.prints.apogee_conditions()
            #flight_reanalysis.prints.impact_conditions()
            #flight_reanalysis.prints.maximum_values()
            #flight_reanalysis.plots.trajectory_3d()
            #flight_reanalysis.plots.stability_and_control_data()

            #flight_reanalysis.prints.all()
            #flight_reanalysis.plots.all()
            #flight_reanalysis.all_info()
            attach_meta(flight_reanalysis, {
                k: flight_vals[k]
                for k in variables
                if k in flight_vals
            })
            flights_reanalysis.append(flight_reanalysis)

        if i % step == 0 or i == total:
            print(f"{i / total:.0%}")
    
    if len(flights_normal) > 0:constants, variables = register("normal_flight", flights_normal, constants, variables)
    if len(flights_forecast) > 0:constants, variables = register("forecast_flight", flights_forecast, constants, variables)
    if len(flights_custom) > 0:constants, variables = register("custom_flight", flights_custom, constants, variables)
    if len(flights_reanalysis) > 0:constants, variables = register("reanalysis_flight", flights_reanalysis, constants, variables)
    #flight_forecast.all_info()
    return constants, variables