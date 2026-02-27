import copy
from rocketpy import Rocket, Flight, Parachute

from simulation.utils import *

DEBUG = False

# w_p = without payload
def create_rocket_without_payload(constants, variables):
    parameter_rockets_w_p = []
    required=[
        "payload_mass_total", 
        "rocket"
    ]
    fill_parameters_exact(parameter_rockets_w_p, ["payload_mass_total", "rocket"], constants, variables)
    check_required(parameter_rockets_w_p, required)

    if DEBUG: print(parameter_rockets_w_p)

    rockets_w_p = []

    for rocket_w_p_vals in generate_combinations(parameter_rockets_w_p, constants, variables):
        rocket_w_p = {}
        rocket_w_p["nominal"] = copy.deepcopy(rocket_w_p_vals["rocket"])
        rocket_w_p["nominal"].mass -= rocket_w_p_vals["payload_mass_total"]

        rocket_w_p["no_main"] = copy.deepcopy(rocket_w_p["nominal"])
        rocket_w_p["no_main"].parachutes.pop()
        rocket_w_p["ballistic"] = copy.deepcopy(rocket_w_p["no_main"])
        rocket_w_p["ballistic"].parachutes.clear()

        attach_meta(rocket_w_p["nominal"], {
            k: rocket_w_p_vals[k]
            for k in variables
            if k in rocket_w_p_vals
        })
        attach_meta(rocket_w_p["no_main"], {
            k: rocket_w_p_vals[k]
            for k in variables
            if k in rocket_w_p_vals
        })
        attach_meta(rocket_w_p["ballistic"], {
            k: rocket_w_p_vals[k]
            for k in variables
            if k in rocket_w_p_vals
        })
        rockets_w_p.append(rocket_w_p)
    return register("rocket_without_payload", rockets_w_p, constants, variables)


# w_p = without payload

def create_flight_without_payload(constants, variables):
    all_flights = (
        lookup("normal_flight", constants, variables)[0] +
        lookup("forecast_flight", constants, variables)[0] +
        lookup("custom_flight", constants, variables)[0] +
        lookup("reanalysis_flight", constants, variables)[0]
    )

    register("all_flights", all_flights, constants, variables)

    
    parameter_flight_w_p = []
    
    fill_parameters_exact(parameter_flight_w_p, "rocket_without_payload", constants, variables)
    fill_parameters_exact(parameter_flight_w_p, "all_flights", constants, variables)

    if DEBUG: print(parameter_flight_w_p)

    flights_w_p = []
    flights_w_p_no_main = []
    flights_w_p_ballistic = []
    total = count_combinations(parameter_flight_w_p, constants, variables)
    step = max(1, total // 10)
    print(total*3)

    for i, flight_w_p_vals in enumerate(generate_combinations(parameter_flight_w_p, constants, variables), start=1):
        flight_w_p = Flight(
            rocket        = flight_w_p_vals["rocket_without_payload"]["nominal"],
            environment   = flight_w_p_vals["all_flights"].env,
            rail_length   = flight_w_p_vals["all_flights"].rail_length,
            inclination   = flight_w_p_vals["all_flights"].inclination,
            heading       = flight_w_p_vals["all_flights"].heading,
            terminate_on_apogee = False,
            initial_solution = flight_w_p_vals["all_flights"],
            name          = "Rocket_touchdown_nominal"
        )

        flight_w_p_no_main = Flight(
            rocket        = flight_w_p_vals["rocket_without_payload"]["no_main"],
            environment   = flight_w_p_vals["all_flights"].env,
            rail_length   = flight_w_p_vals["all_flights"].rail_length,
            inclination   = flight_w_p_vals["all_flights"].inclination,
            heading       = flight_w_p_vals["all_flights"].heading,
            terminate_on_apogee = False,
            initial_solution = flight_w_p_vals["all_flights"],
            name          = "Rocket_touchdown_no_main"
        )

        flight_w_p_ballistic = Flight(
            rocket        = flight_w_p_vals["rocket_without_payload"]["ballistic"],
            environment   = flight_w_p_vals["all_flights"].env,
            rail_length   = flight_w_p_vals["all_flights"].rail_length,
            inclination   = flight_w_p_vals["all_flights"].inclination,
            heading       = flight_w_p_vals["all_flights"].heading,
            terminate_on_apogee = False,
            initial_solution = flight_w_p_vals["all_flights"],
            name          = "Rocket_touchdown_ballistic"
        )

        attach_meta(flight_w_p, {
            k: flight_w_p_vals[k]
            for k in variables
            if k in flight_w_p_vals
        })

        attach_meta(flight_w_p_no_main, {
            k: flight_w_p_vals[k]
            for k in variables
            if k in flight_w_p_vals
        })

        attach_meta(flight_w_p_ballistic, {
            k: flight_w_p_vals[k]
            for k in variables
            if k in flight_w_p_vals
        })
        flights_w_p.append(flight_w_p) 
        flights_w_p_no_main.append(flight_w_p_no_main) 
        flights_w_p_ballistic.append(flight_w_p_ballistic)      
        if i % step == 0 or i == total:
            print(f"{i / total:.0%}")
    constants, variables = register("flight_without_payload", flights_w_p, constants, variables)
    constants, variables = register("flight_without_payload_no_main", flights_w_p_no_main, constants, variables)
    constants, variables = register("flight_without_payload_ballistic", flights_w_p_ballistic, constants, variables)
    return constants, variables
# TODO: first calculate all the safe headings, then begin to calculate payload flights based on that

def create_payload_parachute(constants, variables):
    parameter_payload_parachute = []

    required=[
        "cd_s",
        "trigger"
    ]

    defaults = {
        "sampling_rate": 105,        # hz              # preset
        "lag": 1,                    # s               # measured
        "noise": (0, 8.3, 0.5)       # (pa, pa, pa)    # preset
    }
    fill_parameters(parameter_payload_parachute, "parachute_payload_", constants, variables)

    check_required(parameter_payload_parachute, required, "parachute_payload")

    add_defaults(parameter_payload_parachute, defaults, constants, variables, "parachute_payload")

    if DEBUG: print(parameter_payload_parachute)

    payload_parachutes = []

    for parachute_vals in generate_combinations(parameter_payload_parachute, constants, variables):
        parachute_list = {}
        parachute_list[0] = Parachute(
            name = "parachute_payload",
            cd_s = parachute_vals["parachute_payload_cd_s"],
            trigger = parachute_vals["parachute_payload_trigger"],             # m
            sampling_rate = parachute_vals["parachute_payload_sampling_rate"], # hz
            lag = parachute_vals["parachute_payload_lag"],                     # s
            noise = parachute_vals["parachute_payload_noise"],                 # (pa, pa, pa)
        )
        attach_meta(parachute_list[0], {
            k: parachute_vals[k]
            for k in variables
            if k in parachute_vals
        })
        payload_parachutes.append(parachute_list)
    return register("payload_parachute", payload_parachutes, constants, variables)


def create_payload(constants, variables):
    project = lookup("project", constants, variables)[0]
    constants, variables = create_payload_parachute(constants, variables)
    parameter_payload = []

    required=[
        "diameter",
        "mass",
        "length"
    ]
    defaults={
        "payload_moment_of_intertia_XY":0.01,
        "payload_moment_of_intertia_Z":0.01
    }

    fill_parameters(parameter_payload, "payload_", constants, variables)
    check_required(parameter_payload, required, "payload")
    add_defaults(parameter_payload, defaults, constants, variables)

    if DEBUG: print(parameter_payload)

    payloads = []

    for payload_vals in generate_combinations(parameter_payload, constants, variables):
        payload = {}
        payload["nominal"] = Rocket(
            radius = payload_vals["payload_diameter"] /2 / 1000,                      # m
            mass = payload_vals["payload_mass"] / 1000,                                # m
            inertia = (payload_vals["payload_moment_of_intertia_XY"], payload_vals["payload_moment_of_intertia_XY"], payload_vals["payload_moment_of_intertia_Z"]),  # kg * m^2
            power_off_drag = "./" + project + "/power_off_drag.csv",
            power_on_drag = "./" + project + "/power_off_drag.csv",
            center_of_mass_without_motor = payload_vals["payload_length"] / 2 / 1000,                # m
            coordinate_system_orientation = "tail_to_nose"
        )

        payload["no_chute"] = copy.deepcopy(payload["nominal"])
        
        payload["nominal"].parachutes = list(payload_vals["payload_parachute"].values())



        attach_meta(payload["nominal"], {
            k: payload_vals[k]
            for k in variables
            if k in payload_vals
        })
        attach_meta(payload["no_chute"], {
            k: payload_vals[k]
            for k in variables
            if k in payload_vals
        })
        payloads.append(payload)
    return register("payload", payloads, constants, variables)


def create_payload_flight(constants, variables):
    parameter_payload_flights = []
    fill_parameters_exact(parameter_payload_flights, "all_flights", constants, variables)
    fill_parameters_exact(parameter_payload_flights, "payload", constants, variables)

    if DEBUG: print(parameter_payload_flights)

    flights_payload = []
    flights_payload_no_chute = []

    total = count_combinations(parameter_payload_flights, constants, variables)
    step = max(1, total // 10)
    print(total*2)

    for i, flight_payload_vals in enumerate(generate_combinations(parameter_payload_flights, constants, variables), start=1):
        flight_payload = Flight(
            rocket        = flight_payload_vals["payload"]["nominal"],
            environment   = flight_payload_vals["all_flights"].env,
            rail_length   = flight_payload_vals["all_flights"].rail_length,
            inclination   = flight_payload_vals["all_flights"].inclination,
            heading       = flight_payload_vals["all_flights"].heading,
            terminate_on_apogee = False,
            initial_solution = flight_payload_vals["all_flights"],
            name          = "Payload"
        )
        

        flight_payload_no_chute = Flight(
            rocket        = flight_payload_vals["payload"]["no_chute"],
            environment   = flight_payload_vals["all_flights"].env,
            rail_length   = flight_payload_vals["all_flights"].rail_length,
            inclination   = flight_payload_vals["all_flights"].inclination,
            heading       = flight_payload_vals["all_flights"].heading,
            terminate_on_apogee = False,
            initial_solution = flight_payload_vals["all_flights"],
            name          = "Payload no chute"
        )

        attach_meta(flight_payload, {
            k: flight_payload_vals[k]
            for k in variables
            if k in flight_payload_vals
        })
        attach_meta(flight_payload_no_chute, {
            k: flight_payload_vals[k]
            for k in variables
            if k in flight_payload_vals
        })
        flights_payload.append(flight_payload) 
        flights_payload_no_chute.append(flight_payload_no_chute)     
        if i % step == 0 or i == total:
            print(f"{i / total:.0%}")
    constants, variables = register("flight_payload", flights_payload, constants, variables)
    constants, variables = register("flight_payload_no_chute", flights_payload_no_chute, constants, variables)
    return constants, variables
