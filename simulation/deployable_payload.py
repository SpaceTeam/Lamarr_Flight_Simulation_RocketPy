from rocketpy import Rocket, Flight, Parachute

from simulation.utils import *

DEBUG = False


# =============================================================================
# Payload-as-rocket
# =============================================================================
#
# Payload flights are created automatically by the main simulation pipeline when
# ``payload_mass_total`` is greater than zero.


def has_deployable_payload(constants, variations):
    """
    Return True when the loaded config declares a separated payload mass.
    """
    try:
        payload_mass_total = lookup("payload_mass_total", constants, variations)[0]
    except KeyError:
        return False

    payload_masses = ensure_list(payload_mass_total)
    return any(payload_mass > 0 for payload_mass in payload_masses)


def collect_nominal_ascent_flights(constants, variations):
    """
    Return full-mass ascent flights for payload deployment, falling back to nominal scenario flights.
    """
    try:
        ascent_flights_by_env = lookup("ascent_flights_by_env", constants, variations)[0]
    except KeyError:
        ascent_flights_by_env = None

    if ascent_flights_by_env is not None:
        # Payload deployment uses the full-mass ascent that stops at apogee.
        return [
            ascent_flight
            for ascent_flights in ascent_flights_by_env.values()
            for ascent_flight in ensure_list(ascent_flights)
        ]

    try:
        flights_by_env = lookup("flights_by_env", constants, variations)[0]
    except KeyError:
        return []

    # Without separate ascent flights, fall back to the nominal scenario flight.
    return [
        scenario_set["nominal"]
        for scenario_sets in flights_by_env.values()
        for scenario_set in scenario_sets
    ]


def create_payload_parachute(constants, variations):
    """
    Create the parachute used by the separated payload.
    """
    print("Creating payload parachute...")
    parameter_payload_parachute = []

    required = [
        "cd_s",
        "trigger",
        "sampling_rate",
    ]

    # defaults = {
    #     "sampling_rate": 100,        # hz              # preset
    #     "lag": 1,                    # s               # measured
    # }
    
    fill_parameters(parameter_payload_parachute, "parachutes_payload_", constants, variations)
    check_required(parameter_payload_parachute, required, "parachutes_payload")
    # add_defaults(parameter_payload_parachute, defaults, constants, variations, "parachute_payload")
    if DEBUG:
        print(parameter_payload_parachute)

    payload_parachutes = []

    for parachute_vals in generate_combinations(parameter_payload_parachute, constants, variations):
        parachute_list = {}
        parachute_options = {
            "name": "payload",
            "cd_s": parachute_vals["parachutes_payload_cd_s"],
            "trigger": parachute_vals["parachutes_payload_trigger"],
            "sampling_rate": parachute_vals["parachutes_payload_sampling_rate"],
        }

        if "parachutes_payload_lag" in parachute_vals:
            parachute_options["lag"] = parachute_vals["parachutes_payload_lag"]

        if "parachutes_payload_noise" in parachute_vals:
            parachute_options["noise"] = parachute_vals["parachutes_payload_noise"]

        parachute_list[0] = Parachute(**parachute_options)
        attach_meta(parachute_list[0], {
            k: parachute_vals[k]
            for k in variations
            if k in parachute_vals
        })
        payload_parachutes.append(parachute_list)
    return register("payload_parachute", payload_parachutes, constants, variations)


def create_payload(constants, variations):
    """
    Create the deployed-payload "rocket".
    """
    print("Creating payload...")
    project = constants["project"]
    constants, variations = create_payload_parachute(constants, variations)
    parameter_payload = []

    required = [
        "diameter",
        "mass",
        "length"
    ]
    
    # defaults={
    #     "payload_moment_of_intertia_XY":0.01,
    #     "payload_moment_of_intertia_Z":0.01
    # }
    fill_parameters(parameter_payload, "payload_", constants, variations)
    check_required(parameter_payload, required, "payload")
    # add_defaults(parameter_payload, defaults, constants, variations)

    if DEBUG:
        print(parameter_payload)

    payloads = []

    for payload_vals in generate_combinations(parameter_payload, constants, variations):
        payload = {}
        payload["nominal"] = Rocket(
            radius=payload_vals["payload_diameter"] / 2 / 1000,
            mass=payload_vals["payload_mass"] / 1000,
            inertia=(payload_vals["payload_moment_of_intertia_XY"], payload_vals["payload_moment_of_intertia_XY"], payload_vals["payload_moment_of_intertia_Z"]),
            power_off_drag="./" + project + "/power_off_drag.csv",
            power_on_drag="./" + project + "/power_off_drag.csv",
            center_of_mass_without_motor=payload_vals["payload_length"] / 2 / 1000,
            coordinate_system_orientation="tail_to_nose"
        )

        payload["nominal"].parachutes = list(payload_vals["payload_parachute"].values())

        attach_meta(payload["nominal"], {
            k: payload_vals[k]
            for k in variations
            if k in payload_vals
        })

        payloads.append(payload)
    return register("payload", payloads, constants, variations)


def create_payload_flight(constants, variations):
    """
    Simulate the separated payload from the apogee of each nominal ascent flight.
    """
    print("Creating payload flight...")

    payloads = ensure_list(lookup("payload", constants, variations)[0])
    nominal_ascent_flights = collect_nominal_ascent_flights(constants, variations)

    if not nominal_ascent_flights:
        raise ValueError("No nominal ascent flights are available for payload deployment.")

    flights_payload = []
    total = len(payloads) * len(nominal_ascent_flights)
    step = max(1, total // 10)
    print(total)
    i = 0

    for payload in payloads:
        for nominal_flight in nominal_ascent_flights:
            i += 1

            # Start the separated payload at the nominal rocket apogee instead of at impact.
            initial_solution = [nominal_flight.apogee_time, *nominal_flight.apogee_state]

            flight_payload = Flight(
                rocket=payload["nominal"],
                environment=nominal_flight.env,
                rail_length=nominal_flight.rail_length,
                inclination=nominal_flight.inclination,
                heading=nominal_flight.heading,
                terminate_on_apogee=False,
                initial_solution=initial_solution,
                name="Payload",
            )

            attach_meta(flight_payload, {})
            flights_payload.append(flight_payload)

            if i % step == 0 or i == total:
                print(f"{i / total:.0%}")

    constants, variations = register("flight_payload", flights_payload, constants, variations)
    return constants, variations
