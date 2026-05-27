import math

from rocketpy import Rocket, Flight, Parachute

from simulation.utils import *
from simulation.config_schema import SimParams, ParachuteConfig
DEBUG = False


# =============================================================================
# Payload-as-rocket
# =============================================================================
#
# Payload flights are created automatically by the main simulation pipeline when
# payload.mass_total is greater than zero.


def has_deployable_payload(params: SimParams):
    """Return True when the loaded config declares a separated payload mass."""
    mass_total = params.config.payload.mass_total
    if isinstance(mass_total, list):
        return any(m > 0 for m in mass_total)
    return mass_total > 0


def _collect_nominal_ascent_flights(params: SimParams):
    """Return the full-mass ascent-only flights (or nominal scenario flights as fallback) for payload deployment."""
    if params.runtime.ascent_flights_by_env is not None:
        # Payload deployment uses the full-mass ascent that stops at apogee.
        return [
            flight
            for flights in params.runtime.ascent_flights_by_env.values()
            for flight in ensure_list(flights)
        ]

    if params.runtime.flights_by_env is not None:
        # Without separate ascent flights, fall back to the nominal scenario flight.
        return [
            scenario_set["nominal"]
            for scenario_sets in params.runtime.flights_by_env.values()
            for scenario_set in scenario_sets
        ]

    return []


def _get_parachute_cd_s(payload_cfg: ParachuteConfig):
    """Compute cd_s from a ParachuteConfig: directly, or from cd + radius/fabric_area."""
    if payload_cfg.cd_s is not None:
        return payload_cfg.cd_s
    if payload_cfg.cd is not None and payload_cfg.fabric_area is not None:
        return payload_cfg.cd * payload_cfg.fabric_area
    if payload_cfg.cd is not None and payload_cfg.radius is not None:
        return payload_cfg.cd * math.pi * payload_cfg.radius ** 2
    raise ValueError("Payload parachute needs either cd_s, cd + radius, or cd + fabric_area in config.")


def create_payload(params: SimParams):
    """Create the payload parachute and payload Rocket; store in params.runtime.payload."""
    print("Creating payload...")

    payload_cfg = params.config.parachutes.payload
    if payload_cfg is None:
        raise ValueError("parachutes.payload config is required for deployable payload simulation.")

    pl = params.config.payload
    required = {"mass", "diameter", "length"}
    missing = [f for f in required if getattr(pl, f, None) is None]
    if missing:
        raise ValueError(f"payload config is missing required fields: {sorted(missing)}")

    # defaults = {
    #     "sampling_rate": 100,        # hz              # preset
    #     "lag": 1,                    # s               # measured
    # }
    
    options = {
        "name": "payload",
        "cd_s": _get_parachute_cd_s(payload_cfg),
        "trigger": payload_cfg.trigger,
        "sampling_rate": payload_cfg.sampling_rate,
    }
    if payload_cfg.lag is not None:
        options["lag"] = payload_cfg.lag
    if payload_cfg.noise is not None:
        options["noise"] = payload_cfg.noise
    if payload_cfg.radius is not None:
        options["radius"] = payload_cfg.radius
    if payload_cfg.cd is not None:
        options["drag_coefficient"] = payload_cfg.cd
    parachute = Parachute(**options)

    set_labels(parachute)

    # defaults={
    #     "payload_moment_of_intertia_XY":0.01,
    #     "payload_moment_of_intertia_Z":0.01
    # }
    
    inertia_xy = pl.moment_of_intertia_XY
    inertia_z = pl.moment_of_intertia_Z

    power_off_drag = get_project_file(params, params.config.rocket.power_off_drag)

    payload_rocket = Rocket(
        radius=pl.diameter / 2 / 1000,
        mass=pl.mass / 1000,
        inertia=(inertia_xy, inertia_xy, inertia_z),
        power_off_drag=power_off_drag,
        power_on_drag=power_off_drag,
        center_of_mass_without_motor=pl.length / 2 / 1000,
        coordinate_system_orientation="tail_to_nose",
    )
    payload_rocket.parachutes = [parachute]
    set_labels(payload_rocket)

    params.runtime.payload = payload_rocket
    params.runtime.payload_parachute = {0: parachute}

    if DEBUG:
        print(f"Payload rocket: radius={pl.diameter/2/1000:.4f} m, mass={pl.mass/1000:.3f} kg")


def create_payload_flight(params: SimParams):
    """Simulate the separated payload from the apogee of each nominal ascent flight."""
    print("Creating payload flight...")

    payload_rocket = params.runtime.payload
    if payload_rocket is None:
        raise ValueError("No payload has been created yet. Call create_payload first.")

    nominal_ascent_flights = _collect_nominal_ascent_flights(params)
    if not nominal_ascent_flights:
        raise ValueError("No nominal ascent flights are available for payload deployment.")

    flights_payload = []
    total = len(nominal_ascent_flights)

    print(total)

    for i, nominal_flight in enumerate(nominal_ascent_flights, start=1):
        print(f"{i / total:.0%}")
        initial_solution = [nominal_flight.apogee_time, *nominal_flight.apogee_state]

        flight_payload = Flight(
            rocket=payload_rocket,
            environment=nominal_flight.env,
            rail_length=nominal_flight.rail_length,
            inclination=nominal_flight.inclination,
            heading=nominal_flight.heading,
            terminate_on_apogee=False,
            initial_solution=initial_solution,
            name="Payload",
        )
        set_labels(flight_payload)
        flights_payload.append(flight_payload)


    params.runtime.flight_payload = flights_payload
