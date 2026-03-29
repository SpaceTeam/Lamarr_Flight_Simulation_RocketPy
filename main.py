import simulation.simulation as sim
import simulation.deployable_payload as dp
import simulation.utils as utils
import simulation.outputs as outputs
import sys
import json

project = sys.argv[1]
zone_multiplicator = 1.5
DEBUG = True

constants, variables = utils.parse_config(project+"/config.txt")
constants, variables = utils.register("project", project, constants, variables)

if DEBUG:
    print("# **VARIABLES**")
    for var in variables:
        print(var)

    print()
    print()
    print("# **CONSTANTS**")
    for const in constants:
        print(const)


# Define polygons (x = WO, y = NS)
exclusion_zones,buffer_zones = utils.load_zones(project)

constants, variables = utils.register("exclusion_zones", exclusion_zones, constants, variables)
constants, variables = utils.register("buffer_zones", buffer_zones, constants, variables)
constants, variables = utils.register("zones", buffer_zones, constants, variables)


buffer_zones.update(utils.scale_zones(exclusion_zones, zone_multiplicator))


print("[1/8]")
constants, variables = sim.create_environment(constants, variables)
print("[2/8]")
constants, variables = sim.create_engine(constants, variables)
print("[3/8]")
constants, variables = sim.create_rocket(constants, variables)
print("[4/8]")
constants, variables = sim.create_flight(constants, variables)

if project == "CANSAT":
    print("[5/8]")
    constants, variables = dp.create_rocket_without_payload(constants, variables)
    print("[6/8]")
    constants, variables = dp.create_flight_without_payload(constants, variables)
    print("[7/8]")
    constants, variables = dp.create_payload(constants, variables)
    print("[8/8]")
    constants, variables = dp.create_payload_flight(constants, variables)




if project == "CANSAT":
    utils.plot_safe_flights(project, exclusion_zones, buffer_zones, flight_groups={
        "rocket_nominal": (utils.lookup("flight_without_payload", constants, variables)[0], "green"),
        "payload_nominal": (utils.lookup("flight_payload", constants, variables)[0], "blue")
    }, plot_name = "all_flights")

    outputs.calculate_safe_flights(constants, variables, buffer_zones)

    # ===============================
    # Plot safe flights with zones
    # ===============================
    utils.plot_safe_flights(project, exclusion_zones, buffer_zones, flight_groups={
        "rocket_nominal": (utils.lookup("safe_rocket_nominal", constants, variables)[0], "green"),
        "rocket_no_main": (utils.lookup("safe_rocket_no_main", constants, variables)[0], "orange"),
        "rocket_ballistic": (utils.lookup("safe_rocket_ballistic", constants, variables)[0], "red"),
        "payload_nominal": (utils.lookup("safe_payload_nominal", constants, variables)[0], "blue"),
        }
    )

    result = {"constants": constants, "variables":variables}
    print("RESULT_JSON:" + json.dumps(result))