from simulation.utils import *
from simulation.simulation import *
from simulation.deployable_payload import *

project = "LAMARR"
constants, variables = parse_config(project+"/config.txt")
register("project", project, constants, variables)

all_objects = list(generate_objects(constants, variables))

#for obj in all_objects:
#    print(obj)

print("VARIABLES")
for var in variables:
    print(var)

print()
print("///////////////////////////////////////")
print()
print("CONSTANTS")
for const in constants:
    print(const)

constants, variables = create_environment(constants, variables)
constants, variables = create_engine(constants, variables)
constants, variables = create_rocket(constants, variables)
constants, variables = create_flight(constants, variables)

if project == "CANSAT":
    constants, variables = create_rocket_without_payload(constants, variables)
    constants, variables = create_flight_without_payload(constants, variables)
    constants, variables = create_payload(constants, variables)
    constants, variables = create_payload_flight(constants, variables)
