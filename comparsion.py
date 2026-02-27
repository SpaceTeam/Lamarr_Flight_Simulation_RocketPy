from rocketpy import CompareFlights

from simulation.simulation import *
from simulation.deployable_payload import *


def draw_initial_solutions(constants, variables):
        # TODO: add in/output, so user does not have to open the code
    legend = False
    all_flights = lookup("all_flights", constants, variables)[0]

    #for flight in all_flights:
    #    print(flight.name)
        # print(flight._meta)
    comparison_normal = CompareFlights(all_flights)
    comparison_normal.trajectories_3d(legend=legend, filename = "3d.png")
    comparison_normal.trajectories_2d(legend=legend, filename = "2d_xy.png", plane = "xy")
    comparison_normal.trajectories_2d(legend=legend, filename = "2d_xz.png", plane = "xz")
    comparison_normal.trajectories_2d(legend=legend, filename = "2d_yz.png", plane = "yz")




#=============================================================================
project = "CANSAT"
constants, variables = parse_config(project+"/config.txt")
register("project", project, constants, variables)
constants, variables = create_environment(constants, variables)
constants, variables = create_engine(constants, variables)
constants, variables = create_rocket(constants, variables)
constants, variables = create_flight(constants, variables)

if project == "CANSAT":
    constants, variables = create_rocket_without_payload(constants, variables)
    constants, variables = create_flight_without_payload(constants, variables)
    constants, variables = create_payload(constants, variables)
    constants, variables = create_payload_flight(constants, variables)
#=============================================================================




# Define polygons (x = WO, y = NS)
exclusion_zones = {
    "Exclusion Zone Flughafen": [
        (-110, 450),
        (-310, 550),
        (-610, 350),
        (-600, 10),
        (-400, -10),
        (0,0)   # enable to exclude launches heading to tribune
    ],
    "Exclusion Zone Niederham": [
        (-1100, -470),
        (-220, -210),
        (280, -890)
    ],
    "Rehwinkl": [
        (260, -220),
        (750, -150),
        (700, -440),
        (280, -440)
    ]
}
buffer_zones = {
    "Buffer Zone Wald": [
        (-110, 450),
        (200, -110),
        (630, 280)
    ]
}

dummy_zones = {
    "dummy": [
        (-750,400),
        (250,-200),
        (400,400)
    ]
}

buffer_zones.update(scale_zones(exclusion_zones, 1.1))
plot_all(exclusion_zones, buffer_zones, flight_groups={
    "rocket_nominal": (lookup("flight_without_payload", constants, variables)[0], "green"),
    "payload_nominal": (lookup("flight_payload", constants, variables)[0], "blue")
})


flights_payload = lookup("flight_payload" , constants, variables)[0]
flights_payload_no_chute = lookup("flight_payload_no_chute", constants, variables)[0]
flights_w_p = lookup("flight_without_payload", constants, variables)[0]
flights_w_p_no_main = lookup("flight_without_payload_no_main", constants, variables)[0]
flights_w_p_ballistic = lookup("flight_without_payload_ballistic", constants, variables)[0]

zones = buffer_zones
# ===============================
# Prepare coordinates as tuples
# ===============================

all_payload_scenarios = flights_payload + flights_payload_no_chute
all_rocket_scenarios = flights_w_p + flights_w_p_no_main + flights_w_p_ballistic

payload_coords = [(f.x_impact, f.y_impact) for f in flights_payload]
payload_coords_no_chute = [(f.x_impact, f.y_impact) for f in flights_payload_no_chute]
rocket_coords  = [(f.x_impact, f.y_impact) for f in flights_w_p]
rocket_coords_no_main  = [(f.x_impact, f.y_impact) for f in flights_w_p_no_main]
rocket_coords_ballistic  = [(f.x_impact, f.y_impact) for f in flights_w_p_ballistic]

# ===============================
# Check which impacts are in exclusion zones
# ===============================
payload_impacts = is_in_exclusion_zone(payload_coords, zones)
no_chute_impacts = is_in_exclusion_zone(payload_coords_no_chute, zones)
rocket_impacts  = is_in_exclusion_zone(rocket_coords, zones)
no_main_impacts  = is_in_exclusion_zone(rocket_coords_no_main, zones)
ballistic_impacts  = is_in_exclusion_zone(rocket_coords_ballistic, zones)

# ===============================
# Find headings of payloads in exclusion zones
# ===============================
payload_headings_in_zone = list({flights_payload[i].heading 
                        for i, in_zone in enumerate(payload_impacts) if in_zone})
no_chute_headings_in_zone = list({flights_payload_no_chute[i].heading 
                        for i, in_zone in enumerate(no_chute_impacts) if in_zone})

# ===============================
# Find headings of rockets in exclusion zones
# ===============================
rocket_headings_in_zone = list({flights_w_p[i].heading 
                        for i, in_zone in enumerate(rocket_impacts) if in_zone})
no_main_headings_in_zone = list({flights_w_p_no_main[i].heading 
                        for i, in_zone in enumerate(no_main_impacts) if in_zone})
ballistic_headings_in_zone = list({flights_w_p_ballistic[i].heading 
                        for i, in_zone in enumerate(ballistic_impacts) if in_zone})

#print("Headings with payload in exclusion zones:", payload_headings_in_zone)
#print("Headings with payload no chute in exclusion zones:", no_chute_headings_in_zone)
#print("Headings with rocket in exclusion zones:", rocket_headings_in_zone)
#print("Headings with rocket no main in exclusion zones:", no_main_headings_in_zone)
#print("Headings with ballistic rocket in exclusion zones:", ballistic_headings_in_zone)


# ===============================
# Print headings not in exclusion zones (every 20°)
# ===============================
for i in range(0, 360, 10):
    if (i not in payload_headings_in_zone + no_chute_headings_in_zone) and (i not in rocket_headings_in_zone + ballistic_headings_in_zone +  no_main_headings_in_zone):
        print("Safe heading:", i)

# -------------------------------
# 1) Compute forbidden headings
# -------------------------------
safe_payload_headings = [
    h for h in range(0, 360, 10)
    if h not in payload_headings_in_zone
    and h not in no_chute_headings_in_zone
]

safe_rocket_headings = [
    h for h in range(0, 360, 10)
    if h not in rocket_headings_in_zone
    and h not in no_main_headings_in_zone
    and h not in ballistic_headings_in_zone
]
# -------------------------------
# 2) Filter safe flights by heading
# -------------------------------
safe_payload_nominal = [
    f for f in flights_payload
    if f.heading in safe_payload_headings
]

safe_payload_no_chute = [
    f for f in flights_payload_no_chute
    if f.heading in safe_payload_headings
]

safe_rocket_nominal = [
    f for f in flights_w_p
    if f.heading in safe_rocket_headings
]

safe_rocket_no_main = [
    f for f in flights_w_p_no_main
    if f.heading in safe_rocket_headings
]

safe_rocket_ballistic = [
    f for f in flights_w_p_ballistic
    if f.heading in safe_rocket_headings
]

# -------------------------------
# 3) Find initial solutions that are safe for BOTH
# -------------------------------
payload_solutions = (
    [f.initial_solution for f in safe_payload_nominal] +
    [f.initial_solution for f in safe_payload_no_chute]
)

rocket_solutions = (
    [f.initial_solution for f in safe_rocket_nominal] +
    [f.initial_solution for f in safe_rocket_no_main] +
    [f.initial_solution for f in safe_rocket_ballistic]
)

safe_solutions = [s for s in payload_solutions if s in rocket_solutions]

# -------------------------------
# 4) Keep only flights belonging to safe solutions
# -------------------------------
safe_payload_nominal = [
    f for f in safe_payload_nominal
    if f.initial_solution in safe_solutions
]

safe_payload_no_chute = [
    f for f in safe_payload_no_chute
    if f.initial_solution in safe_solutions
]

safe_rocket_nominal = [
    f for f in safe_rocket_nominal
    if f.initial_solution in safe_solutions
]

safe_rocket_no_main = [
    f for f in safe_rocket_no_main
    if f.initial_solution in safe_solutions
]

safe_rocket_ballistic = [
    f for f in safe_rocket_ballistic
    if f.initial_solution in safe_solutions
]


# ===============================
# Plot safe flights with zones
# ===============================
plot_all(exclusion_zones, buffer_zones, flight_groups={
    "rocket_nominal": (safe_rocket_nominal, "green"),
    "rocket_no_main": (safe_rocket_no_main, "orange"),
    "rocket_ballistic": (safe_rocket_ballistic, "red"),
    "payload_nominal": (safe_payload_nominal, "blue"),
    "payload_no_chute": (safe_payload_no_chute, "yellow"),
    }
)