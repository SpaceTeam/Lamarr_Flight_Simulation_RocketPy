import simulation.utils as utils

def calculate_safe_flights(constants, variables, zones):
    flights_payload = utils.ensure_list(utils.lookup("flight_payload" , constants, variables)[0])
    flights_w_p = utils.ensure_list(utils.lookup("flight_without_payload", constants, variables)[0])
    flights_w_p_no_main = utils.ensure_list(utils.lookup("flight_without_payload_no_main", constants, variables)[0])
    flights_w_p_ballistic = utils.ensure_list(utils.lookup("flight_without_payload_ballistic", constants, variables)[0])
    all_headings = utils.ensure_list(utils.lookup("flight_heading", constants, variables)[0])




    # ===============================
    # Prepare coordinates as tuples
    # ===============================

    all_payload_scenarios = flights_payload
    all_rocket_scenarios = flights_w_p + flights_w_p_no_main + flights_w_p_ballistic

    payload_coords = [(f.x_impact, f.y_impact) for f in flights_payload]
    rocket_coords  = [(f.x_impact, f.y_impact) for f in flights_w_p]
    rocket_coords_no_main  = [(f.x_impact, f.y_impact) for f in flights_w_p_no_main]
    rocket_coords_ballistic  = [(f.x_impact, f.y_impact) for f in flights_w_p_ballistic]

    # ===============================
    # Check which impacts are in exclusion zones
    # ===============================
    payload_impacts = utils.is_in_exclusion_zone(payload_coords, zones)
    rocket_impacts  = utils.is_in_exclusion_zone(rocket_coords, zones)
    no_main_impacts  = utils.is_in_exclusion_zone(rocket_coords_no_main, zones)
    ballistic_impacts  = utils.is_in_exclusion_zone(rocket_coords_ballistic, zones)

    # ===============================
    # Find headings in exclusion zones
    # ===============================
    payload_headings_in_zone = list({flights_payload[i].heading 
                            for i, in_zone in enumerate(payload_impacts) if in_zone})
    rocket_headings_in_zone = list({flights_w_p[i].heading 
                            for i, in_zone in enumerate(rocket_impacts) if in_zone})
    no_main_headings_in_zone = list({flights_w_p_no_main[i].heading 
                            for i, in_zone in enumerate(no_main_impacts) if in_zone})
    ballistic_headings_in_zone = list({flights_w_p_ballistic[i].heading 
                            for i, in_zone in enumerate(ballistic_impacts) if in_zone})

    #print("Headings with payload in exclusion zones:", payload_headings_in_zone)
    #print("Headings with rocket in exclusion zones:", rocket_headings_in_zone)
    #print("Headings with rocket no main in exclusion zones:", no_main_headings_in_zone)
    #print("Headings with ballistic rocket in exclusion zones:", ballistic_headings_in_zone)

    # ===============================
    # Compute safe headings
    # ===============================
    safe_payload_headings = [
        h for h in all_headings
        if h not in payload_headings_in_zone
    ]

    safe_rocket_headings = [
        h for h in all_headings
        if h not in rocket_headings_in_zone
        and h not in no_main_headings_in_zone
        and h not in ballistic_headings_in_zone
    ]

    # ===============================
    # Filter safe flights by heading
    # ===============================
    safe_payload_nominal = [
        f for f in flights_payload
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

    # ===============================
    # Find initial solutions that are safe for BOTH
    # ===============================
    payload_solutions = (
        [f.initial_solution for f in safe_payload_nominal]
    )

    rocket_solutions = (
        [f.initial_solution for f in safe_rocket_nominal] +
        [f.initial_solution for f in safe_rocket_no_main] +
        [f.initial_solution for f in safe_rocket_ballistic]
    )

    safe_solutions = [s for s in payload_solutions if s in rocket_solutions]

    # ===============================
    # Keep only flights belonging to safe solutions
    # ===============================
    safe_payload_nominal = [
        f for f in safe_payload_nominal
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

    safe_configurations = list({
        (f.heading, f.inclination)
        for f in safe_rocket_nominal
    })


    # ===============================
    # Print headings not in exclusion zones (every 20°)
    # ===============================
    #for i in all_headings:
    #    if (i not in payload_headings_in_zone) and (i not in rocket_headings_in_zone + ballistic_headings_in_zone +  no_main_headings_in_zone):
    #        utils.printmd(f"**Safe heading: {i}**")

    print("## **Safe Inclinations per heading**")

    grouped = {}

    for f in safe_payload_nominal + safe_rocket_ballistic + safe_rocket_nominal:
        grouped.setdefault(f.heading, set()).add(f.inclination)

    for heading, inclinations in grouped.items():
        print(f"**Heading {heading}: {sorted(inclinations)}**")
    
    constants, variables = utils.register("safe_payload_nominal", safe_payload_nominal, constants, variables)
    constants, variables = utils.register("safe_rocket_nominal", safe_rocket_nominal, constants, variables)
    constants, variables = utils.register("safe_rocket_no_main", safe_rocket_no_main, constants, variables)
    constants, variables = utils.register("safe_rocket_ballistic", safe_rocket_ballistic, constants, variables)
    constants, variables = utils.register("safe_configurations", safe_configurations, constants, variables)
    
    return constants, variables
