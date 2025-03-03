###
#   Matthias Ogris on 06.02.2025 // TU Wien Space Team
#
#   The config generator only works if all the files are accessed locally i.e. the .ods file must be downloaded prior to script execution.
#
#   This script converts a user friendly .ods spreadsheet containing Hedy simulation parameters into a .json config file which is used by the simulation script.
#   Parallel to just rearranging data, units are converted into the SI system to ensure the simulation focuses on simulating and not preparing data.
#
#   Bugs/Fixes: conversion needs to be rounded due to minor python division bug; tuples being passed could also be stored as lists as nothing would likely change,
#   add unit conversion feature when passing tuples
###
import pandas as pd
import json
import sys

file_paths = {}

try:
    if sys.argv[1] == "-h" or sys.argv[1] == "--help":
        print(f"TU Wien Space Team simulation config generator\n\n\
        Command line instruction: \'python {sys.argv[0]} [config_sheet.ods path] [config_export.json path]\'\n\
        (!) Warning: A similarly named file at the export path will be overwritten.")
        sys.exit()
    else:
        file_paths["import"] = sys.argv[1]
        file_paths["export"] = sys.argv[2]

except SystemExit:
    print("Script terminated")
    sys.exit()
except:
    print(f"This script must be executed with arguments, try \'python {sys.argv[0]} -h\' for help.")
    sys.exit()

column_titles = {
    "category": "Category",
    "name": "Name",
    "value": "Value",
    "unit": "Unit",
    "comment": "Comment"
}

def convert_units(data: pd.Series):
    # this statement is "Pfusch" but it seems to work rather elegantly
    # this else statement is a very java approach to solving the tuple conversion but it should work
    # ATTENTION: so far no unit conversion when parsing a tuple
    if "," in str(data.get(column_titles["value"])):
        data_tuple_raw = str(data.get(column_titles["value"])).split(",")

        for pos in range(0, len(data_tuple_raw)):
            data_tuple_raw[pos] = float(data_tuple_raw[pos])

        data.at[column_titles["value"]] = tuple(data_tuple_raw)
        
        return data
    
    else:
        unit = data.get(column_titles["unit"])
        value = data.get(column_titles["value"])

        match unit:
            case "mm":
                value = float(value)
                data.at[column_titles["value"]] = float(round(value / 1000, 5))
                data.at[column_titles["unit"]] = "m"
                return data
            case "g":
                value = float(value)
                data.at[column_titles["value"]] = float(value / 1000)
                data.at[column_titles["unit"]] = "kg"
                return data
            case "bar":
                value = float(value)
                data.at[column_titles["value"]] = float(value * 100000)
                data.at[column_titles["unit"]] = "Pa"
                return data
            case "kN":
                value = float(value)
                data.at[column_titles["value"]] = float(value * 1000)
                data.at[column_titles["unit"]] = "N"
                return data
            case "°C":
                value = float(value)
                data.at[column_titles["value"]] = float(value + 273.15)
                data.at[column_titles["unit"]] = "K"
                return data
            case "text":
                data.at[column_titles["value"]] = str(value)
                return data
            case _:
                data.at[column_titles["value"]] = float(value)
                return data


xl = pd.ExcelFile(file_paths["import"])

config_dict = {}

for sheet in xl.sheet_names:
    # 1st layer: convert each spreadsheet to dataframe, drop comment column and convert to SI units
    df = pd.DataFrame(xl.parse(sheet_name=sheet)\
        .drop(columns=column_titles["comment"])\
        .apply(convert_units, axis=1))
    
    config_dict[sheet] = {}
    
    # 2nd layer: extract unique category labels for later use in nested dictionary
    categories = pd.DataFrame(df).get(column_titles["category"]).to_list()
    category_keys = []

    for key in categories:
        if not key in category_keys:
            category_keys.append(key)
            config_dict[sheet][key] = {}

    # 3rd layer: assign names to previously installed dictionary categories, iterate over every row of each sheet's dataframe
    rows, columns = pd.DataFrame(df).shape

    for pos in range(0, rows):
        entry_row = pd.DataFrame(df).loc[pos]

        config_dict[sheet][entry_row[column_titles["category"]]][entry_row[column_titles["name"]]] = {
            "Value": entry_row[column_titles["value"]],
            "Unit": str(entry_row[column_titles["unit"]])}

json_formatted = json.dumps(config_dict, indent=4)

with open(file_paths["export"], "wt") as f:
    f.write(json_formatted)


