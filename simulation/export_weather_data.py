"""
Download a vertical atmosphere profile and export it for RocketPy and OpenRocket.

Supported sources:
    - Open-Meteo DWD ICON API:
        icon_d2
        icon_eu
        icon_global

Docs:
    - Open-Meteo DWD ICON API:
      https://open-meteo.com/en/docs/dwd-api

    - RocketPy:
      https://docs.rocketpy.org/en/latest/user/environment/1-atm-models/custom_atmosphere.html
      https://docs.rocketpy.org/en/latest/user/environment/3-further/data_csv.html

    - OpenRocket:
      Docu is in its GUI:
      edit simulation -> Wind section -> Multi-level -> Edit levels -> Import levels

Install:
    pip install openmeteo-requests requests-cache retry-requests pandas

Outputs:
    rocketpy_atmosphere_<weather_model>.csv
    openrocket_wind_profile_<weather_model>.csv

RocketPy output columns:
    height, pressure, temperature, wind_u, wind_v

OpenRocket output columns:
    altitude, wind speed, meteorological wind direction, stddev

Units:
    - height / altitude in the output files is above sea level (ASL).
    - pressure is in Pa.
    - temperature is in K.
    - wind_u: eastward component of physical velocity wind vector (positive eastward)
    - wind_v: northward component of physical velocity wind vector (positive northward)
    - wind_speed: magnitude of the wind vector in m/s.
    - Wind direction is meteorological direction in degrees: direction from which the wind is blowing.
"""

import math
import pandas as pd
from pathlib import Path

import openmeteo_requests
import requests_cache
from retry_requests import retry


# ===========================================================================
# User Settings for standalone script execution
# ===========================================================================

# launch site coordinates
latitude = 47.21415
longitude = 9.00425

# m (AGL), only export data up until slightly above this height, if set to None, all heights are exported
max_expected_height_agl_m = 700             

# Select the local launch time of the launch site.
# Used directly for Open-Meteo.
launch_time = "2026-05-14T12:00"
timezone = "Europe/Berlin"

# Available model choices:
# - "icon_d2"       Open-Meteo DWD ICON-D2
# - "icon_eu"       Open-Meteo DWD ICON-EU
# - "icon_global"   Open-Meteo DWD ICON global
WEATHER_MODEL = "icon_d2"


# ===========================================================================
# Constants
# ===========================================================================

NEAR_GROUND_HEIGHTS_M = [0, 80, 120, 180]

# Pressure levels supported by Open-Meteo DWD ICON API.
# https://open-meteo.com/en/docs/dwd-api#pressure_level_variables
PRESSURE_LEVELS_HPA = [
    1000,   # ~110 m (corresponding approx. altitude (ASL))
    975,    # ~320 m
    950,    # ~500 m
    925,    # ~800 m
    900,    # ~1000 m
    850,    # ~1500 m
    800,    # ~1900 m
    700,    # ~3 km
    600,    # ~4.2 km
    500,    # ~5.6 km
    400,    # ~7.2 km
    300,    # ~9.2 km
    250,    # ~10.4 km
    200,    # ~11.8 km
    150,    # ~13.5 km
    100,    # ~15.8 km
    70,     # ~17.7 km
    50,     # ~19.3 km
    30,     # ~22 km
]     

OPENMETEO_MODELS = [
    "icon_d2",
    "icon_eu",
    "icon_global",
]

# RocketPy assumes pressure to be strictly monotonically decreasing with height an throws an error if it doesn't.
# -> skip pressure-level rows that are very close to near-ground rows.
# see self.barometric_height = self.pressure.inverse_function() in environment.py of RocketPy source code
# they use: https://proofwiki.org/wiki/Strictly_Monotone_Real_Function_is_Bijective
MIN_HEIGHT_GAP_FROM_NEAR_GROUND_M = 5


# ===========================================================================
# Conversions
# ===========================================================================

def wind_speed_direction_to_uv(wind_speed, wind_direction_degrees):
    """
    Convert meteorological wind speed and direction to east/north wind components.

    Meteorological wind direction = direction from which the wind is blowing:
    - 0° = wind coming from north, blowing south
    - 90° = wind coming from east, blowing west

    Meteorological wind direction points to where the wind comes from.
    See https://confluence.ecmwf.int/plugins/viewsource/viewpagesrc.action?pageId=133262398
        https://metview.readthedocs.io/en/5.25.0/api/functions/direction.html
        
    
    The physical wind velocity vector points in the opposite direction.
    wind_speed = magnitude of the wind vector
    
    RocketPy components:
    - wind_u: eastward component of physical velocity wind vector (positive eastward)
    - wind_v: northward component of physical velocity wind vector (positive northward)
    """
    wind_direction_radians = math.radians(wind_direction_degrees)

    wind_u = -wind_speed * math.sin(wind_direction_radians)
    wind_v = -wind_speed * math.cos(wind_direction_radians)

    return wind_u, wind_v


def compute_pressure_with_temperature_gradient(
    surface_pressure_pa,
    surface_temperature_k,
    target_temperature_k,
    near_ground_height_m,
):
    """
    Compute pressure at z ∈ NEAR_GROUND_HEIGHTS_M using the barometric formula with a linear temperature gradient:
        P(z) = Pb * (Tb / T(z)) ** (g * M / (R * L))
    see https://en.wikipedia.org/wiki/Barometric_formula#Model_equations

    where:
    - Pb = reference pressure (here surface pressure) [Pa]
    - Tb = reference temperature (here surface temperature) [K]
    - T(z) = temperature at target height z [K]
    - L = temperature gradient [K/m]
    - g = gravitational acceleration [m/s²]
    - M = mean molar mass of air at sea level [kg/kmol]
    - R = universal gas constant [J/(kmol*K)]

    If the temperature gradient is extremely small, the function falls back to the isothermal
    barometric equation to avoid division by a value close to zero. Also found in above link.
        P(z) = Pb * exp((-g * M * near_ground_height_m) / (R * Tb))
    """
    g = 9.80665
    M = 28.9644
    R = 8314.46261815324

    if near_ground_height_m == 0:
        return surface_pressure_pa

    temperature_gradient = (target_temperature_k - surface_temperature_k) / near_ground_height_m

    if abs(temperature_gradient) < 1e-9:
        return surface_pressure_pa * math.exp(-g * M * near_ground_height_m / (R * surface_temperature_k))

    pressure_pa = surface_pressure_pa * (surface_temperature_k / target_temperature_k) ** (g * M / (R * temperature_gradient))

    return pressure_pa


# ===========================================================================
# Shared Helpers
# ===========================================================================

def add_profile_row(
    profile_rows,
    height_asl,
    pressure_pa,
    temperature_k,
    source,
    wind_speed=None,
    wind_direction=None,
):
    """
    Add one profile row if all required values are present.
    """
    if (
        pd.isna(height_asl)
        or pd.isna(pressure_pa)
        or pd.isna(temperature_k)
        or pd.isna(wind_speed)
        or pd.isna(wind_direction)
    ):
        return

    wind_u, wind_v = wind_speed_direction_to_uv(wind_speed, wind_direction)

    profile_rows.append({
        "height": height_asl,
        "pressure": pressure_pa,
        "temperature": temperature_k,
        "wind_speed": wind_speed,
        "wind_direction": wind_direction,
        "wind_u": wind_u,
        "wind_v": wind_v,
        "source": source,
    })


def remove_duplicate_heights(profile_dataframe: pd.DataFrame):
    """
    Remove exact duplicate heights after rounding.
    """
    profile_dataframe = profile_dataframe.sort_values("height", kind="mergesort").reset_index(drop=True)
    # This keeps the first row, which means near-ground rows win over pressure-level rows if they have exactly the same ASL height.
    profile_dataframe["height_rounded_for_duplicate_check"] = profile_dataframe["height"].round(3)

    profile_dataframe = profile_dataframe.drop_duplicates(
        subset="height_rounded_for_duplicate_check",
        keep="first",
    )

    profile_dataframe = profile_dataframe.drop(columns=["height_rounded_for_duplicate_check"])
    return profile_dataframe.reset_index(drop=True)


def filter_profile_to_expected_height(
    profile_dataframe: pd.DataFrame,
    used_launch_site_elevation_m,
    max_expected_height_agl_m,
):
    """
    Keep all profile rows up to the expected maximum height plus the first row above it.
    """
    if max_expected_height_agl_m is None:
        return profile_dataframe.reset_index(drop=True)

    max_expected_height_asl = used_launch_site_elevation_m + max_expected_height_agl_m

    rows_below_or_equal_limit = profile_dataframe[profile_dataframe["height"] <= max_expected_height_asl]
    rows_above_limit = profile_dataframe[profile_dataframe["height"] > max_expected_height_asl]

    if rows_above_limit.empty:
        filtered_dataframe = rows_below_or_equal_limit.copy()
    else:
        first_row_above_limit = rows_above_limit.iloc[[0]]
        filtered_dataframe = pd.concat([rows_below_or_equal_limit, first_row_above_limit], ignore_index=True)

    return filtered_dataframe.reset_index(drop=True)


# ===========================================================================
# Open-Meteo DWD ICON
# ===========================================================================
def build_profile_from_openmeteo_dwd_icon(
    latitude,
    longitude,
    launch_time,
    timezone,
    weather_model,
    max_expected_height_agl_m,
):
    """
    Download Open-Meteo DWD ICON data and build the common ASL atmosphere dataframe.
    """
    # --- Build list of hourly variables to download from Open-Meteo. ---
    hourly_variables = [
        # Surface values for the launch_site_elevation_m row.
        "temperature_2m",
        "surface_pressure",
        "wind_speed_10m",
        "wind_direction_10m",

        # Near-ground fixed-height values.
        "temperature_80m",
        "temperature_120m",
        "temperature_180m",
        "wind_speed_80m",
        "wind_speed_120m",
        "wind_speed_180m",
        "wind_direction_80m",
        "wind_direction_120m",
        "wind_direction_180m",
    ]

    for pressure_level_hpa in PRESSURE_LEVELS_HPA:
        hourly_variables.append(f"temperature_{pressure_level_hpa}hPa")
        hourly_variables.append(f"wind_speed_{pressure_level_hpa}hPa")
        hourly_variables.append(f"wind_direction_{pressure_level_hpa}hPa")
        hourly_variables.append(f"geopotential_height_{pressure_level_hpa}hPa")     # altitude ASL for pressure level rows

    # -----------------------------------------------------------------------
    # Download the weather data using Open-Meteo API
    # -----------------------------------------------------------------------
    cache_session = requests_cache.CachedSession('.cache', expire_after = 3600)
    retry_session = retry(cache_session, retries = 5, backoff_factor = 0.2)
    openmeteo = openmeteo_requests.Client(session = retry_session)

    url = "https://api.open-meteo.com/v1/dwd-icon"

    params = {
        "latitude": latitude,
        "longitude": longitude,
        "hourly": hourly_variables,
        "models": weather_model,
        "wind_speed_unit": "ms",
        "timezone": timezone,
        "start_hour": launch_time,
        "end_hour": launch_time,
    }

    responses = openmeteo.weather_api(url, params = params)
    response = responses[0]

    # -----------------------------------------------------------------------
    # Convert Open-Meteo response object to pandas DataFrame
    # -----------------------------------------------------------------------
    hourly = response.Hourly()

    hourly_data = {
        "date": pd.date_range(
            start=pd.to_datetime(hourly.Time(), unit="s", utc=True),
            end=pd.to_datetime(hourly.TimeEnd(), unit="s", utc=True),
            freq=pd.Timedelta(seconds=hourly.Interval()),
            inclusive="left",
        ).tz_convert(timezone)
    }

    for variable_index, variable_name in enumerate(hourly_variables):
        hourly_data[variable_name] = hourly.Variables(variable_index).ValuesAsNumpy()

    hourly_dataframe = pd.DataFrame(data=hourly_data)       # is only one row due to start_hour = end_hour

    # for debugging
    # print("Hourly raw dataframe transposed:")
    # print(hourly_dataframe.T.to_string())
    # print()

    if hourly_dataframe.empty:
        raise RuntimeError("The API returned no hourly rows. Check launch_time, model, and forecast availability.")

    # -----------------------------------------------------------------------
    # Build the profile dataframe from the hourly data
    # -----------------------------------------------------------------------
    selected_row = hourly_dataframe.iloc[0]
    used_launch_site_elevation_m = float(response.Elevation())

    profile_rows = []

    surface_temperature_k = selected_row["temperature_2m"] + 273.15
    surface_pressure_pa = selected_row["surface_pressure"] * 100.0


    # --- Near-ground rows: 0, 80, 120, 180 m AGL, exported as ASL ---
    for height in NEAR_GROUND_HEIGHTS_M:
        height_asl = used_launch_site_elevation_m + height

        if height == 0:
            # use surface values
            temperature_k = surface_temperature_k
            pressure_pa = surface_pressure_pa
            wind_speed = selected_row["wind_speed_10m"]
            wind_direction = selected_row["wind_direction_10m"]
            source = "near_ground_0m_from_2m_temperature_10m_wind"
        else:
            temperature_k = selected_row[f"temperature_{height}m"] + 273.15
            pressure_pa = compute_pressure_with_temperature_gradient(
                surface_pressure_pa,
                surface_temperature_k,
                temperature_k,
                height,
            )
            wind_speed = selected_row[f"wind_speed_{height}m"]
            wind_direction = selected_row[f"wind_direction_{height}m"]
            source = f"near_ground_{height}m"

        add_profile_row(
            profile_rows,
            height_asl,
            pressure_pa,
            temperature_k,
            source,
            wind_speed=wind_speed,
            wind_direction=wind_direction,
        )

    # Collect the exact ASL heights of the near-ground rows so we can filter pressure-level rows near them.
    near_ground_heights_asl = [used_launch_site_elevation_m + h for h in NEAR_GROUND_HEIGHTS_M]

    # --- Pressure-level rows: use geopotential height directly as ASL ---
    for pressure_level_hpa in PRESSURE_LEVELS_HPA:
        temperature_key = f"temperature_{pressure_level_hpa}hPa"
        wind_speed_key = f"wind_speed_{pressure_level_hpa}hPa"
        wind_direction_key = f"wind_direction_{pressure_level_hpa}hPa"
        height_key = f"geopotential_height_{pressure_level_hpa}hPa"

        temperature_celsius = selected_row[temperature_key]
        wind_speed = selected_row[wind_speed_key]
        wind_direction = selected_row[wind_direction_key]
        height_asl = selected_row[height_key]

        # Skip this pressure level if it lands too close to a near-ground row
        if not pd.isna(height_asl) and any(
            abs(height_asl - ng_h) < MIN_HEIGHT_GAP_FROM_NEAR_GROUND_M
            for ng_h in near_ground_heights_asl
        ):
            continue

        if (
            pd.isna(temperature_celsius)
            or pd.isna(wind_speed)
            or pd.isna(wind_direction)
            or pd.isna(height_asl)
        ):
            continue

        add_profile_row(
            profile_rows,
            height_asl,
            pressure_level_hpa * 100.0,     # hPa to Pa
            temperature_celsius + 273.15,   # °C to K
            f"{pressure_level_hpa}hPa",
            wind_speed=wind_speed,
            wind_direction=wind_direction,
        )

    if not profile_rows:
        raise RuntimeError("No valid profile rows were created. Try another model, time, or location.")

    profile_dataframe = pd.DataFrame(profile_rows)
    profile_dataframe = remove_duplicate_heights(profile_dataframe)

    profile_dataframe = filter_profile_to_expected_height(
        profile_dataframe,
        used_launch_site_elevation_m,
        max_expected_height_agl_m,
    )

    metadata = {
        "source": "Open-Meteo DWD ICON",
        "used_latitude": response.Latitude(),
        "used_longitude": response.Longitude(),
        "used_launch_site_elevation_m": used_launch_site_elevation_m,
        "selected_forecast_time": hourly_dataframe.iloc[0]["date"],
        "weather_model": weather_model,
    }

    return profile_dataframe, metadata


# ===========================================================================
# Export Files
# ===========================================================================

def write_rocketpy_csv(profile_dataframe, output_path):
    """
    Write the RocketPy atmosphere profile CSV.

    Columns:
    - height: height above sea level [m ASL]
    - pressure: pressure [Pa]
    - temperature: temperature [K]
    - wind_u: eastward wind component [m/s]
    - wind_v: northward wind component [m/s]
    """
    profile_dataframe[["height", "pressure", "temperature", "wind_u", "wind_v"]].to_csv(output_path, index=False)


def write_openrocket_csv(profile_dataframe, output_path):
    """
    Write the OpenRocket wind profile CSV.

    Columns:
    - altitude: altitude above sea level [m ASL]
    - speed: wind speed [m/s]
    - direction: meteorological wind direction [deg], direction the wind comes from
    - stddev: left empty (OpenRocket requires this column but we don't have this data from Open-Meteo)
    """
    # create new header and insert empty stddev column
    openrocket_dataframe = pd.DataFrame({
        "altitude": profile_dataframe["height"],
        "speed": profile_dataframe["wind_speed"],
        "direction": profile_dataframe["wind_direction"],
        "stddev": "",
    })

    openrocket_dataframe.to_csv(output_path, index=False)


# ===========================================================================
# Main Script
# ===========================================================================

def export_weather_data(
    latitude: float,
    longitude: float,
    max_expected_height_agl_m: float | None,
    launch_time: str,
    timezone: str,
    weather_model: str,
    output_folder: str | Path = "./weather_csvs",
):
    """
    Download the weather profile and write export CSV files for RocketPy and OpenRocket.
    
    Parameters
    ----------
    latitude : float
        Latitude of the launch site in decimal degrees.
    longitude : float
        Longitude of the launch site in decimal degrees.
    max_expected_height_agl_m : float or None
        Maximum expected height of the rocket above ground level (AGL) in meters. 
        Only profile rows up to this height plus the first row above it are exported. If None, all valid heights are exported.
    launch_time : str
        The time of the launch in local time of the launch site. Format: "YYYY-MM-DDTHH:MM", e.g. "2026-05-14T12:00".
    timezone : str
        The timezone of the launch site. As found here https://en.wikipedia.org/wiki/List_of_tz_database_time_zones#List.
    weather_model : str
        The weather model to use for the API request.
        Choose from: "icon_d2", "icon_eu", "icon_global"
    output_folder : str or Path
        Folder where the RocketPy and OpenRocket CSV files are written.
        
    Returns
    -------
    None
    """
    if weather_model in OPENMETEO_MODELS:
        profile_dataframe, metadata = build_profile_from_openmeteo_dwd_icon(
            latitude,
            longitude,
            launch_time,
            timezone,
            weather_model,
            max_expected_height_agl_m,
        )
    else:
        raise ValueError(
            f"Unknown weather model '{weather_model}'. "
            f"Choose one of: {OPENMETEO_MODELS}"
        )

    output_folder = Path(output_folder)
    output_folder.mkdir(parents=True, exist_ok=True) 
    write_rocketpy_csv(profile_dataframe, output_folder / f"rocketpy_atmosphere_{weather_model}.csv")
    write_openrocket_csv(profile_dataframe, output_folder / f"openrocket_wind_profile_{weather_model}.csv")

    print(f"Requested coordinates: {latitude}°N, {longitude}°E")
    print(f"Used model coordinates: {metadata['used_latitude']}°N, {metadata['used_longitude']}°E")
    print(f"Used launch elevation: {metadata['used_launch_site_elevation_m']} m ASL")
    print(f"Data source: {metadata['source']}")

    if max_expected_height_agl_m is None:
        print("Maximum expected height: not set, exporting all valid heights")
    else:
        max_expected_height_asl = metadata["used_launch_site_elevation_m"] + max_expected_height_agl_m
        print(f"Maximum expected height: {max_expected_height_agl_m} m AGL")
        print(f"Maximum expected height: {max_expected_height_asl} m ASL")
        print("We export all rows up to this height plus the first row above it.")

    selected_forecast_time = metadata["selected_forecast_time"]

    print(f"Selected forecast time: {selected_forecast_time.strftime('%Y-%m-%d %H:%M:%S (%Z)')}")

    print(f"Weather model: {weather_model}")
    print(f"Created profile rows: {len(profile_dataframe)}")
    print()
    print("Profile:")
    print(profile_dataframe[[
        "height",
        "pressure",
        "temperature",
        "wind_speed",
        "wind_direction",
        "wind_u",
        "wind_v",
        "source",
    ]].to_string(index=False))
    print()
    print(f"Wrote rocketpy_atmosphere_{weather_model}.csv")
    print(f"Wrote openrocket_wind_profile_{weather_model}.csv")


if __name__ == "__main__":
    export_weather_data(latitude, longitude, max_expected_height_agl_m, launch_time, timezone, WEATHER_MODEL)