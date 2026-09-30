"""
Read landing zones from a KML file (e.g. drawn in Google Earth) and convert them to the zones.toml format.

The KML file needs a point named "launch_rail" and one polygon per zone. For every polygon corner the distance [m]
and heading [deg] from the launch rail are calculated on the WGS84 ellipsoid (heading: 0 = North, 90 = East, clockwise).
A polygon whose name contains "exclusion" is an exclusion zone, all others are buffer zones; a "(exclusion)" marker is
removed from the zone name, e.g. "Cars (exclusion)" becomes the exclusion zone "Cars".

Usage as a script (prints the zones):
    python simulation_core/kml_zones.py "projects/ALBATROSS_II/ALBATROSS II.kml"
The Streamlit app uses read_kml_zones to write the zones directly into a project's zones.toml.
"""
import json
import re
import sys
import xml.etree.ElementTree as ElementTree
from pathlib import Path

from pyproj import Geod

KML_NAMESPACE = {"kml": "http://www.opengis.net/kml/2.2"}
# Name of the KML point the distances and headings are measured from.
LAUNCH_RAIL_NAME = "launch_rail"
# A polygon whose name contains this word (any case) is an exclusion zone.
EXCLUSION_KEYWORD = "exclusion"
# Removed from zone names: "(exclusion)" with any case and spacing, e.g. "Cars (exclusion)" or "Cars exclusion" -> "Cars" 
EXCLUSION_MARKER_PATTERN = re.compile(r"\s*\(?\s*\bexclusion\b\s*\)?", re.IGNORECASE)
# Decimals of distance and heading in the output, like the hand-written zones.toml files.
DECIMALS = 2
WGS84 = Geod(ellps="WGS84")


def parse_kml_coordinates(coordinates_text: str) -> list[tuple[float, float]]:
    """Turn KML coordinate text like "lon,lat,alt lon,lat,alt" into (longitude, latitude) pairs; the altitude is ignored."""
    return [(float(point.split(",")[0]), float(point.split(",")[1])) for point in coordinates_text.split()]


def read_kml_zones(kml_content: bytes | str) -> tuple[dict[str, list[list[float]]], dict[str, list[list[float]]]]:
    """Return (exclusion_zones, buffer_zones) from a KML file's content, each as {name: [[distance_m, heading_deg], ...]}.

    Raises ValueError if the content is no valid KML or has no "launch_rail" point.
    """
    # Bytes are parsed with the encoding the file declares itself, e.g. <?xml version="1.0" encoding="UTF-8"?>.
    try:
        root = ElementTree.fromstring(kml_content)
    except ElementTree.ParseError as error:
        raise ValueError(f"Not a valid KML file: {error}") from error

    launch_rail = None
    polygons = {}
    for placemark in root.iter(f"{{{KML_NAMESPACE['kml']}}}Placemark"):
        name = placemark.findtext("kml:name", default="", namespaces=KML_NAMESPACE).strip()
        point_text = placemark.findtext(".//kml:Point/kml:coordinates", namespaces=KML_NAMESPACE)
        ring_text = placemark.findtext(".//kml:Polygon/kml:outerBoundaryIs/kml:LinearRing/kml:coordinates", namespaces=KML_NAMESPACE)
        if name.lower() == LAUNCH_RAIL_NAME and point_text:
            launch_rail = parse_kml_coordinates(point_text)[0]
        elif ring_text:
            polygons[name] = parse_kml_coordinates(ring_text)
    if launch_rail is None:
        raise ValueError(f'The KML file has no point named "{LAUNCH_RAIL_NAME}".')

    exclusion_zones, buffer_zones = {}, {}
    launch_longitude, launch_latitude = launch_rail
    for name, corners in polygons.items():
        # KML closes a polygon by repeating its first corner at the end; zones.toml polygons are closed automatically.
        if len(corners) > 1 and corners[0] == corners[-1]:
            corners = corners[:-1]
        polar_corners = []
        for longitude, latitude in corners:
            # inv returns the forward azimuth (-180..180 deg), the back azimuth and the distance along the ellipsoid [m].
            heading, _, distance = WGS84.inv(launch_longitude, launch_latitude, longitude, latitude)
            # "% 360" twice: first to 0..360, then again in case rounding turns e.g. 359.999 into 360.0.
            polar_corners.append([round(distance, DECIMALS), round(heading % 360, DECIMALS) % 360])
        zones = exclusion_zones if EXCLUSION_KEYWORD in name.lower() else buffer_zones
        zones[EXCLUSION_MARKER_PATTERN.sub("", name).strip()] = polar_corners
    return exclusion_zones, buffer_zones


def format_zones(exclusion_zones: dict[str, list[list[float]]], buffer_zones: dict[str, list[list[float]]]) -> str:
    """Format the zones as TOML lines per zone type, buffer zones first, e.g. "Forest North" = [[341.7, 331.0], ...]."""
    sections = []
    for title, zones in [("buffer zones", buffer_zones), ("exclusion zones", exclusion_zones)]:
        # json.dumps gives a quoted name and a [[a, b], ...] list, both valid TOML.
        lines = [f"{json.dumps(name, ensure_ascii=False)} = {json.dumps(corners)}" for name, corners in zones.items()]
        sections.append("\n".join([f"{title}:", *lines]))
    return "\n\n".join(sections)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit('Usage: python simulation_core/kml_zones.py "path/to/zones.kml"')
    exclusion, buffer = read_kml_zones(Path(sys.argv[1]).read_bytes())
    print(format_zones(exclusion, buffer))
