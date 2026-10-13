"""
Landing zones: import them from a KML file (e.g. drawn in Google Earth), load them from zones.toml and check where a
landing point lies.

KML import: the KML file needs a point named "launch_rail" and one polygon per zone. For every polygon corner the
distance [m] and heading [deg] from the launch rail are calculated on the WGS84 ellipsoid (heading: 0 = North, 90 = East,
clockwise). An optional polygon named "border" is the border of the launch area. A polygon whose name contains "exclusion"
is an exclusion zone, all others are buffer zones; a "(exclusion)" marker is removed from the zone name, e.g.
"Cars (exclusion)" becomes the exclusion zone "Cars".

Usage as a script (prints the zones in the zones.toml format), from the repo folder:
    python -m simulation_core.zones "projects/ALBATROSS_II/ALBATROSS_II_zones.kml"
The Streamlit app uses read_kml_zones to write the zones directly into a project's zones.toml.
"""
import json
import math
import re
import sys
import tomllib
import xml.etree.ElementTree as ElementTree
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from pyproj import Geod
from shapely import affinity
from shapely.geometry import MultiPolygon, Point, Polygon

from simulation_core.config_schema import ZonesConfig
from simulation_core.export_satellite_image import DOWNLOAD_KEYWORD, download_satellite_image

KML_NAMESPACE = {"kml": "http://www.opengis.net/kml/2.2"}
# Name of the KML point the distances and headings are measured from.
LAUNCH_RAIL_NAME = "launch_rail"
# Name (any case) of the optional KML polygon that borders the launch area.
BORDER_NAME = "border"
# A polygon whose name contains this word (any case) is an exclusion zone.
EXCLUSION_KEYWORD = "exclusion"
# Removed from zone names: "(exclusion)" with any case and spacing, e.g. "Cars (exclusion)" or "Cars exclusion" -> "Cars"
EXCLUSION_MARKER_PATTERN = re.compile(r"\s*\(?\s*\bexclusion\b\s*\)?", re.IGNORECASE)
# Decimals of distance and heading in the output, like the hand-written zones.toml files.
DECIMALS = 2
WGS84 = Geod(ellps="WGS84")
# Zone types of a landing point (see landing_zone_type) that make its heading unsafe.
UNSAFE_ZONE_TYPES = {"outside border", "unsafe"}


# =============================================================================
# KML import
# =============================================================================

def parse_kml_coordinates(coordinates_text: str) -> list[tuple[float, float]]:
    """Convert KML coordinate text like "lon,lat,alt lon,lat,alt" into (longitude, latitude) pairs; the altitude is ignored."""
    return [(float(point.split(",")[0]), float(point.split(",")[1])) for point in coordinates_text.split()]


def read_kml_zones(kml_content: bytes | str) -> dict[str, Any]:
    """Return the zones of a KML file's content as zones.toml values: {"exclusion_zones": {name: corners},
    "buffer_zones": {name: corners}, "border": corners or None}, each corner as [distance_m, heading_deg].

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

    # Keys are the ZonesConfig field names, so the result can go straight into zones.toml or the zones form.
    kml_zones = {"exclusion_zones": {}, "buffer_zones": {}, "border": None}
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
        if name.lower() == BORDER_NAME:
            kml_zones["border"] = polar_corners
        elif EXCLUSION_KEYWORD in name.lower():
            kml_zones["exclusion_zones"][EXCLUSION_MARKER_PATTERN.sub("", name).strip()] = polar_corners
        else:
            kml_zones["buffer_zones"][name] = polar_corners
    return kml_zones


def format_zones(kml_zones: dict[str, Any]) -> str:
    """Format the zones from read_kml_zones as zones.toml text: the border (if any), then the buffer and exclusion zone tables."""
    sections = []
    if kml_zones["border"] is not None:
        sections.append(f"border = {json.dumps(kml_zones['border'])}")
    for table_name in ("buffer_zones", "exclusion_zones"):
        # json.dumps gives a quoted name and a [[a, b], ...] list, both valid TOML.
        lines = [f"{json.dumps(name, ensure_ascii=False)} = {json.dumps(corners)}" for name, corners in kml_zones[table_name].items()]
        sections.append("\n".join([f"[{table_name}]", *lines]))
    return "\n\n".join(sections)


# =============================================================================
# Zone loading
# =============================================================================

def polar_to_polygon(polar_corners: list[tuple[float, float]]) -> Polygon:
    """Turn corners in polar [distance_m, heading_deg] (heading clockwise from north) into a polygon in (x_east_m, y_north_m)."""
    return Polygon([
        (distance_m * math.sin(math.radians(heading_deg)), distance_m * math.cos(math.radians(heading_deg)))
        for distance_m, heading_deg in polar_corners
    ])


@dataclass
class Zones:
    """The landing zones of a project, loaded by `Zones.load_from_toml`; all polygons are in meters (x east, y north) around
    the launch rail.

    A flight is unsafe if any trajectory (nominal, no_main, ballistic) lands in a buffer zone or outside the optional border,
    with its additional safety margin and suboptimal if it only lands in a suboptimal zone (see `Zones.landing_zone_type`).

    Attributes
    ----------
    exclusion
        Zones where the rocket must not land, keyed by zone name.
    buffer
        The buffer zones (safety margins around exclusion zones, or zones where the rocket should preferably not land,
        such as forests) plus the exclusion zones enlarged by `exclusion_zone_safety_margin`. Landing here is unsafe.
    suboptimal
        The buffer zones when `buffer_zones_are_suboptimal_but_safe` is true (they are then not in buffer), otherwise empty.
        Landing here is safe but should preferably be avoided.
    border
        The optional border of the launch area, or None if zones.toml has none.
    safe_border
        The border shrunk inward by `border_safety_margin_m`; landing outside of it is unsafe. Can split into several
        polygons where the border comes too close to itself.
    satellite_image
        Path of the background GeoTIFF for the landing plots, or None if zones.toml sets none.
        With satellite_image = "download" the image is downloaded while loading.
    """
    exclusion: dict[str, Polygon] = field(default_factory=dict)
    buffer: dict[str, Polygon] = field(default_factory=dict)
    suboptimal: dict[str, Polygon] = field(default_factory=dict)
    border: Polygon | None = None
    safe_border: Polygon | MultiPolygon | None = None
    satellite_image: Path | None = None

    @classmethod
    def load_from_toml(cls, zones_path: Path) -> "Zones":
        """Load the landing zones from a zones.toml file; without the file all zones are empty."""
        if not zones_path.exists():
            return cls()

        with open(zones_path, "rb") as zones_file:
            zones_config = ZonesConfig.model_validate(tomllib.load(zones_file))

        exclusion_zones = {name: polar_to_polygon(corners) for name, corners in zones_config.exclusion_zones.items()}
        buffer_zones = {name: polar_to_polygon(corners) for name, corners in zones_config.buffer_zones.items()}
        # Enlarge each exclusion zone around its own centroid, e.g. by 1.2 = 20% in every direction.
        margin = zones_config.exclusion_zone_safety_margin
        enlarged_exclusion_zones = {
            name: affinity.scale(polygon, xfact=margin, yfact=margin, origin="centroid") for name, polygon in exclusion_zones.items()
        }

        project_path = zones_path.parent
        if zones_config.satellite_image == DOWNLOAD_KEYWORD:
            satellite_image = download_satellite_image(project_path, zones_config.satellite_image_radius)
        elif zones_config.satellite_image:
            satellite_image = project_path / zones_config.satellite_image
        else:
            satellite_image = None

        border = safe_border = None
        if zones_config.border:
            border = polar_to_polygon(zones_config.border)
            # A negative buffer moves every edge inward by the margin; scaling would shrink a large border unevenly.
            safe_border = border.buffer(-zones_config.border_safety_margin_m)

        # The enlarged exclusion zones are always unsafe; the buffer zones only unless they are marked as suboptimal but safe.
        if zones_config.buffer_zones_are_suboptimal_but_safe:
            unsafe_zones, suboptimal_zones = enlarged_exclusion_zones, buffer_zones
        else:
            unsafe_zones, suboptimal_zones = {**buffer_zones, **enlarged_exclusion_zones}, {}
        return cls(exclusion_zones, unsafe_zones, suboptimal_zones, border, safe_border, satellite_image)

    def landing_zone_type(self, x_east_m: float, y_north_m: float) -> str:
        """Return where a landing point lies: "outside border", "unsafe" (buffer zone), "subopt." (suboptimal zone) or "safe"."""
        landing_point = Point(x_east_m, y_north_m)
        # Points within the safety margin of the border count as outside, too.
        if self.safe_border is not None and not self.safe_border.contains(landing_point):
            return "outside border"
        if any(zone.contains(landing_point) for zone in self.buffer.values()):
            return "unsafe"
        if any(zone.contains(landing_point) for zone in self.suboptimal.values()):
            return "subopt."
        return "safe"


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit('Usage: python -m simulation_core.zones "path/to/zones.kml"')
    print(format_zones(read_kml_zones(Path(sys.argv[1]).read_bytes())))
