"""
Download a satellite image around a project's launch rail and save it as GeoTIFF for the landing plots.

The image is a square of 2 * satellite_image_radius (set in zones.toml) around the launch rail coordinates in the project's
config.toml, downloaded at ZOOM_LEVEL. It is saved JPEG-compressed in Web Mercator (EPSG:3857) as satellite_image.tif
in the project folder.

Web Mercator (EPSG:3857) is a projected coordinate system, that flattens Earth onto a 2D plane using meters as units.
EPSG:4326 uses unprojected geographic coordinates (latitude and longitude measured in degrees) on a 3D globe.

Depending on the size downloading can take a few minutes. If the tile server stalls (it throttles after a few hundred
tiles), the download pauses for TILE_STALL_WAIT_S and continues by itself. Only if it keeps stalling it stops with an
error message; tiles that already downloaded are cached in TILE_CACHE_DIR, so a rerun only fetches the missing ones.

Two ways to use it:
- zones.toml: satellite_image = "download" downloads the image on every run (see Zones.load_from_toml in zones.py).
- Terminal, from the repository folder (downloads once and sets satellite_image = "satellite_image.tif" for offline runs):
    python -m simulation_core.export_satellite_image "projects/ALBATROSS_II"
"""
import sys
import tempfile
import time
import tomllib
from pathlib import Path

import mercantile
import numpy as np
import pymap3d
import rasterio
import requests
import tomlkit
import xyzservices.providers as tile_providers

from rasterio.transform import from_bounds

from simulation_core.config_schema import ZonesConfig

DOWNLOAD_KEYWORD = "download"
OUTPUT_FILE_NAME = "satellite_image.tif"
IMAGE_SOURCE = tile_providers.Esri.WorldImagery     # https://www.arcgis.com/apps/mapviewer/index.html?layers=10df2279f9684e4a9f6a7f08febac2a9
# Tile zoom level: 16 is about 1.5 m per pixel at 50° latitude; each step up halves the pixel size and quadruples the
# tif image size.
ZOOM_LEVEL = 16
GEOTIFF_JPEG_QUALITY = 90       # compress the GeoTIFF to this quality (0-100, 100 = no compression)
TILE_REQUEST_TIMEOUT_S = (10, 60)   # (connect, read) seconds before a tile request is aborted with an error
TILE_RETRY_WAIT_S = 15              # seconds to wait before retrying a tile the server rejected (e.g. rate limit)
TILE_MAX_RETRIES = 5                # retries per rejected tile before the download fails
TILE_STALL_WAIT_S = 60              # pause after a timed out / failed tile, so the server stops throttling us
TILE_STALL_MAX_WAITS = 10           # pauses per tile before giving up (server down or no internet)
TILE_CACHE_DIR = Path(tempfile.gettempdir()) / "lamarr_satellite_tiles"     # windows: %TEMP%\lamarr_satellite_tiles


def download_satellite_image(project_path: Path, radius_m: float) -> Path:
    """Download the satellite image around the project's launch rail and save it as satellite_image.tif in the project folder."""
    # only import if download is triggered
    import contextily

    with open(project_path / "config.toml", "rb") as config_file:
        environment = tomllib.load(config_file)["environment"]
    rail_latitude, rail_longitude = environment["latitude"], environment["longitude"]

    # Convert the satellite image edges from meters east/north of the rail into lat/lon (inverse of latlon_to_local_xy in utils.py)
    south, west, _ = pymap3d.enu2geodetic(-radius_m, -radius_m, 0.0, rail_latitude, rail_longitude, 0.0)
    north, east, _ = pymap3d.enu2geodetic(radius_m, radius_m, 0.0, rail_latitude, rail_longitude, 0.0)

    image_path = project_path / OUTPUT_FILE_NAME
    tiles = list(mercantile.tiles(west, south, east, north, [ZOOM_LEVEL]))
    print(f"Downloading satellite image ({2 * radius_m:.0f} m square, zoom {ZOOM_LEVEL}, {len(tiles)} tiles) to {image_path} ...")
    contextily.set_cache_dir(TILE_CACHE_DIR)
    # contextily's internal cached tile fetcher (pinned version): tiles fetched here land in the cache bounds2img reads.
    fetch_tile_cached = contextily.tile.memory.cache(contextily.tile._fetch_tile)
    # Fetch tile by tile to show live progress
    for tile_number, tile in enumerate(tiles, start=1):
        tile_url = IMAGE_SOURCE.build_url(x=tile.x, y=tile.y, z=tile.z)
        # Timeouts, dropped connections and rejected tiles all derive from RequestException: pause, then retry the same tile.
        for wait_number in range(TILE_STALL_MAX_WAITS + 1):
            try:
                fetch_tile_cached(tile_url, TILE_RETRY_WAIT_S, TILE_MAX_RETRIES, {}, timeout=TILE_REQUEST_TIMEOUT_S)
                break
            except requests.RequestException as error:
                # Still stalling after all pauses: give up with a readable message.
                if wait_number == TILE_STALL_MAX_WAITS:
                    print()     # end the progress line so the error message starts on its own line
                    raise RuntimeError(
                        f"Satellite image download failed: {error}\n"
                        f"The tile server did not answer after {TILE_STALL_MAX_WAITS} pauses of {TILE_STALL_WAIT_S} s. "
                        f"Check your internet connection and run the download again; already downloaded tiles are "
                        f"cached, so it continues where it stopped."
                    ) from error
                print(f"\n  Tile server stalled ({error}). Waiting {TILE_STALL_WAIT_S} s, then continuing ...")
                time.sleep(TILE_STALL_WAIT_S)
        print(f"\r  tile {tile_number}/{len(tiles)}", end="", flush=True) # "\r" overwrites the same terminal line each time.
    print()

    # All tiles are cached now, so bounds2img only stitches them into one image in EPSG:3857 (ll=True: bounds are lon/lat).
    image, (min_x, max_x, min_y, max_y) = contextily.bounds2img(
        west, south, east, north, zoom=ZOOM_LEVEL, source=IMAGE_SOURCE, ll=True,
        timeout=TILE_REQUEST_TIMEOUT_S, wait=TILE_RETRY_WAIT_S, max_retries=TILE_MAX_RETRIES,
    )
    height, width = image.shape[:2]

    # The transform ties the pixels to the Web Mercator edges; YCbCr color storage for more compression
    with rasterio.open(
        image_path, "w", driver="GTiff", width=width, height=height, count=3, dtype="uint8", crs="EPSG:3857",
        transform=from_bounds(min_x, min_y, max_x, max_y, width, height),
        compress="JPEG", photometric="YCBCR", jpeg_quality=GEOTIFF_JPEG_QUALITY,
    ) as image_file:
        # contextily returns (row, column, band) with an alpha band; rasterio writes (band, row, column), and JPEG has no alpha.
        image_file.write(np.moveaxis(image[:, :, :3], -1, 0))
    return image_path


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit('Usage: python -m simulation_core.export_satellite_image "projects/<project>"')
    project_folder = Path(sys.argv[1])
    zones_path = project_folder / "zones.toml"

    # tomlkit keeps the comments and layout of zones.toml when the file name is written back below.
    zones_document = tomlkit.parse(zones_path.read_text(encoding="utf-8")) if zones_path.exists() else tomlkit.document()
    zones = ZonesConfig.model_validate(zones_document.unwrap())
    # Show only the readable download error message in the terminal, without the traceback.
    try:
        download_satellite_image(project_folder, zones.satellite_image_radius)
    except RuntimeError as error:
        sys.exit(str(error))

    # Point zones.toml at the saved file, so the following runs use it offline instead of downloading again.
    zones_document["satellite_image"] = OUTPUT_FILE_NAME
    zones_path.write_text(tomlkit.dumps(zones_document), encoding="utf-8")
    print(f'Set satellite_image = "{OUTPUT_FILE_NAME}" in {zones_path}.')
