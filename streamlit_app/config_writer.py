"""
Config file writer for the Streamlit app (app.py): creates and updates `config.toml` and `zones.toml` from the form values.

Raw values are what a user types: a dict like {"flight": {"heading": "340..360:10"}}. They are validated with the
pydantic model first, but the raw values are written, so range strings and dates keep the form the user entered.

Usage:
    write_config_file(Config, values, Path("projects/NEW_PROJECT/config.toml"))   # create a new file
    update_config_file(Config, values, Path("projects/ALBATROSS/config.toml"))    # change an existing file, keep its comments
"""

import os
from pathlib import Path
from types import UnionType
from typing import Any, Union, get_args, get_origin

import tomlkit
from pydantic import BaseModel
from tomlkit.container import Container
from tomlkit.items import Item, Table

from simulation_core.config_schema import CONFIG_SCHEMA_PATH, ZONES_SCHEMA_PATH, Config, ZonesConfig

SCHEMA_PATH_BY_MODEL = {Config: CONFIG_SCHEMA_PATH, ZonesConfig: ZONES_SCHEMA_PATH}

# Comment lines written at the top of a new file
HEADER_COMMENTS_BY_MODEL = {
    Config: ["Hover a key for its meaning and optional unit (Tombi); TOML syntax and variations: see README.md."],
    ZonesConfig: ["Coordinates: [distance_m, heading_deg] from the launch rail. heading: 0 = North, 90 = East, clockwise."],
}


# =============================================================================
# Public functions
# =============================================================================

def read_comments(path: Path) -> dict[str, str]:
    """Return the trailing comment of every key = value line, keyed by its TOML path, e.g. {"rocket.length": "[mm] measured"}."""
    with open(path, encoding="utf-8", newline="") as config_file:
        document = tomlkit.parse(config_file.read())
    comments = {}
    _collect_comments(document, "", comments)
    return comments


def write_config_file(model_class: type[BaseModel], values: dict, path: Path, comments: dict[str, str] | None = None) -> None:
    """Validate the raw values and write them as a new TOML file that starts with its #:schema line and the model's header comments.

    comments are trailing comments per TOML path, e.g. {"rocket.length": "[mm] measured"}.
    Raises pydantic.ValidationError (nothing is written) when the values are invalid.
    """
    # Validate first, so an invalid config never ends up on disk.
    model_class.model_validate(values)

    document = tomlkit.document()
    _fill_container(document, model_class, values, comments or {}, "")

    schema_path = Path(os.path.relpath(SCHEMA_PATH_BY_MODEL[model_class], path.parent)).as_posix()
    header = f"#:schema {schema_path}\n"
    header += "".join(f"# {comment}\n" for comment in HEADER_COMMENTS_BY_MODEL.get(model_class, []))
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(header + tomlkit.dumps(document), encoding="utf-8")


def update_config_file(model_class: type[BaseModel], values: dict, path: Path, comments: dict[str, str] | None = None) -> None:
    """Replace the content of an existing TOML file with the raw values, keeping its comments and the formatting of unchanged values.

    values is the complete new config: keys that are missing (or None) are removed from the file.
    comments sets the trailing comment of the listed TOML paths ("" removes it); comments of paths not listed stay as they are.
    Raises pydantic.ValidationError (the file stays unchanged) when the values are invalid.
    """
    model_class.model_validate(values)

    # newline="" turns off line-ending conversion, so the file keeps its own LF or CRLF endings and unchanged lines stay identical in git.
    with open(path, encoding="utf-8", newline="") as config_file:
        document = tomlkit.parse(config_file.read())
    _merge_into_container(document, model_class, values, comments or {}, "")
    with open(path, "w", encoding="utf-8", newline="") as config_file:
        config_file.write(tomlkit.dumps(document))


# =============================================================================
# Helpers
# =============================================================================

def _section_model(model_class: type[BaseModel] | None, key: str) -> type[BaseModel] | None:
    """Return the model of a [section] field (e.g. EnvironmentConfig for "environment"), or None if the field holds no model.

    For dict fields with model values, like reanalysis.matched_flight.parachutes, the value model is returned, so that
    their entries ([...parachutes.main]) become sections too. Keys that are not model fields are such dict entries.
    """
    if model_class is None:
        return None
    field = model_class.model_fields.get(key)
    if field is None:
        return model_class
    # Optional[X] / Union[X, Y] list their options in get_args; any other type (including dict[str, X]) is the only option.
    annotation = field.annotation
    options = get_args(annotation) if get_origin(annotation) in (Union, UnionType) else (annotation,)
    for option in options:
        if get_origin(option) is dict:
            option = get_args(option)[1]
        if isinstance(option, type) and issubclass(option, BaseModel):
            return option
    return None


def _is_table(model_class: type[BaseModel] | None, key: str, value: Any) -> bool:
    """Decide whether a value becomes a [section]: model sections and dicts holding lists (e.g. zone polygons) do, 
    small dicts of plain values stay inline."""
    if not isinstance(value, dict):
        return False
    if _section_model(model_class, key) is not None:
        return True
    return any(isinstance(entry, (dict, list)) for entry in value.values())


def _ordered_items(model_class: type[BaseModel] | None, values: dict) -> list[tuple[str, Any]]:
    """Return the non-empty values in model field order, plain values before sections, so the file reads like the schema."""
    field_order = list(model_class.model_fields) if model_class is not None else []
    items = [(key, value) for key, value in values.items() if value is not None]
    # Keys that are not model fields (dict entries like "main" or zone names) keep their given order after the known fields.
    items.sort(key=lambda item: field_order.index(item[0]) if item[0] in field_order else len(field_order))
    return sorted(items, key=lambda item: _is_table(model_class, item[0], item[1]))


def _to_toml_value(value: Any) -> Any:
    """Convert a plain value for tomlkit: dicts become inline tables, e.g. reanalysis_csv = { icon_d2 = "x.csv" }."""
    if isinstance(value, dict):
        inline_table = tomlkit.inline_table()
        inline_table.update(value)
        return inline_table
    return value


def _collect_comments(container: Container | Table, path_prefix: str, comments: dict[str, str]) -> None:
    """Add the trailing comments of all key = value lines in a document or table (and its sub-tables) to comments."""
    for key in container.keys():
        # .item() also returns true/false values as tomlkit items; plain access would give a Python bool without its comment.
        item = container.item(key)
        path = f"{path_prefix}.{key}" if path_prefix else key
        if isinstance(item, Table):
            _collect_comments(item, path, comments)
        elif item.trivia.comment:
            comments[path] = item.trivia.comment.removeprefix("#").strip()


def _set_comment(item: Item, comment: str) -> None:
    """Set, change or remove ("") the trailing comment of one key = value line; only a changed comment is rewritten."""
    # Compare the text only, as read_comments returns it, so e.g. trailing spaces after an unchanged comment don't count as a change.
    if item.trivia.comment.removeprefix("#").strip() == comment:
        return
    new_comment = f"# {comment}" if comment else ""
    # A changed comment keeps its spacing, so aligned comments stay aligned; a new one gets two spaces, a removed one none.
    if not new_comment:
        item.trivia.comment_ws = ""
    elif not item.trivia.comment_ws:
        item.trivia.comment_ws = "  "
    item.trivia.comment = new_comment


def _fill_container(
    container: Container | Table, model_class: type[BaseModel] | None, values: dict, comments: dict[str, str], path_prefix: str
) -> None:
    """Add all values to an empty document or table: sections as [section] tables, everything else as key = value (with its comment)."""
    for key, value in _ordered_items(model_class, values):
        path = f"{path_prefix}.{key}" if path_prefix else key
        if _is_table(model_class, key, value):
            table = tomlkit.table()
            _fill_container(table, _section_model(model_class, key), value, comments, path)
            container[key] = table
        else:
            container[key] = _to_toml_value(value)
            if comments.get(path):
                _set_comment(container.item(key), comments[path])


def _merge_into_container(
    container: Container | Table, model_class: type[BaseModel] | None, values: dict, comments: dict[str, str], path_prefix: str
) -> None:
    """Change an existing document or table so it holds exactly the given values (and comments), touching only keys that differ."""
    new_values = {key: value for key, value in values.items() if value is not None}

    # Remove keys the new values no longer contain; comments and commented-out lines around them stay.
    for key in list(container.keys()):
        if key not in new_values:
            del container[key]

    for key, value in _ordered_items(model_class, new_values):
        path = f"{path_prefix}.{key}" if path_prefix else key
        existing = container.get(key)
        is_table = _is_table(model_class, key, value)
        if is_table and isinstance(existing, (Table, Container)):
            _merge_into_container(existing, _section_model(model_class, key), value, comments, path)
        elif is_table:
            # A section that is new in this file: build it the same way as in write_config_file.
            table = tomlkit.table()
            _fill_container(table, _section_model(model_class, key), value, comments, path)
            container[key] = table
        else:
            if existing != value:
                # Only rewrite changed values, so unchanged ones keep their formatting; tomlkit keeps the line's trailing comment.
                container[key] = _to_toml_value(value)
            if path in comments:
                _set_comment(container.item(key), comments[path])
