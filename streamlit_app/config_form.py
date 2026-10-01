"""
Config form for the Streamlit app (app.py): an editable form for a config model (Config or ZonesConfig).

The form is built from the pydantic models, so new config fields show up automatically:
- [sections] become expanders (optional ones get an "Include" checkbox),
- fields with a fixed set of values (Literal) become dropdowns, true/false fields checkboxes,
- every other field takes a TOML value, typed like in the file itself: 350, [0, 90, 180], "84..88:2", 2026-05-23T13:12:00,
- fields holding a dict (e.g. zone polygons) are edited as one "key = value" line per entry.
Each field shows its schema description as help text and its examples as placeholder.
Each single-line field has a "Comment" toggle for its trailing comment in the file (e.g. length = 659  # [mm] measured).
"""
import tomllib
from types import NoneType, UnionType
from typing import Any, Literal, Union, get_args, get_origin

import streamlit as st
import tomlkit
from pydantic import BaseModel
from pydantic.fields import FieldInfo

from simulation_core.config_schema import docstring_summary


EMPTY_CHOICE = ""       # entry for leaving an optional choice empty (the key is then left out of the file).
FIELDS_WITHOUT_COMMENT = {"project", "output_level"}
REQUIRED_MARK = r" \*"


# =============================================================================
# TOML text <-> values
# =============================================================================

def to_toml_text(value: Any) -> str:
    """Show a value as TOML text for a form field: 350, "84..88:2", [0, 90], or one "key = value" line per dict entry."""
    if value is None:
        return ""
    if isinstance(value, dict):
        document = tomlkit.document()
        for key, entry in value.items():
            # Nested dicts (e.g. main = {enabled = false}) stay on their line as inline tables.
            if isinstance(entry, dict):
                inline_table = tomlkit.inline_table()
                inline_table.update(entry)
                entry = inline_table
            document[key] = entry
        return tomlkit.dumps(document).strip()
    return tomlkit.item(value).as_string()


def parse_toml_text(text: str, is_dict: bool) -> Any:
    """Turn the TOML text of a form field back into a value; empty text means "not set" (None).

    Raises tomllib.TOMLDecodeError for invalid TOML, e.g. a string without quotes.
    """
    if not text.strip():
        return None
    if is_dict:
        return tomllib.loads(text)
    return tomllib.loads(f"value = {text}")["value"]


# =============================================================================
# Form
# =============================================================================

def field_options(field: FieldInfo) -> list:
    """Return the types a field accepts, e.g. [MotorConfig, NoneType] for Optional[MotorConfig] or [str] for str."""
    annotation = field.annotation
    return list(get_args(annotation)) if get_origin(annotation) in (Union, UnionType) else [annotation]


def clear_form_state(key_prefix: str) -> None:
    """Forget the form's typed-in values, so the next run shows the values from the file again."""
    for key in [key for key in st.session_state if str(key).startswith(key_prefix)]:
        del st.session_state[key]


def render_model_form(
    model_class: type[BaseModel],
    values: dict,
    key_prefix: str,
    section_path: str = "",
    depth: int = 0,
    hidden_fields: frozenset[str] = frozenset(),
    comments: dict[str, str] | None = None,
) -> tuple[dict, list[str], dict[str, str]]:
    """Show one form field per model field; return the entered values, the paths of fields with invalid TOML and the field comments.

    key_prefix makes the widget keys unique (e.g. per project and file); section_path is the TOML path shown in titles.
    hidden_fields are top-level fields the caller fills in itself (e.g. "project" in the project builder); they are left out.
    comments are the trailing comments from the file per TOML path (see config_writer.read_comments); each single-line field
    gets a "Comment" toggle, on when it has a comment. Turning it off only hides the comment; clearing the text removes it.
    """
    comments = comments or {}
    new_values = {}
    invalid_fields = []
    new_comments = {}
    if depth == 0:
        st.caption(f"{REQUIRED_MARK.strip()} = required")
    for field_name, field in model_class.model_fields.items():
        if field_name in hidden_fields:
            continue
        options = field_options(field)
        current_value = values.get(field_name)
        widget_key = f"{key_prefix}.{field_name}"
        field_path = f"{section_path}.{field_name}" if section_path else field_name
        model_option = next((option for option in options if isinstance(option, type) and issubclass(option, BaseModel)), None)
        literal_options = [option for option in options if option is not NoneType]
        # Required = no default; inside an optional section it means "required when the section is included".
        required_mark = REQUIRED_MARK if field.is_required() else ""
        label = f"{field_name}{required_mark}"

        # [section]: top-level sections as expanders; Streamlit cannot nest expanders, so deeper ones get a bordered box.
        if model_option is not None:
            box = st.expander(f"[{field_path}]{required_mark}") if depth == 0 else st.container(border=True)
            with box:
                if depth > 0:
                    st.markdown(f"**[{field_path}]**{required_mark}")
                # The fields' own help texts already show their descriptions, so the caption skips the docstring's "Attributes" section.
                st.caption(docstring_summary(model_option.__doc__))
                # Sections that may be left out (Optional or with a default) get a checkbox instead of always being written.
                if not field.is_required():
                    include = st.checkbox("Include this section", value=current_value is not None, key=f"{widget_key}.include")
                    if not include:
                        continue
                section_values, section_invalid, section_comments = render_model_form(
                    model_option, current_value or {}, widget_key, field_path, depth + 1, comments=comments
                )
            new_values[field_name] = section_values
            invalid_fields.extend(section_invalid)
            new_comments.update(section_comments)
            continue

        # Dict fields are edited as several "key = value" lines, so they have no single line for a comment; all others get a toggle,
        # except the FIELDS_WITHOUT_COMMENT.
        is_dict = isinstance(current_value, dict) or any(get_origin(option) is dict for option in literal_options)
        if is_dict or field_path in FIELDS_WITHOUT_COMMENT:
            field_column, comment_toggle_column = st.container(), None
        else:
            field_column, comment_toggle_column = st.columns([6, 1], vertical_alignment="bottom")

        with field_column:
            # Fixed set of values (Literal): dropdown, with an empty entry when the field is optional.
            if all(get_origin(option) is Literal for option in literal_options):
                choices = [choice for option in literal_options for choice in get_args(option)]
                if not field.is_required():
                    choices = [EMPTY_CHOICE, *choices]
                index = choices.index(current_value) if current_value in choices else 0
                choice = st.selectbox(label, choices, index=index, key=widget_key, help=field.description)
                new_values[field_name] = None if choice == EMPTY_CHOICE else choice

            elif literal_options == [bool]:
                # An empty form starts with the field's default, e.g. enabled = true.
                checked = current_value if current_value is not None else bool(field.default)
                new_values[field_name] = st.checkbox(label, value=checked, key=widget_key, help=field.description)

            # Everything else as TOML text: a single line, or a text area with one "key = value" line per entry for dicts.
            else:
                placeholder = "e.g. " + ", ".join(to_toml_text(example) for example in field.examples) if field.examples else None
                text_widget = st.text_area if is_dict else st.text_input
                text = text_widget(
                    label, value=to_toml_text(current_value), key=widget_key, help=field.description, placeholder=placeholder
                )
                try:
                    new_values[field_name] = parse_toml_text(text, is_dict)
                except tomllib.TOMLDecodeError as error:
                    st.error(f"Not valid TOML ({error}). Text needs quotes, e.g. \"Europe/Berlin\".")
                    invalid_fields.append(field_path)

        if comment_toggle_column is not None:
            current_comment = comments.get(field_path, "")
            with comment_toggle_column:
                show_comment = st.toggle("Comment", value=bool(current_comment), key=f"{widget_key}.show_comment")
            # A hidden comment keeps the text from the file; a shown one can be edited, and clearing it removes the comment.
            new_comments[field_path] = (
                st.text_input(
                    f"Comment for {field_name}", value=current_comment, key=f"{widget_key}.comment",
                    label_visibility="collapsed", placeholder=f"Comment for {field_name}, e.g. measured",
                )
                if show_comment
                else current_comment
            )

    # Leave empty fields out, so validation uses the field's default (or reports "Field required") instead of rejecting None.
    return {key: value for key, value in new_values.items() if value is not None}, invalid_fields, new_comments
