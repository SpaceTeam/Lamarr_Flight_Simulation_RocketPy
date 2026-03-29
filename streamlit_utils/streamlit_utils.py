from simulation.utils import *
# ----------------------------
# Read config without parsing
# ----------------------------
def read_config_raw(path):

    lines = []
    constants = {}
    variables = {}

    with open(path) as f:

        for raw_line in f:

            stripped = raw_line.strip()

            # blank line
            if not stripped:
                lines.append({"type": "blank"})
                continue

            # full comment
            if stripped.startswith("#"):
                lines.append({
                    "type": "comment",
                    "text": raw_line
                })
                continue

            # split inline comment
            if "#" in raw_line:
                code, comment = raw_line.split(" ", 1)
                comment = comment
            else:
                code = raw_line
                comment = ""

            code = code

            if "=" not in code:
                lines.append({
                    "type": "comment",
                    "text": raw_line
                })
                continue

            name, value_str = map(str, code.split("=", 1))

            parsed = parse_value(value_str)

            entry = {
                "type": "param",
                "name": name,
                "value": value_str,
                "inline_comment": comment
            }

            if isinstance(parsed, list):
                entry["type"] = "variable"
            else:
                entry["type"] = "constant"

            lines.append(entry)

    return lines, constants, variables


# ----------------------------
# Write config exactly as text
# ----------------------------
def write_config_raw(path, lines):

    with open(path, "w") as f:

        for entry in lines:

            if entry["type"] == "blank":
                f.write("\n")

            elif entry["type"] == "comment":
                f.write(entry["text"])

            elif entry["type"] == "constant" or entry["type"] == "variable" :

                line = f'{entry["name"]}={entry["value"]}'

                if entry["inline_comment"]:
                    line += " " + entry["inline_comment"]

                f.write(line)


# ----------------------------
# Validate parameters
# ----------------------------
def validate_params(lines):
    errors = []

    for entry in lines:

        if entry["type"] != "param":
            continue

        name = entry["name"]
        value = entry["value"]

        try:
            parse_value(value)

        except Exception as e:
            errors.append((name, str(e)))

    return errors

