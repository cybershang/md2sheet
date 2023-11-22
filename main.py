"""Module provide the ability to convert markdown into xlsx sheet."""
import sys
import os
import markdown_to_json
import pandas as pd
import typer

app = typer.Typer()


def flatten_structure(key, value, result, current_level):
    """Recursive function, for flat nested dict."""
    if isinstance(value, dict):
        for sub_key, sub_value in value.items():
            flatten_structure(sub_key, sub_value, result, current_level + [key])
    elif isinstance(value, list):
        for item in value:
            add_to_result(result, current_level + [key], item)
    else:
        add_to_result(result, current_level + [key], value)


def add_to_result(result, keys, value):
    """Add key and value into the relevant level."""
    for i, key in enumerate(keys, start=1):
        result[f"h{i}"].append(key)
    result[f"h{len(keys) + 1}"].append(value)


def falten_nest_dict(nested_dict: dict) -> dict:
    """Convert nested dict into flat structure."""
    result = {f"h{i}": [] for i in range(1, 6 + 1)}

    for key, value in nested_dict.items():
        flatten_structure(key, value, result, [])

    return result


@app.command()
def md2xlsx(
    in_file: str = typer.Argument(..., help="Enter the markdown filepath"),
    out_file: str = typer.Argument(..., help="Enter the xlsx filepath"),
):
    """Convert markdown into xlsx sheet."""
    if os.path.exists(in_file):
        with open(in_file, "r", encoding="utf-8") as f:
            md_text = f.read()
    else:
        print("input file not exist!")
        sys.exit(1)

    flat_dict = falten_nest_dict(markdown_to_json.dictify(md_text))
    df = pd.DataFrame(
        {
            "h1": flat_dict["h1"],
            "h2": flat_dict["h2"],
            "h3": flat_dict["h3"],
            "h4": flat_dict["h4"],
            "h5": flat_dict["h5"],
        }
    )
    df.to_excel(out_file, index=False)


if __name__ == "__main__":
    app()
