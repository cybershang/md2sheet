"""Module provide the ability to convert markdown into xlsx sheet."""
import sys
import os
from typing import Optional
import markdown_to_json
import pandas as pd
import typer
from typing_extensions import Annotated

__version__ = "0.1.0"
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


def output_version():
    """Output version number"""
    print(f"md2sheet version: {__version__}")
    raise typer.Exit(0)


@app.command()
def md2xlsx(
    in_file: str = typer.Argument(default=None, help="Enter the markdown filepath"),
    out_file: str = typer.Argument(default=None, help="Enter the xlsx filepath"),
    version: Annotated[
        Optional[bool], typer.Option("--version")
    ] = None,
):
    """Convert markdown into xlsx sheet."""
    if version:
        output_version()

    elif in_file and out_file:
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
