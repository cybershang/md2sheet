<p align="center">

<img src="https://github.com/shangcode/md2sheet/raw/main/docs/img/md2sheet-logo.svg"/>

</p>

<p align="center">
<a href="https://pypi.org/project/md2sheet/"><img alt="PyPI - Version" src="https://img.shields.io/pypi/v/md2sheet"></a>
<a href="https://pypistats.org/packages/md2sheet"><img alt="PyPI - Downloads" src="https://img.shields.io/pypi/dw/md2sheet"></a>
<a href="https://opensource.org/licenses/MIT"><img src="https://img.shields.io/badge/License-MIT-yellow.svg"></a>
</p>

<p align="center">
<a href="https://pypi.org/project/md2sheet/"><img alt="PyPI - Python Version" src="https://img.shields.io/pypi/pyversions/md2sheet"></a>
<a href="https://md2sheet.slack.com/"><img alt="Static Badge" src="https://img.shields.io/badge/slack---?style=social&logo=slack"></a>
</p>

<p align="center">
<a href="https://github.com/shangcode/md2xlsx/actions/workflows/pylint.yml"><img src="https://github.com/shangcode/md2xlsx/actions/workflows/pylint.yml/badge.svg?branch=main"></a>
<a href="https://github.com/shangcode/md2sheet/actions/workflows/scorecard.yml"><img src="https://github.com/shangcode/md2sheet/actions/workflows/scorecard.yml/badge.svg?branch=main"></a>
<a href="https://securityscorecards.dev/viewer/?uri=github.com/shangcode/md2sheet"><img src="https://api.securityscorecards.dev/projects/github.com/shangcode/md2sheet/badge"></a>
</p>
Convert structed markdown content into sheet(XLSX).

## Is this for me?

If you want convert a markdown like this:

```md
# Comics
## Detective Comics (DC)
### Batman Family
#### Batman
...
#### Catwoman
...

## Marvel Comics
### Avengers
#### Iron Man
...
#### Captain America
...
```

into sheet like this:
|h1|h2|h3|h4|h5|
|--|--|--|--|--|
|Comics|Detective Comics (DC)|BatmanFamily|Batman|...|
|Comics|Detective Comics (DC)|BatmanFamily|Cat woman|...|
|Comics|Marvel Comics|Avengers|Iron Man|...|
|Comics|Marvel Comics|Avengers|Captain America|...|

then you are on the right place.

💡Futhuremore, you can use office tools to modifiy the sheet according to your preferences:
![use office tools to modify sheet](https://github.com/shangcode/md2sheet/raw/main/docs/img/modifed-with-office-tools.png)

## Requirements
Python 3.11+

The md2Sheet stands on the shoulders of giants:

- [markdown-to-json](https://github.com/njvack/markdown-to-json/)
- [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/)
- [Typer](https://typer.tiangolo.com/)

## Installation
```console
pip install md2sheet
```

## Usage

```console
md2xlsx [OPTIONS] IN_FILE OUT_FILE
```

**Options**:

- `--install-completion`: Install completion for the current shell.
- `--show-completion`: Show completion for the current shell, to copy it or customize the installation.
- `--help`: Show this message and exit.

**Arguments**:

- `IN_FILE`: Enter the markdown filepath [required]
- `OUT_FILE`: Enter the xlsx filepath [required]

## License

This project is licensed under the terms of the MIT license.

## Contract
<a href="https://md2sheet.slack.com"><img alt="Static Badge" src="https://img.shields.io/badge/slack---?style=for-the-badge&logo=slack"></a>

