# Mic Tagger

Mic Tagger generates printable microphone tag SVGs from a spreadsheet.

It reads an exported Excel file, asks which sheet and row to start from, then creates numbered tag artwork for one-cast or two-cast productions. The generated files include individual SVG tags and a printable HTML grid.

## Requirements

- `uv`
- A spreadsheet exported as `.xlsx`
- A browser for opening the generated print page

## Prepare The Spreadsheet

Start from the template spreadsheet:

[Mic Tagger template](https://docs.google.com/spreadsheets/d/1Ll1G1Qmv-pMB3NbNWADFaHzB3je_3KfsqAQJo5kN6DY)

Make your own copy, fill in the cast/tag information, or use an existing mute sheet that was created from the template. Put the exported `.xlsx` file in the project folder before running the tool.

For one-cast sheets, the script uses the character name from the second spreadsheet column and the performer name from the third column.

For two-cast sheets, the script uses the character name from the second column, performer names from the third and fourth columns, and the third/fourth column headers as cast labels.

## Setup

Install `uv`:

```sh
curl -LsSf https://astral.sh/uv/install.sh | sh
```

Then clone and set up the project:

```sh
git clone https://github.com/thesamgordon/mic-tagger/
cd mic-tagger
uv sync
```

## Run

After placing your `.xlsx` file in the project folder, run:

```sh
uv run python main.py
```

The script will ask you to choose:

- which `.xlsx` file to use, if there is more than one
- which sheet to read
- which row to start from
- whether the sheet is for one cast or two casts
- the title to print on each tag

When asked for the start row, enter the row number where the usable sheet data begins. The default is `1`.

## Output

Generated files are written to:

```text
output/
```

Open this file in a browser to review or print the tag grid:

```text
output/index.html
```

Each run clears the existing contents of `output/` before creating the new tags.

## Notes

- The project includes SVG templates for one-cast, thin one-cast, and two-cast tags.
- If a two-cast row has the same performer for both casts, the script creates a one-cast tag for that row.
- Long names are not automatically resized, so check `output/index.html` before printing.
