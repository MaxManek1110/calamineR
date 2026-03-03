
<!-- README.md is generated from README.Rmd. Please edit that file -->

# calamineR

Fast Excel file reader for R, powered by the Rust
[calamine](https://github.com/tafia/calamine) library.

## Supported Formats

- xlsx (Excel 2007+)
- xlsm (Excel with macros)
- xlsb (Excel Binary)
- xls (Excel 97-2003)
- ods (OpenDocument Spreadsheet)

## System Requirements

- R \>= 4.0
- Rust toolchain (cargo \>= 1.75.0) - install from <https://rustup.rs/>

## Installation

You can install the development version of calamineR from
[GitHub](https://github.com/) with:

``` r
# install.packages("devtools")
devtools::install_github("MaxManek1110/calamineR")
```

## Usage

``` r
library(calamineR)

# Read a sheet as data.frame
df <- read_excel("data.xlsx")
df <- read_excel("data.xlsb", sheet = "Sheet2")
df <- read_excel("data.xlsx", sheet = 2, col_names = FALSE)
df <- read_excel("data.xlsx", skip = 5)  # Skip first 5 rows

# Fill merged cells with the top-left cell value
df <- read_excel("data.xlsx", fill_merged_cells = TRUE)

# Get sheet names
sheets <- excel_sheets("data.xlsx")

# Get sheet dimensions
dims <- sheet_dims("data.xlsx", 1)
# dims["rows"], dims["cols"]

# Read as raw list of rows (for complex layouts)
rows <- read_sheet_raw("data.xlsx", "Sheet1")
```

## Functions

| Function           | Description                       |
|--------------------|-----------------------------------|
| `read_excel()`     | Read sheet as data.frame          |
| `excel_sheets()`   | Get sheet names                   |
| `sheet_dims()`     | Get sheet dimensions (rows, cols) |
| `read_sheet_raw()` | Read as list of character vectors |

## Parameters for `read_excel()`

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `path` | character | (required) | Path to the Excel file |
| `sheet` | character/integer | `1L` | Sheet name or 1-based index |
| `col_names` | logical | `TRUE` | Use first row as column names |
| `skip` | integer | `0L` | Number of rows to skip |
| `fill_merged_cells` | logical | `FALSE` | Fill merged cells with top-left value |

## Merged Cells Support

When `fill_merged_cells = TRUE`, cells that are part of a merged region are filled with the value from the top-left cell of that region. This is useful when reading spreadsheets where headers or data span multiple cells.

Supported formats: xlsx, xlsm, xlsb, xls (ods not yet supported for merged cells).

## Performance

Calamine is written in pure Rust and is significantly faster than many
alternatives:

- 1.75x faster than excelize (Go)
- 7x faster than ClosedXML (C#)
- 9x faster than openpyxl (Python)

## Development

`calamineR` is built using `extendr` and `Rust`. To build the package,
you need to have `Rust` installed on your machine.

``` r
rextendr::document()
devtools::document()
```
