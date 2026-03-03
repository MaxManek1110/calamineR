# calamineR 0.1.1

* Added `fill_merged_cells` parameter to `read_excel()` - when TRUE, fills merged cells with the value from the top-left cell of the merged region
* Merged cell support for xlsx, xlsm, xlsb, and xls formats
* Package renamed from calaminer to calamineR

# calamineR 0.1.0

* Initial CRAN submission
* Support for xlsx, xlsm, xlsb, xls, and ods formats
* Functions: `read_excel()`, `excel_sheets()`, `sheet_dims()`, `read_sheet_raw()`
* Automatic column type detection: numeric, logical, Date, and character types are inferred from cell data
