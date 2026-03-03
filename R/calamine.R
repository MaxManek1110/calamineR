#' Read Excel File Using Calamine
#'
#' Fast Excel reader powered by the Rust calamine library.
#' Supports xlsx, xlsm, xlsb, xls, and ods formats.
#'
#' @param path Path to the Excel file
#' @param sheet Sheet name (character) or index (integer, 1-based). Default: 1
#' @param col_names Use first row as column names. Default: TRUE
#' @param skip Number of rows to skip before reading. Default: 0
#' @param fill_merged_cells If TRUE, fill merged cells with the value from the
#'   top-left cell of the merged region. Default: FALSE
#'
#' @return A data.frame
#' @export
#'
#' @examples
#' # Using package test file
#' test_file <- system.file("extdata", "test.xlsx", package = "calamineR")
#' if (nzchar(test_file)) {
#'   df <- read_excel(test_file)
#'   head(df)
#'
#'   # Read specific sheet by index
#'   df <- read_excel(test_file, sheet = 1)
#'
#'   # Skip header row
#'   df_no_header <- read_excel(test_file, col_names = FALSE)
#'
#'   # Fill merged cells with their value
#'   df_filled <- read_excel(test_file, fill_merged_cells = TRUE)
#' }
read_excel <- function(path, sheet = 1L, col_names = TRUE, skip = 0L,
                       fill_merged_cells = FALSE) {
  path <- normalizePath(path, mustWork = TRUE)
  cal_read_sheet_df(path, sheet, col_names, as.integer(skip), fill_merged_cells)
}

#' Get Sheet Names from Excel File
#'
#' @param path Path to the Excel file
#' @return Character vector of sheet names
#' @export
#'
#' @examples
#' test_file <- system.file("extdata", "test.xlsx", package = "calamineR")
#' if (nzchar(test_file)) {
#'   sheets <- excel_sheets(test_file)
#'   print(sheets)
#' }
excel_sheets <- function(path) {
  path <- normalizePath(path, mustWork = TRUE)
  cal_sheet_names(path)
}

#' Get Sheet Dimensions
#'
#' @param path Path to the Excel file
#' @param sheet Sheet name or index (1-based)
#' @return Named integer vector with "rows" and "cols"
#' @export
#'
#' @examples
#' test_file <- system.file("extdata", "test.xlsx", package = "calamineR")
#' if (nzchar(test_file)) {
#'   dims <- sheet_dims(test_file, 1)
#'   print(dims)  # Named vector: rows, cols
#' }
sheet_dims <- function(path, sheet = 1L) {
  path <- normalizePath(path, mustWork = TRUE)
  dims <- cal_sheet_dims(path, sheet)
  names(dims) <- c("rows", "cols")
  dims
}

#' Read Sheet as Raw Rows
#'
#' Returns sheet data as a list of character vectors (one per row).
#' Useful for complex layouts where data.frame structure doesn't fit.
#'
#' @param path Path to the Excel file
#' @param sheet Sheet name or index (1-based)
#' @return List of character vectors
#' @export
#'
#' @examples
#' test_file <- system.file("extdata", "test.xlsx", package = "calamineR")
#' if (nzchar(test_file)) {
#'   rows <- read_sheet_raw(test_file, 1)
#'   # Returns list of character vectors (one per row)
#'   head(rows, 3)
#' }
read_sheet_raw <- function(path, sheet = 1L) {
  path <- normalizePath(path, mustWork = TRUE)
  cal_read_sheet(path, sheet)
}
