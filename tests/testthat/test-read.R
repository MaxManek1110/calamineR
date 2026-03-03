test_that("excel_sheets returns character vector", {
  test_file <- system.file("extdata", "test.xlsx", package = "calamineR")
  skip_if(test_file == "", "Test file not available")

  sheets <- excel_sheets(test_file)
  expect_type(sheets, "character")
  expect_true(length(sheets) > 0)
})

test_that("sheet_dims returns integer vector", {
  test_file <- system.file("extdata", "test.xlsx", package = "calamineR")
  skip_if(test_file == "", "Test file not available")

  dims <- sheet_dims(test_file, 1)
  expect_type(dims, "integer")
  expect_length(dims, 2)
  expect_named(dims, c("rows", "cols"))
})

test_that("read_excel returns data.frame", {
  test_file <- system.file("extdata", "test.xlsx", package = "calamineR")
  skip_if(test_file == "", "Test file not available")

  df <- read_excel(test_file)
  expect_s3_class(df, "data.frame")
})

test_that("read_sheet_raw returns list", {
  test_file <- system.file("extdata", "test.xlsx", package = "calamineR")
  skip_if(test_file == "", "Test file not available")

  rows <- read_sheet_raw(test_file, 1)
  expect_type(rows, "list")
})

test_that("invalid file path throws error", {
  expect_error(excel_sheets("nonexistent.xlsx"))
})

test_that("fill_merged_cells fills horizontal merge", {
  test_file <- system.file("extdata", "test_merged.xlsx", package = "calamineR")
  skip_if(test_file == "", "Test merged file not available")

  df <- read_excel(test_file, col_names = FALSE, fill_merged_cells = TRUE)
  # A1:C1 should all have same value
  expect_equal(df[1, 1], df[1, 2])
  expect_equal(df[1, 1], df[1, 3])
})

test_that("fill_merged_cells fills vertical merge", {
  test_file <- system.file("extdata", "test_merged.xlsx", package = "calamineR")
  skip_if(test_file == "", "Test merged file not available")

  df <- read_excel(test_file, col_names = FALSE, fill_merged_cells = TRUE)
  # A3:A5 should all have same value
  expect_equal(df[3, 1], df[4, 1])
  expect_equal(df[3, 1], df[5, 1])
})

test_that("fill_merged_cells fills block merge", {
  test_file <- system.file("extdata", "test_merged.xlsx", package = "calamineR")
  skip_if(test_file == "", "Test merged file not available")

  df <- read_excel(test_file, col_names = FALSE, fill_merged_cells = TRUE)
  # D2:E4 should all have same value
  expect_equal(df[2, 4], df[2, 5])
  expect_equal(df[2, 4], df[3, 4])
  expect_equal(df[2, 4], df[4, 5])
})

test_that("fill_merged_cells default is FALSE", {
  test_file <- system.file("extdata", "test_merged.xlsx", package = "calamineR")
  skip_if(test_file == "", "Test merged file not available")

  df_no_fill <- read_excel(test_file, col_names = FALSE)
  df_with_fill <- read_excel(
    test_file,
    col_names = FALSE,
    fill_merged_cells = TRUE
  )

  # With default (no fill), non-origin cells should differ from origin
  # D2:E4 - D2 has value "Block", E2 should have "e1" originally
  expect_false(identical(df_no_fill[2, 4], df_no_fill[2, 5]))

  # With fill, they should be equal
  expect_equal(df_with_fill[2, 4], df_with_fill[2, 5])
})

test_that("fill_merged_cells works with skip parameter", {
  test_file <- system.file("extdata", "test_merged.xlsx", package = "calamineR")
  skip_if(test_file == "", "Test merged file not available")

  df <- read_excel(
    test_file,
    col_names = FALSE,
    skip = 2,
    fill_merged_cells = TRUE
  )
  # After skipping 2 rows, vertical merge A3:A5 becomes rows 1:3
  expect_equal(df[1, 1], df[2, 1])
  expect_equal(df[1, 1], df[3, 1])
})

test_that("fill_merged_cells works with xlsb format", {
  test_file <- system.file("extdata", "test_merged.xlsb", package = "calamineR")
  skip_if(test_file == "", "Test merged xlsb file not available")

  df <- read_excel(test_file, col_names = FALSE, fill_merged_cells = TRUE)

  # Horizontal merge A1:C1
  expect_equal(df[1, 1], df[1, 2])
  expect_equal(df[1, 1], df[1, 3])

  # Vertical merge A3:A5
  expect_equal(df[3, 1], df[4, 1])
  expect_equal(df[3, 1], df[5, 1])

  # Block merge D2:E4
  expect_equal(df[2, 4], df[2, 5])
  expect_equal(df[2, 4], df[3, 4])
})
