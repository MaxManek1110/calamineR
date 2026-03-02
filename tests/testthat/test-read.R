test_that("excel_sheets returns character vector", {
  test_file <- system.file("extdata", "test.xlsx", package = "calaminer")
  skip_if(test_file == "", "Test file not available")

  sheets <- excel_sheets(test_file)
  expect_type(sheets, "character")
  expect_true(length(sheets) > 0)
})

test_that("sheet_dims returns integer vector", {
  test_file <- system.file("extdata", "test.xlsx", package = "calaminer")
  skip_if(test_file == "", "Test file not available")


  dims <- sheet_dims(test_file, 1)
  expect_type(dims, "integer")
  expect_length(dims, 2)
  expect_named(dims, c("rows", "cols"))
})

test_that("read_excel returns data.frame", {
  test_file <- system.file("extdata", "test.xlsx", package = "calaminer")
  skip_if(test_file == "", "Test file not available")

  df <- read_excel(test_file)
  expect_s3_class(df, "data.frame")
})

test_that("read_sheet_raw returns list", {
  test_file <- system.file("extdata", "test.xlsx", package = "calaminer")
  skip_if(test_file == "", "Test file not available")

  rows <- read_sheet_raw(test_file, 1)
  expect_type(rows, "list")
})

test_that("invalid file path throws error", {
  expect_error(excel_sheets("nonexistent.xlsx"))
})
