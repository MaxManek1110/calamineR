# Test input validation and error messages
# Note: extendr wraps Rust errors in "User function panicked" messages,
# but the actual validation messages are printed to console

test_that("non-existent file throws error", {
  expect_error(
    read_excel("/nonexistent/path/file.xlsx")
  )
  expect_error(
    excel_sheets("/nonexistent/path/file.xlsx")
  )
  expect_error(
    sheet_dims("/nonexistent/path/file.xlsx", 1)
  )
})

test_that("unsupported file format throws error", {
  tmp <- tempfile(fileext = ".csv")
  on.exit(unlink(tmp), add = TRUE)
  write.csv(mtcars, tmp, row.names = FALSE)

  expect_error(
    read_excel(tmp)
  )
  expect_error(
    excel_sheets(tmp)
  )
  expect_error(
    sheet_dims(tmp, 1)
  )
})

test_that("sheet as vector throws error", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(
    read_excel(tmp, sheet = 1:2)
  )
  expect_error(
    read_excel(tmp, sheet = c("Sheet1", "Sheet2"))
  )
  expect_error(
    sheet_dims(tmp, sheet = 1:3)
  )
})

test_that("empty sheet argument throws error", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(
    read_excel(tmp, sheet = character(0))
  )
  expect_error(
    read_excel(tmp, sheet = integer(0))
  )
})

test_that("sheet name not found throws error", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(
    read_excel(tmp, sheet = "NonExistentSheet")
  )
})

test_that("sheet index out of range throws error", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(
    read_excel(tmp, sheet = 99)
  )
  expect_error(
    sheet_dims(tmp, sheet = 99)
  )
})

test_that("sheet index less than 1 throws error", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(
    read_excel(tmp, sheet = 0)
  )
  expect_error(
    read_excel(tmp, sheet = -1)
  )
})

test_that("negative skip throws error", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)

  expect_error(
    read_excel(tmp, skip = -1)
  )
})

test_that("valid inputs still work correctly", {
  skip_if_not_installed("writexl")
  tmp <- tempfile(fileext = ".xlsx")
  on.exit(unlink(tmp), add = TRUE)
  writexl::write_xlsx(mtcars, tmp)


  # Numeric sheet index

  df1 <- read_excel(tmp, sheet = 1)
  expect_s3_class(df1, "data.frame")
  expect_equal(nrow(df1), nrow(mtcars))


  # Sheet name

  sheets <- excel_sheets(tmp)
  df2 <- read_excel(tmp, sheet = sheets[1])
  expect_s3_class(df2, "data.frame")

  # Sheet dims

  dims <- sheet_dims(tmp, sheet = 1)
  expect_length(dims, 2)

  # Skip parameter

  df3 <- read_excel(tmp, skip = 5)
  expect_equal(nrow(df3), nrow(mtcars) - 5)
})

test_that("merge_regions validates inputs", {
  expect_error(
    merge_regions("/nonexistent/file.xlsx", 1)
  )

  tmp <- tempfile(fileext = ".csv")
  on.exit(unlink(tmp), add = TRUE)
  write.csv(mtcars, tmp, row.names = FALSE)

  expect_error(
    merge_regions(tmp, 1)
  )
})
