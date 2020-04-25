# Read files of the Tecan Plate reader Infinite 200 into R
#
# Some useful keyboard shortcuts for package authoring:
#
#   Build and Reload Package:  'Ctrl + Shift + B'
#   Check Package:             'Ctrl + Shift + E'
#   Test Package:              'Ctrl + Shift + T'

#' @importFrom magrittr %>%

get_sheet_name <- function(xlsx_path, ...) {
  dots <- list(...)
  sheet <- ifelse(is.null(dots$sheet),
                  readxl::excel_sheets(xlsx_path)[1],
                  dots$sheet)
}

#' Reads one or all sheets of a Tecan Plate reader Infinite 200 excel file.
#' Automatically determines the table format and the number of wells that were recorded.
#' Reads all absorbance values accordingly and records all sheet names and file names.
#' Ignores but warns about empty or malformatted sheets.
#'
#' @param xlsx_path Path to the Tecan file.
#' @param sheet Name of the sheet to read (default: "all").
#' @return A tibble in long format.
#' @examples
#' read_tecan("tecan_example.xlsx")
#' @export
read_tecan <- function(xlsx_path, sheet = "all") {
  if (sheet == "all") {
    sheet <- readxl::excel_sheets(xlsx_path)
  }
  lapply(sheet, function(s)
    autoread_tecan_sheet(xlsx_path, sheet = s)) %>%
    dplyr::bind_rows() %>%
    dplyr::mutate(filename = basename(xlsx_path))
}

get_table_dims <- function(xlsx_path, start_row, ...) {
  # determine the dimensions of the table (e.g. 8x12 or 1x96)
  d <- readxl::read_excel(
    xlsx_path,
    range = cellranger::anchored(paste0("A", start_row), dim = c(100, 100)),
    .name_repair = ~ vctrs::vec_as_names(.x, repair = "unique", quiet = TRUE),
    ...
  ) %>%
    dplyr::select(`<>`, tidyselect::matches("^[A-Z]?\\d+"))

  c(d %>% dplyr::pull(1) %>% match(NA, .),
    ncol(d))
}

sniff_tecan_fmt <- function(xlsx_path, ...) {
  out_fmt <- list("table_layout" = "invalid")
  col_a_tbl <-
    readxl::read_excel(xlsx_path,
                       range = readxl::cell_cols("A"),
                       col_names = "A",
                       ...)

  if (ncol(col_a_tbl) > 0) {
    col_a <- col_a_tbl$A

    # check where the absorbance values are starting, indicated by "<>"
    table_starts <- which(col_a %in% "<>")
    if (length(table_starts) == 2) {
      out_fmt$start_blue <- table_starts[1]
      out_fmt$start_yellow <- table_starts[2]

      # check if the absorbance table is formatted like a plate (rows x columns)
      # or if there is just a long list of values for each well (A1, A2, A3...)
      next_val <- col_a[table_starts + 1]
      if (all(next_val == "Value")) {
        out_fmt$table_layout <- "well"
      } else if (all(stringr::str_detect(next_val, "^[A-Z]$"))) {
        out_fmt$table_layout <- "plate"
      }
      out_fmt$table_dim <-
        get_table_dims(xlsx_path, table_starts[1], ...)
    }
  }
  out_fmt
}

extract_absorbance_platefmt <- function(xlsx_path, range, ...) {
  readxl::read_excel(xlsx_path, range = range, ...) %>%
    dplyr::select(`<>`, tidyselect::matches("^\\d+")) %>%
    tidyr::drop_na() %>%
    dplyr::mutate_at(dplyr::vars(tidyselect::matches("^\\d+")), as.numeric) %>%
    tidyr::pivot_longer(tidyselect::matches("^\\d+"), values_to = "absorbance") %>%
    tidyr::unite(well, c("<>", "name"), sep = "")
}

extract_absorbance_wellfmt <- function(xlsx_path, range, ...) {
  readxl::read_excel(xlsx_path, range = range, ...) %>%
    tidyr::pivot_longer(
      cols = tidyselect::everything(),
      names_to = "well",
      values_to = "absorbance"
    )
}

read_tecan_sheet_platefmt <- function(xlsx_path,
                                      range_blue = cellranger::cell_limits(c(24, 1), c(32, 13)),
                                      range_yellow = cellranger::cell_limits(c(50, 1), c(58, 13)),
                                      ...) {
  blue <-
    extract_absorbance_platefmt(xlsx_path, range_blue, ...) %>%
    dplyr::rename(blue_absorbance = absorbance)

  yellow <-
    extract_absorbance_platefmt(xlsx_path, range_yellow, ...) %>%
    dplyr::rename(yellow_absorbance = absorbance)

  assemble_tbl(blue, yellow, xlsx_path, ...)
}

assemble_tbl <- function(blue, yellow, xlsx_path, ...) {
  sheet_name <- get_sheet_name(xlsx_path, ...)
  tbl <- dplyr::full_join(blue, yellow, by = "well") %>%
    dplyr::mutate(sheetname = sheet_name) %>%
    dplyr::filter(stringr::str_detect(well, "^[A-Z]\\d+"))

  stopifnot(all(stats::complete.cases(tbl)))
  tbl
}

read_tecan_sheet_wellfmt <- function(xlsx_path,
                                     range_blue = cellranger::cell_limits(c(34, 2), c(35, 97)),
                                     range_yellow = cellranger::cell_limits(c(50, 2), c(51, 97)),
                                     ...) {
  blue <- extract_absorbance_wellfmt(xlsx_path, range_blue, ...) %>%
    dplyr::rename(blue_absorbance = absorbance)

  yellow <-
    extract_absorbance_wellfmt(xlsx_path, range_yellow, ...) %>%
    dplyr::rename(yellow_absorbance = absorbance)

  assemble_tbl(blue, yellow, xlsx_path, ...)
}

autoread_tecan_sheet <- function(xlsx_path, ...) {
  tecan_format <- sniff_tecan_fmt(xlsx_path, ...)
  if (tecan_format$table_layout == "well") {
    limits_b <-
      cellranger::cell_limits(
        c(tecan_format$start_blue, 2),
        c(tecan_format$start_blue + 1, tecan_format$table_dim[2])
      )
    limits_y <-
      cellranger::cell_limits(
        c(tecan_format$start_yellow, 2),
        c(tecan_format$start_yellow + 1, tecan_format$table_dim[2])
      )
    tbl <- read_tecan_sheet_wellfmt(xlsx_path,
                                    range_blue = limits_b,
                                    range_yellow = limits_y,
                                    ...)
  } else if (tecan_format$table_layout == "plate") {
    limits_b <- cellranger::cell_limits(
      c(tecan_format$start_blue, 1),
      c(
        tecan_format$start_blue + tecan_format$table_dim[1] - 1,
        tecan_format$table_dim[2]
      )
    )
    limits_y <- cellranger::cell_limits(
      c(tecan_format$start_yellow, 1),
      c(
        tecan_format$start_yellow + tecan_format$table_dim[1] - 1,
        tecan_format$table_dim[2]
      )
    )
    tbl <- read_tecan_sheet_platefmt(xlsx_path,
                                     range_blue = limits_b,
                                     range_yellow = limits_y,
                                     ...)
  } else {
    warning(
      glue::glue(
        "Ignoring sheet '{list(...)$sheet}' of '{basename(xlsx_path)}' since it has the wrong format."
      )
    )
    tbl <- tibble::tibble()
  }
  tbl
}
