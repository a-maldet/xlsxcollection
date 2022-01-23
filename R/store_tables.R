#' @include import.R
#' @include xlsx_collection_list.R
NULL

#' Initialize the function used for storing the tables.
#' 
#' @param storage_dir A string holding the path to a temporary folder where
#'   the [StyledTable][styledTables::styled_table()] objects should be stored.
#' @return A function, used for storing [styledTable][styledTables::styled_table()]
#'   objects. With each call of the function a single table is stored. The function
#'   takes the following arguments:
#'   - `st`: The [StyledTable][styledTables::styled_table()] class object,
#'     which should be stored.
#'   - `sheet_name`: A string, which should be used as xlsx sheet name.
#'     In case this tables origin from a LaTeX table and you want to use
#'     the LaTeX counter later on, you must pass the LaTeX label of the table
#'     to the `sheet_name` argument. When you call
#'     [xlsx_collection_use_latex_table_counter()] the stored `sheet_name` property
#'     will be compared with the LaTeX labels found in the `aux`-file of the
#'     LaTeX report.
#'   - `caption`: A string, which should be used as table caption.
#'   - `footer`: A string, which should be used as table footer.
#' @seealso [xlsx_collection_list_stored_tables()],
#'   [xlsx_collection_read_stored_tables()],
#'   [xlsx_collection_use_latex_table_counter()],
#'   [xlsx_collection_create_excel()]
#' @export
xlsx_collection_init_store_table <- function(storage_dir) {
  function(st, sheet_name, caption, footer = NULL) {
    if (!dir.exists(storage_dir))
      dir.create(storage_dir)
    list(
      st = st,
      sheet_name = sheet_name,
      caption = caption,
      footer = footer
    ) %>%
      saveRDS(file = file.path(storage_dir, paste0(sheet_name, ".rds")))
  }
}

#' List all stored tables
#' 
#' @inheritParams xlsx_collection_init_store_table
#' @return A character vector holding the file paths of all tables stored in
#'   the directory.
#' @export
#' @seealso [xlsx_collection_init_store_table()],
#'   [xlsx_collection_read_stored_tables()],
#'   [xlsx_collection_use_latex_table_counter()],
#'   [xlsx_collection_create_excel()]
xlsx_collection_list_stored_tables <- function(
  storage_dir
) {
  list.files(storage_dir, pattern = "*.rds", full.names = TRUE)
}

#' Read multiple stored tables
#' 
#' @param stored_tables A character vector holding the file paths to all
#'   stored tables, which should be read.
#' @return A [xlsx_collection_list][new_xlsx_collection_list()] class object,
#'   holding all stored tables.
#' @export
#' @seealso [xlsx_collection_list_stored_tables()],
#'   [xlsx_collection_init_store_table()],
#'   [xlsx_collection_use_latex_table_counter()],
#'   [xlsx_collection_create_excel()]
xlsx_collection_read_stored_tables <- function(stored_tables) {
  err_h = composerr("Error while calling 'xlsx_collection_read_stored_tables()'")
  lapply(
    stored_tables,
    function(file_path) {
      tryCatch(
        readRDS(file_path),
        error = function(e)
          paste0("Could not read file\n\t", stringify(file_path), "\n", "Reason: ", e) %>%
          err_h
      )
    }
  ) %>%
    new_xlsx_collection_list
}

#' Update the xlsx-collection by using the generated LaTeX table counter.
#' 
#' This function is only needed when the stored tables were part of a generated
#' LaTeX report (rnw-file) and the numbering of resulting pdf (created by LaTeX)
#' should also be used for the xlsx collection.
#' This function modifies an 
#' [xlsx_collection_list][new_xlsx_collection_list()] class object, such that
#' the defined sheet names and sheet ordering uses the table counter
#' created by LaTeX.
#' @param xlsx_collection_list A [xlsx_collection_list][new_xlsx_collection_list()]
#'   class object, usually created by [xlsx_collection_read_stored_tables()].
#' @param aux_path A string holding the path to the `*.aux` file generated
#'   by LaTeX when calling `pdflatex` (or `lualatex` etc.).
#' @return A modified  [xlsx_collection_list][new_xlsx_collection_list()]
#'   class object.
#' @export
#' @seealso [xlsx_collection_list_stored_tables()],
#'   [xlsx_collection_init_store_table()],
#'   [xlsx_collection_read_stored_tables()],
#'   [xlsx_collection_create_excel()]
xlsx_collection_use_latex_table_counter <- function(
  xlsx_collection_list,
  aux_path
) {
  err_h <- composerr("Error while calling 'xlsx_collection_use_latex_table_counter'")
  validate_xlsx_collection_list(
    xlsx_collection_list,
    err_h = composerr("Passed in 'xlsx_collection_list' is invalid")
  )
  aux_txt <- tryCatch(
    readLines(aux_path, warn = FALSE),
    error = function(e) paste0(
      "Could not read the aux-file:\n\t",
      stringify(aux_path),
      "\nReason: ", e
    ) %>%
      err_h
  )
  aux_txt <- aux_txt[grepl("^\\\\newlabel\\{.+\\}", aux_txt)]
  aux_entries <- data.frame(
    label = gsub("^\\\\newlabel\\{(.+?)\\}.+$", "\\1", aux_txt, perl = TRUE),
    ref = gsub("^\\\\newlabel\\{(.+?)\\}\\{\\{(.+?)\\}.+$", "\\2", aux_txt, perl = TRUE)
  )
  xlsx_collection_list <- lapply(
    xlsx_collection_list,
    function(x) {
      x$latex_ref <- aux_entries[aux_entries$label== x$sheet_name, "ref"] %>%
        as.numeric
      if(length(x$latex_ref) == 0 || (length(x$latex_ref) == 1 && is.na(x$latex_ref))) {
        paste0(
          "No LaTex table found for the following stored table:",
          "\n\t-> with LaTeX-label (sheet_name): ", stringify(x$sheet_name),
          "\n\t-> with table caption: ", stringify(x$caption),
          "\nPlease ensure that every stored table has an unique latex label!\n"
        ) %>%
          err_h
      } else if (length(x$latex_ref) > 1) {
        paste0(
          "Multiple LaTeX-tables with the same LaTeX-label were found in the aux-file:",
          "\n\t-> found LaTeX-tables: ", stringify(x$latex_ref),
          "\n\t-> for LaTeX-label: ", stringify(x$sheet_name),
          "\n\t-> with table caption: ", stringify(x$caption),
          "\nPlease ensure that every assigned latex label is unique!\n"
        ) %>%
          err_h
      }
      x
    }
  )
  xlsx_collection_list <- xlsx_collection_list %>%
    lapply(function(x) {
      x$sheet_name <- paste0("T.", x$latex_ref)
      x$caption <- paste0("Tabelle ", x$latex_ref, ": ", x$caption)
      x
    })
  xlsx_collection_list[order(sapply(xlsx_collection_list, "[[", "latex_ref"))] %>%
    new_xlsx_collection_list
}