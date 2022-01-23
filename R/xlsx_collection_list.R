#' @include import.R
NULL

#' Create an xlsx collection item
#' 
#' This create a `xlsx_collection_item` class object, which can be used
#' in order to store a singel StyledTable together with caption,
#' footer and sheet name.
#' @param st A [StyledTable][styledTables::styled_table()] class ojbect, which
#'   should be stored.
#' @param sheet_name A string, which should be used as xlsx sheet name.
#'   In case this tables origin from a LaTeX table and you want to use
#'   the LaTeX counter later on, you must pass the LaTeX label of the table
#'   to the `sheet_name` argument. When you call
#'   [xlsx_collection_use_latex_table_counter()] the stored `sheet_name` property
#'   will be compared with the LaTeX labels found in the `aux`-file of the
#'   LaTeX report.
#' @param caption A caption string for the table.
#' @param sheet_name A string used as sheet name, when the table is written
#'   to an xlsx file.
#' @param footer A string used as table footer.
#' @export
new_xlsx_collection_item <- function(
  st,
  caption,
  sheet_name,
  footer = NULL
) {
  structure(
    list(
      st = st,
      caption = caption,
      sheet_name = sheet_name,
      footer = footer
    ),
    class = "xlsx_collection_item"
  )  
}

#' Validate the [xlsx_collection_item][new_xlsx_collection_item()] class object
#' 
#' @param obj The object, which should be validated.
#' @param validate_class A logical flag, defining if the class string should
#'   also be evaluated.
#' @param err_h An error handling function.
validate_xlsx_collection_item <- function(
  obj,
  validate_class = TRUE,
  err_h = composerr("Invalid 'xlsx_collection_item'")
) {
  err_h <- composerr(
    err_prior = err_h,
    text_1 = "The passed in object is not a valid 'xlsx_collection_item'",
    text_2 = paste(
      "Please use the function 'new_xlsx_collection_item()' in order",
      "to create a valid 'xlsx_collection_item' class object."
    )
  )
  if (isTRUE(validate_class) && !"xlsx_collection_item" %in% class(obj))
    err_h("The passed in object is not of class 'xlsx_collection_item'.")
  if (!is.list(obj))
    err_h("The object is not a list.")
  missing_entries <- c("st", "caption", "sheet_name", "footer")
  missing_entries <- missing_entries[!missing_entries %in% names(obj)]
  if (length(missing_entries) > 0)
    paste(
      "The object is missing the following list entries:\n",
      stringify(missing_entries, str_collapse = "\n", str_before = "\t", new_line = TRUE)
    ) %>%
    err_h
  non_char_entries <- intersect(c("caption", "sheet_name", "footer"), names(obj))
  non_char_entries <- sapply(
    non_char_entries,
    function(val_name) {
      val <- obj[[val_name]]
      !is.character(val) || length(val) != 1 || is.na(val)
    }
  ) %>%
    unlist %>%
    `[`(non_char_entries, .)
  if (length(non_char_entries) > 0) 
    paste(
      "The following entries should be strings, but are not:\n",
      stringify(non_char_entries, str_collapse = "\n", str_before = "\t", new_line = TRUE)
    ) %>%
    err_h
  if (!"StyledTable" %in% class(obj$st))
    err_h("The list entry 'st' is not a 'StyledTable' class object.")
  obj
}

#' Create an `xlsx_collection_list` class object.
#' 
#' Bundles mutliple
#' [xlsx_collection_item][new_xlsx_collection_item()] into a single
#' object of class `xlsx_collection_list`.
#' @param obj One can eather pass a list of items or a singe item.
#' @param ... additional [xlsx_collection_item][new_xlsx_collection_item()] class objects.
#' @return A `xlsx_collection_list` class object holding all items.
#' @export
new_xlsx_collection_list <- function(obj, ...) {
  UseMethod("new_xlsx_collection_list")
}

#' @export
new_xlsx_collection_list.xlsx_collection_item <- function(obj, ...) {
  obj <- c(
    obj,
    list(...)
  ) %>%
    new_xlsx_collection_list
}

#' @export
new_xlsx_collection_list.default <- function(obj, ...) {
  validate_xlsx_collection_list(
    obj,
    validate_class = FALSE,
    err_h = composerr("Error while calling 'new_xlsx_collection_list()'.")
  ) %>%
    structure(class = "xlsx_collection_list")  
}

#' Validate the [xlsx_collection_list][new_xlsx_collection_list()] class object
#' 
#' @param obj The object, which should be validated.
#' @param validate_class A logical flag, defining if the class string should
#'   also be evaluated.
#' @param err_h An error handling function.
validate_xlsx_collection_list <- function(
  obj,
  validate_class = TRUE,
  err_h = composerr("Invalid 'xlsx_collection_list'")
) {
  err_h <- composerr(
    err_prior = err_h,
    text_1 = "The passed in object is not a valid 'xlsx_collection_list' class object",
    text_2 = paste(
      "Please use the function 'new_xlsx_collection()' in order",
      "to create a valid 'xlsx_collection' class object."
    )
  )
  if (isTRUE(validate_class) && !"xlsx_collection_list" %in% class(obj))
    err_h("The passed in object is not of class 'xlsx_collection_list'.")
  if (!is.list(obj))
    err_h("The object is not a list.")
  lapply(
    seq_along(obj),
    function(i) {
      validate_xlsx_collection_item(
        obj = obj[[i]],
        err_h = paste(
          "The list entry at index ", stringify(i), "is invalid."
        ) %>%
          composerr(err_h)
      )
    }
  )
  obj
}