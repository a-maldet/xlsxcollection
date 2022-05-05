#' @include import.R xlsxcollection_item.R
NULL

#' Create an xlsx collection item
#' 
#' This create a `xlsxcollection_item` class object, which can be used
#' in order to store a singel StyledTable together with caption,
#' footer and sheet name.
#' @param st A [StyledTable][styledTables::styled_table()] class ojbect, which
#'   should be stored.
#' @param sheet_name A string, which should be used as xlsx sheet name.
#'   In case this tables origin from a LaTeX table and you want to use
#'   the LaTeX counter later on, you must pass the LaTeX label of the table
#'   to the `sheet_name` argument. When you call
#'   [xlsxcollection_use_latex_table_counter()] the stored `sheet_name` property
#'   will be compared with the LaTeX labels found in the `aux`-file of the
#'   LaTeX report.
#' @param caption A caption string for the table.
#' @param sheet_name A string used as sheet name, when the table is written
#'   to an xlsx file.
#' @param footer A string used as table footer.
#' @export
new_xlsxcollection_item <- function(
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
    class = "xlsxcollection_item"
  )  
}

#' Validate the [xlsxcollection_item][new_xlsxcollection_item()] class object
#' 
#' @param obj The object, which should be validated.
#' @param validate_class A logical flag, defining if the class string should
#'   also be evaluated.
#' @param err_h An error handler created with [composerr][composerr::composerr()]
validate_xlsxcollection_item <- function(
  obj,
  validate_class = TRUE,
  err_h = composerr("Invalid 'xlsxcollection_item': ")
) {
  err_h <- composerr(
    before = "The passed in object is not a valid 'xlsxcollection_item': ",
    err_h,
    after = paste(
      ": Please use the function 'new_xlsxcollection_item()' in order",
      "to create a valid 'xlsxcollection_item' class object."
    )
  )
  if (isTRUE(validate_class) && !"xlsxcollection_item" %in% class(obj))
    err_h("The passed in object is not of class 'xlsxcollection_item'.")
  if (!is.list(obj))
    err_h("The object is not a list.")
  missing_entries <- c("st", "caption", "sheet_name")
  missing_entries <- missing_entries[!missing_entries %in% names(obj)]
  if (length(missing_entries) > 0)
    paste(
      "The object is missing the following list entries:\n",
      stringify(missing_entries, collapse = "\n", before = "\t", new_line = TRUE)
    ) %>%
    err_h
  non_char_entries <- intersect(c("caption", "sheet_name", "footer"), names(obj))
  non_char_entries <- sapply(
    non_char_entries,
    function(val_name) {
      val <- obj[[val_name]]
      !is.null(val) && (!is.character(val) || length(val) != 1 || is.na(val))
    }
  ) %>%
    unlist %>%
    `[`(non_char_entries, .)
  if (length(non_char_entries) > 0) 
    paste(
      "The following entries should be strings, but are not:\n",
      stringify(non_char_entries, collapse = "\n", before = "\t", new_line = TRUE)
    ) %>%
    err_h
  if (!"StyledTable" %in% class(obj$st))
    err_h("The list entry 'st' is not a 'StyledTable' class object.")
  obj
}

#' Create an `xlsxcollection_list` class object.
#' 
#' Bundles mutliple
#' [xlsxcollection_item][new_xlsxcollection_item()] into a single
#' object of class `xlsxcollection_list`.
#' @param obj One can eather pass a list of items or a singe item.
#' @param ... additional [xlsxcollection_item][new_xlsxcollection_item()] class objects.
#' @return A `xlsxcollection_list` class object holding all items.
#' @export
new_xlsxcollection_list <- function(obj, ...) {
  UseMethod("new_xlsxcollection_list")
}

#' @export
new_xlsxcollection_list.xlsxcollection_item <- function(obj, ...) {
  obj <- c(
    list(obj),
    list(...)
  ) %>%
    new_xlsxcollection_list
}

#' @export
new_xlsxcollection_list.default <- function(obj, ...) {
  validate_xlsxcollection_list(
    obj,
    validate_class = FALSE,
    err_h = composerr("Error while calling 'new_xlsxcollection_list()': ")
  ) %>%
    structure(class = "xlsxcollection_list")  
}

#' Validate the [xlsxcollection_list][new_xlsxcollection_list()] class object
#' 
#' @param obj The object, which should be validated.
#' @param validate_class A logical flag, defining if the class string should
#'   also be evaluated.
#' @param err_h An error handler created with [composerr][composerr::composerr()]
validate_xlsxcollection_list <- function(
  obj,
  validate_class = TRUE,
  err_h = composerr("Invalid 'xlsxcollection_list': ")
) {
  err_h <- composerr(
    before = "The passed in object is not a valid 'xlsxcollection_list' class object: ",
    err_h,
    after = paste(
      ": Please use the function 'new_xlsxcollection()' in order",
      "to create a valid 'xlsxcollection' class object."
    )
  )
  if (isTRUE(validate_class) && !"xlsxcollection_list" %in% class(obj))
    err_h("The passed in object is not of class 'xlsxcollection_list'.")
  if (!is.list(obj))
    err_h("The object is not a list.")
  lapply(
    seq_along(obj),
    function(i) {
      validate_xlsxcollection_item(
        obj = obj[[i]],
        err_h = composerr(
          paste(
            "The list entry at index ", stringify(i), "is invalid."
          ),
          err_h
        )
      )
    }
  )
  obj
}
