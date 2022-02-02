#' @include store_tables.R
NULL

#' Create an xlsx collection
#'
#' This function writes the read tables to an xlsx file.
#' @param xlsx_path File path of the xlsx file.
#' @param xlsx_collection_list A [xlsx_collection_list][new_xlsx_collection_list()]
#'   class object, usually created by [xlsx_collection_read_stored_tables()].
#' @param toc_header A string used as caption for the table of contents.
#' @param format_toc An optional function applying `styledTables` formatting
#'   commands to the [StyledTable][styledTables::styled_table()] class object
#'   holding the entire toc-table. If omitted, then no additional formatting
#'   is applied.
#' @return The file path given in `xlsx_path`.
#' @export
#' @seealso [xlsx_collection_list_stored_tables()],
#'   [xlsx_collection_init_store_table()],
#'   [xlsx_collection_read_stored_tables()],
#'   [xlsx_collection_use_latex_table_counter()],
#'   [xlsx_collection_append_table()],
#'   [xlsx_collection_append_toc()]
xlsx_collection_create_excel <- function(
  xlsx_collection_list,
  xlsx_path,
  toc_header,
  format_toc = NULL
) {
  wb <- xlsx::createWorkbook()
  xlsx_collection_append_toc(
    wb,
    xlsx_collection_list = xlsx_collection_list,
    toc_header = toc_header,
    format_toc = format_toc
  )
  lapply(
    xlsx_collection_list,
    function(x) {
      xlsx_collection_append_table(
        wb = wb,
        sheet_name = x$sheet_name,
        st = x$st,
        caption = x$caption,
        footer = x$footer
      )
    }
  ) %>% invisible
  xlsx::saveWorkbook(wb, xlsx_path)
  xlsx_path
}

#' Append a [xlsx_collection_item][new_xlsx_collection_item()]
#' class object to an [xlsx-workbook][xlsx::createWorkbook()]
#' 
#' @param wb An [xlsx-workbook][xlsx::createWorkbook()]
#' @param st A [StyledTable][styledTables::styled_table()] class object.
#' @param sheet_name A string used as sheet name for the table
#' @param caption A string used as table caption.
#' @param footer An optional string used as table footer.
#' @export
#' @seealso [xlsx_collection_list_stored_tables()],
#'   [xlsx_collection_init_store_table()],
#'   [xlsx_collection_read_stored_tables()],
#'   [xlsx_collection_use_latex_table_counter()],
#'   [xlsx_collection_create_excel()],
#'   [xlsx_collection_append_toc()]
xlsx_collection_append_table <- function(
  wb,
  sheet_name,
  st,
  caption,
  footer = NULL
) {
  sheet <- xlsx::createSheet(wb, sheet_name)
  matrix(c(caption, ""), ncol = 1) %>%
    styled_table %>%
    set_bold(TRUE) %>%
    set_excel_font_name("Arial") %>%
    set_excel_font_size(14) %>%
    set_excel_row_height(20, row_id = 1) %>%
    set_excel_row_height(6, row_id = 2) %>%
    write_excel(sheet, .)
  st %>%
    write_excel(sheet, ., first_row = 3)
  if (!is.null(footer))
    matrix(
      c(rep("", count_cols(st)), footer, rep("", count_cols(st) - 1)),
      byrow = TRUE,
      nrow = 2
    ) %>%
    styled_table %>%
    merge_cells(row_id = 2, col_id = seq_len(count_cols(st))) %>%
    set_excel_wrapped(TRUE, row_id = 2) %>%
    set_excel_vertical("TOP", row_id = 2) %>%
    set_excel_font_name("Arial") %>%
    set_excel_font_size(8) %>%
    set_excel_row_height(12, row_id = 1) %>%
    set_excel_row_height(150, row_id = 2) %>%
    write_excel(sheet, ., first_row = count_rows(st) + 3)
  invisible(NULL)
}

#' Append a table of contents sheet to the xlsx collection file
#' 
#' @param wb An [xlsx-workbook][xlsx::createWorkbook()]
#' @param xlsx_collection_list An 
#'   [xlsx_collection_list][new_xlsx_collection_list()] class object holding
#'   the stored tables.
#' @inheritParams xlsx_collection_create_excel
#' @export
#' @seealso [xlsx_collection_list_stored_tables()],
#'   [xlsx_collection_init_store_table()],
#'   [xlsx_collection_read_stored_tables()],
#'   [xlsx_collection_use_latex_table_counter()],
#'   [xlsx_collection_create_excel()],
#'   [xlsx_collection_append_table()]
xlsx_collection_append_toc <- function(
  wb,
  xlsx_collection_list,
  toc_header,
  format_toc = NULL
) {
  toc_sheet <- xlsx::createSheet(wb, "Inhalt")
  java_createHelper <- rJava::.jcall(
    obj = wb, 
    returnSig = "Lorg/apache/poi/ss/usermodel/CreationHelper;",
    method = "getCreationHelper"
  )
  java_link_doc <- rJava::.jfield(
    o = "org/apache/poi/common/usermodel/Hyperlink",
    name = "LINK_DOCUMENT"
  )
  matrix(c(toc_header, ""), ncol = 1) %>%
    styled_table %>%
    set_bold(TRUE) %>%
    set_excel_font_name("Arial") %>%
    set_excel_font_size(16) %>%
    set_excel_row_height(20, row_id = 1) %>%
    set_excel_row_height(6, row_id = 2) %>%
    set_excel_col_width(8, col_id = 1) %>%
    write_excel(toc_sheet, .)
  data.frame(
    sheet_name = lapply(xlsx_collection_list, `[[`, "sheet_name") %>% unlist,
    caption = lapply(xlsx_collection_list, `[[`, "caption") %>% unlist
  ) %>%
    styled_table %>%
    set_excel_font_name("Arial") %>%
    set_excel_font_size(12) %>%
    set_underline(1, col_id = 1) %>%
    set_font_color("blue", col_id = 1) %>%
    {
      if (!is.null(format_toc)) {
        format_toc(.)
      } else {
        .
      }
    } %>%
    write_excel(toc_sheet, ., first_row = 3)
  rows <- xlsx::getRows(
    toc_sheet,
    rowIndex = 2 + seq_along(xlsx_collection_list)
  )
  for(i in seq_along(xlsx_collection_list)) {
    rows <- xlsx::getRows(
      toc_sheet,
      rowIndex = 2 + i
    )
    cell <- xlsx::getCells(rows,colIndex = 1)
    java_link <- rJava::.jcall(
      java_createHelper, 
      "Lorg/apache/poi/ss/usermodel/Hyperlink;",
      "createHyperlink",
      java_link_doc
    )
    rJava::.jcall(java_link, "V", "setAddress", sprintf("'%s'!A1", xlsx_collection_list[[i]]$sheet_name))
    rJava::.jcall(
      obj = cell[[1]],
      returnSig = "V",
      method = "setHyperlink",
      java_link
    )
  }
  return(NULL)
}
