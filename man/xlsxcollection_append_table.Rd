% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/write_to_excel.R
\name{xlsxcollection_append_table}
\alias{xlsxcollection_append_table}
\title{Append a \link[=new_xlsxcollection_item]{xlsxcollection_item}
class object to an \link[xlsx:Workbook]{xlsx-workbook}}
\usage{
xlsxcollection_append_table(wb, sheet_name, st, caption, footer = NULL)
}
\arguments{
\item{wb}{An \link[xlsx:Workbook]{xlsx-workbook}}

\item{sheet_name}{A string used as sheet name for the table}

\item{st}{A \link[styledTables:StyledTable-styled_table-method]{StyledTable} class object.}

\item{caption}{A string used as table caption.}

\item{footer}{An optional string used as table footer.}
}
\description{
Append a \link[=new_xlsxcollection_item]{xlsxcollection_item}
class object to an \link[xlsx:Workbook]{xlsx-workbook}
}
\seealso{
\code{\link[=xlsxcollection_list_stored_tables]{xlsxcollection_list_stored_tables()}},
\code{\link[=xlsxcollection_init_store_table]{xlsxcollection_init_store_table()}},
\code{\link[=xlsxcollection_read_stored_tables]{xlsxcollection_read_stored_tables()}},
\code{\link[=xlsxcollection_use_latex_table_counter]{xlsxcollection_use_latex_table_counter()}},
\code{\link[=xlsxcollection_create_excel]{xlsxcollection_create_excel()}},
\code{\link[=xlsxcollection_append_toc]{xlsxcollection_append_toc()}}
}
