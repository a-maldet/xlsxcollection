
# xlsxcollection

<!-- badges: start -->

[![GitHub last
commit](https://img.shields.io/github/last-commit/a-maldet/xlsxcollection.svg?logo=github)](https://github.com/a-maldet/funky/commits/master)
[![GitHub code size in
bytes](https://img.shields.io/github/languages/code-size/a-maldet/xlsxcollection.svg?logo=github)](https://github.com/a-maldet/xlsxcollection)
<!-- badges: end -->

``` r
library(xlsxcollection)
```

This package is used in order to write multiple StyledTables to a single
XLSX file including a table of contents sheet. This packages uses the
`styledTalbes` and the `xlsx` package.

# Installation

``` r
# Install development version from GitHub
devtools::install_github('a-maldet/xlsxcollection', build_opts = NULL)
```

# Usage

## Storing multiple StyledTables

``` r
library(styledTables)
library(xlsx)
library(xlsxcollection)

# assume we have 3 StyledTable objects
st1 <- styled_table(matrix(11:13))
st2 <- styled_table(matrix(21:23))
st3 <- styled_table(matrix(31:33))

# init storage function
dir.create("tables_dir")
store_st <- xlsxcollection_init_store_table("./tables_dir")

# store the tables
store_st(
  st = st1,
  sheet_name = "pop",
  caption = "Population since 2010"
)
store_st(
  st = st2,
  sheet_name = "birth",
  caption = "Childs birth since 2010",
  footer = "No misscarriage included."
)
store_st(
  st = st3,
  sheet_name = "death",
  caption = "Deaths since 2010"
)

cat("Stored files:", xlsxcollection_list_stored_tables("./tables_dir"))
```

## Read stored tables and create Excel file

``` r
# read stored tables
xlsx_coll <- xlsxcollection_list_stored_tables("./tables_dir") %>%
  xlsxcollection_read_stored_tables

# !!! THE FOLLOWING CALL IS OPTIONAL !!!
#  (only necessary when the tables were originated from a LaTeX report):
# The following call reads an "aux"-file (this is a LaTeX file, which also holds
# the table counters for labelled LaTeX tables) and uses the corresponding
# LaTeX counters for the sheet names and table of contents.
xlsx_coll <- xlsx_coll %>%
  xlsxcollection_use_latex_table_counter("./my_report.aux")

# Write the tables to an Excel file
xlsx_coll %>%
  xlsxcollection_create_excel(
    xlsx_path = "./my_report.xlsx",
    toc_caption = "MY REPORT TABLES",
    toc_format = function(x) {
      x %>%
        set_excel_col_width(30, col_id = 1)
    }
  )
```

## Directly create collection (without storing) and create Excel file

``` r
# create collection object holding the tables
xlsx_coll2 <- new_xlsxcollection_list(
  new_xlsxcollection_item(
    st = st1,
    sheet_name = "pop",
    caption = "Population since 2010"
  ),
  new_xlsxcollection_item(
    st = st2,
    sheet_name = "birth",
    caption = "Childs birth since 2010",
    footer = "No misscarriage included."
  ),
  new_xlsxcollection_item(
    st = st3,
    sheet_name = "death",
    caption = "Deaths since 2010"
  )
)

# Write the tables to an Excel file
xlsx_coll2 %>%
  xlsxcollection_create_excel(
    xlsx_path = "./my_report.xlsx",
    toc_caption = "MY REPORT TABLES",
    toc_format = function(x) {
      x %>%
        set_excel_col_width(30, col_id = 1)
    }
  )
```
