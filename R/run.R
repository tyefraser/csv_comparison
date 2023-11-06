# load the excel sheet comparison function
repo <- dirname(dirname(rstudioapi::getSourceEditorContext()$path))
source(paste0(repo,"/R/source/excel_sheet_comparison.R"))



# Uses:
# Used to add in comparison between two sheets in an excel file
# wb <- compare_sheets(
#     wb = openxlsx::loadWorkbook(path_results_xlsx),
#     base_sheet_name = base_sheet_name,
#     updated_sheet_name = updated_sheet_name,
#     result_sheet_name = result_sheet_name,
#     column_config = column_config,
#     verticle = TRUE
# )

# Example:
# Values for testing:
xlsx_path <- "C:/Users/tyewf/Documents/csv_comparer_example.xlsx"
sheets_to_compare <- data.table(
    base_sheet_name = c('base_1', 'base_2'),
    updated_sheet_name = c('updated_1', 'updated_2'),
    result_sheet_name = c('result_1', 'result_2')
)
column_config <- data.table(
    base_cols_to_compare = c('words', 'unique', 'amount', 'number'),
    updated_cols_to_compare = c('strings', 'key', 'dollar', 'numeric'),
    data_type = c('character', 'numeric', 'numeric', 'numeric'),
    rounding = c(0, 0, 2, NA),
    key_column = c(FALSE, TRUE, FALSE, FALSE)
)
# run
compare_xl_sheets(
    path_results_xlsx = xlsx_path,
    sheets_to_compare = sheets_to_compare,
    verticle = FALSE,
    column_config = column_config,
    result_summary_sheet_name = 'Comparison Summary'
)


# read "C:/Users/tyewf/Documents/csv_comparer_example.xlsx" with openxlsx
# read "C:/Users/tyewf/Documents/csv_comparer_example.xlsx" with openxlsx
wb <- 

# Compare Sheets on an existing Excel WB

compare_sheets(
    wb = openxlsx::loadWorkbook(
        "C:/Users/tyewf/Documents/csv_comparer_example.xlsx"),
    base_sheet_name = 'old_version',
    updated_sheet_name = 'new_version',
    result_sheet_name = 'version_comparison',
    column_config = NULL,
    verticle = TRUE
)


