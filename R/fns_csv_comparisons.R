
# Values for testing:
xlsx_path <- "C:/Users/tyewf/Documents/csv_comparer_example.xlsx"
sheets_to_compare <- data.table(
    base_sheet_name = c('base_1', 'base_2'),
    updated_sheet_name = c('updated_1', 'updated_2'),
    result_sheet_name = c('result_1', 'result_2')
)





compare_xl_sheets <- function(
    path_results_xlsx,
    base_sheet_name,
    updated_sheet_name,
    result_sheet_name
) {
    print(paste0(
        "Comparing Excel sheets: ", base_sheet_name, " and ",
        updated_sheet_name))

    # Read the Excel file
    wb <- openxlsx::loadWorkbook(path_results_xlsx)

    # Excel Validation - Check if the required sheets are in the xlsx
    expected_sheet_names <- c(
        sheets_to_compare$base_sheet_name,
        sheets_to_compare$updated_sheet_name)    
    if (!xlsx_validation(
    xlsx_path = path_results_xlsx,
    expected_sheet_names = expected_sheet_names
    )) {
        print(paste0("Excel validation failed for: ", path_results_xlsx))
        print(paste0(
            "Required sheets: ", base_sheet_name, " and ", updated_sheet_name
        ))
        return()
    }

    # Remove the result sheets if any exist
    remove_sheets <- function(wb, sheets_to_remove)
    
    
    for(result_sheet_name %in% sheets_to_compare$result_sheet_name) {
        if(result_sheet_name %in% openxlsx::sheets(wb)) {
            openxlsx::removeWorksheet(wb, result_sheet_name)
            print(paste0("Removed existing result sheet: ", result_sheet_name))
        }
    }
    

    # Add the result sheet
    openxlsx::addWorksheet(wb, result_sheet_name)

    # Read the base and updated sheets
    result_sheet_name

    negStyle <- openxlsx::createStyle(fontColour = "#9C0006", bgFill = "#FFC7CE")
    posStyle <- openxlsx::createStyle(fontColour = "#006100", bgFill = "#C6EFCE")

}


#' Get the names of sheets in an Excel file.
#'
#' This function reads an Excel file and returns a list of sheet names present in
#' the Excel workbook.
#'
#' @param xlsx_path A character string representing the path to the Excel file.
#'
#' @return A character vector containing the names of sheets in the Excel file.
#'
#' @import openxlsx
#'
#' @examples
#' xlsx_path <- "path/to/your/excel_file.xlsx"
#' sheets <- get_excel_sheets(xlsx_path)
#' print(sheets)
#'
#' @export
get_excel_sheets <- function(xlsx_path) {
  message(paste0("Getting Excel details for: ", xlsx_path))
  wb <- openxlsx::loadWorkbook(xlsx_path)
  return(openxlsx::sheets(wb))
}



#' Validate the presence of expected sheets in an Excel file.
#'
#' This function validates whether the expected sheet names are present in
#' the specified Excel file. It compares the list of expected sheet names
#' with the actual sheet names found in the Excel file.
#'
#' @param xlsx_path A character string representing the path to the Excel file to validate.
# @param expected_sheet_names A character vector containing the expected sheet names.
# @return A logical value: TRUE if all expected sheets are found, FALSE if any sheet is missing.
# @export
# @examples
# xlsx_path <- "path/to/your/excel_file.xlsx"
# expected_sheets <- c("Sheet1", "Sheet2", "Sheet3")
# result <- xlsx_validation(xlsx_path, expected_sheets)
# if (result) {
#   print("All expected sheets are present.")
# } else {
#   print("Some expected sheets are missing.")
# }
xlsx_validation <- function(
    xlsx_path,
    expected_sheet_names
) {
    xlsx_sheets <- get_excel_sheets(xlsx_path)

    # Check if the expected sheet names are in the xlsx_sheets
    result <- TRUE
    for (sheet in expected_sheet_names) {
        if (!(sheet %in% xlsx_sheets)) {
            result <- FALSE
            message(paste0("Sheet: ", sheet, " not found in: ", xlsx_path))
        }
    }
    if(result) {
        message(paste0("All expected sheets found in: ", xlsx_path))
    } else {
        message(paste0("Some expected sheets missing in: ", xlsx_path))
    }
    return(result)
}


#' Remove specified sheets from an open Excel workbook.
#'
#' This function takes an open Excel workbook and a list of sheet names to remove.
#'
#' @param wb An open Excel workbook loaded using openxlsx::loadWorkbook().
# @param sheets_to_remove A character vector containing sheet names to remove.
# @return The modified Excel workbook with the specified sheets removed.
# @export
# @import openxlsx
# @examples
# # Load the workbook
# wb <- openxlsx::loadWorkbook("path_to_your_workbook.xlsx")
# 
# # Define sheets to remove
# sheets_to_remove <- c("Sheet1", "Sheet2", "Sheet3")
# 
# # Remove sheets and get the modified workbook
# wb <- remove_sheets(wb, sheets_to_remove)
# 
# # Continue working with the modified workbook or save it as needed.
remove_sheets <- function(wb, sheets_to_remove) {
  # Get the list of existing sheet names in the workbook
  existing_sheets <- openxlsx::sheets(wb)
  
  # Iterate through the sheets to remove
  for (sheet_name in sheets_to_remove) {
    if (sheet_name %in% existing_sheets) {
      openxlsx::removeWorksheet(wb, sheet_name)
      cat("Removed sheet:", sheet_name, "\n")
    }
  }
  
  return(wb)  # Return the modified workbook
}
