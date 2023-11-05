#' Compare multiple sheets in an Excel workbook and generate a summary report.
#'
#' This function performs comparisons between multiple pairs of sheets in an Excel workbook and generates a summary report with the results of the comparisons. It validates the input Excel file, compares sheets, and creates a summary sheet with detailed information.
#'
#' @param path_results_xlsx The file path to the Excel workbook to be processed.
#' @param sheets_to_compare A data frame containing the sheet names to be compared, with columns 'base_sheet_name', 'updated_sheet_name', and 'result_sheet_name'.
#' @param verticle If TRUE, comparisons are performed in a vertical manner; if FALSE, in a horizontal manner.
#' @param column_config A configuration specifying the columns to compare and their properties.
#' @param result_summary_sheet_name The name of the summary sheet that will be created in the Excel workbook.
#'
#' @return The modified Excel workbook with the summary report and comparison results.
#'
#' @import openxlsx
# @importFrom base print paste0
# @importFrom utils recalc
# @importFrom openxlsx loadWorkbook saveWorkbook writeData addWorksheet
# @importFrom .file_exists
# @importFrom .xlsx_validation
# @importFrom .remove_sheets
# @importFrom .generate_summary_sheet
# @importFrom .compare_sheets
# @examples
# # Example usage:
# path_to_excel <- "path/to/your/excel.xlsx"
# sheets_to_compare_data <- data.frame(
#   base_sheet_name = c("Sheet1", "Sheet2"),
#   updated_sheet_name = c("Sheet1_updated", "Sheet2_updated"),
#   result_sheet_name = c("Comparison1", "Comparison2")
# )
# column_config <- list(
#   list("Column1", "Type1"),
#   list("Column2", "Type2")
# )
# result_summary_name <- "SummarySheet"
# compare_xl_sheets(
#   path_results_xlsx = path_to_excel,
#   sheets_to_compare = sheets_to_compare_data,
#   verticle = TRUE,
#   column_config = column_config,
#   result_summary_sheet_name = result_summary_name
# )
compare_xl_sheets <- function(
    path_results_xlsx,
    sheets_to_compare,
    verticle = TRUE,
    column_config,
    result_summary_sheet_name = 'Comparison Summary'
) {
    print("Performing Excel Sheet comparisons")

    # Validate inputs
    if(!compare_xl_sheets_validator(
        path_results_xlsx,
        sheets_to_compare,
        column_config)){
        print("Excel Sheet comparison validation failed")
        return()
    }

    # Read the Excel file
    wb <- openxlsx::loadWorkbook(path_results_xlsx)

    # initiate summary sheet

    # Remove the result sheets if any exist
    wb <- remove_sheets(wb, result_summary_sheet_name)

    # Add the summary sheet
    openxlsx::addWorksheet(wb, sheetName = result_summary_sheet_name)

    # Add in the summary sheet header
    cell_formula <- 'Comparison Summary'
    openxlsx::writeData(
        wb, sheet = result_summary_sheet_name, x = cell_formula, startCol = 1, 
        startRow = 1, colNames = FALSE, rowNames = FALSE)
    
    # Add in headers for each result sheet
    headers <- c(
        'Result Sheet', 'Columns Names Match', 'TRUE count', 'FALSE count',
        'N/A count', 'Numeric Differences count')
    for(header in 1:length(headers)){
        cell_formula <- headers[header]
        openxlsx::writeData(
            wb, sheet = result_summary_sheet_name, x = cell_formula, 
            startCol = header, startRow = 2, colNames = FALSE, 
            rowNames = FALSE)
    }

    # compare the sheets
    for(comparison in 1:nrow(sheets_to_compare)){

        summary_row <- comparison + 2

        print(paste0("Comparison: ", comparison))
        result <- compare_sheets(
            wb = wb,
            base_sheet_name = sheets_to_compare$base_sheet_name[comparison],
            updated_sheet_name = sheets_to_compare$updated_sheet_name[
                comparison],
            result_sheet_name = sheets_to_compare$result_sheet_name[comparison],
            column_config = column_config,
            verticle = verticle
        )
        
        # Add the summary sheet
        wb <- generate_summary_sheet(
            wb = wb <- result[[1]],
            result_sheet_name = sheets_to_compare$result_sheet_name[comparison],
            col_range_true_false_headers <- result[[2]],
            row_range_true_false_headers <- result[[3]],
            col_range_true_false <- result[[4]],
            row_range_true_false <- result[[5]],
            col_range_numeric <- result[[6]],
            row_range_numeric <- result[[7]],
            result_summary_row = summary_row,
            result_summary_sheet_name = result_summary_sheet_name
        )

    }

    # Save the workbook
    openxlsx::saveWorkbook(wb, path_results_xlsx, overwrite = TRUE)

    # Reorder the sheets
    order_sheets(
        path_results_xlsx = path_results_xlsx,
        sheets_at_start = c(
            result_summary_sheet_name,
            sheets_to_compare$result_sheet_name,
            sheets_to_compare$base_sheet_name,
            sheets_to_compare$updated_sheet_name)
    )

}


generate_summary_sheet <- function(
    wb,
    result_sheet_name,
    col_range_true_false_headers,
    row_range_true_false_headers,
    col_range_true_false,
    row_range_true_false,
    col_range_numeric,
    row_range_numeric,
    result_summary_row,
    result_summary_sheet_name = 'Comparison Summary'
) {
    
    
    # Add in the result sheet name
    cell_formula = paste0(result_sheet_name)
    openxlsx::writeData(
            wb, sheet = result_summary_sheet_name, x = cell_formula, 
            startCol = 1, startRow = result_summary_row, colNames = FALSE, 
            rowNames = FALSE)

    # Add in the column names match
    cell_formula <- paste0(
        '=AND(\'', result_sheet_name, '\'!',
        openxlsx::int2col(col_range_true_false_headers[1]),
        row_range_true_false_headers, ':',
        openxlsx::int2col(tail(col_range_true_false_headers, n = 1)),
        row_range_true_false_headers, ')=TRUE')
    openxlsx::writeFormula(
        wb, sheet = result_summary_sheet_name, x = cell_formula, 
        startCol = 2, startRow = result_summary_row)

    # Add in the true count
    cell_formula <- paste0(
        '=COUNTIF(\'', result_sheet_name, '\'!',
        openxlsx::int2col(col_range_true_false[1]),
        row_range_true_false[1], ':',
        openxlsx::int2col(tail(col_range_true_false, n = 1)),
        tail(row_range_true_false, n = 1), ',TRUE)')
    openxlsx::writeFormula(
        wb, sheet = result_summary_sheet_name, x = cell_formula, 
        startCol = 3, startRow = result_summary_row)

    # Add in the false count
    cell_formula <- paste0(
        '=COUNTIF(\'', result_sheet_name, '\'!',
        openxlsx::int2col(col_range_true_false[1]),
        row_range_true_false[1], ':',
        openxlsx::int2col(tail(col_range_true_false, n = 1)),
        tail(row_range_true_false, n = 1), ',FALSE)')
    openxlsx::writeFormula(
        wb, sheet = result_summary_sheet_name, x = cell_formula, 
        startCol = 4, startRow = result_summary_row)

    # Add in N/A count
    cell_formula <- paste0(
        '=COUNTIF(\'', result_sheet_name, '\'!',
        openxlsx::int2col(col_range_true_false[1]),
        row_range_true_false[1], ':',
        openxlsx::int2col(tail(col_range_true_false, n = 1)),
        tail(row_range_true_false, n = 1), ',NA())')
    openxlsx::writeFormula(
        wb, sheet = result_summary_sheet_name, x = cell_formula, 
        startCol = 5, startRow = result_summary_row)

    # Add in numeric differences count
    cell_formula <- paste0(
        '=COUNTIF(\'', result_sheet_name, '\'!',
        openxlsx::int2col(col_range_numeric[1]),
        row_range_numeric[1], ':',
        openxlsx::int2col(tail(col_range_numeric, n = 1)),
        tail(row_range_numeric, n = 1), ',">0")')
    openxlsx::writeFormula(
        wb, sheet = result_summary_sheet_name, x = cell_formula, 
        startCol = 6, startRow = result_summary_row)

    return(wb)
}

compare_sheets <- function(
    wb,
    base_sheet_name,
    updated_sheet_name,
    result_sheet_name,
    column_config = NULL,
    verticle = TRUE
) {
    
    # Read the base and updated sheets
    base_sheet <- as.data.table(
        openxlsx::read.xlsx(wb, sheet = base_sheet_name))
    updated_sheet <- as.data.table(
        openxlsx::read.xlsx(wb, sheet = updated_sheet_name))

    if(is.null(column_config)) {
        message("Column config not provided, creating column config")
        message("Assumes:")
        message("    - all columns in base are in the updated sheets")
        message("    - first column is the unique key")
        column_config <- data.table(
            base_cols_to_compare = names(base_sheet),
            updated_cols_to_compare = names(base_sheet),
            data_type = rep('numeric', length(names(base_sheet))),
            rounding = rep(NA, length(names(base_sheet))),
            key_column = c(TRUE, rep(FALSE, length(names(base_sheet))-1))
        )
    }

    # Validate column config
    if(!sum(column_config$key_column) == 1) {
        message("There must be exactly one key column")
        return(base::list(wb))
    }

    # validate base data table
    message("Validating base data table")
    if(!validate_data_table(
        dt = base_sheet,
        columns = column_config$base_cols_to_compare,
        key = column_config[key_column == TRUE, base_cols_to_compare]
    )){
        message("Base data table validation failed")
        return(base::list(wb))
    }
    
    # validate updated data table
    message("Validating updated data table")
    if(!validate_data_table(
        dt = updated_sheet,
        columns = column_config$updated_cols_to_compare,
        key = column_config[key_column == TRUE, updated_cols_to_compare]
    )){
        message("Updated data table validation failed")
        return(base::list(wb))
    }

    # make key column the first column in the column_config
    column_config <- column_config[
        order(column_config$key_column, decreasing = TRUE)]

    # Remove the result sheet if it exists
    wb <- remove_sheets(wb, result_sheet_name)
    # Add the result sheet
    openxlsx::addWorksheet(wb, result_sheet_name)

    # Add Base Data to the results sheet #######################################

    # Add Base Data header
    cell_formula <- 'Base Data'
    openxlsx::writeData(
        wb, sheet = result_sheet_name, x = cell_formula, startCol = 1, 
        startRow = 1, colNames = FALSE, rowNames = FALSE)

    # Select columns to compare
    if(is.null(column_config)){
        base_cols_to_compare <- names(base_sheet)
    } else{
        base_cols_to_compare <- column_config$base_cols_to_compare
    }
    base_cols_ref <- base::match(base_cols_to_compare, names(base_sheet))

    # Add Base Data into results sheet
    for(col in 1:length(base_cols_ref)){
        for(row in 1:(nrow(base_sheet)+1)){
            cell_formula <- paste0(
                "='", base_sheet_name, "'!",
                openxlsx::int2col(base_cols_ref[col]), row)
            openxlsx::writeFormula(
                wb, sheet = result_sheet_name, x = cell_formula, startCol = col, 
                startRow = row + 1)
        }
    }

    # Add Updated Data to the results sheet ####################################

    # Get starting row and column based on verticle or horizontal layout
    if(verticle) {
        row_start_updated <- nrow(base_sheet) + 3
        col_start_updated <- 0
    } else {
        row_start_updated <- 0
        col_start_updated <- length(base_cols_ref) + 1
    }

    # Add Updated Data header
    cell_formula <- 'Updated Data'
    updated_row <- 1
    updated_col <- 1
    openxlsx::writeData(
        wb, sheet = result_sheet_name, x = cell_formula,
        startCol = col_start_updated + updated_col, 
        startRow = row_start_updated + updated_row, colNames = FALSE, 
        rowNames = FALSE)

    # Select columns to compare
    if(is.null(column_config)){
        updated_cols_to_compare <- names(updated_sheet)
    } else{
        updated_cols_to_compare <- column_config$updated_cols_to_compare
    }
    updated_cols_ref <- base::match(
        updated_cols_to_compare, names(updated_sheet))

    # Add key column using base data keys
    col <- 1
    for(row in 1:(nrow(base_sheet)+1)){
        cell_formula <- paste0(
            "=A", row+1)
        openxlsx::writeFormula(
            wb, sheet = result_sheet_name, x = cell_formula,
            startCol = col_start_updated + col,
            startRow = row_start_updated + row + 1)
    }

    # Add column names from updated data
    for(col in 1:length(updated_cols_ref)){
        row <- 1
        cell_formula <- paste0(
            "='", updated_sheet_name, "'!",
            openxlsx::int2col(updated_cols_ref[col]), row)
        openxlsx::writeFormula(
            wb, sheet = result_sheet_name, x = cell_formula, 
            startCol = col_start_updated + col, 
            startRow = row_start_updated + row + 1)
    }

    # Add Updated Data into results sheet using index match
    for(col in 2:length(updated_cols_ref)){
        for(row in 2:(nrow(base_sheet)+1)){
            cell_formula <- paste0(
                "=INDEX('", updated_sheet_name, "'!",
                openxlsx::int2col(updated_cols_ref[col]), ":",
                openxlsx::int2col(updated_cols_ref[col]),
                ",MATCH(A", row_start_updated+row+1, ",'", updated_sheet_name,
                "'!", openxlsx::int2col(updated_cols_ref[1]), ":",
                openxlsx::int2col(updated_cols_ref[1]), ",0))")
            openxlsx::writeFormula(
                wb, sheet = result_sheet_name, x = cell_formula, 
                startCol = col_start_updated + col, 
                startRow = row_start_updated + row + 1)
        }
    }

    # Create true false comparison #############################################
    if(verticle) {
        row_start_comparison_true_false <- (
            row_start_updated + nrow(base_sheet) + 3)
        col_start_comparison_true_false <- 0
    } else {
        row_start_comparison_true_false <- 0
        col_start_comparison_true_false <- (
            col_start_updated + length(base_cols_ref) + 1)
    }

    # Add comparison header
    cell_formula <- 'Comparison - TRUE/FALSE'
    openxlsx::writeData(
        wb, sheet = result_sheet_name, x = cell_formula, 
        startCol = col_start_comparison_true_false + 1, 
        startRow = row_start_comparison_true_false + 1, colNames = FALSE, 
        rowNames = FALSE)

    # Loop through columns to compare
    for(col in 1:length(base_cols_ref)){
        # Loop through rows to compare
        for(row in 1:(nrow(base_sheet)+1)){            
            # Create the comparison formula
            cell_formula <- paste0(
                "=",
                openxlsx::int2col(col),
                (row + 1),
                "=",
                openxlsx::int2col(col_start_updated + col),
                (row_start_updated + row + 1))
            # Write the comparison formula
            openxlsx::writeFormula(
                wb, sheet = result_sheet_name, x = cell_formula, 
                startCol = col_start_comparison_true_false + col, 
                startRow = row_start_comparison_true_false + row + 1)
        }
    }

    # Format true false comparison
    negStyle <- openxlsx::createStyle(
        fontColour = "#9C0006", bgFill = "#FFC7CE")
    posStyle <- openxlsx::createStyle(
        fontColour = "#006100", bgFill = "#C6EFCE")
    col_range_true_false <- (col_start_comparison_true_false + 1):
        (col_start_comparison_true_false + length(base_cols_ref))
    row_range_true_false <- (row_start_comparison_true_false + 2):
        (row_start_comparison_true_false + nrow(base_sheet) + 2)
    openxlsx::conditionalFormatting(
        wb, sheet = result_sheet_name, cols = col_range_true_false, 
        rows = row_range_true_false, rule = "!=TRUE", style = negStyle)

    # Create numeric comparison ################################################
    if(verticle) {
        row_start_comparison_numeric <- (
            row_start_comparison_true_false + nrow(base_sheet) + 3)
        col_start_comparison_numeric <- 0
    } else {
        row_start_comparison_numeric <- 0
        col_start_comparison_numeric <- (
            col_start_comparison_true_false + length(base_cols_ref) + 1)
    }

    # Add comparison header
    cell_formula <- 'Comparison - Numeric'
    openxlsx::writeData(
        wb, sheet = result_sheet_name, x = cell_formula, 
        startCol = col_start_comparison_numeric + 1, 
        startRow = row_start_comparison_numeric + 1, colNames = FALSE, 
        rowNames = FALSE)

    # Header from base data
    for(col in 1:length(base_cols_ref)){
        row <- 2
        cell_formula <- paste0(
            "=", openxlsx::int2col(col), row)
        openxlsx::writeFormula(
            wb, sheet = result_sheet_name, x = cell_formula, 
            startCol = col_start_comparison_numeric + col, 
            startRow = row_start_comparison_numeric + row)
    }

    # Key column from base data
    for(row in 1:(nrow(base_sheet)+1)){
        cell_formula <- paste0("=A", (row + 1))
        openxlsx::writeFormula(
            wb, sheet = result_sheet_name, x = cell_formula, 
            startCol = col_start_comparison_numeric + 1, 
            startRow = row_start_comparison_numeric + row + 1)
    }
    
    # Loop through columns to compare
    for(col in 2:length(base_cols_ref)){
        # Loop through rows to compare
        for(row in 2:(nrow(base_sheet)+1)){

            base_cell <- paste0(
                openxlsx::int2col(base_cols_ref[col]), (row + 1))
            updated_cell <- paste0(
                openxlsx::int2col(col_start_updated + base_cols_ref[col]),
                (row_start_updated + row + 1))

            # Determine if the column is numeric and if rounding is required
            if(column_config$data_type[col] == 'numeric' &
            !is.na(column_config$rounding[col])){
                cell_formula <- paste0(
                    "=ROUND(",
                    base_cell,
                    ",",
                    column_config$rounding[col],
                    ")-",
                    "ROUND(",
                    updated_cell,
                    ",",
                    column_config$rounding[col],
                    ")")
            } else if (column_config$data_type[col] == 'numeric' &
            is.na(column_config$rounding[col])) {
                cell_formula <- paste0(
                    "=",
                    base_cell,
                    "-",
                    updated_cell)
            } else {
                cell_formula <- ""
            }

            # Write the comparison formula
            openxlsx::writeFormula(
                wb, sheet = result_sheet_name, x = cell_formula,
                startCol = col_start_comparison_numeric + col,
                startRow = row_start_comparison_numeric + row + 1)
        }
    }

    # Format numeric comparison
    negStyle <- openxlsx::createStyle(
        fontColour = "#9C0006", bgFill = "#FFC7CE")
    posStyle <- openxlsx::createStyle(
        fontColour = "#006100", bgFill = "#C6EFCE")
    col_range_numeric <- (col_start_comparison_numeric + 2):
        (col_start_comparison_numeric + length(base_cols_ref))
    numeric_columns <- (column_config$data_type == 'numeric')[-1]
    col_range_numeric <- col_range_numeric[numeric_columns]
    row_range_numeric <- (row_start_comparison_numeric + 3):
        (row_start_comparison_numeric + nrow(base_sheet) + 2)

    # Add conditional formatting for numeric comparison
    openxlsx::conditionalFormatting(
        wb, sheet = result_sheet_name, cols = col_range_numeric,
        rows = row_range_numeric, rule = "<>0", style = negStyle)
    openxlsx::conditionalFormatting(
        wb, sheet = result_sheet_name, cols = col_range_numeric,
        rows = row_range_numeric, rule = "=0", style = posStyle)

    message("Comparison complete")    
    return(base::list(
        wb, 
        col_range_true_false, # col_range_true_false_headers
        row_range_true_false[1], # row_range_true_false_headers
        col_range_true_false, # col_range_true_false
        row_range_true_false[-1], # row_range_true_false
        col_range_numeric,
        row_range_numeric
    ))
}





#' Check if a file exists at the specified path.
#'
#' This function checks if a file or directory exists at the given file path. It returns a logical value indicating
#' whether the file or directory exists and prints a corresponding message.
#'
#' @param file_path A character string representing the file path to check.
# @return TRUE if the file or directory exists; FALSE otherwise.
# @export
# @examples
# # Example 1: Check if a file exists
# file_path1 <- "path/to/your/file.txt"
# result1 <- file_exists(file_path1)
# 
# # Example 2: Check if a directory exists
# directory_path <- "path/to/your/directory"
# result2 <- file_exists(directory_path)
# 
# if (result1) {
#   cat("Perform some actions with the file.\n")
# } else {
#   cat("Handle the case where the file does not exist.\n")
# }
# 
# if (result2) {
#   cat("Perform some actions with the directory.\n")
# } else {
#   cat("Handle the case where the directory does not exist.\n")
file_exists <- function(file_path) {
  if (file.exists(file_path)) {
    cat("File exists: ", file_path, "\n")
    return(TRUE)
  } else {
    cat("File does not exist: ", file_path, "\n")
    return(FALSE)
  }
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
    wb,
    expected_sheet_names
) {
    xlsx_sheets <- openxlsx::sheets(wb)

    # Check if the expected sheet names are in the xlsx_sheets
    result <- TRUE
    for (sheet in expected_sheet_names) {
        if (!(sheet %in% xlsx_sheets)) {
            result <- FALSE
            message(paste0("Sheet: ", sheet, " not found."))
        }
    }
    if(result) {
        message(paste0("All expected sheets found."))
    } else {
        message(paste0("Some expected sheets missing."))
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


compare_xl_sheets_validator <- function(
    path_results_xlsx,
    sheets_to_compare,
    column_config
) {
    result <- TRUE

    # path_results_xlsx ########################################################

    # Check if the excel file exists
    if (!file.exists(path_results_xlsx)) {
        message(paste0("The excel file does not exist: ", path_results_xlsx))
        result <- FALSE
    }

    # sheets_to_compare ########################################################

    # Check if the sheets to compare is a data.table
    if (!is.data.table(sheets_to_compare)) {
        message("The sheets to compare is not a data.table")
        result <- FALSE
    }

    # Check if the sheets to compare has the correct columns
    if (!all(c('base_sheet_name', 'updated_sheet_name', 'result_sheet_name') 
    %in% names(sheets_to_compare))) {
        message("The sheets to compare does not have the correct columns")
        message("The sheets to compare should have the following columns:")
        message("base_sheet_name, updated_sheet_name, result_sheet_name")
        result <- FALSE
    }

    # Check if the sheets to compare has the correct number of rows
    if (nrow(sheets_to_compare) == 0) {
        message("The sheets to compare has no rows")
        result <- FALSE
    }

    # Check all the sheets to compare exist in the excel file
    wb <- openxlsx::loadWorkbook(path_results_xlsx)
    if(!xlsx_validation(
    wb = wb, expected_sheet_names = sheets_to_compare$base_sheet_name)){
            result <- FALSE}
    if(!xlsx_validation(
    wb = wb, expected_sheet_names = sheets_to_compare$updated_sheet_name)){
            result <- FALSE}
    
    # check if all sheets are unique
    if (length(unique(c(sheets_to_compare$base_sheet_name,
    sheets_to_compare$updated_sheet_name,
    sheets_to_compare$result_sheet_name))) !=   
    (nrow(sheets_to_compare)*3)) {
        message("The base sheet names are not unique")
        result <- FALSE
    }

    # column config ############################################################

    # Check if the column config is a data.table
    if (!is.data.table(column_config)) {
        message("The column config is not a data.table")
        result <- FALSE
    }

    # Check if the column config has the correct columns
    if (!all(c('base_cols_to_compare', 'updated_cols_to_compare', 'data_type',
    'rounding', 'key_column') %in% names(column_config))) {
        message("The column config does not have the correct columns")
        message("The column config should have the following columns:")
        message("base_cols_to_compare, updated_cols_to_compare, data_type,",
        " rounding, key_column")
        result <- FALSE
    }

    # Check if the column config has the correct number of rows
    if (nrow(column_config) == 0) {
        message("The column config has no rows")
        result <- FALSE
    }

    # Check all columns in config exist in the sheets to compare
    # Base sheets
    for(sheet in sheets_to_compare$base_sheet_name) {
        base_sheet <- names(as.data.table(
            openxlsx::read.xlsx(wb, sheet = base_sheet_name)))
        if (!all(column_config$base_cols_to_compare %in% base_sheet)) {
            message("The base columns to compare do not exist in the base sheet")
            result <- FALSE
        }
    }
    # Updated sheets
    for(sheet in sheets_to_compare$updated_sheet_name) {
        updated_sheet <- names(as.data.table(
            openxlsx::read.xlsx(wb, sheet = updated_sheet_name)))
        if (!all(column_config$updated_cols_to_compare %in% updated_sheet)) {
            message("The updated columns to compare do not exist in the updated sheet")
            result <- FALSE
        }
    }

    # Check if the column config has the correct data types
    if (!all(c('character', 'numeric') %in% column_config$data_type)) {
        message("The column config does not have the correct data types")
        message("The column config should have the following data types:")
        message("character, numeric")
        result <- FALSE
    }

    # Check if the column config has the correct rounding
    for(val in column_config$rounding) {
        if (!(is.numeric(val) | val != 'None')) {
            message("The column config does not have the correct rounding")
            message("The column config should only have numeric or 'None'",
            " values")            
            result <- FALSE
        }
    }

    # Check if the column config has the correct key column
    for(val in column_config$key_column) {
        if (!is.logical(val)) {
            message("The column config does not have the correct key column")
            message("The column config should only have logical values")            
            result <- FALSE
        }
    }

    # check there is only one key column
    if (sum(column_config$key_column) != 1) {
        message("There is not exactly one key column")
        result <- FALSE
    }

    return(result)
}

validate_data_table <- function(
    dt,
    columns,
    key
){
    result <- TRUE

    # Check required columns exist
    for(cols in columns){
        if(!cols %in% names(dt)){
            print(paste0("Column missing: ", cols))
            result <- FALSE
        }
    }

    # Check key is unique
    if(!length(unique(dt[[key]])) == nrow(dt)){
        print("Key column is not unique")
        result <- FALSE
    }

    return(result)
}



#' Reorder sheets within an Excel workbook based on a specified order.
#'
#' This function is designed to reorder the sheets in an Excel workbook based on a specified order.
#' You can specify a vector of sheet names that should appear at the beginning of the workbook,
#' and the function will rearrange the sheets accordingly.
#'
#' @param path_results_xlsx A character string specifying the path to the Excel workbook (XLSX) to be reordered.
#' @param sheets_at_start A character vector containing the sheet names to position at the beginning.
#'
#' @details
#' 1. Loads the Excel workbook specified by \code{path_results_xlsx} using the \code{loadWorkbook} function.
#' 2. Retrieves the names of all sheets in the workbook.
#' 3. Identifies the sheets that are not part of the specified order (\code{sheets_at_start}).
#' 4. Creates the desired order for all sheets by concatenating \code{sheets_at_start} with \code{other_sheets}.
#' 5. Generates a numeric ordered list using the \code{match} function, which reflects the order of sheets based on \code{req_order}.
#' 6. Reorders the sheets within the workbook using the \code{worksheetOrder} function.
#' 7. Saves the modified workbook to the same file (\code{path_results_xlsx}) while overwriting the existing file, if necessary.
#'
#' @dependencies
#' The function relies on the 'openxlsx' package for working with Excel files.
#'
#' @examples
#' # Reorder sheets in an Excel workbook
#' order_sheets("path/to/your/workbook.xlsx", c("SummarySheet", "Sheet1", "Sheet2", "Sheet3"))
#'
#' @note
#' Make sure to install and load the 'openxlsx' package before using this function.
#'
#' @export
order_sheets <- function(
    path_results_xlsx,
    sheets_at_start
) {

    wb <- loadWorkbook(path_results_xlsx)
    sheet = names(wb)

    # Determine the sheets that are not in the specified order
    other_sheets <- setdiff(sheets, sheets_at_start)

    # Create the desired order for all sheets
    req_order <- c(sheets_at_start, other_sheets)

    # Create a numeric ordered list
    numeric_ordered_list <- match(req_order, sheet)
    
    worksheetOrder(wb) <- numeric_ordered_list
    
    saveWorkbook(wb, path_results_xlsx, overwrite = TRUE)
}


