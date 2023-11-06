base_csv_folder <- 'C:/Users/tyewf/Documents/base_csvs'
updated_csv_folder <- 'C:/Users/tyewf/Documents/updated_csvs'

base_csv_folder <- "C:/Users/tyewf/Documents/matching_csvs_base"
updated_csv_folder <- "C:/Users/tyewf/Documents/matching_csvs_updated"

base_csv_list, updated_csv_list
result_xlsx_path = 'C:/Users/tyewf/Documents/csv_comparer_example.xlsx'
result_xlsx_path = 'C:/Users/tyewf/Documents/csv_comparer_matching.xlsx'
compare_csvs <- function(
    base_csv_folder,
    updated_csv_folder,
    result_xlsx_path,
    verticle = TRUE,
    column_config = NULL,
    result_summary_sheet_name = 'Comparison Summary'
) {
    # update name of result_xlsx_path to include current date and time
    result_xlsx_path <- paste0(
        gsub(".xlsx", "", result_xlsx_path),
        "_",
        format(Sys.time(), "%Y%m%d_%H%M%S"),
        ".xlsx"
    )

    # Get the base and updated csv lists
    base_and_update_csv_lists <- get_csv_base_and_update_lists(
        base_csv_folder = base_csv_folder,
        updated_csv_folder = updated_csv_folder
    )

    # Get the base and updated csv lists
    base_csv_list <- base_and_update_csv_lists[[1]]
    updated_csv_list <- base_and_update_csv_lists[[2]]

    # Check that the base and updated csv lists are not empty
    if(is.null(base_csv_list) | is.null(updated_csv_list)) {
        message("The base and updated csv lists are empty")
        return(FALSE)
    }

    # Check that the base and updated csv lists are the same length
    if(!length(base_csv_list) == length(updated_csv_list)) {
        message("The base and updated csv lists are not the same length")
        return(FALSE)
    }

    # Add csvs to the excel
    add_csvs_to_excel(
        csv_list = c(base_csv_list, updated_csv_list),
        result_xlsx_path = result_xlsx_path
    )

    # Get the sheet names
    base_sheet_names <- gsub(".csv", "", basename(base_csv_list))
    updated_sheet_names <- gsub(".csv", "", basename(updated_csv_list))

    # Ensure columns are the same in each sheet
    if(!check_column_names(
    result_xlsx_path = result_xlsx_path, sheet_names = base_sheet_names)) {
        message("The column names are not the same in the base sheets")
        message("Please check the base CSVs and try again")
        return(FALSE)
    }
    if(!check_column_names(
    result_xlsx_path = result_xlsx_path, sheet_names = updated_sheet_names)) {
        message("The column names are not the same in the updated sheets")
        message("Please check the updated CSVs and try again")
        return(FALSE)
    }
    if(!check_column_names(
    result_xlsx_path = result_xlsx_path, sheet_names = c(
    updated_sheet_names, base_sheet_names))) {
        message("The column names are not the same in the base and updated ",
        "sheets")
        message("Please check the CSVs and try again")
        return(FALSE)
    }

    # Now there is an excel that can be tested
    run_standardised_xlsx_sheet_comparison(
        result_xlsx_path = result_xlsx_path,
        base_sheet_names = base_sheet_names,
        updated_sheet_names = updated_sheet_names
    )

}

# Check column names are the same in each sheet
check_column_names <- function(
    result_xlsx_path,
    sheet_names
) {
    wb <- loadWorkbook(result_xlsx_path)

    columns <- list()
    for(sheet in sheet_names) {
        columns[[sheet]] <- names(read_excel(result_xlsx_path, sheet = sheet))
    }

    for(i in columns){
        if(!all(i == columns[[1]])){
            message("The column names are not the same in the following ",
            "sheets: ")
            message(sheet_names)
            return(FALSE)
        }
    }

    return(TRUE)
}


add_csvs_to_excel <- function(
    csv_list,
    result_xlsx_path
) {
    # Loop through the csvs    
    for(csv_path in csv_list) {
        sheet_name <- gsub(".csv", "", basename(csv_path))
        add_csv_to_excel(csv_path, excel_path= result_xlsx_path, sheet_name)
    }
}

# Function to add CSV data to Excel sheet (overwriting or creating)
add_csv_to_excel <- function(csv_path, excel_path, sheet_name) {
  if (file.exists(excel_path)) {
    # If the Excel file exists, load it
    wb <- loadWorkbook(excel_path)
  } else {
    # If the Excel file doesn't exist, create a new workbook
    wb <- createWorkbook()
  }
  wb$sheet_names
  # Check if the specified sheet already exists
  if (sheet_name %in% wb$sheet_names) {
    # If the sheet exists, overwrite it with the CSV data
    csv_data <- read.csv(csv_path)
    writeData(wb, sheet = sheet_name, x = csv_data)
  } else {
    # If the sheet doesn't exist, add a new sheet with the CSV data
    csv_data <- read.csv(csv_path)
    addWorksheet(wb, sheet_name)
    writeData(wb, sheet = sheet_name, x = csv_data)
  }
  
  # Save the updated workbook to the Excel file
  saveWorkbook(wb, excel_path, overwrite = TRUE)
}




#' Get CSV Files from Base and Updated Folders
#'
#' This function retrieves and matches CSV files from two input folders based on
#' specific criteria and returns lists of matched CSV files from both folders.
#'
#' @param base_csv_folder A character string specifying the folder path containing the
#'                        base CSV files.
#' @param updated_csv_folder A character string specifying the folder path containing
#'                           the updated CSV files.
#'
#' @details
#'   1. The function retrieves the list of CSV files from both input folders that have
#'      filenames ending in an underscore followed by a number.
#'   2. It checks that the numbers extracted from the filenames in both lists match.
#'   3. Only the CSV files with matching numbers in both lists are retained.
#'   4. The function checks if the resulting CSV lists are not empty. If they are empty,
#'      a message is displayed, and the function returns FALSE.
#'   5. The CSV lists are sorted based on the extracted numbers.
#'   6. The function checks that the base CSV filenames have consistent names prior to
#'      the underscore. If not, a warning message is displayed.
#'   7. The function extracts and compares the base CSV filenames (without numbers or
#'      ".csv" extensions) and updated CSV filenames for consistency. If any inconsistency
#'      is found, a warning message is displayed.
#'   8. The function checks if the base and updated CSV lists have the same length. If not,
#'      a message is displayed, and the function returns FALSE.
#'   9. Finally, the function constructs the full paths for the matched CSV files in both
#'      folders and returns a list containing the base and updated CSV lists.
#'
#' @return A list with two elements:
#' \item{base_csv_list}{A character vector containing the full paths of matched base CSV files.}
#' \item{updated_csv_list}{A character vector containing the full paths of matched updated CSV files.}
#'
#' @note
#' The function provides warnings and messages to inform the user about any inconsistencies
#' or potential issues with the CSV files in the input folders.
#'
#' @examples
#' \dontrun{
#' # Example usage
#' result <- get_csv_base_and_update_lists("path_to_base_folder", "path_to_updated_folder")
#' }
#'
#' @export
get_csv_base_and_update_lists <- function(
    base_csv_folder,
    updated_csv_folder
) {
    # get the list of csvs
    base_csv_list <- list.files(
        base_csv_folder, pattern = "*.csv", full.names = FALSE)
    updated_csv_list <- list.files(
        updated_csv_folder, pattern = "*.csv", full.names = FALSE)


    # Only include csvs that end in _ then a number
    base_csv_list <- filter_csv_list_numeric_ending(
        csv_list = base_csv_list)
    updated_csv_list <- filter_csv_list_numeric_ending(
        csv_list = updated_csv_list)

    # check all numbers are the same
    base_csv_numbers <- gsub(".*_", "", base_csv_list)
    updated_csv_numbers <- gsub(".*_", "", updated_csv_list)

    # only keep the numbers that are in both lists
    base_csv_list <- base_csv_list[base_csv_numbers %in% updated_csv_numbers]
    updated_csv_list <- updated_csv_list[
        updated_csv_numbers %in% base_csv_numbers]

    # check list is not empty
    if (length(base_csv_list) == 0) {
        message("No csvs match between the base and updated folders")
        return(FALSE)
    }

    # sort the csvs
    base_csv_list <- base_csv_list[order(base_csv_numbers)]
    updated_csv_list <- updated_csv_list[order(updated_csv_numbers)]

    # Check that base CSVs have consistent names prior to the _
    gsub("*_", "", base_csv_list)

    # remove the numbers and .csv from the end of the csvs
    base_csv_list_names <- gsub("_[0-9]+.csv", "", base_csv_list)
    updated_csv_list_names <- gsub("_[0-9]+.csv", "", updated_csv_list)

    # check that the names are the same in each folder
    if(!all(base_csv_list_names == base_csv_list_names[1])) {
        message("Warning, csv names are not consistent in the base CSVs")
        message("Continuing anyway")
    }
    if(!all(updated_csv_list_names == updated_csv_list_names[1])) {
        message("Warning, csv names are not consistent in the updated CSVs")
        message("Continuing anyway")
    }

    # Check CSV lists are the same length
    if(!length(base_csv_list) == length(updated_csv_list)) {
        message("The base and updated CSV lists are not the same length")
        return(FALSE)
    }
    
    base_csv_list <- paste0(base_csv_folder, "/", base_csv_list)
    updated_csv_list <- paste0(updated_csv_folder, "/", updated_csv_list)

    # return the list of csvs
    return(list(base_csv_list, updated_csv_list))
}

filter_csv_list_numeric_ending <- function(
    csv_list
) {
    # remove '.csv' from the end of the names
    csv_names <- gsub(".csv", "", csv_list)

    # Remove all csvs that do not end in _ then a number
    csv_names <- csv_names[grepl("_[0-9]+$", csv_names)]
    if(!length(csv_names[!grepl("_[0-9]+$", csv_names)]) == 0){
        message("The following csvs were removed as they did not end in ",
        "_ then a number:")
        message(csv_names[!grepl("_[0-9]+$", csv_names)])
    }

    # add the .csv back on
    csv_names <- paste0(csv_names, ".csv")

    # return the list of csvs
    return(csv_names)
}




add_csv_to_excel <-function(
    path_results_xlsx,
    sheet_name,
    path_csv,
    column_config = NULL,
    verticle = TRUE,
    result_summary_sheet_name = 'Comparison Summary'
){
    # read the csv
    csv <- read_csv(path_csv)
    # read the excel
    wb <- openxlsx::loadWorkbook(path_results_xlsx)
    # add the csv to the excel
    wb <- add_sheet_to_excel(
        wb = wb,
        sheet_name = sheet_name,
        data = csv,
        column_config = column_config,
        verticle = verticle
    )
    # save the excel
    openxlsx::saveWorkbook(wb, path_results_xlsx, overwrite = TRUE)
    # return the excel
    return(wb)
}