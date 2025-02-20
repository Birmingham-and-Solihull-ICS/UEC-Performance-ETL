# Initialize necessary variables and paths
file_path <- UHB_Sitrep_input_file_path
excel_files <- list.files(path = file_path, full.names = TRUE)
output_file_name <- paste0(output_file_path, "/Processed_UHB_Sitrep.xlsx")


# Data Processing --------------------------------------------------------------

# The function logics:
#1. Remove first 6 rows
#2. Make the first row as headers/columns
#3. Rename the first column to "Metric"
#4. Create a tibble
#5. Pivot the wider dates format to a long format
#6. Format the Excel dates to appropriate "DD/MM/YYYY" format
#7. Create a new column "Day" for the corresponding day
#8. Create a new column "Organisation" to distinguish among trusts
#9. Ensure all columns are in the correct data types
#10.Re-arrange the columns for the final output


# Function to process each sheet in a given Excel file, with error and warning handling
process_sheet_data <- function(excel_file, sheet_name) {
  warnings <- NULL
  
  # Wrap the process in `tryCatch` for error handling
  result <- tryCatch({
    
    # Use `withCallingHandlers` to capture warnings
    withCallingHandlers({
      read_xlsx(excel_file, sheet = sheet_name) %>%
        slice(-1:-6) %>%  # Remove the first 6 rows
        janitor::row_to_names(row_number = 1) %>%  # Set the first row as column names
        rename(Metric = 1) %>%  # Rename the first column to "Metric"
        as_tibble() %>%  # Convert to tibble
        pivot_longer(cols = -Metric, names_to = "Date", values_to = "Value") %>%  # Reshape to long format
        mutate(Date = as.Date(as.numeric(Date), origin = "1899-12-30")) %>%  # Convert Excel date to Date
        filter(Date > as.Date(max_date)) %>%  # Filter based on max_date
        mutate(Day = lubridate::wday(Date, label = TRUE, abbr = FALSE)) %>%  # Add day of the week
        mutate(Provider = sheet_name) %>%  # Add the provider name (sheet name)
        mutate(Value = str_replace_all(Value, "%", "")) %>%  # Remove percentage signs from values
        mutate(Value = readr::parse_number(Value)) %>%  # Parse numeric values from the Value column
        mutate(`File Name` = "Daily SITREP Checksheet") %>%  # Add file name
        mutate(`Metric Category Type` = "Total",
               `Metric Category Name` = "Total") %>%  # Add metric category columns
        select(Metric, `Metric Category Type`, `Metric Category Name`,
               Provider, Date, Day, Value, `File Name`)  # Select relevant columns for final output
    }, warning = function(w) {
      # Capture and log the warning along with file and sheet name
      warnings <<- c(warnings, w$message)
      message(paste("Warning in file:", basename(excel_file), "Sheet:", sheet_name))
      message("Warning details:", w$message)
      invokeRestart("muffleWarning")  # Muffle the warning but continue processing
    })
    
  }, error = function(e) {
    # Log the error along with the file and sheet name
    message(paste("Error processing file:", basename(excel_file), "Sheet:", sheet_name))
    message("Error details:", e$message)
    return(NULL)  # Return NULL to indicate failure
  })
  
  # If there were warnings, print them
  if (!is.null(warnings)) {
    message(paste("Warnings occurred in file:", basename(excel_file), "Sheet:", sheet_name))
  }
  
  return(result)
}

# Initialize an empty list to store data frames from all files
combined_data_list <- list()

# Initialize a list to log errors
error_log <- list()

# Loop through each Excel file and process the sheets
for (excel_file in excel_files) {
  
  # Get the names of all sheets in the current Excel file
  sheet_names <- excel_sheets(excel_file)
  
  # Initialize an empty list to store data frames for the current file
  data_list <- list()
  
  # Loop through each sheet name and process the data
  for (sheet_name in sheet_names) {
    # Wrap the processing in tryCatch to capture any issues
    processed_data <- tryCatch({
      process_sheet_data(excel_file, sheet_name)
    }, error = function(e) {
      # Log the error with file and sheet name
      error_message <- paste("Error processing file:", basename(excel_file), "Sheet:", sheet_name, ":", e$message)
      message(error_message)
      error_log[[paste(basename(excel_file), sheet_name)]] <- e$message
      return(NULL)  # Continue processing other sheets
    })
    
    if (!is.null(processed_data)) {
      data_list[[sheet_name]] <- processed_data
    }
  }
  
  # Combine all sheets from the current Excel file into one data frame
  if (length(data_list) > 0) {
    file_combined_data <- bind_rows(data_list, .id = "Provider")
    # Add the combined data frame for this file to the overall combined data list
    combined_data_list[[basename(excel_file)]] <- file_combined_data
  }
}

# Combine all data frames from all files into one
output <- bind_rows(combined_data_list)

# Filter Date > max_date
UHB_sitrep_output <- output %>% 
  filter(Date > max_date)


# # Add the combined data frame to the list of data under the key "All"
# combined_data_list$All <- output

# Print the error log (if any)
if (length(error_log) > 0) {
  message("The following errors occurred:")
  print(error_log)
} else {
  message("No errors occurred.")
}

write_xlsx(list("All" = UHB_sitrep_output), path = output_file_name)

# Move files from 'New data` to 'Old data' folder`
move_all_files(main_directory = paste0(input_file_path, "/UHB Sitrep"), source_folder = "New data", destination_folder = "Old data")


# # Excel row limit
# excel_row_limit <- 1000000
# 
# # Split the data into chunks that fit within the Excel row limit
# split_data_list <- split(output, ceiling(seq_len(nrow(output)) / excel_row_limit))
# 
# # Output Excel file path template
# output_file_template <- paste0(output_file_path, "/Processed_UHB_Sitrep_part_")
# 
# # Loop through the split data list and save each chunk into a new Excel file
# for (i in seq_along(split_data_list)) {
#   # Generate the file name (e.g., Processed_UHB_Sitrep_part_1.xlsx, part_2.xlsx, etc.)
#   output_file_name <- paste0(output_file_template, i, ".xlsx")
#   
#   # Create a list with the chunk of data and set the sheet name to "All"
#   data_to_write <- list(All = split_data_list[[i]])
#   
#   # Write the chunk of data to a new Excel file with the sheet named "All"
#   write_xlsx(data_to_write, path = output_file_name)
#   
#   message(paste("Data chunk", i, "written to", output_file_name))
# }


# ------------------------------------------------------------------------------




