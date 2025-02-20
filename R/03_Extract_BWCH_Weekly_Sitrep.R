# Initialize necessary variables and paths
file_path <- BWC_Sitrep_input_file_path
excel_files <- list.files(path = file_path, full.names = TRUE)
output_file_name <- paste0(output_file_path, "/Processed_BWC_Sitrep.xlsx")

# Data Processing --------------------------------------------------------------

# The function logic:
#1. Make the row containing "Metric" as header
#2. Select the necessary columns, excluding the column with NAs & the 1st column
#3. Rename the column "Description" to "Metric"
#4. Create a tibble
#5. Pivot the wider dates format to a long format
#6. Format the dates to appropriate "DD/MM/YYYY" format
#7. Create a new column "Day" for the corresponding day
#8. Create a new column "Organisation" to distinguish among trusts
#9. Ensure all columns are in the correct data types
#10.Re-arrange the columns for the final output

# Initialize a global list to store all warnings and errors
warning_log <- list()
error_log <- list()

# Function to process each Excel file and capture warnings/errors
process_data <- function(excel_file) {
  warnings <- NULL
  
  # Wrap the process in `tryCatch` to handle errors
  result <- tryCatch({
    
    # Use `withCallingHandlers` to capture warnings
    processed_data <- withCallingHandlers({
      data <- read_xlsx(excel_file)  # Read the Excel file
      
      # Your existing data processing steps
      data %>%
        janitor::row_to_names(row_number = which(data$`BWC Informatics` == "Metric")) %>%
        select(Description, c(4: ncol(data))) %>%
        rename(Metric = Description) %>%
        as_tibble() %>%
        pivot_longer(
          cols = -Metric,
          names_to = "Date",
          values_to = "Value"
        ) %>%
        mutate(
          Date = dmy(Date),
          Day = lubridate::wday(Date, label = TRUE, abbr = FALSE),
          Provider = "BWC",
          `Metric Category Type` = "Total",
          `Metric Category Name` = "Total",
          `File Name` = "BWC Weekly SITREP Report"
        ) %>%
        select(Metric, `Metric Category Type`, `Metric Category Name`,
               Provider, Date, Day, Value, `File Name`)
    }, warning = function(w) {
      # Capture the warning and the file name where it occurs
      warnings <<- c(warnings, w$message)
      message(paste("Warning in file:", basename(excel_file)))
      message(paste("Warning details:", w$message))
      invokeRestart("muffleWarning")  # Muffle the warning but continue processing
    })
    
    # If there were warnings, store them in the global warning log
    if (!is.null(warnings)) {
      warning_log[[basename(excel_file)]] <- warnings
    }
    
    return(processed_data)
    
  }, error = function(e) {
    # If there's an error, log the error and skip this file
    message(paste("Error in file:", basename(excel_file)))
    message("Error details:", e$message)
    error_log[[basename(excel_file)]] <- e$message
    return(NULL)  # Skip this file and return NULL
  })
  
  return(result)
}


# Initialize an empty list to store processed data frames
all_processed_data <- list()

# Loop through each Excel file and apply the process_data function
for (excel_file in excel_files) {
  processed_data <- process_data(excel_file)
  
  # Only add the processed data if it was successful (i.e., not NULL)
  if (!is.null(processed_data)) {
    all_processed_data[[basename(excel_file)]] <- processed_data
  }
}

# Combine all processed data frames into one
if (length(all_processed_data) > 0) {
  output <- bind_rows(all_processed_data)
  message("Data processing complete.")
} else {
  message("No data processed successfully.")
}

# Access warnings and errors after processing
if (length(warning_log) > 0) {
  message("Warnings summary:")
  for (file in names(warning_log)) {
    message(paste("File:", file))
    message(paste("Warnings:", paste(warning_log[[file]], collapse = "\n")))
  }
}

if (length(error_log) > 0) {
  message("Errors summary:")
  for (file in names(error_log)) {
    message(paste("File:", file))
    message(paste("Error:", error_log[[file]]))
  }
} else {
  message("No errors occurred during the process.")
}


# Update the Date
BWC_sitrep_output <- output %>% 
  filter(Date > max_date)

# Write the combined data to an Excel file
write_xlsx(BWC_sitrep_output, output_file_name)

# Move files from 'New data` to 'Old data' folder`
move_all_files(main_directory = paste0(input_file_path, "/BWC Sitrep"), source_folder = "New data", destination_folder = "Old data")
