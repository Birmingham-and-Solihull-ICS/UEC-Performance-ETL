# Initialize necessary variables and paths
file_path <- ADA_input_file_path
excel_files <- list.files(path = file_path, full.names = TRUE)
output_file_name <- paste0(output_file_path, "/Processed_ADA.xlsx")


list_of_dataframes <- lapply(excel_files, read_excel)

# Data processing --------------------------------------------------------------

# Function to process each dataframe and capture warnings
process_dataframe <- function(df, df_name) {
  warnings <- NULL
  processed_df <- withCallingHandlers({
    # Extract date from the first column (assumes date is always in the first row and column)
    date_string <- names(df)[1]
    date_extracted <- as.Date(sub(".*on\\s([0-9]{2}/[0-9]{2}/[0-9]{4}).*", "\\1", date_string),
                              format = "%d/%m/%Y")
    
    # Find the row with the string 'Site' and use it as the header
    site_row_index <- which(df[[1]] == 'Site')
    names(df) <- as.character(df[site_row_index, ])
    df <- df[(site_row_index + 1):nrow(df), ]
    
    # Remove rows with any NA values
    df <- df %>%
      drop_na() %>%
      mutate(Date = date_extracted,
             Day = lubridate::wday(Date, label = TRUE, abbr = FALSE)) %>%
      mutate(
        `Patients In ADA` = as.numeric(`Patients In ADA`),
        `Patients in ADA Over 4 Hours` = as.numeric(`Patients in ADA Over 4 Hours`)
      ) %>%
      rename(Provider = Site) %>%
      pivot_longer(
        cols = c("Patients In ADA", "Patients in ADA Over 4 Hours"),
        names_to = "Metric",
        values_to = "Value"
      ) %>%
      mutate(
        `Metric Category Type` = "Total",
        `Metric Category Name` = "Total",
        `File Name` = df_name
      ) %>%
      select(Metric, `Metric Category Type`, `Metric Category Name`,
             Provider, Date, Day, Value, `File Name`)
    
    return(df)
  }, warning = function(w) {
    warnings <<- c(warnings, w$message)
    invokeRestart("muffleWarning") # Suppress warnings but capture them
  })
  
  # If warnings were captured, print the file name and the warnings
  if (!is.null(warnings)) {
    message(paste("Warning(s) occurred in file:", df_name))
    message(paste("Warning details:", warnings, collapse = "\n"))
  }
  
  return(processed_df)
}

# Apply the custom function to each df in the list and include the filenames
list_of_processed_dfs <- lapply(seq_along(list_of_dataframes), function(i) {
  process_dataframe(list_of_dataframes[[i]], basename(excel_files[i]))
})

# Filter out any NULL results (if any errors occurred)
list_of_processed_dfs <- Filter(Negate(is.null), list_of_processed_dfs)


# Combine the data from all dfs into a main table
ADA_output <- bind_rows(list_of_processed_dfs) %>% 
  dplyr::arrange(Date) %>%
  clean_names(case = "title",
              abbreviations = c("ADA")) %>% 
  filter(!(is.na(Date)) &
           Date > max_date) %>% 
  mutate(`File Name` = 'Daily ADA Report')

# Write the processed data frame to a new sheet in an Excel workbook
write_xlsx(ADA_output, path = output_file_name)

# Move files from 'New data` to 'Old data' folder`
move_all_files(main_directory = paste0(input_file_path, "/ADA"), source_folder = "New data", destination_folder = "Old data")

