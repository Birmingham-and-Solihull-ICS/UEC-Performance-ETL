
# Read whichever files in the New data folder
# Process these files by running R script
# Once processsed, move these files to the old folder
# The new data folder will contain next data

max_date <- as.date("2024-03-25")

# Path to the Excel files
file_path <- WMAS_input_file_path 

# Get a list of all Excel files in the folder (.xlsx and .xls)
excel_files <- list.files(path = file_path, pattern = "\\.xlsx$|\\.xls$", full.names = TRUE)

# process_data <- function(excel_file){
#   
#   # Read raw Excel file
#   data <- read_xlsx(excel_file[1])
#   
#   # Extract date
#   date <- data %>%
#     select(`Operational Handover Durations`) %>%
#     slice(1:2) %>%
#     mutate(Date = str_extract(`Operational Handover Durations`, "\\d{2}/\\d{2}/\\d{4}")) %>%
#     filter(!is.na(Date)) %>%
#     pull(Date) %>%
#     first() 
#   
#   # Process handover duration data
#   handover_duration_data <- data %>%
#     select(2:ncol(data)) %>% 
#     janitor::row_to_names(row_number = which(data$...2 == "Hospital")) %>% 
#     select(where(~ any(!is.na(.)))) %>% 
#     select(`Hospital`, c(5:19)) %>%  
#     pivot_longer(
#       cols = -Hospital,
#       names_to = "Metric Category Type",
#       values_to = "Value"
#     ) %>% 
#     mutate(Metric = "Operational Handover Duration",
#            `Metric Category Name` = "Total") %>% 
#     select(Metric, `Metric Category Type`, `Metric Category Name`, Hospital, Value) %>% 
#     as_tibble()
#   
#   # Process average & maximum handover duration data
#   avg_max_handover_duration_data <- data %>%
#     select(2:ncol(data)) %>% 
#     janitor::row_to_names(row_number = which(data$...2 == "Hospital")) %>% 
#     select(where(~ any(!is.na(.)))) %>% 
#     select(Hospital, c(20:21)) %>% 
#     clean_names(case = "title") %>% 
#     pivot_longer(
#       cols = -Hospital,
#       names_to = "Metric",
#       values_to = "Value"
#     ) %>%
#     # Separate the time string into hours, minutes, and seconds
#     separate(Value, into = c("Hours", "Minutes", "Seconds"), sep = ":", remove = FALSE) %>%
#     mutate(Hours = as.character(as.integer(Hours))) %>% 
#     select(-Value) %>% 
#     pivot_longer(
#       cols = c("Hours", "Minutes", "Seconds"),
#       names_to = "Metric Category Type",
#       values_to = "Value"
#     ) %>%
#     mutate(`Metric Category Name` = "Total") %>% 
#     select(Metric, `Metric Category Type`, `Metric Category Name`, Hospital, Value) %>% 
#     as_tibble()
#   
#   # Process non-duration data
#   non_duration_data <- data %>%
#     select(2:ncol(data)) %>% 
#     janitor::row_to_names(row_number = which(data$...2 == "Hospital")) %>% 
#     select(where(~ any(!is.na(.)))) %>% 
#     select(-c(5:19, 20:21)) %>% 
#     rename(`Handovers Recorded` = Total,
#            `Handovers Recorded Percent` = `%`) %>%
#     clean_names(case = "title") %>%
#     pivot_longer(
#       cols = -Hospital,
#       names_to = "Metric",
#       values_to = "Value"
#     ) %>% 
#     mutate(`Metric Category Type` = case_when(
#       str_detect(`Metric`, "Percent") ~ "Percentage",
#       TRUE ~ "Total"
#     ))%>%
#     mutate(`Metric Category Name` = case_when(
#       `Metric Category Type` == "Percentage" ~ "Percentage",
#       TRUE ~ "Total"
#     )) %>%
#     select(Metric, `Metric Category Type`, `Metric Category Name`, Hospital, Value) %>% 
#     as_tibble()
#   
#   # Combine data
#   data <- bind_rows(handover_duration_data, avg_max_handover_duration_data, non_duration_data)
#   
#   # Add Date, Day and other columns
#   data <- data %>% 
#     mutate(Date = dmy(date),
#            Day = wday(Date, label = TRUE, abbr = FALSE)) %>%
#     rename(`Provider` = Hospital) %>% 
#     mutate(Value =as.numeric(Value)) %>% 
#   mutate(`File Name` = "WMAS") %>% 
#     select(Metric, `Metric Category Type`, `Metric Category Name`, `Provider`, Date, Day, Value, `File Name`)
#   
#   return(data)
# }

# Function to process each Excel file
process_data <- function(excel_file) { 
  # Read raw Excel file
  data <- read_xlsx(excel_file)
  
  # Extract date
  date <- data %>%
    select(`Operational Handover Durations`) %>%
    slice(1:2) %>%
    mutate(Date = str_extract(`Operational Handover Durations`, "\\d{2}/\\d{2}/\\d{4}")) %>%
    filter(!is.na(Date)) %>%
    pull(Date) %>%
    first()
  
  # Process handover duration data
  handover_duration_data <- data %>%
    select(2:ncol(data)) %>%
    janitor::row_to_names(row_number = which(data$...2 == "Hospital")) %>%
    select(where(~ any(!is.na(.)))) %>%
    select(`Hospital`, c(5:19)) %>%
    pivot_longer(
      cols = -Hospital,
      names_to = "Metric Category Type",
      values_to = "Value"
    ) %>%
    mutate(Metric = "Operational Handover Duration",
           `Metric Category Name` = "Total") %>%
    select(Metric, `Metric Category Type`, `Metric Category Name`, Hospital, Value) %>%
    as_tibble()
  
  # Process average & maximum handover duration data
  avg_max_handover_duration_data <- data %>%
    select(2:ncol(data)) %>%
    janitor::row_to_names(row_number = which(data$...2 == "Hospital")) %>%
    select(where(~ any(!is.na(.)))) %>%
    select(Hospital, c(20:21)) %>%
    clean_names(case = "title") %>%
    pivot_longer(
      cols = -Hospital,
      names_to = "Metric",
      values_to = "Value"
    ) %>%
    separate(Value, into = c("Hours", "Minutes", "Seconds"), sep = ":", remove = FALSE) %>%
    mutate(Hours = as.character(as.integer(Hours))) %>%
    select(-Value) %>%
    pivot_longer(
      cols = c("Hours", "Minutes", "Seconds"),
      names_to = "Metric Category Type",
      values_to = "Value"
    ) %>%
    mutate(`Metric Category Name` = "Total") %>%
    select(Metric, `Metric Category Type`, `Metric Category Name`, Hospital, Value) %>%
    as_tibble()
  
  # Process non-duration data
  non_duration_data <- data %>%
    select(2:ncol(data)) %>%
    janitor::row_to_names(row_number = which(data$...2 == "Hospital")) %>%
    select(where(~ any(!is.na(.)))) %>%
    select(-c(5:19, 20:21)) %>%
    rename(`Handovers Recorded` = Total, `Handovers Recorded Percent` = `%`) %>%
    clean_names(case = "title") %>%
    pivot_longer(
      cols = -Hospital,
      names_to = "Metric",
      values_to = "Value"
    ) %>%
    mutate(`Metric Category Type` = case_when(
      str_detect(`Metric`, "Percent") ~ "Percentage",
      TRUE ~ "Total"
    )) %>%
    mutate(`Metric Category Name` = case_when(
      `Metric Category Type` == "Percentage" ~ "Percentage",
      TRUE ~ "Total"
    )) %>%
    select(Metric, `Metric Category Type`, `Metric Category Name`, Hospital, Value) %>%
    as_tibble()
  
  # Combine data
  combined_data <- bind_rows(handover_duration_data, avg_max_handover_duration_data, non_duration_data)
  
  # Add Date, Day, and other columns
  combined_data <- combined_data %>%
    mutate(Date = dmy(date),  # Convert extracted date string to Date
           Day = wday(Date, label = TRUE, abbr = FALSE)) %>%
    rename(`Provider` = Hospital) %>%
    mutate(Value = as.numeric(Value)) %>%
    mutate(`File Name` = "WMAS") %>%
    select(Metric, `Metric Category Type`, `Metric Category Name`, `Provider`, Date, Day, Value, `File Name`)
  
  return(combined_data)
}


# Apply function to to a list of Excel files in the WMAS folder
list_of_processed_dfs <- lapply(excel_files, process_data)

# Combine all processed data frames into one data frame
output <- bind_rows(list_of_processed_dfs) %>% 
  dplyr::arrange(Date) %>%
  clean_names(case = "title") %>% 
  filter(Date > max_date)

# Output Excel file path
output_file_name <- paste0(output_file_path, "/Processed_WMAS.xlsx")

# Write the processed data frame in an Excel workbook
write_xlsx(output, path = output_file_name) 

# Move the raw data to "Old data" folder once data processing is finished
## Define the destination path explicitly as a subdirectory of `file_path`
destination_path <- paste0(file_path, "/Old data")

# Ensure the destination directory exists; create it if it doesn't
if (!dir.exists(destination_path)) {
  dir.create(destination_path)
}

# List the Excel file(s) we intend to move
excel_files <- list.files(path = paste0(input_file_path, "/WMAS"), pattern = "\\.xlsx$", full.names = TRUE)


# Check if there are Excel files to move
if (length(excel_files) > 0) {
  for (file in excel_files) {
    
    # Move the file to "Old data" folder
    file_move(file, destination_path)
  }
} else {
  message("No files found to move.")
}