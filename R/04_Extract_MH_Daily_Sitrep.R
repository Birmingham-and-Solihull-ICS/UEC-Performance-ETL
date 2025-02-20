
# Initialize necessary variables and paths
file_path <- BSMHFT_Sitrep_input_file_path
excel_files <- list.files(path = file_path, full.names = TRUE)
output_file_name <- paste0(output_file_path, "/Processed_BSMHFT_Sitrep.xlsx")

# Convert date columns ---------------------------------------------------------

## Two formats found: Excel serial numbers ("45593" - number of days since January 1, 1900) 
## and Unix Timestamps ("1650153600" - number of seconds since January 1, 1970)

# Function to detect and convert date formats
convert_date <- function(date) {
  date <- as.numeric(date)  # Ensure date is numeric
  sapply(date, function(d) {
    if (d > 1e9) {  # likely a Unix timestamp
      formatted_date <- as.POSIXct(d, origin = "1970-01-01", tz = "UTC")
    } else {  # likely an Excel date serial
      formatted_date <- as_date(d, origin = "1899-12-30")  # Excel's base date
    }
    return(format(formatted_date, "%Y-%m-%d"))
  })
}

#1. Read Excel MH sitrep data and clean dates columns --------------------------

# Load the data and get column names

MH_data <- read_xlsx(excel_files[1]) %>% 
  janitor::row_to_names(3)

# Get the original column names from 5 to the last column
date_columns <- colnames(MH_data)[5:ncol(MH_data)]

# Convert date column names
new_date_columns <- convert_date(date_columns)

# Rename the original columns from 5 to the last column with the new date names
colnames(MH_data)[5:ncol(MH_data)] <- new_date_columns

anyDuplicated(colnames(MH_data))
print(colnames(MH_data)[duplicated(colnames(MH_data))])


# 2. Process simple metrics without sites --------------------------------------

#1. MFFD Transferred Out
#2. MH Support for older people stepped down into P2
#3. 12 hour breaches in ED
#4. Total awaiting admission in community
#5. PDU referrals
#6. POS referrals
#7. POS Transfers from ED


process_MH_nonsite_metrics <- function(data) {
  # Data frame of metrics and corresponding BSMHFT leads
  metrics_leads <- data.frame(
    Metric = c(
      "MFFD Transferred Out",
      "MH Support for older people stepped down into P2",
      "12 hour breaches in ED",
      "Total awaiting admission in community",
      "PDU referrals",
      "POS referrals",
      "POS Transfers from ED"
    ),
    BSMHFT_Lead = c(
      "Bed Management",
      NA_character_,
      "Informatics",
      "Bed Management",
      "Bed Management",
      "Bed Management",
      "Bed Management"
    ),
    stringsAsFactors = FALSE
  )
  
  # Initialize a list to store processed metrics
  processed_data_list <- list()
  
  # Loop through each metric and process it
  for (i in 1:nrow(metrics_leads)) {
    metric_name <- metrics_leads$Metric[i]
    BSMHFT_lead <- metrics_leads$BSMHFT_Lead[i]
    
    # Filter data and attach BSMHFT Lead
    processed_data <- data %>%
      filter(Metric == metric_name) %>%
      mutate(`BSMHFT Lead` = BSMHFT_lead) %>%  # Set the BSMHFT Lead value
      mutate(
        across(.cols = 5:ncol(data),
               .fns = ~as.character(.))  # Convert date columns to character
      ) %>%
      mutate(across(.cols = 5:ncol(data),
                    .fns = ~ ifelse(is.na(.) | !str_detect(., "^[0-9.]+$"), NA_real_, as.numeric(.)))) %>%
      pivot_longer(
        cols = -c("No.", "Definition", 'Metric', 'BSMHFT Lead'),
        names_to = "Date",
        values_to = "Value"
      ) %>%
      mutate(
        Date = as.Date(Date, format = "%Y-%m-%d"),
        Value = ifelse(Value %in% c("NA", "N/A", "na"), NA_real_, Value),
        Value = ifelse(is.na(Value), NA_real_, as.numeric(Value)),
        Day = wday(Date, label = TRUE, abbr = FALSE),
        `File Name` = "BSMHFT Daily MH SITREP"
      ) %>%
      mutate(`Metric Category Type` = "Total",
             `Metric Category Name` = "Total",
             Provider = "System") %>%
      select(Metric, `Metric Category Type`, `Metric Category Name`, 
             `BSMHFT Lead`, Provider, Date, Day, Value, `File Name`)
    
    # Add the processed data to the list
    processed_data_list[[metric_name]] <- processed_data
  }
  
  # Combine all processed data into one data frame
  final_data <- bind_rows(processed_data_list)
  
  return(final_data)
}


#3. Process metrics with sites -------------------------------------------------

# Process "Total ED awaiting assessment" and other similar metrics -------------
process_MH_ED_metrics <- function(data, row_value_start, row_value_end, BSMHFT_Lead, metric_name) {
  start_row <- which(data$Metric == row_value_start)[1]
  end_row <- which(data$Metric == row_value_end)[1]
  
  data %>%
    slice(start_row: (end_row - 1)) %>%
    mutate(`BSMHFT Lead` = BSMHFT_Lead) %>%
    as_tibble() %>%
    mutate(
      across(.cols = 3:ncol(data),
             .fns = ~as.character(.))
    ) %>%
    pivot_longer(
      cols = -c('No.', 'Metric','Definition', 'BSMHFT Lead'),
      names_to = "Date",
      values_to = "Value"
    ) %>%
    mutate(
      Date = as.Date(Date, format = "%Y-%m-%d"),
      # Replace "NA" with NA_real_ before conversion
      Value = ifelse(Value %in% c("NA", "N/A", "na"), NA_real_, Value), 
      Value = as.numeric(Value),  # Convert to numeric
      Day = wday(Date, label = TRUE, abbr = FALSE),
      `File Name` = "BSMHFT Daily MH SITREP"
    ) %>%
    mutate(Provider = case_when(
      grepl("City", Metric) ~ "City",
      grepl("Good Hope", Metric) ~ "Good Hope",
      grepl("Heartlands", Metric) ~ "Heartlands",
      grepl("QE/UHB", Metric) ~ "QE/UHB",
      TRUE ~ "System"
    )) %>%
    mutate(`Metric Category Type` = "Total",
           `Metric Category Name` = "Total",
           Metric = metric_name) %>%
    select(Metric, `Metric Category Type`, `Metric Category Name`,
           `BSMHFT Lead`, Provider, Date, Day, Value, `File Name`)
}

#4. Process Delayed Discharges ---------------------------------------------------

process_MH_delayed_discharges <- function(data) {
  # Define metric categories
  metric_categories <- c("Delay Discharges Total",
                         "Adult Acute Total", "Adult Acute - NHS", "Adult Acute - Social Care",
                         "Adult Acute - Both Health and Social Care", "Adult Acute - Housing",
                         "Adult Acute - Other",
                         "Older Adult Total", " Older Adult - NHS", "Older Adult - Social Care",
                         "Older Adult - Both Health and Social Care", "Older Adult - Housing")
  
  # Locate the rows corresponding to the start and end of the "Delayed Discharges" section
  start_row <- which(data$Metric == "Delayed Discharges:")[1]
  end_row <- which(data$Metric == "POS Transfers from ED")[1] - 1
  
  # Process delayed discharges data
  delayed_discharges_dt <- data%>%
    slice(start_row:end_row) %>%
    mutate(`BSMHFT Lead` = "Informatics") %>%
    mutate(Metric = metric_categories) %>%
    as_tibble() %>%
    mutate(
      across(.cols = 3:ncol(data),
             .fns = ~as.character(.))) %>%
    pivot_longer(
      -c('No.', 'Metric','Definition', 'BSMHFT Lead'),
      names_to = "Date",
      values_to = "Value"
    ) %>%
    mutate(
      Date = as.Date(Date, format = "%Y-%m-%d"),
      Value = ifelse(Value %in% c("NA", "N/A", "na"), NA_real_, Value),
      Value = as.numeric(Value),  # Convert to numeric
      Day = wday(Date, label = TRUE, abbr = FALSE),
      `File Name` = "BSHMFT Daily MH SITREP"
    ) %>%
    mutate(`Metric Category Type` = case_when(
      grepl("Adult Acute", Metric) ~ "Adult Acute",
      grepl("Older Adult", Metric) ~ "Older Adult",
      TRUE ~ "Total"
    )) %>%
    mutate(`Metric Category Name` = case_when(
      grepl(" - NHS", Metric) ~ "NHS",
      grepl(" - Social Care", Metric) ~ "Social Care",
      grepl(" - Both Health and Social Care", Metric) ~ "Both Health and Social Care",
      grepl(" - Housing", Metric) ~ "Housing",
      grepl(" - Other", Metric) ~ "Other",
      TRUE ~ "Total"
    )) %>%
    mutate(Metric = "Delayed Discharges",
           Provider = case_when(
             (`Metric Category Type` == "Total" & `Metric Category Name` == "Total") ~ "System",
             TRUE ~ NA
           )) %>%
    select(Metric, `Metric Category Type`, `Metric Category Name`,
           `BSMHFT Lead`, Provider, Date, Day, Value, `File Name`)
  
  # Return the processed delayed discharges data
  return(delayed_discharges_dt)
}



#5. Process discharges -----------------------------------------------------------

process_MH_discharges <- function(data) {
  start_row <- which(data$Metric == "Discharges:")[1]
  end_row <- nrow(data)
  
  discharges_dt <- data %>%
    slice(start_row:end_row) %>%
    mutate(`BSMHFT Lead` = "Informatics") %>%
    mutate(Metric = c("Discharges Total", "BSMHFT Total",
                      "Adult Acute - BSMHFT", "Older Adult - BSMHFT",
                      "Out of Area Total", "Adult Acute - Out of Area")) %>%
    as_tibble() %>%
    mutate(across(.cols = 3:ncol(data),
                  .fns = ~as.character(.))) %>%
    pivot_longer(
      cols = -c('No.', 'Metric','Definition', 'BSMHFT Lead'),
      names_to = "Date",
      values_to = "Value"
    ) %>%
    mutate(
      Date = as.Date(Date, format = "%Y-%m-%d"),
      Value = ifelse(Value %in% c("NA", "N/A", "na"), NA_real_, Value), 
      Value = as.numeric(Value),
      Day = wday(Date, label = TRUE, abbr = FALSE),
      Organisation = "BSMHFT",
      `File Name` = "BSMHFT Daily MH SITREP") %>%
    mutate(`Metric Category Type` = case_when(
      grepl("BSMHFT", Metric) ~ "BSMHFT",
      grepl("Out of Area", Metric) ~ "Out of Area",
      TRUE ~ "Total"
    )) %>%
    mutate(`Metric Category Name` = case_when(
      grepl("Adult Acute - ", Metric) ~ "Adult Acute",
      grepl("Older Adult - ", Metric) ~ "Older Adult",
      TRUE ~ "Total"
    )) %>%
    mutate(Metric = "Discharges",
           Provider = case_when(
             (`Metric Category Type` == "Total" & `Metric Category Name` == "Total") ~ "System",
             TRUE ~ NA
           )) %>%
    select(Metric, `Metric Category Type`, `Metric Category Name`,
           `BSMHFT Lead`, Provider, Date, Day, Value, `File Name`)
  
  return(discharges_dt)
}



#6. Process total IP MH that are MFFD --------------------------------------------

process_MH_total_IP_MFFD <- function(data) {
  start_row <- which(data$Metric == "Total IP MH that are MFFD:")[1]
  end_row <- which(data$Metric == "S136 patients in ED:")[1]
  
  total_IP_MH_MFFD <- data %>%
    slice(start_row:(end_row - 1)) %>%
    mutate(`BSMHFT Lead` = "Bed Management") %>%
    mutate(Metric = c("Total IP MH that are MFFD", "City Total",
                      "City - ED", "City - Other wards", "Good Hope Total",
                      "Good Hope - ED", "Good Hope - Other wards",
                      "Heartlands Total", "Heartlands - ED", "Heartlands - Other wards",
                      "QE/UHB Total", "QE/UHB - ED", "QE/UHB - Other wards")) %>%
    as_tibble() %>%
    mutate(across(.cols = 3:ncol(data),
                  .fns = ~as.character(.))) %>%
    pivot_longer(
      cols = -c('No.', 'Metric','Definition', 'BSMHFT Lead'),
      names_to = "Date",
      values_to = "Value"
    ) %>%
    mutate(
      Date = as.Date(Date, format = "%Y-%m-%d"),
      Value = ifelse(Value %in% c("NA", "N/A", "na"), NA_real_, Value),
      Value = as.numeric(Value),
      Day = wday(Date, label = TRUE, abbr = FALSE),
      `File Name` = "BSMHFT Daily MH SITREP") %>%
    mutate(Provider = case_when(
      grepl("City", Metric) ~ "City",
      grepl("Good Hope", Metric) ~ "Good Hope",
      grepl("Heartlands", Metric) ~ "Heartlands",
      grepl("QE/UHB", Metric) ~ "QE/UHB",
      TRUE ~ "System"
    )) %>%
    mutate(`Metric Category Type` = case_when(
      grepl(" - ED", Metric) ~ "ED",
      grepl(" - Other wards", Metric) ~ "Other wards",
      TRUE ~ "Total"
    )) %>%
    mutate(Metric = "Total IP MH that are MFFD",
           `Metric Category Name` = "Total") %>%
    select(Metric, `Metric Category Type`, `Metric Category Name`,
           `BSMHFT Lead`, Provider, Date, Day, Value, `File Name`)
  
  return(total_IP_MH_MFFD)
}



#7. Process total IP MH ----------------------------------------------------------

process_MH_total_IP <- function(data) {
  start_row <- which(data$Metric == "Total IP MH:")[1]
  end_row <- which(data$Metric == "Total IP MH that are MFFD:")[1]
  
  total_IP_MH <- data %>%
    slice(start_row:(end_row - 1)) %>%
    mutate(`BSMHFT Lead` = "Bed Management") %>%
    mutate(Metric = c("Total IP MH", "City Total",
                      "City - ED", "City - Other wards", "Good Hope Total",
                      "Good Hope - ED", "Good Hope - Other wards",
                      "Heartlands Total", "Heartlands - ED", "Heartlands - Other wards",
                      "QE/UHB Total", "QE/UHB - ED", "QE/UHB - Other wards")) %>%
    as_tibble() %>%
    mutate(across(.cols = 3:ncol(data),
                  .fns = ~as.character(.))) %>%
    pivot_longer(
      cols = -c('No.', 'Metric','Definition', 'BSMHFT Lead'),
      names_to = "Date",
      values_to = "Value"
    ) %>%
    mutate(
      Date = as.Date(Date, format = "%Y-%m-%d"),
      Value = str_trim(Value),
      ValidValue = grepl("^[0-9]+(\\.[0-9]+)?$", Value),  # Validate numeric patterns
      Value = ifelse(ValidValue, as.numeric(Value), NA_real_),  # Convert valid values
      # Value = ifelse(Value %in% c("NA", "N/A", "na"), NA_real_, Value),
      # Value = as.numeric(Value),
      Day = wday(Date, label = TRUE, abbr = FALSE),
      `File Name` = "BSMHFT Daily MH SITREP") %>%
    mutate(Provider = case_when(
      grepl("City", Metric) ~ "City",
      grepl("Good Hope", Metric) ~ "Good Hope",
      grepl("Heartlands", Metric) ~ "Heartlands",
      grepl("QE/UHB", Metric) ~ "QE/UHB",
      TRUE ~ "System"
    )) %>%
    mutate(`Metric Category Type` = case_when(
      grepl(" - ED", Metric) ~ "ED",
      grepl(" - Other wards", Metric) ~ "Other wards",
      TRUE ~ "Total"
    )) %>%
    mutate(Metric = "Total IP MH",
           `Metric Category Name` = "Total") %>%
    select(Metric, `Metric Category Type`, `Metric Category Name`,
           `BSMHFT Lead`, Provider, Date, Day, Value, `File Name`)
  
  return(total_IP_MH)
}

# process_MH_total_IP(MH_data) %>% View()


#8. Process  % seen within target ----------------------------------------------

process_MH_pct_seen_within_target <- function(data) {
  # Get the start and end rows
  start_row <- which(data$Metric == "% seen within 1 hour assessment target in ED:")[1]
  end_row <- which(data$Metric == "Total IP MH:")[1] - 1
  
  # Slice the relevant data
  pct_seen_within_target <- data %>%
    slice(start_row:end_row) 
  
  # Process columns with mixed numbers and percentages separately
  pct_cols_to_modify <- pct_seen_within_target %>%
    mutate(across(.cols = 3:ncol(data),
                  .fns = ~as.character(.))) %>%
    select(where(is.character)) %>% 
    pivot_longer(
      cols = -c(`No.`, `Definition`, `Metric`, `BSMHFT Lead`),
      names_to = "Date",
      values_to = "Value"
    ) %>%
    # Extract numbers within brackets and percentages, derive Numerator and Denominator
    mutate(
      Denominator = ifelse(is.na(Value), NA, as.numeric(str_match(Value, "\\((\\d+)\\)")[,2])),
      Percentage = ifelse(is.na(Value), NA, as.numeric(str_match(Value, "(\\d+)%")[,2])) / 100,
      Numerator = as.integer(Percentage * Denominator)
    ) %>%
    select(Metric, `BSMHFT Lead`, Date, Numerator, Denominator, Percentage) %>%
    pivot_longer(
      cols = -c(Metric, `BSMHFT Lead`, Date),
      names_to = "Metric Category Type",
      values_to = "Value"
    ) %>%
    select(Metric, `Metric Category Type`, `BSMHFT Lead`, Date, Value)
  
  # Combine with the rest of the numeric columns and process as normal
  pct_seen_within_target <- pct_seen_within_target %>%
    as_tibble() %>%
    select(Metric, `BSMHFT Lead`, where(is.numeric)) %>%
    pivot_longer(
      cols = -c('Metric', 'BSMHFT Lead'),
      names_to = "Date",
      values_to = "Value"
    ) %>%
    mutate(`Metric Category Type` = "Percentage") %>%
    select(Metric, `Metric Category Type`, `BSMHFT Lead`, Date, Value)
  
  # Combine both 
  output <- bind_rows(pct_cols_to_modify, pct_seen_within_target)
  
  # Add extra columns
  output <- output %>%
    mutate(
      Date = as.Date(Date, format = "%Y-%m-%d"),
      Value = ifelse(Value %in% c("NA", "N/A", "na"), NA_real_, Value),
      Value = as.numeric(Value),
      Day = wday(Date, label = TRUE, abbr = FALSE),
      `File Name` = "BSMHFT Daily MH SITREP"
    ) %>%
    mutate(Provider = case_when(
      grepl("City", Metric) ~ "City",
      grepl("Good Hope", Metric) ~ "Good Hope",
      grepl("Heartlands", Metric) ~ "Heartlands",
      grepl("QE/UHB", Metric) ~ "QE/UHB",
      TRUE ~ "System"
    )) %>%
    mutate(
      Metric = "Pct seen within 1hr assessment target in ED",
      `BSMHFT Lead` = "Informatics",
      `Metric Category Name` = case_when(
        `Metric Category Type` == "Numerator" ~ "Total number of patients seen within target",
        `Metric Category Type` == "Denominator" ~ "Total number of patients presented in ED",
        `Metric Category Type` == "Percentage" ~ "Percentage number of patients seen within target"
      )
    ) %>%
    select(Metric, `Metric Category Type`, `Metric Category Name`,
           `BSMHFT Lead`, Provider, Date, Day, Value, `File Name`) %>% 
    filter(!is.na(Value)) # To remove extra rows created for missing data for Numerator and Denominator (only % is available)

  return(output)
}

#9. Process all metrics with error handling --------------------------------------

# Initialize a list to store processed metrics
processed_metrics <- list()


# Process "Nonsite Metrics" and store the result first
nonsite_metrics_result <- tryCatch({
  process_MH_nonsite_metrics(data = MH_data)
}, error = function(e) {
  message("Error processing Nonsite Metrics:", e$message)
  return(NULL)
})


# Define other metrics to process
metrics_to_process <- list(
  "Total ED awaiting assessment" = function() process_MH_ED_metrics(data = MH_data, row_value_start = "Total ED awaiting assessment:", row_value_end = "Length of wait for ED assessment:", BSMHFT_Lead = "Bed Management", metric_name = "Total ED awaiting assessment"),
  "Length of wait for ED assessment" = function() process_MH_ED_metrics(data = MH_data, row_value_start = "Length of wait for ED assessment:", row_value_end = "% seen within 1 hour assessment target in ED:", BSMHFT_Lead = "Informatics", metric_name = "Length of wait for ED assessment"),
  "S136 patients in ED" = function() process_MH_ED_metrics(data = MH_data, row_value_start = "S136 patients in ED:", row_value_end = "12 hour breaches in ED", BSMHFT_Lead = "Bed Management", metric_name = "S136 patients in ED"),
  "Delayed Discharges" = function() process_MH_delayed_discharges(data = MH_data),
  "Discharges" = function() process_MH_discharges(data = MH_data),
  "Total IP MH that are MFFD" = function() process_MH_total_IP_MFFD(data = MH_data),
  "Total IP MH" = function() process_MH_total_IP(data = MH_data),
  "Pct seen within 1hr assessment target in ED" = function() process_MH_pct_seen_within_target(data = MH_data)
)

# Apply each metric function to MH_data and collect results
processed_metrics <- map(metrics_to_process, function(metric_func) {
  tryCatch({
    metric_func()
  }, error = function(e) {
    message(paste("Error processing metric:", e$message))
    return(NULL)
  })
})

# Combine nonsite metrics with other processed metrics
if (!is.null(nonsite_metrics_result)) {
  processed_metrics[["Nonsite Metrics"]] <- nonsite_metrics_result
}


# Combine all processed metrics into a single data frame
final_output <- bind_rows(processed_metrics)

# Update year 2032 to 2032 and filter Date > max_date
MH_sitrep_output <- final_output %>% 
  mutate(Date = if_else(year(Date) == 2032,
                        update(Date, year = 2023),
                        Date))%>%
  filter(Date > max_date)


# Write data into Excel
write_xlsx(MH_sitrep_output, path = output_file_name)

# Move files from 'New data` to 'Old data' folder`
move_all_files(main_directory = paste0(input_file_path, "/BSMHFT Sitrep"), source_folder = "New data", destination_folder = "Old data")


