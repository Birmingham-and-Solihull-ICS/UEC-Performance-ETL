# Initialize necessary variables and paths
file_path <- BCHC_UCR_input_file_path
excel_files <- list.files(path = file_path, full.names = TRUE)
output_file_name <- paste0(output_file_path, "/Processed_BCHC_UCR.xlsx")

# Function to process data with required column checks --------------------------------------------------------------
process_bchc_ucr_data <- function(data, WMAS_filter = NULL, breach_filter = NULL, team_filter = NULL, metric) {
  # Check if necessary columns exist before processing
  required_cols <- c("Referral Date", "WMAS", "Team", "Breach")
  if (!all(required_cols %in% names(data))) {
    stop("Required columns are missing")
  }
  
  if (!is.null(WMAS_filter)) {
    data <- filter(data, WMAS == WMAS_filter)
  }
  if (!is.null(team_filter)) {
    data <- filter(data, Team == team_filter)
  }
  if (!is.null(breach_filter)) {
    data <- filter(data, Breach == breach_filter)
  }
  
  data %>%
    group_by(`Referral Date`) %>%
    summarise(Value = n()) %>%
    mutate(Metric = metric,
           `Metric Category Type` = "Total",
           `Metric Category Name` = "Total",
           Provider = "BCHC") %>%
    rename(Date = `Referral Date`) %>%
    mutate(Date = as.Date(Date),
           Day = wday(Date, label = TRUE, abbr = FALSE),
           `File Name` = "BCHC UCR") %>%
    select(Metric, `Metric Category Type`, `Metric Category Name`, Provider, Date, Day, Value, `File Name`)
}

# Function to calculate 2hr performance with dynamic column references ---------------------------------------------
calc_bchc_ucr_2hr_performance <- function(numerator_data, denominator_data, numerator_metric, denominator_metric) {
  combined_data <- bind_rows(numerator_data, denominator_data)
  
  performance_data <- combined_data %>%
    pivot_wider(
      names_from = Metric,
      values_from = Value
    )
  
  # Check if the required columns exist before mutating
  if (!(numerator_metric %in% names(performance_data)) || !(denominator_metric %in% names(performance_data))) {
    stop(paste("Column", numerator_metric, "or", denominator_metric, "not found"))
  }
  
  performance_data <- performance_data %>%
    mutate(`2hr Performance` = !!sym(numerator_metric) / !!sym(denominator_metric)) %>%
    pivot_longer(
      cols = c(7:9),
      names_to = "Metric",
      values_to = "Value"
    ) %>%
    mutate(`Metric Category Type` = case_when(
      Metric == "2hr Performance" ~ denominator_metric,
      TRUE ~ "Total"
    )) %>%
    mutate(`Metric Category Name` = case_when(
      Metric == "2hr Performance" ~ "Percentage",
      TRUE ~ "Total"
    )) %>%
    select(Metric, `Metric Category Type`, `Metric Category Name`, Provider, Date, Day, Value, `File Name`)
  
  return(performance_data)
}

# Initialize an empty list to store processed data
all_processed_data <- list()

# Initialize a list to store error messages
error_log <- list()

# Loop through each Excel file with error handling
for (excel_file in excel_files) {
  # Use tryCatch to handle errors and skip problematic files
  tryCatch({
    # Read data from the "Detail" sheet
    data <- read_xlsx(excel_file, sheet = "Detail")
    
    # Apply process_bchc_ucr_data function to create the various metrics
    total_ucr_referrals <- process_bchc_ucr_data(data, metric = "Total UCR referrals")
    total_ucr_referrals_without_breaches <- process_bchc_ucr_data(data, breach_filter = "0", metric = "Total UCR referrals without breaches")
    total_ucr_referrals_with_breaches <- process_bchc_ucr_data(data, breach_filter = "1", metric = "Total UCR referrals with breaches")
    
    DN_team <- process_bchc_ucr_data(data, team_filter = "DN", metric = "Total DN team referrals")
    DN_team_without_breaches <- process_bchc_ucr_data(data, team_filter = "DN", breach_filter = "0", metric = "Total DN team referrals without breaches")
    DN_team_with_breaches <- process_bchc_ucr_data(data, team_filter = "DN", breach_filter = "1", metric = "Total DN team referrals with breaches")
    
    ucr_team <- process_bchc_ucr_data(data, team_filter = "UCR", metric = "Total UCR team referrals")
    ucr_team_without_breaches <- process_bchc_ucr_data(data, team_filter = "UCR", breach_filter = "0", metric = "Total UCR team referrals without breaches")
    ucr_team_with_breaches <- process_bchc_ucr_data(data, team_filter = "UCR", breach_filter = "1", metric = "Total UCR team referrals with breaches")
    
    total_ucr_referrals_via_WMAS <- process_bchc_ucr_data(data, WMAS_filter = "1", metric = "Total UCR referrals via WMAS")
    total_ucr_referrals_via_WMAS_without_breaches <- process_bchc_ucr_data(data, WMAS_filter = "1", breach_filter = "0", metric = "Total UCR referrals via WMAS without breaches")
    total_ucr_referrals_via_WMAS_with_breaches <- process_bchc_ucr_data(data, WMAS_filter = "1", breach_filter = "1", metric = "Total UCR referrals via WMAS with breaches")
    
    # Calculate 2hr performance
    total_ucr_referrals_performance <- calc_bchc_ucr_2hr_performance(
      numerator_data = total_ucr_referrals_without_breaches,
      denominator_data = total_ucr_referrals,
      numerator_metric = "Total UCR referrals without breaches",
      denominator_metric = "Total UCR referrals"
    )
    
    DN_team_performance <- calc_bchc_ucr_2hr_performance(
      numerator_data = DN_team_without_breaches,
      denominator_data = DN_team,
      numerator_metric = "Total DN team referrals without breaches",
      denominator_metric = "Total DN team referrals"
    )
    
    UCR_team_performance <- calc_bchc_ucr_2hr_performance(
      numerator_data = ucr_team_without_breaches,
      denominator_data = ucr_team,
      numerator_metric = "Total UCR team referrals without breaches",
      denominator_metric = "Total UCR team referrals"
    )
    
    total_UCR_referrals_via_WMAS_performance <- calc_bchc_ucr_2hr_performance(
      numerator_data = total_ucr_referrals_via_WMAS_without_breaches,
      denominator_data = total_ucr_referrals_via_WMAS,
      numerator_metric = "Total UCR referrals via WMAS without breaches",
      denominator_metric = "Total UCR referrals via WMAS"
    )
    
    # Combine data from this file
    combined_file_data <- bind_rows(
      total_ucr_referrals_performance, total_ucr_referrals_with_breaches,
      DN_team_performance, DN_team_with_breaches,
      UCR_team_performance, ucr_team_with_breaches,
      total_UCR_referrals_via_WMAS_performance, total_ucr_referrals_via_WMAS_with_breaches
    )
    
    # Add the processed data to the main list
    all_processed_data[[basename(excel_file)]] <- combined_file_data
    
  }, error = function(e) {
    # Log the error with the file name
    error_message <- paste("Error processing file:", basename(excel_file), "-", e$message)
    message(error_message)  # Print the error message to the console
    
    # Store the error message in the error log for later inspection
    error_log[[basename(excel_file)]] <- e$message
  })
}

# Combine all processed data frames into one
final_output <- bind_rows(all_processed_data)

# Clean and arrange the output
BCHC_ucr_output <- final_output %>%
  dplyr::arrange(Date) %>%
  clean_names(case = "title") %>% 
  filter(Date > max_date)

# Print the error log (if any)
if (length(error_log) > 0) {
  message("The following errors occurred:")
  print(error_log)
} else {
  message("No errors occurred.")
}


# Write the processed data frame in an Excel workbook
write_xlsx(BCHC_ucr_output, path = output_file_name)

# Move files from 'New data` to 'Old data' folder`
move_all_files(main_directory = paste0(input_file_path, "/BCHC UCR"), source_folder = "New data", destination_folder = "Old data")




