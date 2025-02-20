# Run this line to import the library 'UECETL' for the first time
# devtools::install_github("SitiHassan/UEC-Performance@v1.0.0", dependencies = TRUE)

library(readxl)         # For reading Excel files
library(writexl)        # For writing Excel files
library(readr)          # For reading CSV and other text data
library(tidyverse)      # Loads dplyr, tidyr, stringr, etc.
library(janitor)        # For cleaning names
library(lubridate)      # For date-time manipulation
library(fs)             # For filesystem operations
library(Microsoft365R)  # For Microsoft 365 services interaction
library(UECETL)         # My specific package for ETL
library(DBI)            # For DB connection
library(odbc)           # For DB connection

# The path to the parent work directory
directory_path = "//mlcsu-bi-fs/csugroupdata$/Commissioning Intelligence And Strategy/BSOLCCG/Reports/01_Adhoc/BSOL_1339_UEC_Performance"

# Define parameters to keep as a global variable
parameter_list <- c("run_master_script", "directory_path", "input_file_path", "output_file_path",
                    "ADA_input_file_path", "BCHC_UCR_input_file_path", "BSMHFT_Sitrep_input_file_path",
                    "BWC_Sitrep_input_file_path", "UHB_discharges_input_file_path", "UHB_Sitrep_input_file_path",
                    "VW_input_file_path", "WMAS_input_file_path", "first_day_this_month", "first_day_last_month", "max_date", "parameter_list", "start_time", "end_time", "time_taken",
                    "connection_bsol", "BWC_Weekly_Sitrep_Data", "UHB_Daily_Sitrep_Data", "Virtual_Wards_Data", "BSMHFT_Data", "WMAS_Data", "UHB_Discharges_Data",
                    "BCHC_UCR_Data", "ADA_Data", "Data_Frames", "Table_Names", "DQ_Check",
                    lsf.str())


# The path to the input excel files
input_file_path <- paste0(directory_path, "/Data/Raw Data/Recent data")
ADA_input_file_path <- paste0(input_file_path, "/ADA/New data")
BCHC_UCR_input_file_path <- paste0(input_file_path, "/BCHC UCR/New data")
BSMHFT_Sitrep_input_file_path <- paste0(input_file_path, "/BSMHFT Sitrep/New data")
BWC_Sitrep_input_file_path <- paste0(input_file_path, "/BWC Sitrep/New data")
UHB_discharges_input_file_path <- paste0(input_file_path, "/UHB Discharges/New data")
UHB_Sitrep_input_file_path <- paste0(input_file_path, "/UHB Sitrep/New data")
VW_input_file_path <- paste0(input_file_path, "/Virtual Wards/New data")
WMAS_input_file_path <- paste0(input_file_path, "/WMAS/New data")

# The path the output/processed excel files
output_file_path <- paste0(directory_path, "/Data/Processed Data/Recent data")

# Date parameter
first_day_this_month <- as.Date(format(Sys.Date(), "%Y-%m-01"))
max_date <- as.Date(format(first_day_this_month - months(1), "%Y-%m-01")) # 1 month prior to today's date

# Define a master script function to run each R script
run_master_script <- function() {


  start_time <- Sys.time()

  # Run each script and clear the environment after each
  print("Script 1: Starting to extract data from SharePoint folders..")
  source(file.path(directory_path, "R scripts/01_Download_SPT_files.R"))
  print("Finished extracting all SPT data.")
  rm(list = setdiff(ls(), parameter_list))  # Clear environment

  print("Script 2: Starting to process ADA data..")
  source(file.path(directory_path, "R scripts/02_Extract_ADA.R"))
  print("Finished processing ADA data.")
  rm(list = setdiff(ls(), parameter_list))  # Clear environment

  print("Script 3: Starting to process BWCH Sitrep data..")
  source(file.path(directory_path, "R scripts/03_Extract_BWCH_Weekly_Sitrep.R"))
  print("Finished processing BWCH Sitrep data.")
  rm(list = setdiff(ls(), parameter_list))  # Clear environment

  print("Script 4: Starting to process MH Sitrep")
  source(file.path(directory_path, "R scripts/04_Extract_MH_Daily_Sitrep.R"))
  print("Finished processing MH Sitrep.")
  rm(list = setdiff(ls(), parameter_list))  # Clear environment

  print("Script 5: Starting to process UHB Sitrep")
  source(file.path(directory_path, "R scripts/05_Extract_UHB_Daily_Sitrep.R"))
  print("Finished processing UHB Sitrep.")
  rm(list = setdiff(ls(), parameter_list))  # Clear environment

  print("Script 6: Starting to process BCHC UCR data")
  source(file.path(directory_path, "R scripts/06_Extract_BCHC_UCR.R"))
  print("Finished processing BCHC UCR data.")
  rm(list = setdiff(ls(), parameter_list))  # Clear environment

  # print("Script 7: Starting to process UHB Discharges data")
  # source(file.path(directory_path, "R scripts/07_Extract_UHB_Discharges.R"))
  # print("Finished processing UHB Discharges data.")
  #
  # print("Script 8: Starting to process Virtual Wards data")
  # source(file.path(directory_path, "R scripts/08_Extract_Virtual_Wards.R"))
  # print("Finished processing Virtual Wards data.")
  #
  # print("Script 9: Starting to process WMAS data")
  # source(file.path(directory_path, "R scripts/09_Extract_WMAS.R"))
  # print("Finished processing VMAS data.")

  print("Script 10: Starting the data loading process")
  source(file.path(directory_path, "R scripts/10_Load_Data_Into_DB.R"))
  print("Finished processing all UEC data.")

  end_time <- Sys.time()
  time_taken <- end_time - start_time

  print(paste0("Total time taken to run all scripts: ", time_taken))
}

# Run the master script function
run_master_script()

unique(DQ_Check$Sitrep_File)
