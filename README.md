# UEC Performance ETL Process

## Overview
This repository contains R scripts designed to extract, transform, and load (ETL) data related to daily Urgent and Emergency Care (UEC) performance. The scripts systematically process datasets from multiple sources, including hospital discharge reports, virtual ward data, and site reports, to generate a unified dataset for analysis.

## Workflow
The ETL process follows a structured pipeline:

1. **Data Extraction**  
   - Extracts data from various sources such as BCHC, UHB, WMAS, and virtual wards.  
   - Handles multiple structured files for daily and weekly performance metrics.  

2. **Data Transformation**  
   - Cleanses and standardizes extracted data.  
   - Merges relevant datasets to create a comprehensive UEC performance dataset.  
   - Applies transformations necessary for analytics and reporting.  

3. **Data Loading**  
   - Loads transformed data into a database or reporting system.  
   - Ensures data integrity and logs execution status.  

## Script Descriptions

| Script Name                      | Description |
|----------------------------------|-------------|
| **01_Download_SPT_files.R**      | Downloads required data files from external sources. |
| **02_Extract_ADA.R**             | Extracts data related to Admissions, Discharges, and Activity. |
| **03_Extract_BWCH_Weekly_Sitrep.R** | Processes weekly situation reports for Birmingham Women's and Children's Hospital. |
| **04_Extract_MH_Daily_Sitrep.R** | Extracts daily Mental Health site report data. |
| **05_Extract_UHB_Daily_Sitrep.R** | Extracts daily site report data from UHB Trust. |
| **06_Extract_BCHC_UCR.R**        | Extracts data from BCHC Urgent Community Response. |
| **07_Extract_UHB_Discharges.R**  | Extracts daily hospital discharge data from UHB. |
| **08_Extract_Virtual_Wards.R**   | Extracts virtual ward activity data. |
| **09_Extract_WMAS.R**            | Extracts West Midlands Ambulance Service (WMAS) data. |
| **10_Load_Data_Into_DB.R**       | Loads the final transformed dataset into the database. |
| **RUN_MASTER_SCRIPT.R**          | Executes the full ETL pipeline in sequential order. |

## Execution Instructions
1. Ensure all required dependencies (R packages) are installed.
2. Modify any file paths or database connection settings in the scripts if necessary.
3. Run **RUN_MASTER_SCRIPT.R** to execute the full ETL process.

## Requirements
- R (Latest Version)
- Required Libraries: `tidyverse`, `readr`, `DBI`, `RSQLite`, etc.
- Database credentials (if applicable)

## Output
The final processed dataset will be stored in a database or exported as a structured file for reporting and analysis.

## Troubleshooting
- Check logs for errors in data extraction or transformation.
- Ensure input files are available and formatted correctly.
- Verify database connectivity for data loading.
