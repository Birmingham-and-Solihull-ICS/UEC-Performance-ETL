print("1. Establishing connection to the database...")

connection_bsol <- dbConnect(
  odbc(),
  driver="SQL Server",
  server="MLCSU-BI-SQL",
  database="EAT_Reporting_BSOL",
  trusted_Connection="TRUE"
)

# BWC Weekly Sitrep Report ---------------------------------------------------------

print("2. Reading processed Excel files...")

BWC_Weekly_Sitrep_Data <- read_xlsx(
  paste0(output_file_path, "/Processed_BWC_Sitrep.xlsx")) %>% 
  mutate(Load_Date = Sys.Date())


# UHB Daily Sitrep Report ----------------------------------------------------------

UHB_Daily_Sitrep_Data <- read_xlsx(
  paste0(output_file_path, "/Processed_UHB_Sitrep.xlsx"),
  "All") %>% 
  mutate(Load_Date = Sys.Date())


# Virtual Wards --------------------------------------------------------------------

Virtual_Wards_Data <- read_xlsx(
  paste0(output_file_path, "/Processed_Virtual_Wards.xlsx")) %>% 
  mutate(Load_Date = Sys.Date())

# BSMHFT --------------------------------------------------------------------

BSMHFT_Data <- read_xlsx(
 paste0(output_file_path, "/Processed_BSMHFT_Sitrep.xlsx"),
  col_types = c("text", "text", "text", "text", "text",
                "date", "text", "numeric", "text")) %>% 
  mutate(Load_Date = Sys.Date())


# WMAS --------------------------------------------------------------------

WMAS_Data <- read_xlsx(
  paste0(output_file_path, "/Processed_WMAS.xlsx")) %>% 
  mutate(Load_Date = Sys.Date())

# UHB Discharges --------------------------------------------------------------------

UHB_Discharges_Data <- read_xlsx(
  paste0(output_file_path, "/Processed_UHB_Discharges.xlsx")) %>% 
  mutate(Load_Date = Sys.Date())

# UCR BCHC --------------------------------------------------------------------

BCHC_UCR_Data <- read_xlsx(
  paste0(output_file_path, "/Processed_BCHC_UCR.xlsx")) %>% 
  mutate(Load_Date = Sys.Date())

# ADA --------------------------------------------------------------------

ADA_Data <- read_xlsx(
  paste0(output_file_path, "/Processed_ADA.xlsx")) %>% 
  mutate(Load_Date = Sys.Date())


# Load Holding tables into database---------------------------------------------------------------

print("3. Loading the processed data into respective holding tables in the database...")

Data_Frames = list(BWC_Weekly_Sitrep_Data,UHB_Daily_Sitrep_Data,Virtual_Wards_Data,BSMHFT_Data
                   ,WMAS_Data,UHB_Discharges_Data
                   ,BCHC_UCR_Data
                   ,ADA_Data)

Table_Names = c("BSOL_1339_BWC_Sitrep_Data_Holding","BSOL_1339_UHB_Sitrep_Data_Holding","BSOL_1339_Virtual_Wards_Holding","BSOL_1339_BSMHFT_Sitrep_Data_Holding"
                ,"BSOL_1339_WMAS_Data_Holding","BSOL_1339_UHB_Discharges_Data_Holding","BSOL_1339_BCHC_UCR_Data_Holding"
                ,"BSOL_1339_ADA_Holding")

for (i in seq_along(Table_Names)) {
  dbExecute(connection_bsol,
            paste0("DROP TABLE IF EXISTS ",
                   Table_Names[i]))
  dbWriteTable(connection_bsol,
               Id(schema = "Development",
                  table = Table_Names[i]),
               value = Data_Frames[[i]],
               overwrite = TRUE)
  
}

# Insert into static tables and deduplicate-------------------------------------------------------

print("4. Inserting data into static tables and deduplicating data...")

dbExecute(connection_bsol,
          "EXEC [Development].[USP_BSOL_UEC_Performance_Sitrep_Load_Process] ")


DQ_Check <- dbGetQuery(
  connection_bsol,
  "SELECT *
     FROM ##Performance_DQ_Records
    ORDER BY 1,2"
) %>% as_tibble()

print("5. Process completed.")