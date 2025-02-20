
## Parameters ------------------------------------------------------------------
# The path to the parent work directory
SPT_site_url <- "https://csucloudservices.sharepoint.com/sites/BSOLEmbeddedBI"

## Return the latest files from the SharePoint folder --------------------------
#1. ADA -------------------------------------------------------------------------

print("1. Extracting ADA data...")

## List all files in the SharePoint folder
recent_files <- list_recent_spt_files(X = 14, folder_name = "Daily UEC Performance/Raw data/Recent data/ADA", 
                                      site_url = SPT_site_url)

## Download the latest files from the SharePoint folder

download_spt_files(recent_files, folder_name = "Daily UEC Performance/Raw data/Recent data/ADA", 
                   directory = ADA_input_file_path, site_url = SPT_site_url)

print("Process completed")

#2. BCHC UCR -------------------------------------------------------------------

print("2. Extracting BCHC UCR data...")

## List all files in the SharePoint folder
recent_files <- list_recent_spt_files(X = 14, folder_name = "Daily UEC Performance/Raw data/Recent data/BCHC UCR", 
                                      site_url = SPT_site_url)

## Download the latest files from the SharePoint folder

download_spt_files(recent_files, folder_name = "Daily UEC Performance/Raw data/Recent data/BCHC UCR", 
                   directory = BCHC_UCR_input_file_path, site_url = SPT_site_url)

print("Process completed")

#3. BSMHFT Sitrep --------------------------------------------------------------

print("3. Extracting MH Sitrep data...")

## List all files in the SharePoint folder
recent_files <- list_recent_spt_files(X = 1, folder_name = "Daily UEC Performance/Raw data/Recent data/BSMHFT Sitrep", 
                                      site_url = SPT_site_url)

## Download the latest files from the SharePoint folder

download_spt_files(recent_files, folder_name = "Daily UEC Performance/Raw data/Recent data/BSMHFT Sitrep", 
                   directory = BSMHFT_Sitrep_input_file_path, site_url = SPT_site_url)

print("Process completed")

#4. BWC Sitrep -----------------------------------------------------------------

print("4. Extracting BWC Sitrep data...")

## List all files in the SharePoint folder
recent_files <- list_recent_spt_files(X = 14, folder_name = "Daily UEC Performance/Raw data/Recent data/BWC Sitrep", 
                                      site_url = SPT_site_url)

## Download the latest files from the SharePoint folder

download_spt_files(recent_files, folder_name = "Daily UEC Performance/Raw data/Recent data/BWC Sitrep", 
                   directory = BWC_Sitrep_input_file_path, site_url = SPT_site_url)

print("Process completed")


#5. UHB Sitrep -----------------------------------------------------------------

print("5. Extracting UHB Sitrep data...")

## List all files in the SharePoint folder
recent_files <- list_recent_spt_files(X = 1, folder_name = "Daily UEC Performance/Raw data/Recent data/UHB Sitrep", 
                                      site_url = SPT_site_url)

## Download the latest files from the SharePoint folder

download_spt_files(recent_files, folder_name = "Daily UEC Performance/Raw data/Recent data/UHB Sitrep", 
                   directory = UHB_Sitrep_input_file_path, site_url = SPT_site_url)

print("Process completed")

#6. UHB Discharges -------------------------------------------------------------

# print("6. Extracting UHB Discharges data...")

# ## List all files in the SharePoint folder
# recent_files <- list_recent_spt_files(X = 14, folder_name = "Daily UEC Performance/Raw data/Recent data/UHB Discharges", 
#                                       site_url = SPT_site_url)
# 
# ## Download the latest files from the SharePoint folder
# 
# download_spt_files(recent_files, folder_name = "Daily UEC Performance/Raw data/Recent data/UHB Discharges", 
#                    directory = UHB_discharges_input_file_path, site_url = SPT_site_url)

# print("Process completed")

#7. Virtual Wards --------------------------------------------------------------

# print("7. Extracting Virtual Wards data...")

# ## List all files in the SharePoint folder
# recent_files <- list_recent_spt_files(X = 14, folder_name = "Daily UEC Performance/Raw data/Recent data/Virtual Wards", 
#                                       site_url = SPT_site_url)
# 
# ## Download the latest files from the SharePoint folder
# 
# download_spt_files(recent_files, folder_name = "Daily UEC Performance/Raw data/Recent data/Virtual Wards", 
#                    directory = VW_input_file_path, site_url = SPT_site_url)

# print("Process completed")

#8. WMAS -----------------------------------------------------------------------

# print("8. Extracting WMAS data...")

# ## List all files in the SharePoint folder
# recent_files <- list_recent_spt_files(X = 14, folder_name = "Daily UEC Performance/Raw data/Recent data/WMAS", 
#                                       site_url = SPT_site_url)
# 
# ## Download the latest files from the SharePoint folder
# 
# download_spt_files(recent_files, folder_name = "Daily UEC Performance/Raw data/Recent data/WMAS", 
#                    directory = WMAS_input_file_path, site_url = SPT_site_url)

# print("Process completed")

