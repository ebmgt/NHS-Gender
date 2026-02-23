# This file is available at https://github.com/ebmgt/NHS-Gender/
# Author: rbadgett@kumc.edu
# Permission: GNU GPLv3 https://choosealicense.com/licenses/gpl-3.0/
# Last edited 2026-02-02

#== Startup ======
library(tcltk) # For interactions and troubleshooting, part of base package so no install needed.

#* Cleanup ======
#* # REMOVE ALL:
rm(list = ls())

## Global variables -----
Pb <- NULL  # For function_progress

## Set working directory -----
if (Sys.getenv("RSTUDIO") != "1"){
  args <- commandArgs(trailingOnly = FALSE)
  script_path <- sub("--file=", "", args[grep("--file=", args)])  
  script_path <- dirname(script_path)
  setwd(script_path)
}else{
  setwd(dirname(rstudioapi::getSourceEditorContext()$path))
  #ScriptsDir <- paste(getwd(),'/Scripts',sep='')
}

# Functions ------
function_progress <- function(progress, titletext){
  #if  (!exists("titletext") || is.null(titletext)){titletext <- 'testing'}
  #if  (!exists("progress") || is.null(progress)){progress <- 0}
  if (progress == 0){
    Pb <<- tkProgressBar(titletext, "", 0, 100, 0)
  }
  info <- sprintf("%d%% done", round(progress))
  setTkProgressBar(Pb, value = progress, title = paste(titletext,sprintf("(%s)", info)), label = info)
  if (progress == 100){
    close(Pb)
  }
}
function_libraries_install <- function(packages){
  install.packages(setdiff(packages, rownames(installed.packages())), repos = "https://cloud.r-project.org/")
  for(package_name in packages)
  {
    library(package_name, character.only=TRUE, quietly = FALSE);
    cat('Installing package: ', package_name,'
        ')
  }
}

function_plot_print <- function (plotname, plotheight, plotwidth){
  plotname <- gsub(":|\\s|\\n|\\?|\\!|\\'", "", plotname)
  rstudioapi::savePlotAsImage(
    paste(plotname,' -- ',current.date(),'.png',sep=''),
    format = "png", width = plotwidth, height = plotheight)
}


# Function to prepend a single quote if one is found anywhere in the field
prepend_quote_if_exists <- function(x) {
  if (is.character(x) && grepl("'", x)) {
    x <- paste0("'", x)
  }
  return(x)
}

# Function to process column names based on merged cells above
process_column_names <- function(sheet_data) {
  # Initialize the prefix variable
  last_valid_prefix <- NULL
  
  # Loop through each column in row 5
  for (col in 1:ncol(sheet_data)) {
    # Check the first row of the merged cells (row 1)
    cell_value <- sheet_data[1, col]
    
    if (!is.na(cell_value) && startsWith(cell_value, "Q") && grepl("-", cell_value)) {
      # Extract the text before the first '-' and trim it
      last_valid_prefix <- trimws(sub("-.*$", "", cell_value))
    }
    
    # Apply the prefix to the current column if a valid prefix is available
    if (!is.null(last_valid_prefix)) {
      # Prepend the column name with the last valid prefix
      sheet_data[5, col] <- paste(last_valid_prefix, sheet_data[5, col], sep = " ")
    }
  }
  colnames(sheet_data) <- sheet_data[5, ]  # Set the new column names
  sheet_data <- sheet_data[-5, ]  # Remove the row with the headers
  return(sheet_data)
}

# Packages/libraries -----
function_progress(0,'Libraries')

packages_essential <- c('stringr','openxlsx','openxlsx2','readr','dplyr')
function_libraries_install(packages_essential)
function_progress(100,'Libraries')

# ______________________________________-----
# Parameters -----
#* Aggregation level (Breakdown)? -----
grouping <- 'organisational results'

#* Year? -----
Year <- tk_select.list(c(2021,2022,2023, 2024, 2025), 
                       preselect = '2021', multiple = FALSE,
                       title = "\n\nWhat year are we studying?\n\n")
Year_short <- substr(as.character(Year), 3, 4)

# ______________________________________-----
# Data grab -----
#* Define the paths to the files -----
file_name_root <- '../data'

input_file_path <- 
  paste0(file_name_root,"/NSS", Year_short, ' ',"detailed spreadsheets organisational results.xlsx")
output_file_header_path <- 
  paste0(file_name_root,"/NSS", Year_short, " detailed spreadsheets organisational results (header).xlsx")
output_file_path <- 
  paste0(file_name_root,"/NSS", Year_short, " detailed spreadsheets organisational results (selected columns).xlsx")
output_csv_path <- 
  paste0(file_name_root," (selected).csv")

#* Load the workbook -----
wb <- loadWorkbook(input_file_path)

# Check file contents ===================================
#* Get the names of all worksheets -----
sheet_names <- getSheetNames(input_file_path)

# Remove the "Notes" worksheet
sheet_names <- sheet_names[sheet_names != "Notes"]

# ______________________________________-----
# Create a workbook for the header file -----
new_wb_header <- createWorkbook()

# Process each worksheet for the header file
index <- 0
for (sheet_name in sheet_names) {
  
  function_progress(100*(index)/length(sheet_names), paste('Processing ',sheet_name))
  if (index+1 == length(sheet_names)){function_progress(100,'Done')}
  # Read the entire sheet to preserve structure and formatting
  sheet_data <- readWorkbook(input_file_path, sheet = sheet_name, colNames = FALSE, skipEmptyRows = FALSE)
  
  # Process the column names in row 5 based on the cells directly above
  sheet_data <- process_column_names(sheet_data)
  
  # Remove rows 1-4 from the data
  sheet_data <- sheet_data[-(1:4), ]
  
  # Keep only the first 5 rows after the column names
  sheet_data <- sheet_data[1:5, ]
  
  ## Column 1 rename if org type file -----
  if (grouping == "organisation type results"){
    colnames(sheet_data)[colnames(sheet_data) == 'Base'] <- 'Benchmarking_group'
  } 
  
  # Add a new sheet to the header workbook
  addWorksheet(new_wb_header, sheetName = sheet_name)
  
  # Write the modified data back to the header workbook
  writeData(new_wb_header, sheet = sheet_name, x = sheet_data, startRow = 1, colNames = TRUE)
  
  # Reapply any merged cells from the original workbook
  original_sheet <- wb[[sheet_name]]
  merged_cells <- original_sheet$mergedCells
  if (!is.null(merged_cells)) {
    for (merge_range in merged_cells) {
      mergeCells(new_wb_header, sheet_name, merge_range)
    }
  }
  index <- index + 1
}

## Save the workbook header -----
saveWorkbook(new_wb_header, output_file_header_path, overwrite = TRUE)

print(paste("New workbook header saved at:", output_file_header_path))

# Create a workbook for the full output file -----
new_wb_selected <- createWorkbook()

#* For each worksheet before combining -----
df_combined <- NULL
for (sheet_name in sheet_names) {
  sheet_data <- readWorkbook(input_file_path, sheet = sheet_name, colNames = FALSE, skipEmptyRows = FALSE)
  
  # Process the column names
  sheet_data <- process_column_names(sheet_data)
  
  #** Remove top 4 rows -----
  sheet_data <- sheet_data[-(1:4), ]
  
  #** Column 1 rename if org type file -----
  if (grouping == "organisation type results"){
    colnames(sheet_data)[colnames(sheet_data) == 'Base'] <- 'Benchmarking_group'
  }
  
  #** EDIT FIELDS TO KEEP HERE -----  
  if (sheet_name == "YOUR JOB Q1-3") {
    # Find the columns through "Weighting"
    end_col_index <- grep("Weighting", colnames(sheet_data))
    if (length(end_col_index) > 0) {
      selected_columns <- sheet_data[, 1:end_col_index]
    } else {
      selected_columns <- NULL
      message("Weighting column not found.")
    }
    Q1_column <- sheet_data[, grep("^Q1 ", colnames(sheet_data))]
    Q3f_column <- sheet_data[, grep("^Q3f ", colnames(sheet_data))]
    Q3g_column <- sheet_data[, grep("^Q3g ", colnames(sheet_data))]
    Q3i_column <- sheet_data[, grep("^Q3i ", colnames(sheet_data))]
    selected_columns <- cbind(selected_columns, Q1_column, Q3f_column, Q3g_column, Q3i_column)
  } else if (sheet_name == "YOUR JOB Q4-6") {
    # Keep only columns that start with "Q6a"
    selected_columns <- sheet_data[, grep("^Q6a", colnames(sheet_data))]
  } else if (sheet_name == "YOUR TEAM Q7") {
    # Keep only columns that start with "Q7h"
    selected_columns <- sheet_data[, grep("^Q7h", colnames(sheet_data))]
  } else if (sheet_name == "HEALTH WELLBEING SAFETY Q10-11") {
    # Keep only columns that start with "Q11c"
    selected_columns <- sheet_data[, grep("^Q11c", colnames(sheet_data))]
  } else if (sheet_name == "HEALTH WELLBEING SAFETY Q12") {
    # Keep only columns that start with "Q12"
    selected_columns <- sheet_data[, grep("^Q12", colnames(sheet_data))]
  } else if (sheet_name == "HEALTH WELLBEING SAFETY Q13-14") {
    physical_columns <- sheet_data[, grep("^Q13", colnames(sheet_data))]
    # Keep only columns that start with "Q13"
    emotional_columns <- sheet_data[, grep("^Q14", colnames(sheet_data))]
    selected_columns <- cbind(emotional_columns,physical_columns)
  } else if (sheet_name == "HEALTH WELLBEING SAFETY Q15-18") {
    # 2021 Discrimination Q16
    discrimination_public_columns <- sheet_data[, grep("^Q16a", colnames(sheet_data))]
    discrimination_internal_columns <- sheet_data[, grep("^Q16b", colnames(sheet_data))]
    selected_columns <- cbind(discrimination_public_columns, discrimination_internal_columns)
  } else if (sheet_name == "HEALTH WELLBEING SAFETY Q15-17") {
    # 2022-2023 Discrimination Q16
    discrimination_public_columns <- sheet_data[, grep("^Q16a", colnames(sheet_data))]
    discrimination_internal_columns <- sheet_data[, grep("^Q16b", colnames(sheet_data))]
    selected_columns <- cbind(discrimination_public_columns, discrimination_internal_columns)
  } else if (sheet_name == "BACKGROUND INFORMATION Q24-27") {
    # 2021 and before # str_detect(input_file_path, "21")
    # Gender: Keep only columns that start with "Q24a" 
    gender_columns <- sheet_data[, grep("^Q24a", colnames(sheet_data))]
    # Age: Keep only columns that start with "Q24c" 
    age_columns <- sheet_data[, grep("^Q24c", colnames(sheet_data))]
    selected_columns <- cbind(gender_columns, age_columns)
    # Ethnicity: Keep only columns that start with "Q25" 
    ethnicity_columns <- sheet_data[, grep("^Q25", colnames(sheet_data))]
    selected_columns <- cbind(gender_columns, age_columns, ethnicity_columns)
    ### Religion NOW needed as already aggregated by rows in the script
    religion_columns <- sheet_data[, grep("^Q27", colnames(sheet_data))]
    selected_columns <- cbind(gender_columns, age_columns, ethnicity_columns, religion_columns)
    # Sexuality: Keep only columns that start with "Q26" 
    sexuality_columns <- sheet_data[, grep("^Q26", colnames(sheet_data))]
    selected_columns <- cbind(gender_columns, age_columns, ethnicity_columns, religion_columns, sexuality_columns)
  } else if (sheet_name == "BACKGROUND INFORMATION Q26-29") {
    # 2022 # str_detect(input_file_path, "22")
    # Gender: Keep only columns that start with "Q26a"
    gender_columns <- sheet_data[, grep("^Q26a", colnames(sheet_data))]
    # Age: keep only columns that start with "Q26c" Age
    age_columns <- sheet_data[, grep("^Q26c", colnames(sheet_data))]
    selected_columns <- cbind(gender_columns, age_columns)
    # Ethnicity: Keep only columns that start with "Q27" 
    ethnicity_columns <- sheet_data[, grep("^Q27", colnames(sheet_data))]
    selected_columns <- cbind(gender_columns, age_columns, ethnicity_columns)
    ### Religion NOW needed as already aggregated by rows in the script
    religion_columns <- sheet_data[, grep("^Q29", colnames(sheet_data))]
    selected_columns <- cbind(gender_columns, age_columns, ethnicity_columns, religion_columns)
    # Sexuality: Keep only columns that start with "Q28" 
    sexuality_columns <- sheet_data[, grep("^Q28", colnames(sheet_data))]
    selected_columns <- cbind(gender_columns, age_columns, ethnicity_columns, religion_columns, sexuality_columns)
  } else if (sheet_name == "BACKGROUND INFORMATION Q27-30") {
    # 2023 and after
    # Gender: Keep only columns that start with "Q27a"
    gender_columns <- sheet_data[, grep("^Q27a", colnames(sheet_data))]
    # Age: keep only columns that start with "Q27c"
    age_columns <- sheet_data[, grep("^Q27c", colnames(sheet_data))]
    # Ethnicity: Keep only columns that start with "Q28" 
    ethnicity_columns <- sheet_data[, grep("^Q28", colnames(sheet_data))]
    ### Religion NOW needed as already aggregated by rows in the script
    religion_columns <- sheet_data[, grep("^Q30", colnames(sheet_data))]
    selected_columns <- cbind(gender_columns, age_columns, ethnicity_columns, religion_columns)
    # Sexuality: Keep only columns that start with "Q29" 
    sexuality_columns <- sheet_data[, grep("^Q29", colnames(sheet_data))]
    selected_columns <- cbind(gender_columns, age_columns, ethnicity_columns, religion_columns, sexuality_columns)
  } else if (sheet_name == "BACKGROUND INFORMATION Q28-31") {
    # Medical-dental staff 2021
    selected_columns <- sheet_data[, grep("^Q31", colnames(sheet_data))]
  } else if (sheet_name == "BACKGROUND INFORMATION Q30-33") {
    # Medical-dental staff 2022
    selected_columns <- sheet_data[, grep("^Q33", colnames(sheet_data))]
  } else if (sheet_name == "BACKGROUND INFORMATION Q31-35") {
    # Medical-dental staff 2023
    selected_columns <- sheet_data[, grep("^Q35", colnames(sheet_data))]
  } else {selected_columns <- NULL}
  
  # Combine the selected columns from different sheets
  if (!is.null(selected_columns)) {
    if (sheet_name == "YOUR JOB Q1-3") {
      df_combined <- selected_columns
    } else {
      df_combined <- cbind(df_combined, selected_columns)
    }
  }
}

#** Keep only Bases rows desired (Trusts and CCGs) -----
# We included all Trust types and Clinical Commissioning Groups (CCG)
# We excluded Benchmarking group in the religion study: 
# Commissioning Support Units, Social Enterprises, and Community Surgical Services due to their smaller size.
# Organisation Data Service (ODS) code 
# Integrated Care Boards (ICBs) replaced clinical commissioning groups (CCGs) in the NHS in England from 1 July 2022.
# Integrated Care System (ICS)
# Commissioning Support Unit (CSU)
#*** df_combined initiated -----

#*** If org type files (RELIGION STUDY)-----
if (grouping == "organisation type results"){
  df_combined <- df_combined[!df_combined$`Benchmarking_group` %in% 
                               c('Community Surgical Services',
                                 'CSUs',
                                 'Social Enterprises - Community'), ]
  if (Breakdown_included!= 'All'){
    df_combined <- df_combined[df_combined$`Breakdown` %in% 
                                     c(Breakdown_included), ]
  }
}

#*** If org files GENDER STUDY -----
#*** keep OCCUPATIONAL GROUP ONLY -----
if (grouping == "organisational results"){
  if ("Benchmarking_group" %in% colnames(df_combined)) {
    df_combined <- df_combined[!df_combined$`Weighting` %in% 
                               c('Occupation Group'), ]
  }
}

#** Keep only Bases rows desired IF UNIVARIATE ANALYSIS -----

# Convert all columns that start with "Q" to numeric
# Line below revised to skip errors from "-", etc 2024-10-19
df_combined[, grep("^Q", colnames(df_combined))] <- 
  lapply(df_combined[, grep("^Q", colnames(df_combined))], function(x) {
    as.numeric(suppressWarnings(as.numeric(x)))  # Convert non-numeric values to NA
  })

# This is not valid for response rate as no Q had high response rate
df_combined$Respondents <- apply(df_combined[, grep("^Q", colnames(df_combined))], 1, max, na.rm = TRUE)

#* Rename fields to friendly (FIRST, EDIT ABOVE *IF* FIELDS CHANGE/ADDed) -----
# Burnout Q12b
if ("Q12b % Often" %in% colnames(df_combined) && "Q12b % Always" %in% colnames(df_combined)) {
  df_combined$Burned_out_rate <- df_combined$`Q12b % Often` + df_combined$`Q12b % Always`
  df_combined$Burned_out_sometimes_rate <- df_combined$Burned_out_rate  + df_combined$`Q12b % Sometimes`
  df_combined$Burned_out_ever_rate <- 100 - df_combined$`Q12b % Never`
  colnames(df_combined)[colnames(df_combined) == 'Q12b Base (number of responses)'] <- 'Burned_out_base'
  df_combined$Burned_out_response_rate <- df_combined$Burned_out_base/df_combined$Respondents
  df_combined$Burned_out_rate_adjusted <- NULL
} else {
  message("One or both columns 'Q12b % Yes' and 'Q12b % Always' are not found in the data.")
}

# Emotionally_exhausting_rate Q12a
if ("Q12a % Often" %in% colnames(df_combined) && "Q12a % Always" %in% colnames(df_combined)) {
  df_combined$Emotionally_exhausting_rate <- df_combined$`Q12a % Often` + df_combined$`Q12a % Always`
  df_combined$Emotionally_exhausting_sometimes_rate <- df_combined$Emotionally_exhausting_rate  + df_combined$`Q12a % Sometimes`
  df_combined$Emotionally_exhausting_ever_rate <- 100 - df_combined$`Q12a % Never`
  colnames(df_combined)[colnames(df_combined) == 'Q12a Base (number of responses)'] <- 'Emotionally_exhausting_base'
  df_combined$Emotionally_exhausting_response_rate <- df_combined$Emotionally_exhausting_base/df_combined$Respondents
  df_combined$Emotionally_exhausting_rate_adjusted <- NULL
} else {
  message("One or both columns 'Q12a % Yes' and 'Q12a % Always' are not found in the data.")
}

# Frustration_rate Q12c
if ("Q12c % Often" %in% colnames(df_combined) && "Q12c % Always" %in% colnames(df_combined)) {
  df_combined$Frustration_rate <- df_combined$`Q12c % Often` + df_combined$`Q12c % Always`
  df_combined$Frustration_sometimes_rate <- df_combined$Frustration_rate  + df_combined$`Q12c % Sometimes`
  df_combined$Frustration_ever_rate <- 100 - df_combined$`Q12c % Never`
  colnames(df_combined)[colnames(df_combined) == 'Q12c Base (number of responses)'] <- 'Frustration_base'
  df_combined$Frustration_response_rate <- df_combined$Frustration_base/df_combined$Respondents
  df_combined$Emotionally_exhausting_rate_adjusted <- NULL
} else {
  message("One or both columns 'Q12c % Yes' and 'Q12c % Always' are not found in the data.")
}

# Work_stress Q11c
if ("Q11c % Yes" %in% colnames(df_combined) && "Q11c % Yes" %in% colnames(df_combined)) {
  df_combined$Stress_rate <- df_combined$`Q11c % Yes`
  colnames(df_combined)[colnames(df_combined) == 'Q11c Base (number of responses)'] <- 'Stress_base'
} else {
  message("One or both columns 'Q11c % Often' and 'Q11c % Always' are not found in the data.")
}

## Demographics -----
### Age Q24c (2021) OR Q26c (2022) or Q27c (2023 or more)-----
if ("Q24c % 41-50" %in% colnames(df_combined)) {
  df_combined$Older_rate <- df_combined$`Q24c % 51-65` + df_combined$`Q24c % 66+`
  colnames(df_combined)[colnames(df_combined) == 'Q24c Base (number of responses)'] <- 'Older_base'
} else if ("Q26c % 41-50" %in% colnames(df_combined)) {
  df_combined$Older_rate <- df_combined$`Q26c % 51-65` + df_combined$`Q26c % 66+`
  colnames(df_combined)[colnames(df_combined) == 'Q26c Base (number of responses)'] <- 'Older_base'
} else if ("Q27c % 41-50" %in% colnames(df_combined)) {
  df_combined$Older_rate <- df_combined$`Q27c % 51-65` + df_combined$`Q27c % 66+`
  colnames(df_combined)[colnames(df_combined) == 'Q27c Base (number of responses)'] <- 'Older_base'
} else {
  message("\033[31m\033[1mOne or both columns 'Q24c % 51-65' OR 'Q27c % 51-65' are not found in the data.\033[0m")
}
###  Gender Q26 (2021) OR Q26A ( 2022) OR Q27A (2023 or more)
if ("Q24a % Female" %in% colnames(df_combined)) {
  df_combined$Female_rate <- df_combined$`Q24a % Female`
  colnames(df_combined)[colnames(df_combined) == 'Q24a Base (number of responses)'] <- 'Female_base'
}else if ("Q26a % Female" %in% colnames(df_combined)){
  df_combined$Female_rate <- df_combined$`Q26a % Female`
  colnames(df_combined)[colnames(df_combined) == 'Q26a Base (number of responses)'] <- 'Female_base'
}else if ("Q27a % Female" %in% colnames(df_combined)){
  df_combined$Female_rate <- df_combined$`Q27a % Female`
  colnames(df_combined)[colnames(df_combined) == 'Q27a Base (number of responses)'] <- 'Female_base'
} else {
  message("\033[31m\033[1mOne or both columns 'Q24a % Female' OR 'Q27a % Female' are not found in the data.\033[0m")
}
###  Sexuality Q26 (2021) OR Q28 (2022) OR Q29 (2023 or more)
if ("Q26 % Heterosexual or Straight" %in% colnames(df_combined)) {
  df_combined$Sexual_minority_rate <- 100 - df_combined$`Q26 % Heterosexual or Straight`
  colnames(df_combined)[colnames(df_combined) == 'Q26 Base (number of responses)'] <- 'Sexual_minority_base'
}else if ("Q28 % Heterosexual or Straight" %in% colnames(df_combined)){
  #browser()
  df_combined$Sexual_minority_rate <- 100 - df_combined$`Q28 % Heterosexual or Straight` # corrected 2025-01-11
  colnames(df_combined)[colnames(df_combined) == 'Q28 Base (number of responses)'] <- 'Sexual_minority_base'
}else if ("Q29 % Heterosexual or Straight" %in% colnames(df_combined)){
  #browser()
  df_combined$Sexual_minority_rate <- 100 - df_combined$`Q29 % Heterosexual or Straight` # corrected 2025-01-11
  colnames(df_combined)[colnames(df_combined) == 'Q29 Base (number of responses)'] <- 'Sexual_minority_base'
}else {
  message("\033[31m\033[1mThe columns 'Q26' or 'Q28', or 'Q29' and 'Heterosexual' are not found in the data.\033[0m")
}

### Ethnicity Q25 (2021) OR Q27 (2022) or Q28 (2023 or more)-----
"Q25 % White"
if ("Q25 % White - Irish" %in% colnames(df_combined)) {
  selected_columns <- df_combined %>% dplyr::select(matches("Q25 % White"))
  df_combined$Ethnic_minority_rate <- 100 - rowSums(selected_columns, na.rm = TRUE)
  colnames(df_combined)[colnames(df_combined) == 'Q25 Base (number of responses)'] <- 'Ethnic_minority_base'
} else if ("Q27 % White - Irish" %in% colnames(df_combined)) {
  selected_columns <- df_combined %>% dplyr::select(matches("Q27 % White"))
  df_combined$Ethnic_minority_rate <- 100 - rowSums(selected_columns, na.rm = TRUE)
  colnames(df_combined)[colnames(df_combined) == 'Q27 Base (number of responses)'] <- 'Ethnic_minority_base'
} else if ("Q28 % White - Irish" %in% colnames(df_combined)) {
  selected_columns <- df_combined %>% dplyr::select(matches("Q28 % White"))
  df_combined$Ethnic_minority_rate <- 100 - rowSums(selected_columns, na.rm = TRUE)
  colnames(df_combined)[colnames(df_combined) == 'Q28 Base (number of responses)'] <- 'Ethnic_minority_base'
} else {
  message("\033[31m\033[1mOne or both columns 'Q25' OR 'Q27' or 'Q28' are not found in the data.\033[0m")
}

### Religion  Q27 (2021) OR Q29 (2022) or Q30 (2023 or more)-----
"Q27 % No religion"
if ("Q27 % No religion" %in% colnames(df_combined)) {
  # None
  selected_columns <- df_combined %>% dplyr::select(matches("Q27 % No religion"))
  df_combined$Religion_None_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Christian
  selected_columns <- df_combined %>% dplyr::select(matches("Q27 % Christian"))
  df_combined$Religion_Christian_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Buddhist
  selected_columns <- df_combined %>% dplyr::select(matches("Q27 % Buddhist"))
  df_combined$Religion_Buddhist_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Hindu
  selected_columns <- df_combined %>% dplyr::select(matches("Q27 % Hindu"))
  df_combined$Religion_Hindu_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Jewish
  selected_columns <- df_combined %>% dplyr::select(matches("Q27 % Jewish"))
  df_combined$Religion_Jewish_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Muslim
  selected_columns <- df_combined %>% dplyr::select(matches("Q27 % Muslim"))
  df_combined$Religion_Muslim_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Sikh
  selected_columns <- df_combined %>% dplyr::select(matches("Q27 % Sikh"))
  df_combined$Religion_Sikh_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Any other religion
  selected_columns <- df_combined %>% dplyr::select(matches("Q27 % Any other religion"))
  df_combined$Religion_Any_other_rate <- rowSums(selected_columns, na.rm = TRUE)
  # I would prefer not to say
  selected_columns <- df_combined %>% dplyr::select(matches("Q27 % I would prefer not to say"))
  df_combined$Religion_Decline_answer_rate <- rowSums(selected_columns, na.rm = TRUE)
  colnames(df_combined)[colnames(df_combined) == 'Q27 Base (number of responses)'] <- 'Religion_base'
} else if ("Q29 % No religion" %in% colnames(df_combined)) {
  # None
  selected_columns <- df_combined %>% dplyr::select(matches("Q29 % No religion"))
  df_combined$Religion_None_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Christian
  selected_columns <- df_combined %>% dplyr::select(matches("Q29 % Christian"))
  df_combined$Religion_Christian_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Buddhist
  selected_columns <- df_combined %>% dplyr::select(matches("Q29 % Buddhist"))
  df_combined$Religion_Buddhist_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Hindu
  selected_columns <- df_combined %>% dplyr::select(matches("Q29 % Hindu"))
  df_combined$Religion_Hindu_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Jewish
  selected_columns <- df_combined %>% dplyr::select(matches("Q29 % Jewish"))
  df_combined$Religion_Jewish_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Muslim
  selected_columns <- df_combined %>% dplyr::select(matches("Q29 % Muslim"))
  df_combined$Religion_Muslim_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Sikh
  selected_columns <- df_combined %>% dplyr::select(matches("Q29 % Sikh"))
  df_combined$Religion_Sikh_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Any other religion
  selected_columns <- df_combined %>% dplyr::select(matches("Q29 % Any other religion"))
  df_combined$Religion_Any_other_rate <- rowSums(selected_columns, na.rm = TRUE)
  # I would prefer not to say
  selected_columns <- df_combined %>% dplyr::select(matches("Q29 % I would prefer not to say"))
  df_combined$Religion_Decline_answer_rate <- rowSums(selected_columns, na.rm = TRUE)
  colnames(df_combined)[colnames(df_combined) == 'Q29 Base (number of responses)'] <- 'Religion_base'
} else if ("Q30 % No religion" %in% colnames(df_combined)) {
  # None
  selected_columns <- df_combined %>% dplyr::select(matches("Q30 % No religion"))
  df_combined$Religion_None_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Christian
  selected_columns <- df_combined %>% dplyr::select(matches("Q30 % Christian"))
  df_combined$Religion_Christian_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Buddhist
  selected_columns <- df_combined %>% dplyr::select(matches("Q30 % Buddhist"))
  df_combined$Religion_Buddhist_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Hindu
  selected_columns <- df_combined %>% dplyr::select(matches("Q30 % Hindu"))
  df_combined$Religion_Hindu_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Jewish
  selected_columns <- df_combined %>% dplyr::select(matches("Q30 % Jewish"))
  df_combined$Religion_Jewish_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Muslim
  selected_columns <- df_combined %>% dplyr::select(matches("Q30 % Muslim"))
  df_combined$Religion_Muslim_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Sikh
  selected_columns <- df_combined %>% dplyr::select(matches("Q30 % Sikh"))
  df_combined$Religion_Sikh_rate <- rowSums(selected_columns, na.rm = TRUE)
  # Any other religion
  selected_columns <- df_combined %>% dplyr::select(matches("Q30 % Any other religion"))
  df_combined$Religion_Any_other_rate <- rowSums(selected_columns, na.rm = TRUE)
  # I would prefer not to say
  selected_columns <- df_combined %>% dplyr::select(matches("Q30 % I would prefer not to say"))
  df_combined$Religion_Decline_answer_rate <- rowSums(selected_columns, na.rm = TRUE)
  colnames(df_combined)[colnames(df_combined) == 'Q30 Base (number of responses)'] <- 'Religion_base'
} else {
  message("\033[31m\033[1mOne or both columns 'Q27' OR 'Q29' or 'Q30' are not found in the data.\033[0m")
}

### Occupation group ------
# MedDental Q31 (2021) OR Q33 (2022) or Q35 (2023 or more)
if ("Q31 % Medical / Dental - Consultant" %in% colnames(df_combined)) {
  df_combined$Med_Dental_rate <- df_combined$`Q31 % Medical / Dental - Consultant` + df_combined$`Q31 % Medical / Dental - In Training` + df_combined$`Q31 % Medical / Dental - Other`
  colnames(df_combined)[colnames(df_combined) == 'Q31 Base (number of responses)'] <- 'Med_Dental_base'
  df_combined$Med_Dental_rate <- df_combined$`Q31 % Medical / Dental - Consultant` + df_combined$`Q31 % Medical / Dental - In Training` + df_combined$`Q31 % Medical / Dental - Other`
  colnames(df_combined)[colnames(df_combined) == 'Q31 Base (number of responses)'] <- 'Med_Dental_base'
} else if ("Q33 % Medical / Dental - Consultant" %in% colnames(df_combined)) {
  df_combined$Med_Dental_rate <- df_combined$`Q33 % Medical / Dental - Consultant` + df_combined$`Q33 % Medical / Dental - In Training` + df_combined$`Q33 % Medical / Dental - Other`
  colnames(df_combined)[colnames(df_combined) == 'Q33 Base (number of responses)'] <- 'Med_Dental_base'
} else if ("Q35 % Medical / Dental - Consultant" %in% colnames(df_combined)) {
  df_combined$Med_Dental_rate <- df_combined$`Q35 % Medical / Dental - Consultant` + df_combined$`Q35 % Medical / Dental - In Training` + df_combined$`Q35 % Medical / Dental - Other`
  colnames(df_combined)[colnames(df_combined) == 'Q35 Base (number of responses)'] <- 'Med_Dental_base'
} else {
  message("\033[31m\033[1mOne or both columns 'Q31 % Medical / Dental - Consultant' OR 'Q33 % Medical / Dental - Consultant' OR 'Q35 % Medical / Dental - Consultant' are not found in the data.\033[0m")
}
### Patient_facing Q1 -----
if ("Q1 % Yes, frequently" %in% colnames(df_combined) && "Q1 % Yes, occasionally" %in% colnames(df_combined)) {
  df_combined$Patient_facing_rate <- df_combined$`Q1 % Yes, frequently` + df_combined$`Q1 % Yes, occasionally`
  colnames(df_combined)[colnames(df_combined) == 'Q1 Base (number of responses)'] <- 'Patient_facing_base'
} else {
  message("\033[31m\033[1mOne or both columns 'Q1 % Yes, frequently' and 'Q1 % Yes, occasionally' are not found in the data.\033[0m")
}

## Environment -----
### Autonomy_rate -----
#3f: I am able to make improvements happen in my area of work.
if ("Q3f % Agree" %in% colnames(df_combined) && "Q3f % Strongly agree" %in% colnames(df_combined)) {
  df_combined$Autonomy_rate <- df_combined$`Q3f % Agree` + df_combined$`Q3f % Strongly agree`
  colnames(df_combined)[colnames(df_combined) == 'Q3f Base (number of responses)'] <- 'Autonomy_base'
} else {
  message("\033[31m\033[1mOne or both columns 'Q3f % Agree' and 'Q3f % Strongly agree' are not found in the data.\033[0m")
}
### Mastery_meaning_rate -----
#6a: I feel that my role makes a difference to patients / service users.
if ("Q6a % Agree" %in% colnames(df_combined) && "Q6a % Strongly agree" %in% colnames(df_combined)) {
  df_combined$Mastery_meaning_rate <- df_combined$`Q6a % Agree` + df_combined$`Q6a % Strongly agree`
  colnames(df_combined)[colnames(df_combined) == 'Q6a Base (number of responses)'] <- 'Mastery_base'
} else {
  message("\033[31m\033[1mOne or both columns 'Q6a % Agree' and 'Q6a % Strongly agree' are not found in the data.\033[0m")
}
### Membership_rate -----
#7h I feel valued by my team.
if ("Q7h % Agree" %in% colnames(df_combined) && "Q7h % Strongly agree" %in% colnames(df_combined)) {
  df_combined$Membership_rate <- df_combined$`Q7h % Agree` + df_combined$`Q7h % Strongly agree`
  colnames(df_combined)[colnames(df_combined) == 'Q7h Base (number of responses)'] <- 'Membership_base'
} else {
  message("\033[31m\033[1mOne or both columns 'Q7h % Agree' and 'Q7h % Strongly agree' are not found in the data.\033[0m")
}
### Time_demands_meet_rate -----
#3g: I am able to meet all the conflicting demands on my time at work.
if ("Q3g % Agree" %in% colnames(df_combined) && "Q3g % Strongly agree" %in% colnames(df_combined)) {
  df_combined$Time_demands_meet_rate <- df_combined$`Q3g % Agree` + df_combined$`Q3g % Strongly agree`
  colnames(df_combined)[colnames(df_combined) == 'Q3g Base (number of responses)'] <- 'Time_demands_meet_base'
} else {
  message("\033[31m\033[1mOne or both columns 'Q3g % Agree' and 'Q3g % Strongly agree' are not found in the data.\033[0m")
}
### Enough_staff_rate -----
#3i: There are enough staff at this organisation for me to do my job properly.
if ("Q3i % Agree" %in% colnames(df_combined) && "Q3i % Strongly agree" %in% colnames(df_combined)) {
  df_combined$Enough_staff_rate <- df_combined$`Q3i % Agree` + df_combined$`Q3i % Strongly agree`
  colnames(df_combined)[colnames(df_combined) == 'Q3i Base (number of responses)'] <- 'Enough_staff_base'
} else {
  message("\033[31m\033[1mOne or both columns 'Q3i % Agree' and 'Q3i % Strongly agree' are not found in the data.\033[0m")
}
###  Physical ------
# Physical.public Q13a
if ("Q13a % Never" %in% colnames(df_combined) && "Q13a % Never" %in% colnames(df_combined)) {
  df_combined$Physical_public_rate <- 100 - df_combined$`Q13a % Never`
  colnames(df_combined)[colnames(df_combined) == 'Q13a Base (number of responses)'] <- 'Physical_public_base'
} else {
  message("One or both columns 'Q13a % Often' and 'Q13a % Always' are not found in the data.")
}
# Physical.manager Q13b
if ("Q13b % Never" %in% colnames(df_combined) && "Q13b % Never" %in% colnames(df_combined)) {
  df_combined$Physical_manager_rate <- 100 - df_combined$`Q13b % Never`
  colnames(df_combined)[colnames(df_combined) == 'Q13b Base (number of responses)'] <- 'Physical_manager_base'
} else {
  message("One or both columns 'Q13b % Often' and 'Q13b % Always' are not found in the data.")
}
# Physical.colleagues Q13c (100 - never)
if ("Q13c % Never" %in% colnames(df_combined) && "Q13c % Never" %in% colnames(df_combined)) {
  df_combined$Physical_colleagues_rate <- 100 - df_combined$`Q13c % Never`
  colnames(df_combined)[colnames(df_combined) == 'Q13c Base (number of responses)'] <- 'Physical_colleagues_base'
} else {
  message("One or both columns 'Q13c % Often' and 'Q13c % Always' are not found in the data.")
}
###  Emotional -----
# Emotional.public Q14a
if ("Q14a % Never" %in% colnames(df_combined) && "Q14a % Never" %in% colnames(df_combined)) {
  df_combined$Emotional_public_rate <- 100 - df_combined$`Q14a % Never`
  colnames(df_combined)[colnames(df_combined) == 'Q14a Base (number of responses)'] <- 'Emotional_public_base'
} else {
  message("One or both columns 'Q14a % Often' and 'Q14a % Always' are not found in the data.")
}
# Emotional.manager Q14b
if ("Q14b % Never" %in% colnames(df_combined) && "Q14b % Never" %in% colnames(df_combined)) {
  df_combined$Emotional_manager_rate <- 100 - df_combined$`Q14b % Never`
  colnames(df_combined)[colnames(df_combined) == 'Q14b Base (number of responses)'] <- 'Emotional_manager_base'
} else {
  message("One or both columns 'Q14b % Often' and 'Q14b % Always' are not found in the data.")
}
# Emotional.colleagues Q14c (100 - never)
if ("Q14c % Never" %in% colnames(df_combined) && "Q14c % Never" %in% colnames(df_combined)) {
  df_combined$Emotional_colleagues_rate <- 100 - df_combined$`Q14c % Never`
  colnames(df_combined)[colnames(df_combined) == 'Q14c Base (number of responses)'] <- 'Emotional_colleagues_base'
} else {
  message("One or both columns 'Q14c % Often' and 'Q14c % Always' are not found in the data.")
}
###  Discrimination -----
if ("Q16a % Yes" %in% colnames(df_combined) && "Q16a % Yes" %in% colnames(df_combined)) {
  df_combined$Discrimination_public_rate <- df_combined$`Q16a % Yes`
  df_combined$Discrimination_internal_rate <- df_combined$`Q16b % Yes`
  colnames(df_combined)[colnames(df_combined) == 'Q16a Base (number of responses)'] <- 'Discrimination_public_base'
  colnames(df_combined)[colnames(df_combined) == 'Q16b Base (number of responses)'] <- 'Discrimination_internal_base'
} else {
  message("One or both columns 'Q16a % Yes' and 'Q16a % Yes' are not found in the data.")
}

#* Add the combined data to the new workbook -----
addWorksheet(new_wb_selected, sheetName = "Selected Data")
writeData(new_wb_selected, sheet = "Selected Data", x = df_combined, startRow = 1, colNames = TRUE)

##* Styles-----

# Freeze header row (row 1)
freezePane(new_wb_selected, sheet = "Selected Data", firstRow = TRUE)

# Dimensions (include header row)
n_rows <- nrow(df_combined) + 1
n_cols <- ncol(df_combined)

# Turn on filter dropdowns (AutoFilter) on the header row
addFilter(
  wb    = new_wb_selected,
  sheet = "Selected Data",
  rows  = 1,
  cols  = 1:n_cols
)

# Styles
style_header_gray  <- createStyle(fgFill = "#D9D9D9")  # light gray
style_female_green <- createStyle(fgFill = "#D9EAD3")  # light green
style_burn_red     <- createStyle(fgFill = "#F4CCCC")  # light red

# Header row: light gray across all columns (will be overridden for matching columns below)
addStyle(
  new_wb_selected,
  sheet = "Selected Data",
  style = style_header_gray,
  rows = 1,
  cols = 1:n_cols,
  gridExpand = TRUE,
  stack = TRUE
)

# Identify columns by name (case-insensitive)
col_names <- names(df_combined)
female_cols <- grep("female", col_names, ignore.case = TRUE)
burn_cols   <- grep("burn",   col_names, ignore.case = TRUE)

# Apply column shading INCLUDING header row
if (length(female_cols) > 0) {
  addStyle(
    new_wb_selected,
    sheet = "Selected Data",
    style = style_female_green,
    rows = 1:n_rows,
    cols = female_cols,
    gridExpand = TRUE,
    stack = TRUE
  )
}

# Apply burn after female so it overrides on any overlap
if (length(burn_cols) > 0) {
  addStyle(
    new_wb_selected,
    sheet = "Selected Data",
    style = style_burn_red,
    rows = 1:n_rows,
    cols = burn_cols,
    gridExpand = TRUE,
    stack = TRUE
  )
}

# Save the workbook with selected data -----
saveWorkbook(new_wb_selected, output_file_path, overwrite = TRUE)

message(paste("\033[32m\033[1mNew workbook with selected fields saved at:", output_file_path, "\033[0m"))

