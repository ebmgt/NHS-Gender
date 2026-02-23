# This file is available at https://github.com/ebmgt/NHS-Gender/
# Author:  rbadgett@kumc.edu
# Permission: GNU GPLv3 https://choosealicense.com/licenses/gpl-3.0/
# Last edited 2026-02-03

### Start =======================================
library(tcltk) # For interactions and troubleshooting

#* Cleanup ======
#* # REMOVE ALL:
rm(list = ls())

##* Functions -----

`%notin%` <- Negate(`%in%`)
`%!=na%` <- function(e1, e2) (e1 != e2 | (is.na(e1) & !is.na(e2)) | (is.na(e2) & !is.na(e1))) & !(is.na(e1) & is.na(e2))
`%==na%` <- function(e1, e2) (e1 == e2 | (is.na(e1) & is.na(e2)))

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
    cat('Installing package: ', package_name)
  }
  #tk_messageBox(type = "ok", paste(packages, collapse="\n"), title="Packages installed")
}

current.date <- function(){
  return (as.character(strftime (Sys.time(), format="%Y-%m-%d", tz="", usetz=FALSE)))
}  

# Packages install ------
function_progress(0,'Libraries')

packages_essential <- c("tcltk",'stringr','openxlsx','openxlsx','readr')
function_libraries_install(packages_essential)
function_progress(50,'Libraries')

packages_other <- c("dplyr",'readxl','readr')
function_libraries_install(packages_other)
function_progress(100,'Libraries')

# If Rstudio
if (Sys.getenv("RSTUDIO") != "1"){
  args <- commandArgs(trailingOnly = FALSE)
  script_path <- sub("--file=", "", args[grep("--file=", args)])  
  script_path <- dirname(script_path)
  setwd(script_path)
}else{
  setwd(dirname(rstudioapi::getSourceEditorContext()$path))
  #ScriptsDir <- paste(getwd(),'/Scripts',sep='')
}
getwd()

##Parameters  -----
#* Breakdown group  -----
Breakdown_included <- tk_select.list(c('Religion','Ethnic background (summary)',
                                       'Age','Gender',
                                       'Occupation group (summary)',
                                       'Patient facing role',
                                       'organisational results',
                                       'All'), 
                                     preselect = 'organisational results', multiple = FALSE,
                                     title =paste0("\n\n",Sys.info()["user"], ":\n\nWhat breakdown group are we studying?\n\n"))

subdirectory_path <- "../data"

# ______________________________________-----
# File list with SELECTED ROWs  ------
files_processed <- list.files(path = subdirectory_path, pattern = "detailed spreadsheets organisational results \\(selected columns\\).*\\.xlsx$", full.names = TRUE)

files_processed <- files_processed[!grepl("~", basename(files_processed))]

files_processed <- files_processed[grepl(substr(Breakdown_included, start = 1, stop = 4), basename(files_processed))]
print(files_processed)

if(length(files_processed)==0){
  tk_messageBox(type = c('ok'), paste0("\n\n", Sys.info()["user"], ":\n\nNo files found for ",Breakdown_included), caption = "Oops!")
  stop()
}
# ______________________________________-----
# File list of original NHS fies  ------
files_original <- list.files(path = subdirectory_path, pattern = "\\.xlsx$", full.names = TRUE)

# Exclude files containing specific terms like "selected columns" or "header"
files_original <- files_original[
  !grepl("selected columns|header", files_original, ignore.case = TRUE)
]

# Exclude the specific "data_trusts_well-being" file
files_original <- files_original[
  !grepl("data_trusts_well-being", files_original, ignore.case = TRUE)
]

# Apply the Breakdown_included filter
library(readxl)
library(dplyr) # Required for bind_rows() and case_when later

# NOTE: Ensure 'subdirectory_path' and 'Breakdown_included' are defined in your environment

# --- 1. Filter the list of original files ------
cat("--- Filtering Files ---\n")
files_original <- list.files(path = subdirectory_path, pattern = "\\.xlsx$", full.names = TRUE)

# Exclude files containing specific terms
files_original <- files_original[
  !grepl("selected columns|header|data_trusts_well-being|~", files_original, ignore.case = TRUE)
]

# Apply the Breakdown_included filter
files_original <- files_original[grepl(substr(Breakdown_included, start = 1, stop = 4), basename(files_original))]

print(files_original)


# --- 2. Load all filtered files into a single master dataframe ------
cat("\n--- Loading and Combining Data ---\n")
df_combined_list <- lapply(files_original, function(file_path) {
  # Read the excel file from sheet 2, skipping the first 4 rows to reach row 5
  # suppressMessages() keeps the console clean from "New names" warnings
  df <- suppressMessages(read_excel(file_path, sheet = 2, skip = 4))
  return(df)
})

# Combine list into one large dataframe
df_master <- bind_rows(df_combined_list)
cat("Data combined into 'df_master' dataframe (Total rows:", nrow(df_master), ")\n")


# --- 3. Calculate and print metrics per file (Counts Matrix) ------
cat("\n--- Results Per File ---\n")
counts_matrix <- sapply(files_original, function(file_path) {
  df <- suppressMessages(read_excel(file_path, sheet = 2, skip = 4)) 
  
  count_weighted <- sum(df$Weighting == "Occupation Group", na.rm = TRUE)
  total_rows <- nrow(df)
  unique_orgs <- length(unique(df$`Organisation name`))
  
  return(c("Count_Occupation_Group" = count_weighted, 
           "Total_Rows" = total_rows,
           "N_Unique_Orgs_In_File" = unique_orgs))
})
print(counts_matrix)


# --- 4. Calculate and print total combined summary metrics ------

# Grand totals of counts (summing rows across all files)
grand_totals <- rowSums(counts_matrix[c("Count_Occupation_Group", "Total_Rows"), ])

cat("\n--- Grand Total Summary (Counts) ---\n")
print(grand_totals)


# ---- Unique organizations ACROSS ALL FILES (de-duplicated) ----

# Defensive: remove NA values first
all_org_names <- df_master$`Organisation name`
all_org_names <- all_org_names[!is.na(all_org_names)]

total_unique_orgs_across_all_files <- length(unique(all_org_names))
orgs_occ_group <- df_master$`Organisation name`[
  df_master$Weighting == "Occupation Group"
]
orgs_occ_group <- orgs_occ_group[!is.na(orgs_occ_group)]
total_unique_orgs_occ_group <- length(unique(orgs_occ_group))

cat("\n--- Grand Total Summary (Unique Organizations with Occupation Group Weighting) ---\n")
cat("N unique organisations (Occupation Group only):",
    total_unique_orgs_occ_group, "\n")


cat("\n--- Grand Total Summary (Unique Organizations Across ALL Files) ---\n")
cat("N unique organisations:", total_unique_orgs_across_all_files, "\n")
