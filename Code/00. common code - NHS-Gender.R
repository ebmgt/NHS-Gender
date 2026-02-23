# This file is available at https://github.com/ebmgt/NHS-Gender/
# Author: rbadgett@kumc.edu
# Permission: GNU GPLv3 https://choosealicense.com/licenses/gpl-3.0/
# Last edited 2026-02-22

# Start =======================================
library(tcltk) # For interactions and troubleshooting, Oart of base so no install needed.
library(crayon)
# Create a hidden root window, then always use
# patent = root,
root <- tktoplevel()
tkwm.withdraw(root)
# Force root to stay on top of all other windows
tcl("wm", "attributes", root, "-topmost", TRUE)

## Set working directory -----
# Get the directory path of the currently executing script
stub <- function() {}
thisPath <- function() {
  cmdArgs <- commandArgs(trailingOnly = FALSE)
  if (length(grep("^-f$", cmdArgs)) > 0) {
    # R console option
    normalizePath(dirname(cmdArgs[grep("^-f", cmdArgs) + 1]))[1]
  } else if (length(grep("^--file=", cmdArgs)) > 0) {
    # Rscript/R console option
    scriptPath <- normalizePath(dirname(sub("^--file=", "", cmdArgs[grep("^--file=", cmdArgs)])))[1]
  } else if (Sys.getenv("RSTUDIO") == "1") {
    # RStudio
    dirname(rstudioapi::getSourceEditorContext()$path)
  } else if (is.null(attr(stub, "srcref")) == FALSE) {
    # 'source'd via R console
    dirname(normalizePath(attr(attr(stub, "srcref"), "srcfile")$filename))
  } else {
    stop("Cannot find file path")
  }
}
thisPath()
setwd <- thisPath()
getwd()


## Set the custom error handler ------
##* Define the custom error handler
function_custom_error_handler <- function() {
  # Get the error message
  last_error <- geterrmessage()
  
  # Print a clean, visible header in the console
  cat("\n--- 🛑 CUSTOM ERROR CAUGHT ---\n")
  cat("MESSAGE: ", last_error, "\n")
  cat("------------------------------\n")
  
  # Show the traceback only if it's not empty
  tb <- capture.output(traceback(3)) # '3' skips the handler itself
  if(length(tb) > 0) {
    cat("CALL STACK:\n")
    cat(paste(tb, collapse = "\n"), "\n")
  }
  cat("------------------------------\n")
}

# Apply the safer option
options(error = function_custom_error_handler)

## Constants ------
Pb <- NULL  # For function_progress
Pallette_RoyGBiv <- c('red','orange','yellow', 'green', 'blue', '#4B0082', 'violet')
Palette_KU <- c("KUBlue" = "#0022B4", "KUSkyBlue" = "#6DC6E7", "KUCrimson" = "#e8000d", "KUYellow" = "#ffc82d")

# ____________________________________________----
# Functions (essential) -----
function_libraries_install <- function(packages, compile = FALSE){
  if (compile) {
    install.packages(setdiff(packages, rownames(installed.packages())), repos = "https://cloud.r-project.org/")
  } else {
    install.packages(setdiff(packages, rownames(installed.packages())), repos = "https://cloud.r-project.org/", INSTALL_opts = c("--no-multiarch", "--no-test-load"))
  }
  for(package_name in packages)
  {
    library(package_name,character.only=TRUE, quietly = FALSE);
    pkg_desc <- packageDescription(package_name)
    cat("Installed:", package_name, " (Version ", pkg_desc$Version, ")\n")
  }
}

Message.summary <- NULL
`%notin%` <- Negate(`%in%`)
LF <- "\n"

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


## Tables, 2xX ------
# call example: function_table_nice2var(data_aamc_slice, 'RESPONDENT_STATUS','age_under44')
function_table_nice <- function(dataframe, row_name = 'RESPONDENT_STATUS', column_name) {
  # Make the table, either one or two var. For one, leave parameter 2 empty
  contingency_table <- base::table(
    dataframe[,row_name], dataframe[,column_name], 
    dnn =  c(row_name, column_name), useNA = 'always')
  
  print(addmargins(contingency_table))
  
  # Print the column percentages
  print (round(prop.table(contingency_table, 2) * 100,1))

    # Perform chi-square test and display results
  chi_sq_test <- chisq.test(contingency_table)
  print(paste("chi-square: ", chi_sq_test$statistic, " p-value: ", chi_sq_test$p.value, sep=' '))
  
  # Suppress warnings
  suppressWarnings(paste("(warnings suppressed)"))
}

function_table_nice1var <- function(dataframe, column_name) {
  # Create the contingency table
age_table <- table(dataframe[,column_name], useNA = 'always')

# Calculate the column percentages
column_percentages <- prop.table(age_table) * 100

# Combine the counts and percentages into a data frame
combined_df <- as.data.frame(rbind(Counts = as.numeric(age_table), Percentages = round(column_percentages, 2)))

# Add proper row and column names
rownames(combined_df) <- c("Counts", "Percentages")
colnames(combined_df) <- names(age_table)

# Print the combined data frame
print(combined_df)
}

## Display content ------
function_display_df_in_viewer <- function(df,
                                          caption       = NULL,   # bold title above table
                                          footer        = NULL,   # left-aligned note below
                                          highlight_row = NULL) { # 1-based index
  
  ## ---- 1. Caption (bold, left-aligned)
  caption_html <- if (!is.null(caption) && nzchar(caption)) {
    htmltools::HTML(
      sprintf("<div style='text-align:left;'>%s</div>", caption)
    )
  } else {
    NULL
  }
  
  ## ---- 2. Optional row-highlight JavaScript
  rowCallback <- NULL
  if (!is.null(highlight_row)) {
    rowCallback <- DT::JS(sprintf(
      "function(row, data, displayNum, displayIndex, dataIndex){
         if (displayIndex === %d){
           $(row).css('background-color','chartreuse');
         }
       }",
      highlight_row - 1          # zero-based for JS
    ))
  }
  
  ## ---- 3. DT options
  dt_opts <- list(pageLength = 25)
  if (!is.null(rowCallback)) dt_opts$rowCallback <- rowCallback
  
  ## ---- 4. Build the DT widget
  dt_widget <- DT::datatable(
    df,
    caption = caption_html,
    options = dt_opts,
    escape  = FALSE
  )
  
  ## ---- 5. Footer block
  # left line: user-supplied footer (if any)
  left_line  <- if (!is.null(footer) && nzchar(footer)) {
    htmltools::HTML(sprintf("<div style='text-align:left;'>%s</div>", footer))
  } else NULL
  
  # right line: always e-mail + date
  right_line <- htmltools::HTML(sprintf(
    "<div style='text-align:right;'>rbadgett@kumc.edu, %s</div>", Sys.Date()
  ))
  
  # combine safely
  footer_html <- htmltools::tagList(left_line, right_line)
  
  
  ## ---- 6. Append footer and return widget
  htmlwidgets::appendContent(dt_widget, footer_html)
}



## Statistics -----
function_lm_r2 <- # for dominance analysis
  # https://cran.r-project.org/web/packages/domir/
  # call with:
  # dominance <- domir(as.formula(formula), lm_r2, data = temp)
  # OR 
  # dominance <- domir(outcome  ~ predictor1 + predictor2, lm_r2, data = temp)
  function(formula, data) { 
    lm_res <- lm(formula, data = data) # data_ammc_byRespondent
    summary(lm_res)[["r.squared"]]
  }
## Other -----
pretty_label_NHS <- function(x) {
  # 1. Replace underscores with spaces
  x <- gsub("_", " ", x)
  # 2. Replacements for labels
  # Using ignore.case = TRUE catches "emotional", "Emotional", etc.
  x <- gsub("meet", "met", x, ignore.case = TRUE)
  x <- gsub("Emotional", "Harassment by", x, ignore.case = TRUE)
  x <- gsub("Physical", "Physical violence by", x, ignore.case = TRUE)
  x <- gsub("Discrimination", "Discrimination by", x, ignore.case = TRUE)
  x <- gsub("Older", "Older (age over 50)", x, ignore.case = TRUE)
  
  return(x)
}

function_coef_style <- function(est, color = TRUE) {
  # green+bold if positive, red+bold if negative, bold always.
  out_col <- "black"
  out_font <- 2
  # Return default immediately if NA
  if (is.na(est)) {
    return(list(col = out_col, font = out_font))
  }
  # Apply conditional color logic only if color == TRUE
  if (color) {
    if (est > 0) {
      out_col <- "darkgreen"
    } else if (est < 0) {
      out_col <- "red3"
    }
  }
  return(list(col = out_col, font = out_font))
}


function_fmt_p <- function(p) {
  if (is.na(p)) return("NA")
  if (p < 1e-4) return(format(p, scientific = TRUE, digits = 2))
  format(p, digits = 3)
}

# ____________________________________________----
# Packages, libraries -----
function_progress(0,'Libraries')

packages_essential <- c('tcltk',
                        'rstudioapi',
                        'stringr',
                        'dplyr'
                        )
function_libraries_install(packages_essential)

packages_files <- c('openxlsx',"openxlsx2"
)
function_libraries_install(packages_files)

function_progress(25,'Libraries')

function_progress(100,'Libraries')

##* Graphics --------------------------
#windows(600, 600, pointsize = 12) # Size 600x600
devAskNewPage(ask = FALSE)
par(mar=c(5.1,4.1,4.1,2.1), mfrow=c(1,1))
old.par <- par(no.readonly=T)
plot.new()
xmin <- par("usr")[1] + strwidth("A")
xmax <- par("usr")[2] - strwidth("A")
ymin <- par("usr")[3] + 1.2*strheight("A")
ymax <- par("usr")[4] - strheight("A")
# grid.raster(mypng, .3, .3, width=.25)

##* Formatted text --------------------
# options(scipen = 999)
function_fmt_coef <- function(x) {
  ifelse(is.na(x), "",
         formatC(x, format = "f", digits = 3))
}

function_fmt_p <- function(p) {
  ifelse(is.na(p), "",
         ifelse(p < 0.001, "<0.001",
                formatC(p, format = "f", digits = 3)))
}

function_pretty_label <- function(x) 
{gsub("_", " ", x)}
