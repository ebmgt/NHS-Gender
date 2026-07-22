# This file is available at the Open Science Foundation via https://osf.io/3qu9v/
# Author:rbadgett@kumc.edu
# Permissions: -----
#* Images/text CC BY-NC-SA 4.0 https://creativecommons.org/licenses/by-nc-sa/4.0/
#** Plain English details https://creativecommons.org/licenses/by-nc-sa/4.0/
#* Code GNU GPLv3 https://www.gnu.org/licenses/gpl-3.0.en.html
#** Plain English details at https://www.gnu.org/licenses/quick-guide-gplv3.html
#* Commercial licensing options are available separately.
# Optimized for coding with R Studio document outline view

## Include common code -----
source("00. common code - NHS-Gender - 2026-04-13.R")

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
getwd()

# randomForest -----
library(randomForest)
df_organizations_weighted$Benchmarking.group <- factor(
  df_organizations_weighted$Benchmarking.group
)

# _____________________________-----
# Data grab -----

#wb.temp <- loadWorkbook("../Data/regdat - 2025-06-16.xlsx")
wb.temp <- loadWorkbook("../Data/df_organizations_weighted - weighted-2026-04-30.xlsx")
df_organizations_weighted <- read.xlsx (wb.temp, sheet = 1, startRow = 1, colNames = TRUE, na.strings = "NA", detectDates = TRUE)
df_organizations_weighted$Weights_combined <- NULL
vars_x <- function_create_vars_x(df_organizations_weighted)
global_vars_x <- vars_x

# 1) weighting -----

# Check both the column existence and the value of global_var_weight_rf
if (!"Weights_combined" %in% names(df_organizations_weighted) && global_var_weight_rf == "Weights_combined") {
  message("Condition met: 'Weights_combined' is required but missing. Running calculation...")
  df_organizations_weighted$Weights_combined <- function_df_organizations_create_weighted_COMBINED(df_organizations_weighted)
} else {
  message("Condition not met or column already exists. No action taken.")
}
summary(df_organizations_weighted$Weights_combined)

vars_x_with_benchmarking_group <- unique(c(global_vars_x, "Benchmarking.group"))

#* rF: Weights.combined, & Benchmarking.group as an unordered facotor -----
# categorical factor predictor. Benchmarking.group is available to the trees
# as a splitting variable, so trust-type differences can be modeled directly.
# The combined weight also affects row sampling into trees and includes both
# Burned_out_base and benchmarking-group-frequency components.

df_organizations_weighted$Benchmarking.group <- factor(
  df_organizations_weighted$Benchmarking.group
)

vars_x_with_benchmarking_group <- unique(c(global_vars_x, "Benchmarking.group"))

model_alldata_random_forest_fit_combined_weights_with_group <- randomForest(
  x = df_organizations_weighted[, vars_x_with_benchmarking_group, drop = FALSE],
  y = df_organizations_weighted$Burned_out_rate,
  weights = df_organizations_weighted$Weights_combined,
  ntree = 1000,
  mtry = max(floor(length(vars_x_with_benchmarking_group) / 3), 1),
  importance = TRUE,
  keep.forest = TRUE
)

#* rF: Weights.combined but no Benchmarking.group as a factor -----
# as a factor predictor. Benchmarking.group still contributes indirectly to
# row sampling through the group-frequency component of Weights_combined, but
# it is not available to the trees as a splitting variable. Any trust-type
# structure must therefore be captured indirectly through correlated predictors.

model_alldata_random_forest_fit_combined_weights <- randomForest(
  x = df_organizations_weighted[, global_vars_x, drop = FALSE],
  y = df_organizations_weighted$Burned_out_rate,
  weights = df_organizations_weighted$Weights_combined,
  ntree = 1000,
  mtry = max(floor(length(global_vars_x) / 3), 1),
  importance = TRUE,
  keep.forest = TRUE
)

model_alldata_random_forest_fit_combined_weights

# _____________________________-----
# Importance of cofactors -----
importance(
  model_alldata_random_forest_fit_combined_weights_with_group
)[
  order(
    importance(model_alldata_random_forest_fit_combined_weights_with_group)[, "%IncMSE"],
    decreasing = TRUE
  ),
]
##* Betareg to test if female is sig with cofactors -----

# Full model with Female_rate
m_beta_with_female <- betareg(
  as.formula(
    paste0(
      global_var_outcome,
      " ~ ",
      paste(unique(c(global_var_x_focal, "Benchmarking.group", global_vars_best_by_1se_adjusted)), collapse = " + ")
    )
  ),
  data = df_organizations_weighted,
  weights = df_organizations_weighted$Weights_combined,
  link = "logit"
)

# Reduced model without Female_rate
m_beta_without_female <- betareg(
  as.formula(
    paste0(
      global_var_outcome,
      " ~ ",
      paste(unique(c("Benchmarking.group", global_vars_best_by_1se_adjusted)), collapse = " + ")
    )
  ),
  data = df_organizations_weighted,
  weights = df_organizations_weighted$Weights_combined,
  link = "logit"
)

##* Female_rate Likelihood ratio testing -----

cat(green$bold("Female_rate adjusted LR test"))
cat(green("If sig: Female_rate remains a statistically significant predictor of burnout in the adjusted beta-regression model.\n"))
cat(green("Compared two models, with female rate versus without and asks whether adding Female_rate improves overall model fit."))
data.frame(
  Test = "Female_rate adjusted LR test",
  LR_chisq = 2 * (
    as.numeric(logLik(m_beta_with_female)) -
      as.numeric(logLik(m_beta_without_female))
  ),
  df = attr(logLik(m_beta_with_female), "df") -
    attr(logLik(m_beta_without_female), "df"),
  p_value = pchisq(
    2 * (
      as.numeric(logLik(m_beta_with_female)) -
        as.numeric(logLik(m_beta_without_female))
    ),
    df = attr(logLik(m_beta_with_female), "df") -
      attr(logLik(m_beta_without_female), "df"),
    lower.tail = FALSE
  )
)

cat(green("Wald's z tests whether the estimated coefficient for Female_rate differs from zero within the full model."))
summary(m_beta_with_female)$coefficients$mean["Female_rate", ]
