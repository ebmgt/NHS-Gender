# This file is available at https://github.com/ebmgt/NHS-Gender/
# Authors: Robert Badgett. rbadgett@kumc.edu
# Permission: GNU GPLv3 https://choosealicense.com/licenses/gpl-3.0/
# Last edited 2026-01-23

# https://nhssurveys.co.uk/nss/questions/national
# https://www.nhsemployers.org/news/nhs-staff-survey-2023-key-findings

# https://nhssurveys.co.uk/nss/response_rate/national
# Exclude MH & LD, and MH, LD & Community

years <- c(2021,2022,2023,2024)

# ___________________________________________-----
# Respondents, overall -----
# Response rate to survey starting 2021:
# https://nhssurveys.co.uk/nss/response_rate/national
respondents <- c(628475, 679088, 678676, 747288)
response.rates <- c(50.14, 47.67, 49.73, 51.65)

# 1. Back-calculate recipients per year to use as weights
recipients_per_year <- (respondents / (response.rates / 100))
# 2. Calculate the weighted average of response rates
# Each year's rate is weighted by that year's number of recipients
response.rate.overall <- weighted.mean(response.rates, w = recipients_per_year)

message(paste("\033[32m\033[1mRespondents overall: ", 
    formatC(sum(respondents), format = "f", big.mark = ",", digits = 0),
    "\033[0m"))
message(paste("\033[32m\033[1mWeighted response rate over study years: ", 
              formatC(response.rate.overall, format = "f", big.mark = ",", digits = 1), "%\033[0m"))

# ___________________________________________-----
# Responses to  burnout -----
# https://nhssurveys.co.uk/nss/questions/national

# below is copied from embedded spreadsheet at https://nhssurveys.co.uk/nss/detailed_questions/national
# When Question is Q12b - How often, if at all, do you feel burnt out because of your work?

txt <- "
National Average Always - 43,936 44,611 41,957 47,190
National Average Often - 166,191 160,462 161,349 175,934
National Average Sometimes - 233,113 229,936 260,868 289,296
National Average Rarely - 129,913 130,688 156,963 172,050
National Average Never - 42,145 43,548 54,657 58,780
"

lines <- trimws(unlist(strsplit(txt, "\n")))
lines <- lines[nzchar(lines)]

pattern <- "^(.+?)\\s+(Always|Often|Sometimes|Rarely|Never)\\s+-\\s+([0-9,]+)\\s+([0-9,]+)\\s+([0-9,]+)\\s+([0-9,]+)\\s*$"
m <- regexec(pattern, lines, perl = TRUE)
parts <- regmatches(lines, m)


df_burnedout <- do.call(
  rbind,
  lapply(parts, function(x) {
    
    # x structure from regmatches():
    # x[2] = Organisation type
    # x[3] = response option (Q12b)
    # x[4:(3+length(years))] = yearly counts as strings with commas
    
    n_years <- length(years)
    idx_counts <- 4:(3 + n_years)
    
    if (length(x) < max(idx_counts)) {
      stop(
        "Parsed row does not have enough year values. Row:\n",
        paste(x, collapse = " | ")
      )
    }
    
    counts <- as.integer(gsub(",", "", x[idx_counts]))
    
    # name the counts using the years vector
    year_list <- setNames(as.list(counts), as.character(years))
    
    data.frame(
      Organisation_type = x[2],
      Q12b = x[3],
      year_list,
      stringsAsFactors = FALSE,
      check.names = FALSE
    )
  })
)

respondents_burnedout_by_year <- colSums(df_burnedout[, as.character(years), drop = FALSE])
respondents_burnedout_total <- sum(respondents_burnedout_by_year)

message (paste("\033[32m\033[1mRespondents to burnout question, per online dashboard: ",formatC(respondents_burnedout_total, format="f", big.mark = ",", digits=0), "\033[0m"))
response.rate.burnout <- 100 * respondents_burnedout_total/sum(respondents)
message (paste("\033[32m\033[1mResponse rate to burnout question: ",formatC(response.rate.burnout, format="f", big.mark = ",", digits=1), "%\033[0m"))

