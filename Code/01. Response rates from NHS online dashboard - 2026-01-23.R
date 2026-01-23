# This file is available at https://github.com/ebmgt/NHS-Religion/
# Authors: Robert Badgett. rbadgett@kumc.edu
# Permission: GNU GPLv3 https://choosealicense.com/licenses/gpl-3.0/
# Last edited 2025-01-09

# https://nhssurveys.co.uk/nss/questions/national
# https://www.nhsemployers.org/news/nhs-staff-survey-2023-key-findings

# https://nhssurveys.co.uk/nss/response_rate/national
# Exclude MH & LD, and MH, LD & Community
year <- c(2021,2022,2023)

# Response rate to survey starting 2021:
# https://nhssurveys.co.uk/nss/response_rate/national
respondents <- c(628475, 679088, 678676, 747288)
response.rates <- c(50.14, 47.67, 49.73, 51.65)

# 1. Back-calculate recipients per year to use as weights
recipients_per_year <- (respondents / (response.rates / 100))
# 2. Calculate the weighted average of response rates
# Each year's rate is weighted by that year's number of recipients
response.rate.overall <- weighted.mean(response.rates, w = recipients_per_year)
# Output formatting
respondents.total <- formatC(sum(respondents), format = "f", big.mark = ",", digits = 0)
message(paste("\033[32m\033[1mRespondents overall: ", respondents.total, "\033[0m"))

message(paste("\033[32m\033[1mWeighted response rate over study years: ", 
              formatC(response.rate.overall, format = "f", big.mark = ",", digits = 1), "%\033[0m"))
# Responses to the patient facing only over three years for all 4 Reporting Group:
# https://nhssurveys.co.uk/nss/questions/national
respondents <- c(622286,	606895,	673334)
(respondents.total <- formatC(sum(respondents), format = "f", big.mark = ",", digits = 0))

# Responses to the burnout only over three years for all 4 Reporting Group:
# https://nhssurveys.co.uk/nss/questions/national
respondents <- c(615298,	609245,	675383)
(respondents.total <- sum(respondents))
message (paste("\033[32m\033[1mRespondents to burnout question, per online dashboard: ",formatC(respondents.total, format="f", big.mark = ",", digits=0), "\033[0m"))
response.rate.burnout <- 100*sum(respondents.total)/recipients
message (paste("\033[32m\033[1mResponse rate to burnout question: ",formatC(response.rate.burnout, format="f", big.mark = ",", digits=1), "\033[0m"))

nhs.responses <- as.data.frame (cbind (year, respondents, response.rates))

nhs.responses$recipients <- round(nhs.responses$respondents/(nhs.responses$response.rate/100),0)
(receipients.total <- sum(nhs.responses$recipients))

(response.rate <- sum(nhs.responses$respondents)/sum(nhs.responses$recipients))

# ___________________----
# Burnout data from NHS online dashboard ------
nhs.responses$burnout.numerator   <- c(615298,	609245,	675383)
nhs.responses$burnout.rate        <- c(34.49, 33.97, 30.38)
nhs.responses$burnout.denominator <- round(nhs.responses$burnout.numerator/(nhs.responses$burnout.rate/100),0) 
(sum(nhs.responses$burnout.numerator))

(nhs.responses$burnout.response.rate <- round(nhs.responses$burnout.denominator/nhs.responses$respondents*100,2))
(sum(nhs.responses$burnout.denominator)/sum(nhs.responses$respondents))

# Add a summary row to the table
summary_row <- data.frame(
  year = "All",
  respondents = sum(nhs.responses$respondents),
  response.rate = round(
    sum(nhs.responses$response.rate * nhs.responses$respondents) / sum(nhs.responses$respondents), 2
  ),
  recipients = sum(nhs.responses$recipients),
  burnout.numerator = sum(nhs.responses$burnout.numerator),
  burnout.rate = round(
    sum(nhs.responses$burnout.rate * nhs.responses$burnout.numerator) / sum(nhs.responses$burnout.numerator), 2
  ),
  burnout.denominator = sum(nhs.responses$burnout.denominator),
  burnout.response.rate = round(
    sum(nhs.responses$burnout.response.rate * nhs.responses$respondents) / sum(nhs.responses$respondents), 2
  )
)

names(nhs.responses)[names(nhs.responses) == "response.rates"] <- "response.rate"
nhs.responses <- rbind(nhs.responses, summary_row)
DT::datatable(nhs.responses, caption = "nhs.responses", options = list(pageLength = 25))
