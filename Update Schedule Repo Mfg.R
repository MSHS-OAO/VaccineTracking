# Code for adding cancel date to existing schedule repository

#Install and load necessary packages --------------------
#install.packages("readraw_dfl")
#install.packages("writeraw_dfl")
#install.packages("ggplot2")
#install.packages("lubridate")
#install.packages("dplyr")
#install.packages("reshape2")
#install.packages("svDialogs")
#install.packages("stringr")
#install.packages("formattable")
# install.packages("ggpubr")
#install.package(zipcodeR)

#Analysis for weekend discharge tracking
library(readxl)
library(writexl)
library(ggplot2)
library(lubridate)
library(dplyr)
library(reshape2)
library(svDialogs)
library(stringr)
library(formattable)
library(scales)
library(ggpubr)
library(reshape2)
library(knitr)
library(kableExtra)
library(rmarkdown)
library(zipcodeR)

# Set working directory and select raw data ----------------------------
rm(list = ls())

# Determine path for working directory
if (list.files("J://") == "Presidents") {
  user_directory <- "J:/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/System Operations/COVID Vaccine"
} else {
  user_directory <- "J:/deans/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/System Operations/COVID Vaccine"
}

# Set path for file selection
user_path <- paste0(user_directory, "\\*.*")

# Store today's date
today <- Sys.Date()

# Import schedule repository
sched_repo <- readRDS(choose.files(default = paste0(user_directory, "/R_Sched_AM_Repo/*.*"), caption = "Select schedule repository"))

# Add a column for cancel date to match new Epic reports
sched_repo <- sched_repo %>%
  mutate(Mfg = ifelse(is.na(Immunizations), "Pfizer",
                       ifelse(str_detect(Immunizations, "Moderna"), "Moderna", "Pfizer")),
         Dose = ifelse(str_detect(Type, "DOSE 1"), 1, ifelse(str_detect(Type, "DOSE 2"), 2, NA)),
         ApptDate = date(Date),
         ApptYear = year(Date),
         ApptMonth = month(Date),
         ApptDay = day(Date),
         ApptWeek = week(Date),
         Department = ifelse(str_detect(Department, ","), substr(Department, 1, str_locate(Department, ",") - 1), Department))

# Export repository to new .RDS 
saveRDS(sched_repo, paste0(user_directory, "/R_Sched_AM_Repo/Sched Updated Mfg ",
                           format(min(sched_repo$ApptDate), "%m%d%y"), " to ",
                           format(max(sched_repo$ApptDate), "%m%d%y"),
                           " on ", format(Sys.time(), "%m%d%y %H%M"), ".rds"))


# Format and analyze schedule to date for dashboards ---------------------------
sched_to_date <- sched_repo

sched_to_date <- sched_to_date %>%
  mutate(Status = ifelse(`Appt Status` %in% c("Arrived", "Comp", "Checked Out"), "Arr", `Appt Status`), # Group various arrival statuses as "Arr"
         Status2 = ifelse(ApptDate == today & Status == "Arr", "Sch", # Categorize any arrivals from today as scheduled for easier manipulation
                          ifelse(ApptDate < today & Status == "Sch", "No Show", Status)), # Categorize any appts still in sch status from prior days as no shows
         Status3 = ifelse(ApptDate <today & Status == "Arr" & (is.na(`Level Of Service`) | str_detect(`Level Of Service`, "ERRONEOUS")), "Left", Status2),
         Test = Status2 == Status3,
         # Modify status 2 for running report late in evening instead of early morning
         # Status2 = ifelse(ApptDate == today & Status == "Arr", "Arr", # Categorize any arrivals from today as scheduled for easier manipulation
         #                  ifelse(ApptDate <= today & Status == "Sch", "No Show", Status)), # Categorize any appts still in sch status from prior days as no shows
         WeekNum = format(ApptDate, "%U"),
         DOW = weekdays(ApptDate))

change_summary <- sched_to_date %>%
  group_by(Status2, Status3, Test) %>%
  summarize(Count = n())
