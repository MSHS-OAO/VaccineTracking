# Code to look at 2nd dose Moderna vs Pfizer

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

# Determine whether or not to update walk-in analysis
update_walkins <- FALSE

# Import reference data for site and pod mappings
site_mappings <- read_excel(paste0(user_directory, "/ScheduleData/Automation Ref 2021-01-20.xlsx"), sheet = "Site Mappings")
pod_mappings <- read_excel(paste0(user_directory, "/ScheduleData/Automation Ref 2021-01-20.xlsx"), sheet = "Pod Mappings Simple")

# Store today's date
today <- Sys.Date()

# Create dataframe with dates and weeks
all_dates <- data.frame("Date" = seq.Date(as.Date("1/1/20", format = "%m/%d/%y"), 
                                          as.Date("12/31/21", format = "%m/%d/%y"), 
                                          by = 1))
all_dates <- all_dates %>%
  mutate(Year = year(Date),
         WeekNum = format(Date, "%U"))

# Site order
sites <- c("MSB", "MSBI", "MSH", "MSM", "MSQ", "MSW", "MSHS")

# Pod type order
pod_type <- c("Employee", "Patient", "All")

# NY zip codes
ny_zips <- search_state("NY")

# Import schedule repository
sched_repo <- readRDS(choose.files(default = paste0(user_directory, "/R_Sched_AM_Repo/*.*"), caption = "Select schedule repository"))

# Format and analyze schedule to date for dashboards ---------------------------
sched_to_date <- sched_repo

sched_to_date <- left_join(sched_to_date, site_mappings,
                           by = c("Department" = "Department"))

sched_to_date <- left_join(sched_to_date, pod_mappings[ , c("Provider", "Pod Type")],
                           by = c("Provider/Resource" = "Provider"))

sched_to_date[is.na(sched_to_date$`Pod Type`), "Pod Type"] <- "Patient"

sched_to_date <- sched_to_date %>%
  mutate(Status = ifelse(`Appt Status` %in% c("Arrived", "Comp", "Checked Out"), "Arr", `Appt Status`), # Group various arrival statuses as "Arr"
         Status2 = ifelse(ApptDate == today & Status == "Arr", "Sch", # Categorize any arrivals from today as scheduled for easier manipulation
                          ifelse(ApptDate < today & Status == "Sch", "No Show", Status)), # Categorize any appts still in sch status from prior days as no shows
         # Modify status 2 for running report late in evening instead of early morning
         # Status2 = ifelse(ApptDate == today & Status == "Arr", "Arr", # Categorize any arrivals from today as scheduled for easier manipulation
         #                  ifelse(ApptDate <= today & Status == "Sch", "No Show", Status)), # Categorize any appts still in sch status from prior days as no shows
         WeekNum = format(ApptDate, "%U"),
         DOW = weekdays(ApptDate),
         NYZip = substr(`ZIP Code`, 1, 5) %in% ny_zips$zipcode,
         NameDOBID = ifelse(!(str_detect(Patient, ",\\s[A-Z|a-z]+")), paste(Patient, month(DOB), day(DOB), year(DOB)),
                            paste(str_extract(Patient, "[A-Z|a-z]+,\\s[A-Z|a-z]{1}"), month(DOB), day(DOB), year(DOB))))

# Subset dose 2 scheduled from 2/4/21-2/9/21
start_date <- as.Date("2/4/21", format = "%m/%d/%y")
end_date <- as.Date("2/9/21", format = "%m/%d/%y")

dose2_thru_tue <- sched_to_date %>%
  filter(Dose == 2 & ApptDate >= start_date & ApptDate <= end_date & Status2 == "Sch")

dose2_thru_tue_mrns <- unique(dose2_thru_tue$MRN)

dose2_thru_tue_new_id <- unique(dose2_thru_tue$NameDOBID)

dose1_prior_to_today <- sched_to_date %>%
  filter(Dose == 1 & ApptDate < start_date & Status2 == "Arr") %>%
  mutate(Dose2_MRN = MRN %in% dose2_thru_tue_mrns,
         Dose2_ID = NameDOBID %in% dose2_thru_tue_new_id)

dose1_prior_to_today_mfg_summary <- dose1_prior_to_today %>%
  filter(Dose2_MRN) %>%
  group_by(Site, Mfg) %>%
  summarize(Count = n())









