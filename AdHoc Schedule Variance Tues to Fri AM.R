# Code for comparing schedules ----------------------

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

# Determine whether or not to update walk-in analysis
update_walkins <- FALSE

# Import reference data for site and pod mappings
site_mappings <- read_excel(paste0(user_directory, "/ScheduleData/Automation Ref 2021-02-05.xlsx"), sheet = "Site Mappings")
pod_mappings <- read_excel(paste0(user_directory, "/ScheduleData/Automation Ref 2021-02-05.xlsx"), sheet = "Pod Mappings Simple")

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


# Import schedule repositories
sched_repo_013121 <- readRDS(file = paste0(user_directory, "/R_Sched_AM_Repo/Sched 121520 to 021621 on 013121 0759.rds"))

sched_repo_020521 <- readRDS(file = paste0(user_directory, "/R_Sched_AM_Repo/Sched 121520 to 021921 on 020521 0756.rds"))

# Format and analyze schedule to date for dashboards ---------------------------
# Preprocess Tuesday morning's schedule repository
sched_sun <- sched_repo_013121

sched_sun <- left_join(sched_sun, site_mappings,
                           by = c("Department" = "Department"))

sched_sun <- left_join(sched_sun, pod_mappings[ , c("Provider", "Pod Type")],
                           by = c("Provider/Resource" = "Provider"))

sched_sun[is.na(sched_sun$`Pod Type`), "Pod Type"] <- "Patient"

sched_sun <- sched_sun %>%
  mutate(Status = ifelse(`Appt Status` %in% c("Arrived", "Comp", "Checked Out"), "Arr", `Appt Status`), # Group various arrival statuses as "Arr"
         Status2 = ifelse(ApptDate == as.Date("1/31/21", format = "%m/%d/%y") & Status == "Arr", "Sch", # Categorize any arrivals from today as scheduled for easier manipulation
                          ifelse(ApptDate < as.Date("1/31/21", format = "%m/%d/%y") & Status == "Sch", "No Show", Status)), # Categorize any appts still in sch status from prior days as no shows
         # Modify status 2 for running report late in evening instead of early morning
         # Status2 = ifelse(ApptDate == today & Status == "Arr", "Arr", # Categorize any arrivals from today as scheduled for easier manipulation
         #                  ifelse(ApptDate <= today & Status == "Sch", "No Show", Status)), # Categorize any appts still in sch status from prior days as no shows
         WeekNum = format(ApptDate, "%U"),
         DOW = weekdays(ApptDate),
         NYZip = substr(`ZIP Code`, 1, 5) %in% ny_zips$zipcode)


# Preprocess Friday morning's schedule repository
sched_fri <- sched_repo_020521

sched_fri <- left_join(sched_fri, site_mappings,
                       by = c("Department" = "Department"))

sched_fri <- left_join(sched_fri, pod_mappings[ , c("Provider", "Pod Type")],
                       by = c("Provider/Resource" = "Provider"))

sched_fri[is.na(sched_fri$`Pod Type`), "Pod Type"] <- "Patient"

sched_fri <- sched_fri %>%
  mutate(Status = ifelse(`Appt Status` %in% c("Arrived", "Comp", "Checked Out"), "Arr", `Appt Status`), # Group various arrival statuses as "Arr"
         Status2 = ifelse(ApptDate == as.Date("2/5/21", format = "%m/%d/%y") & Status == "Arr", "Sch", # Categorize any arrivals from today as scheduled for easier manipulation
                          ifelse(ApptDate < as.Date("2/5/21", format = "%m/%d/%y") & Status == "Sch", "No Show", Status)), # Categorize any appts still in sch status from prior days as no shows
         # Modify status 2 for running report late in evening instead of early morning
         # Status2 = ifelse(ApptDate == today & Status == "Arr", "Arr", # Categorize any arrivals from today as scheduled for easier manipulation
         #                  ifelse(ApptDate <= today & Status == "Sch", "No Show", Status)), # Categorize any appts still in sch status from prior days as no shows
         WeekNum = format(ApptDate, "%U"),
         DOW = weekdays(ApptDate),
         NYZip = substr(`ZIP Code`, 1, 5) %in% ny_zips$zipcode)


# Compare schedule outlook as reported on Tuesday vs. Friday -----------------------
start_date <- as.Date("2/1/21", format = "%m/%d/%y")
end_date <- as.Date("2/16/21", format = "%m/%d/%y")

dose1_sch_sun <- sched_sun %>%
  filter(ApptDate >= start_date & ApptDate <= end_date & Dose == 1 & Status2 == "Sch")

dose1_arr_sch_fri <- sched_fri %>%
  filter(ApptDate >= start_date & ApptDate <= end_date & Dose == 1 & Status2 %in% c("Arr", "Sch"))

dose1_sun_summary <- dose1_sch_sun %>%
  group_by(Site, `Pod Type`, ApptDate) %>%
  summarize(Count = n(),
            SchedAsOf = "Morning of 2/2/21")

dose1_sun_summary_cast <- dcast(dose1_sun_summary,
                                SchedAsOf + Site + `Pod Type` ~ ApptDate,
                                value.var = "Count")

dose1_fri_summary <- dose1_arr_sch_fri %>%
  group_by(Site, `Pod Type`, ApptDate) %>%
  summarize(Count = n(),
            SchedAsOf = "Morning of 2/5/21")

dose1_fri_summary_cast <- dcast(dose1_fri_summary,
                                SchedAsOf + Site + `Pod Type` ~ ApptDate,
                                value.var = "Count")

schedule_variance <- dose1_fri_summary_cast %>%
  mutate(SchedAsOf = "Sched as of Fri AM - Sched as of Sun AM")

schedule_variance[ , 4:ncol(schedule_variance)] <- dose1_fri_summary_cast[ , 4:ncol(schedule_variance)] - dose1_sun_summary_cast[ , 4:ncol(schedule_variance)]

colnames(schedule_variance)[1] <- "Variance"

export_list <- list("Dose1_Sched_013121AM" = dose1_sun_summary_cast,
                    "Dose1_ArrSched_020521AM" = dose1_fri_summary_cast,
                    "Dose1_Variance" = schedule_variance)

write_xlsx(export_list, path = paste0(user_directory, "/AdHoc/Dose1 Schedule Variance 2020-02-05.xlsx"))
