#Analysis for analyzing outreach efforts

# Import libraries
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
if ("Presidents" %in% list.files("J://")) {
  user_directory <- paste0("J:/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "System Operations/COVID Vaccine")
} else {
  user_directory <- paste0("J:/deans/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "System Operations/COVID Vaccine")
}

# Set path for file selection
user_path <- paste0(user_directory, "\\*.*")

# Import reference data for site and pod mappings
site_mappings <- read_excel(
  paste0(
    user_directory, "/AdHoc/", "Automation Ref w Outreach 2021-03-04.xlsx"),
  sheet = "Site Mappings")

pod_mappings <- read_excel(
  paste0(
    user_directory, "/AdHoc/", "Automation Ref w Outreach 2021-03-04.xlsx"),
  sheet = "Pod Mappings Simple")

scheduler_mappings <- read_excel(
  paste0(
    user_directory, "/AdHoc/", "Automation Ref w Outreach 2021-03-04.xlsx"),
  sheet = "Outreach Mappings")

# Store today's date
today <- Sys.Date()

# Create dataframe with dates and weeks
all_dates <- data.frame("Date" = seq.Date(as.Date("1/1/20",
                                                  format = "%m/%d/%y"),
                                          as.Date("12/31/21",
                                                  format = "%m/%d/%y"),
                                          by = 1))
all_dates <- all_dates %>%
  mutate(Year = year(Date),
         WeekNum = format(Date, "%U"))

# Site order
sites <- c("MSB", "MSBI", "MSH", "MSM", "MSQ", "MSW", "MSVD", "MSHS")

# Pod type order
pod_type <- c("Employee", "Patient", "All")

# Vaccine manufacturers
mfg <- c("Pfizer", "Moderna", "J&J")

# NY zip codes
ny_zips <- search_state("NY")

# Select dates of interest
start_date <- as.Date("2/24/21", format("%m/%d/%y"))
end_date <- as.Date("2/28/21", format("%m/%d/%y"))

# Import schedule repository
sched_repo <- readRDS(
  choose.files(
    default = paste0(user_directory,
                     "/R_Sched_AM_Repo/*.*"),
    caption = "Select schedule repository"))

# Format and analyze schedule to date for dashboards ---------------------------
sched_to_date <- sched_repo

sched_to_date <- sched_to_date %>%
  mutate(
    # Determine whether appointment is for dose 1 or dose 2
    Dose = ifelse(str_detect(Type, "DOSE 1"), 1,
                  ifelse(str_detect(Type, "DOSE 2"), 2, NA)),
    # Determine appointment date, year, month, etc.
    ApptDate = date(Date),
    ApptYear = year(Date),
    ApptMonth = month(Date),
    ApptDay = day(Date),
    ApptWeek = week(Date),
    # Clean up department if it is duplicated in that column
    Department = ifelse(str_detect(Department, ","),
                        substr(Department, 1,
                               str_locate(Department, ",") - 1),
                        Department),
    # Group arrived, completed, and checked out appts as arrived
    Status = ifelse(`Appt Status` %in% c("Arrived", "Comp", "Checked Out"),
                    "Arr", `Appt Status`),
    # (Only needed when manually pulling data in the morning)
    # Classify any arrivals from today as scheduled for morning export and
    # classify any appts from prior days that are still in scheduled status as
    # no shows
    # Status2 = ifelse(ApptDate == today & Status == "Arr", "Sch",
    #                  ifelse(ApptDate < today & Status == "Sch", "No Show",
    # Status)),
    # Update appointment status: classify any appts from prior days that are
    # still in scheduled as No Shows
    Status2 = ifelse(ApptDate < (today) & Status == "Sch", "No Show", Status),
    # Determine manufacturer based on immunization field
    Mfg = #Keep manufacturer as NA if the appointment hasn't been arrived
      ifelse(Status2 != "Arr", NA,
             # Assume any visits without immunization record are Pfizer
             ifelse(is.na(Immunizations), "Pfizer",
                    # Any immunizations with Moderna in text classified as
                    # Moderna, otherwise Pfizer
                    ifelse(str_detect(Immunizations, "Moderna"), "Moderna",
                           "Pfizer"))),
    # Determine week number of year
    WeekNum = format(ApptDate, "%U"),
    # Determine appointment day of week
    DOW = weekdays(ApptDate),
    # Determine if patient's zip code is in NYS
    NYZip = substr(`ZIP Code`, 1, 5) %in% ny_zips$zipcode,
    DateOfInterest = (ApptDate >= start_date & ApptDate <= end_date),
    #
    # Format scheduling person/method for outreach mappings
    SchedulerLastName = substr(
      str_extract(`Entry Person`, ".+\\,"), 1,
      nchar(str_extract(`Entry Person`, ".+\\,")) - 1),
    SchedulerFirstName = substr(
      str_extract(`Entry Person`, "\\,\\s[A-Z]+\\s"), 3,
      nchar(str_extract(`Entry Person`, "\\,\\s[A-Z]+\\s")) - 1),
    ScheduledBy = paste(SchedulerFirstName, SchedulerLastName),
    #
    SchedulerLastName = NULL,
    SchedulerFirstNamer = NULL)

# Crosswalk sites
sched_to_date <- left_join(sched_to_date, site_mappings,
                           by = c("Department" = "Department"))

# Crosswalk pod type (employee vs patient)
sched_to_date <- left_join(sched_to_date,
                           pod_mappings[, c("Provider", "Pod Type")],
                           by = c("Provider/Resource" = "Provider"))

# Crosswalk scheduled by grouper
scheduler_mappings <- scheduler_mappings %>%
  mutate(Name = toupper(Name))

sched_to_date <- left_join(sched_to_date,
                           scheduler_mappings,
                           by = c("ScheduledBy" = "Name"))

# Update scheduling grouper for same day appointments and other methods
sched_to_date <- sched_to_date %>%
  mutate(SchedulingSource = ifelse(is.na(SchedulingSource),
                                   "Other", SchedulingSource))
           # ifelse(is.na(SchedulingSource) &
           #                            `Same Day?` == "Same day",
           #                          "Same Day",
           #                          ifelse(is.na(SchedulingSource),
           #                                 "Other", SchedulingSource)))
#
# Subset schedule for appointments within date range of interest
outreach_sched <- sched_to_date %>%
  filter(DateOfInterest & Dose == 1)

outreach_summary <- outreach_sched %>%
  group_by(SchedulingSource) %>%
  summarize(Count = n(), .groups = "keep") %>%
  arrange(-Count)

other_summary <- outreach_sched %>%
  filter(SchedulingSource == "Other") %>%
  group_by(`Entry Person`) %>%
  summarize(Count = n()) %>%
  arrange(-Count)

outreach_site_status_summary <- outreach_sched %>%
  group_by(SchedulingSource,
           Site,
           Status2) %>%
  summarize(Count = n(), .groups = "keep")

outreach_list <- list("Summary" = outreach_summary,
                      "OtherBreakout" = other_summary,
                      "SiteAndApptStatus" = outreach_site_status_summary)

write_xlsx(outreach_list,
           path = 
             paste0(user_directory, "/AdHoc/Outreach Sched Analysis ",
                    format(Sys.time(), "%m-%d-%y %H%M"), ".xlsx"))


