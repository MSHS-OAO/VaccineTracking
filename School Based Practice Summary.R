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
library(tidyr)
library(purrr)
library(janitor)

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

# Determine whether or not to update an existing repo
update_repo <- TRUE

# Determine whether or not to update walk-in analysis
update_walkins <- FALSE

# Import reference data for site and pod mappings
site_mappings <- read_excel(paste0(user_directory, "/ScheduleData/",
                                   "Automation Ref 2022-01-31.xlsx"),
                            sheet = "Site Mappings")
pod_mappings <- read_excel(paste0(user_directory, "/ScheduleData/",
                                  "Automation Ref 2022-01-31.xlsx"),
                           sheet = "Pod Mappings Simple")
school_based_practices <- read_excel(paste0(user_directory, "/ScheduleData/",
                                            "Automation Ref 2022-01-31.xlsx"),
                                     sheet = "School Based Practices")

school_based_practices_dept <- (unique(school_based_practices[, c("Epic Department",
                                                                "Vaccine Type")]))

school_based_practices_dept <- school_based_practices_dept %>%
  filter(!is.na(`Epic Department`)) %>%
  rename(Epic_Dept = `Epic Department`,
         Age_Group = `Vaccine Type`)

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
all_sites <- c("MSB",
               "MSBI",
               "MSH",
               "MSM",
               "MSQ",
               "MSW",
               "MSVD",
               "Network LI")#,
# "MSHS")

# Site order
city_sites <- c("MSB",
                "MSBI",
                "MSH",
                "MSM",
                "MSQ",
                "MSW",
                "MSVD")#,
# "MSHS")

# Pod type order
pod_type <- c("Employee", "Patient", "All")

# Vaccine manufacturers
mfg <- c("Pfizer", "Moderna", "J&J")

# Vaccine types
vax_types <- c("Adult", "Peds")

# Pediatric school based practices
peds_school_practice_dept <- c("HOSP SBH PS 38",
                               "HOSP SBH PS 83/ PS 182",
                               "HOSP SBH PS 108",
                               "HOSP SBH TPEC",
                               "HOSP PEDS SCH HEALTH")

# NY zip codes
ny_zips <- search_state("NY")

# Import schedule repository and determine last date it was updated
sched_repo_file <- choose.files(default = paste0(user_directory,
                                                 "/R_Sched_AM_Repo/*.*"))

sched_repo <- readRDS(sched_repo_file)
sched_repo_update_date <- date(file.info(sched_repo_file)$ctime)

# Format and analyze schedule to date for dashboards ---------------------------
sched_to_date <- sched_repo

sched_to_date <- sched_to_date %>%
  rename(IsPtEmployee = `Is the patient a Mount Sinai employee`) %>%
  mutate(
    # Update timezones to EST if imported as UTC
    `Appt Time` = with_tz(`Appt Time`, tzone = "EST"),
    # Determine whether appointment is for dose 1, 2, or 3 based on scheduled visit type
    Dose = ifelse(str_detect(Type, "DOSE 1"), 1,
                  ifelse(str_detect(Type, "DOSE 2"), 2,
                         ifelse(str_detect(Type, "DOSE 3"), 3, NA))),
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
    Status2 = ifelse(ApptDate < (today) & Status == "Sch", "No Show", 
                     # Correct any appts erroneously marked as arrived early
                     ifelse(ApptDate >= today & Status == "Arr", "Sch",
                            Status)),
    # Determine whether the vaccine is an adult or pediatric dose
    VaxType = ifelse(Department %in% c("1468 MADISON PEDIATRIC VACCINE POD") |
                       `Provider/Resource` %in% c("DOSE 1 PEDIATRIC [1324684]",
                                                  "DOSE 2 PEDIATRIC [1324685]") |
                       Department %in% peds_school_practice_dept,
                     "Peds", "Adult"),
    # Determine manufacturer based on immunization field and vaccine type fields
    Mfg = #Keep manufacturer as NA if the appointment hasn't been arrived
      ifelse(Status2 != "Arr", NA,
             # Any blank immunizations and any peds assumed to be Pfizer
             ifelse(is.na(Immunizations) |
                      VaxType %in% c("Peds"), "Pfizer",
                    # Any immunizations with Moderna in text classified as Moderna
                    ifelse(str_detect(Immunizations, "Moderna"), "Moderna",
                           # Any visiting docs or Johnson and Johnson
                           # immunizations classified as J&J
                           ifelse(str_detect(Department, "VISITING DOCS") |
                                    str_detect(Immunizations,
                                               "Johnson and Johnson"),
                                  "J&J", "Pfizer")))),
    # Determine week number of year
    WeekNum = format(ApptDate, "%U"),
    # Determine appointment day of week
    DOW = weekdays(ApptDate),
    # Determine if patient's zip code is in NYS
    NYZip = substr(`ZIP Code`, 1, 5) %in% ny_zips$zipcode,
    # Determine if patient is a MSHS employee based on response to question
    MSHSEmployee =
      # First see if patient responded "YES" to question about MSHS employee
      str_detect(
        str_replace_na(IsPtEmployee, ""),
        regex("yes", ignore_case = TRUE)) |
      # Next, see if there is a MSHS site listed for patients who were vaccinated
      # prior to the addition of the MSHS employee field
      str_detect(
        str_replace_na(`MOUNT SINAI HEALTH SYSTEM`, " "),
        regex("[a-z]", ignore_case = TRUE)),
    School_Practice = Department %in% school_based_practices_dept$Epic_Dept)

# Crosswalk sites
sched_to_date <- left_join(sched_to_date, site_mappings,
                           by = c("Department" = "Department"))

# Manually update discrepancies in site/mfg/dose
sched_to_date <- sched_to_date %>%
  mutate(# Roll up any Moderna administered at MSB for YWCA to MSH
    Site = ifelse(Site %in% c("MSB") & Mfg %in% c("Moderna"), "MSH", Site),
    # Manually correct any J&J doses erroneously documented as dose 2 and change to dose 1
    Dose = ifelse(Mfg %in% c("J&J") & Dose %in% c(2), 1, Dose))

# Crosswalk pod type (employee vs patient)
sched_to_date <- left_join(sched_to_date,
                           unique(pod_mappings[, c("Provider", "Pod Type")]),
                           by = c("Provider/Resource" = "Provider"))

sched_sites <- unique(sched_to_date$Site)

# School based practice summary
school_practice_summary <- sched_to_date %>%
  filter(School_Practice) %>%
  group_by(Department, Status2, ApptDate, Dose) %>%
  summarize(Count = n()) %>%
  arrange(Department, ApptDate, Dose, Status2)
