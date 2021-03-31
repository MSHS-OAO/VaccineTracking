# Code for identifying MRNs with lapsed Dose 2 -----------------

#Install and load necessary packages --------------------
# install.packages("readxl")
# install.packages("writexl")
# install.packages("ggplot2")
# install.packages("lubridate")
# install.packages("dplyr")
# install.packages("reshape2")
# install.packages("svDialogs")
# install.packages("stringr")
# install.packages("formattable")
# install.packages("kableExtra")
# install.packages("ggpubr")
# install.packages("zipcodeR")
# install.packages("tidyr")

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

# Import reference data for site and pod mappings
site_mappings <- read_excel(paste0(user_directory, "/ScheduleData/",
                                   "Automation Ref 2021-03-12.xlsx"),
                            sheet = "Site Mappings")
pod_mappings <- read_excel(paste0(user_directory, "/ScheduleData/",
                                  "Automation Ref 2021-03-12.xlsx"),
                           sheet = "Pod Mappings Simple")

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
               "Network LI")
# Site order
city_sites <- c("MSB",
                "MSBI",
                "MSH",
                "MSM",
                "MSQ",
                "MSW",
                "MSVD")

# Pod type order
pod_type <- c("Employee", "Patient", "All")

# Vaccine manufacturers
mfg <- c("Pfizer", "Moderna", "J&J")

# NY zip codes
ny_zips <- search_state("NY")

# Dose 1 cutoff date
dose1_date <- as.Date("2/28/21", format = "%m/%d/%y")

# Import schedule repository
sched_repo <- readRDS(choose.files(default = paste0(user_directory,
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
    Status2 = ifelse(ApptDate < (today) & Status == "Sch", "No Show", 
                     # Correct any appts erroneously marked as arrived early
                     ifelse(ApptDate >= today & Status == "Arr", "Sch",
                            Status)),
    # Determine manufacturer based on immunization field
    Mfg = #Keep manufacturer as NA if the appointment hasn't been arrived
      ifelse(Status2 != "Arr", NA,
             # Any immunizations with Moderna in text classified as Moderna
             ifelse(str_detect(Immunizations, "Moderna"), "Moderna",
                    # Any visiting docs or Johnson and Johnson immunizations
                    # classified as J&J
                    ifelse(str_detect(Department, "VISITING DOCS") |
                             str_detect(Immunizations, "Johnson and Johnson"),
                           "J&J", "Pfizer"))),
    # Determine week number of year
    WeekNum = format(ApptDate, "%U"),
    # Determine appointment day of week
    DOW = weekdays(ApptDate),
    # Determine if patient's zip code is in NYS
    NYZip = substr(`ZIP Code`, 1, 5) %in% ny_zips$zipcode)

# Crosswalk sites
sched_to_date <- left_join(sched_to_date, site_mappings,
                           by = c("Department" = "Department"))

# Determine MRN who received Dose 1 prior to cutoff date
dose1_mrns <- sched_to_date %>%
  filter(Dose == 1 &
           Status2 == "Arr" &
           ApptDate <= dose1_date) %>%
  select(Site, MRN)

dose1_mrns <- unique(dose1_mrns)

# Count the number of dose 1 received for each MRN
dose1_counts <- sched_to_date %>%
  filter(Dose == 1 &
           Status2 == "Arr") %>%
  group_by(MRN) %>%
  summarize(Count = n())

# Identify MRNs with more than 1 administered Dose 1 (likely scheduling error)
mult_dose1 <- dose1_counts %>%
  filter(Count > 1)

# Remove any MRNs with more than 1 administered Dose 1
dose1_mrns <- dose1_mrns %>%
  filter(!(MRN %in% mult_dose1$MRN))

# Identify MRNs with an administered Dose 2
dose2_mrns <- sched_to_date %>%
  filter(Dose == 2 &
           Status2 == "Arr")

# Determine if patients with administered Dose 1 also received a Dose 2
dose1_mrns <- dose1_mrns %>%
  mutate(ArrivedDose2 = MRN %in% dose2_mrns$MRN)

dose1_mrns_summary <- dose1_mrns %>%
  group_by(Site, ArrivedDose2) %>%
  summarize(Count = n(),
            .groups = "keep") %>%
  mutate(ArrivedDose2 = ifelse(ArrivedDose2, "ArrDose2", "NoArrDose2")) %>%
  pivot_wider(id_cols = Site,
              names_from = ArrivedDose2,
              values_from = Count) %>%
  adorn_totals(where = "row",
               fill = "-",
               na.rm = TRUE,
               name = "MSHS") %>%
  mutate(PercentNoArrDose2 = percent(NoArrDose2 / (NoArrDose2 + ArrDose2), 0.1))

# Create list of MRN without Dose 2
dose1_mrns_no_dose2 <- dose1_mrns %>%
  filter(!ArrivedDose2)

random <- dose1_mrns_no_dose2[sample.int(nrow(dose1_mrns_no_dose2), 10), "MRN"]

random_mrns