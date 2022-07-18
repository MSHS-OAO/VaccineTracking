# Script to summarize arrived COVID-19 vaccine appointments for 2020-2021

# Clear environment and load libraries ---------------
# Clear environment
rm(list = ls())

# Load libraries
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
library(DT)

# Reference data -----------------------
# Determine directory
if ("Presidents" %in% list.files("J://")) {
  user_directory <- paste0("J:/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "System Operations/COVID Vaccine")
} else {
  user_directory <- paste0("J:/deans/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "System Operations/COVID Vaccine")
}

update_repo <- FALSE

# Set path for file selection
user_path <- paste0(user_directory, "\\*.*")

# Import reference data for site and pod mappings
dept_mappings <- read_excel(paste0(
  user_directory,
  "/Weekly Dashboard Doses Administered",
  "/MSHS COVID-19 Vaccine Department Mappings 2022-04-11.xlsx"))

dept_mappings <- dept_mappings %>%
  select(-`Site (Historical)`, -Notes)

# Store today's date
today <- Sys.Date()

# Import schedule repo --------------------------
sched_repo_file <- paste0(user_directory,
                          "/R_Sched_AM_Repo",
                          "/Auto Epic Report Sched 121520 to 062922 on 061422 0930.rds")

sched_repo <- readRDS(sched_repo_file)

# Process and format schedule repository for standard reporting ---------------
sched_to_date <- sched_repo

sched_to_date <- sched_to_date %>%
  # rename(IsPtEmployee = `Is the patient a Mount Sinai employee`) %>%
  mutate(
    # Update timezones to EST if imported as UTC
    `Appt Time` = with_tz(`Appt Time`, tzone = "EST"),
    # Determine whether appointment is for dose 1, 2, or 3 based on scheduled visit type
    Dose = ifelse(str_detect(Type, "DOSE 1"), "Dose 1",
                  ifelse(str_detect(Type, "DOSE 2"), "Dose 2",
                         ifelse(str_detect(Type, "(DOSE 3)|(BOOSTER)"),
                                "Dose 3/Boosters", NA))),
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
    # Update appointment status: classify any appts from prior days that are
    # still in scheduled as No Shows
    Status2 = ifelse(ApptDate < (today) & Status == "Sch", "No Show", 
                     # Correct any appts erroneously marked as arrived early
                     ifelse(ApptDate >= today & Status == "Arr", "Sch",
                            Status)),
    # Determine week number of year
    WeekNum = format(ApptDate, "%U"),
    # Determine appointment day of week
    DOW = weekdays(ApptDate)
  ) %>%
  filter(ApptYear %in% c(2020, 2021) &
           Status2 %in% "Arr")

# Crosswalk sites and department grouper
sched_to_date <- left_join(sched_to_date, dept_mappings,
                           by = c("Department" = "Department"))

sched_to_date <- sched_to_date %>%
  rename(Site = `New Site`)

sched_to_date <- sched_to_date %>%
  mutate(
    # Create a column for age grouper
    Age_Grouper = ifelse(str_detect(Age, "\\ y.o."), "years",
                         ifelse(str_detect(Age, "\\ m.o."), "months",
                                ifelse(str_detect(Age, "\\ days"), "days",
                                       "unknown"))),
    # Convert string with age to numeric
    Age_Numeric = 
      ifelse(Age_Grouper == "years",
             as.numeric(str_extract(Age,
                                    # pattern = "(?<=Deceased\\ \\().+(?=\\ y.o.)")),
                                    paste("(?<=Deceased\\ \\()[0-9]+(?=\\ y.o.)",
                                          "[0-9]+(?=\\ y.o.)",
                                          sep = "|"))),
             ifelse(Age_Grouper == "months",
                    as.numeric(str_extract(Age, "[0-9]+(?=\\ m.o.)")) / 12,
                    ifelse(Age_Grouper == "days",
                           as.numeric(
                             str_extract(Age, "[0-9]+(?=\\ days)")) / 365,
                           ifelse(Age_Grouper == "unknown", 10000,
                                  NA_integer_)))),
    # Determine vaccine type based on department, provider/resource, and age if necessary
    VaxType = ifelse(
      `Vaccine Age Group` %in% "Peds" |
        # MSBI Peds POD
        `Provider/Resource` %in% c("DOSE 1 PEDIATRIC [1324684]",
                                   "DOSE 2 PEDIATRIC [1324685]") |
        # Practics that administer adult and pediatric vaccines
        (`Vaccine Age Group` %in% "Both" &
           `Setting Grouper` %in% c("School Based Practice", "MSHS Practice") &
           Age_Numeric < 12),
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
    # Site =
    #   # Roll up any Moderna administered at MSB for YWCA to MSH
    #   ifelse(Site %in% c("MSB") & Mfg %in% c("Moderna"), "MSH",
    #          # Manually classify school practices as peds or adolescent for schools administering both
    #          ifelse(`Setting Grouper` %in% "School Based Practice" &
    #                   VaxType %in% "Peds", "Peds School Based Practice",
    #                 ifelse(`Setting Grouper` %in% "School Based Practice" &
    #                          VaxType %in% "Adult",
    #                        "Adolescent School Based Practice",
    #                        Site))),
    # Manually correct any J&J doses erroneously documented as dose 2 and change to dose 1
    Dose = ifelse(Mfg %in% c("J&J") & Dose %in% c("Dose 2"), "Dose 1", Dose),
    `Setting Grouper` = ifelse(`Setting Grouper` %in% "School Based Practice",
                               "MSHS Practice", `Setting Grouper`),
    # Rename vaccine type with new standard name
    VaxType = case_when(VaxType %in% "Adult" ~ "Adult (12 yo +)",
                        VaxType %in% "Peds" ~ "Peds (5-11 yo)"))

# Sched repo summary
sched_repo_summary <- sched_to_date %>%
  group_by(ApptDate, Department, `Setting Grouper`, Site,
           VaxType, Dose, Mfg) %>%
  summarize(Count = n()) %>%
  rename(Date = ApptDate,
         Setting_Type = `Setting Grouper`)

# Save standardized, static schedule repository as .RDS for use in future reporting --------------------------
saveRDS(sched_repo_summary,
        paste0(user_directory,
               "/Hybrid Repo Sched & Admin Data",
               "/Sched Repo Summary 2020-2021 2022-07-18.RDS"))
