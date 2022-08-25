# Code for exporting vaccines administered by week for Jen Paez

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
#install.packages("ggpubr")
#install.packages("zipcodeR")
#install.packages("tidyr")
#install.packages("janitor")
#install.packages("knitr")
#install.packages("rmarkdown")
#install.packages("kableExtra")

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
library(knitr)
library(rmarkdown)
library(kableExtra)

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
dept_mappings <- read_excel(paste0(
  user_directory,
  "/Weekly Dashboard Doses Administered",
  "/MSHS COVID-19 Vaccine Department Mappings 2022-08-25.xlsx"))

dept_mappings <- dept_mappings %>%
  select(-`New Site`, -Notes) %>%
  rename(Site = `Site (Historical)`)


# Store today's date
today <- Sys.Date()
yesterday <- today - 1

# Site order
all_sites <- c("MSB",
               "MSBI",
               "MSH",
               "MSM",
               "MSQ",
               "MSW",
               "MSVD",
               "Network LI",
               "MSSN")

peds_sites <- c("MSH", "MSBI")

# Vaccine manufacturers
mfg <- c("Pfizer", "Moderna", "J&J")

# Vaccine types
vax_types <- c("Adult", "Peds")

# NY zip codes
ny_zips <- search_state("NY")

# Create dataframe with start dates for adult and peds vaccination
vacc_start_df <- data.frame("VaxType" = c("Adult", "Peds"),
                            "StartDate" = c(as.Date("12/15/20",
                                                    format = "%m/%d/%y"),
                                            as.Date("11/04/21",
                                                    format = "%m/%d/%y")))

vacc_start_df <- vacc_start_df %>%
  mutate(
    # Determine the Sunday of the first week of vaccination
    FirstSun = StartDate - (wday(StartDate) - 1),
    # Determine this week number
    ThisWeek = as.numeric(floor((today - FirstSun) / 7) + 1))

# Create dataframe with dates and weeks since vaccines started
all_dates <- data.frame("Date" = seq.Date(min(vacc_start_df$StartDate),
                                          today + 7,
                                          by = 1
)
)

all_dates <- all_dates %>%
  mutate(
    # Determine vaccine week for adult vaccinations
    Adult_VaccWeek = as.numeric(
      floor((Date -
               vacc_start_df$FirstSun[
                 which(vacc_start_df$VaxType == "Adult")]) / 7) + 1),
    # Determine vaccine week for pediatric vaccinations
    Peds_VaccWeek = ifelse(
      as.numeric(
        floor((Date -
                 vacc_start_df$FirstSun[
                   which(vacc_start_df$VaxType == "Peds")]) / 7) + 1) < 1, NA,
      as.numeric(
        floor((Date -
                 vacc_start_df$FirstSun[
                   which(vacc_start_df$VaxType == "Peds")]) / 7) + 1)
    )
  ) %>%
  pivot_longer(cols = contains("_VaccWeek"),
               names_to = "VaxType",
               values_to = "VaccWeek") %>%
  mutate(VaxType = str_remove(VaxType, "_VaccWeek"),
         WeekStart = Date - (wday(Date) - 1),
         WeekEnd = Date + (7 - wday(Date)),
         WeekDates = paste0(
           format(WeekStart, "%m/%d/%y"),
           "-",
           format(WeekEnd, "%m/%d/%y")
         ),
         WeekStart = NULL,
         WeekEnd = NULL
  ) %>%
  filter(!is.na(VaccWeek)) %>%
  arrange(VaxType, Date)


# Import Epic schedule repository
sched_repo <- readRDS(
  choose.files(default = paste0(user_directory,
                                  "/R_Sched_AM_Repo/*.*"),
                 caption = "Select this morning's schedule repository"))

# Format and analyze Epic schedule to date for dashboards ---------------------
sched_to_date <- sched_repo

sched_to_date <- sched_to_date %>%
  mutate(
    # Determine whether appointment is for dose 1 or dose 2
    Dose = ifelse(str_detect(Type, "DOSE 1"), "Dose 1",
                  ifelse(str_detect(Type, "DOSE 2"), "Dose 2",
                         ifelse(str_detect(Type, "(DOSE 3)|(BOOSTER)"), 
                                "Dose 3/Booster", NA))),
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
    # # Determine whether the vaccine is an adult or pediatric dose
    # VaxType = ifelse(Department %in% c("1468 MADISON PEDIATRIC VACCINE POD",
    #                                    "1468 MADISON HOSP PEDS GEN",
    #                                    "10 UNION SQ E PEDS GEN",
    #                                    "234 E 85TH ST PEDS GEN",
    #                                    "5 E 98  PEDS GENERAL") |
    #                    `Provider/Resource` %in% c("DOSE 1 PEDIATRIC [1324684]",
    #                                               "DOSE 2 PEDIATRIC [1324685]"),
    #                  "Peds", "Adult"),
    # # Determine manufacturer based on immunization and vaccine type fields
    # Mfg = #Keep manufacturer as NA if the appointment hasn't been arrived
    #   ifelse(Status2 != "Arr", NA,
    #          # Assume any visits without immunization record are Pfizer
    #          ifelse(is.na(Immunizations) |
    #                   VaxType %in% c("Peds"), "Pfizer",
    #                 # Any immunizations with Moderna in text classified as
    #                 # Moderna, otherwise Pfizer
    #                 ifelse(str_detect(Immunizations, "Moderna"), "Moderna",
    #                        # Any immunizations with Johnson and Johnson are
    #                        # classified as J&J
    #                        ifelse(str_detect(Department, "VISITING DOCS") |
    #                                 str_detect(Immunizations,
    #                                            "Johnson and Johnson"),
    #                               "J&J", "Pfizer")))),
    # # Determine week number of year
    # WeekNum = format(ApptDate, "%U"),
    # Determine appointment day of week
    DOW = weekdays(ApptDate),
    # # Determine if patient's zip code is in NYS
    # NYZip = substr(`ZIP Code`, 1, 5) %in% ny_zips$zipcode
  )

# Crosswalk sites
sched_to_date <- left_join(sched_to_date, dept_mappings,
                           by = c("Department" = "Department"))

# Determine if there are any missing departments and notify user
unmapped_dept <- sched_to_date %>%
  filter(is.na(Site)) %>%
  select(Department) %>%
  distinct()

if(nrow(unmapped_dept) > 0) {
  stop(paste0("Check department mappings. Missing departments include ",
              paste(as.vector(unmapped_dept$Department), collapse = ", "),
              "."))
}

# Determine vaccine type (adult vs peds) administered based on department, provider, and age, as appropriate
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
        (`Vaccine Age Group` %in% "Both" &
           # Determine if provider/resource at MSBI is peds
           (`Provider/Resource` %in% c("DOSE 1 PEDIATRIC [1324684]",
                                       "DOSE 2 PEDIATRIC [1324685]") |
              (`Setting Grouper` %in% c("School Based Practice", "MSHS Practice") &
                 Age_Numeric < 12))),
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
    Site =
      # Roll up any Moderna administered at MSB for YWCA to MSH
      ifelse(Site %in% c("MSB") & Mfg %in% c("Moderna"), "MSH",
             Site),
    # Manually correct any J&J doses erroneously documented as dose 2 and change to dose 1
    Dose = ifelse(Mfg %in% c("J&J") & Dose %in% c("Dose 2"), "Dose 1", Dose))

# Determine vaccine week based on appointment date
sched_to_date <- left_join(sched_to_date,
                           all_dates,
                           by = c("ApptDate" = "Date",
                                  "VaxType" = "VaxType"))

# Summarize schedule data for Epic sites by key stratification
sched_summary <- sched_to_date %>%
  filter(Status2 %in% c("Arr") &
           ApptDate <= today) %>%
  group_by(Site,
           VaxType,
           Dose,
           VaccWeek,
           WeekDates) %>%
  summarize(Count = n(),
            .groups = "keep") %>%
  pivot_wider(names_from = Dose,
              values_from = Count) %>%
  ungroup() %>%
  mutate(AllDoses = select(., contains("Dose")) %>%
           rowSums(na.rm = TRUE)) %>%
  arrange(VaxType, VaccWeek, Site)

adult_sched_summary <- sched_summary %>%
  filter(VaxType %in% "Adult")

peds_sched_summary <- sched_summary %>%
  filter(VaxType %in% "Peds")

export_list <- list("Adult" = adult_sched_summary,
                    "Peds" = peds_sched_summary)

write_xlsx(export_list,
           path = paste0(user_directory,
                         "/Arr Visits Weekly Summary",
                         "/Vaccine Arrivals as of ",
                         format(Sys.Date(), "%Y-%m-%d"),
                         ".xlsx"))

