# Code for creating repository of COVID vaccine administration -----------------

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
update_repo <- FALSE

# Import reference data for site and pod mappings
site_mappings <- read_excel(paste0(user_directory, "/ScheduleData/",
                                   "Automation Ref 2022-03-21.xlsx"),
                            sheet = "Site Mappings")
pod_mappings <- read_excel(paste0(user_directory, "/ScheduleData/",
                                  "Automation Ref 2022-03-21.xlsx"),
                           sheet = "Pod Mappings Simple")

dept_grouping <- read_excel(paste0(user_directory,
                                   "/ScheduleData/",
                                   "Automation Ref 2022-03-21.xlsx"),
                                   sheet = "Dept Mapping and Grouping")

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

# Dept grouper
dept_group <- c("Hospital POD", "School Based Practice", "MSHS Practice")

# # Pediatric school based practices
# peds_school_practice_dept <- c("HOSP SBH PS 38",
#                                "HOSP SBH PS 83/ PS 182",
#                                "HOSP SBH PS 108",
#                                "HOSP SBH TPEC",
#                                "HOSP PEDS SCH HEALTH")


# Import schedule repository and determine last date it was updated
sched_repo_file <- choose.files(default = paste0(user_directory,
                                                 "/R_Sched_AM_Repo/*.*"))

sched_repo <- readRDS(sched_repo_file)

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
    # # Determine whether the vaccine is an adult or pediatric dose
    # VaxType = ifelse(Department %in% c("1468 MADISON PEDIATRIC VACCINE POD",
    #                                    "1468 MADISON HOSP PEDS GEN") |
    #                    `Provider/Resource` %in% c("DOSE 1 PEDIATRIC [1324684]",
    #                                               "DOSE 2 PEDIATRIC [1324685]") |
    #                    Department %in% peds_school_practice_dept,
    #                  "Peds", "Adult"),
    # # Determine manufacturer based on immunization field and vaccine type fields
    # Mfg = #Keep manufacturer as NA if the appointment hasn't been arrived
    #   ifelse(Status2 != "Arr", NA,
    #          # Any blank immunizations and any peds assumed to be Pfizer
    #          ifelse(is.na(Immunizations) |
    #                   VaxType %in% c("Peds"), "Pfizer",
    #                 # Any immunizations with Moderna in text classified as Moderna
    #                 ifelse(str_detect(Immunizations, "Moderna"), "Moderna",
    #                        # Any visiting docs or Johnson and Johnson
    #                        # immunizations classified as J&J
    #                        ifelse(str_detect(Department, "VISITING DOCS") |
    #                                 str_detect(Immunizations,
    #                                            "Johnson and Johnson"),
    #                               "J&J", "Pfizer")))),
    # Determine week number of year
    WeekNum = format(ApptDate, "%U"),
    # Determine appointment day of week
    DOW = weekdays(ApptDate),
    # # Determine if patient's zip code is in NYS
    # NYZip = substr(`ZIP Code`, 1, 5) %in% ny_zips$zipcode,
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
        regex("[a-z]", ignore_case = TRUE)))

# Crosswalk sites and department grouper
sched_to_date <- left_join(sched_to_date, dept_grouping,
                           by = c("Department" = "Department"))

unmapped_dept <- sched_to_date %>%
  filter(is.na(Site)) %>%
  select(Department) %>%
  distinct()

if(nrow(unmapped_dept) > 0) {
  stop(paste0("Check department mappings. Missing departments include ",
             paste(as.vector(unmapped_dept$Department), collapse = ", "),
             "."))
}

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
      AgeGroup %in% "Peds" |
        (AgeGroup %in% "Both" &
           # Determine if provider/resource at MSBI is peds
           (`Provider/Resource` %in% c("DOSE 1 PEDIATRIC [1324684]",
                                       "DOSE 2 PEDIATRIC [1324685]") |
              (Grouping %in% c("School Based Practice", "MSHS Practice") &
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
    # Roll up any Moderna administered at MSB for YWCA to MSH
    Site = ifelse(Site %in% c("MSB") & Mfg %in% c("Moderna"), "MSH", Site),
    # Manually correct any J&J doses erroneously documented as dose 2 and change to dose 1
    Dose = ifelse(Mfg %in% c("J&J") & Dose %in% c(2), 1, Dose),
    Site_Subgroup = ifelse(Grouping %in% "School Based Practice", Department,
                           Site))

test_df <- sched_to_date %>%
  select(Department, `Provider/Resource`, Site, Grouping,
         Age, AgeGroup, Age_Numeric, VaxType) %>%
  distinct()

# # Manually update discrepancies in site/mfg/dose
# sched_to_date <- sched_to_date %>%
#   mutate(# Roll up any Moderna administered at MSB for YWCA to MSH
#     Site = ifelse(Site %in% c("MSB") & Mfg %in% c("Moderna"), "MSH", Site),
#     # Manually correct any J&J doses erroneously documented as dose 2 and change to dose 1
#     Dose = ifelse(Mfg %in% c("J&J") & Dose %in% c(2), 1, Dose))
# 
# # Crosswalk pod type (employee vs patient)
# sched_to_date <- left_join(sched_to_date,
#                            unique(pod_mappings[, c("Provider", "Pod Type")]),
#                            by = c("Provider/Resource" = "Provider"))
# 
# # Assume any pods without a mapping are patient pods
# sched_to_date[is.na(sched_to_date$`Pod Type`), "Pod Type"] <- "Patient"

# Summarize data for dashboard visualizations
summary_all_doses <- sched_to_date %>%
  filter(Status2 %in% "Arr") %>%
  group_by(Grouping, Site_Subgroup, VaxType) %>%
  summarize(Count = n()) %>%
  rename(Setting = Grouping,
         Site = Site_Subgroup) %>%
  arrange(Setting, VaxType, -Count) %>%
  pivot_wider(names_from = VaxType,
              values_from = Count) %>%
  ungroup() %>%
  split(.$Setting) %>%
    # relocate(Setting, .after = Site) %>%
    map_df(., function(x) {
      adorn_totals(x, where = "row",
                        name = unique(x$Setting),
                        fill = "MSHS Total")})
  



# MSH Summary
msh_summary <- sched_to_date %>%
  filter(Site %in% c("MSH") &
           Status2 %in% c("Arr")) %>%
  group_by(Site, Department, Dose, VaxType) %>%
  summarize(Count = n()) %>%
  mutate(Dose = paste("Dose", Dose)) %>%
  arrange(VaxType, Dose, Department) %>%
  pivot_wider(names_from = Dose,
              values_from = Count) %>%
  adorn_totals(where = "col", na.rm = TRUE, name = "All Doses")

# Export key data tables to Excel for reporting -------------------------------
# Export schedule summary, schedule breakdown, and cumulative administered
# doses to excel file
export_list <- list("MSH Summary" = msh_summary)


write_xlsx(export_list, path = paste0(user_directory,
                                      "/AdHoc/",
                                      "MSH Admin Dose Summary for BC ",
                                      format(Sys.Date(), "%Y-%m-%d"),
                                      ".xlsx"))
