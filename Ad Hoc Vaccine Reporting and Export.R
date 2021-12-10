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
update_repo <- TRUE

# Determine whether or not to update walk-in analysis
update_walkins <- FALSE

# Import reference data for site and pod mappings
site_mappings <- read_excel(paste0(user_directory, "/ScheduleData/",
                                   "Automation Ref 2021-12-10.xlsx"),
                            sheet = "Site Mappings")
pod_mappings <- read_excel(paste0(user_directory, "/ScheduleData/",
                                  "Automation Ref 2021-12-10.xlsx"),
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

# NY zip codes
ny_zips <- search_state("NY")

# Import schedule repository and determine last date it was updated
sched_repo_file <- choose.files(default = paste0(user_directory,
                                                 "/R_Sched_AM_Repo/*.*"))

sched_repo <- readRDS(sched_repo_file)
sched_repo_update_date <- date(file.info(sched_repo_file)$ctime)

if (update_repo) {
  
  # Determine date range of reports to add to repository
  new_report_dates <- seq.Date(sched_repo_update_date, today - 1, by = 1)
  new_report_dates_pattern <- format(new_report_dates, "%Y%m%d")
  
  # Find list of reports
  file_list_vacc_sched <- list.files(
    path = paste0(user_directory,
                  "/Epic Auto Report Sched All Status"),
    pattern = paste0("^(Covid_vaccine_all_status_ZipCode_)",
                     new_report_dates_pattern,
                     ".*.csv",
                     collapse = "|"))
  
  # Import reports and add a column with report name and report date
  import_new_reports <- map_df(
    file_list_vacc_sched,
    .f = function(x){
      sched_data <- read.csv(
        paste0(user_directory,
               "/Epic Auto Report Sched All Status/",
               x),
        stringsAsFactors = FALSE,
        check.names = FALSE) %>%
        mutate(Filename = x,
               ReportDate = as.Date(str_extract(x, "[0-9]{8}"),
                                    "%Y%m%d"),
               DOB = as.Date(DOB, format = "%m/%d/%Y"),
               `Made Date` = as.Date(`Made Date`, format = "%m/%d/%y"),
               Date = as.Date(Date, "%m/%d/%Y"),
               `Appt Time` = as.POSIXct(paste("1899-12-31 ", `Appt Time`),
                                        tz = "", format = "%Y-%m-%d %H:%M %p"),
               `ZIP Code` = substr(`ZIP Code`, 2, nchar(`ZIP Code`)))})
  
  # Filter schedule data from new reports to only keep appointments in latest reports
  sched_data_new_reports <- import_new_reports %>%
    group_by(Date) %>%
    slice_max(ReportDate) %>%
    mutate(Filename = NULL,
           ReportDate = NULL) %>%
    ungroup()
  
  new_reports_dates <- sort(unique(sched_data_new_reports$Date))
  
  # Remove vaccine appoints from any overlapping dates in repository and new reports
  sched_repo <- sched_repo %>%
    filter(!(Date >= new_reports_dates[1]))
  
  # Bind schedule repository and schedule from new reports
  sched_repo <- rbind(sched_repo, sched_data_new_reports)
  
  # Save new repository
  saveRDS(sched_repo, paste0(user_directory, "/R_Sched_AM_Repo/",
                             "Auto Epic Report Sched ",
                             format(min(sched_repo$Date), "%m%d%y"), " to ",
                             format(max(sched_repo$Date), "%m%d%y"),
                             " on ", format(Sys.time(), "%m%d%y %H%M"), ".rds"))

} else {
  
  sched_repo <- sched_repo
  
}


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
                                                  "DOSE 2 PEDIATRIC [1324685]"),
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
        regex("[a-z]", ignore_case = TRUE)))

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

if(sum(is.na(sched_sites)) > 0) {
  stop("Check site and department mappings.")
}

# Assume any pods without a mapping are patient pods
sched_to_date[is.na(sched_to_date$`Pod Type`), "Pod Type"] <- "Patient"

# Summarize schedule data by key stratification
sched_summary <- sched_to_date %>%
  group_by(Site, `Pod Type`, VaxType, Dose, ApptDate, NYZip, Status2) %>%
  summarize(Count = n())

# Summarize arrivals by patient type (employees vs. non-employees)
arr_pt_type_summary <- sched_to_date %>%
  filter(Status2 %in% c("Arr")) %>%
  mutate(PatientType = ifelse(MSHSEmployee, "Employee", "Non-Employee")) %>%
  group_by(Site, ApptDate, Dose, PatientType) %>%
  summarize(Count = n()) %>%
  ungroup

# Summarize schedule breakdown for next 2 weeks and export --------------------
# Format data
sched_breakdown <- sched_to_date %>%
  filter(Status2 %in% c("Arr", "Sch") &
           ApptDate >= (today - 1) &
           ApptDate <= (today + 14)) %>%
  group_by(Dose, Site, VaxType, Status2, ApptDate) %>%
  # group_by(Dose, `Pod Type`, Site, Status2, ApptDate) %>%
  summarize(Count = n(), .groups = "keep") %>%
  ungroup()

sched_breakdown <- sched_breakdown[order(sched_breakdown$Dose,
                                         # sched_breakdown$`Pod Type`,
                                         sched_breakdown$Site,
                                         sched_breakdown$VaxType), ]

# Create template dataframe of all sites, dose, pod, and vaccine type combinations to
# ensure all are included even if there are no scheduled appointments
# Repeat dose types
doses_templ <- data.frame(
  "Dose" = rep(c(1:3),
               length(all_sites)))

doses_templ <- doses_templ %>%
  arrange(Dose)

# Repeat pod types
pods_templ <- data.frame(
  "PodType" = rep(pod_type[1:2],
                  length(all_sites)),
  stringsAsFactors = FALSE)

pods_templ <- pods_templ %>%
  arrange(PodType)

# Repeat dates for next 2 weeks
dates_14days <- seq(from = today - 1, to = today + 14, by = 1)

dates_templ <- data.frame(
  "Date" = rep(dates_14days,
               length(all_sites)),
  stringsAsFactors = FALSE)

dates_templ <- dates_templ %>%
  arrange(Date)

# Repeat vaccine manufacturers
mfg_templ <- data.frame(
  "Mfg" = rep(mfg,
              length(all_sites)),
  stringsAsFactors = FALSE)

mfg_templ <- mfg_templ %>%
  mutate(Mfg = factor(Mfg, levels = mfg, ordered = TRUE)) %>%
  arrange(Mfg) %>%
  mutate(Mfg = as.character(Mfg))

# Repeat vaccine types
vax_type_templ <- data.frame(
  "VaxType" = rep(vax_types,
                  length(all_sites)),
  stringsAsFactors = FALSE)

vax_type_templ <- vax_type_templ %>%
  arrange(VaxType)

# Combine site and doses
site_doses_templ <- data.frame(
  "Site" = rep(all_sites, 3),
  "Dose" = doses_templ,
  stringsAsFactors = FALSE)

# Combine site and pods
site_pods_templ <- data.frame(
  "Site" = rep(all_sites, 2),
  "PodType" = pods_templ,
  stringsAsFactors = FALSE)

# Combine site and dates
site_dates_templ <- data.frame(
  "Site" = rep(all_sites, length(dates_14days)),
  "Date" = dates_templ,
  stringsAsFactors = FALSE)

# Combine site and manufacturer
site_mfg_templ <- data.frame(
  "Site" = rep(all_sites, length(mfg)),
  "Mfg" = mfg_templ)

# Combine site and vaccine type and keep only relevant combinations
site_vax_templ <- data.frame(
  "Site" = rep(all_sites, length(vax_types)),
  "VaxType" = vax_type_templ
)

site_vax_templ <- site_vax_templ %>%
  filter(VaxType %in% c("Adult") |
           (VaxType %in% c("Peds") & Site %in% c("MSH", "MSBI"))
  )

# Combine site, dose, and pods
site_dose_pod_templ <- left_join(site_doses_templ,
                                 site_pods_templ,
                                 by = c("Site" = "Site"))

# # Combine site, dose, pods, and dates
# site_dose_pod_date_templ <- left_join(site_dates_templ,
#                                       site_dose_pod_templ,
#                                       by = c("Site" = "Site"))
# 
# site_dose_pod_date_templ <- site_dose_pod_date_templ %>%
#   # select(Dose, PodType, Site) %>%
#   mutate(Site = factor(Site, levels = all_sites, ordered = TRUE)) %>%
#   arrange(Dose, PodType, Site, Date)

# Combine site, dose, and dates
site_dose_date_templ <- left_join(site_doses_templ,
                                  site_dates_templ,
                                  by = c("Site" = "Site"))

# Combine site, dose, and manufacturer
site_dose_mfg_templ <- left_join(site_doses_templ,
                                 site_mfg_templ,
                                 by = c("Site" = "Site"))

# Combine site, vaccine type, and dose
site_vax_dose_templ <- left_join(site_vax_templ,
                                 site_doses_templ,
                                 by = c("Site" = "Site"))

# Remove dose 3
site_vax_dose_templ <- site_vax_dose_templ %>%
  filter(!(VaxType %in% c("Peds") & Dose %in% c(3)))

# Combine site, dose, manufacturer, and dates
site_dose_mfg_date_templ <- left_join(site_dates_templ,
                                      site_dose_mfg_templ,
                                      by = c("Site" = "Site"))

# Manually remove J&J dose 2 and 3 from template
site_dose_mfg_date_templ <- site_dose_mfg_date_templ %>%
  filter(!(Mfg == "J&J" & Dose > 1))

# Combine site, vaccine type, dose, manufacturer, and dates
site_vax_dose_mfg_date_templ <- left_join(site_vax_templ,
                                          site_dose_mfg_date_templ,
                                          by = c("Site" = "Site"))

# Remove Dose 3, Moderna, and J&J Peds vaccines
site_vax_dose_mfg_date_templ <- site_vax_dose_mfg_date_templ %>%
  filter(!(VaxType %in% c("Peds") &
             (Mfg %in% c("J&J", "Moderna") | Dose == 3)))

# Combine site, vaccine type, dose, and dates
site_vax_dose_date_templ <- left_join(site_vax_templ,
                                      site_dose_date_templ,
                                      by = c("Site" = "Site"))

# Remove booster shots from peds vaccine type
site_vax_dose_date_templ <- site_vax_dose_date_templ %>%
  filter(!(VaxType == "Peds" & Dose == 3))

# Merge template dataframe with schedule breakdown data
sched_breakdown <- left_join(site_vax_dose_date_templ,
                             # site_dose_date_templ,
                             sched_breakdown,
                             by = c("Dose" = "Dose",
                                    # "PodType" = "Pod Type",
                                    "VaxType" = "VaxType",
                                    "Site" = "Site",
                                    "Date" = "ApptDate"))

# Replace NA in schedule breakdown and pivot wider for desired output
sched_breakdown <- sched_breakdown %>%
  mutate(Status2 = ifelse(is.na(Status2) & Date < today, "Arr",
                          ifelse(is.na(Status2) & Date >= today, "Sch",
                                 Status2)),
         Count = ifelse(is.na(Count), 0, Count))

sched_breakdown_cast <- sched_breakdown %>%
  pivot_wider(id_cols = c(Dose, Site, VaxType),
              names_from = c(Status2, Date),
              values_from = Count) %>%
  arrange(VaxType, Dose, Site)

# Summarize schedule breakdown for next two weeks for Pfizer only -------------
# Assumptions:
# 1. Any scheduled dose 1 appointments will receive Pfizer
# 2. Only scheduled dose 2 appointments with an arrived Moderna dose 1
# appointment will receive Moderna. Any other scheduled dose 2 appointments will
# receive Pfizer
# 
# First, subset schedule for city sites
city_sites_sched <- sched_to_date %>%
  filter(Site %in% city_sites)

# Determine MRNs that received Moderna for first dose
dose1_moderna <- city_sites_sched %>%
  filter(Status2 %in% c("Arr") & Mfg == "Moderna" & Dose == 1) %>%
  select(MRN)

city_sites_sched <- city_sites_sched %>%
  #Determine expected manufacturer for scheduled appointment based on above
  # assumptions
  mutate(ExpSchMfg = ifelse(Status2 != "Sch", NA,
                            # Assume all MSVD scheduled appointments receive J&J
                            ifelse(Site == "MSVD", "J&J",
                                   # Assume all peds and all scheduled Dose 1 and Dose 3 appointments receive Pfizer
                                   ifelse(VaxType %in% c("Peds") |
                                            Dose %in% c(1, 3), "Pfizer",
                                          # Assume scheduled Dose 2 only receive Moderna if Dose 1 was Moderna
                                          ifelse(MRN %in% dose1_moderna$MRN,
                                                 "Moderna", "Pfizer")))),
         MfgRollUp = ifelse(is.na(Mfg) & is.na(ExpSchMfg), NA,
                            ifelse(is.na(ExpSchMfg), Mfg, ExpSchMfg)))

mfg_sched_breakdown <- city_sites_sched %>%
  filter(!is.na(MfgRollUp) &
           ApptDate >= (today - 1) &
           ApptDate <= (today + 14)) %>%
  group_by(Dose, VaxType, MfgRollUp, Site, Status2, ApptDate) %>%
  summarize(Count = n(), .groups = "keep") %>%
  ungroup()

mfg_templ <- site_vax_dose_mfg_date_templ %>%
  # site_dose_mfg_date_templ %>%
  filter(Site %in% city_sites)

mfg_sched_breakdown <- left_join(mfg_templ,
                                 mfg_sched_breakdown,
                                 by = c("Dose" = "Dose",
                                        "VaxType" = "VaxType",
                                        "Mfg" = "MfgRollUp",
                                        "Site" = "Site",
                                        "Date" = "ApptDate"))

# Replace NA in schedule breakdown and pivot wider for desired output
mfg_sched_breakdown_cast <- mfg_sched_breakdown %>%
  mutate(Status2 = ifelse(is.na(Status2) & Date < today, "Arr",
                          ifelse(is.na(Status2) & Date >= today, "Sch",
                                 Status2)),
         Count = replace_na(Count, 0)) %>%
  pivot_wider(id_cols = c(Dose, VaxType, Mfg, Site),
              names_from = c(Status2, Date),
              values_from = Count) %>%
  mutate(Site = factor(Site, levels = city_sites, ordered = TRUE),
         Mfg = factor(Mfg, levels = mfg, ordered = TRUE)) %>%
  arrange(VaxType, Dose, Mfg, Site)

# Create schedule breakdown for Pfizer only
pfizer_sched_breakdown <- city_sites_sched %>%
  filter((Mfg == "Pfizer" | ExpSchMfg == "Pfizer") &
           ApptDate >= (today - 1) & ApptDate <= (today + 14)) %>%
  group_by(Dose, VaxType, Site, Status2, ApptDate) %>%
  # group_by(Dose, `Pod Type`, Site, Status2, ApptDate) %>%
  summarize(Count = n(), .groups = "keep") %>%
  ungroup()

# Subset site, dose, pod type, date template for sites that have administered Pfizer
pfizer_templ <- site_vax_dose_date_templ %>%
  filter(Site %in% unique(pfizer_sched_breakdown$Site))

# Merge template with schedule data
pfizer_sched_breakdown <- left_join(pfizer_templ,
                                    pfizer_sched_breakdown,
                                    by = c("Dose" = "Dose",
                                           "VaxType" = "VaxType",
                                           "Site" = "Site",
                                           "Date" = "ApptDate"))

# Replace NA in schedule breakdown and pivot wider for desired output
pfizer_sched_breakdown <- pfizer_sched_breakdown %>%
  mutate(Status2 = ifelse(is.na(Status2) & Date < today, "Arr",
                          ifelse(is.na(Status2) & Date >= today, "Sch",
                                 Status2)),
         Count = ifelse(is.na(Count), 0, Count))

pfizer_sched_breakdown_cast <- pfizer_sched_breakdown %>%
  pivot_wider(id_cols = c(Dose, Site, VaxType),
              names_from = c(Status2, Date),
              values_from = Count) %>%
  arrange(VaxType, Dose, Site)

# Summarize schedule for next vaccine inventory cycle for target analysis -----
# Determine dates in this inventory cycle based on today's date and
# Tuesday of the next week
sched_inv_cycle_dates <- seq(today,
                             floor_date(today,
                                        unit = "week",
                                        week_start = 7) + 16,
                             by = 1)

# Create dataframe of repeated pod types
inv_cycle_pods_site <- data.frame(
  "PodType" = rep(pod_type[1:2], length(all_sites)),
  stringsAsFactors = FALSE)

inv_cycle_pods_site <- inv_cycle_pods_site %>%
  arrange(PodType)

# Create dataframe of repeated doses
inv_cycle_dose_site <- data.frame(
  "Dose" = rep(c(1:3), length(all_sites)),
  stringsAsFactors = FALSE)

inv_cycle_dose_site <- inv_cycle_dose_site %>%
  arrange(Dose)

# Create dataframe of repeated dates
inv_cycle_dates_site <- data.frame(
  "Date" = rep(sched_inv_cycle_dates, length(all_sites)),
  stringsAsFactors = FALSE)

inv_cycle_dates_site <- inv_cycle_dates_site %>%
  arrange(Date)

# Add in sites to repeated dataframe for pod type, dose, and date
inv_cycle_pods_site_templ <- data.frame(
  "Site" = rep(all_sites, 2),
  inv_cycle_pods_site)

inv_cycle_dose_site_templ <- data.frame(
  "Site" = rep(all_sites, 3),
  inv_cycle_dose_site)

inv_cycle_dates_site_templ <- data.frame(
  "Site" = rep(all_sites, length(sched_inv_cycle_dates)),
  inv_cycle_dates_site)

# Combine all templates into single template dataframe
# inv_cycle_site_templ <- left_join(
#   left_join(
#     inv_cycle_pods_site_templ,
#     inv_cycle_dose_site_templ,
#     by = c("Site" = "Site")),
#   inv_cycle_dates_site_templ,
#   by = c("Site" = "Site"))
inv_cycle_site_templ <- left_join(inv_cycle_dose_site_templ,
                                  inv_cycle_dates_site_templ,
                                  by = c("Site" = "Site"))

# Summarize schedule data for each site through next 2 inventory cycles
sched_inv_cycle_site <- sched_to_date %>%
  filter(ApptDate %in% sched_inv_cycle_dates &
           Status2 == "Sch") %>%
  group_by(Site, Dose, ApptDate) %>%
  # group_by(Site, Dose, `Pod Type`, ApptDate) %>%
  summarize(Count = n(), .groups = "keep") %>%
  ungroup()

# Combine with template
sched_inv_cycle_site <- left_join(inv_cycle_site_templ,
                                  sched_inv_cycle_site,
                                  by = c("Site" = "Site",
                                         "Date" = "ApptDate",
                                         "Dose" = "Dose"))

# sched_inv_cycle_site <- sched_inv_cycle_site %>%
#   mutate(Count = ifelse(is.na(Count), 0, Count),
#          Site = factor(Site, levels = all_sites, ordered = TRUE)) %>%
#   pivot_wider(id_cols = c(Site, Dose, Date),
#               names_from = PodType,
#               values_from = Count) %>%
#   mutate(Total = Employee + Patient) %>%
#   pivot_longer(cols = c(Employee, Patient, Total),
#                names_to = "PodType",
#                values_to = "Count") %>%
#   arrange(Site, Date, Dose, desc(PodType)) %>%
#   pivot_wider(id_cols = c(Site, Date),
#               names_from = c(Dose, PodType),
#               values_from = Count,
#               names_sep = "_")

sched_inv_cycle_site <- sched_inv_cycle_site %>%
  mutate(Count = ifelse(is.na(Count), 0, Count),
         Site = factor(Site, levels = all_sites, ordered = TRUE)) %>%
  # pivot_wider(id_cols = c(Site, Dose, Date),
  #             names_from = PodType,
  #             values_from = Count) %>%
  # mutate(Total = Employee + Patient) %>%
  # pivot_longer(cols = c(Employee, Patient, Total),
  #              names_to = "PodType",
  #              values_to = "Count") %>%
  arrange(Site, Date, Dose) %>%
  pivot_wider(id_cols = c(Site, Date),
              names_from = c(Dose),
              names_prefix = "Dose ",
              values_from = Count)

# Determine schedule across system within this inventory cycle's date range
# Create a dataframe of repeated pod types
inv_cycle_pods_sys <- data.frame(
  "PodType" = rep(pod_type[1:2], length(sched_inv_cycle_dates)),
  stringsAsFactors = FALSE)

inv_cycle_pods_sys <- inv_cycle_pods_sys %>%
  arrange(PodType)

# Create a dataframe of repeated doses
inv_cycle_dose_sys <- data.frame(
  "Dose" = rep(c(1:3), length(sched_inv_cycle_dates)),
  stringsAsFactors = FALSE)

inv_cycle_dose_sys <- inv_cycle_dose_sys %>%
  arrange(inv_cycle_dose_sys)

# Create template with pod types and doses for dates of interest
inv_cycle_pods_date_sys_templ <- data.frame(
  "Date" = rep(sched_inv_cycle_dates, 2),
  inv_cycle_pods_sys)

inv_cycle_dose_date_sys_templ <- data.frame(
  "Date" = rep(sched_inv_cycle_dates, 3),
  inv_cycle_dose_sys)

inv_cycle_sys_templ <- inv_cycle_dose_date_sys_templ
# left_join(inv_cycle_pods_date_sys_templ,
#                                inv_cycle_dose_date_sys_templ,
#                                by = c("Date" = "Date"))

sched_inv_cycle_sys <- sched_to_date %>%
  filter(ApptDate %in% sched_inv_cycle_dates &
           Status2 == "Sch" &
           Site %in% city_sites) %>%
  group_by(Dose, ApptDate) %>%
  # group_by(Dose, `Pod Type`, ApptDate) %>%
  summarize(Count = n(), .groups = "keep") %>%
  ungroup()

# Combine with template
sched_inv_cycle_sys <- left_join(inv_cycle_sys_templ,
                                 sched_inv_cycle_sys,
                                 by = c("Date" = "ApptDate",
                                        "Dose" = "Dose"))

sched_inv_cycle_sys <- sched_inv_cycle_sys %>%
  mutate(Count = ifelse(is.na(Count), 0, Count),
         DOW = wday(Date, label = TRUE, abbr = TRUE)) %>%
  arrange(Date, Dose) %>%
  pivot_wider(id_cols = c(DOW, Date),
              names_from = c(Dose),
              names_prefix = "Dose ",
              values_from = Count)

# sched_inv_cycle_sys <- sched_inv_cycle_sys %>%
#   mutate(Count = ifelse(is.na(Count), 0, Count)) %>%
#   pivot_wider(id_cols = c(Date, Dose),
#               names_from = PodType,
#               values_from = Count) %>%
#   mutate(Total = Employee + Patient,
#          DOW = wday(Date, label = TRUE, abbr = TRUE)) %>%
#   pivot_longer(cols = c(Employee, Patient, Total),
#                names_to = "PodType",
#                values_to = "Count") %>%
#   arrange(Date, Dose, desc(PodType)) %>%
#   pivot_wider(id_cols = c(DOW, Date),
#               names_from = c(Dose, PodType),
#               values_from = Count,
#               names_sep = "_")

# Summarize doses administered to date --------------------------------
# Summarize administered doses prior to today and stratify by site and pod type
admin_pods_site <- data.frame(
  "PodType" = rep(pod_type, length(city_sites)),
  stringsAsFactors = FALSE)

admin_pods_site <- admin_pods_site %>%
  arrange(PodType)

# Create dataframe of repeated doses
admin_dose_site <- data.frame(
  "Dose" = rep(c(1:3), length(city_sites)),
  stringsAsFactors = FALSE)

admin_dose_site <- admin_dose_site %>%
  arrange(Dose)

# Add in sites to repeated dataframe for pod type, dose, and date
admin_pods_site_templ <- data.frame(
  "Site" = rep(city_sites, length(pod_type)),
  admin_pods_site)

admin_dose_site_templ <- data.frame(
  "Site" = rep(city_sites, 3),
  admin_dose_site)

admin_pods_dose_site_templ <- left_join(admin_pods_site_templ,
                                        admin_dose_site_templ,
                                        by = c("Site" = "Site"))

# Summarize administered doses prior to today and stratify by site and pod type
admin_doses_site_pod <- sched_to_date %>%
  filter(Status2 == "Arr" &
           Site %in% city_sites) %>%
  group_by(Site, `Pod Type`, Dose) %>%
  summarize(Count = n(),
            .groups = "keep") %>%
  pivot_wider(names_from = `Pod Type`,
              values_from = Count) %>%
  mutate(All = sum(Employee, Patient, na.rm = TRUE)) %>%
  pivot_longer(cols = c(Employee, Patient, All),
               names_to = "PodType",
               values_to = "Count")

admin_doses_site_pod <- left_join(admin_pods_dose_site_templ,
                                  admin_doses_site_pod,
                                  by = c("Site" = "Site",
                                         "Dose" = "Dose",
                                         "PodType" = "PodType"))

admin_doses_site_pod <- admin_doses_site_pod %>%
  mutate(PodType = factor(PodType, levels = pod_type, ordered = TRUE),
         Site = factor(Site, levels = city_sites, ordered = TRUE),
         Count = replace_na(Count, 0)) %>%
  arrange(Site, Dose, PodType) %>%
  mutate(PodType = paste0(PodType, "Pods")) %>%
  pivot_wider(names_from = c(Dose, PodType),
              values_from = Count,
              names_prefix = "Dose",
              names_sep = "_") %>%
  adorn_totals(where = "row", na.rm = TRUE, name = "MSHS")

# Stratify administered doses to date by manufacturer ------------------------
# Create daily summary of administered doses by manufacturer
admin_mfg_summary <- sched_to_date %>%
  filter(Status2 == "Arr" &
           Site %in% city_sites) %>%
  group_by(Site, VaxType, Mfg, Dose, ApptDate, Status2) %>%
  summarize(Count = n(), .groups = "keep")

# Summarize doses administered to date by vaccine type (adult vs peds) ---------
admin_doses_site_vaxtype <- sched_to_date %>%
  filter(Status2 == "Arr") %>%
  group_by(Site, VaxType, Dose) %>%
  summarize(Count = n(),
            .groups = "keep")

admin_doses_site_vaxtype <- left_join(site_vax_dose_templ,
                                      admin_doses_site_vaxtype,
                                      by = c("Site" = "Site",
                                             "VaxType" = "VaxType",
                                             "Dose" = "Dose"))

admin_doses_site_vaxtype_cast <- admin_doses_site_vaxtype %>%
  mutate(Count = replace_na(Count, 0),
         Dose = paste0("Dose", Dose)) %>%
  pivot_wider(names_from = c(VaxType, Dose),
              values_from = Count,
              names_sep = "_") %>%
  adorn_totals(where = "row", na.rm = TRUE, name = "MSHS") %>%
  pivot_longer(cols = contains("_"),
               names_to = c("VaxType", "Dose"),
               names_sep = "_",
               values_to = "Count") %>%
  mutate(Count = replace_na(Count, 0)) %>%
  pivot_wider(names_from = Dose,
              values_from = Count) %>%
  adorn_totals(where = "col",
               na.rm = TRUE,
               name = "AllDoses") %>%
  pivot_longer(cols = contains("Dose"),
               names_to = "Dose",
               values_to = "Count") %>%
  mutate(Count = replace_na(Count, 0)) %>%
  filter(!(VaxType %in% c("Peds") & Dose %in% c("Dose3"))) %>%
  pivot_wider(names_from = c(VaxType, Dose),
              values_from = Count,
              names_sep = "_",
              names_sort = FALSE)

# # Create summary of administered doses prior to yesterday and yesterday
# admin_prior_yest_site <- sched_to_date %>%
#   filter(Status2 == "Arr" &
#            ApptDate <= (today - 2) &
#            Site %in% city_sites) %>%
#   group_by(Site, Mfg, Dose) %>%
#   summarize(DateRange = paste0("Admin 12/15/20-",
#                                format(today - 2, "%m/%d/%y")),
#             Count = n(), .groups = "keep")
# 
# admin_yest_site <- sched_to_date %>%
#   filter(Status2 == "Arr" &
#            ApptDate == (today - 1) &
#            Site %in% city_sites) %>%
#   group_by(Site, Mfg, Dose) %>%
#   summarize(DateRange = paste0("Admin ", format(today - 1, "%m/%d/%y")),
#             Count = n(),
#             .groups = "keep")
# 
# admin_to_date_mfg <- rbind(admin_prior_yest_site, admin_yest_site)
# 
# admin_to_date_mfg <- admin_to_date_mfg %>%
#   ungroup() %>%
#   arrange(Site, Dose, desc(Mfg), desc(DateRange))
# 
# admin_to_date_mfg <-
#   admin_to_date_mfg[, c("Dose", "Mfg", "Site", "DateRange", "Count")]
# 
# admin_to_date_mfg <- admin_to_date_mfg %>%
#   mutate(DateRange = factor(DateRange, level = unique(DateRange),
#                             ordered = TRUE))
# 
# admin_to_date_mfg_cast <- dcast(admin_to_date_mfg,
#                                 Dose + Mfg + Site ~ DateRange,
#                                 value.var = "Count")
# 
# # Create template dataframe of sites, doses, and manufacturers to ensure all
# # combinations are included
# rep_doses <- data.frame("Dose" = rep(c(1:3),
#                                      length(unique(city_sites_sched$Site))))
# rep_doses <- rep_doses %>%
#   arrange(Dose)
# 
# rep_mfg <- data.frame("Mfg" = rep(mfg,
#                                   length(unique(city_sites_sched$Site))),
#                       stringsAsFactors = FALSE)
# 
# rep_mfg <- rep_mfg %>%
#   arrange(desc(Mfg))
# 
# rep_sites_mfg <- data.frame("Site" = rep(unique(city_sites_sched$Site),
#                                          length(mfg)), "Mfg" = rep_mfg,
#                             stringsAsFactors = FALSE)
# 
# rep_sites_doses <- data.frame("Site" = rep(unique(city_sites_sched$Site), 3),
#                               "Dose" = rep_doses,
#                               stringsAsFactors = FALSE)
# 
# # Combine template with administration data
# city_sites_doses_mfg <- left_join(rep_sites_mfg, rep_sites_doses,
#                                   by = c("Site" = "Site"))
# 
# city_sites_doses_mfg <- city_sites_doses_mfg %>%
#   mutate(Site = factor(Site, levels = city_sites, ordered = TRUE)) %>%
#   arrange(Dose, desc(Mfg), Site)
# 
# city_sites_doses_mfg <- city_sites_doses_mfg[, c("Dose", "Mfg", "Site")]
# 
# admin_to_date_mfg_export <- left_join(city_sites_doses_mfg, admin_to_date_mfg_cast,
#                                       by = c("Dose" = "Dose",
#                                              "Mfg" = "Mfg",
#                                              "Site" = "Site"))

# Export key data tables to Excel for reporting -------------------------------
# Export schedule summary, schedule breakdown, and cumulative administered
# doses to excel file
export_list <- list("SchedSummary" = sched_summary,
                    "Arrivals_PatientType" = arr_pt_type_summary,
                    "SchedBreakdown" = sched_breakdown_cast,
                    "Mfg_SchedBreakdown" = mfg_sched_breakdown_cast,
                    "Pfizer_SchedBreakdown" = pfizer_sched_breakdown_cast,
                    # "Sched_Targets_7Days" = sched_7days_format_jm,
                    "Sched_InvCycle_Site" = sched_inv_cycle_site,
                    "Sched_InvCycle_Sys" = sched_inv_cycle_sys,
                    "CumDosesAdmin_PodType" = admin_doses_site_pod,
                    "CumDosesAdmin_VaxType" = admin_doses_site_vaxtype_cast,
                    "AdminMfgSummary" = admin_mfg_summary)
                    # "DosesAdminMfg" = admin_to_date_mfg_export)

write_xlsx(export_list, path = paste0(user_directory,
                                      "/R_Sched_Analysis_AM_Export/",
                                      "Auto Epic Report Sched Data Export ",
                                      format(Sys.time(), "%m%d%y %H%M"),
                                      ".xlsx"))

# rmarkdown::render(input = "Empl_Arrival_Tracking.Rmd",
#                   output_file = paste0(user_directory,
#                                        "/Daily Dashboard Drafts",
#                                        "/Employee Arrival Tracking Tool ",
#                                        format(Sys.Date(), "%Y-%m-%d")))
