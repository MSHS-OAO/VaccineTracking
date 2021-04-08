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
initial_run <- FALSE
update_repo <- TRUE

# Determine whether or not to update walk-in analysis
update_walkins <- FALSE

# Import reference data for site and pod mappings
site_mappings <- read_excel(paste0(user_directory, "/ScheduleData/",
                                   "Automation Ref 2021-04-08.xlsx"),
                            sheet = "Site Mappings")
pod_mappings <- read_excel(paste0(user_directory, "/ScheduleData/",
                                  "Automation Ref 2021-04-08.xlsx"),
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

# NY zip codes
ny_zips <- search_state("NY")

# Import raw data from Epic
if (update_repo) {
  if (initial_run) {
    raw_df <- read_excel(choose.files(default = paste0(user_directory,
                                                       "/ScheduleData/*.*"),
                                      caption = "Select initial Epic schedule"),
                         col_names = TRUE, na = c("", "NA"))
  } else {
    # sched_repo <- read_excel(choose.files(default = user_path,
    # caption = "Select schedule repository"),
    #                       col_names = TRUE, na = c("", "NA"))
    sched_repo <- readRDS(choose.files(default = paste0(user_directory,
                                                        "/R_Sched_AM_Repo/*.*"),
                                       caption = "Select schedule repository"))

    # # Convert appointment date in schedule repository from posixct to date
    # sched_repo <- sched_repo %>%
    #   mutate(ApptDate = date(ApptDate))

    raw_df <- read.csv(choose.files(
      default =
        paste0(user_directory, "/Epic Auto Report Sched All Status/*.*"),
      caption = "Select current Epic schedule"),
      stringsAsFactors = FALSE,
      check.names = FALSE)
  }

  new_sched <- raw_df

  # Remove test patient
  new_sched <- new_sched[new_sched$Patient != "Patient, Test", ]

  # Update data classes from .csv schedule import to match repository structure
  new_sched <- new_sched %>%
    mutate(DOB = as.Date(DOB, format = "%m/%d/%Y"),
           `Made Date` = as.Date(`Made Date`, format = "%m/%d/%y"),
           Date = as.Date(Date, format = "%m/%d/%Y"),
           `Appt Time` = as.POSIXct(paste("1899-12-31 ", `Appt Time`),
                                    tz = "", format = "%Y-%m-%d %H:%M %p"),
           `ZIP Code` = substr(`ZIP Code`, 2, nchar(`ZIP Code`)))

  # Determine dates in new report
  current_dates <- sort(unique(new_sched$Date))

  if (initial_run) {
    sched_repo <- new_sched
  } else {
    # Update schedule repository by removing duplicate dates and adding data
    # from new report
    sched_repo <- sched_repo %>%
      filter(!(Date >= current_dates[1]))
    sched_repo <- rbind(sched_repo, new_sched)
  }

  saveRDS(sched_repo, paste0(user_directory, "/R_Sched_AM_Repo/",
                             "Auto Epic Report Sched ",
                             format(min(sched_repo$Date), "%m%d%y"), " to ",
                             format(max(sched_repo$Date), "%m%d%y"),
                             " on ", format(Sys.time(), "%m%d%y %H%M"), ".rds"))

} else {

  sched_repo <- readRDS(choose.files(default = paste0(user_directory,
                                                      "/R_Sched_AM_Repo/*.*"),
                                     caption = "Select schedule repository"))

}

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
             # Any blank immunizations assumed to be Pfizer
             ifelse(is.na(Immunizations), "Pfizer",
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
    NYZip = substr(`ZIP Code`, 1, 5) %in% ny_zips$zipcode)

# Crosswalk sites
sched_to_date <- left_join(sched_to_date, site_mappings,
                           by = c("Department" = "Department"))

# Roll up any Moderna administered at MSB for YWCA to MSH
sched_to_date <- sched_to_date %>%
  mutate(Site = ifelse(Site %in% c("MSB") & Mfg %in% c("Moderna"), "MSH", Site))

# Crosswalk pod type (employee vs patient)
sched_to_date <- left_join(sched_to_date,
                           pod_mappings[, c("Provider", "Pod Type")],
                           by = c("Provider/Resource" = "Provider"))

# Assume any pods without a mapping are patient pods
sched_to_date[is.na(sched_to_date$`Pod Type`), "Pod Type"] <- "Patient"

# Summarize schedule data by key stratification
sched_summary <- sched_to_date %>%
  group_by(Site, `Pod Type`, Dose, ApptDate, NYZip, Status2) %>%
  summarize(Count = n())

# # Remove Network LI from sched_to_date now that data has been summarized
# sched_to_date <- sched_to_date %>%
#   filter(!(Site %in% "Network LI"))

# Summarize schedule breakdown for next 2 weeks and export --------------------
# Format data
sched_breakdown <- sched_to_date %>%
  filter(Status2 %in% c("Arr", "Sch") &
           ApptDate >= (today - 1) &
           ApptDate <= (today + 14)) %>%
  group_by(Dose, `Pod Type`, Site, Status2, ApptDate) %>%
  summarize(Count = n(), .groups = "keep") %>%
  ungroup()

sched_breakdown <- sched_breakdown[order(sched_breakdown$Dose,
                                         sched_breakdown$`Pod Type`,
                                         sched_breakdown$Site), ]

# Create template dataframe of all sites, dose, and pod type combinations to
# ensure all are included even if there are no scheduled appointments
# Repeat dose types
doses_templ <- data.frame(
  "Dose" = rep(c(1:2),
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

# Combine site and doses
site_doses_templ <- data.frame(
  "Site" = rep(all_sites, 2),
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

# Combine site, dose, and pods
site_dose_pod_templ <- left_join(site_doses_templ,
                                 site_pods_templ,
                                 by = c("Site" = "Site"))

# Combine site, dose, pods, and dates
site_dose_pod_date_templ <- left_join(site_dates_templ,
                                      site_dose_pod_templ,
                                      by = c("Site" = "Site"))

site_dose_pod_date_templ <- site_dose_pod_date_templ %>%
  # select(Dose, PodType, Site) %>%
  mutate(Site = factor(Site, levels = all_sites, ordered = TRUE)) %>%
  arrange(Dose, PodType, Site, Date)

# Combine site, dose, and manufacturer
site_dose_mfg_templ <- left_join(site_doses_templ,
                                 site_mfg_templ,
                                 by = c("Site" = "Site"))

# Combine site, dose, manufacturer, and dates
site_dose_mfg_date_templ <- left_join(site_dates_templ,
                                      site_dose_mfg_templ,
                                      by = c("Site" = "Site"))

# Merge template dataframe with schedule breakdown data
sched_breakdown <- left_join(site_dose_pod_date_templ,
                               sched_breakdown,
                               by = c("Dose" = "Dose",
                                      "PodType" = "Pod Type",
                                      "Site" = "Site",
                                      "Date" = "ApptDate"))

# Replace NA in schedule breakdown and pivot wider for desired output
sched_breakdown <- sched_breakdown %>%
  mutate(Status2 = ifelse(is.na(Status2) & Date < today, "Arr",
                          ifelse(is.na(Status2) & Date >= today, "Sch",
                                 Status2)),
         Count = ifelse(is.na(Count), 0, Count))

sched_breakdown_cast <- sched_breakdown %>%
  pivot_wider(id_cols = c(Dose, PodType, Site),
              names_from = c(Status2, Date),
              values_from = Count)

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
                            ifelse(Site == "MSVD", "J&J",
                                   ifelse(Dose == 1, "Pfizer",
                                          ifelse(MRN %in% dose1_moderna$MRN,
                                                 "Moderna", "Pfizer")))),
         MfgRollUp = ifelse(is.na(Mfg) & is.na(ExpSchMfg), NA,
                            ifelse(is.na(ExpSchMfg), Mfg, ExpSchMfg)))

mfg_sched_breakdown <- city_sites_sched %>%
  filter(!is.na(MfgRollUp) &
         ApptDate >= (today - 1) &
           ApptDate <= (today + 14)) %>%
  group_by(Dose, MfgRollUp, Site, Status2, ApptDate) %>%
  summarize(Count = n(), .groups = "keep") %>%
  ungroup()

mfg_templ <- site_dose_mfg_date_templ %>%
  filter(Site %in% city_sites)

mfg_sched_breakdown <- left_join(mfg_templ,
                                 mfg_sched_breakdown,
                                 by = c("Dose" = "Dose",
                                        "Mfg" = "MfgRollUp",
                                        "Site" = "Site",
                                        "Date" = "ApptDate"))

# Replace NA in schedule breakdown and pivot wider for desired output
mfg_sched_breakdown_cast <- mfg_sched_breakdown %>%
  mutate(Status2 = ifelse(is.na(Status2) & Date < today, "Arr",
                          ifelse(is.na(Status2) & Date >= today, "Sch",
                                 Status2)),
         Count = replace_na(Count, 0)) %>%
  pivot_wider(id_cols = c(Dose, Mfg, Site),
              names_from = c(Status2, Date),
              values_from = Count) %>%
  mutate(Site = factor(Site, levels = city_sites, ordered = TRUE),
         Mfg = factor(Mfg, levels = mfg, ordered = TRUE)) %>%
  arrange(Dose, Mfg, Site)

# Create schedule breakdown for Pfizer only
pfizer_sched_breakdown <- city_sites_sched %>%
  filter((Mfg == "Pfizer" | ExpSchMfg == "Pfizer") &
           ApptDate >= (today - 1) & ApptDate <= (today + 14)) %>%
  group_by(Dose, `Pod Type`, Site, Status2, ApptDate) %>%
  summarize(Count = n(), .groups = "keep") %>%
  ungroup()

# Subset site, dose, pod type, date template for sites that have administered Pfizer
pfizer_templ <- site_dose_pod_date_templ %>%
  filter(Site %in% unique(pfizer_sched_breakdown$Site))

# Merge template with schedule data
pfizer_sched_breakdown <- left_join(pfizer_templ,
                                    pfizer_sched_breakdown,
                                    by = c("Dose" = "Dose",
                                           "PodType" = "Pod Type",
                                           "Site" = "Site",
                                           "Date" = "ApptDate"))

# Replace NA in schedule breakdown and pivot wider for desired output
pfizer_sched_breakdown <- pfizer_sched_breakdown %>%
  mutate(Status2 = ifelse(is.na(Status2) & Date < today, "Arr",
                          ifelse(is.na(Status2) & Date >= today, "Sch",
                                 Status2)),
         Count = ifelse(is.na(Count), 0, Count))

pfizer_sched_breakdown_cast <- pfizer_sched_breakdown %>%
  pivot_wider(id_cols = c(Dose, PodType, Site),
              names_from = c(Status2, Date),
              values_from = Count)

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
  "Dose" = rep(c(1:2), length(all_sites)),
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
  "Site" = rep(all_sites, 2),
  inv_cycle_dose_site)

inv_cycle_dates_site_templ <- data.frame(
  "Site" = rep(all_sites, length(sched_inv_cycle_dates)),
  inv_cycle_dates_site)

# Combine all templates into single template dataframe
inv_cycle_site_templ <- left_join(
  left_join(
    inv_cycle_pods_site_templ,
    inv_cycle_dose_site_templ,
    by = c("Site" = "Site")),
  inv_cycle_dates_site_templ,
  by = c("Site" = "Site"))

# Summarize schedule data for each site through next 2 inventory cycles
sched_inv_cycle_site <- sched_to_date %>%
  filter(ApptDate %in% sched_inv_cycle_dates &
           Status2 == "Sch") %>%
  group_by(Site, Dose, `Pod Type`, ApptDate) %>%
  summarize(Count = n(), .groups = "keep") %>%
  ungroup()

# Combine with template
sched_inv_cycle_site <- left_join(inv_cycle_site_templ,
                                  sched_inv_cycle_site,
                                  by = c("Site" = "Site",
                                         "Date" = "ApptDate",
                                         "Dose" = "Dose",
                                         "PodType" = "Pod Type"))

sched_inv_cycle_site <- sched_inv_cycle_site %>%
  mutate(Count = ifelse(is.na(Count), 0, Count),
         Site = factor(Site, levels = all_sites, ordered = TRUE)) %>%
  pivot_wider(id_cols = c(Site, Dose, Date),
              names_from = PodType,
              values_from = Count) %>%
  mutate(Total = Employee + Patient) %>%
  pivot_longer(cols = c(Employee, Patient, Total),
               names_to = "PodType",
               values_to = "Count") %>%
  arrange(Site, Date, Dose, desc(PodType)) %>%
  pivot_wider(id_cols = c(Site, Date),
              names_from = c(Dose, PodType),
              values_from = Count,
              names_sep = "_")

# Determine schedule across system within this inventory cycle's date range
# Create a dataframe of repeated pod types
inv_cycle_pods_sys <- data.frame(
  "PodType" = rep(pod_type[1:2], length(sched_inv_cycle_dates)),
  stringsAsFactors = FALSE)

inv_cycle_pods_sys <- inv_cycle_pods_sys %>%
  arrange(PodType)

# Create a dataframe of repeated doses
inv_cycle_dose_sys <- data.frame(
  "Dose" = rep(c(1:2), length(sched_inv_cycle_dates)),
  stringsAsFactors = FALSE)

inv_cycle_dose_sys <- inv_cycle_dose_sys %>%
  arrange(inv_cycle_dose_sys)

# Create template with pod types and doses for dates of interest
inv_cycle_pods_date_sys_templ <- data.frame(
  "Date" = rep(sched_inv_cycle_dates, 2),
  inv_cycle_pods_sys)

inv_cycle_dose_date_sys_templ <- data.frame(
  "Date" = rep(sched_inv_cycle_dates, 2),
  inv_cycle_dose_sys)

inv_cycle_sys_templ <- left_join(inv_cycle_pods_date_sys_templ,
                                          inv_cycle_dose_date_sys_templ,
                                          by = c("Date" = "Date"))

sched_inv_cycle_sys <- sched_to_date %>%
  filter(ApptDate %in% sched_inv_cycle_dates &
           Status2 == "Sch" &
           Site %in% city_sites) %>%
  group_by(Dose, `Pod Type`, ApptDate) %>%
  summarize(Count = n(), .groups = "keep") %>%
  ungroup()

# Combine with template
sched_inv_cycle_sys <- left_join(inv_cycle_sys_templ,
                                 sched_inv_cycle_sys,
                                 by = c("Date" = "ApptDate",
                                        "PodType" = "Pod Type",
                                        "Dose" = "Dose"))

sched_inv_cycle_sys <- sched_inv_cycle_sys %>%
  mutate(Count = ifelse(is.na(Count), 0, Count)) %>%
  pivot_wider(id_cols = c(Date, Dose),
              names_from = PodType,
              values_from = Count) %>%
  mutate(Total = Employee + Patient,
         DOW = wday(Date, label = TRUE, abbr = TRUE)) %>%
  pivot_longer(cols = c(Employee, Patient, Total),
               names_to = "PodType",
               values_to = "Count") %>%
  arrange(Date, Dose, desc(PodType)) %>%
  pivot_wider(id_cols = c(DOW, Date),
              names_from = c(Dose, PodType),
              values_from = Count,
              names_sep = "_")

# Summarize doses administered to date --------------------------------
# Summarize administered doses prior to today and stratify by site and pod type
admin_doses_site_pod_type <- sched_to_date %>%
  filter(Status2 == "Arr" &
           Site %in% city_sites) %>%
  group_by(Site, `Pod Type`, Dose) %>%
  summarize(DateRange = paste0(format(min(ApptDate), "%m/%d/%y"), "-",
                               format(max(ApptDate), "%m/%d/%y")),
            Count = n(),
            .groups = "keep") %>%
  ungroup()

# Summarize administered doses prior to today and stratify only by site
admin_doses_site_all_pods <- sched_to_date %>%
  filter(Status2 == "Arr" &
           Site %in% city_sites) %>%
  group_by(Site, Dose) %>%
  summarize("Pod Type" = "All",
            DateRange = paste0(format(min(ApptDate), "%m/%d/%y"), "-",
                               format(max(ApptDate), "%m/%d/%y")),
            Count = n(),
            .groups = "keep") %>%
  ungroup()

admin_doses_site_all_pods <-
  admin_doses_site_all_pods[, colnames(admin_doses_site_pod_type)]

# Summarize administered doses across system and stratify by pod type
# admin_doses_system_pod_type <- sched_to_date %>%
#   filter(Status2 == "Arr") %>%
#   group_by(`Pod Type`, Dose) %>%
#   summarize(Site = "MSHS", DateRange = paste0(format(min(ApptDate), "%m/%d/%y"),
#                                               "-",
#                                               format(max(ApptDate), "%m/%d/%y")),
#             Count = n()) %>%
#   ungroup()
# 
# admin_doses_system_pod_type <-
#   admin_doses_system_pod_type[, colnames(admin_doses_site_pod_type)]

# Summarize administered doses across system
# admin_doses_system_all_pods <- sched_to_date %>%
#   filter(Status2 == "Arr") %>%
#   group_by(Dose) %>%
#   summarize(Site = "MSHS",
#             `Pod Type` = "All",
#             DateRange = paste0(format(min(ApptDate), "%m/%d/%y"), "-",
#                                format(max(ApptDate), "%m/%d/%y")),
#             Count = n()) %>%
#   ungroup()
# 
# admin_doses_system_all_pods <-
#   admin_doses_system_all_pods[, colnames(admin_doses_site_pod_type)]

# Combine administered data by pod type, across pods, and across system
admin_doses_summary <- rbind(admin_doses_site_pod_type,
                             admin_doses_site_all_pods)
                             # admin_doses_system_pod_type,
                             # admin_doses_system_all_pods)

admin_doses_summary <- admin_doses_summary %>%
  mutate(Dose = ifelse(Dose == 1, "Dose1", "Dose2"))

admin_2nd_dose_remaining <- dcast(admin_doses_summary,
                                  Site + `Pod Type` ~ Dose, value.var = "Count")

admin_2nd_dose_remaining <- admin_2nd_dose_remaining %>%
  mutate(RemainingDose2 = ifelse(Site %in% c("MSVD"), NA, Dose1 - Dose2),
         Site = factor(Site, levels = city_sites, ordered = TRUE),
         `Pod Type` = factor(`Pod Type`, levels = pod_type, ordered = TRUE))

admin_2nd_dose_remaining <-
  admin_2nd_dose_remaining[order(admin_2nd_dose_remaining$Site,
                                 admin_2nd_dose_remaining$`Pod Type`), ]

admin_doses_table_export <- melt(admin_2nd_dose_remaining,
                                 id = c("Site", "Pod Type"))

admin_doses_table_export <- dcast(admin_doses_table_export,
                                  Site ~ variable + `Pod Type`,
                                  value.var = "value")

# Determine MSHS total by summing columns
admin_doses_system_summary <- data.frame(
  t(colSums(
    admin_doses_table_export[, c(2:ncol(admin_doses_table_export))],
    na.rm = TRUE, dims = 1)))

admin_doses_system_summary <- admin_doses_system_summary %>%
  mutate(Site = "MSHS")

admin_doses_system_summary <- admin_doses_system_summary[, colnames(admin_doses_table_export)]

# Bind site and system data
admin_doses_table_export <- rbind(admin_doses_table_export,
                                  admin_doses_system_summary)

# Arrange dataframe by site
admin_doses_table_export <- admin_doses_table_export %>%
  arrange(Site)

# Stratify administered doses to date by manufacturer ------------------------
# Create daily summary of administered doses by manufacturer
admin_mfg_summary <- sched_to_date %>%
  filter(Status2 == "Arr" &
           Site %in% city_sites) %>%
  group_by(Site, Mfg, Dose, ApptDate, Status2) %>%
  summarize(Count = n(), .groups = "keep")

# Create summary of administered doses prior to yesterday and yesterday
admin_prior_yest_site <- sched_to_date %>%
  filter(Status2 == "Arr" &
           ApptDate <= (today - 2) &
           Site %in% city_sites) %>%
  group_by(Site, Mfg, Dose) %>%
  summarize(DateRange = paste0("Admin 12/15/20-",
                               format(today - 2, "%m/%d/%y")),
            Count = n(), .groups = "keep")

admin_yest_site <- sched_to_date %>%
  filter(Status2 == "Arr" &
           ApptDate == (today - 1) &
           Site %in% city_sites) %>%
  group_by(Site, Mfg, Dose) %>%
  summarize(DateRange = paste0("Admin ", format(today - 1, "%m/%d/%y")),
            Count = n(),
            .groups = "keep")

admin_to_date_mfg <- rbind(admin_prior_yest_site, admin_yest_site)

admin_to_date_mfg <- admin_to_date_mfg %>%
  ungroup() %>%
  arrange(Site, Dose, desc(Mfg), desc(DateRange))

admin_to_date_mfg <-
  admin_to_date_mfg[, c("Dose", "Mfg", "Site", "DateRange", "Count")]

admin_to_date_mfg <- admin_to_date_mfg %>%
  mutate(DateRange = factor(DateRange, level = unique(DateRange),
                            ordered = TRUE))

admin_to_date_mfg_cast <- dcast(admin_to_date_mfg,
                                Dose + Mfg + Site ~ DateRange,
                                value.var = "Count")

# Create template dataframe of sites, doses, and manufacturers to ensure all
# combinations are included
rep_doses <- data.frame("Dose" = rep(c(1:2),
                                     length(unique(city_sites_sched$Site))))
rep_doses <- rep_doses %>%
  arrange(Dose)

rep_mfg <- data.frame("Mfg" = rep(mfg,
                                  length(unique(city_sites_sched$Site))),
                      stringsAsFactors = FALSE)

rep_mfg <- rep_mfg %>%
  arrange(desc(Mfg))

rep_sites_mfg <- data.frame("Site" = rep(unique(city_sites_sched$Site),
                                         length(mfg)), "Mfg" = rep_mfg,
                            stringsAsFactors = FALSE)

rep_sites_doses <- data.frame("Site" = rep(unique(city_sites_sched$Site), 2),
                              "Dose" = rep_doses,
                              stringsAsFactors = FALSE)

# Combine template with administration data
city_sites_doses_mfg <- left_join(rep_sites_mfg, rep_sites_doses,
                             by = c("Site" = "Site"))

city_sites_doses_mfg <- city_sites_doses_mfg %>%
  mutate(Site = factor(Site, levels = city_sites, ordered = TRUE)) %>%
  arrange(Dose, desc(Mfg), Site)

city_sites_doses_mfg <- city_sites_doses_mfg[, c("Dose", "Mfg", "Site")]

admin_to_date_mfg_export <- left_join(city_sites_doses_mfg, admin_to_date_mfg_cast,
                                      by = c("Dose" = "Dose",
                                             "Mfg" = "Mfg",
                                             "Site" = "Site"))


# Walk-In stats for last 7 days ------------------------------------------
# Determine daily walk-ins by pod type
dose1_7day_walkins_type <- sched_to_date %>%
  filter(Status2 == "Arr" &
           Dose == 1 &
           ApptDate >= (today - 7) &
           ApptDate <= today) %>%
  group_by(Site, `Pod Type`, Dose,
           ApptYear, WeekNum, ApptDate, DOW) %>%
  summarize(DailyArrivals = n(),
            DailyWalkIns = sum(`Same Day?` == "Same day"),
            DailyWalkInPercent = percent(DailyWalkIns / DailyArrivals))

# Determine daily walk-ins for each site
dose1_7day_walkins_totals <- sched_to_date %>%
  filter(Status2 == "Arr" &
           Dose == 1 &
           ApptDate >= (today - 7) &
           ApptDate <= today) %>%
  group_by(Site, Dose,
           ApptYear, WeekNum, ApptDate, DOW) %>%
  summarize(`Pod Type` = "All",
            DailyArrivals = n(),
            DailyWalkIns = sum(`Same Day?` == "Same day"),
            DailyWalkInPercent = percent(DailyWalkIns / DailyArrivals))

# Reorder columns
dose1_7day_walkins_totals <-
  dose1_7day_walkins_totals[, colnames(dose1_7day_walkins_type)]

# Combine daily walk-ins for each site by pod type and site total
dose1_7day_walkins_summary <- rbind(dose1_7day_walkins_type,
                                    dose1_7day_walkins_totals)

dose1_7day_walkins_summary <- dose1_7day_walkins_summary %>%
  ungroup()

# Reorder rows based on site and pod type
dose1_7day_walkins_summary <- dose1_7day_walkins_summary %>%
  mutate(Site = factor(Site, levels = all_sites, ordered = TRUE),
         `Pod Type` = factor(`Pod Type`, levels = pod_type, ordered = TRUE))

dose1_7day_walkins_summary <- dose1_7day_walkins_summary %>%
  arrange(Site, `Pod Type`)

# Calculate average walk-in volume and average walk-in percentage
dose1_7day_walkins_avg <- dose1_7day_walkins_summary %>%
  group_by(Site, Dose, `Pod Type`) %>%
  summarize(AvgWalkInVolume = sum(DailyWalkIns) / sum(DailyWalkIns > 0),
            WalkInPercent = percent(sum(DailyWalkIns) / sum(DailyArrivals))) %>%
  ungroup()

cast_dose1_7day_walkins_arr <- dcast(dose1_7day_walkins_summary,
                                     Site + `Pod Type` ~ ApptDate,
                                     value.var = "DailyWalkIns")

cast_dose1_7day_all_arr <- dcast(dose1_7day_walkins_summary,
                                 Site + `Pod Type` ~ ApptDate,
                                 value.var = "DailyArrivals")

cast_dose1_7day_walkins_stats <- melt(dose1_7day_walkins_summary,
                                      id.vars = c("Site",
                                                  "Pod Type",
                                                  "ApptDate"),
                                      measure.vars = c("DailyWalkIns",
                                                       "DailyArrivals",
                                                       "DailyWalkInPercent"))

cast_dose1_7day_walkins_stats <- dcast(cast_dose1_7day_walkins_stats,
                                       Site + `Pod Type` ~ variable + ApptDate,
                                       value.var = "value")

cast_dose1_7day_walkins_stats <-
  left_join(cast_dose1_7day_walkins_stats,
            dose1_7day_walkins_avg[, c("Site",
                                        "Pod Type",
                                        "AvgWalkInVolume",
                                        "WalkInPercent")],
            by = c("Site" = "Site", "Pod Type" = "Pod Type"))

cast_melt_test <- melt(dose1_7day_walkins_summary,
                       id.vars = c("Site", "Pod Type", "ApptDate"),
                       measure.vars = c("DailyWalkIns",
                                        "DailyArrivals",
                                        "DailyWalkInPercent"))


# Export key data tables to Excel for reporting -------------------------------
# Export schedule summary, schedule breakdown, and cumulative administered
# doses to excel file
export_list <- list("SchedSummary" = sched_summary,
                    "SchedBreakdown" = sched_breakdown_cast,
                    "Mfg_SchedBreakdown" = mfg_sched_breakdown_cast,
                    "Pfizer_SchedBreakdown" = pfizer_sched_breakdown_cast,
                    # "Sched_Targets_7Days" = sched_7days_format_jm,
                    "Sched_InvCycle_Site" = sched_inv_cycle_site,
                    "Sched_InvCycle_Sys" = sched_inv_cycle_sys,
                    "CumDosesAdministered" = admin_doses_table_export,
                    "AdminMfgSummary" = admin_mfg_summary,
                    "DosesAdminMfg" = admin_to_date_mfg_export,
                    "1stDoseArrivals_7Days" = cast_dose1_7day_all_arr,
                    "1stDoseWalkIns_7Days" = cast_dose1_7day_walkins_arr,
                    "WalkInsStats_7Days" = cast_dose1_7day_walkins_stats)

write_xlsx(export_list, path = paste0(user_directory,
                                      "/R_Sched_Analysis_AM_Export/",
                                      "Auto Epic Report Sched Data Export ",
                                      format(Sys.time(), "%m%d%y %H%M"),
                                      ".xlsx"))
