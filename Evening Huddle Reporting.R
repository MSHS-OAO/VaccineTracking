# Code for updating data analysis for afternoon huddle

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
#install.packages(zipcodeR)
#install.packages("tidyr")
# install.packages("janitor")

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

# Set path for file selection
user_path <- paste0(user_directory, "\\*.*")

# Determine whether or not to update an existing repo
update_repo <- TRUE

# Determine whether or not to update an existing repo
initial_run <- FALSE
update_repo <- TRUE

# Import reference data for site and pod mappings
site_mappings <- read_excel(paste0(user_directory, "/ScheduleData/",
                                   "Automation Ref 2021-03-12.xlsx"),
                            sheet = "Site Mappings")
pod_mappings <- read_excel(paste0(user_directory, "/ScheduleData/",
                                  "Automation Ref 2021-03-12.xlsx"),
                           sheet = "Pod Mappings Simple")

# Store today's date
today <- Sys.Date() - 2

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

# Pod type order
pod_type <- c("Employee", "Patient", "All")

# Vaccine manufacturers
mfg <- c("Pfizer", "Moderna", "J&J")

# NY zip codes
ny_zips <- search_state("NY")

# First vaccine date
vacc_start_date <- as.Date("12/15/20", format = "%m/%d/%y")
vacc_start_first_sun <- vacc_start_date - (wday(vacc_start_date) - 1)

# Determine current week of vaccine effort
this_week <- as.numeric(floor((today - vacc_start_first_sun) / 7) + 1)

# Create dataframe with dates and weeks since vaccines started
all_dates <- data.frame("Date" = seq.Date(as.Date("12/15/20",
                                                  format = "%m/%d/%y"),
                                          today + 7,
                                          by = 1))

all_dates <- all_dates %>%
  mutate(VaccWeek = as.numeric(
    floor((Date - vacc_start_first_sun) / 7) + 1))

vacc_week_dates <- all_dates %>%
  group_by(VaccWeek) %>%
  summarize(Dates = 
              paste0(format(min(Date), "%m/%d/%y"),
                     "-",
                     format(max(Date), "%m/%d/%y")))

# Create variable with doses
doses <- c("Dose1", "Dose2", "AllDoses")

# Import raw data from Epic
if (update_repo) {
  sched_repo <- readRDS(
    choose.files(default = paste0(user_directory,
                                  "/R_Sched_AM_Repo/*.*"),
                 caption = "Select this morning's schedule repository"))
  
  raw_df <- read.csv(choose.files(
    default =
      paste0(user_directory, "/Epic 340PM Sched All Status/*.*"),
    caption = "Select this afternoon's Epic schedule"),
    stringsAsFactors = FALSE,
    check.names = FALSE)
  
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
  
  # Update schedule repository by removing duplicate dates and adding data
  # from new report
  sched_repo <- sched_repo %>%
    filter(!(Date >= current_dates[1]))
  sched_repo <- rbind(sched_repo, new_sched)
  
  # Save schedule
  saveRDS(sched_repo, paste0(user_directory, "/R_Sched_PM_Huddle_Repo/",
                             "Epic Sched 340PM ",
                             format(min(sched_repo$Date), "%m%d%y"), " to ",
                             format(max(sched_repo$Date), "%m%d%y"),
                             " on ", format(Sys.time(), "%m%d%y %H%M"), ".rds"))
  
} else {
  
  sched_repo <- readRDS(
    choose.files(default = paste0(user_directory,
                                  "/R_Sched_AM_Repo/*.*"),
                 caption = "Select this morning's schedule repository"))
  
}

# Import MSSN data from Tableau
mssn_admin <- read_excel(
  path = choose.files(
    default = paste0(user_directory, "/MSSN Tableau Doses/*.*"),
    caption = "Select latest MSSN Tableau Administration Data"),
  col_names = TRUE,
  na = c("", NA),
  skip = 1)

mssn_admin <- mssn_admin[1:nrow(mssn_admin) - 1, ]

colnames(mssn_admin) <- c("ApptDate", "Dose1", "Dose2", "AllDoses")


# Format and analyze Epic schedule to date for dashboards ---------------------
epic_sched_to_date <- sched_repo

epic_sched_to_date <- epic_sched_to_date %>%
  mutate(
    # Determine whether appointment is for dose 1 or dose 2
    Dose = ifelse(str_detect(Type, "DOSE 1"), "Dose1",
                  ifelse(str_detect(Type, "DOSE 2"), "Dose2", NA)),
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
    Status2 = ifelse(ApptDate < (today) & Status == "Sch", "No Show", Status),
    # Determine manufacturer based on immunization field
    Mfg = #Keep manufacturer as NA if the appointment hasn't been arrived
      ifelse(Status2 != "Arr", NA,
             # Assume any visits without immunization record are Pfizer
             ifelse(is.na(Immunizations), "Pfizer",
                    # Any immunizations with Moderna in text classified as
                    # Moderna, otherwise Pfizer
                    ifelse(str_detect(Immunizations, "Moderna"), "Moderna",
                           # Any immunizations with Johnson and Johnson are
                           # classified as J&J
                           ifelse(str_detect(Immunizations,
                                             "Johnson and Johnson"),
                                  "J&J", "Pfizer")))),
    # Determine week number of year
    WeekNum = format(ApptDate, "%U"),
    # Determine appointment day of week
    DOW = weekdays(ApptDate),
    # Determine if patient's zip code is in NYS
    NYZip = substr(`ZIP Code`, 1, 5) %in% ny_zips$zipcode,
    # Determine week number of vaccine efforts
    VaccWeek = as.numeric(floor((ApptDate - vacc_start_first_sun) / 7) + 1))

# Crosswalk sites
epic_sched_to_date <- left_join(epic_sched_to_date, site_mappings,
                           by = c("Department" = "Department"))

# Roll up any Moderna administered at MSB for YWCA to MSH
epic_sched_to_date <- epic_sched_to_date %>%
  mutate(Site = ifelse(Site %in% c("MSB") & Mfg %in% c("Moderna"), "MSH", Site))

# Crosswalk pod type (employee vs patient)
epic_sched_to_date <- left_join(epic_sched_to_date,
                           pod_mappings[, c("Provider", "Pod Type")],
                           by = c("Provider/Resource" = "Provider"))

# Assume any pods without a mapping are patient pods
epic_sched_to_date[is.na(epic_sched_to_date$`Pod Type`), "Pod Type"] <- "Patient"

# Summarize schedule data for Epic sites by key stratification
epic_sched_summary <- epic_sched_to_date %>%
  filter(Status2 %in% c("Arr", "Sch") &
           ApptDate <= today + 7) %>%
  group_by(Site,
           Dose,
           VaccWeek,
           ApptDate,
           Status2) %>%
  summarize(Count = n(),
            .groups = "keep") %>%
  pivot_wider(names_from = Dose,
              values_from = Count) %>%
  mutate(AllDoses = sum(Dose1, Dose2, na.rm = TRUE)) %>%
  pivot_longer(cols = c("Dose1", "Dose2", "AllDoses"),
               names_to = "Dose",
               values_to = "Count") %>%
  ungroup()


# Melt MSSN data to format similar to Epic sites
mssn_admin_to_date <- mssn_admin %>%
  mutate(Site = "MSSN",
         Status2 = "Arr",
         ApptDate = as.Date(ApptDate, format = "%B %d, %Y"),
         VaccWeek = as.numeric(
           floor((ApptDate - vacc_start_first_sun) / 7 +1))) %>%
  pivot_longer(cols = c("Dose1", "Dose2", "AllDoses"),
               names_to = "Dose",
               values_to = "Count")

mssn_admin_to_date <- mssn_admin_to_date[, colnames(epic_sched_summary)]

# Bind MSSN data with Epic sites data
daily_sched_summary <- rbind(epic_sched_summary, mssn_admin_to_date)

daily_sched_summary <- daily_sched_summary %>%
  left_join(vacc_week_dates, by = c("VaccWeek" = "VaccWeek"))

# Summarize total doses to date through today
admin_to_date <- daily_sched_summary %>%
  filter(ApptDate <= today) %>%
  group_by(Site, Dose) %>%
  summarize(AdminDoses = sum(Count, na.rm = TRUE),
            .groups = "keep") %>%
  ungroup()

# Summarize administered doses for prior weeks -----------------
prior_weeks_admin <- daily_sched_summary %>%
  filter(VaccWeek < this_week) %>%
  group_by(Site, Dose, VaccWeek, Dates) %>%
  summarize(AdminDoses = sum(Count, na.rm = TRUE),
            .groups = "keep") %>%
  ungroup() %>%
  pivot_wider(names_from = c(VaccWeek, Dates),
              names_prefix = "Wk ",
              names_sep = ": ",
              values_from = AdminDoses)

# Summarize total doses administered and scheduled this week through today -----
this_week_admin_to_date <- daily_sched_summary %>%
  filter(VaccWeek == this_week &
           ApptDate <= today) %>%
  group_by(Site, Dose, VaccWeek, Dates) %>%
  summarize(AdminDoses = sum(Count, na.rm = TRUE),
            .groups = "keep") %>%
  ungroup() %>%
  pivot_wider(names_from = c(VaccWeek, Dates),
              names_prefix = "Wk ",
              names_sep = ": ",
              values_from = AdminDoses)

# Summarize data for this week to date -------------------
this_week_daily_admin <- daily_sched_summary %>%
  filter(VaccWeek == this_week &
           ApptDate < today) %>%
  group_by(Site, Dose, VaccWeek, ApptDate) %>%
  summarize(AdminDoses = sum(Count, na.rm = TRUE),
            .groups = "keep") %>%
  ungroup() %>%
  pivot_wider(names_from = c(VaccWeek, ApptDate),
              names_prefix = "Wk ",
              names_sep = c(": "),
              values_from = AdminDoses,
              names_sort = TRUE)

# Summarize today's arrived and scheduled doses -------------
todays_sched <- daily_sched_summary %>%
  filter(ApptDate == today) %>%
  group_by(Site, Dose, ApptDate, Status2) %>%
  summarize(AdminDoses = sum(Count, na.rm = TRUE),
            .groups = "keep") %>%
  ungroup() %>%
  pivot_wider(names_from = Status2,
              values_from = AdminDoses) %>%
  mutate(Total = sum(Arr, Sch, na.rm = TRUE),
         VaccWeek = NULL) %>%
  pivot_longer(cols = c(Arr, Sch, Total),
               names_to = "ApptStatus",
               values_to = "Count") %>%
  pivot_wider(names_from = c(ApptDate, ApptStatus),
              names_prefix = "Today's Sched: ",
              values_from = c(Count))

# Summarize scheduled data for the next 7 days ---------------
next_7days_sched <- daily_sched_summary %>%
  filter(ApptDate %in% seq(from = today + 1, to = today + 7, by = 1)) %>%
  group_by(Site, Dose, ApptDate) %>%
  summarize(SchedVisits = sum(Count, na.rm = TRUE),
            .groups = "keep") %>%
  ungroup() %>%
  pivot_wider(names_from = ApptDate,
              values_from = SchedVisits)
  
# Create a template of all sites, vaccine dates, and doses ---------
site_rep <- data.frame("Site" = sort(rep(all_sites,
                                         length(doses))))

dose_rep <- data.frame("Dose" = rep(doses,
                                    length(all_sites)))

site_dose_rep <- cbind(site_rep, dose_rep)

site_dose_rep <- site_dose_rep %>%
  mutate(Site = factor(Site, levels = all_sites, ordered = TRUE)) %>%
  arrange(Site, Dose)

# Combine all tables
site_data_summary <- 
  left_join(
    left_join(
      left_join(
        left_join(
          left_join(
            left_join(
              #First, join site and dose template with administered doses to date
              site_dose_rep,
              admin_to_date,
              by = c("Site" = "Site", "Dose" = "Dose")),
            # Second, join with administered doses from prior weeks
            prior_weeks_admin,
            by = c("Site" = "Site", "Dose" = "Dose")),
          # Third, join this week's doses administered/scheduled through today
          this_week_admin_to_date,
          by = c("Site" = "Site", "Dose" = "Dose")),
        # Fourth, join this week's daily data through today
        this_week_daily_admin,
        by = c("Site" = "Site", "Dose" = "Dose")),
      # Fifth, join today's arrived and scheduled appts
      todays_sched,
      by = c("Site" = "Site", "Dose" = "Dose")),
    # Last, join scheduled doses for the next 7 days
    next_7days_sched,
    by = c("Site" = "Site", "Dose" = "Dose"))

# Determine system-wide totals for each dose
system_wide_summary <- do.call(
  #Combine list of dataframes into single dataframe
  rbind, lapply(
    # First, split dataframe into list of dataframes by dose type
    site_data_summary %>%
      group_split(Dose),
    # Next, adorn totals for each dataframe
    function (x)
      x %>%
      adorn_totals(where = "row",
                   fill = unique(x$Dose),
                   na.rm = TRUE,
                   name = "MSHS",
                   -Dose)))
      
# Subset for first dose, second dose, and all doses
all_doses <- system_wide_summary %>%
  filter(Dose == "AllDoses")

first_doses <- system_wide_summary %>%
  filter(Dose == "Dose1")

second_doses <- system_wide_summary %>%
  filter(Dose == "Dose2")
