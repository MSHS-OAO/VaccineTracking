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

# Import reference data for site and pod mappings
site_mappings <- read_excel(paste0(user_directory, "/ScheduleData/",
                                   "Automation Ref 2021-02-25.xlsx"),
                            sheet = "Site Mappings")
pod_mappings <- read_excel(paste0(user_directory, "/ScheduleData/",
                                  "Automation Ref 2021-02-25.xlsx"),
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
sites <- c("MSB", "MSBI", "MSH", "MSM", "MSQ", "MSW", "MSHS")

# Pod type order
pod_type <- c("Employee", "Patient", "All")

# Vaccine manufacturers
mfg <- c("Pfizer", "Moderna")

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
    NYZip = substr(`ZIP Code`, 1, 5) %in% ny_zips$zipcode)

# Crosswalk sites
sched_to_date <- left_join(sched_to_date, site_mappings,
                           by = c("Department" = "Department"))

# Crosswalk pod type (employee vs patient)
sched_to_date <- left_join(sched_to_date,
                           pod_mappings[, c("Provider", "Pod Type")],
                           by = c("Provider/Resource" = "Provider"))

# Assume any pods without a mapping are patient pods
sched_to_date[is.na(sched_to_date$`Pod Type`), "Pod Type"] <- "Patient"

# Subset patients who have arrived or are scheduled for dose 1 appt
dose1_arr_sch <- sched_to_date %>%
  filter(Status2 %in% c("Arr", "Sch") & Dose == 1) %>%
  select(Type, Site, Department,
         `Provider/Resource`, `Pod Type`,
         MRN, ApptDate, Status2) %>%
  mutate(DuplMRN = duplicated(MRN)) %>%
  arrange(Site, `Pod Type`)

write_xlsx(x = dose1_arr_sch,
           path = paste0(user_directory, "/AdHoc/Dose1 ArrSch Pts ",
                         format(min(dose1_arr_sch$ApptDate), "%m%d%Y"),
                         "-",
                         format(max(dose1_arr_sch$ApptDate), "%m%d%Y"),
                         " As Of 340PM ", format(Sys.Date(), "%m%d%Y"),
                         ".xlsx"))


