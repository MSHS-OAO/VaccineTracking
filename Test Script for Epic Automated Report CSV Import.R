# Code for creating repository of COVID vaccine administration ----------------------

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
# install.packages("ggpubr")
#install.package(zipcodeR)

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
if (list.files("J://") == "Presidents") {
  user_directory <- "J:/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/System Operations/COVID Vaccine"
} else {
  user_directory <- "J:/deans/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/System Operations/COVID Vaccine"
}

# Set path for file selection
user_path <- paste0(user_directory, "\\*.*")

# Determine whether or not to update an existing repo
initial_run <- FALSE
update_repo <- TRUE

# Reports to add to repo
num_new_reports <- 1

# Determine whether or not to update walk-in analysis
update_walkins <- FALSE

# Import reference data for site and pod mappings
site_mappings <- read_excel(paste0(user_directory, "/ScheduleData/Automation Ref 2021-01-20.xlsx"), sheet = "Site Mappings")
pod_mappings <- read_excel(paste0(user_directory, "/ScheduleData/Automation Ref 2021-01-20.xlsx"), sheet = "Pod Mappings Simple")

# Store today's date
today <- Sys.Date()

# Create dataframe with dates and weeks
all_dates <- data.frame("Date" = seq.Date(as.Date("1/1/20", format = "%m/%d/%y"), 
                                          as.Date("12/31/21", format = "%m/%d/%y"), 
                                          by = 1))
all_dates <- all_dates %>%
  mutate(Year = year(Date),
         WeekNum = format(Date, "%U"))

# Site order
sites <- c("MSB", "MSBI", "MSH", "MSM", "MSQ", "MSW", "MSHS")

# Pod type order
pod_type <- c("Employee", "Patient", "All")

# NY zip codes
ny_zips <- search_state("NY")


sched_repo <- readRDS(choose.files(default = paste0(user_directory, "/R_Sched_AM_Repo/*.*"), caption = "Select schedule repository"))

raw_epic_csv <- read.csv(choose.files(default = paste0(user_directory, "/Auto Epic Sched Reports/*.*"),
                                      caption = "Select current Epic schedule"), stringsAsFactors = FALSE)

raw_df1 <- read_excel(choose.files(default = paste0(user_directory, "/ScheduleData/*.*"), caption = "Select 1st new Epic schedule"), 
                      col_names = TRUE, na = c("", "NA"))

new_sched <- raw_df1
  
# Remove test patient
new_sched <- new_sched[new_sched$Patient != "Patient, Test", ]
  
# Create column with vaccine location
new_sched <- new_sched %>% 
  mutate(Mfg = ifelse(is.na(Immunizations), "Unknown", ifelse(str_detect(Immunizations, "Pfizer"), "Pfizer", "Moderna")),
         Dose = ifelse(str_detect(Type, "DOSE 1"), 1, ifelse(str_detect(Type, "DOSE 2"), 2, NA)),
         ApptDate = date(Date),
         ApptYear = year(Date),
         ApptMonth = month(Date),
         ApptDay = day(Date),
         ApptWeek = week(Date))
  
# Determine dates in new report
current_dates <- sort(unique(new_sched$ApptDate))

# Update schedule repository by removing duplicate dates and adding data from new report
sched_repo <- sched_repo %>%
  filter(!(ApptDate >= current_dates[1]))
sched_repo <- rbind(sched_repo, new_sched)

# Format and analyze schedule to date for dashboards ---------------------------
sched_to_date <- sched_repo

# Format new epic report
new_epic_sched <- raw_epic_csv %>%
  mutate(DOB = as.POSIXct(DOB, format = "%m/%d/%Y"),
         Made.Date = as.POSIXct(Made.Date, format = "%m/%d/%y"),
         Date = as.POSIXct(Date, format = "%m/%d/%Y"),
         Appt.Time = as.POSIXct(paste("1899-12-31", Appt.Time), format = "%Y-%m-%d %H:%M %p"),
         Mfg = ifelse(is.na(Immunizations), "Unknown", ifelse(str_detect(Immunizations, "Pfizer"), "Pfizer", "Moderna")),
         Dose = ifelse(str_detect(Type, "DOSE 1"), 1, ifelse(str_detect(Type, "DOSE 2"), 2, NA)),
         ApptDate = date(Date),
         ApptYear = year(Date),
         ApptMonth = month(Date),
         ApptDay = day(Date),
         ApptWeek = week(Date),
         Canc.Date = NULL)

colnames(new_epic_sched) <- colnames(new_sched)

test_df <- rbind(sched_to_date, new_sched, new_epic_sched)
