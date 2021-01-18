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

# Set working directory and select raw data ----------------------------
rm(list = ls())

# Determine path for working directory
if (list.files("J://") == "Presidents") {
  user_directory <- "J:/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/System Operations/COVID Vaccine/ScheduleData/Automation Files"
} else {
  user_directory <- "J:/deans/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/System Operations/COVID Vaccine/ScheduleData/Automation Files"
}

# Set path for file selection
user_path <- paste0(user_directory, "\\*.*")

# Determine whether or not to update an existing repo
initial_run <- FALSE
update_repo <- TRUE

# Import reference data for site and pod mappings
site_mappings <- read_excel(paste0(user_directory, "/Reference 2021-01-15.xlsx"), sheet = "Site Mappings")
pod_mappings <- read_excel(paste0(user_directory, "/Reference 2021-01-15.xlsx"), sheet = "Pod Mappings Simple")

# Store today's date
today <- Sys.Date()

# Site order
sites <- c("MSB", "MSBI", "MSH", "MSM", "MSQ", "MSW")

# Pod type order
pod_type <- c("Employee", "Patient", "All")


# Import raw data from Epic
if (update_repo) {
  if (initial_run) {
    raw_df <- read_excel(choose.files(default = user_path, caption = "Select initial Epic schedule"), 
                                  col_names = TRUE, na = c("", "NA"))
    } else {
      # sched_repo <- read_excel(choose.files(default = user_path, caption = "Select schedule repository"), 
      #                       col_names = TRUE, na = c("", "NA"))
      sched_repo <- readRDS(choose.files(default = user_path, caption = "Select schedule repository"))
      
      # Convert appointment date in schedule repository from posixct to date
      sched_repo <- sched_repo %>%
        mutate(ApptDate = date(ApptDate))
      
      raw_df <- read_excel(choose.files(default = user_path, caption = "Select current Epic schedule"), 
                       col_names = TRUE, na = c("", "NA"))
    }
  
  new_sched <- raw_df
  
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
  
  if (initial_run) {
    sched_repo <- new_sched
  } else {
    # Update schedule repository by removing duplicate dates and adding data from new report
    sched_repo <- sched_repo %>%
      filter(!(ApptDate >= current_dates[1]))
    sched_repo <- rbind(sched_repo, new_sched)
  }
  
  # Export updated schedule repository
  # write_xlsx(sched_repo, paste0(user_directory, "/Sched Repo ", 
  #                               format(min(sched_repo$ApptDate), "%m%d%y"), " to ", 
  #                               format(max(sched_repo$ApptDate), "%m%d%y"), 
  #                               " on ", format(Sys.time(), "%m%d%y %H%M"), ".xlsx"))
  
  saveRDS(sched_repo, paste0(user_directory, "/Sched Repo ", 
                             format(min(sched_repo$ApptDate), "%m%d%y"), " to ", 
                             format(max(sched_repo$ApptDate), "%m%d%y"), 
                             " on ", format(Sys.time(), "%m%d%y %H%M"), ".rds"))
  
} else {
  # sched_repo <- read_excel(choose.files(default = user_path, caption = "Select schedule repository"), 
  #                          col_names = TRUE, na = c("", "NA"))
  sched_repo <- readRDS(choose.files(default = user_path, caption = "Select schedule repository"))
  
  
}

# saveRDS(sched_repo, file = paste0(user_directory, "/Sched Repo ", 
#                                   format(min(sched_repo$ApptDate), "%m%d%y"), " to ", 
#                                   format(max(sched_repo$ApptDate), "%m%d%y"), 
#                                   " on ", format(Sys.time(), "%m%d%y %H%M"), ".rds"))
# 
# a <- readRDS(file = choose.files(default = user_path, caption = "Select schedule repository"))
# 
# test_a <- a %>%
#   group_by(Department, ApptDate, `Appt Status`) %>%
#   summarize(Count = n())
# 
# test_sched_repo <- sched_repo %>%
#   group_by(Department, ApptDate, `Appt Status`) %>%
#   summarize(Count = n())

# Format and analyze schedule to date for dashboards
sched_to_date <- sched_repo

sched_to_date <- left_join(sched_to_date, site_mappings,
                           by = c("Department" = "Department"))

sched_to_date <- left_join(sched_to_date, pod_mappings[ , c("Provider", "Pod Type")],
                           by = c("Provider/Resource" = "Provider"))

sched_to_date[is.na(sched_to_date$`Pod Type`), "Pod Type"] <- "Patient"

sched_to_date <- sched_to_date %>%
  mutate(Status = ifelse(`Appt Status` %in% c("Arrived", "Comp", "Checked Out"), "Arr", `Appt Status`),
         WeekNum = format(ApptDate, "%U"),
         DOW = weekdays(ApptDate))

# Start summarizing data
dose1_7days_walkins_type <- sched_to_date %>%
  filter(Status == "Arr" & Dose == 1 & ApptDate >= (today - 7) & ApptDate < today) %>%
  group_by(Site, `Pod Type`, Dose,
           ApptYear, WeekNum, ApptDate, DOW) %>%
  summarize(TotalArrivals = n(),
            WalkIns = sum(`Same Day?` == "Same day"),
            WalkInPercent = WalkIns / TotalArrivals * 100)

dose1_7day_walkins_totals <- sched_to_date %>%
  filter(Status == "Arr" & Dose == 1 & ApptDate >= (today - 7) & ApptDate < today) %>%
  group_by(Site, Dose,
           ApptYear, WeekNum, ApptDate, DOW) %>%
  summarize(`Pod Type` = "All",
            TotalArrivals = n(),
            WalkIns = sum(`Same Day?` == "Same day"),
            WalkInPercent = WalkIns / TotalArrivals * 100)

dose1_7day_walkins_totals <- dose1_7day_walkins_totals[ , colnames(dose1_7days_walkins_type)]

dose1_7day_walkins_summary <- rbind(dose1_7days_walkins_type, dose1_7day_walkins_totals)

dose1_7day_walkins_summary <- dose1_7day_walkins_summary %>%
  ungroup()

dose1_7day_walkins_summary <- dose1_7day_walkins_summary %>%
  mutate(Site = factor(Site, levels = sites, ordered = TRUE),
         `Pod Type` = factor(`Pod Type`, levels = pod_type, ordered = TRUE))

dose1_7day_walkins_avg <- dose1_7day_walkins_summary %>%
  group_by(Site, Dose, `Pod Type`) %>%
  summarize(AvgWalkInVolume = mean(WalkIns),
            AvgWalkInPercent = mean(WalkInPercent)) %>%
  ungroup()

cast_dose1_7day_walkins_volume <- dcast(dose1_7day_walkins_summary,
                                         Site + `Pod Type` ~ ApptDate, value.var = "WalkIns")

cast_dose1_7day_all_volume <- dcast(dose1_7day_walkins_summary,
                                        Site + `Pod Type` ~ ApptDate, value.var = "TotalArrivals")

cast_dose1_7day_walkins_percent <- dcast(dose1_7day_walkins_summary,
                                        Site + `Pod Type` ~ ApptDate, value.var = "WalkInPercent")

cast_dose1_7day_walkins_volume_summary <- left_join(cast_dose1_7day_walkins_volume, 
                                             dose1_7day_walkins_avg[ , c("Site", "Pod Type", "AvgWalkInVolume")],
                                             by = c("Site" = "Site", "Pod Type" = "Pod Type"))

cast_dose1_7day_walkins_percent_summary <- left_join(cast_dose1_7day_walkins_percent, 
                                                    dose1_7day_walkins_avg[ , c("Site", "Pod Type", "AvgWalkInPercent")],
                                                    by = c("Site" = "Site", "Pod Type" = "Pod Type"))

walkins_list <- list("WalkInVolume" = cast_dose1_7day_walkins_volume,
                     "TotalArrivals" = cast_dose1_7day_all_volume,
                     "WalkInVolumeSummary" = cast_dose1_7day_walkins_volume_summary,
                     "WalkInPercentSummary" = cast_dose1_7day_walkins_percent_summary)

# Export walk in data to excel file
write_xlsx(walkins_list, path = paste0(user_directory, "/WalkIns Last 7 Days as of ", today, ".xlsx"))

