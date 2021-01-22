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
  user_directory <- "J:/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/System Operations/COVID Vaccine/ScheduleData/Automation Files"
} else {
  user_directory <- "J:/deans/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/System Operations/COVID Vaccine/ScheduleData/Automation Files"
}

# Set path for file selection
user_path <- paste0(user_directory, "\\*.*")

# Determine whether or not to update an existing repo
update_repo <- TRUE

# Import reference data for site and pod mappings
site_mappings <- read_excel(paste0(user_directory, "/Reference 2021-01-20.xlsx"), sheet = "Site Mappings")
pod_mappings <- read_excel(paste0(user_directory, "/Reference 2021-01-20.xlsx"), sheet = "Pod Mappings Simple")

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

# Import schedule repository (likely from morning analysis)
sched_repo <- readRDS(choose.files(default = user_path, caption = "Select schedule repository"))

# Import new schedule data from Epic
if (update_repo) {
  raw_df <- read_excel(choose.files(default = user_path, caption = "Select this afternoon's Epic schedule"), 
                       col_names = TRUE, na = c("", "NA"))

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
  
  # Update schedule repository by removing duplicate dates and adding data from new report
  sched_repo <- sched_repo %>%
    filter(!(ApptDate >= current_dates[1]))
  
  # Combine repository with updated schedule
  sched_repo <- rbind(sched_repo, new_sched)
  
  # Save new schedule repository
  saveRDS(sched_repo, paste0(user_directory, "/PM Huddle Sched w Repo ",
                             format(min(sched_repo$ApptDate), "%m%d%y"), "-",
                             format(max(sched_repo$ApptDate), "%m%d%y"),
                             " on ", format(Sys.time(), "%m%d%y"), ".rds"))
  
}

# Format and analyze schedule to date for dashboards ---------------------------
sched_to_date <- sched_repo

sched_to_date <- left_join(sched_to_date, site_mappings,
                           by = c("Department" = "Department"))

sched_to_date <- left_join(sched_to_date, pod_mappings[ , c("Provider", "Pod Type")],
                           by = c("Provider/Resource" = "Provider"))

sched_to_date[is.na(sched_to_date$`Pod Type`), "Pod Type"] <- "Patient"

sched_to_date <- sched_to_date %>%
  mutate(Status = ifelse(`Appt Status` %in% c("Arrived", "Comp", "Checked Out"), "Arr", `Appt Status`), # Group various arrival statuses as "Arr"
         Status2 = ifelse(ApptDate == today & Status == "Arr", "Sch", # Categorize any arrivals from today as scheduled for easier manipulation
                          ifelse(ApptDate < today & Status == "Sch", "No Show", Status)), # Categorize any appts still in sch status from prior days as no shows
         WeekNum = format(ApptDate, "%U"),
         AdjYear = ifelse(ApptYear == 2021 & WeekNum == "00", 2020, ApptYear),
         AdjWeekNum = ifelse(ApptYear == 2021 & WeekNum == "00", 52, as.integer(WeekNum)),
         DOW = weekdays(ApptDate))

WeekDates <- sched_to_date %>%
  group_by(AdjYear, AdjWeekNum) %>%
  arrange(AdjYear, AdjWeekNum) %>%
  summarize(StartDate = min(ApptDate),
            EndDate = max(ApptDate))

WeekDates$VaccWk <- row.names(WeekDates)

# Determine week of vaccine distribution for each appointment
sched_to_date <- left_join(sched_to_date, WeekDates[ , c("AdjYear", "AdjWeekNum", "VaccWk")],
                           by = c("AdjYear" = "AdjYear", "AdjWeekNum" = "AdjWeekNum"))


# Summarize schedule data for possible stratifications and export
sched_summary <- sched_to_date %>%
  group_by(Site, `Pod Type`, Dose, ApptDate, Status2) %>%
  summarize(Count = n())


# Summarize schedule breakdown for next 2 weeks and export ------------------------------
sched_breakdown <- sched_to_date %>%
  filter(Status2 %in% c("Arr", "Sch") & ApptDate >= (today - 1) & ApptDate <= (today + 13)) %>%
  group_by(Dose, `Pod Type`, Site, Status2, ApptDate) %>%
  summarize(Count = n()) %>%
  ungroup()

sched_breakdown <- sched_breakdown[order(sched_breakdown$Dose,
                                         sched_breakdown$`Pod Type`,
                                         sched_breakdown$Site), ]

sched_breakdown_cast <- dcast(sched_breakdown,
                              Dose + `Pod Type` + Site ~ Status2 + ApptDate,
                              value.var = "Count")

# Summarize doses administered to date --------------------------------
# Summarize administered doses prior to today and stratify by site and pod type
admin_doses_site_pod_type <- sched_to_date %>%
  filter(Status2 == "Arr") %>%
  group_by(Site, `Pod Type`, Dose) %>%
  summarize(DateRange = paste0(format(min(ApptDate), "%m/%d/%y"), "-", format(max(ApptDate), "%m/%d/%y")),
            Count = n()) %>%
  ungroup()

# Summarize administered doses prior to today and stratify only by site
admin_doses_site_all_pods <- sched_to_date %>%
  filter(Status2 == "Arr") %>%
  group_by(Site, Dose) %>%
  summarize("Pod Type" = "All",
            DateRange = paste0(format(min(ApptDate), "%m/%d/%y"), "-", format(max(ApptDate), "%m/%d/%y")),
            Count = n()) %>%
  ungroup()

admin_doses_site_all_pods <- admin_doses_site_all_pods[ , colnames(admin_doses_site_pod_type)]

# Summarize administered doses across system and stratify by pod type
admin_doses_system_pod_type <- sched_to_date %>%
  filter(Status2 == "Arr") %>%
  group_by(`Pod Type`, Dose) %>%
  summarize(Site = "MSHS", DateRange = paste0(format(min(ApptDate), "%m/%d/%y"), "-", format(max(ApptDate), "%m/%d/%y")),
            Count = n()) %>%
  ungroup()

admin_doses_system_pod_type <- admin_doses_system_pod_type[ , colnames(admin_doses_site_pod_type)]

# Summarize administered doses across system
admin_doses_system_all_pods <- sched_to_date %>%
  filter(Status2 == "Arr") %>%
  group_by(Dose) %>%
  summarize(Site = "MSHS", 
            `Pod Type` = "All",
            DateRange = paste0(format(min(ApptDate), "%m/%d/%y"), "-", format(max(ApptDate), "%m/%d/%y")),
            Count = n()) %>%
  ungroup()

admin_doses_system_all_pods <- admin_doses_system_all_pods[ , colnames(admin_doses_site_pod_type)]

# Combine administered data by pod type, across pods, and across system
admin_doses_summary <- rbind(admin_doses_site_pod_type, 
                             admin_doses_site_all_pods,
                             admin_doses_system_pod_type,
                             admin_doses_system_all_pods)

admin_doses_summary <- admin_doses_summary %>%
  mutate(Dose = ifelse(Dose == 1, "Dose1", "Dose2"))

admin_2nd_dose_remaining <- dcast(admin_doses_summary,
                                  Site + `Pod Type` ~ Dose, value.var = "Count")


admin_2nd_dose_remaining <- admin_2nd_dose_remaining %>%
  mutate(RemainingDose2 = Dose1 - Dose2,
         Site = factor(Site, levels = sites, ordered = TRUE),
         `Pod Type` = factor(`Pod Type`, levels = pod_type, ordered = TRUE))

admin_2nd_dose_remaining <- admin_2nd_dose_remaining[order(admin_2nd_dose_remaining$Site, 
                                                           admin_2nd_dose_remaining$`Pod Type`), ]

admin_doses_table_export <- melt(admin_2nd_dose_remaining, 
                                 id = c("Site", "Pod Type"))

admin_doses_table_export <- dcast(admin_doses_table_export,
                                  Site ~ variable + `Pod Type`, value.var = "value")