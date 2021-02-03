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


# Import raw data from Epic
if (update_repo) {
  if (initial_run) {
    raw_df <- read_excel(choose.files(default = paste0(user_directory, "/ScheduleData/*.*"), caption = "Select initial Epic schedule"), 
                         col_names = TRUE, na = c("", "NA"))
  } else {
    # sched_repo <- read_excel(choose.files(default = user_path, caption = "Select schedule repository"), 
    #                       col_names = TRUE, na = c("", "NA"))
    sched_repo <- readRDS(choose.files(default = paste0(user_directory, "/R_Sched_AM_Repo/*.*"), caption = "Select schedule repository"))
    
    # Convert appointment date in schedule repository from posixct to date
    sched_repo <- sched_repo %>%
      mutate(ApptDate = date(ApptDate))
    
    if (num_new_reports == 1) {
      raw_df <- read_excel(choose.files(default = paste0(user_directory, "/ScheduleData/*.*"), caption = "Select current Epic schedule"), 
                           col_names = TRUE, na = c("", "NA"))
    } else if (num_new_reports == 2) {
      raw_df1 <- read_excel(choose.files(default = paste0(user_directory, "/ScheduleData/*.*"), caption = "Select 1st new Epic schedule"), 
                            col_names = TRUE, na = c("", "NA"))
      
      raw_df2 <- read_excel(choose.files(default = paste0(user_directory, "/ScheduleData/*.*"), caption = "Select 2nd new Epic schedule"), 
                            col_names = TRUE, na = c("", "NA"))
      
      raw_df <- rbind(raw_df1, raw_df2)
    } else {
      print("Cannot accept more than 2 reports")
      break
    }
    # raw_df <- read_excel(choose.files(default = user_path, caption = "Select current Epic schedule"), 
    #                  col_names = TRUE, na = c("", "NA"))
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
           ApptWeek = week(Date),
           Department = ifelse(str_detect(Department, ","), substr(Department, 1, str_locate(Department, ",") - 1), Department))
  
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
  
  saveRDS(sched_repo, paste0(user_directory, "/R_Sched_AM_Repo/Sched ",
                             format(min(sched_repo$ApptDate), "%m%d%y"), " to ",
                             format(max(sched_repo$ApptDate), "%m%d%y"),
                             " on ", format(Sys.time(), "%m%d%y %H%M"), ".rds"))
  
} else {
  
  sched_repo <- readRDS(choose.files(default = paste0(user_directory, "/R_Sched_AM_Repo/*.*"), caption = "Select schedule repository"))
  
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
         # Modify status 2 for running report late in evening instead of early morning
         # Status2 = ifelse(ApptDate == today & Status == "Arr", "Arr", # Categorize any arrivals from today as scheduled for easier manipulation
         #                  ifelse(ApptDate <= today & Status == "Sch", "No Show", Status)), # Categorize any appts still in sch status from prior days as no shows
         WeekNum = format(ApptDate, "%U"),
         DOW = weekdays(ApptDate),
         NYZip = substr(`ZIP Code`, 1, 5) %in% ny_zips$zipcode)

# Summarize schedule data for possible stratifications and export
sched_summary <- sched_to_date %>%
  group_by(Site, `Pod Type`, Dose, ApptDate, NYZip, Status2) %>%
  summarize(Count = n())

# Summarize schedule breakdown for next 2 weeks and export ------------------------------
sched_breakdown <- sched_to_date %>%
  filter(Status2 %in% c("Arr", "Sch") & ApptDate >= (today - 1) & ApptDate <= (today + 14)) %>%
  group_by(Dose, `Pod Type`, Site, Status2, ApptDate) %>%
  summarize(Count = n()) %>%
  ungroup()

sched_breakdown <- sched_breakdown[order(sched_breakdown$Dose,
                                         sched_breakdown$`Pod Type`,
                                         sched_breakdown$Site), ]

sched_breakdown_cast <- dcast(sched_breakdown,
                              Dose + `Pod Type` + Site ~ Status2 + ApptDate,
                              value.var = "Count")

# Summarize schedule for next vaccine inventory cycle for daily schedule target analysis ----------------------
# Determine dates in this inventory cycle based on today's date and Tuesday of the next week
sched_inv_cycle_dates <- seq(as.Date("2/16/21", format = "%m/%d/%y"), as.Date("2/23/21", format = "%m/%d/%y"), by = 1)

sched_inv_cycle_dates_rep <- rep(sched_inv_cycle_dates, length(unique(sched_to_date$Site)))

# Determine schedule for each site within this inventory cycle's date range
sched_inv_cycle_pod_type_site <- sched_to_date %>%
  filter(ApptDate %in% sched_inv_cycle_dates & Status2 == "Sch") %>%
  group_by(Site, Dose, `Pod Type`, ApptDate) %>%
  summarize(Count = n()) %>%
  ungroup()

sched_inv_cycle_all_pods_site <- sched_to_date %>%
  filter(ApptDate %in% sched_inv_cycle_dates & Status2 == "Sch") %>%
  group_by(Site, Dose, ApptDate) %>%
  summarize(`Pod Type` = "Total",
            Count = n()) %>%
  ungroup()

sched_inv_cycle_all_pods_site <- sched_inv_cycle_all_pods_site[ , colnames(sched_inv_cycle_pod_type_site)]

sched_inv_cycle_pod_order <- c("Total", "Employee", "Patient")

sched_inv_cycle_site_summary <- rbind(sched_inv_cycle_pod_type_site, sched_inv_cycle_all_pods_site)

sched_inv_cycle_site_summary <- sched_inv_cycle_site_summary %>%
  mutate(`Pod Type` = factor(`Pod Type`, levels = sched_inv_cycle_pod_order, ordered = TRUE)) %>%
  arrange(Site, ApptDate, `Pod Type`, Dose)

sched_inv_cycle_site_summary_cast <- dcast(sched_inv_cycle_site_summary,
                                           Site + ApptDate ~ Dose + `Pod Type`,
                                           value.var = "Count")


sched_inv_cycle_sites <- sort(rep(unique(sched_to_date$Site), length(sched_inv_cycle_dates)))

sched_inv_cycle_all_sites_dates <- data.frame("Site" = sched_inv_cycle_sites,
                                              "ApptDate" = sched_inv_cycle_dates_rep,
                                              stringsAsFactors = FALSE)

sched_inv_cycle_sites_herds <- left_join(sched_inv_cycle_all_sites_dates, sched_inv_cycle_site_summary_cast,
                                      by = c("Site" = "Site",
                                             "ApptDate" = "ApptDate"))

# Determine schedule across system within this inventory cycle's date range
sched_inv_cycle_pod_type_sys <- sched_to_date %>%
  filter(ApptDate %in% sched_inv_cycle_dates & Status2 == "Sch") %>%
  group_by(Dose, `Pod Type`, ApptDate) %>%
  summarize(Count = n()) %>%
  ungroup()

sched_inv_cycle_all_pods_sys <- sched_to_date %>%
  filter(ApptDate %in% sched_inv_cycle_dates & Status2 == "Sch") %>%
  group_by(Dose, ApptDate) %>%
  summarize(`Pod Type` = "Total",
            Count = n()) %>%
  ungroup()

sched_inv_cycle_all_pods_sys <- sched_inv_cycle_all_pods_sys[ , colnames(sched_inv_cycle_pod_type_sys)]

sched_inv_cycle_sys_summary <- rbind(sched_inv_cycle_pod_type_sys, sched_inv_cycle_all_pods_sys)

sched_inv_cycle_sys_summary <- sched_inv_cycle_sys_summary %>%
  mutate(`Pod Type` = factor(`Pod Type`, levels = sched_inv_cycle_pod_order, ordered = TRUE),
         DOW = wday(ApptDate, label = TRUE, abbr = TRUE)) %>%
  arrange(ApptDate, `Pod Type`, Dose)

sched_inv_cycle_sys_summary_cast <- dcast(sched_inv_cycle_sys_summary,
                                          DOW + ApptDate ~ Dose + `Pod Type`,
                                          value.var = "Count")

sched_inv_cycle_sys_summary_cast <- sched_inv_cycle_sys_summary_cast %>%
  arrange(ApptDate)

sched_inv_cycle_sys_herds <- sched_inv_cycle_sys_summary_cast


# Export key data tables to Excel for reporting ----------------------------------------
# Export schedule summary, schedule breakdown, and cumulative administed doses to excel file
export_list <- list("SchedSummary" = sched_summary,
                    "SchedBreakdown" = sched_breakdown_cast,
                    "Sched_InvCycle_Site" = sched_inv_cycle_sites_herds,
                    "Sched_InvCycle_Sys" = sched_inv_cycle_sys_herds)

write_xlsx(export_list, path = paste0(user_directory, 
                                      "/AdHoc/HERDS Sched Data Export ", 
                                      format(Sys.time(), "%m%d%y %H%M"), ".xlsx"))

