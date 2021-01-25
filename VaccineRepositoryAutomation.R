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
update_repo <- FALSE

# Reports to add to repo
num_new_reports <- 2

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

# Summarize schedule for next 7 days for daily schedule target analysis ----------------------
sched_7days_pod_type <- sched_to_date %>%
  filter(ApptDate > today & Status2 == "Sch") %>%
  group_by(Site, Dose, `Pod Type`, ApptDate) %>%
  summarize(Count = n()) %>%
  ungroup()

sched_7days_all_pods <- sched_to_date %>%
  filter(ApptDate > today & Status2 == "Sch") %>%
  group_by(Site, Dose, ApptDate) %>%
  summarize(`Pod Type` = "Total",
            Count = n()) %>%
  ungroup()

sched_7days_all_pods <- sched_7days_all_pods[ , colnames(sched_7days_pod_type)]

sched_7day_pod_order <- c("Total", "Employee", "Patient")

sched_7days_summary <- rbind(sched_7days_pod_type, sched_7days_all_pods)

sched_7days_summary <- sched_7days_summary %>%
  mutate(Site = factor(Site, levels = sites, ordered = TRUE),
         `Pod Type` = factor(`Pod Type`, levels = sched_7day_pod_order, ordered = TRUE)) %>%
  arrange(Site, ApptDate, `Pod Type`, Dose)

sched_7days_summary_cast <- dcast(sched_7days_summary,
                                  Site + ApptDate ~ Dose + `Pod Type`,
                                  value.var = "Count")

sched_7days_dates <- rep(seq(min(sched_7days_summary$ApptDate),
                             min(sched_7days_summary$ApptDate) + 6,
                             by = 1), length(unique(sched_7days_summary$Site)))

sched_7days_sites <- sort(rep(unique(sched_7days_summary$Site), 7))

sched_7days_all_sites_dates <- data.frame("Site" = sched_7days_sites,
                                          "ApptDate" = sched_7days_dates)

sched_7days_format_jm <- left_join(sched_7days_all_sites_dates, sched_7days_summary_cast,
                                   by = c("Site" = "Site",
                                          "ApptDate" = "ApptDate"))


# Summarize schedule for next vaccine inventory cycle for daily schedule target analysis ----------------------
# Determine dates in this inventory cycle based on today's date and Tuesday of the next week
sched_inv_cycle_dates <- seq(today, floor_date(today, unit = "week", week_start = 7) + 9, by = 1)

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

sched_inv_cycle_sites_jm <- left_join(sched_inv_cycle_all_sites_dates, sched_inv_cycle_site_summary_cast,
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
  arrange(App)

sched_inv_cycle_sys_jm <- sched_inv_cycle_sys_summary_cast


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



# Determine NYS vs NYS patients based on zip code ----------------------------------
nys_dose1_to_date <- sched_to_date %>%
  filter(Status2 == "Arr" & Dose == 1) %>%
  group_by(Site, `Pod Type`) %>%
  summarize(TotalArrivals = n(),
            NYSArrivals = sum(NYZip),
            PercentInState = percent(NYSArrivals / TotalArrivals)) %>%
  ungroup()

nys_dose1_14days <- sched_to_date %>%
  filter(Status2 == "Sch" & Dose == 1 & ApptDate <= (today + 13)) %>%
  group_by(Site, `Pod Type`) %>%
  summarize(TotalSchedAppts = n(),
            NYSSched = sum(NYZip),
            PercentInState = percent(NYSSched / TotalSchedAppts)) %>%
  ungroup()


# Walk-In stats for last 7 days ------------------------------------------
# Determine daily walk-ins by pod type
dose1_7day_walkins_type <- sched_to_date %>%
  filter(Status == "Arr" & Dose == 1 & ApptDate > (today - 7) & ApptDate <= today) %>%
  group_by(Site, `Pod Type`, Dose,
           ApptYear, WeekNum, ApptDate, DOW) %>%
  summarize(DailyArrivals = n(),
            DailyWalkIns = sum(`Same Day?` == "Same day"),
            DailyWalkInPercent = percent(DailyWalkIns / DailyArrivals))

# Determine daily walk-ins for each site
dose1_7day_walkins_totals <- sched_to_date %>%
  filter(Status == "Arr" & Dose == 1 & ApptDate > (today - 7) & ApptDate <= today) %>%
  group_by(Site, Dose,
           ApptYear, WeekNum, ApptDate, DOW) %>%
  summarize(`Pod Type` = "All",
            DailyArrivals = n(),
            DailyWalkIns = sum(`Same Day?` == "Same day"),
            DailyWalkInPercent = percent(DailyWalkIns / DailyArrivals))

# Reorder columns
dose1_7day_walkins_totals <- dose1_7day_walkins_totals[ , colnames(dose1_7day_walkins_type)]

# Combine daily walk-ins for each site by pod type and site total
dose1_7day_walkins_summary <- rbind(dose1_7day_walkins_type, dose1_7day_walkins_totals)

dose1_7day_walkins_summary <- dose1_7day_walkins_summary %>%
  ungroup()

# Reorder rows based on site and pod type
dose1_7day_walkins_summary <- dose1_7day_walkins_summary %>%
  mutate(Site = factor(Site, levels = sites, ordered = TRUE),
         `Pod Type` = factor(`Pod Type`, levels = pod_type, ordered = TRUE))

dose1_7day_walkins_summary <- dose1_7day_walkins_summary %>%
  arrange(Site, `Pod Type`)

# Calculate average walk-in volume and average walk-in percentage
dose1_7day_walkins_avg <- dose1_7day_walkins_summary %>%
  group_by(Site, Dose, `Pod Type`) %>%
  summarize(AvgWalkInVolume = sum(DailyWalkIns) / 7,
            WalkInPercent = percent(sum(DailyWalkIns) / sum(DailyArrivals))) %>%
  ungroup()

cast_dose1_7day_walkins_arr <- dcast(dose1_7day_walkins_summary,
                                        Site + `Pod Type` ~ ApptDate, value.var = "DailyWalkIns")

cast_dose1_7day_all_arr <- dcast(dose1_7day_walkins_summary,
                                    Site + `Pod Type` ~ ApptDate, value.var = "DailyArrivals")

cast_dose1_7day_walkins_stats <- melt(dose1_7day_walkins_summary,
                                      id.vars = c("Site", "Pod Type", "ApptDate"),
                                      measure.vars = c("DailyWalkIns", "DailyArrivals", "DailyWalkInPercent"))

cast_dose1_7day_walkins_stats <- dcast(cast_dose1_7day_walkins_stats,
                                       Site + `Pod Type` ~ variable + ApptDate,
                                       value.var ="value")

cast_dose1_7day_walkins_stats <- left_join(cast_dose1_7day_walkins_stats, 
                                             dose1_7day_walkins_avg[ , c("Site", "Pod Type", "AvgWalkInVolume", "WalkInPercent")],
                                             by = c("Site" = "Site", "Pod Type" = "Pod Type"))

cast_melt_test <- melt(dose1_7day_walkins_summary,
                       id.vars = c("Site", "Pod Type", "ApptDate"),
                       measure.vars = c("DailyWalkIns", "DailyArrivals", "DailyWalkInPercent"))


# Export key data tables to Excel for reporting ----------------------------------------
# Export schedule summary, schedule breakdown, and cumulative administed doses to excel file
export_list <- list("SchedSummary" = sched_summary,
                    "SchedBreakdown" = sched_breakdown_cast,
                    # "Sched_Targets_7Days" = sched_7days_format_jm,
                    "Sched_InvCycle_Site" = sched_inv_cycle_sites_jm,
                    "Sched_InvCycle_Sys" = sched_inv_cycle_sys_jm,
                    "CumDosesAdministered" = admin_doses_table_export,
                    "1stDoseArrivals_7Days" = cast_dose1_7day_all_arr,
                    "1stDoseWalkIns_7Days" = cast_dose1_7day_walkins_arr,
                    "WalkInsStats_7Days" = cast_dose1_7day_walkins_stats)

write_xlsx(export_list, path = paste0(user_directory, 
                                      "/R_Sched_Analysis_AM_Export/Sched Data Export ", 
                                      format(Sys.time(), "%m%d%y %H%M"), ".xlsx"))


# No show analysis -------------------------------------
dose1_noshows_daily <- sched_to_date %>%
  filter(ApptDate < today & Dose == 1) %>%
  group_by(Site, `Pod Type`, ApptDate) %>%
  summarize(NoShow = sum(Status2 == "No Show"),
            Total = sum(Status2 == "No Show") + sum(Status2 == "Arr"),
            NoShowPercent = NoShow / Total)

dose1_noshows_3days <- sched_to_date %>%
  filter(ApptDate < today & ApptDate >= (today - 3) & Dose == 1) %>%
  group_by(Site, `Pod Type`) %>%
  summarize(StartDate = min(ApptDate),
            EndDate = max(ApptDate),
            NoShow = sum(Status2 == "No Show"),
            Total = sum(Status2 == "No Show") + sum(Status2 == "Arr"),
            NoShowPercent = percent(NoShow / Total))


# Analysis of employee and patient pods ------------------------
sched_to_date <- sched_to_date %>%
  mutate(EmployeeSite = ifelse(is.na(`MOUNT SINAI HEALTH SYSTEM`), "Not Employee",
                               ifelse(str_detect(`MOUNT SINAI HEALTH SYSTEM`, " "), 
                                      substr(`MOUNT SINAI HEALTH SYSTEM`, 1, str_locate(`MOUNT SINAI HEALTH SYSTEM`, " ") - 1),
                                      `MOUNT SINAI HEALTH SYSTEM`)),
         ApptClassification = ifelse(EmployeeSite == "Not Employee", "Not MSHS Employee", "MSHS Employee"))

pod_type_breakdown_weekly <- sched_to_date %>%
  filter(ApptDate < today & Dose == 1 & Status2 == "Arr") %>%
  group_by(Site, `Pod Type`,`Provider/Resource`, ApptYear, WeekNum, EmployeeSite, ApptClassification) %>%
  summarize(Count = n())

pod_type_breakdown_total <- sched_to_date %>%
  filter(ApptDate < today & Dose == 1 & Status2 == "Arr") %>%
  group_by(Site, `Pod Type`,`Provider/Resource`, EmployeeSite, ApptClassification) %>%
  summarize(Total = n())

pod_type_breakdown_cast <- dcast(pod_type_breakdown_weekly,
                                 Site + `Pod Type` + `Provider/Resource` + EmployeeSite + ApptClassification ~ ApptYear + WeekNum,
                                 value.var = "Count")

pod_type_breakdown_cast <- left_join(pod_type_breakdown_cast, pod_type_breakdown_total,
                                     by = c("Site" = "Site",
                                            "Pod Type" = "Pod Type",
                                            "Provider/Resource" = "Provider/Resource",
                                            "EmployeeSite" = "EmployeeSite",
                                            "ApptClassification" = "ApptClassification"))

pod_type_appt_type_summary <- pod_type_breakdown_cast %>%
  group_by(Site, `Pod Type`, `Provider/Resource`, ApptClassification) %>%
  summarize(Total = sum(Total))

pod_type_appt_type_summary_cast <- dcast(pod_type_appt_type_summary,
                                         Site + `Pod Type` + `Provider/Resource` ~ ApptClassification,
                                         value.var = "Total")

pod_type_appt_type_summary_cast <- pod_type_appt_type_summary_cast %>%
  mutate(ErrorVolume = ifelse(`Pod Type` == "Employee", `Not MSHS Employee`, `MSHS Employee`),
           PercentApptError = percent(ErrorVolume / (`MSHS Employee` + `Not MSHS Employee`)))

# write_xlsx(pod_type_appt_type_summary_cast, paste0(user_directory, "/Pod and Appt Type Conflicts ", Sys.Date(),".xlsx"))
# 
# 
# employee_test <- sched_to_date %>%
#   filter(Status2 == "Arr" & !is.na(`MOUNT SINAI HEALTH SYSTEM`) & Dose == 1) 
# 
# test <- sched_to_date %>%
#   filter(Status2 == "Arr" & Dose == 1) %>%
#   group_by(Site, `Pod Type`, `MOUNT SINAI HEALTH SYSTEM`, ApptClassification) %>%
#   summarize(Count = n())


# Analysis of patient ages
sched_to_date <- sched_to_date %>%
  mutate(Over65 = `Age >= 65 Check` == "Yes")

age_breakdown <- sched_to_date %>%
  filter(Dose == 1 & Status2 == "Sch" & ApptDate >= as.Date("1/27/21", format = "%m/%d/%y") & ApptDate <= as.Date("2/2/21", format = "%m/%d/%y")) %>%
  group_by(Site, `Pod Type`, ApptDate) %>%
  summarize(AgeEligible = sum(Over65),
            NotAgeEligible = sum(!Over65),
            Total = n(),
            Over65Percent = percent(AgeEligible / Total),
            Under65Percent = percent(NotAgeEligible / Total))

age_breakdown_summary <- sched_to_date %>%
  filter(Dose == 1 & Status2 == "Sch" & ApptDate >= as.Date("1/27/21", format = "%m/%d/%y") & ApptDate <= as.Date("2/2/21", format = "%m/%d/%y")) %>%
  group_by(Site, `Pod Type`) %>%
  summarize(AgeEligible = sum(Over65),
            NotAgeEligible = sum(!Over65),
            Total = n(),
            Over65Percent = percent(AgeEligible / Total),
            Under65Percent = percent(NotAgeEligible / Total))

age_breakdown_summary2 <- age_breakdown_summary %>%
  group_by(`Pod Type`) %>%
  summarize(AgeEligible = sum(AgeEligible),
            NotAgeEligible = sum(NotAgeEligible),
            Total = sum(Total))

# Calculate day with highest arrived volume --------------------------------------------
daily_arrivals <- sched_to_date %>%
  filter(Status2 == "Arr") %>%
  group_by(Site, ApptDate) %>%
  summarize(AdminVolume = n())

daily_max <- daily_arrivals %>%
  group_by(Site, ) %>%
  summarize(MaxDailyAdmin = max(AdminVolume),
            DateOfMax = ApptDate[which(AdminVolume == max(AdminVolume))])

# Determine patients with multiple first dose appointments -----------------
dose1_sch_arr <- sched_to_date %>%
  filter(Status2 %in% c("Arr", "Sch") & Dose == 1)


dose1_sch_arr_summary <- dose1_sch_arr %>%
  group_by(Patient, DOB, Dose) %>%
  summarize(Count = n())

mult_dose1 <- dose1_sch_arr_summary %>%
  filter(Count > 1)
