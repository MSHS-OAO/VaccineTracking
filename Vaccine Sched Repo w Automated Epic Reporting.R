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

# Determine whether or not to update walk-in analysis
update_walkins <- FALSE

# Import reference data for site and pod mappings
site_mappings <- read_excel(paste0(user_directory, "/ScheduleData/Automation Ref 2021-02-13.xlsx"), sheet = "Site Mappings")
pod_mappings <- read_excel(paste0(user_directory, "/ScheduleData/Automation Ref 2021-02-13.xlsx"), sheet = "Pod Mappings Simple")

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

# Vaccine manufacturers
mfg <- c("Pfizer", "Moderna")

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
    
    # # Convert appointment date in schedule repository from posixct to date
    # sched_repo <- sched_repo %>%
    #   mutate(ApptDate = date(ApptDate))

    raw_df <- read.csv(choose.files(default = paste0(user_directory, "/Epic Auto Report Sched All Status/*.*"), 
                                    caption = "Select current Epic schedule"), stringsAsFactors = FALSE, 
                       check.names = FALSE)
  }
  
  new_sched <- raw_df
  
  # Remove test patient
  new_sched <- new_sched[new_sched$Patient != "Patient, Test", ]
  
  # Update data classes from .csv schedule import to match repository for binding
  new_sched <- new_sched %>% 
    mutate(DOB = as.Date(DOB, format = "%m/%d/%Y"),
           `Made Date` = as.Date(`Made Date`, format = "%m/%d/%y"),
           Date = as.Date(Date, format = "%m/%d/%Y"),
           `Appt Time` = as.POSIXct(paste("1899-12-31 ", `Appt Time`), tz = "", format = "%Y-%m-%d %H:%M %p"),
           `ZIP Code` = substr(`ZIP Code`, 2, nchar(`ZIP Code`)))
  
  # Determine dates in new report
  current_dates <- sort(unique(new_sched$Date))
  
  if (initial_run) {
    sched_repo <- new_sched
  } else {
    # Update schedule repository by removing duplicate dates and adding data from new report
    sched_repo <- sched_repo %>%
      filter(!(Date >= current_dates[1]))
    sched_repo <- rbind(sched_repo, new_sched)
  }
  
  saveRDS(sched_repo, paste0(user_directory, "/R_Sched_AM_Repo/Auto Epic Report Sched ",
                             format(min(sched_repo$Date), "%m%d%y"), " to ",
                             format(max(sched_repo$Date), "%m%d%y"),
                             " on ", format(Sys.time(), "%m%d%y %H%M"), ".rds"))
  
} else {
  
  sched_repo <- readRDS(choose.files(default = paste0(user_directory, "/R_Sched_AM_Repo/*.*"), caption = "Select schedule repository"))
  
}

# Format and analyze schedule to date for dashboards ---------------------------
sched_to_date <- sched_repo

sched_to_date <- sched_to_date %>%
  mutate(Dose = ifelse(str_detect(Type, "DOSE 1"), 1, ifelse(str_detect(Type, "DOSE 2"), 2, NA)),
         ApptDate = date(Date),
         ApptYear = year(Date),
         ApptMonth = month(Date),
         ApptDay = day(Date),
         ApptWeek = week(Date),
         Department = ifelse(str_detect(Department, ","), substr(Department, 1, str_locate(Department, ",") - 1), Department), 
         Status = ifelse(`Appt Status` %in% c("Arrived", "Comp", "Checked Out"), "Arr", `Appt Status`), # Group various arrival statuses as "Arr"
         # Status2 = ifelse(ApptDate == today & Status == "Arr", "Sch", # Categorize any arrivals from today as scheduled for easier manipulation
         #                  ifelse(ApptDate < today & Status == "Sch", "No Show", Status)), # Categorize any appts still in sch status from prior days as no shows
         # Modify status 2 for running report late in evening instead of early morning
         Status2 = ifelse(ApptDate < (today) & Status == "Sch", "No Show", Status), # Categorize any appts still in sch status from prior days as no shows
         Mfg = ifelse(Status2 != "Arr", NA, #Keep manufacturer as NA if the appointment hasn't been arrived
                      ifelse(is.na(Immunizations), "Pfizer", # Assume any visits without immunization record are Pfizer
                             ifelse(str_detect(Immunizations, "Moderna"), "Moderna", "Pfizer"))),
         WeekNum = format(ApptDate, "%U"),
         DOW = weekdays(ApptDate),
         NYZip = substr(`ZIP Code`, 1, 5) %in% ny_zips$zipcode)

sched_to_date <- left_join(sched_to_date, site_mappings,
                           by = c("Department" = "Department"))

sched_to_date <- left_join(sched_to_date, pod_mappings[ , c("Provider", "Pod Type")],
                           by = c("Provider/Resource" = "Provider"))

sched_to_date[is.na(sched_to_date$`Pod Type`), "Pod Type"] <- "Patient"


# Summarize schedule data for possible stratifications and export
sched_summary <- sched_to_date %>%
  group_by(Site, `Pod Type`, Dose, ApptDate, NYZip, Status2) %>%
  summarize(Count = n())

# Summarize schedule breakdown for next 2 weeks and export ------------------------------
# Format data
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

# Create template dataframe of all sites, dose, and pod type combinations to ensure all are included even if there are no scheduled appointments
sch_breakdown_doses <- data.frame("Dose" = rep(c(1:2), length(unique(sched_to_date$Site))))
sch_breakdown_doses <- sch_breakdown_doses %>%
  arrange(Dose)

sch_breakdown_pods <- data.frame("PodType" = rep(pod_type[1:2], length(unique(sched_to_date$Site))), stringsAsFactors = FALSE)
sch_breakdown_pods <- sch_breakdown_pods %>%
  arrange(PodType)

sch_breakdown_site_doses <- data.frame("Site" = rep(unique(sched_to_date$Site), 2),
                                       "Dose" = sch_breakdown_doses,
                                       stringsAsFactors = FALSE)

sch_breakdown_site_pods <- data.frame("Site" = rep(unique(sched_to_date$Site), 2),
                                      "PodType" = sch_breakdown_pods,
                                      stringsAsFactors = FALSE)

sch_breakdown_site_dose_pod <- left_join(sch_breakdown_site_doses, sch_breakdown_site_pods, 
                                         by = c("Site" = "Site"))

sch_breakdown_site_dose_pod <- sch_breakdown_site_dose_pod %>%
  select(Dose, PodType, Site) %>%
  arrange(Dose, PodType, Site)

# Merge template dataframe with schedule breakdown data
sched_breakdown_cast <- left_join(sch_breakdown_site_dose_pod, sched_breakdown_cast,
                                  by = c("Dose" = "Dose",
                                         "PodType" = "Pod Type",
                                         "Site" = "Site"))

# Summarize schedule breakdown for next two weeks for Pfizer only -------------
# Assumptions:
# 1. Any scheduled dose 1 appointments will receive Pfizer
# 2. Only scheduled dose 2 appointments with an arrived Moderna dose 1 
# appointment will receive Moderna. Any other scheduled dose 2 appointments will
# receive Pfizer

# Determine MRNs that received Moderna for first dose
dose1_moderna <- sched_to_date %>%
  filter(Status2 %in% c("Arr") & Mfg == "Moderna" & Dose == 1) %>%
  select(MRN)

sched_to_date <- sched_to_date %>%
  #Determine expected manufacturer for scheduled appointment based on above 
  # assumptions
  mutate(ExpSchMfg = ifelse(Status2 != "Sch", NA,
                              ifelse(Dose == 1, "Pfizer",
                                     ifelse(MRN %in% dose1_moderna$MRN,
                                            "Moderna", "Pfizer"))))
pfizer_sched_breakdown <- sched_to_date %>%
  filter((Mfg == "Pfizer" | ExpSchMfg == "Pfizer") & 
           ApptDate >= (today - 1) & ApptDate <= (today + 14)) %>%
  group_by(Dose, `Pod Type`, Site, Status2, ApptDate) %>%
  summarize(Count = n()) %>%
  ungroup()

pfizer_sched_breakdown_cast <- dcast(pfizer_sched_breakdown,
                                     Dose + `Pod Type` + Site ~ Status2 + ApptDate,
                                     value.var = "Count")

# Merge Pfizer schedule with site-dose-pod template
pfizer_sched_breakdown_cast <- left_join(sch_breakdown_site_dose_pod, 
                                         pfizer_sched_breakdown_cast,
                                         by = c("Dose" = "Dose",
                                                "PodType" = "Pod Type",
                                                "Site" = "Site"))

# Summarize schedule for next vaccine inventory cycle for daily schedule target analysis ----------------------
# Determine dates in this inventory cycle based on today's date and Tuesday of the next week
sched_inv_cycle_dates <- seq(today, floor_date(today, unit = "week", week_start = 7) + 16, by = 1)

# Create template for all site and date combinations
sched_inv_cycle_dates_rep <- rep(sched_inv_cycle_dates, length(unique(sched_to_date$Site)))

sched_inv_cycle_sites <- sort(rep(unique(sched_to_date$Site), length(sched_inv_cycle_dates)))

sched_inv_cycle_all_sites_dates <- data.frame("Site" = sched_inv_cycle_sites,
                                              "ApptDate" = sched_inv_cycle_dates_rep,
                                              stringsAsFactors = FALSE)

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
  arrange(ApptDate)

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

# Stratify administered doses to date by manufacturer -------------------------------------------
# Create daily summary of administered doses by manufacturer
admin_mfg_summary <- sched_to_date %>%
  filter(Status2 == "Arr") %>%
  group_by(Site, Mfg, Dose, ApptDate, Status2) %>%
  summarize(Count = n())

# Create summary of administered doses prior to yesterday and yesterday
admin_prior_yest_site <- sched_to_date %>%
  filter(Status2 == "Arr" & ApptDate <= (today - 2)) %>%
  group_by(Site, Mfg, Dose) %>%
  summarize(DateRange = paste0("Admin 12/15/20-", format(today - 2, "%m/%d/%y")), 
            Count = n())

admin_yest_site <- sched_to_date %>%
  filter(Status2 == "Arr" & ApptDate == (today - 1)) %>%
  group_by(Site, Mfg, Dose) %>%
  summarize(DateRange = paste0("Admin ", format(today - 1, "%m/%d/%y")),
            Count = n())

# admin_prior_yest_system <- sched_to_date %>%
#   filter(Status2 == "Arr" & ApptDate <= (today - 2)) %>%
#   group_by(Mfg, Dose) %>%
#   summarize(Site = "MSHS", 
#             DateRange = paste0("Admin 12/15/20-", format(today - 2, "%m/%d/%y")), 
#             Count = n())
# 
# admin_yest_system <- sched_to_date %>%
#   filter(Status2 == "Arr" & ApptDate == (today - 1)) %>%
#   group_by(Mfg, Dose) %>%
#   summarize(Site = "MSHS", 
#             DateRange = paste0("Admin ", format(today - 1, "%m/%d/%y")),
#             Count = n())
# 
# admin_prior_yest_system <- admin_prior_yest_system[ , colnames(admin_prior_yest_site)]
# admin_yest_system <- admin_yest_system[ , colnames(admin_yest_site)]

admin_to_date_mfg <- rbind(admin_prior_yest_site, admin_yest_site)
# admin_prior_yest_system, admin_yest_system)

admin_to_date_mfg <- admin_to_date_mfg %>%
  ungroup() %>%
  # mutate(Site = factor(Site, levels = sites, ordered = TRUE),
  #        Mfg = factor(Mfg, levels = mfg, ordered = TRUE)) %>%
  arrange(Site, Dose, desc(Mfg), desc(DateRange))

admin_to_date_mfg <- admin_to_date_mfg[ , c("Dose", "Mfg", "Site", "DateRange", "Count")]

admin_to_date_mfg <- admin_to_date_mfg %>%
  mutate(DateRange = factor(DateRange, level = unique(DateRange), ordered = TRUE))

admin_to_date_mfg_cast <- dcast(admin_to_date_mfg, 
                                Dose + Mfg + Site ~ DateRange,
                                value.var ="Count")

# Create template dataframe of sites, doses, and manufacturers to ensure all combinations are included
rep_doses <- data.frame("Dose" = rep(c(1:2), length(unique(sched_to_date$Site))))
rep_doses <- rep_doses %>%
  arrange(Dose)

rep_mfg <- data.frame("Mfg" = rep(mfg, length(unique(sched_to_date$Site))), stringsAsFactors = FALSE)
rep_mfg <- rep_mfg %>%
  arrange(desc(Mfg))

rep_sites_mfg <- data.frame("Site" = rep(unique(sched_to_date$Site), length(mfg)), "Mfg" = rep_mfg, stringsAsFactors = FALSE)

rep_sites_doses <- data.frame("Site" = rep(unique(sched_to_date$Site), 2), "Dose" = rep_doses, stringsAsFactors = FALSE)

# Combine template with administration data
sites_doses_mfg <- left_join(rep_sites_mfg, rep_sites_doses,
                             by = c("Site" = "Site"))

sites_doses_mfg <- sites_doses_mfg %>%
  arrange(Dose, desc(Mfg), Site)

sites_doses_mfg <- sites_doses_mfg[ , c("Dose", "Mfg", "Site")]

admin_to_date_mfg_export <- left_join(sites_doses_mfg, admin_to_date_mfg_cast, 
                                      by = c("Dose" = "Dose", "Mfg" = "Mfg", "Site" = "Site"))


# Walk-In stats for last 7 days ------------------------------------------
# Determine daily walk-ins by pod type
dose1_7day_walkins_type <- sched_to_date %>%
  filter(Status2 == "Arr" & Dose == 1 & ApptDate >= (today - 7) & ApptDate <= today) %>%
  group_by(Site, `Pod Type`, Dose,
           ApptYear, WeekNum, ApptDate, DOW) %>%
  summarize(DailyArrivals = n(),
            DailyWalkIns = sum(`Same Day?` == "Same day"),
            DailyWalkInPercent = percent(DailyWalkIns / DailyArrivals))

# Determine daily walk-ins for each site
dose1_7day_walkins_totals <- sched_to_date %>%
  filter(Status2 == "Arr" & Dose == 1 & ApptDate >= (today - 7) & ApptDate <= today) %>%
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
  summarize(AvgWalkInVolume = sum(DailyWalkIns) / sum(DailyWalkIns > 0),
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
                    "Pfizer_SchedBreakdown" = pfizer_sched_breakdown_cast,
                    # "Sched_Targets_7Days" = sched_7days_format_jm,
                    "Sched_InvCycle_Site" = sched_inv_cycle_sites_jm,
                    "Sched_InvCycle_Sys" = sched_inv_cycle_sys_jm,
                    "CumDosesAdministered" = admin_doses_table_export,
                    "AdminMfgSummary" = admin_mfg_summary,
                    "DosesAdminMfg" = admin_to_date_mfg_export,
                    "1stDoseArrivals_7Days" = cast_dose1_7day_all_arr,
                    "1stDoseWalkIns_7Days" = cast_dose1_7day_walkins_arr,
                    "WalkInsStats_7Days" = cast_dose1_7day_walkins_stats)

write_xlsx(export_list, path = paste0(user_directory, 
                                      "/R_Sched_Analysis_AM_Export/Auto Epic Report Sched Data Export ", 
                                      format(Sys.time(), "%m%d%y %H%M"), ".xlsx"))

# Test line