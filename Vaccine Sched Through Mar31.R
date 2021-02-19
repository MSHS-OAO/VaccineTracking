# Code to determine patients with dose 2 scheduled and no dose 1 visits

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

# Import reference data for site and pod mappings
site_mappings <- read_excel(paste0(user_directory, "/ScheduleData/Automation Ref 2021-02-13.xlsx"), sheet = "Site Mappings")
pod_mappings <- read_excel(paste0(user_directory, "/ScheduleData/Automation Ref 2021-02-13.xlsx"), sheet = "Pod Mappings Simple")

# Store today's date
today <- Sys.Date() # Update system date since initial analysis was completed yesterday

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

# Import schedule repository
sched_repo <- readRDS(choose.files(default = paste0(user_directory, "/R_Sched_AM_Repo/*.*"), caption = "Select schedule repository"))

# Import future schedule
future_sched <- read_excel(choose.files(default = paste0(user_directory, "/ScheduleData/*.*"), caption = "Select future schedule"))

# Format and combine schedule repository and future schedule
sched_repo <- sched_repo %>%
  mutate(`Canc Date` = as.Date(`Canc Date`, format = "%m/%d/%Y"))

comb_sched <- rbind(sched_repo, future_sched)

sched_hist_outlook <- comb_sched

sched_hist_outlook <- sched_hist_outlook %>%
  mutate(Dose = ifelse(str_detect(Type, "DOSE 1"), 1, ifelse(str_detect(Type, "DOSE 2"), 2, NA)),
         ApptDate = date(Date),
         ApptYear = year(Date),
         ApptMonth = month(Date),
         ApptDay = day(Date),
         ApptWeek = week(Date),
         Department = ifelse(str_detect(Department, ","), substr(Department, 1, str_locate(Department, ",") - 1), Department), 
         Status = ifelse(`Appt Status` %in% c("Arrived", "Comp", "Checked Out"), "Arr", `Appt Status`), # Group various arrival statuses as "Arr"
         Status2 = ifelse(ApptDate == (today) & Status == "Arr", "Sch", # Categorize any arrivals from today as scheduled for easier manipulation; using "today - 1" since this is 1 day behind
                          ifelse(ApptDate < (today) & Status == "Sch", "No Show", Status)), # Categorize any appts still in sch status from prior days as no shows; using "today - 1" since this is 1 day behind
                                 # ifelse(Status == "Arr" & (is.na(`Level Of Service`) | str_detect(`Level Of Service`, "ERRONEOUS")), "Left", Status))), # Correct patient with Arr status but incomplete visit
         # Modify status 2 for running report late in evening instead of early morning
         # Status2 = ifelse(ApptDate == today & Status == "Arr", "Arr", # Categorize any arrivals from today as scheduled for easier manipulation
         #                  ifelse(ApptDate <= today & Status == "Sch", "No Show", Status)), # Categorize any appts still in sch status from prior days as no shows
         Mfg = ifelse(Status2 != "Arr", NA, #Keep manufacturer as NA if the appointment hasn't been arrived
                      ifelse(is.na(Immunizations), "Pfizer", # Assume any visits without immunization record are Pfizer
                             ifelse(str_detect(Immunizations, "Moderna"), "Moderna", "Pfizer"))),
         WeekNum = as.integer(format(ApptDate, "%U")),
         DOW = weekdays(ApptDate),
         NYZip = substr(`ZIP Code`, 1, 5) %in% ny_zips$zipcode,
         # Create unique identifier with patient's last name, first initial of first name, and DOB since there are many patients with duplicate MRNs
         NameDOBID = ifelse(!(str_detect(Patient, ",\\s[A-Z|a-z]+")), paste(Patient, month(DOB), day(DOB), year(DOB)),
                            paste(str_extract(Patient, ".+,\\s[A-Z|a-z]{1}"), month(DOB), day(DOB), year(DOB))), 
         Name = substr(NameDOBID, 1, str_locate(NameDOBID, "[0-9]+\\s[0-9]+\\s[0-9]+") - 1),
         DOBDate = as.Date(paste0(month(DOB), "/", day(DOB), "/", year(DOB)), format = "%m/%d/%Y"),
         ScheduledBy = ifelse(str_detect(`Entry Person`, "ZOCDOC"), "ZocDoc",
                              ifelse(str_detect(`Entry Person`, "MYCHART"), "MyChart", "Employee")))

sched_hist_outlook <- left_join(sched_hist_outlook, site_mappings,
                                by = c("Department" = "Department"))

sched_hist_outlook <- left_join(sched_hist_outlook, pod_mappings[ , c("Provider", "Pod Type")],
                                by = c("Provider/Resource" = "Provider"))

sched_hist_outlook[is.na(sched_hist_outlook$`Pod Type`), "Pod Type"] <- "Patient"

# Create dataframe with week numbers and dates for each week
week_num <- sched_hist_outlook %>%
  group_by(WeekNum) %>%
  summarize(StartDate = min(ApptDate),
            EndDate = max(ApptDate),
            Dates = paste0(format(StartDate, "%m/%d/%y"), "-", format(EndDate, "%m/%d/%y")))

sched_hist_outlook <- left_join(sched_hist_outlook, week_num[ , c("WeekNum", "Dates")],
                               by = c("WeekNum" = "WeekNum"))

# Summarize schedule data for possible stratifications and export
sched_summary_site <- sched_hist_outlook %>%
  group_by(Site, 
           `Pod Type`, 
           Dose, 
           ApptMonth, 
           WeekNum, 
           Dates, 
           ApptDate, 
           Status2) %>%
  summarize(Count = n()) %>%
  ungroup()

sched_summary_system <-  sched_hist_outlook %>%
  group_by(`Pod Type`, 
           Dose, 
           ApptMonth, 
           WeekNum, 
           Dates, 
           ApptDate, 
           Status2) %>%
  summarize(Site = "MSHS",
            Count = n()) %>%
  ungroup()

sched_summary_system <- sched_summary_system[ , colnames(sched_summary_site)]

sched_summary <- rbind(sched_summary_site, sched_summary_system)

sched_status <- sched_summary %>%
  filter(Status2 == "Sch")

sched_status_totals <- sched_status %>%
  group_by(Site, 
           Dose, 
           `Pod Type`) %>%
  summarize(Total = sum(Count))

sched_status_cast <- dcast(sched_status,
                           Site + Dose + `Pod Type` ~ WeekNum + Dates,
                           fun.aggregate = sum,
                           value.var = "Count")

sched_status_cast <- left_join(sched_status_cast, sched_status_totals,
                               by = c("Site" = "Site",
                                      "Dose" = "Dose",
                                      "Pod Type" = "Pod Type"))


# Create template dataframe
site_dose <- data.frame("Site" = (sort(rep(sites, 
                                          length(unique(sched_hist_outlook$Dose))))),
                        "Dose" = rep(unique(sched_hist_outlook$Dose), 
                                     length(sites)),
                        stringsAsFactors = FALSE)

site_pod <- data.frame("Site" = sort(rep(sites, 
                                         length(unique(sched_hist_outlook$`Pod Type`)))),
                       "Pod" = rep(unique(sched_hist_outlook$`Pod Type`), 
                                   length(sites)),
                       stringsAsFactors = FALSE)

site_dose_pod <- left_join(site_dose, site_pod,
                           by = c("Site" = "Site"))

sched_status_cast <- left_join(site_dose_pod, sched_status_cast,
                               by = c("Site" = "Site",
                                      "Dose" = "Dose",
                                      "Pod" = "Pod Type"))

sched_status_cast <- sched_status_cast %>%
  mutate(Site = factor(Site, levels = sites, ordered = TRUE),
         Pod = factor(Pod, levels = pod_type, ordered = TRUE)) %>%
  arrange(Site, Dose, Pod)
                        

write_xlsx(sched_status_cast, path = paste0(user_directory, "/AdHoc/Vaccine Schedule 021921 to 033121.xlsx"))
