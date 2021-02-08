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
site_mappings <- read_excel(paste0(user_directory, "/ScheduleData/Automation Ref 2021-02-05.xlsx"), sheet = "Site Mappings")
pod_mappings <- read_excel(paste0(user_directory, "/ScheduleData/Automation Ref 2021-02-05.xlsx"), sheet = "Pod Mappings Simple")

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

# Import existing schedule repository
sched_repo <- readRDS(choose.files(default = paste0(user_directory, "/R_Sched_AM_Repo/*.*"), caption = "Select schedule repository"))

num_reports <- 2

# Import schedule for next few weeks
if (num_reports == 2) {
  raw_df_1 <- read_excel(choose.files(default = paste0(user_directory, "/ScheduleData/*.*"), caption = "Select first Epic schedule"), 
                       col_names = TRUE, na = c("", "NA"))
  raw_df_2 <- read_excel(choose.files(default = paste0(user_directory, "/ScheduleData/*.*"), caption = "Select second Epic schedule"), 
                         col_names = TRUE, na = c("", "NA"))
  raw_df <- rbind(raw_df_1, raw_df_2)
} else if (num_reports == 1) {
  raw_df <- read_excel(choose.files(default = paste0(user_directory, "/ScheduleData/*.*"), caption = "Select current Epic schedule"), 
                       col_names = TRUE, na = c("", "NA"))
} else {
  raw_df <- NULL
}


# Modify and format columns in Epic schedule to match schedule repository
new_sched <- raw_df %>%
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

# Filter out any redundant dates
sched_to_date <- sched_repo %>%
  filter(!(ApptDate >= current_dates[1]))

# Combine dataframes
sched_hist_outlook <- rbind(sched_to_date, new_sched)

# Continue preprocessing data
sched_hist_outlook <- left_join(sched_hist_outlook, site_mappings,
                           by = c("Department" = "Department"))

sched_hist_outlook <- left_join(sched_hist_outlook, pod_mappings[ , c("Provider", "Pod Type")],
                           by = c("Provider/Resource" = "Provider"))

sched_hist_outlook[is.na(sched_hist_outlook$`Pod Type`), "Pod Type"] <- "Patient"

sched_hist_outlook <- sched_hist_outlook %>%
  mutate(Status = ifelse(`Appt Status` %in% c("Arrived", "Comp", "Checked Out"), "Arr", `Appt Status`), # Group various arrival statuses as "Arr"
         Status2 = ifelse(ApptDate == (today) & Status == "Arr", "Sch", # Categorize any arrivals from today as scheduled for easier manipulation; using "today - 1" since this is 1 day behind
                          ifelse(ApptDate < (today) & Status == "Sch", "No Show", # Categorize any appts still in sch status from prior days as no shows; using "today - 1" since this is 1 day behind
                                 ifelse(Status == "Arr" & (is.na(`Level Of Service`) | str_detect(`Level Of Service`, "ERRONEOUS")), "Left", Status))), # Correct patient with Arr status but incomplete visit
         # Modify status 2 for running report late in evening instead of early morning
         # Status2 = ifelse(ApptDate == today & Status == "Arr", "Arr", # Categorize any arrivals from today as scheduled for easier manipulation
         #                  ifelse(ApptDate <= today & Status == "Sch", "No Show", Status)), # Categorize any appts still in sch status from prior days as no shows
         WeekNum = format(ApptDate, "%U"),
         DOW = weekdays(ApptDate),
         NYZip = substr(`ZIP Code`, 1, 5) %in% ny_zips$zipcode,
         # Create unique identifier with patient's last name, first initial of first name, and DOB since there are many patients with duplicate MRNs
         NameDOBID = ifelse(!(str_detect(Patient, ",\\s[A-Z|a-z]+")), paste(Patient, month(DOB), day(DOB), year(DOB)),
                              paste(str_extract(Patient, "[A-Z|a-z]+,\\s[A-Z|a-z]{1}"), month(DOB), day(DOB), year(DOB))), 
         Name = substr(NameDOBID, 1, str_locate(NameDOBID, "[0-9]+\\s[0-9]+\\s[0-9]+") - 1),
         DOBDate = as.Date(paste0(month(DOB), "/", day(DOB), "/", year(DOB)), format = "%m/%d/%Y"),
         ScheduledBy = ifelse(str_detect(`Entry Person`, "ZOCDOC"), "ZocDoc",
                              ifelse(str_detect(`Entry Person`, "MYCHART"), "MyChart", "Employee")))

# Summarize data by patient name and DOB identifier
pt_appt_summary <- sched_hist_outlook %>%
  group_by(NameDOBID, Name, DOBDate) %>%
  summarize(Arr_Sch_Dose1 = sum(Dose == 1 & Status2 %in% c("Arr", "Sch")),
            NoShow_Canc_Dose1 = sum(Dose == 1 & Status2 %in% c("No Show", "Can")),
            Arr_Dose2 = sum(Dose == 2 & Status2 %in% c("Arr")),
            Sch_Dose2 = sum(Dose == 2 & Status2 %in% c("Sch")),
            Arr_Sch_Dose2 = sum(Dose == 2 & Status2 %in% c("Arr", "Sch")))

dose2_no_dose1 <- pt_appt_summary %>%
  filter(Arr_Sch_Dose1 == 0 & Arr_Sch_Dose2 > 0) %>%
  mutate(Scenario = ifelse(Arr_Dose2 == 0 & NoShow_Canc_Dose1 >= 1 & Sch_Dose2 >= 1, "No Show/Canc Dose 1 & Dose 2 Still On Sched",
                           ifelse(Arr_Dose2 == 1 & Sch_Dose2 == 0, "Received Dose 2 w/o Arr/Sch Dose 1 at MSHS",
                                  ifelse(Arr_Dose2 == 0 & Sch_Dose2 == 1, "Sched Dose 2 w/o Arr/Sch Dose 1 at MSHS",
                                         ifelse(Arr_Sch_Dose2 > 1, "Multiple Arr/Sched Dose 2 - Scheduling Error", "Other")))))

#Summarize dose 2 scenario statistics
dose2_no_dose1_stats <- dose2_no_dose1 %>%
  group_by(Scenario) %>%
  summarize(Count = n())

dose2_no_dose1_stats <- dose2_no_dose1_stats %>%
  mutate(Percent = percent(Count / sum(Count)))

# Pull in additional patient details including MRN, CSN, Site, and Pod Type for each scheduled dose 2 appointment
dose2_mrn_csn_site_pod <- sched_hist_outlook %>%
  filter(Dose == 2 & Status2 %in% c("Arr", "Sch")) %>%
  select(NameDOBID, MRN, CSN, Site, `Pod Type`)

dose2_mrn_csn_site_pod <- unique(dose2_mrn_csn_site_pod)

dose2_no_dose1_pt_detail <- left_join(dose2_no_dose1, dose2_mrn_csn_site_pod,
                                      by = c("NameDOBID" = "NameDOBID"))

# Create flag to identify patients with multiple dose 2 locations
pt_id_dose2_loc <- dose2_mrn_csn_site_pod %>%
  select(NameDOBID, Site)

pt_id_dose2_loc <- unique(pt_id_dose2_loc)

pt_id_dose2_loc <- pt_id_dose2_loc %>%
  mutate(MultDose2Loc = duplicated(NameDOBID))

mult_dose2_loc <- pt_id_dose2_loc %>%
  filter(MultDose2Loc) %>%
  mutate(Site = NULL)

dose2_no_dose1_pt_detail <- left_join(dose2_no_dose1_pt_detail, mult_dose2_loc,
                            by = c("NameDOBID" = "NameDOBID"))

# Create flag to identify patients with multiple charts
pt_id_mrn <- as.data.frame(unique(sched_hist_outlook[ , c("NameDOBID", "MRN")]))

pt_id_mrn <- pt_id_mrn %>%
  mutate(DuplChart = duplicated(NameDOBID))

dupl_charts <- unique(pt_id_mrn %>%
                        filter(DuplChart) %>%
                        mutate(MRN = NULL))

dose2_no_dose1_pt_detail <- left_join(dose2_no_dose1_pt_detail, dupl_charts,
                                       by = c("NameDOBID" = "NameDOBID"))


# Ungroup dataframe and rename columns
dose2_no_dose1_pt_detail <- dose2_no_dose1_pt_detail %>%
  ungroup()

colnames(dose2_no_dose1_pt_detail) <- c("Identifier", "LastName_FirstInitial", "DOB",
                                      "Count_ArrSch_Dose1", "Count_NoShowCanc_Dose1", 
                                      "Count_Arr_Dose2", "Count_Sch_Dose2", "Count_ArrSch_Dose2",
                                      "Scenario", "MRN", "CSN", "Site", "PodType", 
                                      "MultDose2Loc", "MultPtChart")

# Summarize data
dose2_no_dose1_errors_summary <- dose2_no_dose1_pt_detail %>%
  group_by(Site, Scenario) %>%
  summarize(Count = n()) %>%
  arrange(Site, Scenario)

dose2_no_dose1_errors_summary_cast <- dcast(dose2_no_dose1_errors_summary,
                                            Site ~ Scenario, value.var = "Count")

# Create data frame with patient details for each scenario
receive_dose2_wo_dose1 <- dose2_no_dose1_pt_detail %>%
  filter(Scenario == "Received Dose 2 w/o Arr/Sch Dose 1 at MSHS") %>%
  select(Site, PodType, 
         Identifier, LastName_FirstInitial, DOB, 
         MRN, CSN, MultPtChart, MultDose2Loc,
         Count_NoShowCanc_Dose1, Count_Arr_Dose2, Count_Sch_Dose2, Count_ArrSch_Dose2)

sched_dose2_wo_dose1 <- dose2_no_dose1_pt_detail %>%
  filter(Scenario == "Sched Dose 2 w/o Arr/Sch Dose 1 at MSHS") %>%
  select(Site, PodType, 
         Identifier, LastName_FirstInitial, DOB, 
         MRN, CSN, MultPtChart, MultDose2Loc,
         Count_NoShowCanc_Dose1, Count_Arr_Dose2, Count_Sch_Dose2, Count_ArrSch_Dose2)

sched_dose2_noshow_dose1 <- dose2_no_dose1_pt_detail %>%
  filter(Scenario == "No Show/Canc Dose 1 & Dose 2 Still On Sched") %>%
  select(Site, PodType, 
         Identifier, LastName_FirstInitial, DOB, 
         MRN, CSN, MultPtChart, MultDose2Loc,
         Count_NoShowCanc_Dose1, Count_Arr_Dose2, Count_Sch_Dose2, Count_ArrSch_Dose2)

mult_dose2 <- dose2_no_dose1_pt_detail %>%
  filter(Scenario == "Multiple Arr/Sched Dose 2 - Scheduling Error") %>%
  select(Site, PodType, 
         Identifier, LastName_FirstInitial, DOB, 
         MRN, CSN, MultPtChart, MultDose2Loc,
         Count_NoShowCanc_Dose1, Count_Arr_Dose2, Count_Sch_Dose2, Count_ArrSch_Dose2)


# Create a dataframe with patients with multiple charts
all_pt_charts <- left_join(pt_id_mrn[ , c("NameDOBID", "MRN")], dupl_charts,
                           by = c("NameDOBID" = "NameDOBID"))

all_pt_charts <- all_pt_charts %>%
  arrange(DuplChart, NameDOBID)

all_pt_dupl_charts <- all_pt_charts %>%
  filter(DuplChart)

all_pt_dupl_count <- all_pt_dupl_charts %>%
  group_by(NameDOBID) %>%
  summarize(Count_MRN = n())

all_pt_dupl_charts <- left_join(all_pt_dupl_charts, all_pt_dupl_count,
                                by = c("NameDOBID" = "NameDOBID"))

all_pt_dupl_charts <- all_pt_dupl_charts %>%
  arrange(Count_MRN, NameDOBID)

# Create list of data to export

error_export_list <- list("Dose2_Error_SummaryBySite" = dose2_no_dose1_errors_summary,
                          "Dose2_Error_SummaryTable" = dose2_no_dose1_errors_summary_cast,
                          "ReceivedDose2_NoDose1" = receive_dose2_wo_dose1,
                          "SchedDose2_NoDose1" = sched_dose2_wo_dose1,
                          "NoShowCancDose1_SchedDose2" = sched_dose2_noshow_dose1,
                          "MultipleDose2Sched" = mult_dose2,
                          "All_Pt_Mult_Charts" = all_pt_dupl_charts)

write_xlsx(error_export_list, path = paste0(user_directory, 
                                                        "/AdHoc/Dose2 No Dose1 Though ", format(max(sched_hist_outlook$ApptDate), "%m%d%y"), " Error Summary & Pt List ", format(Sys.Date(), "%m%d%y"), ".xlsx"))


