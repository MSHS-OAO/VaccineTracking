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

# Import data with canceled appointments
canc_dose1_dec <- read_excel(paste0(user_directory, "/Canc Dose 1 Reports/Dose 1 Canc Appt Dec20 Pulled 1022AM 021721.xlsx"))
canc_dose1_jan <- read_excel(paste0(user_directory, "/Canc Dose 1 Reports/Dose 1 Canc Appt Jan21 Pulled 1100AM 021721.xlsx"))
canc_dose1_feb <- read_excel(paste0(user_directory, "/Canc Dose 1 Reports/Dose 1 Canc Appt Feb21 Pulled 1115AM 021721.xlsx"))
canc_dose1_mar <- read_excel(paste0(user_directory, "/Canc Dose 1 Reports/Dose 1 Canc Appt Mar21 Pulled 1119AM 021721.xlsx"))

canc_dose1 <- rbind(canc_dose1_dec,
                    canc_dose1_jan,
                    canc_dose1_feb,
                    canc_dose1_mar)

# Dates of interest
start_date <- today
end_date <- as.Date("2/20/21", format = "%m/%d/%y")

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
                          ifelse(ApptDate < (today) & Status == "Sch", "No Show", # Categorize any appts still in sch status from prior days as no shows; using "today - 1" since this is 1 day behind
                                 ifelse(Status == "Arr" & (is.na(`Level Of Service`) | str_detect(`Level Of Service`, "ERRONEOUS")), "Left", Status))), # Correct patient with Arr status but incomplete visit
         # Modify status 2 for running report late in evening instead of early morning
         # Status2 = ifelse(ApptDate == today & Status == "Arr", "Arr", # Categorize any arrivals from today as scheduled for easier manipulation
         #                  ifelse(ApptDate <= today & Status == "Sch", "No Show", Status)), # Categorize any appts still in sch status from prior days as no shows
         Mfg = ifelse(Status2 != "Arr", NA, #Keep manufacturer as NA if the appointment hasn't been arrived
                      ifelse(is.na(Immunizations), "Pfizer", # Assume any visits without immunization record are Pfizer
                             ifelse(str_detect(Immunizations, "Moderna"), "Moderna", "Pfizer"))),
         WeekNum = format(ApptDate, "%U"),
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

# saveRDS(sched_hist_outlook, file = paste0(user_directory, "/Dose2 Scheduling Errors/Sched Repo ",
#                                           format(min(sched_hist_outlook$Date), "%m%d%y"), " to ",
#                                           format(max(sched_hist_outlook$Date), "%m%d%y"),
#                                           " as of ", format(Sys.time(), "%m%d%y %H%M"), ".rds"))


# Create flag to determine if dose 2 appointment is within dates of interest
sched_hist_outlook <- sched_hist_outlook %>%
  mutate(DateRange = ifelse(Dose == 2 & Status2 == "Sch" & ApptDate >= start_date & ApptDate <= end_date, 
                            paste0("Dose 2 Sched from ", format(start_date, "%m%d%y"), " to ", format(end_date, "%m%d%y")),
                            ifelse(Dose == 2 & Status2 == "Sch" & ApptDate > end_date,
                                   paste0("Dose 2 Sched from ", format(end_date + 1, "%m%d%y"), " to ", format(max(ApptDate), "%m%d%y")),
                                   NA)))

# Create flag to identify patients with multiple charts
pt_id_mrn <- as.data.frame(unique(sched_hist_outlook[ , c("NameDOBID", "MRN")]))

# Create flag for patients with duplicate charts (ie, more than 1 MRN for our name DOB identifier) 
pt_id_mrn <- pt_id_mrn %>%
  mutate(DuplChart = duplicated(NameDOBID))

dupl_charts <- unique(pt_id_mrn %>%
                        filter(DuplChart) %>%
                        mutate(MRN = NULL))

pt_id_mrn <- pt_id_mrn %>%
  mutate(DuplChart = NULL)

pt_id_mrn <- left_join(pt_id_mrn, dupl_charts,
                       by = c("NameDOBID" = "NameDOBID"))

pt_id_mrn <- pt_id_mrn %>%
  mutate(DuplChart = ifelse(is.na(DuplChart), FALSE, DuplChart)) %>%
  arrange(DuplChart, MRN)

unique_mrn <- unique(pt_id_mrn[ , c("MRN", "DuplChart")])

unique_mrn <- unique_mrn %>%
  arrange(desc(DuplChart)) %>%
  mutate(Dupl = duplicated(MRN))

unique_mrn <- unique_mrn %>%
  filter(!Dupl) %>%
  mutate(Dupl = NULL)


# Summarize data by MRN
mrn_appt_summary <- sched_hist_outlook %>%
  group_by(MRN) %>%
  summarize(Arr_Sch_Dose1 = sum(Dose == 1 & Status2 %in% c("Arr", "Sch")),
            NoShow_Dose1 = sum(Dose == 1 & Status2 %in% c("No Show")),
            Canc_Dose1 = sum(Dose == 1 & Status2 %in% c("Can")),
            Left_Dose1 = sum(Dose == 1 & Status2 %in% c("Left")),
            NoShow_Canc_Dose1 = sum(Dose == 1 & Status2 %in% c("No Show", "Can", "Left")),
            Arr_Dose2 = sum(Dose == 2 & Status2 %in% c("Arr")),
            Sch_Dose2 = sum(Dose == 2 & Status2 %in% c("Sch")),
            Arr_Sch_Dose2 = sum(Dose == 2 & Status2 %in% c("Arr", "Sch")))

mrn_appt_summary <- left_join(mrn_appt_summary, unique_mrn,
                              by = c("MRN" = "MRN"))

mrn_appt_summary <- mrn_appt_summary %>%
  arrange(-DuplChart)

mrn_dose2_no_dose1 <- mrn_appt_summary %>%
  filter(Arr_Sch_Dose1 == 0 & Arr_Sch_Dose2 > 0) %>%
  mutate(Scenario = ifelse(Arr_Dose2 == 0 & NoShow_Canc_Dose1 >= 1 & Sch_Dose2 >= 1, "No Show/Canc Dose 1 & Dose 2 Still on Sched",
                           ifelse(Arr_Dose2 == 1 & Sch_Dose2 == 0, "Received Dose 2 w/o Arr/Sch Dose 1 at MSHS",
                                  ifelse(Arr_Dose2 == 0 & Sch_Dose2 == 1, "Sched Dose 2 w/o Arr/Sch Dose 1 at MSHS",
                                         ifelse(Arr_Sch_Dose2 > 1, "Multiple Arr/Sched Dose 2 - Scheduling Error", "Other")))))

# Pull in additional patient details including MRN, CSN, Site, and Pod Type for each scheduled dose 2 appointment
dose2_appt_details <- sched_hist_outlook %>%
  filter(Dose == 2 & Status2 %in% c("Arr", "Sch")) %>%
  select(MRN, CSN, Site, `Pod Type`, ApptDate, DateRange)

dose2_appt_details <- dose2_appt_details %>%
  mutate(DateRange = ifelse(is.na(DateRange), "Before Dates of Interest", DateRange))

dose2_appt_details <- unique(dose2_appt_details)

dose2_no_dose1_pt_detail <- left_join(mrn_dose2_no_dose1, dose2_appt_details,
                                      by = c("MRN" = "MRN"))

# Create flag to identify patients with multiple dose 2 locations
mrn_dose2_loc <- dose2_appt_details %>%
  select(MRN, Site)

mrn_dose2_loc <- unique(mrn_dose2_loc)

mrn_dose2_loc <- mrn_dose2_loc %>%
  mutate(MultDose2Loc = duplicated(MRN))

mult_dose2_loc <- mrn_dose2_loc %>%
  filter(MultDose2Loc) %>%
  mutate(Site = NULL)

mult_dose2_loc <- unique(mult_dose2_loc)

dose2_no_dose1_pt_detail <- left_join(dose2_no_dose1_pt_detail, mult_dose2_loc,
                                      by = c("MRN" = "MRN"))

dose2_no_dose1_pt_detail <- dose2_no_dose1_pt_detail %>%
  mutate(MultDose2Loc = ifelse(is.na(MultDose2Loc), FALSE, MultDose2Loc))

# Ungroup dataframe and rename columns
dose2_no_dose1_pt_detail <- dose2_no_dose1_pt_detail %>%
  ungroup()

colnames(dose2_no_dose1_pt_detail) <- c("MRN", 
                                        "Count_ArrSch_Dose1", "Count_NoShow_Dose1", 
                                        "Count_Canc_Dose1", "Count_Left_Dose1", "Count_NoShowCanc_Dose1", 
                                        "Count_Arr_Dose2", "Count_Sch_Dose2", "Count_ArrSch_Dose2",
                                        "MultPtChart",
                                        "Scenario", "CSN", "Site", "PodType", "ApptDate", "DateRange", "MultDose2Loc")

dose2_no_dose1_pt_detail <- dose2_no_dose1_pt_detail[ , c("Scenario", "Site", "MRN", 
                                                          "MultPtChart", "DateRange", "MultDose2Loc",
                                                          "CSN", "ApptDate", "PodType",
                                                          "Count_ArrSch_Dose1", 
                                                          "Count_NoShow_Dose1", "Count_Canc_Dose1", "Count_Left_Dose1",
                                                          "Count_NoShowCanc_Dose1",
                                                          "Count_Arr_Dose2", "Count_Sch_Dose2", "Count_ArrSch_Dose2")]

dose2_no_dose1_pt_detail <- dose2_no_dose1_pt_detail %>%
  arrange(Scenario, Site, MRN)

# Pull in appointment made date
csn_made_date <- sched_hist_outlook %>%
  select(CSN, `Made Date`)

csn_made_date <- unique(csn_made_date)

dose2_no_dose1_pt_detail <- left_join(dose2_no_dose1_pt_detail, csn_made_date,
                                      by = c("CSN" = "CSN"))

# Summarize data
dose2_no_dose1_errors_site <- dose2_no_dose1_pt_detail %>%
  group_by(Site, DateRange, Scenario) %>%
  summarize(Count = n()) %>%
  arrange(Site, Scenario) %>%
  ungroup()

dose2_no_dose1_errors_system <- dose2_no_dose1_pt_detail %>%
  group_by(DateRange, Scenario) %>%
  summarize(Site = "MSHS",
            Count = n()) %>%
  ungroup()

dose2_no_dose1_errors_system <- dose2_no_dose1_errors_system[ , colnames(dose2_no_dose1_errors_site)]

# dose2_no_dose1_errors_system_this_wk <- dose2_no_dose1_errors_system_this_wk[ , colnames(dose2_no_dose1_errors_site_this_wk)]

dose2_no_dose1_errors_summary <- rbind(dose2_no_dose1_errors_site, dose2_no_dose1_errors_system)

dose2_no_dose1_errors_summary <- dose2_no_dose1_errors_summary %>%
  filter(DateRange != "Before Dates of Interest")

dose2_no_dose1_errors_export <- dcast(dose2_no_dose1_errors_summary,
                                      Site + Scenario ~ DateRange,
                                      value.var = "Count")

# dose2_no_dose1_errors_export <- dose2_no_dose1_errors_export[ , c(1:2, 4, 3)]

dose2_no_dose1_errors_export <- dose2_no_dose1_errors_export %>%
  mutate(Site = factor(Site, levels = sites, ordered = TRUE)) %>%
  arrange(Site, Scenario)

# Crosswalk cancel reason for patients with 1 Dose 1 cancellation
# First classify number of no shows/cancellations
dose2_no_dose1_pt_detail <- dose2_no_dose1_pt_detail %>%
  mutate(RollUp = ifelse(Scenario != "No Show/Canc Dose 1 & Dose 2 Still on Sched", Scenario,
                         ifelse(Count_NoShowCanc_Dose1 > 1, "Multiple No Show/Canc",
                                ifelse(Count_NoShow_Dose1 == 1, "Single No Show",
                                       ifelse(Count_Canc_Dose1 == 1, "Single Canc",
                                              ifelse(Count_Left_Dose1 == 1, "Single Left", "Other"))))))

# Subset data from canceled dose 1 reports
canc_dose1_subset <- canc_dose1 %>%
  mutate(Dose1_CSN = CSN,
         Dose1_MadeDate = `Made Date`,
         Dose1_ApptDate = Date,
         Dose1_CancDate = `Canc Date`,
         Dose1_CancReason = `Canc Reason`) %>%
  select(MRN,
         Dose1_CSN,
         Dose1_MadeDate,
         Dose1_ApptDate,
         Dose1_CancDate,
         Dose1_CancReason)

# Subset patients with 1 dose 1 cancellation
single_dose1_cancel <- dose2_no_dose1_pt_detail %>%
  filter(RollUp == "Single Canc")

# Crosswalk canceled dose 1 information
single_dose1_cancel <- left_join(single_dose1_cancel, canc_dose1_subset,
                                 by = c("MRN" = "MRN"))

single_dose1_cancel_subset <- single_dose1_cancel %>%
  select(MRN, CSN,
         Dose1_CSN, Dose1_MadeDate, Dose1_ApptDate, Dose1_CancDate, Dose1_CancReason)

dose2_no_dose1_pt_detail <- left_join(dose2_no_dose1_pt_detail, single_dose1_cancel_subset,
                                           by = c("MRN" = "MRN",
                                                  "CSN" = "CSN"))

# Create column with action items based on conversation w/ Amb Ops at 1:30PM on 2/17/21
# Acceptable reasons for mass cancellations
mass_cancel_reasons <- c("Covid Vaccine Supply",
                         "Cancelled Via ZocDoc",
                         "Canceled via MyChart",
                         "Patient Cancelled",
                         "Patient")

dose2_no_dose1_pt_detail <- dose2_no_dose1_pt_detail %>%
  mutate(Action =  ifelse(DateRange == "Before Dates of Interest", "No Action - Historical", 
                          ifelse(Scenario == "Sched Dose 2 w/o Arr/Sch Dose 1 at MSHS", "High Priority Chart Review",
                                 ifelse(MultPtChart | 
                                          RollUp == "Multiple No Show/Canc" |
                                          Count_Sch_Dose2 > 1, "Chart Review",
                                               ifelse(RollUp == "Single No Show" | Dose1_CancReason %in% mass_cancel_reasons, "Mass Cancel", 
                                                      ifelse(Scenario == "Multiple Arr/Sched Dose 2 - Scheduling Error" & Count_Sch_Dose2 == 1, 
                                                             "No Action - Manually Reconciled", "TBD"))))))

# Summarize actions
action_summary <- dose2_no_dose1_pt_detail %>%
  filter(DateRange != "Before Dates of Interest") %>%
  group_by(Action) %>%
  summarize(Count = n())

# Create data frame with patient details for each scenario -------------------------------
# receive_dose2_wo_dose1 <- dose2_no_dose1_pt_detail %>%
#   filter(DateOfInterest &
#            Scenario == "Received Dose 2 w/o Arr/Sch Dose 1 at MSHS") %>%
#   select(Site, MRN, 
#          MultPtChart, MultDose2Loc,
#          CSN, PodType, 
#          Count_NoShow_Dose1, Count_Canc_Dose1, Count_Left_Dose1, 
#          Count_NoShowCanc_Dose1, 
#          Count_Arr_Dose2, Count_Sch_Dose2, 
#          Count_ArrSch_Dose2)

sched_dose2_wo_dose1 <- dose2_no_dose1_pt_detail %>%
  filter(DateRange != "Before Dates of Interest" &
           Scenario == "Sched Dose 2 w/o Arr/Sch Dose 1 at MSHS") %>%
  select(Site, MRN, DateRange, Action,
         MultPtChart, MultDose2Loc,
         CSN, ApptDate, PodType, 
         Count_NoShow_Dose1, Count_Canc_Dose1, Count_Left_Dose1, 
         Count_NoShowCanc_Dose1, 
         Count_Arr_Dose2, Count_Sch_Dose2, 
         Count_ArrSch_Dose2) %>%
  arrange(Action, desc(DateRange), Site, MRN)

sched_dose2_noshow_dose1 <- dose2_no_dose1_pt_detail %>%
  filter(DateRange != "Before Dates of Interest" &
           Scenario == "No Show/Canc Dose 1 & Dose 2 Still on Sched") %>%
  select(Site, MRN, DateRange, Action,
         MultPtChart, MultDose2Loc, 
         CSN, ApptDate, PodType, 
         Count_NoShow_Dose1, Count_Canc_Dose1, Count_Left_Dose1, 
         Count_NoShowCanc_Dose1, 
         Count_Arr_Dose2, Count_Sch_Dose2, 
         Count_ArrSch_Dose2,
         RollUp, Dose1_CSN, Dose1_ApptDate, Dose1_CancDate, Dose1_CancReason) %>%
  arrange(Action, desc(DateRange), Site, MRN)

mult_dose2 <- dose2_no_dose1_pt_detail %>%
  filter(DateRange != "Before Dates of Interest" &
           Scenario == "Multiple Arr/Sched Dose 2 - Scheduling Error") %>%
  select(Site, MRN, DateRange, Action,
         MultPtChart, MultDose2Loc, 
         CSN, ApptDate, PodType, 
         Count_NoShow_Dose1, Count_Canc_Dose1, Count_Left_Dose1, 
         Count_NoShowCanc_Dose1, 
         Count_Arr_Dose2, Count_Sch_Dose2, 
         Count_ArrSch_Dose2) %>%
  arrange(Action, desc(DateRange), Site, MRN)

# Create a dataframe with patients with multiple charts
mult_pt_charts <- pt_id_mrn %>%
  filter(DuplChart) %>%
  select(MRN, NameDOBID) %>%
  arrange(NameDOBID)

# Summarize stats for patients with no show/cancelled dose 1
noshow_canc_dose1_summary <- dose2_no_dose1_pt_detail %>%
  filter(DateRange != "Before Dates of Interest" &
           Scenario == "No Show/Canc Dose 1 & Dose 2 Still on Sched") %>%
  group_by(Scenario, RollUp, Dose1_CancReason) %>%
  summarize(Count = n()) %>%
  arrange(RollUp, desc(Count))


# Summarize patients with Dose 2 scheduled w/o ever having Dose 1 on schedule by pod type
sched_dose2_nodose1_summary <- dose2_no_dose1_pt_detail %>%
  filter(DateRange != "Before Dates of Interest" &
           Scenario == "Sched Dose 2 w/o Arr/Sch Dose 1 at MSHS") %>%
  group_by(Scenario, PodType, Action) %>%
  summarize(Count = n())

# # Look at made date stats
# made_date_summary <- dose2_no_dose1_pt_detail %>%
#   filter(DateRange != "Before Dates of Interest") %>%
#   group_by(Scenario, `Made Date`) %>%
#   summarize(Count = n())
# 
# made_date_summary_cast <- dcast(made_date_summary,
#                                 `Made Date` ~ Scenario,
#                                 value.var = "Count")
# 
# made_date_total <- dose2_no_dose1_pt_detail %>%
#   filter(DateRange != "Before Dates of Interest") %>%
#   group_by(`Made Date`) %>%
#   summarize(Total = n())
# 
# made_date_summary_stats <- left_join(made_date_summary_cast, made_date_total,
#                                      by = c("Made Date" = "Made Date"))
# 
# made_date_summary_stats <- made_date_summary_stats %>%
#   mutate(Percent = percent(Total / sum(Total))) %>%
#   arrange(desc(Percent), `Made Date`)

# Create list of data to export
error_export_list <- list("Dose2_Error_SummaryTable" = dose2_no_dose1_errors_export,
                          "Action Summary" = action_summary,
                          "Pt_w_Mult_Charts" = mult_pt_charts,
                          # "ReceivedDose2_NoDose1" = receive_dose2_wo_dose1,
                          "SchedDose2_NoDose1" = sched_dose2_wo_dose1,
                          "NoShowCancDose1_SchedDose2" = sched_dose2_noshow_dose1,
                          "MultipleDose2Sched" = mult_dose2,
                          "NoShowCancDose1_CancReason" = noshow_canc_dose1_summary,
                          "SchedDose2_NoDose1_PodType" = sched_dose2_nodose1_summary)

write_xlsx(error_export_list, path = paste0(user_directory, 
                                            "/Dose2 Scheduling Errors/Dose2 No Dose1 Errors & Actions Through ", 
                                            format(max(sched_hist_outlook$ApptDate), format = "%m%d%y"), " Created On ", format(Sys.Date(), "%m%d%y"), ".xlsx"))


