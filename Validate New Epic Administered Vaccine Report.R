# Clear environment
rm(list = ls())

# Load libraries
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
library(tidyr)
library(purrr)
library(janitor)
library(DT)

# Determine directory
if ("Presidents" %in% list.files("J://")) {
  user_directory <- paste0("J:/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "System Operations/COVID Vaccine")
} else {
  user_directory <- paste0("J:/deans/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "System Operations/COVID Vaccine")
}

update_repo <- FALSE

# Set path for file selection
user_path <- paste0(user_directory, "\\*.*")

# # Determine whether or not to update an existing repo
# update_repo <- FALSE

# Import reference data for site and pod mappings
dept_mappings <- read_excel(paste0(
  user_directory,
  "/Weekly Dashboard Doses Administered",
  "/MSHS COVID-19 Vaccine Department Mappings 2022-04-11.xlsx"))

dept_mappings <- dept_mappings %>%
  select(-`Site (Historical)`, -Notes)

# Store today's date
today <- Sys.Date()

# Site order
all_sites <- c("MSB",
               "MSBI",
               "MSH",
               "MSM",
               "MSQ",
               "MSW",
               "MSVD",
               "Network LI")

# Vaccine types
vax_types <- c("Adult", "Peds")

# Dept grouper
dept_group <- c("Hospital POD", "MSHS Practice")


epic_df <- read_excel(path = paste0(user_directory,
                                    "/Epic Vaccines Administered Report",
                                    "/Sample Report 2022-07-11.xls"))

epic_df <- epic_df %>%
  mutate(Existing_Dept = DEPARTMENT_NAME %in% dept_mappings$Department,
         ImmunizationDate = as.Date(
           str_extract(IMMUNE_DATE,
                       "[0-9]{2}\\-[A-Z]{3}\\-[0-9]{4}"),
           format = "%d-%b-%Y"))

missing_depts <- epic_df %>%
  filter(!Existing_Dept) %>%
  select(DEPARTMENT_NAME) %>%
  distinct()

epic_df <- left_join(epic_df,
                     dept_mappings,
                     by = c("DEPARTMENT_NAME" = "Department"))

epic_df_summary <- epic_df %>%
  group_by(Existing_Dept, `New Site`, DEPARTMENT_NAME, ImmunizationDate) %>%
  summarize(Count = n())

existing_sites_summary <- epic_df_summary %>%
  filter(Existing_Dept) %>%
  group_by(`New Site`) %>%
  summarize(Count = sum(Count, na.rm = TRUE))

# Import schedule repo --------------------------
sched_repo_file <- choose.files(default = paste0(user_directory,
                                                 "/R_Sched_AM_Repo/*.*"))

sched_repo <- readRDS(sched_repo_file)

sched_to_date <- sched_repo

sched_to_date <- sched_to_date %>%
  # rename(IsPtEmployee = `Is the patient a Mount Sinai employee`) %>%
  mutate(
    # Update timezones to EST if imported as UTC
    `Appt Time` = with_tz(`Appt Time`, tzone = "EST"),
    # Determine whether appointment is for dose 1, 2, or 3 based on scheduled visit type
    Dose = ifelse(str_detect(Type, "DOSE 1"), "Dose 1",
                  ifelse(str_detect(Type, "DOSE 2"), "Dose 2",
                         ifelse(str_detect(Type, "(DOSE 3)|(BOOSTER)"),
                                "Dose 3/Boosters", NA))),
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
    # Update appointment status: classify any appts from prior days that are
    # still in scheduled as No Shows
    Status2 = ifelse(ApptDate < (today) & Status == "Sch", "No Show", 
                     # Correct any appts erroneously marked as arrived early
                     ifelse(ApptDate >= today & Status == "Arr", "Sch",
                            Status)),
    # # Determine whether the vaccine is an adult or pediatric dose
    # VaxType = ifelse(Department %in% c("1468 MADISON PEDIATRIC VACCINE POD",
    #                                    "1468 MADISON HOSP PEDS GEN") |
    #                    `Provider/Resource` %in% c("DOSE 1 PEDIATRIC [1324684]",
    #                                               "DOSE 2 PEDIATRIC [1324685]") |
    #                    Department %in% peds_school_practice_dept,
    #                  "Peds", "Adult"),
    # # Determine manufacturer based on immunization field and vaccine type fields
    # Mfg = #Keep manufacturer as NA if the appointment hasn't been arrived
    #   ifelse(Status2 != "Arr", NA,
    #          # Any blank immunizations and any peds assumed to be Pfizer
    #          ifelse(is.na(Immunizations) |
    #                   VaxType %in% c("Peds"), "Pfizer",
    #                 # Any immunizations with Moderna in text classified as Moderna
    #                 ifelse(str_detect(Immunizations, "Moderna"), "Moderna",
    #                        # Any visiting docs or Johnson and Johnson
    #                        # immunizations classified as J&J
    #                        ifelse(str_detect(Department, "VISITING DOCS") |
    #                                 str_detect(Immunizations,
    #                                            "Johnson and Johnson"),
    #                               "J&J", "Pfizer")))),
    # Determine week number of year
    WeekNum = format(ApptDate, "%U"),
    # Determine appointment day of week
    DOW = weekdays(ApptDate)
  )

# Crosswalk sites and department grouper
sched_to_date <- left_join(sched_to_date, dept_mappings,
                           by = c("Department" = "Department"))

sched_to_date <- sched_to_date %>%
  rename(Site = `New Site`)

sched_to_date <- sched_to_date %>%
  mutate(
    # Create a column for age grouper
    Age_Grouper = ifelse(str_detect(Age, "\\ y.o."), "years",
                         ifelse(str_detect(Age, "\\ m.o."), "months",
                                ifelse(str_detect(Age, "\\ days"), "days",
                                       "unknown"))),
    # Convert string with age to numeric
    Age_Numeric = 
      ifelse(Age_Grouper == "years",
             as.numeric(str_extract(Age,
                                    # pattern = "(?<=Deceased\\ \\().+(?=\\ y.o.)")),
                                    paste("(?<=Deceased\\ \\()[0-9]+(?=\\ y.o.)",
                                          "[0-9]+(?=\\ y.o.)",
                                          sep = "|"))),
             ifelse(Age_Grouper == "months",
                    as.numeric(str_extract(Age, "[0-9]+(?=\\ m.o.)")) / 12,
                    ifelse(Age_Grouper == "days",
                           as.numeric(
                             str_extract(Age, "[0-9]+(?=\\ days)")) / 365,
                           ifelse(Age_Grouper == "unknown", 10000,
                                  NA_integer_)))),
    # Determine vaccine type based on department, provider/resource, and age if necessary
    VaxType = ifelse(
      `Vaccine Age Group` %in% "Peds" |
        # MSBI Peds POD
        `Provider/Resource` %in% c("DOSE 1 PEDIATRIC [1324684]",
                                   "DOSE 2 PEDIATRIC [1324685]") |
        # Practics that administer adult and pediatric vaccines
        (`Vaccine Age Group` %in% "Both" &
           `Setting Grouper` %in% c("School Based Practice", "MSHS Practice") &
           Age_Numeric < 12),
      "Peds", "Adult"),
    # Determine manufacturer based on immunization field and vaccine type fields
    Mfg = #Keep manufacturer as NA if the appointment hasn't been arrived
      ifelse(Status2 != "Arr", NA,
             # Any blank immunizations and any peds assumed to be Pfizer
             ifelse(is.na(Immunizations) |
                      VaxType %in% c("Peds"), "Pfizer",
                    # Any immunizations with Moderna in text classified as Moderna
                    ifelse(str_detect(Immunizations, "Moderna"), "Moderna",
                           # Any visiting docs or Johnson and Johnson
                           # immunizations classified as J&J
                           ifelse(str_detect(Department, "VISITING DOCS") |
                                    str_detect(Immunizations,
                                               "Johnson and Johnson"),
                                  "J&J", "Pfizer")))),
    # Site =
    #   # Roll up any Moderna administered at MSB for YWCA to MSH
    #   ifelse(Site %in% c("MSB") & Mfg %in% c("Moderna"), "MSH",
    #          # Manually classify school practices as peds or adolescent for schools administering both
    #          ifelse(`Setting Grouper` %in% "School Based Practice" &
    #                   VaxType %in% "Peds", "Peds School Based Practice",
    #                 ifelse(`Setting Grouper` %in% "School Based Practice" &
    #                          VaxType %in% "Adult",
    #                        "Adolescent School Based Practice",
    #                        Site))),
    # Manually correct any J&J doses erroneously documented as dose 2 and change to dose 1
    Dose = ifelse(Mfg %in% c("J&J") & Dose %in% c("Dose 2"), "Dose 1", Dose),
    `Setting Grouper` = ifelse(`Setting Grouper` %in% "School Based Practice",
                               "MSHS Practice", `Setting Grouper`))

# Dates to include -------------
epic_df_min_date <- min(epic_df$ImmunizationDate)
epic_df_max_date <- max(epic_df$ImmunizationDate)

sched_repo_min_date <- min(sched_to_date$ApptDate[
  which(sched_to_date$Status2 %in% "Arr")])
sched_repo_max_date <- max(sched_to_date$ApptDate[
  which(sched_to_date$Status2 %in% "Arr")])

validate_start_date <- max(epic_df_min_date, sched_repo_min_date)
validate_end_date <- min(epic_df_max_date, sched_repo_max_date)

# Filter data frames for relevant dates
epic_df_validate <- epic_df %>%
  filter(Existing_Dept &
           ImmunizationDate >= validate_start_date &
           ImmunizationDate <= validate_end_date) %>%
  mutate(MRN_VaxDate = paste(MRN, ImmunizationDate))

sched_to_date_validate <- sched_to_date %>%
  filter(Status2 %in% "Arr" &
           ApptDate >= validate_start_date &
           ApptDate <= validate_end_date) %>%
  mutate(MRN_VaxDate = paste(MRN, ApptDate),
         MRN_VaxDate_Epic = MRN_VaxDate %in% epic_df_validate$MRN_VaxDate)

comparison_df <- sched_to_date_validate %>%
  group_by(Site) %>%
  summarize(Total_Arrived = n(),
            Matched_In_Epic = sum(MRN_VaxDate_Epic)) %>%
  adorn_totals(where = "row",
               name = "MSHS") %>%
  mutate(PercentMatch = (Matched_In_Epic / Total_Arrived))

write_xlsx(comparison_df,
           path = paste0(user_directory,
                         "/Epic Vaccines Administered Report",
                         "/Schedule and Admin Comparison Data ",
                         format(Sys.Date(), "%Y-%m-%d"),
                         ".xlsx"))

# Summarize data from new Epic report and schedule repository
epic_summary <- epic_df_validate %>%
  group_by(`New Site`, ImmunizationDate) %>%
  summarize(Epic_Admin_Count = n()) %>%
  arrange(`New Site`, ImmunizationDate)

sched_repo_summary <-sched_to_date_validate %>%
  group_by(Site, ApptDate) %>%
  summarize(Sched_Repo_Count = n()) %>%
  arrange(Site, ApptDate)

comparison_df <- full_join(epic_summary, sched_repo_summary,
                           by = c("New Site" = "Site",
                                  "ImmunizationDate" = "ApptDate"))

# # Summarize data for dashboard visualizations
# summary_all_doses <- sched_to_date %>%
#   filter(Status2 %in% "Arr") %>%
#   group_by(`Setting Grouper`, Site, VaxType) %>%
#   summarize(Count = n()) %>%
#   rename(Setting = `Setting Grouper`) %>%
#   arrange(Setting, VaxType, -Count) %>%
#   pivot_wider(names_from = VaxType,
#               values_from = Count) %>%
#   ungroup() %>%
#   arrange(Setting, Site) %>%
#   group_split(Setting) %>%
#   # split(.$Setting) %>%
#   # relocate(Setting, .after = Site) %>%
#   map_df(., function(x) {
#     adorn_totals(x, where = "row",
#                  name = unique(x$Setting),
#                  fill = "MSHS Total")})
# 
# summary_all_doses <- as.data.frame(summary_all_doses)
# 
# summary_all_doses <- summary_all_doses %>%
#   mutate(across(.cols = where(is.numeric),
#                 .fns = ~replace_na(.x, 0)))
# 
# #