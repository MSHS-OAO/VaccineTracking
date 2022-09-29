# Code for exporting vaccines administered by week for Jen Paez

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
#install.packages("zipcodeR")
#install.packages("tidyr")
#install.packages("janitor")
#install.packages("knitr")
#install.packages("rmarkdown")
#install.packages("kableExtra")

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
library(tidyr)
library(janitor)
library(knitr)
library(rmarkdown)
library(kableExtra)

# Set working directory and select raw data ----------------------------
rm(list = ls())

# Determine path for working directory
if ("Presidents" %in% list.files("J://")) {
  user_directory <- paste0("J:/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "System Operations/COVID Vaccine")
} else {
  user_directory <- paste0("J:/deans/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "System Operations/COVID Vaccine")
}

# Set path for file selection
user_path <- paste0(user_directory, "\\*.*")

# Import reference data for site and pod mappings
dept_mappings <- read_excel(paste0(
  user_directory,
  "/Hybrid Repo Sched & Admin Data",
  "/Hybrid Reporting Dept Mapping 2022-09-27.xlsx"),
  sheet = "Hybrid Dept Mapping")

vax_admin_data_mappings <- dept_mappings %>%
  select(-Department) %>%
  unique()

# Store today's date
today <- Sys.Date()
yesterday <- today - 1

# Site order
all_sites <- c("MSB",
               "MSBI",
               "MSH",
               "MSM",
               "MSQ",
               "MSW",
               "MSVD",
               "Network LI",
               "MSSN")

peds_sites <- c("MSH", "MSBI")

# Vaccine manufacturers
mfg <- c("Pfizer", "Moderna", "J&J")

# Vaccine types
vax_types <- c("Adult", "Peds")

# NY zip codes
ny_zips <- search_state("NY")

# Create dataframe with start dates for adult and peds vaccination
vacc_start_df <- data.frame("VaxType" = c("Adult", "Peds"),
                            "StartDate" = c(as.Date("12/15/20",
                                                    format = "%m/%d/%y"),
                                            as.Date("11/04/21",
                                                    format = "%m/%d/%y")))

vacc_start_df <- vacc_start_df %>%
  mutate(
    # Determine the Sunday of the first week of vaccination
    FirstSun = StartDate - (wday(StartDate) - 1),
    # Determine this week number
    ThisWeek = as.numeric(floor((today - FirstSun) / 7) + 1))

# Create dataframe with dates and weeks since vaccines started
all_dates <- data.frame("Date" = seq.Date(min(vacc_start_df$StartDate),
                                          today + 7,
                                          by = 1
)
)

all_dates <- all_dates %>%
  mutate(
    # Determine vaccine week for adult vaccinations
    Adult_VaccWeek = as.numeric(
      floor((Date -
               vacc_start_df$FirstSun[
                 which(vacc_start_df$VaxType == "Adult")]) / 7) + 1),
    # Determine vaccine week for pediatric vaccinations
    Peds_VaccWeek = ifelse(
      as.numeric(
        floor((Date -
                 vacc_start_df$FirstSun[
                   which(vacc_start_df$VaxType == "Peds")]) / 7) + 1) < 1, NA,
      as.numeric(
        floor((Date -
                 vacc_start_df$FirstSun[
                   which(vacc_start_df$VaxType == "Peds")]) / 7) + 1)
    )
  ) %>%
  pivot_longer(cols = contains("_VaccWeek"),
               names_to = "VaxType",
               values_to = "VaccWeek") %>%
  mutate(VaxType = str_remove(VaxType, "_VaccWeek"),
         WeekStart = Date - (wday(Date) - 1),
         WeekEnd = Date + (7 - wday(Date)),
         WeekDates = paste0(
           format(WeekStart, "%m/%d/%y"),
           "-",
           format(WeekEnd, "%m/%d/%y")
         ),
         WeekStart = NULL,
         WeekEnd = NULL
  ) %>%
  filter(!is.na(VaccWeek)) %>%
  arrange(VaxType, Date)


# Import hybrid schedule and administration data repository
repo_file <- choose.files(
  default = 
    paste0(user_directory,
           "/Epic Vaccines Administered Report",
           "/*.*"),
  caption = "Select hybrid vaccine administration repository.")

hybrid_repo <- readRDS(repo_file)

# Format and analyze Epic schedule to date for dashboards ---------------------
admin_to_date <- hybrid_repo

admin_to_date <- admin_to_date %>%
  mutate(
    VaxAgeGroup = ifelse(str_detect(VaxType, "Peds"), "Peds", "Adult")
  )

# Verify if age group lines up with vaccine start dates
admin_to_date <- left_join(admin_to_date,
                           vacc_start_df %>%
                             select(VaxType, StartDate),
                           by = c("VaxAgeGroup" = "VaxType"))

admin_to_date <- admin_to_date %>%
  mutate(VerifyVaxType = Date >= StartDate,
         VaxAgeGroup = ifelse(!VerifyVaxType, "Adult", VaxAgeGroup)) %>%
  select(-StartDate, -VerifyVaxType)

# Determine vaccine week based on appointment date
admin_to_date <- left_join(admin_to_date,
                           all_dates,
                           by = c("Date" = "Date",
                                  "VaxAgeGroup" = "VaxType"))

# Summarize schedule data for Epic sites by key stratification
admin_to_date_summary <- admin_to_date %>%
  group_by(`Hospital Roll Up`,
           VaxAgeGroup,
           Dose,
           VaccWeek,
           WeekDates) %>%
  summarize(Count = sum(Count, na.rm = TRUE),
            .groups = "keep") %>%
  pivot_wider(names_from = Dose,
              values_from = Count) %>%
  ungroup() %>%
  rename(Site = `Hospital Roll Up`) %>%
  mutate(AllDoses = select(., contains("Dose")) %>%
           rowSums(na.rm = TRUE)) %>%
  arrange(VaxAgeGroup, VaccWeek, Site)

adult_admin_summary <- admin_to_date_summary %>%
  filter(VaxAgeGroup %in% "Adult")

peds_admin_summary <- admin_to_date_summary %>%
  filter(VaxAgeGroup %in% "Peds")

export_list <- list("Adult" = adult_admin_summary,
                    "Peds" = peds_admin_summary)

write_xlsx(export_list,
           path = paste0(user_directory,
                         "/Arr Visits Weekly Summary",
                         "/Vaccines Administered as of ",
                         format(Sys.Date(), "%Y-%m-%d"),
                         ".xlsx"))

