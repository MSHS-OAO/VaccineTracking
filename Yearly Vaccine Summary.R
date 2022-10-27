# Code to summarize doses by year

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
  "/Hybrid Sched & Admin Reporting",
  "/Hybrid Reporting Dept Mapping 2022-10-27.xlsx"),
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

# Import hybrid schedule and administration data repository
repo_file <- choose.files(
  default = 
    paste0(user_directory,
           "/Hybrid Sched & Admin Reporting",
           "/Hybrid Model Repo",
           "/*.*"),
  caption = "Select hybrid vaccine administration repository.")

hybrid_repo <- readRDS(repo_file)

# Format and analyze Epic schedule to date for dashboards ---------------------
admin_to_date <- hybrid_repo

admin_to_date <- admin_to_date %>%
  mutate(
    VaxAgeGroup = ifelse(str_detect(VaxType, "Peds"), "Peds", "Adult"),
    Year = year(Date),
    DoseGroup = ifelse(Dose %in% c("Dose 1", "Dose 2"), "Primary Series",
                       ifelse(Dose %in% c("Dose 3/Booster"), "Dose 3/Booster",
                              NA))
  )


# Summarize doses administered by year
yearly_summary <- admin_to_date %>%
  group_by(VaxAgeGroup, DoseGroup, Year) %>%
  summarize(Volume = sum(Count, na.rm = TRUE)) %>%
  ungroup() %>%
  arrange(Year, DoseGroup, VaxAgeGroup) %>%
  pivot_wider(names_from = Year,
              values_from = Volume) %>%
  mutate(VaxAgeGroup = factor(VaxAgeGroup,
                              levels = unique(VaxAgeGroup),
                              ordered = TRUE),
         DoseGroup = factor(DoseGroup,
                            levels = c("Primary Series",
                                       "Dose 3/Booster"),
                            ordered = TRUE)) %>%
  arrange(VaxAgeGroup, DoseGroup) %>%
  adorn_totals(where = "col", na.rm = TRUE, name = 'Total')

export_list <- list("Annual Summary" = yearly_summary)

# write_xlsx(export_list,
#            path = paste0(user_directory,
#                          "/AdHoc",
#                          "/Vaccines Administered by Year as of ",
#                          format(Sys.Date(), "%Y-%m-%d"),
#                          ".xlsx"))

