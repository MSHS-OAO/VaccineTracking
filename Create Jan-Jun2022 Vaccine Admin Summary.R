# Script to summarize administered COVID-19 vaccine appointments from 1/1/22-6/30/22

# Clear environment and load libraries ---------------
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

# Reference data -----------------------
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

# Import reference data for site and pod mappings
dept_mappings <- read_excel(paste0(
  user_directory,
  "/Hybrid Repo Sched & Admin Data",
  "/Hybrid Reporting Dept Mapping 2022-08-08.xlsx"),
  sheet = "Hybrid Dept Mapping")

vax_admin_data_mappings <- dept_mappings %>%
  select(-Department) %>%
  unique()

# Store today's date
today <- Sys.Date()

# Import historical data for vaccines administered ----------------------
hist_vax_admin <- read_excel(paste0(
  user_directory,
  "/Epic Vaccines Administered Report",
  "/COVID-8.19.2022.xls"
))

vax_admin_to_date <- unique(hist_vax_admin)

# Map on Epic ID
vax_admin_to_date <- left_join(vax_admin_to_date,
                               vax_admin_data_mappings,
                               by = c("EPICDEPID" = "EpicID"))

vax_admin_to_date <- vax_admin_to_date %>%
  select(-DOSE) %>%
  mutate(Mfg = ifelse(str_detect(IMMUNE_NAME, "PFIZER"), "Pfizer",
                      ifelse(str_detect(IMMUNE_NAME, "MODERNA"), "Moderna",
                             ifelse(str_detect(IMMUNE_NAME, "JOHNSON"), "J&J",
                                    NA))),
         VaxType = ifelse(str_detect(IMMUNE_NAME, "(\\(5-11 YEARS)\\)"),
                          "Peds (5-11 yo)",
                          ifelse(str_detect(IMMUNE_NAME,
                                            "(\\(6 MONTHS - 4 YEARS\\))"),
                                 "Peds (6 mo - 4 yo)", 
                          "Adult (12 yo +)")),
         Dose = ifelse(str_detect(replace_na(PRC_NAME, ""), "DOSE 1"),
                       "Dose 1",
                       ifelse(str_detect(replace_na(PRC_NAME, ""),
                                         "DOSE 2"), "Dose 2",
                              ifelse(str_detect(replace_na(PRC_NAME, ""),
                                                "(DOSE 3)|(BOOSTER)"),
                                     "Dose 3/Booster",
                                     ifelse(str_detect(RESPONSE_LIST,
                                                       "First Dose"),
                                            "Dose 1",
                                            ifelse(str_detect(RESPONSE_LIST,
                                                              "Second Dose"),
                                                   "Dose 2",
                                                   ifelse(str_detect(RESPONSE_LIST,
                                                                     "(Third Dose) | (Booster Dose)"),
                                                          "Dose 3/Booster", "Dose 3/Booster")))))),
         Date = as.Date(str_extract(IMMUNE_DATE,
                                    "[0-9]{2}\\-[A-Za-z]{3}\\-[0-9]{4}"),
                        format = "%d-%b-%Y")
         ) %>%
  rename(Department = DEPARTMENT_NAME)

vax_admin_repo_summary <- vax_admin_to_date %>%
  filter(!is.na(Department)) %>%
  group_by(Date, Department, `Setting Grouper`, 
           Site, `Hospital Roll Up`, 
           VaxType, Dose, Mfg) %>%
  summarize(Count = n()) %>%
  ungroup()

saveRDS(vax_admin_repo_summary,
        paste0(user_directory,
               "/Hybrid Repo Sched & Admin Data",
               "/Admin Immune Jan-Jun2022 Daily Summary ",
               format(today, "%Y-%m-%d"),
               ".rds"))



# # Save standardized, static vaccines administered repository as .RDS for use in future reporting ----------
# saveRDS(vax_admin_repo_summary,
#         paste0(user_directory,
#                "/Hybrid Repo Sched & Admin Data",
#                "/Vaccines Admin Summary Jan-Jun2022 ",
#                format(today),
#                ".RDS"))
