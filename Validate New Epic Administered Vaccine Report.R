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