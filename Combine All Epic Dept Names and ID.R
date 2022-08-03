# Create a list of unique Epic departments with names and IDs from historical data

rm(list = ls())

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

# Import schedule repository and determine last date it was updated
sched_repo_file <- choose.files(default = paste0(user_directory,
                                                 "/R_Sched_AM_Repo/*.*"))

sched_repo <- readRDS(sched_repo_file)

sched_repo_dept_names_id <- sched_repo %>%
  select(Department, Dept) %>%
  unique() %>%
  mutate(EpicID = as.numeric(str_extract(Dept, "(?<=\\[)[0-9]*(?=\\])")))

# Import vaccine administration data
hist_vax_admin <- read_excel(paste0(
  user_directory,
  "/Epic Vaccines Administered Report",
  "/Version4-Jan-June2022.xls"
))

vax_admin_dept_names_id <- hist_vax_admin %>%
  select(DEPARTMENT_NAME, EPICDEPID) %>%
  unique() %>%
  mutate(Dept = NA) %>%
  rename(Department = DEPARTMENT_NAME,
         EpicID = EPICDEPID) %>%
  relocate(Dept, .after = Department)

hybrid_data_dept_list <- rbind(sched_repo_dept_names_id,
                              vax_admin_dept_names_id)

hybrid_data_dept_list <- hybrid_data_dept_list %>%
  select(-Dept) %>%
  unique()

write_xlsx(hybrid_data_dept_list,
           path = paste0(user_directory,
                         "/Hybrid Repo Sched & Admin Data",
                         "/Epic Department Names and IDs 2022-08-03.xlsx"))
