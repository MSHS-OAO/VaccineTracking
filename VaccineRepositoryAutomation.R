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

# Set working directory and select raw data ----------------------------
rm(list = ls())

# Determine path for working directory
if (list.files("J://") == "Presidents") {
  user_directory <- "J:/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/System Operations/COVID Vaccine/ScheduleData"
} else {
  user_directory <- "J:/deans/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/System Operations/COVID Vaccine/ScheduleData"
}

# Set path for file selection
user_path <- paste0(user_directory, "\\*.*")

# Determine whether or not to update an existing repo
initial_run <- TRUE
update_repo <- FALSE

# Import raw data from Epic
raw_df <- read_excel(choose.files(default = user_path, caption = "Select Epic report"), 
                     col_names = TRUE, na = c("", "NA"))
vx_sched <- raw_df

# Remove test patient
vx_sched <- vx_sched[vx_sched$Patient != "Patient, Test", ]

# Create column with vaccine location
vx_sched <- vx_sched %>% 
  mutate(Location = str_replace(str_replace(`Provider/Resource`, "COVID-19 VACCINE ", ""), "\\s\\[[0-9]+\\]", ""),
         Mfg = ifelse(is.na(Immunizations), "Unknown", ifelse(str_detect(Immunizations, "Pfizer"), "Pfizer", "Moderna")),
         Dose = ifelse(str_detect(Type, "DOSE 1"), 1, ifelse(str_detect(Type, "DOSE 2"), 2, NA)),
         ApptYear = year(Date),
         ApptMonth = month(Date),
         ApptDay = day(Date),
         ApptWeek = week(Date))

# Subset completed vaccines based on appointment status
vx_compl <- vx_sched %>%
  filter(`Appt Status` %in% c("Comp", "Arrived", "Checked Out"))

vx_compl_summary <- vx_compl %>%
  group_by(Location, Date) %>%
  summarize(Count = n())


### Next steps
# Count scheduled appointments, check to make sure patient hasn't already received that dose


