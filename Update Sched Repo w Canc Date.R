# Code for adding cancel date to existing schedule repository

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

# Store today's date
today <- Sys.Date()

# Import schedule repository
sched_repo <- readRDS(choose.files(default = paste0(user_directory, "/R_Sched_AM_Repo/*.*"), caption = "Select schedule repository"))
    
# Add a column for cancel date to match new Epic reports
sched_repo <- sched_repo %>%
      mutate(`Canc Date` = NA,
             ApptDate = date(ApptDate))

# Reorder columns to match new Epic reports
sched_repo <- sched_repo[ , c(1:10, 33, 11:32)]

# Export repository to new .RDS 
saveRDS(sched_repo, paste0(user_directory, "/R_Sched_AM_Repo/Sched w Canc Date ",
                           format(min(sched_repo$ApptDate), "%m%d%y"), " to ",
                           format(max(sched_repo$ApptDate), "%m%d%y"),
                           " on ", format(Sys.time(), "%m%d%y %H%M"), ".rds"))
