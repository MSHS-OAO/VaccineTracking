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
  user_directory <- "J:/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/Service Lines/Lab Kpi/Data"
} else {
  user_directory <- "J:/deans/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/Service Lines/Lab Kpi/Data"
}

# 
# # Set working directory to script location
# script_path <- dirname(rstudioapi::getActiveDocumentContext()$path)
# setwd(script_path)
# # Change working directory to folder with daily reports
# setwd("..\\Data\\Epic Census Data")
# user_wd <- getwd()
# 
# # user_wd <- "J:\\Presidents\\HSPI-PM\\Operations Analytics and Optimization\\Projects\\System Operations\\Covid IP Staffing Model\\Data\\Epic Census Data"
# user_path <- paste0(user_wd, "\\*.*")
# 
# initial_run <- TRUE
# update_excel_repo <- TRUE