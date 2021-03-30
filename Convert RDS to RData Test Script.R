# Code to import .RDS file and save as .RData

rm(list = ls())

library(dplyr)
library()

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

rds_df <- readRDS(choose.files(default = paste0(user_directory,
                                                "/R_Sched_AM_Repo/*.*"),
                               caption = "Select schedule repository"))

# test_df <- data.frame("ApptNote" = unique(rds_df$`Appt Notes`))

# rds_df <- rds_df %>%
#   mutate(TestNull = ApptNote == "",
#          NewNote = gsub("^$|^ $", NA, ApptNote))

rds_df2 <- as.data.frame(rds_df)

rds_df2 <- rds_df2 %>%
  mutate(ApptNote2 = gsub("^$|^ $", NA, `Appt Notes`))

test_df <- rds_df %>%
  # filter(`Appt Notes` ==)
  mutate(`Appt Notes` = ifelse(is.na(`Appt Notes`), "", `Appt Notes`),
         `Canc Date` = ifelse(is.na(`Canc Date`), "", `Canc Date`),
         `MyChart Status` = ifelse(is.na(`MyChart Status`), "", `MyChart Status`),
         Race = ifelse(is.na(Race), "", Race),
         Immunizations = ifelse(is.na(Immunizations), "", Immunizations),
         `MOUNT SINAI HEALTH SYSTEM` = ifelse(is.na(`MOUNT SINAI HEALTH SYSTEM`),
                                              "", `MOUNT SINAI HEALTH SYSTEM`),
         `PATIENTS LIFE NUMBER OR ORACLE NUMBER` = ifelse(
           is.na(`PATIENTS LIFE NUMBER OR ORACLE NUMBER`), "",
           `PATIENTS LIFE NUMBER OR ORACLE NUMBER`),
         `Level Of Service` = ifelse(is.na(`Level Of Service`), "",
                                     `Level Of Service`))
         

# 
# rds_df2 <- rds_df2 %>%
#   mutate(`Appt Notes` = NULL)
# NOt SURE HOW TABLEAU HANDLES NA

save(rds_df2, file = paste0(user_directory,
                           "/RData Exports/",
                           "RData Test DF2.rda"))

test_df <- data.frame("Site" = c("Site1", "Site2", "Site3", "Site4/"),
                      "Value" = c(1, NA, 7, 150))

save(test_df, file = paste0(user_directory,
                            "/RData Exports/",
                            "Simple DF Test.rda"))

raw_df1 <- read.csv(choose.files(
  default =
    paste0(user_directory, "/Epic Auto Report Sched All Status/*.*"),
  caption = "Select current Epic schedule"),
  stringsAsFactors = FALSE,
  check.names = FALSE)

raw_df2 <- read.csv(choose.files(
  default =
    paste0(user_directory, "/Epic Auto Report Sched All Status/*.*"),
  caption = "Select current Epic schedule"),
  stringsAsFactors = FALSE,
  check.names = FALSE,
  na.strings = c(""))

# raw_df2 <- raw_df2 %>%
#   mutate(`MOUNT SINAI HEALTH SYSTEM` = NULL,
#          `PATIENTS LIFE NUMBER OR ORACLE NUMBER` = NULL)

raw_df_test <- raw_df2

raw_df_test <- raw_df_test %>%
  mutate(DOB = as.Date(DOB, "%m/%d/%Y"),
         `Made Date` = as.Date(`Made Date`, "%m/%d/%y"),
         Date = as.Date(Date, "%m/%d/%Y"),
         `Canc Date` = as.Date(`Canc Date`, "%m/%d/%Y"),
         MyChartStatus = ifelse(`MyChart Status` %in%
                                     c("Non-Standard Mychart Status",
                                       "Activation Code Generated, But Disabled"),
                                   "Other", `MyChart Status`),
         `MyChart Status` = NULL)

raw_df_test <- raw_df_test[, c(10:14, 26)]

# raw_df_test <- raw_df_test[complete.cases(raw_df_test), ]

save(raw_df_test, file = paste0(user_directory,
                            "/RData Exports/",
                            "ImportCSVWnastrings.rdx")),
     ascii = FALSE)
