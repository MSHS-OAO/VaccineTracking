# Code to import .RDS file and save as .RData

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

rds_df <- readRDS(choose.files(default = paste0(user_directory,
                                                "/R_Sched_AM_Repo/*.*"),
                               caption = "Select schedule repository"))

save(rds_df, file = paste0(user_directory,
                           "/RData Exports/",
                           "RData Test.RData"))
