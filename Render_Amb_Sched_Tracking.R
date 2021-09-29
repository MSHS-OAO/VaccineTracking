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

rmarkdown::render(input = "Amb_Sched_Tracking.Rmd",
                  output_file = paste0(user_directory,
                                       "/Dose3 Amb Sched Tracking",
                                       "/Dose3 Schedule Tracking ",
                                       format(Sys.time(), "%Y-%m-%d %I%M%p")))