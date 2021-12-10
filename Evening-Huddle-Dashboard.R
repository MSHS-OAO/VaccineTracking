rm(list = ls())

if ("Presidents" %in% list.files("J://")) {
  user_directory <- paste0("J:/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "System Operations/COVID Vaccine")
} else {
  user_directory <- paste0("J:/deans/Presidents/HSPI-PM/",
                           "Operations Analytics and Optimization/Projects/",
                           "System Operations/COVID Vaccine")
}

rmarkdown::render("Combined Dashboard Arrived and Sched Appts.Rmd",
                  output_file = paste0(user_directory,
                                       "/Daily PM Huddle Output/",
                                       "MSHS Vaccine Summary ",
                                       format(Sys.Date(), "%m-%d-%y")))

