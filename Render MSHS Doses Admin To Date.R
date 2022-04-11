# This code renders the COVID-19 Vaccine Administration data for the health system
# This dashboard looks at historical doses administered by setting type (hospital POD vs practices)
# and is less focused on scheduled appointments in the coming weeks.

# Clear environment
rm(list = ls())

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

# Render markdown file with dashboard code and save with today's date
rmarkdown::render(here::here("MSH Admin Dose Summary.Rmd"), 
                  output_file = paste0(
                    user_directory,
                    "/Weekly Dashboard Doses Administered",
                    "/MSHS COVID-19 Vaccine Administration Summary ",     
                    format(Sys.Date(), "%m-%d-%y")))
