rmarkdown::render(input = "MSHS_Depts_NotReporting_VaccineStatus.Rmd",
                  output_file = paste0(
                    user_directory,
                    "/AdHoc",
                    "/Vaccine Status Non Reporting Depts",
                    "/MSHS Vaccination Status - Depts Not Reporting ",
                    format(Sys.Date(), "%m-%d-%y")))
