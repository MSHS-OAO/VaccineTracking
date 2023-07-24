# This code renders the COVID-19 Vaccine Administration data for the health system
# This dashboard looks at historical doses administered by setting type (hospital POD vs practices)
# and is less focused on scheduled appointments in the coming weeks.

# Clear environment
rm(list = ls())

# Determine directory
#Function to determine path to share drive on R Workbench or R Studio

define_root_path <- function(){
  #Check if directory is from R Workbench; starts with '/home'
  if(grepl("^/home", dirname(getwd()))){
    #Check if mapped Sharedrvie starts at folder Presidents or deans
    ifelse(list.files("/SharedDrive/") == "Presidents",
           #Define prefix of path to share drive with R Workbench format
           output <- "/SharedDrive/Presidents/", 
           output <- "/SharedDrive/deans/Presidents/")
  }#Check if directory is from R Studio; starts with an uppercase letter than ':'
  else if(grepl("^[[:upper:]]+:", dirname(getwd()))){
    #Determine which drive is mapped to Sharedrive (x)
    for(i in LETTERS){
      if(any(grepl("deans|Presidents", list.files(paste0(i, "://"))))){x <- i}
    }
    #Check if mapped Sharedrvie starts at folder Presidents or deans
    ifelse(list.files(paste0(x, "://")) == "Presidents",
           #Define prefix of path to share drive with R Studio format
           output <- paste0(x, ":/Presidents/"),
           output <- paste0(x, ":/deans/Presidents/"))
    
  }
  return(output)
}

user_directory <- define_root_path()

initial_run <- TRUE
update_repo <- FALSE

# Render markdown file with dashboard code and save with today's date
rmarkdown::render(here::here("MSHS Vaccines Administered Hybrid Data Model.Rmd"), 
                  output_file = paste0(
                    user_directory,
                    "/Hybrid Sched & Admin Reporting",
                    "/Hybrid Model Dashboard",
                    "/MSHS COVID-19 Vaccines Administered ",     
                    format(Sys.Date(), "%m-%d-%y")))
