library(tidyverse)
library(dplyr)
library(lubridate)
library(stringr)
library(openxlsx)

### USER INPUTS

# Choose your week reference; this should match what's in the week column in 
# your timetable file.

week <- 1

# Choose the main folder in which you're working. Make sure this contains subfolders
# labelled 'Archive', 'Templates', and 'This Week'

active_folder <- "C://Chris//Teaching"

# Choose the name for the archive subfolder in which you want archival lessons
# to be placed.

archive_folder_name <- "2022-2023"



### EDITABLE INPUTS

lesson_template_name <- "Lesson Slide Template.pptx"


### PROGRAM

# generates folder names from your main folder

target_folder <- paste0(active_folder, "//This Week//")
archive_folder <- paste0(active_folder, "//Archive//", archive_folder_name, "//")
lesson_template_file_location <- paste0(active_folder, "//Templates//", lesson_template_name)


# import timetable
timetable <- read.xlsx("C://Chris//Teaching//Timetable.xlsx") %>% 
  rename(Week = 1) %>% 
  filter(Week == week)




# move last week's pptx into archive

for(xxx in unique(timetable[,4])){
  lw_powerpoints <- list.files(target_folder)
  powerpoints_to_archive <- lw_powerpoints[str_detect(lw_powerpoints,paste0(xxx,".pptx"))]
  for(yyy in powerpoints_to_archive){
    date_check <- yyy %>% str_sub(1, 10) %>% as.Date() # check that they're past lessons!
    if(date_check < Sys.Date()+9){
      file.rename(from = paste0(target_folder,"//",yyy),
                  to = paste0(archive_folder,"//",xxx,"//",yyy))
    }}}


# function for finding date of any given next weekday
nextweekday <- function(date, wday) {
  date <- as.Date(date)
  diff <- wday - wday(date)
  if( diff < 0 )
    diff <- diff + 7
  return(date + diff)
}


# create this week's pptx files

for(xxx in 1:nrow(timetable)){
  file.copy(lesson_template_file_location, target_folder)
  file.rename(paste0(target_folder, "//", lesson_template_name),
              paste0(target_folder, "//", nextweekday(Sys.Date(), timetable[xxx,2]+2),
                     " P", timetable[xxx,3], " ", timetable[xxx,4], ".pptx"))
}

