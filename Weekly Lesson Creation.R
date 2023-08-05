library(tidyverse)
library(dplyr)
library(lubridate)
library(stringr)
library(openxlsx)

# Choose your week reference; this should match what's in the week column in 
# your timetable file.
week <- 1


lesson_template_name <- "Lesson Slide Template.pptx"
lesson_template_folder <- "C://Chris//Teaching//Templates"
target_folder <- "C://Chris//Teaching//This Week"
archive_folder <- "C://Chris//Teaching//Archive//2022-2023"


# find timetable
timetable <- read.xlsx("C://Chris//Teaching//Timetable.xlsx") %>% 
  rename(Week = 1) %>% 
  filter(Week == week)


# function for finding date of any given next weekday
nextweekday <- function(date, wday) {
  date <- as.Date(date)
  diff <- wday - wday(date)
  if( diff < 0 )
    diff <- diff + 7
  return(date + diff)
}


lesson_template_file_location <- paste0(lesson_template_folder, "//", lesson_template_name)



### move last week's pptx into archive

for(xxx in unique(timetable[,4])){
  lw_powerpoints <- list.files(target_folder)
  powerpoints_to_archive <- lw_powerpoints[str_detect(lw_powerpoints,paste0(xxx,".pptx"))]
  for(yyy in powerpoints_to_archive){
    date_check <- yyy %>% str_sub(1, 10) %>% as.Date() # check that they're past lessons!
    if(date_check < Sys.Date()){
      file.rename(from = paste0(target_folder,"//",yyy),
                  to = paste0(archive_folder,"//",xxx,"//",yyy))
    }}}



### create this week's pptx files

for(xxx in 1:nrow(timetable)){
  file.copy(lesson_template_file_location, target_folder)
  file.rename(paste0(target_folder, "//", lesson_template_name),
              paste0(target_folder, "//", nextweekday(Sys.Date(), timetable[xxx,2]+2),
                     " P", timetable[xxx,3], " ", timetable[xxx,4], ".pptx"))
}

