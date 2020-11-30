#MSHQ Payroll

##Libraries and Functions######################################################
library(dplyr)
library(tidyr)
library(lubridate)
library(readxl)
library(xlsx)

#Read most recent MSHQ payroll file, filter on dates and format columns
labor <- function(start,end){
  paycode <- read_excel("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Useful Tools & Templates/Pay Code Mappings/MSHQ_Paycode Vlookup_GL.xlsx") %>%
    select(1,2)
  df <- file.info(list.files("C:/Users/lenang01/Documents/MSH-MSQ-Payroll/MSHQ Oracle", full.names = T))
  df <- read.csv(rownames(df)[which.max(df$mtime)], header = T,sep = "~",stringsAsFactors = F,colClasses = rep("character",32)) %>%
    filter(as.Date(End.Date,format = "%m/%d/%Y") <= as.Date(end,format = "%m/%d/%Y"),
           as.Date(Start.Date,format = "%m/%d/%Y") >= as.Date(start,format = "%m/%d/%Y"),
           !is.na(Job.Code))
  df <- left_join(df,paycode,by=c("Pay.Code"="Full Paycode")) %>%
    mutate(Pay.Code = `Pay Code in Premier`,
           `Pay Code in Premier` = NULL,
           Employee.Name = substr(Employee.Name,1,30),
           Department.Name.Worked.Dept = substr(Department.Name.Worked.Dept,1,50),
           Department.Name.Home.Dept = substr(Department.Name.Home.Dept,1,50),
           Department.ID.Home.Department = paste0(substr(Full.COA.for.Home,1,3),substr(Full.COA.for.Home,41,44),substr(Full.COA.for.Home,5,7),substr(Full.COA.for.Home,12,16)))
  return(df)
}
#Create JCdict and find new job codes
jcdict <- function(end){
  jobcode <- read_excel("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Useful Tools & Templates/Job Code Mappings/MSH MSQ Position Mappings.xlsx")
  df <- left_join(df,jobcode,by=c("Job.Code"="J.C"))
  newjc <- filter(df,is.na(J.C.DESCRIPTION))
  if(nrow(newjc) > 0){
    newjc <- newjc %>% select(Job.Code,Position.Code.Description) %>% distinct() 
    write.xlsx(newjc,paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/MSH/Payroll/MSH Labor/Calculation Worksheets/NewJC/New_Job_Codes_",Sys.Date(),".xlsx"))
    message("There are new Job Codes that need to be added to the Job Code mappings File J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Useful Tools & Templates/Job Code Mappings/MSH MSQ Position Mappings.xlsx")
  } else {
    df <- df %>%
      filter(is.na(Provider)) %>%
      mutate(Position.Code.Description = `Description in Premier (50 character limit)`)
  }
  jcdict <- df %>%
    select(PartnerOR.Health.System.ID,Home.FacilityOR.Hospital.ID,Department.IdWHERE.Worked,Job.Code,Position.Code.Description) %>%
    distinct()
  mon <- toupper(month.abb[month(as.Date(end,format = "%m/%d/%Y"))])
  write.table(jcdict,paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/MSH/Payroll/MSH Labor/Calculation Worksheets/JCDict/MSHQ_JCdict_",substr(end,4,5),mon,substr(end,7,11),".csv"),sep=",",row.names = F,col.names = F)
  return(df)
}
#Create Department dict
depdict <- function(end){
  home <- df %>% select(PartnerOR.Health.System.ID,Home.FacilityOR.Hospital.ID,Department.ID.Home.Department,Department.Name.Home.Dept)
  worked <- df %>% select(PartnerOR.Health.System.ID,Facility.Hospital.Id_Worked,Department.IdWHERE.Worked,Department.Name.Worked.Dept)
  col <- c("Partner","Hosp","CC","CC.Description")
  colnames(home) <- col
  colnames(worked) <- col
  depdict <- rbind(home,worked) %>% distinct()
  mon <- toupper(month.abb[month(as.Date(end,format = "%m/%d/%Y"))])
  write.table(depdict,paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/MSH/Payroll/MSH Labor/Calculation Worksheets/DepDict/MSHQ_DepDict_",substr(end,4,5),mon,substr(end,7,11),".csv"),sep=",",row.names = F,col.names = F)
}
#Creates Department Mapping file for new departments
depmap <- function(end){
  depmap <- file.info(list.files("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/MSH/Payroll/MSH Labor/Dep Mapping Downloads", full.names = T,pattern = ".csv"))
  depmap <- read.csv(rownames(depmap)[which.max(depmap$mtime)], header = F,stringsAsFactors = F) %>% distinct()
  depmap <- left_join(df,depmap,by=c("Department.IdWHERE.Worked"="V3")) %>%
    mutate(Effective = "01012010")
  newdep <- depmap %>% filter(is.na(V5)) %>% select(Effective,PartnerOR.Health.System.ID,Facility.Hospital.Id_Worked,Department.IdWHERE.Worked,V5) %>% distinct()
  if(nrow(newdep) > 0){
    newdep <- newdep %>% mutate(V5 = "10095")
    depmap <- depmap %>% filter(!is.na(V5)) %>% select(Effective,PartnerOR.Health.System.ID,Facility.Hospital.Id_Worked,Department.IdWHERE.Worked,V5) %>% distinct()
    depmap <- rbind(depmap,newdep)
  } else {
    depmap <- depmap %>% filter(!is.na(V5)) %>% select(Effective,PartnerOR.Health.System.ID,Facility.Hospital.Id_Worked,Department.IdWHERE.Worked,V5) %>% distinct() 
  }
  mon <- toupper(month.abb[month(as.Date(end,format = "%m/%d/%Y"))])
  write.table(depmap,paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/MSH/Payroll/MSH Labor/Calculation Worksheets/DepMap/MSHQ_DepMap_",substr(end,4,5),mon,substr(end,7,11),".csv"),sep=",",row.names = F,col.names = F)
  return(depmap)
}
#Create JC mapping file
jcmap <- function(end){
  jcmap <- left_join(df,depmap,by=c("Department.IdWHERE.Worked"="Department.IdWHERE.Worked")) 
  newdep <- filter(jcmap,is.na(V5)) 
  if(nrow(newdep) > 0){
    message("Deparment mapping was not updated correctly")
  }
  jcmap <- jcmap %>% select(Effective,PartnerOR.Health.System.ID.x,Facility.Hospital.Id_Worked.x,Department.IdWHERE.Worked,Job.Code,V5,`Premier ID Code`) %>% mutate(Allocation = "100") %>% distinct()
  mon <- toupper(month.abb[month(as.Date(end,format = "%m/%d/%Y"))])
  write.table(jcmap,paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/MSH/Payroll/MSH Labor/Calculation Worksheets/JCmap/MSHQ_JCMap_",substr(end,4,5),mon,substr(end,7,11),".csv"),sep=",",row.names = F,col.names = F)
}
#Create payroll upload
upload <- function(start,end){
  payroll <- df %>%
    mutate(Approved = "0",
           Hours = round(as.numeric(Hours),2),
           Expense = round(as.numeric(Expense),2)) %>% 
    group_by(PartnerOR.Health.System.ID,Home.FacilityOR.Hospital.ID,Department.ID.Home.Department,Facility.Hospital.Id_Worked,Department.IdWHERE.Worked,Start.Date,End.Date,Employee.ID,Employee.Name,Approved,Job.Code,Pay.Code) %>%
    summarise(Hours = sum(Hours,na.rm = T),
              Expense = sum(Expense,na.rm = T))
  return(payroll)
}
#Trend worked Hours by cost center
worktrend <- function(){
  paycycle <- read_excel("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Useful Tools & Templates/Pay Cycle Calendar.xlsx",col_types = "date") %>%
    select(10,12) %>%
    mutate(Date = as.Date(Date,format="%Y-%M-%D"),
            End.Date = as.Date(End.Date,format="%Y-%M-%D"))
  paycode <- read_excel("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Useful Tools & Templates/Pay Code Mappings/MSHQ_Paycode Vlookup_GL.xlsx") %>%
    select(c(2:5)) %>%
    filter(nchar(`Pay Code in Premier`) > 3) %>%
    distinct()
  payroll <- payroll %>% ungroup() %>% mutate(End.Date = as.Date(End.Date,format="%m/%d/%Y"))
  trend <- left_join(payroll,paycycle,by=c("End.Date"="Date"))
  trend <- left_join(trend,paycode,by=c("Pay.Code"="Pay Code in Premier")) 
  trend <- trend %>%
    filter(`Premier Pay Map` %in% c("REGULAR","OTHER_WORKED","OVETIME"),
           `Include Hours` == 1) %>%
    group_by(Department.IdWHERE.Worked,End.Date.y) %>%
    summarise(Hours = sum(Hours,na.rm=T)) %>%
    rename(PP.END.DATE = End.Date.y) 
  oldtrend <- readRDS("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/MSH/Payroll/MSH Labor/Calculation Worksheets/Worked Trend/trend.RDS")
  oldtrend <- mutate(oldtrend,PP.END.DATE = as.Date(PP.END.DATE,formate="%Y-%m-%d"))
  trend <- rbind(oldtrend,trend) %>%
    arrange(PP.END.DATE) %>%
    mutate(PP.END.DATE = factor(PP.END.DATE))
  trend <<- trend
  new_trend <- trend %>%pivot_wider(id_cols = Department.IdWHERE.Worked,names_from = PP.END.DATE,values_from = Hours)
  return(new_trend)
}
#Save payroll file
save_payroll <- function(start,end){
  payroll <- payroll %>% mutate(End.Date = paste0(substr(End.Date,6,7),"/",substr(End.Date,9,10),"/",substr(End.Date,1,4)))
  smon <- toupper(month.abb[month(as.Date(start,format = "%m/%d/%Y"))])
  emon <- toupper(month.abb[month(as.Date(end,format = "%m/%d/%Y"))])
  write.table(payroll,paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/MSH/Payroll/MSH Labor/Calculation Worksheets/Uploads/MSHQ_Payroll_",substr(start,4,5),smon,substr(start,7,11)," to ",substr(end,4,5),emon,substr(end,7,11),".csv"),sep=",",row.names = F,col.names = F)
  saveRDS(trend,"J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/MSH/Payroll/MSH Labor/Calculation Worksheets/Worked Trend/trend.RDS")
  write.table(new_trend,paste0("J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Labor - Data/MSH/Payroll/MSH Labor/Calculation Worksheets/Worked Trend/CC Worked Trend_",substr(end,4,5),mon,substr(end,7,11),".csv"),sep=",",row.names = F,col.names = T)
}
###############################################################################

#Enter start and end date needed for payroll upload
start <-"09/27/2020" 
end <- "10/24/2020"
df <- labor(start,end)
#If you need to update jobcode list for new jobcodes leave R and do that in excel
#"J:/deans/Presidents/SixSigma/MSHS Productivity/Productivity/Useful Tools & Templates/Job Code Mappings/MSH MSQ Position Mappings.xlsx"
df <- jcdict(end)
depdict(end)
#Download and place department mapping file in MSH Labor folder
depmap <- depmap(end)
jcmap(end)
payroll <- upload(start,end)
new_trend <- worktrend()
#If new_trend looks good then save upload
save_payroll(start,end)
