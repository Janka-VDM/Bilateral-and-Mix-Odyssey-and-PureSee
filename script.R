####LIBRARIES####

library("stringr")
library("DBI")
library("odbc")
library("excel.link")
library("dplyr")
library("wktmo")
library("excel.link")
library("lubridate")
library("ggplot2")
library("tidyverse")
library("reshape2")
library("patchwork")
library('readxl')
library('janitor')
library('openxlsx')

####FILE PATH####

path <- "S:/Clinical Services/Research Assistants/Tecnis Odyssey/Bilateral Odyssey vs Mix PureSee/"

con <- DBI::dbConnect(odbc::odbc(), 
                      Driver = "SQL Server", 
                      Server = "GB-SQL-N2\\I02", 
                      Trusted_Connection = "True")

mindate <- '2024-01-01'
maxdate <- '2026-09-30'
surgerytype <- 2

#### SQL Query ####

source("query.R")
data_iolscript <- dbGetQuery(con, query)
data <- data_iolscript

query1 <- paste('select distinct
s.PatientID
from emr.dbo.tSurgery as s

left join emr.dbo.tEyeSurgeryEye as ese on ese.SurgeryID=s.SurgeryID
left join Estate.dbo.tLocation_Region as lr on lr.LocationID = s.SurgeryLocationID
left join emr.dbo.tLookupSurgeryTreatmentType as lstt on lstt.SurgeryTreatmentTypeID = ese.SurgeryTreatmentTypeID
left join emr.dbo.tLookupSurgeryType as lst on lst.SurgeryTypeID = ese.SurgeryTypeID
Left Join EMR.dbo.tEyeSurgeryEyeIOL ESEI On ESEI.SurgeryID = S.SurgeryID and ESEI.EyeID = ESE.EyeID
LEFT JOIN EMR.dbo.tIOLLensCalculation AS ilc ON ilc.IOLLensCalculationID = esei.Type1IOLLensCalculationID

where s.SurgeryDateTime between ', "'", mindate , "'", ' AND ', "'", maxdate , "'",'
and s.IsDeleted IN (0)
and lr.RegionID in (1,2,3,4,5,6)
and lst.SurgeryTypeGroupID in (', surgerytype, ')
and lstt.SurgeryTreatmentTypeID = 1
and ilc.IOLLensID in (101,103,106,105,108,109,110,111,112)
and ese.SurgeryTypeID not in (23,24) --exchanges
and s.PatientID not in (63197340, 63250189, 63275221, 63277379, 63301854, 63302830, 63317555, 63322931, 63345674) --these patients do not have manifest data
and ese.IsComplete IN (1)',sep='')

checking <- dbGetQuery(con,query1)
check <- checking[which(!(checking$PatientID %in% data_iolscript$PatientID)),] # check for any patients that are in checking but not in data
if (length(check)>0){
  print(check)
  stop("These patients are missing from the data.")
}

#### DATA CLEANING ####
##### Dominant EyeIDs ####
data$DominantEyeID[which(data$PatientID=="63259396")] <- 2
data$DominantEyeID[which(data$PatientID=="63222368")] <- 1
data$DominantEyeID[which(data$PatientID=="63424225")] <- 2
data <- subset(data, !(data$PatientID=="63269887" & data$SurgeryID==830801))
data <- subset(data, !(data$PatientID=="63298212" & data$SurgeryID==843125))
data <- subset(data, !(data$PatientID=="63297875" & data$SurgeryID==845648))
data <- subset(data, !(data$PatientID=="60716226" & data$SurgeryID==848106))
data <- subset(data, !(data$PatientID=="63288348" & data$SurgeryID==863404))
data <- subset(data, !(data$PatientID=="63119724" & data$SurgeryID==778069))
data <- subset(data, !(data$PatientID=="62516032" & as.Date(data$ExamDateTime)=='2019-09-30'))
data <- subset(data, data$SurgeryID!=858290)
data <- data %>% ungroup()

#DominantEyeID you cannot get from Clinical Reviews, so lets just fill that in from the first table:
#checking that everyone only has one DominantEyeID (if they ever got it measured):
DECheck <- data %>%
  ungroup() %>%
  select(PatientID, DominantEyeID) %>%
  filter(!is.na(DominantEyeID)) %>%
  unique() %>%
  group_by(PatientID) %>%
  summarise(Count=n())
if (dim(DECheck[which(DECheck$Count>1),])[1] >0){
  stop("Patients have more than one DominantEyeID! Check 'DECheck'.")
}
#for patients that never had a DominantEyeID (like some LVC patients), these next few commands will not do anything to them, and they are not included in DE
DE <- data %>%
  select(PatientID, DominantEyeID) %>%
  unique %>%
  group_by(PatientID) %>%
  filter(!is.na(DominantEyeID)) %>%
  rename(DominantEyeLU = DominantEyeID)

data <- merge(data,DE, by="PatientID", all.x=TRUE)
data$DominantEyeID <- data$DominantEyeLU
data <- data %>%
  select(-DominantEyeLU) %>%
  unique()

##### Getting max amount of PreOp Data ####
data$max <- rowSums(!is.na(data %>% select(ExamDateTime, EyeID, DominantEyeID, Motivation, Target, 
                                           Sphere, Cyl, Keratometry_Power_1, Keratometry_Power_2, 
                                           UCDVA_SM, UCDVA_LogMAR, UCDVA_N, UCIVA_SM, UCIVA_LogMAR, 
                                           UCIVA_N, UCNVA_SM, UCNVA_LogMAR, UCNVA_N, Binoc_UCDVA_SM, 
                                           Binoc_UCDVA_LM, Binoc_UCIVA_SM, Binoc_UCIVA_LM, Binoc_UCIVA_N, 
                                           Binoc_UCNVA_SM, Binoc_UCNVA_LM, Binoc_UCNVA_N, BCDVA_LM, 
                                           Lens)))

data <- data %>%
  group_by(PatientID, EyeID, SurgeryDateTime) %>%
  filter(max==max(max)) %>%
  select(-max)

# if both exam types have the same number of pre op entries, so both remain from 
# the previous code, so we get the latest one:
data <- data %>%
  group_by(PatientID, EyeID) %>%
  filter(ExamDateTime==max(ExamDateTime))

data$Model[which(data$PatientID=="60381493" & data$EyeID=="1")] <- "Rayner Sulcoflex Multifocal Toric"
data$Target[which(data$PatientID=="63294791")] <- 0
data$Target[which(data$PatientID=="63323587")] <- 0
data$Target[which(data$PatientID=="63336246")] <- 0

##### Data Preparation ####
data$MSE <- round(data$Sphere + (data$Cyl/2), digits = 2)
data$MSE_M1 <- round(data$Sph_M1 + (data$Cyl_M1/2), digits = 2)

data$Sph_ach_M1 <- round(data$Sphere - data$Sph_M1, digits = 2)

data$KTest <- data$Keratometry_Power_1 - data$Keratometry_Power_2

data <- data %>%
  mutate(M1_Binoc_UCDVA_D = case_when(M1_Binoc_UCDVA_D=="CF" ~ 0.016,
                                      M1_Binoc_UCDVA_D=="HM" ~ 0.004,
                                      M1_Binoc_UCDVA_D=="POL" ~ 0.002,
                                      M1_Binoc_UCDVA_D=="NPOL" ~ 0.001,
                                      .default = as.numeric(M1_Binoc_UCDVA_D)),
         M1_Binoc_BCDVA_D = case_when(M1_Binoc_BCDVA_D=="CF" ~ 0.016,
                                      M1_Binoc_BCDVA_D=="HM" ~ 0.004,
                                      M1_Binoc_BCDVA_D=="POL" ~ 0.002,
                                      M1_Binoc_BCDVA_D=="NPOL" ~ 0.001,
                                      .default = as.numeric(M1_Binoc_BCDVA_D)),
         Binoc_BCDVA_D = case_when(Binoc_BCDVA_D=="CF" ~ 0.016,
                                      Binoc_BCDVA_D=="HM" ~ 0.004,
                                      Binoc_BCDVA_D=="POL" ~ 0.002,
                                      Binoc_BCDVA_D=="NPOL" ~ 0.001,
                                      .default = as.numeric(Binoc_BCDVA_D)))

data$Model <- gsub("\\bAMO\\b", "J&J", data$Model)
data$Exch_Model <- gsub("\\bAMO\\b", "J&J", data$Exch_Model)

data <- data %>% ungroup()

##### Last Checks ####
# this is not at the end because otherwise a lot of patients would pop up due
# to the data conditions
test <- checking[which(!(checking$PatientID %in% data$PatientID)),]
if (length(test)>0){
  print(test)
  stop("We have lost these patients!")
}
#checking if we have any duplicates
test <- as.data.frame(table(data$PatientID, data$EyeID, data$SurgeryDateTime))
if (dim(test[which(test$Freq>1),])[1]>1){
  print(test[which(test$Freq>1),])
  stop("There are patients with multiple entries!")
}

#### DATA PREPARATION ####
##### Data Conditions ####
# data <- data %>% filter(PriorLVC==1)
data <- data[which(is.na(data$Complication)),]
data <- data[which(!is.na(data$Month1PostOp)),]
data <- data[which(!is.na(data$M1_SurveyDate)),]
data <- data[which(data$Age>=45),]
data$MSE <- round(data$Sphere + (data$Cyl/2), digits = 2)
data$MSE_M1 <- round(data$Sph_M1 + (data$Cyl_M1/2), digits = 2)

data <- data %>%
  filter((EyeID==DominantEyeID & abs(MSE_M1) <= 0.5) | (EyeID!=DominantEyeID & abs(MSE_M1) <= 1))

data$LensType <- NA
data$LensType[grep("Odyssey", data$Model,ignore.case=TRUE)] <- "Odyssey"
data$LensType[grep("PureSee", data$Model,ignore.case=TRUE)] <- "PureSee"

##### Creating Bilat/Mix ####
bilat1 <- data %>%
  filter(LensType=="Odyssey") %>%
  filter(n_distinct(EyeID) == 2) %>%
  mutate(SurgeryLens = "Bilateral Odyssey") %>%
  select(PatientID, SurgeryDateTime, SurgeryLens)

bilat2 <- data %>%
  filter(LensType=="PureSee") %>%
  filter(n_distinct(EyeID) == 2) %>%
  mutate(SurgeryLens = "Bilateral Odyssey") %>%
  select(PatientID, SurgeryDateTime, SurgeryLens)

# puresee in dominant eye, odyssey in non dominant eye
mix <- data %>%
  mutate(Combo = case_when(EyeID == DominantEyeID & LensType=="PureSee" ~ "PS Dominant",
                           EyeID != DominantEyeID & LensType=="Odyssey" ~ "OD Not-Dominant",
                           .default=NA)) %>%
  group_by(PatientID,SurgeryDateTime) %>%
  arrange(PatientID, Combo) %>%
  filter(!is.na(Combo)) %>%
  filter(n_distinct(Combo)==2)%>%
  mutate(SurgeryLens = "Mix and Match") %>%
  select(PatientID, SurgeryDateTime, SurgeryLens)

temp <- rbind(bilat1, bilat2, mix)

data <- merge(data, temp, by=c("PatientID","SurgeryDateTime")) %>% unique()

test <- as.data.frame(table(data$PatientID, data$EyeID)) %>% filter(Freq>1)
if (dim(test)[1]>0){view(test)
  stop("There are duplicates")}

surgeons <- data %>% select(SurgeonID) %>% unique() %>% nrow()
locations <- data %>% select(SurgeryLocationID) %>% unique() %>% nrow()

##### Reasons Cleaning ####
data$Motivation[which(data$PatientID=="63183343")] <- "Rid of Spectacles and Contact Lenses"
data$Reason <- data$Motivation
data$Reason[which(data$Lens=="Cataract")] <- "Cataract"

r <- c("Rid of Spectacles", "Rid of Contact Lenses", "Rid of Spectacles and Contact Lenses", "Convenience", "Sport / Hobbies Related",
       "Cosmetic", "Occupational")
data <- data %>%
  mutate(Reason = case_when((Reason %in% r) & Sphere < 0 ~ "Improve Distance Vision",
                            (Reason %in% r) & Sphere > 0 ~ "Improve Near Vision",
                            (Reason %in% r) & Sphere == 0 & AddNear > 0 ~ "Improve Near Vision",
                            !(Reason %in% r) ~ Reason))

test <- data[which(is.na(data$Motivation)),]
if (dim(test)[1]>0){
  view(test)
  stop("There are patients who do not have any Motivations.")
}

write.xlsx(data %>% select(PatientID, EyeID, Reason, M1_UCDVA_LM),
           "S:\\Clinical Services\\Research Assistants\\Tecnis Odyssey\\Bilateral Odyssey vs Mix PureSee\\tables\\pxsreasons.xlsx")

##### quick peq test ####
cids <- toString(unique(data$PatientID))
querytest <- paste('select distinct
rh.PatientID,
rh.SurveyDate,
s.SurveyID,
s.Name as Survey
  FROM peq.dbo.tResultHeader rh (NOLOCK)

JOIN peq.dbo.tResultLine rl (NOLOCK) ON rl.ResultHeaderID = rh.ResultHeaderID
left join PEQ.dbo.tSurvey as s on s.SurveyID = rh.SurveyID

WHERE COALESCE(rl.IsDeleted, 0) = 0
and rh.Participated in (1)
and rh.PatientID in (', cids, ')', sep='')

test <- dbGetQuery(con,querytest)

test <- merge(test, data %>% select(PatientID, EyeID, SurgeryDateTime), by="PatientID")
test <- test[which(test$SurveyDate >= test$SurgeryDateTime),]
test <- test %>%
  select(-EyeID) %>%
  unique()
test$DateDiff <- difftime(test$SurveyDate, test$SurgeryDateTime, units="days")

write.xlsx(data,"newdata.xlsx")

#### Patients Saving for PEQs ####
data <- data %>% ungroup() %>% unique()
write.xlsx(data, paste(path,"tables/pxs.xlsx", sep=""))

#### DEMOGRAPHICS ####
Info <- c("Surgeons","Location")
Number <- c(surgeons,locations)
surgloc <- data.frame(Info, Number)

eyes <- as.data.frame(nrow(data[which(data$SurgeryLens=="Bilateral Odyssey"),]))
patients <- as.data.frame(unique(data$PatientID[which(data$SurgeryLens=="Bilateral Odyssey")]))
patients <- as.data.frame(nrow(patients))
totalo <- cbind(patients, eyes)
colnames(totalo) <- c("Patients", "Eyes")

eyes <- as.data.frame(nrow(data[which(data$SurgeryLens=="Bilateral PureSee"),]))
patients <- as.data.frame(unique(data$PatientID[which(data$SurgeryLens=="Bilateral PureSee")]))
patients <- as.data.frame(nrow(patients))
totalp <- cbind(patients, eyes)
colnames(totalp) <- c("Patients", "Eyes")

eyes <- as.data.frame(nrow(data[which(data$SurgeryLens=="Mix and Match"),]))
patients <- as.data.frame(unique(data$PatientID[which(data$SurgeryLens=="Mix and Match")]))
patients <- as.data.frame(nrow(patients))
totalm <- cbind(patients, eyes)
colnames(totalm) <- c("Patients", "Eyes")

u <- data %>%
  filter(SurgeryLens=="Bilateral Odyssey") %>%
  select(PatientID, Age, Gender) %>%
  unique()

mean <- round(mean(u$Age), digits = 2)
min <- min(u$Age)
max <- max(u$Age)
sd <- round(sd(u$Age), digits = 2)
ageo <- cbind(mean, min, max, sd)

gendero <- as.data.frame(table(u$Gender))
gendero$percent <- round(gendero$Freq/sum(gendero$Freq), digits = 4)

datao <- data %>%
  filter(SurgeryLens=="Bilateral Odyssey") %>%
  select(PatientID, EyeID, Sphere, Cyl, MSE, KTest, Toric) %>%
  unique()

mean <- round(mean(datao$Sphere, na.rm = TRUE), digits = 2)
min <- round(min(datao$Sphere, na.rm = TRUE), digits = 2)
max <- round(max(datao$Sphere, na.rm = TRUE), digits = 2)
sd <- round(sd(datao$Sphere, na.rm = TRUE), digits = 2)
sphereo <- cbind(mean, min, max, sd)

mean <- round(mean(datao$Cyl, na.rm = TRUE), digits = 2)
min <- round(min(datao$Cyl, na.rm = TRUE), digits = 2)
max <- round(max(datao$Cyl, na.rm = TRUE), digits = 2)
sd <- round(sd(datao$Cyl, na.rm = TRUE), digits = 2)
cylo <- cbind(mean, min, max, sd)

mean <- round(mean(datao$MSE, na.rm = TRUE), digits = 2)
min <- round(min(datao$MSE, na.rm = TRUE), digits = 2)
max <- round(max(datao$MSE, na.rm = TRUE), digits = 2)
sd <- round(sd(datao$MSE, na.rm = TRUE), digits = 2)
mseo <- cbind(mean, min, max, sd)

mean <- round(mean(datao$KTest, na.rm = TRUE), digits = 2)
min <- round(min(datao$KTest, na.rm = TRUE), digits = 2)
max <- round(max(datao$KTest, na.rm = TRUE), digits = 2)
sd <- round(sd(datao$KTest, na.rm = TRUE), digits = 2)
cornealo <- cbind(mean, min, max, sd)

torico <- as.data.frame(table(datao$Toric))
torico$percent <- round(torico$Freq/sum(torico$Freq), digits = 4)

u <- data %>%
  filter(SurgeryLens=="Bilateral PureSee") %>%
  select(PatientID, Age, Gender) %>%
  unique()

mean <- round(mean(u$Age), digits = 2)
min <- min(u$Age)
max <- max(u$Age)
sd <- round(sd(u$Age), digits = 2)
agep <- cbind(mean, min, max, sd)

genderp <- as.data.frame(table(u$Gender))
genderp$percent <- round(genderp$Freq/sum(genderp$Freq), digits = 4)

datap <- data %>%
  filter(SurgeryLens=="Bilateral PureSee") %>%
  select(PatientID, EyeID, Sphere, Cyl, MSE, KTest, Toric) %>%
  unique()

mean <- round(mean(datap$Sphere, na.rm = TRUE), digits = 2)
min <- round(min(datap$Sphere, na.rm = TRUE), digits = 2)
max <- round(max(datap$Sphere, na.rm = TRUE), digits = 2)
sd <- round(sd(datap$Sphere, na.rm = TRUE), digits = 2)
spherep <- cbind(mean, min, max, sd)

mean <- round(mean(datap$Cyl, na.rm = TRUE), digits = 2)
min <- round(min(datap$Cyl, na.rm = TRUE), digits = 2)
max <- round(max(datap$Cyl, na.rm = TRUE), digits = 2)
sd <- round(sd(datap$Cyl, na.rm = TRUE), digits = 2)
cylp <- cbind(mean, min, max, sd)

mean <- round(mean(datap$MSE, na.rm = TRUE), digits = 2)
min <- round(min(datap$MSE, na.rm = TRUE), digits = 2)
max <- round(max(datap$MSE, na.rm = TRUE), digits = 2)
sd <- round(sd(datap$MSE, na.rm = TRUE), digits = 2)
msep <- cbind(mean, min, max, sd)

mean <- round(mean(datap$KTest, na.rm = TRUE), digits = 2)
min <- round(min(datap$KTest, na.rm = TRUE), digits = 2)
max <- round(max(datap$KTest, na.rm = TRUE), digits = 2)
sd <- round(sd(datap$KTest, na.rm = TRUE), digits = 2)
cornealp <- cbind(mean, min, max, sd)

toricp <- as.data.frame(table(datap$Toric))
toricp$percent <- round(toricp$Freq/sum(toricp$Freq), digits = 4)

u <- data %>%
  filter(SurgeryLens=="Mix and Match") %>%
  select(PatientID, Age, Gender) %>%
  unique()

mean <- round(mean(u$Age), digits = 2)
min <- min(u$Age)
max <- max(u$Age)
sd <- round(sd(u$Age), digits = 2)
agem <- cbind(mean, min, max, sd)

datam <- data %>%
  filter(SurgeryLens=="Mix and Match") %>%
  select(PatientID, EyeID, Sphere, Cyl, MSE, KTest, Toric) %>%
  unique()

genderm <- as.data.frame(table(u$Gender))
genderm$percent <- round(genderm$Freq/sum(genderm$Freq), digits = 4)

mean <- round(mean(datam$Sphere, na.rm = TRUE), digits = 2)
min <- round(min(datam$Sphere, na.rm = TRUE), digits = 2)
max <- round(max(datam$Sphere, na.rm = TRUE), digits = 2)
sd <- round(sd(datam$Sphere, na.rm = TRUE), digits = 2)
spherem <- cbind(mean, min, max, sd)

mean <- round(mean(datam$Cyl, na.rm = TRUE), digits = 2)
min <- round(min(datam$Cyl, na.rm = TRUE), digits = 2)
max <- round(max(datam$Cyl, na.rm = TRUE), digits = 2)
sd <- round(sd(datam$Cyl, na.rm = TRUE), digits = 2)
cylm <- cbind(mean, min, max, sd)

mean <- round(mean(datam$MSE, na.rm = TRUE), digits = 2)
min <- round(min(datam$MSE, na.rm = TRUE), digits = 2)
max <- round(max(datam$MSE, na.rm = TRUE), digits = 2)
sd <- round(sd(datam$MSE, na.rm = TRUE), digits = 2)
msem <- cbind(mean, min, max, sd)

mean <- round(mean(datam$KTest, na.rm = TRUE), digits = 2)
min <- round(min(datam$KTest, na.rm = TRUE), digits = 2)
max <- round(max(datam$KTest, na.rm = TRUE), digits = 2)
sd <- round(sd(datam$KTest, na.rm = TRUE), digits = 2)
cornealm <- cbind(mean, min, max, sd)

toricm <- as.data.frame(table(datam$Toric))
toricm$percent <- round(toricm$Freq/sum(toricm$Freq), digits = 4)

table <- list("Surgery n Location" = surgloc,
              "PatientEye O" = totalo,
              "Age O" = ageo,
              "Gender O" = gendero,
              "Sphere O" = sphereo, 
              "Cyl O" = cylo,
              "MSE O" = mseo,
              "Corneal O" = cornealo,
              "Toric O" = torico,
              "PatientEye P" = totalp,
              "Age P" = agep,
              "Gender P" = genderp,
              "Sphere P" = spherep, 
              "Cyl P" = cylp,
              "MSE P" = msep,
              "Corneal P" = cornealp,
              "Toric P" = toricp,
              "PatientEye M" = totalm,
              "Age M" = agem,
              "Gender M" = genderm,
              "Sphere M" = spherem, 
              "Cyl M" = cylm,
              "MSE M" = msem,
              "Corneal M" = cornealm,
              "Toric M" = toricm)

write.xlsx(table, paste(path, "/tables/", "Demo.xlsx", sep = ""))

#### IMPLANTED LENS ####

lens <- as.data.frame(table(data$Model, data$SurgeryLens)) %>%
  pivot_wider(names_from="Var2", values_from="Freq")
lens$bilatpercent <- round(lens$`Bilateral Odyssey`/sum(lens$`Bilateral Odyssey`), digits = 4)
lens$pspercent <- round(lens$`Bilateral PureSee`/sum(lens$`Bilateral PureSee`), digits = 4)
lens$mixpercent <- round(lens$`Mix and Match`/sum(lens$`Mix and Match`), digits = 4)
lens <- lens %>% arrange(-`Bilateral PureSee`)

write.xlsx(lens, paste(path, "tables/", "Lens Table.xlsx", sep = ""))

#### Mix and Match Combinations ####
mix <- data %>%
  filter(SurgeryLens=="Mix and Match") %>%
  select(PatientID, EyeID, SurgeryDateTime, Model)

#### Pre-Op Sphere Distribution ####

sphere <- data %>%
  select(PatientID, EyeID, Sphere, SurgeryLens)
sphere$Sphere <- round(sphere$Sphere,1)
sphere <- sphere[complete.cases(sphere),]

sphere$plot<-c('')
for (i in 1:nrow(sphere)) {
  if (sphere$Sphere[i] > 3){
    sphere$plot[i] <- ">3D"
  } else if (sphere$Sphere[i] < -8) {
    sphere$plot[i] <- "<-8D"
  } else {sphere$plot[i] <- paste(sphere$Sphere[i],'D',sep = "")
  }
}

sph <- as.data.frame(table(sphere$plot, sphere$SurgeryLens))

format_values <- function(x) {
  if (str_detect(x, ">")) {
    x <- str_remove(x, ">")
    x <- as.numeric(str_remove(x, "D"))
    formatted <- sprintf(">%.2fD", x)
  } else if (str_detect(x, "<")) {
    x <- str_remove(x, "<")
    x <- as.numeric(str_remove(x, "D"))
    formatted <- sprintf("<%.2fD", x)
  } else {
    x <- as.numeric(str_remove(x, "D"))
    formatted <- sprintf("%.2fD", x)
  }
  return(formatted)
}

sph <- sph %>% 
  mutate(Var1 = sapply(Var1, format_values)) %>%
  group_by(Var2) %>%
  mutate(percent = Freq/sum(Freq)) %>%
  select(-Freq) %>%
  pivot_wider(names_from = "Var2", values_from = "percent")

#### Pre-Op MSE Distribution ####

mse <- data %>%
  select(PatientID, EyeID, MSE, SurgeryLens)
mse$MSE <- round(mse$MSE,1)
mse <- mse[complete.cases(mse),]

mse$plot<-c('')
for (i in 1:nrow(mse)) {
  if (mse$MSE[i] > 3){
    mse$plot[i] <- ">3D"
  } else if (mse$MSE[i] < -8) {
    mse$plot[i] <- "<-8D"
  } else {mse$plot[i] <- paste(mse$MSE[i],'D',sep = "")
  }
}

mse <- as.data.frame(table(mse$plot, mse$SurgeryLens))

format_values <- function(x) {
  if (str_detect(x, ">")) {
    x <- str_remove(x, ">")
    x <- as.numeric(str_remove(x, "D"))
    formatted <- sprintf(">%.2fD", x)
  } else if (str_detect(x, "<")) {
    x <- str_remove(x, "<")
    x <- as.numeric(str_remove(x, "D"))
    formatted <- sprintf("<%.2fD", x)
  } else {
    x <- as.numeric(str_remove(x, "D"))
    formatted <- sprintf("%.2fD", x)
  }
  return(formatted)
}

mse <- mse %>% 
  mutate(Var1 = sapply(Var1, format_values)) %>%
  group_by(Var2) %>%
  mutate(percent = Freq/sum(Freq)) %>%
  select(-Freq) %>%
  pivot_wider(names_from = "Var2", values_from = "percent")

#### Pre-Op Cylinder Distribution ####

cylinder <- data %>%
  select(PatientID, EyeID, Cyl, SurgeryLens)
cylinder$Cyl <- round(cylinder$Cyl, 0.5)
cylinder <- cylinder[complete.cases(cylinder),]

cylinder$plot <- c('')
for (i in 1:nrow(cylinder)) {
  if (cylinder$Cyl[i] < -3.5 ){
    cylinder$plot[i] <- ">-3.5D"
  } else {cylinder$plot[i] <- paste(cylinder$Cyl[i],'D',sep = "")
  }
}

cyl <- as.data.frame(table(cylinder$plot, cylinder$SurgeryLens))

format_values <- function(x) {
  if (str_detect(x, ">")) {
    x <- str_remove(x, ">")
    x <- as.numeric(str_remove(x, "D"))
    formatted <- sprintf(">%.2fD", x)
  } else if (str_detect(x, "<")) {
    x <- str_remove(x, "<")
    x <- as.numeric(str_remove(x, "D"))
    formatted <- sprintf("<%.2fD", x)
  } else {
    x <- as.numeric(str_remove(x, "D"))
    formatted <- sprintf("%.2fD", x)
  }
  return(formatted)
}

cyl <- cyl %>% 
  mutate(Var1 = sapply(Var1, format_values)) %>%
  group_by(Var2) %>%
  mutate(percent = Freq/sum(Freq)) %>%
  select(-Freq) %>%
  pivot_wider(names_from = "Var2", values_from = "percent")

#### Pre UCDVA vs Post UCDVA 1 Month Bilateral Odyssey ####

month1 <- data %>%
  filter(SurgeryLens=="Bilateral Odyssey") %>%
  select(PatientID, Binoc_UCDVA_SM, M1_Binoc_UCDVA_SM) %>%
  filter(complete.cases(.)) %>%
  unique()

month <- month1 %>% select(PatientID, Binoc_UCDVA_SM)
month <- as.data.frame(table(month$Binoc_UCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")
mu26 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120")
mu28 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120"| month$Var1 == "6/240"| month$Var1 == "CF")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120"| month$Var1 == "6/240"| month$Var1 == "CF"| month$Var1 == "HM"| month$Var1 == "POL")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu26$Group <- c('6/120 or better')
mu26 <- mu26 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu28$Group <- c('CF or better')
mu28 <- mu28 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('HM or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

pre <- rbind(mu5, mu6, mu7, mu12, mu23, mu26, mu28, mu30)
pre$Type <- "Pre-Op UCDVA"

month <- month1 %>% select(PatientID, M1_Binoc_UCDVA_SM)
month <- as.data.frame(table(month$M1_Binoc_UCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")
mu26 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120")
mu28 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120"| month$Var1 == "6/240"| month$Var1 == "CF")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120"| month$Var1 == "6/240"| month$Var1 == "CF"| month$Var1 == "HM"| month$Var1 == "POL")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu26$Group <- c('6/120 or better')
mu26 <- mu26 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu28$Group <- c('CF or better')
mu28 <- mu28 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('HM or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

post <- rbind(mu5, mu6, mu7, mu12, mu23, mu26, mu28, mu30)
post$Type <- "Post-Op UCDVA"

prepost_ucdva_o <- rbind(pre, post)
prepost_ucdva_o <- subset(prepost_ucdva_o, select = -c(Freq))
prepost_ucdva_o <- pivot_wider(prepost_ucdva_o, names_from = "Type", values_from = "percent")


#### Pre UCDVA vs Post UCDVA 1 Month Bilateral PureSee ####
month1 <- data %>%
  filter(SurgeryLens=="Bilateral PureSee") %>%
  select(PatientID, Binoc_UCDVA_SM, M1_Binoc_UCDVA_SM) %>%
  filter(complete.cases(.)) %>%
  unique()

month <- month1 %>% select(PatientID, Binoc_UCDVA_SM)
month <- as.data.frame(table(month$Binoc_UCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")
mu26 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120")
mu28 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120"| month$Var1 == "6/240"| month$Var1 == "CF")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120"| month$Var1 == "6/240"| month$Var1 == "CF"| month$Var1 == "HM"| month$Var1 == "POL")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu26$Group <- c('6/120 or better')
mu26 <- mu26 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu28$Group <- c('CF or better')
mu28 <- mu28 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('HM or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

pre <- rbind(mu5, mu6, mu7, mu12, mu23, mu26, mu28, mu30)
pre$Type <- "Pre-Op UCDVA"

month <- month1 %>% select(PatientID, M1_Binoc_UCDVA_SM)
month <- as.data.frame(table(month$M1_Binoc_UCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")
mu26 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120")
mu28 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120"| month$Var1 == "6/240"| month$Var1 == "CF")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120"| month$Var1 == "6/240"| month$Var1 == "CF"| month$Var1 == "HM"| month$Var1 == "POL")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu26$Group <- c('6/120 or better')
mu26 <- mu26 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu28$Group <- c('CF or better')
mu28 <- mu28 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('HM or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

post <- rbind(mu5, mu6, mu7, mu12, mu23, mu26, mu28, mu30)
post$Type <- "Post-Op UCDVA"

prepost_ucdva_p <- rbind(pre, post)
prepost_ucdva_p <- subset(prepost_ucdva_p, select = -c(Freq))
prepost_ucdva_p <- pivot_wider(prepost_ucdva_p, names_from = "Type", values_from = "percent")

#### Pre UCDVA vs Post UCDVA 1 Month Mix and Match ####

month1 <- data %>%
  filter(SurgeryLens=="Mix and Match") %>%
  select(PatientID, Binoc_UCDVA_SM, M1_Binoc_UCDVA_SM) %>%
  filter(complete.cases(.)) %>%
  unique()

month <- month1 %>% select(PatientID, Binoc_UCDVA_SM)
month <- as.data.frame(table(month$Binoc_UCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")
mu26 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120")
mu28 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120"| month$Var1 == "6/240"| month$Var1 == "CF")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120"| month$Var1 == "6/240"| month$Var1 == "CF"| month$Var1 == "HM"| month$Var1 == "POL")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu26$Group <- c('6/120 or better')
mu26 <- mu26 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu28$Group <- c('CF or better')
mu28 <- mu28 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('HM or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

pre <- rbind(mu5, mu6, mu7, mu12, mu23, mu26, mu28, mu30)
pre$Type <- "Pre-Op UCDVA"

month <- month1 %>% select(PatientID, M1_Binoc_UCDVA_SM)
month <- as.data.frame(table(month$M1_Binoc_UCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")
mu26 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120")
mu28 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120"| month$Var1 == "6/240"| month$Var1 == "CF")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60"| month$Var1 == "6/75"| month$Var1 == "6/95"
               | month$Var1 == "6/120"| month$Var1 == "6/240"| month$Var1 == "CF"| month$Var1 == "HM"| month$Var1 == "POL")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu26$Group <- c('6/120 or better')
mu26 <- mu26 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu28$Group <- c('CF or better')
mu28 <- mu28 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('HM or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

post <- rbind(mu5, mu6, mu7, mu12, mu23, mu26, mu28, mu30)
post$Type <- "Post-Op UCDVA"

prepost_ucdva_m <- rbind(pre, post)
prepost_ucdva_m <- subset(prepost_ucdva_m, select = -c(Freq))
prepost_ucdva_m <- pivot_wider(prepost_ucdva_m, names_from = "Type", values_from = "percent")

#### Pre BCDVA vs Post UCDVA 1 Month Bilateral Odyssey ####
month1 <- data %>%
  filter(SurgeryLens=="Bilateral Odyssey") %>%
  select(PatientID, Binoc_BCDVA_SM, M1_Binoc_UCDVA_SM) %>%
  filter(complete.cases(.)) %>%
  unique()

month <- month1 %>% select(PatientID, Binoc_BCDVA_SM)
month <- as.data.frame(table(month$Binoc_BCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

pre <- rbind(mu5, mu6, mu7, mu12)
pre$Type <- "Pre-Op BCDVA"

month <- month1 %>% select(PatientID, M1_Binoc_UCDVA_SM)
month <- as.data.frame(table(month$M1_Binoc_UCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

post <- rbind(mu5, mu6, mu7, mu12)
post$Type <- "Post-Op UCDVA"

prebpostu_cdva_o <- rbind(pre, post)
prebpostu_cdva_o <- subset(prebpostu_cdva_o, select = -c(Freq))
prebpostu_cdva_o <- pivot_wider(prebpostu_cdva_o, names_from = "Type", values_from = "percent")

#### Pre BCDVA vs Post UCDVA 1 Month Bilateral PureSee ####
month1 <- data %>%
  filter(SurgeryLens=="Bilateral PureSee") %>%
  select(PatientID, Binoc_BCDVA_SM, M1_Binoc_UCDVA_SM) %>%
  filter(complete.cases(.)) %>%
  unique()

month <- month1 %>% select(PatientID, Binoc_BCDVA_SM)
month <- as.data.frame(table(month$Binoc_BCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

pre <- rbind(mu5, mu6, mu7, mu12)
pre$Type <- "Pre-Op BCDVA"

month <- month1 %>% select(PatientID, M1_Binoc_UCDVA_SM)
month <- as.data.frame(table(month$M1_Binoc_UCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

post <- rbind(mu5, mu6, mu7, mu12)
post$Type <- "Post-Op UCDVA"

prebpostu_cdva_p <- rbind(pre, post)
prebpostu_cdva_p <- subset(prebpostu_cdva_p, select = -c(Freq))
prebpostu_cdva_p <- pivot_wider(prebpostu_cdva_p, names_from = "Type", values_from = "percent")

#### Pre BCDVA vs Post UCDVA 1 Month Mix and Match ####
month1 <- data %>%
  filter(SurgeryLens=="Mix and Match") %>%
  select(PatientID, Binoc_BCDVA_SM, M1_Binoc_UCDVA_SM) %>%
  filter(complete.cases(.)) %>%
  unique()

month <- month1 %>% select(PatientID, Binoc_BCDVA_SM)
month <- as.data.frame(table(month$Binoc_BCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

pre <- rbind(mu5, mu6, mu7, mu12)
pre$Type <- "Pre-Op BCDVA"

month <- month1 %>% select(PatientID, M1_Binoc_UCDVA_SM)
month <- as.data.frame(table(month$M1_Binoc_UCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

post <- rbind(mu5, mu6, mu7, mu12)
post$Type <- "Post-Op UCDVA"

prebpostu_cdva_m <- rbind(pre, post)
prebpostu_cdva_m <- subset(prebpostu_cdva_m, select = -c(Freq))
prebpostu_cdva_m <- pivot_wider(prebpostu_cdva_m, names_from = "Type", values_from = "percent")

#### Pre BCDVA vs Post BCDVA 1 Month Bilateral Odyssey ####

y <- data %>%
  filter(SurgeryLens=="Bilateral Odyssey") %>%
  select(PatientID, EyeID, Binoc_BCDVA_SM, Binoc_BCDVA_D, M1_Binoc_BCDVA_SM, M1_Binoc_BCDVA_D, M1_Binoc_UCDVA_D) %>%
  filter(complete.cases(.)) %>%
  unique()

premean <- as.data.frame(mean(y$Binoc_BCDVA_D))
presd <- as.data.frame(sd(y$Binoc_BCDVA_D))
prebcdva <- do.call("cbind", c(premean,presd))
colnames(prebcdva) <- c("Mean", "SD")

postmean <- as.data.frame(mean(y$M1_Binoc_BCDVA_D))
postsd <- as.data.frame(sd(y$M1_Binoc_BCDVA_D))
postbcdva <- do.call("cbind", c(postmean,postsd))
colnames(postbcdva) <- c("Mean", "SD")

y$safety <- y$M1_Binoc_BCDVA_D/y$Binoc_BCDVA_D
smean <- as.data.frame(mean(y$safety))
ssd <- as.data.frame(sd(y$safety))
bcdva_safety <- do.call("cbind", c(smean,ssd))
colnames(bcdva_safety) <- c("Mean", "SD")

y$efficacy <- y$M1_Binoc_UCDVA_D/y$Binoc_BCDVA_D
emean <- as.data.frame(mean(y$efficacy))
esd <- as.data.frame(sd(y$efficacy))
bcdva_efficacy <- do.call("cbind", c(emean,esd))
colnames(bcdva_efficacy) <- c("Mean", "SD")

bilatlist <- list("Pre BCDVA O" = prebcdva,
                  "Post BCDVA O" = postbcdva,
                  "Safety O" = bcdva_safety,
                  "Efficacy O" = bcdva_efficacy)

month <- y
month <- as.data.frame(table(month$Binoc_BCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu20 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu20$Group <- c('6/20 or better')
mu20 <- mu20 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

pre <-  rbind(mu5, mu6, mu7, mu12, mu20, mu30, mu23)
pre$Tag <- "Pre-Op BCDVA"
pre <- subset(pre, select = -c(Freq))

month <- y
month <- as.data.frame(table(month$M1_Binoc_BCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu20 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu20$Group <- c('6/20 or better')
mu20 <- mu20 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

post <- rbind(mu5, mu6, mu7, mu12, mu20, mu30, mu23)
post$Tag <- "Post-Op BCDVA"
post <- subset(post, select = -c(Freq))

prepost_bcdva_o <- union(pre, post)
prepost_bcdva_o <- prepost_bcdva_o %>%
  pivot_wider(names_from = "Tag", values_from = "percent")


#### Pre BCDVA vs Post BCDVA 1 Month Bilateral PureSee ####

y <- data %>%
  filter(SurgeryLens=="Bilateral PureSee") %>%
  select(PatientID, EyeID, Binoc_BCDVA_SM, Binoc_BCDVA_D, M1_Binoc_BCDVA_SM, M1_Binoc_BCDVA_D, M1_Binoc_UCDVA_D) %>%
  filter(complete.cases(.)) %>%
  unique()

premean <- as.data.frame(mean(y$Binoc_BCDVA_D))
presd <- as.data.frame(sd(y$Binoc_BCDVA_D))
prebcdva <- do.call("cbind", c(premean,presd))
colnames(prebcdva) <- c("Mean", "SD")

postmean <- as.data.frame(mean(y$M1_Binoc_BCDVA_D))
postsd <- as.data.frame(sd(y$M1_Binoc_BCDVA_D))
postbcdva <- do.call("cbind", c(postmean,postsd))
colnames(postbcdva) <- c("Mean", "SD")

y$safety <- y$M1_Binoc_BCDVA_D/y$Binoc_BCDVA_D
smean <- as.data.frame(mean(y$safety))
ssd <- as.data.frame(sd(y$safety))
bcdva_safety <- do.call("cbind", c(smean,ssd))
colnames(bcdva_safety) <- c("Mean", "SD")

y$efficacy <- y$M1_Binoc_UCDVA_D/y$Binoc_BCDVA_D
emean <- as.data.frame(mean(y$efficacy))
esd <- as.data.frame(sd(y$efficacy))
bcdva_efficacy <- do.call("cbind", c(emean,esd))
colnames(bcdva_efficacy) <- c("Mean", "SD")

bilatplist <- list("Pre BCDVA P" = prebcdva,
                  "Post BCDVA P" = postbcdva,
                  "Safety P" = bcdva_safety,
                  "Efficacy P" = bcdva_efficacy)

month <- y
month <- as.data.frame(table(month$Binoc_BCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu20 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu20$Group <- c('6/20 or better')
mu20 <- mu20 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

pre <-  rbind(mu5, mu6, mu7, mu12, mu20, mu30, mu23)
pre$Tag <- "Pre-Op BCDVA"
pre <- subset(pre, select = -c(Freq))

month <- y
month <- as.data.frame(table(month$M1_Binoc_BCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu20 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu20$Group <- c('6/20 or better')
mu20 <- mu20 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

post <- rbind(mu5, mu6, mu7, mu12, mu20, mu30, mu23)
post$Tag <- "Post-Op BCDVA"
post <- subset(post, select = -c(Freq))

prepost_bcdva_p <- union(pre, post)
prepost_bcdva_p <- prepost_bcdva_p %>%
  pivot_wider(names_from = "Tag", values_from = "percent")

#### Pre BCDVA vs Post BCDVA 1 Month Mix and Match ####

y <- data %>%
  filter(SurgeryLens=="Mix and Match") %>%
  select(PatientID, EyeID, Binoc_BCDVA_SM, Binoc_BCDVA_D, M1_Binoc_BCDVA_SM, M1_Binoc_BCDVA_D, M1_Binoc_UCDVA_D) %>%
  filter(complete.cases(.)) %>%
  unique()

premean <- as.data.frame(mean(y$Binoc_BCDVA_D))
presd <- as.data.frame(sd(y$Binoc_BCDVA_D))
prebcdva <- do.call("cbind", c(premean,presd))
colnames(prebcdva) <- c("Mean", "SD")

postmean <- as.data.frame(mean(y$M1_Binoc_BCDVA_D))
postsd <- as.data.frame(sd(y$M1_Binoc_BCDVA_D))
postbcdva <- do.call("cbind", c(postmean,postsd))
colnames(postbcdva) <- c("Mean", "SD")

y$safety <- y$M1_Binoc_BCDVA_D/y$Binoc_BCDVA_D
smean <- as.data.frame(mean(y$safety))
ssd <- as.data.frame(sd(y$safety))
bcdva_safety <- do.call("cbind", c(smean,ssd))
colnames(bcdva_safety) <- c("Mean", "SD")

y$efficacy <- y$M1_Binoc_UCDVA_D/y$Binoc_BCDVA_D
emean <- as.data.frame(mean(y$efficacy))
esd <- as.data.frame(sd(y$efficacy))
bcdva_efficacy <- do.call("cbind", c(emean,esd))
colnames(bcdva_efficacy) <- c("Mean", "SD")

mixlist <- list("Pre BCDVA M" = prebcdva,
                "Post BCDVA M" = postbcdva,
                "Safety M" = bcdva_safety,
                "Efficacy M" = bcdva_efficacy)
list <- c(bilatlist, bilatplist, mixlist)
write.xlsx(list, paste(path, "/tables/", "safetynefficacy.xlsx", sep = ""))

month <- y
month <- as.data.frame(table(month$Binoc_BCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu20 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu20$Group <- c('6/20 or better')
mu20 <- mu20 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

pre <-  rbind(mu5, mu6, mu7, mu12, mu20, mu30, mu23)
pre$Tag <- "Pre-Op BCDVA"
pre <- subset(pre, select = -c(Freq))

month <- y
month <- as.data.frame(table(month$M1_Binoc_BCDVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu20 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")

mu5$Group <- c('6/5 or better')
mu5 <- mu5 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu20$Group <- c('6/20 or better')
mu20 <- mu20 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

post <- rbind(mu5, mu6, mu7, mu12, mu20, mu30, mu23)
post$Tag <- "Post-Op BCDVA"
post <- subset(post, select = -c(Freq))

prepost_bcdva_m <- union(pre, post)
prepost_bcdva_m <- prepost_bcdva_m %>%
  pivot_wider(names_from = "Tag", values_from = "percent")

#### 1 Month Change in UCDVA ####
y <- data %>%
  select(PatientID, EyeID, UCDVA_LogMAR, M1_UCDVA_LM, SurgeryLens)
y <- y[complete.cases(y),]
y$UCDVA_LogMAR <- as.numeric(y$UCDVA_LogMAR)
y$M1_UCDVA_LM <- as.numeric(y$M1_UCDVA_LM)
y$UCDVA_LogMAR[is.na(y$UCDVA_LogMAR)] <- 1.8
y$M1_UCDVA_LM[is.na(y$M1_UCDVA_LM)] <- 1.8

y$m1_ucdlc <- 10*(y$UCDVA_LogMAR - y$M1_UCDVA_LM)
y$m1_ucdlc <- round(y$m1_ucdlc)

y$LogMarChange <- y$UCDVA_LogMAR - y$M1_UCDVA_LM
ymeano<-round(mean(y$LogMarChange[y$SurgeryLens=="Bilateral Odyssey"]), digits = 2)
ymeanp<-round(mean(y$LogMarChange[y$SurgeryLens=="Bilateral PureSee"]), digits = 2)
ymeanm<-round(mean(y$LogMarChange[y$SurgeryLens=="Mix and Match"]), digits = 2)
meanchange <- data.frame(cbind(c("Bilateral Odyssey"),c(ymeano))) %>% rename("UCDVA"="X2")
meanchange <- rbind(meanchange, data.frame(cbind(c("Bilateral PureSee"),c(ymeanp))) %>% rename("UCDVA"="X2"))
meanchange <- rbind(meanchange, data.frame(cbind(c("Mix and Match"),c(ymeanm))) %>% rename("UCDVA"="X2"))

y <- y %>% mutate(gain = case_when(m1_ucdlc < -2 ~ "Loss > -2",
                                     m1_ucdlc == -2 ~ "Loss -2",
                                     m1_ucdlc == -1 ~ "Loss -1",
                                     m1_ucdlc == 0 ~ "No Change",
                                     m1_ucdlc == 1 ~ "Gain 1",
                                     m1_ucdlc == 2 ~ "Gain 2",
                                     m1_ucdlc == 3 ~ "Gain 3",
                                     m1_ucdlc == 4 ~ "Gain 4", 
                                     m1_ucdlc == 5 ~ "Gain 5",
                                     m1_ucdlc == 6 ~ "Gain 6",
                                     m1_ucdlc == 7 ~ "Gain 7",
                                     m1_ucdlc == 8 ~ "Gain 8",
                                     m1_ucdlc == 9 ~ "Gain 9",
                                     m1_ucdlc == 10 ~ "Gain 10",
                                     m1_ucdlc == 11 ~ "Gain 11",
                                     m1_ucdlc == 12 ~ "Gain 12",
                                     m1_ucdlc == 13 ~ "Gain 13",
                                     m1_ucdlc == 14 ~ "Gain 14",
                                     m1_ucdlc > 14 ~ "Gain > 14"))

y <- as.data.frame(table(y$gain, y$SurgeryLens)) %>%
  pivot_wider(names_from="Var2", values_from="Freq")
y$BilatOPercent <- round(y$`Bilateral Odyssey`/sum(y$`Bilateral Odyssey`), digits = 4)
y$BilatPPercent <- round(y$`Bilateral PureSee`/sum(y$`Bilateral PureSee`), digits = 4)
y$MixPercent <- round(y$`Mix and Match`/sum(y$`Mix and Match`), digits = 4)

ucdva_lc <- y %>% select(-`Bilateral Odyssey`,-`Bilateral PureSee`, -`Mix and Match`)

#### 1 Month Binocular UCIVA ####

month1 <- data %>%
  filter(SurgeryLens=="Bilateral Odyssey") %>%
  select(PatientID, M1_Binoc_UCIVA_SM) %>%
  filter(complete.cases(.)) %>%
  unique()

month <- as.data.frame(table(month1$M1_Binoc_UCIVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
# mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              # | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30")

# mu5$Group <- c('6/5 or better')
# mu5 <- mu5 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

# mu6$Group <- c('6/6 or better')
# mu6 <- mu6 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

od <- rbind(mu7, mu12, mu30)
od$Type <- "Odyssey"

month1 <- data %>%
  filter(SurgeryLens=="Bilateral PureSee") %>%
  select(PatientID, M1_Binoc_UCIVA_SM)

month1 <- month1[complete.cases(month1),]
month <- as.data.frame(table(month1$M1_Binoc_UCIVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
# mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
# | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30")

# mu5$Group <- c('6/5 or better')
# mu5 <- mu5 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

# mu6$Group <- c('6/6 or better')
# mu6 <- mu6 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

ps <- rbind(mu7, mu12, mu30)
ps$Type <- "PureSee"

month1 <- data %>%
  filter(SurgeryLens=="Mix and Match") %>%
  select(PatientID, M1_Binoc_UCIVA_SM)

month1 <- month1[complete.cases(month1),]
month <- as.data.frame(table(month1$M1_Binoc_UCIVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
# mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              # | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30")

# mu5$Group <- c('6/5 or better')
# mu5 <- mu5 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

# mu6$Group <- c('6/6 or better')
# mu6 <- mu6 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mix <- rbind(mu7, mu12, mu30)
mix$Type <- "Mix and Match"

binoc_uciva <- union(od, ps)
binoc_uciva <- union(binoc_uciva, mix)
binoc_uciva <- subset(binoc_uciva, select = -c(Freq))
binoc_uciva <- pivot_wider(binoc_uciva, names_from = "Type", values_from = "percent")

#### 1 Month Change in UCIVA ####

y <- data %>%
  select(PatientID, EyeID, UCIVA_LogMAR, M1_UCIVA_LM, SurgeryLens)
y <- y[complete.cases(y),]
y$UCIVA_LogMAR <- as.numeric(y$UCIVA_LogMAR)
y$M1_UCIVA_LM <- as.numeric(y$M1_UCIVA_LM)
y$UCIVA_LogMAR[is.na(y$UCIVA_LogMAR)] <- 1.8
y$M1_UCIVA_LM[is.na(y$M1_UCIVA_LM)] <- 1.8

y$m1_ucdlc <- 10*(y$UCIVA_LogMAR - y$M1_UCIVA_LM)
y$m1_ucdlc <- round(y$m1_ucdlc)

y$LogMarChange <- y$UCIVA_LogMAR - y$M1_UCIVA_LM
ymeano <- cbind("Bilateral Odyssey", round(mean(y$LogMarChange[y$SurgeryLens=="Bilateral Odyssey"]), digits = 2))
ymeanp <- cbind("Bilateral PureSee", round(mean(y$LogMarChange[y$SurgeryLens=="Bilateral PureSee"]), digits = 2))
ymeanm <- cbind("Mix and Match", round(mean(y$LogMarChange[y$SurgeryLens=="Mix and Match"]), digits = 2))
ymean <- rbind(ymeano, ymeanp, ymeanm)
meanchange <- merge(meanchange, data.frame(ymean), by="X1") %>% rename("UCIVA"="X2")

y <- y %>% mutate(gain = case_when(m1_ucdlc < -2 ~ "Loss > -2",
                                     m1_ucdlc == -2 ~ "Loss -2",
                                     m1_ucdlc == -1 ~ "Loss -1",
                                     m1_ucdlc == 0 ~ "No Change",
                                     m1_ucdlc == 1 ~ "Gain 1",
                                     m1_ucdlc == 2 ~ "Gain 2",
                                     m1_ucdlc == 3 ~ "Gain 3",
                                     m1_ucdlc == 4 ~ "Gain 4", 
                                     m1_ucdlc == 5 ~ "Gain 5",
                                     m1_ucdlc == 6 ~ "Gain 6",
                                     m1_ucdlc == 7 ~ "Gain 7",
                                     m1_ucdlc == 8 ~ "Gain 8",
                                     m1_ucdlc == 9 ~ "Gain 9",
                                     m1_ucdlc == 10 ~ "Gain 10",
                                     m1_ucdlc == 11 ~ "Gain 11",
                                     m1_ucdlc == 12 ~ "Gain 12",
                                     m1_ucdlc == 13 ~ "Gain 13",
                                     m1_ucdlc == 14 ~ "Gain 14",
                                     m1_ucdlc > 14 ~ "Gain > 14"))

y <- as.data.frame(table(y$gain, y$SurgeryLens))%>%
  pivot_wider(names_from="Var2", values_from="Freq")
y$BilatOPercent <- round(y$`Bilateral Odyssey`/sum(y$`Bilateral Odyssey`), digits = 4)
y$BilatPPercent <- round(y$`Bilateral PureSee`/sum(y$`Bilateral PureSee`), digits = 4)
y$MixPercent <- round(y$`Mix and Match`/sum(y$`Mix and Match`), digits = 4)

uciva_lc <- y %>% select(-`Bilateral Odyssey`, -`Bilateral PureSee`, - `Mix and Match`)

#### 1 Month Binocular UCNVA ####
month1 <- data %>% select(PatientID, M1_Binoc_UCNVA_SM, SurgeryLens) %>%
  filter(complete.cases(.)) %>%
  unique()

month <- as.data.frame(table(month1$M1_Binoc_UCNVA_SM[which(month1$SurgeryLens=="Bilateral Odyssey")]))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu15 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15")
mu19 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/24"| month$Var1 == "6/30")

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu15$Group <- c('6/15 or better')
mu15 <- mu15 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu19$Group <- c('6/19 or better')
mu19 <- mu19 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

od <- rbind(mu6, mu7, mu12, mu15, mu19, mu30)
od$Type <- "Bilateral Odyssey"

month <- as.data.frame(table(month1$M1_Binoc_UCNVA_SM[which(month1$SurgeryLens=="Bilateral PureSee")]))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu15 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15")
mu19 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/24"| month$Var1 == "6/30")

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu15$Group <- c('6/15 or better')
mu15 <- mu15 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu19$Group <- c('6/19 or better')
mu19 <- mu19 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

ps <- rbind(mu6, mu7, mu12, mu15, mu19, mu30)
ps$Type <- "Bilateral PureSee"

month <- as.data.frame(table(month1$M1_Binoc_UCNVA_SM[which(month1$SurgeryLens=="Mix and Match")]))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)
month$Group <- c('')

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
# mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
#               | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu15 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15")
mu19 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/24"| month$Var1 == "6/30")

# mu6$Group <- c('6/6 or better')
# mu6 <- mu6 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu15$Group <- c('6/15 or better')
mu15 <- mu15 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu19$Group <- c('6/19 or better')
mu19 <- mu19 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mix <- rbind(mu7, mu12, mu15, mu19, mu30)
mix$Type <- "Mix and Match"

binoc_ucnva <- rbind(od,ps)
binoc_ucnva <- rbind(binoc_ucnva,mix)
binoc_ucnva <- subset(binoc_ucnva, select = -c(Freq))
binoc_ucnva <- pivot_wider(binoc_ucnva, names_from = "Type", values_from = "percent")

#### 1 Month Monocular and Binocular UCNVA Bilateral Odyssey ####
month1 <- data %>%
  filter(SurgeryLens=="Bilateral Odyssey") %>%
  select(PatientID, EyeID, M1_UCNVA_SM, M1_Binoc_UCNVA_SM)
month1 <- month1[complete.cases(month1),]

month <- month1 %>% select(PatientID,EyeID, M1_UCNVA_SM)
month <- as.data.frame(table(month$M1_UCNVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu15 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15")
mu20 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/24"| month$Var1 == "6/30")

# mu5$Group <- c('6/5 or better')
# mu5 <- mu5 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu15$Group <- c('6/15 or better')
mu15 <- mu15 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu20$Group <- c('6/20 or better')
mu20 <- mu20 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

monoc <-  rbind(mu6, mu7, mu12, mu15, mu20, mu30)
monoc$Tag <- "Monocular"
monoc <- subset(monoc, select = -c(Freq))

month <- month1 %>% select(PatientID,M1_Binoc_UCNVA_SM) %>% unique()
month <- as.data.frame(table(month$M1_Binoc_UCNVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu15 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15")
mu20 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/24"| month$Var1 == "6/30")

# mu5$Group <- c('6/5 or better')
# mu5 <- mu5 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu15$Group <- c('6/15 or better')
mu15 <- mu15 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu20$Group <- c('6/20 or better')
mu20 <- mu20 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

binoc <-  rbind(mu6, mu7, mu12, mu15, mu20, mu30)
binoc$Tag <- "Binocular"
binoc <- subset(binoc, select = -c(Freq))

ucnva_o <- rbind(monoc,binoc)
ucnva_o <- ucnva_o %>%
  pivot_wider(names_from = "Tag", values_from = "percent")


#### 1 Month Monocular and Binocular UCNVA Bilateral PureSee ####
month1 <- data %>%
  filter(SurgeryLens=="Bilateral PureSee") %>%
  select(PatientID, EyeID, M1_UCNVA_SM, M1_Binoc_UCNVA_SM)
month1 <- month1[complete.cases(month1),]

month <- month1 %>% select(PatientID,EyeID, M1_UCNVA_SM)
month <- as.data.frame(table(month$M1_UCNVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu15 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15")
mu20 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/24"| month$Var1 == "6/30")

# mu5$Group <- c('6/5 or better')
# mu5 <- mu5 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu15$Group <- c('6/15 or better')
mu15 <- mu15 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu20$Group <- c('6/20 or better')
mu20 <- mu20 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

monoc <-  rbind(mu6, mu7, mu12, mu15, mu20, mu30)
monoc$Tag <- "Monocular"
monoc <- subset(monoc, select = -c(Freq))

month <- month1 %>% select(PatientID,M1_Binoc_UCNVA_SM) %>% unique()
month <- as.data.frame(table(month$M1_Binoc_UCNVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu15 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15")
mu20 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/24"| month$Var1 == "6/30")

# mu5$Group <- c('6/5 or better')
# mu5 <- mu5 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu15$Group <- c('6/15 or better')
mu15 <- mu15 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu20$Group <- c('6/20 or better')
mu20 <- mu20 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

binoc <-  rbind(mu6, mu7, mu12, mu15, mu20, mu30)
binoc$Tag <- "Binocular"
binoc <- subset(binoc, select = -c(Freq))

ucnva_p <- rbind(monoc,binoc)
ucnva_p <- ucnva_p %>%
  pivot_wider(names_from = "Tag", values_from = "percent")

#### 1 Month Monocular and Binocular UCNVA Mix and Match ####
month1 <- data %>%
  filter(SurgeryLens=="Mix and Match") %>%
  select(PatientID, EyeID, M1_UCNVA_SM, M1_Binoc_UCNVA_SM)
month1 <- month1[complete.cases(month1),]

month <- month1 %>% select(PatientID,EyeID, M1_UCNVA_SM)
month <- as.data.frame(table(month$M1_UCNVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
# mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
#               | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu15 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15")
mu20 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/24"| month$Var1 == "6/30")

# mu5$Group <- c('6/5 or better')
# mu5 <- mu5 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

# mu6$Group <- c('6/6 or better')
# mu6 <- mu6 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu15$Group <- c('6/15 or better')
mu15 <- mu15 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu20$Group <- c('6/20 or better')
mu20 <- mu20 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

monoc <-  rbind(mu7, mu12, mu15, mu20, mu30)
monoc$Tag <- "Monocular"
monoc <- subset(monoc, select = -c(Freq))

month <- month1 %>% select(PatientID,M1_Binoc_UCNVA_SM) %>% unique()
month <- as.data.frame(table(month$M1_Binoc_UCNVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
# mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
#               | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu15 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15")
mu20 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20")
mu30 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/24"| month$Var1 == "6/30")

# mu5$Group <- c('6/5 or better')
# mu5 <- mu5 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

# mu6$Group <- c('6/6 or better')
# mu6 <- mu6 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu15$Group <- c('6/15 or better')
mu15 <- mu15 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu20$Group <- c('6/20 or better')
mu20 <- mu20 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu30$Group <- c('6/30 or better')
mu30 <- mu30 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

binoc <-  rbind(mu7, mu12, mu15, mu20, mu30)
binoc$Tag <- "Binocular"
binoc <- subset(binoc, select = -c(Freq))

ucnva_m <- rbind(monoc,binoc)
ucnva_m <- ucnva_m %>%
  pivot_wider(names_from = "Tag", values_from = "percent")

#### 1 Month Monocular UCNVA Bilateral Odyssey ####
month1 <- data %>%
  filter(SurgeryLens=="Bilateral Odyssey") %>%
  select(PatientID, EyeID, UCNVA_SM, M1_UCNVA_SM, SurgeryLens)
month1 <- month1[complete.cases(month1),]

month <- month1 %>% select(PatientID,UCNVA_SM)
month <- as.data.frame(table(month$UCNVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
# mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
#               | month$Var1 == "6/6")
# mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
#               | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu15 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15")
mu18 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")

# mu5$Group <- c('6/5 or better')
# mu5 <- mu5 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

# mu6$Group <- c('6/6 or better')
# mu6 <- mu6 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

# mu7$Group <- c('6/7.5 or better')
# mu7 <- mu7 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu15$Group <- c('6/15 or better')
mu15 <- mu15 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu18$Group <- c('6/30 or better')
mu18 <- mu18 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mono <-  rbind(mu12, mu15, mu18, mu23)
mono$Tag <- "Pre Op"
mono <- subset(mono, select = -c(Freq))

month <- month1 %>% select(PatientID,EyeID, M1_UCNVA_SM)
month <- as.data.frame(table(month$M1_UCNVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu15 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15")
mu18 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")

# mu5$Group <- c('6/5 or better')
# mu5 <- mu5 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu15$Group <- c('6/15 or better')
mu15 <- mu15 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu18$Group <- c('6/30 or better')
mu18 <- mu18 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

bi <-  rbind(mu6, mu7, mu12, mu15, mu18, mu23)
bi$Tag <- "3 Month Post Op"
bi <- subset(bi, select = -c(Freq))

monoc_ucnva_o <- union(mono, bi)
monoc_ucnva_o <- monoc_ucnva_o %>%
  pivot_wider(names_from = "Tag", values_from = "percent")


#### 1 Month Monocular UCNVA Bilateral PureSee ####
month1 <- data %>%
  filter(SurgeryLens=="Bilateral PureSee") %>%
  select(PatientID, EyeID, UCNVA_SM, M1_UCNVA_SM, SurgeryLens)
month1 <- month1[complete.cases(month1),]

month <- month1 %>% select(PatientID,EyeID, UCNVA_SM)
month <- as.data.frame(table(month$UCNVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu15 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15")
mu18 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")

# mu5$Group <- c('6/5 or better')
# mu5 <- mu5 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu15$Group <- c('6/15 or better')
mu15 <- mu15 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu18$Group <- c('6/30 or better')
mu18 <- mu18 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mono <-  rbind(mu6, mu7, mu12, mu15, mu18, mu23)
mono$Tag <- "Pre Op"
mono <- subset(mono, select = -c(Freq))

month <- month1 %>% select(PatientID,EyeID, M1_UCNVA_SM)
month <- as.data.frame(table(month$M1_UCNVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu15 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15")
mu18 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")

# mu5$Group <- c('6/5 or better')
# mu5 <- mu5 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

mu6$Group <- c('6/6 or better')
mu6 <- mu6 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu15$Group <- c('6/15 or better')
mu15 <- mu15 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu18$Group <- c('6/30 or better')
mu18 <- mu18 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

bi <-  rbind(mu6, mu7, mu12, mu15, mu18, mu23)
bi$Tag <- "3 Month Post Op"
bi <- subset(bi, select = -c(Freq))

monoc_ucnva_p <- union(mono, bi)
monoc_ucnva_p <- monoc_ucnva_p %>%
  pivot_wider(names_from = "Tag", values_from = "percent")

#### 1 Month Monocular UCNVA Mix and Match ####
month1 <- data %>%
  filter(SurgeryLens=="Mix and Match") %>%
  select(PatientID, EyeID, UCNVA_SM, M1_UCNVA_SM, SurgeryLens)
month1 <- month1[complete.cases(month1),]

month <- month1 %>% select(PatientID,UCNVA_SM)
month <- as.data.frame(table(month$UCNVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
# mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
#               | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu15 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15")
mu18 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")

# mu5$Group <- c('6/5 or better')
# mu5 <- mu5 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)
# 
# mu6$Group <- c('6/6 or better')
# mu6 <- mu6 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu15$Group <- c('6/15 or better')
mu15 <- mu15 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu18$Group <- c('6/30 or better')
mu18 <- mu18 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mono <-  rbind(mu7, mu12, mu15, mu18, mu23)
mono$Tag <- "Pre Op"
mono <- subset(mono, select = -c(Freq))

month <- month1 %>% select(PatientID,EyeID, M1_UCNVA_SM)
month <- as.data.frame(table(month$M1_UCNVA_SM))
month$percent <- round(month$Freq/sum(month$Freq), digits = 4)

# mu5 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5")
# mu6 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
#               | month$Var1 == "6/6")
mu7 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
              | month$Var1 == "6/6" | month$Var1 == "6/7.5")
mu12 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12")
mu15 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12" | month$Var1 == "6/15")
mu18 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30")
mu23 <- subset(month, month$Var1 == "6/3" | month$Var1 == "6/3.8" | month$Var1 == "6/4" | month$Var1 == "6/4.8" | month$Var1 == "6/5"
               | month$Var1 == "6/6" | month$Var1 == "6/7.5" | month$Var1 == "6/9" | month$Var1 == "6/9.5" | month$Var1 == "6/10"
               | month$Var1 == "6/12"| month$Var1 == "6/15"| month$Var1 == "6/18"| month$Var1 == "6/19"| month$Var1 == "6/20"
               | month$Var1 == "6/21"| month$Var1 == "6/24"| month$Var1 == "6/30" | month$Var1 == "6/34.2"| month$Var1 == "6/36"
               | month$Var1 == "6/38"| month$Var1 == "6/48"| month$Var1 == "6/60")

# mu5$Group <- c('6/5 or better')
# mu5 <- mu5 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)
# 
# mu6$Group <- c('6/6 or better')
# mu6 <- mu6 %>%
#   group_by(Group) %>%
#   summarise_at(c('Freq', 'percent'), sum)

mu7$Group <- c('6/7.5 or better')
mu7 <- mu7 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu12$Group <- c('6/12 or better')
mu12 <- mu12 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu15$Group <- c('6/15 or better')
mu15 <- mu15 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu18$Group <- c('6/30 or better')
mu18 <- mu18 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

mu23$Group <- c('6/60 or better')
mu23 <- mu23 %>%
  group_by(Group) %>%
  summarise_at(c('Freq', 'percent'), sum)

bi <-  rbind(mu7, mu12, mu15, mu18, mu23)
bi$Tag <- "3 Month Post Op"
bi <- subset(bi, select = -c(Freq))

monoc_ucnva_m <- union(mono, bi)
monoc_ucnva_m <- monoc_ucnva_m %>%
  pivot_wider(names_from = "Tag", values_from = "percent")

#### 1 Month Change in UCNVA ####
y <- data %>%
  select(PatientID, EyeID, UCNVA_LogMAR, M1_UCNVA_LM, SurgeryLens)
y <- y[complete.cases(y),]
y$UCNVA_LogMAR <- as.numeric(y$UCNVA_LogMAR)
y$M1_UCNVA_LM <- as.numeric(y$M1_UCNVA_LM)
y$UCNVA_LogMAR[is.na(y$UCNVA_LogMAR)] <- 1.8
y$M1_UCNVA_LM[is.na(y$M1_UCNVA_LM)] <- 1.8

y$m1_ucdlc <- 10*(y$UCNVA_LogMAR - y$M1_UCNVA_LM)
y$m1_ucdlc <- round(y$m1_ucdlc)

y$LogMarChange <- y$UCNVA_LogMAR - y$M1_UCNVA_LM
ymeano <- cbind("Bilateral Odyssey", round(mean(y$LogMarChange[y$SurgeryLens=="Bilateral Odyssey"]), digits = 2))
ymeanp <- cbind("Bilateral PureSee", round(mean(y$LogMarChange[y$SurgeryLens=="Bilateral PureSee"]), digits = 2))
ymeanm <- cbind("Mix and Match", round(mean(y$LogMarChange[y$SurgeryLens=="Mix and Match"]), digits = 2))
ymean <- rbind(ymeano, ymeanp, ymeanm)
meanchange <- merge(meanchange, data.frame(ymean), by="X1") %>% rename("UCNVA"="X2")

y <- y %>% mutate(gain = case_when(m1_ucdlc < -2 ~ "Loss > -2",
                                     m1_ucdlc == -2 ~ "Loss -2",
                                     m1_ucdlc == -1 ~ "Loss -1",
                                     m1_ucdlc == 0 ~ "No Change",
                                     m1_ucdlc == 1 ~ "Gain 1",
                                     m1_ucdlc == 2 ~ "Gain 2",
                                     m1_ucdlc == 3 ~ "Gain 3",
                                     m1_ucdlc == 4 ~ "Gain 4", 
                                     m1_ucdlc == 5 ~ "Gain 5",
                                     m1_ucdlc == 6 ~ "Gain 6",
                                     m1_ucdlc == 7 ~ "Gain 7",
                                     m1_ucdlc == 8 ~ "Gain 8",
                                     m1_ucdlc == 9 ~ "Gain 9",
                                     m1_ucdlc == 10 ~ "Gain 10",
                                     m1_ucdlc == 11 ~ "Gain 11",
                                     m1_ucdlc == 12 ~ "Gain 12",
                                     m1_ucdlc == 13 ~ "Gain 13",
                                     m1_ucdlc == 14 ~ "Gain 14",
                                     m1_ucdlc > 14 ~ "Gain > 14"))

y <- as.data.frame(table(y$gain, y$SurgeryLens))%>%
  pivot_wider(names_from="Var2", values_from="Freq")
y$BilatOPercent <- round(y$`Bilateral Odyssey`/sum(y$`Bilateral Odyssey`), digits = 4)
y$BilatPPercent <- round(y$`Bilateral PureSee`/sum(y$`Bilateral PureSee`), digits = 4)
y$MixPercent <- round(y$`Mix and Match`/sum(y$`Mix and Match`), digits = 4)

ucnva_lc <- y %>% select(-`Bilateral Odyssey`, -`Bilateral PureSee`, - `Mix and Match`)

#### 1 Month Refractive Predictability (MSE) #### 
refpred <- data %>%
  select(MSE_M1, SurgeryLens) %>%
  filter(!is.na(MSE_M1))

refprednos <- refpred %>%
  group_by(SurgeryLens) %>%
  summarise(Count = n())

r <- as.data.frame(table(refpred$SurgeryLens))
r <- r %>%
  pivot_wider(names_from = "Var1", values_from = "Freq")

mseo5 <- refpred %>% filter(MSE_M1 >= -0.5 & MSE_M1 <= 0.5)
mse1 <- refpred %>% filter(MSE_M1 >= -1 & MSE_M1 <= 1)

mseo5 <- as.data.frame(table(mseo5$SurgeryLens))
colnames(mseo5) <- c("Group", "Within 0.50D")
mse1 <- as.data.frame(table(mse1$SurgeryLens))
colnames(mse1) <- c("Group", "Within 1.00D")

refpred <- merge(mseo5, mse1, by = "Group")
refpred <- pivot_longer(refpred, cols = c("Within 0.50D","Within 1.00D"),
                    names_to = "Within",
                    values_to = "Count")
refpred <- pivot_wider(refpred, names_from = "Group", values_from = "Count")

refpred$ODPercent <- round(refpred$`Bilateral Odyssey`/r$`Bilateral Odyssey`, digits = 4)
refpred$PSPercent <- round(refpred$`Bilateral PureSee`/r$`Bilateral PureSee`, digits = 4)
refpred$MixPercent <- round(refpred$`Mix and Match`/r$`Mix and Match`, digits = 4)

refpred <- refpred %>% select(-`Bilateral Odyssey`, -`Bilateral PureSee`, -`Mix and Match`)
colnames(refpred) <- c("Group", "Bilateral Odyssey", "Bilateral PureSee", "Mix and Match")

#### Att vs Ach Sphere Patients 1 Months Odyssey ####
sach_m1_o <- data %>%
  filter(SurgeryLens=="Bilateral Odyssey") %>%
  select(Sphere, Target, Sph_ach_M1, Sph_M1)
sach_m1_o <- sach_m1_o[complete.cases(sach_m1_o),]

sach_m1_o$att <- NA
sach_m1_o$att[which(sach_m1_o$Target==0)] <- sach_m1_o$Sphere[which(sach_m1_o$Target==0)]
sach_m1_o$att[which(sach_m1_o$Target!=0)] <- sach_m1_o$Sphere[which(sach_m1_o$Target!=0)] - sach_m1_o$Target[which(sach_m1_o$Target!=0)]

attvach3_o <- data.frame('Preop'= c(mean(sach_m1_o$Sphere,na.rm=TRUE),sd(sach_m1_o$Sphere)), 'M1' = c(mean(sach_m1_o$Sph_M1),sd(sach_m1_o$Sph_M1)),row.names=c('Mean', 'SD'))
sach_m1_o <- sach_m1_o %>% select(att, Sph_ach_M1)

#### Att vs Ach Sphere Patients 1 Months PureSee ####
sach_m1_p <- data %>%
  filter(SurgeryLens=="Bilateral PureSee") %>%
  select(Sphere, Target, Sph_ach_M1, Sph_M1)
sach_m1_p <- sach_m1_p[complete.cases(sach_m1_p),]

sach_m1_p$att <- NA
sach_m1_p$att[which(sach_m1_p$Target==0)] <- sach_m1_p$Sphere[which(sach_m1_p$Target==0)]
sach_m1_p$att[which(sach_m1_p$Target!=0)] <- sach_m1_p$Sphere[which(sach_m1_p$Target!=0)] - sach_m1_p$Target[which(sach_m1_p$Target!=0)]

attvach3_p <- data.frame('Preop'= c(mean(sach_m1_p$Sphere,na.rm=TRUE),sd(sach_m1_p$Sphere)), 'M1' = c(mean(sach_m1_p$Sph_M1),sd(sach_m1_p$Sph_M1)),row.names=c('Mean', 'SD'))
sach_m1_p <- sach_m1_p %>% select(att, Sph_ach_M1)

#### Att vs Ach Sphere Patients 1 Months Mix and Match ####
sach_m1_m <- data %>%
  filter(SurgeryLens=="Mix and Match") %>%
  select(Sphere, Target, Sph_ach_M1, Sph_M1)
sach_m1_m <- sach_m1_m[complete.cases(sach_m1_m),]

sach_m1_m$att <- NA
sach_m1_m$att[which(sach_m1_m$Target==0)] <- sach_m1_m$Sphere[which(sach_m1_m$Target==0)]
sach_m1_m$att[which(sach_m1_m$Target!=0)] <- sach_m1_m$Sphere[which(sach_m1_m$Target!=0)] - sach_m1_m$Target[which(sach_m1_m$Target!=0)]

attvach3_m <- data.frame('Preop'= c(mean(sach_m1_m$Sphere,na.rm=TRUE),sd(sach_m1_m$Sphere,na.rm=TRUE)), 'M1' = c(mean(sach_m1_m$Sph_M1,na.rm=TRUE),sd(sach_m1_m$Sph_M1,na.rm=TRUE)),row.names=c('Mean', 'SD'))
sach_m1_m <- sach_m1_m %>% select(att, Sph_ach_M1)

#### 1 Month BCDVA Line Loss of 2 or more ####

y <- data %>% 
  select(PatientID, EyeID, Age, Sphere, Cyl, BCDVA_LM, M1_BCDVA_LM, Sph_M1, SurgeryLens)
y <- y[complete.cases(y),]
y$BCDVA_LM <- as.numeric(y$BCDVA_LM)
y$M1_BCDVA_LM <- as.numeric(y$M1_BCDVA_LM)
y$BCDVA_LM[is.na(y$BCDVA_LM)] <- 1.8
y$M1_BCDVA_LM[is.na(y$M1_BCDVA_LM)] <- 1.8

y$m1_bcdlc <- 10*(y$BCDVA_LM - y$M1_BCDVA_LM)
y$m1_bcdlc <- round(y$m1_bcdlc, 0)

y <- y %>% mutate(m1_gain_bcdlc = case_when(m1_bcdlc <= -2 ~ "Loss 2 or more",
                                            m1_bcdlc == -1 ~ "Loss -1",
                                            m1_bcdlc == 0 ~ "No Change",
                                            m1_bcdlc == 1 ~ "Gain 1",
                                            m1_bcdlc >= 2 ~ "Gain 2 or more"))

bcdva1 <- as.data.frame(table(y$m1_gain_bcdlc,y$SurgeryLens)) %>%
  pivot_wider(names_from="Var2", values_from="Freq")
bcdva1$BilatOPercent <- round(bcdva1$`Bilateral Odyssey`/sum(bcdva1$`Bilateral Odyssey`), digits = 4)
bcdva1$BilatPPercent <- round(bcdva1$`Bilateral PureSee`/sum(bcdva1$`Bilateral PureSee`), digits = 4)
bcdva1$MixPercent <- round(bcdva1$`Mix and Match`/sum(bcdva1$`Mix and Match`), digits = 4)
bcdva1 <- bcdva1 %>% select(-`Bilateral Odyssey`, -`Bilateral PureSee`, - `Mix and Match`)

bcdvam1 <- y %>%
  filter(m1_gain_bcdlc == "Loss 2 or more") %>%
  select(PatientID, EyeID, Age, Sphere, Cyl, m1_gain_bcdlc, Sph_M1, M1_BCDVA_LM, SurgeryLens)

bcdva_loss_px1 <- unique(bcdvam1[c(1,2)])
  
p <- bcdvam1 %>%
  select(PatientID, SurgeryLens) %>%
  group_by(SurgeryLens) %>%
  summarise(Freq = n())

e <- bcdvam1 %>%
  select(PatientID, EyeID, SurgeryLens) %>%
  group_by(SurgeryLens) %>%
  summarise(Freq = n())

u <- bcdvam1 %>%
  select(PatientID, Age, SurgeryLens) %>%
  unique()

age <- u %>%
  group_by(SurgeryLens) %>%
  summarise(mean = mean(Age, na.rm=TRUE),
            min = min(Age, na.rm=TRUE),
            max = max(Age, na.rm=TRUE),
            sd = sd(Age, na.rm=TRUE))

sphere <- bcdvam1 %>%
  group_by(SurgeryLens) %>%
  summarise(mean = mean(Sphere, na.rm=TRUE),
            min = min(Sphere, na.rm=TRUE),
            max = max(Sphere, na.rm=TRUE),
            sd = sd(Sphere, na.rm=TRUE))

sphere_m1 <- bcdvam1 %>%
  group_by(SurgeryLens) %>%
  summarise(mean = mean(Sph_M1, na.rm=TRUE),
            min = min(Sph_M1, na.rm=TRUE),
            max = max(Sph_M1, na.rm=TRUE),
            sd = sd(Sph_M1, na.rm=TRUE))

cylinder <- bcdvam1 %>%
  group_by(SurgeryLens) %>%
  summarise(mean = mean(Cyl, na.rm=TRUE),
            min = min(Cyl, na.rm=TRUE),
            max = max(Cyl, na.rm=TRUE),
            sd = sd(Cyl, na.rm=TRUE))

bcdva <- list( "Patient" = p,
               "Eyes" = e,
               "Age" = age,
               "Sphere" = sphere,
               "Sphere M1" = sphere_m1,
               "Cyl" = cylinder
)

write.xlsx(bcdva, paste(path, "/tables/", "1M BCDVA.xlsx", sep = ""))


#### Enhancement Rates - Post IOL ####
total_enh6 <- data %>% #getting total for 6 months ago
  filter(SurgeryDateTime <= Sys.Date()-183) %>%
  distinct(PatientID, EyeID, Surgery) %>%
  group_by(Surgery) %>%
  summarise(Total=n())

enh6 <- data %>%
  filter(SurgeryDateTime <= Sys.Date()-183) %>%
  select(PatientID, EyeID, SurgeryDateTime, Surgery, EnhDate, Enh_SurgeryGroup, 
         Enh_Surgery) %>%
  filter(!is.na(EnhDate)) %>%
  mutate(DayDiff = as.numeric(EnhDate - SurgeryDateTime)) %>%
  filter(DayDiff > 90 & DayDiff <= 183) %>%
  distinct(PatientID, EyeID, Surgery) %>%
  group_by(Surgery) %>%
  count() %>%
  merge(total_enh6,by="Surgery") %>%
  mutate(Percent = n/Total,
         Tag = "Within 6 Months") %>%
  select(Surgery, Tag, Percent)

enh6_p <- data %>%
  filter(SurgeryDateTime <= Sys.Date()-183) %>%
  select(PatientID, EyeID, SurgeryDateTime, Surgery, EnhDate, Enh_SurgeryGroup, 
         Enh_Surgery) %>%
  mutate(DayDiff = as.numeric(EnhDate - SurgeryDateTime)) %>%
  filter((DayDiff > 90 & DayDiff <= 183) | is.na(DayDiff)) %>%
  mutate(Enhanced = case_when(!is.na(DayDiff) ~ 1, .default=0)) %>%
  summarise(chi = chisq_test(Enhanced, Surgery)$p,
            chisq_warning = chisq_warns(Enhanced, Surgery),
            fisher = fisher.test(Enhanced, Surgery)$p.value)

total_enh12 <- data %>% #getting total for 12 months ago
  filter(SurgeryDateTime <= Sys.Date()-365) %>%
  distinct(PatientID, EyeID, Surgery) %>%
  group_by(Surgery) %>%
  summarise(Total=n())

enh12 <- data %>%
  filter(SurgeryDateTime <= Sys.Date()-365) %>%
  select(PatientID, EyeID, SurgeryDateTime, Surgery, EnhDate, Enh_SurgeryGroup, 
         Enh_Surgery) %>%
  filter(!is.na(EnhDate)) %>%
  mutate(DayDiff = as.numeric(EnhDate - SurgeryDateTime)) %>%
  filter(DayDiff > 90 & DayDiff <= 365) %>%
  distinct(PatientID, EyeID, Surgery) %>%
  group_by(Surgery) %>%
  count() %>%
  merge(total_enh12,by="Surgery") %>%
  mutate(Percent = n/Total,
         Tag = "Within 1 Year") %>%
  select(Surgery, Tag, Percent)

enh1 <- rbind(enh6 ,enh12)



enh <- data %>% select(PatientID, EyeID, SurgeryLens, SurgeryDateTime, Surgery, EnhDate, Enh_SurgeryGroup, Enh_Surgery)

enh <- enh[which(!is.na(enh$EnhDate)),]
enh <- enh[which(enh$Enh_SurgeryGroup!="YAG"),]
enh$DayDiff <- as.numeric(difftime(as.Date(enh$EnhDate), as.Date(enh$SurgeryDateTime),units="days"))
enh <- enh[which(enh$DayDiff>0),]

test <- enh %>%
  filter(DayDiff > 90 & DayDiff <= 365) #Within 1 Year
test6m <- enh %>%
  filter(DayDiff > 90 & DayDiff <= 183) #Within 6 Months

latest_enh <- test %>%
  filter(EnhDate==max(EnhDate)) %>%
  select(EnhDate, DayDiff)

total <- rbind(totalo, totalp, totalm)
total <- cbind(total, c("Bilateral Odyssey", "Bilateral PureSee", "Mix and Match")) %>%
  rename("SurgeryLens" = "c(\"Bilateral Odyssey\", \"Bilateral PureSee\", \"Mix and Match\")")

enh_6m <- test6m %>%
  group_by(SurgeryLens) %>%
  summarise(Freq = n())
enh_6m <- merge(enh_6m, total, by="SurgeryLens")
enh_6m$percent <- round(enh_6m$Freq/enh_6m$Eyes, digits = 6)
enh_6m$Tag <- "Within 6 Months"
enh_6m <- enh_6m %>%
  select(SurgeryLens, percent,Tag)

enh_1y <- test %>%
  group_by(SurgeryLens) %>%
  summarise(Freq = n())
enh_1y <- merge(enh_1y, total, by="SurgeryLens")
enh_1y$percent <- round(enh_1y$Freq/enh_1y$Eyes, digits = 6)
enh_1y$Tag <- "Within 1 Year"
enh_1y <- enh_1y %>%
  select(SurgeryLens, percent,Tag)

enh <- rbind(enh_6m, enh_1y)
enh <- enh %>%
  pivot_wider(names_from = "Tag", values_from = "percent")

# ## Exchanges ###
# exch <- data[grep("Exchange", data$Enh_Surgery), ]
# exch$ExDateDiff <- difftime(exch$EnhDate,exch$SurgeryDateTime, units = "days")
# 
# exch <- exch[which(exch$ExDateDiff>90), ]
# exch$DateDiff<-NA
# exch$DateDiff[which(exch$ExDateDiff>365)] <- "Over 12 Months"
# exch$DateDiff[which(exch$ExDateDiff>183 & exch$ExDateDiff<=365)] <- "Within 12 Months"
# exch$DateDiff[which(exch$ExDateDiff<=183)] <- "Within 6 Months"
# 
# if (dim(exch)[1] > dim(exch[which(!is.na(exch$Enh_Model)), ])[1] ){
#   print("WARNING: There are NA's in the Enhancement Models.")
# }
# 
# exch <- as.data.frame(table(exch$Model, exch$Enh_Model))
# exch <- exch[which(exch$Freq>0),]
# exch$Percent <- exch$Freq/total[1,2]
# colnames(exch) <- c("Primary Lens", "Exchange Lens", "Freq", )
# 
#### YAGS ####
yags <- data[grep("YAG", data$Enh_SurgeryGroup), ]

yags$DayDiff <- difftime(yags$EnhDate,yags$SurgeryDateTime, units = "days")

test <- yags %>%
  filter(DayDiff > 90 & DayDiff <= 365) #Within 1 Year
test6m <- yags %>%
  filter(DayDiff > 90 & DayDiff <= 183) #Within 6 Months

yags_6m <- test6m %>%
  group_by(SurgeryLens) %>%
  summarise(Freq=n()) %>%
  left_join(total, by="SurgeryLens") %>%
  mutate(percent=Freq/Eyes, Tag="Within 6 Months") %>%
  select(SurgeryLens,percent,Tag)

yags_1y <- test %>%
  group_by(SurgeryLens) %>%
  summarise(Freq=n()) %>%
  left_join(total, by="SurgeryLens") %>%
  mutate(percent=Freq/Eyes, Tag="Within 1 Year") %>%
  select(SurgeryLens,percent,Tag)

yags <- rbind(yags_6m, yags_1y)%>%
  pivot_wider(names_from = "SurgeryLens", values_from = "percent")

#### Complications ####
sid <- toString(unique(data$SurgeryID))

query <- paste('SET TRANSACTION ISOLATION LEVEL READ UNCOMMITTED
SELECT DISTINCT
S.PatientID,
ese.EyeID,
lst.DescriptionEnGB as Surgery,
lstg.DescriptionEnGB AS SurgeryGroup,

/* Complications */
spo.PostOpDateTime,
lpoefp.DescriptionEnGB AS Complication,
lspoevt.DescriptionEnGB AS PostOp
From EMR.dbo.tSurgery S

Left Join EMR.dbo.tEyeSurgeryEye ESE
On ESE.SurgeryID = S.SurgeryID

Left Join EMR.dbo.tLookupSurgeryType AS lst
On ese.SurgeryTypeID = lst.SurgeryTypeID

LEFT JOIN EMR.dbo.tLookupSurgeryTypeGroup as lstg
on lstg.SurgeryTypeGroupID = lst.SurgeryTypeGroupID

Left Join EMR.dbo.tLookupSurgeryTreatmentType AS lstt 
ON lstt.SurgeryTreatmentTypeID = ESE.SurgeryTreatmentTypeID

LEFT JOIN EMR.dbo.tSurgeryPostOp as spo
ON spo.SurgeryID = s.SurgeryID

LEFT JOIN EMR.dbo.tEyeSurgeryPostOpEye as espoe
On espoe.SurgeryPostOpID = SPO.SurgeryPostOpID
and espoe.EyeID = ese.EyeID

LEFT JOIN EMR.dbo.tEyeSurgeryPostOpEyeFlapPocket as poefp
ON poefp.SurgeryPostOpID = espoe.SurgeryPostOpID
AND poefp.EyeID = ese.EyeID
                
LEFT JOIN EMR.dbo.tLookupPostopFlapPocket as lpoefp
ON lpoefp.PostOpFlapPocketID = poefp.PostOpFlapPocketID

LEFT JOIN EMR.dbo.tEyeSurgeryPostOp as espo
ON espo.SurgeryPostOpID = poefp.SurgeryPostOpID

LEFT JOIN EMR.dbo.tLookupSurgeryPostOpEyeVisitType as lspoevt
ON lspoevt.SurgeryPostOpEyeVisitTypeID = espo.SurgeryPostOpEyeVisitTypeID

Left Join Estate.dbo.LocationView Surgery_Lv
On Surgery_Lv.LocationID = S.SurgeryLocationID

LEFT JOIN Estate.dbo.tLocation_Region as lr
on lr.LocationID = s.SurgeryLocationID

LEFT JOIN EMR.dbo.tLookupSurgeryPostOpEyeVisitType as pot
	ON pot.SurgeryPostOpEyeVisitTypeID = espo.SurgeryPostOpEyeVisitTypeID
	
WHERE s.SurgeryDateTime BETWEEN ', "'", mindate , "'", ' AND ', "'", maxdate , "'",'
and s.SurgeryID IN (', sid, ')
and lr.RegionID in (1,2,3,4,5,6)
and lstt.SurgeryTreatmentTypeID = 1
and S.IsDeleted IN (0)
and lstg.SurgeryTypeGroupID IN (2)
and espo.SurgeryPostOpEyeVisitTypeID NOT IN (0, 11, 12)
and poefp.PostOpFlapPocketID not in (125)--normal post operative appearance', sep='')

comp1 <- dbGetQuery(con, query)
COMP <- comp1
COMP <- merge(COMP,data %>% select(PatientID, EyeID, SurgeryLens, SurgeryDateTime), by=c("PatientID","EyeID"),all.x=TRUE)

# Complications cleaning
remove <- read_excel("S:\\Clinical Services\\Research Assistants\\Surgeon Outcomes - NEW\\Complication List - Surgeon Outcomes Packs - 17022026 v3.xlsx",
                     sheet="IOL")
remove <- remove %>% filter(`Include In Outcomes` == "Yes") %>% select(-Combine, -`Include In Outcomes`) %>% unique() %>%
  mutate(`Timeline 1` = case_when(`Timeline 1`== "Surgery" ~ 0,
                                  `Timeline 1`== "20 Days PO" ~ 20,
                                  .default=NA),
         `Timeline 2` = case_when(`Timeline 2`== "90 Days PO" ~ 90,
                                  `Timeline 2`== "No Time Limit" ~ 900000,
                                  `Timeline 2`== "No Outer Limit" ~ 900000,
                                  .default=NA))
COMP <- merge(COMP, remove, by="Complication",all.x=TRUE)
COMP$Diff <- difftime(COMP$PostOpDateTime, COMP$SurgeryDateTime)
COMP <- COMP %>%
  filter(Diff <= `Timeline 2` & Diff >= `Timeline 1`) %>%
  select(PatientID, EyeID, SurgeryLens, Complication, PostOp) %>%
  unique()

#### PCO's ####
PCO <- subset(COMP, COMP$Complication=='Posterior Capsular Opacification (PCO)')

comp_6m <- subset(PCO, PCO$PostOp =='1 Day' | PCO$PostOp =='1 Week' | PCO$PostOp =='1 Month' | PCO$PostOp =='3 Months' | PCO$PostOp =='6 Months')
comp_6m <- comp_6m %>%
  select(PatientID, EyeID, Complication, SurgeryLens) %>%
  unique()
comp_6m$Tag <- 'Within 6 Months'

comp_1y <- subset(PCO, PCO$PostOp =='1 Day' | PCO$PostOp =='1 Week' | PCO$PostOp =='1 Month' | PCO$PostOp =='3 Months' | PCO$PostOp =='6 Months' | PCO$PostOp =='12 Months')
comp_1y <- comp_1y %>%
  select(PatientID, EyeID, Complication, SurgeryLens) %>%
  unique()
comp_1y$Tag <- 'Within 1 Year'

pcos <- rbind(comp_6m, comp_1y)
pcos <- as.data.frame(table(pcos$Tag, pcos$SurgeryLens)) %>%
  pivot_wider(names_from="Var2", values_from="Freq") %>%
  adorn_totals("col")
pcos$BilatOPercent <- round(pcos$`Bilateral Odyssey`/pcos$Total, digits = 4)
pcos$BilatPPercent <- round(pcos$`Bilateral PureSee`/pcos$Total, digits = 4)
pcos$MixPercent <- round(pcos$`Mix and Match`/pcos$Total, digits = 4)
pcos$BilatOPercentPatients <- round(pcos$`Bilateral Odyssey`/total$Eyes[which(total$SurgeryLens=="Bilateral Odyssey")], digits = 4)
pcos$BilatPPercentPatients <- round(pcos$`Bilateral PureSee`/total$Eyes[which(total$SurgeryLens=="Bilateral PureSee")], digits = 4)
pcos$MixPercentPatients <- round(pcos$`Mix and Match`/total$Eyes[which(total$SurgeryLens=="Mix and Match")], digits = 4)
pcos <- pcos %>%
  select(-`Bilateral Odyssey`, -`Bilateral PureSee`, -`Mix and Match`, -Total)

# Complications table
comptab <- COMP %>%
  select(PatientID, EyeID, SurgeryLens, Complication) %>%
  unique() %>%
  group_by(Complication, SurgeryLens) %>%
  summarise(Count=n()) %>%
  pivot_wider(names_from="SurgeryLens", values_from="Count") %>%
  select(Complication, `Bilateral Odyssey`, `Bilateral PureSee`, `Mix and Match`)
comptab$BilatOPercent <- round(comptab$`Bilateral Odyssey`/sum(comptab$`Bilateral Odyssey`,na.rm=TRUE), digits = 4)
comptab$BilatPPercent <- round(comptab$`Bilateral PureSee`/sum(comptab$`Bilateral PureSee`, na.rm=TRUE), digits = 4)
comptab$MixPercent <- round(comptab$`Mix and Match`/sum(comptab$`Mix and Match`), digits = 4)
comptab$BilatOPercentPatients <- round(comptab$`Bilateral Odyssey`/total$Eyes[which(total$SurgeryLens=="Bilateral Odyssey")], digits = 4)
comptab$BilatPPercentPatients <- round(comptab$`Bilateral PureSee`/total$Eyes[which(total$SurgeryLens=="Bilateral PureSee")], digits = 4)
comptab$MixPercentPatients <- round(comptab$`Mix and Match`/total$Eyes[which(total$SurgeryLens=="Mix and Match")], digits = 4)
comptab <- comptab %>% arrange(-`Bilateral PureSee`)

#### BCDVA Loss 2+ Comps ####
comps <- merge(COMP, bcdva_loss_px1, by=c("PatientID", "EyeID"))
month1 <- subset(comps, comps$PostOp=='1 Month')
month1 <- month1 %>%
  select(PatientID, EyeID, Complication, SurgeryLens) %>%
  unique()

month1 <- as.data.frame(table(month1$Complication, month1$SurgeryLens))
month1$percent <- month1$Freq/(dim(bcdva_loss_px1)[1])
month1 <- month1 %>%
  arrange(-Freq) %>%
  slice_head(n=3)
month1$Tag <- "1 Month"

bcdva_loss_comps1_o <- month1[which(month1$Var2=="Bilateral Odyssey"),]
bcdva_loss_comps1_p <- month1[which(month1$Var2=="Bilateral PureSee"),]
bcdva_loss_comps1_m <- month1[which(month1$Var2=="Mix and Match"),]

# #### UCDVA >6/6 PCO ###
# CID <- data %>% select(PatientID, focal, M3_Binoc_UCDVA_SM)
# CID <- CID[complete.cases(CID),]
# month <- CID %>% select(PatientID, focal, M3_Binoc_UCDVA_SM)
# month <- subset(month, month$focal == "Monofocal")
# month <- subset(month, month$M3_Binoc_UCDVA_SM == "6/3" | month$M3_Binoc_UCDVA_SM == "6/3.8" | month$M3_Binoc_UCDVA_SM == "6/4" | month$M3_Binoc_UCDVA_SM == "6/4.8" | month$M3_Binoc_UCDVA_SM == "6/5" | month$M3_Binoc_UCDVA_SM == "6/6")
# 
# comps <- subset(PCO, PCO$PostOp =='1 Day' | PCO$PostOp =='1 Week' | PCO$PostOp =='1 Month' | PCO$PostOp =='3 Months')
# month$PCO <- 0
# month$PCO[which(month$PatientID %in% comps$PatientID)] <- 1
# 
# ucdvaPCO <- sum(month$PCO)/dim(month)[1]

#### GRAPHS ####

graph <- list("Sphere" = sph,
              "MSE" = mse,
              "Cyl" = cyl, 
              "Pre vs Post UCDVA BO" = prepost_ucdva_o,
              "Pre vs Post UCDVA BP" = prepost_ucdva_p,
              "Pre vs Post UCDVA MM" = prepost_ucdva_m,
              "Pre B vs Post UCDVA BO" = prebpostu_cdva_o,
              "Pre B vs Post UCDVA BP" = prebpostu_cdva_p,
              "Pre B vs Post UCDVA MM" = prebpostu_cdva_m,
              "Pre vs Post BCDVA BO" = prepost_bcdva_o,
              "Pre vs Post BCDVA BP" = prepost_bcdva_p,
              "Pre vs Post BCDVA MM" = prepost_bcdva_m,
              "UCDVA Line Change" = ucdva_lc,
              "Binoc UCIVA" = binoc_uciva,
              "UCIVA Line Change" = uciva_lc,
              "Binoc UCNVA" = binoc_ucnva,
              "UCNVA BO" = ucnva_o,
              "UCNVA BP" = ucnva_p,
              "UCNVA MM" = ucnva_m,
              "UCNVA Monoc BO" = monoc_ucnva_o,
              "UCNVA Monoc BP" = monoc_ucnva_p,
              "UCNVA Monoc MM" = monoc_ucnva_m,
              "UCNVA Line Change" = ucnva_lc,
              "Mean changes" = meanchange,
              "Refractive Pred" = refpred,
              "RefPred Tested" = refprednos,
              "Att vs Ach BO" = sach_m1_o,
              "stats BO" = attvach3_o,
              "Att vs Ach BP" = sach_m1_p,
              "stats BP" = attvach3_p,
              "Att vs Ach MM" = sach_m1_m,
              "stats MM" = attvach3_m,
              "BCDVA 1M" = bcdva1,
              "BCDVA loss comps 1 BO" = bcdva_loss_comps1_o,
              "BCDVA loss comps 1 BP" = bcdva_loss_comps1_p,
              "BCDVA loss comps 1 MM" = bcdva_loss_comps1_m,
              "Enhancement" = enh,
              "Latest Enh Date" = latest_enh,
              "PCO" = pcos,
              "YAGS" = yags,
              "Complications Table" = comptab
)

write.xlsx(graph, paste(path, "/tables/", "CRO Data.xlsx", sep = ""))

# write.xlsx(data, paste(path, "test.xlsx", sep = ""),sheet="test")
# 
# 
# wbpath <- paste(path, "test.xlsx", sep = "")
# wb <- loadWorkbook(file=wbpath)
# 
# 
# dates <- as.data.frame(cbind(mindate,maxdate))
# writeData(wb=wb, sheet="test", x=dates, xy=c(1,1))
# writeData(wb=wb, sheet="test", x=data, xy=c(1,4))
# saveWorkbook(wb, wbpath, overwrite = TRUE)







