### ------------ Install and load packages -------------------------------------
#install packages
install.packages("openxlsx")

#load packages
library(openxlsx)

#set working directory
setwd("/Users/tilmannhahn/Desktop/Projekt_Herzchirurgie/Working_directory")

#import tables
lab_values <- read.xlsx("lab_values_20250110.xlsx")
clinical_data <- read.xlsx("clinical_data_20250110.xlsx")

###--------------filtering blood chemistry for false o missing values-----------
#isolation of column 6, "befundtext"
Befundtext <- c(lab_values$Befundtext)

#filtering befundtext for false or missing values
#find empty cells and replace those with NA
Befundtext[trimws(Befundtext) == ""] <- NA
#find cells that contain no number and replace those with NA
Befundtext[!suppressWarnings(!is.na(as.numeric(Befundtext)))] <- NA


### ----------------Reference point---------------------------------------------
# Determination of the reference point for measuring 12h, 24h, ... post-operation intervals
# ------------------------------------------------------------------------------
# Many operation end times were not recorded, necessitating an alternative approach
# To address this, the average duration of all recorded operations is calculated
# Half of this average operation duration is then added to the start time of each operation
# This method approximates the reference point as the estimated midpoint of the operation
# ------------------------------------------------------------------------------
# isolation of column "PreopDiagTimingSurgSkinIncision" (start of operation)
PreopDiagTimingSurgSkinIncision <- c(clinical_data$PreopDiagTimingSurgSkinIncision) 
# isolation of column "PreopDiagTimingSurgSutureEnd" (end of operation)
PreopDiagTimingSurgSutureEnd <- c(clinical_data$PreopDiagTimingSurgSutureEnd) 

#filtering for false or missing values
#start of operation
PreopDiagTimingSurgSkinIncision[trimws(PreopDiagTimingSurgSkinIncision) == ""] <- NA
PreopDiagTimingSurgSkinIncision[!suppressWarnings(!is.na(as.numeric(PreopDiagTimingSurgSkinIncision)))] <- NA
#end of operation
PreopDiagTimingSurgSutureEnd[trimws(PreopDiagTimingSurgSutureEnd) == ""] <- NA
PreopDiagTimingSurgSutureEnd[!suppressWarnings(!is.na(as.numeric(PreopDiagTimingSurgSutureEnd)))] <- NA

# average length of operation
# base date for Excel times is January 1, 1900
excel_base_date <- as.Date("1899-12-30")  # Excel counts from 1900, -2 days correction

# konvertion from Excel-Format into POSIXct (date and time)
PreopDiagTimingSurgSkinIncision <- as.POSIXct(excel_base_date + PreopDiagTimingSurgSkinIncision, origin = "1970-01-01", tz = "UTC")
PreopDiagTimingSurgSutureEnd <- as.POSIXct(excel_base_date + PreopDiagTimingSurgSutureEnd, origin = "1970-01-01", tz = "UTC")

# calculation of the lengt of the operation (end - start) in minutes
OperationDurations <- as.numeric(difftime(PreopDiagTimingSurgSutureEnd, PreopDiagTimingSurgSkinIncision, units = "mins"))

# removal of NA values from the 'OperationDurations' variable
OperationDurations <- OperationDurations[!is.na(OperationDurations)]

#removal of negative durations (patient SUBJ_48707)
OperationDurations <- OperationDurations[OperationDurations >= 0]

# calculation of the mean of the operation duration
MeanOperationDuration <- mean(OperationDurations, na.rm = TRUE)

# Creation of ReferencePoints:
# For known values of PreopDiagTimingSurgSutureEnd, use this value
# For missing values (NA), use PreopDiagTimingSurgSkinIncision + MeanOperationDuration
ReferencePoints <- ifelse(
  !is.na(PreopDiagTimingSurgSutureEnd),  
  PreopDiagTimingSurgSutureEnd,         
  PreopDiagTimingSurgSkinIncision + MeanOperationDuration * 60  
)

### ----------------Classification----------------------------------------------
# Deviding the blood chemistry into classes depending on the time of sampling in relation to the operation
# ------------------------------------------------------------------------------
# Deviation into: prior to operation, to 12h after operation, 12-24h after operation, 24-48h after operation and > 48h after operation
# ------------------------------------------------------------------------------
#test if ReferencePoints has the same length as PreopDiagTimingSurgSkinIncision to make sure
# the ReferencePoints are connected to the correct SubjectIDs
# if TRUE ReferencePoints and SubjectIDs can be connected
SubjectIDs_clinical <- c(clinical_data$subjectID)
if (length(SubjectIDs_clinical) == length(ReferencePoints)) {
  ReferenceData <- data.frame(
    SubjectIDs_clinical = SubjectIDs_clinical,
    ReferencePoints = ReferencePoints
  )
} else {
  cat("Die Länge von SubjectIDs und ReferencePoints stimmt NICHT überein: FALSE\n")
}

# isolation of subjectID, CODE and Auftragsdatum_converted
SubjectIDs_laboratory <- c(lab_values$subjectID)
Code <- c(lab_values$CODE)
Auftragsdatum <- c(lab_values$Auftragsdatum)
Auftragsdatum_converted <- as.POSIXct(excel_base_date + Auftragsdatum, origin = "1970-01-01", tz = "UTC")

# dataframe containing subjectID, CODE, Befundtext and Auftragsdatum_converted
LaboratoryData <- data.frame(
  SubjectIDs_laboratory = SubjectIDs_laboratory,
  Code = Code,
  Befundtext = Befundtext,
  Auftragsdatum_converted = Auftragsdatum_converted
)

# Merge Laboratory and ReferenceData based on SubjectID
merged_data <- merge(LaboratoryData, ReferenceData, 
                     by.x = "SubjectIDs_laboratory", 
                     by.y = "SubjectIDs_clinical", 
                     all.x = TRUE)

# Convert Auftragsdatum_converted to POSIXct for comparison
merged_data$Auftragsdatum_converted <- as.POSIXct(
  merged_data$Auftragsdatum_converted,
  format = "%Y.%m.%d %H:%M:%S",  # Format of the input string
  tz = "UTC"
)

# Calculate time difference in hours between ReferencePoint and Auftragsdatum_converted
merged_data$TimeDifference <- as.numeric(difftime(
  merged_data$Auftragsdatum_converted, 
  merged_data$ReferencePoints, 
  units = "hours"
))

# Classify based on TimeDifference
merged_data$Class <- cut(
  merged_data$TimeDifference,
  breaks = c(-Inf, 0, 12, 24, 48, 96, Inf),
  labels = c("prior", 
             "to 12h", 
             "12-24h", 
             "24-48h",
             "48-96h",
             ">96h"),
  right = FALSE
)

# Inspection of the result
head(merged_data[, c("SubjectIDs_laboratory", "Auftragsdatum_converted", 
                     "ReferencePoints", "TimeDifference", "Code", "Class")])

# Export into Excel-sheet
write.xlsx(merged_data, file = "Lab_values_classes.xlsx", rowNames = FALSE)

# print 
cat("The dataset 'merged_data.xlsx' was succesfully stored in working directory.\n")


