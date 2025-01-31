### ------------ Install and load packages -------------------------------------
# Load packages
library(openxlsx)
library(dplyr)
library(tidyr)
library(ggplot2)

# Set working directory
setwd("/Users/tilmannhahn/Desktop/Projekt_Herzchirurgie/Working_directory")

# Import Excel file containing lab values divided into classes
Lab_values_classes <- read.xlsx("Lab_values_classes.xlsx")
Clinical_data <- read.xlsx("clinical_data_20250110.xlsx")

###------------------- Count samples per class ---------------------------------
Classes <- c(Lab_values_classes$Class)

# Count occurrences of each class
ClassCounts <- table(Classes)

# Display the counts as a data frame
ClassCounts_df <- data.frame(Class = names(ClassCounts), Count = as.vector(ClassCounts))

# Show the result
print(ClassCounts_df)

###------------- Filter for patients with two or more operations ---------------
SubjectIDs <- c(Lab_values_classes$SubjectIDs_laboratory)
ReferencePoints <- c(Lab_values_classes$ReferencePoints)

# Create a data frame combining SubjectIDs and ReferencePoints
SubjectReference_df <- data.frame(SubjectID = SubjectIDs, ReferencePoint = ReferencePoints)

# Find SubjectIDs with multiple unique ReferencePoints
SubjectsWithMultipleOperations <- SubjectReference_df %>%
  group_by(SubjectID) %>%                # Group by SubjectID
  summarize(UniqueReferencePoints = n_distinct(ReferencePoint)) %>% # Count unique ReferencePoints
  filter(UniqueReferencePoints > 1)     # Filter for those with more than one unique ReferencePoint

# Show the Subject IDs with multiple ReferencePoints (10 in total)
print(SubjectsWithMultipleOperations)

# After reviewing patients with multiple (two) operations, it was found that
# for all of them, blood parameters were also measured multiple times. For this reason,
# these patients were excluded from the next test.

###----Regression between blood chemistry prior to operation and organ damage---

# Merge subjectIDs, Classes, and Code
blood_data <- Lab_values_classes[, c("SubjectIDs_laboratory", "Class", "Code", "Befundtext")]

organ_data <- Clinical_data[, c("subjectID", 
                                "PostopICUOrganFailRenalFail", 
                                "PostopICUOrganFailSepsis", 
                                "PostopICUOrganFailLiverFail", 
                                "PostopICUOrganFailVisceralIschemia", 
                                "PostopICUOrganFailPulmoInsufficiency", 
                                "PostopICUOrganFailLowCardiacOutputSyndrome", 
                                "PostopICUOrganFailRethoracotomy",
                                "PostopNeuroStatus")]

# Convert "yes" to 1, "no" to 0, and empty fields to NA for organ failure columns
organ_failure_cols <- c("PostopICUOrganFailRenalFail", 
                        "PostopICUOrganFailSepsis", 
                        "PostopICUOrganFailLiverFail", 
                        "PostopICUOrganFailVisceralIschemia", 
                        "PostopICUOrganFailPulmoInsufficiency", 
                        "PostopICUOrganFailLowCardiacOutputSyndrome", 
                        "PostopICUOrganFailRethoracotomy")

organ_data[, organ_failure_cols] <- lapply(organ_data[, organ_failure_cols], 
                                           function(x) ifelse(x == "yes", 1, ifelse(x == "no", 0, NA)))

# Convert "pathologic" to 1, "without findings" to 0, and empty fields to NA in PostopNeuroStatus
organ_data$PostopNeuroStatus <- ifelse(organ_data$PostopNeuroStatus == "pathologic", 1, 
                                       ifelse(organ_data$PostopNeuroStatus == "without findings", 0, NA))

# Merge blood chemistry data with organ damage data
blood_organ_damage_data <- merge(blood_data, organ_data,
                                 by.x = "SubjectIDs_laboratory",
                                 by.y = "subjectID",
                                 all.x = TRUE)

# Only blood parameters determined prior to operation
blood_organ_damage_data <- subset(blood_organ_damage_data, Class == "prior")

# Patients with two surgeries were excluded
exclude_ids <- c("SUBJ_02864", "SUBJ_04827", "SUBJ_05751", "SUBJ_06060", 
                 "SUBJ_07024", "SUBJ_10577", "SUBJ_11539", "SUBJ_11910", 
                 "SUBJ_12660", "SUBJ_13283")
blood_organ_damage_data <- subset(blood_organ_damage_data, !(SubjectIDs_laboratory %in% exclude_ids))

# List of all organ failure variables
organ_failure_cols <- c("PostopICUOrganFailRenalFail", 
                        "PostopICUOrganFailSepsis", 
                        "PostopICUOrganFailLiverFail", 
                        "PostopICUOrganFailVisceralIschemia", 
                        "PostopICUOrganFailPulmoInsufficiency", 
                        "PostopICUOrganFailLowCardiacOutputSyndrome", 
                        "PostopICUOrganFailRethoracotomy",
                        "PostopNeuroStatus")

# Calculate the mean values of blood parameters per patient
MeansBloodOrganData <- blood_organ_damage_data %>%
  group_by(SubjectIDs_laboratory, Code, across(all_of(organ_failure_cols))) %>%
  summarise(mean_value = mean(Befundtext, na.rm = TRUE), .groups = 'drop') %>%
  pivot_wider(names_from = Code, values_from = mean_value)

# Standardization of blood parameters
MeansBloodOrganDataZ <- MeansBloodOrganData %>%
  mutate(across(where(is.numeric) & !c(SubjectIDs_laboratory, all_of(organ_failure_cols)), 
                ~ (.-mean(., na.rm = TRUE)) / sd(., na.rm = TRUE)))

# Extract all blood parameter columns (only Z-values)
parameters <- setdiff(names(MeansBloodOrganDataZ), c("SubjectIDs_laboratory", organ_failure_cols))

# Store results
all_results <- list()

# Regression for each organ variable
for (organ in organ_failure_cols) {
  organ_results <- lapply(parameters, function(param) {
    RegressionData <- MeansBloodOrganDataZ %>%
      filter(!is.na(.data[[param]]), !is.na(.data[[organ]]))  # Remove NA values
    
    if (nrow(RegressionData) > 5) {  # Minimum number of observations required
      model <- glm(
        formula = as.formula(paste(organ, "~", paste0("`", param, "`"))),
        data = RegressionData,
        family = binomial(link = "logit")
      )
      
      coef_summary <- summary(model)$coefficients
      return(data.frame(
        Organ = organ,
        Parameter = param,
        Estimate = coef_summary[2, "Estimate"],
        Std_Error = coef_summary[2, "Std. Error"],
        Z_Value = coef_summary[2, "z value"],
        P_Value = coef_summary[2, "Pr(>|z|)"]
      ))
    }
  })
  
  # Combine results for the organ
  all_results[[organ]] <- do.call(rbind, organ_results)
}

# Combine all results
final_results <- do.call(rbind, all_results)

# Sort results by P-value
final_results <- final_results %>%
  arrange(Organ, P_Value) %>%
  mutate(Significant = ifelse(P_Value < 0.05, "Significant", "Not significant"),
         Significance = case_when(
           P_Value < 0.001 ~ "***",
           P_Value < 0.01 ~ "**",
           P_Value < 0.05 ~ "*",
           TRUE ~ ""
         ))

# Save results as Excel files
for (organ in organ_failure_cols) {
  organ_data <- final_results %>% filter(Organ == organ)  
  if (nrow(organ_data) > 0) {  
    file_name <- paste0("Regression_Results_", organ, ".xlsx")  
    write.xlsx(organ_data, file_name, rowNames = FALSE)
  }
}

# separate plots for each organ
for (organ in organ_failure_cols) {
  organ_data <- final_results %>% filter(Organ == organ)  
  
  if (nrow(organ_data) > 0) {  
    p <- ggplot(organ_data, aes(x = reorder(Parameter, Estimate), y = Estimate, fill = Significant)) +
      geom_bar(stat = "identity", color = "black") +
      geom_errorbar(aes(ymin = Estimate - Std_Error, ymax = Estimate + Std_Error), width = 0.2) +
      geom_text(aes(label = Significance), hjust = -0.2, size = 5) +
      coord_flip() +
      labs(
        title = paste("Impact of Blood Parameters on", organ),
        x = "Blood Parameter",
        y = "Coefficient (Estimate)",
        fill = "Significance"
      ) +
      scale_fill_manual(values = c("Significant" = "skyblue", "Not significant" = "lightgray")) +
      theme_minimal() +
      theme(plot.title = element_text(hjust = 0.5))
    
    
    print(p)
  }
}