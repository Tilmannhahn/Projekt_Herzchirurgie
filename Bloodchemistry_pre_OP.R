### ------------ Install and load packages -------------------------------------
# Install packages
install.packages("dplyr")

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

###----- Correlation between blood chemistry prior to operation and mortality --
# Merge subjectIDs, Classes, and Code
BloodData <- Lab_values_classes[, c("SubjectIDs_laboratory", "Class", "Code", "Befundtext")]

MortalityData <- Clinical_data[, c("subjectID", "Dead")]

BloodMortalityData <- merge(BloodData, MortalityData,
                            by.x = "SubjectIDs_laboratory",
                            by.y = "subjectID",
                            all.x = TRUE)

# Only blood parameters determined prior to operation
BloodMortalityPrior <- subset(BloodMortalityData, Class == "prior")

# Patients with two surgeries were excluded
exclude_ids <- c("SUBJ_02864", "SUBJ_04827", "SUBJ_05751", "SUBJ_06060", 
                 "SUBJ_07024", "SUBJ_10577", "SUBJ_11539", "SUBJ_11910", 
                 "SUBJ_12660", "SUBJ_13283")
BloodMortality <- subset(BloodMortalityPrior, !(SubjectIDs_laboratory %in% exclude_ids))

# Reshape data
MeansBloodMortality <- BloodMortality %>%
  group_by(SubjectIDs_laboratory, Dead, Code) %>% # Group by SubjectID, Dead, and Code
  summarise(mean_value = mean(Befundtext, na.rm = TRUE), .groups = 'drop') %>% # Calculate mean values
  pivot_wider(names_from = Code, values_from = mean_value) # Transform Codes into columns

# Z-standardization of blood parameters
MeansBloodMortalityZ <- MeansBloodMortality %>%
  mutate(across(where(is.numeric) & !c(SubjectIDs_laboratory, Dead), # Apply to numeric columns except IDs and Dead
                ~ (.-mean(., na.rm = TRUE)) / sd(., na.rm = TRUE)))

# Create results table
results_table <- lapply(parameters, function(param) {
  # Subset without NA values for the specific parameter
  RegressionData <- MeansBloodMortalityZ %>%
    filter(!is.na(.data[[param]])) # Ignore NA values for the current parameter
  
  # Logistic regression
  model <- glm(
    formula = as.formula(paste("Dead ~", paste0("`", param, "`"))), # Backticks around column names
    data = RegressionData,
    family = binomial(link = "logit")
  )
  
  # Extract relevant results
  coef_summary <- summary(model)$coefficients
  data.frame(
    Parameter = param,
    Estimate = coef_summary[2, "Estimate"],
    Std_Error = coef_summary[2, "Std. Error"],
    Z_Value = coef_summary[2, "z value"],
    P_Value = coef_summary[2, "Pr(>|z|)"]
  )
})

# Combine results
results_table <- do.call(rbind, results_table)

# Sort table by p-value (optional)
results_table <- results_table %>%
  arrange(P_Value)

print(results_table)

# Add significance column and significance levels
results_table <- results_table %>%
  mutate(
    Significant = ifelse(P_Value < 0.05, "Significant", "Not significant"), # Categorize significance
    Significance = case_when( # Add stars based on p-value
      P_Value < 0.001 ~ "***",
      P_Value < 0.01 ~ "**",
      P_Value < 0.05 ~ "*",
      TRUE ~ ""
    )
  )

# Create bar chart
ggplot(results_table, aes(x = reorder(Parameter, Estimate), y = Estimate, fill = Significant)) +
  geom_bar(stat = "identity", color = "black") + # Create bars with fill color based on significance
  geom_errorbar(aes(ymin = Estimate - Std_Error, ymax = Estimate + Std_Error), width = 0.2) + # Add error bars
  geom_text(aes(label = Significance), hjust = -0.2, size = 5) + # Add stars as text
  coord_flip() + # Rotate parameters to y-axis
  labs(
    title = "Impact of Blood Parameters on Postoperative Mortality",
    x = "Parameter",
    y = "Coefficient (Estimate)",
    fill = "Significance"
  ) +
  scale_fill_manual(values = c("Significant" = "skyblue", "Not significant" = "lightgray")) + # Define colors
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5)) # Center the title

