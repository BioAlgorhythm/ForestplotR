library(forestplot)
library(readxl)

svg("Early Single Biomarkers test.svg", 
    width = 25, 
    height = 15)

### (A) Import the Excel file with the data.
# The headers, in the first row, must be: Subtitle	Study	Biomarker	ES	LowerCI	UpperCI	Weight	IsSubtitle	IsSummary
# Subtitle (the class of outcome, such as graft-related)  Study	Biomarker (the biomarker name)	ES (measure of effect)	LowerCI	UpperCI	Weight (from 0 to 100)	IsSubtitle (if the row is a subtitle)	IsSummary (if the row refers to an overall)
# Of note, blank rows can be added in the Excel sheet, but they must be flagged as a subtitle (IsSubtitle == 1) to remove NA values.
my_data <- read_excel(
  "forest plot generation.xlsx", 
  sheet = "EarlySingle"
)
my_data$IsSubtitle <- my_data$IsSubtitle == 1 # IsSubtitle flag is used to remove numeric values (in columns 4 and 5) for rows meant only as section headers.
my_data$IsSummary  <- my_data$IsSummary == 1 # IsSummary flag is used to bold the text and generate diamond shapes.
my_data$ES         <- as.numeric(my_data$ES)
my_data$LowerCI    <- as.numeric(my_data$LowerCI)
my_data$UpperCI    <- as.numeric(my_data$UpperCI)
my_data$Weight     <- as.numeric(my_data$Weight)
# For subtitle rows, these will become NA if the Excel cells are empty.

# Replace any carriage return/newline characters in the Biomarker column with an actual newline ("\n")
# or, if you prefer to remove them, replace with a space.
my_data$Biomarker <- gsub("[\r\n]+", "\n", my_data$Biomarker)

### (B) Define the Header Row for the Forest Plot Table
# 
# The header row is defined as:
# c("", "Study", "Biomarker", "HR/OR [95% CI]", "Weight")
#
# - The first element is empty (serving as a placeholder).
# - The second column is "Study".
# - The third column is "Biomarker".
# - The fourth column is labeled "HR/OR [95% CI]" to reflect the effect measure.
# - The fifth column is "Weight".
col_header <- c("", "Study", "Biomarker", "HR/OR [95% CI]", "Weight")

### (C) The data rows come from my_data, one row per entry
# Explanation: The last part of the ifelse is not necessary at the moment, but it would be required if the weight is omitted in some rows of the Excel document, leaving NA values.
# In our dataset, the weight value is not provided (i.e., it's NA) for some rows, typically the overall summary rows (is.summary == TRUE).
# When we use sprintf("%.2f", my_data$Weight) on an NA value, it produces the literal string "NA", which we do not want to display in our forest plot table.
# To avoid this, we use an ifelse() statement that checks if a weight is NA.
# If it is, we replace it with an empty string ""; otherwise, we format the weight normally (with two decimal places).
data_rows <- cbind(
  as.character(my_data$Subtitle),
  as.character(my_data$Study),
  as.character(my_data$Biomarker),
  paste0(
    sprintf("%.2f", my_data$ES), " [",
    sprintf("%.2f", my_data$LowerCI), ", ",
    sprintf("%.2f", my_data$UpperCI), "]"
  ),
  ifelse(is.na(my_data$Weight), "", sprintf("%.2f", my_data$Weight))
)

### (D) Combine them with rbind
# IsSubtitle flag is used to remove numeric values (in columns 4 and 5) for rows meant only as section headers.
# For rows flagged as pure subtitles (using my_data$IsSubtitle), we don't want any numeric data.
# Thus, for these rows (excluding the header), we set columns 4 (HR/OR [95% CI]) and 5 (Weight) to an empty string. This ensures that subtitle rows display only text and no numerical values.
tabletext <- rbind(col_header, data_rows)
tabletext[-1, ][my_data$IsSubtitle, 4] <- ""
tabletext[-1, ][my_data$IsSubtitle, 5] <- ""
# For rows flagged as summary (my_data$IsSummary == 1), even though a weight value is inputted and calculated to get a big dimension of the diamond, we do not want to display the value itself.
# Therefore, for these rows (excluding the header), we set the weight column (column 5) to an empty string.
tabletext[-1, ][my_data$IsSummary, 5] <- ""

### (E) Create vectors for the forest plot data
# First element is NA for the "header" row, then each row is your data.
mean  <- c(NA, my_data$ES)
lower <- c(NA, my_data$LowerCI)
upper <- c(NA, my_data$UpperCI)

### (F) Copy weights for boxsize calculation
weights_for_boxsize <- as.numeric(my_data$Weight)  # Create a copy of my_data$Weight
weights_for_boxsize[is.na(weights_for_boxsize)] <- 0.1  # Replace NA with a small value
# Log-transform and scale weights to avoid extreme differences in square sizes when displaying rows with very different weights.
log_weights <- log(weights_for_boxsize + 0.1)  # Handle small weights with +0.1
normalized_weights <- (log_weights - min(log_weights)) / (max(log_weights) - min(log_weights))
boxsize <- 0.15 + (normalized_weights * 0.1)  # Scale to range [0.15, 0.25]
boxsize <- c(NA, boxsize)  # Add NA for the header row

### (G) Add a label attribute to the xticks vector (this is for your ES scale on the X axis, adjust accordingly to your needs):
my_ticks <- c(0.020, 1.000, 4.000, 16.000, 64.000, 256.000, 1000.000)
attr(my_ticks, "labels") <- c("0.02", "1", "4", "16", "64", "256", "1000")

### (H) Generate the forest plot.
forestplot::forestplot(
  labeltext = tabletext,
  mean      = mean,
  lower     = lower,
  upper     = upper,
  clip      = c(0.02, 1000),  # Direct range in exponential form
  xticks    = my_ticks,
  
  txt_gp = fpTxtGp(
    label = gpar(
      cex = 1 # vertical spacing between labels
    ),
    ticks = gpar(
      cex = 1 # x-axis tick labels
    ),
    xlab  = gpar(
      cex = 1 # x-axis label
    ),
    title = gpar(
      cex = 1.4 # main title
    )
  ),
  
  # The is.summary flag marks a row as a summary row. This is the only function that can make the text bold.
  # It is used for both:
  #   - Overall summary rows to generate the diamond figure and make the text bold.
  #   - Subtitle rows to make the text bold.
  is.summary = c(TRUE, my_data$IsSummary), # TRUE here stands for the first row, which is the header row and has to be bold.
  
  # If your effect measure is HR or OR, the "no-effect" line is at 1:
  zero      = 1,         
  graph.pos = 4,
  xlog      = TRUE,
  
  # Align columns: 
  #   - columns 1 & 2 left-aligned, 
  #   - columns 3, 4 & 5 right-aligned
  #   (You can tweak further to your taste)
  align = c("l", "l", "l", "r", "r"),
  
  # Appearance:
  boxsize    = boxsize,
  xlab       = "log2",
  title      = "Early Single Biomarkers",
  colgap     = unit((6), "mm"), # Smaller gap for the first column
  graphwidth = unit(15, "cm"),
  lwd.ci     = 3,     # CI line width; adjust graph width
  lineheight = unit(0.7, "cm")
)

dev.off()
