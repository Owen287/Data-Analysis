# You should get a prompt to upload your data and then a prompt within the R
# console to type in the number of samples, the names of each sample, the size
# each sample.
# The console will also ask you for the well positions for each sample. Please
# type them in one at a time or else it won't work


# Press the source button to run all of the code at once

# Remove hashtag for any packages you don't have installed

#install.packages("tcltk")
#install.packages("tidyverse")
#install.packages("zoo")
#install.packages("writexl")
library(tcltk)
library(tidyverse)
library(zoo)
library(writexl)

# Load in the data
data <- read.csv(file.choose(), header = TRUE, stringsAsFactors = FALSE)

# Creating a vector containing new names of columns
names <- c("WellPosition" = "ImageSceneContainerName..Image.Scene.Container.Name",
           "Area" = "RegionsArea..Area..R", 
           "IntensityMean" = "RegionsIntensityMean_Cy3.T1..Intensity.Mean.Value.of.channel..Cy3.T1...R", 
           "IntensitySum" = "RegionsIntensitySum1_Cy3.T1..Intensity.Sum.of.channel..Cy3.T1...R",
           "ImageChannel" = "ImageChannelName..Image.Channel.Name",
           "RegionCount" = "RegionsCount..Count..I"
)

# Inserting the vector names into the column headers
data <- data%>%
  rename(all_of(names))
rm(names)
# Filling in the blank cells in WellPosition with the last valid cell above
# Have to delete row one as it only contains units which is not relevant
data <- data[-1,] 
data$WellPosition[data$WellPosition == ""] <- NA
data$WellPosition[data$WellPosition == " "] <- NA
data$WellPosition <- na.locf(data$WellPosition)

# Going to set the Area, IntensitySum and RegionCount columns to be numeric to allow for mean and sum functions
data[c("Area", 
       "IntensitySum", 
       "RegionCount")] <- lapply(data[c("Area", 
                                        "IntensitySum", 
                                        "RegionCount")], as.numeric)

# Creating 2 new columns with Area/RegionCount and IntensitySum/RegionCount
data <- data %>%
  mutate(
    AreaAvg = as.numeric(Area/RegionCount),
    Intensity = as.numeric(IntensitySum/RegionCount)
  )

# Split data into Cy3 and DAPI channels
foci_data <- data %>% filter(ImageChannel == "Cy3-T1")
nuclei_data <- data %>% filter(ImageChannel == "DAPI-T2")

# Creating a list that will separate each dataframe into separate wells
foci_wells <- split(foci_data, foci_data$WellPosition)
nuclei_wells <- split(nuclei_data, nuclei_data$WellPosition)

#Creating a matrix that will now have the user input: number of samples, name of
# the samples, sample size, well position of the samples 
# Get the console to ask the user the number of samples being used
sample_number <- readline(prompt = paste("Enter the number of samples: "))
sample_size <- numeric(sample_number) #creates blank numeric vector of size sample_number
sample_name <- vector("character", sample_number) #creates blank character vector of size sample_number
sample_positions <- vector("character") #creates blank character vector
# gets the samples names
for(n in 1:sample_number){
  sample_name[n] <- readline(prompt = paste("Enter the name of sample", n, ": ")) 
}
# gets the sample sizes
for(n in 1:sample_number){
  sample_size[n] <- readline(prompt = paste("Enter size of sample", sample_name[n], ": "))
}  
#preparing vectors for loop
sample_size <- as.numeric(sample_size) #makes sample size numeric so it can work in the loop
all_positions <- vector("list", sample_number) #creates an empty list of size "sample_size" 
# gets the well positions and creates a list (use a list as it allows for missing data)
for (n in 1:sample_number){ #outer loop. cycles through each sample
  sample_positions <- vector("character", sample_size[n]) #creates empty vector that will change size dependent on the sample its size
  for(w in 1:sample_size[n]){ # inner loop. w will be the size of sample n. 
    sample_positions[w] <- readline(prompt = paste("Enter the well positions for", sample_name[n], ": ")) #puts user values into said vector
    all_positions[[n]] <- sample_positions #pastes the vector containing user values into the nth element of the list
  }
}
# preparing vectors for matrix
sample_number <- as.numeric(sample_number) #turning sample number numeric for matrix
names(all_positions) <- sample_name #pastes the sample names into the list containing the positions
#creating blank matrix big enough to contain the number of samples and sample sizes
SampleMatrix <- matrix(NA, 
                       nrow = max(sample_size), ncol = sample_number)
#Pasting position values into the blank matrix and 
for (i in 1:sample_number) {
  matrix_wells <- all_positions[[i]] #retrieves the vector of well positions for the current sample (i) from the all_positions list and stores in matrix_wells
  SampleMatrix[1:length(matrix_wells), i] <- matrix_wells #pastes the matrix_wells values into the correct locations
}
# 1:length(matrix_wells) creates a sequence of index's up to the sample size of the current sample
# so the first part accesses the rows from 1 to the length of matrix_wells in column "i" 
# When the correct location has been accessed it then pastes in the matrix_wells values into the correct positions in the matrix
#pasting sample names into the matrix
colnames(SampleMatrix) <- sample_name
#clean up
rm(all_positions,i,matrix_wells,n,sample_name,sample_number,sample_positions,
   sample_size,w)


# Creating a dataframe for Nuclei_Count using the matrix
intl_count <- lapply(nuclei_wells, function(nuclei_wells) sum(nuclei_wells$RegionCount, na.rm = TRUE))
NucleiCount <- data.frame(WellPosition = names(intl_count), N_Count = as.numeric(unlist(intl_count)))
rm(intl_count)
# Extract count values for each well and put into a dataframe
CountMatrix <- SampleMatrix
index <- match(CountMatrix, NucleiCount$WellPosition)
valid_index <- !is.na(index)
CountMatrix[valid_index] <- NucleiCount$N_Count[index[valid_index]]
CountMatrixDirty <- as.data.frame(CountMatrix)
# Removing unused columns
NumericCount <- suppressWarnings(apply(CountMatrixDirty, 2, function(x) as.numeric(x))) #Makes the matrix numeric (using suppressWarnings as R thinks turning the unused WellPositions to NA is an error but its actually the intention)
Has_Data_Count <- apply(NumericCount, 2, function(x) !all(is.na(x))) #Creates a logical vector for if a column has numeric data inside/ isnt just full of NA values
CleanedCountMatrix <- NumericCount[, Has_Data_Count] #Creates a matrix from NumericCount but only includes columns from the previous line
Nuclei_Count <- as.data.frame(CleanedCountMatrix) #Turns matrix into a dataframe
rm(CleanedCountMatrix,CountMatrix,CountMatrixDirty,NumericCount,Has_Data_Count,index,valid_index)

# Creating a dataframe for Foci_Area using the created matrix
# Sum of AreaAvg for each well. Creating a dataframe to then paste into the blank matrix
intl_area <- lapply(foci_wells, function(foci_wells) mean(foci_wells$AreaAvg, na.rm = TRUE))
AreaWells <- data.frame(WellPosition = names(intl_area), Area = as.numeric(unlist(intl_area)))
# inserting AvgArea values into the SampleMatrix
AreaMatrix <- SampleMatrix #Creating a copy of the framework matrix
index <- match(AreaMatrix, AreaWells$WellPosition) # Getting index values for each result
valid_index <- !is.na(index) # removing NA values 
AreaMatrix[valid_index] <- AreaWells$Area[index[valid_index]] #Puts these valid index values into the matrix 
NumericArea <- suppressWarnings(apply(AreaMatrix, 2, function(x) as.numeric(x))) #Converts the values to numeric ones
Foci_Area <- as.data.frame(NumericArea) # Turns the data into a dataframe
rm(intl_area,AreaMatrix,AreaWells,NumericArea,index,valid_index) # cleaning up

# Creating a dataframe for Foci_Intensity
# Sum of Intensity 
intl_intensity <- lapply(foci_wells, function(foci_wells) sum(foci_wells$Intensity, na.rm = TRUE))
IntensityWells <- data.frame(WellPosition = names(intl_intensity), Intensity = as.numeric(unlist(intl_intensity)))
# Inserting Intensity values into the SampleMatrix
IntensityMatrix <- SampleMatrix
index <- match(IntensityMatrix, IntensityWells$WellPosition)
valid_index <- !is.na(index)
IntensityMatrix[valid_index] <- IntensityWells$Intensity[index[valid_index]]
NumericIntensity <- suppressWarnings(apply(IntensityMatrix, 2, function(x) as.numeric(x)))
Foci_Intensity <- as.data.frame(NumericIntensity)
rm(intl_intensity,IntensityMatrix,IntensityWells,NumericIntensity,index,valid_index)

# Creating a dataframe for Foci_per_Nuclei
intl_count_foci <- lapply(foci_wells, function(foci_wells) sum(foci_wells$RegionCount, na.rm = TRUE))
FociCount <- data.frame(WellPosition = names(intl_count_foci), F_Count = as.numeric(unlist(intl_count_foci)))
Wells_FpN <- merge(FociCount,NucleiCount) #Merging this and the previously generated NucleiCount from doing the Nuclei_Count dataframe
Wells_FpN <- Wells_FpN %>%
  mutate(FpN = F_Count/N_Count) # Creates a new column of foci/nuclei
Wells_FpN <- Wells_FpN %>% select(-F_Count, -N_Count) # removes the foci and nuclei count columns
rm(intl_count_foci, FociCount,NucleiCount)
# Putting the FpN values for each well into the SampleMatrix
FpN_Matrix <- SampleMatrix
index <- match(FpN_Matrix, Wells_FpN$WellPosition)
valid_index <- !is.na(index)
FpN_Matrix[valid_index] <- Wells_FpN$FpN[index[valid_index]]
NumericFpN <- suppressWarnings(apply(FpN_Matrix, 2, function(x) as.numeric(x)))
Foci_per_Nuclei <- as.data.frame(NumericFpN)
rm(FpN_Matrix,SampleMatrix,Wells_FpN,NumericFpN,index,valid_index)
rm(foci_data,foci_wells,nuclei_data,nuclei_wells)

# Exporting the data as a single .xlsx file with multiple sheets
sheet_list <- list(
  "Foci Area" = Foci_Area,
  "Foci Intensity" = Foci_Intensity,
  "Foci per Nuclei" = Foci_per_Nuclei,
  "Nuclei Count" = Nuclei_Count
)

# Generates the directory according to the user input
save_path <- save_path <- tclvalue(tkgetSaveFile(
  initialfile = "Foci Data.xlsx",
  filetypes = "{{Excel Files} {.xlsx}}"
))
# Saving the file according to the generated directory
if (save_path != ""){
  if(!grepl("\\.xlsx$", save_path)){ # This checks the files is .xlsx and if it isn#tKB it will paste the .xlsx format on the end of it
    save_path <- paste0(save_path, ".xlsx")
  }
  write_xlsx(sheet_list, path = save_path) # This is the code that writes the excel file to the user chosen destination
} else {
  cat ("No file selected")
}
rm(sheet_list,save_path) #clean up