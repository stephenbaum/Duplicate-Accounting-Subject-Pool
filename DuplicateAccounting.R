
### duplicate detection in subject pool data ###
### code written by stephen baum ###
### last updated - 12.20.23 ###

#### LOAD IN REQUIRED PACKAGED ####

# install and load the required packages if not already installed
if (!require(dplyr, readxl, writexl)) {
  install.packages("dplyr", "readxl", "writexl")
}
library(dplyr)
library(readxl)
library(writexl)

#### LOAD IN SHEETS WITH SUBJECT POOL STUDENTS ####

# load in all three sheets
allsheets = sapply(readxl::excel_sheets("Rosters for Human Subject Pool - 9.11.23.xlsx"), # change the file name if needed
                    simplify = F, USE.NAMES = T,
                    function(X) readxl::read_excel("Rosters for Human Subject Pool - 9.11.23.xlsx", # change the file name here, too
                                                   sheet = X))

#### COMBINE SHEETS AND IDENTIFY DUPLICATES ####

# combine the datasets from each tab into one data frame
combined_data = bind_rows(
  allsheets$MGT100,   # these are names of individual tabs; might need to be changed
  allsheets$MKT370,   
  allsheets$OB360    
)

# identify and display duplicates based on the ID variable
duplicates = combined_data %>%
  group_by(ID) %>%    # do this by ID (or email, I suppose)
  filter(n() > 1)     # doing it by last name is inefficient, because many different people have the same last name

# print the duplicates
print(duplicates)
nrow(duplicates)

# assuming each duplicate is in two courses
total_count = nrow(duplicates)/2
total_count # there should be this many duplicates
# that is, unless someone is in all 3 courses

# checking if someone is in all three courses; can verify by just scanning sheet
triples = duplicates %>%
  group_by(ID) %>%
  filter(n() == 3) # in this case, no one has an ID that appears thrice

#### EXPORT FILE WITH NAMES OF DUPLICATES ####

# specify the path for the output Excel file on the desktop 
duplicates_output = data.frame(duplicates)

# actually create the file and save it to the desktop 
write_xlsx( 
  duplicates_output,"duplicates_output.xlsx", 
  col_names = TRUE, 
  format_headers = TRUE)


