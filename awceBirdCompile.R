############################Scripts to merge eBird data from AWC and AWC forms#####################
library(lubridate)
library(tidyverse)
library(readxl)
library(openxlsx)
library(stringr)
source("config.R")

# Toggle this to avoid unzipping the file again & again. Saves time.
unzipebd <- 1
unzipmyebd <- 0


#################Google Form Preprocessing###################################
form <- read_excel(paste0(inputdir,googleform))

# Add Year and Month
form$Year <- as.numeric(format(as.Date(form$`Date of visit`), "%Y"))
form$Month <- as.numeric(format(as.Date(form$`Date of visit`), "%m"))

# Create a field where checklist ID is obtained
form$ChecklistID <- sub(".*[=/]", "", form$`Link to your eBird List`)

# Remove entries where the submission is not during AWC months
form_awc <- form %>% filter (Month %in% awc_months)
form_not_awc <- form %>% filter(!(Month %in% awc_months))


################My Ebird File Preprocessing#################################
if (unzipmyebd)
{
  unzip(paste0(inputdir,myEbdFile, exdir = inputdir))
}

myEbdData <- readxl::read_excel(paste0(inputdir, "MyEBirdData.xlsx"))

# Add Month & Year
myEbdData$Year <- as.numeric(format(as.Date(myEbdData$Date), "%Y"))
myEbdData$Month <- as.numeric(format(as.Date(myEbdData$Date), "%m"))

# Take data from only AWC months
# Months are a bit hard-coded right now. Can change later.
myEbdData_year <- myEbdData %>%
                    filter(
                      (Month %in% awc_2nd_half_months & Year == startYear) |
                        (Month %in% awc_1st_half_months  & Year == endYear)
                    )

##################eBird Public Download Pre-processing##########################################

ebdfilename <- paste0("ebd_IN_", startYear, "12", "_",endYear,"02","_unv_rel", downloadMonth,"-", downloadYear)


# List the interested columns
preimp <-  c( 
  "GLOBAL.UNIQUE.IDENTIFIER",
  "TAXONOMIC.ORDER",
  "CATEGORY",
  "COMMON.NAME",
  "SCIENTIFIC.NAME",
  "SUBSPECIES.SCIENTIFIC.NAME",
  "EXOTIC.CODE",
  "AGE.SEX",
  "IBA.CODE",
  "LOCALITY",
  "LOCALITY.ID",
  "LOCALITY.TYPE",
  "STATE.CODE",
  "COUNTY.CODE",
  "OBSERVATION.DATE",
  "OBSERVATION.COUNT",
  "COUNTY",
  "STATE",
  "PROTOCOL.NAME",
  "OBSERVER.ID",
  "TIME.OBSERVATIONS.STARTED",
  "DURATION.MINUTES",
  "EFFORT.DISTANCE.KM",
  "SAMPLING.EVENT.IDENTIFIER",
  "ALL.SPECIES.REPORTED",
  "GROUP.IDENTIFIER",
  "LATITUDE",
  "LONGITUDE",
  "HAS.MEDIA",
  "APPROVED",
  "REVIEWED",
  "CHECKLIST.COMMENTS"
)

if (unzipebd)
{
  unzip(paste(inputdir, ebdfilename,'.zip',sep=''), exdir = inputdir)
}


# Read the header plus first row
nms <- read.delim( paste0 (inputdir, ebdfilename,".txt"),
                   nrows = 1, 
                   sep = '\t', 
                   header = T, 
                   quote = "", 
                   stringsAsFactors = F, 
                   na.strings = c ("", " ",NA)) 
nms <- names(nms)
nms [!(nms %in% preimp)] <- "NULL"
nms [nms %in% preimp] <- NA

# Basic ebd Dataset - all confirmed records
ebd <- read.delim(paste0(inputdir, ebdfilename,".txt"),
                  colClasses = nms,
                  #                  nrows = 100000, # For testing, this is useful
                  sep = '\t', 
                  header = T, 
                  quote = "", 
                  stringsAsFactors = F, 
                  na.strings = c ("", " ",NA)) 

# Records still in the review queue
ebd_unvetted <- read.delim(paste0(inputdir, ebdfilename,"_unvetted.txt"),
                           colClasses = nms,
                           #                  nrows = 100000, # For testing, this is useful
                           sep = '\t', 
                           header = T, 
                           quote = "", 
                           stringsAsFactors = F, 
                           na.strings = c ("", " ",NA)) 

stateTable <- ebd %>% distinct(STATE, STATE.CODE)

################JOIN steps starts####################################################
# See https://docs.google.com/presentation/d/1Jp1L8LCcbgBAFAwEnDTpLspffVpwFUSSvpPS9ebE1aU

# Step 1: Join eBird Dataset with AWC form and add a new field called GROUP.ID

form_plus_ebd <-  inner_join(form_awc, ebd, 
                           by = c("ChecklistID" = "SAMPLING.EVENT.IDENTIFIER"), 
                           relationship = 'many-to-many') %>%
                           mutate (GROUP.ID = ifelse (is.na(GROUP.IDENTIFIER), 
                                                      ChecklistID, GROUP.IDENTIFIER)) %>%
                           mutate(OBSERVATION.COUNT = case_when(
                            OBSERVATION.COUNT == "X" ~ 1,
                            TRUE ~ as.numeric(OBSERVATION.COUNT)
                           )) %>% 
                          mutate (`Date of visit` = format(`Date of visit`, "%Y-%m-%d"))


# Step 1a: Summary of all visits

visit_summary <-  form_plus_ebd %>% 
                    distinct(GROUP.ID, .keep_all = TRUE) %>%
                    select (LOCALITY,
                            `Name of Wetland Site counted`,
                            `Name of Parent Site`,
                            LOCALITY.TYPE,
                            COUNTY,
                            STATE,
                            `Your name`,
                            `Names of Participants`,
                            `Date of visit`,
                            OBSERVATION.DATE,
                            DURATION.MINUTES,
                            PROTOCOL.NAME,
                            EFFORT.DISTANCE.KM,
                            `Link to your eBird List`
                            )

# Step 1b: List of entries where checklist field is not usable. E.g. wrong checklist, incorrect format, missing
bad_checklists <- anti_join(form_awc, ebd, 
                                  by = c("ChecklistID" = "SAMPLING.EVENT.IDENTIFIER")) %>%
                          filter(
                              (Month %in% awc_2nd_half_months & Year == startYear) |
                              (Month %in% awc_1st_half_months  & Year == endYear)
                          )


# Step 2: Join eBird Dataset with MyEbirdData and add a new field called GROUP.ID

MyEbddata_plus_ebd <- inner_join(myEbdData_year, ebd, 
                                          by = c("Submission ID" = "SAMPLING.EVENT.IDENTIFIER",
                                                "Taxonomic Order" = "TAXONOMIC.ORDER")) %>% 
                      mutate (GROUP.ID = ifelse (is.na(GROUP.IDENTIFIER), ChecklistID, GROUP.IDENTIFIER))

# Step 2a: Some checklists are completely missing in eBird. Mostly its because the checklist is hidden due to problems.
# Note: This is not there in the figure.
problemChecklists <-  anti_join(myEbdData_year, ebd, 
                                by = c("Submission ID" = "SAMPLING.EVENT.IDENTIFIER")) %>%
                                  anti_join (ebd_unvetted, by = c("Submission ID" = "SAMPLING.EVENT.IDENTIFIER")) %>%
                                  select (`Submission ID`, County, Location, `State/Province`, Date) %>%
                                  distinct_all()
  

# Step 3: Join two datasets - combining form, ebd and MyEbirdData using GROUP.ID
full_data_set  <- inner_join (MyEbddata_plus_ebd, 
                             form_plus_ebd, 
                             by = c("GROUP.ID" = "GROUP.ID",
                                    "Taxonomic Order" = "TAXONOMIC.ORDER"))



# Step 3a: Anti join the two joined datasets - combining form, ebd and MyEbirdData using GROUP.ID
# This will give missing lists in forms while being present in MyEbirData
missing_checklists  <- anti_join (MyEbddata_plus_ebd %>% distinct(GROUP.ID, .keep_all = TRUE), 
                               form_plus_ebd %>% distinct(GROUP.ID, .keep_all = TRUE), 
                                      by = c("GROUP.ID" = "GROUP.ID")) %>%
                                        distinct(GROUP.ID, .keep_all = TRUE) 

#missing_records  <- anti_join (MyEbddata_plus_ebd, 
#                               form_plus_ebd, 
#                               by = c("GROUP.ID" = "GROUP.ID",
#                                      "Taxonomic Order" = "TAXONOMIC.ORDER")) %>% 
#                               anti_join(# remove rows whose GROUP.ID is in missing_lists
#                                        missing_lists %>% distinct(GROUP.ID),
#                                        by = "GROUP.ID"
#                                      )


# Steps 3a and 3b are combined in the presentation. 
# Step 3a: If the missing records are in unvetted file, then they need to be reviewed.
#unreviewed_records <- inner_join (MyEbddata_plus_ebd, ebd_unvetted, 
#                                                  by = c("Submission ID" = "SAMPLING.EVENT.IDENTIFIER",
#                                                           "Taxonomic Order" = "TAXONOMIC.ORDER"))

# Step 3b: If the missing records are not in unvetted file, then they have been unconfirmed by reviewers
#unconfirmed_records <- anti_join (missing_records, ebd_unvetted, 
#                                                  by = c("Submission ID" = "SAMPLING.EVENT.IDENTIFIER",
#                                                  "Taxonomic Order" = "TAXONOMIC.ORDER"))


###############WRITING OUTPUTS##########################

years <- unique(sort(form_awc$Year))
states <- unique(sort(MyEbddata_plus_ebd$STATE))

# Format all the dataframes that need to be written
f_form_not_awc <- form_not_awc %>% 
  select(`Your name`, 
         `Email address`, 
         `Name of Wetland Site counted`, 
         `Name of Parent Site`, 
         `Name of State in which the waterbird count was made`,
         `Date of visit`, 
         `Link to your eBird List`) %>%
  rename(Submitter = `Your name`,
         Email = `Email address`,
         Site = `Name of Wetland Site counted`,
         `Parent Site` = `Name of Parent Site`,
         State = `Name of State in which the waterbird count was made`,
         Date = `Date of visit`,
         List = `Link to your eBird List`)

# Format: Form + eBird merged data
f_form_plus_ebd <- form_plus_ebd %>%
  select(`Name of Wetland Site counted`,  
         `Name of Parent Site`, 
         LOCALITY,
         LOCALITY.TYPE,
         LATITUDE,
         LONGITUDE,
         COUNTY,
         STATE,
         OBSERVATION.DATE,
         `Date of visit`,
         TIME.OBSERVATIONS.STARTED,
         DURATION.MINUTES,
         PROTOCOL.NAME,
         EFFORT.DISTANCE.KM,
         COMMON.NAME,
         SCIENTIFIC.NAME,
         CATEGORY,
         OBSERVATION.COUNT,
         AGE.SEX,
         `Your name`,
         `Email address`,
         `Names of Participants`,
         `Link to your eBird List`) %>%
  rename(Site = `Name of Wetland Site counted`,
         `Parent Site` = `Name of Parent Site`,
         LocalityInEbird = LOCALITY,
         LocalityType = LOCALITY.TYPE,
         Latitude = LATITUDE,
         Longitude = LONGITUDE,
         District = COUNTY,
         State = STATE,
         DateInEbird = OBSERVATION.DATE,
         DateInForm = `Date of visit`,
         Time = TIME.OBSERVATIONS.STARTED,
         Duration = DURATION.MINUTES,
         Protocol = PROTOCOL.NAME,
         Distance = EFFORT.DISTANCE.KM,
         CommonName = COMMON.NAME,
         ScientificName = SCIENTIFIC.NAME,
         Category = CATEGORY,
         Count = OBSERVATION.COUNT,
         `Age/Sex` = AGE.SEX,
         Submitter = `Your name`,
         `Email address` = `Email address`,
         Participants = `Names of Participants`,
         List = `Link to your eBird List`)

f_visit_summary <- visit_summary %>%
      rename(Site = `Name of Wetland Site counted`,
             `Parent Site` = `Name of Parent Site`,
             LocalityInEbird = LOCALITY,
             LocalityType = LOCALITY.TYPE,
             District = COUNTY,
             State = STATE,
             DateInEbird = OBSERVATION.DATE,
             DateInForm = `Date of visit`,
             Duration = DURATION.MINUTES,
             Protocol = PROTOCOL.NAME,
             Distance = EFFORT.DISTANCE.KM,
             Submitter = `Your name`,
             Participants = `Names of Participants`,
             List = `Link to your eBird List`)

# Format: problem checklists
f_problemChecklists <- problemChecklists
colnames(f_problemChecklists) <- c("List", "District", "Locality", "State", "Date")

# Format: bad checklists
f_bad_checklists <- bad_checklists %>%
  select(`Your name`, 
         `Email address`, 
         `Name of Wetland Site counted`, 
         `Name of Parent Site`, 
         `Name of State in which the waterbird count was made`,
         `Date of visit`, 
         `Link to your eBird List`) %>%
  rename(Submitter = `Your name`,
         Email = `Email address`,
         Site = `Name of Wetland Site counted`,
         `Parent Site` = `Name of Parent Site`,
         State = `Name of State in which the waterbird count was made`,
         Date = `Date of visit`,
         List = `Link to your eBird List`)

f_missing_checklists <- missing_checklists %>% 
  select (`Submission ID`, 
          STATE, 
          COUNTY, 
          LOCALITY, 
          `OBSERVATION.DATE`, 
          `Duration (Min)`)   %>%
  rename(List = `Submission ID`,
         State = STATE,
         District = COUNTY,
         LocalityInEbird = LOCALITY,
         DateInEbird = OBSERVATION.DATE,
         Duration = `Duration (Min)`)

# Loop through each state and write Excel
for (state in states) {
  
  df_form_not_awc <- f_form_not_awc %>%
    filter(str_detect(State, fixed(state, ignore_case = TRUE)))
  
  df_form_plus_ebd <- f_form_plus_ebd %>%
    filter(State == state)
  
  df_visit_summary <- f_visit_summary %>%
    filter(State == state)
  
  df_counts <- df_form_plus_ebd %>% 
                      group_by(CommonName, Site) %>%                           # rows × columns to keep
                      summarise(max_count = max(Count, na.rm = TRUE),   # or COUNT, if that’s your field
                                .groups = "drop") %>% 
                      pivot_wider(
                        names_from  = Site,     # columns
                        values_from = max_count,    # values
                        values_fill = 0             # fill missing combos with 0
                      )
  
  
  df_missing_checklists <- f_missing_checklists %>%
    filter(State == state)
  
  df_bad_checklists <- f_bad_checklists %>%
    filter(str_detect(State, fixed(state, ignore_case = TRUE)))
  
  state_code <- stateTable %>%
    filter(STATE == state) %>%
    pull(STATE.CODE) %>%
    unique()
  
  if (length(state_code) > 0) {
    df_problemChecklists <- f_problemChecklists %>%
      filter(State %in% state_code)
  } else {
    df_problemChecklists <- tibble()
  }
  
  # File name
  outfile <- paste0(outputdir, "output_", startYear, "-", endYear, "-", gsub("\\s+", "_", state), ".xlsx")
  
  
  wb <- createWorkbook()
  
  addWorksheet(wb, "README")
  
  # Define sheet names and descriptions
  sheet_info <- tibble::tibble(
    SheetName = c("Valid_Visit Summary", "Valid_CountSummary", "Valid_FullData", "Error_ListswithNoFormEntry", "Error_BadListInGoogleForm", "Error_NotAwcLists", "Error_ListsWithIssues"),
    Description = c(
      "Summary of AWC visits",
      "Summary of AWC counts",
      "All date of checklists from eBird matched with form submissions. Review for accuracy and completeness.",
      "eBird checklists shared with awcindia that do not have any corresponding google form entry.",
      "Checklists provided by the submitter in the Google Form is not a valid link or valid checklist.",
      "Checklists submitted to the form are outside the December to Febuary period of AWC.",
      "Checklists flagged with issues like too long a list, duplicate entry or other issues by the editors of eBird and hence missing in data download."
    )
  )
  
  writeData(wb, "README", c("Sheet", "Description"), startCol = 1, startRow = 1)
  
  # Write formulas row by row so hyperlinks render properly
  for (i in seq_along(sheet_info$SheetName)) {
    sheet_name <- sheet_info$SheetName[i]
    desc <- sheet_info$Description[i]
    link_formula <- paste0('HYPERLINK("#\'', sheet_name, '\'!A1", "', sheet_name, '")')
    
    writeFormula(wb, sheet = "README", x = link_formula, startCol = 1, startRow = i + 1)
    writeData(wb, sheet = "README", x = desc, startCol = 2, startRow = i + 1, colNames = FALSE)
  }
  
  # Set column widths
  setColWidths(wb, "README", cols = 1:2, widths = "auto")

  addWorksheet(wb, "Valid_Visit Summary")
  writeData(wb, "Valid_Visit Summary", df_visit_summary)
  setColWidths(wb, "Valid_Visit Summary", cols = 1:ncol(df_visit_summary), widths = "auto")
  freezePane(wb, sheet = "Valid_Visit Summary", firstActiveRow = 2, firstActiveCol = 2)
  if(nrow(df_visit_summary) > 0)
  {
    addStyle(
      wb, 
      sheet = "Valid_Visit Summary", 
      style = grey_fill, 
      rows = seq(2, nrow(df_visit_summary) + 1, 2),  # +1 accounts for header row
      cols = 1:ncol(df_visit_summary), 
      gridExpand = TRUE
    )
  }
  
  addWorksheet(wb, "Valid_CountSummary")
  writeData(wb, "Valid_CountSummary", df_counts)
  setColWidths(wb, "Valid_CountSummary", cols = 1:ncol(df_counts), widths = "auto")
  headerStyle <- createStyle(textRotation = 60, halign = "center", valign = "center", wrapText = TRUE)
  setRowHeights(wb, sheet = "Valid_CountSummary", rows = 1, heights = 100)
  freezePane(wb, sheet = "Valid_CountSummary", firstActiveRow = 2, firstActiveCol = 2)
  
  # Apply style to header row
  addStyle(wb, 
           sheet = "Valid_CountSummary", 
           style = headerStyle, 
           rows = 1, 
           cols = 1:ncol(df_counts), 
           gridExpand = TRUE)
  
  if(nrow(df_counts) > 0)
  {
    addStyle(
      wb, 
      sheet = "Valid_CountSummary", 
      style = grey_fill, 
      rows = seq(2, nrow(df_counts) + 1, 2),  # +1 accounts for header row
      cols = 1:ncol(df_counts), 
      gridExpand = TRUE
    )
  }
  
  
  # Add worksheets and write data with auto width
  addWorksheet(wb, "Valid_FullData")
  writeData(wb, "Valid_FullData", df_form_plus_ebd)
  setColWidths(wb, "Valid_FullData", cols = 1:ncol(df_form_plus_ebd), widths = "auto")
  freezePane(wb, sheet = "Valid_FullData", firstActiveRow = 2, firstActiveCol = 2)
  
  if(nrow(df_form_plus_ebd) > 0)
  {
    addStyle(
      wb, 
      sheet = "Valid_FullData", 
      style = grey_fill, 
      rows = seq(2, nrow(df_form_plus_ebd) + 1, 2),  # +1 accounts for header row
      cols = 1:ncol(df_form_plus_ebd), 
      gridExpand = TRUE
    )
  }
  
  addWorksheet(wb, "Error_ListswithNoFormEntry")
  writeData(wb, "Error_ListswithNoFormEntry", df_missing_checklists)
  setColWidths(wb, "Error_ListswithNoFormEntry", cols = 1:ncol(df_missing_checklists), widths = "auto")
  
  addWorksheet(wb, "Error_BadListInGoogleForm")
  writeData(wb, "Error_BadListInGoogleForm", df_bad_checklists)
  setColWidths(wb, "Error_BadListInGoogleForm", cols = 1:ncol(df_bad_checklists), widths = "auto")
  
  addWorksheet(wb, "Error_NotAwcLists")
  writeData(wb, "Error_NotAwcLists", df_form_not_awc)
  setColWidths(wb, "Error_NotAwcLists", cols = 1:ncol(df_form_not_awc), widths = "auto")
  
  addWorksheet(wb, "Error_ListsWithIssues")
  writeData(wb, "Error_ListsWithIssues", df_problemChecklists)
  setColWidths(wb, "Error_ListsWithIssues", cols = 1:ncol(df_problemChecklists), widths = "auto")
  
  # Save workbook
  saveWorkbook(wb, outfile, overwrite = TRUE)
  
}