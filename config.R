############################Configurations for scripts to merge eBird data from AWC and AWC forms#####################
googleform <- "AWC India-ebird Wetland Assessment Form 2022 (Responses) - downloaded on 15April2025 (2).xlsx"
myEbdFile <- "MyEBirdData.csv"
startYear <- 2021
endYear <- 2022

# For this version, endYear should be startYear + 1
if ( (endYear - startYear) != 1) 
{
  stop("This version does not support multiple AWC seasons")
}

inputdir <- ".\\inputs\\"
outputdir <- ".\\outputs\\"
downloadMonth <- "May"
downloadYear <- "2025"
awc_1st_half_months <- c (1, 2) # January & February. Add 3 if you want to include March
awc_2nd_half_months <- c (12) # December. Add 11 if you want to include November
awc_months <- c(awc_1st_half_months, awc_2nd_half_months)

grey_fill <- createStyle(fgFill = "#D9D9D9")  # light grey