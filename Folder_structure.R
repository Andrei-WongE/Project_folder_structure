## ---------------------------
##
## Script name: Folder_structure
##
## Project: GPE_RF_2025_Indicators
##
## Purpose of script: Create folder structure and templates for GPE RF          
##                     indicators
##
## Author: Andrei Wong Espejo
##
## Date Created: 2022-03-02
##
## Email: awongespejo@worldbank.org
##
## ---------------------------
##
## Notes: Place and run script in folder where you want to create folder        
##         structure and templates.
##   
##
## ---------------------------

# Program Set-up ------------

options(scipen = 100, digits = 2) # Prefer non-scientific notation

# Load required packages ----

if (!require("pacman")) {
  install.packages("pacman")
}
pacman::p_load(tidyverse, data.table, openxlsx, here, datapasta, styler, purrr)

gc() #Free memory

# Runs the following --------
# 0. Set up the folder structure
# 1. Add templates to the folders
# 2. Modify templates according to business requirement doc


# 0. Set up the folder structure ----

folder_structure <- (
  # Main folders
  c(
    "Indicator_1",
    "Indicator_2",
    "Indicator_3i",
    "Indicator_3ii",
    "Indicator_4i",
    "Indicator_4iia_5iia,c_8iia,c_8iiia",
    "Indicator_4iib_5iib_8iib_8iiib",
    "Indicator_5i",
    "Indicator_6",
    "Indicator_7i",
    "Indicator_7ii",
    "Indicator_8i",
    "Indicator_8iiic",
    "Indicator_9i, 10i, 11, 13i",
    "Indicator_9ii, 10ii, 13ii",
    "Indicator_12i_12ii",
    "Indicator_14i",
    "Indicator_14ii",
    "Indicator_15",
    "Indicator_16iii",
    "Indicator_17",
    "Indicator_18i",
    "Indicator_18ii",
    # Sub-folders
    "Indicator_1/Base_2020",
    "Indicator_2/Base_2020",
    "Indicator_3i/Base_2020",
    "Indicator_3ii/Base_2020",
    "Indicator_4i/Base_2020",
    "Indicator_4iia_5iia,c_8iia,c_8iiia/CY_2021",
    "Indicator_5i/Base_2020",
    "Indicator_6/Base_2020",
    "Indicator_7i/Base_2020",
    "Indicator_8i/Base_2020",
    "Indicator_8iiic/CY_2021",
    "Indicator_8iiic/CY_2022",    
    "Indicator_12i_12ii/Base_2020",
    "Indicator_12i_12ii/FY_2021",
    "Indicator_12i_12ii/FY_2022"
  ))

for (j in seq_along(folder_structure)) {
  dir.create(folder_structure[j])
}

# 1. Add files to the folders ----

# File names

wb_names <- c(
  "GPE2025_Indicator_1",
  "GPE2025_Indicator_2",
  "GPE2025_Indicator_3i",
  "GPE2025_Indicator_3ii",
  "GPE2025_Indicator_4i",
  "GPE2025_Indicator_4iia_5iia,c_8iia,c_8iiia",
  "GPE2025_Indicator_4iib_5iib_8iib_8iiib",
  "GPE2025_Indicator_5i",
  "GPE2025_Indicator_6",
  "GPE2025_Indicator_7i",
  "GPE2025_Indicator_7ii",
  "GPE2025_Indicator_8i",
  "GPE2025_Indicator_8iiic",
  "GPE2025_Indicator_9i, 10i, 11, 13i",
  "GPE2025_Indicator_9ii, 10ii, 13ii",
  "GPE2025_Indicator_12i_12ii",
  "GPE2025_Indicator_14i",
  "GPE2025_Indicator_14ii",
  "GPE2025_Indicator_15",
  "GPE2025_Indicator_16iii",
  "GPE2025_Indicator_17",
  "GPE2025_Indicator_18i",
  "GPE2025_Indicator_18ii"
)

# Sheet names

sheet_names <- c(
  "data_country",
  "data_aggregate",
  "metadata",
  "admin_data",
  "raw_data"
)

# Basic template structure

data_country_db <- data.frame(
  "country" = rep("", times = 76),
  "entity"  = rep("", times = 76),
  "region"  = sprintf("VLOOKUP(A%s, admin_data!$A$2:$E$77,2,TRUE)", 2:77),
  "income_group" = sprintf("VLOOKUP(A%s,admin_data!$A$2:$E$77,3,TRUE)", 2:77),
  "pcfc" = sprintf("VLOOKUP(A%s, admin_data!$A$2:$E$77,4,TRUE)", 2:77),
  "iso"  = sprintf("VLOOKUP(A%s, admin_data!$A$2:$E$77,5,TRUE)", 2:77),
  "indi_2"    = rep("", times = 76),
  "indi_2_f"  = rep("", times = 76),
  "indi_2_m"  = rep("", times = 76),
  "data_year" = rep("", times = 76)
)

# Declare it is a formula
class(data_country_db$region) <- c(class(data_country_db$region), "formula")
class(data_country_db$income_group) <- c(class(data_country_db$income_group), "formula")
class(data_country_db$pcfc) <- c(class(data_country_db$pcfc), "formula")
class(data_country_db$iso) <- c(class(data_country_db$iso), "formula")

data_aggregate_db <- data.frame(
  "type" = character(),
  "value" = character(), "n" = numeric(),
  "measurement" = character()
)

metadata_db <- data.frame(
  "var_name" = character(),
  "var_label" = character(),
  "var_level" = character()
)

admin_data_db <- data.frame(
  stringsAsFactors = FALSE,
  country = c(
    "Afghanistan", "Albania",
    "Bangladesh", "Benin", "Bhutan", "Burkina Faso",
    "Burundi", "Cabo Verde", "Cambodia",
    "Cameroon", "Central African Republic",
    "Chad", "Comoros",
    "Congo,Democratic Republic of", "Congo,Republic of",
    "Coted'Ivoire", "Djibouti", "Dominica", "Eritrea",
    "Ethiopia", "Gambia, The", "Georgia",
    "Ghana", "Grenada", "Guinea",
    "Guinea-Bissau", "Guyana", "Haiti", "Honduras",
    "Kenya", "Kiribati", "Kyrgyz Republic",
    "LaoPeople's Democratic Republic",
    "Lesotho", "Liberia", "Madagascar", "Malawi",
    "Maldives", "Mali", "Marshall Islands",
    "Mauritania",
    "Micronesia, Federated States of", "Moldova", "Mongolia",
    "Mozambique", "Myanmar", "Nepal", "Nicaragua",
    "Niger", "Nigeria", "Pakistan",
    "Papua New Guinea", "Rwanda", "Samoa",
    "SaoTome and Principe", "Senegal", "Sierra Leone",
    "Solomon Islands", "Somalia",
    "South Sudan", "St.Lucia",
    "St.Vincent and the Grenadines", "Sudan", "Tajikistan",
    "Tanzania", "Timor-Leste", "Togo", "Tonga",
    "Tuvalu", "Uganda", "Uzbekistan",
    "Vanuatu", "Vietnam", "Yemen, Republic of",
    "Zambia", "Zimbabwe"
  ),
  region = c(
    "South Asia", "Europe & Central Asia",
    "South Asia", "Sub-Saharan Africa",
    "South Asia", "Sub-Saharan Africa",
    "Sub-Saharan Africa", "Sub-Saharan Africa",
    "East Asia & Pacific",
    "Sub-Saharan Africa", "Sub-Saharan Africa",
    "Sub-Saharan Africa", "Sub-Saharan Africa",
    "Sub-Saharan Africa", "Sub-Saharan Africa",
    "Sub-Saharan Africa",
    "Middle East & North Africa", "Latin America & Caribbean",
    "Sub-Saharan Africa", "Sub-Saharan Africa",
    "Sub-Saharan Africa",
    "Europe & Central Asia", "Sub-Saharan Africa",
    "Latin America & Caribbean", "Sub-Saharan Africa",
    "Sub-Saharan Africa",
    "Latin America & Caribbean", "Latin America & Caribbean",
    "Latin America & Caribbean",
    "Sub-Saharan Africa", "East Asia & Pacific",
    "Europe & Central Asia", "East Asia & Pacific",
    "Sub-Saharan Africa",
    "Sub-Saharan Africa", "Sub-Saharan Africa",
    "Sub-Saharan Africa", "South Asia",
    "Sub-Saharan Africa", "East Asia & Pacific",
    "Sub-Saharan Africa", "East Asia & Pacific",
    "Europe & Central Asia", "East Asia & Pacific",
    "Sub-Saharan Africa",
    "East Asia & Pacific", "South Asia",
    "Latin America & Caribbean", "Sub-Saharan Africa",
    "Sub-Saharan Africa", "South Asia",
    "East Asia & Pacific", "Sub-Saharan Africa",
    "East Asia & Pacific", "Sub-Saharan Africa",
    "Sub-Saharan Africa", "Sub-Saharan Africa",
    "East Asia & Pacific",
    "Sub-Saharan Africa", "Sub-Saharan Africa",
    "Latin America & Caribbean",
    "Latin America & Caribbean", "Sub-Saharan Africa",
    "Europe & Central Asia", "Sub-Saharan Africa",
    "East Asia & Pacific", "Sub-Saharan Africa",
    "East Asia & Pacific",
    "East Asia & Pacific", "Sub-Saharan Africa",
    "Europe & Central Asia", "East Asia & Pacific",
    "East Asia & Pacific",
    "Middle East & North Africa", "Sub-Saharan Africa",
    "Sub-Saharan Africa"
  ),
  income_group = c(
    "Low income", "Upper middle income",
    "Lower middle income",
    "Lower middle income", "Lower middle income",
    "Low income", "Low income", "Lower middle income",
    "Lower middle income",
    "Lower middle income", "Low income", "Low income",
    "Lower middle income", "Low income",
    "Lower middle income", "Lower middle income",
    "Lower middle income",
    "Upper middle income", "Low income", "Low income",
    "Low income", "Upper middle income",
    "Lower middle income", "Upper middle income",
    "Low income", "Low income",
    "Upper middle income", "Lower middle income",
    "Lower middle income", "Lower middle income",
    "Lower middle income", "Lower middle income",
    "Lower middle income",
    "Lower middle income", "Low income", "Low income",
    "Low income", "Upper middle income",
    "Low income", "Upper middle income",
    "Lower middle income", "Lower middle income",
    "Upper middle income", "Lower middle income",
    "Low income", "Lower middle income",
    "Lower middle income",
    "Lower middle income", "Low income", "Lower middle income",
    "Lower middle income",
    "Lower middle income", "Low income",
    "Lower middle income", "Lower middle income",
    "Lower middle income", "Low income",
    "Lower middle income", "Low income", "Low income",
    "Upper middle income", "Upper middle income",
    "Low income", "Lower middle income",
    "Lower middle income",
    "Lower middle income", "Low income", "Upper middle income",
    "Upper middle income", "Low income",
    "Lower middle income",
    "Lower middle income", "Lower middle income", "Low income",
    "Lower middle income",
    "Lower middle income"
  ),
  pcfc = c(
    1L, 0L, 0L, 0L, 0L, 1L, 1L, 0L, 0L,
    1L, 1L, 1L, 1L, 1L, 1L, 0L, 0L, 0L, 1L,
    0L, 1L, 0L, 0L, 0L, 0L, 1L, 0L, 1L, 0L,
    1L, 1L, 0L, 0L, 0L, 1L, 0L, 0L, 0L, 1L,
    1L, 0L, 1L, 0L, 0L, 0L, 1L, 0L, 0L, 1L,
    1L, 1L, 1L, 1L, 0L, 0L, 0L, 0L, 1L, 1L,
    1L, 0L, 0L, 1L, 0L, 0L, 1L, 0L, 0L, 1L,
    1L, 0L, 0L, 0L, 1L, 0L, 1L
  ),
  iso = c(
    "AFG", "ALB", "BGD", "BEN", "BTN",
    "BFA", "BDI", "CPV", "KHM", "CMR", "CAF",
    "TCD", "COM", "COD", "COG", "CIV", "DJI",
    "DMA", "ERI", "ETH", "GMB", "GEO",
    "GHA", "GRD", "GIN", "GNB", "GUY", "HTI",
    "HND", "KEN", "KIR", "KGZ", "LAO", "LSO",
    "LBR", "MDG", "MWI", "MDV", "MLI",
    "MHL", "MRT", "FSM", "MDA", "MNG", "MOZ",
    "MMR", "NPL", "NIC", "NER", "NGA", "PAK",
    "PNG", "RWA", "WSM", "STP", "SEN", "SLE",
    "SLB", "SOM", "SSD", "LCA", "VCT",
    "SDN", "TJK", "TZA", "TLS", "TGO", "TON",
    "TUV", "UGA", "UZB", "VUT", "VNM", "YEM",
    "ZMB", "ZWE"
  )
)


raw_data_db <- data.frame(NULL
)


# Sheet header style

hs1 <- createStyle(
  halign = "CENTER",
  valign = "CENTER",
  border = "TopBottomLeftRight",
  borderColour = "black"
  # wrapText = TRUE
)

## Create Workbooks and add worksheets

lapply(seq_along(wb_names), function(j) {
  
  wb <- createWorkbook()
  
  lapply(seq_along(sheet_names), function(i) {
    
    # Add sheets
    
    addWorksheet(wb = wb, sheetName = sheet_names[i])
    data <- get(paste0(sheet_names[i], "_db", sep = ""))
    
    # Write all the create data
    writeData(wb,
              sheet = sheet_names[i], data,
              colNames = TRUE, xy = c(1, 1),
              headerStyle = hs1
    ) ## black border + bold header
  })
  
  # Save Workbook
  
  saveWorkbook(wb,
               file = here(
                 folder_structure[j],
                 paste0(wb_names[j],
                        ".xlsx",
                        sep = ""
                 )
               ),
               overwrite = TRUE
  )
})

# 2. Modify templates according to business requirement doc ----

## GPE2025_indicator_1 ----

# rm(wb)

# Load workbook and set parameters
i <- 1 # Workbook sequence
j <- 6 # Start Column of new data

wb <- loadWorkbook(here(
  folder_structure[i],
  paste0(wb_names[i],
         ".xlsx",
         sep = ""
  )
), )

# Add new columns

data <- data.frame(
  "indi_1_yr_comp" = numeric(),
  "ind_1_yr_free" = numeric(),
  "indi_1" = numeric(),
  "data_year" = numeric()
)

writeData(wb,
          sheet = sheet_names[1], data,
          colNames = TRUE, xy = c(j, 1),
          headerStyle = hs1
) ## black border + bold header

# Save workbook with new columns

saveWorkbook(wb,
             file = here(
               folder_structure[i],
               paste0(wb_names[i],
                      ".xlsx",
                      sep = ""
               )
             ),
             overwrite = TRUE
)
rm(data)

## GPE2025_indicator_3i ----

# rm(wb)

# Load workbook and set parameters
i <- 3 # Workbook sequence
j <- 6 # Start Column of new data

wb <- loadWorkbook(here(
  folder_structure[i],
  paste0(wb_names[i],
         ".xlsx",
         sep = ""
  )
), )

# Add new columns

data <- data.frame(
     "indi_3ia" = numeric(),
     "indi_3ia_f" = numeric(),
     "indi_3ia_m" = numeric(),
     "indi_3ib" = numeric(),
     "indi_3ib_f" = numeric(),
     "indi_3ib_m" = numeric(),
     "data_year_3ia" = numeric(),
     "data_year_3ib" = numeric()
      )

writeData(wb,
          sheet = sheet_names[1], data,
          colNames = TRUE, xy = c(j, 1),
          headerStyle = hs1
) ## black border + bold header

# Save workbook with new columns

saveWorkbook(wb,
             file = here(
               folder_structure[i],
               paste0(wb_names[i],
                      ".xlsx",
                      sep = ""
               )
             ),
             overwrite = TRUE
)
rm(data)

## GPE2025_indicator_3ii ----

# rm(wb)

# Load workbook and set parameters
i <- 4 # Workbook sequence
j <- 6 # Start Column of new data

wb <- loadWorkbook(here(
  folder_structure[i],
  paste0(wb_names[i],
         ".xlsx",
         sep = ""
  )
), )

# Add new columns

data <- data.frame(
  "indi_3iia" = numeric(),
  "indi_3iia_f" = numeric(),
  "indi_3iia_m" = numeric(),
  "indi_3iia_rural" = numeric(),
  "indi_3iia_poor" = numeric(),
  "indi_3iib" = numeric(),
  "indi_3iib_f" = numeric(),
  "indi_3iib_m" = numeric(),
  "indi_3iib_rural" = numeric(),
  "indi_3iib_poor" = numeric(),
  "indi_3iic" = numeric(),
  "indi_3iic_f" = numeric(),
  "indi_3iic_m" = numeric(),
  "indi_3iic_rural" = numeric(),
  "indi_3iic_poor" = numeric(),
  "data_year" = numeric()
)

writeData(wb,
          sheet = sheet_names[1], data,
          colNames = TRUE, xy = c(j, 1),
          headerStyle = hs1
) ## black border + bold header

# Save workbook with new columns

saveWorkbook(wb,
             file = here(
               folder_structure[i],
               paste0(wb_names[i],
                      ".xlsx",
                      sep = ""
               )
             ),
             overwrite = TRUE
)
rm(data)

## GPE2025_indicator_4i ----

# rm(wb)

# Load workbook and set parameters
i <- 5 # Workbook sequence
j <- 6 # Start Column of new data

wb <- loadWorkbook(here(
  folder_structure[i],
  paste0(wb_names[i],
         ".xlsx",
         sep = ""
  )
), )

# Add new columns

data <- data.frame(
  "Indi_4i_base" = numeric(),
  "Indi_4i_current" = numeric(),
  "Indi_4i_a" = numeric(),
  "Indi_4i_b" = numeric(),
  "Indi_4i" = numeric()
)

writeData(wb,
          sheet = sheet_names[1], data,
          colNames = TRUE, xy = c(j, 1),
          headerStyle = hs1
) ## black border + bold header

# Save workbook with new columns

saveWorkbook(wb,
             file = here(
               folder_structure[i],
               paste0(wb_names[i],
                      ".xlsx",
                      sep = ""
               )
             ),
             overwrite = TRUE
)
rm(data)

## GPE2025_Indicator_4iia_5iia,c_8iia,c_8iiia ----

rm(wb)

# Load workbook and set parameters
i <- 6 # Workbook sequence
j <- 6 # Start Column of new data

wb <- loadWorkbook(here(
  folder_structure[i],
  paste0(wb_names[i],
         ".xlsx",
         sep = ""
  )
), )

# Add new columns

data <- data.frame(
  "entity" = numeric(),
  "indi_4iia" = numeric(),
  "indi_4iia_priority" = numeric(),
  "indi_5iia" = numeric(),
  "indi_5iia_priority" = numeric(),
  "indi_5iic" = numeric(),
  "indi_8iia" = numeric(),
  "indi_8iia_priority" = numeric(),
  "indi_8iic" = numeric(),
  "indi_8iiia" = numeric(),
  "indi_8iiia_priority" = numeric()
)


writeData(wb,
          sheet = sheet_names[1], data,
          colNames = TRUE, xy = c(j, 1),
          headerStyle = hs1
) ## black border + bold header

# Add new sheets and columns

addWorksheet(wb = wb, sheetName = "Tool_ITAP")
data_1 <- data.frame(NULL)

writeData(wb,
          sheet = "Tool_ITAP", data_1,
          colNames = TRUE, xy = c(1, 1),
          headerStyle = hs1
) ## black border + bold header


addWorksheet(wb = wb, sheetName = "ref_iso")
data_4 <- data.frame(
  "country" = character(),
  "Iso" = character()
)

writeData(wb,
          sheet = "ref_iso", data_4,
          colNames = TRUE, xy = c(1, 1),
          headerStyle = hs1
) ## black border + bold header

# Worksheet order

worksheetOrder(wb)
names(wb)
worksheetOrder(wb) <- c(6, 1, 2, 3, 7, 4, 5)

# Save workbook with new columns and sheets, and template name

saveWorkbook(wb,
             file = here(
               folder_structure[i],
               paste0("ITAP tool for RF indicator ", wb_names[i],
                      ".xlsx",
                      sep = ""
               )
             ),
             overwrite = TRUE
)

rm(data, data_1, data_4)
file.remove(here(
  folder_structure[i],
  paste0(wb_names[i],
         ".xlsx",
         sep = ""
  )
))

## GPE2025_indicator_5i ----

# rm(wb)

# Load workbook and set parameters
i <- 8 # Workbook sequence
j <- 6 # Start Column of new data

wb <- loadWorkbook(here(
  folder_structure[i],
  paste0(wb_names[i],
         ".xlsx",
         sep = ""
  )
), )

# Add new columns

data <- data.frame(
     "indi_5i" = numeric(),
     "indi_5i_pop" = numeric(),
     "data_year" = numeric()
)

writeData(wb,
          sheet = sheet_names[1], data,
          colNames = TRUE, xy = c(j, 1),
          headerStyle = hs1
) ## black border + bold header

# Save workbook with new columns

saveWorkbook(wb,
             file = here(
               folder_structure[i],
               paste0(wb_names[i],
                      ".xlsx",
                      sep = ""
               )
             ),
             overwrite = TRUE
)
rm(data)


## GPE2025_indicator_6 ----

# rm(wb)

# Load workbook and set parameters
i <- 9 # Workbook sequence
j <- 6 # Start Column of new data

wb <- loadWorkbook(here(
  folder_structure[i],
  paste0(wb_names[i],
         ".xlsx",
         sep = ""
  )
), )

# Add new columns

data <- data.frame(
  "indi_6a1" = numeric(),
  "indi_6a1_f" = numeric(),
  "indi_6a1_m" = numeric(),
  "indi_6a2" = numeric(),
  "indi_6a2_f" = numeric(),
  "indi_6a2_m" = numeric(),
  "indi_6b1" = numeric(),
  "indi_6b1_f" = numeric(),
  "indi_6b1_m" = numeric(),
  "indi_6b2" = numeric(),
  "indi_6b2_f" = numeric(),
  "indi_6b2_m" = numeric(),
  "indi_6c1" = numeric(),
  "indi_6c1_f" = numeric(),
  "indi_6c1_m" = numeric(),
  "indi_6c2" = numeric(),
  "indi_6c2_f" = numeric(),
  "indi_6c2_m" = numeric(),
  "data_year_6a1" = numeric(),
  "data_year_6a2" = numeric(),
  "data_year_6b1" = numeric(),
  "data_year_6b2" = numeric(),
  "data_year_6c1" = numeric(),
  "data_year_6c2" = numeric()
)

writeData(wb,
          sheet = sheet_names[1], data,
          colNames = TRUE, xy = c(j, 1),
          headerStyle = hs1
) ## black border + bold header

# Save workbook with new columns

saveWorkbook(wb,
             file = here(
               folder_structure[i],
               paste0(wb_names[i],
                      ".xlsx",
                      sep = ""
               )
             ),
             overwrite = TRUE
)
rm(data)

## GPE2025_indicator_7i ----

rm(wb)

# Load workbook and set parameters
i <- 10 # Workbook sequence
j <- 6 # Start Column of new data

wb <- loadWorkbook(here(
  folder_structure[i],
  paste0(wb_names[i],
         ".xlsx",
         sep = ""
  )
), )

# Add new columns

data <- data.frame(
  "indi_7ia" = numeric(),
  "indi_7ia_f" = numeric(),
  "indi_7ia_m" = numeric(),
  "indi_7ib" = numeric(),
  "indi_7ib_f" = numeric(),
  "indi_7ib_m" = numeric(),
  "indi_7ic" = numeric(),
  "indi_7ic_f" = numeric(),
  "indi_7ic_m" = numeric(),
  "indi_7id" = numeric(),
  "indi_7id_f" = numeric(),
  "indi_7id_m" = numeric(),
  "data_year_7ia" = numeric(),
  "data_year_7ib" = numeric(),
  "data_year_7ic" = numeric(),
  "data_year_7id" = numeric()
)


writeData(wb,
          sheet = sheet_names[1], data,
          colNames = TRUE, xy = c(j, 1),
          headerStyle = hs1
) ## black border + bold header

# Save workbook with new columns

saveWorkbook(wb,
             file = here(
               folder_structure[i],
               paste0(wb_names[i],
                      ".xlsx",
                      sep = ""
               )
             ),
             overwrite = TRUE
)
rm(data)

## GPE2025_indicator_8i ----

rm(wb)

# Load workbook and set parameters
i <- 12 # Workbook sequence
j <- 6 # Start Column of new data

wb <- loadWorkbook(here(
  folder_structure[i],
  paste0(wb_names[i],
         ".xlsx",
         sep = ""
  )
), )

# Add new columns

data <- data.frame(
  "indi_8i" = numeric(),
  "indi_8i_n" = numeric(),
  "indi_8i_outcome_1" = numeric(),
  "indi_8i_outcome_2" = numeric(),
  "indi_8i_outcome_3" = numeric(),
  "indi_8i_outcome_4" = numeric(),
  "indi_8i_outcome_5" = numeric(),
  "indi_8i_service_6" = numeric(),
  "indi_8i_service_7" = numeric(),
  "indi_8i_service_8" = numeric(),
  "indi_8i_service_9" = numeric(),
  "indi_8i_financing_10" = numeric(),
  "indi_8i_financing_11" = numeric(),
  "indi_8i_financing_12" = numeric()
)


writeData(wb,
          sheet = sheet_names[1], data,
          colNames = TRUE, xy = c(j, 1),
          headerStyle = hs1
) ## black border + bold header

# Save workbook with new columns

saveWorkbook(wb,
             file = here(
               folder_structure[i],
               paste0(wb_names[i],
                      ".xlsx",
                      sep = ""
               )
             ),
             overwrite = TRUE
)

rm(data)

## GPE2025_Indicator_8iiic ----

rm(wb)

# Load workbook and set parameters
i <- 13 # Workbook sequence
j <- 1 # Start Column of new data

wb <- loadWorkbook(here(
  folder_structure[i],
  paste0(wb_names[i],
         ".xlsx",
         sep = ""
  )
), )

# Add new sheets and columns

addWorksheet(wb = wb, sheetName = "Datasheet")
data <- data.frame(NULL) ##!!!!!! Need to include variable names!!!!

writeData(wb,
          sheet = "Datasheet", data,
          colNames = TRUE, xy = c(j, 1),
          headerStyle = hs1
) ## black border + bold header

addWorksheet(wb = wb, sheetName = "data_leg")
data_1 <- data.frame(
  "country" = character(),
  "entity" = character(),
  "iso" = character(),
  "region" = character(),
  "income_group" = character(),
  "pcfc" = numeric(),
  "indi_8iiic_cso" = numeric(),
  "indi_8iiic_ta" = numeric(),
  "indi_8iiic" = numeric()
)

writeData(wb,
          sheet = "data_leg", data_1,
          colNames = TRUE, xy = c(j, 1),
          headerStyle = hs1
) ## black border + bold header

# Worksheet order

worksheetOrder(wb)
names(wb)
removeWorksheet(wb, sheet = 1)
names(wb)
worksheetOrder(wb) <- c(5, 6, 1, 2, 3, 4)

# Save workbook with new columns and sheets, and template name

saveWorkbook(wb,
             file = here(
               folder_structure[i],
               paste0("Local Education Group-2020",
                      ".xlsx",
                      sep = ""
               )
             ),
             overwrite = TRUE
)

rm(data, data_1)
file.remove(here(
  folder_structure[i],
  paste0(wb_names[i],
         ".xlsx",
         sep = ""
  )
))

## GPE2025_Indicator_12i_12ii ----

rm(wb)

# Load workbook and set parameters
i <- 16 # Workbook sequence
j <- 6 # Start Column of new data

wb <- loadWorkbook(here(
  folder_structure[i],
  paste0(wb_names[i],
         ".xlsx",
         sep = ""
  )
), )

# Add new columns

data <- data.frame(
  "grant_id" = character(),
  "grant_amount" = numeric(),
  "ind_12i_education_service_plan" = numeric(),
  "ind_12i_medium_term_expenditure_framework" = numeric(),
  "ind_12i_national_budget_information" = numeric(),
  "ind_12i_specific_budget_appropriations" = numeric(),
  "ind_12i_treasury" = numeric(),
  "ind_12i_pfm_expenditure_process" = numeric(),
  "ind_12i_procurement_rules" = numeric(),
  "ind_12i_accounting_system" = numeric(),
  "ind_12i_national_external_audit" = numeric(),
  "ind_12i_esp_annual_implementation_report" = numeric(),
  "ind_12i_aligmend_score" = numeric(),
  "ind_12i_aligned" = numeric(),
  "ind_12ii_modality_label" = character(),
  "ind_12ii_modality" = numeric(),
  "data_year" = numeric()
)

writeData(wb,
          sheet = sheet_names[1], data,
          colNames = TRUE, xy = c(j, 1),
          headerStyle = hs1
) ## black border + bold header

# Rename worksheet
renameWorksheet(wb, 1, "data_country_grant")

# Worksheet order

worksheetOrder(wb)
names(wb)

# Save workbook with new columns, and sheetname

saveWorkbook(wb,
             file = here(
               folder_structure[i],
               paste0(wb_names[i],
                      ".xlsx",
                      sep = ""
               )
             ),
             overwrite = TRUE
)

rm(data) 

## GPE2025_indicator_15 ----

rm(wb)

# Load workbook and set parameters
i <- 19 # Workbook sequence
j <- 6 # Start Column of new data

wb <- loadWorkbook(here(
  folder_structure[i],
  paste0(wb_names[i],
         ".xlsx",
         sep = ""
  )
), ) 

# Add new sheets and columns

cloneWorksheet(wb, "data_case" , clonedSheet = sheet_names[1])
data <- data.frame("case_number" = character(),
                   "gesi" = character(),
                   "gesi_related" = numeric(),
                   "kix_indicator" = character(),
                   "kix_activity" = character(),
                   "kix_region" = character()
)

writeData(wb,
          sheet = "data_case", data,
          colNames = TRUE, xy = c(j, 1),
          headerStyle = hs1
) ## black border + bold header

data_1 <- data.frame("indi_15" = numeric(),
                     "indi_15_gesi" = numeric()
)

writeData(wb,
          sheet = sheet_names[1], data_1,
          colNames = TRUE, xy = c(j, 1),
          headerStyle = hs1
) ## black border + bold header

# Worksheet order and deletion

worksheetOrder(wb)
names(wb)
#rm_sheets %>% 
#  walk(~ removeWorksheet(wb, sheet = .))
names(wb)
worksheetOrder(wb) <- c(6, 1, 2, 3, 4, 5)

# Save workbook with new columns and sheets, and template name

saveWorkbook(wb,
             file = here(
               folder_structure[i],
               paste0(wb_names[i],
                      ".xlsx",
                      sep = ""
               )
             ),
             overwrite = TRUE
)

rm(data, data_1) 


## GPE2025_indicator_17 ----

rm(wb)

# Load workbook and set parameters
i <- 21 # Workbook sequence
j <- 6 # Start Column of new data

wb <- loadWorkbook(here(
  folder_structure[i],
  paste0(wb_names[i],
         ".xlsx",
         sep = ""
  )
), )

# Add new sheets and columns

addWorksheet(wb = wb, sheetName = "data_policy")
data <- data.frame(
  "country" = character(),
  "region" = character(),
  "income_group" = character(),
  "pcfc" = numeric(),
  "iso" = character(),
  "policy_change_name" = character(),
  "TPR_reporting_period" = character(),
  "policy_change_approval_date" = numeric(),
  "policy_change_info" = character(),
  "policy_change_CSO_importance" = character(),
  "CSO_participation" = character(),
  "CSO_participation_detail" = character()
)

writeData(wb,
          sheet = "data_policy", data,
          colNames = TRUE, xy = c(1, 1),
          headerStyle = hs1
) ## black border + bold header

data_1 <- data.frame(
  "indi_17_count" = numeric(),
  "indi_17"       = numeric()
)

writeData(wb,
          sheet = sheet_names[1], data_1,
          colNames = TRUE, xy = c(j, 1),
          headerStyle = hs1
) ## black border + bold header

addWorksheet(wb = wb, sheetName = "list_eol")
data_2 <- data.frame("country" = character(),
                     "region" = character(),
                     "income_group" = character(),
                     "pcfc" = numeric(), 
                     "iso" = character()         
)

writeData(wb,
          sheet = "list_eol", data_2,
          colNames = TRUE, xy = c(1, 1),
          headerStyle = hs1
) ## black border + bold header

# Save workbook with new columns

saveWorkbook(wb,
             file = here(
               folder_structure[i],
               paste0(wb_names[i],
                      ".xlsx",
                      sep = ""
               )
             ),
             overwrite = TRUE
)

# Worksheet order

worksheetOrder(wb)
names(wb)
worksheetOrder(wb) <- c(6, 1, 2, 3, 7, 4, 5)  

# Save workbook with new columns and sheets, and template name

saveWorkbook(wb,
             file = here(
               folder_structure[i],
               paste0(wb_names[i],
                      ".xlsx",
                      sep = ""
               )
             ),
             overwrite = TRUE
)


rm(data, data_1, data_2)


# Create a Readme file
#file.create("README.text")




