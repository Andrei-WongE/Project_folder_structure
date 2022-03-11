## ---------------------------
##
## Script name: Folder_structure
##
## Project: GPE_RF_2025_Indicators
##
## Purpose of script: Create folder structure and templates for GPE RF          ##                     indicators
##
## Author: Andrei Wong Espejo
##
## Date Created: 2022-03-02
##
## Email: awongespejo@worldbank.org
##
## ---------------------------
##
## Notes: Place and run script in folder where you want to create folder        ##         structure and templates
##   
##
## ---------------------------

# Program Set-up ------------

  options(scipen = 100, digits = 2) # Prefer non-scientific notation

# Load required packages ----

  if (!require("pacman")) {
    install.packages("pacman")
  }
  pacman::p_load(tidyverse, data.table, openxlsx, here, purrr)

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
      "Indicator_12i",
      "Indicator_12ii",
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
      "Indicator_5i/Base_2020",
      "Indicator_6/Base_2020",
      "Indicator_7i/Base_2020",
      "Indicator_8i/Base_2020",
      "Indicator_8iiic/Base_2020",
      "Indicator_12i/Base_2020",
      "Indicator_12ii/Base_2020"
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
    "GPE2025_Indicator_12i",
    "GPE2025_Indicator_12ii",
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
    "metadata"
  )

  # Basic template structure

  data_country_db <- data.frame(
    "country" = character(),
    "region" = character(),
    "income_group" = character(),
    "pcfc" = numeric(),
    "iso" = character(),
    "indi_2" = numeric(),
    "indi_2_f" = numeric(),
    "indi_2_m" = numeric(),
    "data_year" = numeric()
  )

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
  worksheetOrder(wb) <- c(4, 1, 2, 3, 5)

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
      ".x0lsx",
      sep = ""
    )
  ))

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
  data <- data.frame(NULL)

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
  worksheetOrder(wb) <- c(4, 5, 1, 2, 3)

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
  # file.remove(here(
  #   folder_structure[i],
  #   paste0(wb_names[i],
  #     ".xlsx",
  #     sep = ""
  #   )
  # ))

  ## GPE2025_indicator_15 ----
  
  rm(wb)
  
  # Load workbook and set parameters
  i <- 20 # Workbook sequence
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
  #rm_sheets <- c(4,6)
  #rm_sheets %>% 
  #  walk(~ removeWorksheet(wb, sheet = .))
  worksheetOrder(wb) <- c(4, 1, 2, 3)
  
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
  i <- 22 # Workbook sequence
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
  worksheetOrder(wb) <- c(4, 1, 2, 3, 5)  
  
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

  
