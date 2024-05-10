if(!require(pacman)){ install.packages("pacman") } else { library(pacman) }

# Import ####
p_load(
  here, 
  dplyr, 
  readxl,
  tibble,
  purrr,
  stringr,
  tidyr,
  lubridate,
  openxlsx,
  janitor,
  data.table
)
options(scipen = 999)
source("function.R")

year <- "2012-2023"
site <- "Barranquilla"
site_id <- "BAQ"

# set up file paths ####
fol_main <- here::here()
fol_data <- file.path(fol_main, "data", site)
fol_remaps <- file.path(fol_main, "data",  "remapping")
path_anl_edd <- here::here(fol_remaps, "Corteva_EDD_Format_Best_Value_Desc.xlsx")
path_fs_edd <- here::here(fol_remaps, "Corteva_FieldSample_EDD_Format_Desc.xlsx")
path_fm_edd <- here::here(fol_remaps, "Corteva_Field_Meas-Added_Tech_EDD_Format_Desc.xlsx")

# read in files ####
lst_esbasic <- list.files(file.path(fol_data, "ESBasic"), pattern = "*.xlsx", full.names = TRUE)
esbasic_in <- lapply(lst_esbasic, function(x){
  df <- read_excel(x, sheet = 2, guess_max = 10000)
}) %>%
  bind_rows()

# set up remappings ####
if(site == "Colon-Argentina"){
  df_cas_remaps <- read_excel(file.path(fol_remaps, "PI_MinisterioDeAmbient_report_remappings.xlsx"),
                              sheet = 2)
  df_meth_remaps <- read_excel(file.path(fol_remaps, "PI_MinisterioDeAmbient_report_remappings.xlsx"),
                               sheet = 3)
  df_unit_remaps <- read_excel(file.path(fol_remaps, "PI_MinisterioDeAmbient_report_remappings.xlsx"),
                               sheet = 4)
} else if(site == "Palotina") {
  df_cas_remaps <- read_excel(file.path(fol_remaps, "2015-2022_remap_guide.xlsx"),
                              sheet = paste0("analyte-", year))
  df_meth_remaps <- read_excel(file.path(fol_remaps, "2015-2022_remap_guide.xlsx"),
                               sheet = paste0("method-", year))
  df_unit_remaps <- read_excel(file.path(fol_remaps, "2015-2022_remap_guide.xlsx"),
                               sheet = "units")
} else {
  df_cas_remaps <- read_excel(file.path(fol_remaps, "master_remap_guide.xlsx"),
                              sheet = 2)
  df_meth_remaps <- read_excel(file.path(fol_remaps, "master_remap_guide.xlsx"),
                               sheet = 3)
  df_unit_remaps <- read_excel(file.path(fol_remaps, "master_remap_guide.xlsx"),
                               sheet = 4)
  df_lab_name_remaps <- read_excel(file.path(fol_remaps, "master_remap_guide.xlsx"),
                                   sheet = 5)
}


anl_meth_maps <- list(df_meth_remaps$EIM_analytic_method) %>%
  unlist() %>%
  setNames(df_meth_remaps$equis_analytic_method)

prep_meth_maps <- list(df_meth_remaps$EIM_prep_method) %>%
  unlist() %>%
  setNames(df_meth_remaps$equis_prep_method)

cas_maps <- list(df_cas_remaps$EIM_parameter_code) %>%
  unlist() %>%
  setNames(df_cas_remaps$equis_cas_rn)

unit_maps <- list(df_unit_remaps$EIM_unit) %>%
  unlist() %>%
  setNames(df_unit_remaps$equis_unit)

lab_name_maps <- list(df_lab_name_remaps$EIM_lab_id) %>%
  unlist() %>%
  setNames(df_lab_name_remaps$equis_lab_name)

names(esbasic_in) <- toupper(names(esbasic_in))
esbasic <- esbasic_in %>%
  filter(
    !str_detect(`#SYS_SAMPLE_CODE`, "MB |LCS |LCSD |LB|MS|MSD"),
    !str_detect(SAMPLE_TYPE_CODE, "SD|MS|LR")
  ) %>%
  mutate(
    `#SYS_SAMPLE_CODE` = str_replace_all(`#SYS_SAMPLE_CODE`, ",", "."),
    SAMPLE_DATE = format(as.Date(SAMPLE_DATE, "%Y/%m/%d"), "%m/%d/%Y"),
    ANALYSIS_DATE = format(as.Date(ANALYSIS_DATE, "%Y/%m/%d"), "%m/%d/%Y"),
    PREP_DATE = format(as.Date(PREP_DATE, "%Y/%m/%d"), "%m/%d/%Y"),
  )

# set up analytical EDD ####
corteva_anl_edd_prep <- esbasic %>%
  mutate(
    SITE_ID = site_id,
    FIELD_SAMPLE_ID = case_when(
      FRACTION == "D" ~ paste0(`#SYS_SAMPLE_CODE`, "-Z"),
      TRUE ~ `#SYS_SAMPLE_CODE`
    ),
    LAB_RESULT = case_when(
      DETECT_FLAG == "Y" ~ as.numeric(RESULT_VALUE),
      DETECT_FLAG == "N" ~ as.numeric(REPORTING_DETECTION_LIMIT)
    ),
    LAB_DETECTION_LIMIT = case_when(
      !is.na(QUANTITATION_LIMIT) ~ as.character(QUANTITATION_LIMIT),
      TRUE ~ as.character(METHOD_DETECTION_LIMIT)
    ),
    LAB_REPORTING_LIMIT_TYPE = case_when(
      as.numeric(REPORTING_DETECTION_LIMIT) == as.numeric(METHOD_DETECTION_LIMIT) ~ "MDL",
      as.numeric(REPORTING_DETECTION_LIMIT) == as.numeric(QUANTITATION_LIMIT) ~ "PQL",
      TRUE ~ "PQL" #AECOM team confirmed that if detection limit is unknown, default to PQL
    ),
    ANALYSIS_TYPE_CODE = case_when(
      TEST_TYPE == "(?i)INITIAL" ~ "INIT",
      TEST_TYPE == "(?i)Reanalysis" ~ "REANL"
    ),
    FILTERED_FLAG = case_when(
      FRACTION == "D" ~ "Y",
      TRUE ~ "N"
    ),
    LEACHED_FLAG = "N",
    SAMPLE_PURPOSE = case_when(
      SAMPLE_TYPE_CODE == "N" ~ "FS",
      SAMPLE_TYPE_CODE == "FD" ~ "DUP",
      TRUE ~ SAMPLE_TYPE_CODE
    ),
    ORIGINAL_LAB_RESULT = LAB_RESULT,
    LAB_MATRIX = case_when(
      grepl("W", SAMPLE_MATRIX_CODE) ~ "LIQUID",
      SAMPLE_MATRIX_CODE == "LO" ~ "LIQUID",
      grepl("S", SAMPLE_MATRIX_CODE) ~ "SOLID",
      grepl("A", SAMPLE_MATRIX_CODE) ~ "AIR"
    ),
    SAMPLE_TIME = format(as.POSIXct(SAMPLE_TIME, format = "%H:%M:%S"), "%H:%M"),
    ANALYSIS_TIME = format(as.POSIXct(ANALYSIS_TIME, format = "%H:%M:%S"), "%H:%M"),
    PREP_TIME = format(as.POSIXct(PREP_TIME, format = "%H:%M:%S"), "%H:%M"),
    method_check = LAB_ANL_METHOD_NAME,
    units_check = RESULT_UNIT,
    temp_anl_meth = LAB_ANL_METHOD_NAME,
    temp_units = RESULT_UNIT,
    temp_lab = LAB_NAME_CODE
  ) %>%
  rename(
    LAB_ID = LAB_NAME_CODE,
    ANALYTICAL_METHOD = LAB_ANL_METHOD_NAME,
    PARAMETER_CODE = CAS_RN,
    LAB_UNITS = RESULT_UNIT,
    prelim_matrix_code = SAMPLE_MATRIX_CODE,
    PREP_METHOD = LAB_PREP_METHOD_NAME,
    PARAMETER_NAME = CHEMICAL_NAME,
    LAB_QUALIFIER = LAB_QUALIFIERS
  )

corteva_anl_edd_prep$ANALYTICAL_METHOD <- anl_meth_maps[corteva_anl_edd_prep$ANALYTICAL_METHOD]
corteva_anl_edd_prep$PREP_METHOD <- prep_meth_maps[corteva_anl_edd_prep$PREP_METHOD]
corteva_anl_edd_prep$PARAMETER_CODE <- cas_maps[corteva_anl_edd_prep$PARAMETER_CODE]
corteva_anl_edd_prep$LAB_UNITS <- unit_maps[corteva_anl_edd_prep$LAB_UNITS]
corteva_anl_edd_prep$LAB_ID <- lab_name_maps[corteva_anl_edd_prep$LAB_ID]

corteva_anl_edd_prep$PARAMETER_CODE[corteva_anl_edd_prep$PARAMETER_NAME == "Picloram"] <- "1918-02-1"

corteva_anl_edd <- corteva_anl_edd_prep %>%
  filter(
    !str_detect(PARAMETER_NAME, "Field"),
    RESULT_TYPE_CODE == "TRG"
  ) %>%
  add_edd_columns(path_anl_edd, "Sheet0")

#batch edds if necessary to fit EIM's 15000-row loading max
if (nrow(corteva_anl_edd) > 14000) {
  max_rows <- 14000
  rows <- nrow(corteva_anl_edd)
  batch_size <- rep(1:ceiling(rows/max_rows), each = max_rows)[1:rows]
  batch_df <- split(corteva_anl_edd, batch_size)
  
  lapply(1:length(batch_df), function(x){
    write.xlsx(batch_df[x], file.path(fol_data, "EIM_EDDs", paste0("pt", x, "-", length(batch_df), "_", site,"_", year, "-Analytical.", site_id, ".Corteva_EDD_Format_Best_Value.xlsx")))
  })
} else {
  write.xlsx(corteva_anl_edd, file.path(fol_data, "EIM_EDDs", paste0("pt1-1_", site,"_", year, "-Analytical.", site_id, ".Corteva_EDD_Format_Best_Value.xlsx")))
}

# create Corteva_FieldSample EDD format to load task codes and sample info ####
#this EDD is considered the COC info EDD and is filled out for ALL results,
#not just field results

corteva_fs_edd_prep <- esbasic %>%
  select(
    `#SYS_SAMPLE_CODE`,
    SYS_LOC_CODE,
    SAMPLE_DATE,
    SAMPLE_TIME,
    SAMPLE_MATRIX_CODE,
    SAMPLE_TYPE_CODE,
    START_DEPTH,
    END_DEPTH,
    DEPTH_UNIT,
    COMMENT,
    TASK_CODE,
    FRACTION,
    SAMPLE_TYPE_CODE
  ) %>%
  rename(
    SAMPLE_START_DEPTH = START_DEPTH,
    SAMPLE_END_DEPTH = END_DEPTH,
    SAMPLE_DEPTH_UNITS = DEPTH_UNIT,
    SAMPLING_PROGRAM = TASK_CODE
  ) %>%
  mutate(
    LOCATION_ID = case_when(
      SAMPLE_TYPE_CODE == "TB" ~ "TRIPBLK",
      SAMPLE_TYPE_CODE == "EB" ~ "EQBLK",
      SAMPLE_TYPE_CODE == "FB" ~ "FLDBLK",
      TRUE ~ SYS_LOC_CODE
    ),
    FIELD_SAMPLE_ID = case_when(
      FRACTION == "D" ~ paste0(`#SYS_SAMPLE_CODE`, "-Z"),
      TRUE ~ `#SYS_SAMPLE_CODE`
    ),
    SAMPLE_TIME = format(as.POSIXct(SAMPLE_TIME, format = "%H:%M:%S"), "%H:%M"),
    SAMPLE_MATRIX = case_when(
      grepl("W", SAMPLE_MATRIX_CODE) ~ "Liquid",
      grepl("S", SAMPLE_MATRIX_CODE) ~ "Solid"
    ),
    SAMPLE_TYPE = case_when(
      grepl("W", SAMPLE_MATRIX_CODE) ~ "Groundwater",
      grepl("SO", SAMPLE_MATRIX_CODE) ~ "Soil",
      grepl("SQ", SAMPLE_MATRIX_CODE) ~ "Soil",
      SAMPLE_MATRIX_CODE == "A" ~ "Air",
      SAMPLE_MATRIX_CODE == "AF" ~ "Sub-Slab",
      SAMPLE_MATRIX_CODE == "LO" ~ "Oil"
    ),
    FIELD_PREPARATION_CODE = case_when(
      FRACTION == "D" ~ "D",
      TRUE ~ "T"
    ),
    SAMPLE_PURPOSE = case_when(
      SAMPLE_TYPE_CODE == "N" ~ "FS",
      SAMPLE_TYPE_CODE == "FD" ~ "DUP",
      TRUE ~ SAMPLE_TYPE_CODE
    ),
  )

corteva_fs_edd <- corteva_fs_edd_prep %>%
  add_edd_columns(path_fs_edd, "Sheet0") %>%
  distinct()

write.xlsx(corteva_fs_edd, file.path(fol_data, "EIM_EDDs", paste0(site,"_", year, ".", site_id, ".Corteva_FieldSample.xlsx")))

# create Corteva FieldMeas-Added Tech EDD format to load field results ####
#(turbidity, etc.). these results do not get loaded in the analytical EDD

corteva_fm_edd_prep <- corteva_anl_edd_prep %>%
  filter(
    str_detect(PARAMETER_NAME, "Field")
  ) %>%
  select(
    SITE_ID,
    TASK_CODE,
    SYS_LOC_CODE,
    FIELD_SAMPLE_ID,
    SAMPLE_DATE,
    SAMPLE_TIME,
    START_DEPTH,
    END_DEPTH,
    DEPTH_UNIT,
    PARAMETER_NAME,
    LAB_RESULT,
    LAB_UNITS,
  ) %>%
  rename(
    SAMPLING_PROGRAM = TASK_CODE,
    FIELD_MEASUREMENT_DATE = SAMPLE_DATE,
    FIELD_MEASUREMENT_TIME = SAMPLE_TIME,
    FIELD_MEASUREMENT_START_DEPTH = START_DEPTH,
    FIELD_MEASUREMENT_END_DEPTH = END_DEPTH,
    FIELD_MEASUREMENT_DEPTH_UNITS = DEPTH_UNIT,
    FIELD_PARAMETER = PARAMETER_NAME,
    FIELD_MEASUREMENT_VALUE = LAB_RESULT,
    FIELD_MEASUREMENT_UNITS = LAB_UNITS
  )

if (nrow(corteva_fm_edd_prep) > 0) {
  corteva_fm_edd <- corteva_fm_edd_prep %>%
    add_edd_columns(path_fm_edd, "Sheet0")
  
  write.xlsx(corteva_fm_edd, file.path(fol_data, "EIM_EDDs", paste0(site,"_", year, ".", site_id, ".Corteva_FieldMeasure.xlsx")))
}







