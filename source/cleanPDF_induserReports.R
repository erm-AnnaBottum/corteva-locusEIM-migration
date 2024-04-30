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

data_year <- "2018"
facility <- "Mercedes"
task_type <- "2023_Barros_Residuos-Induser"
get_refvals <- TRUE

fol_main <- here::here()
fol_data <- file.path(fol_main, "data", facility, "analytical_data/extracted_data", task_type)
files_list <- list.files(fol_data, pattern = ".xlsx")

#handle Induser lab data
df_full <- lapply(files_list, function(x){
  print(x)
  df_in <- read_excel(file.path(fol_data, x),
                      col_names = FALSE)
  
  lab_name <- df_in[[9,2]]
  original_sample_date <- str_replace(df_in[[5,2]], ".*FECHA MUESTREO ", "")
  sample_matrix <- str_replace(df_in[[5,2]], " - .*", "")
  lab_id <- str_replace(df_in[[1,1]], "PROTOCOLO DE ANÁLISIS   ", "")
  sample_id <- str_replace(df_in[[10,1]], "Muestra: ", "")
  us_sample_date <- format(as.Date(original_sample_date, "%d/%m/%Y"), "%m/%d/%Y")
  fmt_date <- format(as.Date(original_sample_date, "%d/%m/%Y"), "%Y%m%d")
  
  # method_heads <- which(grepl("Límite detección", df_in$...3)) #version for 2023 Aguas Residuos
  method_heads <- which(grepl("Limite de Detección", df_in$...3)) #version for 2023 Barros Residuos
  method_ends <- nrow(df_in) - 3
  
  df_results <- df_in[12:method_heads-1, -5] %>%
    row_to_names(1) %>%
    rename(
      Result_Unit = Unidad
    ) %>%
    mutate(
      remark = case_when(
        str_detect(Parámetro, "(//*)") ~ "Starred Analyte"
      ),
      Chemical_Name = str_replace(Parámetro, " (//*)", "")
    ) %>%
    select(-Parámetro)
  
  df_methods <- df_in[method_heads:method_ends, -5] %>%
    row_to_names(1) %>%
    rename(
      detection_limit_unit = Unidad
    )
  
  df_prep <- df_results %>%
    left_join(
      df_methods,
      by = c("Chemical_Name" = "Parámetros")
    )
  
  df_out <- df_prep %>%
    filter(
      !is.na(Chemical_Name),
      !is.na(Result_Unit),
      !is.na(`Valor Obtenido`),
      !is.na(Método),
      !str_detect(Chemical_Name, "Parámetro")
    ) %>%
    mutate(
      sample_date = us_sample_date,
      sample_id = sample_id,
      lab_name = lab_name,
      Result_Value = case_when(
        str_detect(`Valor Obtenido`, "(?i)detectado") ~ NA_real_,
        TRUE ~ as.numeric(str_replace(`Valor Obtenido`, " .*", ""))
      ),
      Detect_Flag = case_when(
        is.na(Result_Value) ~ "N",
        TRUE ~ "Y"
      ),
      # Reporting_Detection_Limit = as.numeric(`Límite detección`), #version for 2023 Aguas Residuos
      Reporting_Detection_Limit = as.numeric(`Limite de Detección`), #version for 2023 Barros Residuos
      # quantification_limit = str_replace(`Límite cuantificación`, " .*", ""), #version for 2023 Aguas Residuos
      quantification_limit = as.numeric(`Limite de Cuantificación`), #version for 2023 Barros Residuos
      Sample_Matrix_Code = case_when(
        str_detect(sample_matrix, "(?i)AGUA SUBTERRANEA") ~ "WG",
        str_detect(sample_matrix, "(?i)BARRO") ~ "SO"
      ),
      Lab_Sample_ID = lab_id,
      lab_sdg = str_replace(lab_id, "Q ", ""),
      Sys_Loc_Code = str_replace(sample_id, " - Q.*", ""),
      Sys_Sample_Code = paste0(Sys_Loc_Code, "-", Sample_Matrix_Code, "-", fmt_date),
      sampling_event_type = task_type
    ) %>%
    rename(
      Lab_Anl_Method_Name = Método
    )
}) %>%
  bind_rows()

write.xlsx(df_full, file.path(fol_data, "output", paste0(facility, "_", task_type, "_eddPrep.xlsx")))

if (get_refvals) {
  analyte <- df_full %>%
    select(
      Chemical_Name
    ) %>%
    distinct()
  
  method <- df_full %>%
    select(
      Lab_Anl_Method_Name
    ) %>%
    distinct()
  
  unit <- df_full %>%
    select(
      Result_Unit
    ) %>%
    distinct()
  
  lst_out <- list("analyte" = analyte, "method" = method, "unit" = unit)
  write.xlsx(lst_out, file.path(fol_data, "output", paste0(facility, "_", task_type, "_RefVals-check.xlsx")))
}

if (get_locations) {
  fol_locs <- file.path(fol_main, "data", facility, "analytical_data/extracted_data")
  list_loc_files <- list.files(fol_locs, pattern = "*eddPrep.xlsx", recursive = TRUE, full.names = TRUE)
  
  df_locs <- lapply(list_loc_files, function(x) {
    df <- read_excel(x) %>%
      select(Sys_Loc_Code) %>%
      distinct()
  }) %>%
    bind_rows()
  
  write.xlsx(df_locs, file.path(fol_data, "locID_mappings.xlsx"))
}
