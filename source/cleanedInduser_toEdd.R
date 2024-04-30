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

data_year <- "2018"
facility <- "Mercedes"
task_type <- "2023_Barros_Residuos-Induser"
get_refvals <- TRUE

fol_main <- here::here()
fol_data <- file.path(fol_main, "data", facility, "analytical_data/extracted_data")
fol_remaps <- file.path(fol_main, "reference_materials", "EQuIS")
fol_edd <- file.path(fol_main, "data")
path_edd <- file.path(fol_edd, "ESBasic_ERM_Desc.xlsx")

if (FALSE){
  files_list <- list.files(fol_data, pattern = "eddPrep.xlsx", recursive = TRUE, full.names = TRUE)
  
  df_full <- lapply(files_list, function(x){
    print(x)
    df_in <- read_excel(x) %>%
      mutate(
        quantification_limit = as.character(quantification_limit)
      )
  }) %>%
    bind_rows()
  
  write.xlsx(df_full, file.path(fol_data, "combined_Mercedes_Data.xlsx"))
}

if (TRUE) {
  df_in <- read_excel(file.path(fol_data, "combined_Mercedes_Data.xlsx"), sheet = 1, guess_max = 10000)
  
  df_edd_prep <- df_in %>%
    mutate(
      fmt_date = format(as.Date(sample_date, "%d/%m/%Y"), "%Y%m%d"),
      Sample_Date = format(as.Date(sample_date, "%d/%m/%Y"), "%m/%d/%Y"),
      Lab_Name_Code = "INDUSER",
      Sys_Sample_Code = case_when(
        str_detect(original_loc_id, "BK|DUP") ~ paste0(original_loc_id, "-", Sample_Matrix_Code, "-", fmt_date),
        TRUE ~ paste0(Sys_Loc_Code, "-", Sample_Matrix_Code, "-", fmt_date)
      ),
      Sample_Type_Code = case_when(
        str_detect(Sys_Sample_Code, "DUP") ~ "FD",
        str_detect(Sys_Sample_Code, "BK") ~ "TB",
        TRUE ~ "N"
      ),
      Analysis_Date = Sample_Date,
      Analysis_Time = "00:00:00",
      Test_Type = "INITIAL",
      Lab_Prep_Method_Name = Lab_Anl_Method_Name,
      Cas_Rn = Chemical_Name,
      Lab_Qualifiers = case_when(
        Detect_Flag == "N" ~ "U"
      ),
      Fraction = Chemical_Name,
      column_number = "NA",
      result_type_code = case_when(
        str_detect(Chemical_Name, "p-Bromofluorobenceno") ~ "SUR",
        TRUE ~ "TRG"
      ),
      reportable_result = "Y",
      validated_yn = "N",
      lab_matrix_code = Sample_Matrix_Code,
      method_detection_limit = Reporting_Detection_Limit,
      Reporting_Detection_Limit = case_when(
        str_detect(result_field, "cuantificable") ~ quantitation_limit, #as.numeric(quantitation_limit),
        !is.na(quantitation_limit) ~ quantitation_limit, #as.numeric(quantitation_limit),
        !is.na(`Limite de Cuantificación`) ~ `Limite de Cuantificación`,
        !is.na(`Límite cuantificación`) ~ `Límite cuantificación`,
        TRUE ~ as.character(Reporting_Detection_Limit)
      ),
      interim_anl_meth = Lab_Prep_Method_Name
    ) %>%
    rename(
      sample_delivery_group = lab_sdg
      )
  
  #do some remapping
  df_cas_remaps <- read_excel(file.path(fol_remaps, "master_remap_guide.xlsx"),
                              sheet = 2)
  df_meth_remaps <- read_excel(file.path(fol_remaps, "master_remap_guide.xlsx"),
                               sheet = 3)
  df_unit_remaps <- read_excel(file.path(fol_remaps, "master_remap_guide.xlsx"),
                               sheet = 4)
  
  anl_meth_maps <- list(df_meth_remaps$equis_analytic_method) %>%
    unlist() %>%
    setNames(df_meth_remaps$source_method_value)
  
  prep_meth_maps <- list(df_meth_remaps$equis_prep_method) %>%
    unlist() %>%
    setNames(df_meth_remaps$source_method_value)
  
  cas_maps <- list(df_cas_remaps$equis_cas_rn) %>%
    unlist() %>%
    setNames(df_cas_remaps$source_chemical_name)
  
  frac_maps <- list(df_cas_remaps$fraction) %>%
    unlist() %>%
    setNames(df_cas_remaps$source_chemical_name)
  
  unit_maps <- list(df_unit_remaps$equis_unit) %>%
    unlist() %>%
    setNames(df_unit_remaps$source_units)

  df_edd_prep$Lab_Anl_Method_Name <- anl_meth_maps[df_edd_prep$Lab_Anl_Method_Name]
  df_edd_prep$Lab_Prep_Method_Name <- prep_meth_maps[df_edd_prep$Lab_Prep_Method_Name]
  df_edd_prep$Cas_Rn <- cas_maps[df_edd_prep$Cas_Rn]
  df_edd_prep$Fraction <- frac_maps[df_edd_prep$Fraction]
  
  df_edd_out <- df_edd_prep %>%
    mutate(
      Prep_Date = case_when(
        !is.na(Lab_Prep_Method_Name) ~ Sample_Date
      ),
      Prep_Time = case_when(
        !is.na(Lab_Prep_Method_Name) ~ "00:00:00"
      ),
      detection_limit_unit = case_when(
        !is.na(method_detection_limit) | !is.na(quantitation_limit) ~ Result_Unit,
        TRUE ~ NA_character_
      )
    ) %>%
    add_edd_columns(path_edd, "ESBasic_ERM") %>%
    rename(`#Sys_Sample_Code` = Sys_Sample_Code)
    
 
  lst_out <- list("ESBasic_ERM" = df_edd_out)
  write.xlsx(lst_out, file.path(fol_data, "Mercedes2023Results.CA_MERCEDES-SITE-AR.ESBasic_ERM.xlsx"))
  
  
}



  