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
  data.table,
  numbers
)
options(scipen = 999)

#notes ####
#fill out the following fields for each individual dataset
facility = "Mercedes" #which site is this data for
task_type = "2023_Agua_OPDS" #if no specific task type, put 'NA' and the standard default will be generated
lab = "UNK" #TAPEN = Test America Pensacola
#filename = "400-212739-1_Results" #filename to be read
#method_lookupfile = "400-212739-1_LabChronicle" #lab chronicle section extracted for lab report
write_edd = TRUE #do you want to write an EDD draft containing all the information provided so far?
get_refvals = FALSE #do you want to get refvals summary from data
source("function.R")

fol_main <- here::here()
fol_edd <- file.path(fol_main, "data")
fol_data <- file.path(fol_main, "data", facility, "analytical_data/extracted_data", task_type)
fol_data_b <- file.path(fol_main, "data", facility, "analytical_data/extracted_data", task_type, "B_reports")
files_list_a_rpts <- list.files(fol_data, pattern = ".xlsx")
files_list_b_rpts <- list.files(file.path(fol_data, "B_reports"), pattern = ".xlsx")

# set up data ####
df_full_a_rpts <- lapply(files_list, function(x){
  df_in <- read_excel(file.path(fol_data, x),
                      col_names = FALSE)
  
  lab_name <- df_in[[4,5]]
  sample_date <- df_in[[7,5]]
  lab_sdg <- str_replace(x, " .*", "")
  sample_id <- df_in[[20,4]]
  
  df_prep <- df_in[,-3]
  names(df_prep) <- c("Chemical_Name", "result_field", "Lab_Anl_Method_Name", "detection_limit_field", "quantification_limit_field")
  
  df_out <- df_prep %>%
    tail(-24) %>%
    filter(
      !str_detect(Chemical_Name, "(?i)Página|ANEXO|Nombre|Firma y Sello "),
      !is.na(Chemical_Name),
      !is.na(result_field)
    ) %>%
    mutate(
      sample_date = sample_date,
      sample_id = sample_id,
      lab_sdg = lab_sdg,
      lab_name = lab_name,
      Result_Value = case_when(
        str_detect(result_field, "(?i)detectado") ~ NA_real_,
        TRUE ~ as.numeric(str_replace(result_field, " .*", ""))
      ),
      Result_Unit = case_when(
        str_detect(result_field, "(?i)detectado") ~ NA_character_,
        TRUE ~ str_replace(result_field, ".* ", "")
      ),
      Detect_Flag = case_when(
        is.na(Result_Value) ~ "N",
        TRUE ~ "Y"
      ),
      Reporting_Detection_Limit = as.numeric(str_replace(detection_limit_field, " .*", "")),
      detection_limit_unit = str_replace(detection_limit_field, ".* ", ""),
      quantification_limit = str_replace(quantification_limit_field, " .*", ""),
      Sample_Matrix_Code = case_when(
        str_detect(detection_limit_field, "(?i)g\\/l") ~ "WG"
      ),
      Lab_Sample_ID = paste0("Q", str_replace(sample_id, ".* Q", "")),
      Sys_Loc_Code = str_replace(sample_id, " - Q.*", ""),
      sampling_event_type = task_type
    )
}) %>%
  bind_rows()

df_sample_info <- df_full_a_rpts %>%
  select(
    sample_date,
    sample_id,
    lab_sdg,
    lab_name,
    Sample_Matrix_Code,
    Lab_Sample_ID,
    Sys_Loc_Code,
    sampling_event_type
  ) %>%
  distinct()

df_prep_b_rpts <- lapply(files_list_b_rpts, function(x) {
  df_in <- read_excel(file.path(fol_data, "B_reports", x),
                      col_names = FALSE)
  
  Lab_Sample_ID <- str_replace(df_in[[1,2]], "PROTOCOLO DE ANÁLISIS   ", "")
  sample_lookup_id <- str_replace(Lab_Sample_ID, " B", "")
  # lab_name <- df_in[[9,2]]
  # sample_info <- df_in[[5,2]]
  # sample_date <- df_in[[7,5]]
  # lab_sdg <- str_replace(x, " .*", "")
  # sample_id <- df_in[[20,4]]
  
  df_prep <- df_in[,c(-2, -6)]
  names(df_prep) <- c("Chemical_Name", "Result_Unit", "result_field", "Lab_Anl_Method_Name")
  
  df_out <- df_prep %>%
    tail(-11) %>%
    filter(
      !str_detect(Chemical_Name, "(?i)Página|ANEXO|Nombre|Firma y Sello "),
      !is.na(Chemical_Name),
      !is.na(result_field)
    ) %>%
    mutate(
      Result_Value = case_when(
        str_detect(result_field, "<") ~ NA_real_,
        TRUE ~ as.numeric(result_field)
      ),
      Reporting_Detection_Limit = case_when(
        str_detect(result_field, "<") ~ as.numeric(str_replace_all(result_field, "<| ", "")),
        TRUE ~ NA_real_
      ),
      Detect_Flag = case_when(
        is.na(Result_Value) ~ "N",
        TRUE ~ "Y"
      ),
      detection_limit_unit = Result_Unit,
      Lab_Sample_ID = Lab_Sample_ID,
      sample_lookup_id = sample_lookup_id
    )
}) %>%
  bind_rows()

#merge B reports data with sample metadata from A reports
df_full_b_rpts <- df_prep_b_rpts %>%
  left_join(
    df_sample_info,
    by = c("sample_lookup_id" = "Lab_Sample_ID")
  ) %>%
  select(
    -sample_lookup_id
  )

df_all_rpts <- bind_rows(df_full_b_rpts, df_full_a_rpts)

write.xlsx(df_all_rpts, file.path(fol_data, paste0(facility, "_", task_type, "_eddPrep.xlsx")))

if (get_refvals) {
  analyte <- df_all_rpts %>%
    select(
      Chemical_Name
    ) %>%
    distinct()
  
  method <- df_all_rpts %>%
    select(
      Lab_Anl_Method_Name
    ) %>%
    distinct()
  
  unit <- df_all_rpts %>%
    select(
      Result_Unit
    ) %>%
    distinct()
  
  lst_out <- list("analyte" = analyte, "method" = method, "unit" = unit)
  write.xlsx(lst_out, file.path(fol_data, paste0(facility, "_", task_type, "_RefVals-check.xlsx")))
}
