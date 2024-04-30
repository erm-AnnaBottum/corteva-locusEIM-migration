

clean_pi_data <- function(file_name) {
  print(file_name)
  df_in <- readxl::read_excel(file_name, #file.path(fol_data, "357117 PI.xlsx"),
                              #range = readxl::anchored(1, c(6, 140)), #specify range of cells to read in
                              col_names = paste(1:6) #add column indices
                              #col_names = paste(1:5) #add column indices- update to 5 for 2019 Colon data
  )
  
  df_temp <- df_in %>%
    mutate(
      protocol_for_report = df_in[2,5],
      expedition_date = df_in[3,5],
      lab_name = df_in[4,5],
      qualification_certificate_number = df_in[5,5],
      coc_number = df_in[6,5],
      sample_date = df_in[7,5],
      sample_receipt_date = df_in[8,5],
      cuit_tax_id = df_in[10,2],
      id_set = df_in[11,2],
      facility = df_in[11,5],
      sample_matrix = df_in[16,1],
      sample_preservation = df_in[18,4],
      sample_name = df_in[20,4]
    ) %>%
    tail(-24) %>% #get rid of first 24 rows of header info now that its extracted
    filter(
      !grepl("gina", `1`),
      !grepl("(?i)GASEOSO", `1`),
      !grepl("(?i)ESPECTROMETRO", `1`),
      is.na(`3`),
      !is.na(`2`) | !is.na(`5`) | grepl("mg", `6`)
      #change made for Colon 2019 data
      # is.na(`2`),
      # !is.na(`3`) | !is.na(`5`) | grepl("mg", `6`)
    ) %>%
    left_join(df_remaps, by = c("1" = "format_fit_analyte_name"))
  
  #remove empty column 3
  df_out <- data.frame(df_temp[,-3])
  
  #remove empty column 2 - updated for 2019 Colon data
  # df_out <- data.frame(df_temp[,-2])
  
  names(df_out) <- c(
    "original_chem_name",
    "result_field",
    "analytic_method",
    "detection_limit",
    "quantification_limit",
    "protocol_for_report",
    "expedition_date",
    "lab_name",
    "qualification_certificate_number",
    "coc_number",
    "sample_date",
    "sample_receipt_date",
    "cuit_tax_id",
    "id_set",
    "facility",
    "sample_matrix",
    "sample_preservation",
    "sample_name",
    "chemical_name"
  )
  
  return(df_out)
}






get_edd_column_names <- function(edd_format_path, edd_ws) {
  vec_fields <- read_excel(edd_format_path, edd_ws) %>%
    pull(`Field Name`)
  
  n_fields <- min(which(is.na(vec_fields)) - 1, length(vec_fields))
  
  vec_fields[1:n_fields]
}

add_edd_columns <- function(df, edd_format_path, edd_ws) {
  edd_column_names <- get_edd_column_names(edd_format_path, edd_ws)
  
  current_column_names <- names(df)
  
  missing_column_names <- base::setdiff(
    edd_column_names, current_column_names
  ) %>%
    setNames(nm = .)
  
  missing_column_names[] <- NA
  
  tibble::add_column(df, !!!missing_column_names) %>%
    select(all_of(edd_column_names))
}
