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
facility = "Barranquilla" #which site is this data for
task_type = "2018-SoilInvestigation" #if no specific task type, put 'NA' and the standard default will be generated
lab = "TAPEN" #TAPEN = Test America Pensacola
filename = "400-212739-1_Results" #filename to be read
method_lookupfile = "400-212739-1_LabChronicle" #lab chronicle section extracted for lab report
write_edd = TRUE #do you want to write an EDD draft containing all the information provided so far?
get_refvals = FALSE #do you want to get refvals summary from data
source("function.R")

fol_main <- here::here()
fol_edd <- file.path(fol_main, "data")
fol_data <- file.path(fol_main, "data", facility, "Corteva Barranquilla Reports", "ERM Soil Investigation 2016")

# set up data ####
df_in <- read_excel(file.path(fol_data, paste0(filename, ".xlsx")),
                    col_names = FALSE)

#get indices of start of each section to which metadata can be shared
sample_heads <- which(str_detect(df_in$...1, "Client Sample ID: "))
method_heads <- which(str_detect(df_in$...1, "Method: "))

#to get method ends, find first blank row after non-blank block
#loop over method_heads values and find first empty value after each
sample_ends <- c(0, sample_heads-1, nrow(df_in))
method_ends <- c(0, method_heads-1, nrow(df_in))[-2]
#method_ends <- which(is.na(df_in$...1))

sample_lengths <- diff(sample_ends)
method_lengths <- diff(method_ends)

#create dataframes with sample id and the group number (equal to rownum)
#to join tables on group/row number
sample_ids <- data.frame(df_in[sample_heads, 1], sample_heads) %>%
  mutate(sample_group_num = row_number()) %>%
  rename(sample_info = ...1)

method_names <- data.frame(df_in[method_heads, 1], method_heads) %>%
  mutate(method_group_num = row_number()) %>%
  rename(analytic_method_name = ...1)

#split dataframe by sample id group
df_split_prep <- df_in %>%
  mutate(
    sample_group_num = rep(seq_along(sample_lengths), sample_lengths)-1,
    method_group_num = rep(seq_along(method_lengths), method_lengths)
    #seq() retunrs a sequence (of numbers) that is the length of its input
    #rep() replicates the values in x (seq_along(report_lengths)), report_lengths
    #number of times
  ) %>%
  left_join(
    sample_ids, by = c("sample_group_num" = "sample_group_num")
  ) %>%
  left_join(
    method_names, by = c("method_group_num" = "method_group_num")
  )

split_lst <- split(df_split_prep, df_split_prep$sample_group_num)[-1]

full_df <- lapply(split_lst, function(x) {
  print(x)
  #separate first six rows from dataframe, this block will contain metadata
  #for proceeding dataframe
  info_block <- head(x, 6)
  
  #surrogates don't give all the col headers, so need to fill in manually in those cases
  if(x[4,1] == "Surrogate" | x[6,1] == "Surrogate") {
    colnames(x)[1:11] <- c("Analyte", "Result", "Qualifier", "RL", "MDL", "Unit", "D", "Prepared", "Analyzed", "Dil Fac", "5")
  } else {
    # colnames(x)[1:11] <- x[7,1:11]
    colnames(x)[1:11] <- x[which(x[1] == "Analyte")[1],1:11]
  }
  
  #handle cases where pdf extraction may have messed up col headers
  # if(sum(is.na(names(x))) > 0) {
  #   sapply(1:ncol(x), function(y) {
  #     if(is.na(names(x)[y])) {
  #       if(str_detect(names(x)[y-1], "   ")) {
  #         names(x)[y] = str_extract(names(x)[y-1], "   .*")
  #         names(x)[y-1] = str_extract(names(x)[y-1], ".*   ")
  #       } else if(str_detect(names(x)[y+1], "   ")) {
  #         names(x)[y] = str_extract(names(x)[y+1], ".*   ")
  #         names(x)[y+1] = str_extract(names(x)[y+1], "   .*")
  #       }
  #     }
  #     return(x)
  #   })
  # }
  
  m <- nrow(info_block)
  n <- ncol(info_block)
  x_unlst <- as.vector(t(as.matrix(info_block)))

  sample_idx <- which(str_detect(x_unlst, "Lab Sample ID: "))
  sample_row <- (div(sample_idx, n) + 1)[[1]]
  sample_col <- ((sample_idx %% (n+1)))[[1]]
  
  matrix_idx <- which(str_detect(x_unlst, "Matrix:"))
  matrix_row <- (div(matrix_idx, n) + 1)[[1]]
  matrix_col <- ((matrix_idx %% (n)))[[1]]
  
  
  #populate dataframe with info block and do some housekeeping
  df_out <- x %>%
    # tail(-5) %>%
    mutate(
      Sys_Sample_Code = str_replace(info_block$...1[[1]], "Client Sample ID: ", ""),
      # Lab_Sample_ID = str_replace(info_block$...8[[1]], "Lab Sample ID: ", ""),
      Lab_Sample_ID = str_replace(info_block[sample_row, sample_col], "Lab Sample ID: ", ""),
      # Sample_Matrix = str_replace(info_block$...9[[2]], "Matrix: ", ""),
      Sample_Matrix = str_replace(info_block[matrix_row, matrix_col], "Matrix: ", ""),
      Sample_Date_Time = str_replace(info_block$...1[[2]], "Date Collected: ", ""),
      Date_Received = str_replace(info_block$...1[[3]], "Date Received: ", ""),
      sample_delivery_group = "400-123152-1",
      result_type_code = case_when(
        str_detect(Analyte, "(Surr)") ~ "SUR",
        Analyte == "2-Fluorobiphenyl" ~ "SUR",
        Analyte == "Tetrachloro-m-xylene" ~ "SUR",
        str_detect(RL, "- ") ~ "SUR",
        TRUE ~ "TRG"
      ),
      percent_solids = str_replace(info_block$...2[[4]], "Percent Solids: ", ""),
      prep_date = substr(Prepared, nchar(Prepared)-13, nchar(Prepared)),
      Lab_Anl_Method_Name = case_when(
        Analyte == "Percent Moisture" ~ "Moisture",
        TRUE ~ str_match(analytic_method_name, "Method:\\s*(.*?)\\s* -")[,2]
      ),
    ) %>%
    # head(-5) %>% #last five rows are footer and beginning of next section's header
    filter(
      !is.na(Analyte),
      !(Analyte == "Analyte" | Analyte == "Surrogate" | Analyte == "General Chemistry"),
      !str_detect(Analyte, "Client|Date|Project|Page"),
      !str_detect(Result, "TestAmerica|Client")
    )
  
  df_out <- df_out[-c(11)]
  
  return(df_out)
  
}) %>%
  bind_rows()

# set up method lookups ####
df_lookup_in <- read_excel(file.path(fol_data, paste0(method_lookupfile, ".xlsx")),
                           col_names = FALSE)

#get indices of start of each section to which metadata can be shared
lookup_sample_heads <- which(str_detect(df_lookup_in$...1, "Client Sample ID: "))
lookup_sample_ends <- c(0, lookup_sample_heads-1, nrow(df_lookup_in))[-2]
lookup_sample_lengths <- diff(lookup_sample_ends)

lookup_sample_names <- data.frame(df_lookup_in[lookup_sample_heads, 1], lookup_sample_heads) %>%
  mutate(
    sample_group_num = row_number(),
    lookup_sample_name = str_replace(...1, "Client Sample ID: ", "")
    ) %>%
  select(-c(...1))

#find column with dates and rename it
names(df_lookup_in)[which(grepl("(?i)Or Anal", df_lookup_in))] <- "analysis_prep_date"
names(df_lookup_in)[which(grepl("(?i)Batch", df_lookup_in) & grepl("Number", df_lookup_in))] <- "batch"

#group by sample id
df_lookup <- df_lookup_in %>%
  mutate(
    sample_group_num = rep(seq_along(lookup_sample_lengths), lookup_sample_lengths),
    #seq() retunrs a sequence (of numbers) that is the length of its input
    #rep() replicates the values in x (seq_along(report_lengths)), report_lengths
    #number of times
    batch = as.integer(batch) #remove trailing zeroes added by read_excel
  ) %>%
  left_join(
    lookup_sample_names, by = c("sample_group_num" = "sample_group_num")
  ) %>%
  distinct()

df_prep_lookup <- df_lookup %>%
  filter(
  ...2 == "Prep"
) %>%
  rename(
    Lab_Prep_Method_Name = ...3
  )

df_anl_lookup <- df_lookup %>%
  filter(
    ...2 == "Analysis"
  ) %>%
  rename(
    Test_Batch_ID = batch,
    Lab_Anl_Method_Name = ...3
  )

#join data to prep methods by sample id and prep date/time
df_add_info <- full_df %>%
  left_join(
    df_prep_lookup,
    by = c("Sys_Sample_Code" = "lookup_sample_name",
           "prep_date" = "analysis_prep_date"),
    keep = FALSE,
    multiple = "first" #all matches are the same, take just the first
  ) %>%
  left_join(
    df_anl_lookup,
    by = c("Sys_Sample_Code" = "lookup_sample_name",
           "Lab_Anl_Method_Name" = "Lab_Anl_Method_Name",
           "Analyzed" = "analysis_prep_date"),
    keep = FALSE,
    multiple = "first"
  )

# prepare for edd format ####
df_edd_prep <- df_add_info %>%
  mutate(
    Result_Value = case_when(
      # is.na(Result) ~ NA_real_,
      str_detect(Qualifier, "U") ~ NA_character_,
      !is.na(as.numeric(Result)) ~ as.character(as.numeric(Result))#,
      # TRUE ~ "check lab report"
    ),
    Detect_Flag = case_when(
      # !is.na(Result) ~ "Y",
      str_detect(Qualifier, "U") ~ "N",
      TRUE ~ "Y"
    ),
    #this is assuming a naming structure where sample depth is surrounded by 'V' and 'N'
    #within the sample ID, may need to change around for other reports/sites
    sample_depth = str_extract(Sys_Sample_Code, "V\\s*(.*?)\\s*N"),
    start_depth = case_when(
      !str_detect(sample_depth, "-") ~ gsub( "V|N", "", sample_depth),
      TRUE ~ str_replace_all(sample_depth, ".*-|N|V", "")
    ),
    end_depth = case_when(
      !str_detect(sample_depth, "-") ~ NA_character_,
      TRUE ~ str_replace_all(sample_depth, ".*-|N|V", "")
    ),
    depth_unit = case_when(
      !is.na(start_depth) ~ "m",
      TRUE ~ NA_character_
    ),
    Sample_Date = str_replace(Sample_Date_Time, " .*", ""),
    Sample_Time = str_replace(Sample_Date_Time, ".* ", ""),
    Analysis_Date = str_replace(Analyzed, " .*", ""),
    Analysis_Time = str_replace(Analyzed, ".* ", ""),
    Prep_Date = str_replace(prep_date, " .*", ""),
    Prep_Time = str_replace(prep_date, ".* ", ""),
    Reporting_Detection_Limit = case_when(
      !is.na(as.numeric(MDL)) ~ as.character(as.numeric(MDL)),
      TRUE ~ RL
    ),
    method_detection_limit = case_when(
      !is.na(as.numeric(MDL)) ~ as.character(as.numeric(MDL)), 
      TRUE ~ MDL
    ),
    Sample_Matrix_Code = case_when(
      Sample_Matrix == "Solid" ~ "SO",
      Sample_Matrix == "Liquid" ~ "WG",
      Sample_Matrix == "Water" ~ "WG"
    ),
    Test_Type = "INITIAL",
    column_number = "NA",
    validated_yn = "N",
    task_code = case_when(
      is.na(task_type) & Sample_Matrix == "Solid" ~ paste0(
        "20",
        format(as.Date(Sample_Date, "%m/%d/%y"), "%y"),
        quarters(as.Date(Sample_Date, "%m/%d/%y")),
        "-SoilBoring"
      ),
      is.na(task_type) & Sample_Matrix == "Liquid" | Sample_Matrix == "Water" ~ paste0(
        "20",
        format(as.Date(Sample_Date, "%m/%d/%y"), "%y"),
        quarters(as.Date(Sample_Date, "%m/%d/%y")),
        "-GWMonitor"
      ),
      TRUE ~ task_type
    ),
    Lab_Name_Code = lab,
    reportable_result = "Y",
    quantitation_limit = case_when(
      !is.na(as.numeric(RL)) ~ as.character(as.numeric(RL)),
      TRUE ~ RL
    )
  ) %>%
  rename(
    Chemical_Name = Analyte,
    Result_Unit = Unit,
    Lab_Qualifiers = Qualifier,
    dilution_factor = `Dil Fac`
  ) %>%
  filter(
    !str_detect(Chemical_Name, "Method")
  )

if(get_refvals) {
  anl_method <- df_edd_prep %>% select(Lab_Anl_Method_Name) %>% distinct()
  prep_method <- df_edd_prep %>% select(Lab_Prep_Method_Name) %>% distinct()
  analyte <- df_edd_prep %>% select(Chemical_Name) %>% distinct()
  
  lst_rv <- list("analytic method" = anl_method,
                 "prep method" = prep_method,
                 "analyte" = analyte)
  write.xlsx(lst_rv, file.path("output", paste0("refVals_", filename, ".xlsx")))
}

if(write_edd) {
  path_edd <- here::here(fol_edd, "ESBasic_ERM_Desc.xlsx")
  fol_remaps <- here::here(fol_main, "data", facility, "remapping")
  df_chem_remaps <- read_excel(file.path(fol_remaps, "testAmerica_remapGuide.xlsx"), sheet = 1)
  df_anl_method_remaps <- read_excel(file.path(fol_remaps, "testAmerica_remapGuide.xlsx"), sheet = 2)
  df_prep_method_remaps <- read_excel(file.path(fol_remaps, "testAmerica_remapGuide.xlsx"), sheet = 3)
  df_unit_remaps <- read_excel(file.path(fol_remaps, "testAmerica_remapGuide.xlsx"), sheet = 4)
  
  edd_out <- df_edd_prep  %>%
    mutate(
      #since these two fields are mapped based on chemical name, populate with that
      #for next step
      Fraction = Chemical_Name,
      Cas_Rn = Chemical_Name
    ) %>%
    add_edd_columns(path_edd, "ESBasic_ERM") %>%
    rename(`#Sys_Sample_Code` = Sys_Sample_Code)
  
  cas_maps <- list(df_chem_remaps$equis_cas_rn) %>%
    unlist() %>%
    setNames(df_chem_remaps$external_value)
  
  fraction_maps <- list(df_chem_remaps$fraction) %>%
    unlist() %>%
    setNames(df_chem_remaps$external_value)
  
  anl_meth_maps <- list(df_anl_method_remaps$equis_analytic_method) %>%
    unlist() %>%
    setNames(df_anl_method_remaps$external_value)
  
  prep_meth_maps <- list(df_prep_method_remaps$equis_prep_method) %>%
    unlist() %>%
    setNames(df_prep_method_remaps$external_value)
  
  unit_maps <- list(df_unit_remaps$equis_unit) %>%
    unlist() %>%
    setNames(df_unit_remaps$external_value)
  
  edd_out$Cas_Rn <- cas_maps[edd_out$Cas_Rn]
  edd_out$Fraction <- fraction_maps[edd_out$Fraction]
  edd_out$Lab_Anl_Method_Name <- anl_meth_maps[edd_out$Lab_Anl_Method_Name]
  edd_out$Lab_Prep_Method_Name <- prep_meth_maps[edd_out$Lab_Prep_Method_Name]
  edd_out$Result_Unit <- unit_maps[edd_out$Result_Unit]
  
  #now that units are remapped, populate detection limit units
  edd_out$detection_limit_unit <- edd_out$Result_Unit
  
  trg_edd_out <- edd_out %>%
    filter(
      result_type_code == "TRG"
    )
  
  sur_edd_out <- edd_out %>%
    filter(
      result_type_code == "SUR"
    )
  lst_edd_out <- list("Target Analytes" = trg_edd_out,
                      "Surrogates" = sur_edd_out)
  
  write.xlsx(lst_edd_out, file.path("output", paste0(facility, "-", filename, ".ESBasic_ERM-", Sys.Date(), ".xlsx")))
} else {
  write.xlsx(df_edd_prep, file.path("output", paste0(facility, "-", filename, ".ESBasic_ERM-", Sys.Date(), ".xlsx")))
}

