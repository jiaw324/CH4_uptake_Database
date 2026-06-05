# =========================================================
# Quantitative support for temperature/moisture framework
# Standardized multiple regression only
# Joint-control threshold = 2.0
# If object long_all does not exist, read data from file
# Output: only one sheet named "standardized_regression"
# =========================================================

suppressPackageStartupMessages({
  library(dplyr)
  library(openxlsx)
  library(readxl)
  library(readr)
})

# ---------------------------------------------------------
# 0. Set input and output paths
# ---------------------------------------------------------

# !!! 修改这里：换成你实际包含 Ecosystem, CH4_flux, ST, SM 的文件
# 可以是 .xlsx 或 .csv
input_file <- "D:/233 CH4 uptake_Database/CH4 uptake_Database_Write/Supplymentary_code/your_input_file.xlsx"

out_dir <- "D:/233 CH4 uptake_Database/CH4 uptake_Database_Write/Supplymentary_code/Table S1"

if (!dir.exists(out_dir)) {
  dir.create(out_dir, recursive = TRUE)
  message("Output directory created: ", out_dir)
} else {
  message("Output directory already exists: ", out_dir)
}

# ---------------------------------------------------------
# 1. Load data
# ---------------------------------------------------------

if (exists("long_all")) {
  
  message("Object 'long_all' found in current R environment.")
  data_all <- long_all
  
} else {
  
  message("Object 'long_all' not found. Trying to read data from input_file...")
  
  if (!file.exists(input_file)) {
    stop(
      "Object 'long_all' not found, and input_file does not exist:\n",
      input_file,
      "\n\nPlease either:\n",
      "1) Run the previous data-processing script first to create long_all; or\n",
      "2) Change input_file to the correct Excel/CSV file path."
    )
  }
  
  file_ext <- tools::file_ext(input_file)
  
  if (tolower(file_ext) %in% c("xlsx", "xls")) {
    data_all <- readxl::read_excel(input_file)
  } else if (tolower(file_ext) == "csv") {
    data_all <- readr::read_csv(input_file, show_col_types = FALSE)
  } else {
    stop("Unsupported file type: ", file_ext, ". Please use .xlsx, .xls, or .csv.")
  }
  
  message("Data loaded from: ", input_file)
}

# ---------------------------------------------------------
# 2. Check required columns
# ---------------------------------------------------------

need_cols <- c("Ecosystem", "CH4_flux", "ST", "SM")
miss_cols <- setdiff(need_cols, names(data_all))

if (length(miss_cols) > 0) {
  stop(
    "Missing columns in input data: ",
    paste(miss_cols, collapse = ", "),
    "\n\nCurrent columns are:\n",
    paste(names(data_all), collapse = ", ")
  )
}

# ---------------------------------------------------------
# 3. Data preparation
# ---------------------------------------------------------

reg_data <- data_all %>%
  dplyr::select(Ecosystem, CH4_flux, ST, SM) %>%
  dplyr::filter(
    !is.na(Ecosystem),
    !is.na(CH4_flux),
    !is.na(ST),
    !is.na(SM)
  )

# ---------------------------------------------------------
# 4. Standardized multiple regression helper
# ---------------------------------------------------------

safe_standardized_lm <- function(dat, min_n = 10) {
  
  dat <- dat %>%
    dplyr::filter(
      !is.na(CH4_flux),
      !is.na(ST),
      !is.na(SM)
    )
  
  n_obs <- nrow(dat)
  
  if (n_obs < min_n) {
    return(tibble(
      n = n_obs,
      beta_ST = NA_real_,
      p_ST = NA_real_,
      beta_SM = NA_real_,
      p_SM = NA_real_,
      abs_beta_ST = NA_real_,
      abs_beta_SM = NA_real_,
      Dominance = "Insufficient data"
    ))
  }
  
  if (length(unique(dat$CH4_flux)) < 2 ||
      length(unique(dat$ST)) < 2 ||
      length(unique(dat$SM)) < 2) {
    return(tibble(
      n = n_obs,
      beta_ST = NA_real_,
      p_ST = NA_real_,
      beta_SM = NA_real_,
      p_SM = NA_real_,
      abs_beta_ST = NA_real_,
      abs_beta_SM = NA_real_,
      Dominance = "Insufficient variation"
    ))
  }
  
  dat_std <- dat %>%
    dplyr::mutate(
      CH4_flux_z = as.numeric(scale(CH4_flux)),
      ST_z = as.numeric(scale(ST)),
      SM_z = as.numeric(scale(SM))
    )
  
  fit <- lm(CH4_flux_z ~ ST_z + SM_z, data = dat_std)
  sm <- summary(fit)
  coef_tab <- sm$coefficients
  
  beta_ST <- coef_tab["ST_z", "Estimate"]
  p_ST    <- coef_tab["ST_z", "Pr(>|t|)"]
  
  beta_SM <- coef_tab["SM_z", "Estimate"]
  p_SM    <- coef_tab["SM_z", "Pr(>|t|)"]
  
  abs_beta_ST <- abs(beta_ST)
  abs_beta_SM <- abs(beta_SM)
  
  beta_ratio_threshold <- 2.0
  
  Dominance <- dplyr::case_when(
    is.na(beta_ST) | is.na(beta_SM) ~ "Unclassified",
    
    p_ST < 0.05 & p_SM >= 0.05 ~ "Temperature-dominated",
    
    p_ST >= 0.05 & p_SM < 0.05 ~ "Moisture-dominated",
    
    p_ST < 0.05 & p_SM < 0.05 &
      abs_beta_ST > abs_beta_SM * beta_ratio_threshold ~ "Temperature-dominated",
    
    p_ST < 0.05 & p_SM < 0.05 &
      abs_beta_SM > abs_beta_ST * beta_ratio_threshold ~ "Moisture-dominated",
    
    p_ST < 0.05 & p_SM < 0.05 ~ "Jointly controlled",
    
    TRUE ~ "No clear dominant control"
  )
  
  tibble(
    n = n_obs,
    beta_ST = beta_ST,
    p_ST = p_ST,
    beta_SM = beta_SM,
    p_SM = p_SM,
    abs_beta_ST = abs_beta_ST,
    abs_beta_SM = abs_beta_SM,
    Dominance = Dominance
  )
}

# ---------------------------------------------------------
# 5. Standardized multiple regression by ecosystem
# ---------------------------------------------------------

reg_summary <- reg_data %>%
  dplyr::group_by(Ecosystem) %>%
  dplyr::group_modify(~ safe_standardized_lm(.x, min_n = 10)) %>%
  dplyr::ungroup()

# ---------------------------------------------------------
# 6. Keep selected columns only
# ---------------------------------------------------------

standardized_regression <- reg_summary %>%
  dplyr::mutate(
    `abs_beta_ST/abs_beta_SM` = dplyr::case_when(
      is.na(abs_beta_ST) | is.na(abs_beta_SM) ~ NA_real_,
      abs_beta_SM == 0 ~ NA_real_,
      TRUE ~ abs_beta_ST / abs_beta_SM
    )
  ) %>%
  dplyr::select(
    Ecosystem,
    n,
    beta_ST,
    p_ST,
    beta_SM,
    p_SM,
    abs_beta_ST,
    abs_beta_SM,
    `abs_beta_ST/abs_beta_SM`,
    Dominance
  )

print(standardized_regression)

# ---------------------------------------------------------
# 7. Export Excel: only one sheet
# ---------------------------------------------------------

out_xlsx <- file.path(
  out_dir,
  "Standardized_regression_temperature_moisture_2p0_threshold.xlsx"
)

wb <- openxlsx::createWorkbook()

openxlsx::addWorksheet(wb, "standardized_regression")

openxlsx::writeData(
  wb,
  sheet = "standardized_regression",
  x = standardized_regression
)

openxlsx::setColWidths(
  wb,
  sheet = "standardized_regression",
  cols = 1:ncol(standardized_regression),
  widths = "auto"
)

openxlsx::saveWorkbook(
  wb,
  out_xlsx,
  overwrite = TRUE
)

message("Excel exported: ", out_xlsx)
message("✅ Standardized multiple regression analysis with 2.0-fold threshold finished.")