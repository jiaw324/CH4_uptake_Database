# =========================================================
# Count and proportion of Q01-Q07 flags
# Output only Q_summary sheet
# =========================================================

library(readxl)
library(dplyr)
library(tidyr)
library(writexl)

# -----------------------------
# 1. File paths
# -----------------------------
root_dir <- "D:/233 CH4 uptake_Database/CH4 uptake_Database_Write/Supplymentary_code"

input_file <- file.path(root_dir, "Studies and Fluxes.xlsx")

output_dir <- file.path(root_dir, "Count")

if (!dir.exists(output_dir)) {
  dir.create(output_dir, recursive = TRUE)
}

output_file <- file.path(output_dir, "Q01_Q07_Q_summary.xlsx")

# -----------------------------
# 2. Read data
# -----------------------------
df <- read_excel(input_file, sheet = 1)

q_cols <- paste0("Q", sprintf("%02d", 1:7))

# Check whether Q01-Q07 exist
missing_cols <- setdiff(q_cols, names(df))
if (length(missing_cols) > 0) {
  stop("Missing columns: ", paste(missing_cols, collapse = ", "))
}

total_records <- nrow(df)

# -----------------------------
# 3. Clean Q01-Q07 flags
# -----------------------------
df_q <- df %>%
  mutate(across(
    all_of(q_cols),
    ~ trimws(as.character(.x))
  )) %>%
  mutate(across(
    all_of(q_cols),
    ~ ifelse(is.na(.x) | .x == "" | .x == "NA", NA, .x)
  ))

# -----------------------------
# 4. Count each Q flag
# -----------------------------
q_summary <- df_q %>%
  summarise(across(
    all_of(q_cols),
    ~ sum(!is.na(.x))
  )) %>%
  pivot_longer(
    cols = everything(),
    names_to = "Q_flag",
    values_to = "n_flagged_records"
  ) %>%
  mutate(
    total_records = total_records,
    proportion = n_flagged_records / total_records,
    percentage = proportion * 100
  )

# -----------------------------
# 5. Export only Q_summary sheet
# -----------------------------
write_xlsx(
  list(
    Q_summary = q_summary
  ),
  output_file
)

cat("Done!\n")
cat("Output file:\n", output_file, "\n")