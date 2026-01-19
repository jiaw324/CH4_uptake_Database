# ============================================
# FigS1: Plot MAT/MAP comparison (Original vs ERA5) with filtering
#
# INPUT DIR:
#   D:/Users/jiaweiChiang/Desktop/Supplymentary_code
#
# INPUT FILES:
#   1) CH4 uptake data_CH4 FLUX_1_1410.xlsx   (only check exists)
#   2) Studies and Fluxes.xlsx               (main data source for plot)
#
# FILTERS:
#   - Keep only: management == "control" (case-insensitive, trim)
#   - Exclude if Q01/Q02/Q03/Q04 is checked (√/✓/✔/yes/1/true etc.)
#
# OUTPUT DIR:
#   D:/Users/jiaweiChiang/Desktop/Supplymentary_code/FigS1
# ============================================

suppressPackageStartupMessages({
  req <- c("readxl", "dplyr", "ggplot2", "stringr", "patchwork", "ggpp")
  for (p in req) if (!requireNamespace(p, quietly = TRUE)) install.packages(p)
  library(readxl)
  library(dplyr)
  library(ggplot2)
  library(stringr)
  library(patchwork)
  library(ggpp)
})

# -----------------------
# Paths
# -----------------------
in_dir <- "D:/Users/jiaweiChiang/Desktop/Supplymentary_code"

# 输入文件1：CH4 uptake data_CH4 FLUX_1_1410.xlsx（只检查存在）
in_file_ch4 <- file.path(in_dir, "CH4 uptake data_CH4 FLUX_1_1410.xlsx")

# 输入文件2：Studies and Fluxes.xlsx（画图数据源）
in_file_studies <- file.path(in_dir, "Studies and Fluxes.xlsx")

# 输出文件夹
out_dir <- file.path(in_dir, "FigS1")
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)

# 检查文件存在
stopifnot(file.exists(in_file_ch4))
stopifnot(file.exists(in_file_studies))

message("[INFO] Input CH4 file exists: ", in_file_ch4)
message("[INFO] Input Studies file: ", in_file_studies)
message("[INFO] Output dir: ", out_dir)

# -----------------------
# Helper: normalize column name
# -----------------------
norm_name <- function(x) gsub("[^a-z0-9]+", "", tolower(x))

pick_col <- function(df, target) {
  nn <- norm_name(names(df))
  idx <- which(nn == norm_name(target))
  if (length(idx) == 0) return(NULL)
  names(df)[idx[1]]
}

# -----------------------
# Helper: find sheet that contains required columns
# (support both old and new names)
# -----------------------
find_sheet_with_cols <- function(xlsx_path, needed_cols_any) {
  shs <- excel_sheets(xlsx_path)
  for (sh in shs) {
    df0 <- read_excel(xlsx_path, sheet = sh, n_max = 5)
    cols <- names(df0)
    cols_n <- norm_name(cols)
    
    ok <- TRUE
    for (cands in needed_cols_any) {
      # cands is a vector of possible names for the same field
      hit <- any(norm_name(cands) %in% cols_n)
      if (!hit) { ok <- FALSE; break }
    }
    
    if (ok) return(sh)
  }
  return(NULL)
}

# 需要找到 MAT_1/MAT_2/MAP_1/MAP_2 的工作表（兼容新旧列名）
needed_any <- list(
  c("MAT_1", "MAT_1 (original)", "MAT_1_original", "MAT1"),
  c("MAT_2", "MAT_2 (ERA5)",     "MAT_2_ERA5",     "MAT2"),
  c("MAP_1", "MAP_1 (original)", "MAP_1_original", "MAP1"),
  c("MAP_2", "MAP_2 (ERA5)",     "MAP_2_ERA5",     "MAP2")
)

sheet_use <- find_sheet_with_cols(in_file_studies, needed_any)
if (is.null(sheet_use)) {
  stop("❌ 未找到同时包含 MAT/MAP（Original 与 ERA5）所需列的工作表。请检查 Studies and Fluxes.xlsx。")
}
message("[INFO] Using sheet: ", sheet_use)

# -----------------------
# Read data
# -----------------------
df <- read_excel(in_file_studies, sheet = sheet_use)

# -----------------------
# Column matching for filter columns
# new mapping:
#   laboratory -> Q01
#   repeat     -> Q02
#   dataquality-> Q03
#   Q7         -> Q04
#   control    -> management
# -----------------------
col_q01 <- pick_col(df, "Q01")
col_q02 <- pick_col(df, "Q02")
col_q03 <- pick_col(df, "Q03")
col_q04 <- pick_col(df, "Q04")
col_mgt <- pick_col(df, "management")

# -----------------------
# Column matching for MAT/MAP fields (new names)
# -----------------------
col_mat1 <- {
  c1 <- pick_col(df, "MAT_1 (original)")
  if (is.null(c1)) c1 <- pick_col(df, "MAT_1")
  c1
}
col_mat2 <- {
  c2 <- pick_col(df, "MAT_2 (ERA5)")
  if (is.null(c2)) c2 <- pick_col(df, "MAT_2")
  c2
}
col_map1 <- {
  c3 <- pick_col(df, "MAP_1 (original)")
  if (is.null(c3)) c3 <- pick_col(df, "MAP_1")
  c3
}
col_map2 <- {
  c4 <- pick_col(df, "MAP_2 (ERA5)")
  if (is.null(c4)) c4 <- pick_col(df, "MAP_2")
  c4
}

if (any(sapply(list(col_mat1, col_mat2, col_map1, col_map2), is.null))) {
  stop("❌ MAT/MAP 列名匹配失败：请确认存在 MAT_1(original)/MAT_2(ERA5)/MAP_1(original)/MAP_2(ERA5)")
}

message("[INFO] MAT/MAP columns detected:")
message("  MAT1 = ", col_mat1)
message("  MAT2 = ", col_mat2)
message("  MAP1 = ", col_map1)
message("  MAP2 = ", col_map2)

# -----------------------
# Flag parser: checked?
# -----------------------
is_checked <- function(x) {
  s <- tolower(trimws(as.character(x)))
  s <- gsub("\\s+", "", s)
  s %in% c("√","✓","✔","✅","☑","v","true","t","1","yes","y","checked")
}

is_control <- function(x) {
  s <- tolower(trimws(as.character(x)))
  s == "control"
}

# -----------------------
# Apply filtering
# -----------------------
n0 <- nrow(df)

# 1) Keep only management == control
if (!is.null(col_mgt)) {
  df <- df %>% filter(is_control(.data[[col_mgt]]))
} else {
  stop("❌ 未找到 management 列，无法进行 control 筛选。请检查表头。")
}

# 2) Exclude checked Q01-Q04
if (!is.null(col_q01)) df <- df %>% filter(!is_checked(.data[[col_q01]]))
if (!is.null(col_q02)) df <- df %>% filter(!is_checked(.data[[col_q02]]))
if (!is.null(col_q03)) df <- df %>% filter(!is_checked(.data[[col_q03]]))
if (!is.null(col_q04)) df <- df %>% filter(!is_checked(.data[[col_q04]]))

n1 <- nrow(df)
message(sprintf("[INFO] Filtered rows: %d -> %d (removed %d)", n0, n1, n0 - n1))

# -----------------------
# Numeric conversion
# -----------------------
to_num <- function(x) suppressWarnings(as.numeric(x))

df <- df %>%
  mutate(
    MAT_1_plot = to_num(.data[[col_mat1]]),
    MAT_2_plot = to_num(.data[[col_mat2]]),
    MAP_1_plot = to_num(.data[[col_map1]]),
    MAP_2_plot = to_num(.data[[col_map2]])
  )

# -----------------------
# Plot settings
# -----------------------
BASE_SIZE   <- 16
ANN_SIZE    <- 5.0
POINT_SIZE  <- 1.6
LINE_W_LM   <- 1.1
LINE_W_11   <- 0.9

NPC_X <- 0.05
NPC_Y <- 0.95

# -----------------------
# Plot function
# -----------------------
plot_scatter_lm <- function(dat, xcol, ycol, xlabel, ylabel, add_1to1 = TRUE) {
  
  dd <- dat %>%
    select(x = all_of(xcol), y = all_of(ycol)) %>%
    filter(is.finite(x), is.finite(y))
  
  n <- nrow(dd)
  if (n < 3) stop("Not enough valid pairs for: ", xcol, " vs ", ycol)
  
  fit <- lm(y ~ x, data = dd)
  co  <- coef(fit)
  slope <- unname(co["x"])
  intercept <- unname(co["(Intercept)"])
  r2 <- summary(fit)$r.squared
  
  ann <- sprintf("y = %.3fx + %.3f\nR² = %.3f\nN = %d", slope, intercept, r2, n)
  
  p <- ggplot(dd, aes(x = x, y = y)) +
    geom_point(size = POINT_SIZE, alpha = 0.65, color = "orange") +
    geom_smooth(method = "lm", se = FALSE, linewidth = LINE_W_LM) +
    labs(title = NULL, x = xlabel, y = ylabel) +
    theme_bw(base_size = BASE_SIZE) +
    theme(
      plot.title = element_blank(),
      panel.grid.minor = element_blank()
    ) +
    ggpp::annotate(
      "text_npc",
      npcx = NPC_X, npcy = NPC_Y,
      label = ann,
      hjust = 0, vjust = 1,
      size = ANN_SIZE
    )
  
  if (add_1to1) {
    p <- p + geom_abline(intercept = 0, slope = 1,
                         linetype = "dashed", linewidth = LINE_W_11)
  }
  
  return(p)
}

# -----------------------
# Make plots
# -----------------------
p_mat <- plot_scatter_lm(
  df,
  xcol = "MAT_1_plot", ycol = "MAT_2_plot",
  xlabel = "MAT (Original, ℃)",
  ylabel = "MAT (ERA5, ℃)",
  add_1to1 = TRUE
)

p_map <- plot_scatter_lm(
  df,
  xcol = "MAP_1_plot", ycol = "MAP_2_plot",
  xlabel = "MAP (Original, mm)",
  ylabel = "MAP (ERA5, mm)",
  add_1to1 = TRUE
)

# -----------------------
# Save single panels
# -----------------------
ggsave(file.path(out_dir, "FigS1_scatter_MAT_vs_ERA5.png"),
       p_mat, width = 6.5, height = 6.5, dpi = 600)

ggsave(file.path(out_dir, "FigS1_scatter_MAP_vs_ERA5.png"),
       p_map, width = 6.5, height = 6.5, dpi = 600)

# -----------------------
# Combine side-by-side
# -----------------------
p_combo <- p_mat + p_map + plot_layout(ncol = 2)

ggsave(file.path(out_dir, "FigS1_scatter_MAT_MAP_vs_ERA5_side_by_side.png"),
       p_combo, width = 13.5, height = 6.5, dpi = 600)

message("✅ All done. Output dir: ", out_dir)