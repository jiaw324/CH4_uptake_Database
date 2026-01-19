# make_NS_overall_triptych_UNIFIED_COLORS__FigS5__StudiesFluxes.R
suppressPackageStartupMessages({
  req <- c("readxl","dplyr","stringr","tidyr","ggplot2","cowplot",
           "readr","rlang","scales","openxlsx")
  for (p in req) if (!requireNamespace(p, quietly = TRUE)) install.packages(p)
  invisible(lapply(req, library, character.only = TRUE))
})

# ------------------------- Paths -------------------------
root_dir <- "D:/Users/jiaweiChiang/Desktop/Supplymentary_code"

# 输入1（附件1，仅检查存在，不读取）
in_file_pairs <- file.path(root_dir, "CH4 uptake data_CH4 FLUX_1_1410.xlsx")

# 输入2（附件2，主数据源）
in_file_main  <- file.path(root_dir, "Studies and Fluxes.xlsx")

# 输出路径（FigS5）
out_dir <- file.path(root_dir, "FigS5")
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)

stopifnot(file.exists(in_file_pairs))
stopifnot(file.exists(in_file_main))

# ------------------------- Helpers -------------------------
safe_save <- function(path, plot, width, height, dpi=600, bg="white"){
  dir.create(dirname(path), recursive = TRUE, showWarnings = FALSE)
  ggplot2::ggsave(filename = path, plot = plot, width = width, height = height,
                  dpi = dpi, bg = bg, device = grDevices::jpeg, limitsize = FALSE)
}

norm_txt <- function(x){
  if (is.null(x)) return(NA_character_)
  stringr::str_squish(as.character(x))
}

is_checked <- function(x){
  y <- stringr::str_to_lower(norm_txt(x))
  !is.na(y) & stringr::str_detect(y, "^(√|✓|✔|yes|y|1|true)$")
}

is_control <- function(x){
  y <- stringr::str_to_lower(norm_txt(x))
  !is.na(y) & stringr::str_detect(y, "^control(?:\\s*group)?$")
}

safe_filename <- function(s){
  s <- as.character(s)
  s <- iconv(s, from = "UTF-8", to = "ASCII//TRANSLIT", sub = "_")
  s <- gsub("[/\\\\:*?\"<>|]", "_", s)
  s <- gsub("[^A-Za-z0-9._-]+", "_", s)
  s <- gsub("^_+|_+$","", s)
  if (nchar(s) == 0) s <- "unknown"
  s
}

colmap_exact <- function(target, pool){
  hit <- pool[tolower(pool) == tolower(target)]
  if (length(hit) == 0) stop(sprintf("Column not found: %s", target))
  hit[1]
}

to_num <- function(x) suppressWarnings(as.numeric(readr::parse_number(as.character(x))))

# ------------------------- Read (Studies and Fluxes.xlsx) -------------------------
sheets <- readxl::excel_sheets(in_file_main)

# 优先读 row；否则如果有 Sheet1 就读；否则读第一个
sheet_to_read <- if ("row" %in% sheets) {
  "row"
} else if ("Sheet1" %in% sheets) {
  "Sheet1"
} else {
  sheets[1]
}

dat <- readxl::read_excel(in_file_main, sheet = sheet_to_read)
if (nrow(dat) == 0) stop("Main input sheet has 0 rows: ", sheet_to_read)

# ------------------------- Columns -------------------------
cn <- names(dat)

# ✅ 新列名：Q01 Q02 Q03 Q04 management
col_q01 <- colmap_exact("Q01", cn)
col_q02 <- colmap_exact("Q02", cn)
col_q03 <- colmap_exact("Q03", cn)
col_q04 <- colmap_exact("Q04", cn)
col_mgmt<- colmap_exact("management", cn)

# latitude auto-detect（优先匹配 Latitude）
lat_aliases <- c("Latitude","lat","latitude","site_lat","lat_dd","latitude_dd","lat_deg","lat_dec","纬度","北纬","南纬")
cn_lower <- tolower(trimws(cn))
lat_hit  <- match(tolower(lat_aliases), cn_lower, nomatch = 0)
lat_hit  <- lat_hit[lat_hit != 0]
lat_col  <- if (length(lat_hit)) cn[lat_hit[1]] else NA_character_
if (is.na(lat_col)) {
  fidx <- grep("\\b(lat|latitude)\\b", cn_lower, perl = TRUE)
  if (length(fidx)) lat_col <- cn[fidx[1]]
}
if (is.na(lat_col)) stop("Latitude column not found.")

# data columns
flux_months <- c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
flux_cols   <- flux_months
sm_cols     <- paste0("SM_m", 1:12)
st_cols     <- paste0("ST_m", 1:12)

need_cols <- c(col_q01, col_q02, col_q03, col_q04, col_mgmt, lat_col, flux_cols, sm_cols, st_cols)
miss <- setdiff(need_cols, cn)
if (length(miss) > 0) stop("Missing columns: ", paste(miss, collapse = ", "))

# ------------------------- Filter + Hemisphere -------------------------
dat_f <- dat %>%
  dplyr::mutate(
    q01_chk = is_checked(.data[[col_q01]]),
    q02_chk = is_checked(.data[[col_q02]]),
    q03_chk = is_checked(.data[[col_q03]]),
    q04_chk = is_checked(.data[[col_q04]]),
    ctrl_ok = is_control(.data[[col_mgmt]]),
    lat_num = to_num(.data[[lat_col]]),
    Hemisphere = dplyr::case_when(
      is.na(lat_num) ~ NA_character_,
      lat_num >= 0   ~ "North",
      TRUE           ~ "South"
    )
  ) %>%
  dplyr::filter(!q01_chk, !q02_chk, !q03_chk, !q04_chk, ctrl_ok, !is.na(Hemisphere)) %>%
  dplyr::select(-q01_chk, -q02_chk, -q03_chk, -q04_chk, -ctrl_ok)

if (nrow(dat_f) == 0) stop("No data after filtering / hemisphere assignment.")

# ------------------------- Long data (>0) -------------------------
flux_long <- dat_f %>%
  dplyr::mutate(.row_id = dplyr::row_number()) %>%
  tidyr::pivot_longer(cols = dplyr::all_of(flux_cols),
                      names_to = "Month", values_to = "CH4_flux") %>%
  dplyr::select(.row_id, Month, CH4_flux, Hemisphere)

sm_long <- dat_f %>%
  dplyr::mutate(.row_id = dplyr::row_number()) %>%
  tidyr::pivot_longer(cols = dplyr::all_of(sm_cols),
                      names_to = "m", values_to = "SM") %>%
  dplyr::mutate(
    Month = paste0("m", readr::parse_number(m)),
    Month = dplyr::recode(Month,
                          m1="Jan", m2="Feb", m3="Mar", m4="Apr", m5="May", m6="Jun",
                          m7="Jul", m8="Aug", m9="Sep", m10="Oct", m11="Nov", m12="Dec")
  ) %>%
  dplyr::select(.row_id, Month, SM)

st_long <- dat_f %>%
  dplyr::mutate(.row_id = dplyr::row_number()) %>%
  tidyr::pivot_longer(cols = dplyr::all_of(st_cols),
                      names_to = "m", values_to = "ST") %>%
  dplyr::mutate(
    Month = paste0("m", readr::parse_number(m)),
    Month = dplyr::recode(Month,
                          m1="Jan", m2="Feb", m3="Mar", m4="Apr", m5="May", m6="Jun",
                          m7="Jul", m8="Aug", m9="Sep", m10="Oct", m11="Nov", m12="Dec")
  ) %>%
  dplyr::select(.row_id, Month, ST)

long_all <- flux_long %>%
  dplyr::left_join(sm_long, by = c(".row_id","Month")) %>%
  dplyr::left_join(st_long, by = c(".row_id","Month")) %>%
  dplyr::mutate(
    CH4_flux = to_num(CH4_flux),
    SM       = to_num(SM),
    ST       = to_num(ST),
    Month    = factor(Month, levels = flux_months)
  ) %>%
  dplyr::filter(!is.na(CH4_flux), CH4_flux > 0)

if (nrow(long_all) == 0) stop("No CH4_flux > 0 after filtering + parsing.")

# ------------------------- Labels & ranges -------------------------
ylab_CH4 <- expression(atop(bold(CH[4]~uptake), "(" * mu * "g CH"[4]*"-C m"^-2*" h"^-1*")"))
ylab_SM  <- "WFPS (%)"
ylab_ST  <- "T (°C)"

range_CH4_default <- c(-5, 120)
range_SM          <- c(0, 100)
range_ST          <- c(-15, 40)

# 统一色：北半球 & 南半球
hemi_cols <- c(North = "#3C78D8", South = "#E69138")

# ------------------------- Panel fn -------------------------
plot_box_panel <- function(df, var, y_lab = NULL, y_fix = NULL,
                           axes = c("none","x","y","xy"),
                           fill_col = "#999999",
                           base_size = 18, n_size = 4.6){
  axes <- match.arg(axes)
  var_sym <- rlang::sym(var)
  
  n_df <- df %>%
    dplyr::group_by(Month) %>%
    dplyr::summarise(n = sum(!is.na(!!var_sym)), .groups = "drop")
  
  if (!is.null(y_fix)) {
    n_df$ypos <- y_fix[1] + 0.001 * (y_fix[2] - y_fix[1])
  } else {
    vals <- df[[var]]
    if (all(is.na(vals))) {
      n_df$ypos <- 0
    } else {
      vmin <- min(vals, na.rm = TRUE)
      vmax <- max(vals, na.rm = TRUE)
      rng  <- vmax - vmin
      if (rng == 0) rng <- abs(vmin) + 1
      n_df$ypos <- vmin + 0.01 * rng
    }
  }
  
  p <- ggplot2::ggplot(df, ggplot2::aes(x = Month, y = !!var_sym)) +
    ggplot2::geom_boxplot(
      outlier.shape = 21, outlier.fill = NA, outlier.size = 2.1, outlier.stroke = 0.9,
      linewidth = 0.9, width = 0.55, fill = fill_col, color = "black", na.rm = TRUE
    ) +
    ggplot2::geom_text(
      data = n_df,
      mapping = ggplot2::aes(x = Month, y = ypos, label = n),
      inherit.aes = FALSE, size = n_size, vjust = 1
    ) +
    ggplot2::labs(x = NULL, y = NULL) +
    ggplot2::theme_minimal(base_size = base_size) +
    ggplot2::theme(
      panel.grid.minor = ggplot2::element_blank(),
      panel.grid.major = ggplot2::element_line(linewidth = 0.3),
      axis.text.x = ggplot2::element_text(size = base_size, margin = ggplot2::margin(t = 4)),
      axis.text.y = ggplot2::element_text(size = base_size),
      plot.margin = ggplot2::margin(8, 12, 8, 12),
      panel.background = ggplot2::element_rect(fill = "white", color = NA),
      plot.background  = ggplot2::element_rect(fill = "white", color = NA)
    )
  
  if (!is.null(y_fix)) {
    p <- p + ggplot2::scale_y_continuous(
      limits = y_fix,
      expand = ggplot2::expansion(mult = c(0.13, 0.02)),
      oob = scales::oob_squish
    )
  } else {
    p <- p + ggplot2::scale_y_continuous(
      expand = ggplot2::expansion(mult = c(0.13, 0.02)),
      oob = scales::oob_squish
    )
  }
  
  if (axes == "none") {
    p <- p + ggplot2::theme(
      axis.title.x = ggplot2::element_blank(), axis.title.y = ggplot2::element_blank(),
      axis.text.x  = ggplot2::element_blank(), axis.text.y  = ggplot2::element_blank(),
      axis.ticks.x = ggplot2::element_blank(), axis.ticks.y = ggplot2::element_blank()
    )
  }
  if (axes == "x") {
    p <- p + ggplot2::labs(x = "Month") +
      ggplot2::theme(axis.title.y = ggplot2::element_blank(),
                     axis.text.y  = ggplot2::element_blank(),
                     axis.ticks.y = ggplot2::element_blank())
  }
  if (axes == "y") {
    p <- p + ggplot2::labs(y = y_lab) +
      ggplot2::theme(axis.title.x = ggplot2::element_blank(),
                     axis.text.x  = ggplot2::element_blank(),
                     axis.ticks.x = ggplot2::element_blank())
  }
  if (axes == "xy") {
    p <- p + ggplot2::labs(x = "Month", y = y_lab)
  }
  
  p
}

# ------------------------- 小图（各半球；统一色 + 标题） -------------------------
make_hemi_triptych_noaxes <- function(dat_h, hemi){
  col_fill <- hemi_cols[[hemi]]
  title_txt <- if (hemi == "North") "All (Northern Hemisphere)" else "All (Southern Hemisphere)"
  
  p1 <- plot_box_panel(dat_h, "CH4_flux", y_fix = range_CH4_default, axes = "none", fill_col = col_fill)
  p2 <- plot_box_panel(dat_h, "SM",       y_fix = range_SM,          axes = "none", fill_col = col_fill)
  p3 <- plot_box_panel(dat_h, "ST",       y_fix = range_ST,          axes = "none", fill_col = col_fill)
  
  body  <- cowplot::plot_grid(p1, p2, p3, ncol = 1, align = "v", rel_heights = c(1,1,1))
  title <- cowplot::ggdraw() + cowplot::draw_label(title_txt, fontface = "bold", size = 18)
  cowplot::plot_grid(title, body, ncol = 1, rel_heights = c(0.1, 1))
}

# ------------------------- 合并大图（3×2；左右统一色） -------------------------
make_NS_combined <- function(dat_N, dat_S){
  col_N <- hemi_cols[["North"]]
  col_S <- hemi_cols[["South"]]
  
  # North column（左列 y，底行 x）
  p_ch4_N <- plot_box_panel(dat_N, "CH4_flux", y_lab = ylab_CH4, y_fix = range_CH4_default, axes = "y",  fill_col = col_N)
  p_sm_N  <- plot_box_panel(dat_N, "SM",       y_lab = ylab_SM,  y_fix = range_SM,          axes = "y",  fill_col = col_N)
  p_st_N  <- plot_box_panel(dat_N, "ST",       y_lab = ylab_ST,  y_fix = range_ST,          axes = "xy", fill_col = col_N)
  
  # South column（不显示 y；底行显示 x）
  p_ch4_S <- plot_box_panel(dat_S, "CH4_flux", y_lab = NULL, y_fix = range_CH4_default, axes = "none", fill_col = col_S)
  p_sm_S  <- plot_box_panel(dat_S, "SM",       y_lab = NULL, y_fix = range_SM,          axes = "none", fill_col = col_S)
  p_st_S  <- plot_box_panel(dat_S, "ST",       y_lab = NULL, y_fix = range_ST,          axes = "x",    fill_col = col_S)
  
  col_title_N <- cowplot::ggdraw() + cowplot::draw_label("All (Northern Hemisphere)", fontface = "bold", size = 18)
  col_title_S <- cowplot::ggdraw() + cowplot::draw_label("All (Southern Hemisphere)", fontface = "bold", size = 18)
  header <- cowplot::plot_grid(col_title_N, col_title_S, ncol = 2)
  
  row1 <- cowplot::plot_grid(p_ch4_N, p_ch4_S, ncol = 2, rel_widths = c(1,1), align = "h")
  row2 <- cowplot::plot_grid(p_sm_N,  p_sm_S,  ncol = 2, rel_widths = c(1,1), align = "h")
  row3 <- cowplot::plot_grid(p_st_N,  p_st_S,  ncol = 2, rel_widths = c(1,1), align = "h")
  
  grid <- cowplot::plot_grid(row1, row2, row3, ncol = 1, align = "v", rel_heights = c(1,1,1))
  cowplot::plot_grid(header, grid, ncol = 1, rel_heights = c(0.12, 1))
}

# ------------------------- 导出：两张小图 + 合并大图 -------------------------
hemi_data <- list()

for (hemi in c("North","South")) {
  dat_h <- long_all %>% dplyr::filter(Hemisphere == hemi)
  if (nrow(dat_h) == 0) { message("No data for hemisphere: ", hemi); next }
  
  g_small <- make_hemi_triptych_noaxes(dat_h, hemi)
  
  safe_save(
    file.path(out_dir, paste0("FigS5_", safe_filename(hemi), ".jpg")),
    g_small, width = 9, height = 12, dpi = 600, bg = "white"
  )
  
  hemi_data[[hemi]] <- dat_h
}

if (all(c("North","South") %in% names(hemi_data))) {
  g_comb <- make_NS_combined(hemi_data[["North"]], hemi_data[["South"]])
  
  safe_save(
    file.path(out_dir, "FigS5.jpg"),
    g_comb, width = 18, height = 12, dpi = 600, bg = "white"
  )
}

# ------------------------- Excel（各半球；sheet名缩短） -------------------------
for (hemi in c("North","South")) {
  dat_h <- hemi_data[[hemi]]
  if (is.null(dat_h) || nrow(dat_h) == 0) next
  
  flux_months <- levels(dat_h$Month)
  
  monthly_stats_tbl <- dat_h %>%
    dplyr::group_by(Month) %>%
    dplyr::summarise(
      n          = sum(!is.na(CH4_flux)),
      min_CH4    = if (all(is.na(CH4_flux))) NA_real_ else min(CH4_flux, na.rm = TRUE),
      q25_CH4    = if (all(is.na(CH4_flux))) NA_real_ else quantile(CH4_flux, 0.25, na.rm = TRUE, names = FALSE),
      median_CH4 = if (all(is.na(CH4_flux))) NA_real_ else median(CH4_flux, na.rm = TRUE),
      q75_CH4    = if (all(is.na(CH4_flux))) NA_real_ else quantile(CH4_flux, 0.75, na.rm = TRUE, names = FALSE),
      max_CH4    = if (all(is.na(CH4_flux))) NA_real_ else max(CH4_flux, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    dplyr::mutate(Month = factor(Month, levels = flux_months)) %>%
    dplyr::arrange(Month)
  
  all_monthly_tbl <- dat_h %>%
    dplyr::arrange(Month) %>%
    dplyr::select(Hemisphere, Month, CH4_flux, SM, ST)
  
  xlsx_file <- file.path(out_dir, paste0("CH4_stats_overall_", safe_filename(hemi), ".xlsx"))
  
  wb <- openxlsx::createWorkbook()
  
  add_sh <- function(name, df, freeze_row = 1) {
    name <- substr(name, 1, 31)
    openxlsx::addWorksheet(wb, name)
    openxlsx::writeData(wb, name, df)
    if (freeze_row >= 1) openxlsx::freezePane(wb, name, firstActiveRow = freeze_row + 1, firstActiveCol = 1)
    openxlsx::addFilter(wb, name, row = 1, cols = seq_len(ncol(df)))
    openxlsx::setColWidths(wb, name, cols = 1:ncol(df), widths = "auto")
    
    num_cols <- which(vapply(df, is.numeric, logical(1)))
    if (length(num_cols)) {
      fmt <- openxlsx::createStyle(numFmt = "0.00")
      openxlsx::addStyle(wb, name, style = fmt, rows = 2:(nrow(df)+1),
                         cols = num_cols, gridExpand = TRUE, stack = TRUE)
    }
  }
  
  add_sh("monthly_stats_CH4_all", monthly_stats_tbl)
  add_sh("all_rows_all",          all_monthly_tbl)
  
  openxlsx::saveWorkbook(wb, xlsx_file, overwrite = TRUE)
  message("Excel saved (", hemi, "): ", xlsx_file)
}

message("✅ Done. Outputs saved to: ", out_dir)