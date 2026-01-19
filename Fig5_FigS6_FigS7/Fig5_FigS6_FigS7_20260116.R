# make_ecosystem_vertical_boxplots_fig5_ranges_hemispheres.R
# ✅ 更新点（按你这条消息）：
# - 输入文件：Studies and Fluxes.xlsx（附件2）
# - 输入路径：D:/Users/jiaweiChiang/Desktop/Supplymentary_code
# - 输出路径：D:/Users/jiaweiChiang/Desktop/Supplymentary_code/Fig5_FigS6_FigS7
# - 列名更新：
#     laboratory -> Q01
#     repeat     -> Q02
#     dataquality-> Q03
#     Q7         -> Q04
#     control    -> management
#
# 其他功能保持：
# - Excel 原始行数据拆成 CH4/SM/ST 3 个 sheet
# - 拼图：每列“最后一个非空行”的 ST 面板显示 Jan–Dec；
#        如果该列全空（没有生态系统），则该列第一行 ST 显示 Jan–Dec
# - 南半球：固定 6 个生态系统，3×2 拼图
# - 北半球：Tundra 强制保留 1–12 月空格
# - 拼图：CH4/SM 不显示横坐标；只有 ST 按规则显示 Jan–Dec
# - Excel：南北半球分别输出，每生态系统×月导出 CH4/SM/ST 的统计

suppressPackageStartupMessages({
  req <- c("readxl","dplyr","stringr","tidyr","ggplot2",
           "purrr","cowplot","readr","rlang","scales",
           "openxlsx","writexl")
  for (p in req) if (!requireNamespace(p, quietly=TRUE)) install.packages(p)
  invisible(lapply(req, library, character.only = TRUE))
})

# --- Global typography knobs ---
eco_label_size <- 28
af_label_size  <- 30
label_band_h   <- 0.10

# n 标签字号
n_size_default <- 6.6
n_size_big     <- 6.6

# -------------------------
# Paths (UPDATED)
# -------------------------
in_dir   <- "D:/Users/jiaweiChiang/Desktop/Supplymentary_code"
in_file  <- file.path(in_dir, "Studies and Fluxes.xlsx")

out_dir  <- file.path(in_dir, "Fig5_FigS6_FigS7")
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE, showWarnings = FALSE)
setwd(out_dir)

stopifnot(file.exists(in_file))

# -------------------------
# Helpers
# -------------------------
safe_save <- function(path, plot, width, height, dpi=600, bg="white"){
  dir.create(dirname(path), recursive = TRUE, showWarnings = FALSE)
  ggplot2::ggsave(filename = path, plot = plot, width = width, height = height,
                  dpi = dpi, bg = bg, device = grDevices::jpeg, limitsize = FALSE)
}

norm_txt <- function(x) {
  if (is.null(x)) return(NA_character_)
  as.character(x) %>% stringr::str_trim() %>% stringr::str_squish()
}

is_checked <- function(x){
  y <- norm_txt(x)
  !is.na(y) & stringr::str_detect(stringr::str_to_lower(y), "^(√|✓|✔|yes|y|1|true)$")
}

is_control <- function(x){
  y <- norm_txt(x)
  !is.na(y) & stringr::str_detect(stringr::str_to_lower(y), "^control(?:\\s*group)?$")
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
  if (length(hit) == 0) stop(sprintf("Column not found: %s (check template headers)", target))
  hit[1]
}

to_num <- function(x) suppressWarnings(as.numeric(readr::parse_number(as.character(x))))

# -------------------------
# Read (prefer sheet "row")
# -------------------------
sheets <- readxl::excel_sheets(in_file)
sheet_to_read <- if ("row" %in% sheets) "row" else sheets[1]
dat <- readxl::read_excel(in_file, sheet = sheet_to_read)

# -------------------------
# Column checks (UPDATED)
# -------------------------
cn <- names(dat)

col_lab  <- colmap_exact("Q01", cn)
col_rep  <- colmap_exact("Q02", cn)
col_dq   <- colmap_exact("Q03", cn)
col_q07  <- colmap_exact("Q04", cn)
col_ctrl <- colmap_exact("management", cn)

# latitude auto-detect (新版通常是 Latitude)
lat_aliases <- c("Latitude","latitude","lat","site_lat","lat_dd","latitude_dd","lat_deg","lat_dec","纬度","北纬","南纬")
cn_lower <- tolower(trimws(cn))
lat_hit <- match(tolower(lat_aliases), cn_lower, nomatch = 0)
lat_hit <- lat_hit[lat_hit != 0]
lat_col <- if (length(lat_hit)) cn[lat_hit[1]] else NA_character_
if (is.na(lat_col)) {
  fidx <- grep("\\b(lat|latitude)\\b", cn_lower, perl = TRUE)
  if (length(fidx)) lat_col <- cn[fidx[1]]
}
if (is.na(lat_col)) stop("Latitude column not found. Please ensure a column like 'Latitude' exists.")

# data columns (新版 CH4 直接是 Jan–Dec)
flux_months <- c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
flux_cols   <- flux_months
sm_cols     <- paste0("SM_m", 1:12)
st_cols     <- paste0("ST_m", 1:12)

need_cols   <- c(col_lab, col_rep, col_dq, col_q07, col_ctrl, lat_col, flux_cols, sm_cols, st_cols)
miss <- setdiff(need_cols, cn)
if (length(miss) > 0) stop("Missing columns: ", paste(miss, collapse = ", "))

# -------------------------
# ✅ Ecosystem column (UPDATED: prefer Biome first)
#   优先级：Biome -> Biomes -> Boimes (兼容拼写)
# -------------------------
eco_col <- NULL
preferred_eco <- c("Biome", "Biomes", "Boimes")   # ✅ 你最新是 Biome

# 1) 优先精确匹配（不区分大小写）
for (nm in preferred_eco) {
  if (tolower(nm) %in% tolower(cn)) {
    eco_col <- cn[match(tolower(nm), tolower(cn))]
    message("Using ecosystem column: ", eco_col, " (preferred match: ", nm, ")")
    break
  }
}

# 2) 如果没有精确命中，则走原来的模糊逻辑
if (is.null(eco_col)) {
  cn_trim  <- trimws(cn)
  cn_lower <- tolower(cn_trim)
  
  eco_aliases <- c(
    "biome","biomes","boimes",
    "ecosystem","ecosystem_english","ecosystem (english)","ecosystem type","ecosystem_type",
    "ecosystem category","ecosystem_category","ecosys","ecosys_type",
    "biome_name","biome type","biome_type","biome_category",
    "ecosystem_eng","ecosystem_en","ecosystem_cn","ecosystem_chinese",
    "ecosystemclass","ecosystem_class","ecotype","eco_type",
    "生态系统","生态系统英文","生态系统（英文）","生态系统类型","生态类型","生态系统类别",
    "生物群系","生态类群","生态系统_英文","生态系统_中文","生态系统_类别"
  )
  
  eco_hit_idx <- match(eco_aliases, cn_lower, nomatch = 0)
  eco_hit_idx <- eco_hit_idx[eco_hit_idx != 0]
  
  choose_col_by_idx <- function(idx) if (length(idx) >= 1) cn[idx[1]] else NULL
  eco_col <- choose_col_by_idx(eco_hit_idx)
  
  if (is.null(eco_col)) {
    pat <- "\\b(ecosystem|biome|biomes|boimes|ecosys)\\b"
    fuzzy_idx <- which(grepl(pat, cn_lower, perl = TRUE))
    eco_col <- choose_col_by_idx(fuzzy_idx)
  }
  
  if (is.null(eco_col)) {
    message("⚠ Ecosystem column not found. Using temporary Ecosystem='Unknown'.")
    dat$Ecosystem <- "Unknown"
    eco_col <- "Ecosystem"
  } else {
    message(sprintf("Using ecosystem column: %s (fuzzy match)", eco_col))
  }
}

# -------------------------
# Filter (remove √; keep control) + Hemisphere
# -------------------------
dat_f <- dat %>%
  dplyr::mutate(
    lab_chk = is_checked(.data[[col_lab]]),
    rep_chk = is_checked(.data[[col_rep]]),
    dq_chk  = is_checked(.data[[col_dq]]),
    q7_chk  = is_checked(.data[[col_q07]]),
    ctrl_ok = is_control(.data[[col_ctrl]]),
    lat_num = to_num(.data[[lat_col]]),
    Hemisphere = dplyr::case_when(
      is.na(lat_num) ~ NA_character_,
      lat_num >= 0   ~ "North",
      TRUE           ~ "South"
    )
  ) %>%
  dplyr::filter(!lab_chk, !rep_chk, !dq_chk, !q7_chk, ctrl_ok, !is.na(Hemisphere)) %>%
  dplyr::select(-lab_chk, -rep_chk, -dq_chk, -q7_chk, -ctrl_ok)

if (nrow(dat_f) == 0) stop("No data after filtering / hemisphere assignment. Check flags and latitude.")

# -------------------------
# Long data; keep CH4_flux > 0
# -------------------------
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
  dplyr::left_join(
    dat_f %>%
      dplyr::mutate(.row_id = dplyr::row_number()) %>%
      dplyr::select(.row_id, Ecosystem = dplyr::all_of(eco_col)),
    by = ".row_id"
  ) %>%
  dplyr::mutate(
    CH4_flux  = to_num(CH4_flux),
    SM        = to_num(SM),
    ST        = to_num(ST),
    Ecosystem = dplyr::if_else(is.na(Ecosystem) | Ecosystem=="",
                               "Unknown", as.character(Ecosystem)),
    Month     = factor(Month, levels = flux_months)
  ) %>%
  dplyr::filter(!is.na(CH4_flux), CH4_flux > 0)

# Fixed ecosystem order
eco_levels <- c(
  "Tundra","Boreal forest","Temperate seasonal forest","Temperate rainforest",
  "Temperate grassland","Shrubland","Woodland","Savanna","Tropical seasonal forest",
  "Tropical rainforest","Subtropical desert","Desert","Alpine","Wetland",
  "Agriculture","Urban","Bare","Others","Unknown"
)
long_all <- long_all %>%
  dplyr::mutate(Ecosystem = factor(Ecosystem, levels = eco_levels)) %>%
  dplyr::arrange(Hemisphere, Ecosystem, Month)

# -------------------------
# Axis labels
# -------------------------
ylab_CH4 <- expression(atop(bold(CH[4]~uptake),
                            "(" * mu * "g CH"[4]*"-C m"^-2*" h"^-1*")"))
ylab_SM  <- "WFPS (%)"
ylab_ST  <- "T (°C)"

# -------------------------
# Ecosystem colors
# -------------------------
eco_pal <- c(
  "Tundra"="#a7d3df","Boreal forest"="#4f9f4f","Temperate seasonal forest"="#8cbf70",
  "Temperate rainforest"="#e8c8a1","Tropical rainforest"="#c49b50","Tropical seasonal forest"="#2e7032",
  "Savanna"="#5fa6bc","Subtropical desert"="#f2c300","Temperate grassland"="#2e6f89",
  "Desert"="#fff2b2","Woodland"="#cd7a50","Shrubland"="#a34030",
  "Alpine"="#b27874","Agriculture"="#886b99","Bare"="#357266",
  "Wetland"="#69778c","Urban"="#6b575a","Others"="#b3b5ae","Unknown"="#b3b5ae"
)
get_eco_col <- function(eco) if (!is.null(eco_pal[[eco]])) eco_pal[[eco]] else eco_pal[["Unknown"]]

# -------------------------
# Plot function
# -------------------------
plot_box_panel <- function(df, var, y_lab, y_fix = NULL,
                           axes = c("none","x","y","xy"),
                           fill_col = "#999999",
                           add_title = NULL,
                           base_size = 18, title_size = 20,
                           n_size = n_size_default,
                           base_family = "sans") {
  axes <- match.arg(axes)
  var_sym <- rlang::sym(var)
  
  # n labels: 只给有数据的月份（n>0）标数字；缺月不显示“0”
  n_df <- df %>%
    dplyr::group_by(Month) %>%
    dplyr::summarise(n = sum(!is.na(!!var_sym)), .groups = "drop") %>%
    dplyr::filter(n > 0)
  
  if (nrow(n_df) > 0) {
    if (!is.null(y_fix)) {
      n_df$ypos <- y_fix[1] + 0.02 * (y_fix[2] - y_fix[1])
    } else {
      vals <- df[[var]]
      if (all(is.na(vals))) {
        n_df$ypos <- 0
      } else {
        vmin <- min(vals, na.rm = TRUE); vmax <- max(vals, na.rm = TRUE)
        rng  <- vmax - vmin; if (rng == 0) rng <- abs(vmin) + 1
        n_df$ypos <- vmin + 0.02 * rng
      }
    }
  }
  
  p <- ggplot2::ggplot(df, ggplot2::aes(x = Month, y = !!var_sym)) +
    ggplot2::geom_boxplot(
      outlier.shape = 21, outlier.fill = NA,
      outlier.size = 2.1, outlier.stroke = 0.9,
      linewidth = 0.9, width = 0.55, fill = fill_col, color = "black",
      na.rm = TRUE
    )
  
  if (nrow(n_df) > 0) {
    p <- p + ggplot2::geom_text(
      data = n_df,
      ggplot2::aes(x = Month, y = ypos, label = n),
      inherit.aes = FALSE,
      size = n_size,
      vjust = 1
    )
  }
  
  p <- p +
    ggplot2::labs(x = NULL, y = NULL, title = add_title) +
    ggplot2::theme_minimal(base_size = base_size) +
    ggplot2::theme(
      panel.grid.minor = ggplot2::element_blank(),
      panel.grid.major = ggplot2::element_line(linewidth = 0.3),
      axis.text.x = ggplot2::element_text(size = base_size,
                                          margin = ggplot2::margin(t = 4)),
      axis.text.y = ggplot2::element_text(size = base_size),
      plot.title  = ggplot2::element_text(hjust = 0.5, face = "bold",
                                          size = title_size),
      plot.margin = ggplot2::margin(8, 12, 8, 12),
      panel.background = ggplot2::element_rect(fill = "white", color = NA),
      plot.background  = ggplot2::element_rect(fill = "white", color = NA),
      text = ggplot2::element_text(family = base_family)
    ) +
    ggplot2::scale_x_discrete(drop = FALSE, limits = flux_months)
  
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
      axis.title.x = ggplot2::element_blank(),
      axis.title.y = ggplot2::element_blank(),
      axis.text.x  = ggplot2::element_blank(),
      axis.text.y  = ggplot2::element_blank(),
      axis.ticks.x = ggplot2::element_blank(),
      axis.ticks.y = ggplot2::element_blank()
    )
  } else if (axes == "x") {
    p <- p + ggplot2::labs(x = "Month") +
      ggplot2::theme(
        axis.title.y = ggplot2::element_blank(),
        axis.text.y  = ggplot2::element_blank(),
        axis.ticks.y = ggplot2::element_blank()
      )
  } else if (axes == "y") {
    p <- p + ggplot2::labs(y = y_lab) +
      ggplot2::theme(
        axis.title.x = ggplot2::element_blank(),
        axis.text.x  = ggplot2::element_blank(),
        axis.ticks.x = ggplot2::element_blank()
      )
  } else if (axes == "xy") {
    p <- p + ggplot2::labs(x = "Month", y = y_lab)
  }
  
  p
}

# -------------------------
# Ranges
# -------------------------
range_CH4_default <- c(-5, 120)
range_SM          <- c(0, 100)
range_ST          <- c(-15, 40)
eco_limits <- list(
  "Temperate rainforest" = c(-5, 330),
  "Bare"                 = c(-5, 380),
  "Shrubland"            = c(-5, 300)
)

# -------------------------
# Montage builder（ST 规则：每列最后一个非空行显示 Jan–Dec；该列全空则第一行显示）
# -------------------------
build_montage_grid <- function(dat_h, ecos_vec, ncol, nrow, label_start_letter, outfile,
                               width, height,
                               big_label_ecos = NULL,
                               first_col_rel_width = 1.0,
                               show_x_on_ST = TRUE) {
  stopifnot(ncol >= 1, nrow >= 1)
  n_target <- ncol * nrow
  
  filled_names <- rep(NA_character_, n_target)
  if (length(ecos_vec) > 0) {
    filled_names[seq_len(min(length(ecos_vec), n_target))] <-
      ecos_vec[seq_len(min(length(ecos_vec), n_target))]
  }
  filled_names <- as.character(filled_names)
  filled_names[!nzchar(trimws(filled_names))] <- NA_character_
  
  filled_mat <- matrix(filled_names, nrow = nrow, ncol = ncol, byrow = TRUE)
  st_x_row_by_col <- vapply(seq_len(ncol), function(j) {
    rr <- which(!is.na(filled_mat[, j]))
    if (length(rr) == 0) 1L else max(rr)   # ✅ 全空列：第一行显示横坐标
  }, integer(1))
  
  make_alpha_labels <- function(k, len) {
    base <- c(LETTERS, outer(LETTERS, LETTERS, paste0))
    base[seq.int(k, k + len - 1)]
  }
  letter_index <- function(ch) {
    letters_seq <- c(LETTERS, outer(LETTERS, LETTERS, paste0))
    match(ch, letters_seq)
  }
  
  lab_vec <- rep("", n_target)
  start_idx <- letter_index(label_start_letter)
  if (any(!is.na(filled_names))) {
    letters_needed <- sum(!is.na(filled_names))
    seq_letters <- make_alpha_labels(start_idx, letters_needed)
    lab_vec[!is.na(filled_names)] <- seq_letters
  }
  
  cell_plot <- vector("list", n_target)
  
  for (i in seq_len(n_target)) {
    eco_name <- filled_names[[i]]
    col_i    <- ((i - 1) %% ncol) + 1
    row_i    <- ((i - 1) %/% ncol) + 1
    is_left  <- (col_i == 1)
    
    if (is.na(eco_name)) {
      title_band <- cowplot::ggdraw() +
        ggplot2::theme(plot.background = ggplot2::element_rect(fill = "white", color = NA))
      body_blank <- cowplot::ggdraw() + ggplot2::theme_void()
      cell_plot[[i]] <- cowplot::plot_grid(title_band, body_blank, ncol = 1,
                                           rel_heights = c(label_band_h, 1))
      next
    }
    
    sub <- dat_h %>% dplyr::filter(Ecosystem == eco_name)
    eco_chr <- as.character(eco_name)
    
    range_CH4_eco <- eco_limits[[eco_chr]]
    if (is.null(range_CH4_eco)) range_CH4_eco <- range_CH4_default
    eco_col_fill <- get_eco_col(eco_chr)
    
    use_n_size <- if (!is.null(big_label_ecos) && eco_chr %in% big_label_ecos) n_size_big else n_size_default
    axes_y <- if (is_left) "y" else "none"
    
    axes_ch4 <- axes_y
    axes_sm  <- axes_y
    
    axes_st <- axes_y
    show_x_here <- isTRUE(show_x_on_ST) &&
      !is.na(st_x_row_by_col[col_i]) &&
      row_i == st_x_row_by_col[col_i]
    if (show_x_here) axes_st <- if (is_left) "xy" else "x"
    
    p_ch4 <- plot_box_panel(sub, "CH4_flux", y_lab = ylab_CH4,
                            y_fix = range_CH4_eco, axes = axes_ch4,
                            fill_col = eco_col_fill, n_size = use_n_size)
    p_sm  <- plot_box_panel(sub, "SM", y_lab = ylab_SM,
                            y_fix = range_SM, axes = axes_sm,
                            fill_col = eco_col_fill, n_size = use_n_size)
    p_st  <- plot_box_panel(sub, "ST", y_lab = ylab_ST,
                            y_fix = range_ST, axes = axes_st,
                            fill_col = eco_col_fill, n_size = use_n_size)
    
    body <- cowplot::plot_grid(p_ch4, p_sm, p_st, ncol = 1, align = "v", rel_heights = c(1,1,1))
    body <- cowplot::ggdraw(body) +
      ggplot2::theme(plot.background = ggplot2::element_rect(fill = "white", color = NA))
    
    title_band <- cowplot::ggdraw() +
      ggplot2::theme(plot.background = ggplot2::element_rect(fill = "white", color = NA)) +
      cowplot::draw_text(eco_chr, x = 0.5, y = 0.5,
                         hjust = 0.5, vjust = 0.5,
                         fontface = "bold", size = eco_label_size)
    
    cell_plot[[i]] <- cowplot::plot_grid(title_band, body, ncol = 1, rel_heights = c(label_band_h, 1))
  }
  
  relw <- rep(1, ncol)
  relw[1] <- first_col_rel_width
  
  grid <- cowplot::plot_grid(
    plotlist = cell_plot, ncol = ncol, align = "hv",
    labels = lab_vec, label_size = af_label_size,
    label_fontface = "bold",
    label_x = 0.012, label_y = 0.988,
    hjust = 0, vjust = 1,
    rel_widths = relw
  )
  
  safe_save(outfile, grid, width = width, height = height, dpi = 600, bg = "white")
  message("Exported montage: ", outfile)
}

# -------------------------
# Excel stats helper
# -------------------------
summ_stats_monthly <- function(dat_h, var, prefix) {
  v <- rlang::sym(var)
  dat_h %>%
    dplyr::group_by(Ecosystem, Month) %>%
    dplyr::summarise(
      n      = sum(!is.na(!!v)),
      min    = if (all(is.na(!!v))) NA_real_ else min(!!v, na.rm = TRUE),
      q25    = if (all(is.na(!!v))) NA_real_ else as.numeric(stats::quantile(!!v, 0.25, na.rm = TRUE, names = FALSE)),
      median = if (all(is.na(!!v))) NA_real_ else stats::median(!!v, na.rm = TRUE),
      q75    = if (all(is.na(!!v))) NA_real_ else as.numeric(stats::quantile(!!v, 0.75, na.rm = TRUE, names = FALSE)),
      max    = if (all(is.na(!!v))) NA_real_ else max(!!v, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    dplyr::mutate(Month = factor(Month, levels = flux_months)) %>%
    dplyr::arrange(Ecosystem, Month) %>%
    dplyr::rename(
      !!paste0("n_", prefix)      := n,
      !!paste0("min_", prefix)    := min,
      !!paste0("q25_", prefix)    := q25,
      !!paste0("median_", prefix) := median,
      !!paste0("q75_", prefix)    := q75,
      !!paste0("max_", prefix)    := max
    )
}

# ✅ 原始行数据拆成 3 张 sheet（CH4/SM/ST）
write_excel_per_hemi <- function(hemi, dat_h, out_dir) {
  stats_ch4 <- summ_stats_monthly(dat_h, "CH4_flux", "CH4")
  stats_sm  <- summ_stats_monthly(dat_h, "SM",       "SM")
  stats_st  <- summ_stats_monthly(dat_h, "ST",       "ST")
  
  all_rows_ch4 <- dat_h %>%
    dplyr::arrange(Ecosystem, Month) %>%
    dplyr::select(Hemisphere, Ecosystem, Month, CH4_flux)
  
  all_rows_sm <- dat_h %>%
    dplyr::arrange(Ecosystem, Month) %>%
    dplyr::select(Hemisphere, Ecosystem, Month, SM)
  
  all_rows_st <- dat_h %>%
    dplyr::arrange(Ecosystem, Month) %>%
    dplyr::select(Hemisphere, Ecosystem, Month, ST)
  
  xlsx_file <- file.path(out_dir, paste0("Monthly_stats_CH4_SM_ST_", safe_filename(hemi), ".xlsx"))
  
  wb <- openxlsx::createWorkbook()
  add_sh <- function(name, df, freeze_row = 1) {
    openxlsx::addWorksheet(wb, name)
    openxlsx::writeData(wb, name, df)
    if (freeze_row >= 1)
      openxlsx::freezePane(wb, name, firstActiveRow = freeze_row + 1, firstActiveCol = 1)
    openxlsx::addFilter(wb, name, row = 1, cols = seq_len(ncol(df)))
    openxlsx::setColWidths(wb, name, cols = 1:ncol(df), widths = "auto")
    num_cols <- which(vapply(df, is.numeric, logical(1)))
    if (length(num_cols)) {
      fmt <- openxlsx::createStyle(numFmt = "0.00")
      openxlsx::addStyle(wb, name, style = fmt,
                         rows = 2:(nrow(df)+1), cols = num_cols,
                         gridExpand = TRUE, stack = TRUE)
    }
  }
  
  add_sh("monthly_stats_CH4", stats_ch4)
  add_sh("monthly_stats_SM",  stats_sm)
  add_sh("monthly_stats_ST",  stats_st)
  add_sh("all_rows_CH4",      all_rows_ch4)
  add_sh("all_rows_SM",       all_rows_sm)
  add_sh("all_rows_ST",       all_rows_st)
  
  openxlsx::saveWorkbook(wb, xlsx_file, overwrite = TRUE)
  message("Excel saved (", hemi, "): ", xlsx_file)
}

# -------------------------
# === Loop per Hemisphere ===
# -------------------------
hemis <- c("North","South")

# 南半球固定要画的 6 个生态系统（顺序固定）
south_target6 <- c(
  "Temperate grassland","Temperate seasonal forest","Agriculture",
  "Temperate rainforest","Tropical seasonal forest","Savanna"
)

for (hemi in hemis) {
  message("==== Hemisphere: ", hemi, " ====")
  dat_h <- long_all %>% dplyr::filter(Hemisphere == hemi)
  if (nrow(dat_h) == 0) { message("No data for hemisphere: ", hemi); next }
  
  ecos_present <- dat_h %>%
    dplyr::distinct(Ecosystem) %>%
    dplyr::filter(!is.na(Ecosystem)) %>%
    dplyr::pull() %>%
    as.character()
  
  # 单生态系统竖排图
  ecos_for_single <- if (hemi == "South") south_target6 else ecos_present
  
  for (eco in ecos_for_single) {
    sub <- dat_h %>% dplyr::filter(Ecosystem == eco)
    if (nrow(sub) == 0) next
    
    eco_chr <- as.character(trimws(eco))
    range_CH4_eco <- eco_limits[[eco_chr]]
    if (is.null(range_CH4_eco)) range_CH4_eco <- range_CH4_default
    eco_col_fill <- get_eco_col(eco_chr)
    
    p1 <- plot_box_panel(sub, "CH4_flux", y_lab = ylab_CH4,
                         y_fix = range_CH4_eco, axes = "none",
                         fill_col = eco_col_fill,
                         add_title = paste0(eco_chr, " (", hemi, ")"))
    p2 <- plot_box_panel(sub, "SM", y_lab = ylab_SM,
                         y_fix = range_SM, axes = "none",
                         fill_col = eco_col_fill)
    p3 <- plot_box_panel(sub, "ST", y_lab = ylab_ST,
                         y_fix = range_ST, axes = "none",
                         fill_col = eco_col_fill)
    
    g <- cowplot::plot_grid(p1, p2, p3, ncol = 1, align = "v", rel_heights = c(1,1,1))
    g <- cowplot::ggdraw(g) +
      ggplot2::theme(plot.background = ggplot2::element_rect(fill = "white", color = NA))
    
    outfile <- file.path(out_dir,
                         paste0(safe_filename(hemi), "_", safe_filename(eco_chr),
                                "_vertical_boxplot_fixedranges_CH4_SM_ST.jpg"))
    safe_save(outfile, g, width = 7.5, height = 10, dpi = 600, bg = "white")
  }
  
  # 拼图输出
  if (hemi == "South") {
    file_south_6 <- file.path(out_dir, paste0("FigS7_", safe_filename(hemi), ".jpg"))
    build_montage_grid(
      dat_h = dat_h,
      ecos_vec = south_target6,
      ncol = 3, nrow = 2,
      label_start_letter = "A",
      outfile = file_south_6,
      width = 30, height = 20,
      big_label_ecos = NULL,
      first_col_rel_width = 1.0,
      show_x_on_ST = TRUE
    )
  } else {
    target5 <- c("Temperate grassland","Tropical rainforest","Boreal forest",
                 "Tropical seasonal forest","Savanna")
    fig1_5wide <- target5[target5 %in% ecos_present]
    
    file_fig1_5wide <- file.path(out_dir, paste0("Fig5_", safe_filename(hemi), ".jpg"))
    build_montage_grid(
      dat_h = dat_h,
      ecos_vec = fig1_5wide,
      ncol = 5, nrow = 1,
      label_start_letter = "A",
      outfile = file_fig1_5wide,
      width = 40, height = 16,
      big_label_ecos = fig1_5wide,
      first_col_rel_width = 1.18,
      show_x_on_ST = TRUE
    )
    
    rest_north <- setdiff(ecos_present, fig1_5wide)
    rest_north <- setdiff(rest_north, c("Wetland","Others"))
    if (length(rest_north) > 0) {
      rest_north <- head(rest_north, 11)
      
      rest_north_6 <- head(rest_north, 6)
      if (length(rest_north_6) > 0) {
        file_north_grid_6 <- file.path(out_dir, paste0("FigS6_1_", safe_filename(hemi), ".jpg"))
        build_montage_grid(
          dat_h = dat_h,
          ecos_vec = rest_north_6,
          ncol = 3, nrow = 2,
          label_start_letter = "F",
          outfile = file_north_grid_6,
          width = 30, height = 20,
          big_label_ecos = NULL,
          first_col_rel_width = 1.0,
          show_x_on_ST = TRUE
        )
      }
      
      if (length(rest_north) > 6) {
        rest_north_5 <- rest_north[7:min(11, length(rest_north))]
        if (length(rest_north_5) > 0) {
          file_north_grid_5 <- file.path(out_dir, paste0("FigS6_2_", safe_filename(hemi), ".jpg"))
          build_montage_grid(
            dat_h = dat_h,
            ecos_vec = rest_north_5,
            ncol = 3, nrow = 2,
            label_start_letter = "L",
            outfile = file_north_grid_5,
            width = 30, height = 20,
            big_label_ecos = NULL,
            first_col_rel_width = 1.0,
            show_x_on_ST = TRUE
          )
        }
      }
    }
  }
  
  # Excel outputs
  write_excel_per_hemi(hemi, dat_h, out_dir)
}

message("✅ Done. Used sheet: ", sheet_to_read)
