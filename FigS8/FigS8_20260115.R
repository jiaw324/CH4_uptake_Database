# ================================
# Monthly CH4 median tiles + mosaic (smooth) with SD & CV annotations
# —— 以 Biomes 列分类；分南/北半球；剔除 Wetland 与 Others
# 仅对：Tundra (South) 与 Boreal forest (South) 取消平滑线（只画点）
#
# ✅ Updated for NEW headers:
#   Q01 Q02 Q03 Q04 + management + Latitude + Biomes + Jan..Dec
# ✅ Input:
#   Studies and Fluxes.xlsx
# ✅ Output:
#   D:/Users/jiaweiChiang/Desktop/Supplymentary_code/FigS8
# ================================

suppressPackageStartupMessages({
  req <- c("readxl","dplyr","tidyr","stringr","ggplot2","cowplot",
           "writexl","glue","purrr","rlang")
  to_install <- req[!sapply(req, requireNamespace, quietly = TRUE)]
  if (length(to_install)) install.packages(to_install)
  invisible(lapply(req, library, character.only = TRUE))
})

# -------------------------
# 路径（按你要求改）
# -------------------------
root_dir <- "D:/Users/jiaweiChiang/Desktop/Supplymentary_code"

# 你说的输入文件1（这里不强制使用，但保留检查）
raw_file1 <- file.path(root_dir, "CH4 uptake data_CH4 FLUX_1_1410.xlsx")
if (!file.exists(raw_file1)) {
  warning("找不到输入文件：CH4 uptake data_CH4 FLUX_1_1410.xlsx（但本脚本绘图主要使用 Studies and Fluxes.xlsx）")
}

# 你说的模板（这里作为主数据源）
raw_file <- file.path(root_dir, "Studies and Fluxes.xlsx")
stopifnot(file.exists(raw_file))

# 输出文件夹（按你要求改）
work_dir <- file.path(root_dir, "FigS8")
if (!dir.exists(work_dir)) dir.create(work_dir, recursive = TRUE)

# -------------------------
# 配色与顺序（与 Biomes 名称一致）
# -------------------------
eco_colors <- c(
  "Tundra"="#a7d3df","Boreal forest"="#4f9f4f","Temperate seasonal forest"="#8cbf70",
  "Temperate rainforest"="#e8c8a1","Tropical rainforest"="#c49b50","Tropical seasonal forest"="#2e7032",
  "Savanna"="#5fa6bc","Subtropical desert"="#f2c300","Temperate grassland"="#2e6f89",
  "Desert"="#fff2b2","Woodland"="#cd7a50","Shrubland"="#a34030",
  "Alpine"="#b27874","Agriculture"="#886b99","Bare"="#357266",
  "Wetland"="#69778c","Urban"="#6b575a","Others"="#b3b5ae","Unknown"="#b3b5ae"
)
eco_levels <- names(eco_colors)

# -------------------------
# 列名匹配 + 判定函数
# -------------------------
pick_col <- function(nms, candidates, must = TRUE, label = "column") {
  nms_l <- tolower(nms)
  for (cand in candidates) {
    hit <- which(nms_l == tolower(cand))
    if (length(hit) == 1) return(nms[hit])
  }
  for (cand in candidates) {
    pat <- paste0("^(?:", cand, ")$|\\b", cand, "\\b")
    hit <- which(grepl(pat, nms_l, perl = TRUE))
    if (length(hit) > 0) return(nms[hit[1]])
  }
  if (must)
    stop(sprintf("找不到【%s】列。表头示例：%s",
                 label, paste(head(nms, 40), collapse = ", ")), call. = FALSE)
  NA_character_
}

norm_txt <- function(x) { if (is.null(x)) return(NA_character_); stringr::str_squish(as.character(x)) }
is_checked <- function(x){
  y <- stringr::str_to_lower(norm_txt(x))
  !is.na(y) & stringr::str_detect(y, "^(√|✓|✔|yes|y|1|true|t)$")
}

# ✅ management == control（不区分大小写，trim）
is_control <- function(x){
  y <- stringr::str_to_lower(norm_txt(x))
  !is.na(y) & stringr::str_detect(y, "^control$")
}

# -------------------------
# 读取原始模板（优先 row）
# -------------------------
sheets_raw <- readxl::excel_sheets(raw_file)
sheet_raw  <- if ("row" %in% sheets_raw) "row" else sheets_raw[1]

raw0 <- readxl::read_excel(raw_file, sheet = sheet_raw)
names(raw0) <- trimws(names(raw0))
cols <- names(raw0)

# --- 强制使用 Biomes 列 ---
biomes_col <- pick_col(cols, c("Biomes","biomes","Biome","biome"), must = TRUE, label = "Biomes")

# --- ✅ 新列名：Q01/Q02/Q03/Q04/management/Latitude ---
q01_col <- pick_col(cols, c("Q01","q01"), TRUE, "Q01")
q02_col <- pick_col(cols, c("Q02","q02"), TRUE, "Q02")
q03_col <- pick_col(cols, c("Q03","q03"), TRUE, "Q03")
q04_col <- pick_col(cols, c("Q04","q04"), TRUE, "Q04")

mgmt_col <- pick_col(cols, c("management","Management"), TRUE, "management")
lat_col  <- pick_col(cols, c("Latitude","latitude","lat","site_latitude"), TRUE, "Latitude")

# --- ✅ 月份列：Jan..Dec ---
col_Jan <- pick_col(cols, c("Jan","January","^1$"),  TRUE, "Jan")
col_Feb <- pick_col(cols, c("Feb","February","^2$"), TRUE, "Feb")
col_Mar <- pick_col(cols, c("Mar","March","^3$"),    TRUE, "Mar")
col_Apr <- pick_col(cols, c("Apr","April","^4$"),    TRUE, "Apr")
col_May <- pick_col(cols, c("May","^5$"),            TRUE, "May")
col_Jun <- pick_col(cols, c("Jun","June","^6$"),     TRUE, "Jun")
col_Jul <- pick_col(cols, c("Jul","July","^7$"),     TRUE, "Jul")
col_Aug <- pick_col(cols, c("Aug","August","^8$"),   TRUE, "Aug")
col_Sep <- pick_col(cols, c("Sep","Sept","September","^9$"), TRUE, "Sep")
col_Oct <- pick_col(cols, c("Oct","October","^10$"), TRUE, "Oct")
col_Nov <- pick_col(cols, c("Nov","November","^11$"), TRUE, "Nov")
col_Dec <- pick_col(cols, c("Dec","December","^12$"), TRUE, "Dec")

month_cols   <- c(col_Jan,col_Feb,col_Mar,col_Apr,col_May,col_Jun,
                  col_Jul,col_Aug,col_Sep,col_Oct,col_Nov,col_Dec)
month_levels <- c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")

# -------------------------
# 清洗 + 半球 + 月度长表（仅 >0）
# -------------------------
monthly_long <- raw0 %>%
  dplyr::mutate(
    q01 = .data[[q01_col]],
    q02 = .data[[q02_col]],
    q03 = .data[[q03_col]],
    q04 = .data[[q04_col]],
    mgmt = .data[[mgmt_col]],
    lat = suppressWarnings(as.numeric(.data[[lat_col]]))
  ) %>%
  dplyr::filter(
    !is_checked(q01),
    !is_checked(q02),
    !is_checked(q03),
    !is_checked(q04)
  ) %>%
  dplyr::filter(is_control(mgmt), !is.na(lat)) %>%
  dplyr::mutate(Hemisphere = dplyr::if_else(lat >= 0, "North", "South")) %>%
  dplyr::select(Hemisphere, Biomes = dplyr::all_of(biomes_col), dplyr::all_of(month_cols)) %>%
  tidyr::pivot_longer(cols = dplyr::all_of(month_cols), names_to = "Month_raw", values_to = "CH4") %>%
  dplyr::mutate(
    Month = dplyr::case_when(
      Month_raw == col_Jan ~ "Jan",
      Month_raw == col_Feb ~ "Feb",
      Month_raw == col_Mar ~ "Mar",
      Month_raw == col_Apr ~ "Apr",
      Month_raw == col_May ~ "May",
      Month_raw == col_Jun ~ "Jun",
      Month_raw == col_Jul ~ "Jul",
      Month_raw == col_Aug ~ "Aug",
      Month_raw == col_Sep ~ "Sep",
      Month_raw == col_Oct ~ "Oct",
      Month_raw == col_Nov ~ "Nov",
      Month_raw == col_Dec ~ "Dec",
      TRUE ~ NA_character_
    ),
    CH4 = suppressWarnings(as.numeric(CH4))
  ) %>%
  dplyr::filter(!is.na(Month), !is.na(CH4), CH4 > 0) %>%
  dplyr::mutate(
    Ecosystem = as.character(Biomes),
    Month     = factor(Month, levels = month_levels, ordered = TRUE)
  )

if (!nrow(monthly_long))
  stop("筛选后月度数据为空（检查：Q01–Q04、management==control、Latitude、Jan–Dec 是否存在且有值）。")

# -------------------------
# ✅ 仅这两幅图不画平滑线：Tundra (South), Boreal forest (South)
# -------------------------
NO_SMOOTH_SOUTH <- c("Tundra", "Boreal forest")

# -------------------------
# 单生态系统 tile（平滑 + SD/CV + 右端均值）
# draw_smooth=FALSE -> 只画点
# -------------------------
build_tile <- function(df_one, title_text, color_hex,
                       draw_smooth = TRUE,
                       w = 6, h = 3.8, dpi = 600, out_file = NULL) {
  
  df_one <- df_one %>% dplyr::arrange(Month)
  if (nrow(df_one) == 0) return(NULL)
  
  avg_val <- df_one$median_avg12[1]
  sd_val  <- df_one$sd_median12[1]
  cv_val  <- df_one$cv_median12[1]
  
  sd_lab <- if (is.na(sd_val)) "SD = NA" else glue::glue("SD = {round(sd_val, 3)}")
  cv_lab <- if (is.na(cv_val)) "CV = NA" else glue::glue("CV = {round(cv_val, 1)}%")
  
  # 平滑曲线数据
  if (isTRUE(draw_smooth)) {
    base_x <- as.numeric(df_one$Month)
    base_y <- df_one$median_CH4
    xx <- seq(1, 12, by = 0.05)
    
    if (length(unique(base_x)) >= 2 && length(unique(base_y)) >= 2) {
      full_vals <- approx(x = base_x, y = base_y, xout = 1:12, method = "linear", rule = 2)$y
      if (length(unique(full_vals)) >= 2) {
        sf <- splinefun(x = 1:12, y = full_vals, method = "monoH.FC")
        yy <- sf(xx)
      } else {
        yy <- approx(x = base_x, y = base_y, xout = xx, method = "linear", rule = 2)$y
      }
    } else {
      yy <- rep(base_y[1], length(xx))
    }
    sp <- data.frame(x = xx, y = yy)
  } else {
    sp <- NULL
  }
  
  base_sz <- 16; title_sz <- 18; axis_t_sz <- 16; axis_txt <- 14
  ann_sd_sz <- 13; ann_avg_sz <- 13
  line_w <- 1.6; pt_sz <- 3.4; avg_lw <- 1.2
  
  y_max <- max(df_one$median_CH4, na.rm = TRUE)
  y_min <- min(df_one$median_CH4, na.rm = TRUE)
  y_rng <- max(1e-9, y_max - y_min)
  y_sd  <- y_max - 0.03 * y_rng
  y_cv  <- y_sd  - 0.07 * y_rng
  
  month_labels <- c("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
  
  p <- ggplot2::ggplot()
  
  if (isTRUE(draw_smooth) && !is.null(sp)) {
    p <- p + ggplot2::geom_line(
      data = sp,
      ggplot2::aes(x = x, y = y),
      linewidth = line_w, color = color_hex
    )
  }
  
  p <- p +
    ggplot2::geom_point(
      data = df_one,
      ggplot2::aes(x = as.numeric(Month), y = median_CH4),
      shape = 16, size = pt_sz, color = color_hex
    ) +
    ggplot2::geom_hline(
      yintercept = avg_val, linetype = 2,
      linewidth = avg_lw, color = color_hex
    ) +
    ggplot2::scale_x_continuous(
      breaks = 1:12, labels = month_labels,
      limits = c(1, 12),
      expand = ggplot2::expansion(mult = c(0.01, 0.01))
    ) +
    ggplot2::labs(
      title = title_text,
      x = NULL,
      y = expression(paste("CH"[4], " uptake (", mu, "g CH"[4], "-C m"^-2, " h"^-1, ")"))
    ) +
    ggplot2::annotate(
      "text", x = 1.05, y = y_sd, label = sd_lab,
      hjust = 0, vjust = 1, size = ann_sd_sz/3
    ) +
    ggplot2::annotate(
      "text", x = 1.05, y = y_cv, label = cv_lab,
      hjust = 0, vjust = 1, size = ann_sd_sz/3
    ) +
    ggplot2::annotate(
      "text", x = 12.0, y = avg_val,
      label = as.character(round(avg_val, 2)),
      hjust = 1, vjust = -0.6, size = ann_avg_sz/3
    ) +
    ggplot2::theme_bw(base_size = base_sz) +
    ggplot2::theme(
      plot.title = ggplot2::element_text(size = title_sz, face = "bold",
                                         hjust = 0.5, margin = ggplot2::margin(b = 4)),
      axis.title = ggplot2::element_text(size = axis_t_sz),
      axis.text  = ggplot2::element_text(size = axis_txt),
      panel.grid.major = ggplot2::element_line(linetype = "dotted", linewidth = 0.5),
      panel.grid.minor = ggplot2::element_blank(),
      plot.margin = ggplot2::margin(10, 14, 10, 10)
    )
  
  if (!is.null(out_file)) {
    ggplot2::ggsave(out_file, p, width = w, height = h, dpi = dpi,
                    bg = "white", limitsize = FALSE)
  }
  p
}

# -------------------------
# 半球流水线（以 Biomes 分类；剔除 Wetland/Others）
# -------------------------
run_for_hemi <- function(hemi) {
  d_hemi <- monthly_long %>% dplyr::filter(Hemisphere == hemi)
  if (nrow(d_hemi) == 0) { warning(sprintf("半球 %s 无数据。", hemi)); return(invisible(NULL)) }
  
  grp <- d_hemi %>%
    dplyr::group_by(Ecosystem, Month) %>%
    dplyr::summarise(median_CH4 = stats::median(CH4), .groups = "drop")
  
  grp$Ecosystem <- ifelse(grp$Ecosystem %in% eco_levels, grp$Ecosystem, "Others")
  grp$Ecosystem <- factor(grp$Ecosystem, levels = eco_levels)
  
  grp <- grp %>% dplyr::filter(!(Ecosystem %in% c("Wetland","Others")))
  if (nrow(grp) == 0) { warning(sprintf("%s：全部被剔除（仅剩 Wetland/Others）。", hemi)); return(invisible(NULL)) }
  
  eco_stats <- grp %>%
    dplyr::group_by(Ecosystem) %>%
    dplyr::summarise(
      median_mean12 = mean(median_CH4, na.rm = TRUE),
      sd_median12   = stats::sd(median_CH4, na.rm = TRUE),
      n_months      = sum(!is.na(median_CH4)),
      .groups = "drop"
    ) %>%
    dplyr::mutate(
      cv_median12 = dplyr::if_else(!is.na(median_mean12) & median_mean12 > 0,
                                   sd_median12 / median_mean12 * 100, NA_real_)
    )
  
  grp_plot <- grp %>%
    dplyr::left_join(eco_stats %>% dplyr::select(Ecosystem, median_mean12,
                                                 sd_median12, cv_median12),
                     by = "Ecosystem") %>%
    dplyr::rename(median_avg12 = median_mean12)
  
  hemi_dir   <- file.path(work_dir, hemi)
  tiles_dir  <- file.path(hemi_dir, "tiles_smooth")
  if (!dir.exists(tiles_dir)) dir.create(tiles_dir, recursive = TRUE)
  
  mosaic_file <- file.path(hemi_dir, paste0("FigS8_", hemi, ".jpg"))
  coef_out    <- file.path(hemi_dir, paste0("ecosystem_variation_coefficients_", hemi, ".xlsx"))
  
  eco_span <- grp %>%
    dplyr::group_by(Ecosystem) %>%
    dplyr::summarise(
      median_min12  = min(median_CH4, na.rm = TRUE),
      median_max12  = max(median_CH4, na.rm = TRUE),
      median_mean12 = mean(median_CH4, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    dplyr::left_join(eco_stats %>% dplyr::select(Ecosystem, sd_median12,
                                                 cv_median12, n_months),
                     by = "Ecosystem") %>%
    dplyr::arrange(desc(sd_median12))
  
  writexl::write_xlsx(
    list(
      coefficients             = eco_span,
      monthly_medians_for_plot = grp_plot
    ),
    path = coef_out
  )
  
  tile_w <- 6; tile_h <- 3.8; dpi <- 600
  ecos_in_data <- unique(as.character(grp_plot$Ecosystem))
  ecos_order   <- eco_levels[eco_levels %in% ecos_in_data]
  
  plot_list <- purrr::map(ecos_order, function(eco){
    sub <- grp_plot %>% dplyr::filter(Ecosystem == eco)
    if (nrow(sub) == 0) return(NULL)
    
    col_hex <- eco_colors[[eco]]
    safe    <- gsub("[^A-Za-z0-9_\\-]+", "_", eco)
    out_jpg <- file.path(tiles_dir, paste0(safe, ".jpg"))
    
    draw_smooth <- TRUE
    if (identical(hemi, "South") && (eco %in% NO_SMOOTH_SOUTH)) {
      draw_smooth <- FALSE
    }
    
    build_tile(sub, paste0(eco, " (", hemi, ")"), col_hex,
               draw_smooth = draw_smooth,
               w = tile_w, h = tile_h, dpi = dpi, out_file = out_jpg)
  }) %>% purrr::compact()
  
  if (length(plot_list) == 0) { warning(sprintf("%s：无可用小图。", hemi)); return(invisible(NULL)) }
  
  ncol_mosaic <- 3; nrow_mosaic <- 6
  target_slots <- ncol_mosaic * nrow_mosaic
  if (length(plot_list) < target_slots) {
    blanks_needed <- target_slots - length(plot_list)
    blank_plot <- ggplot2::ggplot() + ggplot2::theme_void()
    plot_list <- c(plot_list, rep(list(blank_plot), blanks_needed))
  } else if (length(plot_list) > target_slots) {
    plot_list <- plot_list[1:target_slots]
  }
  
  pg <- cowplot::plot_grid(plotlist = plot_list, ncol = ncol_mosaic, align = "hv")
  
  ggplot2::ggsave(filename = mosaic_file, plot = pg,
                  width  = tile_w * ncol_mosaic,
                  height = tile_h * nrow_mosaic,
                  dpi = dpi, bg = "white", limitsize = FALSE)
  
  message("✅ ", hemi, " tiles 目录: ", tiles_dir)
  message("✅ ", hemi, " mosaic : ", mosaic_file)
  message("✅ ", hemi, " 统计表 : ", coef_out)
}

# -------------------------
# 执行：North & South
# -------------------------
run_for_hemi("North")
run_for_hemi("South")

message("✅ 全部完成！输出目录：", work_dir)