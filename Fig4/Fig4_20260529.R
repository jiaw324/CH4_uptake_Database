# ============================================================
# Fig4: Seasons + Annual (from Studies and Fluxes template)
# - Input:
#   1) CH4 uptake data_CH4 FLUX_1_1410.xlsx  (only check exists)
#   2) Studies and Fluxes.xlsx              (main data source)
#
# - Filters:
#   Exclude if Q01/Q02/Q03/Q04 is checked (√/✓/✔/yes/1/true)
#   Keep only management == "control" (case-insensitive, trim)
#   Keep only value > 0 for stats & plotting
#   Exclude wetland & others (double insurance)
#
# - Output:
#   D:/233 CH4 uptake_Database/CH4 uptake_Database_Write/Supplymentary_code/Fig4/
#     Fig4.jpg
#     Fig4.xlsx
#
# - Fig4.xlsx sheets:
#   1) Annual
#   2) Spring
#   3) Summer
#   4) Autumn
#   5) Winter
#   6) Exclusion_summary
#   7) Annual_excluded_by_biome
#
# - Exclusion_summary / Annual_excluded_by_biome only keep:
#   dataset
#   total_valid_records
#   n_emission_lt0
#   percent_emission_lt0
#   mean_emission_lt0
#   median_emission_lt0
# ============================================================

# ===============================
# 🎟️ 必要包
# ===============================
required_packages <- c(
  "readxl", "dplyr", "stringr", "ggplot2",
  "tidyr", "cowplot", "scales", "writexl"
)

for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) install.packages(pkg)
}
invisible(lapply(required_packages, library, character.only = TRUE))

# ===============================
# 📁 路径
# ===============================
ROOT <- "D:/233 CH4 uptake_Database/CH4 uptake_Database_Write/Supplymentary_code"

# 输入文件
in_raw <- file.path(ROOT, "CH4 uptake data_CH4 FLUX_1_1410.xlsx")   # 附件1，仅检查存在
in_tpl <- file.path(ROOT, "Studies and Fluxes.xlsx")                # 附件2，真正用于画图

stopifnot(file.exists(in_raw))
stopifnot(file.exists(in_tpl))

# 输出路径
out_dir <- file.path(ROOT, "Fig4")
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
setwd(out_dir)

out_file_img   <- file.path(out_dir, "Fig4.jpg")
out_file_stats <- file.path(out_dir, "Fig4.xlsx")

# 显示范围，仅控制图上显示，不影响统计
y_min_plot <- -15
y_max_plot <- 200

# ===============================
# 🧭 读取 Studies and Fluxes
# 优先读取 row sheet，否则读取第一个 sheet
# ===============================
sheets <- readxl::excel_sheets(in_tpl)
sheet_to_read <- if ("row" %in% sheets) "row" else sheets[1]
df <- readxl::read_excel(in_tpl, sheet = sheet_to_read)

# ===============================
# 🔎 自动匹配列名
# ===============================
pick_col <- function(nms, candidates, must = TRUE, label = "column") {
  nms_l <- tolower(trimws(nms))
  candidates <- trimws(candidates)
  cand_l <- tolower(candidates)
  
  # 1) 精确匹配
  for (cand in cand_l) {
    hit <- which(nms_l == cand)
    if (length(hit) >= 1) return(nms[hit[1]])
  }
  
  # 2) 边界匹配 / 模糊匹配
  for (cand in cand_l) {
    pat <- paste0("(^", cand, "$)|\\b", cand, "\\b")
    hit <- which(grepl(pat, nms_l, perl = TRUE))
    if (length(hit) >= 1) return(nms[hit[1]])
  }
  
  # 3) 包含匹配
  hit <- which(Reduce(`|`, lapply(cand_l, function(x) grepl(x, nms_l))))
  if (length(hit) >= 1) return(nms[hit[1]])
  
  if (must) {
    stop(
      sprintf(
        "找不到【%s】列。表头示例：%s",
        label,
        paste(head(nms, 40), collapse = ",")
      ),
      call. = FALSE
    )
  }
  
  NA_character_
}

cols <- names(df)

# 关键字段
biome_col <- pick_col(cols, c("Biome", "Biomes", "Boimes", "biome", "biome_name"), TRUE, "Biome/Biomes")
lat_col   <- pick_col(cols, c("Latitude", "lat", "latitude"), TRUE, "Latitude")
lon_col   <- pick_col(cols, c("Longitude", "lon", "long", "longitude"), TRUE, "Longitude")

# QC 字段
q01_col <- pick_col(cols, c("Q01", "q01"), TRUE, "Q01")
q02_col <- pick_col(cols, c("Q02", "q02"), TRUE, "Q02")
q03_col <- pick_col(cols, c("Q03", "q03"), TRUE, "Q03")
q04_col <- pick_col(cols, c("Q04", "q04"), TRUE, "Q04")

# management 字段
mgmt_col <- pick_col(cols, c("management", "Management"), TRUE, "management")

# 季节/年字段
col_Spring <- pick_col(cols, c("Spring", "spring"), TRUE, "Spring")
col_Summer <- pick_col(cols, c("Summer", "summer"), TRUE, "Summer")
col_Autumn <- pick_col(cols, c("Autumn", "autumn", "Fall", "fall"), TRUE, "Autumn")
col_Winter <- pick_col(cols, c("Winter", "winter"), TRUE, "Winter")
col_Annual <- pick_col(cols, c("Annual", "annual", "1-12", "Year", "year"), TRUE, "Annual")

message(sprintf(
  "✅ 使用列名：Biome=%s, lat=%s, lon=%s, Q01=%s, Q02=%s, Q03=%s, Q04=%s, management=%s ; sheet=%s",
  biome_col, lat_col, lon_col,
  q01_col, q02_col, q03_col, q04_col,
  mgmt_col, sheet_to_read
))

# ===============================
# ✔️ 过滤辅助函数
# ===============================
norm_txt <- function(x) {
  if (is.null(x)) return(NA_character_)
  stringr::str_squish(as.character(x))
}

is_checked <- function(x) {
  y <- stringr::str_to_lower(norm_txt(x))
  !is.na(y) & stringr::str_detect(y, "^(√|✓|✔|yes|y|1|true)$")
}

is_control <- function(x) {
  y <- stringr::str_to_lower(norm_txt(x))
  !is.na(y) & stringr::str_detect(y, "^control(?:\\s*group)?$")
}

# ===============================
# 🧼 生态系统色板
# 不包含 others / wetland
# ===============================
biome_palette <- c(
  "tundra" = "#a7d3df",
  "boreal forest" = "#4f9f4f",
  "temperate seasonal forest" = "#8cbf70",
  "temperate rainforest" = "#e8c8a1",
  "tropical rainforest" = "#c49b50",
  "tropical seasonal forest" = "#2e7032",
  "savanna" = "#5fa6bc",
  "subtropical desert" = "#f2c300",
  "temperate grassland" = "#2e6f89",
  "desert" = "#fff2b2",
  "woodland" = "#cd7a50",
  "shrubland" = "#a34030",
  "alpine" = "#b27874",
  "agriculture" = "#886b99",
  "bare" = "#357266",
  "urban" = "#6b575a"
)

# 缩写标签
biome_labels <- c(
  "savanna" = "Sav",
  "boreal forest" = "BF",
  "agriculture" = "Agri",
  "bare" = "Bare",
  "subtropical desert" = "SDes",
  "desert" = "Des",
  "urban" = "Urban",
  "woodland" = "Wood",
  "temperate seasonal forest" = "TSF",
  "temperate grassland" = "TG",
  "alpine" = "Alp",
  "tundra" = "Tun",
  "tropical rainforest" = "TRF",
  "temperate rainforest" = "TRRF",
  "shrubland" = "Shrub",
  "tropical seasonal forest" = "TRSF"
)

# ===============================
# 🧼 清洗 + QC + management 筛选
# df0 是基础数据：
# 已排除 Q01-Q04 和非 control，
# 但还没有排除 value <= 0。
# ===============================
df0 <- df %>%
  dplyr::mutate(
    q01 = .data[[q01_col]],
    q02 = .data[[q02_col]],
    q03 = .data[[q03_col]],
    q04 = .data[[q04_col]],
    mgmt = .data[[mgmt_col]],
    lat = suppressWarnings(as.numeric(.data[[lat_col]])),
    lon = suppressWarnings(as.numeric(.data[[lon_col]]))
  ) %>%
  # Q01-Q04 只要勾选就排除
  dplyr::filter(!is_checked(q01)) %>%
  dplyr::filter(!is_checked(q02)) %>%
  dplyr::filter(!is_checked(q03)) %>%
  dplyr::filter(!is_checked(q04)) %>%
  # 只保留 management == control
  dplyr::filter(is_control(mgmt)) %>%
  dplyr::filter(!is.na(lat) & !is.na(lon)) %>%
  dplyr::mutate(
    biomes_clean = stringr::str_trim(stringr::str_to_lower(.data[[biome_col]])),
    
    # 不在色板内的 biome 变成 NA，后面不参与分析
    biomes_clean = dplyr::if_else(
      is.na(biomes_clean) | !(biomes_clean %in% names(biome_palette)),
      NA_character_,
      biomes_clean
    ),
    
    Spring = suppressWarnings(as.numeric(.data[[col_Spring]])),
    Summer = suppressWarnings(as.numeric(.data[[col_Summer]])),
    Autumn = suppressWarnings(as.numeric(.data[[col_Autumn]])),
    Winter = suppressWarnings(as.numeric(.data[[col_Winter]])),
    Annual = suppressWarnings(as.numeric(.data[[col_Annual]]))
  )

message("✅ QC + management + biome 清洗后记录数: ", nrow(df0))

# ============================================================
# 🧮 统计 <0 emission records
# 输出两个精简统计表：
#   1) Exclusion_summary
#   2) Annual_excluded_by_biome
#
# 两个表只保留：
#   dataset
#   total_valid_records
#   n_emission_lt0
#   percent_emission_lt0
#   mean_emission_lt0
#   median_emission_lt0
# ============================================================

# Annual 全部有效记录，包括 >0, <0, =0
annual_all_valid_df <- df0 %>%
  dplyr::select(biomes_clean, value = Annual) %>%
  dplyr::filter(
    !is.na(biomes_clean),
    !is.na(value),
    !biomes_clean %in% c("wetland", "others")
  )

# Seasonal 全部有效记录，包括 >0, <0, =0
# 这里只用于 Exclusion_summary 的 Spring/Summer/Autumn/Winter 总体统计
season_all_valid_df <- df0 %>%
  dplyr::select(biomes_clean, Spring, Summer, Autumn, Winter) %>%
  tidyr::pivot_longer(
    cols = Spring:Winter,
    names_to = "season",
    values_to = "value"
  ) %>%
  dplyr::filter(
    !is.na(biomes_clean),
    !is.na(value),
    !biomes_clean %in% c("wetland", "others")
  )

# Annual 总体统计：输出到 Exclusion_summary
annual_exclusion_summary <- annual_all_valid_df %>%
  dplyr::summarise(
    dataset = "Annual",
    total_valid_records = dplyr::n(),
    n_emission_lt0 = sum(value < 0, na.rm = TRUE),
    percent_emission_lt0 = round(100 * n_emission_lt0 / total_valid_records, 2),
    mean_emission_lt0 = ifelse(
      n_emission_lt0 > 0,
      round(mean(value[value < 0], na.rm = TRUE), 3),
      NA_real_
    ),
    median_emission_lt0 = ifelse(
      n_emission_lt0 > 0,
      round(stats::median(value[value < 0], na.rm = TRUE), 3),
      NA_real_
    )
  )

# 四季总体统计：输出到 Exclusion_summary
season_exclusion_summary <- season_all_valid_df %>%
  dplyr::group_by(season) %>%
  dplyr::summarise(
    dataset = unique(season),
    total_valid_records = dplyr::n(),
    n_emission_lt0 = sum(value < 0, na.rm = TRUE),
    percent_emission_lt0 = round(100 * n_emission_lt0 / total_valid_records, 2),
    mean_emission_lt0 = ifelse(
      n_emission_lt0 > 0,
      round(mean(value[value < 0], na.rm = TRUE), 3),
      NA_real_
    ),
    median_emission_lt0 = ifelse(
      n_emission_lt0 > 0,
      round(stats::median(value[value < 0], na.rm = TRUE), 3),
      NA_real_
    ),
    .groups = "drop"
  ) %>%
  dplyr::select(
    dataset,
    total_valid_records,
    n_emission_lt0,
    percent_emission_lt0,
    mean_emission_lt0,
    median_emission_lt0
  )

# 汇总表：Annual + Spring/Summer/Autumn/Winter
exclusion_summary <- dplyr::bind_rows(
  annual_exclusion_summary,
  season_exclusion_summary
) %>%
  dplyr::select(
    dataset,
    total_valid_records,
    n_emission_lt0,
    percent_emission_lt0,
    mean_emission_lt0,
    median_emission_lt0
  )

# Annual 按 biome 统计：输出到 Annual_excluded_by_biome
# 注意：这里把 biomes_clean 重命名为 dataset，
# 这样两个 emission 统计表的列名完全一致。
annual_exclusion_stats <- annual_all_valid_df %>%
  dplyr::group_by(biomes_clean) %>%
  dplyr::summarise(
    dataset = unique(biomes_clean),
    total_valid_records = dplyr::n(),
    n_emission_lt0 = sum(value < 0, na.rm = TRUE),
    percent_emission_lt0 = round(100 * n_emission_lt0 / total_valid_records, 2),
    mean_emission_lt0 = ifelse(
      n_emission_lt0 > 0,
      round(mean(value[value < 0], na.rm = TRUE), 3),
      NA_real_
    ),
    median_emission_lt0 = ifelse(
      n_emission_lt0 > 0,
      round(stats::median(value[value < 0], na.rm = TRUE), 3),
      NA_real_
    ),
    .groups = "drop"
  ) %>%
  dplyr::select(
    dataset,
    total_valid_records,
    n_emission_lt0,
    percent_emission_lt0,
    mean_emission_lt0,
    median_emission_lt0
  ) %>%
  dplyr::arrange(dplyr::desc(n_emission_lt0), dataset)

message("✅ 已完成 <0 emission records 精简统计")

# ===============================
# 🧼 只用 >0 的通量参与 Fig4 作图和原始统计
# ===============================
annual_df <- df0 %>%
  dplyr::select(biomes_clean, value = Annual) %>%
  dplyr::filter(!is.na(biomes_clean), !is.na(value), value > 0) %>%
  dplyr::filter(!biomes_clean %in% c("wetland", "others"))

season_df <- df0 %>%
  dplyr::select(biomes_clean, Spring, Summer, Autumn, Winter) %>%
  tidyr::pivot_longer(
    cols = Spring:Winter,
    names_to = "season",
    values_to = "value"
  ) %>%
  dplyr::filter(!is.na(biomes_clean), !is.na(value), value > 0) %>%
  dplyr::filter(!biomes_clean %in% c("wetland", "others"))

all_biomes_present <- sort(unique(c(annual_df$biomes_clean, season_df$biomes_clean)))
stopifnot(length(all_biomes_present) > 0)

message("✅ Annual positive records: ", nrow(annual_df))
message("✅ Seasonal positive records: ", nrow(season_df))

# ===============================
# 🧮 Annual 基准顺序
# ===============================
compute_biome_order <- function(data_annual) {
  d <- data_annual %>% dplyr::filter(value > 0)
  
  if (nrow(d) == 0) return(character(0))
  
  d %>%
    dplyr::group_by(biomes_clean) %>%
    dplyr::summarise(
      med = stats::median(value, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    dplyr::arrange(med) %>%
    dplyr::pull(biomes_clean) %>%
    as.character()
}

# ===============================
# 🖌️ 绘图 + 统计输出
# 返回 list(plot, stats)
# stats 用于导出 Annual / Spring / Summer / Autumn / Winter 五个 sheet
# ===============================
make_smart_plot <- function(data, title_text, show_y = TRUE, preferred_order = NULL) {
  d <- data %>% dplyr::filter(value > 0)
  
  stat_raw <- d %>%
    dplyr::group_by(biomes_clean) %>%
    dplyr::summarise(
      n = dplyr::n(),
      median = stats::median(value, na.rm = TRUE),
      mean = mean(value, na.rm = TRUE),
      .groups = "drop"
    )
  
  # 排序
  if (is.null(preferred_order) || length(preferred_order) == 0) {
    ord_levels <- stat_raw %>%
      dplyr::arrange(median) %>%
      dplyr::pull(biomes_clean)
  } else {
    present <- stat_raw$biomes_clean
    extras <- setdiff(present, preferred_order)
    master_order <- names(biome_palette)
    extras_ordered <- intersect(master_order, extras)
    ord_levels <- c(extras_ordered, intersect(preferred_order, present))
  }
  
  stat_for_labels <- stat_raw %>%
    dplyr::mutate(biomes_clean_chr = as.character(biomes_clean)) %>%
    dplyr::slice(match(ord_levels, biomes_clean_chr)) %>%
    dplyr::mutate(
      label = paste0(biome_labels[biomes_clean_chr], " (", n, ")")
    )
  
  d2 <- d %>%
    dplyr::mutate(biomes_clean_chr = as.character(biomes_clean)) %>%
    dplyr::inner_join(
      stat_for_labels %>%
        dplyr::select(biomes_clean_chr, n, median, mean, label),
      by = "biomes_clean_chr"
    ) %>%
    dplyr::mutate(
      biomes_clean = factor(biomes_clean_chr, levels = ord_levels)
    )
  
  # 统计表输出到 Excel
  stat_tbl <- stat_for_labels %>%
    dplyr::transmute(
      biome = biomes_clean_chr,
      n = n,
      median = round(median, 3),
      mean = round(mean, 3)
    )
  
  # 根据样本量选择点图 / 箱线图 / 小提琴图
  df_pts <- d2 %>% dplyr::filter(n < 10)
  df_box <- d2 %>% dplyr::filter(n >= 10 & n < 20)
  df_vln <- d2 %>% dplyr::filter(n >= 20)
  
  safe_pal <- biome_palette
  missing_keys <- setdiff(unique(d2$biomes_clean_chr), names(safe_pal))
  if (length(missing_keys)) safe_pal[missing_keys] <- "#b3b5ae"
  
  p <- ggplot2::ggplot(
    d2,
    ggplot2::aes(x = label, y = value, fill = biomes_clean)
  )
  
  if (nrow(df_vln) > 0) {
    p <- p +
      ggplot2::geom_violin(
        data = df_vln,
        scale = "width",
        width = 0.9,
        color = "black",
        alpha = 0.9,
        adjust = 3
      )
  }
  
  if (nrow(df_box) > 0) {
    p <- p +
      ggplot2::geom_boxplot(
        data = df_box,
        width = 0.6,
        outlier.shape = 21,
        outlier.size = 1.8,
        outlier.stroke = 0.3,
        color = "black",
        alpha = 0.9
      )
  }
  
  if (nrow(df_pts) > 0) {
    p <- p +
      ggplot2::geom_point(
        data = df_pts,
        position = ggplot2::position_jitter(width = 0.12, height = 0),
        shape = 21,
        size = 2.4,
        stroke = 0.3,
        color = "black",
        alpha = 0.9
      )
  }
  
  # 中位数线
  p <- p +
    ggplot2::geom_errorbar(
      data = stat_for_labels,
      ggplot2::aes(x = label, ymin = median, ymax = median),
      inherit.aes = FALSE,
      width = 0.8,
      linewidth = 0.9,
      color = "black"
    )
  
  # 均值 ×
  p <- p +
    ggplot2::geom_point(
      data = stat_for_labels,
      ggplot2::aes(x = label, y = mean),
      inherit.aes = FALSE,
      shape = 4,
      size = 3.8,
      stroke = 1.1,
      color = "black"
    )
  
  p <- p +
    ggplot2::scale_x_discrete(limits = stat_for_labels$label) +
    ggplot2::scale_fill_manual(values = safe_pal, drop = FALSE) +
    ggplot2::scale_y_continuous(
      expand = ggplot2::expansion(mult = c(0.02, 0.08))
    ) +
    ggplot2::coord_cartesian(
      ylim = c(y_min_plot, y_max_plot),
      clip = "on"
    ) +
    ggplot2::theme_bw(base_size = 18) +
    ggplot2::theme(
      axis.text.x = ggplot2::element_text(
        angle = 45,
        hjust = 1,
        size = 16,
        color = "black"
      ),
      axis.text.y = ggplot2::element_text(size = 16, color = "black"),
      axis.title.y = ggplot2::element_text(size = 16),
      axis.title.x = ggplot2::element_blank(),
      plot.margin = ggplot2::margin(20, 20, 30, 20),
      legend.position = "none"
    ) +
    ggplot2::labs(
      y = expression(CH[4]~uptake~" ("*mu*"g CH"[4]*"-C m"^{-2}*" h"^{-1}*")")
    ) +
    ggplot2::annotate(
      "text",
      x = -Inf,
      y = y_max_plot - 0.02 * (y_max_plot - y_min_plot),
      label = title_text,
      hjust = -0.1,
      vjust = 1,
      size = 7,
      fontface = "bold"
    )
  
  if (!show_y) {
    p <- p +
      ggplot2::theme(
        axis.text.y = ggplot2::element_blank(),
        axis.title.y = ggplot2::element_blank()
      )
  }
  
  list(plot = p, stats = stat_tbl)
}

# ===============================
# 🖼️ 组合绘图
# Annual + 四季
# ===============================
annual_order <- compute_biome_order(annual_df)

if (length(annual_order) == 0) {
  seasonal_order <- season_df %>%
    dplyr::group_by(biomes_clean) %>%
    dplyr::summarise(
      med = median(value, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    dplyr::arrange(med) %>%
    dplyr::pull(biomes_clean)
  
  annual_order <- if (length(seasonal_order)) seasonal_order else names(biome_palette)
  annual_order <- setdiff(annual_order, c("wetland", "others"))
}

res_annual <- make_smart_plot(
  annual_df,
  "Annual",
  show_y = TRUE
)

res_spring <- make_smart_plot(
  dplyr::filter(season_df, season == "Spring"),
  "Spring",
  show_y = TRUE,
  preferred_order = annual_order
)

res_summer <- make_smart_plot(
  dplyr::filter(season_df, season == "Summer"),
  "Summer",
  show_y = FALSE,
  preferred_order = annual_order
)

res_autumn <- make_smart_plot(
  dplyr::filter(season_df, season == "Autumn"),
  "Autumn",
  show_y = TRUE,
  preferred_order = annual_order
)

res_winter <- make_smart_plot(
  dplyr::filter(season_df, season == "Winter"),
  "Winter",
  show_y = FALSE,
  preferred_order = annual_order
)

p_annual <- res_annual$plot
p_spring <- res_spring$plot
p_summer <- res_summer$plot
p_autumn <- res_autumn$plot
p_winter <- res_winter$plot

season_grid <- cowplot::plot_grid(
  p_spring,
  p_summer,
  p_autumn,
  p_winter,
  ncol = 2
)

final_plot <- cowplot::plot_grid(
  p_annual,
  season_grid,
  ncol = 1,
  rel_heights = c(1.2, 2)
)

ggplot2::ggsave(
  filename = out_file_img,
  plot = final_plot,
  width = 20,
  height = 14,
  dpi = 600
)

message("✅ 图已保存: ", out_file_img)

# ===============================
# 📊 导出统计到 Excel
# 保留原始 Fig4 正值统计表：
#   Annual / Spring / Summer / Autumn / Winter
#
# 同时新增两个精简 emission 统计表：
#   Exclusion_summary
#   Annual_excluded_by_biome
#
# 不再输出：
#   Seasonal_excluded_by_biome
# ===============================
stat_list <- list(
  # 原始 Fig4 正值 uptake 统计
  Annual = res_annual$stats,
  Spring = res_spring$stats,
  Summer = res_summer$stats,
  Autumn = res_autumn$stats,
  Winter = res_winter$stats,
  
  # 新增：总体 <0 emission records 精简统计
  Exclusion_summary = exclusion_summary,
  
  # 新增：Annual 按 biome 的 <0 emission records 精简统计
  Annual_excluded_by_biome = annual_exclusion_stats
)

writexl::write_xlsx(stat_list, path = out_file_stats)

message("✅ 统计已保存: ", out_file_stats)
message("✅ Fig4.xlsx 包含 7 个 sheet:")
message("   Annual / Spring / Summer / Autumn / Winter / Exclusion_summary / Annual_excluded_by_biome")
message("✅ Exclusion_summary 和 Annual_excluded_by_biome 只保留 6 列:")
message("   dataset / total_valid_records / n_emission_lt0 / percent_emission_lt0 / mean_emission_lt0 / median_emission_lt0")

# ===============================
# ✅ 额外输出：告诉你实际用了哪个 sheet
# ===============================
message("✅ 读取的 Studies and Fluxes sheet: ", sheet_to_read)

# ===============================
# ✅ 控制台打印关键汇总，方便直接看结果
# ===============================
message("------------------------------------------------------------")
message("✅ Exclusion summary:")
print(exclusion_summary)
message("------------------------------------------------------------")
message("✅ Annual excluded by biome:")
print(annual_exclusion_stats)
message("------------------------------------------------------------")