# ===============================
# 🎟️ 必要包
# ===============================
required_packages <- c("readxl","dplyr","stringr","ggplot2","tidyr")
for (pkg in required_packages) if (!requireNamespace(pkg, quietly = TRUE)) install.packages(pkg)
invisible(lapply(required_packages, library, character.only = TRUE))

# ===============================
# 📁 输入/输出路径（按你要求）
# ===============================
in_dir <- "D:/Users/jiaweiChiang/Desktop/Supplymentary_code"

# ✅ 附件2：主数据源（读取它）
in_file <- file.path(in_dir, "Studies and Fluxes.xlsx")
stopifnot(file.exists(in_file))

# ✅ 附件1：只检查存在（不读取）
file_check <- file.path(in_dir, "CH4 uptake data_CH4 FLUX_1_1410.xlsx")
stopifnot(file.exists(file_check))

# ✅ 输出目录：FigS4
out_dir <- file.path(in_dir, "FigS4")
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)

out_file <- file.path(out_dir, "FigS4.jpg")

# ===============================
# 🔧 可调参数（↑字号）
# ===============================
ymin_plot        <- -15
ymax_plot        <- 120
base_fontsize    <- 20
axis_tick_size   <- 16
strip_title_size <- 18
axis_title_y_sz  <- 18
label_textsize   <- 4.6
n_label_size     <- 4.2

# ===============================
# 🧭 读取（优先 row；否则第一个sheet）
# ===============================
sheets <- readxl::excel_sheets(in_file)
sheet_to_read <- if ("row" %in% sheets) "row" else sheets[1]
df <- readxl::read_excel(in_file, sheet = sheet_to_read)

# ===============================
# 🔎 自动匹配列名函数
# ===============================
pick_col <- function(nms, candidates, must=TRUE, label="column") {
  nms_l <- tolower(nms)
  for (cand in candidates) {
    hit <- which(nms_l == tolower(cand))
    if (length(hit) == 1) return(nms[hit])
  }
  for (cand in candidates) {
    c_low <- tolower(cand)
    pat <- paste0("^(?:", c_low, ")$|\\b", c_low, "\\b")
    hit <- which(grepl(pat, nms_l, perl = TRUE))
    if (length(hit) > 0) return(nms[hit[1]])
  }
  if (must) stop(sprintf("找不到【%s】列。表头示例：%s", label, paste(head(nms, 30), collapse=", ")), call. = FALSE)
  NA_character_
}

cols <- names(df)

# ===============================
# ✅ 按新表头匹配列名（Biome 改名兼容）
# ===============================
# ⭐ 这里改了：Biome / Biomes / Boimes 都能识别
biome_col <- pick_col(cols, c("Biome","biome","Biomes","biomes","Boimes","boimes"), TRUE, "Biome")
lat_col   <- pick_col(cols, c("Latitude","latitude","lat"), TRUE, "Latitude")
lon_col   <- pick_col(cols, c("Longitude","longitude","lon","long"), TRUE, "Longitude")

q01_col <- pick_col(cols, c("Q01","q01"), TRUE, "Q01")
q02_col <- pick_col(cols, c("Q02","q02"), TRUE, "Q02")
q03_col <- pick_col(cols, c("Q03","q03"), TRUE, "Q03")
q04_col <- pick_col(cols, c("Q04","q04"), TRUE, "Q04")

mgmt_col <- pick_col(cols, c("management","Management"), TRUE, "management")

col_Spring <- pick_col(cols, c("Spring","spring"), TRUE, "Spring")
col_Summer <- pick_col(cols, c("Summer","summer"), TRUE, "Summer")
col_Autumn <- pick_col(cols, c("Autumn","autumn","Fall","fall"), TRUE, "Autumn")
col_Winter <- pick_col(cols, c("Winter","winter"), TRUE, "Winter")
col_Annual <- pick_col(cols, c("Annual","annual","1-12","Year","year"), TRUE, "Annual")

# ===============================
# ✔️ 判定函数
# ===============================
norm_txt <- function(x) { if (is.null(x)) return(NA_character_); stringr::str_squish(as.character(x)) }

is_checked <- function(x){
  y <- stringr::str_to_lower(norm_txt(x))
  !is.na(y) & stringr::str_detect(y, "^(√|✓|✔|yes|y|1|true)$")
}

is_control <- function(x){
  y <- stringr::str_to_lower(norm_txt(x))
  !is.na(y) & stringr::str_detect(y, "^control(?:\\s*group)?$")
}

# ===============================
# 🎨 配色（小写键）
# ===============================
biome_colors <- c(
  "tundra"="#a7d3df","boreal forest"="#4f9f4f","temperate seasonal forest"="#8cbf70",
  "temperate rainforest"="#e8c8a1","tropical rainforest"="#c49b50","tropical seasonal forest"="#2e7032",
  "savanna"="#5fa6bc","subtropical desert"="#f2c300","temperate grassland"="#2e6f89",
  "desert"="#fff2b2","woodland"="#cd7a50","shrubland"="#a34030",
  "alpine"="#b27874","agriculture"="#886b99","bare"="#357266",
  "wetland"="#69778c","urban"="#6b575a","others"="#b3b5ae"
)

to_full <- function(x) tools::toTitleCase(as.character(x))

# ===============================
# 📊 筛选与清洗（新规则：Q01-Q04 + management）
# ===============================
df_base <- df %>%
  dplyr::mutate(
    Q01  = .data[[q01_col]],
    Q02  = .data[[q02_col]],
    Q03  = .data[[q03_col]],
    Q04  = .data[[q04_col]],
    mgmt = .data[[mgmt_col]],
    lat = suppressWarnings(as.numeric(.data[[lat_col]])),
    lon = suppressWarnings(as.numeric(.data[[lon_col]]))
  ) %>%
  # ✅ 过滤 Q01-Q04 被勾选的数据
  dplyr::filter(!is_checked(Q01)) %>%
  dplyr::filter(!is_checked(Q02)) %>%
  dplyr::filter(!is_checked(Q03)) %>%
  dplyr::filter(!is_checked(Q04)) %>%
  # ✅ 只保留 management == control
  dplyr::filter(is_control(mgmt)) %>%
  # ✅ 经纬度必须有效
  dplyr::filter(!is.na(lat) & !is.na(lon)) %>%
  # ✅ biome + 数值列
  dplyr::mutate(
    biomes_clean = stringr::str_trim(stringr::str_to_lower(.data[[biome_col]])),
    biomes_clean = ifelse(is.na(biomes_clean) | !(biomes_clean %in% names(biome_colors)), "others", biomes_clean),
    Spring = suppressWarnings(as.numeric(.data[[col_Spring]])),
    Summer = suppressWarnings(as.numeric(.data[[col_Summer]])),
    Autumn = suppressWarnings(as.numeric(.data[[col_Autumn]])),
    Winter = suppressWarnings(as.numeric(.data[[col_Winter]])),
    Annual = suppressWarnings(as.numeric(.data[[col_Annual]]))
  ) %>%
  # ✅ 排除 wetland + others
  dplyr::filter(!biomes_clean %in% c("wetland","others"))

# —— 仅保留 >0 的通量参与分析 —— #
annual_df <- df_base %>%
  dplyr::filter(!is.na(Annual) & Annual > 0) %>%
  dplyr::transmute(biomes_clean, value = Annual)

season_df <- df_base %>%
  dplyr::select(biomes_clean, Spring, Summer, Autumn, Winter) %>%
  tidyr::pivot_longer(Spring:Winter, names_to = "season", values_to = "value") %>%
  dplyr::filter(!is.na(value) & value > 0)

# ===============================
# 📈 统计（仅 >0）
# ===============================
season_levels <- c("Spring","Summer","Autumn","Winter")

season_stats <- season_df %>%
  dplyr::mutate(season = factor(season, levels = season_levels)) %>%
  dplyr::group_by(biomes_clean, season) %>%
  dplyr::summarise(n = dplyr::n(), med = stats::median(value), .groups = "drop")

annual_stats <- annual_df %>%
  dplyr::group_by(biomes_clean) %>%
  dplyr::summarise(med_ann = stats::median(value), .groups = "drop")

# 基于四季中位数的 SD / CV
season_sd_tbl <- season_stats %>%
  dplyr::group_by(biomes_clean) %>%
  dplyr::summarise(
    k           = sum(!is.na(med)),
    season_mean = mean(med),
    sd_sample   = if (k > 1) sqrt(sum((med - season_mean)^2) / (k - 1)) else NA_real_,
    .groups = "drop"
  ) %>%
  dplyr::mutate(
    cv_standard_pct = dplyr::if_else(is.na(season_mean) | season_mean == 0,
                                     NA_real_, sd_sample / season_mean * 100)
  )

# 排序（按 annual median）
annual_order <- annual_stats %>% dplyr::arrange(med_ann) %>% (\(x) as.character(x$biomes_clean))()
season_only <- setdiff(unique(as.character(season_stats$biomes_clean)), annual_order)

season_only_order <- season_stats %>%
  dplyr::filter(as.character(biomes_clean) %in% season_only) %>%
  dplyr::group_by(biomes_clean) %>%
  dplyr::summarise(median_of_medians = stats::median(med), .groups = "drop") %>%
  dplyr::arrange(median_of_medians) %>%
  (\(x) as.character(x$biomes_clean))()

ord_levels <- c(annual_order, season_only_order)

plot_df <- season_stats %>%
  dplyr::mutate(
    biomes_clean_chr = as.character(biomes_clean),
    biomes_clean = factor(biomes_clean_chr, levels = ord_levels),
    label = factor(to_full(biomes_clean_chr), levels = to_full(ord_levels))
  )

annual_for_plot <- plot_df %>%
  dplyr::distinct(label, biomes_clean) %>%
  dplyr::left_join(annual_stats, by = "biomes_clean") %>%
  dplyr::filter(!is.na(med_ann))

season_sd_for_plot <- season_sd_tbl %>%
  dplyr::mutate(
    biomes_clean_chr = as.character(biomes_clean),
    label = factor(to_full(biomes_clean_chr), levels = levels(plot_df$label)),
    stats_label = sprintf("SD=%.2f  CV=%.1f%%", sd_sample, cv_standard_pct)
  ) %>%
  dplyr::filter(!is.na(label))

season_n_for_plot <- season_stats %>%
  dplyr::mutate(
    biomes_clean_chr = as.character(biomes_clean),
    label = factor(tools::toTitleCase(biomes_clean_chr), levels = levels(plot_df$label)),
    n_label = as.character(n)
  ) %>%
  dplyr::select(label, season, n_label)

right_nudge  <- 0.42
bottom_vjust <- 1.6

# ===============================
# 🖼️ 作图
# ===============================
p_fluct <- ggplot2::ggplot(plot_df, ggplot2::aes(x = season, y = med, group = 1)) +
  ggplot2::geom_line(linewidth = 1.1, color = "black") +
  ggplot2::geom_point(
    ggplot2::aes(fill = biomes_clean),
    size = 3.6, shape = 21, color = "black", alpha = 0.95, show.legend = FALSE
  ) +
  ggplot2::scale_fill_manual(values = biome_colors, guide = "none") +
  ggplot2::scale_x_discrete(drop = FALSE) +
  ggplot2::scale_y_continuous(expand = ggplot2::expansion(mult = c(0.16, 0.10))) +
  ggplot2::facet_wrap(~ label, ncol = 4) +
  ggplot2::geom_hline(
    data = annual_for_plot,
    ggplot2::aes(yintercept = med_ann),
    linetype = "dashed", color = "grey30", linewidth = 0.7, inherit.aes = FALSE
  ) +
  ggplot2::geom_text(
    data = annual_for_plot,
    ggplot2::aes(x = "Winter", y = pmin(med_ann, ymax_plot - 1), label = round(med_ann, 1)),
    inherit.aes = FALSE, position = ggplot2::position_nudge(x = right_nudge),
    hjust = 1, vjust = -0.35, size = label_textsize, color = "grey30"
  ) +
  ggplot2::geom_text(
    data = season_sd_for_plot,
    ggplot2::aes(x = "Spring", y = ymax_plot, label = stats_label),
    hjust = 0, vjust = 1.2, size = label_textsize, color = "grey20",
    inherit.aes = FALSE
  ) +
  ggplot2::geom_text(
    data = season_n_for_plot,
    ggplot2::aes(x = season, y = ymin_plot, label = n_label),
    inherit.aes = FALSE, vjust = bottom_vjust, size = n_label_size, color = "grey20"
  ) +
  ggplot2::coord_cartesian(ylim = c(ymin_plot, ymax_plot), clip = "off") +
  ggplot2::theme_bw(base_size = base_fontsize) +
  ggplot2::theme(
    strip.text   = ggplot2::element_text(face = "bold", size = strip_title_size),
    axis.text.x  = ggplot2::element_text(size = axis_tick_size, color = "black"),
    axis.text.y  = ggplot2::element_text(size = axis_tick_size, color = "black"),
    axis.title.x = ggplot2::element_blank(),
    axis.title.y = ggplot2::element_text(size = axis_title_y_sz),
    legend.position = "none",
    plot.margin     = ggplot2::margin(20, 28, 28, 20)
  ) +
  ggplot2::labs(y = expression(Median~CH[4]~uptake~(mu*g~CH[4]-C~m^{-2}~h^{-1})))

ggplot2::ggsave(out_file, plot = p_fluct, width = 20, height = 14, dpi = 600)
message("✅ 已保存图片: ", out_file)

# ===============================
# 📤 Excel 输出（FigS4）
# ===============================
season_wide <- season_stats %>%
  dplyr::select(biomes_clean, season, med) %>%
  tidyr::pivot_wider(names_from = season, values_from = med) %>%
  dplyr::rename(med_Spring = Spring, med_Summer = Summer, med_Autumn = Autumn, med_Winter = Winter)

summary_tbl <- season_wide %>%
  dplyr::left_join(annual_stats %>% dplyr::rename(med_Annual = med_ann), by = "biomes_clean") %>%
  dplyr::left_join(season_sd_tbl %>% dplyr::select(biomes_clean, season_mean, sd_sample, cv_standard_pct), by = "biomes_clean") %>%
  dplyr::mutate(
    Biome = tools::toTitleCase(as.character(biomes_clean)),
    mean_of_seasonal_medians = season_mean,
    sd_of_seasonal_medians   = sd_sample
  ) %>%
  dplyr::select(
    Biome, med_Annual, med_Spring, med_Summer, med_Autumn, med_Winter,
    mean_of_seasonal_medians, sd_of_seasonal_medians, CV_standard_pct = cv_standard_pct
  )

biome_order_out <- tools::toTitleCase(ord_levels)
summary_tbl <- summary_tbl %>%
  dplyr::mutate(Biome = factor(Biome, levels = biome_order_out)) %>%
  dplyr::arrange(Biome) %>%
  dplyr::mutate(Biome = as.character(Biome))

excel_path <- file.path(out_dir, "FigS4_seasonal_medians_summary.xlsx")

if (!requireNamespace("writexl", quietly = TRUE)) install.packages("writexl")
writexl::write_xlsx(
  list(
    summary_by_biome    = summary_tbl,
    season_medians_long = season_stats,
    annual_medians_long = annual_stats
  ),
  path = excel_path
)

message("📘 Excel 已输出到: ", excel_path)
message("✅ 读取的 Studies and Fluxes sheet: ", sheet_to_read)
