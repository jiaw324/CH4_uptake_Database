# ===============================
# 📦 必要包
# ===============================
required_packages <- c("readxl","dplyr","stringr","ggplot2","tidyr","cowplot","scales")
for (pkg in required_packages) if (!requireNamespace(pkg, quietly = TRUE)) install.packages(pkg)
invisible(lapply(required_packages, library, character.only = TRUE))

# ===============================
# 📁 路径（按你要求）
# ===============================
in_dir <- "D:/Users/jiaweiChiang/Desktop/Supplymentary_code"

# ✅ 输入文件1：只检查存在（不读取）
in_file_flux <- file.path(in_dir, "CH4 uptake data_CH4 FLUX_1_1410.xlsx")
stopifnot(file.exists(in_file_flux))

# ✅ 主数据源：Studies and Fluxes.xlsx
in_file <- file.path(in_dir, "Studies and Fluxes.xlsx")
stopifnot(file.exists(in_file))

# ✅ 输出文件夹（Count）
work_dir <- file.path(in_dir, "FigS3")
if (!dir.exists(work_dir)) dir.create(work_dir, recursive = TRUE)
setwd(work_dir)

out_file_merge    <- file.path(work_dir, "fig4_Climate_violin_TemperatePlusSubtropicMerged.jpg")
out_file_separate <- file.path(work_dir, "fig4_Climate_violin_Temperate_Subtropic_Separate.jpg")

# ===============================
# 🔎 智能列名匹配（与你之前一致）
# ===============================
pick_col <- function(nms, candidates, must=TRUE, label="column") {
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
  if (must) stop(sprintf("找不到【%s】列。表头示例：%s", label, paste(head(nms, 30), collapse=", ")), call. = FALSE)
  NA_character_
}

# ===============================
# 📖 读取 Studies and Fluxes.xlsx
# ===============================
sheets <- readxl::excel_sheets(in_file)

# ✅ 优先读 ST，其次 row，否则第一个 sheet
sheet_to_read <- if ("ST" %in% sheets) "ST" else if ("row" %in% sheets) "row" else sheets[1]
df <- readxl::read_excel(in_file, sheet = sheet_to_read)

cols <- names(df)

# Climate 与 Annual 列
clim_col   <- pick_col(cols, c("Climate","climate","climate_class","climate zone","climate_zone","koppen","köppen","koppen_geiger"), TRUE, "Climate")
col_Annual <- pick_col(cols, c("Annual","annual","1-12","Year","year","Annual_annual","Sea_Annual"), TRUE, "Annual")

# （可选）management 列：新版表头里叫 management
management_col <- pick_col(cols, c("management","Management","control"), must = FALSE, label = "management")

# ===============================
# 🗺️ 目标等级 & 颜色
# ===============================
# 分开版本的 5 类
levels_separate <- c("Alpine","Boreal","Subtropic","Temperate","Tropic")
colors_separate <- c(
  "Alpine"    = "#7f8c8d",
  "Boreal"    = "#756bb1",
  "Subtropic" = "#41b6c4",
  "Temperate" = "#3182bd",
  "Tropic"    = "#2ca25f"
)

# 合并版本（Subtropic -> Temperate），4 类
levels_merged <- c("Alpine","Boreal","Temperate","Tropic")
colors_merged <- c(
  "Alpine"    = "#7f8c8d",
  "Boreal"    = "#756bb1",
  "Temperate" = "#3182bd",
  "Tropic"    = "#2ca25f"
)

# ===============================
# ✅ 基础清洗
# - 只用 Annual 数值与 Climate 文本
# - value > 0 且 <= 200
# - （可选）如果存在 management 列，则只保留 control（你之前常用）
# ===============================
base_df <- df %>%
  dplyr::transmute(
    climate_raw = stringr::str_to_title(stringr::str_trim(as.character(.data[[clim_col]]))),
    value       = suppressWarnings(as.numeric(.data[[col_Annual]])),
    management  = if (!is.na(management_col)) stringr::str_to_lower(stringr::str_trim(as.character(.data[[management_col]]))) else NA_character_
  ) %>%
  dplyr::filter(!is.na(value)) %>%
  dplyr::filter(value > 0, value <= 200)

# ✅ 如果 management 列存在，就只保留 control（不区分大小写）
if (!all(is.na(base_df$management))) {
  base_df <- base_df %>% dplyr::filter(management == "control")
}

# ===============================
# 🌍 Climate 五类归一
# ===============================
normalize5 <- function(x) {
  x0 <- stringr::str_to_lower(stringr::str_trim(as.character(x)))
  out <- dplyr::case_when(
    grepl("^alpin", x0)                     ~ "Alpine",
    grepl("^boreal|subarctic|taiga", x0)    ~ "Boreal",
    grepl("^subtrop", x0)                   ~ "Subtropic",
    grepl("^temper", x0)                    ~ "Temperate",
    grepl("^tropic|equator", x0)            ~ "Tropic",
    TRUE ~ NA_character_
  )
  out
}

# ===============================
# 🎻 通用绘图函数
# ===============================
plot_violin <- function(dat, group_col, levels_keep, color_map, outfile){
  stopifnot(group_col %in% names(dat))
  dat <- dat %>% dplyr::filter(!is.na(.data[[group_col]]))
  
  if (nrow(dat) == 0) stop("过滤后无可用数据。")
  
  dat[[group_col]] <- factor(dat[[group_col]], levels = levels_keep)
  
  # 计算排序依据（中位数）
  stat_by_group <- dat %>%
    dplyr::group_by(.data[[group_col]]) %>%
    dplyr::summarise(
      n   = dplyr::n(),
      med = stats::median(value, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    dplyr::filter(!is.na(.data[[group_col]]))
  
  if (nrow(stat_by_group) == 0) stop("没有任何分组含数据。")
  
  final_groups <- stat_by_group %>%
    dplyr::arrange(med) %>%
    dplyr::pull(1) %>%
    as.character()
  
  stat_by_group <- stat_by_group %>%
    dplyr::mutate(
      !!group_col := factor(.data[[group_col]], levels = final_groups),
      xlab = paste0(as.character(.data[[group_col]]), " (", n, ")")
    )
  
  dat2 <- dat %>%
    dplyr::mutate(!!group_col := factor(.data[[group_col]], levels = final_groups)) %>%
    dplyr::left_join(
      stat_by_group %>% dplyr::select(1, n, med, xlab),
      by = group_col
    )
  
  # 三分支（按 n）
  df_pts <- dat2 %>% dplyr::filter(n < 10)
  df_box <- dat2 %>% dplyr::filter(n >= 10 & n < 20)
  df_vln <- dat2 %>% dplyr::filter(n >= 20)
  
  # 平均值
  mean_by_group <- dat2 %>%
    dplyr::group_by(.data[[group_col]]) %>%
    dplyr::summarise(mu = mean(value, na.rm = TRUE), .groups = "drop")
  
  aes_group <- aes_string(x = group_col, y = "value", fill = group_col)
  
  p <- ggplot(dat2, aes_group) +
    { if (nrow(df_vln) > 0)
      geom_violin(
        data  = df_vln,
        scale = "width", width = 0.9,
        color = "black", alpha = 0.9, adjust = 3
      )
    } +
    { if (nrow(df_box) > 0)
      geom_boxplot(
        data = df_box, width = 0.6,
        outlier.shape  = 21,
        outlier.size   = 2.0,
        outlier.stroke = 0.3,
        color = "black", alpha = 0.9
      )
    } +
    { if (nrow(df_pts) > 0)
      geom_point(
        data = df_pts,
        position = position_jitter(width = 0.12, height = 0),
        shape = 21, size = 2.6, stroke = 0.3,
        color = "black", alpha = 0.9
      )
    } +
    geom_crossbar(
      data = stat_by_group,
      aes_string(x = group_col, y = "med", ymin = "med", ymax = "med"),
      inherit.aes = FALSE,
      width = 0.7, linewidth = 0.5, color = "black"
    ) +
    scale_x_discrete(
      limits = final_groups,
      labels = setNames(stat_by_group$xlab, stat_by_group[[group_col]])
    ) +
    scale_fill_manual(values = color_map, drop = FALSE) +
    scale_y_continuous(
      limits = c(-5, 200),
      expand = expansion(mult = c(0.02, 0.08))
    ) +
    theme_bw(base_size = 18) +
    theme(
      legend.position = "none",
      axis.text.x = element_text(angle = 35, hjust = 1, size = 14, color = "black"),
      axis.text.y = element_text(size = 14, color = "black"),
      axis.title.x = element_blank(),
      plot.margin  = margin(20, 20, 30, 20)
    ) +
    labs(
      y = expression(CH[4]~uptake~(mu*g~CH[4]-C~m^{-2}~h^{-1}))
    ) +
    geom_point(
      data = mean_by_group,
      aes_string(x = group_col, y = "mu"),
      inherit.aes = FALSE,
      shape = 4, size = 3.2, stroke = 0.8, color = "black"
    )
  
  ggsave(outfile, plot = p, width = 14, height = 8, dpi = 600)
  message("✅ 图已保存: ", outfile)
}

# ===============================
# ▶️ 构造两套分组并作图
# ===============================

# 1) 不合并的版本：5 类
df_sep <- base_df %>%
  dplyr::mutate(Climate5 = normalize5(climate_raw)) %>%
  dplyr::filter(!is.na(Climate5))

plot_violin(df_sep, "Climate5", levels_separate, colors_separate, out_file_separate)

# 2) 合并版本：Subtropic 并入 Temperate → 4 类
df_merge <- df_sep %>%
  dplyr::mutate(
    ClimateMerge = dplyr::recode(
      Climate5,
      "Subtropic" = "Temperate",
      .default    = Climate5
    )
  )

plot_violin(df_merge, "ClimateMerge", levels_merged, colors_merged, out_file_merge)