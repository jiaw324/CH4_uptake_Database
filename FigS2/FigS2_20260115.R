# ===============================
# 📦 必要包
# ===============================
required_packages <- c("readxl","dplyr","stringr","ggplot2","tidyr","cowplot","scales")
for (pkg in required_packages) if (!requireNamespace(pkg, quietly = TRUE)) install.packages(pkg)
invisible(lapply(required_packages, library, character.only = TRUE))

# ===============================
# 📁 路径（✅已按你要求修改）
# ===============================
ROOT <- "D:/Users/jiaweiChiang/Desktop/Supplymentary_code"

# ✅ 输入文件1：CH4 uptake data_CH4 FLUX_1_1410.xlsx（附件1）——这里只做存在性检查
in_file_flux <- file.path(ROOT, "CH4 uptake data_CH4 FLUX_1_1410.xlsx")
stopifnot(file.exists(in_file_flux))

# ✅ 输入文件2：Studies and Fluxes.xlsx（附件2）——本图的真正数据来源
in_file <- file.path(ROOT, "Studies and Fluxes.xlsx")
stopifnot(file.exists(in_file))

# ✅ 输出路径：FigS2
out_dir <- file.path(ROOT, "FigS2")
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)
setwd(out_dir)

out_file <- file.path(out_dir, "FigS2.jpg")

# ===============================
# 🔎 智能列名匹配
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
# 📥 读取数据（Studies and Fluxes.xlsx）
# ===============================
sheets <- readxl::excel_sheets(in_file)
sheet_to_read <- sheets[1]  # 默认第一张
df <- readxl::read_excel(in_file, sheet = sheet_to_read)

cols <- names(df)

# ✅ 新表头：Ecosystem / Biomes
eco_col <- pick_col(cols, c("Ecosystem","ecosystem","Biomes","biomes","biome","biome_name"), TRUE, "Ecosystem/Biomes")

# ✅ 新表头：Annual
col_Annual <- pick_col(cols, c("Annual","annual","1-12","Year","year"), TRUE, "Annual")

# ===============================
# 🎯 固定生态系统类别（Wetland→Others；新增 Savanna）
# ===============================
target_ecos <- c(
  "Agriculture","Bare","Desert","Forest","Grassland",
  "Others","Rainforest","Savanna","Shrub","Tundra","Urban","Woodland"
)

# 大小写容错 + 仅 Wetland → Others；其余严格匹配
map_ecos <- function(x) {
  x_raw <- as.character(x)
  xl <- stringr::str_to_lower(stringr::str_trim(x_raw))
  
  # wetlands -> others（合并）
  xl[ xl %in% c("wetland","marsh","swamp","bog","fen","peatland") ] <- "others"
  
  # 回到目标拼写
  tl <- stringr::str_to_lower(target_ecos)
  idx <- match(xl, tl)
  out <- ifelse(is.na(idx), NA_character_, target_ecos[idx])
  out
}

# ===============================
# 🎨 颜色（含 Savanna）
# ===============================
group_colors <- c(
  "Agriculture"="#8AAE5D",
  "Bare"       ="#9E9E9E",
  "Desert"     ="#F2C300",
  "Forest"     ="#2E7D32",
  "Grassland"  ="#2E6F89",
  "Others"     ="#BDBDBD",
  "Rainforest" ="#C4A46A",
  "Savanna"    ="#DAA520",
  "Shrub"      ="#A34A38",
  "Tundra"     ="#90CAF9",
  "Urban"      ="#6B575A",
  "Woodland"   ="#CD7A50"
)

# ===============================
# 📊 提取 Annual & 映射
# ===============================
annual_df <- df %>%
  dplyr::transmute(
    group = map_ecos(.data[[eco_col]]),
    value = suppressWarnings(as.numeric(.data[[col_Annual]]))
  ) %>%
  dplyr::filter(!is.na(group), !is.na(value)) %>%
  dplyr::mutate(group = factor(group, levels = target_ecos))

# 过滤：吸收 > 0 且 ≤ 200
annual_df <- annual_df %>% dplyr::filter(value > 0, value <= 200)

# ===============================
# 📐 左→右按“中位数”升序
# ===============================
stat_by_group <- annual_df %>%
  dplyr::group_by(group) %>%
  dplyr::summarise(
    n   = dplyr::n(),
    med = stats::median(value, na.rm = TRUE),
    .groups = "drop"
  )

final_groups <- stat_by_group %>%
  dplyr::arrange(med) %>%
  dplyr::pull(group) %>%
  as.character()

stat_by_group <- stat_by_group %>%
  dplyr::mutate(
    group = factor(group, levels = final_groups),
    # ✅ 数据量标签只保留数字： " (n=12)" -> " (12)"
    xlab  = paste0(as.character(group), " (", n, ")")
  )

annual_df2 <- annual_df %>%
  dplyr::mutate(group = factor(group, levels = final_groups)) %>%
  dplyr::left_join(
    stat_by_group %>% dplyr::select(group, n, med, xlab),
    by = "group"
  )

# 三分支（按 n）
df_pts <- annual_df2 %>% dplyr::filter(n < 10)
df_box <- annual_df2 %>% dplyr::filter(n >= 10 & n < 20)
df_vln <- annual_df2 %>% dplyr::filter(n >= 20)

# ===============================
# 🖌️ 绘图（中位数黑横线；均值“×”）
# ===============================
p <- ggplot(annual_df2, aes(x = group, y = value, fill = group)) +
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
    aes(x = group, y = med, ymin = med, ymax = med),
    inherit.aes = FALSE,
    width = 0.7, linewidth = 0.5, color = "black"
  ) +
  scale_x_discrete(
    limits = final_groups,
    labels = setNames(stat_by_group$xlab, stat_by_group$group)
  ) +
  scale_fill_manual(values = group_colors, drop = FALSE) +
  scale_y_continuous(
    limits = c(-5, 200),
    expand = expansion(mult = c(0.02, 0.08))
  ) +
  theme_bw(base_size = 18) +
  theme(
    legend.position = "none",
    axis.text.x = element_text(
      angle = 35, hjust = 1, size = 14, color = "black"
    ),
    axis.text.y = element_text(size = 14, color = "black"),
    axis.title.x = element_blank(),
    plot.margin  = margin(20, 20, 30, 20)
  ) +
  labs(
    y = expression(
      CH[4]~uptake~(mu*g~CH[4]-C~m^{-2}~h^{-1})
    )
  )

# 平均值“×”
mean_by_group <- annual_df2 %>%
  dplyr::group_by(group) %>%
  dplyr::summarise(mu = mean(value, na.rm = TRUE), .groups = "drop")

p <- p +
  geom_point(
    data = mean_by_group,
    aes(x = group, y = mu),
    inherit.aes = FALSE,
    shape = 4, size = 3.2, stroke = 0.8, color = "black"
  )

# ===============================
# ✅ 导出
# ===============================
ggsave(out_file, plot = p, width = 14, height = 8, dpi = 600)
message("✅ 图已保存: ", out_file)