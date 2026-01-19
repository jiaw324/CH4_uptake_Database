# ================================
# 📦 必要包
# ================================
required_packages <- c("readxl","dplyr","ggplot2","plotbiomes","cowplot","stringr","tidyr","sf","readr")
for (pkg in required_packages) if (!requireNamespace(pkg, quietly = TRUE)) install.packages(pkg)
invisible(lapply(required_packages, library, character.only = TRUE))

# ================================
# 📁 路径（输出固定到：D:/Users/jiaweiChiang/Desktop/Supplymentary_code/Fig3）
# ================================
ROOT     <- "D:/Users/jiaweiChiang/Desktop/Supplymentary_code"
in_dir   <- ROOT
out_dir  <- file.path(ROOT, "Fig3")
work_dir <- out_dir
if (!dir.exists(work_dir)) dir.create(work_dir, recursive = TRUE)
setwd(work_dir)

# ================================
# 📄 输入文件：Studies and Fluxes（优先精确匹配 Studies and Fluxes.xlsx）
# ================================
preferred <- file.path(in_dir, "Studies and Fluxes.xlsx")
if (file.exists(preferred)) {
  in_file <- preferred
} else {
  cand <- list.files(in_dir, pattern = "^Studies and Fluxes.*\\.xlsx$", full.names = TRUE, ignore.case = TRUE)
  if (length(cand) == 0) stop("在目录中找不到 'Studies and Fluxes*.xlsx'。请确认文件放在：", in_dir, call. = FALSE)
  in_file <- cand[1]
}
message("Using input file: ", in_file)

# ================================
# 🔎 列名匹配函数
# ================================
pick_col <- function(nms, candidates, must = TRUE, label = "column") {
  nms <- trimws(nms)
  nms_l <- tolower(nms)
  
  # 1) 精确匹配
  cand_l <- tolower(trimws(candidates))
  hit <- cand_l[cand_l %in% nms_l]
  if (length(hit) > 0) return(nms[match(hit[1], nms_l)])
  
  # 2) 单词边界匹配
  for (pat in cand_l) {
    idx <- grepl(paste0("\\b", pat, "\\b"), nms_l)
    if (any(idx)) return(nms[which(idx)[1]])
  }
  
  # 3) 模糊包含匹配
  idx <- Reduce(`|`, lapply(cand_l, function(p) grepl(p, nms_l)))
  if (any(idx)) return(nms[which(idx)[1]])
  
  if (must) stop(sprintf("找不到【%s】列。表头示例：%s", label, paste(head(nms, 40), collapse = ", ")), call. = FALSE)
  NA_character_
}

# ================================
# 🧭 自动选对 sheet：必须包含 Biome/Biomes + MAT + MAP + Q01/Q02/Q03 + management
# ================================
sheets <- readxl::excel_sheets(in_file)

score_sheet <- function(sh) {
  hdr <- tryCatch({
    suppressMessages(readxl::read_excel(in_file, sheet = sh, n_max = 1))
  }, error = function(e) NULL)
  if (is.null(hdr)) return(-999)
  
  nms <- trimws(names(hdr))
  nms_l <- tolower(nms)
  
  # ✅ 兼容 Biome / Biomes
  has_biome <- any(nms_l %in% c("biome", "biomes"))
  
  has_mat    <- any(nms_l %in% c("mat_2 (era5)", "mat_1 (original)", "mat"))
  has_map    <- any(nms_l %in% c("map_2 (era5)", "map_1 (original)", "map"))
  has_q01    <- any(nms_l == "q01")
  has_q02    <- any(nms_l == "q02")
  has_q03    <- any(nms_l == "q03")
  has_mgmt   <- any(nms_l == "management")
  has_pid    <- any(nms_l == "paper_number")
  has_sid    <- any(nms_l == "study_number")
  
  # 打分：越符合越高
  score <- 0
  score <- score + ifelse(has_biome, 5, 0)
  score <- score + ifelse(has_mat,    2, 0)
  score <- score + ifelse(has_map,    2, 0)
  score <- score + ifelse(has_q01,    2, 0)
  score <- score + ifelse(has_q02,    2, 0)
  score <- score + ifelse(has_q03,    2, 0)
  score <- score + ifelse(has_mgmt,   2, 0)
  score <- score + ifelse(has_pid,    1, 0)
  score <- score + ifelse(has_sid,    1, 0)
  
  # 若关键列缺失则降权
  if (!has_biome) score <- score - 10
  if (!has_mat)   score <- score - 5
  if (!has_map)   score <- score - 5
  if (!has_q01)   score <- score - 5
  if (!has_q02)   score <- score - 5
  if (!has_q03)   score <- score - 5
  if (!has_mgmt)  score <- score - 5
  
  score
}

scores <- vapply(sheets, score_sheet, numeric(1))
best_i <- which.max(scores)
sheet_to_read <- sheets[best_i]

if (length(sheet_to_read) == 0 || is.infinite(scores[best_i]) || scores[best_i] < 0) {
  stop("没有找到包含 Biome/MAT/MAP/Q01/Q02/Q03/management 的 sheet。请检查工作簿内是否有正确数据表。", call. = FALSE)
}
message("Auto-selected sheet: ", sheet_to_read)

df <- readxl::read_excel(in_file, sheet = sheet_to_read)

# ================================
# 🔎 识别你指定的表头（Biome/Biomes 都可）
# ================================
cols <- trimws(names(df))

# ✅ 关键修改：Biomes -> Biome（兼容两者）
biome_col <- pick_col(cols, c("Biome", "Biomes"), TRUE, "Biome/Biomes")

MAT_col   <- pick_col(cols, c("MAT_2 (ERA5)", "MAT_1 (original)", "MAT"), TRUE, "MAT")
MAP_col   <- pick_col(cols, c("MAP_2 (ERA5)", "MAP_1 (original)", "MAP"), TRUE, "MAP")

# 你要求的列名映射：laboratory->Q01 repeat->Q02 dataquality->Q03 control->management
q01_col   <- pick_col(cols, c("Q01"), TRUE, "Q01")
q02_col   <- pick_col(cols, c("Q02"), TRUE, "Q02")
q03_col   <- pick_col(cols, c("Q03"), TRUE, "Q03")
mgmt_col  <- pick_col(cols, c("management"), TRUE, "management")

# 可选列：Paper/Study/Lat/Lon（缺失不报错）
paper_col <- pick_col(cols, c("Paper_number"), FALSE, "Paper_number")
study_col <- pick_col(cols, c("Study_number"), FALSE, "Study_number")
lat_col   <- pick_col(cols, c("Latitude"), FALSE, "Latitude")
lon_col   <- pick_col(cols, c("Longitude"), FALSE, "Longitude")

message(sprintf(
  "使用列名：Biome=%s, MAT=%s, MAP=%s, Q01=%s, Q02=%s, Q03=%s, management=%s ; sheet=%s",
  biome_col, MAT_col, MAP_col, q01_col, q02_col, q03_col, mgmt_col, sheet_to_read
))

# ================================
# 🔁 统一到 Whittaker 官方类目
# ================================
map_to_whittaker <- function(x) {
  x <- stringr::str_trim(as.character(x))
  dplyr::recode(x,
                "Tundra"="Tundra",
                "Boreal forest"="Boreal forest",
                "Temperate seasonal forest"="Temperate seasonal forest",
                "Temperate rain forest"="Temperate rain forest",
                "Tropical rain forest"="Tropical rain forest",
                "Tropical seasonal forest/savanna"="Tropical seasonal forest/savanna",
                "Subtropical desert"="Subtropical desert",
                "Temperate grassland/desert"="Temperate grassland/desert",
                "Woodland/shrubland"="Woodland/shrubland",
                
                # 常见同义/变体
                "Temperate rainforest"="Temperate rain forest",
                "Tropical rainforest"="Tropical rain forest",
                "Tropical seasonal forest"="Tropical seasonal forest/savanna",
                "Savanna"="Tropical seasonal forest/savanna",
                "Desert"="Temperate grassland/desert",
                "Temperate grassland"="Temperate grassland/desert",
                "Woodland"="Woodland/shrubland",
                "Shrubland"="Woodland/shrubland",
                .default=NA_character_
  )
}

# ================================
# 📋 小图生态系统（去掉 Wetland）
# ================================
inset_biomes <- c("Alpine","Agriculture","Bare","Urban","Others")

# ================================
# 📊 小图数据
# ================================
inset_df <- df %>%
  mutate(
    biomes = stringr::str_trim(.data[[biome_col]]),
    MAT = suppressWarnings(as.numeric(.data[[MAT_col]])),
    MAP = suppressWarnings(as.numeric(.data[[MAP_col]])),
    MAP_cm = MAP / 10,
    MAT_r = round(MAT, 1),
    MAP_r = round(MAP, 0)
  ) %>%
  filter(biomes %in% inset_biomes, !is.na(MAT), !is.na(MAP_cm)) %>%
  mutate(biomes = factor(biomes, levels = inset_biomes))

used_points_coords <- inset_df %>% distinct(MAT_r, MAP_r)

# ================================
# ✅ 主图数据（按 Q01/Q02/Q03/management 筛选）
# ================================
main_df <- df %>%
  mutate(
    Q01 = trimws(as.character(.data[[q01_col]])),
    Q02 = trimws(as.character(.data[[q02_col]])),
    Q03 = trimws(as.character(.data[[q03_col]])),
    management = trimws(as.character(.data[[mgmt_col]])),
    
    Latitude     = if (!is.na(lat_col)   && nzchar(lat_col))   suppressWarnings(as.numeric(.data[[lat_col]]))   else NA_real_,
    Longitude    = if (!is.na(lon_col)   && nzchar(lon_col))   suppressWarnings(as.numeric(.data[[lon_col]]))   else NA_real_,
    Paper_number = if (!is.na(paper_col) && nzchar(paper_col)) .data[[paper_col]] else NA,
    Study_number = if (!is.na(study_col) && nzchar(study_col)) .data[[study_col]] else NA
  ) %>%
  filter(
    is.na(Q01) | Q01 != "√",
    is.na(Q02) | Q02 != "√",
    is.na(Q03) | Q03 != "√",
    tolower(management) == "control"
  ) %>%
  mutate(
    MAT = suppressWarnings(as.numeric(.data[[MAT_col]])),
    MAP = suppressWarnings(as.numeric(.data[[MAP_col]])),
    temperature   = MAT,
    precipitation = MAP / 10,
    biomes_raw    = stringr::str_trim(.data[[biome_col]]),
    biomes        = map_to_whittaker(biomes_raw),
    MAT_r = round(MAT, 1),
    MAP_r = round(MAP, 0)
  ) %>%
  filter(
    !is.na(biomes),
    !is.na(temperature), !is.na(precipitation),
    dplyr::between(temperature, -15, 30),
    dplyr::between(precipitation, 0, 450)
  ) %>%
  anti_join(used_points_coords, by = c("MAT_r","MAP_r"))

# ================================
# 🎨 颜色
# ================================
main_colors <- c(
  "Tundra"="#a7d3df","Boreal forest"="#7fbf7f","Temperate seasonal forest"="#8cbf70",
  "Temperate rain forest"="#5a9d5a","Tropical rain forest"="#157f1d",
  "Tropical seasonal forest/savanna"="#a5891d",
  "Subtropical desert"="#c6a857","Temperate grassland/desert"="#c49b50",
  "Woodland/shrubland"="#cd7a50"
)
inset_colors <- c("Alpine"="#b27874","Agriculture"="#886b99","Bare"="#357266","Urban"="#6b575a","Others"="#b3b5ae")

# ================================
# 📈 主图 & 小图
# ================================
p1 <- ggplot() +
  geom_polygon(
    data = plotbiomes::Whittaker_biomes,
    aes(x = temp_c, y = precp_cm, group = biome, fill = biome),
    color = "white", linewidth = 0.6, alpha = 0.8
  ) +
  geom_point(
    data = main_df,
    aes(x = temperature, y = precipitation),
    shape = 21, fill = "gray60", color = "black",
    size = 2.8, stroke = 0.5, alpha = 0.85
  ) +
  scale_fill_manual(values = main_colors, name = NULL, drop = FALSE) +
  coord_cartesian(xlim = c(-15, 30), ylim = c(0, 450)) +
  labs(x = "MAT (°C)", y = "MAP (cm)") +
  theme_bw(base_size = 20) +
  theme(
    panel.grid = element_blank(),
    panel.border = element_rect(linewidth = 1.2),
    legend.position = c(0.17, 0.43),
    legend.background = element_rect(fill = "transparent", color = NA),
    legend.key = element_rect(fill = "transparent", color = NA),
    legend.text = element_text(size = 13),
    legend.key.size = unit(1, "lines"),
    axis.title = element_text(size = 20),
    axis.text = element_text(size = 14)
  )

p2 <- ggplot(inset_df, aes(x = MAT, y = MAP_cm, fill = biomes)) +
  geom_point(shape = 21, size = 3.5, stroke = 0.4, color = "black", alpha = 0.7) +
  scale_fill_manual(values = inset_colors, name = NULL, drop = FALSE) +
  coord_cartesian(xlim = c(-15, 30), ylim = c(-5, 350)) +
  theme_bw(base_size = 16) +
  theme(
    panel.border = element_rect(color = "black", linewidth = 0.6),
    panel.grid = element_blank(),
    panel.background = element_rect(fill = NA, color = NA),
    plot.background = element_rect(fill = NA, color = NA),
    axis.title = element_blank(),
    axis.text = element_text(size = 10),
    axis.ticks = element_line(color = "black"),
    axis.line = element_line(color = "black", linewidth = 0.4),
    legend.position = c(0.03, 0.95),
    legend.justification = c(0, 1),
    legend.background = element_rect(fill = "transparent", color = NA),
    legend.key = element_rect(fill = "transparent", color = NA),
    legend.text = element_text(size = 11)
  )

# ================================
# 🧩 合并 & 输出（输出到 Fig3）
# ================================
final_plot <- cowplot::ggdraw() +
  cowplot::draw_plot(p1, 0, 0, 1, 1) +
  cowplot::draw_plot(p2, x = 0.073, y = 0.60, width = 0.50, height = 0.38)

out_file <- file.path(out_dir, "Fig3.jpg")
ggsave(out_file, final_plot, width = 12, height = 8, dpi = 600)
message("Done: ", out_file)
print(final_plot)
