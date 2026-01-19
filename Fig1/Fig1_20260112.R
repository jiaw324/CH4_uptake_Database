# ================== English / Chinese；右轴 CH[4]；深绿色半透明折线；仅柱体图例 ==================
# 筛选：去掉 Q01=="√"、Q02=="√"、Q03=="√"；仅保留 management=="control"
suppressPackageStartupMessages({
  library(readxl); library(dplyr); library(tidyr); library(ggplot2)
  library(readr);  library(stringr)
})

# ===== 路径设置 =====
ROOT    <- "D:/Users/jiaweiChiang/Desktop/Supplymentary_code"
FIG1DIR <- file.path(ROOT, "fig1")   # ✅ 你指定的输出目录 + CH4 文件目录

in_dir   <- ROOT
work_dir <- ROOT
out_dir  <- FIG1DIR                  # ✅ 输出改到 fig1

if (!dir.exists(work_dir)) dir.create(work_dir, recursive = TRUE)
if (!dir.exists(out_dir))  dir.create(out_dir,  recursive = TRUE)

setwd(work_dir)

# ===== 工具：强列名匹配（忽略符号；支持前缀/包含匹配）=====
norm_key <- function(x) tolower(gsub("[^a-z0-9]+", "", as.character(x)))

pick_col <- function(nms, candidates) {
  nms0 <- as.character(nms)
  nk <- norm_key(nms0)
  for (cand in candidates) {
    ck <- norm_key(cand)
    
    hit <- which(nk == ck)
    if (length(hit) > 0) return(nms0[hit[1]])
    
    hit <- which(startsWith(nk, ck))
    if (length(hit) > 0) return(nms0[hit[1]])
    
    hit <- which(grepl(ck, nk, fixed = TRUE))
    if (length(hit) > 0) return(nms0[hit[1]])
  }
  return(NA_character_)
}

need_candidates <- list(
  Paper_number   = c("Paper_number","Paper","paper","PaperID","papernumber"),
  Paper_year     = c("Paper_year","Year","paper_year","paperyear"),
  Paper_language = c("Paper_language","Language","paperlanguage"),
  Q01            = c("Q01","laboratory"),
  Q02            = c("Q02","repeat"),
  Q03            = c("Q03","dataquality"),
  management     = c("management","control")
)

has_required_cols <- function(colnames_vec) {
  nms <- as.character(colnames_vec)
  found <- vapply(names(need_candidates), function(k){
    !is.na(pick_col(nms, need_candidates[[k]]))
  }, logical(1))
  all(found)
}

# =============================================================================
# 1) 自动定位输入文件：Studies and Fluxes*.xlsx（扫描所有sheet找含必要列的）
# =============================================================================
cand_in <- list.files(in_dir, pattern = "(?i)^Studies and Fluxes.*\\.xlsx$", full.names = TRUE)
if (length(cand_in) == 0) stop("未找到输入文件：Studies and Fluxes*.xlsx\n目录：", in_dir)

score_file <- function(fp){
  bn <- tolower(basename(fp))
  sc <- 0
  if (bn == "studies and fluxes.xlsx") sc <- sc + 1000
  if (grepl("unit", bn)) sc <- sc - 50
  if (grepl("副本|copy", bn)) sc <- sc - 20
  sc
}
cand_in <- cand_in[order(vapply(cand_in, score_file, numeric(1)), decreasing = TRUE)]

picked_file  <- NA_character_
picked_sheet <- NA_character_

for (f in cand_in) {
  sheets <- tryCatch(readxl::excel_sheets(f), error = function(e) character(0))
  if (length(sheets) == 0) next
  for (sh in sheets) {
    hdr <- tryCatch(readxl::read_excel(f, sheet = sh, n_max = 0), error = function(e) NULL)
    if (is.null(hdr)) next
    if (has_required_cols(names(hdr))) {
      picked_file  <- f
      picked_sheet <- sh
      break
    }
  }
  if (!is.na(picked_file)) break
}

if (is.na(picked_file)) {
  stop("在所有 Studies and Fluxes*.xlsx 的所有 sheet 中，都没找到同时包含这些列：\n",
       paste(names(need_candidates), collapse = ", "))
}

message("✅ Using file : ", picked_file)
message("✅ Using sheet: ", picked_sheet)

df <- readxl::read_excel(picked_file, sheet = picked_sheet)

# ===== 正式解析列名 =====
nms <- names(df)
col_pnum <- pick_col(nms, need_candidates$Paper_number)
col_year <- pick_col(nms, need_candidates$Paper_year)
col_lang <- pick_col(nms, need_candidates$Paper_language)
col_Q01  <- pick_col(nms, need_candidates$Q01)
col_Q02  <- pick_col(nms, need_candidates$Q02)
col_Q03  <- pick_col(nms, need_candidates$Q03)
col_mgt  <- pick_col(nms, need_candidates$management)

need <- c(Paper_number=col_pnum, Paper_year=col_year, Paper_language=col_lang,
          Q01=col_Q01, Q02=col_Q02, Q03=col_Q03, management=col_mgt)
if (any(is.na(need))) {
  miss <- names(need)[is.na(need)]
  stop("缺失必要列：", paste(miss, collapse=", "),
       "\n当前表头：\n", paste(names(df), collapse=", "))
}

# =============================================================================
# 2) 过滤 + 语言映射
# =============================================================================
df_filtered <- df %>%
  mutate(
    Q01 = trimws(as.character(.data[[col_Q01]])),
    Q02 = trimws(as.character(.data[[col_Q02]])),
    Q03 = trimws(as.character(.data[[col_Q03]])),
    management = trimws(as.character(.data[[col_mgt]])),
    Paper_language = trimws(as.character(.data[[col_lang]]))
  ) %>%
  filter(
    is.na(Q01) | Q01 != "√",
    is.na(Q02) | Q02 != "√",
    is.na(Q03) | Q03 != "√",
    !is.na(management) & tolower(management) == "control"
  ) %>%
  mutate(
    language = case_when(
      tolower(Paper_language) == "chinese" ~ "Chinese",
      tolower(Paper_language) == "english" ~ "English",
      TRUE ~ NA_character_
    )
  ) %>%
  filter(!is.na(language)) %>%
  distinct(.data[[col_pnum]], .keep_all = TRUE) %>%
  mutate(language = factor(language, levels = c("English", "Chinese")))

# =============================================================================
# 3) 年份 × 语言计数（补齐 1989–2024）
# =============================================================================
year_min  <- 1989L
year_max  <- 2024L
all_years <- seq.int(year_min, year_max)

df_count <- df_filtered %>%
  mutate(Paper_year = suppressWarnings(as.integer(.data[[col_year]]))) %>%
  filter(!is.na(Paper_year), Paper_year <= year_max, Paper_year >= year_min) %>%
  group_by(Paper_year, language) %>%
  summarise(count = n(), .groups = "drop") %>%
  complete(Paper_year = all_years, language, fill = list(count = 0)) %>%
  mutate(i = match(Paper_year, all_years)) %>%
  arrange(Paper_year, language)

year_total <- df_count %>%
  group_by(Paper_year) %>%
  summarise(total_count = sum(count), .groups = "drop") %>%
  mutate(i = match(Paper_year, all_years))

# =============================================================================
# 4) 读取 CH4（ppb）：在 fig1 目录下找 ch4_con_1980_2025(.xlsx/.csv)
# =============================================================================
cand_ch4 <- list.files(
  FIG1DIR,
  pattern = "(?i)^ch4_con_1980_2025(.*)?\\.(xlsx|csv)$",
  full.names = TRUE
)
if (length(cand_ch4) == 0) stop("未找到 CH4 文件：ch4_con_1980_2025.xlsx/csv\n目录：", FIG1DIR)

ch4_file <- cand_ch4[1]
message("✅ CH4 file: ", ch4_file)

ch4_raw <- if (grepl("(?i)\\.xlsx$", ch4_file)) {
  readxl::read_excel(ch4_file)
} else {
  readr::read_csv(ch4_file, show_col_types = FALSE)
}

# 自动识别 year 与 ppb 列
ch4_nms <- names(ch4_raw)
ch4_year_col <- pick_col(ch4_nms, c("year","Year","YYYY","yr"))
ch4_ppb_col  <- pick_col(ch4_nms, c(
  "ch4 concentration","ch4_concentration","ch4_ppb","ch4ppb",
  "ch4 mole fraction","mole fraction","concentration","ppb","ch4"
))
if (is.na(ch4_year_col)) stop("CH4 表无法识别年份列。CH4 表头：\n", paste(ch4_nms, collapse=", "))
if (is.na(ch4_ppb_col))  stop("CH4 表无法识别浓度列。CH4 表头：\n", paste(ch4_nms, collapse=", "))
message("CH4 year col: ", ch4_year_col)
message("CH4 ppb  col: ", ch4_ppb_col)

ch4 <- ch4_raw %>%
  transmute(
    Paper_year = suppressWarnings(as.integer(.data[[ch4_year_col]])),
    CH4_ppb    = suppressWarnings(as.numeric(gsub("[^0-9\\.-]", "", as.character(.data[[ch4_ppb_col]]))))
  ) %>%
  filter(!is.na(Paper_year), !is.na(CH4_ppb)) %>%
  filter(Paper_year >= year_min, Paper_year <= year_max) %>%
  distinct(Paper_year, .keep_all = TRUE) %>%
  arrange(Paper_year) %>%
  mutate(i = match(Paper_year, all_years))

if (nrow(ch4) == 0) stop("CH4 浓度数据在 ", year_min, "–", year_max, " 范围内为空。")

# =============================================================================
# 5) 折线映射到左轴 + 右轴 sec.axis
# =============================================================================
ymax_left <- max(year_total$total_count, na.rm = TRUE)
if (!is.finite(ymax_left) || ymax_left <= 0) stop("左轴最大值异常：", ymax_left)

ppb_min <- min(ch4$CH4_ppb, na.rm = TRUE)
ppb_max <- max(ch4$CH4_ppb, na.rm = TRUE)
if (!is.finite(ppb_min) || !is.finite(ppb_max) || ppb_max <= ppb_min) stop("ppb 范围异常。")

a <- (ymax_left - 0) / (ppb_max - ppb_min)
b <- -a * ppb_min
ch4 <- ch4 %>% mutate(y_left_scaled = a * CH4_ppb + b)

# =============================================================================
# 6) 绘图 + 输出到 fig1
# =============================================================================
fill_colors <- c("English" = "#ED702D", "Chinese" = "#F6B403")
line_color  <- "#006400"

bar_width <- 0.7
gap <- 1 - bar_width
left_edge  <- 1 - bar_width/2 - gap
n_years    <- length(all_years)
right_edge <- n_years + bar_width/2 + gap

p <- ggplot(df_count, aes(x = i, y = count, fill = language)) +
  geom_col(width = bar_width, color = "black", position = position_stack(reverse = TRUE)) +
  geom_line(data = ch4, aes(x = i, y = y_left_scaled), inherit.aes = FALSE,
            linewidth = 0.9, color = line_color, alpha = 0.6) +
  geom_point(data = ch4, aes(x = i, y = y_left_scaled), inherit.aes = FALSE,
             size = 1.8, color = line_color, stroke = 0, alpha = 0.6) +
  scale_fill_manual(values = fill_colors,
                    breaks = c("Chinese", "English"),
                    labels = c("Chinese", "English"),
                    name = NULL) +
  guides(fill = guide_legend(reverse = FALSE)) +
  scale_x_continuous(breaks = seq_len(n_years), labels = all_years, expand = c(0, 0)) +
  scale_y_continuous(
    name = "Publications (count) ",
    expand = expansion(mult = c(0, 0.03)),
    sec.axis = sec_axis(~ (. - b) / a, name = expression(CH[4]~mole~fraction~"(ppb)"))
  ) +
  coord_cartesian(xlim = c(left_edge, right_edge), ylim = c(0, NA), clip = "off") +
  labs(x = "Year") +
  theme_classic(base_size = 18) +
  theme(
    legend.position = c(0.05, 0.95),
    legend.justification = c("left", "top"),
    legend.background = element_rect(color = "black", fill = "white", linewidth = 0.4),
    legend.text = element_text(size = 16),
    legend.key.height = unit(20, "pt"),
    legend.key.width  = unit(28, "pt"),
    legend.box = "vertical",
    panel.grid.major.y = element_blank(),
    panel.grid.minor.y = element_blank(),
    panel.grid.major.x = element_line(linetype = "dashed", linewidth = 0.3, colour = "gray75"),
    axis.text.x = element_text(angle = 45, hjust = 1, size = 14),
    axis.text.y.left  = element_text(size = 14),
    axis.title.y.left = element_text(size = 18),
    axis.text.y.right  = element_text(size = 14, colour = line_color),
    axis.title.y.right = element_text(size = 18, colour = line_color),
    axis.line.y.right  = element_line(color = line_color, linewidth = 0.5),
    axis.ticks.y.right = element_line(color = line_color, linewidth = 0.3),
    axis.line.y.left  = element_line(color = "black", linewidth = 0.5),
    axis.ticks.y.left = element_line(color = "black", linewidth = 0.3),
    axis.title.x = element_text(size = 18),
    axis.line.x  = element_line(color = "black", linewidth = 0.5),
    axis.ticks.x = element_line(color = "black", linewidth = 0.3),
    axis.ticks.length = unit(-3, "pt")
  )

# ===== 导出图到 fig1 =====
out_file <- file.path(out_dir, "Fig 1.jpg")
ggsave(out_file, p, width = 12, height = 6, dpi = 600, units = "in", device = "jpeg")
message("Done: ", out_file)

# ======= 末尾输出总计（Total / English / Chinese），并导出汇总 CSV =======
lang_totals <- df_count %>%
  group_by(language) %>%
  summarise(total = sum(count), .groups = "drop") %>%
  complete(language = factor(c("English","Chinese"), levels = c("English","Chinese")),
           fill = list(total = 0)) %>%
  arrange(language)

total_all <- sum(lang_totals$total)
eng_total <- lang_totals$total[lang_totals$language == "English"]
chn_total <- lang_totals$total[lang_totals$language == "Chinese"]

cat("\n================ Language Totals (", year_min, "–", year_max, ") ================\n", sep = "")
cat(sprintf("Total papers : %d\n", total_all))
cat(sprintf("English      : %d\n", eng_total))
cat(sprintf("Chinese      : %d\n", chn_total))
cat("===========================================================\n\n")

write.csv(
  data.frame(
    Metric = c("Total","English","Chinese"),
    Count  = c(total_all, eng_total, chn_total),
    Range  = paste0(year_min, "–", year_max)
  ),
  file.path(out_dir, "fig1_language_totals_summary.csv"),
  row.names = FALSE
)
message("Totals CSV: ", file.path(out_dir, "fig1_language_totals_summary.csv"))
# ================== 结束 ==================

