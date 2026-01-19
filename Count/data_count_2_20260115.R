# ===============================
# 📦 必要包（用 writexl 替代 openxlsx）
# ===============================
required_packages <- c("readxl","dplyr","stringr","writexl")
for (pkg in required_packages) {
  if (!requireNamespace(pkg, quietly = TRUE)) install.packages(pkg)
}
invisible(lapply(required_packages, library, character.only = TRUE))

# ===============================
# 📁 输入/输出路径
# ===============================
in_dir <- "D:/Users/jiaweiChiang/Desktop/Supplymentary_code"
in_file <- file.path(in_dir, "Studies and Fluxes.xlsx")
stopifnot(file.exists(in_file))

out_dir <- file.path(in_dir, "Count")
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)

# ✅ 建议输出新文件名，避免你打开旧文件导致覆盖写入异常
excel_out <- file.path(out_dir, "nonempty_counts_AFTER_FILTER_only_WRITEXL.xlsx")

# 是否输出筛选后的完整表（很大时可设 FALSE）
export_filtered_sheet <- TRUE

# ===============================
# 🔎 智能列名匹配
# ===============================
pick_col <- function(nms, candidates, must=FALSE, label="column") {
  nms_l <- tolower(nms)
  
  # 精确命中优先
  for (cand in candidates) {
    hit <- which(nms_l == tolower(cand))
    if (length(hit) == 1) return(nms[hit])
  }
  
  # 模糊匹配
  for (cand in candidates) {
    pat <- paste0("(^", cand, "$)|\\b", cand, "\\b")
    hit <- which(grepl(pat, nms_l, perl = TRUE))
    if (length(hit) > 0) return(nms[hit[1]])
  }
  
  if (must) {
    stop(sprintf("找不到【%s】列；表头示例：%s",
                 label, paste(head(nms, 50), collapse=", ")), call. = FALSE)
  }
  NA_character_
}

# ===============================
# ✔️ 判定函数
# ===============================
norm_txt <- function(x) stringr::str_squish(as.character(x))

is_checked <- function(x){
  y <- stringr::str_to_lower(norm_txt(x))
  pos <- c("√","✓","✔","yes","y","1","true","t","是","對","对","勾","✅","☑","✔️")
  !is.na(y) & (y %in% tolower(pos))
}

is_control_strict <- function(x){
  y <- stringr::str_to_lower(norm_txt(x))
  !is.na(y) & stringr::str_detect(y, "^control$")
}

is_nonempty <- function(v){
  if (is.numeric(v)) return(!is.na(v))
  s <- stringr::str_trim(as.character(v))
  !(is.na(s) | s == "" | s %in% c("NA","N/A","-"))
}

# ===============================
# 🧼 Excel 写入前强力清理（关键：彻底避免“修复弹窗”）
# ===============================
clean_excel_text <- function(x){
  # 去掉 XML 不允许的控制字符
  x <- gsub("[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F]", "", x)
  
  # 截断超长字符串（Excel 单元格上限 32767）
  # 用 32000 做安全值
  too_long <- nchar(x, type = "chars", allowNA = TRUE) > 32000
  x[!is.na(too_long) & too_long] <- substr(x[!is.na(too_long) & too_long], 1, 32000)
  
  x
}

clean_df_for_excel <- function(df0){
  df1 <- as.data.frame(df0, check.names = FALSE, stringsAsFactors = FALSE)
  
  # 列名清理 + 强制唯一（重复列名也会让 Excel 很不爽）
  names(df1) <- clean_excel_text(names(df1))
  names(df1) <- make.unique(names(df1), sep = "_dup_")
  
  for (j in seq_along(df1)) {
    # 任何日期/时间都转字符，避免底层类型写入冲突
    if (inherits(df1[[j]], "POSIXct") || inherits(df1[[j]], "POSIXt") || inherits(df1[[j]], "Date")) {
      df1[[j]] <- as.character(df1[[j]])
    }
    
    # list 列强制转字符（writexl 不接受 list）
    if (is.list(df1[[j]]) && !is.data.frame(df1[[j]])) {
      df1[[j]] <- sapply(df1[[j]], function(z) {
        if (length(z) == 0) return(NA_character_)
        paste0(z, collapse = "; ")
      })
    }
    
    # 字符列清理
    if (is.character(df1[[j]])) {
      df1[[j]] <- clean_excel_text(df1[[j]])
    }
  }
  
  df1
}

# ===============================
# 📖 读取 Studies and Fluxes
# ===============================
sheets <- readxl::excel_sheets(in_file)
sheet_to_read <- if ("Sheet1" %in% sheets) "Sheet1" else sheets[1]
df <- readxl::read_excel(in_file, sheet = sheet_to_read)

cols <- names(df)

# —— 必需列：Q01~Q04 + management
q01_col <- pick_col(cols, c("q01"), TRUE, "Q01")
q02_col <- pick_col(cols, c("q02"), TRUE, "Q02")
q03_col <- pick_col(cols, c("q03"), TRUE, "Q03")
q04_col <- pick_col(cols, c("q04"), TRUE, "Q04")
mgmt_col <- pick_col(cols, c("management"), TRUE, "management")

# —— Daily 标记：Q07（可缺省）
q07_col <- pick_col(cols, c("q07"), must = FALSE, label = "Q07 (daily)")

# —— sites 计数：Latitude / Longitude（必须）
lat_col <- pick_col(cols, c("latitude","lat"), must = TRUE, label = "Latitude")
lon_col <- pick_col(cols, c("longitude","lon","long","lng"), must = TRUE, label = "Longitude")

# ===============================
# 🔧 过滤（你的规则）
#   - Q01/Q02/Q03/Q04 任意勾选 -> 排除
#   - management == "control" -> 保留
# ===============================
df_filt <- df %>%
  dplyr::filter(!is_checked(.data[[q01_col]])) %>%
  dplyr::filter(!is_checked(.data[[q02_col]])) %>%
  dplyr::filter(!is_checked(.data[[q03_col]])) %>%
  dplyr::filter(!is_checked(.data[[q04_col]])) %>%
  dplyr::filter(is_control_strict(.data[[mgmt_col]]))

total_rows <- nrow(df_filt)

# ===============================
# ✅ Site 计数：不重复 Latitude+Longitude
# ===============================
site_round_digits <- 6

count_unique_sites_latlon <- function(data){
  if (nrow(data) == 0) return(0L)
  
  lat <- suppressWarnings(as.numeric(data[[lat_col]]))
  lon <- suppressWarnings(as.numeric(data[[lon_col]]))
  
  ok <- !is.na(lat) & !is.na(lon)
  if (!any(ok)) return(0L)
  
  lat2 <- round(lat[ok], site_round_digits)
  lon2 <- round(lon[ok], site_round_digits)
  
  key <- paste0(lat2, "_", lon2)
  length(unique(key))
}

# ===============================
# 📅 识别 Monthly / Seasonal / Annual 列
# ===============================
pick_by_prefix <- function(data, prefixes){
  nms_l <- tolower(names(data))
  hits <- c()
  for (p in prefixes) {
    pat <- paste0("^", p)  # ✅ 兼容 Jan_1 / m1_1 / m1_conv
    hits <- c(hits, names(data)[grepl(pat, nms_l, perl = TRUE)])
  }
  unique(hits)
}

# Monthly：Jan–Dec + m1–m12
month_cols <- pick_by_prefix(df_filt, c(
  "jan","january","feb","february","mar","march","apr","april","may",
  "jun","june","jul","july","aug","august","sep","sept","september",
  "oct","october","nov","november","dec","december",
  "m1","m2","m3","m4","m5","m6","m7","m8","m9","m10","m11","m12"
))

# Seasonal
season_cols <- pick_by_prefix(df_filt, c(
  "spring","spr","summer","sum","autumn","fall","aut","winter","win"
))

# Annual（取最像的一个）
annual_col <- NA_character_
nms_f <- tolower(names(df_filt))
annual_candidates <- c("annual","annual_mean","annualavg","annual_flux",
                       "annual_rate","annual_annual","sea_annual",
                       "year","yearly","year_mean")
for (cand in annual_candidates) {
  eq_hit <- which(nms_f == cand)
  if (length(eq_hit) == 1) { annual_col <- names(df_filt)[eq_hit]; break }
  if (is.na(annual_col)) {
    pr_hit <- which(grepl(paste0("^", cand), nms_f, perl = TRUE))
    if (length(pr_hit) > 0) { annual_col <- names(df_filt)[pr_hit[1]] }
  }
}

# ===============================
# ✅ 子集：某类列里“任意一列非空”就选该行
# ===============================
subset_any_nonempty_rows <- function(data, cols_vec){
  if (length(cols_vec) == 0 || nrow(data) == 0) return(data[0, , drop=FALSE])
  m <- sapply(cols_vec, function(cn) is_nonempty(data[[cn]]))
  if (is.null(dim(m))) m <- matrix(m, ncol = 1)
  data[rowSums(m) > 0, , drop = FALSE]
}

monthly_rows <- subset_any_nonempty_rows(df_filt, month_cols)
season_rows  <- subset_any_nonempty_rows(df_filt, season_cols)

annual_rows <- df_filt[0, , drop=FALSE]
if (!is.na(annual_col) && annual_col %in% names(df_filt)) {
  annual_rows <- df_filt[is_nonempty(df_filt[[annual_col]]), , drop=FALSE]
}

# ===============================
# ✅ Daily：Q07 勾选行
# ===============================
daily_rows <- df_filt[0, , drop=FALSE]
if (!is.na(q07_col) && q07_col %in% names(df_filt)) {
  daily_rows <- df_filt %>% dplyr::filter(is_checked(.data[[q07_col]]))
}

# ===============================
# 📊 统计输出
# ===============================
n_daily   <- nrow(daily_rows)
n_month   <- nrow(monthly_rows)
n_season  <- nrow(season_rows)
n_annual  <- nrow(annual_rows)

sites_daily  <- count_unique_sites_latlon(daily_rows)
sites_month  <- count_unique_sites_latlon(monthly_rows)
sites_season <- count_unique_sites_latlon(season_rows)
sites_annual <- count_unique_sites_latlon(annual_rows)

coverage_pct_safe <- function(x, total){
  if (total > 0) round(100 * x / total, 2) else 0
}

stats_fill <- dplyr::tibble(
  Category = c(
    "Daily (Q07 checked)",
    "Monthly (Jan–Dec / m1–m12 any)",
    "Seasonal (Spring–Winter any)",
    "Annual"
  ),
  Rows_with_values = c(n_daily, n_month, n_season, n_annual),
  Unique_sites_LatLon = c(sites_daily, sites_month, sites_season, sites_annual),
  Total_rows = total_rows,
  Coverage_pct = c(
    coverage_pct_safe(n_daily, total_rows),
    coverage_pct_safe(n_month, total_rows),
    coverage_pct_safe(n_season, total_rows),
    coverage_pct_safe(n_annual, total_rows)
  ),
  Columns_used = c(
    if (!is.na(q07_col)) q07_col else "(Q07 not found)",
    if (length(month_cols) > 0) paste(month_cols, collapse=", ") else "(month columns not found)",
    if (length(season_cols) > 0) paste(season_cols, collapse=", ") else "(season columns not found)",
    if (!is.na(annual_col)) annual_col else "(annual column not found)"
  )
)

# ===============================
# 💾 写入 Excel（writexl 最稳）
# ===============================
stats_out <- clean_df_for_excel(stats_fill)

out_list <- list(
  Counts_Rows_and_Sites = stats_out
)

if (isTRUE(export_filtered_sheet)) {
  out_list$Filtered_Data <- clean_df_for_excel(df_filt)
}

writexl::write_xlsx(out_list, path = excel_out)

message("✅ 已导出（writexl 稳定版）：", excel_out)
message("📌 读取文件：", in_file, " (sheet=", sheet_to_read, ")")
message("📌 筛选后行数 total_rows = ", total_rows)
message("📌 Latitude列 = ", lat_col, " | Longitude列 = ", lon_col)
message("📌 Q07列 = ", ifelse(is.na(q07_col), "(not found)", q07_col))
