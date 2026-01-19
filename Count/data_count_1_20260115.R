# ===============================
# 📦 必要包
# ===============================
required_packages <- c("readxl","dplyr","stringr","readr","openxlsx")
for (pkg in required_packages) if (!requireNamespace(pkg, quietly = TRUE)) install.packages(pkg)
invisible(lapply(required_packages, library, character.only = TRUE))

# ===============================
# 📁 路径（按你的最新要求）
# ===============================
in_dir1  <- "D:/Users/jiaweiChiang/Desktop/Supplymentary_code"
file_flux1410 <- file.path(in_dir1, "CH4 uptake data_CH4 FLUX_1_1410.xlsx")   # 附件1：只检查存在
file_studies  <- file.path(in_dir1, "Studies and Fluxes.xlsx")               # 附件2：主数据来源

stopifnot(file.exists(file_flux1410))
stopifnot(file.exists(file_studies))

out_dir <- file.path(in_dir1, "Count")
if (!dir.exists(out_dir)) dir.create(out_dir, recursive = TRUE)

# ===============================
# 🔎 智能列名匹配（更稳）
# ===============================
pick_col <- function(nms, candidates, must=TRUE, label="column") {
  nms_l <- tolower(nms)
  # 1) 完全匹配优先
  for (cand in candidates) {
    hit <- which(nms_l == tolower(cand))
    if (length(hit) == 1) return(nms[hit])
  }
  # 2) 模糊匹配兜底
  for (cand in candidates) {
    pat <- paste0("^(?:", cand, ")$|\\b", cand, "\\b")
    hit <- which(grepl(pat, nms_l, perl = TRUE))
    if (length(hit) > 0) return(nms[hit[1]])
  }
  if (must) stop(sprintf("找不到【%s】列。表头示例：%s", label, paste(head(nms, 30), collapse=", ")), call. = FALSE)
  NA_character_
}

# ===============================
# 📖 读取 Studies and Fluxes.xlsx（优先 sheet='row'）
# ===============================
sheets <- readxl::excel_sheets(file_studies)
sheet_to_read <- if ("row" %in% sheets) "row" else sheets[1]
df <- readxl::read_excel(file_studies, sheet = sheet_to_read)

cols <- names(df)

# --- 核心列 ---
cont_col <- pick_col(cols, c("Continent","continent"), TRUE, "Continent")
ctry_col <- pick_col(cols, c("Country","country","nation"), TRUE, "Country")
clim_col <- pick_col(cols, c("Climate","climate","climate_zone","koppen"), TRUE, "Climate")
eco_col  <- pick_col(cols, c("Ecosystem","ecosystem"), TRUE, "Ecosystem")
biom_col <- pick_col(cols, c("Biomes","biomes","biome","biome_name"), TRUE, "Biomes")

# --- 新列名映射（按你要求） ---
q01_col <- pick_col(cols, c("Q01","q01"), TRUE, "Q01")
q02_col <- pick_col(cols, c("Q02","q02"), TRUE, "Q02")
q03_col <- pick_col(cols, c("Q03","q03"), TRUE, "Q03")
q04_col <- pick_col(cols, c("Q04","q04"), TRUE, "Q04")
mgmt_col <- pick_col(cols, c("management","Management","control","Control"), TRUE, "management")

# ===============================
# ✔️ 判定函数
# ===============================
norm_txt <- function(x) { stringr::str_squish(as.character(x)) }

is_checked <- function(x){
  y <- stringr::str_to_lower(norm_txt(x))
  y %in% c("√","✓","✔","yes","y","1","true","t","是","對","对","勾","checked")
}

is_control_strict <- function(x){
  y <- stringr::str_to_lower(norm_txt(x))
  stringr::str_detect(y, "^control(?:\\s*group)?$")
}

# ===============================
# 🌿 仅对 Ecosystem 做 Wetland→Others（其余不变）
# ===============================
target_ecos <- c("Agriculture","Bare","Desert","Forest","Grassland",
                 "Others","Rainforest","Shrub","Tundra","Urban","Woodland","Savanna")

map_ecos <- function(x) {
  xl <- stringr::str_to_lower(stringr::str_trim(as.character(x)))
  xl[ xl %in% c("wetland","marsh","swamp","bog","fen","peatland") ] <- "others"
  tl <- stringr::str_to_lower(target_ecos)
  idx <- match(xl, tl)
  ifelse(is.na(idx), NA_character_, target_ecos[idx])
}

# ===============================
# 🔧 筛选（去掉 Q01/Q02/Q03/Q04 = √；仅保留 management="control"）
# ===============================
df_filt <- df %>%
  dplyr::mutate(
    Q01v = .data[[q01_col]],
    Q02v = .data[[q02_col]],
    Q03v = .data[[q03_col]],
    Q04v = .data[[q04_col]],
    MGMT = .data[[mgmt_col]]
  ) %>%
  dplyr::filter(!is_checked(Q01v)) %>%
  dplyr::filter(!is_checked(Q02v)) %>%
  dplyr::filter(!is_checked(Q03v)) %>%
  dplyr::filter(!is_checked(Q04v)) %>%
  dplyr::filter(is_control_strict(MGMT)) %>%
  dplyr::transmute(
    Continent = norm_txt(.data[[cont_col]]),
    Country   = norm_txt(.data[[ctry_col]]),
    Climate   = norm_txt(.data[[clim_col]]),
    Ecosystem = map_ecos(.data[[eco_col]]),   # Wetland → Others（只改这里）
    Biomes    = norm_txt(.data[[biom_col]])
  )

# ===============================
# 📊 统计各列：数量与百分比（分母=该列非空有效值）
# ===============================
tally_one <- function(df, colname) {
  v <- df[[colname]]
  v <- if (is.character(v)) ifelse(v %in% c("", "NA", "N/A", "-"), NA_character_, v) else v
  dplyr::tibble(Item = v) |>
    dplyr::filter(!is.na(Item)) |>
    dplyr::count(Item, name = "count") |>
    dplyr::mutate(percent = round(100 * count / sum(count), 2)) |>
    dplyr::arrange(dplyr::desc(count))
}

tbl_continent <- tally_one(df_filt, "Continent")
tbl_country   <- tally_one(df_filt, "Country")
tbl_climate   <- tally_one(df_filt, "Climate")
tbl_ecosys    <- tally_one(df_filt, "Ecosystem")
tbl_biomes    <- tally_one(df_filt, "Biomes")

# ===============================
# 💾 导出 Excel（多 sheet）
# ===============================
excel_out <- file.path(out_dir, "counts_after_filter(Q01_Q02_Q03_Q04_management_control).xlsx")

wb <- openxlsx::createWorkbook()

add_sheet <- function(name, dat){
  openxlsx::addWorksheet(wb, name)
  openxlsx::writeData(wb, name, dat)
  if (nrow(dat) > 0) {
    openxlsx::addFilter(wb, name, row=1, cols=1:ncol(dat))
    openxlsx::setColWidths(wb, name, cols=1:ncol(dat), widths="auto")
  }
}

add_sheet("Continent", tbl_continent)
add_sheet("Country",   tbl_country)
add_sheet("Climate",   tbl_climate)
add_sheet("Ecosystem", tbl_ecosys)   # 已 Wetland→Others
add_sheet("Biomes",    tbl_biomes)   # 原样

openxlsx::saveWorkbook(wb, excel_out, overwrite = TRUE)

message("✅ 已导出：", excel_out)
message("📌 读取文件：", file_studies, " | sheet=", sheet_to_read)
message("📌 附件1存在性检查通过：", file_flux1410)
message("📌 筛选后行数：", nrow(df_filt))