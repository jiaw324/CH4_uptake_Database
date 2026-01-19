# ============================================
# Packages
# ============================================
suppressPackageStartupMessages({
  library(ggplot2); library(sf); library(readxl); library(dplyr)
  library(rnaturalearth); library(rnaturalearthdata)
  library(grid); library(stringr); library(rlang); library(patchwork)
})

# ============================================
# Paths  (UPDATED: output to .../Fig2)
# ============================================
ROOT     <- "D:/Users/jiaweiChiang/Desktop/Supplymentary_code"
work_dir <- file.path(ROOT, "Fig2")   # ✅ 输出目录
out_dir  <- work_dir
if (!dir.exists(work_dir)) dir.create(work_dir, recursive = TRUE)
setwd(work_dir)

# ---- auto-detect input file: Studies and Fluxes*.xlsx ----
cand <- list.files(ROOT, pattern = "^Studies and Fluxes.*\\.xlsx$", full.names = TRUE, ignore.case = TRUE)
if (!length(cand)) stop("在目录中未找到输入文件：Studies and Fluxes*.xlsx ；目录=", ROOT)

pref <- cand[grepl("^Studies and Fluxes\\.xlsx$", basename(cand), ignore.case = TRUE)]
in_file <- if (length(pref)) pref[1] else cand[1]
stopifnot(file.exists(in_file))
message("✅ Using input: ", in_file)

# ============================================
# Read (auto sheet)
# ============================================
sheets_all <- readxl::excel_sheets(in_file)
sheet_use  <- if ("row" %in% sheets_all) "row" else if ("Sheet1" %in% sheets_all) "Sheet1" else sheets_all[1]
df <- read_excel(in_file, sheet = sheet_use, col_names = TRUE)

# ============================================
# Helpers (column finding)
# ============================================
cn_raw <- names(df)
cn_key <- str_to_lower(str_trim(gsub("[\\._\\s\\(\\)]+", "", cn_raw)))

find_col_exact <- function(cands){
  idx <- which(cn_key %in% cands)
  if (length(idx)) cn_raw[idx[1]] else NA_character_
}

find_cols_fuzzy <- function(patterns){
  hit <- Reduce(`|`, lapply(patterns, function(p) grepl(p, cn_key, perl = TRUE)))
  cn_raw[which(hit)]
}

# ============================================
# Base columns (UPDATED mapping: Biome)
# ============================================
# ✅ 这里改成优先找 Biome，同时兼容老的 Biomes
col_biome <- find_col_exact(c("biome","biomes","biome(s)"))

col_lat    <- find_col_exact(c("latitude","lat"))
col_lon    <- find_col_exact(c("longitude","lon","long"))

col_labo   <- find_col_exact(c("q01"))         # laboratory -> Q01
col_repeat <- find_col_exact(c("q02"))         # repeat     -> Q02
col_dq     <- find_col_exact(c("q03"))         # dataquality-> Q03
col_ctrl   <- find_col_exact(c("management"))  # control    -> management

needed <- c(col_biome, col_lat, col_lon, col_labo, col_repeat, col_ctrl, col_dq)
if (any(is.na(needed))) {
  miss <- c("Biome","Latitude","Longitude","Q01","Q02","management","Q03")[is.na(needed)]
  stop("缺少必要列：", paste(miss, collapse=", "),
       "。请检查表头是否与新版 Studies and Fluxes 一致。")
}

# ============================================
# Base filter（去实验室/重复/差质量；仅保留 control）
# ============================================
df_base <- df %>%
  mutate(
    lab  = trimws(as.character(.data[[col_labo]])),
    rep  = trimws(as.character(.data[[col_repeat]])),
    dq   = trimws(as.character(.data[[col_dq]])),
    ctrl = str_to_lower(trimws(as.character(.data[[col_ctrl]])))
  ) %>%
  filter(is.na(lab) | lab != "√") %>%
  filter(is.na(rep) | rep != "√") %>%
  filter(is.na(dq)  | dq  != "√") %>%
  filter(ctrl == "control")

# ============================================
# Palette & fixed legend order
# ============================================
custom_colors <- c(
  "Tundra"="#a7d3df","Boreal forest"="#4f9f4f","Temperate seasonal forest"="#8cbf70",
  "Temperate rainforest"="#e8c8a1","Tropical rainforest"="#c49b50","Tropical seasonal forest"="#2e7032",
  "Savanna"="#5fa6bc","Subtropical desert"="#f2c300","Temperate grassland"="#2e6f89",
  "Desert"="#fff2b2","Woodland"="#cd7a50","Shrubland"="#a34030",
  "Alpine"="#b27874","Agriculture"="#886b99","Bare"="#357266","Urban"="#6b575a","Others"="#b3b5ae"
)

legend_levels <- c(
  "Tundra","Boreal forest","Temperate seasonal forest",
  "Temperate rainforest","Tropical rainforest","Tropical seasonal forest",
  "Savanna","Subtropical desert","Temperate grassland",
  "Desert","Woodland","Shrubland",
  "Alpine","Agriculture","Bare","Urban","Others"
)

# ============================================
# Time-scale columns & flags（任意一列非 NA 即纳入）
# ============================================
month_cols <- find_cols_fuzzy(c(
  "^jan(uary)?$","^feb(ruary)?$","^mar(ch)?$","^apr(il)?$","^may$",
  "^jun(e)?$","^jul(y)?$","^aug(ust)?$","^sep(t)?(ember)?$",
  "^oct(ober)?$","^nov(ember)?$","^dec(ember)?$"
))
if (!length(month_cols)) month_cols <- find_cols_fuzzy(paste0("^m", 1:12, "$"))

season_cols <- unique(c(
  find_cols_fuzzy(c("^spring$","^summer$","^autumn$","^winter$")),
  find_cols_fuzzy(c("^fall$"))
))
annual_cols <- find_cols_fuzzy(c("^annual$","^annual.*$"))

has_any_non_na <- function(df_in, cols){
  if (!length(cols)) return(rep(FALSE, nrow(df_in)))
  sub <- df_in[, cols, drop=FALSE]
  sub[] <- lapply(sub, function(x) suppressWarnings(as.numeric(as.character(x))))
  apply(!is.na(sub), 1, any)
}

flag_month  <- has_any_non_na(df_base, month_cols)
flag_season <- has_any_non_na(df_base, season_cols)
flag_annual <- has_any_non_na(df_base, annual_cols)

prep_points <- function(dfin){
  dfin %>%
    filter(!is.na(.data[[col_lat]]), !is.na(.data[[col_lon]])) %>%
    mutate(Biome = factor(as.character(.data[[col_biome]]), levels = legend_levels)) %>%
    filter(!is.na(Biome))
}

dfm <- prep_points(df_base[flag_month, , drop=FALSE])
dfs <- prep_points(df_base[flag_season, , drop=FALSE])
dfa <- prep_points(df_base[flag_annual, , drop=FALSE])
stopifnot(nrow(dfm)>0, nrow(dfs)>0, nrow(dfa)>0)

# ============================================
# Make sf + push Agriculture to bottom (draw first)
# ============================================
push_biome_to_bottom <- function(sf_obj, bottom="Agriculture"){
  ord <- c(which(sf_obj$Biome==bottom), which(sf_obj$Biome!=bottom))
  sf_obj[ord, ]
}

sf_m <- st_as_sf(dfm, coords=c(col_lon, col_lat), crs=4326) |> push_biome_to_bottom("Agriculture")
sf_s <- st_as_sf(dfs, coords=c(col_lon, col_lat), crs=4326) |> push_biome_to_bottom("Agriculture")
sf_a <- st_as_sf(dfa, coords=c(col_lon, col_lat), crs=4326) |> push_biome_to_bottom("Agriculture")

# ============================================
# Basemap + Robinson bbox (方案B关键)
# ============================================
world <- rnaturalearth::ne_countries(scale="medium", returnclass="sf")
wb <- st_bbox(st_transform(world, "+proj=robin"))

ocean_color  <- "#C7E9F1"
land_color   <- "#DDDDDD"
border_color <- "#222222"
border_size  <- 0.5

make_map <- function(sf_pts, title_text, x_title=NULL, y_title=NULL){
  ggplot() +
    geom_sf(data=world, fill=land_color, color=border_color, linewidth=border_size) +
    geom_sf(data=sf_pts, aes(fill=Biome), shape=21, size=3.2, alpha=0.95, color="black", stroke=0.4) +
    coord_sf(crs="+proj=robin", expand=FALSE) +
    labs(title=title_text, x=x_title, y=y_title) +
    theme_minimal() +
    theme(
      plot.title       = element_text(size=18, face="bold", hjust=0.5, margin=margin(b=6)),
      panel.background = element_rect(fill=ocean_color, color=NA),
      panel.grid       = element_line(color="gray30", linewidth=0.2),
      axis.title       = element_text(size=16),
      axis.text        = element_text(size=12),
      legend.position  = "bottom",
      legend.text      = element_text(size=14),
      legend.box       = "horizontal",
      legend.key       = element_rect(fill="white", color=NA),
      legend.spacing.x = unit(0.6, "cm"),
      plot.margin      = margin(4,4,4,4)
    )
}

p_m <- make_map(sf_m, "Monthly",  x_title=NULL,        y_title="Latitude")
p_s <- make_map(sf_s, "Seasonal", x_title=NULL,        y_title="Latitude")
p_a <- make_map(sf_a, "Annual",   x_title="Longitude", y_title="Latitude")

# ============================================
# Fixed fill scale
# ============================================
fill_scale <- scale_fill_manual(
  values = custom_colors[legend_levels],
  limits = legend_levels,
  breaks = legend_levels,
  drop   = FALSE,
  name   = NULL
)

p_m <- p_m + fill_scale + theme(legend.position = "none")
p_s <- p_s + fill_scale + theme(legend.position = "none")

# ---- Annual panel: ghost points + lock bbox ----
legend_ghost <- st_as_sf(
  data.frame(Biome = factor(legend_levels, levels = legend_levels),
             lon = 0, lat = -90),
  coords = c("lon","lat"), crs = 4326
)

p_a <- p_a + fill_scale +
  geom_sf(
    data = legend_ghost, aes(fill = Biome),
    shape = 21, size = 3.2, color = "black", stroke = 0.4,
    alpha = 0, show.legend = TRUE, inherit.aes = FALSE
  ) +
  guides(
    fill = guide_legend(
      nrow = 3, byrow = TRUE,
      override.aes = list(shape = 21, size = 3.0, colour = "black", stroke = 0.35, alpha = 0.95)
    )
  ) +
  coord_sf(
    crs = "+proj=robin",
    xlim = c(wb["xmin"], wb["xmax"]),
    ylim = c(wb["ymin"], wb["ymax"]),
    expand = FALSE
  ) +
  theme(
    legend.position      = "bottom",
    legend.box           = "horizontal",
    legend.justification = "center",
    legend.text          = element_text(size = 12),
    legend.key.width     = unit(7,  "pt"),
    legend.key.height    = unit(7,  "pt"),
    legend.spacing.x     = unit(4,  "pt"),
    legend.box.margin    = margin(t = 4, r = 4, b = 4, l = 4)
  )

# ============================================
# Combine & Save
# ============================================
combined <- (p_m / p_s / p_a) +
  plot_layout(heights = c(1,1,1)) &
  theme(panel.spacing = unit(0.6, "lines"),
        plot.margin   = margin(4,4,4,4))

print(combined)

out_file <- file.path(out_dir, "Fig2.jpg")
ggsave(out_file, combined, width = 16, height = 18, dpi = 600)
message("✅ 已保存三联图：", out_file)

ggsave(file.path(out_dir, "Fig2 monthly.jpg"),  p_m, width = 16, height = 8, dpi = 600)
ggsave(file.path(out_dir, "Fig2 seasonal.jpg"), p_s, width = 16, height = 8, dpi = 600)
ggsave(file.path(out_dir, "Fig2 annual.jpg"),   p_a + theme(legend.position="none"), width = 16, height = 8, dpi = 600)
message("✅ 已保存单图（不含图例）：monthly / seasonal / annual")
