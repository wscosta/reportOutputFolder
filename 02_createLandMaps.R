# 02_createLandMaps.R ----------------------------------------------------
# Generates land cover and crop share maps from MAgPIE raster outputs.
# Produces PNG images saved under maps/ directory,
# which are later embedded in the report (04_buildReport.R).
#
# Input:  magpie/output/{outputFolder}/cell.land_0.5.mz
#         magpie/output/{outputFolder}/cell.croparea_0.5_share.mz
#         reportOutputFolder/data/csv/  (IBGE historical crop maps)
#         reportOutputFolder/data/shapefiles/  (Brazil states, biomes, MATOPIBA)
#
# Output: reportOutputFolder/maps/{outputFolder}/land/y{year}/{region}/*.png
#         reportOutputFolder/maps/{outputFolder}/y{year}/*.png
#         reportOutputFolder/maps/IBGE/y{year}/*.png
#
# Run order: SECOND — after 01_convertClusterPDF.R.


# Config ------------------------------------------------------------------

source(here::here("config.R"))


# Libraries ---------------------------------------------------------------

library(here)
library(magclass)
library(terra)
library(stringr)
library(tidyverse)
library(leaflet)
library(RColorBrewer)


# Functions ---------------------------------------------------------------

# Load Brazil shapefiles (states, biomes, MATOPIBA)
load_shapefiles <- function() {
  list(
    states   = vect(here::here("data", "shapefiles", "br_states.shp")),
    biomes   = vect(here::here("data", "shapefiles", "br_biomes.shp")),
    matopiba = vect(here::here("data", "shapefiles", "matopiba.shp"))
  )
}

# Plot a single land use class for a given year and region
plot_land <- function(raster_object, class, class_title, year, region, unit, palette, shp) {
  ceil          <- 0.30914
  terra_pattern <- paste0("y", year, "..", class)
  title         <- paste0(class_title, " (", year, ") ")

  if (region == "Brazil") {
    filetitle     <- paste0(class_title, " ", year, " Brazil.png")
    raster_object <- terra::crop(raster_object,
                                 ext(-73.9872354804, -33.7683777809, -34.7299934555, 5.24448639569))
  } else {
    filetitle <- paste0(class_title, " ", year, " World.png")
  }

  output_dir <- here::here("maps", outputFolder, "land", paste0("y", year), region)
  if (!dir.exists(output_dir)) {
    dir.create(output_dir, recursive = TRUE)
    cat("Folder created:", output_dir, "\n")
  }

  png(filename = file.path(output_dir, filetitle))
  terra::plot(raster_object[[terra_pattern]],
              col        = palette,
              range      = c(0, ceil),
              plg        = list(title = unit, title.cex = 1, cex = 1),
              main       = title,
              cex.main   = 1,
              background = "lightblue")

  if (region == "Brazil") {
    terra::plot(shp$states,   border = "lightgray", lwd = 0.1, add = TRUE)
    terra::plot(shp$biomes,   border = "black",     lwd = 1,   add = TRUE)
    terra::plot(shp$matopiba, border = "cyan",      lwd = 0.1, add = TRUE)
  }
  dev.off()
}

# Plot a single crop share map for a given year (Brazil only)
plot_crop <- function(raster_object, class, class_title, year, unit, shp) {
  ceil          <- 1
  terra_pattern <- paste0("y", year, "..", class)
  title         <- paste0(class_title, " (", year, ")")
  filetitle     <- paste0(class_title, year, ".png")

  pal <- leaflet::colorNumeric(palette = "Reds", domain = c(0, 1), reverse = FALSE)
  b   <- seq(0, 1, 0.001)

  output_dir <- here::here("maps", outputFolder, "crop", paste0("y", year))
  if (!dir.exists(output_dir)) {
    dir.create(output_dir, recursive = TRUE)
    cat("Folder created:", output_dir, "\n")
  }

  png(filename = file.path(output_dir, filetitle))
  terra::plot(raster_object[[terra_pattern]],
              col        = pal(b),
              range      = c(0, ceil),
              plg        = list(title = unit, title.cex = 1, cex = 1),
              main       = title,
              cex.main   = 1,
              background = "lightblue")

  terra::plot(shp$states,   border = "lightgray", lwd = 0.1, add = TRUE)
  terra::plot(shp$biomes,   border = "black",     lwd = 1,   add = TRUE)
  terra::plot(shp$matopiba, border = "cyan",      lwd = 0.1, add = TRUE)
  dev.off()
}

# Plot a historical crop map from a SpatRaster object (IBGE source)
plot_hist_crop <- function(raster_object, class, class_title, year, unit, shp,
                           source_label = "IBGE") {
  terra_pattern <- paste0("y", year, "..", class)
  title         <- paste0(class_title, " (", year, ") ", source_label)
  filetitle     <- paste0(class_title, " ", source_label, " (", year, ").png")

  ceil <- 0.30914
  pal  <- leaflet::colorNumeric(palette = "Reds", domain = c(0, 1), reverse = FALSE)
  b    <- seq(0, 1, 0.001)

  output_dir <- here::here("maps", "IBGE", paste0("y", year))
  if (!dir.exists(output_dir)) {
    dir.create(output_dir, recursive = TRUE)
    cat("Folder created:", output_dir, "\n")
  }

  png(filename = file.path(output_dir, filetitle))
  terra::plot(raster_object[[terra_pattern]],
              col        = pal(b),
              range      = c(0, ceil),
              plg        = list(title = unit, title.cex = 1, cex = 1),
              main       = title,
              cex.main   = 1,
              background = "white")

  terra::plot(shp$states,   border = "lightgray", lwd = 0.1, add = TRUE)
  terra::plot(shp$biomes,   border = "black",     lwd = 1,   add = TRUE)
  terra::plot(shp$matopiba, border = "cyan",      lwd = 0.1, add = TRUE)
  dev.off()
}

# Helper: aggregate rainfed + irrigated layers for a crop into a single layer
aggregate_crop_layers <- function(rCropsBrazil, crop_rainfed, crop_irrigated, crop_output) {
  desired <- c(crop_rainfed, crop_irrigated)
  pattern <- paste0(".*\\.", desired, collapse = "|")
  layers  <- grep(pattern, names(rCropsBrazil), value = TRUE)
  rCrop   <- terra::subset(rCropsBrazil, layers)
  yrs     <- unique(sub(".*(y[0-9]{4}).*", "\\1", names(rCrop)))

  for (yr in yrs) {
    rf_layer  <- paste0(yr, "..", crop_rainfed)
    irr_layer <- paste0(yr, "..", crop_irrigated)
    if (rf_layer %in% names(rCrop) && irr_layer %in% names(rCrop)) {
      rCrop[[paste0(yr, "..", crop_output)]] <- rCrop[[rf_layer]] + rCrop[[irr_layer]]
    }
  }
  return(rCrop)
}

# Helper: load a CSV crop file and convert to a cropped Brazil SpatRaster
load_csv_crop <- function(csv_path, value_col = "value", ceil = 0.30914) {
  df        <- read.csv2(csv_path)
  df[[value_col]] <- as.numeric(df[[value_col]])

  mp <- magclass::as.magpie(df, spatial = "x.y.iso", tidy = TRUE)
  magclass::getItems(mp, 1, raw = TRUE) <- df[["x.y.iso"]]

  r        <- as.SpatRaster(mp)
  r_brazil <- terra::crop(r, ext(-73.9872354804, -33.7683777809, -34.7299934555, 5.24448639569))
  return(r_brazil)
}


# Main --------------------------------------------------------------------

shp <- load_shapefiles()


# Land cover maps (MAgPIE output) -----------------------------------------

land   <- read.magpie(file.path(magpie_output_path, "cell.land_0.5.mz"))
rLand  <- as.SpatRaster(land)

for (year in land_years) {

  rLand_year <- subset(rLand, grep(paste0("^y", year), names(rLand)))

  # Derive combined forest layer
  rLand_year[[paste0("y", year, "..forest")]] <-
    rLand_year[[paste0("y", year, "..primforest")]] +
    rLand_year[[paste0("y", year, "..secdforest")]]

  for (i in seq_along(land_classes)) {
    for (region in land_regions) {
      plot_land(
        raster_object = rLand_year,
        class         = land_classes[i],
        class_title   = land_titles[i],
        year          = year,
        region        = region,
        unit          = "Mha",
        palette       = land_palettes[[i]],
        shp           = shp
      )
      cat("Land map created:", land_titles[i], "(", year, ") -", region, "\n")
    }
  }
}


# Crop share maps (MAgPIE output) -----------------------------------------

crops        <- read.magpie(file.path(magpie_output_path, "cell.croparea_0.5_share.mz"))
#crops        <- crops * land[,getItems(crops, dim = 2), "crop", drop = FALSE]
rCrops       <- as.SpatRaster(crops)
#names(rCrops) <- sub("\\.crop$", "", names(rCrops))
rCropsBrazil <- terra::crop(rCrops,
                            ext(-73.9872354804, -33.7683777809, -34.7299934555, 5.24448639569))

rSoybean   <- aggregate_crop_layers(rCropsBrazil, "soybean.rainfed",   "soybean.irrigated",   "soybean")
rMaize     <- aggregate_crop_layers(rCropsBrazil, "maiz.rainfed",      "maiz.irrigated",      "maize")
rSugarcane <- aggregate_crop_layers(rCropsBrazil, "sugr_cane.rainfed", "sugr_cane.irrigated", "sugr_cane")

for (year in crop_years) {
  plot_crop(rSoybean,   "soybean",   "MAgPIE - Soybean",   year, "Cell share", shp)
  plot_crop(rMaize,     "maize",     "MAgPIE - Maize",     year, "Cell share", shp)
  plot_crop(rSugarcane, "sugr_cane", "MAgPIE - Sugarcane", year, "Cell share", shp)
  cat("Crop maps created for year:", year, "\n")
}


# Historical crop maps (IBGE source, from CSV) ----------------------------
# These files are static and do not change between scenarios.
# Each map is only generated if the output file does not yet exist.

ibge_crops <- list(
  list(file = "soybean_planted_area_2020.csv", class = "soybean",   title = "Soybean",   year = 2020),
  list(file = "maiz_planted_area_2020.csv",    class = "maiz",      title = "Maize",     year = 2020),
  list(file = "sugr_cane_planted_area_2020.csv", class = "sugr_cane", title = "Sugarcane", year = 2020)
)

ibge_cropland <- list(
  list(file = "cropland_1995_all.csv", class = "cropland", title = "Cropland (ALL)", year = 1995),
  list(file = "cropland_2000_all.csv", class = "cropland", title = "Cropland (ALL)", year = 2000),
  list(file = "cropland_2020_all.csv", class = "cropland", title = "Cropland (ALL)", year = 2020),
  list(file = "cropland_2024_all.csv", class = "cropland", title = "Cropland (ALL)", year = 2024)
)

for (entry in ibge_crops) {
  output_file <- here::here(
    "maps", "IBGE", paste0("y", entry$year),
    paste0(entry$title, " IBGE (", entry$year, ").png")
  )
  if (file.exists(output_file)) next
  r <- load_csv_crop(here::here("data", "csv", entry$file))
  plot_hist_crop(r, entry$class, entry$title, entry$year, "Mha", shp)
  cat("IBGE crop map created:", entry$title, "(", entry$year, ")\n")
}

for (entry in ibge_cropland) {
  output_file <- here::here(
    "maps", "IBGE", paste0("y", entry$year),
    paste0(entry$title, " IBGE (", entry$year, ").png")
  )
  if (file.exists(output_file)) next
  df       <- read.csv2(here::here("data", "csv", entry$file))
  df$value <- as.numeric(df$value)
  df       <- df %>% filter(kcr == "cropland") %>% mutate(value = value / 1e6)
  df$value <- pmin(df$value, 0.30914)
  
  mp <- magclass::as.magpie(df, spatial = "x.y.iso", tidy = TRUE)
  magclass::getItems(mp, 1, raw = TRUE) <- df[["x.y.iso"]]
  
  r        <- as.SpatRaster(mp)
  r_brazil <- terra::crop(r, ext(-73.9872354804, -33.7683777809, -34.7299934555, 5.24448639569))
  plot_hist_crop(r_brazil, entry$class, entry$title, entry$year, "Mha", shp)
  cat("IBGE cropland map created:", entry$title, "(", entry$year, ")\n")
}

# Check if all IBGE maps were skipped (all already existed)
all_ibge <- c(ibge_crops, ibge_cropland)
all_exist <- all(sapply(all_ibge, function(entry) {
  file.exists(here::here(
    "maps", "IBGE", paste0("y", entry$year),
    paste0(entry$title, " IBGE (", entry$year, ").png")
  ))
}))
if (all_exist) cat("All IBGE maps already exist, skipping.\n")