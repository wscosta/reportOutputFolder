# 03_createDiffMaps.R ----------------------------------------------------
# Generates temporal Primary Forest loss/gain maps within the same scenario.
# Computes raster(year_end) - raster(year_start) for the primforest class.
#
#   Periods: 2000->2020 (historical), 2000->2050 (future projection)
#
# Input:  magpie/output/{outputFolder}/cell.land_0.5.mz
#         reportOutputFolder/data/shapefiles/
#
# Output: reportOutputFolder/maps/{outputFolder}/diff/deforestation/{region}/*.png
#
# Run order: THIRD — after 02_createLandMaps.R.


# Config ------------------------------------------------------------------

source(here::here("config.R"))

# Libraries ---------------------------------------------------------------

library(here)
library(magclass)
library(terra)
library(stringr)
library(tidyverse)
library(leaflet)

# Functions ---------------------------------------------------------------

# Load Brazil shapefiles (states, biomes, MATOPIBA)
load_shapefiles <- function() {
  list(
    states   = vect(here::here("data", "shapefiles", "br_states.shp")),
    biomes   = vect(here::here("data", "shapefiles", "br_biomes.shp")),
    matopiba = vect(here::here("data", "shapefiles", "matopiba.shp"))
  )
}

# Plot temporal Primary Forest loss/gain within the same scenario
# Computes: raster(year_end) - raster(year_start)
# Positive values = gain, negative values = loss
plot_forest_loss <- function(raster_object, year_start, year_end,
                             region = "Brazil", unit = "Mha", output_folder, shp) {
  class         <- "primforest"
  class_title   <- "Primary Forest"
  pattern_start <- paste0("y", year_start, "..", class)
  pattern_end   <- paste0("y", year_end,   "..", class)

  if (!pattern_start %in% names(raster_object)) {
    stop("Layer not found in raster: ", pattern_start)
  }
  if (!pattern_end %in% names(raster_object)) {
    stop("Layer not found in raster: ", pattern_end)
  }

  rDiff <- raster_object[[pattern_end]] - raster_object[[pattern_start]]

  ceil         <- 0.30914
  pal          <- leaflet::colorNumeric(palette = "RdBu", domain = c(-ceil, ceil), reverse = FALSE)
  b            <- seq(-ceil, ceil, 0.001)
  period_label <- paste0(year_start, "-", year_end)
  title        <- paste0(class_title, " Loss/Gain (", period_label, ")")

  if (region == "Brazil") {
    filetitle <- paste0(class_title, " Loss ", period_label, " Brazil.png")
    rDiff     <- terra::crop(rDiff, ext(-73.9872354804, -33.7683777809, -34.7299934555, 5.24448639569))
  } else {
    filetitle <- paste0(class_title, " Loss ", period_label, " World.png")
  }

  output_dir <- here::here("maps", output_folder, "diff", "deforestation", paste0("y", period_label), region)
  if (!dir.exists(output_dir)) {
    dir.create(output_dir, recursive = TRUE)
    cat("Folder created:", output_dir, "\n")
  }

  png(filename = file.path(output_dir, filetitle))
  terra::plot(rDiff,
              range      = c(-ceil, ceil),
              col        = pal(b),
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
  cat("Deforestation map created:", filetitle, "| region:", region, "\n")
}

# Main --------------------------------------------------------------------

shp <- load_shapefiles()

rLand_X <- as.SpatRaster(read.magpie(file.path(magpie_output_path, "cell.land_0.5.mz")))


# Primary Forest loss/gain ---------------------------------------

for (period in forest_loss_periods) {
  for (region in forest_loss_regions) {
    plot_forest_loss(
      raster_object = rLand_X,
      year_start    = period["year_start"],
      year_end      = period["year_end"],
      region        = region,
      unit          = "Mha",
      output_folder = outputFolder,
      shp           = shp
    )
  }
}

cat("All deforestation maps done.\n")
