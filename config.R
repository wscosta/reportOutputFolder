# config.R ---------------------------------------------------------------
# Central configuration file for the MAgPIE Brazil report pipeline.
# Edit ONLY this file when switching scenarios or output folders.
# All pipeline scripts (01_ to 04_) source this file at startup.

# Package check -----------------------------------------------------------

required_packages <- c(
  "here", "reshape2", "ggplot2", "dplyr", "plyr", "data.table",
  "tidyverse", "officer", "rvg", "magrittr", "flextable", "glue",
  "magick", "pdftools", "withr", "magclass", "terra", "stringr",
  "leaflet", "RColorBrewer"
)

missing_packages <- required_packages[!sapply(required_packages, requireNamespace, quietly = TRUE)]

if (length(missing_packages) > 0) {
  stop(
    "The following required packages are not installed:\n",
    paste(" -", missing_packages, collapse = "\n"),
    "\n\nInstall them with:\n",
    "install.packages(c(", paste0('"', missing_packages, '"', collapse = ", "), "))"
  )
}

for (pkg in required_packages) {
  if (!require(pkg, character.only = TRUE)) {
    install.packages(pkg)
    library(pkg, character.only = TRUE)
  }
}

message("All required packages are installed.")

# Scenario settings -------------------------------------------------------

# Name of the output folder containing the MAgPIE run results
#   outputFolder <- "BRA_vSecdForest"
outputFolder <- "BRA_v2SecdForest"

# Where to read the MAgPIE output files from.
# "magpie"  — reads from magpie/output/{outputFolder}/  (default when running inside the MAgPIE repo)
# "project" — reads from reportOutputFolder/data/{outputFolder}/  (use when files were copied locally)
data_source <- "project"

# Path to the MAgPIE output folder for this scenario (resolved automatically)
magpie_output_path <- if (data_source == "magpie") {
  file.path(
    dirname(here::here()),  # magpie/ root (one level above reportOutputFolder/)
    "output",
    outputFolder
  )
} else if (data_source == "project") {
  here::here("output", outputFolder)
} else {
  stop("Invalid data_source in config.R. Must be \"magpie\" or \"project\".")
}


selected_scenario <- as.character(unique(
  readRDS(file.path(magpie_output_path, "report.rds"))$scenario
))

if (length(selected_scenario) != 1) {
  stop(
    "Expected exactly 1 scenario in report.rds, found ", length(selected_scenario), ": ",
    paste(selected_scenario, collapse = ", ")
  )
}


# Cluster map (PDF -> JPG conversion) -------------------------------------

# Name of the spamplot PDF inside the scenario output folder
cluster_pdf_file <- list.files(
  magpie_output_path,
  pattern = "^spamplot.*\\.pdf$",
  full.names = TRUE
)[1]



# Land map settings -------------------------------------------------------

land_years   <- c(1995, 2020)
land_regions <- c("Brazil", "World")

land_classes <- c("primforest", "secdforest", "past", "crop", "other", "urban", "forest")
land_titles  <- c("Primary Forest", "Secondary Forest", "Pastures and Rangelands",
                  "Cropland", "Other Land", "Urban", "Forest")

land_palettes <- list(
  colorRampPalette(brewer.pal(9, "Greens"))(100),     # primforest
  colorRampPalette(brewer.pal(9, "Greens"))(100),     # secdforest
  colorRampPalette(brewer.pal(9, "Purples"))(100),    # past
  colorRampPalette(brewer.pal(9, "Reds"))(100),       # crop
  colorRampPalette(brewer.pal(9, "RdPu"))(100),       # other
  colorRampPalette(brewer.pal(9, "Greys"))(100),      # urban
  colorRampPalette(brewer.pal(9, "Greens"))(100)      # forest
)


# Crop map settings -------------------------------------------------------

crop_years   <- c(2020)
crop_regions <- c("Brazil")


# Diff map settings -------------------------------------------------------

# Temporal diff: Primary Forest loss/gain periods (03_createDiffMaps.R)
forest_loss_periods <- list(
  c(year_start = 2000, year_end = 2020),
  c(year_start = 2000, year_end = 2050)
)
forest_loss_regions <- c("Brazil", "World")
