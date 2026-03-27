# 04_buildReport.R -------------------------------------------------------
# Main report script. Loads data, produces all charts and tables, and
# assembles the final .docx and .pdf report files.
#
# Input:  data/csv/histdatabrazil.csv
#         data/{outputFolder}/report.rds
#         maps/{outputFolder}/land/  (map JPGs from 02_createLandMaps.R)
#         maps/{outputFolder}/crop/         (crop PNGs from 02_createLandMaps.R)
#         maps/IBGE/                   (historical crop PNGs from 02_createLandMaps.R)
#         magpie/output/{outputFolder}/<spamplot>.jpg  (from 01_convertClusterPDF.R)
#
# Output: reportOutputFolder/report/comparisonsHistdata_{selected_scenario}.docx
#         reportOutputFolder/report/comparisonsHistdata_{selected_scenario}.pdf
#
# Run order: FOURTH — after all previous scripts have been run.

# Config ------------------------------------------------------------------

source(here::here("config.R"))

# Libraries ---------------------------------------------------------------

library(here)
library(reshape2)
library(ggplot2)
library(dplyr)
library(plyr)
library(data.table)
library(tidyverse)
library(officer)
library(rvg)
library(magrittr)
library(flextable)
library(glue)
library(magick)


# Helpers ------------------------------------------------------

source(here::here("R", "plot_functions.R"))
source(here::here("R", "table_functions.R"))


# Helper: prepare data for comparison plots -------------------------------

prepare_plot_data <- function(data_magpie, data_hist, product_type,
                              plot_type = "area", moisture_content = 0,
                              product_levels) {
  filtered_magpie <- data_magpie %>%
    dplyr::filter(type == !!product_type, year <= 2020)

  filtered_hist <- data_hist %>%
    dplyr::filter(type == !!product_type)

  if (plot_type == "production") {
    filtered_magpie <- filtered_magpie %>%
      dplyr::mutate(value = value / (1 - moisture_content),
                    unit  = "million tonnes")
  }

  dplyr::bind_rows(filtered_magpie, filtered_hist) %>%
    dplyr::filter(source %in% product_levels) %>%
    dplyr::mutate(source = factor(source, levels = product_levels))
}


# Helper: build table + flextable for a given variable --------------------

build_table <- function(data, reference_source) {
  tbl <- data %>%
    dplyr::select(-type, -unit) %>%
    dplyr::mutate(source = dplyr::recode(
      source,
      !!!setNames("MAgPIE", selected_scenario),
      !!reference_source := "Reference"
    )) %>%
    dplyr::rename(Diff = source) %>%
    tidyr::pivot_wider(names_from = year, values_from = value) %>%
    dplyr::select(Diff, `1995`, `2000`, `2005`, `2010`, `2015`, `2020`)

  tbl <- add_difference_rows(tbl)
  ft  <- create_flextable(tbl)
  return(ft)
}


# Input data --------------------------------------------------------------

data_hist              <- read.csv(here::here("data", "csv", "histdatabrazil.csv"))
data_magpie_raw        <- readRDS(file.path(magpie_output_path, "report.rds"))

# Cluster allocation JPG (produced by 01_convertClusterPDF.R)
jpg_cluster_file <- sub("\\.pdf$", ".jpg", cluster_pdf_file)


# Process MAgPIE data -----------------------------------------------------

data_bra_all <- data_magpie_raw %>%
  dplyr::filter(region == "BRA",
                scenario == selected_scenario,
                period <= 2050) %>%
  dplyr::select(-region, -model) %>%
  dplyr::rename(year = period, type = variable, source = scenario) %>%
  dplyr::select(type, source, year, value, unit)

data_bra <- data_bra_all %>%
  dplyr::mutate(type = dplyr::case_when(
    type == "Resources|Land Cover|+|Forest"                                          ~ "Forest",
    type == "Resources|Land Cover|Forest|Natural Forest|+|Primary Forest"            ~ "Primary Forest",
    type == "Resources|Land Cover|Forest|Natural Forest|+|Secondary Forest"          ~ "Secondary Forest",
    type == "Resources|Land Cover|+|Cropland"                                        ~ "Cropland",
    type == "Resources|Land Cover|+|Pastures and Rangelands"                         ~ "Pastures and Rangelands",
    type == "Resources|Land Cover|+|Other Land"                                      ~ "Other Land",
    type == "Resources|Land Cover|Cropland|Croparea|Crops|Oil crops|+|Soybean"       ~ "Soybean Area",
    type == "Resources|Land Cover|Cropland|Croparea|Crops|Cereals|+|Maize"           ~ "Maize Area",
    type == "Resources|Land Cover|Cropland|Croparea|Crops|Sugar crops|+|Sugar cane"  ~ "Sugarcane Area",
    type == "Production|Crops|Oil crops|+|Soybean"                                   ~ "Soybean Production",
    type == "Production|Crops|Cereals|+|Maize"                                       ~ "Maize Production",
    type == "Production|Crops|Sugar crops|+|Sugar cane"                              ~ "Sugarcane Production",
    type == "Emissions|CO2|Land|+|Land-use Change"                                   ~ "CO2",
    type == "Emissions|NO2|Land"                                                     ~ "N2O",
    type == "Emissions|CH4|Land"                                                     ~ "CH4",
    type == "Production|Livestock products|+|Ruminant meat"                          ~ "Ruminant Meat",
    type == "Resources|Land Cover|Other Land|Initial"                                ~ "Initial",
    type == "Resources|Land Cover|Other Land|Recovered"                              ~ "Recovered",
    type == "Resources|Land Cover|Other Land|Restored"                               ~ "Restored",
    type == "Resources|Land Cover|Cropland|Croparea|Crops|Oil crops|+|Oilpalms"      ~ "Oilpalms",
    type == "Resources|Land Cover|Cropland|Croparea|Crops|Other crops|+|Fruits Vegetables Nuts" ~ "FruVeg",
    type == "Productivity|Landuse Intensity Indicator Tau"                           ~ "TAU",
    TRUE ~ type
  )) %>%
  dplyr::filter(type %in% c(
    "Forest", "Primary Forest", "Secondary Forest",
    "Cropland", "Pastures and Rangelands", "Other Land",
    "Soybean Area", "Maize Area", "Sugarcane Area",
    "Soybean Production", "Maize Production", "Sugarcane Production",
    "CO2", "CH4", "N2O", "Ruminant Meat",
    "Initial", "Recovered", "Restored",
    "Oilpalms", "FruVeg", "TAU"
  ))

# Aggregate Oilpalms + FruVeg -> Permanent Crops
data_bra <- data_bra %>%
  dplyr::mutate(original_row = dplyr::row_number())

data_bra <- data_bra %>%
  dplyr::filter(!type %in% c("Oilpalms", "FruVeg")) %>%
  dplyr::bind_rows(
    data_bra %>%
      dplyr::filter(type %in% c("Oilpalms", "FruVeg")) %>%
      dplyr::group_by(year) %>%
      dplyr::summarise(type   = "Permanent Crops",
                       value  = sum(value),
                       source = unique(source),
                       unit   = unique(unit))
  ) %>%
  dplyr::arrange(original_row, year) %>%
  dplyr::select(-original_row)

# Derive Temporary Crops = Cropland - Permanent Crops
data_bra <- data_bra %>%
  dplyr::bind_rows(
    data_bra %>%
      dplyr::filter(type %in% c("Permanent Crops", "Cropland")) %>%
      dplyr::group_by(year) %>%
      dplyr::summarise(type   = "Temporary Crops",
                       value  = abs(diff(value)),
                       source = unique(source),
                       unit   = unique(unit))
  )

# Replace NA with 0
data_bra <- data_bra %>%
  dplyr::mutate(dplyr::across(dplyr::everything(), ~ tidyr::replace_na(., 0)))


# Charts: MAgPIE outputs --------------------------------------------------

# Primary + Secondary Forest (stacked bar)
data_primsecdforest <- data_bra %>%
  dplyr::filter(type %in% c("Primary Forest", "Secondary Forest")) %>%
  dplyr::mutate(type = factor(type, levels = c("Secondary Forest", "Primary Forest")))

primsecdforest <- ggplot2::ggplot(data_primsecdforest) +
  ggplot2::geom_bar(ggplot2::aes(x = year, y = value, fill = type),
                    alpha = 3/4, position = "stack", stat = "identity",
                    colour = "black", width = 3, linewidth = 0.6) +
  ggplot2::scale_fill_manual(values = c("#B4E5A2", "#3B7D23")) +
  theme_report() +
  ggplot2::guides(fill = ggplot2::guide_legend(nrow = 1, reverse = TRUE)) +
  ggplot2::labs(y = "Million hectares") +
  ggplot2::scale_x_continuous(breaks = seq(1995, 2050, by = 5)) +
  ggplot2::scale_y_continuous(limits = c(0, 600), expand = c(0, 0),
                               breaks = seq(0, 600, by = 100))

# Temporary + Permanent Crops (stacked bar)
data_temppermcrops <- data_bra %>%
  dplyr::filter(type %in% c("Temporary Crops", "Permanent Crops")) %>%
  dplyr::mutate(type = factor(type, levels = c("Permanent Crops", "Temporary Crops")))

temppermcrops <- ggplot2::ggplot(data_temppermcrops) +
  ggplot2::geom_bar(ggplot2::aes(x = year, y = value, fill = type),
                    alpha = 3/4, position = "stack", stat = "identity",
                    colour = "black", width = 3, linewidth = 0.6) +
  ggplot2::scale_fill_manual(values = c("#fec44f", "#d95f0e")) +
  theme_report() +
  ggplot2::guides(fill = ggplot2::guide_legend(nrow = 1, reverse = TRUE)) +
  ggplot2::labs(y = "Million hectares") +
  ggplot2::scale_x_continuous(breaks = seq(1995, 2050, by = 5)) +
  ggplot2::scale_y_continuous(limits = c(0, 200), expand = c(0, 0),
                               breaks = seq(0, 200, by = 50))


# Charts + Tables: historical comparisons ---------------------------------

# Secondary Forest
data_secdforest_area <- prepare_plot_data(data_bra, data_hist,
                                          product_type   = "Secondary Forest",
                                          product_levels = c(selected_scenario, "Mapbiomas"))
secdforest_area <- plot_comparisons(data_secdforest_area, ylab = "Area (Mha)")
ft_secdforest_area <- build_table(data_secdforest_area, "Mapbiomas")

# Cropland (total)
data_cropland_area <- prepare_plot_data(data_bra, data_hist,
                                        product_type   = "Cropland",
                                        product_levels = c(selected_scenario, "IBGE"))
cropland_area <- plot_comparisons(data_cropland_area, ylab = "Area (Mha)")
ft_cropland_area <- build_table(data_cropland_area, "IBGE")

# Pasture
data_pasture_area <- prepare_plot_data(data_bra, data_hist,
                                       product_type   = "Pastures and Rangelands",
                                       product_levels = c(selected_scenario, "LAPIG"))
pasture_area <- plot_comparisons(data_pasture_area, ylab = "Area (Mha)")
ft_pasture_area <- build_table(data_pasture_area, "LAPIG")

# Ruminant Meat
data_rummeat_prod <- prepare_plot_data(data_bra, data_hist,
                                       product_type    = "Ruminant Meat",
                                       plot_type       = "production",
                                       moisture_content = 0.5,
                                       product_levels  = c(selected_scenario, "FAOSTAT"))
rummeat_prod <- plot_comparisons(data_rummeat_prod, ylab = "Production (Mt)")
ft_rummeat_prod <- build_table(data_rummeat_prod, "FAOSTAT")

# Soybean Area
data_soybean_area <- prepare_plot_data(data_bra, data_hist,
                                       product_type   = "Soybean Area",
                                       product_levels = c(selected_scenario, "IBGE"))
soybean_area <- plot_comparisons(data_soybean_area, ylab = "Area (Mha)")
ft_soybean_area <- build_table(data_soybean_area, "IBGE")

# Soybean Production
data_soybean_prod <- prepare_plot_data(data_bra, data_hist,
                                       product_type    = "Soybean Production",
                                       plot_type       = "production",
                                       moisture_content = 0.13,
                                       product_levels  = c(selected_scenario, "IBGE"))
soybean_prod <- plot_comparisons(data_soybean_prod, ylab = "Production (Mt)")
ft_soybean_prod <- build_table(data_soybean_prod, "IBGE")

# Maize Area
data_maize_area <- prepare_plot_data(data_bra, data_hist,
                                     product_type   = "Maize Area",
                                     product_levels = c(selected_scenario, "IBGE"))
maize_area <- plot_comparisons(data_maize_area, ylab = "Area (Mha)")
ft_maize_area <- build_table(data_maize_area, "IBGE")

# Maize Production
data_maize_prod <- prepare_plot_data(data_bra, data_hist,
                                     product_type    = "Maize Production",
                                     plot_type       = "production",
                                     moisture_content = 0.15,
                                     product_levels  = c(selected_scenario, "IBGE"))
maize_prod <- plot_comparisons(data_maize_prod, ylab = "Production (Mt)")
ft_maize_prod <- build_table(data_maize_prod, "IBGE")

# Sugarcane Area
data_sugarcane_area <- prepare_plot_data(data_bra, data_hist,
                                         product_type   = "Sugarcane Area",
                                         product_levels = c(selected_scenario, "IBGE"))
sugarcane_area <- plot_comparisons(data_sugarcane_area, ylab = "Area (Mha)")
ft_sugarcane_area <- build_table(data_sugarcane_area, "IBGE")

# Sugarcane Production
data_sugarcane_prod <- prepare_plot_data(data_bra, data_hist,
                                         product_type    = "Sugarcane Production",
                                         plot_type       = "production",
                                         moisture_content = 0.5,
                                         product_levels  = c(selected_scenario, "IBGE"))
sugarcane_prod <- plot_comparisons(data_sugarcane_prod, ylab = "Production (Mt)")
ft_sugarcane_prod <- build_table(data_sugarcane_prod, "IBGE")

# CO2 Emissions
data_CO2 <- prepare_plot_data(data_bra, data_hist,
                              product_type   = "CO2",
                              product_levels = c(selected_scenario, "SEEG")) %>%
  dplyr::mutate(value = value / 1000, unit = "Gt CO2/yr")
ylab_CO2   <- parse(text = as.character(expression(paste("Gt ", CO[2], " /yr"))))
CO2_emissions <- plot_comparisons(data_CO2, ylab = ylab_CO2)
ft_CO2_emissions <- build_table(data_CO2, "SEEG")


# Image file paths --------------------------------------------------------

# dir_land_1995      <- file.path("maps", outputFolder, "land", "y1995", "Brazil")
# dir_land_2020      <- file.path("maps", outputFolder, "land", "y2020", "Brazil")
# dir_crops_1995     <- file.path("maps", outputFolder, "crop","y1995")
# dir_crops_2020     <- file.path("maps", outputFolder, "crop","y2020")
# dir_hist_crops_1995 <- file.path("maps", "IBGE", "y1995")
# dir_hist_crops_2020 <- file.path("maps", "IBGE", "y2020")
# 
# fp <- function(dir, filename) file.path(dir, filename)
fp <- function(...) file.path(...)
# 
# file_path_forest_1995        <- fp(dir_land_1995, "Forest 1995 Brazil.png")
# file_path_forest_2020        <- fp(dir_land_2020, "Forest 2020 Brazil.png")
# file_path_primforest_1995    <- fp(dir_land_1995, "Primary Forest 1995 Brazil.png")
# file_path_primforest_2020    <- fp(dir_land_2020, "Primary Forest 2020 Brazil.png")
# file_path_secdforest_1995    <- fp(dir_land_1995, "Secondary Forest 1995 Brazil.png")
# file_path_secdforest_2020    <- fp(dir_land_2020, "Secondary Forest 2020 Brazil.png")
# file_path_otherland_1995     <- fp(dir_land_1995, "Other Land 1995 Brazil.png")
# file_path_otherland_2020     <- fp(dir_land_2020, "Other Land 2020 Brazil.png")
# file_path_cropland_1995      <- fp(dir_land_1995, "Cropland 1995 Brazil.png")
# file_path_cropland_2020      <- fp(dir_land_2020, "Cropland 2020 Brazil.png")
# file_path_pasture_1995       <- fp(dir_land_1995, "Pastures and Rangelands 1995 Brazil.png")
# file_path_pasture_2020       <- fp(dir_land_2020, "Pastures and Rangelands 2020 Brazil.png")
# file_path_soybean_2020       <- fp(dir_crops_2020, "MAgPIE - Soybean2020.png")
# file_path_maize_2020         <- fp(dir_crops_2020, "MAgPIE - Maize2020.png")
# file_path_sugarcane_2020     <- fp(dir_crops_2020, "MAgPIE - Sugarcane2020.png")
# file_path_cropland_hist_1995 <- fp(dir_hist_crops_1995, "Cropland (ALL) IBGE (1995) All.png")
# file_path_cropland_hist_2020 <- fp(dir_hist_crops_2020, "Cropland (ALL) IBGE (2020).png")
# file_path_soybean_hist_2020  <- fp(dir_hist_crops_2020, "Soybean IBGE (2020).png")
# file_path_maize_hist_2020    <- fp(dir_hist_crops_2020, "Maize IBGE (2020).png")
# file_path_sugarcane_hist_2020 <- fp(dir_hist_crops_2020, "Sugarcane IBGE (2020).png")


# Land maps — built dynamically from land_years defined in config.R
land_map_paths <- lapply(land_years, function(year) {
  dir <- file.path("maps", outputFolder, "land", paste0("y", year), "Brazil")
  list(
    forest     = fp(dir, paste0("Forest ",                    year, " Brazil.png")),
    primforest = fp(dir, paste0("Primary Forest ",            year, " Brazil.png")),
    secdforest = fp(dir, paste0("Secondary Forest ",          year, " Brazil.png")),
    otherland  = fp(dir, paste0("Other Land ",                year, " Brazil.png")),
    cropland   = fp(dir, paste0("Cropland ",                  year, " Brazil.png")),
    pasture    = fp(dir, paste0("Pastures and Rangelands ",   year, " Brazil.png"))
  )
})
names(land_map_paths) <- paste0("y", land_years)

# Deforestation maps — built dynamically from forest_loss_periods defined in config.R
deforestation_map_paths <- lapply(forest_loss_periods, function(period) {
  dir <- file.path("maps", outputFolder, "diff", "deforestation", paste0("y", period["year_start"], 
                                                                         "-", period["year_end"]), "Brazil")
  list(
    primforest = fp(dir, paste0("Primary Forest ",  "Loss ",  period["year_start"], 
                                "-", period["year_end"], " Brazil.png"))
  )
})
names(deforestation_map_paths) <- sapply(forest_loss_periods, function(period) {
  paste0("y", period["year_start"], "_", period["year_end"])
})

# Crop share maps — built dynamically from crop_years defined in config.R
crop_map_paths <- lapply(crop_years, function(year) {
  dir <- file.path("maps", outputFolder, "crop", paste0("y", year))
  list(
    soybean   = fp(dir, paste0("MAgPIE - Soybean",   year, ".png")),
    maize     = fp(dir, paste0("MAgPIE - Maize",     year, ".png")),
    sugarcane = fp(dir, paste0("MAgPIE - Sugarcane", year, ".png"))
  )
})
names(crop_map_paths) <- paste0("y", crop_years)

# Historical IBGE maps — static, not scenario-dependent
hist_map_paths <- list(
  cropland_all_1995  = fp("maps", "IBGE", "y1995", "Cropland (ALL) IBGE (1995).png"),
  cropland_all_2020  = fp("maps", "IBGE", "y2020", "Cropland (ALL) IBGE (2020).png"),
  soybean_2020       = fp("maps", "IBGE", "y2020", "Soybean IBGE (2020).png"),
  maize_2020         = fp("maps", "IBGE", "y2020", "Maize IBGE (2020).png"),
  sugarcane_2020     = fp("maps", "IBGE", "y2020", "Sugarcane IBGE (2020).png")
)


# Word document assembly --------------------------------------------------

par_style <- fp_par(text.align = "left")

doc_word <- read_docx()

# Title
doc_word <- body_add_fpar(
  doc_word,
  fpar(
    ftext("Comparisons - MAgPIE and Historical Data",
          prop = fp_text(font.size = 15, color = "black", bold = TRUE,
                         font.family = "Calibri (Headings)")),
    fp_p = fp_par(text.align = "center")
  )
)
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_toc(doc_word)
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_break(doc_word)


# Cluster allocation ------------------------------------------------------

doc_word <- body_add_par(doc_word, "Cluster allocation", style = "heading 1")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_img(doc_word, src = jpg_cluster_file,
                         width = 6.5, height = 4.5, style = "centered")
doc_word <- body_add_break(doc_word)


# MAgPIE Outputs ----------------------------------------------------------

doc_word <- body_add_par(doc_word, "MAgPIE Outputs", style = "heading 1")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "Primary Forest and Secondary Forest", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_gg(doc_word, value = primsecdforest, res = 600,
                        width = 6, height = 3.33, style = "centered")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "Cropland Area (Permanent + Temporary)", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_gg(doc_word, value = temppermcrops, res = 600,
                        width = 6, height = 3.33, style = "centered")
doc_word <- body_add_break(doc_word)


# Historical Data Comparisons ---------------------------------------------

doc_word <- body_add_par(doc_word, "Historical Data Comparisons", style = "heading 1")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")

doc_word <- body_add_par(doc_word, "Secondary Forest Area", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_gg(doc_word, value = secdforest_area, res = 600,
                        width = 6, height = 3.33, style = "centered")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_flextable(doc_word, value = ft_secdforest_area)
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Variable: Resources|Land Cover|Forest|Natural Forest|+|Secondary Forest", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Green:  [diff] <= 10%",       prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Orange: 10% < [diff] <= 20%", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Red:    [diff] > 20%",        prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_break(doc_word)

doc_word <- body_add_par(doc_word, "Cropland Area (Total)", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_gg(doc_word, value = cropland_area, res = 600,
                        width = 6, height = 3.33, style = "centered")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_flextable(doc_word, value = ft_cropland_area)
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Variable: Resources|Land Cover|+|Cropland", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Green:  [diff] <= 10%",       prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Orange: 10% < [diff] <= 20%", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Red:    [diff] > 20%",        prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_break(doc_word)

doc_word <- body_add_par(doc_word, "Pasture Area", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_gg(doc_word, value = pasture_area, res = 600,
                        width = 6, height = 3.33, style = "centered")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_flextable(doc_word, value = ft_pasture_area)
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Variable: Resources|Land Cover|+|Pastures and Rangelands", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Green:  [diff] <= 10%",       prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Orange: 10% < [diff] <= 20%", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Red:    [diff] > 20%",        prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_break(doc_word)

doc_word <- body_add_par(doc_word, "Ruminant Meat Production", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_gg(doc_word, value = rummeat_prod, res = 600,
                        width = 6, height = 3.33, style = "centered")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_flextable(doc_word, value = ft_rummeat_prod)
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Variable: Production|Livestock products|+|Ruminant meat", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Green:  [diff] <= 10%",       prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Orange: 10% < [diff] <= 20%", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Red:    [diff] > 20%",        prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_break(doc_word)

doc_word <- body_add_par(doc_word, "Soybean Area", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_gg(doc_word, value = soybean_area, res = 600,
                        width = 6, height = 3.33, style = "centered")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_flextable(doc_word, value = ft_soybean_area)
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Variable: Resources|Land Cover|Cropland|Croparea|Crops|Oil crops|+|Soybean", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Green:  [diff] <= 10%",       prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Orange: 10% < [diff] <= 20%", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Red:    [diff] > 20%",        prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_break(doc_word)

doc_word <- body_add_par(doc_word, "Soybean Production", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_gg(doc_word, value = soybean_prod, res = 600,
                        width = 6, height = 3.33, style = "centered")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_flextable(doc_word, value = ft_soybean_prod)
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Variable: Production|Crops|Oil crops|+|Soybean", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Green:  [diff] <= 10%",       prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Orange: 10% < [diff] <= 20%", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Red:    [diff] > 20%",        prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_break(doc_word)

doc_word <- body_add_par(doc_word, "Maize Area", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_gg(doc_word, value = maize_area, res = 600,
                        width = 6, height = 3.33, style = "centered")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_flextable(doc_word, value = ft_maize_area)
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Variable: Resources|Land Cover|Cropland|Croparea|Crops|Cereals|+|Maize", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Green:  [diff] <= 10%",       prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Orange: 10% < [diff] <= 20%", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Red:    [diff] > 20%",        prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_break(doc_word)

doc_word <- body_add_par(doc_word, "Maize Production", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_gg(doc_word, value = maize_prod, res = 600,
                        width = 6, height = 3.33, style = "centered")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_flextable(doc_word, value = ft_maize_prod)
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Variable: Production|Crops|Cereals|+|Maize", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Green:  [diff] <= 10%",       prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Orange: 10% < [diff] <= 20%", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Red:    [diff] > 20%",        prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_break(doc_word)

doc_word <- body_add_par(doc_word, "Sugarcane Area", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_gg(doc_word, value = sugarcane_area, res = 600,
                        width = 6, height = 3.33, style = "centered")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_flextable(doc_word, value = ft_sugarcane_area)
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Variable: Resources|Land Cover|Cropland|Croparea|Crops|Sugar crops|+|Sugar cane", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Green:  [diff] <= 10%",       prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Orange: 10% < [diff] <= 20%", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Red:    [diff] > 20%",        prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_break(doc_word)

doc_word <- body_add_par(doc_word, "Sugarcane Production", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_gg(doc_word, value = sugarcane_prod, res = 600,
                        width = 6, height = 3.33, style = "centered")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_flextable(doc_word, value = ft_sugarcane_prod)
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Variable: Production|Crops|Sugar crops|+|Sugar cane", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Green:  [diff] <= 10%",       prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Orange: 10% < [diff] <= 20%", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Red:    [diff] > 20%",        prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_break(doc_word)

formatted_CO2 <- fpar(
  ftext("CO",              fp_text(font.size = 13, bold = TRUE, font.family = "Calibri (Headings)")),
  ftext("2",               fp_text(font.size = 13, bold = TRUE, font.family = "Calibri (Headings)", vertical.align = "subscript")),
  ftext(" emissions (AFOLU)", fp_text(font.size = 13, bold = TRUE, font.family = "Calibri (Headings)"))
)
doc_word <- body_add_fpar(doc_word, formatted_CO2, style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_gg(doc_word, value = CO2_emissions, res = 600,
                        width = 6, height = 3.33, style = "centered")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_flextable(doc_word, value = ft_CO2_emissions)
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Variable: Emissions|CO2|Land|+|Land-use Change", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_fpar(doc_word, fpar(ftext("Green:  [diff] <= 10%",       prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Orange: 10% < [diff] <= 20%", prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_fpar(doc_word, fpar(ftext("Red:    [diff] > 20%",        prop = fp_text(font.size = 9)), fp_p = par_style))
doc_word <- body_add_break(doc_word)


# # Land Use Maps -----------------------------------------------------------
# 
# doc_word <- body_add_par(doc_word, "Land Use Maps", style = "heading 1")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# 
# doc_word <- body_add_par(doc_word, "Forest (Primary + Secondary)", style = "heading 2")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_forest_1995,        width = 4, height = 4, style = "centered")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_forest_2020,        width = 4, height = 4, style = "centered")
# doc_word <- body_add_break(doc_word)
# 
# doc_word <- body_add_par(doc_word, "Primary Forest", style = "heading 2")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_primforest_1995,    width = 4, height = 4, style = "centered")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_primforest_2020,    width = 4, height = 4, style = "centered")
# doc_word <- body_add_break(doc_word)
# 
# doc_word <- body_add_par(doc_word, "Secondary Forest", style = "heading 2")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_secdforest_1995,    width = 4, height = 4, style = "centered")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_secdforest_2020,    width = 4, height = 4, style = "centered")
# doc_word <- body_add_break(doc_word)
# 
# doc_word <- body_add_par(doc_word, "Other Land", style = "heading 2")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_otherland_1995,     width = 4, height = 4, style = "centered")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_otherland_2020,     width = 4, height = 4, style = "centered")
# doc_word <- body_add_break(doc_word)
# 
# doc_word <- body_add_par(doc_word, "Pastures and Rangelands", style = "heading 2")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_pasture_1995,       width = 4, height = 4, style = "centered")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_pasture_2020,       width = 4, height = 4, style = "centered")
# doc_word <- body_add_break(doc_word)
# 
# doc_word <- body_add_par(doc_word, "Cropland", style = "heading 2")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_cropland_1995,      width = 4, height = 4, style = "centered")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_cropland_2020,      width = 4, height = 4, style = "centered")
# doc_word <- body_add_break(doc_word)
# 
# 
# # Crop Maps 2020 ----------------------------------------------------------
# 
# doc_word <- body_add_par(doc_word, "Crop Maps 2020", style = "heading 1")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# 
# doc_word <- body_add_par(doc_word, "Cropland (Total)", style = "heading 2")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_cropland_2020,      width = 4, height = 4, style = "centered")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_cropland_hist_2020, width = 4, height = 4, style = "centered")
# doc_word <- body_add_break(doc_word)
# 
# doc_word <- body_add_par(doc_word, "Soybean (2020)", style = "heading 2")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_soybean_2020,       width = 4, height = 4, style = "centered")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_soybean_hist_2020,  width = 4, height = 4, style = "centered")
# doc_word <- body_add_break(doc_word)
# 
# doc_word <- body_add_par(doc_word, "Maize (2020)", style = "heading 2")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_maize_2020,         width = 4, height = 4, style = "centered")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_maize_hist_2020,    width = 4, height = 4, style = "centered")
# doc_word <- body_add_break(doc_word)
# 
# doc_word <- body_add_par(doc_word, "Sugarcane (2020)", style = "heading 2")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_sugarcane_2020,     width = 4, height = 4, style = "centered")
# doc_word <- body_add_par(doc_word, "", style = "Normal")
# doc_word <- body_add_img(doc_word, src = file_path_sugarcane_hist_2020, width = 4, height = 4, style = "centered")
# doc_word <- body_add_break(doc_word)


# Land Use Maps -----------------------------------------------------------

doc_word <- body_add_par(doc_word, "Land Use Maps", style = "heading 1")
doc_word <- body_add_par(doc_word, "", style = "Normal")
#doc_word <- body_add_par(doc_word, "", style = "Normal")

doc_word <- body_add_par(doc_word, "Forest (Primary + Secondary)", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
for (yr in paste0("y", land_years)) {
  doc_word <- body_add_img(doc_word, src = land_map_paths[[yr]]$forest, width = 4, height = 4, style = "centered")
  doc_word <- body_add_par(doc_word, "", style = "Normal")
}
doc_word <- body_add_break(doc_word)

doc_word <- body_add_par(doc_word, "Primary Forest", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
for (yr in paste0("y", land_years)) {
  doc_word <- body_add_img(doc_word, src = land_map_paths[[yr]]$primforest, width = 4, height = 4, style = "centered")
  doc_word <- body_add_par(doc_word, "", style = "Normal")
}
doc_word <- body_add_break(doc_word)

doc_word <- body_add_par(doc_word, "Secondary Forest", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
for (yr in paste0("y", land_years)) {
  doc_word <- body_add_img(doc_word, src = land_map_paths[[yr]]$secdforest, width = 4, height = 4, style = "centered")
  doc_word <- body_add_par(doc_word, "", style = "Normal")
}
doc_word <- body_add_break(doc_word)

doc_word <- body_add_par(doc_word, "Other Land", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
for (yr in paste0("y", land_years)) {
  doc_word <- body_add_img(doc_word, src = land_map_paths[[yr]]$otherland, width = 4, height = 4, style = "centered")
  doc_word <- body_add_par(doc_word, "", style = "Normal")
}
doc_word <- body_add_break(doc_word)

doc_word <- body_add_par(doc_word, "Pastures and Rangelands", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
for (yr in paste0("y", land_years)) {
  doc_word <- body_add_img(doc_word, src = land_map_paths[[yr]]$pasture, width = 4, height = 4, style = "centered")
  doc_word <- body_add_par(doc_word, "", style = "Normal")
}
doc_word <- body_add_break(doc_word)

doc_word <- body_add_par(doc_word, "Cropland", style = "heading 2")
doc_word <- body_add_par(doc_word, "", style = "Normal")
for (yr in paste0("y", land_years)) {
  doc_word <- body_add_img(doc_word, src = land_map_paths[[yr]]$cropland, width = 4, height = 4, style = "centered")
  doc_word <- body_add_par(doc_word, "", style = "Normal")
}
doc_word <- body_add_break(doc_word)


# Deforestation Maps -----------------------------------------------------------

doc_word <- body_add_par(doc_word, "Deforestation Maps", style = "heading 1")
doc_word <- body_add_par(doc_word, "", style = "Normal")
for (yr in names(deforestation_map_paths)) {
  doc_word <- body_add_img(doc_word, src = deforestation_map_paths[[yr]]$primforest, width = 4, height = 4, style = "centered")
  doc_word <- body_add_par(doc_word, "", style = "Normal")
}
doc_word <- body_add_break(doc_word)

# Crop Maps ---------------------------------------------------------------

doc_word <- body_add_par(doc_word, paste("Crop Maps", paste(crop_years, collapse = " & ")), style = "heading 1")
doc_word <- body_add_par(doc_word, "", style = "Normal")
doc_word <- body_add_par(doc_word, "", style = "Normal")

for (yr in paste0("y", crop_years)) {
  year_label <- sub("y", "", yr)
  
  doc_word <- body_add_par(doc_word, paste0("Cropland (Total) (", year_label, ")"), style = "heading 2")
  doc_word <- body_add_par(doc_word, "", style = "Normal")
  doc_word <- body_add_img(doc_word, src = land_map_paths[[yr]]$cropland,       width = 4, height = 4, style = "centered")
  doc_word <- body_add_par(doc_word, "", style = "Normal")
  doc_word <- body_add_img(doc_word, src = hist_map_paths$cropland_all_2020,     width = 4, height = 4, style = "centered")
  doc_word <- body_add_break(doc_word)
  
  doc_word <- body_add_par(doc_word, paste0("Soybean (", year_label, ")"), style = "heading 2")
  doc_word <- body_add_par(doc_word, "", style = "Normal")
  doc_word <- body_add_img(doc_word, src = crop_map_paths[[yr]]$soybean,         width = 4, height = 4, style = "centered")
  doc_word <- body_add_par(doc_word, "", style = "Normal")
  doc_word <- body_add_img(doc_word, src = hist_map_paths$soybean_2020,          width = 4, height = 4, style = "centered")
  doc_word <- body_add_break(doc_word)
  
  doc_word <- body_add_par(doc_word, paste0("Maize (", year_label, ")"), style = "heading 2")
  doc_word <- body_add_par(doc_word, "", style = "Normal")
  doc_word <- body_add_img(doc_word, src = crop_map_paths[[yr]]$maize,           width = 4, height = 4, style = "centered")
  doc_word <- body_add_par(doc_word, "", style = "Normal")
  doc_word <- body_add_img(doc_word, src = hist_map_paths$maize_2020,            width = 4, height = 4, style = "centered")
  doc_word <- body_add_break(doc_word)
  
  doc_word <- body_add_par(doc_word, paste0("Sugarcane (", year_label, ")"), style = "heading 2")
  doc_word <- body_add_par(doc_word, "", style = "Normal")
  doc_word <- body_add_img(doc_word, src = crop_map_paths[[yr]]$sugarcane,       width = 4, height = 4, style = "centered")
  doc_word <- body_add_par(doc_word, "", style = "Normal")
  doc_word <- body_add_img(doc_word, src = hist_map_paths$sugarcane_2020,        width = 4, height = 4, style = "centered")
  doc_word <- body_add_break(doc_word)
}



# Save .docx --------------------------------------------------------------

if (!dir.exists(here::here("report"))) dir.create(here::here("report"))

output_path_docx <- glue::glue("comparisonsHistdata_{selected_scenario}.docx")
print(doc_word, target = here::here("report", output_path_docx))
message("Word document saved: report/", output_path_docx)


# Convert to PDF (Windows only via PowerShell) ----------------------------

docx_path <- normalizePath(here::here("report", output_path_docx))
pdf_path  <- sub("\\.docx$", ".pdf", docx_path)

ps_script <- glue::glue(
  '
$word = New-Object -ComObject Word.Application
$word.Visible = $false
$doc = $word.Documents.Open("{docx_path}")
$doc.SaveAs2("{pdf_path}", 17)
$doc.Close()
$word.Quit()
'
)

ps_file <- tempfile(fileext = ".ps1")
writeLines(ps_script, ps_file)
system2("powershell", args = c("-ExecutionPolicy", "Bypass", "-File", shQuote(ps_file)))

if (Sys.info()["sysname"] == "Windows") {
  shell.exec(pdf_path)
}

message("PDF saved: ", pdf_path)
