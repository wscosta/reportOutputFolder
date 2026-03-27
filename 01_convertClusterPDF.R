# 01_convertClusterPDF.R -------------------------------------------------
# Converts the MAgPIE cluster allocation PDF (spamplot) to JPG.
# The resulting JPG is used directly in the report (04_buildReport.R).
#
# Input:  magpie/output/{outputFolder}/spamplot_*.pdf
# Output: magpie/output/{outputFolder}/spamplot_*.jpg
#
# Run order: FIRST — before any other pipeline script.


# Libraries ---------------------------------------------------------------

library(here)
library(pdftools)
library(magick)
library(withr)


# Config ------------------------------------------------------------------

source(here::here("config.R"))


# Function ----------------------------------------------------------------

pdf_to_jpg <- function(pdf_path, dpi = 300, max_width = NULL) {
  pdf_path <- normalizePath(pdf_path)
  jpg_path <- sub("\\.pdf$", ".jpg", pdf_path)

  bitmap <- suppressWarnings(pdf_render_page(pdf_path, page = 1, dpi = dpi))
  img    <- image_read(bitmap)

  if (!is.null(max_width)) {
    img <- image_scale(img, paste0(max_width))
  }

  image_write(img, path = jpg_path, format = "jpg", quality = 100)
  message("JPG created at: ", jpg_path)
  return(jpg_path)
}


# Run ---------------------------------------------------------------------

withr::with_output_sink(nullfile(), {
  withr::with_message_sink(nullfile(), {
    jpg_cluster_file <- pdf_to_jpg(cluster_pdf_file, dpi = 300, max_width = 2000)
  })
})
message("JPG created at: ", jpg_cluster_file)