# run_all.R --------------------------------------------------------------
# Master script: runs the full report pipeline in sequence.
# Stops immediately if any step fails and logs all output to file.
#
# Usage:
#   - RStudio:        open this file and click Source
#   - Command line:   Rscript run_all.R


# Log setup ---------------------------------------------------------------

if (!dir.exists("logs")) dir.create("logs")

log_file <- file.path(
  "logs",
  format(Sys.time(), "pipeline_%Y-%m-%d_%H-%M-%S.log")
)

# Tee: writes to both console and log file
log <- function(...) {
  msg <- paste0("[", format(Sys.time(), "%Y-%m-%d %H:%M:%S"), "] ", ...)
  message(msg)
  cat(msg, "\n", file = log_file, append = TRUE)
}


# Pipeline steps ----------------------------------------------------------

steps <- list(
  list(script = "01_convertClusterPDF.R", label = "Convert cluster PDF to JPG"),
  list(script = "02_createLandMaps.R",    label = "Create land cover and crop maps"),
  list(script = "03_createDiffMaps.R",    label = "Create deforestation maps"),
  list(script = "04_buildReport.R",       label = "Build Word and PDF report")
)


# Run ---------------------------------------------------------------------

log("========================================")
log("Pipeline started")
log("Working directory: ", getwd())
log("========================================")

pipeline_start <- proc.time()

for (step in steps) {

  log("----------------------------------------")
  log("STEP: ", step$label)
  log("Script: ", step$script)

  step_start <- proc.time()

  tryCatch(
    {
      source(step$script, echo = FALSE, local = FALSE)

      elapsed <- round((proc.time() - step_start)[["elapsed"]])
      log("OK — completed in ", elapsed, "s")
    },
    error = function(e) {
      log("FAILED — ", conditionMessage(e))
      log("----------------------------------------")
      log("Pipeline aborted at: ", step$script)
      log("See log for details: ", log_file)
      stop(
        "\n\nPipeline failed at: ", step$script,
        "\nReason: ", conditionMessage(e),
        "\nFull log saved at: ", log_file,
        call. = FALSE
      )
    }
  )
}

total <- round((proc.time() - pipeline_start)[["elapsed"]])
log("========================================")
log("Pipeline completed successfully in ", total, "s")
log("Report saved in: report/")
log("========================================")
