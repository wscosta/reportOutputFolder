# MAgPIE Report Generation Workflow

This repository contains an R-based workflow to process outputs from MAgPIE (Model of Agricultural Production and its Impact on the Environment), generate spatial maps, and build automated reports for land-use scenario analysis.

## 📁 Project Structure
```
.
├── data/ # Input datasets (CSV, shapefiles, etc.)
│ ├── csv/
│ └── shapefiles/
├── R/ # Helper functions (plots and tables)
├── maps/ # Generated maps
├── output/ # Model outputs and intermediate files (not tracked)
├── report/ # Final reports (not tracked)
├── logs/ # Logs (not tracked)
│
├── 01_convertClusterPDF.R
├── 02_createLandMaps.R
├── 03_createDiffMaps.R
├── 04_buildReport.R
├── config.R # Configuration file
├── run_all.R # Main pipeline script
└── reportOutputFolder.Rproj
```

## ⚙️ Workflow Overview

The workflow is organized into sequential steps:

1. **Data preparation**
   - Load and organize input datasets (CSV, shapefiles, MAgPIE outputs)

2. **Map generation**
   - Create land-use maps
   - Generate crop-specific maps
   - Produce difference maps across scenarios and time periods

3. **Report building**
   - Compile results into structured outputs (PDF/DOCX)

4. **Automation**
   - The entire pipeline can be executed via a single script

## ▶️ How to Run

To execute the full workflow:

```r
Rscript run_all.R
```

or in an R session:

```r
source("run_all.R")
```

Make sure to adjust paths and parameters in:

```r
config.R
```

## 📦 Dependencies

This project relies on R and commonly used packages for:

- Data manipulation
- Spatial data processing
- Plotting and visualization
- Report generation

## 📊 Inputs

- MAgPIE output files (e.g., `.rds`, `.mz`)
- CSV datasets (e.g., crop areas, historical data) **Historical data can be inserted in `data/csv/histdatabrazil.csv`**
- Shapefiles for spatial boundaries (Brazil regions, biomes, states, etc.)

## 📤 Outputs

Generated outputs include:

- Spatial maps (PNG)
- Scenario comparison maps
- Automated reports (PDF/DOCX)

**Note:** Output folders (`maps/`, `output/`, `report/`) are excluded from version control and can be regenerated.

## 🔁 Reproducibility

This repository is designed to be reproducible:

- All results are generated from scripts
- No large outputs are stored in the repository
- Configuration is centralized in `config.R`

## 🧑‍💻 Author

Wanderson Costa

## 📌 Notes

- Ensure all required input data is available before running the pipeline
