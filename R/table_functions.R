# R/table_functions.R ----------------------------------------------------
# Reusable functions for building comparison tables in the MAgPIE report.
# Sourced by 04_buildReport.R.


# Add absolute and relative difference rows to a wide table ---------------

add_difference_rows <- function(data) {
  # Absolute difference
  data <- data %>%
    dplyr::bind_rows(
      data.frame(
        Diff = "Absolute Difference",
        t(abs(as.numeric(data[1, -1]) - as.numeric(data[2, -1])))
      ) %>%
        setNames(names(data))
    )

  # Relative difference (%)
  data <- data %>%
    dplyr::bind_rows(
      data.frame(
        Diff = "Relative Difference (%)",
        t(((as.numeric(data[1, -1]) - as.numeric(data[2, -1])) /
             as.numeric(data[2, -1])) * 100)
      ) %>%
        setNames(names(data))
    )

  data <- data %>%
    dplyr::mutate(Diff = factor(
      Diff,
      levels = c("Relative Difference (%)", "Absolute Difference", "MAgPIE", "Reference")
    )) %>%
    dplyr::arrange(Diff) %>%
    dplyr::mutate(dplyr::across(`1995`:`2020`, ~ round(.x, 2)))

  return(data)
}


# Create styled flextable -------------------------------------------------
# Color coding for the Relative Difference (%) row:
#   Green  : |diff| <= 10%
#   Orange : 10% < |diff| <= 20%
#   Red    : |diff| > 20%

create_flextable <- function(data) {
  year_cols <- c("1995", "2000", "2005", "2010", "2015", "2020")

  ft <- flextable::flextable(data) %>%
    flextable::width(j = "Diff", width = 1.5) %>%
    flextable::bg(j = 1, bg = "white") %>%
    flextable::set_header_labels(Diff = "") %>%
    flextable::align(align = "center", part = "all") %>%
    flextable::bg(bg = "#A9CBE1", part = "header") %>%
    flextable::bg(j = "Diff", bg = "white", part = "header") %>%
    flextable::border(j = "Diff",
                      border.top = officer::fp_border(color = "white"), part = "header") %>%
    flextable::border(border.top = officer::fp_border(color = "black"), part = "body") %>%
    flextable::bold(part = "header") %>%
    flextable::fontsize(size = 9, part = "all")

  # Apply color coding per year column
  for (col in year_cols) {
    col_sym <- paste0("`", col, "`")

    ft <- ft %>%
      flextable::bg(
        i   = stats::as.formula(paste0("~ abs(`", col, "`) > 20 & Diff == 'Relative Difference (%)'") ),
        j   = col_sym,
        bg  = "#F28D8D",
        part = "body"
      ) %>%
      flextable::bg(
        i   = stats::as.formula(paste0("~ abs(`", col, "`) > 10 & abs(`", col, "`) <= 20 & Diff == 'Relative Difference (%)'") ),
        j   = col_sym,
        bg  = "wheat",
        part = "body"
      ) %>%
      flextable::bg(
        i   = stats::as.formula(paste0("~ abs(`", col, "`) <= 10 & Diff == 'Relative Difference (%)'") ),
        j   = col_sym,
        bg  = "#8BC34A",
        part = "body"
      )
  }

  return(ft)
}
