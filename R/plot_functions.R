# R/plot_functions.R -----------------------------------------------------
# Reusable ggplot2 plotting functions for the MAgPIE Brazil report.
# Sourced by 04_buildReport.R.


# Set y-axis ceiling ------------------------------------------------------
# Returns the nearest value >= x that is a multiple of the step size.
# Used to define clean y-axis upper limits.

set_y_axis_ceiling <- function(x) {
  y <- dplyr::case_when(
    x <= 4   ~ 0.4,
    x <= 40  ~ 4,
    x <= 400 ~ 40,
    TRUE     ~ 80
  )

  remainder <- ceiling(x / y)

  if (remainder * y == x) {
    return(x + y)
  } else {
    return(remainder * y)
  }
}


# Shared theme ------------------------------------------------------------

theme_report <- function() {
  ggplot2::theme_bw() +
    ggplot2::theme(
      strip.placement      = "outside",
      strip.background     = ggplot2::element_rect(fill = NA, color = "white"),
      panel.spacing        = ggplot2::unit(-.01, "cm"),
      legend.position      = "bottom",
      legend.title         = ggplot2::element_blank(),
      legend.text          = ggplot2::element_text(size = 9),
      axis.text.x          = ggplot2::element_text(size = 9, colour = "black"),
      axis.text.y          = ggplot2::element_text(size = 9, colour = "black"),
      axis.title.y         = ggplot2::element_text(size = 10, vjust = +3),
      axis.title.x         = ggplot2::element_blank(),
      plot.title           = ggplot2::element_text(size = 12, hjust = 0.5),
      strip.text           = ggplot2::element_text(size = 10),
      panel.grid.minor.x   = ggplot2::element_blank(),
      panel.grid.major.x   = ggplot2::element_blank(),
      panel.grid.major.y   = ggplot2::element_line(size = .1),
      panel.background     = ggplot2::element_rect(size = 1),
      panel.border         = ggplot2::element_rect(linetype = "solid", colour = "black", size = 1)
    )
}


# Bar chart comparison (MAgPIE vs reference) ------------------------------

plot_comparisons <- function(data, xlab = "Year", ylab) {
  c_palette       <- c("#3366CC", "#ffbe0b")
  yaxis_ceiling   <- set_y_axis_ceiling(max(data$value))
  yaxis_interval  <- yaxis_ceiling / 4

  ggplot2::ggplot(data) +
    ggplot2::geom_bar(ggplot2::aes(x = year, y = value, fill = source),
                      alpha    = 3/4,
                      position = ggplot2::position_dodge(),
                      stat     = "identity",
                      colour   = "black",
                      width    = 3,
                      linewidth = 0.6) +
    ggplot2::scale_fill_manual(values = c_palette) +
    theme_report() +
    ggplot2::guides(fill = ggplot2::guide_legend(nrow = 1, reverse = FALSE)) +
    ggplot2::labs(x = xlab, y = ylab) +
    ggplot2::scale_x_continuous(
      breaks = seq(from = min(data$year), to = max(data$year), by = 5)
    ) +
    ggplot2::scale_y_continuous(
      limits = c(0, yaxis_ceiling),
      expand = c(0, 0),
      breaks = seq(0, yaxis_ceiling, by = yaxis_interval)
    )
}


# Line chart (e.g. TAU indicator) -----------------------------------------

plot_line <- function(data, xlab = "Year", ylab) {
  yaxis_ceiling  <- set_y_axis_ceiling(max(data$value))
  yaxis_interval <- yaxis_ceiling / 4

  ggplot2::ggplot(data, ggplot2::aes(x = year, y = value, color = source, group = source)) +
    ggplot2::geom_line(linewidth = 0.5) +
    ggplot2::geom_point(size = 1.2, shape = 18) +
    theme_report() +
    ggplot2::guides(fill = ggplot2::guide_legend(nrow = 1, reverse = FALSE)) +
    ggplot2::labs(x = xlab, y = ylab) +
    ggplot2::scale_x_continuous(
      breaks = seq(from = min(data$year), to = max(data$year), by = 5)
    ) +
    ggplot2::scale_y_continuous(
      limits = c(0, yaxis_ceiling),
      expand = c(0, 0),
      breaks = seq(0, yaxis_ceiling, by = yaxis_interval)
    )
}
