#!/usr/bin/env Rscript

if (!requireNamespace("renv", quietly = TRUE)) {
  install.packages("renv", repos = "https://cloud.r-project.org")
}

rendered_posts <- Sys.glob("posts/*/index.qmd")
dependencies <- renv::dependencies(
  path = c(rendered_posts, "scripts"),
  progress = FALSE
)
packages <- sort(unique(stats::na.omit(dependencies$Package)))

renv::snapshot(
  packages = packages,
  library = .libPaths(),
  prompt = FALSE
)
message("Updated renv.lock for ", length(packages), " direct website dependencies.")
