#!/usr/bin/env Rscript
source("scripts/site_helpers.R")

new_post <- function(title, section, category, topic, type = "article", status = "draft") {
  if (!type %in% c("article", "tutorial", "resource")) stop("type must be article, tutorial, or resource")
  slug <- slugify(title)
  root <- file.path("posts", slug)
  if (dir.exists(root)) stop("Post already exists: ", root)
  dir.create(file.path(root, "files"), recursive = TRUE)
  dir.create(file.path(root, "images"), recursive = TRUE)
  template <- readLines("templates/post-template.qmd", warn = FALSE)
  values <- c(TITLE = title, DATE = as.character(Sys.Date()), SECTION = section, CATEGORY = category, TOPIC = topic, TYPE = type, STATUS = status)
  for (key in names(values)) template <- gsub(paste0("{{", key, "}}"), values[[key]], template, fixed = TRUE)
  write_utf8(template, file.path(root, "index.qmd"))
  message("Created ", root, "/ with index.qmd, files/, and images/.")
  invisible(root)
}

if (sys.nframe() == 0) {
  args <- commandArgs(trailingOnly = TRUE)
  if (length(args)) {
    if (length(args) < 4) stop("Usage: Rscript scripts/new_post.R TITLE SECTION CATEGORY TOPIC [TYPE] [STATUS]")
    do.call(new_post, as.list(args[seq_len(min(length(args), 6))]))
  }
}
