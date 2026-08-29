#!/usr/bin/env Rscript
suppressPackageStartupMessages(library(yaml))
source("scripts/site_helpers.R")
errors <- character()
nav <- tryCatch(yaml::read_yaml("config/navigation.yml")$navigation, error = function(e) { errors <<- c(errors, paste("Malformed config/navigation.yml:", e$message)); list() })
sections <- Filter(function(x) !is.null(x$slug), nav)
all_posts <- tryCatch(post_records(TRUE), error = function(e) { errors <<- c(errors, e$message); list() })

required <- c("title", "section", "category", "topic", "type", "status")
valid_types <- c("article", "tutorial", "resource")
valid_status <- c("draft", "published")
for (p in all_posts) {
  missing <- required[vapply(required, function(k) is.null(p[[k]]) || !nzchar(as.character(p[[k]])), logical(1))]
  if (length(missing)) errors <- c(errors, paste0(p$path, ": missing required metadata: ", paste(missing, collapse = ", ")))
  section <- Filter(function(s) identical(s$title, p$section), sections)
  if (!length(section)) errors <- c(errors, paste0(p$path, ': invalid section "', p$section %||% "", '"'))
  else if (!p$category %in% vapply(section[[1]]$children, `[[`, "", "title")) errors <- c(errors, paste0(p$path, ': invalid category "', p$category, '" for section "', p$section, '"'))
  if (!is.null(p$type) && !p$type %in% valid_types) errors <- c(errors, paste0(p$path, ": type must be article, tutorial, or resource"))
  if (!is.null(p$status) && !p$status %in% valid_status) errors <- c(errors, paste0(p$path, ": status must be draft or published"))
  for (field in c("downloads", "pdf")) for (item in p[[field]] %||% list()) {
    target <- file.path(dirname(p$path), item$file %||% "")
    if (!file.exists(target)) errors <- c(errors, paste0("Post:\n", p$path, "\nreferences missing file:\n", item$file %||% "<empty>"))
  }
}
slugs <- vapply(all_posts, function(p) p$slug, "")
if (anyDuplicated(slugs)) errors <- c(errors, paste("Duplicate post slugs:", paste(unique(slugs[duplicated(slugs)]), collapse = ", ")))
drafts <- Filter(function(p) identical(p$status, "draft"), all_posts)
if (file.exists("_generated/home.qmd")) for (p in drafts) if (any(grepl(p$slug, readLines("_generated/home.qmd", warn = FALSE), fixed = TRUE))) errors <- c(errors, paste0("Draft leaked into homepage: ", p$path))
if (length(errors)) { cat("VALIDATION FAILED\n\n", paste(errors, collapse = "\n\n"), "\n", sep = ""); quit(status = 1) }
cat("Validation passed:", length(all_posts), "posts checked (", length(drafts), " draft).\n")
