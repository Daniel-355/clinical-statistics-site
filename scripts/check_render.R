#!/usr/bin/env Rscript
source("scripts/site_helpers.R")
drafts <- Filter(function(p) identical(p$status, "draft"), post_records(TRUE))
leaked <- vapply(drafts, function(p) file.exists(file.path("_site", "posts", p$slug, "index.html")), logical(1))
if (any(leaked)) stop("Draft output detected: ", paste(vapply(drafts[leaked], function(p) p$slug, ""), collapse = ", "))
message("Post-render draft check passed.")
