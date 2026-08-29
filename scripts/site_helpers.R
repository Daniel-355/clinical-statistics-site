slugify <- function(x) {
  x <- iconv(x, to = "ASCII//TRANSLIT")
  x <- tolower(gsub("[^a-zA-Z0-9]+", "-", x))
  gsub("(^-+|-+$)", "", x)
}

read_front_matter <- function(path) {
  lines <- readLines(path, warn = FALSE, encoding = "UTF-8")
  if (length(lines) < 3 || trimws(lines[1]) != "---") stop("Missing YAML front matter: ", path)
  ends <- which(trimws(lines[-1]) == "---")
  if (!length(ends)) stop("Unclosed YAML front matter: ", path)
  yaml::yaml.load(paste(lines[2:ends[1]], collapse = "\n"))
}

post_records <- function(include_drafts = FALSE) {
  paths <- Sys.glob("posts/*/index.qmd")
  out <- lapply(paths, function(path) {
    meta <- read_front_matter(path)
    meta$path <- gsub("\\\\", "/", path)
    meta$slug <- basename(dirname(path))
    meta
  })
  if (!include_drafts) out <- Filter(function(x) identical(x$status %||% "published", "published"), out)
  out
}

`%||%` <- function(x, y) if (is.null(x) || length(x) == 0) y else x

q <- function(x) paste0('"', gsub('"', '\\"', x, fixed = TRUE), '"')

write_utf8 <- function(lines, path) {
  dir.create(dirname(path), recursive = TRUE, showWarnings = FALSE)
  writeLines(enc2utf8(lines), path, useBytes = TRUE)
}
