# Listing 1: Adverse Events
# Standalone program that creates simulated ADSL/ADAE data, Excel workbooks,
# and a clinical-style adverse event listing. No external R scripts are required.

set.seed(4101)

script_arg <- grep("^--file=", commandArgs(), value = TRUE)
output_dir <- if (length(script_arg)) {
  dirname(normalizePath(gsub("~[+]~", " ", sub("^--file=", "", script_arg[1])), mustWork = TRUE))
} else {
  getwd()
}
output_stem <- "01_adverse_events_listing"

# Original source-style subject data. CSG has been replaced with PILOT.
original_adsl <- data.frame(
  USUBJID = sprintf("PILOT-%04d", 1001:1009),
  TRT01A = c("Arm B", "Arm A", "Arm C", "Arm B", "Arm C", "Arm B", "Arm A", "Arm C", "Arm A"),
  TRT01AN = c(2, 1, 3, 2, 3, 2, 1, 3, 1),
  AGE = c(81, 84, 77, 74, 82, 84, 65, 77, 88),
  RACE = "WHITE",
  SEX = c("F", "M", "M", "F", "M", "F", "M", "M", "F"),
  stringsAsFactors = FALSE
)

# Simulate twice the original population and create AE records with the
# same source-style variables and controlled categorical distributions.
adsl <- rbind(original_adsl, transform(original_adsl, USUBJID = sprintf("PILOT-%04d", 2001:2009)))

ae_terms <- c("ARTHRALGIA", "CELLULITIS", "ERYTHEMA", "LOCALISED INFECTION", "MICTURITION URGENCY")
ae_soc <- c("MUSCULOSKELETAL AND CONNECTIVE TISSUE DISORDERS", "INFECTIONS AND INFESTATIONS", "SKIN AND SUBCUTANEOUS TISSUE DISORDERS", "INFECTIONS AND INFESTATIONS", "RENAL AND URINARY DISORDERS")

adae <- do.call(rbind, lapply(seq_len(nrow(adsl)), function(i) {
  n_events <- sample(1:3, 1)
  event_index <- sample(seq_along(ae_terms), n_events)
  data.frame(
    USUBJID = adsl$USUBJID[i],
    AETERM = ae_terms[event_index],
    AEDECOD = ae_terms[event_index],
    AEBODSYS = ae_soc[event_index],
    AESEV = sample(c("MILD", "MODERATE", "SEVERE"), n_events, replace = TRUE),
    AESER = "N",
    AEACN = "",
    AEREL = "NONE",
    AEOUT = sample(c("RECOVERED/RESOLVED", "NOT RECOVERED/NOT RESOLVED"), n_events, replace = TRUE),
    AESTDTC = as.character(as.Date("2012-09-01") + sample(-70:40, n_events)),
    AEENDTC = "",
    AESTDY = sample(-70:40, n_events),
    AEENDY = NA,
    TRT01A = adsl$TRT01A[i],
    TRT01AN = adsl$TRT01AN[i],
    stringsAsFactors = FALSE
  )
}))

stopifnot(nrow(adsl) == 2L * nrow(original_adsl))
stopifnot(all(names(adsl) == toupper(names(adsl))), all(names(adae) == toupper(names(adae))))

# Locate LibreOffice without requiring `soffice` to be on PATH.
find_soffice <- function() {
  candidates <- c(
    Sys.which("soffice"),
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/Applications/OpenOffice.app/Contents/MacOS/soffice",
    "/Users/dianhe/.cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override/soffice"
  )
  candidates <- candidates[nzchar(candidates) & file.exists(candidates)]
  if (!length(candidates)) stop("LibreOffice was not found. Install LibreOffice or add soffice to PATH.")
  candidates[1]
}

# Create a standards-compliant XLSX workbook via LibreOffice.
write_xlsx <- function(data, path) {
  csv_path <- tempfile(fileext = ".csv")
  write.csv(data, csv_path, row.names = FALSE, na = "", fileEncoding = "UTF-8")
  converted_path <- file.path(tempdir(), sub("\\.csv$", ".xlsx", basename(csv_path)))
  if (file.exists(converted_path)) unlink(converted_path)

  status <- system2(find_soffice(), c("--headless", "--convert-to", "xlsx", "--outdir", shQuote(tempdir()), shQuote(csv_path)))
  if (status != 0 || !file.exists(converted_path)) stop("Unable to create Excel workbook.")
  if (file.exists(path)) unlink(path)
  if (!file.rename(converted_path, path)) stop("Unable to save Excel workbook.")
  unlink(csv_path)
}

# Escape RTF control characters and write a bordered clinical-style listing.
escape_rtf <- function(x) {
  x <- gsub("\\\\", "\\\\\\\\", as.character(x))
  gsub("([{}])", "\\\\\\1", x)
}

write_rtf_listing <- function(data, title, path) {
  con <- file(path, "wb"); on.exit(close(con), add = TRUE)
  w <- function(...) writeLines(paste0(...), con, useBytes = TRUE)
  w("{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\landscape\\paperw15840\\paperh12240\\margl600\\margr600\\margt500\\margb600")
  w("\\pard\\f0\\fs15 PILOT\\tab Confidential\\tab Page {\\field{\\*\\fldinst PAGE}} of {\\field{\\*\\fldinst NUMPAGES}}\\par")
  w("\\pard\\qc\\b\\fs19 ", escape_rtf(title), "\\b0\\par\\pard\\qc\\fs15 Safety Population\\par\\par")
  n <- ncol(data)
  # Use the full landscape text width (15,840 twips less 600-twip margins)
  # and allocate extra space to the columns containing longer clinical text.
  widths <- c(1400, 1100, 3000, 1800, 1100, 700, 1100, 2000, 1320, 1120)
  stopifnot(length(widths) == n, sum(widths) == 14640)
  ends <- cumsum(widths)
  for (i in 0:nrow(data)) {
    values <- if (i == 0) names(data) else as.character(data[i, ])
    border <- if (i == 0) "\\clbrdrt\\brdrs\\brdrw15\\clbrdrb\\brdrs\\brdrw15" else if (i == nrow(data)) "\\clbrdrb\\brdrs\\brdrw15" else ""
    w("\\trowd\\trleft0\\trgaph60", if (i == 0) "\\trhdr" else "", paste0(border, "\\cellx", ends, collapse = ""))
    for (j in seq_len(n)) w("\\intbl", if (j <= 4) "\\ql" else "\\qc", if (i == 0) "\\b" else "\\b0", "\\fs14 ", escape_rtf(values[j]), "\\cell")
    w("\\row")
  }
  w("\\pard\\ql\\fs12\\par Source: simulated data generated by this standalone program.\\par}")
}

# Assemble the presentation listing and create all deliverables.
listing <- adae[order(adae$USUBJID, adae$AESTDY), c("USUBJID", "TRT01A", "AEBODSYS", "AEDECOD", "AESEV", "AESER", "AEREL", "AEOUT", "AESTDTC", "AEENDTC")]
names(listing) <- c("Subject", "Treatment", "System Organ Class", "Preferred Term", "Severity", "Serious", "Relationship", "Outcome", "Start Date", "End Date")

write_xlsx(adsl, file.path(output_dir, paste0(output_stem, "_simulated_data.xlsx")))
write_xlsx(adae, file.path(output_dir, paste0(output_stem, "_simulated_adae_data.xlsx")))
write_rtf_listing(listing, "Listing 1. Adverse Events", file.path(output_dir, paste0(output_stem, ".rtf")))

message("Created ", output_stem, ": original N=9; simulated N=18")
