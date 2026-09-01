# =============================================================================
# PILOT clinical figure: Forest Plot of Subgroup Analysis
#
# Standalone program: embeds the original source data, simulates a dataset with
# twice the original number of records/subjects, performs the same analysis,
# and creates XLSX, PNG, and clinical-format RTF outputs.
# =============================================================================

options(stringsAsFactors = FALSE)

if (!requireNamespace("ggplot2", quietly = TRUE)) {
  stop("The 'ggplot2' package is required.")
}
if (!requireNamespace("zip", quietly = TRUE)) {
  stop("The 'zip' package is required to create the Excel workbook.")
}

get_script_directory <- function() {
  args <- commandArgs(trailingOnly = FALSE)
  file_arg <- grep("^--file=", args, value = TRUE)
  if (length(file_arg) == 1L) {
    return(dirname(normalizePath(sub("^--file=", "", file_arg))))
  }
  normalizePath(getwd())
}

output_directory <- get_script_directory()
stem <- "02_forest_subgroup_analysis"
set.seed(2026082902)

# -----------------------------------------------------------------------------
# Original data copied from the corresponding DOCX code segment
# -----------------------------------------------------------------------------
original_csv <- paste(c(
  "PARAMCD,SUBGROUP,TRT1,TRT2,YVAL,HR,LCL,UCL,SUBGRPGR1,INDENT",
  "ORR,Age Group,,,0,NA,NA,NA,,0",
  "ORR,< 65 Years,35/50 (70%),30/50 (60%),1,0.75,0.5,1.12,Age Grou,2",
  "ORR,= 65 Years,25/40 (62.5%),20/40 (50%),2,0.68,0.42,1.1,Age Grou,2",
  "ORR,Sex,,,3,NA,NA,NA,,0",
  "ORR,Male,30/45 (66.7%),25/45 (55.6%),4,0.7,0.45,1.09,Sex,2",
  "ORR,Female,30/45 (66.7%),25/45 (55.6%),5,0.78,0.49,1.2,Sex,2",
  "ORR,Disease Severity,,,6,NA,NA,NA,,0",
  "ORR,Mild,28/40 (70%),20/40 (50%),7,0.6,0.38,0.95,Disease,2",
  "ORR,Severe Disease,32/50 (64%),30/50 (60%),8,0.82,0.54,1.25,Disease,2"
), collapse = "\n")
original_data <- read.csv(
  text = original_csv,
  na.strings = c("NA", ""),
  check.names = FALSE,
  stringsAsFactors = FALSE
)
names(original_data) <- toupper(names(original_data))

# -----------------------------------------------------------------------------
# Reproducible simulation with twice the original sample size
# -----------------------------------------------------------------------------
simulated_data <- rbind(original_data, original_data)
simulated_data$SIMREC <- seq_len(nrow(simulated_data))
numeric_rows <- !is.na(simulated_data$HR)
simulated_data$HR[numeric_rows] <- pmax(0.05, simulated_data$HR[numeric_rows] * exp(rnorm(sum(numeric_rows), 0, 0.025)))
simulated_data$LCL[numeric_rows] <- pmax(0.01, simulated_data$LCL[numeric_rows] * exp(rnorm(sum(numeric_rows), 0, 0.02)))
simulated_data$UCL[numeric_rows] <- simulated_data$UCL[numeric_rows] * exp(rnorm(sum(numeric_rows), 0, 0.02))

names(simulated_data) <- toupper(names(simulated_data))
stopifnot(
  nrow(simulated_data) == 2L * nrow(original_data),
  all(names(simulated_data) == toupper(names(simulated_data)))
)

# -----------------------------------------------------------------------------
# Minimal standards-compliant XLSX writer (no LibreOffice, Java, or Python)
# -----------------------------------------------------------------------------
xml_escape <- function(x) {
  x <- gsub("&", "&amp;", as.character(x), fixed = TRUE)
  x <- gsub("<", "&lt;", x, fixed = TRUE)
  x <- gsub(">", "&gt;", x, fixed = TRUE)
  x <- gsub('"', "&quot;", x, fixed = TRUE)
  x
}

excel_column <- function(index) {
  value <- character()
  while (index > 0L) {
    remainder <- (index - 1L) %% 26L
    value <- c(LETTERS[remainder + 1L], value)
    index <- (index - 1L) %/% 26L
  }
  paste(value, collapse = "")
}

write_xlsx <- function(data, path) {
  workbook_dir <- tempfile("pilot_xlsx_")
  dir.create(file.path(workbook_dir, "_rels"), recursive = TRUE)
  dir.create(file.path(workbook_dir, "docProps"), recursive = TRUE)
  dir.create(file.path(workbook_dir, "xl", "_rels"), recursive = TRUE)
  dir.create(file.path(workbook_dir, "xl", "worksheets"), recursive = TRUE)
  on.exit(unlink(workbook_dir, recursive = TRUE, force = TRUE), add = TRUE)

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>',
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>',
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>',
    '</Types>'
  ), file.path(workbook_dir, "[Content_Types].xml"), useBytes = TRUE)

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>',
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>',
    '</Relationships>'
  ), file.path(workbook_dir, "_rels", ".rels"), useBytes = TRUE)

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">',
    '<sheets><sheet name="SIMULATED_DATA" sheetId="1" r:id="rId1"/></sheets>',
    '</workbook>'
  ), file.path(workbook_dir, "xl", "workbook.xml"), useBytes = TRUE)

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>',
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>',
    '</Relationships>'
  ), file.path(workbook_dir, "xl", "_rels", "workbook.xml.rels"), useBytes = TRUE)

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    '<fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Arial"/></font></fonts>',
    '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills>',
    '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>',
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>',
    '<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/></cellXfs>',
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>',
    '</styleSheet>'
  ), file.path(workbook_dir, "xl", "styles.xml"), useBytes = TRUE)

  cell_xml <- function(value, reference, header = FALSE) {
    style <- if (header) ' s="1"' else ""
    if ((is.numeric(value) || is.integer(value)) && !is.na(value)) {
      sprintf('<c r="%s"%s><v>%s</v></c>', reference, style, value)
    } else {
      text_value <- if (is.na(value)) "" else xml_escape(value)
      sprintf('<c r="%s" t="inlineStr"%s><is><t>%s</t></is></c>', reference, style, text_value)
    }
  }

  rows <- character(nrow(data) + 1L)
  header_cells <- vapply(seq_along(data), function(column) {
    cell_xml(names(data)[column], paste0(excel_column(column), "1"), TRUE)
  }, character(1))
  rows[1] <- paste0('<row r="1">', paste(header_cells, collapse = ""), '</row>')

  for (row in seq_len(nrow(data))) {
    cells <- vapply(seq_along(data), function(column) {
      cell_xml(data[row, column], paste0(excel_column(column), row + 1L))
    }, character(1))
    rows[row + 1L] <- paste0('<row r="', row + 1L, '">', paste(cells, collapse = ""), '</row>')
  }

  worksheet <- c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
    paste0('<dimension ref="A1:', excel_column(ncol(data)), nrow(data) + 1L, '"/>'),
    '<sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>',
    '<sheetFormatPr defaultRowHeight="15"/>',
    paste0('<cols><col min="1" max="', ncol(data), '" width="18" customWidth="1"/></cols>'),
    '<sheetData>', rows, '</sheetData>',
    paste0('<autoFilter ref="A1:', excel_column(ncol(data)), nrow(data) + 1L, '"/>'),
    '</worksheet>'
  )
  writeLines(worksheet, file.path(workbook_dir, "xl", "worksheets", "sheet1.xml"), useBytes = TRUE)

  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>PILOT Simulated Data</dc:title><dc:creator>PILOT</dc:creator></cp:coreProperties>'
  ), file.path(workbook_dir, "docProps", "core.xml"), useBytes = TRUE)
  writeLines(c(
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
    '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>R</Application></Properties>'
  ), file.path(workbook_dir, "docProps", "app.xml"), useBytes = TRUE)

  if (file.exists(path)) unlink(path)
  zip::zipr(path, files = ".", root = workbook_dir, include_directories = FALSE)
  if (!file.exists(path) || file.info(path)$size == 0L) stop("Unable to create Excel workbook: ", path)
}

excel_path <- file.path(output_directory, paste0(stem, "_simulated_data.xlsx"))
write_xlsx(simulated_data, excel_path)

# -----------------------------------------------------------------------------
# Analysis and figure creation, retaining the original ggplot logic
# -----------------------------------------------------------------------------
double_count_text <- function(x) {
  if (is.na(x) || x == "") return(x)
  matched <- regmatches(x, regexec("^([0-9]+)/([0-9]+)", x))[[1]]
  if (length(matched) != 3L) return(x)
  replacement <- paste0(as.integer(matched[2]) * 2L, "/", as.integer(matched[3]) * 2L)
  sub("^[0-9]+/[0-9]+", replacement, x)
}

plot_data <- original_data
plot_data$SUBGROUP <- gsub("PILOT", "PILOT", plot_data$SUBGROUP, fixed = TRUE)
plot_data$TRT1 <- vapply(plot_data$TRT1, double_count_text, character(1))
plot_data$TRT2 <- vapply(plot_data$TRT2, double_count_text, character(1))

finalplot <- ggplot2::ggplot(plot_data, ggplot2::aes(x = HR, y = YVAL)) +
  ggplot2::geom_vline(xintercept = 1, linetype = "dashed", color = "grey50") +
  ggplot2::geom_errorbar(ggplot2::aes(xmin = LCL, xmax = UCL), orientation = "y", color = "#005A9C", width = 0) +
  ggplot2::geom_point(color = "#0072B2", size = 2.5, na.rm = TRUE) +
  ggplot2::geom_text(ggplot2::aes(x = 0.02, label = SUBGROUP), hjust = 0, size = 3) +
  ggplot2::geom_text(ggplot2::aes(x = 1.62, label = TRT1), hjust = 0, size = 3, na.rm = TRUE) +
  ggplot2::geom_text(ggplot2::aes(x = 2.08, label = TRT2), hjust = 0, size = 3, na.rm = TRUE) +
  ggplot2::annotate("text", x = 0.02, y = -1, label = "Subgroup", hjust = 0, fontface = "bold", size = 3) +
  ggplot2::annotate("text", x = 1.62, y = -1, label = "PILOT\nn/m (%)", hjust = 0, fontface = "bold", size = 3) +
  ggplot2::annotate("text", x = 2.08, y = -1, label = "Placebo\nn/m (%)", hjust = 0, fontface = "bold", size = 3) +
  ggplot2::scale_x_continuous("Hazard Ratio", breaks = seq(0.4, 1.4, 0.2), limits = c(0, 2.4), expand = c(0, 0)) +
  ggplot2::scale_y_reverse(labels = NULL, limits = c(max(plot_data$YVAL, na.rm = TRUE) + 1, -1.5), expand = c(0, 0)) +
  ggplot2::labs(title = "Subgroup Analysis", subtitle = "Analysis Population") +
  ggplot2::theme_minimal(base_size = 11) +
  ggplot2::theme(panel.grid = ggplot2::element_blank(), axis.title.y = ggplot2::element_blank(), axis.ticks.y = ggplot2::element_blank(), plot.title = ggplot2::element_text(face = "bold"))

png_path <- file.path(output_directory, paste0(stem, ".png"))
ggplot2::ggsave(
  filename = png_path,
  plot = finalplot,
  width = 10.5,
  height = 6.3,
  units = "in",
  dpi = 220,
  bg = "white"
)

# -----------------------------------------------------------------------------
# Create a landscape, single-page clinical RTF with the PNG embedded
# -----------------------------------------------------------------------------
rtf_escape <- function(x) {
  x <- gsub("\\\\", "\\\\\\\\", x)
  x <- gsub("\\{", "\\\\{", x)
  x <- gsub("\\}", "\\\\}", x)
  x
}

write_figure_rtf <- function(image_path, rtf_path, title, figure_number) {
  image_raw <- readBin(image_path, what = "raw", n = file.info(image_path)$size)
  image_hex <- paste(sprintf("%02x", as.integer(image_raw)), collapse = "")
  starts <- seq(1L, nchar(image_hex), by = 120L)
  image_lines <- substring(image_hex, starts, pmin(starts + 119L, nchar(image_hex)))

  rtf <- c(
    "{\\rtf1\\ansi\\deff0",
    "{\\fonttbl{\\f0 Arial;}}",
    "\\landscape\\paperw15840\\paperh12240\\margl720\\margr720\\margt540\\margb540",
    "{\\header\\pard\\ql\\fs16 PILOT\\tab Confidential\\par}",
    "{\\footer\\pard\\ql\\fs16 Source: Simulated data\\tab\\qr Page {\\field{\\*\\fldinst PAGE}} of {\\field{\\*\\fldinst NUMPAGES}}\\par}",
    paste0("\\pard\\qc\\b\\fs22 Figure ", figure_number, "\\par"),
    paste0("\\pard\\qc\\b\\fs22 ", rtf_escape(title), "\\par"),
    "\\pard\\qc\\b0\\fs18 Analysis Population\\par\\par",
    "{\\pict\\pngblip\\picwgoal10880\\pichgoal6528",
    image_lines,
    "}",
    "\\par\\pard\\ql\\fs16 Note: Results are based on reproducibly simulated data with twice the original sample size.\\par",
    "}"
  )
  writeLines(rtf, rtf_path, useBytes = TRUE)
}

rtf_path <- file.path(output_directory, paste0(stem, ".rtf"))
write_figure_rtf(png_path, rtf_path, "Forest Plot of Subgroup Analysis", "14.1.2")

message("Created: ", excel_path)
message("Created: ", png_path)
message("Created: ", rtf_path)
