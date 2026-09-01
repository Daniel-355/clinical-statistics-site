#!/usr/bin/env Rscript

# PILOT - Demographics and Baseline Characteristics
# Standalone program: embedded original data -> doubled simulation -> XLSX -> analysis -> RTF

options(stringsAsFactors = FALSE)
set.seed(20260827)
suppressPackageStartupMessages(library(tidyverse))

script_dir <- function() {
  a <- commandArgs(trailingOnly = FALSE)
  f <- sub("^--file=", "", a[grepl("^--file=", a)])
  f <- gsub("~\\+~", " ", f)
  if (length(f)) dirname(normalizePath(f[1])) else getwd()
}
out_dir <- script_dir()
stem <- "01_demographics_baseline_summary"

# Original 41-subject ADSL data from the source DOCX.
original_data <- tribble(
  ~USUBJID,~TRT01AN,~AGE,~SEX,~SEXN,~ARACE,~ARACEN,~WEIGHTBL,~HEIGHTBL,~BMIBL,~SAFFL,
  "DBAA100001",2,68,"M",1,"White",1,100.7,186,29.11,"Y",
  "DBAA300004",1,64,"M",1,"White",1,91.2,172,30.83,"Y",
  "DBAA400001",2,72,"M",1,"White",1,105.8,178,33.39,"Y",
  "DEAA300001",2,71,"M",1,"White",1,77.9,162,29.68,"Y",
  "DEAA300003",1,79,"M",1,"White",1,87.8,164,32.64,"Y",
  "DEAA400001",2,61,"M",1,"White",1,92.7,174,30.62,"Y",
  "DEAA400002",2,55,"F",2,"White",1,72.3,155,30.09,"Y",
  "DEAA800002",1,58,"M",1,"White",1,92,176,29.7,"Y",
  "DEAB200003",1,78,"M",1,"White",1,80.4,164,29.89,"Y",
  "DGAA300001",2,75,"F",2,"White",1,81.5,152,35.28,"Y",
  "DGAA300002",1,56,"F",2,"White",1,53.5,146,25.1,"Y",
  "DGAA300006",2,66,"M",1,"White",1,109.1,178,34.43,"Y",
  "DGAA300012",2,70,"F",2,"White",1,101.2,166,36.73,"Y",
  "DGAA300013",1,66,"M",1,"White",1,116.8,167,41.88,"Y",
  "DGAA300014",2,70,"F",2,"White",1,92.5,157,37.53,"Y",
  "DGAB200001",1,70,"F",2,"White",1,66,165,24.24,"Y",
  "DGAB200003",1,75,"M",1,"White",1,74,173,24.73,"Y",
  "DGAB200010",2,66,"M",1,"White",1,118,176,38.09,"Y",
  "ECAA200001",2,59,"F",2,"White",1,87.6,168,31.04,"Y",
  "ECAA200002",1,64,"M",1,"White",1,102.7,182,31,"Y",
  "ECAA200003",2,70,"M",1,"White",1,81,171,27.7,"Y",
  "ECAA200006",1,74,"M",1,"White",1,105,183,31.35,"Y",
  "ECAA500004",1,84,"F",2,"White",1,67,168,23.74,"Y",
  "ECAA500005",2,74,"M",1,"White",1,93.7,178,29.57,"Y",
  "EEAA100002",1,73,"F",2,"White",1,132.9,153,56.77,"Y",
  "EEAA300001",2,76,"M",1,"White",1,96,183,28.67,"Y",
  "EEAA400001",2,60,"M",1,"Asian",3,72.5,168,25.69,"Y",
  "EFAA100002",2,70,"F",2,"White",1,75.5,162,28.77,"Y",
  "EFAA100003",2,71,"M",1,"White",1,96,169,33.61,"Y",
  "EFAA100004",2,66,"M",1,"Asian",3,86,167,30.84,"Y",
  "EFAA100005",2,73,"M",1,"White",1,83.2,178,26.26,"Y",
  "EFAA200001",1,66,"M",1,"White",1,124.2,178,39.2,"Y",
  "EFAA200003",2,80,"M",1,"White",1,80.1,170,27.72,"Y",
  "EFAA400001",1,69,"M",1,"White",1,141.4,180,43.64,"Y",
  "EFAA400002",1,73,"M",1,"White",1,74.6,176,24.08,"Y",
  "EFAA400003",1,71,"M",1,"White",1,116.6,181,35.59,"Y",
  "EFAA700002",2,75,"M",1,"White",1,75.3,163,28.34,"Y",
  "EIAA200005",1,61,"F",2,"White",1,82.5,155,34.34,"Y",
  "EJAA300002",1,62,"F",2,"White",1,96,157,38.95,"Y",
  "EJAA300004",1,74,"M",1,"White",1,90.8,174,29.99,"Y",
  "EJAA400009",2,72,"M",1,"White",1,91.6,177,29.24,"Y"
)

# Simulate 82 new subjects while preserving treatment, sex, and race structure.
# Each source subject contributes two new subjects; continuous variables receive
# controlled perturbations and BMI is recalculated from weight and height.
simulated_data <- original_data[rep(seq_len(nrow(original_data)), each = 2), ]
simulated_data$USUBJID <- sprintf("PILOT-%04d", seq_len(nrow(simulated_data)))
simulated_data$AGE <- pmax(18L, round(simulated_data$AGE + rnorm(nrow(simulated_data), 0, 2.5)))
simulated_data$HEIGHTBL <- round(pmax(135, simulated_data$HEIGHTBL + rnorm(nrow(simulated_data), 0, 1.5)))
simulated_data$WEIGHTBL <- round(pmax(40, simulated_data$WEIGHTBL + rnorm(nrow(simulated_data), 0, 3.5)), 1)
simulated_data$BMIBL <- round(simulated_data$WEIGHTBL / (simulated_data$HEIGHTBL / 100)^2, 2)

stopifnot(
  nrow(original_data) == 41L,
  nrow(simulated_data) == 82L,
  all(names(simulated_data) == toupper(names(simulated_data))),
  max(abs(simulated_data$BMIBL - simulated_data$WEIGHTBL / (simulated_data$HEIGHTBL / 100)^2)) < 0.006
)

# Minimal standalone XLSX writer (no external Excel package required).
xml_escape <- function(x) {
  x[is.na(x)] <- ""
  x <- gsub("&", "&amp;", x, fixed = TRUE)
  x <- gsub("<", "&lt;", x, fixed = TRUE)
  x <- gsub(">", "&gt;", x, fixed = TRUE)
  gsub('"', "&quot;", x, fixed = TRUE)
}
excel_col <- function(n) {
  out <- character(length(n))
  for (j in seq_along(n)) {
    k <- n[j]; s <- ""
    while (k > 0) { k <- k - 1; s <- paste0(intToUtf8(65 + k %% 26), s); k <- k %/% 26 }
    out[j] <- s
  }
  out
}
write_xlsx_base <- function(dat, path) {
  td <- tempfile("xlsx_"); dir.create(td)
  dirs <- c("_rels", "docProps", "xl", "xl/_rels", "xl/worksheets")
  invisible(vapply(file.path(td, dirs), dir.create, logical(1), recursive = TRUE))
  writeLines('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>', file.path(td, "[Content_Types].xml"))
  writeLines('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>', file.path(td, "_rels/.rels"))
  writeLines('<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="SIMULATED_DATA" sheetId="1" r:id="rId1"/></sheets></workbook>', file.path(td, "xl/workbook.xml"))
  writeLines('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>', file.path(td, "xl/_rels/workbook.xml.rels"))
  writeLines('<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="10"/><name val="Arial"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf xfId="0"/></cellXfs></styleSheet>', file.path(td, "xl/styles.xml"))
  all_rows <- rbind(names(dat), as.matrix(dat))
  rows <- vapply(seq_len(nrow(all_rows)), function(i) {
    cells <- vapply(seq_len(ncol(all_rows)), function(j) sprintf('<c r="%s%d" t="inlineStr"><is><t>%s</t></is></c>', excel_col(j), i, xml_escape(as.character(all_rows[i,j]))), character(1))
    sprintf('<row r="%d">%s</row>', i, paste0(cells, collapse = ""))
  }, character(1))
  writeLines(paste0('<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>', paste0(rows, collapse = ""), '</sheetData></worksheet>'), file.path(td, "xl/worksheets/sheet1.xml"))
  writeLines('<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>PILOT</dc:creator></cp:coreProperties>', file.path(td, "docProps/core.xml"))
  writeLines('<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>R</Application></Properties>', file.path(td, "docProps/app.xml"))
  old <- setwd(td); on.exit(setwd(old), add = TRUE)
  if (file.exists(path)) unlink(path)
  utils::zip(path, list.files(".", recursive = TRUE, all.files = TRUE, no.. = TRUE), flags = "-q")
}

write_xlsx_base(simulated_data, file.path(out_dir, paste0(stem, "_simulated_data.xlsx")))

# Original analysis structure retained: Safety Set, treatment rows plus Total,
# numeric summaries, categorical counts/percentages, and treatment columns.
adsl01 <- simulated_data %>% filter(SAFFL == "Y")
adsl02 <- adsl01 %>% mutate(TREATMENT = TRT01AN) %>%
  bind_rows(adsl01 %>% mutate(TREATMENT = 4L))
trt_levels <- 1:4
trt_n <- vapply(trt_levels, function(g) sum(adsl02$TREATMENT == g), integer(1))

fmt_num <- function(x, dp) {
  x <- x[!is.na(x)]
  if (!length(x)) return(c("0 (0)", "-", "-", "-", "-"))
  f0 <- paste0("%.", dp, "f"); f1 <- paste0("%.", dp + 1, "f"); f2 <- paste0("%.", dp + 2, "f")
  c(
    sprintf("%d (%d)", length(x), 0L),
    sprintf("%s (%s)", sprintf(f1, mean(x)), sprintf(f2, sd(x))),
    sprintf(f1, median(x)),
    sprintf("%s, %s", sprintf(f1, quantile(x, .25, type = 2)), sprintf(f1, quantile(x, .75, type = 2))),
    sprintf("%s, %s", sprintf(f0, min(x)), sprintf(f0, max(x)))
  )
}
numeric_specs <- tribble(
  ~GROUPLABEL, ~VAR, ~DP,
  "Age (years)", "AGE", 0,
  "Weight (kg)", "WEIGHTBL", 1,
  "Height (cm)", "HEIGHTBL", 0,
  "BMI (kg/m^2)", "BMIBL", 2
)
stat_labels <- c("n (missing)", "Mean (SD)", "Median", "Q1, Q3", "Min, Max")
numeric_rows <- list()
for (i in seq_len(nrow(numeric_specs))) {
  block <- sapply(trt_levels, function(g) fmt_num(adsl02[adsl02$TREATMENT == g, ][[numeric_specs$VAR[i]]], numeric_specs$DP[i]))
  numeric_rows[[i]] <- cbind(LABEL = c(numeric_specs$GROUPLABEL[i], rep("", 4)), STATISTIC = stat_labels, block)
}

fmt_cat <- function(n, d) if (n == 0 || d == 0) "0" else sprintf("%d (%.1f%%)", n, 100*n/d)
cat_specs <- tribble(
  ~GROUPLABEL, ~STATISTIC, ~VAR, ~VALUE,
  "Sex, n (%)", "Male", "SEX", "M",
  "", "Female", "SEX", "F",
  "", "Missing", "SEX", NA_character_,
  "Race, n (%)", "White", "ARACE", "White",
  "", "Black or African American", "ARACE", "Black or African American",
  "", "Asian", "ARACE", "Asian",
  "", "Other", "ARACE", "Other",
  "", "Missing", "ARACE", NA_character_
)
cat_rows <- t(apply(cat_specs, 1, function(s) {
  vals <- vapply(trt_levels, function(g) {
    z <- adsl02[adsl02$TREATMENT == g, ]
    n <- if (is.na(s[["VALUE"]])) sum(is.na(z[[s[["VAR"]]]]) | z[[s[["VAR"]]]] == "") else sum(z[[s[["VAR"]]]] == s[["VALUE"]], na.rm = TRUE)
    fmt_cat(n, nrow(z))
  }, character(1))
  c(s[["GROUPLABEL"]], s[["STATISTIC"]], vals)
}))

body <- rbind(do.call(rbind, numeric_rows), cat_rows)
headers <- c("Parameter", "Statistic", sprintf("Treatment 1\n(N=%d)", trt_n[1]), sprintf("Treatment 2\n(N=%d)", trt_n[2]), sprintf("Treatment 3\n(N=%d)", trt_n[3]), sprintf("Total\n(N=%d)", trt_n[4]))
table_data <- rbind(headers, body)

rtf_escape <- function(x) {
  x <- gsub("\\", "\\\\", x, fixed = TRUE)
  x <- gsub("{", "\\{", x, fixed = TRUE)
  x <- gsub("}", "\\}", x, fixed = TRUE)
  gsub("\n", "\\line ", x, fixed = TRUE)
}
write_clinical_rtf <- function(tab, path) {
  con <- file(path, "wb"); on.exit(close(con))
  w <- function(...) writeLines(paste0(...), con, useBytes = TRUE)
  w("{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\landscape\\paperw15840\\paperh12240\\margl600\\margr600\\margt500\\margb600")
  w("\\pard\\f0\\fs15 PILOT\\tab Confidential\\tab Page {\\field{\\*\\fldinst PAGE}} of {\\field{\\*\\fldinst NUMPAGES}}\\par")
  w("\\pard\\qc\\b\\fs20 Table 1. Demographics and Baseline Characteristics\\b0\\par")
  w("\\pard\\qc\\fs16 Safety Analysis Set\\par\\par")
  ends <- cumsum(c(3100, 2300, rep(2150, 4)))
  for (i in seq_len(nrow(tab))) {
    border <- if (i == 1) "\\clbrdrt\\brdrs\\brdrw15\\clbrdrb\\brdrs\\brdrw15" else if (i == nrow(tab)) "\\clbrdrb\\brdrs\\brdrw15" else ""
    w("\\trowd\\trgaph60", if (i == 1) "\\trhdr" else "", paste0(border, "\\cellx", ends, collapse = ""))
    for (j in seq_len(ncol(tab))) {
      align <- if (j <= 2) "\\ql" else "\\qc"
      bold <- if (i == 1 || (j == 1 && nzchar(tab[i,1]))) "\\b" else "\\b0"
      w("\\intbl", align, bold, "\\f0\\fs14 ", rtf_escape(as.character(tab[i,j])), "\\cell")
    }
    w("\\row")
  }
  w("\\pard\\ql\\fs13\\par Percentages are based on the number of subjects in the Safety Analysis Set for each treatment group.\\par")
  w("\\pard\\ql\\fs13 BMI = body mass index; N = number of subjects; n = number in category; Q1 = first quartile; Q3 = third quartile; SD = standard deviation.\\par")
  w("\\pard\\ql\\fs13 Source: Simulated data generated by this standalone program. Generated: ", format(Sys.Date(), "%d%b%Y"), "\\par}")
}

write_clinical_rtf(table_data, file.path(out_dir, paste0(stem, ".rtf")))
cat("Created", paste0(stem, ".rtf"), "and", paste0(stem, "_simulated_data.xlsx"), "with 82 subjects.\n")
