# Listing 1. Screen Failures
# Standalone program: embeds the original data, simulates 2 x N, and writes XLSX/RTF.

set.seed(20260830)
args <- grep("^--file=", commandArgs(), value = TRUE)
out_dir <- if (length(args)) dirname(normalizePath(sub("^--file=", "", args[1]))) else getwd()
stem <- "01_screen_failures_listing"

# Original source data transcribed from Listings- Disposition.docx.
original <- data.frame(
  USUBJID = sprintf("PILOT-%04d", 1001:1014),
  TRT01PN = NA_integer_,
  SITEID = c("7103","7104","7001","7001","7002","2014","4003","4006","4007","5907","8011","9037","0004","0006"),
  SITENAM = paste("Site", c("7103","7104","7001","7001","7002","2014","4003","4006","4007","5907","8011","9037","0004","0006")),
  AGE = c(58,73,41,38,62,82,65,79,55,55,73,75,76,58), AGEU = "YEARS",
  SEX = c("F","M","F","F","M","M","M","F","M","F","F","M","M","M"),
  ARACE = c("OTHER","OTHER","BLACK OR AFRICAN AMERICAN","BLACK OR AFRICAN AMERICAN","ASIAN",rep("WHITE",9)),
  DSFFL = "Y",
  DSFREAS = c("SCREEN FAILURE","SCREEN FAILURE","LOST TO FOLLOW-UP","SCREEN FAILURE","OTHER","WITHDRAWAL BY SUBJECT","OTHER","WITHDRAWAL BY SUBJECT","WITHDRAWAL BY SUBJECT","ADVERSE EVENT","OTHER","ADVERSE EVENT","ADVERSE EVENT","LOST TO FOLLOW-UP"),
  stringsAsFactors = FALSE
)

# Sample source records with replacement and apply small, controlled age variation.
n_original <- nrow(original)
simulated <- original[sample(seq_len(n_original), 2L * n_original, replace = TRUE), ]
simulated$USUBJID <- sprintf("PILOT-%04d", 2001:(2000 + nrow(simulated)))
simulated$AGE <- pmax(18L, simulated$AGE + sample(-3:3, nrow(simulated), TRUE))
rownames(simulated) <- NULL

# Minimal standards-compliant XLSX writer; no LibreOffice or add-on R package required.
xml_escape <- function(x) {
  x <- ifelse(is.na(x), "", as.character(x))
  x <- gsub("&", "&amp;", x, fixed = TRUE)
  x <- gsub("<", "&lt;", x, fixed = TRUE)
  gsub(">", "&gt;", x, fixed = TRUE)
}
excel_col <- function(n) {
  out <- character(n)
  for (i in seq_len(n)) {
    k <- i; s <- ""
    while (k > 0) { k <- k - 1L; s <- paste0(LETTERS[k %% 26L + 1L], s); k <- k %/% 26L }
    out[i] <- s
  }
  out
}
write_xlsx <- function(data, path) {
  td <- tempfile("xlsx_"); dir.create(td); dir.create(file.path(td, "_rels"));
  dir.create(file.path(td, "xl")); dir.create(file.path(td, "xl", "_rels")); dir.create(file.path(td, "xl", "worksheets"))
  writeLines('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>', file.path(td,"[Content_Types].xml"))
  writeLines('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>', file.path(td,"_rels",".rels"))
  writeLines('<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="SIMULATED_DATA" sheetId="1" r:id="rId1"/></sheets></workbook>', file.path(td,"xl","workbook.xml"))
  writeLines('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>', file.path(td,"xl","_rels","workbook.xml.rels"))
  writeLines('<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf fontId="0" fillId="0" borderId="0"/><xf fontId="1" fillId="2" borderId="0" applyFont="1" applyFill="1"/></cellXfs></styleSheet>', file.path(td,"xl","styles.xml"))
  all <- rbind(names(data), lapply(data, as.character) |> as.data.frame(stringsAsFactors=FALSE)); names(all) <- names(data)
  rows <- vapply(seq_len(nrow(all)), function(i) paste0('<row r="',i,'">',paste0('<c r="',excel_col(ncol(all)),i,'" t="inlineStr" s="',if(i==1)1 else 0,'"><is><t>',xml_escape(unlist(all[i,],use.names=FALSE)),'</t></is></c>',collapse=''),'</row>'), "")
  sheet <- paste0('<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData>',paste(rows,collapse=''),'</sheetData><autoFilter ref="A1:',excel_col(ncol(data))[ncol(data)],nrow(data)+1L,'"/></worksheet>')
  writeLines(sheet, file.path(td,"xl","worksheets","sheet1.xml"), useBytes=TRUE)
  old <- setwd(td); on.exit(setwd(old), add=TRUE); if (file.exists(path)) unlink(path)
  utils::zip(path, list.files(".", recursive=TRUE, all.files=TRUE, no..=TRUE), flags="-q")
}

# Clinical-style RTF listing, grouped by investigator site.
rtf_escape <- function(x) { x <- gsub("\\\\", "\\\\\\\\", as.character(x)); gsub("([{}])", "\\\\\\1", x) }
write_rtf <- function(data, path) {
  con <- file(path, "wb"); on.exit(close(con), add=TRUE); w <- function(x) writeLines(x, con, useBytes=TRUE)
  w('{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\landscape\\paperw15840\\paperh12240\\margl600\\margr600\\margt450\\margb500')
  w('\\pard\\fs15 PILOT\\tab Confidential\\tab Page {\\field{\\*\\fldinst PAGE}} of {\\field{\\*\\fldinst NUMPAGES}}\\par')
  w('\\pard\\qc\\b\\fs20 Listing 1. Screen Failures\\b0\\par\\pard\\qc\\fs15 All Screen Failure Subjects\\par\\par')
  labels <- c("Subject ID","Age (Years)","Sex","Race","Reason for Discontinuation"); widths <- c(2200,1600,1000,3600,6240); ends <- cumsum(widths)
  header <- paste0('\\trowd\\trhdr\\trgaph60',paste0('\\clbrdrt\\brdrs\\brdrw15\\clbrdrb\\brdrs\\brdrw15\\cellx',ends,collapse=''),paste0('\\intbl\\ql\\b\\fs15 ',rtf_escape(labels),'\\cell',collapse=''),'\\row'); w(header)
  sites <- unique(data$SITEID)
  for (site in sites) {
    d <- data[data$SITEID == site, ]; w(paste0('\\trowd\\trgaph60\\cellx14640\\intbl\\ql\\b\\fs15 Investigator Site = ',rtf_escape(site),', ',rtf_escape(d$SITENAM[1]),'\\b0\\cell\\row'))
    for (i in seq_len(nrow(d))) { v <- c(d$USUBJID[i],d$AGE[i],d$SEX[i],d$ARACE[i],d$DSFREAS[i]); w(paste0('\\trowd\\trgaph60',paste0('\\cellx',ends,collapse=''),paste0('\\intbl\\ql\\fs15 ',rtf_escape(v),'\\cell',collapse=''),'\\row')) }
  }
  w('\\trowd\\trgaph60\\clbrdrb\\brdrs\\brdrw15\\cellx14640\\intbl\\cell\\row\\pard\\fs12\\par Source: simulated data generated from the original DOCX data logic.\\par}')
}

simulated <- simulated[order(simulated$SITEID, simulated$USUBJID), ]
write_xlsx(simulated, file.path(out_dir, paste0(stem, "_simulated_data.xlsx")))
write_rtf(simulated, file.path(out_dir, paste0(stem, ".rtf")))
stopifnot(nrow(simulated) == 2L*n_original, all(names(simulated) == toupper(names(simulated))))
message("Created ", stem, ": original N=", n_original, "; simulated N=", nrow(simulated))
