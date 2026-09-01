# Standalone time-to-event analysis generated from code segment 3.

suppressPackageStartupMessages(library(survival))
set.seed(20260828)
arg <- grep("^--file=", commandArgs(FALSE), value=TRUE)
script_raw <- if(length(arg)) sub("^--file=", "", arg[1]) else "."
script_path <- normalizePath(gsub("~\\+~", " ", script_raw))
out_dir <- if(file.info(script_path)$isdir) script_path else dirname(script_path)
stem <- tools::file_path_sans_ext(basename(script_path))

make_simulated <- function(adsl, adtte) {
  trts <- sort(unique(adsl$TRT01AN))
  pieces <- lapply(trts, function(trt) {
    src <- adtte[adtte$TRT01AN == trt, , drop=FALSE]
    n_new <- 2L * nrow(src)
    picked <- src[sample(seq_len(nrow(src)), n_new, replace=TRUE), , drop=FALSE]
    jitter <- exp(rnorm(n_new, 0, 0.12))
    picked$AVAL <- pmax(1L, as.integer(round(picked$AVAL * jitter)))
    event_rate <- mean(src$CNSR == 0)
    picked$CNSR <- as.integer(runif(n_new) >= event_rate)
    picked$USUBJID <- sprintf("PILOT-%04d", seq_len(n_new) + sum(vapply(pieces_before(trt, trts, adsl), nrow, integer(1))))
    picked
  })
  adtte_new <- do.call(rbind, pieces)
  rownames(adtte_new) <- NULL
  adsl_new <- unique(adtte_new[c("USUBJID", "TRT01AN", "SAFFL")])
  list(ADSL=adsl_new, ADTTE=adtte_new)
}
pieces_before <- function(trt, trts, adsl) {
  earlier <- trts[trts < trt]
  lapply(earlier, function(x) adsl[rep(which(adsl$TRT01AN == x), each=2), , drop=FALSE])
}

xml_escape <- function(x){x[is.na(x)]<-"";x<-gsub("&","&amp;",x,fixed=TRUE);x<-gsub("<","&lt;",x,fixed=TRUE);x<-gsub(">","&gt;",x,fixed=TRUE);gsub('"',"&quot;",x,fixed=TRUE)}
excel_col <- function(n){z<-character(length(n));for(j in seq_along(n)){k<-n[j];s<-"";while(k>0){k<-k-1;s<-paste0(intToUtf8(65+k%%26),s);k<-k%/%26};z[j]<-s};z}
write_xlsx <- function(sheets, path) {
  sheets <- lapply(sheets, function(d){names(d)<-toupper(names(d));d})
  sn <- substr(gsub("[^A-Za-z0-9_]","_",names(sheets)),1,31)
  td <- tempfile("xlsx_"); dir.create(td)
  for(z in c("_rels","docProps","xl","xl/_rels","xl/worksheets")) dir.create(file.path(td,z),recursive=TRUE,showWarnings=FALSE)
  types <- paste0('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',paste0('<Override PartName="/xl/worksheets/sheet',seq_along(sheets),'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',collapse=""),'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>')
  writeLines(types,file.path(td,"[Content_Types].xml"))
  writeLines('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>',file.path(td,"_rels/.rels"))
  sh <- paste0('<sheet name="',sn,'" sheetId="',seq_along(sn),'" r:id="rId',seq_along(sn),'"/>',collapse="")
  writeLines(paste0('<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>',sh,'</sheets></workbook>'),file.path(td,"xl/workbook.xml"))
  rels <- paste0('<Relationship Id="rId',seq_along(sn),'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet',seq_along(sn),'.xml"/>',collapse="")
  writeLines(paste0('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',rels,'<Relationship Id="rId',length(sn)+1,'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'),file.path(td,"xl/_rels/workbook.xml.rels"))
  writeLines('<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf xfId="0" fontId="0" fillId="0"/><xf xfId="0" fontId="1" fillId="1" applyFont="1" applyFill="1"/></cellXfs></styleSheet>',file.path(td,"xl/styles.xml"))
  for(k in seq_along(sheets)) {
    d<-sheets[[k]]; a<-rbind(names(d),as.matrix(d))
    rows<-vapply(seq_len(nrow(a)),function(i){cc<-vapply(seq_len(ncol(a)),function(j){v<-as.character(a[i,j]);if(i>1 && is.numeric(d[[j]]) && !is.na(d[i-1,j]))sprintf('<c r="%s%d"><v>%s</v></c>',excel_col(j),i,d[i-1,j])else sprintf('<c r="%s%d" t="inlineStr" s="%d"><is><t>%s</t></is></c>',excel_col(j),i,ifelse(i==1,1,0),xml_escape(v))},character(1));sprintf('<row r="%d">%s</row>',i,paste0(cc,collapse=""))},character(1))
    writeLines(paste0('<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData>',paste0(rows,collapse=""),'</sheetData></worksheet>'),file.path(td,"xl/worksheets",paste0("sheet",k,".xml")))
  }
  writeLines('<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>PILOT</dc:creator></cp:coreProperties>',file.path(td,"docProps/core.xml"));writeLines('<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>R</Application></Properties>',file.path(td,"docProps/app.xml"))
  old<-setwd(td);on.exit(setwd(old),add=TRUE);if(file.exists(path))unlink(path);utils::zip(path,list.files(".",recursive=TRUE,all.files=TRUE,no..=TRUE),flags="-q")
}

rtf_escape <- function(x){x<-gsub("\\\\","\\\\\\\\",as.character(x));gsub("([{}])","\\\\\\1",x)}
write_rtf <- function(tab, title, path, original_n, simulated_n) {
  con<-file(path,"wb");on.exit(close(con));w<-function(...)writeLines(paste0(...),con,useBytes=TRUE)
  w("{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\landscape\\paperw15840\\paperh12240\\margl700\\margr700\\margt500\\margb600")
  w("\\pard\\f0\\fs15 PILOT\\tab Confidential\\tab Page {\\field{\\*\\fldinst PAGE}} of {\\field{\\*\\fldinst NUMPAGES}}\\par")
  w("\\pard\\qc\\b\\fs20 ",rtf_escape(title),"\\b0\\par\\pard\\qc\\fs16 Safety Population\\par\\par")
  n<-ncol(tab);ends<-cumsum(c(3400,rep((14440-3400)/(n-1),n-1)))
  for(i in 0:nrow(tab)){vals<-if(i==0)names(tab)else unlist(tab[i,],use.names=FALSE);bd<-if(i==0)"\\clbrdrt\\brdrs\\brdrw15\\clbrdrb\\brdrs\\brdrw15"else if(i==nrow(tab))"\\clbrdrb\\brdrs\\brdrw15"else"";w("\\trowd\\trgaph65",if(i==0)"\\trhdr"else"",paste0(bd,"\\cellx",round(ends),collapse=""));for(j in seq_len(n))w("\\intbl",if(j==1)"\\ql"else"\\qc",if(i==0)"\\b"else"\\b0","\\f0\\fs14 ",rtf_escape(vals[j]),"\\cell");w("\\row")}
  w("\\pard\\ql\\fs13\\par Note: Event is defined as CNSR=0; censored observation as CNSR=1. Median and confidence limits are based on the Kaplan-Meier method.\\par")
  w("Original N=",original_n,"; simulated N=",simulated_n,". Source: simulated data generated by this standalone program. Generated: ",format(Sys.Date(),"%d%b%Y"),"\\par}")
}

fmt_ci <- function(x,l,u) if(any(is.na(c(x,l,u)))) "NE" else sprintf("%.1f (%.1f, %.1f)",x,l,u)
label_trt <- function(x) ifelse(x==0,"Placebo",paste0("Treatment ",x," mg"))
km_row <- function(d) {
  fit<-survfit(Surv(AVAL,1-CNSR)~1,data=d,conf.type="log-log")
  t<-summary(fit)$table
  c(N=nrow(d),Events=sum(d$CNSR==0),Censored=sum(d$CNSR==1),Median=fmt_ci(t["median"],t["0.95LCL"],t["0.95UCL"]))
}
analyze <- function(d) {
  groups<-split(d,d$TRT01AN); all_groups<-c(groups,list(Total=d))
  rows<-lapply(seq_along(all_groups),function(i){x<-all_groups[[i]];z<-km_row(x);data.frame(Treatment=if(i<=length(groups))label_trt(as.numeric(names(groups)[i]))else"Total",`N`=z["N"],`Events n (%)`=sprintf("%d (%.1f%%)",as.integer(z["Events"]),100*as.integer(z["Events"])/as.integer(z["N"])),`Censored n (%)`=sprintf("%d (%.1f%%)",as.integer(z["Censored"]),100*as.integer(z["Censored"])/as.integer(z["N"])),`Median Time (95% CI)`=z["Median"],check.names=FALSE)})
  do.call(rbind,rows)
}

# Original source data embedded from the DOCX with PILOT identifiers.
adsl <- data.frame(USUBJID=c("PILOT-1001","PILOT-1002","PILOT-1003","PILOT-1004","PILOT-1005","PILOT-1006","PILOT-1007","PILOT-1008","PILOT-1009","PILOT-1010","PILOT-1011","PILOT-1012","PILOT-1013","PILOT-1014","PILOT-1015","PILOT-1016","PILOT-1017","PILOT-1018","PILOT-1019","PILOT-1020","PILOT-1021","PILOT-1022","PILOT-1023","PILOT-1024","PILOT-1025","PILOT-1026","PILOT-1027","PILOT-1028","PILOT-1029","PILOT-1030","PILOT-1031","PILOT-1032","PILOT-1033","PILOT-1034","PILOT-1035","PILOT-1036","PILOT-1037","PILOT-1038","PILOT-1039","PILOT-1040","PILOT-1041","PILOT-1042","PILOT-1043","PILOT-1044","PILOT-1045","PILOT-1046","PILOT-1047","PILOT-1048","PILOT-1049","PILOT-1050","PILOT-1051","PILOT-1052","PILOT-1053","PILOT-1054","PILOT-1055","PILOT-1056","PILOT-1057","PILOT-1058","PILOT-1059","PILOT-1060"),TRT01AN=c(0,54,54,0,54,0,0,0,54,0,54,54,54,54,0,0,0,0,0,54,0,0,54,0,0,0,0,54,54,0,54,54,0,0,54,54,0,0,0,54,54,54,0,0,54,54,0,0,54,54,54,54,54,0,0,54,54,54,54,0),SAFFL=c("Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y"),check.names=FALSE)
adtte <- data.frame(USUBJID=c("PILOT-1001","PILOT-1002","PILOT-1003","PILOT-1004","PILOT-1005","PILOT-1006","PILOT-1007","PILOT-1008","PILOT-1009","PILOT-1010","PILOT-1011","PILOT-1012","PILOT-1013","PILOT-1014","PILOT-1015","PILOT-1016","PILOT-1017","PILOT-1018","PILOT-1019","PILOT-1020","PILOT-1021","PILOT-1022","PILOT-1023","PILOT-1024","PILOT-1025","PILOT-1026","PILOT-1027","PILOT-1028","PILOT-1029","PILOT-1030","PILOT-1031","PILOT-1032","PILOT-1033","PILOT-1034","PILOT-1035","PILOT-1036","PILOT-1037","PILOT-1038","PILOT-1039","PILOT-1040","PILOT-1041","PILOT-1042","PILOT-1043","PILOT-1044","PILOT-1045","PILOT-1046","PILOT-1047","PILOT-1048","PILOT-1049","PILOT-1050","PILOT-1051","PILOT-1052","PILOT-1053","PILOT-1054","PILOT-1055","PILOT-1056","PILOT-1057","PILOT-1058","PILOT-1059","PILOT-1060"),AVAL=c(182,2,17,177,15,48,183,183,43,64,25,15,41,39,116,195,57,175,31,15,64,182,35,189,43,177,182,119,69,148,27,27,77,28,30,34,21,26,183,126,12,107,181,42,60,5,33,135,8,112,41,15,15,110,1,19,24,17,49,70),CNSR=c(1,0,0,1,0,0,1,1,1,1,0,0,1,0,1,1,1,1,1,0,0,1,1,1,0,0,1,0,1,1,0,0,0,0,0,0,0,0,1,0,0,0,1,0,0,0,0,1,1,0,0,0,0,0,0,0,1,0,1,1),SAFFL=c("Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y","Y"),TRT01AN=c(0,54,54,0,54,0,0,0,54,0,54,54,54,54,0,0,0,0,0,54,0,0,54,0,0,0,0,54,54,0,54,54,0,0,54,54,0,0,0,54,54,54,0,0,54,54,0,0,54,54,54,54,54,0,0,54,54,54,54,0),check.names=FALSE)
TITLE <- "Table 3. Kaplan-Meier Summary of Time-to-Event Endpoint - Dataset 3"

original_n <- length(unique(adsl$USUBJID))
sim <- make_simulated(adsl, adtte)
stopifnot(nrow(sim$ADSL)==2L*original_n, all(names(sim$ADSL)==toupper(names(sim$ADSL))), all(names(sim$ADTTE)==toupper(names(sim$ADTTE))))
xlsx_path <- file.path(out_dir,paste0(stem,"_simulated_data.xlsx"))
rtf_path <- file.path(out_dir,paste0(stem,".rtf"))
write_xlsx(sim,xlsx_path)
result <- analyze(sim$ADTTE)
write_rtf(result, TITLE, rtf_path, original_n, nrow(sim$ADSL))
message(sprintf("Created %s: original N=%d; simulated N=%d",stem,original_n,nrow(sim$ADSL)))
