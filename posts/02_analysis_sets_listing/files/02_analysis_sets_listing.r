# Listing 2. Inclusion/Exclusion From Analysis Sets
# Standalone program: original data, 2 x N simulation, XLSX export, and clinical RTF.

set.seed(20260831)
a <- grep("^--file=", commandArgs(), value=TRUE)
out_dir <- if(length(a)) dirname(normalizePath(sub("^--file=", "", a[1]))) else getwd()
stem <- "02_analysis_sets_listing"

# Original 30-subject analysis-set data from the DOCX.
site <- c(rep("002",2),rep("003",12),rep("004",9),rep("007",7))
trt <- c("A","A","B","B","A","A","B","A","A","B","B","B","A","A","B","A","A","B","A","B","B","B","B","B","A","B","A","B","A","A")
fas <- c("Y","Y","Y","N","N","N","Y","N","Y","N","N","N","N","N","N","Y","N","Y","N","N","N","N","N","Y","N","N","N","N","N","Y")
original <- data.frame(
  USUBJID=sprintf("PILOT-%04d",1001:1030), TRT01P=paste("Arm",trt), TRT01PN=ifelse(trt=="A",1L,2L),
  SITEID=site, SITENAM=paste("Site",site), FASFL=fas, SAFFL="Y", PPROTFL=fas, PKASFL="Y",
  stringsAsFactors=FALSE
)

# Double the source sample while preserving the original joint flag/treatment pattern.
n_original <- nrow(original)
simulated <- original[sample(seq_len(n_original),2L*n_original,replace=TRUE),]
simulated$USUBJID <- sprintf("PILOT-%04d",2001:(2000+nrow(simulated)))
rownames(simulated) <- NULL

# Self-contained XLSX writer (Open XML); avoids soffice, writexl, and openxlsx.
xml <- function(x){x<-ifelse(is.na(x),"",as.character(x));x<-gsub("&","&amp;",x,fixed=TRUE);x<-gsub("<","&lt;",x,fixed=TRUE);gsub(">","&gt;",x,fixed=TRUE)}
xcol <- function(n){z<-character(n);for(i in seq_len(n)){k<-i;s<-"";while(k>0){k<-k-1L;s<-paste0(LETTERS[k%%26L+1L],s);k<-k%/%26L};z[i]<-s};z}
write_xlsx <- function(d,path){
  td<-tempfile("xlsx_");dir.create(td);dir.create(file.path(td,"_rels"));dir.create(file.path(td,"xl"));dir.create(file.path(td,"xl","_rels"));dir.create(file.path(td,"xl","worksheets"))
  writeLines('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>',file.path(td,"[Content_Types].xml"))
  writeLines('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',file.path(td,"_rels",".rels"))
  writeLines('<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="SIMULATED_DATA" sheetId="1" r:id="rId1"/></sheets></workbook>',file.path(td,"xl","workbook.xml"))
  writeLines('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',file.path(td,"xl","_rels","workbook.xml.rels"))
  writeLines('<?xml version="1.0"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><name val="Arial"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf/><xf fontId="1" fillId="2" applyFont="1" applyFill="1"/></cellXfs></styleSheet>',file.path(td,"xl","styles.xml"))
  m<-rbind(names(d),as.data.frame(lapply(d,as.character),stringsAsFactors=FALSE));rows<-vapply(seq_len(nrow(m)),function(i)paste0('<row r="',i,'">',paste0('<c r="',xcol(ncol(m)),i,'" t="inlineStr" s="',if(i==1)1 else 0,'"><is><t>',xml(unlist(m[i,],use.names=FALSE)),'</t></is></c>',collapse=''),'</row>'),"")
  writeLines(paste0('<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" state="frozen"/></sheetView></sheetViews><sheetData>',paste(rows,collapse=''),'</sheetData><autoFilter ref="A1:',tail(xcol(ncol(d)),1),nrow(d)+1L,'"/></worksheet>'),file.path(td,"xl","worksheets","sheet1.xml"))
  old<-setwd(td);on.exit(setwd(old),add=TRUE);if(file.exists(path))unlink(path);utils::zip(path,list.files(".",recursive=TRUE,all.files=TRUE,no..=TRUE),flags="-q")
}

# RTF listing grouped by site; Yes/No labels match the reference layout.
esc <- function(x){x<-gsub("\\\\","\\\\\\\\",as.character(x));gsub("([{}])","\\\\\\1",x)}
write_rtf <- function(d,path){
  con<-file(path,"wb");on.exit(close(con),add=TRUE);w<-function(x)writeLines(x,con,useBytes=TRUE)
  w('{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\landscape\\paperw15840\\paperh12240\\margl600\\margr600\\margt450\\margb500')
  w('\\pard\\fs15 PILOT\\tab Confidential\\tab Page {\\field{\\*\\fldinst PAGE}} of {\\field{\\*\\fldinst NUMPAGES}}\\par\\pard\\qc\\b\\fs20 Listing 2. Inclusion/Exclusion From Analysis Sets\\b0\\par\\pard\\qc\\fs15 All Subjects\\par\\par')
  lab<-c("Subject ID","Treatment Group","Full Analysis Set","Per Protocol Set","Safety Analysis Set","Pharmacokinetic Analysis Set");ends<-cumsum(c(2100,2200,2200,2200,2200,3740))
  w(paste0('\\trowd\\trhdr\\trgaph60',paste0('\\clbrdrt\\brdrs\\brdrw15\\clbrdrb\\brdrs\\brdrw15\\cellx',ends,collapse=''),paste0('\\intbl\\ql\\b\\fs15 ',esc(lab),'\\cell',collapse=''),'\\row'))
  yn<-function(x)ifelse(x=="Y","Yes","No")
  for(s in unique(d$SITEID)){q<-d[d$SITEID==s,];w(paste0('\\trowd\\cellx14640\\intbl\\ql\\b Investigator Site = ',esc(s),'\\b0\\cell\\row'));for(i in seq_len(nrow(q))){v<-c(q$USUBJID[i],q$TRT01P[i],yn(q$FASFL[i]),yn(q$PPROTFL[i]),yn(q$SAFFL[i]),yn(q$PKASFL[i]));w(paste0('\\trowd',paste0('\\cellx',ends,collapse=''),paste0('\\intbl\\ql\\fs15 ',esc(v),'\\cell',collapse=''),'\\row'))}}
  w('\\trowd\\clbrdrb\\brdrs\\brdrw15\\cellx14640\\intbl\\cell\\row\\pard\\fs12\\par Source: simulated data generated from the original DOCX data logic.\\par}')
}

simulated<-simulated[order(simulated$SITEID,simulated$USUBJID),]
write_xlsx(simulated,file.path(out_dir,paste0(stem,"_simulated_data.xlsx")))
write_rtf(simulated,file.path(out_dir,paste0(stem,".rtf")))
stopifnot(nrow(simulated)==60L,all(names(simulated)==toupper(names(simulated))))
message("Created ",stem,": original N=30; simulated N=60")
