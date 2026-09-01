# Listing 4. Dates of First and Last Evaluations
# Standalone program based on the fourth code block in Listings- Disposition.docx.

set.seed(20260833)
a <- grep("^--file=",commandArgs(),value=TRUE)
out_dir <- if(length(a)) dirname(normalizePath(sub("^--file=","",a[1]))) else getwd()
stem <- "04_dates_first_last_evaluations_listing"

# Original 21-subject data; RFICDT/RFPENDT are SAS dates (origin 1960-01-01).
original <- data.frame(
  USUBJID=sprintf("PILOT-%04d",1001:1021),
  TRT01P=c("Arm B","Arm A","Screen Failure","Arm B","Arm A","Arm B","Arm A","Screen Failure","Arm B","Screen Failure","Arm B","Arm A","Arm A","Arm B","Screen Failure","Arm A","Screen Failure","Screen Failure","Arm A","Screen Failure","Arm B"),
  TRT01PN=c(2,1,NA,2,1,2,1,NA,2,NA,2,1,1,2,NA,1,NA,NA,1,NA,2),
  SITEID=c("002",rep("003",11),rep("004",4),rep("005",2),rep("007",2),"009"),
  SITENAM=paste("Site",c("002",rep("003",11),rep("004",4),rep("005",2),rep("007",2),"009")),
  AMENDNO=c(2,1,1,1,1,2,2,2,2,2,1,2,2,1,2,2,2,2,2,1,1),
  RFICDT=c(7284,6922,6690,6748,7042,7139,7397,7119,7342,7364,6804,7385,7357,7033,7015,7021,7384,7390,7224,6543,6963),
  RFPENDT=c(7552,7205,6852,7010,7345,7340,7545,7309,7502,7515,7026,7548,7678,7342,7163,7329,7676,7571,7543,6884,7118), stringsAsFactors=FALSE
)

# Resample full records to retain logical relationships, then shift both dates together.
n_original <- nrow(original)
simulated <- original[sample(seq_len(n_original),2L*n_original,replace=TRUE),]
simulated$USUBJID <- sprintf("PILOT-%04d",2001:(2000+nrow(simulated)))
shift <- sample(-45:45,nrow(simulated),replace=TRUE)
simulated$RFICDT <- simulated$RFICDT + shift; simulated$RFPENDT <- simulated$RFPENDT + shift
rownames(simulated) <- NULL

# Package-free Open XML workbook writer.
xe<-function(x){x<-ifelse(is.na(x),"",as.character(x));x<-gsub("&","&amp;",x,fixed=TRUE);x<-gsub("<","&lt;",x,fixed=TRUE);gsub(">","&gt;",x,fixed=TRUE)}
xc<-function(n){z<-character(n);for(i in seq_len(n)){k<-i;s<-"";while(k>0){k<-k-1L;s<-paste0(LETTERS[k%%26L+1L],s);k<-k%/%26L};z[i]<-s};z}
write_xlsx<-function(d,path){td<-tempfile("xlsx_");dir.create(td);dir.create(file.path(td,"_rels"));dir.create(file.path(td,"xl"));dir.create(file.path(td,"xl","_rels"));dir.create(file.path(td,"xl","worksheets"));writeLines('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',file.path(td,"[Content_Types].xml"));writeLines('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',file.path(td,"_rels",".rels"));writeLines('<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="SIMULATED_DATA" sheetId="1" r:id="rId1"/></sheets></workbook>',file.path(td,"xl","workbook.xml"));writeLines('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',file.path(td,"xl","_rels","workbook.xml.rels"));m<-rbind(names(d),as.data.frame(lapply(d,as.character),stringsAsFactors=FALSE));rows<-vapply(seq_len(nrow(m)),function(i)paste0('<row r="',i,'">',paste0('<c r="',xc(ncol(m)),i,'" t="inlineStr"><is><t>',xe(unlist(m[i,],use.names=FALSE)),'</t></is></c>',collapse=''),'</row>'),"");writeLines(paste0('<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" state="frozen"/></sheetView></sheetViews><sheetData>',paste(rows,collapse=''),'</sheetData><autoFilter ref="A1:',tail(xc(ncol(d)),1),nrow(d)+1L,'"/></worksheet>'),file.path(td,"xl","worksheets","sheet1.xml"));old<-setwd(td);on.exit(setwd(old),add=TRUE);if(file.exists(path))unlink(path);utils::zip(path,list.files(".",recursive=TRUE,all.files=TRUE,no..=TRUE),flags="-q")}

# RTF follows the screenshot's site grouping and date presentation.
esc<-function(x){x<-gsub("\\\\","\\\\\\\\",as.character(x));gsub("([{}])","\\\\\\1",x)}
sas_date<-function(x)format(as.Date(x,origin="1960-01-01"),"%Y-%m-%d")
write_rtf<-function(d,path){con<-file(path,"wb");on.exit(close(con),add=TRUE);w<-function(x)writeLines(x,con,useBytes=TRUE);w('{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\landscape\\paperw15840\\paperh12240\\margl650\\margr650\\margt450\\margb500');w('\\pard\\fs15 PILOT\\tab Confidential\\tab Page {\\field{\\*\\fldinst PAGE}} of {\\field{\\*\\fldinst NUMPAGES}}\\par\\pard\\qc\\b\\fs20 Listing 4. Dates of First and Last Evaluations\\b0\\par\\pard\\qc\\fs15 All Subjects\\par\\par');lab<-c("Subject ID","Treatment Group","Date of First Evaluation [1]","Date of Last Evaluation [1]","Subject Enrolled Under Protocol Version");ends<-cumsum(c(2300,2300,2900,2900,2840));w(paste0('\\trowd\\trhdr',paste0('\\clbrdrt\\brdrs\\brdrw15\\clbrdrb\\brdrs\\brdrw15\\cellx',ends,collapse=''),paste0('\\intbl\\ql\\b\\fs15 ',esc(lab),'\\cell',collapse=''),'\\row'));for(s in unique(d$SITEID)){q<-d[d$SITEID==s,];w(paste0('\\trowd\\cellx13240\\intbl\\ql\\b Investigator Site = ',esc(s),'\\b0\\cell\\row'));for(i in seq_len(nrow(q))){v<-c(q$USUBJID[i],q$TRT01P[i],sas_date(q$RFICDT[i]),sas_date(q$RFPENDT[i]),q$AMENDNO[i]);w(paste0('\\trowd',paste0('\\cellx',ends,collapse=''),paste0('\\intbl\\ql\\fs15 ',esc(v),'\\cell',collapse=''),'\\row'))}};w('\\trowd\\clbrdrb\\brdrs\\brdrw15\\cellx13240\\intbl\\cell\\row\\pard\\fs12\\par [1] Dates are displayed as YYYY-MM-DD.\\par Source: simulated data generated from the original DOCX data logic.\\par}')}

simulated<-simulated[order(simulated$SITEID,simulated$USUBJID),]
write_xlsx(simulated,file.path(out_dir,paste0(stem,"_simulated_data.xlsx")))
write_rtf(simulated,file.path(out_dir,paste0(stem,".rtf")))
stopifnot(nrow(simulated)==42L,all(names(simulated)==toupper(names(simulated))),all(simulated$RFPENDT>=simulated$RFICDT))
message("Created ",stem,": original N=21; simulated N=42")
