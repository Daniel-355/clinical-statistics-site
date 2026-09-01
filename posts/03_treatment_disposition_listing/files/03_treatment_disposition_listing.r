# Listing 3. Treatment Discontinuation
# Standalone program based on the third code block in Listings- Disposition.docx.

set.seed(20260832)
a <- grep("^--file=",commandArgs(),value=TRUE)
out_dir <- if(length(a)) dirname(normalizePath(sub("^--file=","",a[1]))) else getwd()
stem <- "03_treatment_disposition_listing"

# Original 23-subject data. Day variables are SAS study-day/date values.
original <- data.frame(
  USUBJID=sprintf("PILOT-%04d",1001:1023),
  TRT01A=c("Arm B","Arm B","Arm A","Arm A","Arm A","Arm A","Arm B","Arm A","Arm B","Arm A","Arm A","Arm B","Arm B","Arm A","Arm B","Arm A","Arm A","Arm B","Arm A","Arm A","Arm A","Arm A","Arm A"),
  TRT01AN=c(2,2,1,1,1,1,2,1,2,1,1,2,2,1,2,1,1,2,1,1,1,1,1),
  TRTEDT=c(228,85,77,97,31,311,305,104,566,151,57,505,423,38,514,125,NA,26,55,339,158,208,1),
  TRTEDY=c(228,85,77,97,31,311,305,104,566,151,57,505,423,38,514,125,NA,26,55,339,158,208,1),
  SITEID=c(rep("002",5),rep("003",16),rep("004",2)),
  SITENAM=c(rep("Site 002",5),rep("Site 003",16),rep("Site 004",2)),
  AGE=c(48,67,72,55,48,51,84,66,67,72,72,63,59,59,65,75,59,78,64,71,48,67,56),
  SEX=c("F","F","F","F","M","M","M","M","M","F","F","M","M","M","F","M","M","M","F","F","M","F","M"),
  ARACE=c("BLACK OR AFRICAN AMERICAN",rep("ASIAN",4),"BLACK OR AFRICAN AMERICAN",rep("WHITE",17)), DSTFL="Y",
  DSTREAS=c("NON-COMPLIANCE WITH STUDY DRUG","WITHDRAWAL BY SUBJECT","OTHER","NON-COMPLIANCE WITH STUDY DRUG","NON-COMPLIANCE WITH STUDY DRUG","LACK OF EFFICACY","PHYSICIAN DECISION","LACK OF EFFICACY","DEATH","LOST TO FOLLOW-UP","PHYSICIAN DECISION","ADVERSE EVENT","PROTOCOL DEVIATION","WITHDRAWAL BY SUBJECT","ADVERSE EVENT","PHYSICIAN DECISION","LOST TO FOLLOW-UP","DEATH","ADVERSE EVENT","LOST TO FOLLOW-UP","OTHER","OTHER","PROTOCOL DEVIATION"),
  DSTDT=c(325,85,78,141,92,314,308,107,581,152,71,531,429,38,517,127,113,28,56,375,269,208,8),
  DSTDY=c(325,85,78,141,92,314,308,107,581,152,71,531,429,38,517,127,113,28,56,375,269,208,8), stringsAsFactors=FALSE
)

# Preserve coherent source rows, double N, and add small age/day perturbations.
n_original <- nrow(original)
simulated <- original[sample(seq_len(n_original),2L*n_original,replace=TRUE),]
simulated$USUBJID <- sprintf("PILOT-%04d",2001:(2000+nrow(simulated)))
simulated$AGE <- pmax(18L,simulated$AGE+sample(-3:3,nrow(simulated),TRUE))
shift <- sample(-10:10,nrow(simulated),TRUE); simulated$TRTEDT <- simulated$TRTEDY <- simulated$TRTEDT+shift
simulated$DSTDT <- simulated$DSTDY <- pmax(simulated$TRTEDT,simulated$DSTDT+shift,na.rm=TRUE)
rownames(simulated) <- NULL

# Package-free Open XML workbook writer.
xe<-function(x){x<-ifelse(is.na(x),"",as.character(x));x<-gsub("&","&amp;",x,fixed=TRUE);x<-gsub("<","&lt;",x,fixed=TRUE);gsub(">","&gt;",x,fixed=TRUE)}
xc<-function(n){z<-character(n);for(i in seq_len(n)){k<-i;s<-"";while(k>0){k<-k-1L;s<-paste0(LETTERS[k%%26L+1L],s);k<-k%/%26L};z[i]<-s};z}
write_xlsx<-function(d,path){td<-tempfile("xlsx_");dir.create(td);dir.create(file.path(td,"_rels"));dir.create(file.path(td,"xl"));dir.create(file.path(td,"xl","_rels"));dir.create(file.path(td,"xl","worksheets"));writeLines('<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>',file.path(td,"[Content_Types].xml"));writeLines('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',file.path(td,"_rels",".rels"));writeLines('<?xml version="1.0"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="SIMULATED_DATA" sheetId="1" r:id="rId1"/></sheets></workbook>',file.path(td,"xl","workbook.xml"));writeLines('<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',file.path(td,"xl","_rels","workbook.xml.rels"));m<-rbind(names(d),as.data.frame(lapply(d,as.character),stringsAsFactors=FALSE));rows<-vapply(seq_len(nrow(m)),function(i)paste0('<row r="',i,'">',paste0('<c r="',xc(ncol(m)),i,'" t="inlineStr"><is><t>',xe(unlist(m[i,],use.names=FALSE)),'</t></is></c>',collapse=''),'</row>'),"");writeLines(paste0('<?xml version="1.0"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" state="frozen"/></sheetView></sheetViews><sheetData>',paste(rows,collapse=''),'</sheetData><autoFilter ref="A1:',tail(xc(ncol(d)),1),nrow(d)+1L,'"/></worksheet>'),file.path(td,"xl","worksheets","sheet1.xml"));old<-setwd(td);on.exit(setwd(old),add=TRUE);if(file.exists(path))unlink(path);utils::zip(path,list.files(".",recursive=TRUE,all.files=TRUE,no..=TRUE),flags="-q")}

# Clinical RTF, with date values displayed from the SAS origin used by the DOCX.
esc<-function(x){x<-gsub("\\\\","\\\\\\\\",as.character(x));gsub("([{}])","\\\\\\1",x)}
day_date<-function(day){ifelse(is.na(day),"",paste0(day,"/",format(as.Date(day,origin="1960-01-01"),"%Y-%m-%d")))}
write_rtf<-function(d,path){con<-file(path,"wb");on.exit(close(con),add=TRUE);w<-function(x)writeLines(x,con,useBytes=TRUE);w('{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\landscape\\paperw15840\\paperh12240\\margl420\\margr420\\margt400\\margb450');w('\\pard\\fs14 PILOT\\tab Confidential\\tab Page {\\field{\\*\\fldinst PAGE}} of {\\field{\\*\\fldinst NUMPAGES}}\\par\\pard\\qc\\b\\fs19 Listing 3. Treatment Discontinuation\\b0\\par\\pard\\qc\\fs14 Safety Analysis Set\\par\\par');lab<-c("Subject ID","Treatment Group","Age (Years)","Sex","Race","Last Dose Day/Date","Primary Reason for End of Treatment","Discontinuation Day/Date");ends<-cumsum(c(1500,1500,1100,600,2600,2200,3300,1500));w(paste0('\\trowd\\trhdr',paste0('\\clbrdrt\\brdrs\\brdrw15\\clbrdrb\\brdrs\\brdrw15\\cellx',ends,collapse=''),paste0('\\intbl\\ql\\b\\fs13 ',esc(lab),'\\cell',collapse=''),'\\row'));for(s in unique(d$SITEID)){q<-d[d$SITEID==s,];w(paste0('\\trowd\\cellx14300\\intbl\\ql\\b Investigator Site = ',esc(s),'\\b0\\cell\\row'));for(i in seq_len(nrow(q))){v<-c(q$USUBJID[i],q$TRT01A[i],q$AGE[i],q$SEX[i],q$ARACE[i],day_date(q$TRTEDT[i]),q$DSTREAS[i],day_date(q$DSTDT[i]));w(paste0('\\trowd',paste0('\\cellx',ends,collapse=''),paste0('\\intbl\\ql\\fs13 ',esc(v),'\\cell',collapse=''),'\\row'))}};w('\\trowd\\clbrdrb\\brdrs\\brdrw15\\cellx14300\\intbl\\cell\\row\\pard\\fs11\\par Source: simulated data generated from the original DOCX data logic.\\par}')}

simulated<-simulated[order(simulated$SITEID,simulated$USUBJID),]
write_xlsx(simulated,file.path(out_dir,paste0(stem,"_simulated_data.xlsx")))
write_rtf(simulated,file.path(out_dir,paste0(stem,".rtf")))
stopifnot(nrow(simulated)==46L,all(names(simulated)==toupper(names(simulated))))
message("Created ",stem,": original N=23; simulated N=46")
