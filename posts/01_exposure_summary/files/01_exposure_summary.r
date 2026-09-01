#!/usr/bin/env Rscript

# PILOT TEX L101 - Exposure Summary
# Standalone: embedded source data -> doubled simulation -> XLSX -> analysis -> RTF.
options(stringsAsFactors=FALSE)
set.seed(20260827)
suppressPackageStartupMessages(library(tidyverse))

script_dir<-function(){a<-commandArgs(FALSE);f<-sub("^--file=","",a[grepl("^--file=",a)]);f<-gsub("~\\+~"," ",f);if(length(f))dirname(normalizePath(f[1]))else getwd()}
OUT_DIR<-script_dir();STEM<-"01_exposure_summary"

# Original 25-subject ADSL data from the DOCX.
original_data<-tribble(
~USUBJID,~TRT01AN,~SAFFL,~TRDURN,~TRDGR1,~TRDGR1N,~TRCPL,
"DEAA300001",2,"Y",86,">=42",5,100,
"DEAA400001",2,"Y",85,">=42",5,100,
"DGAA300001",2,"Y",83,">=42",5,98.8,
"DGAA300002",1,"Y",82,">=42",5,100,
"DGAA300006",2,"Y",84,">=42",5,94,
"DGAA300012",2,"Y",85,">=42",5,102.4,
"DGAA300013",1,"Y",48,">=42",5,91.7,
"DGAA300014",2,"Y",85,">=42",5,97.6,
"DGAB200001",1,"Y",85,">=42",5,100,
"DGAB200003",1,"Y",86,">=42",5,98.8,
"DGAB200010",2,"Y",85,">=42",5,100,
"ECAA200001",2,"Y",83,">=42",5,100,
"ECAA200002",1,"Y",85,">=42",5,100,
"ECAA200003",2,"Y",85,">=42",5,101.2,
"ECAA500004",1,"Y",85,">=42",5,102.4,
"ECAA500005",2,"Y",88,">=42",5,100,
"EEAA400001",2,"Y",83,">=42",5,95.2,
"EFAA100002",2,"Y",87,">=42",5,95.4,
"EFAA100003",2,"Y",49,">=42",5,95.9,
"EFAA100004",2,"Y",84,">=42",5,101.2,
"EFAA200001",1,"Y",83,">=42",5,101.2,
"EFAA400001",1,"Y",84,">=42",5,100,
"EFAA400002",1,"Y",83,">=42",5,100,
"EFAA400003",1,"Y",18,">=14 to <28",3,94.4,
"EJAA300002",1,"Y",85,">=42",5,97.6)

duration_group<-function(x){case_when(x<7~"<7",x<14~">=7 to <14",x<28~">=14 to <28",x<42~">=28 to <42",TRUE~">=42")}
duration_group_n<-function(x){case_when(x<7~1L,x<14~2L,x<28~3L,x<42~4L,TRUE~5L)}

# Each source record contributes two new subjects. Duration and compliance are
# perturbed within plausible bounds; duration categories are then re-derived.
simulated_data<-original_data[rep(seq_len(nrow(original_data)),each=2),]
simulated_data$USUBJID<-sprintf("PILOT-%04d",seq_len(nrow(simulated_data)))
simulated_data$TRDURN<-pmax(1L,round(simulated_data$TRDURN+rnorm(nrow(simulated_data),0,4)))
simulated_data$TRCPL<-round(pmin(110,pmax(75,simulated_data$TRCPL+rnorm(nrow(simulated_data),0,2.5))),1)
simulated_data$TRDGR1<-duration_group(simulated_data$TRDURN)
simulated_data$TRDGR1N<-duration_group_n(simulated_data$TRDURN)
stopifnot(nrow(original_data)==25L,nrow(simulated_data)==50L,all(names(simulated_data)==toupper(names(simulated_data))))

# Standalone XLSX writer.
xe<-function(x){x[is.na(x)]<-"";x<-gsub("&","&amp;",x,fixed=TRUE);x<-gsub("<","&lt;",x,fixed=TRUE);x<-gsub(">","&gt;",x,fixed=TRUE);gsub('"',"&quot;",x,fixed=TRUE)}
xc<-function(n){z<-character(length(n));for(j in seq_along(n)){k<-n[j];s<-"";while(k>0){k<-k-1;s<-paste0(intToUtf8(65+k%%26),s);k<-k%/%26};z[j]<-s};z}
write_xlsx<-function(d,p){td<-tempfile("x");dir.create(td);ds<-c("_rels","docProps","xl","xl/_rels","xl/worksheets");invisible(vapply(file.path(td,ds),dir.create,logical(1),recursive=TRUE));writeLines('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>',file.path(td,"[Content_Types].xml"));writeLines('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>',file.path(td,"_rels/.rels"));writeLines('<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="SIMULATED_DATA" sheetId="1" r:id="rId1"/></sheets></workbook>',file.path(td,"xl/workbook.xml"));writeLines('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',file.path(td,"xl/_rels/workbook.xml.rels"));writeLines('<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="10"/><name val="Arial"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf xfId="0"/></cellXfs></styleSheet>',file.path(td,"xl/styles.xml"));a<-rbind(names(d),as.matrix(d));rows<-vapply(seq_len(nrow(a)),function(i){cc<-vapply(seq_len(ncol(a)),function(j)sprintf('<c r="%s%d" t="inlineStr"><is><t>%s</t></is></c>',xc(j),i,xe(as.character(a[i,j]))),character(1));sprintf('<row r="%d">%s</row>',i,paste0(cc,collapse=""))},character(1));writeLines(paste0('<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>',paste0(rows,collapse=""),'</sheetData></worksheet>'),file.path(td,"xl/worksheets/sheet1.xml"));writeLines('<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>PILOT</dc:creator></cp:coreProperties>',file.path(td,"docProps/core.xml"));writeLines('<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>R</Application></Properties>',file.path(td,"docProps/app.xml"));old<-setwd(td);on.exit(setwd(old),add=TRUE);if(file.exists(p))unlink(p);utils::zip(p,list.files(".",recursive=TRUE,all.files=TRUE,no..=TRUE),flags="-q")}
write_xlsx(simulated_data,file.path(OUT_DIR,paste0(STEM,"_simulated_data.xlsx")))

# Original analysis structure retained.
adsl01<-simulated_data%>%filter(SAFFL=="Y")
adsl02<-adsl01%>%mutate(TREATMENT=TRT01AN)%>%bind_rows(adsl01%>%mutate(TREATMENT=4L))
trt_levels<-1:4;trt_n<-vapply(trt_levels,function(g)sum(adsl02$TREATMENT==g),integer(1))
fmt_num<-function(x,dp){nmiss<-sum(is.na(x));x<-x[!is.na(x)];if(!length(x))return(c(sprintf("0 (%d)",nmiss),rep("-",4)));f0<-paste0("%.",dp,"f");f1<-paste0("%.",dp+1,"f");f2<-paste0("%.",dp+2,"f");c(sprintf("%d (%d)",length(x),nmiss),sprintf("%s (%s)",sprintf(f1,mean(x)),sprintf(f2,sd(x))),sprintf(f1,median(x)),sprintf("%s, %s",sprintf(f1,quantile(x,.25,type=2)),sprintf(f1,quantile(x,.75,type=2))),sprintf("%s, %s",sprintf(f0,min(x)),sprintf(f0,max(x))))}
specs<-tribble(~LABEL,~VAR,~DP,"Duration (days)","TRDURN",0,"Treatment Compliance (%)","TRCPL",1)
stat_labels<-c("n (missing)","Mean (SD)","Median","Q1, Q3","Min, Max")
num_rows<-list();for(i in seq_len(nrow(specs))){block<-sapply(trt_levels,function(g)fmt_num(adsl02[adsl02$TREATMENT==g,][[specs$VAR[i]]],specs$DP[i]));num_rows[[i]]<-cbind(LABEL=c(specs$LABEL[i],rep("",4)),STATISTIC=stat_labels,block)}
fmt_cat<-function(n,d)if(n==0||d==0)"0"else sprintf("%d (%.1f%%)",n,100*n/d)
cat_levels<-c("<7",">=7 to <14",">=14 to <28",">=28 to <42",">=42","Missing")
cat_rows<-do.call(rbind,lapply(seq_along(cat_levels),function(i){lev<-cat_levels[i];vals<-vapply(trt_levels,function(g){z<-adsl02%>%filter(TREATMENT==g);n<-if(lev=="Missing")sum(is.na(z$TRDGR1)|z$TRDGR1=="")else sum(z$TRDGR1==lev,na.rm=TRUE);fmt_cat(n,nrow(z))},character(1));c(if(i==1)"Duration (days), n (%)"else"",lev,vals)}))
body<-rbind(do.call(rbind,num_rows),cat_rows)
tab<-rbind(c("Parameter","Statistic",sprintf("Treatment 1\n(N=%d)",trt_n[1]),sprintf("Treatment 2\n(N=%d)",trt_n[2]),sprintf("Treatment 3\n(N=%d)",trt_n[3]),sprintf("Total\n(N=%d)",trt_n[4])),body)

re<-function(x){x<-gsub("\\","\\\\",x,fixed=TRUE);x<-gsub("{","\\{",x,fixed=TRUE);x<-gsub("}","\\}",x,fixed=TRUE);gsub("\n","\\line ",x,fixed=TRUE)}
write_rtf<-function(tab,p){con<-file(p,"wb");on.exit(close(con));w<-function(...)writeLines(paste0(...),con,useBytes=TRUE);w("{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\landscape\\paperw15840\\paperh12240\\margl600\\margr600\\margt500\\margb600");w("\\pard\\f0\\fs15 PILOT\\tab Confidential\\tab Page {\\field{\\*\\fldinst PAGE}} of {\\field{\\*\\fldinst NUMPAGES}}\\par");w("\\pard\\qc\\b\\fs20 Table 1. Summary of Treatment Exposure and Compliance\\b0\\par");w("\\pard\\qc\\fs16 Safety Analysis Set\\par\\par");ends<-cumsum(c(3100,2300,rep(2150,4)));for(i in seq_len(nrow(tab))){bd<-if(i==1)"\\clbrdrt\\brdrs\\brdrw15\\clbrdrb\\brdrs\\brdrw15"else if(i==nrow(tab))"\\clbrdrb\\brdrs\\brdrw15"else"";w("\\trowd\\trgaph65",if(i==1)"\\trhdr"else"",paste0(bd,"\\cellx",ends,collapse=""));for(j in seq_len(ncol(tab)))w("\\intbl",if(j<=2)"\\ql"else"\\qc",if(i==1||j==1&&nzchar(tab[i,1]))"\\b"else"\\b0","\\f0\\fs14 ",re(as.character(tab[i,j])),"\\cell");w("\\row")};w("\\pard\\ql\\fs13\\par Percentages are based on subjects in the Safety Analysis Set within each treatment group.\\par");w("\\pard\\ql\\fs13 N = number of subjects; n = number in category; Q1 = first quartile; Q3 = third quartile; SD = standard deviation.\\par");w("\\pard\\ql\\fs13 Source: Simulated data generated by this standalone program. Generated: ",format(Sys.Date(),"%d%b%Y"),"\\par}")}
write_rtf(tab,file.path(OUT_DIR,paste0(STEM,".rtf")))
cat("Created",STEM,"with 50 simulated subjects.\n")
