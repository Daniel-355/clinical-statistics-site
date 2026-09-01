# Standalone clinical inferential analysis generated from the corresponding DOCX code segment.

suppressPackageStartupMessages({library(tidyverse); library(broom)})
set.seed(20260827)
script_arg <- grep("^--file=", commandArgs(FALSE), value=TRUE)
SCRIPT_RAW <- if(length(script_arg)) sub("^--file=", "", script_arg[1]) else "."
SCRIPT_RAW <- gsub("~\\+~", " ", SCRIPT_RAW)
SCRIPT_PATH <- normalizePath(SCRIPT_RAW)
OUT_DIR <- if(file.info(SCRIPT_PATH)$isdir) SCRIPT_PATH else dirname(SCRIPT_PATH)
STEM <- tools::file_path_sans_ext(basename(SCRIPT_PATH))

make_id_maps <- function(data_list) {
  ids <- unique(unlist(lapply(data_list, function(d) {
    id <- intersect(c("usubjid", "USUBJID"), names(d))[1]
    if(is.na(id)) character() else as.character(d[[id]])
  })))
  list(
    a=setNames(sprintf("PILOT-%04d", seq_along(ids)), ids),
    b=setNames(sprintf("PILOT-%04d", seq_along(ids)+length(ids)), ids),
    original_n=length(ids)
  )
}
double_subjects <- function(d, maps) {
  id <- intersect(c("usubjid", "USUBJID"), names(d))[1]
  if(is.na(id)) return(bind_rows(d,d))
  a <- d; b <- d
  a[[id]] <- unname(maps$a[as.character(d[[id]])])
  b[[id]] <- unname(maps$b[as.character(d[[id]])])
  bind_rows(a,b)
}

xe <- function(x){x[is.na(x)]<-"";x<-gsub("&","&amp;",x,fixed=TRUE);x<-gsub("<","&lt;",x,fixed=TRUE);x<-gsub(">","&gt;",x,fixed=TRUE);gsub('"',"&quot;",x,fixed=TRUE)}
xc <- function(n){z<-character(length(n));for(j in seq_along(n)){k<-n[j];s<-"";while(k>0){k<-k-1;s<-paste0(intToUtf8(65+k%%26),s);k<-k%/%26};z[j]<-s};z}
write_xlsx_list <- function(data_list,p){
  data_list <- lapply(data_list,function(d){names(d)<-toupper(names(d));d})
  sn <- substr(gsub("[^A-Za-z0-9_]","_",names(data_list)),1,31)
  td<-tempfile("xlsx_");dir.create(td);for(z in c("_rels","docProps","xl","xl/_rels","xl/worksheets"))dir.create(file.path(td,z),recursive=TRUE,showWarnings=FALSE)
  types<-paste0('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',paste0('<Override PartName="/xl/worksheets/sheet',seq_along(data_list),'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',collapse=""),'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>')
  writeLines(types,file.path(td,"[Content_Types].xml"))
  writeLines('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>',file.path(td,"_rels/.rels"))
  sheets<-paste0('<sheet name="',sn,'" sheetId="',seq_along(sn),'" r:id="rId',seq_along(sn),'"/>',collapse="")
  writeLines(paste0('<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>',sheets,'</sheets></workbook>'),file.path(td,"xl/workbook.xml"))
  rels<-paste0('<Relationship Id="rId',seq_along(sn),'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet',seq_along(sn),'.xml"/>',collapse="")
  writeLines(paste0('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',rels,'<Relationship Id="rId',length(sn)+1,'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'),file.path(td,"xl/_rels/workbook.xml.rels"))
  writeLines('<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf xfId="0" fontId="0" fillId="0"/><xf xfId="0" fontId="1" fillId="1" applyFont="1" applyFill="1"/></cellXfs></styleSheet>',file.path(td,"xl/styles.xml"))
  for(k in seq_along(data_list)){
    d<-data_list[[k]]; a<-rbind(names(d),as.matrix(d)); rows<-vapply(seq_len(nrow(a)),function(i){cc<-vapply(seq_len(ncol(a)),function(j){v<-dplyr::if_else(i==1,as.character(a[i,j]),as.character(a[i,j])); if(i>1 && (is.numeric(d[[j]])||is.integer(d[[j]])) && !is.na(d[i-1,j])) sprintf('<c r="%s%d"><v>%s</v></c>',xc(j),i,as.character(d[i-1,j])) else sprintf('<c r="%s%d" t="inlineStr" s="%d"><is><t>%s</t></is></c>',xc(j),i,ifelse(i==1,1,0),xe(v))},character(1));sprintf('<row r="%d">%s</row>',i,paste0(cc,collapse=""))},character(1));
    writeLines(paste0('<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData>',paste0(rows,collapse=""),'</sheetData></worksheet>'),file.path(td,"xl/worksheets",paste0("sheet",k,".xml")))
  }
  writeLines('<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>PILOT</dc:creator></cp:coreProperties>',file.path(td,"docProps/core.xml"));writeLines('<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>R</Application></Properties>',file.path(td,"docProps/app.xml"))
  old<-setwd(td);on.exit(setwd(old),add=TRUE);if(file.exists(p))unlink(p);utils::zip(p,list.files(".",recursive=TRUE,all.files=TRUE,no..=TRUE),flags="-q")
}

rtf_escape <- function(x){x[is.na(x)]<-"-";x<-gsub("\\\\","\\\\\\\\",x);x<-gsub("([{}])","\\\\\\1",x);x}
write_clinical_rtf <- function(tab,title,p,footnote){
  tab<-as.data.frame(tab,check.names=FALSE);tab[]<-lapply(tab,as.character);con<-file(p,"wb");on.exit(close(con));w<-function(...)writeLines(paste0(...),con,useBytes=TRUE)
  w("{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\landscape\\paperw15840\\paperh12240\\margl600\\margr600\\margt500\\margb600")
  w("\\pard\\f0\\fs15 PILOT\\tab Confidential\\tab Page {\\field{\\*\\fldinst PAGE}} of {\\field{\\*\\fldinst NUMPAGES}}\\par")
  w("\\pard\\qc\\b\\fs20 ",rtf_escape(title),"\\b0\\par\\pard\\qc\\fs16 Analysis Population\\par\\par")
  n<-ncol(tab);first<-if(n>4)3200 else 4200; ends<-cumsum(c(first,rep((14640-first)/(n-1),n-1)))
  for(i in 0:nrow(tab)){vals<-if(i==0)names(tab) else unlist(tab[i,],use.names=FALSE);bd<-if(i==0)"\\clbrdrt\\brdrs\\brdrw15\\clbrdrb\\brdrs\\brdrw15"else if(i==nrow(tab))"\\clbrdrb\\brdrs\\brdrw15"else"";w("\\trowd\\trgaph65",if(i==0)"\\trhdr"else"",paste0(bd,"\\cellx",round(ends),collapse=""));for(j in seq_len(n))w("\\intbl",if(j==1)"\\ql"else"\\qc",if(i==0)"\\b"else"\\b0","\\f0\\fs14 ",rtf_escape(vals[j]),"\\cell");w("\\row")}
  w("\\pard\\ql\\fs13\\par ",rtf_escape(footnote),"\\par\\pard\\ql\\fs13 Source: Simulated data generated by this standalone program. Generated: ",format(Sys.Date(),"%d%b%Y"),"\\par}")
}
fmt_p <- function(p) ifelse(is.na(p),"-",ifelse(p<.0001,"<.0001",sprintf("%.4f",p)))
fmt_np <- function(n,N) ifelse(n==0,"0",sprintf("%d (%.1f%%)",n,100*n/N))

# Original embedded source data with PILOT identifiers.
adlb<-tribble(
~usubjid,~trtp,~avisit,~avisitn,~ablfl,~aval,~base,~chg,~trtpn,~saffl,~paramcd,~param,~paramn,
"P001","Placebo","Baseline",0,"Y",8.1,8.1,NA,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P001","Placebo","Week 4",1,"N",8.2,8.1,0.1,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P001","Placebo","Week 8",2,"N",7.9,8.1,-0.2,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P001","Placebo","Week 12",3,"N",7.7,8.1,-0.4,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P002","Placebo","Baseline",0,"Y",7.7,7.7,NA,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P002","Placebo","Week 4",1,"N",6.7,7.7,-1,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P002","Placebo","Week 8",2,"N",6.6,7.7,-1.1,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P002","Placebo","Week 12",3,"N",7.7,7.7,0,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P003","Placebo","Baseline",0,"Y",8.4,8.4,NA,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P003","Placebo","Week 4",1,"N",8.2,8.4,-0.2,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P003","Placebo","Week 8",2,"N",7.2,8.4,-1.2,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P003","Placebo","Week 12",3,"N",8.6,8.4,0.2,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P004","Placebo","Baseline",0,"Y",8.7,8.7,NA,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P004","Placebo","Week 4",1,"N",7.8,8.7,-0.9,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P004","Placebo","Week 8",2,"N",7.8,8.7,-0.9,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P004","Placebo","Week 12",3,"N",7.8,8.7,-0.9,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P005","Placebo","Baseline",0,"Y",8,8,NA,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P005","Placebo","Week 4",1,"N",7.5,8,-0.5,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P005","Placebo","Week 8",2,"N",7.4,8,-0.6,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P005","Placebo","Week 12",3,"N",7.2,8,-0.8,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P006","Placebo","Baseline",0,"Y",8.4,8.4,NA,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P006","Placebo","Week 4",1,"N",7.4,8.4,-1,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P006","Placebo","Week 8",2,"N",7.6,8.4,-0.8,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P006","Placebo","Week 12",3,"N",7.7,8.4,-0.7,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P007","Placebo","Baseline",0,"Y",8.2,8.2,NA,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P007","Placebo","Week 4",1,"N",8.1,8.2,-0.1,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P007","Placebo","Week 8",2,"N",7.3,8.2,-0.9,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P007","Placebo","Week 12",3,"N",7.7,8.2,-0.5,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P008","Placebo","Baseline",0,"Y",8.4,8.4,NA,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P008","Placebo","Week 4",1,"N",7.3,8.4,-1.1,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P008","Placebo","Week 8",2,"N",8.1,8.4,-0.3,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P008","Placebo","Week 12",3,"N",7.4,8.4,-1,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P009","Placebo","Baseline",0,"Y",7.6,7.6,NA,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P009","Placebo","Week 4",1,"N",7.7,7.6,0.1,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P009","Placebo","Week 8",2,"N",7.8,7.6,0.2,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P009","Placebo","Week 12",3,"N",7.5,7.6,-0.1,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P010","Placebo","Baseline",0,"Y",8,8,NA,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P010","Placebo","Week 4",1,"N",6.9,8,-1.1,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P010","Placebo","Week 8",2,"N",7.8,8,-0.2,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P010","Placebo","Week 12",3,"N",7.4,8,-0.6,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P011","Placebo","Baseline",0,"Y",7.7,7.7,NA,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P011","Placebo","Week 4",1,"N",7.2,7.7,-0.5,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P011","Placebo","Week 8",2,"N",6.5,7.7,-1.2,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P011","Placebo","Week 12",3,"N",7.8,7.7,0.1,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P012","Placebo","Baseline",0,"Y",7.9,7.9,NA,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P012","Placebo","Week 4",1,"N",7.6,7.9,-0.3,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P012","Placebo","Week 8",2,"N",7.1,7.9,-0.8,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P012","Placebo","Week 12",3,"N",7.4,7.9,-0.5,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P013","Placebo","Baseline",0,"Y",8.3,8.3,NA,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P013","Placebo","Week 4",1,"N",7.4,8.3,-0.9,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P013","Placebo","Week 8",2,"N",8.5,8.3,0.2,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P013","Placebo","Week 12",3,"N",8.2,8.3,-0.1,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P014","Placebo","Baseline",0,"Y",8.9,8.9,NA,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P014","Placebo","Week 4",1,"N",9,8.9,0.1,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P014","Placebo","Week 8",2,"N",8.5,8.9,-0.4,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P014","Placebo","Week 12",3,"N",9,8.9,0.1,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P015","Placebo","Baseline",0,"Y",7.6,7.6,NA,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P015","Placebo","Week 4",1,"N",6.7,7.6,-0.9,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P015","Placebo","Week 8",2,"N",6.5,7.6,-1.1,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"P015","Placebo","Week 12",3,"N",6.9,7.6,-0.7,1,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D001","DrugX","Baseline",0,"Y",8.1,8.1,NA,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D001","DrugX","Week 4",1,"N",7.3,8.1,-0.8,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D001","DrugX","Week 8",2,"N",8.1,8.1,0,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D001","DrugX","Week 12",3,"N",7.4,8.1,-0.7,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D002","DrugX","Baseline",0,"Y",7.9,7.9,NA,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D002","DrugX","Week 4",1,"N",7.5,7.9,-0.4,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D002","DrugX","Week 8",2,"N",6.9,7.9,-1,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D002","DrugX","Week 12",3,"N",7.8,7.9,-0.1,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D003","DrugX","Baseline",0,"Y",7.6,7.6,NA,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D003","DrugX","Week 4",1,"N",7.8,7.6,0.2,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D003","DrugX","Week 8",2,"N",7.5,7.6,-0.1,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D003","DrugX","Week 12",3,"N",6.7,7.6,-0.9,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D004","DrugX","Baseline",0,"Y",7.5,7.5,NA,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D004","DrugX","Week 4",1,"N",7.4,7.5,-0.1,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D004","DrugX","Week 8",2,"N",7.3,7.5,-0.2,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D004","DrugX","Week 12",3,"N",7.3,7.5,-0.2,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D005","DrugX","Baseline",0,"Y",8.7,8.7,NA,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D005","DrugX","Week 4",1,"N",7.6,8.7,-1.1,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D005","DrugX","Week 8",2,"N",8,8.7,-0.7,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D005","DrugX","Week 12",3,"N",7.7,8.7,-1,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D006","DrugX","Baseline",0,"Y",8.8,8.8,NA,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D006","DrugX","Week 4",1,"N",8.5,8.8,-0.3,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D006","DrugX","Week 8",2,"N",8.1,8.8,-0.7,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D006","DrugX","Week 12",3,"N",7.7,8.8,-1.1,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D007","DrugX","Baseline",0,"Y",8,8,NA,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D007","DrugX","Week 4",1,"N",7.3,8,-0.7,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D007","DrugX","Week 8",2,"N",7.8,8,-0.2,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D007","DrugX","Week 12",3,"N",7.7,8,-0.3,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D008","DrugX","Baseline",0,"Y",8.8,8.8,NA,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D008","DrugX","Week 4",1,"N",8.3,8.8,-0.5,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D008","DrugX","Week 8",2,"N",7.8,8.8,-1,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D008","DrugX","Week 12",3,"N",8.6,8.8,-0.2,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D009","DrugX","Baseline",0,"Y",8.6,8.6,NA,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D009","DrugX","Week 4",1,"N",8.2,8.6,-0.4,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D009","DrugX","Week 8",2,"N",8.5,8.6,-0.1,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D009","DrugX","Week 12",3,"N",8.1,8.6,-0.5,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D010","DrugX","Baseline",0,"Y",8.3,8.3,NA,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D010","DrugX","Week 4",1,"N",7.7,8.3,-0.6,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D010","DrugX","Week 8",2,"N",7.1,8.3,-1.2,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D010","DrugX","Week 12",3,"N",7.3,8.3,-1,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D011","DrugX","Baseline",0,"Y",7.5,7.5,NA,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D011","DrugX","Week 4",1,"N",7.2,7.5,-0.3,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D011","DrugX","Week 8",2,"N",6.7,7.5,-0.8,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D011","DrugX","Week 12",3,"N",7,7.5,-0.5,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D012","DrugX","Baseline",0,"Y",8.9,8.9,NA,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D012","DrugX","Week 4",1,"N",8,8.9,-0.9,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D012","DrugX","Week 8",2,"N",8.3,8.9,-0.6,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D012","DrugX","Week 12",3,"N",8.8,8.9,-0.1,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D013","DrugX","Baseline",0,"Y",7.8,7.8,NA,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D013","DrugX","Week 4",1,"N",6.7,7.8,-1.1,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D013","DrugX","Week 8",2,"N",7,7.8,-0.8,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D013","DrugX","Week 12",3,"N",6.8,7.8,-1,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D014","DrugX","Baseline",0,"Y",8.9,8.9,NA,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D014","DrugX","Week 4",1,"N",8.8,8.9,-0.1,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D014","DrugX","Week 8",2,"N",8.6,8.9,-0.3,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D014","DrugX","Week 12",3,"N",8.9,8.9,0,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D015","DrugX","Baseline",0,"Y",8.7,8.7,NA,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D015","DrugX","Week 4",1,"N",7.8,8.7,-0.9,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D015","DrugX","Week 8",2,"N",8.7,8.7,0,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
"D015","DrugX","Week 12",3,"N",8.3,8.7,-0.4,2,"Y","HBA1C","Hemoglobin A1C (%)",1,
)

# Analysis and outputs.

suppressPackageStartupMessages(library(nlme));maps<-make_id_maps(list(ADLB=adlb));adlb<-double_subjects(adlb,maps);adlb<-adlb%>%group_by(usubjid)%>%mutate(base=first(base),aval=ifelse(avisitn==0,base,aval+rnorm(n(),0,.12)),chg=ifelse(avisitn==0,NA,aval-base))%>%ungroup();stopifnot(n_distinct(adlb$usubjid)==2L*maps$original_n);write_xlsx_list(list(ADLB=adlb),file.path(OUT_DIR,paste0(STEM,"_simulated_data.xlsx")))
z<-adlb%>%filter(paramcd=="HBA1C",saffl=="Y",avisitn>0)%>%mutate(treatment=factor(trtpn),visit=factor(avisitn));model<-lme(chg~base+treatment*visit,random=~1|usubjid,data=z,method="REML",na.action=na.exclude);V<-vcov(model);b<-fixef(model);ref<-mean(z$base,na.rm=TRUE);levv<-levels(z$visit);rows<-list();for(v in levv){nd<-data.frame(base=ref,treatment=factor(1:2,levels=1:2),visit=factor(v,levels=levv));X<-model.matrix(~base+treatment*visit,nd);est<-drop(X%*%b);se<-sqrt(diag(X%*%V%*%t(X)));for(g in 1:2)rows[[length(rows)+1]]<-data.frame(Visit=unique(z$avisit[z$avisitn==as.numeric(v)])[1],Treatment=c("Placebo","Drug X")[g],N=sum(z$avisitn==as.numeric(v)&z$trtpn==g),`LS Mean (SE)`=sprintf("%.2f (%.3f)",est[g],se[g]),`95% CI`=sprintf("(%.2f, %.2f)",est[g]-1.96*se[g],est[g]+1.96*se[g]),`Difference vs Placebo`=ifelse(g==1,"-",sprintf("%.2f",est[2]-est[1])),`p-value`=ifelse(g==1,"-",fmt_p(2*pnorm(-abs((est[2]-est[1])/sqrt(drop((X[2,]-X[1,])%*%V%*%(X[2,]-X[1,]))))))),check.names=FALSE)};output<-bind_rows(rows)
write_clinical_rtf(output,"Table 8. Repeated-Measures Mixed Model for Change in HbA1c",file.path(OUT_DIR,paste0(STEM,".rtf")),"Model includes baseline, treatment, visit and treatment-by-visit interaction with a subject-level random intercept.")

cat("Created 08_repeated_measures_mixed_model: simulated N =", 2L*maps$original_n, "\n")
