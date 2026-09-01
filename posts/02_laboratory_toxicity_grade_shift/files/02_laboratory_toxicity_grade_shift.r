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
adsl<-tribble(
~usubjid,~saffl,~trt01an,
1001,"Y",2,
1002,"Y",1,
1003,"Y",2,
1004,"Y",2,
1005,"Y",2,
1006,"Y",1,
1007,"Y",2,
1008,"Y",1,
1009,"Y",2,
1010,"Y",2,
)

adlb<-tribble(
~usubjid,~saffl,~trt01an,~paramcd,~paramn,~param,~avisitn,~atoxgrn,~ablfl,~btoxgrn,~anl02fl,~avisit,~shift1,~shift2,~shift1n,
1001,"Y",2,"PARAM 01",1,"Param 01",0,2,"Y",2,"","Baseline","","",NA,
1001,"Y",2,"PARAM 01",1,"Param 01",2,1,"",2,"","Week 2","","",NA,
1001,"Y",2,"PARAM 01",1,"Param 01",4,2,"",2,"Y","Week 4","No change from baseline","",1,
1001,"Y",2,"PARAM 01",1,"Param 01",6,2,"",2,"","Week 6","","",NA,
1001,"Y",2,"PARAM 02",2,"Param 02",0,3,"Y",3,"","Baseline","","",NA,
1001,"Y",2,"PARAM 02",2,"Param 02",2,1,"",3,"","Week 2","","",NA,
1001,"Y",2,"PARAM 02",2,"Param 02",4,2,"",3,"Y","Week 4","Improved from baseline","",7,
1001,"Y",2,"PARAM 02",2,"Param 02",6,0,"",3,"","Week 6","","",NA,
1001,"Y",2,"PARAM 03",3,"Param 03",0,3,"Y",3,"","Baseline","","",NA,
1001,"Y",2,"PARAM 03",3,"Param 03",2,3,"",3,"Y","Week 2","No change from baseline","",1,
1001,"Y",2,"PARAM 03",3,"Param 03",4,1,"",3,"","Week 4","","",NA,
1001,"Y",2,"PARAM 03",3,"Param 03",6,1,"",3,"","Week 6","","",NA,
1001,"Y",2,"PARAM 04",4,"Param 04",0,2,"Y",2,"","Baseline","","",NA,
1001,"Y",2,"PARAM 04",4,"Param 04",2,3,"",2,"Y","Week 2","Any worsening from baseline","Worsening to Grade 3",2,
1001,"Y",2,"PARAM 04",4,"Param 04",4,2,"",2,"","Week 4","","",NA,
1001,"Y",2,"PARAM 04",4,"Param 04",6,2,"",2,"","Week 6","","",NA,
1001,"Y",2,"PARAM 05",5,"Param 05",0,1,"Y",1,"","Baseline","","",NA,
1001,"Y",2,"PARAM 05",5,"Param 05",2,3,"",1,"Y","Week 2","Any worsening from baseline","Worsening to Grade 3",2,
1001,"Y",2,"PARAM 05",5,"Param 05",4,0,"",1,"","Week 4","","",NA,
1001,"Y",2,"PARAM 05",5,"Param 05",6,3,"",1,"","Week 6","","",NA,
1002,"Y",1,"PARAM 01",1,"Param 01",0,2,"Y",2,"","Baseline","","",NA,
1002,"Y",1,"PARAM 01",1,"Param 01",2,0,"",2,"","Week 2","","",NA,
1002,"Y",1,"PARAM 01",1,"Param 01",4,2,"",2,"","Week 4","","",NA,
1002,"Y",1,"PARAM 01",1,"Param 01",6,3,"",2,"Y","Week 6","Any worsening from baseline","Worsening to Grade 3",2,
1002,"Y",1,"PARAM 02",2,"Param 02",0,1,"Y",1,"","Baseline","","",NA,
1002,"Y",1,"PARAM 02",2,"Param 02",2,3,"",1,"Y","Week 2","Any worsening from baseline","Worsening to Grade 3",2,
1002,"Y",1,"PARAM 02",2,"Param 02",4,3,"",1,"","Week 4","","",NA,
1002,"Y",1,"PARAM 02",2,"Param 02",6,3,"",1,"","Week 6","","",NA,
1002,"Y",1,"PARAM 03",3,"Param 03",0,0,"Y",0,"","Baseline","","",NA,
1002,"Y",1,"PARAM 03",3,"Param 03",2,2,"",0,"Y","Week 2","Any worsening from baseline","Worsening to < Grade 3",2,
1002,"Y",1,"PARAM 03",3,"Param 03",4,1,"",0,"","Week 4","","",NA,
1002,"Y",1,"PARAM 03",3,"Param 03",6,0,"",0,"","Week 6","","",NA,
1002,"Y",1,"PARAM 04",4,"Param 04",0,3,"Y",3,"","Baseline","","",NA,
1002,"Y",1,"PARAM 04",4,"Param 04",2,1,"",3,"","Week 2","","",NA,
1002,"Y",1,"PARAM 04",4,"Param 04",4,2,"",3,"Y","Week 4","Improved from baseline","",7,
1002,"Y",1,"PARAM 04",4,"Param 04",6,2,"",3,"","Week 6","","",NA,
1002,"Y",1,"PARAM 05",5,"Param 05",0,2,"Y",2,"","Baseline","","",NA,
1002,"Y",1,"PARAM 05",5,"Param 05",2,0,"",2,"","Week 2","","",NA,
1002,"Y",1,"PARAM 05",5,"Param 05",4,0,"",2,"","Week 4","","",NA,
1002,"Y",1,"PARAM 05",5,"Param 05",6,1,"",2,"Y","Week 6","Improved from baseline","",7,
1003,"Y",2,"PARAM 01",1,"Param 01",0,3,"Y",3,"","Baseline","","",NA,
1003,"Y",2,"PARAM 01",1,"Param 01",2,2,"",3,"","Week 2","","",NA,
1003,"Y",2,"PARAM 01",1,"Param 01",4,3,"",3,"Y","Week 4","No change from baseline","",1,
1003,"Y",2,"PARAM 01",1,"Param 01",6,2,"",3,"","Week 6","","",NA,
1003,"Y",2,"PARAM 02",2,"Param 02",0,2,"Y",2,"","Baseline","","",NA,
1003,"Y",2,"PARAM 02",2,"Param 02",2,0,"",2,"","Week 2","","",NA,
1003,"Y",2,"PARAM 02",2,"Param 02",4,2,"",2,"Y","Week 4","No change from baseline","",1,
1003,"Y",2,"PARAM 02",2,"Param 02",6,1,"",2,"","Week 6","","",NA,
1003,"Y",2,"PARAM 03",3,"Param 03",0,0,"Y",0,"","Baseline","","",NA,
1003,"Y",2,"PARAM 03",3,"Param 03",2,1,"",0,"","Week 2","","",NA,
1003,"Y",2,"PARAM 03",3,"Param 03",4,3,"",0,"Y","Week 4","Any worsening from baseline","Worsening to Grade 3",2,
1003,"Y",2,"PARAM 03",3,"Param 03",6,3,"",0,"","Week 6","","",NA,
1003,"Y",2,"PARAM 04",4,"Param 04",0,0,"Y",0,"","Baseline","","",NA,
1003,"Y",2,"PARAM 04",4,"Param 04",2,2,"",0,"Y","Week 2","Any worsening from baseline","Worsening to < Grade 3",2,
1003,"Y",2,"PARAM 04",4,"Param 04",4,2,"",0,"","Week 4","","",NA,
1003,"Y",2,"PARAM 04",4,"Param 04",6,0,"",0,"","Week 6","","",NA,
1003,"Y",2,"PARAM 05",5,"Param 05",0,3,"Y",3,"","Baseline","","",NA,
1003,"Y",2,"PARAM 05",5,"Param 05",2,3,"",3,"Y","Week 2","No change from baseline","",1,
1003,"Y",2,"PARAM 05",5,"Param 05",4,3,"",3,"","Week 4","","",NA,
1003,"Y",2,"PARAM 05",5,"Param 05",6,1,"",3,"","Week 6","","",NA,
1004,"Y",2,"PARAM 01",1,"Param 01",0,1,"Y",1,"","Baseline","","",NA,
1004,"Y",2,"PARAM 01",1,"Param 01",2,0,"",1,"","Week 2","","",NA,
1004,"Y",2,"PARAM 01",1,"Param 01",4,3,"",1,"Y","Week 4","Any worsening from baseline","Worsening to Grade 3",2,
1004,"Y",2,"PARAM 01",1,"Param 01",6,0,"",1,"","Week 6","","",NA,
1004,"Y",2,"PARAM 02",2,"Param 02",0,1,"Y",1,"","Baseline","","",NA,
1004,"Y",2,"PARAM 02",2,"Param 02",2,3,"",1,"Y","Week 2","Any worsening from baseline","Worsening to Grade 3",2,
1004,"Y",2,"PARAM 02",2,"Param 02",4,3,"",1,"","Week 4","","",NA,
1004,"Y",2,"PARAM 02",2,"Param 02",6,3,"",1,"","Week 6","","",NA,
1004,"Y",2,"PARAM 03",3,"Param 03",0,0,"Y",0,"","Baseline","","",NA,
1004,"Y",2,"PARAM 03",3,"Param 03",2,3,"",0,"Y","Week 2","Any worsening from baseline","Worsening to Grade 3",2,
1004,"Y",2,"PARAM 03",3,"Param 03",4,1,"",0,"","Week 4","","",NA,
1004,"Y",2,"PARAM 03",3,"Param 03",6,1,"",0,"","Week 6","","",NA,
1004,"Y",2,"PARAM 04",4,"Param 04",0,0,"Y",0,"","Baseline","","",NA,
1004,"Y",2,"PARAM 04",4,"Param 04",2,1,"",0,"","Week 2","","",NA,
1004,"Y",2,"PARAM 04",4,"Param 04",4,3,"",0,"Y","Week 4","Any worsening from baseline","Worsening to Grade 3",2,
1004,"Y",2,"PARAM 04",4,"Param 04",6,0,"",0,"","Week 6","","",NA,
1004,"Y",2,"PARAM 05",5,"Param 05",0,0,"Y",0,"","Baseline","","",NA,
1004,"Y",2,"PARAM 05",5,"Param 05",2,3,"",0,"Y","Week 2","Any worsening from baseline","Worsening to Grade 3",2,
1004,"Y",2,"PARAM 05",5,"Param 05",4,1,"",0,"","Week 4","","",NA,
1004,"Y",2,"PARAM 05",5,"Param 05",6,3,"",0,"","Week 6","","",NA,
1005,"Y",2,"PARAM 01",1,"Param 01",0,2,"Y",2,"","Baseline","","",NA,
1005,"Y",2,"PARAM 01",1,"Param 01",2,1,"",2,"","Week 2","","",NA,
1005,"Y",2,"PARAM 01",1,"Param 01",4,0,"",2,"","Week 4","","",NA,
1005,"Y",2,"PARAM 01",1,"Param 01",6,3,"",2,"Y","Week 6","Any worsening from baseline","Worsening to Grade 3",2,
1005,"Y",2,"PARAM 02",2,"Param 02",0,1,"Y",1,"","Baseline","","",NA,
1005,"Y",2,"PARAM 02",2,"Param 02",2,2,"",1,"Y","Week 2","Any worsening from baseline","Worsening to < Grade 3",2,
1005,"Y",2,"PARAM 02",2,"Param 02",4,1,"",1,"","Week 4","","",NA,
1005,"Y",2,"PARAM 02",2,"Param 02",6,1,"",1,"","Week 6","","",NA,
1005,"Y",2,"PARAM 03",3,"Param 03",0,1,"Y",1,"","Baseline","","",NA,
1005,"Y",2,"PARAM 03",3,"Param 03",2,1,"",1,"Y","Week 2","No change from baseline","",1,
1005,"Y",2,"PARAM 03",3,"Param 03",4,1,"",1,"","Week 4","","",NA,
1005,"Y",2,"PARAM 03",3,"Param 03",6,0,"",1,"","Week 6","","",NA,
1005,"Y",2,"PARAM 04",4,"Param 04",0,0,"Y",0,"","Baseline","","",NA,
1005,"Y",2,"PARAM 04",4,"Param 04",2,1,"",0,"Y","Week 2","Any worsening from baseline","Worsening to < Grade 3",2,
1005,"Y",2,"PARAM 04",4,"Param 04",4,1,"",0,"","Week 4","","",NA,
1005,"Y",2,"PARAM 04",4,"Param 04",6,1,"",0,"","Week 6","","",NA,
1005,"Y",2,"PARAM 05",5,"Param 05",0,0,"Y",0,"","Baseline","","",NA,
1005,"Y",2,"PARAM 05",5,"Param 05",2,3,"",0,"Y","Week 2","Any worsening from baseline","Worsening to Grade 3",2,
1005,"Y",2,"PARAM 05",5,"Param 05",4,0,"",0,"","Week 4","","",NA,
1005,"Y",2,"PARAM 05",5,"Param 05",6,3,"",0,"","Week 6","","",NA,
1006,"Y",1,"PARAM 01",1,"Param 01",0,1,"Y",1,"","Baseline","","",NA,
1006,"Y",1,"PARAM 01",1,"Param 01",2,3,"",1,"Y","Week 2","Any worsening from baseline","Worsening to Grade 3",2,
1006,"Y",1,"PARAM 01",1,"Param 01",4,0,"",1,"","Week 4","","",NA,
1006,"Y",1,"PARAM 01",1,"Param 01",6,1,"",1,"","Week 6","","",NA,
1006,"Y",1,"PARAM 02",2,"Param 02",0,3,"Y",3,"","Baseline","","",NA,
1006,"Y",1,"PARAM 02",2,"Param 02",2,0,"",3,"","Week 2","","",NA,
1006,"Y",1,"PARAM 02",2,"Param 02",4,2,"",3,"Y","Week 4","Improved from baseline","",7,
1006,"Y",1,"PARAM 02",2,"Param 02",6,1,"",3,"","Week 6","","",NA,
1006,"Y",1,"PARAM 03",3,"Param 03",0,1,"Y",1,"","Baseline","","",NA,
1006,"Y",1,"PARAM 03",3,"Param 03",2,0,"",1,"","Week 2","","",NA,
1006,"Y",1,"PARAM 03",3,"Param 03",4,3,"",1,"Y","Week 4","Any worsening from baseline","Worsening to Grade 3",2,
1006,"Y",1,"PARAM 03",3,"Param 03",6,3,"",1,"","Week 6","","",NA,
1006,"Y",1,"PARAM 04",4,"Param 04",0,3,"Y",3,"","Baseline","","",NA,
1006,"Y",1,"PARAM 04",4,"Param 04",2,0,"",3,"","Week 2","","",NA,
1006,"Y",1,"PARAM 04",4,"Param 04",4,3,"",3,"Y","Week 4","No change from baseline","",1,
1006,"Y",1,"PARAM 04",4,"Param 04",6,1,"",3,"","Week 6","","",NA,
1006,"Y",1,"PARAM 05",5,"Param 05",0,1,"Y",1,"","Baseline","","",NA,
1006,"Y",1,"PARAM 05",5,"Param 05",2,2,"",1,"Y","Week 2","Any worsening from baseline","Worsening to < Grade 3",2,
1006,"Y",1,"PARAM 05",5,"Param 05",4,2,"",1,"","Week 4","","",NA,
1006,"Y",1,"PARAM 05",5,"Param 05",6,0,"",1,"","Week 6","","",NA,
1007,"Y",2,"PARAM 01",1,"Param 01",0,3,"Y",3,"","Baseline","","",NA,
1007,"Y",2,"PARAM 01",1,"Param 01",2,0,"",3,"","Week 2","","",NA,
1007,"Y",2,"PARAM 01",1,"Param 01",4,2,"",3,"Y","Week 4","Improved from baseline","",7,
1007,"Y",2,"PARAM 01",1,"Param 01",6,2,"",3,"","Week 6","","",NA,
1007,"Y",2,"PARAM 02",2,"Param 02",0,1,"Y",1,"","Baseline","","",NA,
1007,"Y",2,"PARAM 02",2,"Param 02",2,1,"",1,"","Week 2","","",NA,
1007,"Y",2,"PARAM 02",2,"Param 02",4,2,"",1,"","Week 4","","",NA,
1007,"Y",2,"PARAM 02",2,"Param 02",6,3,"",1,"Y","Week 6","Any worsening from baseline","Worsening to Grade 3",2,
1007,"Y",2,"PARAM 03",3,"Param 03",0,2,"Y",2,"","Baseline","","",NA,
1007,"Y",2,"PARAM 03",3,"Param 03",2,0,"",2,"","Week 2","","",NA,
1007,"Y",2,"PARAM 03",3,"Param 03",4,2,"",2,"Y","Week 4","No change from baseline","",1,
1007,"Y",2,"PARAM 03",3,"Param 03",6,2,"",2,"","Week 6","","",NA,
1007,"Y",2,"PARAM 04",4,"Param 04",0,0,"Y",0,"","Baseline","","",NA,
1007,"Y",2,"PARAM 04",4,"Param 04",2,0,"",0,"","Week 2","","",NA,
1007,"Y",2,"PARAM 04",4,"Param 04",4,3,"",0,"Y","Week 4","Any worsening from baseline","Worsening to Grade 3",2,
1007,"Y",2,"PARAM 04",4,"Param 04",6,1,"",0,"","Week 6","","",NA,
1007,"Y",2,"PARAM 05",5,"Param 05",0,0,"Y",0,"","Baseline","","",NA,
1007,"Y",2,"PARAM 05",5,"Param 05",2,2,"",0,"Y","Week 2","Any worsening from baseline","Worsening to < Grade 3",2,
1007,"Y",2,"PARAM 05",5,"Param 05",4,2,"",0,"","Week 4","","",NA,
1007,"Y",2,"PARAM 05",5,"Param 05",6,1,"",0,"","Week 6","","",NA,
1008,"Y",1,"PARAM 01",1,"Param 01",0,2,"Y",2,"","Baseline","","",NA,
1008,"Y",1,"PARAM 01",1,"Param 01",2,2,"",2,"Y","Week 2","No change from baseline","",1,
1008,"Y",1,"PARAM 01",1,"Param 01",4,2,"",2,"","Week 4","","",NA,
1008,"Y",1,"PARAM 01",1,"Param 01",6,1,"",2,"","Week 6","","",NA,
1008,"Y",1,"PARAM 02",2,"Param 02",0,0,"Y",0,"","Baseline","","",NA,
1008,"Y",1,"PARAM 02",2,"Param 02",2,3,"",0,"Y","Week 2","Any worsening from baseline","Worsening to Grade 3",2,
1008,"Y",1,"PARAM 02",2,"Param 02",4,3,"",0,"","Week 4","","",NA,
1008,"Y",1,"PARAM 02",2,"Param 02",6,3,"",0,"","Week 6","","",NA,
1008,"Y",1,"PARAM 03",3,"Param 03",0,0,"Y",0,"","Baseline","","",NA,
1008,"Y",1,"PARAM 03",3,"Param 03",2,2,"",0,"","Week 2","","",NA,
1008,"Y",1,"PARAM 03",3,"Param 03",4,0,"",0,"","Week 4","","",NA,
1008,"Y",1,"PARAM 03",3,"Param 03",6,3,"",0,"Y","Week 6","Any worsening from baseline","Worsening to Grade 3",2,
1008,"Y",1,"PARAM 04",4,"Param 04",0,2,"Y",2,"","Baseline","","",NA,
1008,"Y",1,"PARAM 04",4,"Param 04",2,2,"",2,"Y","Week 2","No change from baseline","",1,
1008,"Y",1,"PARAM 04",4,"Param 04",4,1,"",2,"","Week 4","","",NA,
1008,"Y",1,"PARAM 04",4,"Param 04",6,0,"",2,"","Week 6","","",NA,
1008,"Y",1,"PARAM 05",5,"Param 05",0,1,"Y",1,"","Baseline","","",NA,
1008,"Y",1,"PARAM 05",5,"Param 05",2,3,"",1,"Y","Week 2","Any worsening from baseline","Worsening to Grade 3",2,
1008,"Y",1,"PARAM 05",5,"Param 05",4,0,"",1,"","Week 4","","",NA,
1008,"Y",1,"PARAM 05",5,"Param 05",6,2,"",1,"","Week 6","","",NA,
1009,"Y",2,"PARAM 01",1,"Param 01",0,1,"Y",1,"","Baseline","","",NA,
1009,"Y",2,"PARAM 01",1,"Param 01",2,1,"",1,"","Week 2","","",NA,
1009,"Y",2,"PARAM 01",1,"Param 01",4,0,"",1,"","Week 4","","",NA,
1009,"Y",2,"PARAM 01",1,"Param 01",6,2,"",1,"Y","Week 6","Any worsening from baseline","Worsening to < Grade 3",2,
1009,"Y",2,"PARAM 02",2,"Param 02",0,3,"Y",3,"","Baseline","","",NA,
1009,"Y",2,"PARAM 02",2,"Param 02",2,1,"",3,"","Week 2","","",NA,
1009,"Y",2,"PARAM 02",2,"Param 02",4,3,"",3,"Y","Week 4","No change from baseline","",1,
1009,"Y",2,"PARAM 02",2,"Param 02",6,1,"",3,"","Week 6","","",NA,
1009,"Y",2,"PARAM 03",3,"Param 03",0,2,"Y",2,"","Baseline","","",NA,
1009,"Y",2,"PARAM 03",3,"Param 03",2,2,"",2,"Y","Week 2","No change from baseline","",1,
1009,"Y",2,"PARAM 03",3,"Param 03",4,1,"",2,"","Week 4","","",NA,
1009,"Y",2,"PARAM 03",3,"Param 03",6,1,"",2,"","Week 6","","",NA,
1009,"Y",2,"PARAM 04",4,"Param 04",0,3,"Y",3,"","Baseline","","",NA,
1009,"Y",2,"PARAM 04",4,"Param 04",2,0,"",3,"","Week 2","","",NA,
1009,"Y",2,"PARAM 04",4,"Param 04",4,0,"",3,"","Week 4","","",NA,
1009,"Y",2,"PARAM 04",4,"Param 04",6,1,"",3,"Y","Week 6","Improved from baseline","",7,
1009,"Y",2,"PARAM 05",5,"Param 05",0,0,"Y",0,"","Baseline","","",NA,
1009,"Y",2,"PARAM 05",5,"Param 05",2,2,"",0,"Y","Week 2","Any worsening from baseline","Worsening to < Grade 3",2,
1009,"Y",2,"PARAM 05",5,"Param 05",4,0,"",0,"","Week 4","","",NA,
1009,"Y",2,"PARAM 05",5,"Param 05",6,0,"",0,"","Week 6","","",NA,
1010,"Y",2,"PARAM 01",1,"Param 01",0,1,"Y",1,"","Baseline","","",NA,
1010,"Y",2,"PARAM 01",1,"Param 01",2,3,"",1,"Y","Week 2","Any worsening from baseline","Worsening to Grade 3",2,
1010,"Y",2,"PARAM 01",1,"Param 01",4,0,"",1,"","Week 4","","",NA,
1010,"Y",2,"PARAM 01",1,"Param 01",6,2,"",1,"","Week 6","","",NA,
1010,"Y",2,"PARAM 02",2,"Param 02",0,1,"Y",1,"","Baseline","","",NA,
1010,"Y",2,"PARAM 02",2,"Param 02",2,1,"",1,"","Week 2","","",NA,
1010,"Y",2,"PARAM 02",2,"Param 02",4,1,"",1,"","Week 4","","",NA,
1010,"Y",2,"PARAM 02",2,"Param 02",6,3,"",1,"Y","Week 6","Any worsening from baseline","Worsening to Grade 3",2,
1010,"Y",2,"PARAM 03",3,"Param 03",0,3,"Y",3,"","Baseline","","",NA,
1010,"Y",2,"PARAM 03",3,"Param 03",2,3,"",3,"Y","Week 2","No change from baseline","",1,
1010,"Y",2,"PARAM 03",3,"Param 03",4,2,"",3,"","Week 4","","",NA,
1010,"Y",2,"PARAM 03",3,"Param 03",6,3,"",3,"","Week 6","","",NA,
1010,"Y",2,"PARAM 04",4,"Param 04",0,1,"Y",1,"","Baseline","","",NA,
1010,"Y",2,"PARAM 04",4,"Param 04",2,0,"",1,"Y","Week 2","Improved from baseline","",7,
1010,"Y",2,"PARAM 04",4,"Param 04",4,0,"",1,"","Week 4","","",NA,
1010,"Y",2,"PARAM 04",4,"Param 04",6,0,"",1,"","Week 6","","",NA,
1010,"Y",2,"PARAM 05",5,"Param 05",0,3,"Y",3,"","Baseline","","",NA,
1010,"Y",2,"PARAM 05",5,"Param 05",2,1,"",3,"","Week 2","","",NA,
1010,"Y",2,"PARAM 05",5,"Param 05",4,2,"",3,"Y","Week 4","Improved from baseline","",7,
1010,"Y",2,"PARAM 05",5,"Param 05",6,1,"",3,"","Week 6","","",NA,
)


# Simulate twice the original subject count while preserving treatment,
# parameter and baseline/post-baseline toxicity-grade relationships.
maps <- make_id_maps(list(ADSL=adsl, ADLB=adlb))
adsl <- double_subjects(adsl, maps)
adlb <- double_subjects(adlb, maps)
eligible <- which(!is.na(adlb$atoxgrn) & adlb$avisitn > 0)
change <- eligible[runif(length(eligible)) < 0.05]
if(length(change)) adlb$atoxgrn[change] <- pmin(5, pmax(0, adlb$atoxgrn[change] + sample(c(-1,1),length(change),replace=TRUE)))
stopifnot(n_distinct(adsl$usubjid) == 2L * maps$original_n)
write_xlsx_list(list(ADSL=adsl, ADLB=adlb), file.path(OUT_DIR, paste0(STEM,"_simulated_data.xlsx")))

# Original analysis structure retained from the source document.
# Filter adlb01 dataset
adlb01 <- adlb %>%
 filter(anl02fl == "Y" & !is.na(atoxgrn) & !is.na(btoxgrn) & saffl == "Y")
# Create a variable named 'treatment' and duplicate rows
adlb01 <- adlb01 %>%
 mutate(
 treatment = trt01an
 ) %>%
 bind_rows(adlb01 %>% mutate(treatment = 3))
adsl01 <- adsl %>%
 mutate(
 treatment = trt01an
 ) %>%
 bind_rows(adsl %>% mutate(treatment = 3))
# Duplicate records for total atoxgrn and btoxgrn
adlb02 <- adlb01 %>%
 bind_rows(adlb01 %>% filter(!is.na(atoxgrn)) %>% mutate(atoxgrn = 6))
adlb03 <- adlb02 %>%
 bind_rows(adlb02 %>% filter(!is.na(btoxgrn)) %>% mutate(btoxgrn = 6))
# Get treatment totals
denoms <- adsl01 %>%
 group_by(treatment) %>%
summarise(denom = n_distinct(usubjid), .groups="drop")
# Get actual counts in atoxgrn vs btoxgrn
counts01 <- adlb03 %>%
 group_by(paramn, paramcd, param, treatment, atoxgrn, btoxgrn) %>%
summarise(count = n_distinct(usubjid), .groups="drop")
# Create a dummy to get all levels
dummy01 <- adlb03 %>%
 distinct(paramn, paramcd, param)
dummy02 <- dummy01 %>%
 expand_grid(
 treatment = 1:3,
 atoxgrn = 0:6,
 btoxgrn = 0:6
 )
# Merge dummy and actual counts to have all levels in the table
counts02 <- dummy02 %>%
 left_join(counts01, by = c("paramn", "paramcd", "param", "treatment", "atoxgrn", "btoxgrn")) %>%
 mutate(count = coalesce(count, 0))
# Fetch treatment totals into counts dataset
counts03_pre <- counts02 %>%
 left_join(denoms, by = "treatment")
# Fetch total to total row counts as a column to use as a denominator for calculating percentages
altdenoms <- counts02 %>%
 filter(atoxgrn == 6 & btoxgrn == 6) %>%
 select(paramn, paramcd, param, treatment, count) %>%
 rename(altdenom = count)
counts03 <- counts03_pre %>%
 left_join(altdenoms, by = c("paramn","paramcd", "param", "treatment"))
# Calculate percentages
counts04 <- counts03 %>%
 mutate(
 cp = ifelse(count != 0, paste0(count, " (", sprintf("%.1f", count / altdenom * 100), ")"), " 0")
 ) %>%
 mutate(
 label = case_when(
 treatment == 1 ~ paste0("ARM A (N=", denom, ")"),
 treatment == 2 ~ paste0("ARM B (N=", denom, ")"),
 treatment == 3 ~ paste0("Total (N=", denom, ")")
 )
 )
# Transpose the data so that post-baseline toxicity levels become columns
trans01 <- counts04 %>%
 pivot_wider(
 id_cols = c("paramn", "param", "treatment", "label", "btoxgrn"),
 names_prefix = "tox",
 names_from = c( "atoxgrn"),
 values_from = cp
 )
output <- trans01

grade_label <- c("Grade 0","Grade 1","Grade 2","Grade 3","Grade 4","Grade 5","Total")
display <- output %>%
  arrange(paramn, treatment, btoxgrn) %>%
  mutate(`Baseline Toxicity Grade`=grade_label[btoxgrn+1]) %>%
  mutate(
    parameter_display=if_else(
      row_number()==1L | param!=lag(param),
      param,
      ""
    ),
    treatment_display=if_else(
      row_number()==1L | param!=lag(param) | label!=lag(label),
      label,
      ""
    )
  ) %>%
  transmute(
    Parameter=parameter_display,
    Treatment=treatment_display,
    `Baseline Toxicity Grade`,
    `Post Grade 0`=coalesce(tox0,"0"),
    `Post Grade 1`=coalesce(tox1,"0"),
    `Post Grade 2`=coalesce(tox2,"0"),
    `Post Grade 3`=coalesce(tox3,"0"),
    `Post Grade 4`=coalesce(tox4,"0"),
    `Post Grade 5`=coalesce(tox5,"0"),
    Total=coalesce(tox6,"0")
  )
write_clinical_rtf(display,
  "Table 2. Shift from Baseline in Laboratory Toxicity Grade",
  file.path(OUT_DIR,paste0(STEM,".rtf")),
  "Cells show n (percent). Percentages use subjects with evaluable baseline and post-baseline toxicity grades."
)
cat("Created ",STEM,": simulated N = ",2L*maps$original_n,"\n",sep="")
