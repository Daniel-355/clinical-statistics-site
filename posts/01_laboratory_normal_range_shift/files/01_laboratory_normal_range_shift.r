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
~usubjid,~paramn,~paramcd,~param,~avisitn,~avisit,~anrind,~bnrind,~saffl,~trt01an,
"PILOT-1001",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","NORMAL","HIGH","Y",1,
"PILOT-1001",32,"NEUT","Neutrophils (10^6/L)",8,"Week 4","NORMAL","HIGH","Y",1,
"PILOT-1002",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","NORMAL","NORMAL","Y",2,
"PILOT-1002",32,"NEUT","Neutrophils (10^6/L)",8,"Week 4","NORMAL","NORMAL","Y",2,
"PILOT-1002",32,"NEUT","Neutrophils (10^6/L)",9,"Week 6","NORMAL","NORMAL","Y",2,
"PILOT-1002",32,"NEUT","Neutrophils (10^6/L)",10,"Week 8","NORMAL","NORMAL","Y",2,
"PILOT-1002",32,"NEUT","Neutrophils (10^6/L)",11,"Week 12","NORMAL","NORMAL","Y",2,
"PILOT-1002",32,"NEUT","Neutrophils (10^6/L)",12,"Week 16","NORMAL","NORMAL","Y",2,
"PILOT-1002",32,"NEUT","Neutrophils (10^6/L)",13,"Week 24","NORMAL","NORMAL","Y",2,
"PILOT-1003",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","NORMAL","NORMAL","Y",1,
"PILOT-1003",32,"NEUT","Neutrophils (10^6/L)",8,"Week 4","NORMAL","NORMAL","Y",1,
"PILOT-1003",32,"NEUT","Neutrophils (10^6/L)",9,"Week 6","NORMAL","NORMAL","Y",1,
"PILOT-1003",32,"NEUT","Neutrophils (10^6/L)",10,"Week 8","NORMAL","NORMAL","Y",1,
"PILOT-1003",32,"NEUT","Neutrophils (10^6/L)",11,"Week 12","NORMAL","NORMAL","Y",1,
"PILOT-1003",32,"NEUT","Neutrophils (10^6/L)",12,"Week 16","NORMAL","NORMAL","Y",1,
"PILOT-1003",32,"NEUT","Neutrophils (10^6/L)",13,"Week 24","NORMAL","NORMAL","Y",1,
"PILOT-1004",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","NORMAL","NORMAL","Y",2,
"PILOT-1004",32,"NEUT","Neutrophils (10^6/L)",8,"Week 4","NORMAL","NORMAL","Y",2,
"PILOT-1004",32,"NEUT","Neutrophils (10^6/L)",9,"Week 6","NORMAL","NORMAL","Y",2,
"PILOT-1004",32,"NEUT","Neutrophils (10^6/L)",10,"Week 8","NORMAL","NORMAL","Y",2,
"PILOT-1004",32,"NEUT","Neutrophils (10^6/L)",11,"Week 12","NORMAL","NORMAL","Y",2,
"PILOT-1004",32,"NEUT","Neutrophils (10^6/L)",12,"Week 16","NORMAL","NORMAL","Y",2,
"PILOT-1004",32,"NEUT","Neutrophils (10^6/L)",13,"Week 24","NORMAL","NORMAL","Y",2,
"PILOT-1005",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","NORMAL","NORMAL","Y",2,
"PILOT-1005",32,"NEUT","Neutrophils (10^6/L)",8,"Week 4","NORMAL","NORMAL","Y",2,
"PILOT-1005",32,"NEUT","Neutrophils (10^6/L)",9,"Week 6","NORMAL","NORMAL","Y",2,
"PILOT-1005",32,"NEUT","Neutrophils (10^6/L)",10,"Week 8","NORMAL","NORMAL","Y",2,
"PILOT-1005",32,"NEUT","Neutrophils (10^6/L)",11,"Week 12","NORMAL","NORMAL","Y",2,
"PILOT-1006",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","NORMAL","NORMAL","Y",1,
"PILOT-1006",32,"NEUT","Neutrophils (10^6/L)",8,"Week 4","NORMAL","NORMAL","Y",1,
"PILOT-1006",32,"NEUT","Neutrophils (10^6/L)",9,"Week 6","NORMAL","NORMAL","Y",1,
"PILOT-1006",32,"NEUT","Neutrophils (10^6/L)",10,"Week 8","LOW","NORMAL","Y",1,
"PILOT-1006",32,"NEUT","Neutrophils (10^6/L)",11,"Week 12","NORMAL","NORMAL","Y",1,
"PILOT-1006",32,"NEUT","Neutrophils (10^6/L)",12,"Week 16","NORMAL","NORMAL","Y",1,
"PILOT-1006",32,"NEUT","Neutrophils (10^6/L)",13,"Week 24","NORMAL","NORMAL","Y",1,
"PILOT-1007",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","NORMAL","NORMAL","Y",2,
"PILOT-1007",32,"NEUT","Neutrophils (10^6/L)",8,"Week 4","NORMAL","NORMAL","Y",2,
"PILOT-1007",32,"NEUT","Neutrophils (10^6/L)",9,"Week 6","NORMAL","NORMAL","Y",2,
"PILOT-1007",32,"NEUT","Neutrophils (10^6/L)",10,"Week 8","NORMAL","NORMAL","Y",2,
"PILOT-1007",32,"NEUT","Neutrophils (10^6/L)",11,"Week 12","NORMAL","NORMAL","Y",2,
"PILOT-1007",32,"NEUT","Neutrophils (10^6/L)",12,"Week 16","NORMAL","NORMAL","Y",2,
"PILOT-1007",32,"NEUT","Neutrophils (10^6/L)",13,"Week 24","HIGH","NORMAL","Y",2,
"PILOT-1008",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","HIGH","NORMAL","Y",2,
"PILOT-1008",32,"NEUT","Neutrophils (10^6/L)",8,"Week 4","NORMAL","NORMAL","Y",2,
"PILOT-1008",32,"NEUT","Neutrophils (10^6/L)",9,"Week 6","NORMAL","NORMAL","Y",2,
"PILOT-1008",32,"NEUT","Neutrophils (10^6/L)",10,"Week 8","NORMAL","NORMAL","Y",2,
"PILOT-1008",32,"NEUT","Neutrophils (10^6/L)",11,"Week 12","NORMAL","NORMAL","Y",2,
"PILOT-1008",32,"NEUT","Neutrophils (10^6/L)",12,"Week 16","NORMAL","NORMAL","Y",2,
"PILOT-1008",32,"NEUT","Neutrophils (10^6/L)",13,"Week 24","NORMAL","NORMAL","Y",2,
"PILOT-1009",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","NORMAL","NORMAL","Y",1,
"PILOT-1009",32,"NEUT","Neutrophils (10^6/L)",8,"Week 4","NORMAL","NORMAL","Y",1,
"PILOT-1009",32,"NEUT","Neutrophils (10^6/L)",9,"Week 6","NORMAL","NORMAL","Y",1,
"PILOT-1009",32,"NEUT","Neutrophils (10^6/L)",10,"Week 8","NORMAL","NORMAL","Y",1,
"PILOT-1009",32,"NEUT","Neutrophils (10^6/L)",11,"Week 12","NORMAL","NORMAL","Y",1,
"PILOT-1009",32,"NEUT","Neutrophils (10^6/L)",12,"Week 16","NORMAL","NORMAL","Y",1,
"PILOT-1010",32,"NEUT","Neutrophils (10^6/L)",8,"Week 4","NORMAL","NORMAL","Y",2,
"PILOT-1010",32,"NEUT","Neutrophils (10^6/L)",9,"Week 6","NORMAL","NORMAL","Y",2,
"PILOT-1010",32,"NEUT","Neutrophils (10^6/L)",10,"Week 8","NORMAL","NORMAL","Y",2,
"PILOT-1010",32,"NEUT","Neutrophils (10^6/L)",11,"Week 12","NORMAL","NORMAL","Y",2,
"PILOT-1010",32,"NEUT","Neutrophils (10^6/L)",12,"Week 16","NORMAL","NORMAL","Y",2,
"PILOT-1011",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","NORMAL","NORMAL","Y",1,
"PILOT-1011",32,"NEUT","Neutrophils (10^6/L)",8,"Week 4","NORMAL","NORMAL","Y",1,
"PILOT-1011",32,"NEUT","Neutrophils (10^6/L)",9,"Week 6","NORMAL","NORMAL","Y",1,
"PILOT-1011",32,"NEUT","Neutrophils (10^6/L)",10,"Week 8","NORMAL","NORMAL","Y",1,
"PILOT-1012",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","NORMAL","NORMAL","Y",2,
"PILOT-1012",32,"NEUT","Neutrophils (10^6/L)",8,"Week 4","NORMAL","NORMAL","Y",2,
"PILOT-1012",32,"NEUT","Neutrophils (10^6/L)",9,"Week 6","NORMAL","NORMAL","Y",2,
"PILOT-1012",32,"NEUT","Neutrophils (10^6/L)",10,"Week 8","NORMAL","NORMAL","Y",2,
"PILOT-1012",32,"NEUT","Neutrophils (10^6/L)",11,"Week 12","NORMAL","NORMAL","Y",2,
"PILOT-1012",32,"NEUT","Neutrophils (10^6/L)",12,"Week 16","NORMAL","NORMAL","Y",2,
"PILOT-1012",32,"NEUT","Neutrophils (10^6/L)",13,"Week 24","NORMAL","NORMAL","Y",2,
"PILOT-1013",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","NORMAL","NORMAL","Y",2,
"PILOT-1013",32,"NEUT","Neutrophils (10^6/L)",8,"Week 4","NORMAL","NORMAL","Y",2,
"PILOT-1013",32,"NEUT","Neutrophils (10^6/L)",9,"Week 6","NORMAL","NORMAL","Y",2,
"PILOT-1013",32,"NEUT","Neutrophils (10^6/L)",10,"Week 8","NORMAL","NORMAL","Y",2,
"PILOT-1013",32,"NEUT","Neutrophils (10^6/L)",11,"Week 12","NORMAL","NORMAL","Y",2,
"PILOT-1013",32,"NEUT","Neutrophils (10^6/L)",12,"Week 16","NORMAL","NORMAL","Y",2,
"PILOT-1014",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","HIGH","HIGH","Y",1,
"PILOT-1014",32,"NEUT","Neutrophils (10^6/L)",8,"Week 4","HIGH","HIGH","Y",1,
"PILOT-1014",32,"NEUT","Neutrophils (10^6/L)",9,"Week 6","NORMAL","HIGH","Y",1,
"PILOT-1014",32,"NEUT","Neutrophils (10^6/L)",10,"Week 8","NORMAL","HIGH","Y",1,
"PILOT-1014",32,"NEUT","Neutrophils (10^6/L)",11,"Week 12","HIGH","HIGH","Y",1,
"PILOT-1014",32,"NEUT","Neutrophils (10^6/L)",12,"Week 16","NORMAL","HIGH","Y",1,
"PILOT-1014",32,"NEUT","Neutrophils (10^6/L)",13,"Week 24","HIGH","HIGH","Y",1,
"PILOT-1016",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","NORMAL","NORMAL","Y",1,
"PILOT-1016",32,"NEUT","Neutrophils (10^6/L)",8,"Week 4","NORMAL","NORMAL","Y",1,
"PILOT-1016",32,"NEUT","Neutrophils (10^6/L)",9,"Week 6","NORMAL","NORMAL","Y",1,
"PILOT-1016",32,"NEUT","Neutrophils (10^6/L)",10,"Week 8","NORMAL","NORMAL","Y",1,
"PILOT-1016",32,"NEUT","Neutrophils (10^6/L)",11,"Week 12","NORMAL","NORMAL","Y",1,
"PILOT-1017",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","NORMAL","NORMAL","Y",1,
"PILOT-1018",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","NORMAL","NORMAL","Y",1,
"PILOT-1019",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","NORMAL","NORMAL","Y",1,
"PILOT-1020",32,"NEUT","Neutrophils (10^6/L)",7,"Week 2","NORMAL","NORMAL","Y",2,
)

adsl<-tribble(
~usubjid,~trt01an,~saffl,
"PILOT-1001",1,"Y",
"PILOT-1002",2,"Y",
"PILOT-1003",1,"Y",
"PILOT-1004",2,"Y",
"PILOT-1005",2,"Y",
"PILOT-1006",1,"Y",
"PILOT-1007",2,"Y",
"PILOT-1008",2,"Y",
"PILOT-1009",1,"Y",
"PILOT-1010",2,"Y",
"PILOT-1011",1,"Y",
"PILOT-1012",2,"Y",
"PILOT-1013",2,"Y",
"PILOT-1014",1,"Y",
"PILOT-1015",2,"Y",
"PILOT-1016",1,"Y",
"PILOT-1017",1,"Y",
"PILOT-1018",1,"Y",
"PILOT-1019",1,"Y",
"PILOT-1020",2,"Y",
)


# Simulate twice the original subject count while preserving parameter, visit,
# treatment and baseline/post-baseline normal-range relationships.
maps <- make_id_maps(list(ADSL=adsl, ADLB=adlb))
adsl <- double_subjects(adsl, maps)
adlb <- double_subjects(adlb, maps)
eligible <- which(adlb$anrind %in% c("LOW", "NORMAL", "HIGH"))
change <- eligible[runif(length(eligible)) < 0.05]
if(length(change)) adlb$anrind[change] <- sample(c("LOW", "NORMAL", "HIGH"), length(change), replace=TRUE)
stopifnot(n_distinct(adsl$usubjid) == 2L * maps$original_n)
write_xlsx_list(list(ADSL=adsl, ADLB=adlb), file.path(OUT_DIR, paste0(STEM,"_simulated_data.xlsx")))

# Original analysis structure retained from the source document.
# Read datasets and filter rows
adlb01 <- adlb %>%
 filter(saffl == "Y")
adsl01 <- adsl %>%
 filter(saffl == "Y")
#==============================================================================
# Create a variable named 'treatment' to hold report level column groupings
# Also create duplicate rows for presenting total column
#==============================================================================
adlb02 <- bind_rows(
 mutate(adlb01,treatment = trt01an),
 mutate(adlb01,treatment = 3)
 )
adsl02 <- bind_rows(
 mutate(adsl01,treatment = trt01an),
 mutate(adsl01,treatment = 3)
)
#==============================================================================
# Create numeric variables corresponding to anrind and bnrind
#==============================================================================
adlb03 <- adlb02 %>%
 mutate(anrindn = case_when(
 anrind == "LOW" ~ 1,
 anrind == "NORMAL" ~ 2,
 anrind == "HIGH" ~ 3,
 anrind == "" ~ 99,
 TRUE ~ NA
 ),
 bnrindn = case_when(
 bnrind == "LOW" ~ 1,
 bnrind == "NORMAL" ~ 2,
 bnrind == "HIGH" ~ 3,
 bnrind == "" ~ 99,
 TRUE ~ NA
 ))
#==============================================================================
# Get treatment totals into a dataset and into macro variables (for column headers)
#==============================================================================
trttotal_pre <- adsl02 %>%
 group_by(treatment) %>%
summarise(trttotal = n_distinct(usubjid), .groups="drop")
# Create a dummy dataset for treatment totals
dummy_pre <- tibble(treatment = 1:3)
# Merge actual counts with dummy counts
trttotals <- dummy_pre %>%
 left_join(trttotal_pre, by = "treatment") %>%
 mutate(trttotal = ifelse(is.na(trttotal), 0, trttotal))
#==============================================================================;
#Create a record for each safety for each analysis visit;
#==============================================================================;
dummy01 <- adlb02 %>%
 distinct(paramn, paramcd, param, avisitn, avisit) %>%
 filter(!is.na(paramn))
dummy02 <- adsl02 %>%
 cross_join(dummy01)
# Merge actual and dummy data
adlb04 <- adlb03 %>%
 full_join(dummy02, by = c("usubjid", "treatment", "paramn", "paramcd", "param", "avisitn", "avisit", "saffl", "trt01an")) %>%
 mutate(anrindn = ifelse(is.na(anrindn) & !is.na(param), 99, anrindn),
 bnrindn = ifelse(is.na(bnrindn) & !is.na(param), 99, bnrindn)) %>%
 select(-anrind, -bnrind)
# Duplicating data to create a total row and total column
adlb04_1 <- bind_rows(
 mutate(adlb04),
 adlb04 %>% filter(anrindn !=99) %>% mutate(anrindn=98)
)
adlb05 <- bind_rows(
 mutate(adlb04_1),
 adlb04_1 %>% filter(bnrindn !=99) %>% mutate(bnrindn=98)
)
#==============================================================================;
#Obtain counts;
#==============================================================================;
counts01 <- adlb05 %>%
 group_by(paramn, paramcd, param, avisitn, avisit, anrindn, bnrindn, treatment) %>%
summarise(count = n_distinct(usubjid), .groups="drop") %>%
 ungroup()
# Create dummy for all levels of bnrindn and anrindn
dummyx01 <- adlb05 %>%
 distinct(paramn, paramcd, param, avisitn, avisit)
dummyx02 <- expand_grid(anrindn = c(1, 2, 3, 98, 99), bnrindn = c(1, 2, 3, 98, 99), treatment = 1:3) %>%
 cross_join(dummyx01) %>%
 arrange(paramn, paramcd, param, avisitn, avisit, anrindn, bnrindn, treatment)
# Merge dummy and actual counts
counts02 <- counts01 %>%
 full_join(dummyx02, by = c("paramn", "paramcd", "param", "avisitn", "avisit", "anrindn", "bnrindn", "treatment")) %>%
 mutate(count = ifelse(is.na(count), 0, count))
# Get denominator to calculate percentages
denoms01 <- counts02 %>%
 filter(anrindn == 98) %>%
 rename(denom = count) %>%
 select(-anrindn)
counts03 <- counts02 %>%
 left_join(denoms01,by=c("paramn", "paramcd", "param", "avisitn", "avisit", "bnrindn", "treatment"))
# Bring treatment total into counts dataset to concatenate to the treatment label
counts03 <- counts03 %>%
 left_join(trttotals, by = "treatment")
# Calculate percentages and create a label column
counts04 <- counts03 %>%
 mutate(cp = ifelse(anrindn %in% c(1, 2, 3) & bnrindn %in% c(1,2,3),
 ifelse(count!=0 & anrindn %in% c(1, 2, 3) & bnrindn %in% c(1,2,3),
 paste0(count, " (", sprintf("%.1f", count / denom * 100), ")"),
 as.character(count)
 ),
 as.character(count)
 ),
 label = case_when(anrindn == 1 ~ "Low",
 anrindn == 2 ~ "Normal",
 anrindn == 3 ~ "High",
 anrindn == 98 ~ "Total",
 anrindn == 99 ~ "No Data"),
 treatmentc = case_when(treatment == 1 ~ paste0("Placebo (N=", trttotal, ")"),
 treatment == 2 ~ paste0("Active (N=", trttotal, ")"),
 treatment == 3 ~ paste0("Total (N=", trttotal, ")"))) %>%
 arrange(paramn,avisitn,treatment,anrindn,bnrindn)
# Transpose data such that the baseline indicator becomes columns
trans01 <-counts04 %>%
 pivot_wider(
 id_cols = c(paramn, paramcd, param, avisitn, avisit, anrindn, label, treatment, treatmentc),
 names_from = bnrindn,
 values_from = cp,
 names_prefix = "c") %>%
 arrange(paramn,avisitn,treatment,anrindn)
output<-trans01 %>%
 select(paramn,paramcd, param, avisitn,avisit,treatment,treatmentc,anrindn,label,c1,c2,c3,c98,c99)

display <- output %>%
  mutate(
    parameter_display=if_else(
      row_number()==1L | param!=lag(param),
      param,
      ""
    ),
    visit_display=if_else(
      row_number()==1L | param!=lag(param) | avisit!=lag(avisit),
      avisit,
      ""
    )
  ) %>%
  transmute(
    Parameter=parameter_display,
    Visit=visit_display,
    Treatment=treatmentc,
    `Post-baseline Category`=label,
    `Baseline Low`=coalesce(c1,"0"),
    `Baseline Normal`=coalesce(c2,"0"),
    `Baseline High`=coalesce(c3,"0"),
    Total=coalesce(c98,"0"),
    `No Data`=coalesce(c99,"0")
  )
write_clinical_rtf(display,
  "Table 1. Shift from Baseline in Laboratory Normal-Range Status",
  file.path(OUT_DIR,paste0(STEM,".rtf")),
  "Cells show n (percent) where applicable. Percentages use subjects with available baseline category within parameter, visit and treatment."
)
cat("Created ",STEM,": simulated N = ",2L*maps$original_n,"\n",sep="")
