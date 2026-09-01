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
~USUBJID,~TRT01AN,~RACE,~RACEN,~SEX,~SAFFL,~SEXN,
"PILOT-1001",1,"WHITE",1,"F","Y",2,
"PILOT-1002",2,"WHITE",1,"M","Y",1,
"PILOT-1003",2,"WHITE",1,"M","Y",1,
"PILOT-1004",2,"WHITE",1,"M","Y",1,
"PILOT-1005",2,"WHITE",1,"F","Y",2,
"PILOT-1006",1,"WHITE",1,"M","Y",1,
"PILOT-1007",2,"WHITE",1,"F","Y",2,
"PILOT-1008",1,"WHITE",1,"F","Y",2,
"PILOT-1009",2,"WHITE",1,"F","Y",2,
"PILOT-1010",1,"BLACK OR AFRICAN AMERICAN",2,"F","Y",2,
"PILOT-1011",2,"WHITE",1,"F","Y",2,
"PILOT-1012",2,"WHITE",1,"F","Y",2,
"PILOT-1013",2,"WHITE",1,"F","Y",2,
"PILOT-1014",1,"WHITE",1,"M","Y",1,
"PILOT-1015",2,"WHITE",1,"M","Y",1,
"PILOT-1016",2,"WHITE",1,"F","Y",2,
"PILOT-1017",2,"WHITE",1,"F","Y",2,
"PILOT-1018",2,"BLACK OR AFRICAN AMERICAN",2,"F","Y",2,
"PILOT-1019",1,"WHITE",1,"F","Y",2,
"PILOT-1020",2,"WHITE",1,"F","Y",2,
"PILOT-1021",2,"WHITE",1,"F","Y",2,
"PILOT-1022",1,"WHITE",1,"F","Y",2,
"PILOT-1023",2,"WHITE",1,"F","Y",2,
"PILOT-1024",2,"WHITE",1,"M","Y",1,
"PILOT-1025",1,"WHITE",1,"F","Y",2,
"PILOT-1026",1,"WHITE",1,"M","Y",1,
"PILOT-1027",1,"WHITE",1,"F","Y",2,
"PILOT-1028",1,"WHITE",1,"F","Y",2,
"PILOT-1029",2,"WHITE",1,"M","Y",1,
"PILOT-1030",1,"WHITE",1,"F","Y",2,
"PILOT-1031",1,"WHITE",1,"M","Y",1,
"PILOT-1032",2,"WHITE",1,"F","Y",2,
"PILOT-1033",1,"WHITE",1,"M","Y",1,
"PILOT-1034",2,"WHITE",1,"F","Y",2,
"PILOT-1035",1,"WHITE",1,"F","Y",2,
"PILOT-1036",1,"WHITE",1,"F","Y",2,
"PILOT-1037",1,"WHITE",1,"F","Y",2,
"PILOT-1038",2,"WHITE",1,"M","Y",1,
"PILOT-1039",1,"WHITE",1,"F","Y",2,
"PILOT-1040",2,"WHITE",1,"F","Y",2,
"PILOT-1041",2,"WHITE",1,"M","Y",1,
"PILOT-1042",2,"WHITE",1,"F","Y",2,
"PILOT-1043",1,"WHITE",1,"F","Y",2,
"PILOT-1044",1,"WHITE",1,"M","Y",1,
"PILOT-1045",1,"WHITE",1,"M","Y",1,
"PILOT-1046",1,"WHITE",1,"F","Y",2,
"PILOT-1047",2,"WHITE",1,"M","Y",1,
"PILOT-1048",1,"WHITE",1,"F","Y",2,
"PILOT-1049",1,"WHITE",1,"F","Y",2,
"PILOT-1050",2,"WHITE",1,"M","Y",1,
"PILOT-1051",2,"WHITE",1,"F","Y",2,
"PILOT-1052",2,"BLACK OR AFRICAN AMERICAN",2,"F","Y",2,
"PILOT-1053",2,"WHITE",1,"M","Y",1,
"PILOT-1054",1,"WHITE",1,"F","Y",2,
"PILOT-1055",1,"WHITE",1,"M","Y",1,
"PILOT-1056",1,"WHITE",1,"F","Y",2,
"PILOT-1057",1,"WHITE",1,"F","Y",2,
"PILOT-1058",1,"WHITE",1,"M","Y",1,
"PILOT-1059",2,"WHITE",1,"F","Y",2,
"PILOT-1060",2,"WHITE",1,"M","Y",1,
"PILOT-1061",1,"WHITE",1,"M","Y",1,
"PILOT-1062",1,"WHITE",1,"M","Y",1,
"PILOT-1063",1,"WHITE",1,"M","Y",1,
"PILOT-1064",2,"WHITE",1,"F","Y",2,
"PILOT-1065",2,"WHITE",1,"F","Y",2,
"PILOT-1066",2,"WHITE",1,"M","Y",1,
"PILOT-1067",2,"BLACK OR AFRICAN AMERICAN",2,"F","Y",2,
"PILOT-1068",2,"WHITE",1,"M","Y",1,
"PILOT-1069",2,"WHITE",1,"F","Y",2,
"PILOT-1070",1,"WHITE",1,"F","Y",2,
"PILOT-1071",1,"WHITE",1,"F","Y",2,
"PILOT-1072",2,"WHITE",1,"M","Y",1,
"PILOT-1073",1,"WHITE",1,"F","Y",2,
"PILOT-1074",1,"WHITE",1,"F","Y",2,
"PILOT-1075",1,"WHITE",1,"F","Y",2,
"PILOT-1076",2,"WHITE",1,"F","Y",2,
"PILOT-1077",2,"WHITE",1,"F","Y",2,
"PILOT-1078",2,"WHITE",1,"M","Y",1,
"PILOT-1079",1,"WHITE",1,"F","Y",2,
"PILOT-1080",1,"WHITE",1,"F","Y",2,
"PILOT-1081",2,"WHITE",1,"F","Y",2,
"PILOT-1082",1,"WHITE",1,"M","Y",1,
"PILOT-1083",2,"BLACK OR AFRICAN AMERICAN",2,"F","Y",2,
"PILOT-1084",1,"BLACK OR AFRICAN AMERICAN",2,"F","Y",2,
"PILOT-1085",1,"WHITE",1,"F","Y",2,
"PILOT-1086",1,"BLACK OR AFRICAN AMERICAN",2,"M","Y",1,
"PILOT-1087",2,"BLACK OR AFRICAN AMERICAN",2,"F","Y",2,
"PILOT-1088",2,"WHITE",1,"M","Y",1,
"PILOT-1089",1,"WHITE",1,"M","Y",1,
"PILOT-1090",1,"WHITE",1,"M","Y",1,
"PILOT-1091",1,"WHITE",1,"F","Y",2,
"PILOT-1092",2,"WHITE",1,"M","Y",1,
"PILOT-1093",1,"WHITE",1,"M","Y",1,
"PILOT-1094",2,"WHITE",1,"F","Y",2,
"PILOT-1095",2,"WHITE",1,"F","Y",2,
"PILOT-1096",1,"WHITE",1,"M","Y",1,
"PILOT-1097",1,"WHITE",1,"M","Y",1,
"PILOT-1098",1,"WHITE",1,"M","Y",1,
"PILOT-1099",2,"WHITE",1,"F","Y",2,
"PILOT-1100",2,"WHITE",1,"F","Y",2,
)

# Analysis and outputs.

adsl<-rename_with(adsl,tolower);maps<-make_id_maps(list(ADSL=adsl));adsl<-double_subjects(adsl,maps);stopifnot(n_distinct(adsl$usubjid)==2L*maps$original_n);write_xlsx_list(list(ADSL=adsl),file.path(OUT_DIR,paste0(STEM,"_simulated_data.xlsx")))
z<-adsl%>%filter(saffl=="Y")%>%mutate(treatment=trt01an);den<-z%>%count(treatment,name="N")
make_cat<-function(var,label){x<-z%>%mutate(category=ifelse(is.na(.data[[var]])|.data[[var]]=="","Missing",.data[[var]]));tab<-table(x$category,x$treatment);pv<-tryCatch(chisq.test(tab,correct=FALSE)$p.value,error=function(e)NA_real_);cats<-rownames(tab);bind_rows(lapply(seq_along(cats),function(i){n1<-tab[i,"1"];n2<-tab[i,"2"];data.frame(Parameter=ifelse(i==1,label,""),Category=cats[i],`Treatment 1 n (%)`=fmt_np(n1,den$N[den$treatment==1]),`Treatment 2 n (%)`=fmt_np(n2,den$N[den$treatment==2]),`Chi-square p-value`=ifelse(i==1,fmt_p(pv),""),check.names=FALSE)}))};output<-bind_rows(make_cat("sex","Sex"),make_cat("race","Race"))
write_clinical_rtf(output,"Table 2. Demographic Characteristics with Chi-Square Tests",file.path(OUT_DIR,paste0(STEM,".rtf")),"Percentages use treatment-group Safety Analysis Set denominators; p-values are Pearson chi-square tests.")

cat("Created 02_demographics_chi_square: simulated N =", 2L*maps$original_n, "\n")
