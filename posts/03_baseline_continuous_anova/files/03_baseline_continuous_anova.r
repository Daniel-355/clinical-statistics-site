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
~USUBJID,~TRT01PN,~AGE,~ITTFL,~WEIGHTBL,
"PILOT-1001",2,80,"Y",55.8,
"PILOT-1002",3,77,"Y",72.6,
"PILOT-1003",2,54,"Y",59.9,
"PILOT-1004",3,76,"Y",46.7,
"PILOT-1005",3,67,"Y",81,
"PILOT-1006",3,77,"Y",63.1,
"PILOT-1007",1,69,"Y",49,
"PILOT-1008",2,68,"Y",74.4,
"PILOT-1009",1,83,"Y",70.8,
"PILOT-1010",3,82,"Y",69,
"PILOT-1011",2,56,"Y",71.7,
"PILOT-1012",2,62,"Y",82.6,
"PILOT-1013",1,88,"Y",54.4,
"PILOT-1014",2,72,"Y",55.3,
"PILOT-1015",2,78,"Y",51.3,
"PILOT-1016",2,83,"Y",58.1,
"PILOT-1017",3,60,"Y",57.8,
"PILOT-1018",1,82,"Y",70.3,
"PILOT-1019",3,78,"Y",76.7,
"PILOT-1020",1,76,"Y",49,
"PILOT-1021",1,81,"Y",41.7,
"PILOT-1022",2,67,"Y",50.4,
"PILOT-1023",2,86,"Y",56.5,
"PILOT-1024",1,89,"Y",47.2,
"PILOT-1025",3,79,"Y",73.5,
"PILOT-1026",3,79,"Y",53.5,
"PILOT-1027",1,84,"Y",61.2,
"PILOT-1028",3,82,"Y",80.3,
"PILOT-1029",2,71,"Y",88.9,
"PILOT-1030",2,79,"Y",69.9,
"PILOT-1031",3,74,"Y",53.1,
"PILOT-1032",3,84,"Y",44.5,
"PILOT-1033",2,81,"Y",59.9,
"PILOT-1034",1,84,"Y",41.1,
"PILOT-1035",1,84,"Y",79.4,
"PILOT-1036",3,81,"Y",74.4,
"PILOT-1037",2,74,"Y",64.4,
"PILOT-1038",1,79,"Y",58.1,
"PILOT-1039",2,84,"Y",82.6,
"PILOT-1040",1,59,"Y",63.5,
"PILOT-1041",2,85,"Y",59.4,
"PILOT-1042",2,81,"Y",54.9,
"PILOT-1043",2,68,"Y",62.6,
"PILOT-1044",1,74,"Y",66.2,
"PILOT-1045",3,77,"Y",62.8,
"PILOT-1046",3,56,"Y",89.8,
"PILOT-1047",2,84,"Y",68,
"PILOT-1048",2,80,"Y",63.5,
"PILOT-1049",2,75,"Y",85.7,
"PILOT-1050",1,65,"Y",74.4,
"PILOT-1051",1,78,"Y",49.4,
"PILOT-1052",2,76,"Y",45.4,
"PILOT-1053",3,76,"Y",41.7,
"PILOT-1054",2,77,"Y",68,
"PILOT-1055",1,87,"Y",46.3,
"PILOT-1056",3,56,"Y",95.9,
"PILOT-1057",3,86,"Y",73,
"PILOT-1058",2,78,"Y",82.1,
"PILOT-1059",1,71,"Y",77.1,
"PILOT-1060",1,82,"Y",65.3,
"PILOT-1061",1,78,"Y",81.7,
"PILOT-1062",1,86,"Y",47.6,
"PILOT-1063",3,61,"Y",69,
"PILOT-1064",3,56,"Y",69.4,
"PILOT-1065",3,56,"Y",56.3,
"PILOT-1066",3,88,"Y",56.3,
"PILOT-1067",2,76,"Y",71.7,
"PILOT-1068",1,81,"Y",45.4,
"PILOT-1069",3,65,"Y",78.5,
"PILOT-1070",2,78,"Y",80.7,
"PILOT-1071",1,83,"Y",56.3,
"PILOT-1072",2,68,"Y",66.7,
"PILOT-1073",3,65,"Y",75.8,
"PILOT-1074",3,75,"Y",59.9,
"PILOT-1075",3,67,"Y",55.3,
"PILOT-1076",1,63,"Y",56.3,
"PILOT-1077",2,79,"Y",60.3,
"PILOT-1078",1,86,"Y",44.5,
"PILOT-1079",2,87,"Y",50.4,
"PILOT-1080",3,73,"Y",89.4,
"PILOT-1081",1,79,"Y",58.5,
"PILOT-1082",3,77,"Y",81.3,
"PILOT-1083",3,67,"Y",80.3,
"PILOT-1084",1,81,"Y",64.9,
"PILOT-1085",3,84,"Y",88.9,
"PILOT-1086",1,78,"Y",57.6,
"PILOT-1087",1,67,"Y",70.3,
"PILOT-1088",3,82,"Y",96.2,
"PILOT-1089",1,87,"Y",59.4,
"PILOT-1090",1,81,"Y",58.5,
"PILOT-1091",2,84,"Y",56.3,
"PILOT-1092",1,76,"Y",67.1,
"PILOT-1093",3,77,"Y",68.5,
"PILOT-1094",3,81,"Y",91.6,
"PILOT-1095",2,84,"Y",59,
"PILOT-1096",2,84,"Y",56.7,
"PILOT-1097",1,74,"Y",76.2,
"PILOT-1098",1,85,"Y",57.6,
"PILOT-1099",2,57,"Y",106.1,
"PILOT-1100",1,75,"Y",72.1,
"PILOT-1101",2,75,"Y",47.2,
"PILOT-1102",3,72,"Y",54.4,
)

# Analysis and outputs.

adsl<-rename_with(adsl,tolower);maps<-make_id_maps(list(ADSL=adsl));adsl<-double_subjects(adsl,maps);adsl$age<-pmax(18,round(adsl$age+rnorm(nrow(adsl),0,2)));adsl$weightbl<-round(pmax(35,adsl$weightbl+rnorm(nrow(adsl),0,1.8)),1);stopifnot(n_distinct(adsl$usubjid)==2L*maps$original_n);write_xlsx_list(list(ADSL=adsl),file.path(OUT_DIR,paste0(STEM,"_simulated_data.xlsx")))
z<-adsl%>%filter(ittfl=="Y")%>%mutate(treatment=factor(trt01pn));make_cont<-function(v,label,d=1){pv<-anova(lm(reformulate("treatment",v),z))[1,"Pr(>F)"];s<-z%>%group_by(treatment)%>%summarise(N=sum(!is.na(.data[[v]])),Mean=mean(.data[[v]],na.rm=TRUE),SD=sd(.data[[v]],na.rm=TRUE),Median=median(.data[[v]],na.rm=TRUE),Min=min(.data[[v]],na.rm=TRUE),Max=max(.data[[v]],na.rm=TRUE),.groups="drop");out<-data.frame(Parameter=label,Statistic=c("n","Mean (SD)","Median","Min, Max"),check.names=FALSE);for(g in levels(z$treatment)){q<-filter(s,treatment==g);out[[paste0("Treatment ",g)]]<-c(q$N,sprintf(paste0("%.",d,"f (%.",d+1,"f)"),q$Mean,q$SD),sprintf(paste0("%.",d,"f"),q$Median),sprintf(paste0("%.",d,"f, %.",d,"f"),q$Min,q$Max))};out$`ANOVA p-value`<-c(fmt_p(pv),rep("",3));out};output<-bind_rows(make_cont("age","Age (years)",0),make_cont("weightbl","Baseline Weight (kg)",1))
write_clinical_rtf(output,"Table 3. Baseline Continuous Characteristics with ANOVA",file.path(OUT_DIR,paste0(STEM,".rtf")),"P-values are from one-way analysis of variance across treatment groups.")

cat("Created 03_baseline_continuous_anova: simulated N =", 2L*maps$original_n, "\n")
