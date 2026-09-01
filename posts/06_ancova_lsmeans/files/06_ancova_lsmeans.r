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
~fasfl,~usubjid,~trt01pn,~blhba1c,~trt01p,
"Y","1001",1,6.4,"High Dose",
"Y","1002",2,6.1,"Low Dose",
"Y","1003",2,6.6,"Low Dose",
"Y","1004",3,6.1,"Placebo",
"Y","1005",2,6.3,"Low Dose",
"Y","1006",1,7.6,"High Dose",
"Y","1007",2,7.4,"Low Dose",
"Y","1008",1,8.9,"High Dose",
"Y","1009",2,5.6,"Low Dose",
"Y","1010",1,8.7,"High Dose",
"Y","1011",3,5.8,"Placebo",
"Y","1012",3,5.9,"Placebo",
"Y","1013",2,5.6,"Low Dose",
"Y","1014",3,6.1,"Placebo",
"Y","1015",1,6.2,"High Dose",
"Y","1016",2,4.6,"Low Dose",
"Y","1017",3,4.7,"Placebo",
"Y","1018",1,6.9,"High Dose",
"Y","1019",2,6.1,"Low Dose",
"Y","1020",3,5,"Placebo",
"Y","1021",1,8.7,"High Dose",
"Y","1022",1,7.6,"High Dose",
"Y","1023",2,6.6,"Low Dose",
"Y","1024",2,5.6,"Low Dose",
"Y","1025",2,8.5,"Low Dose",
"Y","1026",1,5.4,"High Dose",
"Y","1027",1,5.7,"High Dose",
"Y","1028",3,5,"Placebo",
"Y","1029",3,7.7,"Placebo",
"Y","1030",1,5.1,"High Dose",
"Y","1031",2,5.5,"Low Dose",
"Y","1032",1,8.6,"High Dose",
"Y","1033",2,7,"Low Dose",
"Y","1034",2,7.5,"Low Dose",
"Y","1035",1,5.3,"High Dose",
"Y","1036",1,5.9,"High Dose",
"Y","1037",3,4.9,"Placebo",
"Y","1038",3,4.7,"Placebo",
"Y","1039",3,7.6,"Placebo",
"Y","1040",3,5.8,"Placebo",
"Y","1041",3,8.5,"Placebo",
"Y","1042",3,6.6,"Placebo",
"Y","1043",3,7.1,"Placebo",
"Y","1044",1,6,"High Dose",
"Y","1045",1,9,"High Dose",
"Y","1046",1,6.8,"High Dose",
"Y","1047",1,8,"High Dose",
"Y","1048",2,7.4,"Low Dose",
"Y","1049",3,7.2,"Placebo",
"Y","1050",3,5.3,"Placebo",
"Y","1051",1,7.5,"High Dose",
"Y","1052",1,4.8,"High Dose",
"Y","1053",3,5.8,"Placebo",
"Y","1054",3,7.5,"Placebo",
"Y","1055",1,8.2,"High Dose",
"Y","1056",1,7.8,"High Dose",
"Y","1057",3,5,"Placebo",
"Y","1058",2,8,"Low Dose",
"Y","1059",2,6.8,"Low Dose",
"Y","1060",2,8.2,"Low Dose",
"Y","1061",1,6.1,"High Dose",
"Y","1062",1,8.7,"High Dose",
"Y","1063",3,4.5,"Placebo",
"Y","1064",1,6.5,"High Dose",
"Y","1065",1,7.2,"High Dose",
"Y","1066",2,4.8,"Low Dose",
"Y","1067",1,6.8,"High Dose",
"Y","1068",3,8.6,"Placebo",
"Y","1069",1,8.2,"High Dose",
"Y","1070",2,5.3,"Low Dose",
"Y","1071",1,6.6,"High Dose",
"Y","1072",1,5.1,"High Dose",
"Y","1073",2,5,"Low Dose",
"Y","1074",3,8.8,"Placebo",
"Y","1075",1,7.8,"High Dose",
"Y","1076",2,8.9,"Low Dose",
"Y","1077",3,7.6,"Placebo",
"Y","1078",2,7.1,"Low Dose",
"Y","1079",1,8.9,"High Dose",
"Y","1080",2,4.9,"Low Dose",
"Y","1081",1,5.4,"High Dose",
"Y","1082",3,6.4,"Placebo",
"Y","1083",1,6.1,"High Dose",
"Y","1084",2,8.3,"Low Dose",
"Y","1085",1,7.9,"High Dose",
"Y","1086",1,5.7,"High Dose",
"Y","1087",3,7.2,"Placebo",
"Y","1088",1,6.1,"High Dose",
"Y","1089",2,6.4,"Low Dose",
"Y","1090",2,8.3,"Low Dose",
"Y","1091",2,4.5,"Low Dose",
"Y","1092",3,5.3,"Placebo",
"Y","1093",3,8.6,"Placebo",
"Y","1094",1,7.8,"High Dose",
"Y","1095",1,5.2,"High Dose",
"Y","1096",3,4.5,"Placebo",
"Y","1097",2,4.8,"Low Dose",
"Y","1098",1,5.3,"High Dose",
"Y","1099",2,8,"Low Dose",
"Y","1100",3,5.8,"Placebo",
"Y","1101",1,4.8,"High Dose",
"Y","1102",3,5.2,"Placebo",
"Y","1103",2,5.7,"Low Dose",
"Y","1104",1,4.7,"High Dose",
"Y","1105",1,5.5,"High Dose",
"Y","1106",3,5.9,"Placebo",
"Y","1107",1,7.5,"High Dose",
"Y","1108",3,4.6,"Placebo",
"Y","1109",3,7.9,"Placebo",
"Y","1110",1,4.8,"High Dose",
"Y","1111",2,4.8,"Low Dose",
"Y","1112",1,5.9,"High Dose",
"Y","1113",2,4.8,"Low Dose",
"Y","1114",3,5.6,"Placebo",
"Y","1115",2,5.1,"Low Dose",
"Y","1116",3,8.1,"Placebo",
"Y","1117",2,8.4,"Low Dose",
"Y","1118",1,7.3,"High Dose",
"Y","1119",2,6.9,"Low Dose",
"Y","1120",1,5,"High Dose",
"Y","1121",3,8.5,"Placebo",
"Y","1122",2,6.1,"Low Dose",
"Y","1123",1,7.4,"High Dose",
"Y","1124",1,8.6,"High Dose",
"Y","1125",1,6.4,"High Dose",
"Y","1126",1,7.4,"High Dose",
"Y","1127",3,8.3,"Placebo",
"Y","1128",2,8,"Low Dose",
"Y","1129",2,5.5,"Low Dose",
"Y","1130",2,6,"Low Dose",
"Y","1131",3,6.7,"Placebo",
"Y","1132",2,8.4,"Low Dose",
"Y","1133",1,6.6,"High Dose",
"Y","1134",1,5.3,"High Dose",
"Y","1135",1,7.1,"High Dose",
"Y","1136",1,4.7,"High Dose",
"Y","1137",2,8.7,"Low Dose",
"Y","1138",3,8.7,"Placebo",
"Y","1139",3,8.8,"Placebo",
"Y","1140",1,5.1,"High Dose",
"Y","1141",2,6.4,"Low Dose",
"Y","1142",1,8.2,"High Dose",
"Y","1143",1,5,"High Dose",
"Y","1144",3,8.3,"Placebo",
"Y","1145",1,4.8,"High Dose",
"Y","1146",2,4.5,"Low Dose",
"Y","1147",1,8.5,"High Dose",
"Y","1148",1,5.6,"High Dose",
"Y","1149",2,8,"Low Dose",
"Y","1150",2,5.8,"Low Dose",
"Y","1151",1,7.6,"High Dose",
"Y","1152",3,5.8,"Placebo",
"Y","1153",2,8.5,"Low Dose",
"Y","1154",3,8,"Placebo",
"Y","1155",2,8.8,"Low Dose",
"Y","1156",1,8.9,"High Dose",
"Y","1157",3,4.7,"Placebo",
"Y","1158",2,7.6,"Low Dose",
"Y","1159",2,4.5,"Low Dose",
"Y","1160",2,7.5,"Low Dose",
)

adlb<-tribble(
~fasfl,~usubjid,~trt01pn,~blhba1c,~trt01p,~paramcd,~avisitn,~chg,
"Y","1001",1,6.4,"High Dose","TRIG",12,-4.7,
"Y","1002",2,6.1,"Low Dose","TRIG",12,-1.8,
"Y","1005",2,6.3,"Low Dose","TRIG",12,-0.3,
"Y","1006",1,7.6,"High Dose","TRIG",12,-7.9,
"Y","1007",2,7.4,"Low Dose","TRIG",12,-2.6,
"Y","1009",2,5.6,"Low Dose","TRIG",12,-1.2,
"Y","1010",1,8.7,"High Dose","TRIG",12,-4.7,
"Y","1011",3,5.8,"Placebo","TRIG",12,3.1,
"Y","1012",3,5.9,"Placebo","TRIG",12,-3.1,
"Y","1013",2,5.6,"Low Dose","TRIG",12,2.2,
"Y","1014",3,6.1,"Placebo","TRIG",12,-3.5,
"Y","1015",1,6.2,"High Dose","TRIG",12,-4.2,
"Y","1016",2,4.6,"Low Dose","TRIG",12,-1.9,
"Y","1017",3,4.7,"Placebo","TRIG",12,-3.4,
"Y","1018",1,6.9,"High Dose","TRIG",12,0.3,
"Y","1019",2,6.1,"Low Dose","TRIG",12,0.8,
"Y","1020",3,5,"Placebo","TRIG",12,-1.1,
"Y","1021",1,8.7,"High Dose","TRIG",12,-1.7,
"Y","1022",1,7.6,"High Dose","TRIG",12,0.8,
"Y","1023",2,6.6,"Low Dose","TRIG",12,2.3,
"Y","1024",2,5.6,"Low Dose","TRIG",12,-3.7,
"Y","1025",2,8.5,"Low Dose","TRIG",12,-2,
"Y","1026",1,5.4,"High Dose","TRIG",12,-2.4,
"Y","1027",1,5.7,"High Dose","TRIG",12,0.3,
"Y","1028",3,5,"Placebo","TRIG",12,-0.4,
"Y","1029",3,7.7,"Placebo","TRIG",12,-1.8,
"Y","1030",1,5.1,"High Dose","TRIG",12,-3.5,
"Y","1031",2,5.5,"Low Dose","TRIG",12,6,
"Y","1032",1,8.6,"High Dose","TRIG",12,2.9,
"Y","1033",2,7,"Low Dose","TRIG",12,-1.6,
"Y","1035",1,5.3,"High Dose","TRIG",12,2.7,
"Y","1036",1,5.9,"High Dose","TRIG",12,1.6,
"Y","1037",3,4.9,"Placebo","TRIG",12,4.4,
"Y","1038",3,4.7,"Placebo","TRIG",12,-0.5,
"Y","1039",3,7.6,"Placebo","TRIG",12,-2.9,
"Y","1040",3,5.8,"Placebo","TRIG",12,-1.5,
"Y","1041",3,8.5,"Placebo","TRIG",12,-2.2,
"Y","1042",3,6.6,"Placebo","TRIG",12,0.1,
"Y","1043",3,7.1,"Placebo","TRIG",12,-4.1,
"Y","1044",1,6,"High Dose","TRIG",12,3.7,
"Y","1045",1,9,"High Dose","TRIG",12,1.6,
"Y","1046",1,6.8,"High Dose","TRIG",12,-2.5,
"Y","1047",1,8,"High Dose","TRIG",12,2.9,
"Y","1048",2,7.4,"Low Dose","TRIG",12,-3.6,
"Y","1049",3,7.2,"Placebo","TRIG",12,-1.4,
"Y","1050",3,5.3,"Placebo","TRIG",12,1.5,
"Y","1051",1,7.5,"High Dose","TRIG",12,0.3,
"Y","1052",1,4.8,"High Dose","TRIG",12,4.6,
"Y","1053",3,5.8,"Placebo","TRIG",12,-1.1,
"Y","1054",3,7.5,"Placebo","TRIG",12,3.6,
"Y","1055",1,8.2,"High Dose","TRIG",12,-5.2,
"Y","1056",1,7.8,"High Dose","TRIG",12,-1.1,
"Y","1057",3,5,"Placebo","TRIG",12,-0.4,
"Y","1058",2,8,"Low Dose","TRIG",12,-2.3,
"Y","1059",2,6.8,"Low Dose","TRIG",12,-1.6,
"Y","1060",2,8.2,"Low Dose","TRIG",12,2.3,
"Y","1061",1,6.1,"High Dose","TRIG",12,-4.2,
"Y","1062",1,8.7,"High Dose","TRIG",12,-0.3,
"Y","1065",1,7.2,"High Dose","TRIG",12,-2.1,
"Y","1066",2,4.8,"Low Dose","TRIG",12,1.4,
"Y","1067",1,6.8,"High Dose","TRIG",12,2.5,
"Y","1068",3,8.6,"Placebo","TRIG",12,1.3,
"Y","1069",1,8.2,"High Dose","TRIG",12,0.8,
"Y","1070",2,5.3,"Low Dose","TRIG",12,-2.2,
"Y","1071",1,6.6,"High Dose","TRIG",12,-2.7,
"Y","1072",1,5.1,"High Dose","TRIG",12,-1.1,
"Y","1073",2,5,"Low Dose","TRIG",12,-2.4,
"Y","1076",2,8.9,"Low Dose","TRIG",12,0.5,
"Y","1077",3,7.6,"Placebo","TRIG",12,-5.5,
"Y","1078",2,7.1,"Low Dose","TRIG",12,-2.9,
"Y","1080",2,4.9,"Low Dose","TRIG",12,-0.8,
"Y","1081",1,5.4,"High Dose","TRIG",12,-2.4,
"Y","1082",3,6.4,"Placebo","TRIG",12,1.9,
"Y","1083",1,6.1,"High Dose","TRIG",12,-0.4,
"Y","1084",2,8.3,"Low Dose","TRIG",12,-2.3,
"Y","1085",1,7.9,"High Dose","TRIG",12,0.2,
"Y","1086",1,5.7,"High Dose","TRIG",12,4.6,
"Y","1087",3,7.2,"Placebo","TRIG",12,1.4,
"Y","1088",1,6.1,"High Dose","TRIG",12,-1.8,
"Y","1089",2,6.4,"Low Dose","TRIG",12,-0.8,
"Y","1090",2,8.3,"Low Dose","TRIG",12,-0.5,
"Y","1091",2,4.5,"Low Dose","TRIG",12,-3.7,
"Y","1092",3,5.3,"Placebo","TRIG",12,2.7,
"Y","1093",3,8.6,"Placebo","TRIG",12,2.2,
"Y","1094",1,7.8,"High Dose","TRIG",12,2,
"Y","1096",3,4.5,"Placebo","TRIG",12,-2,
"Y","1097",2,4.8,"Low Dose","TRIG",12,0.5,
"Y","1098",1,5.3,"High Dose","TRIG",12,-1.7,
"Y","1099",2,8,"Low Dose","TRIG",12,-0.7,
"Y","1100",3,5.8,"Placebo","TRIG",12,-3.2,
"Y","1101",1,4.8,"High Dose","TRIG",12,1.1,
"Y","1102",3,5.2,"Placebo","TRIG",12,2.8,
"Y","1103",2,5.7,"Low Dose","TRIG",12,-5,
"Y","1104",1,4.7,"High Dose","TRIG",12,7.7,
"Y","1105",1,5.5,"High Dose","TRIG",12,1,
"Y","1106",3,5.9,"Placebo","TRIG",12,2.1,
"Y","1107",1,7.5,"High Dose","TRIG",12,-2.5,
"Y","1109",3,7.9,"Placebo","TRIG",12,-4.9,
"Y","1110",1,4.8,"High Dose","TRIG",12,-1.8,
"Y","1112",1,5.9,"High Dose","TRIG",12,4.5,
"Y","1113",2,4.8,"Low Dose","TRIG",12,-1.7,
"Y","1114",3,5.6,"Placebo","TRIG",12,-5.8,
"Y","1115",2,5.1,"Low Dose","TRIG",12,1,
"Y","1116",3,8.1,"Placebo","TRIG",12,2.6,
"Y","1117",2,8.4,"Low Dose","TRIG",12,-1.7,
"Y","1118",1,7.3,"High Dose","TRIG",12,-2.1,
"Y","1119",2,6.9,"Low Dose","TRIG",12,0.4,
"Y","1120",1,5,"High Dose","TRIG",12,2.2,
"Y","1121",3,8.5,"Placebo","TRIG",12,0.3,
"Y","1122",2,6.1,"Low Dose","TRIG",12,0.7,
"Y","1123",1,7.4,"High Dose","TRIG",12,-5.6,
"Y","1124",1,8.6,"High Dose","TRIG",12,-2.6,
"Y","1125",1,6.4,"High Dose","TRIG",12,0.8,
"Y","1126",1,7.4,"High Dose","TRIG",12,0.8,
"Y","1127",3,8.3,"Placebo","TRIG",12,0,
"Y","1128",2,8,"Low Dose","TRIG",12,-0.4,
"Y","1129",2,5.5,"Low Dose","TRIG",12,-2.4,
"Y","1130",2,6,"Low Dose","TRIG",12,-3,
"Y","1131",3,6.7,"Placebo","TRIG",12,2.7,
"Y","1132",2,8.4,"Low Dose","TRIG",12,-3.8,
"Y","1133",1,6.6,"High Dose","TRIG",12,3.1,
"Y","1134",1,5.3,"High Dose","TRIG",12,-3.2,
"Y","1135",1,7.1,"High Dose","TRIG",12,2.1,
"Y","1136",1,4.7,"High Dose","TRIG",12,2.6,
"Y","1137",2,8.7,"Low Dose","TRIG",12,1.3,
"Y","1138",3,8.7,"Placebo","TRIG",12,5.6,
"Y","1139",3,8.8,"Placebo","TRIG",12,1.3,
"Y","1140",1,5.1,"High Dose","TRIG",12,6,
"Y","1141",2,6.4,"Low Dose","TRIG",12,3.3,
"Y","1142",1,8.2,"High Dose","TRIG",12,0.8,
"Y","1143",1,5,"High Dose","TRIG",12,0.8,
"Y","1144",3,8.3,"Placebo","TRIG",12,-5.3,
"Y","1145",1,4.8,"High Dose","TRIG",12,2.8,
"Y","1146",2,4.5,"Low Dose","TRIG",12,-2.6,
"Y","1148",1,5.6,"High Dose","TRIG",12,1.2,
"Y","1149",2,8,"Low Dose","TRIG",12,2.1,
"Y","1150",2,5.8,"Low Dose","TRIG",12,-1.2,
"Y","1151",1,7.6,"High Dose","TRIG",12,-5.1,
"Y","1152",3,5.8,"Placebo","TRIG",12,3.8,
"Y","1153",2,8.5,"Low Dose","TRIG",12,-5.8,
"Y","1154",3,8,"Placebo","TRIG",12,-3.7,
"Y","1155",2,8.8,"Low Dose","TRIG",12,1.7,
"Y","1156",1,8.9,"High Dose","TRIG",12,-2.2,
"Y","1157",3,4.7,"Placebo","TRIG",12,-0.4,
"Y","1158",2,7.6,"Low Dose","TRIG",12,3.8,
"Y","1159",2,4.5,"Low Dose","TRIG",12,-2.7,
"Y","1160",2,7.5,"Low Dose","TRIG",12,-0.3,
)

# Analysis and outputs.

maps<-make_id_maps(list(ADSL=adsl,ADLB=adlb));adsl<-double_subjects(adsl,maps);adlb<-double_subjects(adlb,maps);adsl$blhba1c<-round(pmax(4,adsl$blhba1c+rnorm(nrow(adsl),0,.15)),1);adlb$chg<-round(adlb$chg+rnorm(nrow(adlb),0,.2),2);stopifnot(n_distinct(adsl$usubjid)==2L*maps$original_n);write_xlsx_list(list(ADSL=adsl,ADLB=adlb),file.path(OUT_DIR,paste0(STEM,"_simulated_data.xlsx")))
z<-adlb%>%filter(fasfl=="Y",avisitn==12)%>%left_join(select(adsl,usubjid,blhba1c),by=c("usubjid","blhba1c"))%>%mutate(treatment=factor(trt01pn,levels=1:3));if(!"blhba1c"%in%names(z))z<-adlb%>%filter(fasfl=="Y",avisitn==12)%>%mutate(treatment=factor(trt01pn,levels=1:3));model<-lm(chg~treatment+blhba1c,data=z);ref<-mean(z$blhba1c,na.rm=TRUE);nd<-data.frame(treatment=factor(1:3,levels=1:3),blhba1c=ref);X<-model.matrix(delete.response(terms(model)),nd);est<-drop(X%*%coef(model));se<-sqrt(diag(X%*%vcov(model)%*%t(X)));crit<-qt(.975,df.residual(model));trt_labels<-c("High Dose","Low Dose","Placebo");rows<-lapply(1:3,function(g)data.frame(Treatment=trt_labels[g],N=as.character(sum(z$trt01pn==g)),`LS Mean (SE)`=sprintf("%.2f (%.3f)",est[g],se[g]),`95% CI`=sprintf("(%.2f, %.2f)",est[g]-crit*se[g],est[g]+crit*se[g]),Comparison="",`Difference (SE)`="",`Comparison p-value`="",check.names=FALSE));for(pair in list(c(2,3),c(1,3),c(1,2))){cvec<-X[pair[1],]-X[pair[2],];e<-sum(cvec*coef(model));s<-sqrt(drop(t(cvec)%*%vcov(model)%*%cvec));pv<-2*pt(-abs(e/s),df.residual(model));rows[[length(rows)+1]]<-data.frame(Treatment="",N="",`LS Mean (SE)`="",`95% CI`="",Comparison=paste(trt_labels[pair],collapse=" vs "),`Difference (SE)`=sprintf("%.2f (%.3f)",e,s),`Comparison p-value`=fmt_p(pv),check.names=FALSE)};output<-bind_rows(rows)
write_clinical_rtf(output,"Table 6. ANCOVA Least-Squares Means and Treatment Comparisons",file.path(OUT_DIR,paste0(STEM,".rtf")),"ANCOVA model: change = treatment + baseline HbA1c; LS means are evaluated at the overall mean baseline.")

cat("Created 06_ancova_lsmeans: simulated N =", 2L*maps$original_n, "\n")
