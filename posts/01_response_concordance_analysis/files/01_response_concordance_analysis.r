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
~usubjid,~trt01pn,
1001,2,
1002,1,
1003,2,
1004,2,
1005,2,
1006,1,
1007,2,
1008,1,
1009,2,
1010,2,
1011,1,
1012,1,
1013,2,
1014,2,
1015,2,
1016,2,
1017,1,
1018,2,
1019,1,
1020,2,
1021,2,
1022,1,
1023,2,
1024,2,
1025,1,
1026,2,
1027,2,
1028,2,
1029,1,
1030,2,
1031,1,
1032,1,
1033,2,
1034,1,
1035,2,
1036,2,
1037,2,
1038,1,
1039,1,
1040,1,
1041,2,
1042,2,
1043,2,
1044,2,
1045,2,
1046,1,
1047,2,
1048,1,
1049,1,
1050,1,
1051,2,
1052,2,
1053,1,
1054,2,
1055,2,
1056,1,
1057,2,
1058,2,
1059,2,
1060,1,
1061,1,
1062,1,
1063,2,
1064,1,
1065,1,
1066,2,
1067,2,
1068,2,
1069,1,
1070,2,
1071,1,
1072,1,
1073,1,
1074,1,
1075,2,
1076,1,
1077,1,
1078,2,
1079,1,
1080,2,
1081,2,
1082,1,
1083,1,
1084,2,
1085,1,
1086,2,
1087,1,
1088,1,
1089,1,
1090,1,
1091,1,
1092,1,
1093,1,
1094,1,
1095,1,
1096,1,
1097,1,
1098,2,
1099,1,
1100,2,
1101,1,
1102,2,
1103,1,
1104,1,
1105,2,
1106,1,
1107,2,
1108,1,
1109,1,
1110,1,
1111,2,
1112,2,
1113,2,
1114,1,
1115,2,
1116,1,
1117,1,
1118,2,
1119,2,
1120,1,
1121,2,
1122,1,
1123,2,
1124,2,
1125,1,
1126,1,
1127,2,
1128,2,
1129,2,
1130,1,
1131,2,
1132,2,
1133,1,
1134,1,
1135,2,
1136,1,
1137,1,
1138,2,
1139,2,
1140,1,
1141,2,
1142,2,
1143,2,
1144,1,
1145,1,
1146,2,
1147,2,
1148,2,
1149,1,
1150,2,
1151,1,
1152,2,
1153,2,
1154,2,
1155,1,
1156,1,
1157,1,
1158,2,
1159,1,
1160,2,
1161,1,
1162,1,
1163,1,
1164,2,
1165,2,
1166,1,
1167,2,
1168,1,
1169,2,
)

adrs<-tribble(
~usubjid,~trt01pn,~paramcd,~aval,~avalc,
1001,2,"INV",2,"PR",
1001,2,"IRC",3,"SD",
1002,2,"INV",5,"NE",
1002,2,"IRC",2,"PR",
1003,2,"INV",1,"CR",
1003,2,"IRC",5,"NE",
1004,2,"INV",2,"PR",
1004,2,"IRC",3,"SD",
1005,2,"INV",4,"PD",
1005,2,"IRC",4,"PD",
1006,2,"INV",3,"SD",
1006,2,"IRC",4,"PD",
1007,1,"INV",5,"NE",
1007,1,"IRC",3,"SD",
1008,1,"INV",4,"PD",
1008,1,"IRC",4,"PD",
1009,1,"INV",5,"NE",
1009,1,"IRC",4,"PD",
1010,2,"INV",1,"CR",
1010,2,"IRC",3,"SD",
1011,1,"INV",2,"PR",
1011,1,"IRC",4,"PD",
1012,1,"INV",4,"PD",
1012,1,"IRC",4,"PD",
1013,2,"INV",1,"CR",
1013,2,"IRC",1,"CR",
1014,1,"INV",5,"NE",
1014,1,"IRC",3,"SD",
1015,2,"INV",4,"PD",
1015,2,"IRC",4,"PD",
1016,1,"INV",3,"SD",
1016,1,"IRC",2,"PR",
1017,1,"INV",3,"SD",
1017,1,"IRC",5,"NE",
1018,2,"INV",1,"CR",
1018,2,"IRC",3,"SD",
1019,2,"INV",1,"CR",
1019,2,"IRC",4,"PD",
1020,2,"INV",5,"NE",
1020,2,"IRC",2,"PR",
1021,1,"INV",1,"CR",
1021,1,"IRC",5,"NE",
1022,1,"INV",3,"SD",
1022,1,"IRC",5,"NE",
1023,2,"INV",5,"NE",
1023,2,"IRC",2,"PR",
1024,2,"INV",3,"SD",
1024,2,"IRC",2,"PR",
1025,1,"INV",3,"SD",
1025,1,"IRC",4,"PD",
1026,1,"INV",2,"PR",
1026,1,"IRC",4,"PD",
1027,1,"INV",5,"NE",
1027,1,"IRC",3,"SD",
1028,1,"INV",1,"CR",
1028,1,"IRC",5,"NE",
1029,1,"INV",3,"SD",
1029,1,"IRC",3,"SD",
1030,1,"INV",2,"PR",
1030,1,"IRC",2,"PR",
1031,1,"INV",1,"CR",
1031,1,"IRC",2,"PR",
1032,1,"INV",2,"PR",
1032,1,"IRC",3,"SD",
1033,1,"INV",5,"NE",
1033,1,"IRC",2,"PR",
1034,2,"INV",3,"SD",
1034,2,"IRC",5,"NE",
1035,1,"INV",2,"PR",
1035,1,"IRC",5,"NE",
1036,1,"INV",4,"PD",
1036,1,"IRC",2,"PR",
1037,1,"INV",1,"CR",
1037,1,"IRC",5,"NE",
1038,2,"INV",4,"PD",
1038,2,"IRC",1,"CR",
1039,2,"INV",2,"PR",
1039,2,"IRC",3,"SD",
1040,2,"INV",3,"SD",
1040,2,"IRC",2,"PR",
1041,2,"INV",2,"PR",
1041,2,"IRC",4,"PD",
1042,2,"INV",3,"SD",
1042,2,"IRC",2,"PR",
1043,2,"INV",5,"NE",
1043,2,"IRC",4,"PD",
1044,1,"INV",3,"SD",
1044,1,"IRC",3,"SD",
1045,1,"INV",1,"CR",
1045,1,"IRC",5,"NE",
1046,1,"INV",1,"CR",
1046,1,"IRC",3,"SD",
1047,2,"INV",2,"PR",
1047,2,"IRC",4,"PD",
1048,2,"INV",3,"SD",
1048,2,"IRC",2,"PR",
1049,1,"INV",5,"NE",
1049,1,"IRC",5,"NE",
1050,2,"INV",2,"PR",
1050,2,"IRC",4,"PD",
1051,1,"INV",5,"NE",
1051,1,"IRC",3,"SD",
1052,2,"INV",2,"PR",
1052,2,"IRC",2,"PR",
1053,1,"INV",5,"NE",
1053,1,"IRC",1,"CR",
1054,2,"INV",2,"PR",
1054,2,"IRC",2,"PR",
1055,1,"INV",4,"PD",
1055,1,"IRC",4,"PD",
1056,1,"INV",4,"PD",
1056,1,"IRC",3,"SD",
1057,2,"INV",3,"SD",
1057,2,"IRC",3,"SD",
1058,1,"INV",5,"NE",
1058,1,"IRC",1,"CR",
1059,1,"INV",2,"PR",
1059,1,"IRC",1,"CR",
1060,2,"INV",2,"PR",
1060,2,"IRC",2,"PR",
1061,1,"INV",5,"NE",
1061,1,"IRC",1,"CR",
1062,2,"INV",2,"PR",
1062,2,"IRC",3,"SD",
1063,1,"INV",5,"NE",
1063,1,"IRC",5,"NE",
1064,2,"INV",3,"SD",
1064,2,"IRC",5,"NE",
1065,1,"INV",2,"PR",
1065,1,"IRC",1,"CR",
1066,1,"INV",5,"NE",
1066,1,"IRC",3,"SD",
1067,2,"INV",3,"SD",
1067,2,"IRC",2,"PR",
1068,1,"INV",4,"PD",
1068,1,"IRC",1,"CR",
1069,2,"INV",3,"SD",
1069,2,"IRC",1,"CR",
1070,2,"INV",4,"PD",
1070,2,"IRC",4,"PD",
1071,2,"INV",3,"SD",
1071,2,"IRC",5,"NE",
1072,1,"INV",5,"NE",
1072,1,"IRC",5,"NE",
1073,2,"INV",3,"SD",
1073,2,"IRC",2,"PR",
1074,2,"INV",2,"PR",
1074,2,"IRC",2,"PR",
1075,2,"INV",5,"NE",
1075,2,"IRC",3,"SD",
1076,2,"INV",2,"PR",
1076,2,"IRC",5,"NE",
1077,2,"INV",3,"SD",
1077,2,"IRC",3,"SD",
1078,2,"INV",1,"CR",
1078,2,"IRC",3,"SD",
1079,1,"INV",2,"PR",
1079,1,"IRC",5,"NE",
1080,1,"INV",4,"PD",
1080,1,"IRC",5,"NE",
1081,1,"INV",1,"CR",
1081,1,"IRC",2,"PR",
1082,2,"INV",2,"PR",
1082,2,"IRC",2,"PR",
1083,1,"INV",3,"SD",
1083,1,"IRC",5,"NE",
1084,1,"INV",2,"PR",
1084,1,"IRC",1,"CR",
1085,2,"INV",5,"NE",
1085,2,"IRC",5,"NE",
1086,2,"INV",2,"PR",
1086,2,"IRC",1,"CR",
1087,1,"INV",5,"NE",
1087,1,"IRC",1,"CR",
1088,1,"INV",5,"NE",
1088,1,"IRC",3,"SD",
1089,1,"INV",5,"NE",
1089,1,"IRC",5,"NE",
1090,2,"INV",1,"CR",
1090,2,"IRC",5,"NE",
1091,1,"INV",4,"PD",
1091,1,"IRC",5,"NE",
1092,1,"INV",2,"PR",
1092,1,"IRC",2,"PR",
1093,2,"INV",3,"SD",
1093,2,"IRC",2,"PR",
1094,2,"INV",2,"PR",
1094,2,"IRC",5,"NE",
1095,2,"INV",4,"PD",
1095,2,"IRC",1,"CR",
1096,2,"INV",5,"NE",
1096,2,"IRC",3,"SD",
1097,1,"INV",2,"PR",
1097,1,"IRC",4,"PD",
1098,1,"INV",2,"PR",
1098,1,"IRC",5,"NE",
1099,1,"INV",2,"PR",
1099,1,"IRC",3,"SD",
1100,1,"INV",3,"SD",
1100,1,"IRC",5,"NE",
1101,1,"INV",4,"PD",
1101,1,"IRC",3,"SD",
1102,1,"INV",3,"SD",
1102,1,"IRC",4,"PD",
1103,1,"INV",1,"CR",
1103,1,"IRC",4,"PD",
1104,1,"INV",1,"CR",
1104,1,"IRC",3,"SD",
1105,2,"INV",3,"SD",
1105,2,"IRC",5,"NE",
1106,1,"INV",1,"CR",
1106,1,"IRC",4,"PD",
1107,1,"INV",4,"PD",
1107,1,"IRC",5,"NE",
1108,2,"INV",2,"PR",
1108,2,"IRC",3,"SD",
1109,2,"INV",3,"SD",
1109,2,"IRC",2,"PR",
1110,2,"INV",1,"CR",
1110,2,"IRC",3,"SD",
1111,2,"INV",2,"PR",
1111,2,"IRC",4,"PD",
1112,1,"INV",2,"PR",
1112,1,"IRC",4,"PD",
1113,1,"INV",1,"CR",
1113,1,"IRC",2,"PR",
1114,2,"INV",5,"NE",
1114,2,"IRC",2,"PR",
1115,1,"INV",4,"PD",
1115,1,"IRC",1,"CR",
1116,2,"INV",2,"PR",
1116,2,"IRC",2,"PR",
1117,1,"INV",3,"SD",
1117,1,"IRC",5,"NE",
1118,2,"INV",3,"SD",
1118,2,"IRC",1,"CR",
1119,1,"INV",3,"SD",
1119,1,"IRC",5,"NE",
1120,1,"INV",4,"PD",
1120,1,"IRC",2,"PR",
1121,2,"INV",5,"NE",
1121,2,"IRC",1,"CR",
1122,1,"INV",4,"PD",
1122,1,"IRC",2,"PR",
1123,2,"INV",1,"CR",
1123,2,"IRC",4,"PD",
1124,2,"INV",5,"NE",
1124,2,"IRC",5,"NE",
1125,2,"INV",2,"PR",
1125,2,"IRC",5,"NE",
1126,2,"INV",1,"CR",
1126,2,"IRC",2,"PR",
1127,2,"INV",2,"PR",
1127,2,"IRC",4,"PD",
1128,2,"INV",1,"CR",
1128,2,"IRC",2,"PR",
1129,2,"INV",1,"CR",
1129,2,"IRC",2,"PR",
1130,1,"INV",1,"CR",
1130,1,"IRC",3,"SD",
1131,1,"INV",3,"SD",
1131,1,"IRC",3,"SD",
1132,1,"INV",3,"SD",
1132,1,"IRC",2,"PR",
1133,1,"INV",5,"NE",
1133,1,"IRC",4,"PD",
1134,2,"INV",1,"CR",
1134,2,"IRC",5,"NE",
1135,1,"INV",1,"CR",
1135,1,"IRC",4,"PD",
1136,2,"INV",3,"SD",
1136,2,"IRC",4,"PD",
1137,1,"INV",4,"PD",
1137,1,"IRC",4,"PD",
1138,1,"INV",5,"NE",
1138,1,"IRC",3,"SD",
1139,2,"INV",2,"PR",
1139,2,"IRC",4,"PD",
1140,1,"INV",1,"CR",
1140,1,"IRC",4,"PD",
1141,1,"INV",3,"SD",
1141,1,"IRC",2,"PR",
1142,2,"INV",2,"PR",
1142,2,"IRC",1,"CR",
1143,2,"INV",3,"SD",
1143,2,"IRC",1,"CR",
1144,2,"INV",3,"SD",
1144,2,"IRC",5,"NE",
1145,1,"INV",2,"PR",
1145,1,"IRC",1,"CR",
1146,1,"INV",4,"PD",
1146,1,"IRC",2,"PR",
1147,2,"INV",1,"CR",
1147,2,"IRC",5,"NE",
1148,1,"INV",1,"CR",
1148,1,"IRC",3,"SD",
1149,1,"INV",5,"NE",
1149,1,"IRC",5,"NE",
1150,2,"INV",2,"PR",
1150,2,"IRC",4,"PD",
1151,1,"INV",3,"SD",
1151,1,"IRC",4,"PD",
1152,1,"INV",5,"NE",
1152,1,"IRC",4,"PD",
1153,1,"INV",3,"SD",
1153,1,"IRC",5,"NE",
1154,2,"INV",4,"PD",
1154,2,"IRC",1,"CR",
1155,1,"INV",2,"PR",
1155,1,"IRC",1,"CR",
1156,1,"INV",4,"PD",
1156,1,"IRC",4,"PD",
1157,1,"INV",4,"PD",
1157,1,"IRC",5,"NE",
1158,1,"INV",3,"SD",
1158,1,"IRC",5,"NE",
1159,2,"INV",3,"SD",
1159,2,"IRC",1,"CR",
1160,2,"INV",1,"CR",
1160,2,"IRC",2,"PR",
1161,1,"INV",3,"SD",
1161,1,"IRC",2,"PR",
1162,2,"INV",2,"PR",
1162,2,"IRC",2,"PR",
1163,1,"INV",5,"NE",
1163,1,"IRC",4,"PD",
1164,2,"INV",4,"PD",
1164,2,"IRC",2,"PR",
1165,1,"INV",2,"PR",
1165,1,"IRC",3,"SD",
1166,1,"INV",1,"CR",
1166,1,"IRC",4,"PD",
1167,1,"INV",3,"SD",
1167,1,"IRC",2,"PR",
1168,1,"INV",3,"SD",
1168,1,"IRC",4,"PD",
1169,1,"INV",1,"CR",
1169,1,"IRC",3,"SD",
)


# Simulate twice the original subject count while preserving treatment and the
# paired independent-review/investigator response structure.
maps <- make_id_maps(list(ADSL=adsl, ADRS=adrs))
adsl <- double_subjects(adsl, maps)
adrs <- double_subjects(adrs, maps)
resp_text <- c("CR","PR","SD","PD","NE")
change <- which(runif(nrow(adrs))<0.04)
if(length(change)) {
  adrs$aval[change] <- sample(1:5,length(change),replace=TRUE)
  adrs$avalc[change] <- resp_text[adrs$aval[change]]
}
stopifnot(n_distinct(adsl$usubjid)==2L*maps$original_n)
write_xlsx_list(list(ADSL=adsl,ADRS=adrs),file.path(OUT_DIR,paste0(STEM,"_simulated_data.xlsx")))

# Original concordance analysis retained from the source document.
#==============================================================================;
#Define the response format;
#==============================================================================;
resp_format <- tribble(
 ~resp_code, ~resp_text,
 1, "CR",
 2, "PR",
 3, "SD",
 4, "PD",
 5, "NE"
)
#==============================================================================;
# Process ADRS dataset to transpose parameters as variables
#==============================================================================;
adrs01 <- adrs %>%
 mutate(paramcd=str_to_lower(paramcd)) %>%
 pivot_wider(
 id_cols = c(usubjid, trt01pn),
 names_from = paramcd,
 values_from = avalc
 )
#==============================================================================;
# Process adrs and adsl datasets
#==============================================================================;
adrs02 <- bind_rows(
 adrs01 %>% mutate(treatment=trt01pn),
 adrs01 %>% mutate(treatment=3)
)
adsl02 <- bind_rows(
 adsl %>% mutate(treatment=trt01pn),
 adsl %>% mutate(treatment=3)
)
#==============================================================================;
# Create a dummy dataset to have all levels
#==============================================================================;
dummy01 <- expand_grid(
 treatment = 1:3,
 ircn = 1:5,
 invn = 1:5
) %>%
 mutate(
 irc = resp_format$resp_text[ircn],
 inv = resp_format$resp_text[invn]
 )
#==============================================================================;
# Fetch treatment totals
#==============================================================================;
denoms <- adsl02 %>%
 group_by(treatment) %>%
 summarize(denom = n_distinct(usubjid), .groups="drop")
#==============================================================================;
#Process for counts;
#==============================================================================;
# Create concnum01 and concdenom01 tables
concnum01 <- adrs02 %>%
 filter(irc == inv & inv != "" & irc != "") %>%
 group_by(treatment) %>%
 summarise(num = n_distinct(usubjid), .groups="drop")
concdenom01 <- adrs02 %>%
 filter(inv != "" & irc != "") %>%
 group_by(treatment) %>%
 summarise(denom = n_distinct(usubjid), .groups="drop")
# Create conc01 by joining concnum01 and concdenom01
conc01 <- full_join(concnum01, concdenom01, by = "treatment")
# Create concdummy01
concdummy01 <- tibble(treatment = 1:3)
# Get IRC vs INV counts and merge with dummy dataset containing all levels
counts01 <- adrs02 %>%
 group_by(treatment, irc, inv) %>%
 summarise(count = n(), .groups="drop") %>%
 ungroup() %>%
 mutate(count = ifelse(is.na(count), 0, count))
counts02 <- full_join(dummy01, counts01, by = c("treatment", "irc", "inv")) %>%
 mutate(count = ifelse(is.na(count), 0, count))
# Calculate percentages using the number of subjects with irc and inv data
counts03 <- counts02 %>%
 left_join(concdenom01, by = "treatment") %>%
 mutate(
 cp = ifelse(count != 0,
 paste(count, " (", sprintf("%.1f",count/denom*100), ")", sep = ""), "0")
 ) %>%
 select(treatment, ircn, irc, inv, cp)
# Transpose the data such that investigator result levels become columns
trans01 <- counts03 %>%
 pivot_wider(
 id_cols = c(treatment, ircn, irc),
 values_from = cp,
 names_from = inv,
 ) %>%
 rename_all(tolower)
# Create a row for overall concordance and combine with irc/inv level counts
conc02 <- conc01 %>%
 mutate(
 num = ifelse(is.na(num), 0, num),
 denom = ifelse(is.na(denom), 0, denom),
 sd = paste(num, "(", sprintf("%.1f",num/denom*100), ")", sep = "")
 ) %>%
 select(treatment, sd) %>%
 mutate(ircn=99)
trans02 <- trans01 %>% mutate(group=1) %>%
 bind_rows(conc02 %>% mutate(group=2)) %>%
 mutate(
 irc = ifelse(group == 2, "Overall Concordance Rate", irc)
 )
output <- trans02 %>%
 arrange(treatment,group,ircn)
#==============================================================================;

trt_n <- denoms$denom; names(trt_n)<-denoms$treatment
display <- output %>% mutate(
  Treatment=case_when(
    treatment==1 ~ paste0("Treatment 1 (N=",trt_n["1"],")"),
    treatment==2 ~ paste0("Treatment 2 (N=",trt_n["2"],")"),
    treatment==3 ~ paste0("Total (N=",trt_n["3"],")")
  ),
  Concordance=ifelse(group==2,sd,""),
  `Investigator SD`=ifelse(group==2,"",sd)
) %>%
mutate(
  Treatment=if_else(
    row_number()==1L | Treatment!=lag(Treatment),
    Treatment,
    ""
  )
) %>%
transmute(
  Treatment,
  `Independent Review Committee Response`=irc,
  `Investigator CR`=coalesce(cr,""),
  `Investigator PR`=coalesce(pr,""),
  `Investigator SD`,
  `Investigator PD`=coalesce(pd,""),
  `Investigator NE`=coalesce(ne,""),
  `Overall Concordance n (%)`=Concordance
)
write_clinical_rtf(display,
  "Table 1. Concordance of Independent Review Committee and Investigator Response",
  file.path(OUT_DIR,paste0(STEM,".rtf")),
  "Cells show n (percent) among subjects with both IRC and investigator assessments. CR = complete response; PR = partial response; SD = stable disease; PD = progressive disease; NE = not evaluable."
)
cat("Created ",STEM,": simulated N = ",2L*maps$original_n,"\n",sep="")
