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
adae<-tribble(
~usubjid,~trtan,~saffl,~astdt,~aedecod,~trtemfl,
"PILOT-1001",2,"Y",19723,"HEADACHE","Y",
"PILOT-1001",2,"Y",19724,"STUPOR","Y",
"PILOT-1001",2,"Y",19724,"TRANSIENT ISCHAEMIC ATTACK","Y",
"PILOT-1001",2,"Y",19724,"VISION BLURRED","Y",
"PILOT-1001",2,"Y",19724,"DIZZINESS","Y",
"PILOT-1002",1,"Y",19114,"ANXIETY","N",
"PILOT-1002",1,"Y",19730,"DIARRHOEA","Y",
"PILOT-1002",1,"Y",19737,"EXCORIATION","Y",
"PILOT-1002",1,"Y",19738,"IRRITABILITY","Y",
"PILOT-1002",1,"Y",19737,"EXCORIATION","Y",
"PILOT-1002",1,"Y",19786,"PRURITUS","Y",
"PILOT-1002",1,"Y",19114,"ANXIETY","N",
"PILOT-1002",1,"Y",19738,"IRRITABILITY","Y",
"PILOT-1002",1,"Y",19786,"PRURITUS","Y",
"PILOT-1003",2,"Y",19627,"APPLICATION SITE ERYTHEMA","Y",
"PILOT-1003",2,"Y",19642,"NAUSEA","Y",
"PILOT-1003",2,"Y",19642,"VOMITING","Y",
"PILOT-1003",2,"Y",19627,"APPLICATION SITE ERYTHEMA","Y",
"PILOT-1003",2,"Y",19642,"NAUSEA","Y",
"PILOT-1003",2,"Y",19642,"VOMITING","Y",
"PILOT-1003",2,"Y",19790,"CATARACT OPERATION","Y",
"PILOT-1004",2,"Y",19531,"VOMITING","Y",
"PILOT-1004",2,"Y",19541,"NASOPHARYNGITIS","Y",
"PILOT-1004",2,"Y",19541,"SKIN IRRITATION","Y",
"PILOT-1004",2,"Y",19541,"PRURITUS","Y",
"PILOT-1004",2,"Y",19541,"NASOPHARYNGITIS","Y",
"PILOT-1004",2,"Y",19561,"RASH","Y",
"PILOT-1004",2,"Y",19573,"ERYTHEMA","Y",
"PILOT-1004",2,"Y",19583,"SKIN EXFOLIATION","Y",
"PILOT-1004",2,"Y",19541,"SKIN IRRITATION","Y",
"PILOT-1004",2,"Y",19541,"PRURITUS","Y",
"PILOT-1004",2,"Y",19561,"RASH","Y",
"PILOT-1004",2,"Y",19573,"ERYTHEMA","Y",
"PILOT-1004",2,"Y",19583,"SKIN EXFOLIATION","Y",
"PILOT-1005",2,"Y",19386,"PRURITUS","Y",
"PILOT-1005",2,"Y",19386,"RASH","Y",
"PILOT-1005",2,"Y",19386,"PRURITUS","Y",
"PILOT-1005",2,"Y",19386,"RASH","Y",
"PILOT-1006",1,"Y",19360,"NASAL CONGESTION","Y",
"PILOT-1006",1,"Y",19362,"INSOMNIA","Y",
"PILOT-1006",1,"Y",19362,"INSOMNIA","Y",
"PILOT-1006",1,"Y",19379,"APPLICATION SITE INDURATION","Y",
"PILOT-1006",1,"Y",19381,"APPLICATION SITE PRURITUS","Y",
"PILOT-1006",1,"Y",19389,"APPLICATION SITE PRURITUS","Y",
"PILOT-1006",1,"Y",19397,"FALL","Y",
"PILOT-1006",1,"Y",19398,"FALL","Y",
"PILOT-1007",1,"Y",19704,"APPLICATION SITE DERMATITIS","Y",
"PILOT-1007",1,"Y",19719,"SKIN IRRITATION","Y",
"PILOT-1007",1,"Y",19735,"AGITATION","Y",
"PILOT-1007",1,"Y",NA,"ARTHRITIS","N",
"PILOT-1007",1,"Y",NA,"ARTHRITIS","N",
"PILOT-1008",2,"Y",19271,"CONFUSIONAL STATE","Y",
"PILOT-1008",2,"Y",19271,"APPLICATION SITE IRRITATION","Y",
"PILOT-1008",2,"Y",19271,"APPLICATION SITE IRRITATION","Y",
"PILOT-1009",1,"Y",19377,"HEART RATE IRREGULAR","Y",
"PILOT-1009",1,"Y",19377,"HEART RATE IRREGULAR","Y",
"PILOT-1009",1,"Y",19393,"DIZZINESS","Y",
"PILOT-1009",1,"Y",19405,"HEADACHE","Y",
"PILOT-1009",1,"Y",19433,"HEART RATE IRREGULAR","Y",
"PILOT-1009",1,"Y",19433,"HEART RATE IRREGULAR","Y",
"PILOT-1010",1,"Y",NA,"COUGH","N",
"PILOT-1011",2,"Y",19433,"APPLICATION SITE IRRITATION","Y",
"PILOT-1011",2,"Y",19433,"APPLICATION SITE IRRITATION","Y",
"PILOT-1011",2,"Y",19482,"PRURITUS","Y",
"PILOT-1011",2,"Y",19482,"PRURITUS","Y",
"PILOT-1012",2,"Y",19700,"RASH","Y",
"PILOT-1012",2,"Y",19700,"RASH","Y",
"PILOT-1013",1,"Y",19779,"AGITATION","Y",
"PILOT-1013",1,"Y",19804,"ERYTHEMA","Y",
"PILOT-1014",2,"Y",19314,"PRURITUS","Y",
"PILOT-1014",2,"Y",19314,"SWELLING","Y",
"PILOT-1014",2,"Y",19330,"BLISTER","Y",
"PILOT-1014",2,"Y",19427,"HYPERSENSITIVITY","Y",
"PILOT-1014",2,"Y",19427,"HYPERSENSITIVITY","Y",
"PILOT-1015",2,"Y",19638,"DIZZINESS","Y",
"PILOT-1015",2,"Y",19642,"HYPERHIDROSIS","Y",
"PILOT-1015",2,"Y",19642,"BODY TEMPERATURE INCREASED","Y",
"PILOT-1015",2,"Y",19642,"HYPERTENSION","Y",
"PILOT-1015",2,"Y",19643,"SYNCOPE","Y",
"PILOT-1015",2,"Y",19647,"AGITATION","Y",
"PILOT-1017",1,"Y",19589,"DYSURIA","Y",
"PILOT-1017",1,"Y",19592,"DYSPEPSIA","Y",
"PILOT-1017",1,"Y",19598,"URINARY TRACT INFECTION","Y",
"PILOT-1017",1,"Y",19606,"ASTHENIA","Y",
"PILOT-1017",1,"Y",19598,"URINARY TRACT INFECTION","Y",
"PILOT-1017",1,"Y",19606,"ASTHENIA","Y",
"PILOT-1017",1,"Y",19613,"DYSPEPSIA","Y",
"PILOT-1017",1,"Y",19641,"PELVIC PAIN","Y",
"PILOT-1017",1,"Y",19641,"GASTROENTERITIS VIRAL","Y",
"PILOT-1017",1,"Y",19641,"GASTROOESOPHAGEAL REFLUX DISEASE","Y",
"PILOT-1017",1,"Y",19685,"VAGINAL MYCOSIS","Y",
"PILOT-1017",1,"Y",19685,"CERVICITIS","Y",
"PILOT-1017",1,"Y",19685,"VAGINAL MYCOSIS","Y",
"PILOT-1017",1,"Y",19685,"CERVICITIS","Y",
"PILOT-1017",1,"Y",19641,"PELVIC PAIN","Y",
"PILOT-1018",2,"Y",19662,"APPLICATION SITE ERYTHEMA","Y",
"PILOT-1018",2,"Y",19662,"APPLICATION SITE ERYTHEMA","Y",
"PILOT-1018",2,"Y",19684,"SOMNOLENCE","Y",
"PILOT-1018",2,"Y",19684,"SOMNOLENCE","Y",
"PILOT-1019",1,"Y",19792,"SINUS BRADYCARDIA","Y",
"PILOT-1019",1,"Y",19792,"ATRIOVENTRICULAR BLOCK FIRST DEGREE","Y",
"PILOT-1020",2,"Y",19405,"ERYTHEMA","Y",
"PILOT-1020",2,"Y",19405,"PRURITUS","Y",
"PILOT-1020",2,"Y",19405,"ERYTHEMA","Y",
"PILOT-1020",2,"Y",19405,"PRURITUS","Y",
"PILOT-1020",2,"Y",19421,"URTICARIA","Y",
"PILOT-1020",2,"Y",19421,"URTICARIA","Y",
"PILOT-1020",2,"Y",19438,"APPLICATION SITE URTICARIA","Y",
"PILOT-1020",2,"Y",19438,"URTICARIA","Y",
"PILOT-1021",1,"Y",19687,"SINUS ARRHYTHMIA","Y",
"PILOT-1021",1,"Y",19688,"BLOOD CREATINE PHOSPHOKINASE INCREASED","Y",
"PILOT-1021",1,"Y",19695,"POSTNASAL DRIP","Y",
"PILOT-1021",1,"Y",19688,"BLOOD CREATINE PHOSPHOKINASE INCREASED","Y",
"PILOT-1021",1,"Y",19695,"POSTNASAL DRIP","Y",
"PILOT-1021",1,"Y",19687,"SINUS ARRHYTHMIA","Y",
"PILOT-1021",1,"Y",19762,"CONJUNCTIVITIS","Y",
"PILOT-1021",1,"Y",19762,"CONJUNCTIVITIS","Y",
"PILOT-1024",2,"Y",19319,"COUGH","Y",
"PILOT-1024",2,"Y",19371,"SHOULDER PAIN","Y",
"PILOT-1024",2,"Y",19387,"APPLICATION SITE PRURITUS","Y",
"PILOT-1025",1,"Y",19296,"APPLICATION SITE DERMATITIS","Y",
"PILOT-1025",1,"Y",19296,"APPLICATION SITE DERMATITIS","Y",
"PILOT-1026",2,"Y",19462,"HEMIANOPIA HOMONYMOUS","Y",
"PILOT-1026",2,"Y",19467,"APPLICATION SITE DERMATITIS","Y",
"PILOT-1026",2,"Y",19467,"APPLICATION SITE PRURITUS","Y",
"PILOT-1026",2,"Y",19497,"SINUS BRADYCARDIA","Y",
"PILOT-1026",2,"Y",19467,"APPLICATION SITE DERMATITIS","Y",
"PILOT-1026",2,"Y",19467,"APPLICATION SITE PRURITUS","Y",
"PILOT-1027",2,"Y",19659,"PRURITUS","Y",
"PILOT-1027",2,"Y",19767,"ERYTHEMA","Y",
"PILOT-1028",2,"Y",19470,"ERYTHEMA","Y",
"PILOT-1028",2,"Y",19470,"PRURITUS","Y",
"PILOT-1028",2,"Y",19517,"BLISTER","Y",
"PILOT-1028",2,"Y",19470,"ERYTHEMA","Y",
"PILOT-1028",2,"Y",19470,"PRURITUS","Y",
"PILOT-1028",2,"Y",19517,"BLISTER","Y",
"PILOT-1029",1,"Y",19833,"RASH","Y",
"PILOT-1029",1,"Y",19833,"APPLICATION SITE PRURITUS","Y",
"PILOT-1029",1,"Y",19833,"RASH","Y",
"PILOT-1029",1,"Y",19833,"APPLICATION SITE PRURITUS","Y",
"PILOT-1029",1,"Y",19847,"APPLICATION SITE IRRITATION","Y",
"PILOT-1029",1,"Y",19863,"DRUG ERUPTION","Y",
"PILOT-1029",1,"Y",19833,"RASH","Y",
"PILOT-1029",1,"Y",19833,"APPLICATION SITE PRURITUS","Y",
"PILOT-1029",1,"Y",19847,"APPLICATION SITE IRRITATION","Y",
"PILOT-1030",2,"Y",19871,"ERYTHEMA","Y",
"PILOT-1030",2,"Y",19926,"ABDOMINAL PAIN","Y",
"PILOT-1030",2,"Y",19937,"ULCER","Y",
"PILOT-1030",2,"Y",19871,"ERYTHEMA","Y",
"PILOT-1031",1,"Y",19860,"APPLICATION SITE DERMATITIS","Y",
"PILOT-1031",1,"Y",20015,"ELECTROCARDIOGRAM ST SEGMENT DEPRESSION","Y",
"PILOT-1032",1,"Y",19752,"HEADACHE","Y",
"PILOT-1032",1,"Y",19769,"APPLICATION SITE ERYTHEMA","Y",
"PILOT-1033",2,"Y",19377,"APPLICATION SITE PRURITUS","Y",
"PILOT-1033",2,"Y",19377,"APPLICATION SITE ERYTHEMA","Y",
"PILOT-1033",2,"Y",19384,"APPLICATION SITE PRURITUS","Y",
"PILOT-1033",2,"Y",19384,"APPLICATION SITE ERYTHEMA","Y",
"PILOT-1033",2,"Y",19384,"APPLICATION SITE IRRITATION","Y",
"PILOT-1034",2,"Y",19588,"DIZZINESS","Y",
"PILOT-1034",2,"Y",19588,"DIZZINESS","Y",
"PILOT-1034",2,"Y",19598,"HYPOTENSION","Y",
"PILOT-1035",1,"Y",19447,"SKIN IRRITATION","Y",
"PILOT-1035",1,"Y",19452,"EAR PAIN","Y",
"PILOT-1035",1,"Y",19447,"SKIN IRRITATION","Y",
"PILOT-1035",1,"Y",19452,"EAR PAIN","Y",
"PILOT-1035",1,"Y",19507,"PRURITUS","Y",
"PILOT-1035",1,"Y",19525,"ERYTHEMA","Y",
"PILOT-1035",1,"Y",19507,"PRURITUS","Y",
"PILOT-1035",1,"Y",19525,"ERYTHEMA","Y",
"PILOT-1036",1,"Y",19228,"HEADACHE","N",
"PILOT-1038",1,"Y",19364,"PARKINSON'S DISEASE","Y",
"PILOT-1039",1,"Y",20027,"COMPLETED SUICIDE","Y",
"PILOT-1040",2,"Y",19801,"APPLICATION SITE IRRITATION","Y",
"PILOT-1040",2,"Y",19809,"NAUSEA","Y",
"PILOT-1040",2,"Y",19801,"APPLICATION SITE IRRITATION","Y",
"PILOT-1040",2,"Y",19809,"NAUSEA","Y",
)

adsl<-tribble(
~usubjid,~trt01an,~saffl,
"PILOT-1001",2,"Y",
"PILOT-1002",1,"Y",
"PILOT-1003",2,"Y",
"PILOT-1004",2,"Y",
"PILOT-1005",2,"Y",
"PILOT-1006",1,"Y",
"PILOT-1007",1,"Y",
"PILOT-1008",2,"Y",
"PILOT-1009",1,"Y",
"PILOT-1010",1,"Y",
"PILOT-1011",2,"Y",
"PILOT-1012",2,"Y",
"PILOT-1013",1,"Y",
"PILOT-1014",2,"Y",
"PILOT-1015",2,"Y",
"PILOT-1016",1,"Y",
"PILOT-1017",1,"Y",
"PILOT-1018",2,"Y",
"PILOT-1019",1,"Y",
"PILOT-1020",2,"Y",
"PILOT-1021",1,"Y",
"PILOT-1022",1,"Y",
"PILOT-1023",1,"Y",
"PILOT-1024",2,"Y",
"PILOT-1025",1,"Y",
"PILOT-1026",2,"Y",
"PILOT-1027",2,"Y",
"PILOT-1028",2,"Y",
"PILOT-1029",1,"Y",
"PILOT-1030",2,"Y",
"PILOT-1031",1,"Y",
"PILOT-1032",1,"Y",
"PILOT-1033",2,"Y",
"PILOT-1034",2,"Y",
"PILOT-1035",1,"Y",
"PILOT-1036",1,"Y",
"PILOT-1037",2,"Y",
"PILOT-1038",1,"Y",
"PILOT-1039",1,"Y",
"PILOT-1040",2,"Y",
)

# Analysis and outputs.

maps<-make_id_maps(list(ADSL=adsl,ADAE=adae)); adsl<-double_subjects(adsl,maps); adae<-double_subjects(adae,maps)
adae$astdt<-ifelse(is.na(adae$astdt),NA,adae$astdt+sample(-2:2,nrow(adae),replace=TRUE))
stopifnot(n_distinct(adsl$usubjid)==2L*maps$original_n)
write_xlsx_list(list(ADSL=adsl,ADAE=adae),file.path(OUT_DIR,paste0(STEM,"_simulated_data.xlsx")))
adsl01<-adsl%>%filter(saffl=="Y")%>%mutate(treatment=trt01an); adae01<-adae%>%filter(saffl=="Y",trtemfl=="Y")%>%mutate(treatment=trtan)
den<-adsl01%>%count(treatment,name="N"); terms<-c("Overall",sort(unique(adae01$aedecod)))
rows<-lapply(terms,function(term){z<-if(term=="Overall")adae01 else filter(adae01,aedecod==term); cnt<-sapply(1:2,function(g)n_distinct(z$usubjid[z$treatment==g])); ev<-sapply(1:2,function(g)sum(z$treatment==g)); Ns<-sapply(1:2,function(g)den$N[match(g,den$treatment)]); pv<-tryCatch(fisher.test(rbind(cnt,Ns-cnt))$p.value,error=function(e)NA_real_);data.frame(`Adverse Event`=term,`Treatment 1 n (%) [E]`=paste0(fmt_np(cnt[1],Ns[1])," [",ev[1],"]"),`Treatment 2 n (%) [E]`=paste0(fmt_np(cnt[2],Ns[2])," [",ev[2],"]"),`Fisher p-value`=fmt_p(pv),check.names=FALSE)}); output<-bind_rows(rows)
write_clinical_rtf(output,"Table 1. Treatment-Emergent Adverse Events with Fisher Exact Test",file.path(OUT_DIR,paste0(STEM,".rtf")),"n = subjects with event; E = number of events; percentages use treatment-group safety denominators.")

cat("Created 01_adverse_events_fisher_exact: simulated N =", 2L*maps$original_n, "\n")
