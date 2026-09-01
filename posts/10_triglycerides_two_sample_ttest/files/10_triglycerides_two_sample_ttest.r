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
~usubjid,~trt01p,~trt01pn,
"101","Placebo",0,
"101","Placebo",0,
"1010","Placebo",0,
"1010","Placebo",0,
"1011","Placebo",0,
"1011","Placebo",0,
"1012","Placebo",0,
"1012","Placebo",0,
"1013","Placebo",0,
"1013","Placebo",0,
"1014","Placebo",0,
"1014","Placebo",0,
"1015","Placebo",0,
"1015","Placebo",0,
"1016","Placebo",0,
"1016","Placebo",0,
"1017","Placebo",0,
"1017","Placebo",0,
"1018","Placebo",0,
"1018","Placebo",0,
"1019","Placebo",0,
"1019","Placebo",0,
"102","Placebo",0,
"102","Placebo",0,
"1020","Placebo",0,
"1020","Placebo",0,
"1021","Placebo",0,
"1021","Placebo",0,
"1022","Placebo",0,
"1022","Placebo",0,
"1023","Placebo",0,
"1023","Placebo",0,
"1024","Placebo",0,
"1024","Placebo",0,
"1025","Placebo",0,
"1025","Placebo",0,
"1026","Placebo",0,
"1026","Placebo",0,
"1027","Placebo",0,
"1027","Placebo",0,
"1028","Placebo",0,
"1028","Placebo",0,
"1029","Placebo",0,
"1029","Placebo",0,
"103","Placebo",0,
"103","Placebo",0,
"1030","Placebo",0,
"1030","Placebo",0,
"104","Placebo",0,
"104","Placebo",0,
"105","Placebo",0,
"105","Placebo",0,
"106","Placebo",0,
"106","Placebo",0,
"107","Placebo",0,
"107","Placebo",0,
"108","Placebo",0,
"108","Placebo",0,
"109","Placebo",0,
"109","Placebo",0,
"201","Test Drug",1,
"201","Test Drug",1,
"2010","Test Drug",1,
"2010","Test Drug",1,
"2011","Test Drug",1,
"2011","Test Drug",1,
"2012","Test Drug",1,
"2012","Test Drug",1,
"2013","Test Drug",1,
"2013","Test Drug",1,
"2014","Test Drug",1,
"2014","Test Drug",1,
"2015","Test Drug",1,
"2015","Test Drug",1,
"2016","Test Drug",1,
"2016","Test Drug",1,
"2017","Test Drug",1,
"2017","Test Drug",1,
"2018","Test Drug",1,
"2018","Test Drug",1,
"2019","Test Drug",1,
"2019","Test Drug",1,
"202","Test Drug",1,
"202","Test Drug",1,
"2020","Test Drug",1,
"2020","Test Drug",1,
"2021","Test Drug",1,
"2021","Test Drug",1,
"2022","Test Drug",1,
"2022","Test Drug",1,
"2023","Test Drug",1,
"2023","Test Drug",1,
"2024","Test Drug",1,
"2024","Test Drug",1,
"2025","Test Drug",1,
"2025","Test Drug",1,
"2026","Test Drug",1,
"2026","Test Drug",1,
"2027","Test Drug",1,
"2027","Test Drug",1,
"2028","Test Drug",1,
"2028","Test Drug",1,
"2029","Test Drug",1,
"2029","Test Drug",1,
"203","Test Drug",1,
"203","Test Drug",1,
"2030","Test Drug",1,
"2030","Test Drug",1,
"204","Test Drug",1,
"204","Test Drug",1,
"205","Test Drug",1,
"205","Test Drug",1,
"206","Test Drug",1,
"206","Test Drug",1,
"207","Test Drug",1,
"207","Test Drug",1,
"208","Test Drug",1,
"208","Test Drug",1,
"209","Test Drug",1,
"209","Test Drug",1,
)

adlb<-tribble(
~usubjid,~trtp,~trtpn,~avisit,~aval,~chg,
"101","Placebo",0,"Baseline",170,NA,
"101","Placebo",0,"Week 12",170,0,
"102","Placebo",0,"Baseline",150,NA,
"102","Placebo",0,"Week 12",145,-5,
"103","Placebo",0,"Baseline",158,NA,
"103","Placebo",0,"Week 12",157,-1,
"104","Placebo",0,"Baseline",157,NA,
"104","Placebo",0,"Week 12",156,-1,
"105","Placebo",0,"Baseline",173,NA,
"105","Placebo",0,"Week 12",173,0,
"106","Placebo",0,"Baseline",171,NA,
"106","Placebo",0,"Week 12",166,-5,
"107","Placebo",0,"Baseline",178,NA,
"107","Placebo",0,"Week 12",174,-4,
"108","Placebo",0,"Baseline",152,NA,
"108","Placebo",0,"Week 12",148,-4,
"109","Placebo",0,"Baseline",163,NA,
"109","Placebo",0,"Week 12",163,0,
"1010","Placebo",0,"Baseline",150,NA,
"1010","Placebo",0,"Week 12",150,0,
"1011","Placebo",0,"Baseline",156,NA,
"1011","Placebo",0,"Week 12",155,-1,
"1012","Placebo",0,"Baseline",166,NA,
"1012","Placebo",0,"Week 12",162,-4,
"1013","Placebo",0,"Baseline",150,NA,
"1013","Placebo",0,"Week 12",146,-4,
"1014","Placebo",0,"Baseline",156,NA,
"1014","Placebo",0,"Week 12",151,-5,
"1015","Placebo",0,"Baseline",170,NA,
"1015","Placebo",0,"Week 12",165,-5,
"1016","Placebo",0,"Baseline",167,NA,
"1016","Placebo",0,"Week 12",164,-3,
"1017","Placebo",0,"Baseline",157,NA,
"1017","Placebo",0,"Week 12",154,-3,
"1018","Placebo",0,"Baseline",168,NA,
"1018","Placebo",0,"Week 12",166,-2,
"1019","Placebo",0,"Baseline",175,NA,
"1019","Placebo",0,"Week 12",175,0,
"1020","Placebo",0,"Baseline",174,NA,
"1020","Placebo",0,"Week 12",173,-1,
"1021","Placebo",0,"Baseline",172,NA,
"1021","Placebo",0,"Week 12",169,-3,
"1022","Placebo",0,"Baseline",160,NA,
"1022","Placebo",0,"Week 12",158,-2,
"1023","Placebo",0,"Baseline",154,NA,
"1023","Placebo",0,"Week 12",153,-1,
"1024","Placebo",0,"Baseline",180,NA,
"1024","Placebo",0,"Week 12",178,-2,
"1025","Placebo",0,"Baseline",153,NA,
"1025","Placebo",0,"Week 12",153,0,
"1026","Placebo",0,"Baseline",162,NA,
"1026","Placebo",0,"Week 12",162,0,
"1027","Placebo",0,"Baseline",161,NA,
"1027","Placebo",0,"Week 12",159,-2,
"1028","Placebo",0,"Baseline",169,NA,
"1028","Placebo",0,"Week 12",167,-2,
"1029","Placebo",0,"Baseline",175,NA,
"1029","Placebo",0,"Week 12",175,0,
"1030","Placebo",0,"Baseline",173,NA,
"1030","Placebo",0,"Week 12",170,-3,
"201","Test Drug",1,"Baseline",167,NA,
"201","Test Drug",1,"Week 12",149,-18,
"202","Test Drug",1,"Baseline",179,NA,
"202","Test Drug",1,"Week 12",152,-27,
"203","Test Drug",1,"Baseline",152,NA,
"203","Test Drug",1,"Week 12",128,-24,
"204","Test Drug",1,"Baseline",176,NA,
"204","Test Drug",1,"Week 12",150,-26,
"205","Test Drug",1,"Baseline",168,NA,
"205","Test Drug",1,"Week 12",147,-21,
"206","Test Drug",1,"Baseline",172,NA,
"206","Test Drug",1,"Week 12",155,-17,
"207","Test Drug",1,"Baseline",151,NA,
"207","Test Drug",1,"Week 12",129,-22,
"208","Test Drug",1,"Baseline",174,NA,
"208","Test Drug",1,"Week 12",150,-24,
"209","Test Drug",1,"Baseline",152,NA,
"209","Test Drug",1,"Week 12",130,-22,
"2010","Test Drug",1,"Baseline",177,NA,
"2010","Test Drug",1,"Week 12",159,-18,
"2011","Test Drug",1,"Baseline",162,NA,
"2011","Test Drug",1,"Week 12",139,-23,
"2012","Test Drug",1,"Baseline",164,NA,
"2012","Test Drug",1,"Week 12",138,-26,
"2013","Test Drug",1,"Baseline",155,NA,
"2013","Test Drug",1,"Week 12",129,-26,
"2014","Test Drug",1,"Baseline",161,NA,
"2014","Test Drug",1,"Week 12",140,-21,
"2015","Test Drug",1,"Baseline",171,NA,
"2015","Test Drug",1,"Week 12",148,-23,
"2016","Test Drug",1,"Baseline",172,NA,
"2016","Test Drug",1,"Week 12",155,-17,
"2017","Test Drug",1,"Baseline",169,NA,
"2017","Test Drug",1,"Week 12",149,-20,
"2018","Test Drug",1,"Baseline",167,NA,
"2018","Test Drug",1,"Week 12",145,-22,
"2019","Test Drug",1,"Baseline",155,NA,
"2019","Test Drug",1,"Week 12",126,-29,
"2020","Test Drug",1,"Baseline",162,NA,
"2020","Test Drug",1,"Week 12",139,-23,
"2021","Test Drug",1,"Baseline",179,NA,
"2021","Test Drug",1,"Week 12",157,-22,
"2022","Test Drug",1,"Baseline",171,NA,
"2022","Test Drug",1,"Week 12",146,-25,
"2023","Test Drug",1,"Baseline",176,NA,
"2023","Test Drug",1,"Week 12",160,-16,
"2024","Test Drug",1,"Baseline",157,NA,
"2024","Test Drug",1,"Week 12",141,-16,
"2025","Test Drug",1,"Baseline",175,NA,
"2025","Test Drug",1,"Week 12",150,-25,
"2026","Test Drug",1,"Baseline",162,NA,
"2026","Test Drug",1,"Week 12",139,-23,
"2027","Test Drug",1,"Baseline",152,NA,
"2027","Test Drug",1,"Week 12",131,-21,
"2028","Test Drug",1,"Baseline",179,NA,
"2028","Test Drug",1,"Week 12",154,-25,
"2029","Test Drug",1,"Baseline",156,NA,
"2029","Test Drug",1,"Week 12",126,-30,
"2030","Test Drug",1,"Baseline",162,NA,
"2030","Test Drug",1,"Week 12",133,-29,
)

# Analysis and outputs.

maps<-make_id_maps(list(ADSL=adsl,ADLB=adlb));adsl<-double_subjects(adsl,maps);adlb<-double_subjects(adlb,maps);adlb<-adlb%>%group_by(usubjid)%>%mutate(aval=round(pmax(1,aval+rnorm(n(),0,2)),1),chg=ifelse(avisit=="Baseline",NA,aval-first(aval[avisit=="Baseline"])))%>%ungroup();stopifnot(n_distinct(adlb$usubjid)==2L*maps$original_n);write_xlsx_list(list(ADSL=adsl,ADLB=adlb),file.path(OUT_DIR,paste0(STEM,"_simulated_data.xlsx")))
z<-adlb%>%filter(avisit=="Week 12")%>%mutate(treatment=factor(trtpn));tt<-t.test(chg~treatment,data=z,var.equal=TRUE);s<-z%>%group_by(treatment)%>%summarise(N=sum(!is.na(chg)),Mean=mean(chg,na.rm=TRUE),SD=sd(chg,na.rm=TRUE),Median=median(chg,na.rm=TRUE),Q1=quantile(chg,.25,type=2,na.rm=TRUE),Q3=quantile(chg,.75,type=2,na.rm=TRUE),Min=min(chg,na.rm=TRUE),Max=max(chg,na.rm=TRUE),.groups="drop");output<-data.frame(Statistic=c("n (missing)","Mean (SD)","Median","Q1, Q3","Min, Max","Two-sample t-test p-value"),check.names=FALSE);for(g in levels(z$treatment)){q<-filter(s,treatment==g);output[[paste0("Treatment ",g)]]<-c(sprintf("%d (%d)",q$N,sum(is.na(z$chg[z$treatment==g]))),sprintf("%.1f (%.2f)",q$Mean,q$SD),sprintf("%.1f",q$Median),sprintf("%.1f, %.1f",q$Q1,q$Q3),sprintf("%.0f, %.0f",q$Min,q$Max),ifelse(g==levels(z$treatment)[1],fmt_p(tt$p.value),""))}
write_clinical_rtf(output,"Table 10. Triglyceride Change and Two-Sample t-Test",file.path(OUT_DIR,paste0(STEM,".rtf")),"Two-sided pooled-variance t-test compares Week 12 change between treatment groups.")

cat("Created 10_triglycerides_two_sample_ttest: simulated N =", 2L*maps$original_n, "\n")
