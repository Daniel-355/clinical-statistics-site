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
~usubjid,~trt01a,~pkfl,~trt01an,
"PILOT-10001","PILOT 20 mg","Y",1,
"PILOT-10002","PILOT 20 mg","Y",1,
"PILOT-10003","PILOT 20 mg","Y",1,
"PILOT-10004","PILOT 20 mg","Y",1,
"PILOT-10005","PILOT 20 mg","Y",1,
"PILOT-20001","PILOT 40 mg","Y",2,
"PILOT-20002","PILOT 40 mg","Y",2,
"PILOT-20003","PILOT 40 mg","Y",2,
"PILOT-20004","PILOT 40 mg","Y",2,
"PILOT-20005","PILOT 40 mg","Y",2,
)

adpc<-tribble(
~usubjid,~trt01a,~avisit,~atpt,~aval,~pkfl,~trt01an,~avisitn,~atptn,~paramcd,
"PILOT-10001","PILOT 20 mg","Day 1","Predose",0,"Y",1,1,0,"PILOT",
"PILOT-10001","PILOT 20 mg","Day 1","1 hour",48.62,"Y",1,1,1,"PILOT",
"PILOT-10001","PILOT 20 mg","Day 1","2 hours",56.48,"Y",1,1,2,"PILOT",
"PILOT-10001","PILOT 20 mg","Month 1","Predose",65.23,"Y",1,2,0,"PILOT",
"PILOT-10001","PILOT 20 mg","Month 1","1 hour",47.66,"Y",1,2,1,"PILOT",
"PILOT-10001","PILOT 20 mg","Month 1","2 hours",47.66,"Y",1,2,2,"PILOT",
"PILOT-10002","PILOT 20 mg","Day 1","Predose",0,"Y",1,1,0,"PILOT",
"PILOT-10002","PILOT 20 mg","Day 1","1 hour",57.67,"Y",1,1,1,"PILOT",
"PILOT-10002","PILOT 20 mg","Day 1","2 hours",45.31,"Y",1,1,2,"PILOT",
"PILOT-10002","PILOT 20 mg","Month 1","Predose",55.43,"Y",1,2,0,"PILOT",
"PILOT-10002","PILOT 20 mg","Month 1","1 hour",45.37,"Y",1,2,1,"PILOT",
"PILOT-10002","PILOT 20 mg","Month 1","2 hours",45.34,"Y",1,2,2,"PILOT",
"PILOT-10003","PILOT 20 mg","Day 1","Predose",0,"Y",1,1,0,"PILOT",
"PILOT-10003","PILOT 20 mg","Day 1","1 hour",30.87,"Y",1,1,1,"PILOT",
"PILOT-10003","PILOT 20 mg","Day 1","2 hours",32.75,"Y",1,1,2,"PILOT",
"PILOT-10003","PILOT 20 mg","Month 1","Predose",44.38,"Y",1,2,0,"PILOT",
"PILOT-10003","PILOT 20 mg","Month 1","1 hour",39.87,"Y",1,2,1,"PILOT",
"PILOT-10003","PILOT 20 mg","Month 1","2 hours",53.14,"Y",1,2,2,"PILOT",
"PILOT-10004","PILOT 20 mg","Day 1","Predose",0,"Y",1,1,0,"PILOT",
"PILOT-10004","PILOT 20 mg","Day 1","1 hour",35.88,"Y",1,1,1,"PILOT",
"PILOT-10004","PILOT 20 mg","Day 1","2 hours",64.66,"Y",1,1,2,"PILOT",
"PILOT-10004","PILOT 20 mg","Month 1","Predose",47.74,"Y",1,2,0,"PILOT",
"PILOT-10004","PILOT 20 mg","Month 1","1 hour",50.68,"Y",1,2,1,"PILOT",
"PILOT-10004","PILOT 20 mg","Month 1","2 hours",35.75,"Y",1,2,2,"PILOT",
"PILOT-10005","PILOT 20 mg","Day 1","Predose",0,"Y",1,1,0,"PILOT",
"PILOT-10005","PILOT 20 mg","Day 1","1 hour",51.11,"Y",1,1,1,"PILOT",
"PILOT-10005","PILOT 20 mg","Day 1","2 hours",38.49,"Y",1,1,2,"PILOT",
"PILOT-10005","PILOT 20 mg","Month 1","Predose",53.76,"Y",1,2,0,"PILOT",
"PILOT-10005","PILOT 20 mg","Month 1","1 hour",43.99,"Y",1,2,1,"PILOT",
"PILOT-10005","PILOT 20 mg","Month 1","2 hours",47.08,"Y",1,2,2,"PILOT",
"PILOT-20001","PILOT 40 mg","Day 1","Predose",0,"Y",2,1,0,"PILOT",
"PILOT-20001","PILOT 40 mg","Day 1","1 hour",127.78,"Y",2,1,1,"PILOT",
"PILOT-20001","PILOT 40 mg","Day 1","2 hours",99.8,"Y",2,1,2,"PILOT",
"PILOT-20001","PILOT 40 mg","Month 1","Predose",84.13,"Y",2,2,0,"PILOT",
"PILOT-20001","PILOT 40 mg","Month 1","1 hour",112.34,"Y",2,2,1,"PILOT",
"PILOT-20001","PILOT 40 mg","Month 1","2 hours",81.69,"Y",2,2,2,"PILOT",
"PILOT-20002","PILOT 40 mg","Day 1","Predose",0,"Y",2,1,0,"PILOT",
"PILOT-20002","PILOT 40 mg","Day 1","1 hour",70.6,"Y",2,1,1,"PILOT",
"PILOT-20002","PILOT 40 mg","Day 1","2 hours",80.08,"Y",2,1,2,"PILOT",
"PILOT-20002","PILOT 40 mg","Month 1","Predose",102.95,"Y",2,2,0,"PILOT",
"PILOT-20002","PILOT 40 mg","Month 1","1 hour",111.08,"Y",2,2,1,"PILOT",
"PILOT-20002","PILOT 40 mg","Month 1","2 hours",102.57,"Y",2,2,2,"PILOT",
"PILOT-20003","PILOT 40 mg","Day 1","Predose",0,"Y",2,1,0,"PILOT",
"PILOT-20003","PILOT 40 mg","Day 1","1 hour",95.48,"Y",2,1,1,"PILOT",
"PILOT-20003","PILOT 40 mg","Day 1","2 hours",77.82,"Y",2,1,2,"PILOT",
"PILOT-20003","PILOT 40 mg","Month 1","Predose",89.2,"Y",2,2,0,"PILOT",
"PILOT-20003","PILOT 40 mg","Month 1","1 hour",93.09,"Y",2,2,1,"PILOT",
"PILOT-20003","PILOT 40 mg","Month 1","2 hours",115.86,"Y",2,2,2,"PILOT",
"PILOT-20004","PILOT 40 mg","Day 1","Predose",0,"Y",2,1,0,"PILOT",
"PILOT-20004","PILOT 40 mg","Day 1","1 hour",73.55,"Y",2,1,1,"PILOT",
"PILOT-20004","PILOT 40 mg","Day 1","2 hours",104.86,"Y",2,1,2,"PILOT",
"PILOT-20004","PILOT 40 mg","Month 1","Predose",94.22,"Y",2,2,0,"PILOT",
"PILOT-20004","PILOT 40 mg","Month 1","1 hour",89.85,"Y",2,2,1,"PILOT",
"PILOT-20004","PILOT 40 mg","Month 1","2 hours",109.18,"Y",2,2,2,"PILOT",
"PILOT-20005","PILOT 40 mg","Day 1","Predose",0,"Y",2,1,0,"PILOT",
"PILOT-20005","PILOT 40 mg","Day 1","1 hour",113.97,"Y",2,1,1,"PILOT",
"PILOT-20005","PILOT 40 mg","Day 1","2 hours",87.41,"Y",2,1,2,"PILOT",
"PILOT-20005","PILOT 40 mg","Month 1","Predose",95.36,"Y",2,2,0,"PILOT",
"PILOT-20005","PILOT 40 mg","Month 1","1 hour",104.97,"Y",2,2,1,"PILOT",
"PILOT-20005","PILOT 40 mg","Month 1","2 hours",114.63,"Y",2,2,2,"PILOT",
)


# Simulate twice the original subject count while preserving treatment, day,
# nominal sampling time and concentration-time profile relationships.
maps <- make_id_maps(list(ADSL=adsl, ADPC=adpc))
adsl <- double_subjects(adsl, maps)
adpc <- double_subjects(adpc, maps)
positive <- adpc$aval > 0 & !is.na(adpc$aval)
adpc$aval[positive] <- round(adpc$aval[positive] * exp(rnorm(sum(positive),0,0.08)),3)
adpc$aval[!positive & !is.na(adpc$aval)] <- 0
stopifnot(n_distinct(adsl$usubjid) == 2L * maps$original_n)
write_xlsx_list(list(ADSL=adsl, ADPC=adpc), file.path(OUT_DIR,paste0(STEM,"_simulated_data.xlsx")))

# Original analysis structure retained from the source document.
#==============================================================================
# Process input data
#==============================================================================
adsl01 <- adsl %>%
 filter(pkfl == "Y") %>%
 mutate(treatment = trt01an)
adpc01 <- adpc %>%
 filter(pkfl == "Y") %>%
 mutate(treatment = trt01an,
 logaval = if_else(aval > 0, log(aval), NA_real_))
#==============================================================================
# Obtain descriptive statistics
#==============================================================================
stats01 <- adpc01 %>%
 arrange(paramcd, avisitn, avisit, atptn, atpt) %>%
 group_by(paramcd, avisitn, avisit, atptn, atpt, treatment) %>%
 summarise(
 an = n(),
 amean = mean(aval, na.rm = TRUE),
 astd = sd(aval, na.rm = TRUE),
 amin = min(aval, na.rm = TRUE),
 amedian = median(aval, na.rm = TRUE),
 amax = max(aval, na.rm = TRUE),
 ln = n(),
 lmean = mean(logaval, na.rm = TRUE),
 lstd = sd(logaval, na.rm = TRUE),
lmin = if(all(is.na(logaval))) NA_real_ else min(logaval, na.rm = TRUE),
lmedian = if(all(is.na(logaval))) NA_real_ else median(logaval, na.rm = TRUE),
lmax = if(all(is.na(logaval))) NA_real_ else max(logaval, na.rm = TRUE),
.groups = "drop"
 ) %>%
 ungroup()
#==============================================================================
# Process the statistics as per presentation requirements
#==============================================================================
stats02 <- stats01 %>%
 mutate(
 n = if_else(!is.na(an), as.character(an), ""),
 mean = if_else(!is.na(amean), sprintf("%.2f", amean), ""),
 std = if_else(!is.na(astd), sprintf("%.3f", astd), ""),
 median = if_else(!is.na(amedian), sprintf("%.2f", amedian), ""),
 min = if_else(!is.na(amin), sprintf("%.1f", amin), ""),
 max = if_else(!is.na(amax), sprintf("%.1f", amax), ""),
 gmean = if_else(!is.na(lmean), sprintf("%.2f", exp(lmean)), ""),
 gstd = if_else(!is.na(lstd), sprintf("%.3f", exp(lstd)), ""),
 minmax = if_else(!is.na(amin) & !is.na(amax), str_c(min, ", ", max), ""),
 variance = if_else(!is.na(lstd), lstd^2, NA_real_),
 varexp = exp(variance),
 var_1 = varexp - 1,
 varsqrt = sqrt(var_1),
 gcv = if_else(!is.na(lstd), sprintf("%.1f", varsqrt * 100), "")
 )
#==============================================================================
# Keep only the required variables
#==============================================================================
stats03 <- stats02 %>%
 select(paramcd, avisitn, avisit, atptn, atpt, treatment, n, mean, std, gmean, gstd, gcv, median, minmax)
#==============================================================================
# Restructure the statistics such that they appear as 'rows' - using pivot_longer
#==============================================================================
stats04 <- stats03 %>%
 pivot_longer(cols = c(n, mean, std, gmean, gstd, gcv, median, minmax),
 names_to = "statistic",
 values_to = "value") %>%
 mutate(statistic = str_to_upper(statistic))
#==============================================================================
# Create some supporting variables as per sorting and presentation requirements
#==============================================================================
stats05 <- stats04 %>%
 mutate(
 intord = case_when(
 statistic == "N"      ~ 1,
 statistic == "MEAN"   ~ 2,
 statistic == "STD"    ~ 3,
 statistic == "GMEAN"  ~ 4,
 statistic == "GSTD"   ~ 5,
 statistic == "GCV"    ~ 6,
 statistic == "MEDIAN" ~ 7,
 statistic == "MINMAX" ~ 8,
 TRUE                  ~ NA_integer_
 )
 ) %>%
 arrange(paramcd, avisitn, avisit, atptn, atpt, intord)
#==============================================================================
# Restructure the data such that treatments appear as 'columns' - using pivot_wider
#==============================================================================
stats06 <- stats05 %>%
 pivot_wider(names_from = treatment, values_from = value, names_prefix = "trt")
#==============================================================================
# Create final dataset - keeping only required variables
#==============================================================================
final <- stats06 %>%
 select(paramcd, avisitn, avisit, atptn, atpt, intord, statistic, starts_with("trt"))
output <- final
#==============================================================================;

n1 <- n_distinct(adsl$usubjid[adsl$trt01an==1 & adsl$pkfl=="Y"])
n2 <- n_distinct(adsl$usubjid[adsl$trt01an==2 & adsl$pkfl=="Y"])
stat_labels <- c(N="n",MEAN="Mean",STD="SD",GMEAN="Geometric Mean",GSTD="Geometric SD",GCV="Geometric CV (%)",MEDIAN="Median",MINMAX="Min, Max")
display <- output %>%
  mutate(
    parameter_display=if_else(
      row_number()==1L | paramcd!=lag(paramcd),
      paramcd,
      ""
    ),
    visit_display=if_else(
      row_number()==1L | paramcd!=lag(paramcd) | avisit!=lag(avisit),
      avisit,
      ""
    )
  ) %>%
  transmute(
  Parameter=parameter_display,
  Visit=visit_display,
  `Nominal Time`=atpt,
  Statistic=unname(stat_labels[statistic]),
  !!paste0("PILOT 20 mg (N=",n1,")") := trt1,
  !!paste0("PILOT 40 mg (N=",n2,")") := trt2
)
write_clinical_rtf(display,
  "Table 1. Summary of Pharmacokinetic Concentrations by Nominal Time",
  file.path(OUT_DIR,paste0(STEM,".rtf")),
  "PK concentrations are summarized for the PK Analysis Set. Geometric statistics exclude non-positive concentrations. CV = coefficient of variation; SD = standard deviation."
)
cat("Created ",STEM,": simulated N = ",2L*maps$original_n,"\n",sep="")
