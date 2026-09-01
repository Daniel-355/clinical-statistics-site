# Standalone vital-signs analysis generated from DOCX code segment 1.

suppressPackageStartupMessages({library(dplyr);library(tidyr)})
set.seed(20260828)
arg <- grep("^--file=", commandArgs(FALSE), value=TRUE)
script_raw <- if(length(arg)) sub("^--file=", "", arg[1]) else "."
script_path <- normalizePath(gsub("~\\+~", " ", script_raw))
out_dir <- if(file.info(script_path)$isdir) script_path else dirname(script_path)
stem <- tools::file_path_sans_ext(basename(script_path))

b64decode <- function(s){a<-strsplit(gsub("\\s","",s),"")[[1]];tab<-c(LETTERS,letters,0:9,"+","/");v<-match(a,tab)-1L;v[is.na(v)]<-0L;bits<-as.integer(unlist(lapply(v,function(x)intToBits(x)[6:1])));n<-floor(length(bits)/8);as.raw(vapply(seq_len(n),function(i)sum(bits[((i-1)*8+1):(i*8)]*2^(7:0)),numeric(1)))}
unpack_csv <- function(s) read.csv(text=rawToChar(memDecompress(base64enc::base64decode(s),type="gzip")),stringsAsFactors=FALSE,check.names=FALSE,na.strings=c("NA",""))

xml_escape <- function(x){x[is.na(x)]<-"";x<-gsub("&","&amp;",x,fixed=TRUE);x<-gsub("<","&lt;",x,fixed=TRUE);x<-gsub(">","&gt;",x,fixed=TRUE);gsub('"',"&quot;",x,fixed=TRUE)}
excel_col <- function(n){z<-character(length(n));for(j in seq_along(n)){k<-n[j];s<-"";while(k>0){k<-k-1;s<-paste0(intToUtf8(65+k%%26),s);k<-k%/%26};z[j]<-s};z}
write_xlsx <- function(sheets,path){
 sheets<-lapply(sheets,function(d){names(d)<-toupper(names(d));d});sn<-substr(gsub("[^A-Za-z0-9_]","_",names(sheets)),1,31);td<-tempfile("xlsx_");dir.create(td);for(z in c("_rels","docProps","xl","xl/_rels","xl/worksheets"))dir.create(file.path(td,z),recursive=TRUE,showWarnings=FALSE)
 types<-paste0('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>',paste0('<Override PartName="/xl/worksheets/sheet',seq_along(sheets),'.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>',collapse=""),'<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');writeLines(types,file.path(td,"[Content_Types].xml"))
 writeLines('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>',file.path(td,"_rels/.rels"))
 sh<-paste0('<sheet name="',sn,'" sheetId="',seq_along(sn),'" r:id="rId',seq_along(sn),'"/>',collapse="");writeLines(paste0('<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>',sh,'</sheets></workbook>'),file.path(td,"xl/workbook.xml"));rels<-paste0('<Relationship Id="rId',seq_along(sn),'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet',seq_along(sn),'.xml"/>',collapse="");writeLines(paste0('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',rels,'<Relationship Id="rId',length(sn)+1,'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>'),file.path(td,"xl/_rels/workbook.xml.rels"))
 writeLines('<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="10"/><name val="Arial"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="10"/><name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="2"><xf xfId="0" fontId="0" fillId="0"/><xf xfId="0" fontId="1" fillId="1" applyFont="1" applyFill="1"/></cellXfs></styleSheet>',file.path(td,"xl/styles.xml"))
 for(k in seq_along(sheets)){d<-sheets[[k]];a<-rbind(names(d),as.matrix(d));rows<-vapply(seq_len(nrow(a)),function(i){cc<-vapply(seq_len(ncol(a)),function(j){v<-as.character(a[i,j]);if(i>1&&is.numeric(d[[j]])&&!is.na(d[i-1,j]))sprintf('<c r="%s%d"><v>%s</v></c>',excel_col(j),i,d[i-1,j])else sprintf('<c r="%s%d" t="inlineStr" s="%d"><is><t>%s</t></is></c>',excel_col(j),i,ifelse(i==1,1,0),xml_escape(v))},character(1));sprintf('<row r="%d">%s</row>',i,paste0(cc,collapse=""))},character(1));writeLines(paste0('<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetViews><sheetView showGridLines="0" workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetData>',paste0(rows,collapse=""),'</sheetData></worksheet>'),file.path(td,"xl/worksheets",paste0("sheet",k,".xml")))}
 writeLines('<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>PILOT</dc:creator></cp:coreProperties>',file.path(td,"docProps/core.xml"));writeLines('<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>R</Application></Properties>',file.path(td,"docProps/app.xml"));old<-setwd(td);on.exit(setwd(old),add=TRUE);if(file.exists(path))unlink(path);utils::zip(path,list.files(".",recursive=TRUE,all.files=TRUE,no..=TRUE),flags="-q")
}

rtf_escape<-function(x){x<-gsub("\\\\","\\\\\\\\",as.character(x));gsub("([{}])","\\\\\\1",x)}
write_rtf<-function(tab,title,path,foot){tab<-as.data.frame(tab,check.names=FALSE);tab[]<-lapply(tab,as.character);con<-file(path,"wb");on.exit(close(con));w<-function(...)writeLines(paste0(...),con,useBytes=TRUE);w("{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\landscape\\paperw15840\\paperh12240\\margl600\\margr600\\margt500\\margb600");w("\\pard\\f0\\fs15 PILOT\\tab Confidential\\tab Page {\\field{\\*\\fldinst PAGE}} of {\\field{\\*\\fldinst NUMPAGES}}\\par");w("\\pard\\qc\\b\\fs19 ",rtf_escape(title),"\\b0\\par\\pard\\qc\\fs15 Safety Population\\par\\par");n<-ncol(tab);first<-if(n>=6)2700 else 3600;ends<-cumsum(c(first,rep((14640-first)/(n-1),n-1)));for(i in 0:nrow(tab)){vals<-if(i==0)names(tab)else unlist(tab[i,],use.names=FALSE);bd<-if(i==0)"\\clbrdrt\\brdrs\\brdrw15\\clbrdrb\\brdrs\\brdrw15"else if(i==nrow(tab))"\\clbrdrb\\brdrs\\brdrw15"else"";w("\\trowd\\trgaph55",if(i==0)"\\trhdr"else"",paste0(bd,"\\cellx",round(ends),collapse=""));for(j in seq_len(n))w("\\intbl",if(j<=2)"\\ql"else"\\qc",if(i==0)"\\b"else"\\b0","\\f0\\fs13 ",rtf_escape(vals[j]),"\\cell");w("\\row")};w("\\pard\\ql\\fs12\\par ",rtf_escape(foot),"\\par Source: simulated data generated by this standalone program. Generated: ",format(Sys.Date(),"%d%b%Y"),"\\par}")}

simulate_profiles<-function(adsl,advs,original_n){
 trts<-sort(unique(adsl$TRT01AN));new_adsl<-list();new_advs<-list();next_id<-1L
 for(trt in trts){n_src<-sum(adsl$TRT01AN==trt);n_new<-2L*n_src;src_ids<-unique(advs$USUBJID[advs$TRT01AN==trt]);if(length(src_ids)==0)src_ids<-unique(advs$USUBJID);picked<-sample(src_ids,n_new,replace=TRUE)
  for(k in seq_len(n_new)){prof<-advs[advs$USUBJID==picked[k]&advs$TRT01AN==trt,,drop=FALSE];if(!nrow(prof))prof<-advs[advs$USUBJID==picked[k],,drop=FALSE];id<-sprintf("PILOT-%04d",next_id);next_id<-next_id+1L;prof$USUBJID<-id;if("AVAL"%in%names(prof)){sdv<-ave(advs$AVAL,advs$PARAMCD,FUN=function(x)max(sd(x,na.rm=TRUE),1));prof$AVAL<-round(prof$AVAL+rnorm(nrow(prof),0,0.04*sdv[match(prof$PARAMCD,advs$PARAMCD)]),1)};new_advs[[length(new_advs)+1]]<-prof;new_adsl[[length(new_adsl)+1]]<-data.frame(USUBJID=id,TRT01AN=trt,SAFFL="Y",stringsAsFactors=FALSE)}
 }
  a<-bind_rows(new_adsl);v<-bind_rows(new_advs);if("BASE"%in%names(v)){base_tbl<-v%>%group_by(USUBJID,PARAMCD)%>%arrange(AVISITN,.by_group=TRUE)%>%summarise(BASE=round(first(AVAL),1),.groups="drop");v<-v%>%select(-any_of(c("BASE","CHG")))%>%left_join(base_tbl,by=c("USUBJID","PARAMCD"))%>%group_by(USUBJID,PARAMCD)%>%mutate(CHG=ifelse(AVISITN==min(AVISITN,na.rm=TRUE),NA_real_,round(AVAL-BASE,1)))%>%ungroup()};list(ADSL=a,ADVS=v)
}

EMBED_ADSL<-"H4sIAJeqkWoC/42SzQrCMBCE7z5LhMwmqXqsSKGlqNj04Pu/iDYRDbOhNsePmf2ZzTzN56G/mPiIFu3VTG3Xjbt7P97iHvb9YMQ8P8AtwBt8ga8qFHBrlgSkAMcMfhbhGvUuwoqGFBAGXBS+UEjusjLHAmC1hUBlF1I0hSJwyAmEwoLqtipTaAsBx8AzCBvyUG3BtyVQRnjgwSS3xZ/bql94SjVegA6VGs8CAAA="
EMBED_ADVS<-"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"
ORIGINAL_N<-41L

make_output <- function(adsl, advs) {
  safety <- adsl %>% filter(SAFFL == "Y")
  vital <- advs %>% filter(SAFFL == "Y")

  treatment_totals <- safety %>% count(TRT01AN, name = "N")
  treatment_labels <- setNames(
    paste0(
      ifelse(treatment_totals$TRT01AN == 1, "Placebo", "Active"),
      " (N=", treatment_totals$N, ")"
    ),
    treatment_totals$TRT01AN
  )
  treatment_labels <- c(
    treatment_labels,
    Total = paste0("Total (N=", n_distinct(safety$USUBJID), ")")
  )

  vital_with_total <- bind_rows(
    vital %>% mutate(TREATMENT = as.character(TRT01AN)),
    vital %>% mutate(TREATMENT = "Total")
  )

  analysis_long <- bind_rows(
    vital_with_total %>%
      transmute(
        PARAMN, PARAM, AVISITN, AVISIT, TREATMENT,
        TYPE = "Observed Value",
        RESULT = AVAL
      ),
    vital_with_total %>%
      filter(!is.na(CHG)) %>%
      transmute(
        PARAMN, PARAM, AVISITN, AVISIT, TREATMENT,
        TYPE = "Change from Baseline",
        RESULT = CHG
      )
  )

  output <- analysis_long %>%
    group_by(PARAMN, PARAM, AVISITN, AVISIT, TYPE, TREATMENT) %>%
    summarise(
      N = sum(!is.na(RESULT)),
      MEAN = mean(RESULT, na.rm = TRUE),
      SD = sd(RESULT, na.rm = TRUE),
      MIN = min(RESULT, na.rm = TRUE),
      MEDIAN = median(RESULT, na.rm = TRUE),
      MAX = max(RESULT, na.rm = TRUE),
      .groups = "drop"
    ) %>%
    mutate(
      TREATMENT_ORDER = match(TREATMENT, c("1", "2", "Total")),
      TREATMENT_LABEL = unname(treatment_labels[TREATMENT])
    ) %>%
    arrange(PARAMN, AVISITN, TYPE, TREATMENT_ORDER) %>%
    mutate(
      PARAMETER_DISPLAY = if_else(
        row_number() == 1L | PARAM != lag(PARAM), PARAM, ""
      ),
      VISIT_DISPLAY = if_else(
        row_number() == 1L | PARAM != lag(PARAM) | AVISIT != lag(AVISIT),
        AVISIT,
        ""
      ),
      TYPE_DISPLAY = if_else(
        row_number() == 1L | PARAM != lag(PARAM) |
          AVISIT != lag(AVISIT) | TYPE != lag(TYPE),
        TYPE,
        ""
      ),
      SD_DISPLAY = if_else(N > 1L & is.finite(SD), sprintf("%.2f", SD), "-"),
      MEAN_DISPLAY = if_else(N > 0L & is.finite(MEAN), sprintf("%.2f", MEAN), ""),
      MIN_DISPLAY = if_else(N > 0L & is.finite(MIN), sprintf("%.1f", MIN), ""),
      MEDIAN_DISPLAY = if_else(N > 0L & is.finite(MEDIAN), sprintf("%.2f", MEDIAN), ""),
      MAX_DISPLAY = if_else(N > 0L & is.finite(MAX), sprintf("%.1f", MAX), "")
    ) %>%
    transmute(
      Parameter = PARAMETER_DISPLAY,
      Visit = VISIT_DISPLAY,
      Type = TYPE_DISPLAY,
      Treatment = TREATMENT_LABEL,
      n = N,
      Mean = MEAN_DISPLAY,
      SD = SD_DISPLAY,
      Min = MIN_DISPLAY,
      Median = MEDIAN_DISPLAY,
      Max = MAX_DISPLAY
    )

  output
}
TITLE<-"Table 1. Summary of Vital Signs and Change from Baseline"
FOOT<-"Mean, standard deviation, median, minimum and maximum are calculated from non-missing values."

adsl<-unpack_csv(EMBED_ADSL);advs<-unpack_csv(EMBED_ADVS)
original_n<-ORIGINAL_N
sim<-simulate_profiles(adsl,advs,original_n)
names(sim$ADSL)<-toupper(names(sim$ADSL));names(sim$ADVS)<-toupper(names(sim$ADVS))
stopifnot(n_distinct(sim$ADSL$USUBJID)==2L*original_n,all(names(sim$ADSL)==toupper(names(sim$ADSL))),all(names(sim$ADVS)==toupper(names(sim$ADVS))))
write_xlsx(sim,file.path(out_dir,paste0(stem,"_simulated_data.xlsx")))
output<-make_output(sim$ADSL,sim$ADVS)
write_rtf(output,TITLE,file.path(out_dir,paste0(stem,".rtf")),FOOT)
message(sprintf("Created %s: original N=%d; simulated N=%d",stem,original_n,n_distinct(sim$ADSL$USUBJID)))
