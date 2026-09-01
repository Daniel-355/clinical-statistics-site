#!/usr/bin/env Rscript
# PILOT L102: Summary of Electrocardiogram Parameters and Change from Baseline
# Standalone workflow: source-logical data -> doubled simulation -> XLSX -> analysis -> RTF.
options(stringsAsFactors=FALSE);set.seed(20260827);suppressPackageStartupMessages(library(tidyverse))
TYPE<-"L102";STEM<-"02_ecg_numeric_summary";TITLE<-"Summary of Electrocardiogram Parameters and Change from Baseline";N_ORIGINAL<-10L
script_dir<-function(){a<-commandArgs(FALSE);f<-sub("^--file=","",a[grepl("^--file=",a)]);f<-gsub("~\\+~"," ",f);if(length(f))dirname(normalizePath(f[1]))else getwd()};OUT_DIR<-script_dir()
xe<-function(x){x[is.na(x)]<-"";x<-gsub("&","&amp;",x,fixed=TRUE);x<-gsub("<","&lt;",x,fixed=TRUE);x<-gsub(">","&gt;",x,fixed=TRUE);gsub('"',"&quot;",x,fixed=TRUE)}
xc<-function(n){z<-character(length(n));for(j in seq_along(n)){k<-n[j];s<-"";while(k>0){k<-k-1;s<-paste0(intToUtf8(65+k%%26),s);k<-k%/%26};z[j]<-s};z}
write_xlsx<-function(d,p){names(d)<-toupper(names(d));td<-tempfile("x");dir.create(td);ds<-c("_rels","docProps","xl","xl/_rels","xl/worksheets");invisible(vapply(file.path(td,ds),dir.create,logical(1),recursive=TRUE));writeLines('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>',file.path(td,"[Content_Types].xml"));writeLines('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>',file.path(td,"_rels/.rels"));writeLines('<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="SIMULATED_DATA" sheetId="1" r:id="rId1"/></sheets></workbook>',file.path(td,"xl/workbook.xml"));writeLines('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>',file.path(td,"xl/_rels/workbook.xml.rels"));writeLines('<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><sz val="10"/><name val="Arial"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf xfId="0"/></cellXfs></styleSheet>',file.path(td,"xl/styles.xml"));a<-rbind(names(d),as.matrix(d));rows<-vapply(seq_len(nrow(a)),function(i){cc<-vapply(seq_len(ncol(a)),function(j)sprintf('<c r="%s%d" t="inlineStr"><is><t>%s</t></is></c>',xc(j),i,xe(as.character(a[i,j]))),character(1));sprintf('<row r="%d">%s</row>',i,paste0(cc,collapse=""))},character(1));writeLines(paste0('<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>',paste0(rows,collapse=""),'</sheetData></worksheet>'),file.path(td,"xl/worksheets/sheet1.xml"));writeLines('<?xml version="1.0" encoding="UTF-8"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:creator>PILOT</dc:creator></cp:coreProperties>',file.path(td,"docProps/core.xml"));writeLines('<?xml version="1.0" encoding="UTF-8"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>R</Application></Properties>',file.path(td,"docProps/app.xml"));old<-setwd(td);on.exit(setwd(old),add=TRUE);if(file.exists(p))unlink(p);utils::zip(p,list.files(".",recursive=TRUE,all.files=TRUE,no..=TRUE),flags="-q")}
re<-function(x){x<-gsub("\\","\\\\",x,fixed=TRUE);x<-gsub("{","\\{",x,fixed=TRUE);x<-gsub("}","\\}",x,fixed=TRUE);gsub("\n","\\line ",x,fixed=TRUE)}
write_rtf<-function(tab,p){con<-file(p,"wb");on.exit(close(con));w<-function(...)writeLines(paste0(...),con,useBytes=TRUE);w("{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}\\landscape\\paperw15840\\paperh12240\\margl600\\margr600\\margt450\\margb600");w("\\pard\\par");w("\\pard\\f0\\fs15 PILOT\\tab Confidential\\tab Page {\\field{\\*\\fldinst PAGE}} of {\\field{\\*\\fldinst NUMPAGES}}\\par");w("\\pard\\qc\\b\\fs20 ",re(TITLE),"\\b0\\par");w("\\pard\\qc\\fs16 Safety Analysis Set\\par\\par");nc<-ncol(tab);widths<-if(nc==6)c(1400,3000,rep(2400,4))else{first<-if(nc<=4)6200 else 3600;c(first,rep(floor((14100-first)/(nc-1)),nc-1))};ends<-cumsum(widths);for(i in seq_len(nrow(tab))){bd<-if(i==1)"\\clbrdrt\\brdrs\\brdrw15\\clbrdrb\\brdrs\\brdrw15"else if(i==nrow(tab))"\\clbrdrb\\brdrs\\brdrw15"else"";w("\\trowd\\trgaph60",if(i==1)"\\trhdr"else"",paste0(bd,"\\cellx",ends,collapse=""));for(j in seq_len(nc))w("\\intbl",if(j<=2)"\\ql"else"\\qc",if(i==1)"\\b"else"\\b0","\\f0\\fs14 ",re(as.character(tab[i,j])),"\\cell");w("\\row")};w("\\pard\\ql\\fs13\\par Percentages are based on subjects with available data within treatment and visit. N = subjects in the Safety Analysis Set; n = subjects in category.\\par");w("\\pard\\ql\\fs13 Source: Simulated data generated by this standalone program. Generated: ",format(Sys.Date(),"%d%b%Y"),"\\par}")}

# Write the numeric ECG table with separate statistics under Result and Change.
write_rtf_numeric <- function(tab, path) {
  con <- file(path, "wb")
  on.exit(close(con))
  write_text <- function(...) writeLines(paste0(...), con, useBytes = TRUE)

  write_text("{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Arial;}}",
             "\\landscape\\paperw15840\\paperh12240",
             "\\margl600\\margr600\\margt450\\margb600")
  write_text("\\pard\\par")
  write_text("\\pard\\f0\\fs15 PILOT\\tab Confidential\\tab Page ",
             "{\\field{\\*\\fldinst PAGE}} of ",
             "{\\field{\\*\\fldinst NUMPAGES}}\\par")
  write_text("\\pard\\qc\\b\\fs20 ", re(TITLE), "\\b0\\par")
  write_text("\\pard\\qc\\fs16 Safety Analysis Set\\par\\par")

  # Parameter, Visit, Treatment, followed by two groups of six statistics.
  widths <- c(2100, 1150, 1050, rep(815, 12))
  ends <- cumsum(widths)

  for (i in seq_len(nrow(tab))) {
    row_border <- if (i == 1L) {
      "\\clbrdrt\\brdrs\\brdrw15"
    } else if (i == 2L || i == nrow(tab)) {
      "\\clbrdrb\\brdrs\\brdrw15"
    } else {
      ""
    }

    cell_definitions <- character(ncol(tab))
    for (j in seq_len(ncol(tab))) {
      merge_control <- ""
      group_underline <- ""
      if (i == 1L && j %in% 4:9) {
        merge_control <- if (j == 4L) "\\clmgf" else "\\clmrg"
        group_underline <- "\\clbrdrb\\brdrs\\brdrw15"
      }
      if (i == 1L && j %in% 10:15) {
        merge_control <- if (j == 10L) "\\clmgf" else "\\clmrg"
        group_underline <- "\\clbrdrb\\brdrs\\brdrw15"
      }
      cell_definitions[j] <- paste0(
        merge_control,
        row_border,
        group_underline,
        "\\cellx",
        ends[j]
      )
    }

    write_text(
      "\\trowd\\trgaph45",
      if (i <= 2L) "\\trhdr" else "",
      paste0(cell_definitions, collapse = "")
    )

    for (j in seq_len(ncol(tab))) {
      alignment <- if (j <= 3L && i > 2L) "\\ql" else "\\qc"
      font_weight <- if (i <= 2L) "\\b" else "\\b0"
      write_text(
        "\\intbl", alignment, font_weight, "\\f0\\fs12 ",
        re(as.character(tab[i, j])), "\\cell"
      )
    }
    write_text("\\row")
  }

  write_text(
    "\\pard\\ql\\fs13\\par Descriptive statistics are based on nonmissing ",
    "observations within parameter, visit, and treatment group. SD = standard deviation.\\par"
  )
  write_text(
    "\\pard\\ql\\fs13 Source: Simulated data generated by this standalone program. Generated: ",
    format(Sys.Date(), "%d%b%Y"), "\\par}"
  )
}
visits<-tibble(AVISITN=c(6L,8L,11L,12L),AVISIT=c("Baseline","Week 4","Week 12","Week 16"))
subjects<-tibble(USUBJID=sprintf("PILOT-O%04d",seq_len(N_ORIGINAL)),TRT01AN=rep(1:2,length.out=N_ORIGINAL),SAFFL="Y")
if (TYPE == "L102") {
  params <- tibble(
    PARAMN = 1:4,
    PARAMCD = c("PR", "QRS", "QT", "QTCF"),
    PARAM = c(
      "PR Interval (msec)", "QRS Duration (msec)",
      "QT Interval (msec)", "QTcF Interval (msec)"
    ),
    MU = c(170, 92, 405, 420)
  )

  # Generate one baseline value for each subject and ECG parameter.
  # That same BASE value is then carried to every visit for the group.
  subject_parameter_base <- crossing(subjects, params) %>%
    mutate(BASE = round(rnorm(n(), mean = MU, sd = 12)))

  d <- crossing(subjects, visits, params) %>%
    select(-MU) %>%
    left_join(
      subject_parameter_base %>% select(USUBJID, PARAMCD, BASE),
      by = c("USUBJID", "PARAMCD")
    ) %>%
    mutate(
      AVAL = if_else(
        AVISITN == 6L,
        BASE,
        round(BASE + rnorm(n(), mean = 2, sd = 8))
      ),
      CHG = AVAL - BASE
    )

  # Validation: one BASE per subject/parameter and zero change at Baseline.
  baseline_check <- d %>%
    group_by(USUBJID, PARAMCD) %>%
    summarise(N_BASE_VALUES = n_distinct(BASE), .groups = "drop")

  stopifnot(
    all(baseline_check$N_BASE_VALUES == 1L),
    all(d$CHG[d$AVISITN == 6L] == 0)
  )
} else {d<-crossing(subjects,visits);basecat<-sample(c("NORMAL","ABNORMAL - NOT CLINICALLY SIGNIFICANT","ABNORMAL - CLINICALLY SIGNIFICANT"),N_ORIGINAL,TRUE,c(.42,.50,.08));si<-match(d$USUBJID,subjects$USUBJID);d$PARAMCD<-"INTP";d$PARAM<-"Interpretation";d$AVALC<-basecat[si];post<-d$AVISITN>6&runif(nrow(d))<.18;d$AVALC[post]<-sample(c("NORMAL","ABNORMAL - NOT CLINICALLY SIGNIFICANT","ABNORMAL - CLINICALLY SIGNIFICANT"),sum(post),TRUE,c(.42,.50,.08));d$AVAL<-round(rnorm(nrow(d),420,25));d$BASE<-ave(d$AVAL,d$USUBJID,FUN=function(x)x[1]);d$CHG<-d$AVAL-d$BASE}
pick<-c(seq_len(N_ORIGINAL),sample(seq_len(N_ORIGINAL),N_ORIGINAL,TRUE));parts<-vector("list",length(pick));for(i in seq_along(pick)){z<-d[d$USUBJID==subjects$USUBJID[pick[i]],];z$USUBJID<-sprintf("PILOT-S%04d",i);parts[[i]]<-z};simulated_data<-bind_rows(parts)
simulated_baseline_check <- simulated_data %>%
  group_by(USUBJID, PARAMCD) %>%
  summarise(N_BASE_VALUES = n_distinct(BASE), .groups = "drop")

stopifnot(
  n_distinct(simulated_data$USUBJID) == 2L * N_ORIGINAL,
  all(names(simulated_data) == toupper(names(simulated_data))),
  all(simulated_baseline_check$N_BASE_VALUES == 1L),
  all(simulated_data$CHG[simulated_data$AVISIT == "Baseline"] == 0)
)
write_xlsx(simulated_data,file.path(OUT_DIR,paste0(STEM,"_simulated_data.xlsx")))
fmt<-function(n,d)if(n==0||d==0)"0"else sprintf("%d (%.1f%%)",n,100*n/d);den<-c(sum(subjects$TRT01AN[pick]==1),sum(subjects$TRT01AN[pick]==2));groups<-1:3;gh<-c(sprintf("Placebo\n(N=%d)",den[1]),sprintf("Active\n(N=%d)",den[2]),sprintf("Total\n(N=%d)",sum(den)))
if(TYPE=="L101"){lev<-c("n","Normal","Abnormal - NCS","Abnormal - CS");rows<-list();for(v in visits$AVISIT){z<-simulated_data%>%filter(AVISIT==v);for(k in lev){vals<-vapply(groups,function(g){q<-if(g==3)z else z%>%filter(TRT01AN==g);dd<-n_distinct(q$USUBJID);if(k=="n")as.character(dd)else{map<-c(Normal="NORMAL","Abnormal - NCS"="ABNORMAL - NOT CLINICALLY SIGNIFICANT","Abnormal - CS"="ABNORMAL - CLINICALLY SIGNIFICANT");fmt(n_distinct(q$USUBJID[q$AVALC==map[k]]),dd)}},character(1));rows[[length(rows)+1]]<-c(if(k=="n")v else paste0("   ",k),vals)}};tab<-rbind(c("Visit / Interpretation",gh),do.call(rbind,rows))
}else if(TYPE=="L102"){
  # Return each descriptive statistic as a separate display column.
  descriptive_statistics <- function(x) {
    x <- x[!is.na(x)]
    n <- length(x)
    c(
      as.character(n),
      if (n) sprintf("%.1f", mean(x)) else "-",
      if (n > 1L) sprintf("%.2f", sd(x)) else "-",
      if (n) sprintf("%.0f", min(x)) else "-",
      if (n) sprintf("%.1f", median(x)) else "-",
      if (n) sprintf("%.0f", max(x)) else "-"
    )
  }

  rows <- list()
  previous_parameter <- ""
  previous_visit <- ""

  for (parameter in unique(simulated_data$PARAM)) {
    for (visit in visits$AVISIT) {
      for (group in groups) {
        subset_data <- simulated_data %>%
          filter(PARAM == parameter, AVISIT == visit)

        if (group < 3L) {
          subset_data <- subset_data %>% filter(TRT01AN == group)
        }

        display_parameter <- if (identical(parameter, previous_parameter)) "" else parameter
        display_visit <- if (
          identical(parameter, previous_parameter) && identical(visit, previous_visit)
        ) "" else visit

        rows[[length(rows) + 1L]] <- c(
          display_parameter,
          display_visit,
          c("Placebo", "Active", "Total")[group],
          descriptive_statistics(subset_data$AVAL),
          descriptive_statistics(subset_data$CHG)
        )

        previous_parameter <- parameter
        previous_visit <- visit
      }
    }
  }

  header_group <- c(
    "Parameter", "Visit", "Treatment",
    "Result", rep("", 5),
    "Change", rep("", 5)
  )
  header_statistics <- c(
    "", "", "",
    "n", "Mean", "SD", "Min", "Median", "Max",
    "n", "Mean", "SD", "Min", "Median", "Max"
  )
  tab <- rbind(header_group, header_statistics, do.call(rbind, rows))
}else if(TYPE%in%c("L103","L104","L105")){if(TYPE=="L103"){labs<-c("n",">450",">480",">500");fun<-list(function(x)TRUE,function(x)x>450,function(x)x>480,function(x)x>500)};if(TYPE=="L104"){labs<-c("n","<=450",">450 to <=480",">480 to <=500",">500");fun<-list(function(x)TRUE,function(x)x<=450,function(x)x>450&x<=480,function(x)x>480&x<=500,function(x)x>500)};if(TYPE=="L105"){labs<-c("n","<0",">=0",">30",">60");fun<-list(function(x)TRUE,function(x)x<0,function(x)x>=0,function(x)x>30,function(x)x>60)};rows<-list();for(v in visits$AVISIT){z0<-simulated_data%>%filter(AVISIT==v);xn<-if(TYPE=="L105")"CHG"else"AVAL";for(k in seq_along(labs)){vals<-vapply(groups,function(g){z<-if(g==3)z0 else z0%>%filter(TRT01AN==g);x<-z[[xn]];if(k==1)as.character(length(x))else fmt(sum(fun[[k]](x)),length(x))},character(1));rows[[length(rows)+1]]<-c(if(k==1)v else paste0("   ",labs[k]),vals)}};tab<-rbind(c("Visit / Category",gh),do.call(rbind,rows))
}else{miss<-TYPE!="L106";lev<-c("Normal","Abnormal - NCS","Abnormal - CS",if(miss)"Missing");code<-c("NORMAL","ABNORMAL - NOT CLINICALLY SIGNIFICANT","ABNORMAL - CLINICALLY SIGNIFICANT",NA);b<-simulated_data%>%filter(AVISIT=="Baseline")%>%select(USUBJID,BASEC=AVALC);z<-simulated_data%>%filter(AVISITN>6)%>%left_join(b,by="USUBJID");if(miss)z$AVALC[runif(nrow(z))<.03]<-NA;rows<-list();trts<-if(TYPE=="L106B")1:2 else 0;for(g in trts){q<-if(g==0)z else z%>%filter(TRT01AN==g);if(TYPE=="L106B")rows[[length(rows)+1]]<-c(paste0(c("Placebo","Active")[g]," treatment"),rep("",length(lev)+1));for(v in unique(q$AVISIT))for(r in seq_along(lev)){vals<-vapply(seq_along(lev),function(k)fmt(sum((if(is.na(code[r]))is.na(q$AVALC)else q$AVALC==code[r])&(if(is.na(code[k]))is.na(q$BASEC)else q$BASEC==code[k])&q$AVISIT==v,na.rm=TRUE),sum(q$AVISIT==v)),character(1));rows[[length(rows)+1]]<-c(if(r==1)v else "",lev[r],vals)}};tab<-rbind(c("Visit","Post-baseline Result",paste0("Baseline: ",lev)),do.call(rbind,rows))}
if (TYPE == "L102") {
  write_rtf_numeric(tab, file.path(OUT_DIR, paste0(STEM, ".rtf")))
} else {
  write_rtf(tab, file.path(OUT_DIR, paste0(STEM, ".rtf")))
}
cat("Created",STEM,"with",n_distinct(simulated_data$USUBJID),"simulated subjects.\n")
