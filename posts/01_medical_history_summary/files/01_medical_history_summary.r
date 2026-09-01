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
~usubjid,~trt01an,~saffl,
"PILOT-1001",54,"Y",
"PILOT-1002",0,"Y",
"PILOT-1003",54,"Y",
"PILOT-1004",81,"Y",
"PILOT-1005",81,"Y",
"PILOT-1006",54,"Y",
"PILOT-1007",0,"Y",
"PILOT-1008",81,"Y",
"PILOT-1009",54,"Y",
"PILOT-1010",0,"Y",
"PILOT-1011",81,"Y",
"PILOT-1012",0,"Y",
"PILOT-1013",81,"Y",
"PILOT-1014",0,"Y",
"PILOT-1015",54,"Y",
)

admh<-tribble(
~usubjid,~mhdecod,~mhbodsys,~trt01an,~saffl,
"PILOT-1001","WRIST FRACTURE","INJURY, POISONING AND PROCEDURAL COMPLICATIONS",54,"Y",
"PILOT-1001","RECTAL HAEMORRHAGE","GASTROINTESTINAL DISORDERS",54,"Y",
"PILOT-1001","HAEMORRHOIDS","GASTROINTESTINAL DISORDERS",54,"Y",
"PILOT-1001","CYSTITIS","INFECTIONS AND INFESTATIONS",54,"Y",
"PILOT-1001","TUBERCULOSIS","INFECTIONS AND INFESTATIONS",54,"Y",
"PILOT-1001","MACULAR DEGENERATION","EYE DISORDERS",54,"Y",
"PILOT-1001","DEAFNESS","EAR AND LABYRINTH DISORDERS",54,"Y",
"PILOT-1001","HIATUS HERNIA","GASTROINTESTINAL DISORDERS",54,"Y",
"PILOT-1001","EATING DISORDER","PSYCHIATRIC DISORDERS",54,"Y",
"PILOT-1001","FOOT FRACTURE","INJURY, POISONING AND PROCEDURAL COMPLICATIONS",54,"Y",
"PILOT-1001","OEDEMA PERIPHERAL","GENERAL DISORDERS AND ADMINISTRATION SITE CONDITIONS",54,"Y",
"PILOT-1001","ACROCHORDON","NEOPLASMS BENIGN, MALIGNANT AND UNSPECIFIED (INCL CYSTS AND POLYPS)",54,"Y",
"PILOT-1001","COUGH","RESPIRATORY, THORACIC AND MEDIASTINAL DISORDERS",54,"Y",
"PILOT-1001","BASAL CELL CARCINOMA","NEOPLASMS BENIGN, MALIGNANT AND UNSPECIFIED (INCL CYSTS AND POLYPS)",54,"Y",
"PILOT-1001","TINNITUS","EAR AND LABYRINTH DISORDERS",54,"Y",
"PILOT-1001","OSTEOPOROSIS","MUSCULOSKELETAL AND CONNECTIVE TISSUE DISORDERS",54,"Y",
"PILOT-1001","LOCALISED INFECTION","INFECTIONS AND INFESTATIONS",54,"Y",
"PILOT-1001","LOWER LIMB FRACTURE","INJURY, POISONING AND PROCEDURAL COMPLICATIONS",54,"Y",
"PILOT-1001","ARTHRITIS","MUSCULOSKELETAL AND CONNECTIVE TISSUE DISORDERS",54,"Y",
"PILOT-1001","UPPER LIMB FRACTURE","INJURY, POISONING AND PROCEDURAL COMPLICATIONS",54,"Y",
"PILOT-1001","CORRECTIVE LENS USER","SOCIAL CIRCUMSTANCES",54,"Y",
"PILOT-1001","RIB FRACTURE","INJURY, POISONING AND PROCEDURAL COMPLICATIONS",54,"Y",
"PILOT-1001","HYPOTHYROIDISM","ENDOCRINE DISORDERS",54,"Y",
"PILOT-1001","INCONTINENCE","RENAL AND URINARY DISORDERS",54,"Y",
"PILOT-1001","BACK PAIN","MUSCULOSKELETAL AND CONNECTIVE TISSUE DISORDERS",54,"Y",
"PILOT-1001","VARICOSE VEIN","VASCULAR DISORDERS",54,"Y",
"PILOT-1001","HYPERTENSION","VASCULAR DISORDERS",54,"Y",
"PILOT-1001","SINUSITIS","INFECTIONS AND INFESTATIONS",54,"Y",
"PILOT-1001","APPENDICECTOMY","SURGICAL AND MEDICAL PROCEDURES",54,"Y",
"PILOT-1002","DEAFNESS UNILATERAL","EAR AND LABYRINTH DISORDERS",0,"Y",
"PILOT-1002","PROSTATECTOMY","SURGICAL AND MEDICAL PROCEDURES",0,"Y",
"PILOT-1002","PROSTATE CANCER","NEOPLASMS BENIGN, MALIGNANT AND UNSPECIFIED (INCL CYSTS AND POLYPS)",0,"Y",
"PILOT-1002","HAEMORRHOIDS","GASTROINTESTINAL DISORDERS",0,"Y",
"PILOT-1002","FEMUR FRACTURE","INJURY, POISONING AND PROCEDURAL COMPLICATIONS",0,"Y",
"PILOT-1002","OTITIS EXTERNA","INFECTIONS AND INFESTATIONS",0,"Y",
"PILOT-1002","SKIN LACERATION","INJURY, POISONING AND PROCEDURAL COMPLICATIONS",0,"Y",
"PILOT-1002","DENTURE WEARER","SOCIAL CIRCUMSTANCES",0,"Y",
"PILOT-1002","PNEUMONIA","INFECTIONS AND INFESTATIONS",0,"Y",
"PILOT-1002","PNEUMONIA","INFECTIONS AND INFESTATIONS",0,"Y",
"PILOT-1002","PARAESTHESIA","NERVOUS SYSTEM DISORDERS",0,"Y",
"PILOT-1002","DRY SKIN","SKIN AND SUBCUTANEOUS TISSUE DISORDERS",0,"Y",
"PILOT-1002","TONSILLECTOMY","SURGICAL AND MEDICAL PROCEDURES",0,"Y",
"PILOT-1002","ASTIGMATISM","EYE DISORDERS",0,"Y",
"PILOT-1003","HEADACHE","NERVOUS SYSTEM DISORDERS",54,"Y",
"PILOT-1003","ABORTION","PREGNANCY, PUERPERIUM AND PERINATAL CONDITIONS",54,"Y",
"PILOT-1003","PRESBYOPIA","EYE DISORDERS",54,"Y",
"PILOT-1003","CYST","GENERAL DISORDERS AND ADMINISTRATION SITE CONDITIONS",54,"Y",
"PILOT-1003","UPPER RESPIRATORY TRACT INFECTION","INFECTIONS AND INFESTATIONS",54,"Y",
"PILOT-1003","DUODENAL ULCER","GASTROINTESTINAL DISORDERS",54,"Y",
"PILOT-1003","MYOCARDIAL INFARCTION","CARDIAC DISORDERS",54,"Y",
"PILOT-1003","BRONCHITIS","INFECTIONS AND INFESTATIONS",54,"Y",
"PILOT-1003","ARTHRITIS","MUSCULOSKELETAL AND CONNECTIVE TISSUE DISORDERS",54,"Y",
"PILOT-1003","NASAL CYST REMOVAL","SURGICAL AND MEDICAL PROCEDURES",54,"Y",
"PILOT-1003","COUGH","RESPIRATORY, THORACIC AND MEDIASTINAL DISORDERS",54,"Y",
"PILOT-1003","GOITRE","ENDOCRINE DISORDERS",54,"Y",
"PILOT-1003","HIP ARTHROPLASTY","SURGICAL AND MEDICAL PROCEDURES",54,"Y",
"PILOT-1003","HIP ARTHROPLASTY","SURGICAL AND MEDICAL PROCEDURES",54,"Y",
"PILOT-1003","GONORRHOEA","INFECTIONS AND INFESTATIONS",54,"Y",
"PILOT-1003","DIZZINESS","NERVOUS SYSTEM DISORDERS",54,"Y",
"PILOT-1003","MALAISE","GENERAL DISORDERS AND ADMINISTRATION SITE CONDITIONS",54,"Y",
"PILOT-1003","IMPAIRED HEALING","GENERAL DISORDERS AND ADMINISTRATION SITE CONDITIONS",54,"Y",
"PILOT-1003","DEAFNESS BILATERAL","EAR AND LABYRINTH DISORDERS",54,"Y",
"PILOT-1003","VERTIGO","EAR AND LABYRINTH DISORDERS",54,"Y",
"PILOT-1003","TONSILLECTOMY","SURGICAL AND MEDICAL PROCEDURES",54,"Y",
"PILOT-1004","ARTHRALGIA","MUSCULOSKELETAL AND CONNECTIVE TISSUE DISORDERS",81,"Y",
"PILOT-1004","DEAFNESS","EAR AND LABYRINTH DISORDERS",81,"Y",
"PILOT-1004","CONSTIPATION","GASTROINTESTINAL DISORDERS",81,"Y",
"PILOT-1004","FRACTURE TREATMENT","SURGICAL AND MEDICAL PROCEDURES",81,"Y",
"PILOT-1004","OSTEITIS DEFORMANS","MUSCULOSKELETAL AND CONNECTIVE TISSUE DISORDERS",81,"Y",
"PILOT-1004","HYPOTHYROIDISM","ENDOCRINE DISORDERS",81,"Y",
"PILOT-1005","CARDIAC DISORDER","CARDIAC DISORDERS",81,"Y",
"PILOT-1005","LIBIDO INCREASED","PSYCHIATRIC DISORDERS",81,"Y",
"PILOT-1005","ELECTROCARDIOGRAM ST SEGMENT ELEVATION","INVESTIGATIONS",81,"Y",
"PILOT-1005","MYOCARDIAL INFARCTION","CARDIAC DISORDERS",81,"Y",
"PILOT-1005","MYOCARDIAL INFARCTION","CARDIAC DISORDERS",81,"Y",
"PILOT-1005","ELECTROCARDIOGRAM ST SEGMENT DEPRESSION","INVESTIGATIONS",81,"Y",
"PILOT-1005","TRIPLE VESSEL BYPASS GRAFT","SURGICAL AND MEDICAL PROCEDURES",81,"Y",
"PILOT-1005","EXTRASYSTOLES","CARDIAC DISORDERS",81,"Y",
"PILOT-1006","JOINT DISLOCATION","INJURY, POISONING AND PROCEDURAL COMPLICATIONS",54,"Y",
"PILOT-1006","DIPHTHERIA","INFECTIONS AND INFESTATIONS",54,"Y",
"PILOT-1006","BREAST COSMETIC SURGERY","SURGICAL AND MEDICAL PROCEDURES",54,"Y",
"PILOT-1007","HAEMORRHOIDS","GASTROINTESTINAL DISORDERS",0,"Y",
"PILOT-1007","HEADACHE","NERVOUS SYSTEM DISORDERS",0,"Y",
"PILOT-1007","GLAUCOMA","EYE DISORDERS",0,"Y",
"PILOT-1007","CONSTIPATION","GASTROINTESTINAL DISORDERS",0,"Y",
"PILOT-1007","OSTEOARTHRITIS","MUSCULOSKELETAL AND CONNECTIVE TISSUE DISORDERS",0,"Y",
"PILOT-1007","DEPRESSION","PSYCHIATRIC DISORDERS",0,"Y",
"PILOT-1007","DERMATITIS","SKIN AND SUBCUTANEOUS TISSUE DISORDERS",0,"Y",
"PILOT-1007","GANGLION","MUSCULOSKELETAL AND CONNECTIVE TISSUE DISORDERS",0,"Y",
"PILOT-1007","CHOLECYSTECTOMY","SURGICAL AND MEDICAL PROCEDURES",0,"Y",
"PILOT-1007","BENIGN PROSTATIC HYPERPLASIA","REPRODUCTIVE SYSTEM AND BREAST DISORDERS",0,"Y",
"PILOT-1007","APPENDICECTOMY","SURGICAL AND MEDICAL PROCEDURES",0,"Y",
"PILOT-1008","PROSTATE CANCER","NEOPLASMS BENIGN, MALIGNANT AND UNSPECIFIED (INCL CYSTS AND POLYPS)",81,"Y",
"PILOT-1008","HYPERTENSION","VASCULAR DISORDERS",81,"Y",
"PILOT-1008","HERNIA REPAIR","SURGICAL AND MEDICAL PROCEDURES",81,"Y",
"PILOT-1009","OOPHORECTOMY PARTIAL","SURGICAL AND MEDICAL PROCEDURES",54,"Y",
"PILOT-1009","BUNDLE BRANCH BLOCK RIGHT","CARDIAC DISORDERS",54,"Y",
"PILOT-1009","ATRIAL HYPERTROPHY","CARDIAC DISORDERS",54,"Y",
"PILOT-1010","MASTOIDITIS","INFECTIONS AND INFESTATIONS",0,"Y",
"PILOT-1010","HYPERLIPIDAEMIA","METABOLISM AND NUTRITION DISORDERS",0,"Y",
"PILOT-1011","SHOULDER OPERATION","SURGICAL AND MEDICAL PROCEDURES",81,"Y",
"PILOT-1011","SUPRAVENTRICULAR EXTRASYSTOLES","CARDIAC DISORDERS",81,"Y",
"PILOT-1011","MASTOIDECTOMY","SURGICAL AND MEDICAL PROCEDURES",81,"Y",
"PILOT-1011","EXTRASYSTOLES","CARDIAC DISORDERS",81,"Y",
"PILOT-1011","DEAFNESS UNILATERAL","EAR AND LABYRINTH DISORDERS",81,"Y",
"PILOT-1011","URINARY INCONTINENCE","RENAL AND URINARY DISORDERS",81,"Y",
"PILOT-1012","VASCULAR BYPASS GRAFT","SURGICAL AND MEDICAL PROCEDURES",0,"Y",
"PILOT-1012","HEAD INJURY","INJURY, POISONING AND PROCEDURAL COMPLICATIONS",0,"Y",
"PILOT-1012","CATARACT OPERATION","SURGICAL AND MEDICAL PROCEDURES",0,"Y",
"PILOT-1012","MALIGNANT TUMOUR EXCISION","SURGICAL AND MEDICAL PROCEDURES",0,"Y",
"PILOT-1012","BUNDLE BRANCH BLOCK RIGHT","CARDIAC DISORDERS",0,"Y",
"PILOT-1012","CONSTIPATION","GASTROINTESTINAL DISORDERS",0,"Y",
"PILOT-1012","AORTIC VALVE REPLACEMENT","SURGICAL AND MEDICAL PROCEDURES",0,"Y",
"PILOT-1012","HEARING AID USER","SOCIAL CIRCUMSTANCES",0,"Y",
"PILOT-1012","HYDROCELE","CONGENITAL, FAMILIAL AND GENETIC DISORDERS",0,"Y",
"PILOT-1012","HERNIA REPAIR","SURGICAL AND MEDICAL PROCEDURES",0,"Y",
"PILOT-1012","PNEUMONIA","INFECTIONS AND INFESTATIONS",0,"Y",
"PILOT-1012","ANGINA PECTORIS","CARDIAC DISORDERS",0,"Y",
"PILOT-1012","ATRIOVENTRICULAR BLOCK","CARDIAC DISORDERS",0,"Y",
"PILOT-1014","CARDIAC FAILURE CONGESTIVE","CARDIAC DISORDERS",0,"Y",
"PILOT-1014","RECTAL HAEMORRHAGE","GASTROINTESTINAL DISORDERS",0,"Y",
"PILOT-1014","DEAFNESS","EAR AND LABYRINTH DISORDERS",0,"Y",
"PILOT-1014","PNEUMONIA","INFECTIONS AND INFESTATIONS",0,"Y",
"PILOT-1014","HYPOTHYROIDISM","ENDOCRINE DISORDERS",0,"Y",
"PILOT-1014","SCOLIOSIS","MUSCULOSKELETAL AND CONNECTIVE TISSUE DISORDERS",0,"Y",
"PILOT-1015","CAROTID ENDARTERECTOMY","SURGICAL AND MEDICAL PROCEDURES",54,"Y",
"PILOT-1015","HIP ARTHROPLASTY","SURGICAL AND MEDICAL PROCEDURES",54,"Y",
"PILOT-1015","HIP ARTHROPLASTY","SURGICAL AND MEDICAL PROCEDURES",54,"Y",
"PILOT-1015","CORONARY ARTERY SURGERY","SURGICAL AND MEDICAL PROCEDURES",54,"Y",
)


# Simulate twice the original subject count while preserving treatment, SOC,
# preferred-term and safety-population relationships.
maps <- make_id_maps(list(ADSL=adsl, ADMH=admh))
adsl <- double_subjects(adsl, maps)
admh <- double_subjects(admh, maps)
for(soc in unique(admh$mhbodsys)) {
  ix <- which(admh$mhbodsys == soc)
  terms <- unique(admh$mhdecod[ix])
  change <- ix[runif(length(ix)) < 0.05]
  if(length(change) && length(terms)>1) admh$mhdecod[change] <- sample(terms,length(change),replace=TRUE)
}
stopifnot(n_distinct(adsl$usubjid) == 2L * maps$original_n)
write_xlsx_list(list(ADSL=adsl, ADMH=admh), file.path(OUT_DIR,paste0(STEM,"_simulated_data.xlsx")))

# Original analysis structure retained from the source document.
#==============================================================================;
#Filter required records;
#==============================================================================;
adsl01<-adsl %>%
 filter(saffl=="Y")
admh01<-admh  %>%
 filter(saffl=="Y")
#==============================================================================;
#Create treatment variable;
#==============================================================================;
adsl02<-adsl01 %>% mutate(treatment=trt01an) %>%
 bind_rows(adsl01 %>% mutate(treatment=99))
admh02<-admh01 %>% mutate(treatment=trt01an) %>%
 bind_rows(admh01 %>% mutate(treatment=99))
#==============================================================================;
#Get treatment totals;
#==============================================================================;
#------------------------------------------------------------------------------;
#Create a dummy dataset containing all possible treatment levels;
#------------------------------------------------------------------------------;
dummy_trttotals <- tibble(treatment = c(0, 54, 81, 99))
#------------------------------------------------------------------------------;
#Get actual treatment totals;
#------------------------------------------------------------------------------;
trttotals_pre<-adsl02 %>%
 count(treatment)
#------------------------------------------------------------------------------;
#Merge with dummy treatment totals;
#------------------------------------------------------------------------------;
trttotals<-dummy_trttotals  %>%
 left_join(trttotals_pre,by=c("treatment")) %>%
 mutate(trttotal=if_else(is.na(n),0,n)) %>%
 select(-n)
#==============================================================================;
#Obtaining actual counts-for the table;
#==============================================================================;
#------------------------------------------------------------------------------;
#TOP row;
#------------------------------------------------------------------------------;
sub_count <- admh02 %>%
 group_by(treatment) %>%
 summarise(label = "Overall", count = n_distinct(usubjid), events=n()) %>%
 ungroup()
#------------------------------------------------------------------------------;
#SOC rows;
#------------------------------------------------------------------------------;
soc_count <- admh02 %>%
 group_by(mhbodsys, treatment) %>%
summarise(count = n_distinct(usubjid),events = n(), .groups="drop") %>%
 ungroup()
#------------------------------------------------------------------------------;
#PT rows;
#------------------------------------------------------------------------------;
pt_count <- admh02 %>%
 group_by(mhbodsys, mhdecod, treatment) %>%
summarise(count = n_distinct(usubjid),events = n(), .groups="drop") %>%
 ungroup()
#------------------------------------------------------------------------------;
#Combine toprow, SOC, and PT level counts into single dataset;
#Replace NAs with blanks
#------------------------------------------------------------------------------;
counts01 <- bind_rows(sub_count, soc_count, pt_count) %>%
 mutate(across(c(label, mhbodsys, mhdecod), ~ if_else(is.na(.), "", .)))
#==============================================================================;
#Create zero counts if an event is not present in a treatment;
#==============================================================================;
#------------------------------------------------------------------------------;
#Get all the available SOC and PT values;
#Create a row for each treatment;
#------------------------------------------------------------------------------;
dummy01<-counts01 %>%distinct(label,mhbodsys,mhdecod)
dummy02<-cross_join(dummy01,dummy_trttotals)
#==============================================================================;
#Merge dummy counts with actual counts;
#Set subject count and event count to 0 for missing rows;
#==============================================================================;
counts02<-dummy02 %>%
 left_join(counts01,by=c("label","mhbodsys", "mhdecod", "treatment")) %>%
 mutate(across(c(count,events),~if_else(is.na(.),0,.)))
#==============================================================================;
#Calculate percentages;
#==============================================================================;
#------------------------------------------------------------------------------;
#Fetch denominators;
#------------------------------------------------------------------------------;
counts03<-counts02 %>%
 left_join(trttotals,by="treatment")
#------------------------------------------------------------------------------;
#Create a single variable to hold count and percentage;
#------------------------------------------------------------------------------;
counts04<-counts03 %>%
 mutate(
 percent=if_else(trttotal!=0,count/trttotal*100,0),
 percentc=str_c(" (",sprintf("%.1f",percent),"%)"),
 cp=if_else(count==0,"0",str_c(count, percentc))
 )
#==============================================================================;
#Create label column;
#==============================================================================;
counts05<-counts04 %>%
 mutate(
 label=case_when(
 mhbodsys=="" & mhdecod=="" ~ label,
 mhbodsys!="" & mhdecod=="" ~ mhbodsys,
 mhbodsys!="" & mhdecod!="" ~ str_c("   ",mhdecod),
 TRUE~""
 )
 )
#==============================================================================;
#Transpose to obtain treatment as columns;
#==============================================================================;
counts06 <- counts05 %>%
 pivot_wider(
 id_cols = c(mhbodsys, mhdecod, label),
 names_from = treatment,
 values_from = cp,
 names_prefix = "trt"
 ) %>%
 arrange(mhbodsys,mhdecod,label)
output<-counts06 %>%
 select(mhbodsys,mhdecod,trt0,trt54, trt81)

n0 <- n_distinct(adsl$usubjid[adsl$trt01an==0 & adsl$saffl=="Y"])
n54 <- n_distinct(adsl$usubjid[adsl$trt01an==54 & adsl$saffl=="Y"])
n81 <- n_distinct(adsl$usubjid[adsl$trt01an==81 & adsl$saffl=="Y"])
display <- output %>% transmute(
  `System Organ Class / Preferred Term`=case_when(
    mhbodsys=="" & mhdecod=="" ~ "Overall",
    mhdecod=="" ~ mhbodsys,
    TRUE ~ paste0("   ",mhdecod)
  ),
  !!paste0("Treatment 0 (N=",n0,")") := trt0,
  !!paste0("Treatment 54 (N=",n54,")") := trt54,
  !!paste0("Treatment 81 (N=",n81,")") := trt81
)
write_clinical_rtf(display,
  "Table 1. Summary of Medical History by System Organ Class and Preferred Term",
  file.path(OUT_DIR,paste0(STEM,".rtf")),
  "n = number of subjects with medical history; percentages use treatment-group Safety Analysis Set denominators. Subjects may be counted in more than one category."
)
cat("Created ",STEM,": simulated N = ",2L*maps$original_n,"\n",sep="")
