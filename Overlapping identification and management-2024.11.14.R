library(openxlsx)
library(tidyr)
rm(list=ls()); gc()

##Phase one - step one - Importing standardized PICO data----

data_PICO <- read.xlsx("Example standardized PICO of meta-analysis.xlsx",sheet=1,na.strings = "")
data_PICO <- unite(data_PICO, Meta_analysis, c("Meta-analysis.ID", "Author.(first.author’s.last.name)","Year.of.Publication"), remove = F)
colnames(data_PICO)

##Phase one - Step two - categorizing topic clusters and identifying overlaps----
#p_begin: 8; p_last: 9 Indicates the column to which the patient belongs in "data_PICO"
#i_begin: 10; i_last: 14 Indicates the column to which the intervention belongs in "data_PICO"
#c_begin: 15; c_last: 16 Indicates the column to which the control belongs in "data_PICO"
#o_begin: 17; o_last: 18 Indicates the column to which the outcome belongs in "data_PICO"

#1.categorizing topic clusters
PICO_cluster <- function(data_PICO = data_PICO, p_begin, p_last, i_begin, i_last, c_begin, c_last,o_begin, o_last) {
  PICO_cluster <- list()
  PICO_combination <- data.frame(matrix(ncol = 2, nrow = 0))
  colnames(PICO_combination) <- c("number_PICO", "number_PICO_afterde")
  for (p in c(1:length(c(p_begin:p_last)))) {
    patient <- p_begin-1 + p
    data_PICO1 <- data_PICO[which(data_PICO[, patient] ==1),]
    for (i in c(1:length(c(i_begin:i_last)))) {
      intervention <- i_begin-1 + i
      data_PICO2 <-data_PICO1
      data_PICO2 <- data_PICO2[which(data_PICO2[, intervention] ==1),]
      for (c in c(1:length(c(c_begin:c_last)))) {
        control <- c_begin-1 + c
        data_PICO3 <-data_PICO2
        data_PICO3 <- data_PICO3[which(data_PICO3[, control] ==1),]
        for (o in c(1:length(c(o_begin:o_last)))) {
          outcome <- o_begin-1 + o
          data_PICO4 <-data_PICO3
          data_PICO4 <- data_PICO4[which(data_PICO4[, outcome] ==1 | is.na(data_PICO4[, outcome]) == F),]
          data_PICO_de4 <- data_PICO4[!duplicated(data_PICO4[c(patient,intervention,control,outcome)]),]
          print(nrow(data_PICO4))
          name <- paste(colnames(data_PICO4[patient]), colnames(data_PICO4[intervention]), colnames(data_PICO4[control]), colnames(data_PICO4[outcome]),sep = "_")
          
          rownames(PICO_combination[nrow(PICO_combination)+1,]) <- name
          PICO_combination[nrow(PICO_combination),"number_PICO"] <- nrow(data_PICO4)
          PICO_combination[nrow(PICO_combination),"number_PICO_de"] <- nrow(data_PICO_de4)
          if (nrow(data_PICO4) != 0 ) {
            PICO_cluster <- c(PICO_cluster, list(o = data_PICO4[c("Meta-analysis.ID","Meta_analysis","Author.(first.author’s.last.name)","Year.of.Publication")])) 
            names(PICO_cluster)[length(PICO_cluster)] <- name}
        }
      }
    }
  }
  return(PICO_cluster)
}
PICO_cluster_example <- PICO_cluster(data_PICO, p_begin=8, p_last=9, i_begin=10, i_last=14, c_begin=15, c_last=16,o_begin=17, o_last=18)

#2.identifying overlaps
# PICO_cluster: from the step two
identify_overlappingMA <- function(PICO_cluster){
  overlappingMA <- data.frame(matrix(ncol = 5, nrow = 0))
  colnames(overlappingMA) <- c("Meta-analysis.ID","Meta_analysis","Author.(first.author’s.last.name)","Year.of.Publication","PICO")
  for (i in c(1:length(PICO_cluster))) {
    data <- PICO_cluster[i]
    data_1 <- as.data.frame(data)
    colnames(data_1) <- c("Meta-analysis.ID","Meta_analysis","Author.(first.author’s.last.name)","Year.of.Publication")
    data_1$PICO <- names(data)
    if (nrow(data_1) > 1 ) {
      overlappingMA <- rbind (overlappingMA, data_1)
    }
  }
  return(overlappingMA)
}
overlappingMA_example <- identify_overlappingMA(PICO_cluster=PICO_cluster_example)

##Phase one - Step three - managing overlaps----
#1.Importing primary studies cited in meta-analyses
data_primarystudy <- read.xlsx("Example basic information of included primary study.xlsx",sheet=1,na.strings = "")
colnames(data_primarystudy)
data_primarystudy <- unite(data_primarystudy, primary_study_journal, c("Primary.study.Year.of.publication", "Primary.study.Author.(first.author’s.last.name)","Primary.study.Journal.name.(full.name)"), remove = F)

#2.creating citation matrix and calculate pairwise corrected covered area (CCA) index
# PICO_cluster: from the step two:categorizing topic clusters and identifying overlaps
# data_PICO: from step one: standardized PICO of meta-analysis
# data_primarystudy: from step three2: basic information of included primary study
creat_citationmatrix_CCA <- function(PICO_cluster, data_PICO, data_primarystudy){
  opencm <- list()
  for (i in c(1:length(PICO_cluster))) {
    data <- PICO_cluster[i]
    data <- as.data.frame(data)
    colnames(data) <- c("Meta-analysis.ID","Meta_analysis","Author.(first.author’s.last.name)","Year.of.Publication") 
    if (nrow(data) > 1 ) {
      data_MA <- subset(data_PICO, data_PICO$"Meta-analysis.ID" %in% c(data$"Meta-analysis.ID"))
      CCA_C <- nrow(data_MA)
      data_pr <- subset(data_primarystudy, data_primarystudy$"Meta-analysis.ID" %in% c(data_MA$"Meta-analysis.ID"))
      primary_study_journal <- data_primarystudy[which(data_primarystudy$"Meta-analysis.ID" %in% c(data_MA$"Meta-analysis.ID")),"primary_study_journal"]
      CCA_N <- length(primary_study_journal)
      primary_study_journal <- unique(primary_study_journal)
      CCA_r <- length(primary_study_journal)
      CCA <- (CCA_N - CCA_r)/(CCA_r*CCA_C - CCA_r)
      table_CCA <- data.frame(matrix(ncol = 1, nrow = 1))
      colnames(table_CCA) <- names(PICO_cluster)[i]
      rownames(table_CCA) <- "allma"
      table_CCA[1,1] <- CCA
      opencm_order <- data.frame(matrix(ncol = length(primary_study_journal), nrow = length(data_MA$"Meta-analysis.ID")))
      rownames(opencm_order) <- data_MA$"Meta_analysis"
      colnames(opencm_order) <- primary_study_journal
      for (j in c(1: nrow(opencm_order))) {
        ID <- data_MA$"Meta-analysis.ID"[j]
        included_primary_study <- data_pr$primary_study_journal[which(data_pr$"Meta-analysis.ID" == ID)]
        for (k in c(1:length(included_primary_study))) {
          for (t in c(1: length(primary_study_journal))) {
            if (included_primary_study[k] == primary_study_journal[t]) {opencm_order[j,t] <- included_primary_study[k]}
          }
        }
      }
      opencm <- c(opencm, list(o = opencm_order))
      names(opencm)[length(opencm)] <- colnames(table_CCA)
      for (p in c(1:(nrow(opencm_order)-1))) {
        primary_study_journal1 <- as.character(na.omit(as.character(opencm_order[p,])))
        for (q in c((p+1):nrow(opencm_order))) {
          primary_study_journal2 <- as.character(na.omit(as.character(opencm_order[q,])))
          if (as.character(rownames(opencm_order)[p]) != as.character(rownames(opencm_order)[q])) {
            primary_study_journal3 <- c(primary_study_journal1, primary_study_journal2)
          CCA_N <- length(primary_study_journal3)
          primary_study_journal3 <- unique(primary_study_journal3)
          CCA_r <- length(unique(primary_study_journal3))
          CCA <- (CCA_N - CCA_r)/(CCA_r*2 - CCA_r)
          table_CCA[nrow(table_CCA)+1,1] <- CCA
          rownames(table_CCA)[nrow(table_CCA)] <- paste(as.character(rownames(opencm_order)[p]), as.character(rownames(opencm_order)[q]), sep = "+")
          }
        }
      }
      opencm <- c(opencm, list(o = table_CCA))
      names(opencm)[length(opencm)] <- paste(names(PICO_cluster)[i], "CCA", sep = "_")
    } 
  }
  return(opencm)
}
opencm_example <- creat_citationmatrix_CCA(PICO_cluster_example, data_PICO, data_primarystudy)

#3.remove duplicates based on pre-defined rule
# opencm: from step three2: citation matrix and pairwise CCA
# data_PICO: from step one: standardized PICO of meta-analysis
# overlappingMA: from step two2:identifying overlaps
opencm <- opencm_example
i=2
remove_overlappingMA <- function(opencm, data_PICO,overlappingMA){
  retain_nonoverlappingMA <- list()
  for (i in c(1:(length(opencm)/2))) {
    opencm_CCA_orig <- opencm[[i*2]]
    opencm_CCA_orig$"MA_compare_all" <- rownames(opencm_CCA_orig)
    opencm_CCA_orig <- opencm_CCA_orig[-1,]
    opencm_CCA <- data.frame(separate_wider_delim(opencm_CCA_orig[c(1:nrow(opencm_CCA_orig)),], "MA_compare_all", "+", names = c("MA1","MA2")))
    opencm_CCA <- opencm_CCA[order(opencm_CCA[,1],decreasing = T), ]
    process_overlappingMA <- data.frame(matrix(ncol = 4, nrow = 0))
    colnames(process_overlappingMA) <- c("PICO","Included MA", "retain_MA", "CCA")
 
    if (opencm_CCA[1,1] <= 0.05 ) {
      All_MA <- unique(c(opencm_CCA$MA1, opencm_CCA$MA2))
      for (j in 1:length(All_MA)) {
        MA <- All_MA[j]
        process_overlappingMA[nrow(process_overlappingMA)+1,"PICO"] <- colnames(opencm_CCA_orig)[1]
        process_overlappingMA[nrow(process_overlappingMA),"Included MA"] <- paste(overlappingMA[which(overlappingMA$PICO == colnames(opencm_CCA_orig)[1]),"Meta_analysis"], collapse = "; ")
        process_overlappingMA[nrow(process_overlappingMA),"retain_MA"] <- MA
        process_overlappingMA[nrow(process_overlappingMA),"CCA"] <- "CCA<=0.05"
      }
    } else if (opencm_CCA[nrow(opencm_CCA),1] > 0.05){
      
      compareMA <- unique(c(opencm_CCA$MA1, opencm_CCA$MA2))
      comparebasic <- subset(data_PICO[,c("Meta_analysis","Year.of.Publication","Number.of.included.primary.study")], data_PICO$"Meta_analysis" %in% compareMA)
      comparebasic$Year.of.Publication <- as.numeric(comparebasic$Year.of.Publication)
      comparebasic$Number.of.included.primary.study <- as.numeric(comparebasic$Number.of.included.primary.study)
      retainMA <- subset(comparebasic, comparebasic$Year.of.Publication == max(comparebasic$Year.of.Publication))
      if (nrow(retainMA) > 1) {
        retainMA <- subset(retainMA, retainMA$Number.of.included.primary.study == max(retainMA$Number.of.included.primary.study))
      }
      for (j in 1:nrow(retainMA)) {
        CCA_MA <- retainMA[j]
        process_overlappingMA[nrow(process_overlappingMA)+1,"PICO"] <- colnames(opencm_CCA_orig)[1]
        process_overlappingMA[nrow(process_overlappingMA),"Included MA"] <- paste(overlappingMA[which(overlappingMA$PICO == colnames(opencm_CCA_orig)[1]),"Meta_analysis"], collapse = "; ")
        process_overlappingMA[nrow(process_overlappingMA),"retain_MA"] <- CCA_MA
        if (nrow(retainMA) > 1) {
          process_overlappingMA[nrow(process_overlappingMA),"CCA"] <- "CCA>0.05_same"
        } else {process_overlappingMA[nrow(process_overlappingMA),"CCA"] <- "CCA>0.05"}
      }
    } else {
      
      CCAhigh <- subset(opencm_CCA, opencm_CCA[1] >0.05)
      compareCCAhigh <- unique(c(CCAhigh$MA1, CCAhigh$MA2))
      
      comparebasic <- subset(data_PICO[,c("Meta_analysis","Year.of.Publication","Number.of.included.primary.study")], data_PICO$"Meta_analysis" %in% compareCCAhigh)
      comparebasic$Year.of.Publication <- as.numeric(comparebasic$Year.of.Publication)
      comparebasic$Number.of.included.primary.study <- as.numeric(comparebasic$Number.of.included.primary.study)
      retainMA <- subset(comparebasic, comparebasic$Year.of.Publication == max(comparebasic$Year.of.Publication))
      if (nrow(retainMA) > 1) {
        retainMA <- subset(retainMA, retainMA$Number.of.included.primary.study == max(retainMA$Number.of.included.primary.study))
      }
      CCAhigh_retain <-retainMA$Meta_analysis
      for (j in 1:length(CCAhigh_retain)) {
        CCAhigh_MA <- CCAhigh_retain[j]
        process_overlappingMA[nrow(process_overlappingMA)+1,"PICO"] <- colnames(opencm_CCA_orig)[1]
        process_overlappingMA[nrow(process_overlappingMA),"Included MA"] <- paste(overlappingMA[which(overlappingMA$PICO == colnames(opencm_CCA_orig)[1]),"Meta_analysis"], collapse = "; ")
        process_overlappingMA[nrow(process_overlappingMA),"retain_MA"] <- CCAhigh_MA
        if (length(CCAhigh_retain) > 1) {
          process_overlappingMA[nrow(process_overlappingMA),"CCA"] <- "CCA>0.05_same"
        } else {process_overlappingMA[nrow(process_overlappingMA),"CCA"] <- "CCA>0.05"}
      }

      CCAlow <- subset(opencm_CCA, opencm_CCA[1] <=0.05)
      CCAlow_includeretainMA <- subset(CCAlow[,c("MA1","MA2")], CCAlow$MA1 %in% CCAhigh_retain | CCAlow$MA2 %in% CCAhigh_retain)
      CCAlow_includeretainMA <- unique(c(CCAlow$MA1, CCAlow$MA2))
      CCAlow_includeretainMA <- CCAlow_includeretainMA[-which(CCAlow_includeretainMA == CCAhigh_retain)]
      for (j in 1:length(CCAlow_includeretainMA)) {
        CCAlow_MA <- CCAlow_includeretainMA[j]
        process_overlappingMA[nrow(process_overlappingMA)+1,"PICO"] <- colnames(opencm_CCA_orig)[1]
        process_overlappingMA[nrow(process_overlappingMA),"Included MA"] <- paste(overlappingMA[which(overlappingMA$PICO == colnames(opencm_CCA_orig)[1]),"Meta_analysis"], collapse = "; ")
        process_overlappingMA[nrow(process_overlappingMA),"retain_MA"] <- CCAlow_MA
        process_overlappingMA[nrow(process_overlappingMA),"CCA"] <- "CCA<=0.05"
      }
    }
    
    retain_nonoverlappingMA <- c(retain_nonoverlappingMA, list(o = process_overlappingMA))
    names(retain_nonoverlappingMA)[length(retain_nonoverlappingMA)] <- colnames(opencm[[i*2]])
  }
  return(retain_nonoverlappingMA)
}
remove_overlappingMA_example <- remove_overlappingMA(opencm_example, data_PICO,overlappingMA_example)

##Phase one - Step four - constructing data location table----
# PICO_cluster: from the step two:categorizing topic clusters and identifying overlaps
# data_PICO: from step one: standardized PICO of meta-analysis
# remove_overlappingMA: from step three4: remove overlapping MAs
collect_needextractMA <- function(PICO_cluster, data_PICO, remove_overlappingMA) {
  nooverlappingcluster <- data.frame(matrix(ncol = 3, nrow = 0))
  colnames(nooverlappingcluster) <- c("PICO", "retain_MA","CCA")
  for (i in c(1:length(PICO_cluster))) {
    data <- PICO_cluster[i]
    data <- as.data.frame(data)
    colnames(data) <- c("Meta-analysis.ID","Meta_analysis","Author.(first.author’s.last.name)","Year.of.Publication")
    if (nrow(data) == 1) {
      nooverlappingcluster[nrow(nooverlappingcluster)+1, "retain_MA"] <- data_PICO[which(data_PICO$"Meta-analysis.ID" == data$"Meta-analysis.ID"),c("Meta_analysis")]
      nooverlappingcluster[nrow(nooverlappingcluster), "PICO"] <- names(PICO_cluster)[i]
    }
  }
  
  needextractMA <- data.frame(matrix(ncol = length(names(PICO_cluster)), nrow = nrow(data_PICO)))
  colnames(needextractMA) <- names(PICO_cluster)
  rownames(needextractMA) <- data_PICO$"Meta_analysis"
  for (i in 1:(length(remove_overlappingMA))) {
    retain_MA <- remove_overlappingMA[[i]]
    for (j in c(1: nrow(retain_MA))) {
      need_MA <- retain_MA$retain_MA[j]
      for (k in c(1:nrow(needextractMA))){
        if (rownames(needextractMA[k,]) == need_MA) {needextractMA[k,c(unique(retain_MA$PICO))] <- need_MA}
      }
    }
  }
  for (i in c(1:(nrow(nooverlappingcluster)))) {
    for (j in c(1: nrow(needextractMA))) {
      if (rownames(needextractMA[j,]) == nooverlappingcluster[i,"retain_MA"]) {needextractMA[j,c(nooverlappingcluster[i,"PICO"])] <- nooverlappingcluster[i,"retain_MA"]}
    }
  }
  needextractMA$"Review.ID" <- substr(rownames(needextractMA),1,1)
  needextractMA <- aggregate(needextractMA[1:length(names(PICO_cluster))], needextractMA["Review.ID"], FUN= function(X){paste(unique(X[is.na(X)==F]))})
  needextractMA[needextractMA == "character(0)"] <- NA
  rownames(needextractMA) <- needextractMA$"Review.ID"
  needextractMA <- needextractMA[-1]
  needextractMA$total <- rowSums(is.na(needextractMA[,c(1:length(names(PICO_cluster)))]) == F)
  needextractMA[nrow(needextractMA)+1, -ncol(needextractMA)] <- colSums(is.na(needextractMA[c(1:length(unique(data_PICO$"Review.ID"))),-ncol(needextractMA)]) == F)
  needextractMA[nrow(needextractMA), "total"] <- sum(needextractMA[c(1:(nrow(needextractMA)-1)),"total"])
  needextractMA <- needextractMA[needextractMA$total != 0,]
  
  return(needextractMA)
}
collect_needextractMA_example <- collect_needextractMA(PICO_cluster_example, data_PICO, remove_overlappingMA_example)

##Phase one - supplementary - constructing supplementary data location table----
# opencm: from step three2: citation matrix and pairwise CCA
# PICO_cluster: from the step two:categorizing topic clusters and identifying overlaps
# data_PICO: from step one: standardized PICO of meta-analysis
collect_needextractMA_S <- function(opencm, PICO_cluster, data_PICO){
  retain_alloverlappingMA <- list()
  for (i in c(1:(length(opencm)/2))) {
    opencm_CCA <- opencm[[i*2]]
    opencm_CCA$"MA_compare_all" <- rownames(opencm_CCA)
    opencm_CCA <- data.frame(separate_wider_delim(opencm_CCA[c(2:nrow(opencm_CCA)),], "MA_compare_all", "+", names = c("MA1","MA2")))
    opencm_CCA <- opencm_CCA[order(opencm_CCA[,1],decreasing = T), ]
    process_overlappingMA <- data.frame(matrix(ncol = 3, nrow = 0))
    colnames(process_overlappingMA) <- c("PICO", "retain_MA", "CCA")
    for (j in c(1:nrow(opencm_CCA))) {
      CCA <- opencm_CCA[j,1]
      process_overlappingMA[nrow(process_overlappingMA)+1,"retain_MA"] <- opencm_CCA[j,"MA1"]
      process_overlappingMA[nrow(process_overlappingMA),"CCA"] <- CCA
      process_overlappingMA[nrow(process_overlappingMA),"PICO"] <- colnames(opencm[[i*2]])
      process_overlappingMA[nrow(process_overlappingMA)+1,"retain_MA"] <- opencm_CCA[j,"MA2"]
      process_overlappingMA[nrow(process_overlappingMA),"CCA"] <- CCA
      process_overlappingMA[nrow(process_overlappingMA),"PICO"] <- colnames(opencm[[i*2]])
    }
    retain_alloverlappingMA <- c(retain_alloverlappingMA, list(o = process_overlappingMA))
    names(retain_alloverlappingMA)[length(retain_alloverlappingMA)] <- colnames(opencm[[i*2]])
    process_overlappingMA <- process_overlappingMA[!duplicated(process_overlappingMA["retain_MA"]),]
    retain_alloverlappingMA <- c(retain_alloverlappingMA, list(o = process_overlappingMA))
    names(retain_alloverlappingMA)[length(retain_alloverlappingMA)] <- paste(colnames(opencm[[i*2]]),"de",sep = "_")
  }
  
  nooverlappingcluster <- data.frame(matrix(ncol = 3, nrow = 0))
  colnames(nooverlappingcluster) <- c("PICO", "retain_MA","CCA")
  for (i in c(1:length(PICO_cluster))) {
    data <- PICO_cluster[i]
    data <- as.data.frame(data)
    colnames(data) <- c("Meta-analysis.ID","Meta_analysis","Author.(first.author’s.last.name)","Year.of.Publication")
    if (nrow(data) == 1) {
      nooverlappingcluster[nrow(nooverlappingcluster)+1, "retain_MA"] <- data_PICO[which(data_PICO$"Meta-analysis.ID" == data$"Meta-analysis.ID"),c("Meta_analysis")]
      nooverlappingcluster[nrow(nooverlappingcluster), "PICO"] <- names(PICO_cluster)[i]
    }
  }
  
  needextractMA_S <- data.frame(matrix(ncol = length(names(PICO_cluster)), nrow = nrow(data_PICO)))
  colnames(needextractMA_S) <- names(PICO_cluster)
  rownames(needextractMA_S) <- data_PICO$"Meta_analysis"
  for (i in c(1:(length(retain_alloverlappingMA)/2))) {
    retain_MA <- retain_alloverlappingMA[[i*2]]
    for (j in c(1: nrow(retain_MA))) {
      need_MA <- retain_MA$retain_MA[j]
      for (k in c(1:nrow(needextractMA_S))){
        if (rownames(needextractMA_S[k,]) == need_MA) {needextractMA_S[k,c(unique(retain_MA$PICO))] <- need_MA}
      }
    }
  }
  for (i in c(1:(nrow(nooverlappingcluster)))) {
    for (j in c(1: nrow(needextractMA_S))) {
      if (rownames(needextractMA_S[j,]) == nooverlappingcluster[i,"retain_MA"]) {needextractMA_S[j,c(nooverlappingcluster[i,"PICO"])] <- nooverlappingcluster[i,"retain_MA"]}
    }
  }
  needextractMA_S$"Review.ID" <- substr(rownames(needextractMA_S),1,1)
  needextractMA_S <- aggregate(needextractMA_S[1:length(names(PICO_cluster))], needextractMA_S["Review.ID"], FUN= function(X){paste(unique(X[is.na(X)==F]))})
  needextractMA_S[needextractMA_S == "character(0)"] <- NA
  rownames(needextractMA_S) <- needextractMA_S$"Review.ID"
  needextractMA_S <- needextractMA_S[-1]
  needextractMA_S$total <- rowSums(is.na(needextractMA_S[,c(1:length(names(PICO_cluster)))]) == F)
  needextractMA_S[nrow(needextractMA_S)+1, -ncol(needextractMA_S)] <- colSums(is.na(needextractMA_S[c(1:length(unique(data_PICO$"Review.ID"))),-ncol(needextractMA_S)]) == F)
  needextractMA_S[nrow(needextractMA_S), "total"] <- sum(needextractMA_S[c(1:(nrow(needextractMA_S)-1)),"total"])
  needextractMA_S <- needextractMA_S[needextractMA_S$total != 0,]
  
  return(needextractMA_S)
}
collect_needextractMA_S_example <- collect_needextractMA_S(opencm_example, PICO_cluster_example, data_PICO)

##Phase one - outputting data location table----
outputdata <- createWorkbook()
addWorksheet(outputdata, "MA_after_removing")
writeData(outputdata, sheet = "MA_after_removing", collect_needextractMA_example, rowNames = TRUE)
addWorksheet(outputdata, "MA_before_removing")
writeData(outputdata, sheet = "MA_before_removing", collect_needextractMA_S_example, rowNames = TRUE)
saveWorkbook(outputdata, file = "collect_needextractMA.xlsx", overwrite = TRUE)
