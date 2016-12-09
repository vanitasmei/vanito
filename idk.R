rm(list = ls())
library(dplyr)
base_folder <- setwd("~/Vanitas/Assessment TN1")
get_summary <- function(MCO = "AG"){
  file_names <- list.files(path = file.path(base_folder, MCO))
  file_paths <- paste(base_folder, MCO, file_names, sep = "/")
  for (i in 1:length(file_paths)) {
    data_temp <- xlsx::read.xlsx(file_paths[i],startRow = 12, colIndex = c(2, 3, 4, 5, 10, 13, 14, 15), sheetIndex = 1)
    if (substr(file_names[i], 1, 2) =="TN") {
      data_temp <- rbind(data_temp,
                         xlsx::read.xlsx(file_paths[i],startRow = 12, colIndex = c(2, 3, 4, 5, 10, 13, 14, 15), sheetIndex = 2),
                         xlsx::read.xlsx(file_paths[i],startRow = 12, colIndex = c(2, 3, 4, 5, 10, 13, 14, 15), sheetIndex = 3))
    }
    data_temp <- filter(data_temp, !is.na(Episode.Type))
    summary_temp <- filter(data_temp, Included.in.reports....1.Yes..0.No. == 1) %>%
      group_by(.,Reporting.Period, Payer, Episode.Type) %>%
      summarize(., N.Valid.Episodes = sum(Num.of.Valid.Episodes, na.rm = T),
                Unadj.Mean.Spend = sum(Unadjusted.Mean.Spend.of.Valid..Episodes * Num.of.Valid.Episodes, na.rm = T) / sum(Num.of.Valid.Episodes, na.rm = T),
                NProviders = length(QBName)) 
    summary_temp[,1:3] <- lapply(summary_temp[,1:3], as.character)
    if (i==1){summary_table <- summary_temp} else {summary_table <- rbind(summary_table,summary_temp)}
  }
  summary_table
}

summary_table <- rbind(get_summary(MCO = "AG"),
                       get_summary(MCO = "UHC"))
summary_table <- as.data.frame(summary_table)

xlsx::write.xlsx(summary_table, file = "Summary.xlsx")