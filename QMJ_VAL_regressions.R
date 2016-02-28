## loop through all files in directory
# http://www.r-bloggers.com/looping-through-files/

path = "S:/Neptune Investment Team/Analyst_Patrick/Dev/Projects/QMJ/OutputFilesAVG"
setwd(path)

out.file <- ""
QUAL_file.names <- dir(path = path, pattern = "QUAL.csv")
VAL_file.names <- dir(path = path, ppattern = "VAL.csv")
RAY <- data.frame(read.csv("RAY_Index.csv"))
RAY$Date <- as.factor(RAY$Date)

# cross sectional regresions
res <- data.frame()
val_ratios <- c("PE_RATIO","CURR_EV_TO_T12M_EBITDA_C","PX_TO_BOOK_RATIO")
for(i in 1:length(QUAL_file.names)){
  fileq <- read.csv(QUAL_file.names[i], stringsAsFactors = FALSE)
  filev <- read.csv(VAL_file.names[i], stringsAsFactors = FALSE)
  # exclude values outside of 5 std devs
  fileq <- fileq[which(fileq$Z_QUAL < 5 & fileq$Z_QUAL > -5), ]
  # merge quality and valuation data sets
  fileq <- merge(fileq, filev, by = "Ticker")
  fileq <- fileq[!is.na(fileq$PE_RATIO), ]
  fileq <- fileq[!is.na(fileq$CURR_EV_TO_T12M_EBITDA_C), ]
  fileq <- fileq[!is.na(fileq$PX_TO_BOOK_RATIO), ]
  fileq <- fileq[!is.na(fileq$Z_QUAL), ]
  # regression
  reg_vector <- data.frame(substr(QUAL_file.names[i],1,8))
  for(j in 1:length(val_ratios)){
    reg_formula <- paste(val_ratios[j]," ~ Z_QUAL", sep = "")
    reg <- lm(as.formula(reg_formula), data = fileq)
    # bind results to result array
    reg_vector <- cbind(reg_vector, data.frame(summary(reg)$coef[1,1], summary(reg)$coef[2,1],
                                       summary(reg)$r.squared))
    colnames(reg_vector)[(j*3-1):((j*3-1)+2)] <- c(paste(val_ratios[j],"$Intercept", sep = ""),
                                           paste(val_ratios[j],"$coef(Z_QUAL)", sep = ""),
                                           paste(val_ratios[j],"$R-squared",sep=""))
  }
  res <- rbind(res, reg_vector)
}
res <- cbind(res, reg_array)
colnames(res)[1] <- "DATE" 



