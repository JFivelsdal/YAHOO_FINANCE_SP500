

#Program Name: StockScript-S&P500Update
#Description: Creates a time stamped spreadsheet of S&P 500 stock data retrieved from Yahoo Finance
#Created by: Jon Fivelsdal


library(RCurl)

library(XML)

library(stringr)

library(RHTMLForms)

library(plyr)

library(reshape2)

library(xlsx)


sourceSAndP <- "http://en.wikipedia.org/wiki/List_of_S%26P_500_companies"

parseSAndP <- htmlParse(sourceSAndP)

#Table of S&P Companies

compTable <- readHTMLTable(parseSAndP,which = 1) #Retrieve a table of companies listed on the
#S&P 500 and there corresponding ticker symbol


compTable <- data.frame(compTable,stringsAsFactors = "FALSE")


newCompTable <- as.data.frame(as.matrix(compTable),stringsAsFactors = FALSE)

newCompTable[75,1] <- "BF-B" #Changes the stock symbol from the one present in the original table

newCompTable[64,1] <- "BRK-A" #Changes the stock symbol from the one present in the original table



#Yahoo Finance website

yFinanceSite <- "http://finance.yahoo.com/"

yForm <- getHTMLFormDescription(yFinanceSite)

getStockInfo <- createFunction(yForm[[1]]) #getStockInfo allows you to query for info on
#any paticular stock



#Retrieves information about a particular companies stock


######

#getStockProps: Retrieves stock information using the getStockInfo function
# and then cleans up the data
#####

getStockProps <- function(stock)
{
  
  stockTable <- readHTMLTable(htmlParse(getStockInfo(stock)))[2:7]#Tables 2-7 are relevant
  
  
  
  StockFrame <- rbind.fill(stockTable,stringsAsFactors = FALSE) #creates a dataframe
  
  
  
  StockFrame <- as.matrix(StockFrame)
  
  StockFrame[which(StockFrame == "N/A",arr.ind = TRUE)] <- ""
  
  
  validPos <- which(is.na(StockFrame) == FALSE,arr.ind = TRUE)
  
  StockFrame <- StockFrame[validPos]
  
  #Using gsub to get rid of extraneous spaces, newline characters and colons
  
  StockFrame <- gsub(pattern = "([\n])",replacement = "", x = StockFrame,fixed = TRUE)
  
  StockFrame <- gsub(pattern = "([:])",replacement = "", x = StockFrame )
  
  StockFrame <- gsub(pattern = "(\\s{2,})",replacement = "", x = StockFrame )
  
  
  #colNameList <- c("Open", "Bid","Ask","1y Target Est","Beta","Earnings Date",
                   #"Range- 52 Weeks","Volume","Avg Vol (3m)","Market Cap", "P/E (ttm)",
                   #"EPS (ttm)","Div & Yield","P/S (ttm)","P/S (ttm)","Ex-Dividend Date",
                   #"Quarterly","Mean Recommendation*","PEG Ratio (5 yr expected)")
  
  colNameList <- c("Open", "Bid","Ask","1y Target Est","Beta","Earnings Date",
    "52wk Range","Volume","Avg Vol (3m)","Market Cap", "P/E (ttm)",
    "EPS (ttm)","Div & Yield","P/S (ttm)","Ex-Dividend Date",
    "Quarterly","Mean Recommendation*","PEG Ratio (5 yr expected)")
  
  
  
  perPos <- grep(pattern = "%)", x = StockFrame, fixed = TRUE) + 1
  
  notAvPos <- grep(pattern = "N/A (N/A)", x = StockFrame, fixed = TRUE) + 1
  
  psPos <- grep(pattern = "P/S (ttm)", x = StockFrame, fixed = TRUE) - 1
  
  StockFrame[which(StockFrame == "N/A (N/A)",arr.ind = TRUE)] <- "NA"
  
  if(length(perPos) > 0 & length(psPos) > 0)
  {
    StockFrame <- StockFrame[-(perPos:psPos)]
  }
  
  if(length(notAvPos) > 0 & length(psPos) > 0)
  {
    StockFrame <- StockFrame[-(notAvPos:psPos)]
  }
  
  
  stockValPos <- grep(pattern = "^[0-9][^y]|[0-9]{3,}|(Est\\.)", x = StockFrame)
  
  blankPos <- which(StockFrame == "") #include blanks
  
  negNumPos <- grep(pattern = "^-[0-9]|NA", x = StockFrame) #include negative numbers
  
  numPos <- grep(pattern = "^[0-9]", x = StockFrame) #look for lines that begin with numbers
  
  stockValPos <- append(x = stockValPos, values = blankPos)
  
  stockValPos <- append(x = stockValPos, values = negNumPos)
  
  stockValPos <- append(x = stockValPos, values = numPos)
  
  stockValPos <- sort(stockValPos)
  
  stockValPos <- unique(stockValPos) #Make sure the positions are unique
  #so you don't get the same variable reported twice in the stock results
  
  wkPos <- grep(pattern = "52wk Range", x = StockFrame,fixed = TRUE)
  
  oneYearTarget <- grep(pattern = "1y Target Est", x = StockFrame,fixed = TRUE)
  
  
  stockValPos <- stockValPos[-(which(stockValPos == wkPos))] #Remove 52wk Range from the
  #data values vector
  
  
  stockValPos <- stockValPos[-(which(stockValPos == oneYearTarget))]#Remove 1y Target Est from the
  #data values vector
  stockVals <- StockFrame[stockValPos] #a vector of the numerical data
  
  colNames <- StockFrame[-stockValPos] #a vector of the column names
  
  StockFrame <- matrix(data = stockVals,nrow = 1,byrow = TRUE) #1 row matrix of numerical
  #stock data
  
  StockFrame <- as.data.frame(StockFrame,stringsAsFactors = FALSE)
  

  
  
  
  try(colnames(StockFrame) <- colNames) #set the column names of the 1 row data frame of
  #numerical stock data
  
  #c("Open", "Bid","Ask","1y Target Est","Beta","Earnings Date",
  #"52wk Range","Volume","Avg Vol (3m)","Market Cap", "P/E (ttm)",
  #"EPS (ttm)","Div & Yield","P/S (ttm)","Ex-Dividend Date",
  #"Quarterly","Mean Recommendation*","PEG Ratio (5 yr expected)")

  StockFrame$Open <- as.numeric(StockFrame$Open) 
  StockFrame[,4] <- as.numeric(StockFrame[,4])
  StockFrame$Beta <- as.numeric(StockFrame$Beta)
  StockFrame[,8] <- as.numeric(gsub(pattern = ",",x =  StockFrame[,8],replacement=""))
  StockFrame[,9] <- as.numeric(gsub(pattern = ",",x =  StockFrame[,9],replacement=""))
  StockFrame[,14] <- as.numeric(gsub(pattern = ",",x =  StockFrame[,14],replacement=""))
# Column 15 is the Ex Dividend date
  StockFrame[,16] <- as.numeric(gsub(pattern = ",",x =  StockFrame[,16],replacement=""))
 StockFrame[,17] <- as.numeric(gsub(pattern = ",",x =  StockFrame[,17],replacement=""))
  stockTime <- as.data.frame(format(Sys.time(), "%a %b %d %X %Y"))
  
  StockFrame <- cbind(stockTime,StockFrame)
  
  colnames(StockFrame)[1] <- "Time Retrieved"
  
  return(StockFrame)
  
}

###

#getSAndPProps: Checcks to see in a stock is listed on the S&P 500
#and then uses the function getStockProps to retrieve stock information


####

getSAndPProps <- function(stock)
{
  
  if(stock %in% newCompTable[,1])
  {
    getStockProps(stock)
    
  }
  else
    stop("Error: Invalid stock symbol")
}


require(parallel)
require(doSNOW)

cl<-makeCluster(detectCores(),type="SOCK")
registerDoSNOW(cl)

require(foreach)

system.time(
  storeSandP <- foreach(i = 1:nrow(newCompTable),.packages = c('XML','RHTMLForms','RCurl','plyr')) %dopar%
{
  
  try(getSAndPProps(newCompTable[i,1]))
  
  
}
)


largeStockDataFrame <- ldply(storeSandP,data.frame) #Create a dataframe of the 500 stocks

newlargeStockDataFrame <- largeStockDataFrame[,-c(17:ncol(largeStockDataFrame))] #Get rid of columns with NA


largeCompanyFrame <- as.data.frame(newCompTable[,2])

newlargeStockDataFrame <- cbind(largeCompanyFrame,newlargeStockDataFrame)


colnames(newlargeStockDataFrame)[1] <- "Company Name"

colnames(newlargeStockDataFrame)[6] <- "Estimated One Year Target"

colnames(newlargeStockDataFrame)[9] <- "Range - 52 Weeks"

colnames(newlargeStockDataFrame)[11] <- "Average Volume (3m)"

colnames(newlargeStockDataFrame)[12] <- "Market Cap"

colnames(newlargeStockDataFrame)[13] <- "P/E (ttm)"

colnames(newlargeStockDataFrame)[14] <- "EPS (ttm)"

colnames(newlargeStockDataFrame)[15] <- "Dividend and Yield"

colnames(newlargeStockDataFrame)[16] <- "P/S (ttm)"

newlargeStockDataFrame[,13] <- as.numeric(gsub(pattern = ",",x =  newlargeStockDataFrame[,13],replacement=""))

newlargeStockDataFrame[,14] <- as.numeric(gsub(pattern = ",",x =  newlargeStockDataFrame[,14],replacement=""))

newlargeStockDataFrame[,16] <- as.numeric(gsub(pattern = ",",x =  newlargeStockDataFrame[,16],replacement=""))

colnames(newlargeStockDataFrame) <- gsub(pattern = ".",x = colnames(newlargeStockDataFrame),replacement = " ",fixed = TRUE)


currentTime <- format(Sys.time(), "%a %b %d %X %Y")


printTime <- gsub(pattern = "\\s",x = currentTime,replacement = "-")

printTime <- gsub(pattern = ":", x = currentTime,replacement = "+")


write.xlsx(x = newlargeStockDataFrame, file = paste0("S&P500StockDataEdit",".xlsx"))


#Commented out. Gives a unique time stamp to the Excel file
#write.xlsx(x = newlargeStockDataFrame, file = paste0("S&P500StockDataEdit",printTime,".xlsx"))


