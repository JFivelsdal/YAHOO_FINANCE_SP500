
#Program Name: StockScript-S&P500
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

newCompTable[75,1] <- "BF-B"  #Changes the stock symbol from the one present in the original table

newCompTable[64,1] <- "BRK-A" #Changes the stock symbol from the one present in the original table



#Yahoo Finance website

yFinanceSite <- "http://finance.yahoo.com/"

yForm <- getHTMLFormDescription(yFinanceSite)

getStockInfo <- createFunction(yForm[[1]]) #getStockInfo allows you to query for info on 
                                           #any paticular stock



#Retrieves information about a particular companies stock


######

#getStockProps: Retrieves stock information using the getStockInfo function 
#               and then cleans up the data
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
 
 StockFrame  <- gsub(pattern = "([\n])",replacement = "", x = StockFrame,fixed = TRUE)
 
 StockFrame  <- gsub(pattern = "([:])",replacement = "", x = StockFrame )
 
 StockFrame  <- gsub(pattern = "(\\s{2,})",replacement = "", x = StockFrame )
 
 
 colNameList <- c("Open", "Bid","Ask","1y Target Est","Beta","Earnings Date",
                  "52wk Range","Volume","Avg Vol (3m)","Market Cap", "P/E (ttm)",
                  "EPS (ttm)","Div & Yield","P/S (ttm)","P/S (ttm)","Ex-Dividend Date",
                  "Quarterly","Mean Recommendation*","PEG Ratio (5 yr expected)")
 
 
 
 perPos <- grep(pattern = "%)", x = StockFrame, fixed = TRUE) + 1
 
notAvPos <-  grep(pattern = "N/A (N/A)", x = StockFrame, fixed = TRUE) + 1
 
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

 numPos <-  grep(pattern = "^[0-9]", x = StockFrame) #look for lines that begin with numbers

 stockValPos <- append(x = stockValPos, values = blankPos)

 stockValPos <- append(x = stockValPos, values = negNumPos)

 stockValPos <- append(x = stockValPos, values = numPos)

 stockValPos <- sort(stockValPos)

 stockValPos <- unique(stockValPos) #Make sure the positions are unique 
                                    #so you don't get the same variable reported twice
                                    #in the stock results
 
 wkPos <- grep(pattern = "52wk Range", x = StockFrame,fixed = TRUE)

 oneYearTarget <- grep(pattern = "1y Target Est", x = StockFrame,fixed = TRUE)
 

stockValPos <- stockValPos[-(which(stockValPos == wkPos))] #Remove 52wk Range from the 
                                                           #data values vector
                                                          

stockValPos <- stockValPos[-(which(stockValPos == oneYearTarget))]#Remove 1y Target Est from the 
                                                                  #data values vector
stockVals <- StockFrame[stockValPos] #a vector of the numerical data
 
colNames <- StockFrame[-stockValPos] #a vector of the column names

StockFrame <- matrix(data =   stockVals,nrow = 1,byrow = TRUE) #1 row matrix of numerical
                                                               #stock data

StockFrame <- as.data.frame(StockFrame,stringsAsFactors = FALSE)
 
try(colnames(StockFrame) <- colNames) #set the column names of the 1 row data frame of 
                                      #numerical stock data

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

colnames(newlargeStockDataFrame)[5] <- "Estimated One Year Target"

colnames(newlargeStockDataFrame)[8] <- "Range - 52 Weeks"

colnames(newlargeStockDataFrame)[10] <- "Average Volume"

colnames(newlargeStockDataFrame)[11] <- "Market Cap"

colnames(newlargeStockDataFrame)[12] <- "P/E (ttm)"

colnames(newlargeStockDataFrame)[13] <- "EPS (ttm)"

colnames(newlargeStockDataFrame)[14] <- "Dividend and Yield"

colnames(newlargeStockDataFrame)[15] <- "P/S (ttm)"


colnames(newlargeStockDataFrame) <- gsub(pattern = ".",x = colnames(newlargeStockDataFrame),replacement = " ",fixed = TRUE)


currentTime <- format(Sys.time(), "%a %b %d %X %Y")

newlargeStockDataFrame <- cbind(currentTime,newlargeStockDataFrame)

colnames(newlargeStockDataFrame)[1] <- "Time Retrieved"

write.xlsx(x = newlargeStockDataFrame, file = paste0("S&P500StockData",".xlsx"))

