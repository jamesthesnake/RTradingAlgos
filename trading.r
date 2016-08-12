
library(xlsx);library(zoo);library(quantmod);library(rowr);library(TTR);
 
 
 
 
<strong># Set the parameters here </strong>
 
pc_up_level = 6  # Set the high percentage threshold level
 
noDays = 2       # No of lookback days
 
 
 
 
<strong># Read the csv file that contains the Stock tickers to be scanned</strong>
 
test_tickers = read.csv("F&amp;O Stock List.csv")
 
symbol = as.character(test_tickers[,4])          
 
leverage = as.character(test_tickers[,5])        
 
 
 
 
sig_finds = data.frame(0,0,0,0,0,0)
 
colnames(sig_finds) = c("Ticker","Leverage","Previous Price","Current Price","Abs Change","% Change")
for(s in 1:length(symbol)){
 
 print(s) 
 
 # Custom function to download stock data from Google finance
 source("Stock price data.R") 
 intraday_price_data(symbol[s],noDays)
 
 dirPath = paste(getwd(),"/",sep="") # Specify the R directory 
 fileName = paste(dirPath,symbol[s],".csv",sep="") # Specify the filename 
 data = as.data.frame(read.csv(fileName)) # Read the downloaded file
 
 data_close =subset(data, select=(CLOSE), subset=(TIME == 1530))
 pdp = tail(data_close,1)[,1]
 lp = tail(data,1)[,3]
 
 # Compute absolute and the percentage change
 apc = round((lp - pdp),2)
 pc = round(((lp - pdp) / pdp) * 100,2)
 
 if (pc &gt; pc_up_level ){
 
 # Create a temporary empty data frame, only if the Percentage change is significant
 sig_finds_temp = data.frame(0,0,0,0,0,0) 
 colnames(sig_finds_temp) = c("Ticker","Leverage","Previous Price","Current Price","Abs Change","% Change")
 
 sig_finds_temp = c(symbol[s],leverage[s],pdp,lp,apc,pc) 
 sig_finds = rbind(sig_finds, sig_finds_temp) # combines the previous and current data frame
 
 rm(sig_finds_temp) # deletes the temporary data frame 
 
 print(sig_finds)
 
 }
 
 # Deletes the downloaded stock price files for each ticker from the directory
 unlink(fileName) 
 
} 
 
# Write the results in an excel sheet
write.xlsx(sig_finds,"Shorting at High.xlsx")

# Pull NIFTY data for the last 1 year
source("Stock price data.R") 
daily_price_data("NIFTY",1) 
dirPath = paste(getwd(),"/",sep="") 
fileName = paste(dirPath,"NIFTY",".csv",sep="")
nifty_data = as.data.frame(read.csv(file = fileName))
 
# Read the stock symbols from "Shorting at High.xlsx" workbook
test_tickers = read.xlsx("Shorting at High.xlsx",header=TRUE, 1, startRow=1, as.data.frame=TRUE)
t = nrow(test_tickers)
 
test_tickers$Count = 0; test_tickers$Avg_high = 0;
test_tickers$Avg_decline = 0; test_tickers$Next_3d_Change = 0;
test_tickers$Low_1Yr = 0; test_tickers$High_1Yr = 0;
test_tickers$Neg_Last15_days = 0 ; test_tickers$RSI = 0;
test_tickers$NIFTY_Correlation = 0
 
symbol = as.character(test_tickers[2:t,2])
noDays = 1
 
# Compute the metrics for each stock 
for (s in 1:length(symbol)){
 
 print(s)
 
 table_sig = data.frame(0,0,0,0,0,0,0) 
 colnames(table_sig) = c("Date","Days High in %","Days Close in %","RSI","Entry Price","Exit Price","Profit/Loss")
 
 source("Stock price data.R")
 daily_price_data(symbol[s],noDays) 
 dirPath = paste(getwd(),"/",sep="") 
 fileName = paste(dirPath,symbol[s],".csv",sep="") 
 data = as.data.frame(read.csv(file = fileName)) 
 N = nrow(data)
 
 # Initializing variables to zero
 nc = 0 ; Cum_hc_diff = 0;
 Cum_threshold_per = 0;Cum_next_3d_change = 0; 
 
 diff = data$HIGH - data$CLOSE
 
 rsi_data = RSI(data$CLOSE, n = 14, maType="WMA", wts=data[,"VOLUME"])
 
 for (i in 2:N) {
 
 days_close = (data$CLOSE[i] - data$CLOSE[i-1])*100/data$CLOSE[i-1]
 days_high = (data$HIGH[i] - data$CLOSE[i-1])*100/data$CLOSE[i-1]
 
 condition = ((days_high &gt; pc_up_level) == TRUE)
 if (condition)
 {
 nc = nc + 1 
 
 date_table = data$DATE[i]
 
 high_per_table = round(days_high,2)
 close_price_per_table = format(round(days_close,2),nsmall = 2)
 
 rsi_table = format(round(rsi_data[i],2),nsmall = 2)
 
 entry_price = format(round(data$CLOSE[i-1]*(1+(pc_up_level/100)),2),nsmall = 2)
 exit_price = format(round(data$CLOSE[i],2),nsmall = 2)
 pl_trade = format(round(((data$CLOSE[i-1]*(1+(pc_up_level/100))) - data$CLOSE[i]),2),nsmall = 2)
 
 hc_diff = diff[i] # difference between day's high and day's close whenever the price crossed the threshold level.
 Cum_hc_diff = Cum_hc_diff + hc_diff # Cumulative of the difference for all trades
 Average_hc_diff = round(Cum_hc_diff / nc , 2) # Computing the average for difference between high and low.
 
 # Computes the total of Day's high's whenever price crossed the threshold.
 Cum_threshold_per = Cum_threshold_per + days_high 
 # Average of High's
 Average_threshold_per = round(Cum_threshold_per / nc , 2) 
 
 if ( i &lt; (N - 5)){ # Captures the average next 3-day price movement
 next_3d_change = data$CLOSE[i+3] - data$CLOSE[i]
 Cum_next_3d_change = Cum_next_3d_change + (data$CLOSE[i+3] - data$CLOSE[i])
 Average_next_3d_change = round(Cum_next_3d_change / nc,2)
 }
 
 # Create a temporary dataframe to add values and then later merge with the final table 
 table_temp = data.frame(0,0,0,0,0,0,0) 
 colnames(table_temp) = c("Date","Days High in %","Days Close in %","RSI","Entry Price","Exit Price","Profit/Loss")
 table_temp = c(date_table,high_per_table,close_price_per_table,rsi_table,entry_price,exit_price,pl_trade) 
 table_sig = rbind(table_sig, table_temp) 
 rm(table_temp)
 
 } 
 
 
 }
 
 table_sig = table_sig[order(table_sig$Date),]
 
 # Write the individual results in an excel sheet
 name = paste(symbol[s]," past performance",".xlsx",sep="")
 write.xlsx(table_sig,name) 
 
 # load it back to format the excel sheet for column width
 wb = loadWorkbook(name)
 sheets = getSheets(wb)
 autoSizeColumn(sheets[[1]], colIndex=1:16)
 saveWorkbook(wb,name)
 
 # Price performance in the last 15 days, checks for many days the price change was negative. 
 nc_last_15 = 0 
 for ( p in (N-15):N){
 
 condition_1 = (((data$CLOSE[p] - data$CLOSE[p-1])&lt; 0 ) == TRUE )
 if (condition_1)
 {
 nc_last_15 = nc_last_15 + 1
 }
 } 
 
 # Correlation between the ticker and NIFTY 
 Cor_coeff = round(cor(data$CLOSE , nifty_data$CLOSE),2)
 
 # Filling the “Shorting at High.xlsx” with the computed metrics
 
 e = s + 1
 test_tickers$Count[e] = nc
 test_tickers$Avg_high[e] = format(Average_threshold_per,nsmall = 2) # Indicates the Avg. percentage above the threshold in the past
 test_tickers$Avg_decline[e] = format(Average_hc_diff,nsmall = 2) # Indicates the Avg. decline in Absolute rupee terms from the Avg. high to the closing price in the past test_tickers$Next_3d_Change[e] = round(Average_next_3d_change,2) # Indicates the next 3 day price movement.
 test_tickers$Next_3d_Change[e] = format(Average_next_3d_change,nsmall = 2)
 test_tickers$Low_1Yr[e] = format(min(data$CLOSE),nsmall = 2) # 1 year low price
 test_tickers$High_1Yr[e] = format(max(data$CLOSE),nsmall = 2) # 1 year high price
 test_tickers$Neg_Last15_days[e] = nc_last_15 # Indicates how many times in the last 15 days has the price been negative.
 test_tickers$RSI[e] = round(tail(rsi_data,1),2)
 test_tickers$NIFTY_Correlation[e] = Cor_coeff # Computes the correlation between the Stock and NIFTY.
 
 unlink(fileName) # Deletes the downloaded stock price files for each ticker after processing the data
 
 rm("data","table_sig")
 
} # This closes the for loop 
 
test_tickers = test_tickers[-1,-1]
colnames(test_tickers) = c("Ticker", "Leverage","Previous Price","Current Price","Abs. Change","% Change","Count","Avg.High %","Abs Avg.Decline","Next 3-d price move","1-yr low","1-yr high","-Ve last 15-days","RSI","Nifty correlation")
 
# Write and format the final output file
 
write.xlsx(test_tickers,"Shorting at High.xlsx") 
 
wb = loadWorkbook("Shorting at High.xlsx")
sheets = getSheets(wb)
autoSizeColumn(sheets[[1]], colIndex=1:19)
saveWorkbook(wb,"Shorting at High.xlsx")
