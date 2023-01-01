###############Enviroment clearing and installing/loading packages and function declearing###############
rm(list = ls())

#install.packages("ggplot2")
#install.packages("dplyr")
#install.packages("lubridate")
#install.packages("xts")
#install.packages("data.table")
#install.packages("friedland") # https://github.com/efriedland/friedland
#install.packages("tibble")
#install.packages("olsrr")
#install.packages("tseries")
#install.packages("writexl")

library(ggplot2)     #plotting
library(dplyr)       #functional programming
library(lubridate)   #data frame function package
library(xts)         #for time series objects
library(data.table)  #for table
library(friedland)   #for unitroot tests and dynamic ols calculating cay variable
library(tibble)      #for tables
library(olsrr)       #stepwise backward regression
library(tseries)     #basic time series functions
library(writexl)     #needed to make excel files


#monthly calculations
#Important remarks: Some variables are only reported in the middle of the month in datastream
#Thus, I treat them the same as if they would have been reported at the end of the month to avoid look ahead bias. E.g. data from 15.01 is used to predict returns in februrary

###############Read in Data from csv###############
#dependent Variable. I am forecasting from January 1990 to September 2020 -> consistent with quarterly
dependentVariable <- read.csv2(file="depVariable.csv", header=TRUE, sep=',')
dependentVariable[,-1] <- lapply(dependentVariable[,-1], function(x) {as.numeric(gsub(",", ".", x))})
colnames(dependentVariable)[1] <- "Date"
dependentVariable$Date = as.Date(dependentVariable$Date, format = "%d/%m/%Y")
dependentVariable <- as.xts(dependentVariable[,-1], order.by=dependentVariable[,1])


#this contains all data from DS, where I extracted monthly data
monthlyIndex <- read.csv2(file="Monthly.csv", header=TRUE, sep=',')
monthlyIndex[,-1] <- lapply(monthlyIndex[,-1], function(x) {as.numeric(gsub(",", ".", x))})
colnames(monthlyIndex)[1] <- "Date"
monthlyIndex$Date = as.Date(monthlyIndex$Date, format = "%d/%m/%Y")


#This contains all data from DS, where I need one date prior to my regression, as some variables are needed in changes from t to t-1
changeVariables <- read.csv2(file="changeVariable.csv", header=TRUE, sep=',')
changeVariables[,-1] <- lapply(changeVariables[,-1], function(x) {as.numeric(gsub(",", ".", x))})
colnames(changeVariables)[1] <- "Date"
changeVariables$Date = as.Date(changeVariables$Date, format = "%d/%m/%Y")
changeVariables <- as.xts(changeVariables[,-1], order.by=changeVariables[,1])


#This file has the data needed for the smoothed 10 years earnings to price ratio -> 10 years earlier
smoothedEP <- tail(read.csv2("smoothedPE.csv", header= TRUE, sep=','),-1)
smoothedEP[,-1] <- lapply(smoothedEP[,-1], function(x) {as.numeric(gsub(",", ".", x))})
smoothedEP$Name = as.Date(smoothedEP$Name, format = "%d/%m/%Y")


#This file has the data needed for the variance. Seperate file, as we need daily data and one period prior to the regression
stockVariance <- read.csv2(file="stockVar.csv", header=TRUE, sep=',')
stockVariance[2] <- lapply(stockVariance[2], function(x) {as.numeric(gsub(",", ".", x))})
colnames(stockVariance)[1] <- "Date"
stockVariance$Date = as.Date(stockVariance$Date, format = "%d/%m/%Y")
stockVariance <- as.xts(stockVariance[,-1], order.by=stockVariance[,1])
colnames(stockVariance) <- "svr"


#net equity expansion, different file because also different t needed
netIssue <- read.csv2(file="netEquExp.csv", header=TRUE, sep=',')
netIssue[,-1] <- lapply(netIssue[,-1], function(x) {as.numeric(gsub(",", ".", x))})
colnames(netIssue)[1] <- "Date"
netIssue$Date = as.Date(netIssue$Date, format = "%d/%m/%Y")
netIssue <- as.xts(netIssue[,-1], order.by=netIssue[,1])


#data from the chair
lehrstuhlIndex <- read.csv2(file="ICC_data_1990_2020_final.csv", header=TRUE, sep=',')



###############Data transformations###############
#IMPORTANT: All data in DS, which are given in percentages, are transformed into decimals
#Furthermore, as we should follow the data manipulation of Li et al. (2013) I do not take the logarithm of D/Y, D/P, E/P, E10/P, even though it is done often in other literature
Date <- index(dependentVariable)

#calculation of not continously compounded excess market return for summary statistics -> following Li
notCompVwretd <- (dependentVariable$JAPAN.DS.Market...TOT.RETURN.IND-stats::lag(dependentVariable$JAPAN.DS.Market...TOT.RETURN.IND))/
  stats::lag(dependentVariable$JAPAN.DS.Market...TOT.RETURN.IND) -
  (dependentVariable$JP.OVERNIGHT.UNCOLLATERISED.CALL.MONEY.RATE..AVG...NADJ/100)/12 #divide by 100, because it is given in % in DS and by 12, because it is annnualized in DS
  

#calculation of excess market return following Li et al (2013) for regression: log(1+ret%) - log (1+tbill%)
dependentVariable$JAPAN.DS.Market...TOT.RETURN.IND <- log((dependentVariable$JAPAN.DS.Market...TOT.RETURN.IND/stats::lag(dependentVariable$JAPAN.DS.Market...TOT.RETURN.IND)))
#divide tbill by 100 because given % in DS and divide by 12, because tbill is annualized in DS
dependentVariable$JP.OVERNIGHT.UNCOLLATERISED.CALL.MONEY.RATE..AVG...NADJ <- log(1 + (dependentVariable$JP.OVERNIGHT.UNCOLLATERISED.CALL.MONEY.RATE..AVG...NADJ/100)/12)
dependentVariable$JAPAN.DS.Market...TOT.RETURN.IND <- dependentVariable$JAPAN.DS.Market...TOT.RETURN.IND - dependentVariable$JP.OVERNIGHT.UNCOLLATERISED.CALL.MONEY.RATE..AVG...NADJ
dependentVariable <- data.frame(date = Date, coredata(dependentVariable$JAPAN.DS.Market...TOT.RETURN.IND))
names(dependentVariable)[1] <- "date"
names(dependentVariable)[2] <- "totalReturnChange"
dependentVariable <- tail(dependentVariable, -1) #ommit first value -> no look ahead bias

#get values such as price, dividends, earnings etc. to calculate valuation ratios such as the dividends-earnings-ratio
#monthlyIndex <- data.frame(date=index(monthlyIndex), coredata(monthlyIndex))
dividends <- monthlyIndex$JAPAN.DS.Market...PRICE.INDEX/monthlyIndex$JAPAN.DS.Market...TOT.RETURN.IND
price <- dividends/monthlyIndex$JAPAN.DS.Market...DIVIDEND.YIELD*100 #multiply by 100, because DY is given in % in DS

#here we take data from the changeVariables table, as we need the lag of 1 period to construct the dividendYield according to Welch/Goyal (2003)
dividendsY <- changeVariables$JAPAN.DS.Market...PRICE.INDEX/changeVariables$JAPAN.DS.Market...TOT.RETURN.IND
priceY <- dividendsY/changeVariables$JAPAN.DS.Market...DIVIDEND.YIELD*100 #multiply by 100, because DY is given in % in DS
earnings <- price/monthlyIndex$JAPAN.DS.Market...PER
#using head because data is one month to much, as I take the lag
dividendsY <- head(dividendsY, -1)
priceY <- head(priceY, -1)

#calculating independent variables - I take the log following the sources specified in my thesis
dividendPriceRatio <- (monthlyIndex$JAPAN.DS.Market...DIVIDEND.YIELD/100) #divide by 100, DY is given in % in DS
dividendYield <- drop(coredata(dividendsY/priceY))
dividendPayoutRatio <- log(dividends/earnings) #this in log following literature
bookToMarketRatio <- 1/monthlyIndex$JAPAN.DS.Market...PRICE.BOOK.RATIO 
earningsToPriceRatio <- 1/monthlyIndex$JAPAN.DS.Market...PER

#sentiment variables
volatilityIndex <- log(1+diff(changeVariables$NIKKEI.STOCK.AVERAGE.VOLATILITY.INDEX...PRICE.INDEX
     , lag = 1, difference = 1)/stats::lag(changeVariables$NIKKEI.STOCK.AVERAGE.VOLATILITY.INDEX...PRICE.INDEX))
volatilityIndex <- tail(drop(coredata(volatilityIndex)),-1)
stockVariance <- tail((diff(stockVariance$svr, lag = 1, difference = 1)/stats::lag(stockVariance$svr)),-1)
stockVariance <- apply.monthly(stockVariance, var)
stockVariance <- drop (coredata(stockVariance$svr))

#calculation of smoothed PE following Robert Shiller's website and Aono/Iwaisako (2011): PE = Price t / [AVG: Earnings t-120-1] - I take reverse - Earnings to price ratio
#first get prices und earnings individually
sDividends <- smoothedEP$JAPAN.DS.Market...PRICE.INDEX/smoothedEP$JAPAN.DS.Market...TOT.RETURN.IND
sPrice <- sDividends/smoothedEP$JAPAN.DS.Market...DIVIDEND.YIELD*100 #multiply by 100, because DY is given in % in DS
sEarnings <- sPrice/smoothedEP$JAPAN.DS.Market...PER
sEarnings <- rollmean(sEarnings, k = 119, fill = NA, align = 'left')
smoothedEP <- head(sEarnings, -119)
smoothedEP <- 1/price/smoothedEP #*12 in order to annualize

#Further independent variables. Names should be self explanatory + see description if thesis if something is not clear
icc <- as.numeric(lehrstuhlIndex[2591:2960, 5])
longTermYield <- as.numeric(lehrstuhlIndex[2591:2960, 4])
tBill <- monthlyIndex$JP.OVERNIGHT.UNCOLLATERISED.CALL.MONEY.RATE..AVG...NADJ/100 #divided by 100, because DY is given in % in DS
icc <- icc - tBill #we use excess return
termSpread <- longTermYield - tBill
percentEquityIssuing <- monthlyIndex$JP.STOCKS..PUBLIC.OFFERINGS...AMOUNT.RAISED.CURN/
  (monthlyIndex$JP.ISSUES..CORPORATE.STRAIGHT.BONDS.CURN+monthlyIndex$JP.STOCKS..PUBLIC.OFFERINGS...AMOUNT.RAISED.CURN)
percentEquityIssuing[is.nan(percentEquityIssuing)] <- 0 #I manually set 01.1994 to zero, as both data yielded 0 -> dividing by 0 not possible
percentEquityIssuing <- percentEquityIssuing
inflationRate <- monthlyIndex$JP.CPI..NATIONAL.MEASURE...ANNUAL.INFLATION.RATE.NADJ/100 #divided by 100, because DY is given in % in DS


#independent variables, where i need to calculate the changes. e.g. t to t-1
realMoneySupply <- diff(changeVariables$JP.MONEY.SUPPLY..M2..METHO.BREAK..APR..2003..CURA
                        , lag = 1, difference = 1)/stats::lag(changeVariables$JP.MONEY.SUPPLY..M2..METHO.BREAK..APR..2003..CURA)
realMoneySupply <- tail(drop(coredata(realMoneySupply)),-1)
petroleumConsumption <- diff(changeVariables$JP.PETROLEUM..CONSUMPTION.VOLN
                        , lag = 1, difference = 1)/stats::lag(changeVariables$JP.PETROLEUM..CONSUMPTION.VOLN)
petroleumConsumption <- tail(drop(coredata(petroleumConsumption)),-1)
unemploymentRate <- diff(changeVariables$JP.UNEMPLOYMENT.RATE..METHO.BREAK.OCT.2010..SADJ
                         , lag = 1, difference = 1)/stats::lag(changeVariables$JP.UNEMPLOYMENT.RATE..METHO.BREAK.OCT.2010..SADJ)
unemploymentRate <- tail(drop(coredata(unemploymentRate)),-1)
industrialProduction <- diff(changeVariables$JP.INDUSTRIAL.PRODUCTION...MINING...MANUFACTURING.VOLA
                             , lag = 1, difference = 1)/stats::lag(changeVariables$JP.INDUSTRIAL.PRODUCTION...MINING...MANUFACTURING.VOLA)
industrialProduction <- tail(drop(coredata(industrialProduction)),-1)
crudeOilPriceChange <- diff(changeVariables$US.REFINERS.ACQUISITION.COST.OF.DOM....IMPORTED.CRUDE.OIL.CURN, lag = 1, 
                      difference = 1)/stats::lag(changeVariables$US.REFINERS.ACQUISITION.COST.OF.DOM....IMPORTED.CRUDE.OIL.CURN)
crudeOilPriceChange <- drop(coredata(crudeOilPriceChange))
crudeOilPriceChange <- tail(crudeOilPriceChange,-1)
crudeOilProduction <- diff(changeVariables$WD.CRUDE.OIL.PRODUCTION...WORLD.VOLN, lag = 1, 
                           difference = 1)/stats::lag(changeVariables$WD.CRUDE.OIL.PRODUCTION...WORLD.VOLN)
crudeOilProduction <- drop(coredata(crudeOilProduction))
crudeOilProduction <- tail(crudeOilProduction, -1)


#net equity expansion
mv <- tail(netIssue$JAPAN.DS.Market...MARKET.VALUE, -14)
temp9 <- diff(netIssue$JAPAN.DS.Market...PRICE.INDEX, lag = 1, differences = 1)/stats::lag(netIssue$JAPAN.DS.Market...PRICE.INDEX)
netIssue <- netIssue$JAPAN.DS.Market...MARKET.VALUE-stats::lag(netIssue$JAPAN.DS.Market...MARKET.VALUE)*(1+temp9)
netIssue <- head(tail(rollsum(netIssue, k = 12, fill = NA, align = 'left'), -3),-11)
netIssue <- drop(coredata(netIssue$JAPAN.DS.Market...MARKET.VALUE))/drop(coredata(mv$JAPAN.DS.Market...MARKET.VALUE)) 

#netPayoutYield
netPY <- c(NA, monthlyIndex$JAPAN.DS.Market...DIVIDEND.YIELD/100)+((( #weil dy in % angegeben ist
  stats::lag(changeVariables$JAPAN.DS.Market...MARKET.VALUE)*((changeVariables$JAPAN.DS.Market...PRICE.INDEX/stats::lag(changeVariables$JAPAN.DS.Market...PRICE.INDEX))))-
    changeVariables$JAPAN.DS.Market...MARKET.VALUE)/changeVariables$JAPAN.DS.Market...MARKET.VALUE)
netPY <- log(0.1 + tail(netPY, -1))


###############Finalindex construction###############
Date <- tail(Date, -1) #putting the dependent variable at time t in the same row with independent variable at time t-1 -> easier regressions
#Calculate with K = 1, 12, 24, 36, 48 forecast horizon
mk12 <- rollapply(dependentVariable$totalReturnChange, 12, mean, align = "left", fill = NA)
mk12 <- data.frame(date = Date, coredata(mk12))
colnames(mk12)[2] <- "totalReturnChange"
mk24 <- rollapply(dependentVariable$totalReturnChange, 24, mean, align = "left", fill = NA)
mk24 <- data.frame(date = Date, coredata(mk24))
colnames(mk24)[2] <- "totalReturnChange"
mk36 <- rollapply(dependentVariable$totalReturnChange, 36, mean, align = "left", fill = NA)
mk36 <- data.frame(date = Date, coredata(mk36))
colnames(mk36)[2] <- "totalReturnChange"
mk48 <- rollapply(dependentVariable$totalReturnChange, 48, mean, align = "left", fill = NA)
mk48 <- data.frame(date = Date, coredata(mk48))
colnames(mk48)[2] <- "totalReturnChange"

#only independent variables
finalmonthlyOnlyIV <- data.frame(monthlyIndex$Date, icc, dividendYield, dividendPriceRatio, dividendPayoutRatio, bookToMarketRatio, earningsToPriceRatio, smoothedEP,
                                percentEquityIssuing, netIssue, netPY, tBill, longTermYield, termSpread, inflationRate, realMoneySupply, #exchangeRate, 
                                crudeOilPriceChange, crudeOilProduction, petroleumConsumption, industrialProduction, unemploymentRate, stockVariance, volatilityIndex)
colnames(finalmonthlyOnlyIV) <- c("date", "Icc", "D/Y", "D/P", "D/E", "B/M", "E/P", "E10/P", "Equis", "Ntis", "Ndy","Tbill", "Lty", "Ts", "Ifl", "Rms",
                                  "C/P", "C/O", "Pcon", "Iprod", "Unem", "Svr", "Vola")
finalmonthlyOnlyIV <- head(finalmonthlyOnlyIV, -1) #omit last row, because i only forecast to september -> Independent variables data only till august needed

#using finalIndexMonthlySS for Summary statistic purposes
finalIndexMonthlySS <- bind_cols(finalmonthlyOnlyIV[,-1], dependentVariable)
temp <- finalIndexMonthlySS
finalIndexMonthlySS$date <- NULL

#final Index
finalIndexMonthly <- finalIndexMonthlySS
finalIndexMonthly <- add_column(finalIndexMonthly, mk12$totalReturnChange)
finalIndexMonthly <- add_column(finalIndexMonthly, mk24$totalReturnChange)
finalIndexMonthly <- add_column(finalIndexMonthly, mk36$totalReturnChange)
finalIndexMonthly <- add_column(finalIndexMonthly, mk48$totalReturnChange)
colnames(finalIndexMonthly)[(ncol(finalIndexMonthly)-4):ncol(finalIndexMonthly)] <- c("K1", "K12", "K24", "K36", "K48")
finalIndexMonthlySS <- temp



###############Unit root tests###############
#I set the max lag to 12, because this is the usual maximum lag considered for monthly data -> then we dynamically pick the best lag using the
#AIC value to find the best corresponding lag
#For better description see unit root table in appendix of thesis
#I(1) = no stationary ------ I(0) = stationary
unitRootTestMonthly <- as.ts(finalIndexMonthlySS)
unitRootTestMonthlyTrend <- UnitRoot(unitRootTestMonthly, 12, drift = T, trend = T)
unitRootTestMonthly <- UnitRoot(unitRootTestMonthly, 12, drift = F, trend = F)
unitRootTestMonthly <- data.frame(unitRootTestMonthly)
unitRootTestMonthly <- unitRootTestMonthly[c("ADFstatistic", "significance", "result")]
unitRootTestMonthlyTrend <- data.frame(unitRootTestMonthlyTrend)
unitRootTestMonthlyTrend <- unitRootTestMonthlyTrend[c("ADFstatistic", "significance", "result")]
unitRootTestMonthly <- cbind(unitRootTestMonthly, unitRootTestMonthlyTrend)



###############Summary statistics###############
#I report market excess return not in compounded following Li -> here replace the value in finalIndexMonthlySS
notCompVwretd <- tail(coredata(notCompVwretd),-1)
finalIndexMonthlySS$totalReturnChange <- notCompVwretd

#following Li report following variables in annualized percentages only in Summary statistics!!!
annualizedLi <- c("Icc", "D/P", "D/Y", "E/P", "E10/P", "Tbill", "Lty", "Ts", "Ifl") #this needs to be *100, because already annualized in DS and I transformed in to decimals
annualizedLi2 <- c("totalReturnChange", "Iprod", "Vola", "Unem", "Rms", "Pcon", "C/P", "C/O") #this needs to be * 1200 -> annualizing *12, percentage *100

finalIndexMonthlySS[,(colnames(finalIndexMonthlySS) %in% annualizedLi)] <- finalIndexMonthlySS[,(colnames(finalIndexMonthlySS) %in% annualizedLi)]*100
finalIndexMonthlySS[,(colnames(finalIndexMonthlySS) %in% annualizedLi2)] <- finalIndexMonthlySS[,(colnames(finalIndexMonthlySS) %in% annualizedLi2)]*1200

#mean and std
meantmp <- sapply(finalIndexMonthlySS, mean)
sdtmp <- sapply(finalIndexMonthlySS, sd)
summaryStat <- data.frame(meantmp, sdtmp)

#autocorrelation
autocorrelation <- as.data.frame(sapply(finalIndexMonthlySS, function (u) c(acf(u, plot = FALSE, lag.max = 60,  na.action = na.pass)$acf)))
sumauto <- tail(autocorrelation, -1)[,"totalReturnChange"] #hier, da dependent variable bis lag zusammengerechnet wird
sumauto <- data.frame(sum(sumauto[1]), sum(sumauto[1:12]), sum(sumauto[1:24]), sum(sumauto[1:36]), sum(sumauto[1:48]), sum(sumauto[1:60]))
autocorrelation <- as.data.frame(t(autocorrelation[c(2, 13, 25, 37, 49, 61),])) #richtige rows dann raussuchen und dann invertieren
colnames(autocorrelation) <- c("1", "12", "24", "36", "48", "60")
summaryStat <- bind_cols(summaryStat, autocorrelation)
summaryStat["totalReturnChange",3:8] <- sumauto[1,]

#correlation matrix
summaryStatCormatrix <- cor(finalmonthlyOnlyIV[,-1])
summaryStatCormatrix <- data.frame(summaryStatCormatrix)


###############Univariate in-sample regressions###############
anzahlIV <- ncol(finalmonthlyOnlyIV) - 1   #number of independent Variables. -1, because date is also in table
veck = c(1,12,24,36,48)
#resultsUV is final table for univariate insample regression. 
#rownumber is 5*anzahlIV + 2*anzahlIV, because there are 5 K's + 1 space for the variable name + 1 space for the average
#colnumer 5, because I report the K's, coeffiecients, Newey West tstat, Newey West pval and R^2
resultsUV = matrix(NA, nrow=(5*anzahlIV + 2*anzahlIV), ncol=5)
resultsUV[,1] = append(append("Variable", veck), "Avg")
T <- dim(finalIndexMonthly)[1]
counterName <- 1 #to take the correct independent variable
counterZeile <- 1 #to bring the value to the correct cell

for(i in seq(from = 1, to = nrow(resultsUV), by = 7)) {
  resultsUV[i,2] <- names(finalIndexMonthly[counterName])
  for (j in 1:length(veck)) {
    J <- veck[j]
    lmreg = lm (finalIndexMonthly[1:(T-J + 1), c(paste("K", J, sep = ""), names(finalIndexMonthly[counterName]))])
    coeffs <- summary(lmreg)$coeff
    nw <- coeftest(lmreg, vcov = NeweyWest(lmreg, prewhite = FALSE, lag = J - 1)) #prewhite needs to be set to false if we want Newey West (1987)
    resultsUV[i+counterZeile,2] = round(coeffs[2,1],2)                  #coefficient
    #resultsUV[i+counterZeile,3] = round(coeffs[2,3],2)                 #tstat
    resultsUV[i+counterZeile,3] = round(nw[2,3],2)                      #newey west t stat
    resultsUV[i+counterZeile,4] = round(nw[2,4],3)                      #pvalue
    resultsUV[i+counterZeile,5] = round(summary(lmreg)$adj.r.squared,2) #r^2 value
    counterZeile <- counterZeile + 1
  }
  #average der coefficients
  resultsUV[i+counterZeile, 2] = round(mean(as.numeric(resultsUV[(i+1):(i+5), 2])),2)
  
  counterZeile <- 1
  counterName <- counterName + 1
}
resultsUV <- data.frame(resultsUV)
colnames(resultsUV) <- c("K", "b", "t-val (adjv)", "pval", "adjR2")

#Addings stars indicating significance into table
withStar <- lapply(resultsUV[,"pval"], function(x) {
  temp <- is.na(x)
  if (temp) {
    x <- NA
  } else {
    temp2 <- abs(as.numeric(x))
    if(temp2 <= 0.01) {
      x <- paste(x, "***", sep = "")

    } else if (temp2 <= 0.05) {
      x <- paste(x, "**", sep = "")

    } else if (temp2 <= 0.1) {
      x <- paste(x, "*", sep = "")

    } else
      x <- x
  }
})
withStar <- unlist(withStar)
resultsUV$`pval` <- withStar



###############Multivariate in-sample regressions###############
#I consider 2 models. First one takes all variables except the ones with multicolinearity.
#Furthermore, I also run AIC/pval stepwise backward regression to get only relevant predictors
#Second model takes the variables, which passes the vif criterion with vif < 10.
anzahlIV <- ncol(finalmonthlyOnlyIV) - 1   #number of independent Variables. -1, because date is also in table
veck = c(1,12,24,36,48)

#Test for perfect multicolinearity: Ts with Tbill perfect colinear
multiVar <- lm(finalIndexMonthly[,"K1"] ~ ., finalIndexMonthly[,1:22]) 
alias <- alias(multiVar) #yield perfect collineraity
perfectMulticol <- c("date", "Ts") #also include date in this vector, because i dont want date in my regression
numberOfPerfectMultiCol <- 1

#First model: #colnumer is anzahlIV * 3, because I report the coeffiecient, Newey West tstat, Newey West pval for each predictor also add 2 for R^2 and K's
#i put all predictors in one row and than transform it later to a more visually pleasing table - as seen in thesis
resultMVMain = matrix(NA, nrow = length(veck) + 3, ncol = anzahlIV * 3 + 2 - numberOfPerfectMultiCol*3) 
resultMVMain[,1] = append(append(append(NA, "K"), veck),"Avg")
T <- dim(finalIndexMonthly)[1]
counterName <- 2 #to find the correct coefficient in the coeffs table
counterZeile <- 3 #to insert the values into right cell

for(i in seq(from = 1, to = length(veck), by = 1)) {
  I <- veck[i]
  lmreg = lm (finalIndexMonthly[1:(T-I + 1), c(paste("K", I, sep = ""), names(finalmonthlyOnlyIV[,!(colnames(finalmonthlyOnlyIV) %in% perfectMulticol)]))])
  coeffs <- summary(lmreg)$coeff
  nw <- coeftest(lmreg, vcov = NeweyWest(lmreg, prewhite = FALSE, lag = I - 1)) #prewhite needs to be set to false if we want Newey West (1987)
  #here also -1*3, becuase 1 variables yield perfect multicolinearity 
  for (j in seq(from = 1, to = anzahlIV*3 - numberOfPerfectMultiCol*3, by = 3)) {
    resultMVMain[counterZeile, j + 1] = round(coeffs[counterName, 1],2)    #coefficient
    #resultMVMain[counterZeile, j + 2] = round(coeffs[counterName, 3],2)   #normal tstat
    resultMVMain[counterZeile, j + 2] = round(nw[counterName, 3],2)        #adjusted t stat
    resultMVMain[counterZeile, j + 3] = round(nw[counterName,4],3)         #pvalue
    
    counterName <- counterName + 1
  }
  resultMVMain[counterZeile, ncol(resultMVMain)] <- round(summary(lmreg)$adj.r.squared,2)
  counterName <- 2
  counterZeile <- counterZeile + 1
}
#average coefficient
for(i in seq(from = 1, to = ncol(resultMVMain), by = 3)) {
  resultMVMain[8, i + 1] = round(mean(as.numeric(resultMVMain[3:7, i+1])),2)
}

#add variable names and correct specifications for row names
t1 <- c(NA, NA)
t2 <- names(finalmonthlyOnlyIV[,!(colnames(finalmonthlyOnlyIV) %in% perfectMulticol)])
t3 <- length(t2)
counter <- 0
for (i in seq(from = 1, to = t3)) {
  t2 <- append(t2, t1, after = i + counter)
  counter <- counter + 2 
}
t2 <- head(append(NA,t2),-1)
resultMVMain[1,2:(ncol(resultMVMain)-1)] <- t2
resultMVMain[2,2:(ncol(resultMVMain)-1)] <- rep(c("b", "Z(b)", "pval"), t3)
resultMVMain[2,ncol(resultMVMain)] <- "adjR2"


#add stars for significance results  
for (i in seq(from = 4, to = ncol(resultMVMain) - 4, by = 3)) {  
  for (j in seq(from = 3, to = 7)) {  
    temp <- as.numeric(resultMVMain[j,i])
    if(temp <= 0.01) {
      resultMVMain[j,i] <- paste(temp, "***", sep = "")

    } else if (temp <= 0.05) {
      resultMVMain[j,i] <- paste(temp, "**", sep = "")

    } else if (temp <= 0.1) {
      resultMVMain[j,i] <- paste(temp, "*", sep = "")

    } else
      temp <- "do nothing"#do nothing
  }
}

#make the table clearer -> make it more horizontal than vertical. Put max 2 variables in a row and add R^2 at the end
temp <- resultMVMain
#initializing
resultMVMain <- temp[,1:7]
for (i in seq(8,ncol(temp) - 6, by = 6)) {
  t1 <-cbind(resultMVMain[1:8,1], temp[,i:(i+5)])
  names(t1) <- names(resultMVMain)
  resultMVMain <- rbind(resultMVMain, t1)
}
#add R^2 into table
t1 <- matrix(NA, nrow = 8, ncol = 7)
t1[1:8,2:4] <- temp[1:8,62:64]
t1[1:8,6] <- temp[1:8, ncol(temp)]
t1[1:8,1] <- temp[1:8, 1]
names(t1) <- names(resultMVMain)
resultMVMain <- rbind(resultMVMain, t1)
resultMVMain <- data.frame(resultMVMain)


#Second model: first check Vif criterion -> omit highest vif until no variable yields vif > 5
I <- 1
lmreg = lm (finalIndexMonthly[1:(T-I + 1), c(paste("K", I, sep = ""), names(finalmonthlyOnlyIV[,!(colnames(finalmonthlyOnlyIV) %in% perfectMulticol)]))])
vif(lmreg) #cutoff >5

vifOmmit <- append(perfectMulticol, "D/P")
lmreg = lm (finalIndexMonthly[1:(T-I + 1), c(paste("K", I, sep = ""), names(finalmonthlyOnlyIV[,!(colnames(finalmonthlyOnlyIV) %in% vifOmmit)]))])
vif(lmreg) #cutoff >5 #


vifOmmit <- append(vifOmmit, "E10/P")
lmreg = lm (finalIndexMonthly[1:(T-I + 1), c(paste("K", I, sep = ""), names(finalmonthlyOnlyIV[,!(colnames(finalmonthlyOnlyIV) %in% vifOmmit)]))])
vif(lmreg) #cutoff >5 #


vifOmmit <- append(vifOmmit, "E/P")
lmreg = lm (finalIndexMonthly[1:(T-I + 1), c(paste("K", I, sep = ""), names(finalmonthlyOnlyIV[,!(colnames(finalmonthlyOnlyIV) %in% vifOmmit)]))])
vif(lmreg) #cutoff >5 #


vifOmmit <- append(vifOmmit, "Icc")
lmreg = lm (finalIndexMonthly[1:(T-I + 1), c(paste("K", I, sep = ""), names(finalmonthlyOnlyIV[,!(colnames(finalmonthlyOnlyIV) %in% vifOmmit)]))])
vif(lmreg) #cutoff >5 #


vifOmmit <- append(vifOmmit, "Lty")
lmreg = lm (finalIndexMonthly[1:(T-I + 1), c(paste("K", I, sep = ""), names(finalmonthlyOnlyIV[,!(colnames(finalmonthlyOnlyIV) %in% vifOmmit)]))])
vif(lmreg) #cutoff >5 #

vifOmmit <- append(vifOmmit, "D/Y")
lmreg = lm (finalIndexMonthly[1:(T-I + 1), c(paste("K", I, sep = ""), names(finalmonthlyOnlyIV[,!(colnames(finalmonthlyOnlyIV) %in% vifOmmit)]))])
vif(lmreg) #cutoff >5 #


vifPredictors <- c("D/E", "B/M", "Equis", "Ntis", "Ndy", "Tbill", "Ifl", "Rms"#, "Exr"
                   , "C/P", "C/O", "Pcon", "Iprod", "Unem", "Svr", "Vola")
resultsMVAppendixVif <- matrix(NA, nrow = length(veck) + 3, ncol = length(vifPredictors) * 3 + 2)
resultsMVAppendixVif[,1] = append(append(append(NA, "K"), veck), "Avg")
counterZeileVif <- 3 #to insert into right cell
counterName <- 2 #to get correct independent varible coefficient
for(i in seq(from = 1, to = length(veck), by = 1)) {
  I <- veck[i]
  
  lmreg = lm (finalIndexMonthly[1:(T-I + 1), c(paste("K", I, sep = ""), names(finalmonthlyOnlyIV[,(colnames(finalmonthlyOnlyIV) %in% vifPredictors)]))])
  coeffs <- summary(lmreg)$coeff
  nw <- coeftest(lmreg, vcov = NeweyWest(lmreg, prewhite = FALSE, lag = I - 1)) #prewhite needs to be set to false if we want Newey West (1987)
  for (j in seq(from = 1, to = length(vifPredictors) * 3 + 2 - 3, by = 3)) {
    resultsMVAppendixVif[counterZeileVif, j + 1] = round(coeffs[counterName, 1],2)    #coefficient
    #resultsMVAppendixVif[counterZeileVif, j + 2] = round(coeffs[counterName, 3],2)   #normal tstat
    resultsMVAppendixVif[counterZeileVif, j + 2] = round(nw[counterName, 3],2)        #adjusted t stat
    resultsMVAppendixVif[counterZeileVif, j + 3] = round(nw[counterName,4],3)         #pvalue:
    
    counterName <- counterName + 1
  }
  resultsMVAppendixVif[counterZeileVif, ncol(resultsMVAppendixVif)] <- round(summary(lmreg)$adj.r.squared,2)
  counterName <- 2
  counterZeileVif <- counterZeileVif + 1
}
#average coefficient
for(i in seq(from = 1, to = ncol(resultsMVAppendixVif), by = 3)) {
  resultsMVAppendixVif[8, i + 1] = round(mean(as.numeric(resultsMVAppendixVif[3:7, i+1])),2)
}

#add variable names and correct specifications
t1 <- c(NA, NA)
t2 <- vifPredictors
t3 <- length(t2)
counter <- 0
for (i in seq(from = 1, to = t3)) {
  t2 <- append(t2, t1, after = i + counter)
  counter <- counter + 2 
}
t2 <- head(append(NA,t2),-1)
resultsMVAppendixVif[1,2:(ncol(resultsMVAppendixVif)-1)] <- t2
resultsMVAppendixVif[2,2:(ncol(resultsMVAppendixVif)-1)] <- rep(c("b", "Z(b)", "pval"), t3)
resultsMVAppendixVif[2,ncol(resultsMVAppendixVif)] <- "adjR2"

#add stars for significance levels      
for (i in seq(from = 4, to = ncol(resultsMVAppendixVif) - 4, by = 3)) {             
  for (j in seq(from = 3, to = 7)) {                    
    temp <- as.numeric(resultsMVAppendixVif[j,i])
    if(temp <= 0.01) {
      resultsMVAppendixVif[j,i] <- paste(temp, "***", sep = "")
      
    } else if (temp <= 0.05) {
      resultsMVAppendixVif[j,i] <- paste(temp, "**", sep = "")
      
    } else if (temp <= 0.1) {
      resultsMVAppendixVif[j,i] <- paste(temp, "*", sep = "")
      
    } else
      temp <- "do nothing"#do nothing
  }
}

#make the table clearer -> make it more horizontal than vertical. Put max 2 variables in a row and add Vola and R^2 at the end
temp <- resultsMVAppendixVif
#initializing
resultsMVAppendixVif <- temp[,1:7]
for (i in seq(8,ncol(temp) - 6, by = 6)) {
  t1 <-cbind(resultsMVAppendixVif[1:8,1], temp[,i:(i+5)])
  names(t1) <- names(resultsMVAppendixVif)
  resultsMVAppendixVif <- rbind(resultsMVAppendixVif, t1)
}
t1 <- matrix(NA, nrow = 8, ncol = 7)
t1[1:8,2:4] <- temp[1:8, (ncol(temp)-3):(ncol(temp)-1)] #add vola
t1[1:8,6] <- temp[1:8, ncol(temp)]                      #add R^2
t1[1:8,1] <- temp[1:8, 1]
names(t1) <- names(resultsMVAppendixVif)
resultsMVAppendixVif <- rbind(resultsMVAppendixVif, t1)
resultsMVAppendixVif <- data.frame(resultsMVAppendixVif)

########Individual out-of-sample Forecast###########
#Start with predictive model
m <- 121   #estimation period         #forecasting starts at 2000-01-31 
q <- T-m   #out of sample period   
I <- 1     #set I to 1, because out-of-sample is only with 1 horizon forecast

#realReturns
realWerte <- as.vector(finalIndexMonthly[(m):(T-I+1), c(paste("K", I, sep = ""))])

#historical average market return
hmean <- NULL
for (a in seq(1:(T-I+1))) {
  hmean <- append(hmean, mean(finalIndexMonthly[1:a, paste ("K", I, sep = "")]))
}
rhmean <- hmean[(m):(T-I+1)]


#utility gain: here portfolio with benchmark
y <- 3 #risk aversion parameter
varianceUT <- frollapply(finalIndexMonthly[,"K1"]*100, 12*10, var)[m:T] #here 12*10, because i have 10 years rolling windows
wVariable <- (1/y)*(rhmean*100/varianceUT)
#restrict wVariable to max 0% or 150%
for (i in seq(1:length(wVariable))) {
  if (wVariable[i] < 0)
    wVariable[i] <- 0
  else if (wVariable[i]>1.5)
    wVariable[i] <- 1.5
  else
    a <- NULL # do nothing
}
utilityHistoric <- mean(wVariable)-(1/2*y*var(wVariable))


#resultmatrix: 
#has a row for each predictor
#has 5 columns -> variable name, R^2os, pval, sign restriction, utility gain
anzahlIV <- ncol(finalmonthlyOnlyIV) - 1   #exclude date column with - 1
resultsForecastUV = matrix(NA, nrow=anzahlIV, ncol=5)
colnames(resultsForecastUV) <- c("Variable", "R^2OS", "pval", "signRE", "Ugain")

#data for figures
figuresOosUV = list()

#this table contains returns of each individual forecast -> needed for combination approach following Rapach et al. (2010)
rapachTabelle <- matrix(NA, nrow = length(m:(T-I+1)), anzahlIV)
colnames(rapachTabelle) <- as.vector(names(finalIndexMonthly[1:anzahlIV]))
rapachCounter <- 1

#first for loop: each predictor is considered
for (i in 1:anzahlIV) {
  resultsForecastUV[i,1] <- names(finalIndexMonthly[i])
  #second for loop is for calculating oos forecast "recursively"
  oospred <- NULL           #normal
  oospreadCampbell <- NULL  #campbell restriction
  for (j in m:T) {
    lmreg = lm (finalIndexMonthly[1:j, c(paste("K", I, sep = ""), names(finalIndexMonthly[i]))])
    coeffs <- summary(lmreg)$coeff
    oospred <- append(oospred, coeffs[1,1] + coeffs[2,1]*finalIndexMonthly[j,names(finalIndexMonthly[i])])
    
    #safe individual forecast for rapach combination later
    rapachTabelle[rapachCounter, i] <- coeffs[1,1] + coeffs[2,1]*finalIndexMonthly[j,names(finalIndexMonthly[i])]
    rapachCounter <- rapachCounter + 1
    
    #campbell restrictions: first check if variable has theoretically correct sign
    temp <- mean(finalIndexMonthly[,names(finalIndexMonthly[i])])
    temp2 <- NULL
    if (((temp < 0) && (coeffs[2,1]) < 0) || ((temp > 0) && (coeffs[2,1] > 0))) {
      temp2 <- coeffs[1,1] + coeffs[2,1]*finalIndexMonthly[j,names(finalIndexMonthly[i])]
    } else {
      temp2 <- coeffs[1,1] + 0*finalIndexMonthly[j,names(finalIndexMonthly[i])]
    }
    #check if returns are greater or equal 0
    oospreadCampbell <- append(oospreadCampbell, ifelse(temp2 < 0, 0, temp2))
  }
  rapachCounter <- 1
  
  #R^2 os
  rNormalR2os <- 1 - sum((realWerte-oospred)^2)/sum((realWerte-rhmean)^2)
  resultsForecastUV[i,2] <- round(rNormalR2os*100,2)                     #*100 because I report in %
  #adjusted MSPE following Clark and West (2007) to get pval
  rAdjustMspe <- (realWerte-rhmean)^2 - ((realWerte-oospred)^2-(rhmean-oospred)^2)
  rAdjustMspeTest <- t.test(rAdjustMspe, alternative = "greater")      #same as summary(lm(test~1))  ### regressing over constant
  resultsForecastUV[i,3] <- round(rAdjustMspeTest$p.value,3)
  #summary(lm(rep(0, length(rAdjustMspe)) ~ rAdjustMspe))
  
  #R^2os with campbell restrictions
  rNormalMspeCampbell <- 1 - sum((realWerte-oospreadCampbell)^2)/sum((realWerte-rhmean)^2)
  resultsForecastUV[i,4] <- round(rNormalMspeCampbell*100,2)            #*100 wegen %te und dann runden

  #data for figures (normal ones. I do not consider figures for campbell restricted ones.
  #If one wants to, they can replace oospred with oospredCampbell to get Campbell restricted figures)
  sqm <- NULL
  for (a in seq(1:length(oospred))) {
    sqm <- append(sqm, sum((realWerte[1:a]-rhmean[1:a])^2) - sum((realWerte[1:a]-oospred[1:a])^2))
  }
  figuresOosUV <- c(figuresOosUV, list(sqm))
  
  #utility gain: here portfolio with forecast
  wVariable <- (1/y)*(oospred*100/varianceUT)
  #restrict wVariable to max 0% or 150%
  for (j in seq(1:length(wVariable))) {
    if (wVariable[j] < 0)
      wVariable[j] <- 0
    else if (wVariable[j]>1.5)
      wVariable[j] <- 1.5
    else
      a <- NULL # do nothing
  }
  utilityPred <- (mean(wVariable)-(1/2*y*var(wVariable)))
  
  #final utility gain: multiply by 1200, to get annualized in percentages
  gain <- (utilityPred - utilityHistoric) * 1200
  resultsForecastUV[i,5] <- round(gain,2)
}
resultsForecastUV <- data.frame(resultsForecastUV)



########Mutlivariate out-of-sample Forecasts###########
#consider 2 models: combination approach following Rapach et al. (2010) and kitchen sink following Goyal and Welch (2008)
#I do not consider campbell and Thompson (2007) restrictions here, as combination approaches always satisfy theoreical restrictions. See Rapach et al. (2010) p. 832
I <- 1     #set I to 1, because out-of-sample is only with 1 horizon forecast


#First: Different combination methods
#table with individual forecast. values obtained from individual forecast
rapachTabelle <- data.frame(rapachTabelle)
rapachTabelle <- round(rapachTabelle, 5)

#mean weight combination
rMean <- rowMeans(rapachTabelle)

#median combination
rMedian <- apply(rapachTabelle, 1, median)

#trimmed mean combination
tmp <- rapachTabelle
Max <- apply(rapachTabelle, 1, max)
Min <- apply(rapachTabelle, 1, min)
for (i in seq(1:nrow(rapachTabelle))) {
  for(j in seq(1:ncol(rapachTabelle))) {
    if (Max[i] == rapachTabelle[i,j]) {
      tmp[i,j] <- 0
      break
    }
  }
  for (j in seq(1:ncol(rapachTabelle))) {
    if (Min[i] == rapachTabelle[i,j]) {
      tmp[i,j] <- 0
      break
    }
  }
}
rTMean <- rowSums(tmp)/(anzahlIV-2)

weightsV1 <- sum(1^(T-seq(m:T))*(realWerte-rapachTabelle[,1])^2)
weightsV2 <- sum(0.9^(T-seq(m:T))*(realWerte-rapachTabelle[,1])^2)

#DMSPE combination
#calculating phi following from thesis equation 8
for (i in seq(1:(anzahlIV-1))) {
  weightsV1 <- append(weightsV1, sum(1^(T-seq((m):T))*(realWerte-rapachTabelle[,i+1])^2))
  weightsV2 <- append(weightsV2, sum(0.9^(T-seq((m):T))*(realWerte-rapachTabelle[,i+1])^2))
}
tmp1 <- weightsV1
tmp2 <- weightsV2

#calculating omega following from thesis equation 7
for (i in seq(1:anzahlIV)) {
  weightsV1[i] <- weightsV1[i]/sum(tmp1)
  weightsV2[i] <- weightsV2[i]/sum(tmp2)
}

#returns are formed with the corresponding weights
tmp1 <- rapachTabelle
tmp2 <- rapachTabelle
for(i in seq(1:nrow(rapachTabelle))) {
  for (j in seq(1:ncol(rapachTabelle))) {
    tmp1[i,j] <- weightsV1[j]*rapachTabelle[i,j]
    tmp2[i,j] <- weightsV2[j]*rapachTabelle[i,j]
  }
}
combDMPSE1 <- rowSums(tmp1)
combDMPSE2 <- rowSums(tmp2)

#the utility gain, adjusted MSPE, and data for figures ard figures for the forecasts are here determined
#first make lists containing all combination approaches
utilityGainOOScomb <- list (rMean, rMedian, rTMean, combDMPSE1, combDMPSE2)
adjustedMSPEOOScomb <- list (rMean, rMedian, rTMean, combDMPSE1, combDMPSE2)
figuresOosMV <- list ()

#for each combination approach
for(i in seq(1:length(utilityGainOOScomb))) {
  oospred <- unlist(utilityGainOOScomb[i])
  #adjusted MSPE following Clark and West (2007)
  rAdjustMspe <- (realWerte-rhmean)^2 - ((realWerte-oospred)^2-(rhmean-oospred)^2)
  rAdjustMspeTest <- t.test(rAdjustMspe, alternative = "greater") #same as summary(lm(test~1))  #regressing over constant
  adjustedMSPEOOScomb[i] <- round(rAdjustMspeTest$p.value,3) 
  
  #data for figures
  sqm <- NULL
  for (a in seq(1:length(oospred))) {
    sqm <- append(sqm, sum((realWerte[1:a]-rhmean[1:a])^2) - sum((realWerte[1:a]-oospred[1:a])^2))
  }
  figuresOosMV <- c(figuresOosMV, list(sqm))
  
  #portfolio with forecast for utility gain
  wVariable <- (1/y)*(oospred*100/varianceUT)
  #restrict wVariable to max 0% or 150%
  for (j in seq(1:length(wVariable))) {
    if (wVariable[j] < 0)
      wVariable[j] <- 0
    else if (wVariable[j]>1.5)
      wVariable[j] <- 1.5
    else
      a <- NULL # do nothing
  }
  utilityPred <- (mean(wVariable)-(1/2*y*var(wVariable)))
  
  #final utility gain: multiply by 1200, to get annualized in percentages
  gain <- (utilityPred - utilityHistoric) * 1200
  utilityGainOOScomb[i] <- round(gain,2)
}

#constructing R^2os - *100, because I give it in % as Li et al. (2013)
rMean <- round((1 - sum((realWerte-rMean)^2)/sum((realWerte-rhmean)^2))*100,2)
rMedian <-round((1 - sum((realWerte-rMedian)^2)/sum((realWerte-rhmean)^2))*100,2)
rTMean <- round((1 - sum((realWerte-rTMean)^2)/sum((realWerte-rhmean)^2))*100,2)
combDMPSE1 <- round((1 - sum((realWerte-combDMPSE1)^2)/sum((realWerte-rhmean)^2))*100,2)
combDMPSE2 <- round((1 - sum((realWerte-combDMPSE2)^2)/sum((realWerte-rhmean)^2))*100,2)

#First: kitchen sink model
#consider 4 different variations: 
#a) Normal with every predictor except TS, because perfect multicolinear
oospred <- NULL
#b) Only predictors, who passed vif criterion
oospredReduced <- NULL

for (i in m:T) {
  #a) normal
  lmreg = lm (finalIndexMonthly[1:i, c(paste("K", I, sep = ""), names(finalmonthlyOnlyIV[,!(colnames(finalmonthlyOnlyIV) %in% perfectMulticol)]))])
  coeffs <- summary(lmreg)$coeff
  #count coefficients together
  temp <- NULL
  for (j in seq(1:(anzahlIV - numberOfPerfectMultiCol))){
    temp <- append(temp, coeffs[j+1,1]*finalIndexMonthly[i,names(finalmonthlyOnlyIV[,!(colnames(finalmonthlyOnlyIV) %in% perfectMulticol)])[j]])
  }
  
  oospred <- append(oospred, coeffs[1,1] + sum(temp))
  
  #b) vif
  multiVar <- lm(finalIndexMonthly[1:i, c(paste("K", I, sep = ""), names(finalmonthlyOnlyIV[,(colnames(finalmonthlyOnlyIV) %in% vifPredictors)]))])
  coeffs <- summary(multiVar)$coeff
  #count coefficients together
  temp <- NULL
  for (j in seq(1:length(vifPredictors))) {
    temp <- append(temp, coeffs[j+1,1] * finalIndexMonthly[i,vifPredictors[j]])
  }
  oospredReduced <- append(oospredReduced, coeffs[1,1] + sum(temp))
}

#data for figures
sqm <- NULL  #a) normal
sqm2 <- NULL #b) vif
for (a in seq(1:length(oospred))) {
  sqm <- append(sqm, sum((realWerte[1:a]-rhmean[1:a])^2) - sum((realWerte[1:a]-oospred[1:a])^2))
  sqm2 <- append(sqm2, sum((realWerte[1:a]-rhmean[1:a])^2) - sum((realWerte[1:a]-oospredReduced[1:a])^2))
}
figuresOosMV<- c(figuresOosMV, list(sqm))
figuresOosMV <- c(figuresOosMV, list(sqm2))

#contain all KS into a list, so i can calculate the utility gain more easily in a for loop
alle <- list(oospred, oospredReduced)

#for pvalues and utility gain
adjustedMSPEOOSMVKS <- list()
utilityGainOOSMVKS <- list()

for (i in seq(1:length(alle))) {
  #adjusted MSPE following Clark and West (2007)
  helper <- unlist(alle[i])
  rAdjustMspe <- (realWerte-rhmean)^2 - ((realWerte-helper)^2-(rhmean-helper)^2)
  rAdjustMspeTest <- t.test(rAdjustMspe, alternative = "greater") #same as summary(lm(test~1))  ### regressing over constant
  adjustedMSPEOOSMVKS <- c(adjustedMSPEOOSMVKS, round(rAdjustMspeTest$p.value,3))
  
  #portfolio of forecast for utility gain
  wVariable <- (1/y)*(helper*100/varianceUT)
  #restrict wVariable to max 0% or 150%
  for (j in seq(1:length(wVariable))) {
    if (wVariable[j] < 0)
      wVariable[j] <- 0
    else if (wVariable[j]>1.5)
      wVariable[j] <- 1.5
    else
      a <- NULL # do nothing
  }
  utilityPred <- (mean(wVariable)-(1/2*y*var(wVariable)))
  
  #final utility gain: multiply by 1200, to get annualized in percentages
  gain <- (utilityPred - utilityHistoric) * 1200
  utilityGainOOSMVKS <- c(utilityGainOOSMVKS, round(gain,2))
}


#r^2os values for KS
oospred <- 1 - sum((realWerte-oospred)^2)/sum((realWerte-rhmean)^2)
oospredReduced <- 1 - sum((realWerte-oospredReduced)^2)/sum((realWerte-rhmean)^2)


#insert all calculated values into a data frame
resultsForecastMV = matrix(NA, nrow = 7, ncol = 4) #7 rows, because i have 5 combination approaches und 2 kitchen sink approaches
colnames(resultsForecastMV) <- c("variable", "R^2os", "pval", "Ugain")
resultsForecastMV[1:7,1] <- c("rMean", "rMedian", "rTMean", "combDMPSE1", "combDMPSE2", "KS1", "KS2")
resultsForecastMV[1:7,2] <- c(rMean, rMedian, rTMean, combDMPSE1, combDMPSE2, round(oospred*100,2), round(oospredReduced*100,2))
resultsForecastMV[1:7,3] <- c(unlist(adjustedMSPEOOScomb[1]), unlist(adjustedMSPEOOScomb[2]),unlist(adjustedMSPEOOScomb[3]),unlist(adjustedMSPEOOScomb[4]),
                              unlist(adjustedMSPEOOScomb[5]), unlist(adjustedMSPEOOSMVKS[1]), unlist(adjustedMSPEOOSMVKS[2])) 
resultsForecastMV[1:7,4] <- c(unlist(utilityGainOOScomb[1]), unlist(utilityGainOOScomb[2]),unlist(utilityGainOOScomb[3]),unlist(utilityGainOOScomb[4]),
                              unlist(utilityGainOOScomb[5]), unlist(utilityGainOOSMVKS[1]), unlist(utilityGainOOSMVKS[2]))
resultsForecastMV <- data.frame(resultsForecastMV)

#here add stars to significant p values
for (j in seq(from = 1, to = nrow(resultsForecastUV))) {                  
    temp <- as.numeric(resultsForecastUV[j,3])
    if(temp <= 0.01) {
      resultsForecastUV[j,3] <- paste(temp, "***", sep = "")
      
    } else if (temp <= 0.05) {
      resultsForecastUV[j,3] <- paste(temp, "**", sep = "")
      
    } else if (temp <= 0.1) {
      resultsForecastUV[j,3] <- paste(temp, "*", sep = "")
      
    } else
      temp <- "do nothing"#do nothing
}

for (j in seq(from = 1, to = nrow(resultsForecastMV))) {                   
  temp <- as.numeric(resultsForecastMV[j,3])
  if(temp <= 0.01) {
    resultsForecastMV[j,3] <- paste(temp, "***", sep = "")
    
  } else if (temp <= 0.05) {
    resultsForecastMV[j,3] <- paste(temp, "**", sep = "")
    
  } else if (temp <= 0.1) {
    resultsForecastMV[j,3] <- paste(temp, "*", sep = "")
    
  } else
    temp <- "do nothing"#do nothing
}

#plotting figures of OOS and saving them - individual forecasts
for (i in seq(from = 1, to = nrow(resultsForecastUV))) {
  #using t1, to determine file name. Using grepl, because I named some variables with "/". This is not allowed as a png dataname
  t1 <- paste0(resultsForecastUV[i,1],".png")
  test <- grepl("/", t1, fixed = TRUE)
  if (test) {
    t1 <- gsub("/", "", t1)
  }
  name <- paste0("./MonthlyRegressionFigures/", t1)
  t2 <- unlist(figuresOosUV[i])
  #specifications on how I want the figure output to be
  png(filename = name, width = 600, height = 200)
  print(ggplot() + geom_line(aes(x=Date[m:T], y = t2)) + geom_hline(yintercept = 0, color = "blue", alpha = 0.7, size = 1.0, linetype = "dotted") +
          theme_bw() +
          theme(panel.grid.major = element_blank(),
                panel.grid.minor = element_blank(),
                panel.border = element_rect(colour = "black", fill=NA, size=0.5),
                panel.background = element_blank(),
                axis.text=element_text(size=22),
                axis.title = element_text(size=22)
                )+ scale_x_date(date_labels = "%m-%Y") +
        labs(y = "", x = "Date")
          )
  dev.off()
}

#plotting figures of OOS and saving them - multivariate forecasts
for (i in seq(from = 1, to = nrow(resultsForecastMV))) {
  #using t1, to determine file name. Using grepl, because I named some variables with "/". This is not allowed as a png dataname
  t1 <- paste0(resultsForecastMV[i,1],".png")
  test <- grepl("/", t1, fixed = TRUE)
  if (test) {
    t1 <- gsub("/", "", t1)
  }
  name <- paste0("./MonthlyRegressionFigures/", t1)
  t2 <- unlist(figuresOosMV[i])
  #specifications on how I want the figure output to be
  png(filename = name, width = 600, height = 200)
  print(ggplot() + geom_line(aes(x=Date[m:T], y = t2)) + geom_hline(yintercept = 0, color = "blue", alpha = 0.7, size = 1.0, linetype = "dotted") +
          theme_bw() +
          theme(panel.grid.major = element_blank(),
                panel.grid.minor = element_blank(),
                panel.border = element_rect(colour = "black", fill=NA, size=0.5),
                panel.background = element_blank(),
                axis.text=element_text(size=22),
                axis.title = element_text(size=22)
          )+ scale_x_date(date_labels = "%m-%Y") +
          labs(y = "", x = "Date")
  )
  dev.off()
}


###############Get every table as an xlsx file###############
write_xlsx(finalIndexMonthly, "./MonthlyRegressionData/finalIndexMonthly.xlsx") #all data with dependent and independent variables
write_xlsx(summaryStat, "./MonthlyRegressionTables/summaryStat.xlsx") #table with summary statistics
write_xlsx(summaryStatCormatrix, "./MonthlyRegressionTables/summaryStatCormatrix.xlsx") #table with correlation matrix
write_xlsx(unitRootTestMonthly, "./MonthlyRegressionTables/unitRootTestMonthly.xlsx") #table with unitroot tests
write_xlsx(resultsUV, "./MonthlyRegressionTables/resultsUV.xlsx") #table with univariate insample results
write_xlsx(resultMVMain, "./MonthlyRegressionTables/resultsMVMain.xlsx") #table with multivariate insample results full model
write_xlsx(resultsMVAppendixVif, "./MonthlyRegressionTables/resultsMVAppendixVif.xlsx") #table with multivariate insample results vif model
write_xlsx(resultsForecastUV, "./MonthlyRegressionTables/resultsForecastUV.xlsx") #table with univariate out of sample results
write_xlsx(resultsForecastMV, "./MonthlyRegressionTables/resultsForecastMV.xlsx") #table with multivariate out of sample results

#for the data for the figures we need to make dataframes first, as I use lists in my code
figures1 <- t(data.frame(matrix(unlist(figuresOosUV), nrow=length(figuresOosUV), byrow=TRUE))) #data univariate
figures2 <- t(data.frame(matrix(unlist(figuresOosMV), nrow=length(figuresOosMV), byrow=TRUE))) #data multivariate
colnames(figures1) <- resultsForecastUV$Variable
colnames(figures2) <- resultsForecastMV$variable
figures <- cbind(figures1, figures2)
figures <- data.frame(figures)
write_xlsx(figures, "./MonthlyRegressionData/figures.xlsx")


###############Remove unnecessary things###############
rm(list = ls.str(mode = 'numeric'))
rm(list = ls.str(mode = 'character'))
rm(list = ls.str(mode = 'NULL'))
rm(dependentVariable, lehrstuhlIndex, mk12, mk24, mk36, mk48, autocorrelation, sumauto, lmreg, multiVar, tmp1, tmp2, test)
rm(finalmonthlyOnlyIV, alle, monthlyIndex, alias, unitRootTestMonthlyTrend, finalIndexMonthlySS)
rm(adjustedMSPEOOScomb, adjustedMSPEOOSMVKS, rAdjustMspeTest, rapachTabelle, tmp, utilityGainOOScomb, utilityGainOOSMVKS)
rm(figuesOosMV, figuesOosUV)