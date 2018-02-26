# Required libraries for the Regression problem
library(forecast)
library(stats)
library(dplyr)
library(imputeTS)
library(XLConnect)
library(xlsx)
library(stringr)
library(plyr)
library(caret)
library(dummies)
library(DMwR)
library(gbm)
library(inTrees)

# Cleaning the Work Environment
rm(list=ls(all=TRUE))

# UDF to trim both leading and trailing white spaces
trim <- function (x) gsub("^\\s+|\\s+$", "", x)

# Defined a custom "cbind" UDF with open parameters to bind datasets with unequal observations with missing value counts with zeros
cbind.fill <- function(...){
  nm <- list(...) 
  nm <- lapply(nm, as.matrix)
  n <- max(sapply(nm, nrow)) 
  do.call(cbind, lapply(nm, function (x) 
    rbind(x, matrix(, n-nrow(x), ncol(x))))) 
}

# Setting the Working directory
setwd("C:/Users/Sekhar/OneDrive/INSOFE/PHD/TS/")

# Reading the xlsx file to a Dataframe
MED = read.xlsx("MacroEconomicData.xlsx", sheetName = "Sheet1")

# Reading the OneHot transformed Holiday file
HLDYS <- read.csv("Holiday.csv")

# Reading the "Train.csv" file provided
TS_Women  <- read.csv("C:/Users/Sekhar/OneDrive/INSOFE/PHD/TS/Train.csv")

# Renaming the subset of the features to be selected for further processing
colnames(MED)[1] <- "YearMonth"
colnames(MED)[4] <- "CPI"
colnames(MED)[6] <- "UER"
colnames(MED)[7] <- "CCInterest"
colnames(MED)[8] <- "PLRateOfInt"
colnames(MED)[9] <- "WagePerHour"
colnames(MED)[10] <- "Advertising"

# Selecting the subset of the features ONLY
MED = data.frame(MED[,c("YearMonth", "CPI", "UER", "CCInterest", "PLRateOfInt", "WagePerHour", "Advertising")])

# Splitting the first column Year and Month separately and then renaming them accordingly
MED <- cbind(str_split_fixed(MED$YearMonth, "-", 2),MED)
colnames(MED)[1] <- "Year"
colnames(MED)[2] <- "Month"

MED$Year <- trim(MED$Year)

# Selection of the features from Macro Economic Data by omitting the combined YearMonth one
MED = data.frame(MED[,c("Year", "Month", "CPI", "UER", "CCInterest", "PLRateOfInt", "WagePerHour", "Advertising")])

# Replacing all "?" in the Advertising feature with zeros
MED = cbind(MED, data.frame(revalue(MED$Advertising, c("?"="0"))))

# Selection of the featues by omitting the old Advertising feature which has "?"
# ALso to rename the "0" imputed feature as Advertising
MED <- subset(MED, select=-c(Advertising))
colnames(MED)[8] <- "Advertising"

# Renaming the feature "TS_Women$Salesx1000" to "TS_Women$Sales"
colnames(TS_Women)[4] <- "Sales"

#Selecting "WomenClothing" product category alone
TS_Women <- TS_Women[TS_Women$ProductCategory == 'WomenClothing',]

# Selecting only the "Sales" feature alone
TS_Women <- TS_Women[,4,drop=FALSE]

# Imputing the missing values in the "Sales" feature using "Interpolation" method from "imputeTS" package
TS_Women <- na.interpolation(TS_Women)

# Binding the "Holidays"HLDYS" dataset with "TS_Women" using the custom cbind function as the observation count differs
TS_Women <- cbind.fill(HLDYS,TS_Women)

#Binding the Macro Economic Data to the "TS_Women"
TS_Women <- cbind(MED,TS_Women)

# Deleting the "HLDYS" and "MED" datasets from the "Global Environment" as they nolonger needed.
rm(HLDYS)
rm(MED)

# Checking for the NaN values and imputing them with "0"
sum(is.na(TS_Women))
TS_Women[is.na(TS_Women)] <- 0

# Separating the 2016 subset of the original dataset for which the monthly "Sales" values to be predicted
predict2016 <- filter(TS_Women, Year == 2016)
TS_Women <- filter(TS_Women, Year <= 2015)

# Deleting the target ("Sales") feature from the dataset to be predicted
predict2016$Sales <- NULL

# Your answer passed the tests! Your score is 129.93%
# Your MAPE is 7.696243%

################################################################################################
# Building the Decision Trees using Gradient Boosting Machine strategy as it is giving the better
# results over and above "Random Forest" and "Linear Regression"
################################################################################################
set.seed(100)
fitForestGBM <- gbm(Sales~Month+CPI+UER+CCInterest+PLRateOfInt+WagePerHour+Advertising+HLD_NYDY+HLD_MLKJ+HLD_VLND+HLD_PRES+HLD_ESTR+HLD_MTHR+HLD_MOMR+HLD_FTHR+HLD_INDO+HLD_INDA+HLD_LABR+HLD_CLMB+HLD_HLWE+HLD_VTRN+HLD_THKS+HLD_CHEV+HLD_CHRS+HLD_NYRE,
                    data = TS_Women, n.trees = 1000, shrinkage=0.001, n.minobsinnode = 10,
                    cv.folds = 50)

# Predicting the monthly 2016 Sales amounts for the "WomenClothing" product category
preds_basic2016_GBM <- predict(fitForestGBM, predict2016, n.trees = 1000)

# Printing the predictions
preds_basic2016_GBM

################################################################################################
# Try to make the business explicability for the model and results built above
################################################################################################

# Printing the GBM model developed above to get a highlevel idea
pretty.gbm.tree(fitForestGBM)

# Transforming the above GBM model to a list of trees
extract <- GBM2List(fitForestGBM, preds_basic2016_GBM)

# Extracting the rules from the above GBM model to cater to the business explicability
extractRules(extract, TS_Women[1:ncol(TS_Women)-1])