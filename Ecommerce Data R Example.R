#loading in my normal packages
pacman::p_load(pacman, dplyr, GGally, ggplot2, ggthemes, 
               ggvis, httr, lubridate, plotly, rio, rmarkdown, shiny, 
               stringr, tidyr, ExcelFunctionsR, openxlsx, data.table,tidyverse)
library(dplyr)
library(tidyverse)
library(lubridate)
library(ExcelFunctionsR)
library(openxlsx)
library(data.table)
library(ggplot2)
library(rio)

#importing csv data files
addsToCart <- read.csv("C:/Users/Sarah/Desktop/DataAnalyst_Ecom_data_addsToCart.csv")
sessionCounts <- read.csv("C:/Users/Sarah/Desktop/DataAnalyst_Ecom_data_sessionCounts.csv")

#changing date into date format then converting that into date from character
sessionCounts$dim_date <- as.Date(sessionCounts$dim_date, format("%m/%d/%y"), tryFormats = c("%Y-%m-%d", "%Y/%m/%d"),
                         optional = FALSE)
#reducing to month for grouping
sessionCounts$dim_date <- month(sessionCounts$dim_date)

#renaming columns just to make it look more presentable
colnames(sessionCounts)[colnames(sessionCounts) == "dim_date"] <- "Month"
colnames(sessionCounts)[colnames(sessionCounts) == "dim_deviceCategory"] <- "DeviceCategory"
colnames(sessionCounts)[colnames(sessionCounts) == "sessions"] <- "Sessions"
colnames(sessionCounts)[colnames(sessionCounts) == "transactions"] <- "Transactions"

#deleting un-needed column
sessionCounts%>%
  select(-dim_browser) -> sessionCounts
view(sessionCounts)

#concatinating month and year into single column
addsToCart$MonthYear <- paste(addsToCart$dim_month, "/01/", addsToCart$dim_year)

#changing date into date format then converting that into date from character
addsToCart$MonthYear <- as.Date(addsToCart$MonthYear, format("%m /%d/ %Y"), tryFormats = c("%Y-%m-%d", "%Y/%m/%d"),
                         optional = FALSE)

#reducing to month for grouping
addsToCart$MonthYear <- month(addsToCart$MonthYear)

#renaming columns just to make it look more presentable
colnames(addsToCart)[colnames(addsToCart) == "MonthYear"] <- "Month"
colnames(addsToCart)[colnames(addsToCart) == "addsToCart"] <- "AddsToCart"

#deleting un-needed column
addsToCart%>%
  select(-dim_month, -dim_year) -> addsToCart
view(addsToCart)

#Creating the first table
Table1 <- sessionCounts %>%
  select(Sessions, Transactions, QTY, Month, DeviceCategory)%>%
  group_by(Month, DeviceCategory)%>%
  summarise(Sessions=sum(Sessions), Transactions=sum(Transactions), QTY=sum(QTY))%>%
  mutate(ECR = Transactions/Sessions)
class(Table1$ECR) <- "percentage"
view(Table1)

# dropping device category because it cant be used with addtocart data
updatedData <- Table1%>%
  select(-DeviceCategory)%>%
  group_by(Month)%>%
  summarise(Sessions=sum(Sessions), Transactions=sum(Transactions), QTY=sum(QTY))%>%
  mutate(ECR = Transactions/Sessions)
class(updatedData$ECR) <- "percentage"
view(updatedData)
  
#Merging table with adds to cart with Month being the mutual column
SummaryData <- merge(updatedData, addsToCart, by = "Month")

#Creating columns to calculat difference in variable from previous month
SummaryData$SessionsDiff <- -lag(SummaryData$Sessions, 1) + SummaryData$Sessions
SummaryData$TransactionDiff <- -lag(SummaryData$Transactions, 1) + SummaryData$Transactions
SummaryData$QTYDiff <- -lag(SummaryData$QTY, 1) + SummaryData$QTY
SummaryData$AddsToCartDiff <- -lag(SummaryData$AddsToCart, 1) + SummaryData$AddsToCart
SummaryData$ECRDiff <- -lag(SummaryData$ECR, 1) + SummaryData$ECR
class(SummaryData$ECR) <- "percentage"
class(SummaryData$ECRDiff) <- "percentage"
view(SummaryData)

#Creating second table
Table2 <- SummaryData%>%
  select( Month, Sessions, Transactions, QTY, AddsToCart, ECR, SessionsDiff, TransactionDiff, QTYDiff, AddsToCartDiff, ECRDiff)%>%
  filter(Month=="5"|Month=="6")%>%
  group_by(Month)
class(Table2$ECR) <- "percentage"
class(Table2$ECRDiff) <- "percentage"
View(Table2)

#creating a couple plots
plot(SummaryData$Month, SummaryData$ECR, type = "l", col = "#cc0000", main = "Conversion Rate Over Year", xlab = "Month of Year", ylab = "Conversion Rate", ylim = c(0.010, 0.040))
plot(SummaryData$Month, SummaryData$Transactions, type = "l", col = "#cc0000", main = "Yearly Transaction Trends", xlab = "Month of Year", ylab = "Tansactions")

hs <- createStyle(halign = "center", textDecoration = "BOLD", fgFill = "#4F81BD")

#creating xlsx file
IXISproject <- createWorkbook(
  creator = "Sarah Perez",
  title = "IXIS Project",
  subject = NULL,
  category = NULL
)
addWorksheet(IXISproject, "Table1") #adding worksheets
addWorksheet(IXISproject, "Table2")
addWorksheet(IXISproject, "SummaryData")
writeData (IXISproject, "Table1", Table1, startCol = 1, startRow = 1, headerStyle = hs)  #adding data
writeData (IXISproject, "Table2", Table2, startCol = 1, startRow = 1, headerStyle = hs)
writeData (IXISproject, "SummaryData", SummaryData, startCol = 1, startRow = 1, headerStyle = hs)
setColWidths(IXISproject, "Table1", cols = 1:6, widths = "auto")
setColWidths(IXISproject, "Table2", cols = 1:11, widths = "auto")
setColWidths(IXISproject, "SummaryData", cols = 1:11, widths = "auto")
saveWorkbook(IXISproject, "IXISproject.xlsx", overwrite = TRUE, returnValue = FALSE)
