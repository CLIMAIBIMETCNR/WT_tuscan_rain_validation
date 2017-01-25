#!/usr/bin/env Rscript
args = commandArgs(trailingOnly=TRUE)

options(java.parameters = "-Xmx4g" )

library(sp)
library(xts)
library(XLConnect)
library(xts)
library(MASS)
library(dplyr)


setwd("D:\\Documenti\\R-LAMMA\\procedure")

source("ian_aux_function.r")

file=args[1]
# tresh_year=0.2
# tresh_inv=0.2
# tresh_prim=0.2
# tresh_est=0.2
# tresh_aut=0.2

tresh_year=args[2]
tresh_inv=args[3]
tresh_prim=args[4]
tresh_est=args[5]
tresh_aut=args[6]
fileout=args[7]
##########################################################################################################################

wb = loadWorkbook(file)
names_sheets=getSheets(wb)

#[1] "ANNUAL_data"    "ANNUAL_station" "WINTER_data"    "WINTER_station" "SPRING_data"    "SPRING_station" "SUMMER_data"   
#[8] "SUMMER_station" "AUTUMN_data"    "AUTUMN_station"


res=list()
z=1




datawts_year = readWorksheet(wb, sheet = 1)
data_year = readWorksheet(wb, sheet = 2)
data_year[data_year == -999] <- NA

for ( i in 6:length(datawts_year)) {
  for ( j in 6:length(data_year)) {
    res[[z]]=cbind(messeri_iannuccilli_Iclass(data_year[,j],datawts_year[,i],treshpar=tresh_year),class=names(datawts_year)[i],name_staz=names(data_year)[j],period="year")
    z=z+1
    }
}

datawts_year = readWorksheet(wb, sheet = 3)
data_year = readWorksheet(wb, sheet = 4)
data_year[data_year == -999] <- NA

for ( i in 6:length(datawts_year)) {
  for ( j in 6:length(data_year)) {
    res[[z]]=cbind(messeri_iannuccilli_Iclass(data_year[,j],datawts_year[,i],treshpar=tresh_inv),class=names(datawts_year)[i],name_staz=names(data_year)[j],period="inv")
z=z+1
  }
}

datawts_year = readWorksheet(wb, sheet = 5)
data_year = readWorksheet(wb, sheet = 6)
data_year[data_year == -999] <- NA

for ( i in 6:length(datawts_year)) {
  for ( j in 6:length(data_year)) {
    res[[z]]=cbind(messeri_iannuccilli_Iclass(data_year[,j],datawts_year[,i],treshpar=tresh_prim),class=names(datawts_year)[i],name_staz=names(data_year)[j],period="prim")
    z=z+1
  }
}
 
datawts_year = readWorksheet(wb, sheet = 7)
data_year = readWorksheet(wb, sheet = 8)
data_year[data_year == -999] <- NA

for ( i in 6:length(datawts_year)) {
  for ( j in 6:length(data_year)) {
    res[[z]]=cbind(messeri_iannuccilli_Iclass(data_year[,j],datawts_year[,i],treshpar=tresh_est),class=names(datawts_year)[i],name_staz=names(data_year)[j],period="est")
    z=z+1
  }
}

datawts_year = readWorksheet(wb, sheet = 9)
data_year = readWorksheet(wb, sheet = 10)
data_year[data_year == -999] <- NA

for ( i in 6:length(datawts_year)) {
  for ( j in 6:length(data_year)) {
    res[[z]]=cbind(messeri_iannuccilli_Iclass(data_year[,j],datawts_year[,i],treshpar=tresh_aut),class=names(datawts_year)[i],name_staz=names(data_year)[j],period="aut")
z=z+1
  }
}

res_final=do.call("rbind",res)
saveRDS(res_final,"res_template.rds")
# res_final=readRDS("res_template.rds")
if (length(list.files(pattern=fileout)) != 0) { file.remove("res_table_ian.xlsx")}
writeWorksheetToFile(fileout, data = res_final, sheet = "results")

