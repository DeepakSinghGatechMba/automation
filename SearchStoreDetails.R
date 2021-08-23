library(RCurl)
library("rvest")
library(XML)
library('dplyr')
library(readxl)
library('stringr')
library('openxlsx')
library(svDialogs)
require(XML)
require(RCurl)
require(rvest)

SearchSiteDetails = function(searchString)
{
  site <- getForm("http://www.google.com/search", hl="en",lr="", q=searchString, btnG="Search")
  sitexml=read_html(site)
  myvec=sitexml %>% html_nodes('.AVsepf') %>% html_text()
  print(myvec)
  return(myvec)
}

fname <- file.choose()
SiteList <- read_excel(fname,sheet = "FindStoreDetails", col_names = TRUE)
CustomerName= dlg_input("Enter Customer Name")$res
View(SiteList)
#Identifier	Address1	City	State
SiteList$`Search String`=NA
SiteList$`GoogledDetails`=""
SiteList$`Google Searched Details`=""
for(i in 1:nrow(SiteList))
{
  SiteList$`Search String`[i]= paste(CustomerName,SiteList[i,2],SiteList[i,3],SiteList[i,4], sep = " ")
  myvec= SearchSiteDetails(SiteList$`Search String`[i])
  if(length(myvec)>0)
  {
      SiteList$`GoogledDetails`[i]=myvec
      results=""
      if(length(myvec)>=3)
      {
      #check if it has Address field
        if(length(grep("Address",myvec[1]))>0)
        {
          #extract address from index 1
          address=unlist(strsplit(myvec[1],","))
          street_addr=address[1]
          street_addr=trimws(unlist(strsplit(street_addr,":"))[2])
          city=trimws(address[2])
          statezip=address[3]
          statezip=unlist(strsplit(statezip," "))
          state=statezip[2]
          zip=statezip[3]
        }
        #check if it has Phone field
        if(length(grep("Phone",myvec[3]))>0)
        {
          phone_no=trimws(unlist(strsplit(myvec[3],":"))[2])
        }
        results=paste(zip,"@@@",phone_no,sep="")
      }else {
        #case where 2 or 3 results show up - pass/ignore for now
      }
      #case where 4 or more results
      SiteList$`Google Searched Details`[i] = results
  }
  else
  {
    SiteList$`GoogledDetails`[i]=""
    SiteList$`Google Searched Details`[i]=""
  }
  Sys.sleep(2) #sleep for 2 sec
  print(i)
}
View(SiteList)

wb = createWorkbook()
addWorksheet(wb, "StoreDetails") 
writeData(wb, sheet = "StoreDetails", SiteList) # Will default add from cell A1 
date=unlist(strsplit(as.character.Date(Sys.time()),":"))
saveWorkbook(wb,paste(CustomerName,"_StoreDetails",paste(date[1],"_",date[2],sep=""),".xlsx",sep = ""),overwrite = T) 

