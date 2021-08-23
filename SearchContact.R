library(RCurl)
library(XML)
library('dplyr')
library(readxl)
library('stringr')
library('openxlsx')
library(svDialogs)
require(XML)
require(RCurl)

SearchPhoneNoOnGoogle = function(searchString)
{
  site <- getForm("http://www.google.com/search", hl="en",lr="", q=searchString, btnG="Search")
  siteHTML= htmlTreeParse(site)
  indicatorWord = "Phone"
  #print(siteHTML)
  myvec= gregexpr(indicatorWord,siteHTML,ignore.case = FALSE)
  #print(myvec)
  myvec =unlist(myvec)
  myvec= myvec %>% dplyr::na_if("-1")
  myvec =myvec[!is.na(myvec)]
  #print(myvec)
  results=""
  if(length(myvec)>0)
  {
    for(i in 1:length(myvec))
    {
      results= HelperFunction(myvec[i],siteHTML)
      if(results !="")
      {
        break;
      }
    }
  }
  return(results)
}

HelperFunction = function(position,siteHTML)
{
  posExtractStart= as.numeric(position)
  stringExtract <- substring(siteHTML, first=posExtractStart,   last = posExtractStart + 600)
  print(stringExtract)
  stringExtract= stringExtract %>% dplyr::na_if("")
  stringExtract =stringExtract[!is.na(stringExtract)]
  posResults = str_locate(stringExtract,'((\\(\\d{3}\\) ?)|(\\d{3}-))?\\d{3}-\\d{4}')
  tempVec= as.numeric(posResults[!is.na(posResults)])
  results=""
  if(length(tempVec)>0)
  {
    results =substring(stringExtract, first= tempVec[1],last= tempVec[2])
  }
  return(results)
}

fname <- file.choose()
SiteList <- read_excel(fname,sheet = "FindPhoneNo", col_names = TRUE)
CustomerName= dlg_input("Enter Customer Name")$res
View(SiteList)
#Identifier	Address1	City	State	Zip
SiteList$`Search String`=NA
SiteList$`Google Searched Phone`=""
for(i in 1:nrow(SiteList))
{
  SiteList$`Search String`[i]= paste(SiteList[i,2],SiteList[i,3],SiteList[i,4],SiteList[i,5],CustomerName, sep = " ")
  SiteList$`Google Searched Phone`[i] = SearchPhoneNoOnGoogle(SiteList$`Search String`[i])
  Sys.sleep(1.5) #sleep for 1.5s between requests
  print(i)
}
View(SiteList)

wb = createWorkbook()
addWorksheet(wb, "StoreDetails") 
writeData(wb, sheet = "StoreDetails", SiteList) # Will default add from cell A1 
date=unlist(strsplit(as.character.Date(Sys.time()),":"))
saveWorkbook(wb,paste(CustomerName,"_StoreDetails",paste(date[1],"_",date[2],sep=""),".xlsx",sep = ""),overwrite = T) 
