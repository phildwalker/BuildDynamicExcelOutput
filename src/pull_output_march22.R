# pull data for March 2022 view
# Fri May 06 19:05:31 2022 ------------------------------

library(DBI)
library(glue)
library(tidyverse)
library(lubridate)
library(openxlsx)

MonthPull <- "March2022"

cntr_lk <- readxl::read_excel(path = here::here("input", "Center-LookUp.xlsx"))

# --------------------
warehouse <- DBI::dbConnect(odbc::odbc(), dsn = "WAREHOUSESQL")

pull_sql <- glue_sql(
  "
select dimd.fulldate, dimc.CENTERNAME, dimc.BldgGroupName, dimc.[BLDGID], dimt.TangerCatDesc, dimt.TangerSubCatDesc, diml.TenantName, diml.ChainName, diml.leasid 
,dime.PORTID,fas.ISTEMP, fas.IsCurrentlyOpen,fas.ISNONREPORTING, fas.IsOpenForFullYear, diml.zPopUpProgram, lastlsf, [CurrentYearSales12M],
cast(iif(lastlsf = 0 or lastlsf is null, 0, [CurrentYearSales1M]/lastlsf) as money) as AdjustSPSF
from warehouse.dbo.FactAdjustedSales fas 
	left join warehouse.dbo.DimLease diml on fas.LeaseKey = diml.LeaseKey
	left join warehouse.dbo.DimCenter dimc on fas.CenterKey = dimc.CenterKey
	left join warehouse.dbo.DimTenant dimt on fas.TenantKey = dimt.TenantKey
	left join warehouse.dbo.DimEntity dime on fas.EntityKey = dime.EntityKey
	left join warehouse.dbo.DimDate dimd on fas.dateKey = dimd.DateKey
where fas.DateKey in (20190331, 20210331, 20220331) 
and lastlsf > 0
order by CENTERNAME, fas.datekey desc, tenantname
  ",
.con = warehouse)

parm_sql <- dbSendQuery(warehouse, pull_sql)
Sales_center <- dbFetch(parm_sql)

dbClearResult(parm_sql)


# ------------
# unique(Sales_center$BldgGroupName)



#------ Clean sales -----

CleanSales <- 
  Sales_center %>% 
  filter(#BldgGroupName %in% Centers,
    BldgGroupName %in% cntr_lk$BldgGroupName,
    ISTEMP == "N") %>% 
  mutate(Date = as.Date(fulldate),
         Year = lubridate::year(Date)) %>% 
  rename(GrossSales = CurrentYearSales12M) %>% 
  mutate(across(where(is.character), str_trim),
         Center = str_remove_all(BldgGroupName, " Outlet Center| Outlets| - The Arches| Outlet")) %>%
  select(Date, Year, Center, GrossSales, TenantName, ChainName) 

unique(CleanSales$Center)

# ------- Top N - Report Month -------------

RptTop <-
  CleanSales %>%
  group_by(Year, Center) %>%
  arrange(-GrossSales, ChainName) %>%
  mutate(rank = row_number(),
         rank = as.numeric(rank)) %>% 
  filter(rank <= 10)

#------------- List of Centers --------

CntList <- 
  CleanSales %>%
  distinct(Center) %>% 
  pull()



#---- Cleaned to write to xlsx ---------
RankCenter <-
  RptTop %>% 
  select(-Date, -ChainName) %>% 
  pivot_wider(names_from = Year, values_from = c(TenantName, GrossSales)) %>% 
  arrange(Center, rank) %>% 
  select(Center, rank, TenantName_2022, GrossSales_2022, TenantName_2021, GrossSales_2021, TenantName_2019, GrossSales_2019) %>% 
  # write_csv(., here::here("output", glue::glue("Top10_Sales_ByCenter_3years_{MonthPull}.csv"))
  ungroup() %>% 
  rename(Rank = rank,
         Retailer_22 = TenantName_2022,
         `2022` = GrossSales_2022,
         Retailer_21 = TenantName_2021,
         `2021` = GrossSales_2021,
         Retailer_19 = TenantName_2019,
         `2019` = GrossSales_2019)

#---------- overall for center
SumCenter <-
  CleanSales %>% 
  group_by(Center, Year) %>% 
  summarise(TotalSales = sum(GrossSales, na.rm=T)) %>% 
  pivot_wider(names_from = Year, values_from = c(TotalSales), names_prefix = "Gross_") %>% 
  mutate(Type = "Center Gross Sales",
         # Retailer_22 = "",
         Retailer_21 = "",
         Retailer_19 = "") %>% 
  select(Center, Type, Gross_2022, Retailer_21, Gross_2021, Retailer_19, Gross_2019)

TotCenSpl <- split(SumCenter[, c("Type","Gross_2022","Retailer_21","Gross_2021","Retailer_19", "Gross_2019")], SumCenter$Center)

#-------------- Write out multi-tabbed

datSpl <- split(RankCenter[, c("Rank","Retailer_22","2022","Retailer_21","2021","Retailer_19", "2019")], RankCenter$Center)
    
MgrLab <- unique(RankCenter$Center)

    ## Create a blank workbook
wb <- createWorkbook()
    
    ## Loop through the list of split tables as well as their names
    ##   and add each one as a sheet to the workbook
    Map(function(data, dat2, name){
      s <- createStyle(numFmt = "$ #,##0", halign = "left", valign = "center") #"$ #,##0.00"
      
      hs <- createStyle(fontColour = "#030303", fgFill = "#D9D9D9", fontSize = 18,
                        halign = "center", valign = "center", textDecoration = "Bold")
      
      hsc <- createStyle(fontColour = "#030303", fgFill = "#FFFFFF", fontSize = 18,
                        halign = "center", valign = "center", textDecoration = "Bold")
      
      dolS <- createStyle(halign = "left", valign = "center")
      
      # create style, in this case bold header
      header_st <- createStyle(textDecoration = "Bold")
      
      addWorksheet(wb, name)
      addStyle(wb = wb, sheet = name, style = s, rows = 4:20, cols = 3, gridExpand = TRUE)
      addStyle(wb = wb, sheet = name, style = s, rows = 4:20, cols = 5, gridExpand = TRUE)
      addStyle(wb = wb, sheet = name, style = s, rows = 4:20, cols = 7, gridExpand = TRUE)
      
      setColWidths(wb, sheet = name, cols = 1, widths = 7.2) ## set column width for rank

      setColWidths(wb, sheet = name, cols = 2, widths = 23.5) ## set column width for row names column
      setColWidths(wb, sheet = name, cols = 4, widths = 23.5) ## set column width for row names column
      setColWidths(wb, sheet = name, cols = 6, widths = 23.5) ## set column width for row names column
            
      setColWidths(wb, sheet = name, cols = 3, widths = 15) ## set column width for row names column
      setColWidths(wb, sheet = name, cols = 5, widths = 15) ## set column width for row names column
      setColWidths(wb, sheet = name, cols = 7, widths = 15) ## set column width for row names column
      
      writeData(wb, sheet = name, x = "Top 10 RETAILER SALES (Rolling 12 through March 2022)", startRow = 1, startCol = 1)
      addStyle(wb = wb, sheet = name, rows = 1, cols = 1, style = hs)
      mergeCells(wb, sheet = name, cols = 1:7, rows = 1)
      
      writeData(wb, sheet = name, x = toupper(name), startRow = 2, startCol = 1)
      addStyle(wb = wb, sheet = name, rows = 2, cols = 1, style = hsc)
      mergeCells(wb, sheet = name, cols = 1:7, rows = 2)
      
      writeData(wb, name, data, startRow = 3, headerStyle = header_st)
      writeData(wb, name, dat2, startCol = 2, startRow = 15, colNames = FALSE)
      
      # addStyle(wb = wb, sheet = name, rows = 3:15, cols = 3, style = dolS)
      # addStyle(wb = wb, sheet = name, rows = 3:15, cols = 5, style = dolS)
      # addStyle(wb = wb, sheet = name, rows = 3:15, cols = 7, style = dolS)      
      # setHeaderFooter(wb, sheet = name,
      #                 header=c(NA,"This is a header", NA), 
      #                 footer=c(NA, NA, "ALL FOOT RIGHT 2"))
      
    }, data = datSpl, dat2=TotCenSpl, name = names(datSpl))
    
    ## Save workbook to working directory
file_CCname <- glue::glue("Top10_Sales_ByCenter_3years_TTM_{MonthPull}.xlsx")
saveWorkbook(wb, file = here::here("output", file_CCname), overwrite = TRUE) #"Team_Notes","Build_Assoc_Skill_Review",here( CC,FlName)


