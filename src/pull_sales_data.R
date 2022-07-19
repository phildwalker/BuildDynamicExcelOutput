# pull top retailer revenue (sales) - exec sumry
# Thu Jan 06 16:30:24 2022 ------------------------------

library(DBI)
library(glue)
library(tidyverse)
library(lubridate)

# datesLk <- readxl::read_excel(here::here("input", "DateLookUp.xlsx"))

# ----- Sales Report Date

# Today <- Sys.Date()
# # Today <- as.Date("2022-01-23")
# 
# RPTDate <- 
#   if(as.numeric(lubridate::day(Today)) < 22){
#     Today %m-% months(1) %>% lubridate::floor_date(., "month")-1
#   } else {
#     # statement(s) will execute if the boolean expression is false.
#     Today %m-% months(1) %>% lubridate::floor_date(., "month")
#   }


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
where fas.DateKey in (20191231, 20201231, 20211231) 
and lastlsf > 0
order by CENTERNAME, fas.datekey desc, tenantname
  ",
.con = warehouse)

parm_sql <- dbSendQuery(warehouse, pull_sql)
Sales_center <- dbFetch(parm_sql)

dbClearResult(parm_sql)


#------------
unique(Sales_center$BldgGroupName)

Centers <- c("Daytona Beach Outlet Center", "Foley Outlet Center", "Gonzales Outlet Center", "San Marcos Outlet Center", "Houston Outlet Center",
             "Fort Worth Outlet Center", "Branson Outlet Center", "Commerce Outlet Center", "Locust Grove Outlet Center", "Savannah Outlets",
             "Columbus Outlet Center", "Howell Outlet Center", "Grand Rapids Outlet Center", "Pittsburgh Outlet Center", "Tilton Outlet Center")

notInc <- c("Charlotte Outlet Center", "Jeffersonville Outlet Center", "Nags Head Outlet Center", "Ocean City Outlet Center",
            "Park City Outlet Center", "St. Sauveur", "Terrell Outlet Center", "Williamsburg Outlet Center")

#------ Clean sales -----

CleanSales <- 
  Sales_center %>% 
  filter(#BldgGroupName %in% Centers,
    !BldgGroupName %in% notInc,
    ISTEMP == "N") %>% 
  mutate(Date = as.Date(fulldate),
         Year = lubridate::year(Date)) %>% 
  rename(GrossSales = CurrentYearSales12M) %>% 
  mutate(across(where(is.character), str_trim),
         Center = str_remove_all(BldgGroupName, " Outlet Center| Outlets")) %>%
  select(Date, Year, Center, GrossSales, TenantName, ChainName) 

# unique(CleanSales$Center)

  # ------- Top N - Report Month -------------
  
  RptTop <-
    CleanSales %>%
    group_by(Year, Center) %>%
    arrange(-GrossSales, ChainName) %>%
    mutate(rank = row_number(),
           rank = as.numeric(rank)) %>% 
  filter(rank <= 10)




# RankCenter <-
  RptTop %>% 
  select(-Date, -ChainName) %>% 
  pivot_wider(names_from = Year, values_from = c(TenantName, GrossSales)) %>% 
  arrange(Center, rank) %>% 
  select(Center, rank, TenantName_2021, GrossSales_2021, TenantName_2020, GrossSales_2020, TenantName_2019, GrossSales_2019) %>% 
  write_csv(., here::here("output", "Top10_Sales_ByCenter_3years.csv"))
  

  
  
  
  
  
  
  
  
  
  