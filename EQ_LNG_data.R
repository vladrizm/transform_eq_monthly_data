#### Exploring EnergyQuest LNG data
library(openxlsx)
library(purrr)
library(readxl)

file_path_01 <- "\\update path to \\Monthly Spreadsheets"
file_names_01 <- list.files(file_path_01, pattern = ".xlsx")
length(file_names_01)

paste0(file_path_01,'\\',file_names_01[1])
getSheetNames(paste0(file_path_01,'\\',file_names_01[1]))


list_of_sheets <- list()
for(fn in file_names_01) {
  name_01 <- openxlsx::getSheetNames(paste0(file_path_01,'\\',fn))
  list_of_sheets[[fn]] <- name_01
}

list_of_sheets_01 <- list_of_sheets
n_files <- length(list_of_sheets)

####____________________________________________________________________________

# library(lubridate)
# 
# today <- Sys.Date()
# 
# start_date <- seq(as.Date("1/12/2017", "%d/%m/%Y"),  as.Date('1/07/2022', "%d/%m/%Y"), by = "months")

## Update vector with new date for monthly file. UPDATE!!!----

dates_01 <- c("1/12/2017","1/04/2018", "1/05/2018", "1/06/2018", "1/07/2018", "1/08/2018", "1/09/2018", "1/10/2018", "1/11/2018", "1/12/2018","1/01/2019", "1/02/2019", "1/03/2019", "1/04/2019", "1/05/2019", "1/06/2019", "1/07/2019", "1/08/2019", "1/09/2019", "1/10/2019", "1/11/2019", "1/12/2019", "1/01/2020", "1/02/2020", "1/03/2020", "1/04/2020", "1/05/2020", "1/06/2020", "1/07/2020", "1/08/2020", "1/09/2020", "1/10/2020", "1/11/2020", "1/12/2020", "1/01/2021", "1/02/2021", "1/03/2021", "1/04/2021", "1/05/2021", "1/06/2021", "1/07/2021", "1/08/2021", "1/09/2021", "1/10/2021", "1/11/2021", "1/12/2021", "1/01/2022", "1/02/2022", "1/03/2022", "1/04/2022", "1/05/2022", "1/06/2022", "1/07/2022", "1/08/2022", "1/09/2022")

dates_01 <- as.Date(dates_01, "%d/%m/%Y")

####____________________________________________________________________________
### extract-vector-from-list-in-r https://stackoverflow.com/questions/49778153/extract-vector-from-list-in-r
### Exploring data. Quality check----

sheets_01 <- as.vector(sapply(list_of_sheets_01, "[[", 1)) ### Not being used
sheets_02 <- as.vector(sapply(list_of_sheets_01, "[[", 2)) ### Good, they all match.
sheets_03 <- as.vector(sapply(list_of_sheets_01, "[[", 3)) ### Good, they all match.
sheets_04 <- as.vector(sapply(list_of_sheets_01, "[[", 4)) ### Needs to be corrected...
sheets_04_c <- sheets_04[6:n_files] ### Corrected only 48 files. It starts from 2018.08_LNG Report Excel File August 2018.xlsx...
sheets_05 <- as.vector(sapply(list_of_sheets_01, "[[", 5)) ### Needs to be corrected... 
sheets_05_c <-append(sheets_04[1:5],sheets_05[6:n_files]) ### Corrected: Sheet 4 became Sheet 5. "Table 4 LNG imports" added to the list of "Table 5 LNG imports".
sheets_06 <- as.vector(sapply(list_of_sheets_01, "[", 6)) ### Not being used

####### Tables 1-2 LNG exports (1) ----

####____________________________________________________________________________
## Read all "Tables 1-2 LNG exports" data ----
data_lng_exports <- list()
for(i in 1:n_files) {
  exports_data <- readxl::read_excel(paste0(file_path_01,"\\",file_names_01[i]), sheet = sheets_02[i])
  exports_data$date <- dates_01[i]
  exports_data$file_name <- file_names_01[i]
  data_lng_exports[[i]] <- exports_data
}
####____________________________________________________________________________
### Function to get TABLE 1 AUSTRALIAN SHIPMENTS

### Custom function to clean AUSTRALIAN SHIPMENTS
shipment <- function(df){
  which(df[,1]== "PROJECT")[1]
  which(df[,1]== "TOTAL")[1]
  df <- df[(which(df[,1]== "PROJECT")[1]+1):(which(df[,1]== "TOTAL")[1]-1),c(1:3,(length(df)-1),length(df))] 
  ## Add generic column names
  name_02 <- paste0("column",1:length(df))
  colnames(df) <- paste0("column_", 1:length(df))
  ## Change columns with numbers from character to numeric
  df[names(df[2:3])] <- purrr::map_df(df[2:3], as.numeric)
  return(df)
}
# shipment(df_test)

### Prepare to export to excel----
data_shipment <- lapply(data_lng_exports, shipment)
data_shipment_01 <- do.call("rbind",data_shipment)
colnames(data_shipment_01) <- c("PROJECT", "CARGOES","Quantity ('000 tonnes)", "Date", "File name")

####____________________________________________________________________________

### Custom function to clean SPOT CARGOES----
spot_cargoes <- function(df_01){ 
  c <- unlist(df_01[,1])
  a = stringr::str_detect(c,"GLADSTONE SPOT CARGOES")
  a = which(a)+1
  b = a + 2
  df_01 <- df_01[a:b,c(1:2,(length(df_01)-1),length(df_01))] 
  ## Add generic column names
  name_02 <- paste0("column",1:length(df_01))
  colnames(df_01) <- paste0("column_", 1:length(df_01))
  ## Change columns with numbers from character to numeric
  df_01[,2] <- purrr::map_df(df_01[,2], as.numeric)
  df_01[3,1] <- "GLNG"
  return(df_01)
}
# spot_cargoes(df_test)

### Select only files with spot cargo data
data_lng_exports_7_53 <- list()
for(i in 7:n_files) {
  data_lng_exports_7_53[[i]] <- spot_cargoes(data_lng_exports[[i]])
}

### rbind all data
data_lng_spot <- do.call("rbind",data_lng_exports_7_53)
colnames(data_lng_spot) <- c("PROJECT", "CARGOES", "Date", "File name")

#   lapply(data_lng_exports, `[`, c(7:53))
# sapply(data_lng_exports, "[", c(7:53))
# data_spot_cargoes <- lapply(data_lng_exports_7_53, spot_cargoes)

####____________________________________________________________________________

"TABLE 2 AUSTRALIAN DELIVERIES"
### Custom function to clean TABLE 2 AUSTRALIAN DELIVERIES----
au_delivery <- function(df){ 
  c <- unlist(df[,1])
  uniq_01 <- which(stringr::str_detect(c,"PROJECT:"))
  proj <- unique(unlist(df[uniq_01,2:(length(df) - 5)]))
  proj <- proj[!is.na(proj)]
  n_proj <- length(proj)
  a = stringr::str_detect(c,"PROJECT:")
  a = which(a) + 2
  df_sub <- df[a:nrow(df),]
  c_01 <- unlist(df_sub[,1])
  b = which(stringr::str_detect(c_01,"TOTAL")) - 1
  df <- df_sub[1:b,c(1:(length(df_sub) - 4),(length(df_sub) - 1),length(df_sub))]
  ## Add generic column names
  name_02 <- paste0("column",1:length(df))
  colnames(df) <- paste0("column_", 1:length(df))
  ### Columns positions 
  col_num <- as.numeric(c("4","7","10","13","16","19","22","25","28","31","34","37","40","43"))
  for(i in 1:n_proj) {
    df[,col_num[i]] <- proj[i]
  }
  list_01 <- list()
  for(i in 1:n_proj) {
    col_num_01 <- as.numeric(c("2","5","8","11","14","17","20","23","26","29","32","35","38","41","44","47","50","53"))
    list_01[[i]] <- df[,c(1,col_num_01[i]:col_num[i],(length(df) - 1),length(df))]
    colnames(list_01[[i]]) <- paste0("column_", 1:length(df[,c(1,col_num_01[i]:col_num[i],(length(df) - 1),length(df))]))
  }
  ### rbind all data
  df <- do.call("rbind",list_01)
  # Change columns with numbers from character to numeric
  df[,2:3] <- purrr::map_df(df[,2:3], as.numeric)
  return(df)
}
# au_delivery(data_lng_exports[[37]])
# df <- data_lng_exports[[53]]

### Select only files with spot cargo data
data_lng_dest <- list()
for(i in 1:n_files) {
  data_lng_dest[[i]] <- au_delivery(data_lng_exports[[i]])
}

### rbind all data
data_lng_dest <- do.call("rbind",data_lng_dest)
colnames(data_lng_dest) <- c("DESTINATION", "CARGOES","Quantity ('000 tonnes)","PROJECT", "Date", "File name")
####___________________________________________________________________________


###### Table 5 LNG imports (4) ----
####____________________________________________________________________________
## Read all "Table 5 LNG imports" data ----
data_lng_imports <- list()
for(i in 1:n_files) {
  imports_data <- readxl::read_excel(paste0(file_path_01,"\\",file_names_01[i]), sheet = sheets_05_c[i], col_names = FALSE)
  imports_data$date <- dates_01[i]
  imports_data$file_name <- file_names_01[i]
  data_lng_imports[[i]] <- imports_data
}
####____________________________________________________________________________
### Function to get Table 5 LNG imports----

lng_imports <- function(df){ 
  c <- unlist(df[1,2:(length(df) - 2)]) # to remove last two columns with date and file name
  # period <- unique(c)
  period <- c[!is.na(c)]
  state_01 <- unlist(df[4,2:(length(df) - 4)]) # to remove last 4 columns with date and file name
  state_01 <- state_01[!is.na(state_01)]
  n_state <- length(state_01)
  a = stringr::str_detect(unlist(df[,3]),"000t")
  a = which(a) + 1
  b = which(stringr::str_detect(unlist(df[,1]),"Total"))
  b = b[1] - 1
  df <- df[a:b,c(1, 3:(length(df) - 4), (length(df) - 1), length(df))]
  df <- df[, colMeans(is.na(df)) != 1] # https://stackoverflow.com/questions/2643939/remove-columns-from-dataframe-where-all-values-are-na
  l_df <- length(df)
  ## Add generic column names
  name_02 <- paste0("column",1:length(df))
  colnames(df) <- paste0("column_", 1:length(df))
  list_01 <- list()
  for(i in 1:n_state) {
    df_01 <- df[,c(1,(2*i),(2*i+1),(l_df - 1),l_df)]
    df_01$state <- state_01[i]
    df_01$p_date <- period[i]
    df_01 <- df_01[,c(1:3,6,7,4,5)]
    colnames(df_01) <- paste0("column_", 1:length(df_01))
    list_01[[i]] <- df_01
  }
  ### rbind all data
  df <- do.call("rbind",list_01)
  # Change columns with numbers from character to numeric
  df[,2:3] <- purrr::map_df(df[,2:3], as.numeric)
  return(df)
}

# lng_imports(data_lng_imports[[37]])

### Select only files with spot cargo data
data_lng_imp <- list()
for(i in 1:n_files) {
  data_lng_imp[[i]] <- lng_imports(data_lng_imports[[i]])
}

### rbind all data
data_lng_imp <- do.call("rbind",data_lng_imp)
colnames(data_lng_imp) <- c("Country", "000t","US$/Mmbtu","Country_import", "Period", "Date", "File name")

#### Change period to date format
library(zoo)
# month <- data_lng_imp[1,5]
# zoo::as.Date(as.yearmon(month))
data_lng_imp$Period <- zoo::as.Date(as.yearmon(data_lng_imp$Period))


###### Table 4 NEM generation (3) ----
####____________________________________________________________________________
## Read all "Table 4 NEM generation" data ----

data_nem <- list()
for(i in 6:n_files) {
  data_nem_01 <- readxl::read_excel(paste0(file_path_01,"\\",file_names_01[i]), sheet = sheets_04[i], range = "A1:H130", col_names = FALSE)
  data_nem_01$date <- dates_01[i]
  data_nem_01$file_name <- file_names_01[i]
  data_nem[[i]] <- data_nem_01
}

####____________________________________________________________________________
### Function to get Table 4 NEM generation----

lng_nem_fuel <- function(df){ 
  df <- df[rowSums(is.na(df)) != ncol(df[,1:8]), ]
  which(stringr::str_detect(unlist(df[,1]),"Interconnector flows"))[1]
  which(stringr::str_detect(unlist(df[,1]),"Power drawn from storage in battery"))[1]
  # df <- zoo::na.locf(df[,c(1,2,5,9,10)])
  df_1 <- zoo::na.locf(df[which(stringr::str_detect(unlist(df[,1]),"NSW"))[1]:(which(stringr::str_detect(unlist(df[,1]),"East coast")) - 1),c(1,2,5,9,10)])
  df_2 <- df[which(stringr::str_detect(unlist(df[,1]),"Interconnector flows"))[1]:(which(stringr::str_detect(unlist(df[,1]),"Power drawn from storage in battery"))[1] - 1), c(1,2,5,9,10)]
  df_2[-1,2] <- "Interconnector flows"
  df_2 <- df_2[-1,]
  df_1 <- rbind(df_1,df_2)
  #### df[which(stringr::str_detect(unlist(df[,1]),"Power drawn from storage in battery"))[1],c(1,2,5,9,10)] ## to add battery data
  df <- df_1
  df[,3] <- purrr::map_df(df[,3], as.numeric)
  colnames(df) <- paste0("column_", 1:length(df))
  ## Change Wind1 to just Wind
  df[,2][df[,2] == 'Wind1'] <- 'Wind'
  return(df)
}
#### Test function
# lng_nem_fuel(data_nem[[6]])

### Select only files with fuel
data_nem_fuel <- list()
for(i in 6:n_files) {
  data_nem_fuel[[i]] <- lng_nem_fuel(data_nem[[i]])
}

### rbind all data
data_nem_fuel <- do.call("rbind",data_nem_fuel)
colnames(data_nem_fuel) <- c("State", "Fuel_type_or_flow","GWh", "Date", "File_name")


# df <- data_nem[[53]]
lng_nem_station <- function(df){ 
  # df <- df[rowSums(is.na(df)) != ncol(df[,1:8]), ]
  df <- df[which(stringr::str_detect(unlist(df[,1]),"NEM gas use  by station"))[1]:nrow(df),]
  df <- df[which(stringr::str_detect(unlist(df[,1]),"NSW"))[1]:(which(stringr::str_detect(unlist(df[,1]),"Total"))[1] - 1),]
  df[,1] <- zoo::na.locf(df[,1])
  df <- df[,c(1,2,5,9,10)]
  df[,3] <- purrr::map_df(df[,3], as.numeric)
  colnames(df) <- paste0("column_", 1:length(df))
  return(df)
}
#### Test function
# lng_nem_station (data_nem[[6]])

### Select only files with fuel
data_nem_station <- list()
for(i in 6:n_files) {
  data_nem_station[[i]] <- lng_nem_station(data_nem[[i]])
}

### rbind all data
data_nem_station <- do.call("rbind",data_nem_station)
colnames(data_nem_station) <- c("State", "Station","GWh", "Date", "File_name")

####____________________________________________________________________________
### Append new file ----

# ### Update file name----
# new_file_name <- ""
# 
# new_file_path <- paste0("\\update path to \\Monthly Spreadsheets\\", new_file_name) 
# 
# exports_data$date <- as.Date("1/08/2022", "%d/%m/%Y")
# exports_data$file_name <- new_file_name



####____________________________________________________________________________
### Write to excel----

export_list <- list(data_shipment_01,
                    data_lng_dest,
                    data_lng_spot,
                    data_lng_imp,
                    data_nem_fuel,
                    data_nem_station)
names(export_list) <- c("data_shipment", "lng_dest", "lng_spot", "lng_import", "nem_fuel", "nem_station")
openxlsx::write.xlsx(export_list, "EQ_Monthly_Data.xlsx")



