library(readxl)
library(openxlsx)

wb <- createWorkbook()

sheet <- "JAKARTA PUSAT"
data <- read_excel("Dataset.xlsx",sheet = sheet)
head(data)
names(data)

# Periksa Data Awal
unique(as.vector(as.matrix(data[,c(4:14)])))

# Cleaning
data[, 4:14] <- lapply(data[, 4:14], as.character)

data[, 4:14] <- lapply(data[, 4:14], function(x) ifelse(x == "80 cm", "71 s.d 150 cm", x))
data[, 4:14] <- lapply(data[, 4:14], function(x) ifelse(x == "40 s.d 60 cm", "10 s.d 70 cm", x))
data[, 4:14] <- lapply(data[, 4:14], function(x) ifelse(x == "20 s.d 50 cm", "10 s.d 70 cm", x))
data[, 4:14] <- lapply(data[, 4:14], function(x) ifelse(x == "30 cm", "10 s.d 70 cm", x))

# Periksa Hasil Cleaning
unique(as.vector(as.matrix(data[,c(4:14)])))

# Edit Data Ketinggian Air
data[, 4:14] <- lapply(data[, 4:14], function(x) ifelse(is.na(x) == TRUE, 0, x))
data[, 4:14] <- lapply(data[, 4:14], function(x) ifelse(x == "10 s.d 70 cm", 1, x))
data[, 4:14] <- lapply(data[, 4:14], function(x) ifelse(x == "71 s.d 150 cm", 2, x))
data[, 4:14] <- lapply(data[, 4:14], function(x) ifelse(x == ">150 cm", 3, x))
data[, 4:14] <- lapply(data[, 4:14], function(x) as.numeric(x))

# Edit Frekuensi
data[,15] <- ifelse(data[,15] > 1, 1, 0)

addWorksheet(wb,sheet)
writeData(wb, sheet, data)

sheet <- "JAKARTA UTARA"
data <- read_excel("Dataset.xlsx",sheet = sheet)
head(data)
names(data)

# Periksa Data Awal
unique(as.vector(as.matrix(data[,c(4:15)])))

# Cleaning
data[, 4:15] <- lapply(data[, 4:15], as.character)

data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x %in% c("5 s.d 10 cm",
                                                                 "25 s.d 30 cm",
                                                                 "3 s.d 10 s.d 70 cm",
                                                                 "15 s.d 70 cm",
                                                                 "30 s.d 60 cm",
                                                                 "30 cm","40 s.d 50 cm",
                                                                 "50 cm","70 cm","60 cm",
                                                                 "'30 cm","25 cm","36 cm",
                                                                 "20 s.d 10 s.d 70 cm","20 s.d 30 cm",
                                                                 "20 s.d. 30 cm","20 s.d 25 cm",
                                                                 "50 s.d 60 cm","10 cm","10 s.d 55 cm"), "10 s.d 70 cm", x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x %in% c("10 s.d 80 cm", "40 s.d 80 cm","50 s.d 80 cm",
                                                                 "80 cm"), "71 s.d 150 cm", x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x %in% c(), ">150 cm", x))

# Periksa Hasil Cleaning
unique(as.vector(as.matrix(data[,c(4:15)])))

# Edit Data Ketinggian Air
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(is.na(x) == TRUE, 0, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x == "10 s.d 70 cm", 1, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x == "71 s.d 150 cm", 2, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x == ">150 cm", 3, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) as.numeric(x))

# Edit Frekuensi
data[,16] <- ifelse(data[,16] > 1, 1, 0)

addWorksheet(wb,sheet)
writeData(wb, sheet, data)

sheet <- "JAKARTA BARAT"
data <- read_excel("Dataset.xlsx",sheet = sheet)
head(data)
names(data)

# Periksa Data Awal
unique(as.vector(as.matrix(data[,c(4:15)])))

# Cleaning
data[, 4:15] <- lapply(data[, 4:15], as.character)

data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x %in% c("10 s.d 50 cm","60 cm",
                                                                 "40 s.d 50 cm","65 cm",
                                                                 "30 s.d 60 cm","30 cm",
                                                                 "30 s.d 50 cm","50 cm",
                                                                 "70 cm","40 s.d 60 cm","40 s.d 70 cm",
                                                                 "30 s.d 10 s.d 70 cm","'50 cm",
                                                                 "20 s.d 60 cm","25 cm","35 s.d 70 cm",
                                                                 "45 cm","20 cm","60 s.d 70 cm","30 s.d 35 cm",
                                                                 "20 s.d 70 cm","55 cm","30 s.d 55 cm","20 s.d 55 cm",
                                                                 "40 s.d. 60 cm","50 s.d 70 cm","30 s.d. 10 s.d 70 cm"), "10 s.d 70 cm", x))

data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x %in% c("100 cm", "80 cm",
                                                                 "40 s.d 100 cm","40 s.d 80 cm",
                                                                 "35 s.d 100 cm","60 s.d 100 cm",
                                                                 "95 cm","60 s.d 80 cm","50 s.d 100 cm",
                                                                 "90 cm","30 s.d 80 cm","20 s.d 110 cm",
                                                                 "20 s.d 80 cm","20 s.d 120 cm",
                                                                 "20 s.d 90 cm","30 s.d 100 cm",
                                                                 "80 s.d 110 cm","31 s.d 100 cm",
                                                                 "40 s.d. 110 cm"), "71 s.d 150 cm", x))
# Periksa Hasil Cleaning
unique(as.vector(as.matrix(data[,c(4:15)])))

# Edit Data Ketinggian Air
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(is.na(x) == TRUE, 0, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x == "10 s.d 70 cm", 1, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x == "71 s.d 150 cm", 2, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x == ">150 cm", 3, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) as.numeric(x))

# Edit Frekuensi
data[,16] <- ifelse(data[,16] > 1, 1, 0)

addWorksheet(wb,sheet)
writeData(wb, sheet, data)

sheet <- "JAKARTA SELATAN"
data <- read_excel("Dataset.xlsx",sheet = sheet)
head(data)
names(data)

# Periksa Data Awal
unique(as.vector(as.matrix(data[,c(4:15)])))

# Cleaning
data[, 4:15] <- lapply(data[, 4:15], as.character)

data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x %in% c("10 s.d 70 cm","10 s.d 35 cm","50 cm","70 cm",
                                                                 "40 s.d 70 cm","40 s.d 45 cm","60 cm",             
                                                                 "40 s.d 50 cm","40 s.d 60 cm","45 cm",
                                                                 "40","55 cm","40 s.d. 50 cm","30 cm",              
                                                                 "40 s.d. 60 cm","20 cm","35 s.d. 50 cm","35 cm",
                                                                 "30 s.d 70 cm","25 s.d. 60 cm","20 s.d 10 s.d 70 cm",
                                                                 "25 cm","40 s.d. 70 cm","50 Cm","15 s.d 30 cm","30 Cm",              
                                                                 "20 s.d 30 cm","30 s.d 10 s.d 70 cm","30 s.d 60 cm","50 s.d 70 cm",
                                                                 "50 s.d 60 cm","65 cm","25 s.d 10 s.d 70 cm","30 s.d 50 cm"), "10 s.d 70 cm", x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x %in% c("10 s.d 120 cm","10 s.d 250 cm",      
                                                                 "30 s.d 150 cm","55 s.d 100 cm","30 s.d 120 cm",
                                                                 "100 cm","40 s.d 90 cm","40 s.d 100 cm","145 cm",             
                                                                 "40 s.d 150 cm","160 cm","90 cm","100 s.d 200 cm",     
                                                                 "50 s.d 120 cm","50 s.d 80 cm","200 cm","150 cm",             
                                                                 "80 cm","40 s.d 80 cm","50 s.d 100 cm","110 s.d 70 cm",      
                                                                 "120 cm","190 cm","170 cm","110 cm",
                                                                 "120 s.d 130 cm","75 cm","25 s.d 80 cm",       
                                                                 "35 s.d 80 cm","30 s.d 100 cm",
                                                                 "60 s.d. 90 cm","50 s.d 90 cm",
                                                                 "50 s.d 110 cm","40 s.d. 80 cm",
                                                                 "70 s.d 150 cm","50 s.d 170 cm",
                                                                 "40 s.d 200 cm","120 s.d 150 cm",
                                                                 "30 s.d. 150 cm","180 cm","50 s.d 80 Cm",
                                                                 "70 s.d 80 cm","30 s.d 80 cm","20 s.d 160 cm",
                                                                 "10 s.d 160 cm","10 s.d 100 cm","60 s.d 150 cm",
                                                                 "10 s.d 130 cm"), "71 s.d 150 cm", x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x %in% c("> 150 cm"), ">150 cm", x))

# Periksa Hasil Cleaning
unique(as.vector(as.matrix(data[,c(4:15)])))

# Edit Data Ketinggian Air
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(is.na(x) == TRUE, 0, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x == "10 s.d 70 cm", 1, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x == "71 s.d 150 cm", 2, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x == ">150 cm", 3, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) as.numeric(x))

# Edit Frekuensi
data[,16] <- ifelse(data[,16] > 1, 1, 0)

addWorksheet(wb,sheet)
writeData(wb, sheet, data)

sheet <- "JAKARTA TIMUR"
data <- read_excel("Dataset.xlsx",sheet = sheet)
head(data)
names(data)

# Periksa Data Awal
unique(as.vector(as.matrix(data[,c(4:15)])))

# Cleaning
data[, 4:15] <- lapply(data[, 4:15], as.character)

data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x %in% c("70 cm","50 cm","31 s/s 70 cm","45 cm",
                                                                 "31 s/d 70 cm","50 CM","60 cm","40 s/d 70 cm",
                                                                 "10 s/d 30 cm","35 s/d 50 cm",         
                                                                 "45 s/d 70 cm","40 s/d 50 cm",        
                                                                 "20 cm","30 cm","40 s.d. 50 cm",
                                                                 "35 s.d. 60 cm","40 s.d. 70 cm",
                                                                 "25 cm","20 s.d 10 s.d 70 cm","30 Cm",                
                                                                 "10 cm","40 s.d 60 cm","35 s.d 55 cm",
                                                                 "35 cm","30 s.d 70 cm","20 s.d 35 cm"), "10 s.d 70 cm", x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x %in% c("60 s/d 75 cm","30 s/d 100 cm","100 cm",
                                                                 "150 cm","50 s/d 100 cm",
                                                                 "31 s/d 90 cm","40 s/d 100 cm","50 s/d 150 cm",
                                                                 "40 s/d 190 cm","90 s/d 110 cm","230 s/d 300 cm",
                                                                 "250 s/d 300 cm","50 s/d 250 cm","250 cm",
                                                                 "30 s/d 125 cm","40 s/d 120 cm",
                                                                 "80 cm","40 s/d 150 cm","40 s/d 80 cm","90 cm",
                                                                 "120 cm","80 s/d 120 cm","40 s.d 90 cm","'40 s.d 120 cm",       
                                                                 "40 s.d. 100 cm","40 s.d. 90 cm","25 s.d 120 cm","40 s.d 100 cm",
                                                                 "40 s.d 80 cm","30 s.d 100 cm","10 s.d 150 cm",
                                                                 "120 s.d 150 cm","30 s.d 110 s.d 70 cm","40 s/d 75 cm"), "71 s.d 150 cm", x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x %in% c("40 s/d 170 cm","40 s/d 250 cm",
                                                                 "100 s/d 250 cm","40 s/d 200 cm",
                                                                 "30 s/d 300 cm","180 cm","20 s.d. 170 cm",
                                                                 "230 s/d 300 cm","40 s.d. 110 s.d 70 cm",
                                                                 "50 s.d 160 cm","20 s.d 160 cm","40 s.d 190 cm",
                                                                 "40 s.d. 200 cm","40 s.d 200 cm","200 cm",
                                                                 "20 s.d 180 cm","30 s.d 160 cm","40 s.d 200 cm",
                                                                 "75 s.d 175 cm","10 s.d 175 cm","35 s.d 250 cm","35 s.d 230 cm",
                                                                 "30 s.d 180 cm","30 s.d 200 cm","35 s.d 200 cm","250 s/d 275 cm",
                                                                 "40 s/d 300 cm"), ">150 cm", x))

# Periksa Hasil Cleaning
unique(as.vector(as.matrix(data[,c(4:15)])))

# Edit Data Ketinggian Air
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(is.na(x) == TRUE, 0, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x == "10 s.d 70 cm", 1, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x == "71 s.d 150 cm", 2, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x == ">150 cm", 3, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) as.numeric(x))

# Edit Frekuensi
data[,16] <- ifelse(data[,16] > 1, 1, 0)

addWorksheet(wb,sheet)
writeData(wb, sheet, data)

sheet <- "PULAU SERIBU"
data <- read_excel("Dataset.xlsx",sheet = sheet)
head(data)
names(data)

# Periksa Data Awal
unique(as.vector(as.matrix(data[,c(4:15)])))

# Cleaning
data[, 4:15] <- lapply(data[, 4:15], as.character)

data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x %in% c("10 s.d 20 cm","10 s.d 50 cm","10 s.d 45 cm",
                                                                 "30 s.d 50 cm",
                                                                 "5 s.d 20 cm",
                                                                 "40 s.d 50 cm","10 s.d 15 cm",
                                                                 "10 cm","20 cm",
                                                                 "15 s.d 20 cm",
                                                                 "15 cm"), "10 s.d 70 cm", x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x %in% c("10 s.d 10 s.d 70 cm","25 s.d 10 s.d 70 cm"), "71 s.d 150 cm", x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x %in% c(), ">150 cm", x))

# Periksa Hasil Cleaning
unique(as.vector(as.matrix(data[,c(4:15)])))

# Edit Data Ketinggian Air
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(is.na(x) == TRUE, 0, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x == "10 s.d 70 cm", 1, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x == "71 s.d 150 cm", 2, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) ifelse(x == ">150 cm", 3, x))
data[, 4:15] <- lapply(data[, 4:15], function(x) as.numeric(x))

# Edit Frekuensi
data[,16] <- ifelse(data[,16] > 1, 1, 0)

addWorksheet(wb,sheet)
writeData(wb, sheet, data)

saveWorkbook(wb,"Cleaned Dataset.xlsx", overwrite = TRUE)
