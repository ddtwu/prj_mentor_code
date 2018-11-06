#-------------------------------------------------------#
# Install package if not exists
pkgs <- c('data.table', 'plyr', 'dplyr', 'reshape2', 'tidyr', # 資料處理
          'stringr', 'stringi', # 文字處理
          'lubridate', # 日期變數處理
          'magrittr', # pipe
          'readxl', # excel的讀取
          'xlsx', 'writexl',  # excel的回寫
          'readr', # csv的讀取
          'ggplot2', 'scales', 'gridExtra', # 視覺化
          'DescTools', # 提供一些好用的函式
          'Matrix' # 稀疏矩陣
)


# pkgs <- c('tidyverse')
new.pkgs <- pkgs[!(pkgs %in% installed.packages())]
if (length(new.pkgs)) {
  install.packages(new.pkgs, repos = 'http://cran.csie.ntu.edu.tw/')
}

# Require packages
suppressPackageStartupMessages(lapply(pkgs, library, character.only=T, lib.loc = .libPaths()[1]))


# install.packages('AAA')
# library(AAA)
# require(AAA)


#-------------------------------------------------------#


test1 <- read_xls(path = 'E:/Marvin.Wu/myProject/prj_mentor/中保產品分析用_R.xls', sheet = 'Sheet1')



setwd('E:/Marvin.Wu/myProject/prj_mentor')
test1 <- read_xls('中保產品分析用_R.xls', sheet = 'Sheet1')
test2 <- read_xls('中保產品分析用_R.xls', sheet = 'Sheet2')
test3 <- read_xls('中保產品分析用_R.xls', sheet = 'Sheet3')
test4 <- read_xls('中保產品分析用_R.xls', sheet = 'Sheet4')
test5 <- read_xls('中保產品分析用_R.xls', sheet = 'Sheet5')
test6 <- read_xls('中保產品分析用_R.xls', sheet = 'Sheet6')
test7 <- read_xls('中保產品分析用_R.xls', sheet = 'Sheet7')

raw_data <- bind_rows(test1, test2, test3, test4, test5, test6, test7)

# 取出器材編號與器材名稱表
prod_name <- raw_data %>% count(器材名稱)



# 有多少契約編號
raw_data %>% select(契約編號) %>% distinct()
raw_data %>% pull(契約編號) %>% unique() %>% length()
raw_data %>% count(契約編號)  

# 次數統計
tb1 <- raw_data %>% count(., 契約編號, 系統別) %>% count(., 系統別) %>% arrange(., desc(nn)) %>% mutate(prob = nn/sum(nn))





