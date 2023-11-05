# renv
install.packages("renv")

library(renv)
renv::init()

# Packages
install.packages("tidyverse")
library(tidyverse)

install.packages("data.table")
library(data.table)

install.packages("openxlsx")
library(openxlsx)

install.packages("magrittr")
library(magrittr)

install.packages("stringi")
library(stringi)

install.packages("readxl")
library(readxl)

install.packages("tidyr")
library(tidyr)

install.packages("dplyr")
library(dplyr)

install.packages("lubridate")
library(lubridate)

# install.packages("writexl")
# library(writexl)

# renv.lock
renv::snapshot()

options(scipen = 99999)