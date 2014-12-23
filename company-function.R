library(openxlsx)
library(lubridate)
library(dplyr)
library(ggplot2)
library(scales)  

space <- function(x, ...) {format(x, ..., big.mark = " ", scientific = FALSE, trim = TRUE)}

c <- read.xlsx("ФБ2015_Консолидация.xlsX", sheet = 8)
costs <- tbl_df(data.frame(companygroup=as.factor(c[,11]), company=as.factor(c[,1]), fun=as.factor(c[,13]),
                           pl=as.factor(c[,16]), month=c[,15], value=c[,10], ico=as.logical(c[,19])))
costs$month <- dmy("01.01.1900") + days(costs$month) - days(2)
costs$plblock <- strtrim(costs$pl,2)
costs <- filter(costs, fun != 0 & ico == 0)

costs <- group_by(costs, companygroup, fun, month) %>% summarise(value = sum(value))

#options(scipen=999)
png("FUNvsCOMP_BP_freeYscale.png", width=2339, height=3308)
p <- ggplot(costs, aes(x = month, y = value/1000)) +
     facet_grid(fun ~ companygroup, scales = "free_y") +
     geom_point() +
     geom_smooth(method="loess", se = F) +
     scale_y_continuous(labels = space) +
     ylab("Тысяч рублей") +
     xlab("") +
     theme_bw()
print(p)
dev.off()

png("FUNvsCOMP_BP.png", width=2339, height=3308)
p <- ggplot(costs, aes(x = month, y = value/1000)) +
  facet_grid(fun ~ companygroup) +
  geom_point() +
  geom_smooth(method="loess", se = F) +
  scale_y_continuous(labels = space) +
  ylab("Тысяч рублей") +
  xlab("") +
  theme_bw()
print(p)
dev.off()