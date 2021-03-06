library(openxlsx)
library(lubridate)
library(dplyr)
library(ggplot2)
library(scales)  

space <- function(x, ...) {format(x, ..., big.mark = " ", scientific = FALSE, trim = TRUE)}

c <- read.xlsx("ФБ2015_Консолидация_все_затраты.xlsx", sheet = 8)
costs <- tbl_df(data.frame(companygroup=as.factor(c[,11]), company=as.factor(c[,1]), fun=as.factor(c[,13]),
                           pl=as.factor(c[,16]), sz=as.factor(c[,8]), month=c[,15], value=c[,10],
                           ico=as.logical(c[,19])))
costs$month <- dmy("01.01.1900") + days(costs$month) - days(2)
costs$plblock <- strtrim(costs$pl,2)
costs <- filter(costs, fun != 0 & ico == 0)

costs <- group_by(costs, companygroup, fun, month, plblock) %>% summarise(value = sum(value))
costs$pn <- factor(sign(costs$value), labels=c("-", "+"))

costs <- na.omit(costs)

png("FUNvsCOMP_BAR_BP15_freeYscale.png", width=2339, height=3308)
p <- ggplot(costs, aes(x = month, y = value/1000, fill = plblock)) +
  facet_grid(fun ~ companygroup, scales = "free_y") +
  geom_bar(stat = "identity", data = subset(costs, pn == "+"), alpha = 0.7) +
  geom_bar(stat = "identity", data = subset(costs, pn == "-"), alpha = 0.7) +
  geom_hline(yintercept=0) +
  #geom_smooth(method="loess", se = F) +
  scale_y_continuous(labels = space) +
  ylab("Тысяч рублей") +
  xlab("2015") +
  theme_bw() +
  theme(legend.position="top") +
  theme(axis.text.x  = element_text(hjust=1,angle=90, vjust=0.5)) +
  scale_x_datetime(labels = date_format("%B"), breaks = date_breaks("month")) +
#  scale_fill_discrete(name="Вид расходов",
#                     breaks=c("CS", "CE", "GE", "OE"),
#                      labels=c("Производственные", "Коммерческие", "Административные", "Финансовые"))
#scale_fill_hue()
scale_fill_manual(values=c("CS"="FireBrick", "CE"="LimeGreen", "GE"="DarkTurquoise", "OE"="DimGray", "FE"="DarkOrchid"), 
                  name="Вид расходов",
                  breaks=c("CS", "CE", "GE", "OE", "FE"),
                  labels=c("Производственные ", "Коммерческие ", "Административные ", "Прочие доходы/расходы ", "Финансовые"))
print(p)
dev.off()

png("FUNvsCOMP_BAR_BP15.png", width=2339, height=3308)
p <- p + facet_grid(fun ~ companygroup, scales = "fixed")
print(p)
dev.off()