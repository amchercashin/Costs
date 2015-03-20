library(openxlsx)
library(lubridate)
library(dplyr)
library(ggplot2)
library(scales)  
library(grid)

space <- function(x, ...) {format(x, ..., big.mark = " ", scientific = FALSE, trim = TRUE)}

bp <- read.xlsx("КонтрольФБ.xlsx", sheet = 18)
f <- read.xlsx("КонтрольФБ.xlsx", sheet = 19)

bp <- tbl_df(data.frame(companygroup=as.factor(bp[,3]), company=as.factor(bp[,17]), fun=as.factor(bp[,13]),
                        pl=as.factor(bp[,4]), sz=bp[,12], month=bp[,14], value=bp[,15], ico=as.logical(bp[,22])))
bp$scenario <- "Бюджет"
bp$value <- bp$value * 1000

f <- tbl_df(data.frame(companygroup=as.factor(f[,11]), company=as.factor(f[,1]), fun=as.factor(f[,13]),
                       pl=as.factor(f[,16]), sz=f[,8], month=f[,15], value=f[,10], ico=as.logical(f[,18])))
f$scenario <- "Факт"

costs <- rbind(bp, f)
costs <- costs[!is.na(costs$value),]
costs$scenario <- factor(costs$scenario)
costs$month <- dmy("01.01.1900") + days(costs$month) - days(2)
costs$plblock <- strtrim(costs$pl,2)
costs <- filter(costs, plblock != "OE" & plblock != "FE")
costs <- filter(costs, ico == 0)

costs <- group_by(costs, scenario, companygroup, sz, month) %>% summarise(value = sum(value))
costs$value[costs$value<=1]  <- 1
#options(scipen=999)
png("test.png", width=2339*2, height=3308*2)
p <- ggplot(costs, aes(x = month, y = value/1000)) +
  facet_wrap(~ companygroup) +
  #geom_point(aes(colour = scenario), size = 8, alpha = 0.6) +
  geom_line(aes(group = interaction(sz, scenario), colour = scenario), size = 1, alpha = 0.4) +
  #geom_line(data=filter(costs, scenario == "Факт"), aes(group = sz), colour="black", size = 1, alpha = 0.4) +
  scale_y_log10(labels = space, limits = c(1, NA)) +
  ylab("Тысяч рублей") +
  xlab("2014") +
  ggtitle("Постоянные затраты по статьям затрат\n Факт 2014 vs Бюджет 2014") +
  theme_bw(base_size = 36) +
  theme(legend.position="top", legend.title=element_blank(), legend.key.size = unit(2,"cm")) +
  theme(axis.text.x=element_text(size=36), axis.title.y=element_text(size=36), strip.text.x=element_text(size=40), legend.text=element_text(size=40)) +
  theme(axis.text.x  = element_text(hjust=1,angle=90, vjust=0.5, size=32), axis.title.x = element_blank()) +
  scale_x_datetime(labels = date_format("%b"), breaks = date_breaks("month")) +
  scale_colour_manual(values = c("red", "black"))
print(p)
dev.off()