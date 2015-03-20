library(openxlsx)
library(lubridate)
library(dplyr)
library(ggplot2)
library(scales)  
library(grid)
library(gridExtra)

#Тест служебной функции для вывода рашифровок - функция проверки отклонени
dF <- function(f, g) {
        (sum(filter(costs, fun == f, companygroup == g, scenario == "Факт") 
                           %>% select(value)) -  
                               sum(filter(costs, fun == f, companygroup == g, scenario == "Бюджет") 
                                   %>% select(value))) /
                              sum(filter(costs, fun == f, companygroup == g, scenario == "Бюджет") 
                                  %>% select(value))
}

space <- function(x, ...) {format(x, ..., big.mark = " ", scientific = FALSE, trim = TRUE)}
# ЗАГРУЗКА ДАННЫХ ################
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

# ОЧИСТКА ДАННЫХ ##################
costs <- costs[!is.na(costs$value),]
costs <- costs[!is.na(costs$companygroup),]
costs$scenario <- factor(costs$scenario)
costs$month <- dmy("01.01.1900") + days(costs$month) - days(2)
costs$plblock <- strtrim(costs$pl,2)
costs <- filter(costs, plblock == "GE" | plblock == "CS"| plblock == "CE")
costs <- filter(costs, fun != 0 & ico == 0)

costs$sz <- as.character(costs$sz)
# Убираем амортизацию
costs <- costs[!grepl("аморт", costs$sz, ignore.case = T), ]
# Убираем переменные коммерческие
costs <- filter(costs, pl != "CE-02-01-00" & pl != "CE-02-03-00" & pl != "CE-02-04-00" & pl != "CE-02-05-00"
                & pl != "CE-04-09-00" & pl != "CE-03-01-00" & pl != "CE-05-00-00" & pl != "CE-02-02-00")
# Убираем HR (не проч)
costs <- filter(costs, pl != "CE-07-01-00" & pl != "CE-07-02-00" & pl != "CE-07-03-00" & pl != "CE-07-09-00"
                & pl != "CE-07-02-02" & pl != "CE-07-04-00" & pl != "CE-07-05-00" & pl != "CE-07-06-00"
                & pl != "CE-09-08-00" & pl != "CE-10-02-02" & pl != "GE-03-01-00" & pl != "GE-03-02-00"
                & pl != "GE-03-03-00" & pl != "GE-03-09-00" & pl != "GE-03-02-02" & pl != "GE-03-04-00"
                & pl != "GE-03-05-00" & pl != "GE-03-06-00" & pl != "GE-09-08-00" & pl != "GE-06-03-02"
                & pl != "CS-09-01-00" & pl != "CS-09-02-00" & pl != "CS-09-03-00" & pl != "CS-09-09-00"
                & pl != "CS-09-02-02" & pl != "CS-09-04-00" & pl != "CS-09-05-00" & pl != "CS-09-06-00"
                & pl != "CS-09-08-00" & pl != "CS-12-01-02" )
# Убираем налоги (но не аналоги)
costs <- costs[!grepl("налог", costs$sz, ignore.case = T) | grepl("аналог", costs$sz, ignore.case = T), ]
# Убираем производственную переменку
costs <- costs[!grepl("CS-03", costs$pl, ignore.case = T), ]
costs <- costs[!grepl("CS-04", costs$pl, ignore.case = T), ]

# СОЗДАНИЕ ИТОГО И ДАННЫХ ПРО ПРОЧИМ КОМПАНИЯМ #######
# Отдельный датасет для Прочих по компаниям
costs.p <- filter(costs, companygroup == "Прочие")
costs.p <- costs.p[!is.na(costs.p$company),]
costs.p$company <- as.character(costs.p$company)
costs.p$company[costs.p$company == 'ООО "УРАЛХИМ Консалт"'] <- 'ООО УРАЛХИМ Консалт'
costs.p$company <- factor(costs.p$company)
costs.p <- group_by(costs.p, scenario, company, fun, month) %>% summarise(value = sum(value))
costs.p2 <- group_by(costs.p, scenario, compan, month) %>% summarise(value = sum(value)) %>% mutate(fun = "ИТОГО")
costs.p <- bind_rows(costs.p, costs.p2)
rm(costs.p2)

setFunFactorLvls <- function(myFactorVar) {
        factor(myFactorVar, levels = c("Производственная функция", "Функция ИТ", "Коммерческая функция", 
                                       "Транспортная функция", "Юридическая функция", "Функция развития бизнеса",
                                       "Административная функция", "Финансовая функция",
                                       "Функция по связям с общественностью", "Охрана и безопасность",
                                       "Функция по работе с персоналом", "ИТОГО"))
}

costs.p$fun <- setFunFactorLvls(costs.p$fun)

costs <- group_by(costs, scenario, companygroup, fun, month) %>% summarise(value = sum(value))
costs2 <- group_by(costs, scenario, companygroup, month) %>% summarise(value = sum(value)) %>% mutate(fun = "ИТОГО")
costs <- bind_rows(costs, costs2)
rm(costs2)

costs$fun <- setFunFactorLvls(costs$fun)

#options(scipen=999)
# СОЗДАНИЕ КАРТИНОК #################
makeslides <- function(group) {
        png(paste0("slides/",group,".png"), width=3308, height=2339)
        p <- ggplot(filter(costs, companygroup == group), aes(x = month, y = value/1000)) +
                facet_wrap(~ fun, scales = "free_y", drop = FALSE) +
                geom_point(aes(colour = scenario), size = 8, alpha = 0.6) +
                geom_line(aes(colour = scenario), size = 2, alpha = 0.6) +
                scale_y_continuous(labels = space) +
                expand_limits(y = 0) + 
                ylab("Тысяч рублей") +
                #xlab("2014") +
                ggtitle(paste0(group,"\nПостоянные затраты по функциям управления\n Факт 2014 vs Бюджет 2014")) +
                theme_bw(base_size = 36) +
                theme(legend.position="top", legend.title=element_blank(), legend.key.size = unit(2,"cm")) +
                theme(axis.text.x=element_text(size=36), axis.title.y=element_text(size=36), strip.text.x=element_text(size=34), 
                      legend.text=element_text(size=40)) +
                theme(axis.text.x  = element_text(hjust=1,angle=90, vjust=0.5, size=32), axis.title.x = element_blank()) +
                scale_x_datetime(labels = date_format("%b"), breaks = date_breaks("month")) 
        
        p <- arrangeGrob(p, sub = textGrob("* В целя сопоставимости исключены данные по\n  Прочим операционным доходам/(расходам) и\n  Финансовым доходам/(расходам), а так же\n  переменные коммерческие, аренда ж/д\n  транспорта, HR (не проч.), налоги\n  переменные производственные"
                                           , x = 0.8, y = 13.4, hjust = 0, vjust = 0, just = c("left", "top"),
                                           gp = gpar(fontface = "italic", fontsize = 20)))  
        print(p)
        
        dev.off()  
}

#Картинки для Прочих компаний
makeslides.p <- function(comp) {
        png(paste0("slides/Прочие/",comp,".png"), width=3308, height=2339)
        p <- ggplot(filter(costs.p, company == comp), aes(x = month, y = value/1000)) +
                facet_wrap(~ fun, scales = "free_y", drop = FALSE) +
                geom_point(aes(colour = scenario), size = 8, alpha = 0.6) +
                geom_line(aes(colour = scenario), size = 2, alpha = 0.6) +
                scale_y_continuous(labels = space) +
                expand_limits(y = 0) + 
                ylab("Тысяч рублей") +
                #xlab("2014") +
                ggtitle(paste0(comp,"\nПостоянные затраты по функциям управления\n Факт 2014 vs Бюджет 2014")) +
                theme_bw(base_size = 36) +
                theme(legend.position="top", legend.title=element_blank(), legend.key.size = unit(2,"cm")) +
                theme(axis.text.x=element_text(size=36), axis.title.y=element_text(size=36), strip.text.x=element_text(size=30), 
                      legend.text=element_text(size=40)) +
                theme(axis.text.x  = element_text(hjust=1,angle=90, vjust=0.5, size=32), axis.title.x = element_blank()) +
                scale_x_datetime(labels = date_format("%b"), breaks = date_breaks("month")) 
        
        p <- arrangeGrob(p, sub = textGrob("* В целя сопоставимости исключены данные по\n  Прочим операционным доходам/(расходам) и\n  Финансовым доходам/(расходам), а так же\n  переменные коммерческие, аренда ж/д\n  транспорта, HR (не проч.), налоги\n  переменные производственные"
                                           , x = 0.8, y = 13.4, hjust = 0, vjust = 0, just = c("left", "top"),
                                           gp = gpar(fontface = "italic", fontsize = 20)))  
        print(p)
        
        dev.off()  
}

groupList <- c("Группа АЗОТ", "Группа ВМУ", "Группа КЧХК", "ОХК", "ПМУ", "Прочие", "ТД Латвия", "ТД Уралхим", "УХТ",
               "УХФ", "RFT")
compList <- c("HAVENPORT INVESTMENTS LIMITED", "Sanders Enterprises", "Uralchem trading DO brasil Ltd",
              "URALCHEMAssist GmbH", "ООО УРАЛХИМ Консалт", "ООО Карбин", "СПК Богородский",
               "ХимПроект, ООО", "LELADA ENTERPRISES LIMITED")
lapply(groupList, function(group) {makeslides(group)})
lapply(compList, function(comp) {makeslides.p(comp)})
