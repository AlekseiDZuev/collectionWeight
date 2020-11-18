library(lubridate)
library(readxl)
library(xlsx)

#Нормативы по категориям не установлены в 1с. Прописываем вручную. Свободностоящие 76000
#, но разделяются по Мыльницам, Стаканам, Дозаторам, Ёршикам
normColl<-cbind(c("Мягкие коврики", "Занавески", "Свободностоящие аксессуары"),c(14000, 30000, 76000))
Softrugs<-0
Curtains<-0
freeSt<-cbind(c("Мыльницы", "Стаканы для зубных щеток", "Дозаторы", "Ершики" ),c(19000, 19000, 19000, 19000))
Soaps<-0
Glasses<-0
Dispensers<-0
Brushes<-0
#Здесь загружаю файл с данными для обработки. Нужно исправить на SQL запрос к базе 1с
collections <- read_excel("./Для расчёта.xlsx")
#Смотрю коллекции и артикулы. Рассчитываю среднемесячную валовую прибыль по артикулам.
unColl<-unique(collections$Коллекция)
averGross<-0
uNom<-unique(collections$Артикул)
for(i in 1:length(uNom)){
  tradeMonths<-11-month(as.Date(min(collections[collections$Артикул==uNom[i],]$Месяц), "%d.%m.%Y" ))
  gross<-sum(collections[collections$Артикул==uNom[i],]$`Валовая прибыль`)
  averGross[i]<-gross/tradeMonths
}
averGross<-cbind(uNom, averGross)
weight<-0
weightColl<-0
#Цикл по коллециям для расчёта веса ВП
for(i in 1:length(unColl)){
  Softrugs[i]<-0
  Curtains[i]<-0
  Soaps[i]<-0
  Glasses[i]<-0
  Dispensers[i]<-0
  Brushes[i]<-0
  weight[i]<-0
  weightColl[i]<-0
  coll<-collections[collections$Коллекция==unColl[i],]
  #Цикл по категория для расчёта веса ВП
  for(j in 1:length(normColl[,1])){
    #Цикл по "Мягкие коврики", "Занавески"
    if(normColl[j,1]!="Свободностоящие аксессуары"){
      if(isFALSE(identical(unique(coll$`Товарная категория`)[unique(coll$`Товарная категория`)==normColl[j,1]],character(0)))){
        weight[i]<-weight[i]+as.numeric(normColl[j,2])
        tempWeight<-0
        #Считаю средний вес в категории коллекции
        for(a in 1:length(unique(coll[coll$`Товарная категория`==normColl[j,1],]$Артикул))){
          tempWeight<-tempWeight+as.numeric(averGross[averGross[,1]==unique(coll[coll$`Товарная категория`==normColl[j,1],]$Артикул)[a],][2])
        }
        weightColl[i]<-weightColl[i]+tempWeight/length(unique(coll[coll$`Товарная категория`==normColl[j,1],]$Артикул))
        if(j==1){
          Softrugs[i]<-tempWeight/length(unique(coll[coll$`Товарная категория`==normColl[j,1],]$Артикул))
        } else {
          Curtains[i]<-tempWeight/length(unique(coll[coll$`Товарная категория`==normColl[j,1],]$Артикул))
        }
      }
    }
  }
      #Цикл по "Мыльницы", "Стаканы для зубных щеток", "Дозаторы", "Ершики"
      for(j in 1:length(freeSt[,1])){
        if(isFALSE(identical(unique(coll$Свободностоящие)[unique(coll$Свободностоящие)==freeSt[j,1]],character(0)))){
          weight[i]<-weight[i]+as.numeric(freeSt[j,2])
          tempWeight<-0
          #Считаю средний вес в категории коллекции
          for(a in 1:length(unique(coll[coll$Свободностоящие==freeSt[j,1],]$Артикул))){
            tempWeight<-tempWeight+as.numeric(averGross[averGross[,1]==unique(coll[coll$Свободностоящие==freeSt[j,1],]$Артикул)[a],][2])
          }
          weightColl[i]<-weightColl[i]+tempWeight/length(unique(coll[coll$Свободностоящие==freeSt[j,1],]$Артикул))
          if(j==1){
            Soaps[i]<-tempWeight/length(unique(coll[coll$Свободностоящие==freeSt[j,1],]$Артикул))
          } else {
            if(j==2){
              Glasses[i]<-tempWeight/length(unique(coll[coll$Свободностоящие==freeSt[j,1],]$Артикул))
            } else {
              if (j==3){
                Dispensers[i]<-tempWeight/length(unique(coll[coll$Свободностоящие==freeSt[j,1],]$Артикул))            
              }else{
                Brushes[i]<-tempWeight/length(unique(coll[coll$Свободностоящие==freeSt[j,1],]$Артикул))
              }
            }
          }
    }
  }
}
#Записываю в файл
write.xlsx2(cbind(unColl, weight, weightColl, Softrugs, Curtains, Soaps, Glasses, Dispensers, Brushes), file='Weight.xlsx', sep=",")
write.xlsx2(averGross, file='aver.xlsx', sep=",")

