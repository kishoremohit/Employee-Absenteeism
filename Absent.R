####
# Required Libraries for our Dataset
###
x <- c("ggplot2", "corrgram", "DMwR", "caret", "randomForest", "unbalanced", "C50", "dummies", "e1071", "Information",
      "MASS", "rpart", "gbm", "ROSE", 'sampling', 'DataCombine', 'inTrees','readxl')

####
# Load packages(x)
###
lapply(x, require, character.only = TRUE)
rm(x)

####
# Loading data
###
Absenteeism <- read_excel("Absenteeism.xls")
str(Absenteeism)
options(warn = -1)

####
# Missing Data analysis In Percentage
###
missing <- data.frame(apply(Absenteeism,2,function(x){sum(is.na(x))}))
missing$Columns <- row.names(missing)
names(missing)[1] <-  "Missing_PER"
missing$Missing <- (missing$Missing/nrow(Absenteeism)) * 100
missing <- missing[order(-missing$Missing),]
row.names(missing) <- NULL
missing <- missing[,c(2,1)]


####
# Plotting Histogram For Missing Data
###
ggplot(data = missing[1:8,], aes(x=reorder(Columns, -Missing_PER),y = Missing_PER))+
  geom_bar(stat = "identity",fill = "grey")+xlab("Parameter")+
  ggtitle("Missing data percentage") + theme_bw()

ggplot(data = missing[9:15,], aes(x=reorder(Columns, -Missing_PER),y = Missing_PER))+
  geom_bar(stat = "identity",fill = "grey")+xlab("Parameter")+
  ggtitle("Missing data percentage") + theme_bw()

ggplot(data = missing[16:21,], aes(x=reorder(Columns, -Missing_PER),y = Missing_PER))+
  geom_bar(stat = "identity",fill = "grey")+xlab("Parameter")+
  ggtitle("Missing data percentage") + theme_bw()

####
# Imputation Of Missing Data In Dataset Using Median
###
Absenteeism$`Reason for absence`[is.na(Absenteeism$`Reason for absence`)] = median(Absenteeism$`Reason for absence`, na.rm = T)
Absenteeism$`Month of absence`[is.na(Absenteeism$`Month of absence`)] = median(Absenteeism$`Month of absence`, na.rm = T)
Absenteeism$`Disciplinary failure`[is.na(Absenteeism$`Disciplinary failure`)] = median(Absenteeism$`Disciplinary failure`, na.rm = T)
Absenteeism$Education[is.na(Absenteeism$Education)] = median(Absenteeism$Education, na.rm = T)
Absenteeism$`Social drinker`[is.na(Absenteeism$`Social drinker`)] = median(Absenteeism$`Social drinker`, na.rm = T)
Absenteeism$`Social smoker`[is.na(Absenteeism$`Social smoker`)] = median(Absenteeism$`Social smoker`, na.rm = T)
Absenteeism$`Transportation expense`[is.na(Absenteeism$`Transportation expense`)] = median.default(Absenteeism$`Transportation expense`, na.rm = T)
Absenteeism$`Distance from Residence to Work`[is.na(Absenteeism$`Distance from Residence to Work`)] = median(Absenteeism$`Distance from Residence to Work`, na.rm = T)
Absenteeism$`Service time`[is.na(Absenteeism$`Service time`)] = median(Absenteeism$`Service time`, na.rm = T)
Absenteeism$Age[is.na(Absenteeism$Age)] = median(Absenteeism$Age, na.rm = T)
Absenteeism$`Work load Average/day`[is.na(Absenteeism$`Work load Average/day`)] = median(Absenteeism$`Work load Average/day`, na.rm = T)
Absenteeism$`Hit target`[is.na(Absenteeism$`Hit target`)] = median(Absenteeism$`Hit target`, na.rm = T)
Absenteeism$Son[is.na(Absenteeism$Son)] = median(Absenteeism$Son, na.rm = T)
Absenteeism$Pet[is.na(Absenteeism$Pet)] = median(Absenteeism$Pet, na.rm = T)
Absenteeism$Weight[is.na(Absenteeism$Weight)] = median(Absenteeism$Weight, na.rm = T)
Absenteeism$Height[is.na(Absenteeism$Height)] = median(Absenteeism$Height, na.rm = T)
Absenteeism$`Body mass index`[is.na(Absenteeism$`Body mass index`)] = median(Absenteeism$`Body mass index`, na.rm = T)
Absenteeism$`Absenteeism time in hours`[is.na(Absenteeism$`Absenteeism time in hours`)] = median(Absenteeism$`Absenteeism time in hours`, na.rm = T)


####
# Converting variables to their types as factor for categorical variable
###
Absenteeism$`Reason for absence`=as.factor(Absenteeism$`Reason for absence`)
Absenteeism$`Month of absence` = as.factor(Absenteeism$`Month of absence`)
Absenteeism$`Day of the week` = as.factor(Absenteeism$`Day of the week`)
Absenteeism$Seasons = as.factor(Absenteeism$Seasons)
Absenteeism$`Disciplinary failure` = as.factor(Absenteeism$`Disciplinary failure`)
Absenteeism$Education = as.factor(Absenteeism$Education)
Absenteeism$`Social drinker` = as.factor(Absenteeism$`Social drinker`)
Absenteeism$`Social smoker`= as.factor(Absenteeism$`Social smoker`)

str(Absenteeism)

####
# Changing Column Name As Per Need and For Better Understanding
###
colnames(Absenteeism)[colnames(Absenteeism) == 'Absenteeism time in hours'] = 'Absenteeism.time.in.hours'
colnames(Absenteeism)[colnames(Absenteeism) == 'Reason for absence'] = 'Reason.for.absence'
colnames(Absenteeism)[colnames(Absenteeism) == 'Month of absence'] = 'Month.of.absence'
colnames(Absenteeism)[colnames(Absenteeism) == 'Day of the week'] = 'Day.of.week'
colnames(Absenteeism)[colnames(Absenteeism) == 'Transportation expense'] = 'Transportation.expense'
colnames(Absenteeism)[colnames(Absenteeism) == 'Distance from Residence to Work'] = 'Distance.from.residence.to.work'
colnames(Absenteeism)[colnames(Absenteeism) == 'Service time'] = 'Service.time'
colnames(Absenteeism)[colnames(Absenteeism) == 'Work load Average/day'] = 'Work.load.Average.per.day'
colnames(Absenteeism)[colnames(Absenteeism) == 'Hit target'] = 'Hit.target'
colnames(Absenteeism)[colnames(Absenteeism) == 'Disciplinary failure'] = 'Disciplinary.failure'
colnames(Absenteeism)[colnames(Absenteeism) == 'Social drinker'] = 'Social.drinker'
colnames(Absenteeism)[colnames(Absenteeism) == 'Social smoker'] = 'Social.smoker'
colnames(Absenteeism)[colnames(Absenteeism) == 'Body mass index'] = 'Body.mass.index'

####
# Outlier Analysis On our Numerical Variable
###
numeric_index = sapply(Absenteeism,is.numeric) #selecting only numeric
numerical = Absenteeism[,numeric_index]
Numerical = colnames(numerical)

for (i in 1:length(Numerical))
{
  assign(paste0("gn",i), ggplot(aes_string(y = (Numerical[i]), x = "Absenteeism.time.in.hours"), data = subset(Absenteeism))+ 
           stat_boxplot(geom = "errorbar", width = 0.5) +
           geom_boxplot(outlier.colour="red", fill = "blue" ,outlier.shape=18,
                        outlier.size=1, notch=FALSE) +
           theme(legend.position="bottom")+
           labs(y=Numerical[i],x="Absenteeism.time.in.hours")+
           ggtitle(paste("Box plot for",Numerical[i])))
}
####
# Plotting Boxplot For Outliers
###
gridExtra::grid.arrange(gn1,gn2,gn3,ncol=3)
gridExtra::grid.arrange(gn4,gn5,gn6,ncol=3)
gridExtra::grid.arrange(gn7,gn8,gn9,ncol=3)
gridExtra::grid.arrange(gn10,gn11,ncol=2)


####
# Outlier imputation using NA technique then Median for Missing Data
###
Out = Absenteeism$`Transportation.expense`[Absenteeism$`Transportation.expense` %in% boxplot.stats(Absenteeism$`Transportation.expense`)$out]
Absenteeism$`Transportation.expense`[(Absenteeism$`Transportation.expense` %in% Out)] = NA
Absenteeism$`Transportation.expense`[is.na(Absenteeism$`Transportation.expense`)] = median.default(Absenteeism$`Transportation.expense`, na.rm = T)

Out = Absenteeism$`Service.time`[Absenteeism$`Service.time` %in% boxplot.stats(Absenteeism$`Service.time`)$out]
Absenteeism$`Service.time`[(Absenteeism$`Service.time` %in% Out)] = NA
Absenteeism$`Service.time`[is.na(Absenteeism$`Service.time`)] = median(Absenteeism$`Service.time`, na.rm = T)

Out = Absenteeism$Age[Absenteeism$Age %in% boxplot.stats(Absenteeism$Age)$out]
Absenteeism$Age[(Absenteeism$Age %in% Out)] = NA
Absenteeism$Age[is.na(Absenteeism$Age)] = median(Absenteeism$Age, na.rm = T)

Out = Absenteeism$`Work.load.Average.per.day`[Absenteeism$`Work.load.Average.per.day` %in% boxplot.stats(Absenteeism$`Work.load.Average.per.day`)$out]
Absenteeism$`Work.load.Average.per.day`[(Absenteeism$`Work.load.Average.per.day` %in% Out)] = NA
Absenteeism$`Work.load.Average.per.day`[is.na(Absenteeism$`Work.load.Average.per.day`)] = median(Absenteeism$`Work.load.Average.per.day`, na.rm = T)

Out = Absenteeism$`Hit.target`[Absenteeism$`Hit.target` %in% boxplot.stats(Absenteeism$`Hit.target`)$out]
Absenteeism$`Hit.target`[(Absenteeism$`Hit.target` %in% Out)] = NA
Absenteeism$`Hit.target`[is.na(Absenteeism$`Hit.target`)] = median(Absenteeism$`Hit.target`, na.rm = T)

Out = Absenteeism$Pet[Absenteeism$Pet %in% boxplot.stats(Absenteeism$Pet)$out]
Absenteeism$Pet[(Absenteeism$Pet %in% Out)] = NA
Absenteeism$Pet[is.na(Absenteeism$Pet)] = median(Absenteeism$Pet, na.rm = T)

Out = Absenteeism$Height[Absenteeism$Height %in% boxplot.stats(Absenteeism$Height)$out]
Absenteeism$Height[(Absenteeism$Height %in% Out)] = NA
Absenteeism$Height[is.na(Absenteeism$Height)] = median(Absenteeism$Height, na.rm = T)


rm(Out)


####
# Correlation plot on Numerical variable to remove most common correlated Data
###
corrgram(Absenteeism, order = F,
         upper.panel=panel.pie, text.panel=panel.txt, main = "Correlation Plot")

str(Absenteeism)

####
# Making Copy of our Data 
####
Absent = Absenteeism
Absenteeism = Absent

####
# Random Forest For Important Variables in Dataset
####
library(randomForest)
Random.Model = randomForest(Absenteeism$Absenteeism.time.in.hours~.,data = Absenteeism , importance= TRUE,ntree = 500)
print(Random.Model)
attributes(Random.Model)
importance(Random.Model,type = 1)
varUsed(Random.Model) # to find which variables used in random forest
varImpPlot(Random.Model,sort = TRUE,n.var = 10,main="10 Imp. Variable")
varImp(Random.Model)

####
# Anova for categorical data to find important categorical variable for dataset
####
str(absent)
library(ANOVA.TFNs)
library(ANOVAreplication)
result = aov(formula=Absenteeism.time.in.hours~Reason.for.absence+Month.of.absence+Day.of.week+Seasons+Disciplinary.failure
             +Education+Social.smoker+Social.drinker, data = Absenteeism)
summary(result)

####
# Decision Tree to Find Important Variable Participation in Dataset
####
library(rpart.plot)
tree <- rpart(Absenteeism.time.in.hours ~ . , method='class', data = Absenteeism)
printcp(tree)
plot(tree, uniform=TRUE, main="Main Title")
text(tree, use.n=TRUE, all=TRUE)
prp(tree)

#qqnorm(Absenteeism$Transportation.expense)
#hist(Absenteeism$Transportation.expense)
#qqnorm(Absenteeism$Distance.from.residence.to.work)
#hist(Absenteeism$Distance.from.residence.to.work)
#qqnorm(Absenteeism$Service.time)
#hist(Absenteeism$Service.time)
#qqnorm(Absenteeism$Workload.average.perday)
#hist(Absenteeism$Workload.average.perday)
#qqnorm(Absenteeism$Hit.target)
#hist(Absenteeism$Hit.target)
#qqnorm(Absenteeism$Weight)
#hist(Absenteeism$Weight)
#qqnorm(Absenteeism$Absenteeism.time.in.hours)
#hist(Absenteeism$Absenteeism.time.in.hours)


####
# Removing Unnecessary and unimportant Variable from our Dataset
####
Absenteeism = subset(Absenteeism, select = -c(ID,Education,Day.of.week, Pet,Hit.target, Seasons,Social.smoker,
                                              Social.drinker,Weight))

####
# Making New Set of Numerical column
####
Numerical_Col = c("Transportation.expense","Age","Son","Height","Body.mass.index","Service.time","Work.load.Average.per.day",
                  "Distance.from.residence.to.work")

####
# Applying Normalization to bring bring all data to same scale
####
for(i in Numerical_Col){
  print(i)
  Absenteeism[,i] = (Absenteeism[,i] - min(Absenteeism[,i]))/
    (max(Absenteeism[,i] - min(Absenteeism[,i])))
}

rmExcept(c("Absent","Absenteeism"))


####
# Creating Sample Variable for Test and Train data for further Analysis
####
sample = sample(1:nrow(Absenteeism), 0.8 * nrow(Absenteeism))
train = Absenteeism[sample,]
test = Absenteeism[-sample,]


####
# Loading Import library require for tree analysis
####
library(caTools)
library(mltools)

####
# RMSE fuction to calculate Error Rate
####
RMSE <- function(Error)
{
  sqrt(mean(Error^2))
}


####
# Implementing Decision Tree Model on dataset
####
dtree = rpart(Absenteeism.time.in.hours~.,data = train, method = "anova")
dtree.plt = rpart.plot(dtree,type = 3,digits = 3,fallen.leaves = TRUE)
prediction_dtree = predict(dtree,test[,-12])
actual = test[,12]
predicted_dtree = data.frame(prediction_dtree)
summary(dtree)
Error = actual - predicted_dtree
RMSE(Error)
# Error = 13.42  #Accuracy = 85.84


####
# Implementing Random forest Model on dataset with N = 100,500,1000
####

# Random forest 100
Random.Model = randomForest(Absenteeism.time.in.hours~.,train,importance = TRUE,ntree = 100)
rand_pred = predict(Random.Model,test[-12])
actual = test[,12]
predicted_rand = data.frame(rand_pred)
Error = actual - predicted_rand
RMSE(Error)
# Error rate = 12.28  # Accuracy 87.72

# Random forest 500
Random.Model = randomForest(Absenteeism.time.in.hours~.,train,importance = TRUE,ntree = 500)
rand_pred = predict(Random.Model,test[-12])
actual = test[,12]
predicted_rand = data.frame(rand_pred)
Error = actual - predicted_rand
RMSE(Error)
# Error rate = 12.31  # Accuracy 87.69

# Random forest 1000
Random.Model = randomForest(Absenteeism.time.in.hours~.,train,importance = TRUE,ntree = 1000)
rand_pred = predict(Random.Model,test[-12])
actual = test[,12]
predicted_rand = data.frame(rand_pred)
Error = actual - predicted_rand
RMSE(Error)
summary(Random.Model)
# Error rate = 12.37  # Accuracy 87.63

####
# Converting Categorical to numerical for linear Regression Analysis
###
Absenteeism$Reason.for.absence=as.numeric(Absenteeism$Reason.for.absence)
Absenteeism$`Month.of.absence` = as.numeric(Absenteeism$`Month.of.absence`)
Absenteeism$`Disciplinary.failure`= as.numeric(Absenteeism$`Disciplinary.failure`)
str(Absenteeism)


####
#Implementing Linear Regression Model on dataset with new Sample
####
droplevels(Absenteeism$Reason.for.absence)
sample_lm = sample(1:nrow(Absenteeism),0.8*nrow(Absenteeism))
train_lm = Absenteeism[sample_lm,]
test_lm = Absenteeism[-sample_lm,]
linear_model = lm(Absenteeism.time.in.hours~.,data = train_lm)
summary(linear_model)
#vif(linear_model)
predictions_lm = predict(linear_model,test_lm[,1:11])
Predicted_LM = data.frame(predictions_lm)
Actual = test_lm[,12]
Error = Actual - Predicted_LM
RMSE(Error)
#Error rate = 9.81  Accuracy 90.19

###     ###
# PART 2  #
###     ###

####
# Creating New Dataset require to calculate Loss per month company can estimate in coming Year with same data
###
NEW_LOSS_DF = subset(Absent, select = c(Month.of.absence, Service.time, Absenteeism.time.in.hours, 
                                        Work.load.Average.per.day))

#Work loss = (service time - Absenteeism time) / work load per day
NEW_LOSS_DF["Loss"]=with(NEW_LOSS_DF,((NEW_LOSS_DF[,4]*NEW_LOSS_DF[,3])/NEW_LOSS_DF[,2]))

for(i in 0:12)
{
  DataLOSS=NEW_LOSS_DF[which(NEW_LOSS_DF["Month.of.absence"]==i),]
  print(data.frame(sum(DataLOSS$Loss)))
  
}
