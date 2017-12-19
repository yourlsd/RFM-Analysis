##############################
# Datasets and packages needed
##############################

library(readxl)#for reading xls file
library(readr)#for reading csv file
library(ROSE)#for dealing with imbalanced class
library(rpart)#for decision tree
library(randomForest)#for random forest
library(gplots)#for visualization


login2017 <- read_excel("C:/Users/ttlus/Desktop/Gateway & Salesforce Metrics/login2017.xlsx")
tier <- read_excel("C:/Users/ttlus/Desktop/Gateway & Salesforce Metrics/company by tier.xlsx")
searchterm<- read_excel("C:/Users/ttlus/Desktop/Gateway & Salesforce Metrics/searchterm-detail-2017-11-09.xlsx")
clicked<- read_excel("C:/Users/ttlus/Desktop/Gateway & Salesforce Metrics/TrackClickedContent-2017-11-09.xlsx")
Expiration<- read_excel("C:/Users/ttlus/Desktop/Expiration Dates for Subscription Clients.xlsx", 
 col_types = c("text", "text", "text", "date", "date", "date"))



#############################
# Descriptive RFME
#############################


#-----------------------------#
# R -- Company Recency summary#
#-----------------------------#
octfirst<-"2017-10-01"
last_login_date<-substr(login2017$`Last Login`, 1, 10)

login2017$days_since_last_login<-
  round(as.numeric(difftime(octfirst,
                            substr(login2017$`Last Login`, 1, 10)),
                   units = "days"))

#hist(login2017$days_since_last_login, xlab = "Days since Last Login",
#main = "Distribution of All Individual's Recency",breaks = 25)
#quantile(login2017$days_since_last_login)
#   0%  25%  50%  75% 100% 
#   0   25   77  159  273 

#H <- 1.5 * IQR(login2017$days_since_last_login)
#159+H
#IQR upper limit=360

#mean(login2017$days_since_last_login)
# 98.87231 days since last login

#avg_recency  <- aggregate(login2017$days_since_last_login,
#by=list(login2017$Company), mean)
#colnames(avg_recency)<-c("Company","Average Days Since Last Login")

#hist(avg_recency$`Average Days Since Last Login`, xlab = "Days since Last Login",
#main = "Distribution of Average Recency Per Company",breaks = 25)

#med_recency  <- aggregate(login2017$days_since_last_login,
#by=list(login2017$Company), median)

#hist(med_recency$x, xlab = "Days since Last Login",
#main = "Distribution of Median Recency Per Company",breaks = 25)

cutoff<-60


#Company recency summary
active_rec<-login2017[login2017$days_since_last_login <= cutoff,]

avg_active  <- aggregate(active_rec$days_since_last_login,
                         by=list(active_rec$Company), mean)

n_of_active_rec<-data.frame(table(active_rec$Company))

n_of_user<-data.frame(table(login2017$Company))

comb<-merge(n_of_active_rec,n_of_user,by.x = 'Var1',by.y = 'Var1',all = T)

comb[is.na(comb)] <- 0

comb$proportion<-comb$Freq.x/comb$Freq.y

company_recency<-merge(comb,avg_active,by.x = 'Var1',
                       by.y = 'Group.1', all= T)

colnames(company_recency)<-c("Company", "Number of active users(recency)","Total number of users",
                             "Proportion of active users(recency) ",
                             "Avrage Days Since Last Login for active users")


#------------------------------#
# F-- Company Frequency summary#
#------------------------------#
#hist(login2017$Count, xlab = "Login Count in 10 months",
# main = "Distribution of All Individual's Log-in Frequency",breaks = 25)

#quantile(login2017$Count)
#   0%  25%  50%  75% 100% 
#   1    1    2    4  593 
# Extremely skewed!

#mean(login2017$Count)
#4.581813

#H <- 1.5 * IQR(login2017$Count)

#login2017.1<-login2017[login2017$Count<9,]

#hist(login2017.1$Count, xlab = "Login Count in 10 months",
#main = "Distribution of Individual's Log-in Frequency for frequency<9",breaks = 20)

#405/2673
#IQR upper limit=360

total_freq<- aggregate(login2017$Count,
                       by=list(login2017$Company), sum)
freqcutoff<-5

active_freq<-login2017[login2017$Count > freqcutoff,]

n_of_active_freq<-data.frame(table(active_freq$Company))

avg_active_freq  <- aggregate(active_freq$Count,
                              by=list(active_freq$Company), mean)

active_freq_com <- aggregate(active_freq$Count,
                             by=list(active_freq$Company), sum)

comb<-merge(n_of_active_freq,n_of_user,by.x = 'Var1',by.y = 'Var1',all = T)

comb[is.na(comb)] <- 0

comb$proportion<-comb$Freq.x/comb$Freq.y

company_freqency<-merge(comb,avg_active_freq,by.x = 'Var1',
                        by.y = 'Group.1', all= T)

colnames(company_freqency)<-c("Company", "Number of active users(frequency)","Total number of users",
                              "Proportion of active users(frequency) ",
                              "Avrage Login Frequency for active users")
quantile(company_freqency$`Avrage Login Frequency for active users`,na.rm = T)
#------------------==------------#
#  M -- Company Tier summary     #
#--------------------------------#
tier <- tier[,8:9]

company_tier<-merge(login2017,tier,all.x = T,
                 by.x = "Sales Force Id", by.y = "Account ID")

company_tier<-as.data.frame(unique(company_tier[,c(1,2,9)]))


#--------------------------------#
# E -- Company Engagement summary#
#--------------------------------#

########Search Data Preprocessing##########
searchterm$TimeStamp<-substr(searchterm$TimeStamp, 1, 10)

#Remove replicated searches within one way
searchterm<-unique(searchterm)
search_count  <- aggregate(rep(1, length(searchterm$`Search Term`)),
                           by=list(searchterm$Company), sum)

searchaccount<-unique(searchterm[,1:2])
search_count$account<-table(searchaccount$Company)
search_count$'Search Count/Account'<- search_count$x/search_count$account
colnames(search_count) <- c("Company","Total Search Count","Number of Users who searched","Avg search count")

user_search_count<-as.data.frame(table(searchterm$UserName))
#quantile(user_search_count$Freq)
#0%  25%  50%  75% 100% 
#1    2    3    7  214 
#The pattern of search count is consistent with login count
#There are lots of searching counts created by inactive users and we would like to see
#the search counts created by active users as we did before.
search_active_user<-user_search_count[user_search_count$Freq>3,]
active_freq_search<-searchterm[searchterm$UserName %in% search_active_user$Var1,]
#12162 out of 14047

n_of_active_freq_search<-data.frame(table(active_freq_search$Company))
user_count_of_active_freq_search<-unique(active_freq_search[,1:2])
user_count_of_active_freq_search<-data.frame(table(user_count_of_active_freq_search$Company))


company_search<-merge(n_of_active_freq_search,user_count_of_active_freq_search,by.x = 'Var1',by.y = 'Var1',all = T)

company_search[is.na(company_search)] <- 0

company_search$avg<-company_search$Freq.x/company_search$Freq.y

colnames(company_search)<-c("Company", "Total number of searches from active users","Number of active users(search)","Avg search count for active users")

company_search<-merge(search_count,company_search,by.x = 'Company',by.y = 'Company',all = T)

company_search[is.na(company_search)] <- 0

#I also tried to conduct an RFM to search&click but there is no user info in click data
#################click Data preprocessing##############################
click_count  <- aggregate(clicked$Total,
                          by=list(clicked$Company), sum)
click_count_mean  <- aggregate(clicked$Total,
                               by=list(clicked$Company), mean)
click_count<-cbind(click_count[,1:2],click_count_mean[,2])

click_count$'#Users'<- click_count[,2]/click_count[,3]

company_click<-click_count[,c(1,2,4,3)]

colnames(company_click) <- c("Company","Total Click Count","#Users","Avg Clicked count")

#----------------------------------------------#
# Combine R,F,M,E Data and Response            #
#----------------------------------------------#

comb<-merge(company_recency,company_freqency, by.x = 'Company',by.y = 'Company',all = T)

comb<-merge(comb,company_tier, by.x = 'Company',by.y = 'Company',all = T)

comb<-merge(comb,company_search, by.x = 'Company',by.y = 'Company',all = T)

comb<-merge(comb,company_click, by.x = 'Company',by.y = 'Company',all = T)

#Determinhe if the client is subscribing
Expiration$USEX<-
  round(as.numeric(difftime(octfirst,Expiration$`US Expiration Date`),
                   units = "days"))

Expiration$CUEX<-
  round(as.numeric(difftime(octfirst,Expiration$`Current TRU Agreement End Date`),
                   units = "days"))

Expiration$GLEX<-
  round(as.numeric(difftime(octfirst,Expiration$`Global Expiration Date`),
                   units = "days"))

for(i in 1:nrow(Expiration)){
  Expiration$subsc[i]<-ifelse(min(Expiration[i,7:9],na.rm = T) <0 , 1, 0)}

comb<-merge(comb,Expiration, by.x = 'Sales Force Id', by.y = 'Account ID',all.x = T)

c<-c(1,2,4,3,5,6,7,9,10,11:20,29)
combtable<-comb[,c]
colnames(combtable)[3]<-"Total Number of Users who login"
colnames(combtable)[10]<-"Client/Prospect Tier"
colnames(combtable)[18]<-"Total Number of Users who clicked"
colnames(combtable)[20]<-"Subscription"
#colnames(combtable)

#Impute NA with median
combtable$`Avrage Days Since Last Login for active users`[is.na(combtable$`Avrage Days Since Last Login for active users`)] <- 150
combtable$`Avrage Login Frequency for active users`[is.na(combtable$`Avrage Login Frequency for active users`)] <- 1
combtable$`Client/Prospect Tier`[is.na(combtable$`Client/Prospect Tier`)] <- 0
combtable$`Total number of searches from active users`[is.na(combtable$`Total number of searches from active users`)] <- 2
combtable$`Number of active users(search)` [is.na(combtable$`Number of active users(search)`)] <- 0
combtable$`Avg search count for active users` [is.na(combtable$`Avg search count for active users`)] <- 2
combtable$`Total Search Count` [is.na(combtable$`Total Search Count`)] <- 0
combtable$`Total Click Count` [is.na(combtable$`Total Click Count`)] <- 0
combtable$`Number of Users who searched` [is.na(combtable$`Number of Users who searched` )] <- 0
combtable$`Total Number of Users who clicked` [is.na(combtable$`Total Number of Users who clicked`)] <- 0
combtable$`Avg search count` [is.na(combtable$`Avg search count`)] <- 0
combtable$`Avg Clicked count` [is.na(combtable$`Avg Clicked count`)] <- 0

#-----------------------------------#
# Create R,F,M,E Segments           #
#-----------------------------------#
RFME<-NA
RFME$Company<- combtable$Company
RFME<-as.data.frame(RFME)
quantile(company_recency$`Avrage Days Since Last Login for active users`, c(0.0, 0.25, 0.50, 0.75, 1.0),na.rm = T)
quantile(company_freqency$`Avrage Login Frequency for active users`, c(0.0, 0.25, 0.50, 0.75, 1.0),na.rm = T)
quantile(company_click$`Total Click Count`, c(0.0, 0.25, 0.50, 0.75, 1.0),na.rm = T)

RFME$R_segment <- findInterval(combtable$`Avrage Days Since Last Login for active users`, 
                               quantile(company_recency$`Avrage Days Since Last Login for active users`, c(0.0, 0.25, 0.50, 0.75, 1.0),na.rm = T))
#make some transformation of recency and get higher scores --> better recency
RFME$Rsegment <- ifelse(RFME$R_segment == 5, 1,
                        ifelse(RFME$R_segment == 4, 2,
                               ifelse(RFME$R_segment == 2, 4, 
                                      ifelse(RFME$R_segment == 1, 5,3))))
RFME$Recency <- ordered(ifelse(RFME$Rsegment <= 1, "60+",
                               ifelse(RFME$Rsegment <= 2, "29-59", 
                                      ifelse(RFME$Rsegment <= 3, "22-28",
                                             ifelse(RFME$Rsegment <= 4, "16-22",
                                             "0-15")))),
                        levels = c("0-15","16-22","22-28","29-59","60+"))

#Need to do this manually since frequency data is so polarized
RFME$Fsegment <- findInterval(combtable$`Avrage Login Frequency for active users`, 
                             c(1, 6, 9, 12, 15))
RFME$Frequency <- ordered(ifelse(RFME$Fsegment <= 1, "5-",
                                ifelse(RFME$Fsegment <= 2, "6-9",
                                       ifelse(RFME$Fsegment <= 3, "9-12", 
                                              ifelse(RFME$Fsegment <= 4, "12-15",
                                                     "15+")))),
                         levels = c("15+","12-15", "9-12", "6-9", "5-"))

RFME$Monetary <- combtable$`Client/Prospect Tier`
RFME$Msegment <- ifelse(RFME$Monetary == "Global Focus", 5,
                        ifelse(RFME$Monetary == "Platinum", 4, 
                               ifelse(RFME$Monetary == "Gold", 3, 
                                      ifelse(RFME$Monetary == "Silver", 2, 
                                             ifelse(RFME$Monetary == "Bronze", 1, 0)))))

RFME$Esegment <- findInterval(combtable$`Total Click Count`, 
                             c(0, 1, 18, 50, 100))

RFME$Engagement <- ordered(ifelse(RFME$Esegment<=1, "0",
                                 ifelse(RFME$Esegment <= 2, "1-18",
                                        ifelse(RFME$Esegment <= 3, "18-50", 
                                               ifelse(RFME$Esegment <= 4, "50-100",
                                                      "100+")))),
                          levels = c("100+","50-100","18-50", "1-18","0"))

RFME<-RFME[,c(2,4:7,9,8,10,11)]

#Merge to combtable and create a integrated dataset
data<-merge(combtable,RFME,by.x = "Company",by.y = "Company")

#Create an excel to view
write.csv(data,'C:\\Users\\ttlus\\Desktop\\gateway summary.csv',row.names = F)
?write.csv
# Recency by Frequence - Counts
RxF <- as.data.frame(table(RFME$Recency, RFME$Frequency,
                           dnn = c("Recency", "Frequency")),
                     responseName = "Number_Customers")
with(RxF, balloonplot(Recency, Frequency, Number_Customers,
                      zlab = "# Total Companies"))
RxE <- as.data.frame(table(RFME$Recency, RFME$Engagement,
                           dnn = c("Recency", "Engagement")),
                     responseName = "Number_Customers")
with(RxE, balloonplot(Recency, Engagement, Number_Customers,
                      zlab = "# Total Companies"))


#############################
# Prediction Part
#############################

#---------------------#
# Outlier Detection   #
#---------------------#
#unsubscribing clients
unsub<-na.omit(data[data$Subscription==0,])
#Can we check manually here?

#Exclude SABmiller
newdata1<-data[-(data$Company=="SABMiller Plc (EMEA)"),]

#Exclude records with missing response
newdata2<-newdata1[is.na(newdata1$Subscription)==F,]
#170 viable obs

#Need to be checked

newdata2<-newdata2[,c(6,9,25,16,17,20)]
#---------------------#
# Decision Tree       #
#---------------------#
#shuffle a data and divide it into training and test sets with a 60/40 split
n <- nrow(newdata2)
set.seed(1)
shuffled_df <- newdata2[sample(n), ]
train_indices <- 1:round(0.6 * n)
train <- shuffled_df[train_indices, ]
test_indices <- (round(0.6 * n) + 1):n
test <- shuffled_df[test_indices, ] 

colnames(train)<-c('recency','frequency','monetary','search','click','subsc')
colnames(test)<-c('recency','frequency','monetary','search','click','subsc')

#Decision Tree - imbalanced class
set.seed(1)
treeimb <- rpart(subsc ~ recency+frequency+monetary+search+click, data = train)
pred.treeimb <- predict(treeimb, newdata = test)

accuracy.meas(test$subsc, pred.treeimb)
roc.curve(test$subsc, pred.treeimb, plotit = F)

table(test$subsc, round(pred.treeimb))
table(round(pred.treeimb))
table(test$subsc)
temp<-cbind(test,pred.treeimb)
#Over sampling -- have even worse performance, discard it 
#nrow(train[train$subsc==1,])
#14 unsubscribing in total
#105 subscribing
#data_balanced_over <- ovun.sample(subsc ~ recency+frequency+monetary+search+click, data = train,
#                                  method = "over",N = 210)$data
#table(data_balanced_over$subsc)

#Decision Tree
#tree.over <- rpart(subsc ~ recency+frequency+monetary+search+click, data = data_balanced_over)
#pred.tree.over <- predict(tree.over, newdata = test)

#accuracy.meas(test$subsc, pred.tree.over)
#roc.curve(test$subsc, pred.tree.over)


#---------------------#
# Random Forest       #
#---------------------#
#Random Forest
train$subsc<-as.factor(train$subsc)
test$subsc<-as.factor(test$subsc)
rfmodel<- randomForest(subsc~.,data = train,importance=TRUE)

#prediction accuracy
pred.rf <- predict(rfmodel, newdata = test)
pred.rf.prob <- predict(rfmodel, newdata = test, type = "prob")

table(pred.rf, test$subsc)
accuracy.meas(test$subsc, pred.rf)
roc.curve(test$subsc, pred.rf, plotit = F)


#Decision tree has better performance but not stable.
#Here we choose 40% data for testing and 60% for training
#If we hope to test on real data and anticipate, need updated company expiration report

#Monthly prediction
#Frequency - frequency this month? frequency all the time?
#Recency - the same method, need updated recency
#Monetary - the same method, need updated Monetary
#Engagement - Engagement this month?
# Maybe it's just OK to use this model and for monthly analysis, change the start to end date of the input data
# (May still use 10 month interval)
# Change the cut off based on quantiles of new data





