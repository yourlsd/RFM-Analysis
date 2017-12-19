
#----------------------------------#
# Testing - forward, monthly       #
#----------------------------------#
library(readxl)#for reading xls file
library(readr)#for reading csv file
library(ROSE)#for dealing with imbalanced class
library(rpart)#for decision tree
library(randomForest)#for random forest
library(gplots)#for visualization

#Update login report: from 2017-02-01 to 2017-12-01
login_2017_12_08 <- read_excel("C:/Users/ttlus/Desktop/export-login-2017-12-08 (1).xlsx")

# Total Expiration information
Expiration<- read_excel("C:/Users/ttlus/Desktop/Copy of Expiration Dates for Subscription Clients -TOTAL.xlsx", 
                        col_types = c("text", "text", "text", "date", "date", "date"))
#'Expiration_Q1' means the clients expire in Q1 of 2018
Expitation_Q1<-read_excel("C:/Users/ttlus/Desktop/Expiration Dates for Subscription - Q1 2018.xlsx", 
                          col_types = c("text", "text", "text", "date", "date", "date"))

#Note: using 2017-11-08 information here. Assume the tier doesn't change in December
tier <- read_excel("C:/Users/ttlus/Desktop/Gateway & Salesforce Metrics/company by tier.xlsx")

#Update search term report: from 2017-02-01 to 2017-12-01
searchterm<- read_excel("C:/Users/ttlus/Desktop/searchterm-detail-2017-12-08.xlsx")

#Update click report: from 2017-02-01 to 2017-12-01
clicked<- read_excel("C:/Users/ttlus/Desktop/TrackClickedContent-2017-12-08.xlsx")

#############################
# Descriptive RFME
#############################


#-----------------------------#
# R -- Company Recency summary#
#-----------------------------#

#Define the date of checking here. Usually the 1st day of each month
Decfirst<-"2017-12-01"

#Create a new column: days since last login 
last_login_date<-substr(login_2017_12_08$`Last Login`, 1, 10)

login_2017_12_08$days_since_last_login<-
  round(as.numeric(difftime(Decfirst,
                            substr(login_2017_12_08$`Last Login`, 1, 10)),
                   units = "days"))

#Check the quantile for recency, the result here don't change a lot compared to the previous 
# one with date from 2017-01-01 to 2017-10-01
#########################################################
# BUT IF THE QUANTILES CHANGES A LOT, WE NEED TO 
# CHANGE THE CUTOFF IN THE FUTURE ANALYSIS!!!
# So does the cutoff of frequency, clicks, searches
#########################################################

quantile(login_2017_12_08$days_since_last_login)
#    0%  25%  50%  75% 100% 
#    0   25   84  178  303 

#Here I still use 60 days as the cutoff of defining 'active user'
cutoff<-60

#Company recency summary
active_rec<-login_2017_12_08[login_2017_12_08$days_since_last_login <= cutoff,]

avg_active  <- aggregate(active_rec$days_since_last_login,
                         by=list(active_rec$Company), mean)

n_of_active_rec<-data.frame(table(active_rec$Company))

n_of_user<-data.frame(table(login_2017_12_08$Company))

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

#quantile(login_2017_12_08$Count)
#   0%  25%  50%  75% 100% 
#   1    1    2    5  649 



total_freq<- aggregate(login_2017_12_08$Count,
                       by=list(login_2017_12_08$Company), sum)
freqcutoff<-5

active_freq<-login_2017_12_08[login_2017_12_08$Count > freqcutoff,]

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

company_tier<-merge(login_2017_12_08,tier,all.x = T,
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
  round(as.numeric(difftime(Decfirst,Expiration$`US Expiration Date`),
                   units = "days"))

Expiration$CUEX<-
  round(as.numeric(difftime(Decfirst,Expiration$`Current TRU Agreement End Date`),
                   units = "days"))

Expiration$GLEX<-
  round(as.numeric(difftime(Decfirst,Expiration$`Global Expiration Date`),
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
#eg. if there is no active users in a company, impute the NA with median of INACTIVE USER
# These numbers will change in the future! make sure to check the new quantiles!!!
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
#May check manually here to see if there is any unreasonable points

#Issue : we have many duplicated records here
newdata <- unique(data[,-2])

#Exclude records with missing response
newdata2<-newdata[is.na(newdata$Subscription)==F,]
#174 viable obs

#Expire soon -forward test
expiresoon<-data[data$`Sales Force Id` %in%  Expitation_Q1$`Account ID`,]

newdata2<-newdata2[,c(5,8,24,15,16,19)]
#---------------------#
# Decision Tree       #
#---------------------#

################################Random test################################
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


################################Test for ones who expire soon#######################

#Expire soon -forward test
expiresoon<-data[data$`Sales Force Id` %in%  Expitation_Q1$`Account ID`,-2]
testdata<-expiresoon[,c(1,5,8,24,15,16,19)]

traindata<-data[-(data$`Sales Force Id` %in%  Expitation_Q1$`Account ID`),]
traindata <- unique(traindata[,-2])
traindata<-traindata[is.na(traindata$Subscription)==F,]
traindata<-traindata[,c(1,5,8,24,15,16,19)]

colnames(traindata)<-c('company','recency','frequency','monetary','search','click','subsc')
colnames(testdata)<-c('company','recency','frequency','monetary','search','click','subsc')

#Decision Tree - imbalanced class
set.seed(1)
treeimb <- rpart(subsc ~ recency+frequency+monetary+search+click, data = traindata)
pred.treeimb <- predict(treeimb, newdata = testdata)

table(testdata$subsc, round(pred.treeimb))
table(round(pred.treeimb))
table(testdata$subsc)
predicted<-cbind(testdata,pred.treeimb)
colnames(predicted)[8]<-'Probability of Subscribing'

#Define the output address here
# Output the results to 
write.csv(predicted,'C:\\Users\\ttlus\\Desktop\\Predicted-2018-Q1.csv',row.names = F)
write.csv(data,'C:\\Users\\ttlus\\Desktop\\gateway summary from 2017-02-01 to 2017-12-01 .csv',row.names = F)
#2 of them are predicted as not subscribing
