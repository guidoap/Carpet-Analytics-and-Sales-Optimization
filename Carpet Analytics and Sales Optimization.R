library(tidyverse)
library(ggrepel)
library(ggthemes)
library(maps)
library(hrbrthemes)
library(viridis)
library(caret)

## Data load ##
#First Loading the data from the excel file into R and setting up the path
pacman::p_load("pacman","tidyverse","rpart","rpart.plot","readxl","ggplot2")

#Dataset - Raw Data-Order and Sample
OrderSample <- read_xlsx("Champo.xlsx",sheet = 2)
#Dataset - Data Order ONLY 
DataOrder <- read_xlsx("Champo.xlsx", sheet = 3)
#Data on Sample ONLY
DataOnSample<- read_xlsx("Champo.xlsx", sheet = 4)
#Data for Recommendation
DataForRecommendation <- read_xlsx("Champo.xlsx", sheet = 5)
#Data for Clustering
DataForClustering <- read_xlsx("Champo.xlsx",sheet = 6)
#Data-Association Rules A-11
AssociationRules <- read_xlsx("Champo.xlsx", sheet = 7)
#Once our data is loaded we can now start cleaning our data and 
#the way it is done is by replacing the NA values with the most occuring value in the column

####The Data Cleaning ####

####Data Cleaning for Order Sample table####
#Creating backup for Order Sample Dataset
OrderSample2 <- OrderSample

# Convert all columns to factor
OrderSample <- as.data.frame(unclass(OrderSample),                    
                              stringsAsFactors = TRUE)
#Data type check
str(OrderSample)

## How many rows and columns does the dataset have
dim(OrderSample)
names(OrderSample)

#summary before cleaning
summary(OrderSample)

### Missing values

##The following code shows total No. of missing values
sum(is.na(OrderSample))
## There are only 9 missing values. So, we have decided to remove the 9 rows that contain NAs

## Deleting remaining missing value rows (Total deleted rows= 9)
OrderSample <- na.omit(OrderSample)
sum(is.na(OrderSample))

#Converting values of categorical with posix format into date format
OrderSample$Custorderdate <- as.Date(OrderSample$Custorderdate)

## Outlier Treatment 
##One of the best ways of finding outliers is to draw the boxplot of each column.
par(mfrow = c(2,2))
boxplot(OrderSample$QtyRequired,
        main = "Boxplot of QtyRequired", 
        xlab = "QtyRequired")
boxplot(OrderSample$TotalArea,
        main = "Boxplot of TotalArea",
        xlab = "TotalArea")
boxplot(OrderSample$Amount,
        main = "Boxplot of Amount",
        xlab = "Amount")
boxplot(OrderSample$AreaFt, 
        main = "Boxplot of total AreaFt", 
        xlab = "AreaFt")
dev.off()

## Information obtain from the boxplot is not adequate to identify outliers
## One on the reason could be, there are 4 carpet categories and above parameters can vary for each category

#There are outliers in CustomerOrderNo
#Since there are a lot of values in CustomerOrderNo we do the following to extract most occuring value in our data
CON1<-OrderSample$CustomerOrderNo
sort(table(CON1), decreasing = T)[1:3] #By doing this we get the top 3 frequently occuring values in the column
#the value "12985" has occured 114 times so we replace the NA values with it 
a3<-OrderSample$CustomerOrderNo
e3<-replace_na(a3,"12985")
OrderSample$CustomerOrderNo<-as.numeric(e3)


####Data Cleaning for DataOrder table####

##The following code shows total No. of missing values
sum(is.na(DataOrder))
## There is no missing value

# Convert all columns to factor
DataOrder <- as.data.frame(unclass(DataOrder),                    
                             stringsAsFactors = TRUE)

#summary before cleaning
summary(DataOrder)
#There are no outliers in DataOrder table

####Data Cleaning for DataOnSample table ####

# Convert all columns to factor
DataOnSample <- as.data.frame(unclass(DataOnSample),                    
                              stringsAsFactors = TRUE)

##The following code shows total No. of missing values
sum(is.na(DataOnSample))
## There are 273 missing value

### Missing values by column
colSums(is.na(DataOnSample))

#Replacing missing value with "0" Since they are not the country mentioned in the
#CountryName variable:
#USA
DataOnSample$USA[is.na(DataOnSample$USA)] <- 0
DataOnSample$USA <- factor(DataOnSample$USA)

#UK
DataOnSample$UK[is.na(DataOnSample$UK)] <- 0
DataOnSample$UK <- factor(DataOnSample$UK)

#Italy
DataOnSample$Italy[is.na(DataOnSample$Italy)] <- 0
DataOnSample$Italy <- factor(DataOnSample$Italy)

#Belgium
DataOnSample$Belgium[is.na(DataOnSample$Belgium)] <- 0
DataOnSample$Belgium <- factor(DataOnSample$Belgium)

#Romania
DataOnSample$Romania[is.na(DataOnSample$Romania)] <- 0
DataOnSample$Romania <- factor(DataOnSample$Romania)

#Australia
DataOnSample$Australia[is.na(DataOnSample$Australia)] <- 0
DataOnSample$Australia <- factor(DataOnSample$Australia)

#India
DataOnSample$India[is.na(DataOnSample$India)] <- 0
DataOnSample$India <- factor(DataOnSample$India)

#Adding the other 7 countries:
DataOnSample$Poland <- as.factor(ifelse(DataOnSample$CountryName == "POLAND", 1, 0))
DataOnSample$Canada <- as.factor(ifelse(DataOnSample$CountryName == "CANADA", 1, 0))
DataOnSample$UAE <- as.factor(ifelse(DataOnSample$CountryName == "UAE", 1, 0))
DataOnSample$Sounth_Africa <- as.factor(ifelse(DataOnSample$CountryName == "SOUTH AFRICA", 1, 0))
DataOnSample$Brazil <- as.factor(ifelse(DataOnSample$CountryName == "BRAZIL", 1, 0))
DataOnSample$China <- as.factor(ifelse(DataOnSample$CountryName == "CHINA", 1, 0))
DataOnSample$Israel <- as.factor(ifelse(DataOnSample$CountryName == "ISRAEL", 1, 0))

#Turning the other binary variables into factors:
DataOnSample$Hand.Tufted <- factor(DataOnSample$Hand.Tufted)
DataOnSample$Durry <- factor(DataOnSample$Durry)
DataOnSample$Double.Back <- factor(DataOnSample$Double.Back)
DataOnSample$Hand.Woven <- factor(DataOnSample$Hand.Woven)
DataOnSample$Knotted <- factor(DataOnSample$Knotted)
DataOnSample$Jacquard <- factor(DataOnSample$Jacquard)
DataOnSample$Handloom <- factor(DataOnSample$Handloom)
DataOnSample$Other <- factor(DataOnSample$Other)

DataOnSample$REC <- factor(DataOnSample$REC)
DataOnSample$Round <- factor(DataOnSample$Round)
DataOnSample$Square <- factor(DataOnSample$Square)

DataOnSample$Order.Conversion <- factor(DataOnSample$Order.Conversion)

### EDA ######################################################### 

## 1. Order Category (order and sample during 2017-20)

OrderSample$OrderYear <-as.numeric(format(OrderSample$Custorderdate,'%Y'))

OrdSam<- as.data.frame(OrderSample %>% 
                            group_by(OrderYear,OrderCategory) %>%
                            summarise(total=n()) %>%
                            arrange(desc(OrderYear)))

ggplot(data=OrdSam, aes(x=OrderYear, y=total, group=OrderCategory, colour=OrderCategory)) +
  geom_line(aes(linetype=OrderCategory))+
  geom_point(aes(shape=OrderCategory))+
  geom_text(aes(y=total, label=round(total,2)),size=5)+
  scale_color_brewer(palette="Paired")+
  xlab("Year") +
  ylab("Number of Orders") +
  ggtitle("Total Orders and Samples by Year") +
  theme_minimal()

## Insights- We will ignore year 2020 as data is incomplete, information only for January and February are present. 
## Number of orders increased from 2017 to 2018 and again decreased in 2019. 
## There were no samples in 2017 and increased almost 3 times from 2018 to 2019. 

## 2. Revenue Generated

RevenueDistribution<- as.data.frame(OrderSample %>% 
                         group_by(ITEM_NAME) %>%
                         summarise(revenue= sum(Amount)) %>%
                         arrange(desc(revenue)))


ggplot(RevenueDistribution, aes(x="", y=revenue, fill=fct_inorder(ITEM_NAME))) +
  geom_col(width = 1, color = 1) +
  coord_polar(theta = "y") +
  scale_fill_brewer(palette = "Pastel1") +
  geom_label_repel(data = RevenueDistribution,
                   aes(y = round(revenue,0), label = paste(format(round(revenue / 1e3, 1), trim = TRUE), "K")),
                   size = 4.5, nudge_x = 1, show.legend = FALSE) +
  guides(fill = guide_legend(title = "Group")) +
  ggtitle("Revenue Distribution by type of Item") +
  theme_void()

## Insight: Highest contributor of revenue is the “Hand Tufted” carpet type followed by “Durry”. 
## “Gun Tufted” has the lowest sales among all types. 

## 3.  Carpet type and unit sold 

UnitSold<- as.data.frame(OrderSample %>% 
                                      group_by(ITEM_NAME) %>%
                                      summarise(Unit= sum(QtyRequired)) %>%
                                      arrange(desc(Unit)))

ggplot(UnitSold, aes(x=ITEM_NAME, y=Unit, fill=ITEM_NAME)) +
  geom_text(aes(label=paste(format(round(Unit / 1e3, 1), trim = TRUE), "K")), vjust=-0.3, size=2.5)+
  geom_bar(stat="identity")+theme_minimal()+
  theme(legend.position = "right")+
  theme(axis.text.x=element_blank(),
        axis.ticks.x=element_blank(),
        axis.text.y=element_blank(),
        axis.ticks.y=element_blank() 
  )+
  xlab("Carpet Type") +
  ylab("Number of Unit Sold") +
  ggtitle("Total number of unit sold by Carpet Type")

## Insight: Highest number of carpets sold are of “Durry” type 
## but it was the second highest contributor of revenue 
## and considerably lower than “Hand Tufted” in terms of revenue 
## therefore we can say it is much cheaper in price. 
## Since “Hand Tufted” carpet is the highest generator of revenue 
## but number of units sold is much less, we can say that it is towards the expensive side 
## and hence a premium quality carpet.

## 4. Countries vs revenue 

ConRev<- as.data.frame(OrderSample %>% 
                           group_by(CountryName) %>%
                           summarise(revenue= sum(Amount)) %>%
                           arrange(desc(revenue)))

MapData<-map_data("world")

MapData$region<- toupper(MapData$region)

data <- merge(ConRev, 
              MapData, by.x = "CountryName", by.y = "region") %>%
  arrange(CountryName)

ggplot()+ 
  # world map
  geom_polygon(data = MapData, 
               aes(x=long, y = lat, group = group),
               fill = "grey") +
  # custom map
  geom_polygon(data = data,
               aes(x=long, y = lat, group = group, , fill = round(revenue,0)))+
scale_fill_distiller(palette = "Spectral") + 
  labs(fill = "revenue") + 
  theme_map() +
  theme(legend.position = "top")

##Insights: Highest sales of carpets is from USA followed by UK. 
## Other countries such as Romania, Australia, India, Canada and South Africa are also major contributors to the revenue.


## 5.Customer and revenue generated 

CustRev<- as.data.frame(OrderSample %>% 
                         group_by(CustomerCode) %>%
                         summarise(Custrevenue= round(sum(Amount),0)) %>%
                         arrange(desc(Custrevenue)))

ggplot(CustRev, aes(x=CustomerCode, y=round(Custrevenue / 1e3, 1), size=Custrevenue, colour=CustomerCode, fill=CustomerCode )) +
  geom_point(alpha=0.7, shape=21, color="black") +
  scale_fill_viridis(discrete=TRUE, guide=FALSE, option="B") +
  scale_size(range = c(.1, 20), name="Custrevenue") +
  theme_ipsum() +
  ylab("Revenue") +
  xlab("Customer Type") +
  theme(axis.text.x=element_blank(),
        axis.ticks.x=element_blank(),
        axis.text.y=element_blank(),
        axis.ticks.y=element_blank()) +
  geom_text(aes(label= paste(format(round(Custrevenue / 1e3, 1), trim = TRUE), "K")), nudge_y= -3, nudge_x= -2, size=2.5)

## Insight: Sales from the customer “TGT” is remarkably high making him the most important client. 
## There are about 37 customers of which top 20 are major contributors to the revenue. 


### Some model options ######################################################### 

# Champo Carpets can use a variety of analytics and machine learning algorithms to solve their problems and create value.
# 1. Classification: Champo Carpets can use classification algorithms to classify their customers based on their demographics, purchasing habits, and other relevant factors. This can help them identify different customer segments and tailor their marketing strategies accordingly.
# 2. Regression: Regression algorithms can be used by Champo Carpets to forecast future sales and demand for different types of carpets. This can help them optimize their inventory levels, production schedules, and pricing strategies.
# 3. Clustering: Clustering algorithms can be used by Champo Carpets to group similar customers together based on their purchasing patterns. This can help them identify common customer needs and preferences and develop targeted marketing campaigns.
# 4. Recommender systems: Recommender systems can be used by Champo Carpets to recommend carpets to customers based on their past purchases, browsing history, and other relevant factors. This can help them increase customer engagement and loyalty.
# 5. Natural Language Processing: Natural Language Processing (NLP) algorithms can be used by Champo Carpets to analyze customer feedback and reviews. This can help them identify common customer complaints and improve their products and services accordingly.
# Overall, the use of analytics and machine learning algorithms can help Champo Carpets make more informed business decisions, improve customer satisfaction, and increase revenue.

### Modeling #########################################################

#Normalizing the numeric variables
library(dplyr)
normalization <- function(x) {
  (x - min(x)) / (max(x) - min(x))
}
DataOnSample <- DataOnSample %>% mutate_if(is.numeric, normalization)
summary(DataOnSample)

## Preparing the data ----
# Dropping the "China" column since it only had 1 order and was imparing on the 
# creating of the models, since either training or testing would be missing an iniput for it

DataOnSample <- subset(DataOnSample, CustomerCode != "B-2", select = -China)


# creating a dataframe without the "dummy variables"

DataOnSample_NoD <- subset(DataOnSample,select = c("CustomerCode", "CountryName", "QtyRequired",
                                                    "ITEM_NAME", "ShapeName", "AreaFt",
                                                    "Order.Conversion"))
# creating a dataframe with only the "dummy variables" and dropping the "customerCode" variable

DataOnSample_D <- subset(DataOnSample,select = c("USA", "UK", "Italy", "Belgium",
                                                 "Romania", "Australia", "India", "Poland",
                                                 "Canada", "UAE", "Sounth_Africa","Brazil",
                                                 "Israel", "QtyRequired","Hand.Tufted", "Durry",
                                                 "Double.Back", "Hand.Woven", "Knotted","Jacquard",
                                                 "Handloom", "Other", "REC", "Round", "Square",
                                                 "AreaFt", "Order.Conversion"))
DataOnSample_D$USA <- as.numeric(DataOnSample_D$USA) - 1
DataOnSample_D$UK <- as.numeric(DataOnSample_D$UK) 
DataOnSample_D$Italy <- as.numeric(DataOnSample_D$Italy) 
DataOnSample_D$Belgium <- as.numeric(DataOnSample_D$Belgium) 
DataOnSample_D$Romania <- as.numeric(DataOnSample_D$Romania)
DataOnSample_D$Australia <- as.numeric(DataOnSample_D$Australia)
DataOnSample_D$India <- as.numeric(DataOnSample_D$India) - 1
DataOnSample_D$Poland <- as.numeric(DataOnSample_D$Poland)
DataOnSample_D$Canada <- as.numeric(DataOnSample_D$Canada)
DataOnSample_D$UAE <- as.numeric(DataOnSample_D$UAE) 
DataOnSample_D$Sounth_Africa <- as.numeric(DataOnSample_D$Sounth_Africa) 
DataOnSample_D$Brazil <- as.numeric(DataOnSample_D$Brazil) 
DataOnSample_D$Israel <- as.numeric(DataOnSample_D$Israel) 
DataOnSample_D$Hand.Tufted <- as.numeric(DataOnSample_D$Hand.Tufted) - 1
DataOnSample_D$Durry <- as.numeric(DataOnSample_D$Durry)
DataOnSample_D$Double.Back <- as.numeric(DataOnSample_D$Double.Back) - 1
DataOnSample_D$Hand.Woven <- as.numeric(DataOnSample_D$Hand.Woven)
DataOnSample_D$Knotted <- as.numeric(DataOnSample_D$Knotted) 
DataOnSample_D$Jacquard <- as.numeric(DataOnSample_D$Jacquard)
DataOnSample_D$Handloom <- as.numeric(DataOnSample_D$Handloom)
DataOnSample_D$Other <- as.numeric(DataOnSample_D$Other)
DataOnSample_D$REC <- as.numeric(DataOnSample_D$REC) - 1
DataOnSample_D$Round <- as.numeric(DataOnSample_D$Round) 
DataOnSample_D$Square <- as.numeric(DataOnSample_D$Square) 

# creating a version of the table w/o dummys with balanced data:
table(DataOnSample$Order.Conversion)
DataOnSample_NoD_Balanc <- downSample(DataOnSample_NoD, DataOnSample_NoD$Order.Conversion)
table(DataOnSample_NoD_Balanc$Order.Conversion)
DataOnSample_NoD_Balanc <- subset(DataOnSample_NoD_Balanc, select = -Class)

# creating a version of the table w/ dummys with balanced data:
table(DataOnSample_D$Order.Conversion)
DataOnSample_D_Balanc <- downSample(DataOnSample_D, DataOnSample_D$Order.Conversion)
table(DataOnSample_D_Balanc$Order.Conversion)
DataOnSample_D_Balanc <- subset(DataOnSample_D_Balanc, select = -Class)

## Partitioning the 4 data tables ----
# Partitioning the unbalanced data w/o dummys into train and test data
set.seed(256)
indx <- sample(2, nrow(DataOnSample_NoD), replace= TRUE, prob = c(0.8, 0.2))
train_LRU <- DataOnSample_NoD[indx == 1, ]
test_LRU <- DataOnSample_NoD[indx == 2, ]

# partitioning the balanced data w/o dummys into training and test data:
set.seed(256)
indx2 <- sample(2, nrow(DataOnSample_NoD_Balanc), replace= TRUE, prob = c(0.8, 0.2))
train_LRB <- DataOnSample_NoD_Balanc[indx2 == 1, ]
test_LRB <- DataOnSample_NoD_Balanc[indx2 == 2, ]

# Partitioning the unbalanced data w/ dummys into train and test data
set.seed(256)
indx3 <- sample(2, nrow(DataOnSample_D), replace= TRUE, prob = c(0.8, 0.2))
train_DU <- DataOnSample_D[indx3 == 1, ]
test_DU <- DataOnSample_D[indx3 == 2, ]

# Partitioning the balanced data w/ dummys into train and test data
set.seed(256)
indx4 <- sample(2, nrow(DataOnSample_D_Balanc), replace= TRUE, prob = c(0.8, 0.2))
train_DB <- DataOnSample_D_Balanc[indx4 == 1, ]
test_DB <- DataOnSample_D_Balanc[indx4 == 2, ]


## Logistic Regression model ----
#1. Unbalanced data
LR_Model <- glm(Order.Conversion ~ ., data = train_LRU, family = binomial)
summary(LR_Model)

LR_Pred <- predict(LR_Model, newdata = test_LRU, type = "response")
binary_pred <- ifelse(LR_Pred > 0.5, 1, 0)

# Compute the confusion matrix
confusionMatrix(table(binary_pred, test_LRU$Order.Conversion))

# Our logistic regression model had an accuracy of 89%

#2. Balanced Data
LR_Model2 <- glm(Order.Conversion ~ ., data = train_LRB, family = binomial, maxit = 1000)
summary(LR_Model2)

LR_Pred2 <- predict(LR_Model2, newdata = test_LRB, type = "response")
binary_pred2 <- ifelse(LR_Pred2 > 0.5, 1, 0)

# Compute the confusion matrix
confusionMatrix(table(binary_pred2, test_LRB$Order.Conversion))

# Our logistic regression model on balanced data had an accuracy of 82%

## Decision Tree Model ----
#1. Unbalanced data
library(rpart)

DT_Model <- rpart(Order.Conversion ~ ., train_LRU)
DT_Model
rpart.plot(DT_Model)

DT_Pred <- predict(DT_Model, newdata = test_LRU, type = "class")
confusion_matrix <- confusionMatrix(DT_Pred, test_LRU$Order.Conversion)
confusion_matrix

# Our Decision Tree model on unbalanced data had an accuracy of 91%

#2. Balanced data
DT_Model2 <- rpart(Order.Conversion ~ ., train_LRB, control = rpart.control(cp = 0.01))
DT_Model2
rpart.plot(DT_Model2)

DT_Pred2 <- predict(DT_Model2, newdata = test_LRB, type = "class")
confusion_matrix2 <- confusionMatrix(DT_Pred2, test_LRB$Order.Conversion)
confusion_matrix2

# Our Decision Tree model on unbalanced data had an accuracy of 84%

## Random Forest Model ----
#1. Unbalanced data
library(randomForest)

RF_Model <- randomForest(Order.Conversion ~ ., data = train_LRU, ntree = 500)
summary(RF_Model)
importance(RF_Model)
RF_Pred <- predict(RF_Model, newdata = test_LRU)
accuracy <- mean(RF_Pred == test_LRU$Order.Conversion)
accuracy

# Our Random Forest model on unbalanced data had an accuracy of 93%

#2. Balanced data

RF_Model2 <- randomForest(Order.Conversion ~ ., data = train_LRB, ntree = 500)
summary(RF_Model2)
importance(RF_Model2)
RF_Pred2 <- predict(RF_Model2, newdata = test_LRB)
accuracy2 <- mean(RF_Pred2 == test_LRB$Order.Conversion)
accuracy2

# Our Random Forest model on balanced data had an accuracy of 88%

## Neural Network ----
#1. Unbalanced data
library(nnet)
library(NeuralNetTools)
NN_Model <- nnet(Order.Conversion ~ ., data = train_DU, linout = FALSE, 
                size = 10, decay = 0.01, maxit = 1000)
plotnet(NN_Model)
NN_Pred <- predict(NN_Model, test_DU)
NN_Pred_Class <- as.factor(predict(NN_Model, test_DU, type = "class"))
confusionMatrix(NN_Pred_Class, test_DU$Order.Conversion)
varImp(NN_Model)

# Our Neural Network model on unbalanced data had an accuracy of 91%

#2. Balanced data
NN_Model2 <- nnet(Order.Conversion ~ ., data = train_DB, linout = FALSE, 
                 size = 10, decay = 0.01, maxit = 1000)
plotnet(NN_Model2)
NN_Pred2 <- predict(NN_Model2, test_DB)
NN_Pred_Class2 <- as.factor(predict(NN_Model2, test_DB, type = "class"))
confusionMatrix(NN_Pred_Class2, test_DB$Order.Conversion)
varImp(NN_Model2)

# Our Neural Network model on balanced data had an accuracy of 83%




