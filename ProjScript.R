library(xlsx)
library(ggplot2)
library(dplyr)

getwd()
setwd("../Desktop/Uni/6th sem shit/365/Project")
veteranDF <- read.xlsx("Suicide.xlsx", "Veteran",header = TRUE,startRow = 2,endRow = 22, cols=c(1,7,10,11,14))
veterans <- read.xlsx("Suicide.xlsx", "Veteran",header = TRUE,startRow = 2,endRow = 22, cols=c(1,7,10,11,14))
nonveterans <- read.xlsx("Suicide.xlsx", "Non-Veteran", header = T, startRow = 2, endRow = 22)
veteranRace1 <- read.xlsx("Suicide.xlsx", "Veteran Race & Ethnicity", header = F, startRow = 2, endRow = 23)
veteranAge <- read.xlsx("Suicide.xlsx", "Veteran", header = T, startRow = 25, endRow = 125)
vha <- read.xlsx("Suicide.xlsx", "Recent Veteran VHA User", header = T, startRow = 2, endRow = 22)
maleVfemalePop <- read.xlsx("Suicide.xlsx", "Veteran", header = T, startRow = 2, endRow = 22)

veteranRace2 <- veteranRace1

header1 <- as.character(veteranRace1[1,])
header2 <- as.character(veteranRace1[2,])

header1 <- as.character(veteranRace1[1,])
header2 <- as.character(veteranRace1[2,])

veteranRace2
veteranRace1
str(veteranRace1)
vha

colnames(veteranRace1) <- header1
colnames(veteranRace1)[!is.na(header2)] <- header2[!is.na(header2)]

colnames(veteranRace2) <- header1
colnames(veteranRace2)[!is.na(header2)] <- header2[!is.na(header2)]

veteranRace1 <- veteranRace1[-c(1,2),]
veteranRace2 <- veteranRace2[-c(1,2)]

header1 <- header1[header1 != "NA"]
header1 <- header1[seq(2, 12, 2)]
header1

veteranRace1 <- veteranRace1[,c(1,seq(2, 12, 2))]
veteranRace2 <- veteranRace2[,c(seq(1, 11, 2))]

veteranRace <- veteranRace1
veteranRace2 <- veteranRace2[-c(1,2),-c(6)]

veteranRace2

names(veteranRace) <- c("Years",header1)
names(veteranRace1) <- c("Years",header1)
names(veteranRace2) <- c(header1[-c(6)])

veteranRace1

veteranRace2

veteranRace2 <- cbind("Years" = veteranRace$Years, veteranRace2)

veteranRace2 <- as.data.frame(lapply(veteranRace2, as.numeric))

veteranRace2 <- replace(veteranRace2, veteranRace2 == "--", 0)

veteranRace1
veteranRaceDF <- as.data.frame(veteranRace[-c(1)])
vetRaceDF1 <- as.data.frame(veteranRace1)

vetRaceDF1 <- replace(veteranRaceDF, veteranRaceDF == "<10", 5)
veteranRaceDF <- replace(veteranRaceDF, veteranRaceDF == "<10", 5)

veteranRaceDF <- na.omit(veteranRaceDF)

str(vetRaceDF1)
vetRaceDF1
veteranRaceDF

veteranRaceDF <- apply(veteranRaceDF, 2, as.numeric)

values <- colSums(veteranRaceDF)

data <- data.frame(
  name = names(veteranRace)[-c(1)],
  value = values
)

veteranAndNonveteran <- data.frame(veterans$Year.of.Death, veterans$Veteran.Age..and.Sex..Adjusted.Rate.per.100.000, nonveterans$Non.Veteran.Age..and.Sex..Adjusted.Rate.per.100.000)

veteranAndNonveteran

veteranAge

vetAge <- veteranAge %>%
  group_by(Age.Group) %>%
  summarise(VeteranSuicide = sum(Veteran.Suicide.Deaths))

vetAge[5,2] <- 127337
vetAgeDF <- as.data.frame(vetAge)
vetAgeDF <- vetAgeDF[-5,]
vetAgeDF
  
vetAgeDF$percent <- round(vetAgeDF$VeteranSuicide / sum(vetAgeDF$VeteranSuicide) * 100,2)

pie_chart <- ggplot(vetAgeDF, aes(x = "", y = VeteranSuicide, fill = Age.Group)) +
  geom_bar(stat = "identity", width = 1) +
  coord_polar(theta = "y") +
  labs(title = "Veteran Suicide Rate Based on Age Group") +
  theme_void() +
  geom_text(aes(label = paste0(percent, "%")), position = position_stack(vjust = 0.5))
veteranAndNonveteranDF <- data.frame(
  veterans.Year.of.Death = veteranAndNonveteran$veterans.Year.of.Death,
  veterans.Veteran.Age..and.Sex..Adjusted.Rate.per.100.000 = veteranAndNonveteran$veterans.Veteran.Age..and.Sex..Adjusted.Rate.per.100.000,
  nonveterans.Non.Veteran.Age..and.Sex..Adjusted.Rate.per.100.000 = veteranAndNonveteran$nonveterans.Non.Veteran.Age..and.Sex..Adjusted.Rate.per.100.000
)
veteranAndNonveteran <- lapply(veteranAndNonveteran, as.vector)
veteranAndNonveteran <- lapply(veteranAndNonveteran, function(x){
  if (is.factor(x)) as.numeric(x) else x
})

veteranAndNonveteranDF <- fortify(as.data.frame(veteranAndNonveteran))

print(pie_chart)

create_bar_plot <- function(data, x_col, y_col) {
  bar_plot <- ggplot(data, aes(x = x_col, y = y_col)) +
    geom_bar(stat = "identity", fill = "steelblue") +
    labs(title = "Bar Plot") +
    xlab(x_col) +
    ylab(y_col) +
    theme_minimal()
  
  print(bar_plot)
}

create_pie_chart <- function(data, values_col, labels_col, title) {
  data$percent <- round(data[[values_col]] / sum(data[[values_col]]) * 100, 2)
  
  pie_chart <- ggplot(data, aes(x = "", y = .data[[values_col]], fill = .data[[labels_col]])) +
    geom_bar(stat = "identity", width = 1) +
    coord_polar(theta = "y") +
    labs(title = title) +
    theme_void() +
    geom_text(aes(label = paste0(percent, "%")), position = position_stack(vjust = 0.5))
  
  print(pie_chart)
}
create_line_plot <- function(data, x_col, y_col1, y_col2, x_label, y_label, title) {
  data <- data.frame(data)
  line_plot <- ggplot(data) +
    geom_line(aes(x = !!sym(x_col), y = !!sym(y_col1), colour = y_col1)) +
    geom_line(aes(x = !!sym(x_col), y = !!sym(y_col2), colour = y_col2)) +
    xlab(x_label) + ylab(y_label) + ggtitle(title) + labs(color = "Category")
  
  print(line_plot)
}

names(veteranAndNonveteranDF) <- c("YearOfDeath","Veterans", "Nonveterans")

veterans
nonveterans
veteranPopDF <- data.frame("Veterans" = veterans$Veteran.Population.Estimate, "Nonveterans" = nonveterans$Non.Veteran.Population.Estimate)

veteranAndNonveteran <- lapply(veteranAndNonveteran, as.vector)
veteranAndNonveteran <- lapply(veteranAndNonveteran, function(x){
  if (is.factor(x)) as.numeric(x) else x
})

veteranAndNonveteranDF <- fortify(as.data.frame(veteranAndNonveteran))

print(pie_chart)


veterans <- veterans[-c(2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 13)]
names(veterans) <- c("YearOfDeath", "Male", "Female")
veterans
veteranAndNonveteranDF
str(veteranAndNonveteran)
head(veteranAndNonveteran)
typeof(veteranAndNonveteranDF)

vha <- data.frame("Year of Death" = vha$Year.of.Death,"VHA Veterans" = veteranAndNonveteranDF$Veterans, "Non VHA Veterans" = vha$VHA.Veteran.Age..and.Sex..Adjusted.Rate.per.100.000)
vha
nonveterans
# 1 - 5
create_bar_plot(data, data$name, data$value)
create_pie_chart(vetAgeDF, "VeteranSuicide", "Age.Group", "Veteran Suicide Rate Based on Age Group")
create_line_plot(veteranAndNonveteranDF, "YearOfDeath", "Veterans", "Nonveterans", "Year of Death", "Suicide Deaths", "Comparison of Suicide Rate of Veterans and Non-Veterans")
create_line_plot(veterans, "YearOfDeath", "Male", "Female", "Year of Death", "Age Adjusted Rate per 100,000", "Comparison of the Male and Female Veteran Suicide Rates")
create_line_plot(vha, x_col = "Year.of.Death", y_col1 = "VHA.Veterans", y_col2 = "Non.VHA.Veterans", x_label = "Year of Death", "Suicide Rate",title = "Comparison between VHA utilizing and non-utilizing Veterans")
# *******
# Next are altercations on the data types just to see what works
# But I do hope that this will work when it runs on other computers
# After each plot there would be some work to be done on the datasets so that it matches the
# requirement of the function
veteranDF
veteranPopulation <- list("VeteranPop" = veteranDF[c(3)], "NonveteranPop" = nonveterans[,c(3)])
typeof(data)
data
typeof(veteranPopulation)
veteranPopulation
veteranPopulation <- list("Veteran Population" = veteranDF$Veteran.Population.Estimate, "Non-veteran Population" = nonveterans$Non.Veteran.Population.Estimate)
veteranPopulation <- colSums(veteranPopulation)
names(veteranPopulation) <- c("Veteran Population", "Non-veteran Population")
veteranPopulation$Total <- veteranPopulation$Veteran.Population.Estimate + veteranPopulation$NonveteranPop
veteranPopulation[1]

veteranPopDF
sums <- colSums(veteranPopDF)
pop <- data.frame(names = names(veteranPopDF), values = sums)
pop$total <- sum(pop$values)
pop

# Number 6
create_pie_chart(pop, "values", "names", "Population Comparison")
# **********

veterans <- maleVfemalePop
maleVfemalePop
maleVfemalePop <- maleVfemalePop[c(8,12)]
typeof(maleVfemalePop)
mVf <- data.frame(maleVfemalePop)
mVf1 <- mVf
mVf1 <- colSums(mVf)
mVf1
names(mVf1)
df <- data.frame(
  names = names(mVf1),
  values = unlist(mVf1),
  total = sum(mVf1)
)
df

# Number 7
create_pie_chart(df, "values", "names", "Male Vs Female Veteran Population Comparison")
# ***********

nonveterans
nonMvF <- nonveterans[c(1, 7, 11)]
typeof(nonMvF)
names(nonMvF) <- c("YearOfDeath", "Male", "Female")
nonMvF
nonMvFdf <- data.frame(names = names(nonMvF), values = nonMvF)

# Number 8
create_line_plot(nonMvF, "YearOfDeath", "Male", "Female",x_label = "Year Of Death", y_label = "Suicide Deaths", title = "Male Vs Female Non-veteran Suicide Deaths")
#**********

veteranAge
femaleVeteranAge <- data.frame("YearOfDeath" = veteranAge$Year.of.Death, "Age Group" = veteranAge$Age.Group.2, "Female Suicide deaths" = veteranAge$Female.Veteran.Suicide.Deaths)
femaleVeteranAge
femaleVeteranAge <- femaleVeteranAge %>%
  replace(.==".", NA)
femaleVeteranAge <- na.omit(femaleVeteranAge)
femaleVeteranAge
femaleVeteranAge <- as.data.frame(femaleVeteranAge)
femaleVeteranAge$Female.Suicide.Deaths <- ifelse(is.na(femaleVeteranAge$Female.Suicide.deaths) | femaleVeteranAge$Female.Suicide.deaths == "", NA, as.numeric(femaleVeteranAge$Female.Suicide.deaths))
typeof(femaleVeteranAge$Female.Suicide.deaths)
femaleVeteranAge
# Number 9
line_plot <- ggplot(femaleVeteranAge) + 
  geom_line(aes(x = YearOfDeath, y = as.numeric(Female.Suicide.Deaths), group = Age.Group, color = Age.Group)) +
  xlab("Year of Death") +
  ylab("Female Suicide Deaths") +
  ggtitle("Female Suicide Deaths by Age Group")

print(line_plot)
#************

femaleVeteranAge
str(femaleVeteranAge)

veteranAge
maleVeteranAge <- data.frame("YearOfDeath" = veteranAge$Year.of.Death, "Age Group" = veteranAge$Age.Group, "Male Suicide Deaths" = veteranAge$Male.Veteran.Suicide.Deaths)
maleVeteranAge

# Number 10
line_plot <- ggplot(maleVeteranAge) + 
  geom_line(aes(x = YearOfDeath, y = as.numeric(Male.Suicide.Deaths), group = Age.Group, color = Age.Group)) +
  xlab("Year of Death") +
  ylab("Male Suicide Deaths") +
  ggtitle("Male Suicide Deaths by Age Group")

print(line_plot)
#**************

veteranRace <- veteranRace1
veteranRace <- replace(veteranRace, veteranRace == "<10", NA)
veteranRace <- replace(veteranRace, is.na(veteranRace), 5)
veteranRace
veteranRace <- as.data.frame(sapply(veteranRace, as.numeric))

#Number 11
line_plot <- ggplot(veteranRace, aes(x = Years)) + 
  geom_line(aes(y = White, color = "White")) +
  geom_line(aes(y = Black, color = "Black")) +
  geom_line(aes(y = ` American Indian/Alaskan Native`, color = "American Indian/Alaskan Native")) + 
  geom_line(aes(y = `Asian, Native Hawaiian, or Pacific Islander`, color = "Asian, Native Hawaiian, or Pacific Islander")) +
  geom_line(aes(y = `Multiple Race`, color = "Multiple Race")) + 
  geom_line(aes(y = `Unknown Race`, color = "Unknown Race")) +
  xlab("Year of Death") +
  ylab("Suicide Deaths") +
  ggtitle("Suicide Deaths by Race")

print(line_plot)
#**************

str(veteranRaceDF)

library(moments)

calculate_summary <- function(data, column) {
  values <- data[[column]]
  n <- length(values)
  mean_value <- mean(values)
  sd_value <- sd(values)
  median_value <- median(values)
  min_value <- min(values)
  max_value <- max(values)
  range_value <- max_value - min_value
  se_value <- sd_value / sqrt(n)
  q1 <- quantile(values, 0.25)
  q3 <- quantile(values, 0.75)
  skewness_value <- skewness(values)
  kurtosis_value <- kurtosis(values)
  
  summary_df <- data.frame(n = n,
                           mean = mean_value,
                           sd = sd_value,
                           median = median_value,
                           min = min_value,
                           max = max_value,
                           range = range_value,
                           q1 = q1,
                           q3 = q3,
                           skewness = skewness_value,
                           kurtosis = kurtosis_value,
                           se = se_value)
  
  return(summary_df)
}
names(veteranAndNonveteranDF)

vetAndNon <- data.frame(
  group = rep(c("Veterans", "Nonveterans"), each = length(veteranAndNonveteranDF$Veterans)),
  suicide_rate = c(veteranAndNonveteranDF$Veterans, veteranAndNonveteranDF$Nonveterans)
)

vet <- data.frame(
  group = rep(c("Male", "Female"), each = length(veterans$Male.Veteran.Crude.Rate.per.100.000)),
  suicide_rate = c(veterans$Male.Veteran.Age..Adjusted.Rate.per.100.000, veterans$Female.Veteran.Age..Adjusted.Rate.per.100.000)
)

vetVHA <- data.frame(
  group = rep(names(vha)[-c(1)], each = length(vha$VHA.Veterans)),
  suicide_rate = c(vha$VHA.Veterans, vha$Non.VHA.Veterans)
)

#Descriptive Statistics
veterans_data <- subset(vetAndNon, group == "Veterans")
veterans_summary <- calculate_summary(veterans_data, "suicide_rate")
nonveterans_data <- subset(vetAndNon, group == "Nonveterans")
nonveterans_summary <- calculate_summary(nonveterans_data, "suicide_rate")
maleVet_data <- subset(vet, group == "Male")
maleVet_summary <- calculate_summary(maleVet_data, "suicide_rate")
femaleVet_data <- subset(vet, group == "Female")
femaleVet_summary <- calculate_summary(femaleVet_data, "suicide_rate")
vhaVet_data <- subset(vetVHA, group == "VHA.Veterans")
vhaVet_summary <- calculate_summary(vhaVet_data, "suicide_rate")
nonVhaVet_data <- subset(vetVHA, group == "Non.VHA.Veterans")
nonVhaVet_summary <- calculate_summary(nonVhaVet_data, "suicide_rate")
summary_summary <- data.frame()
summary_df <- rbind(veterans_summary, nonveterans_summary, maleVet_summary, femaleVet_summary,
                    vhaVet_summary, nonVhaVet_summary)
rownames(summary_df) <- c("Veterans", "Nonveterans", "Male Veterans", "Female Veterans", 
                          "VHA-utilizing Veterans", "Non-VHA-utilizing Veterans")

# Number 1
summary_df
#**********

veteranRace <- replace(veteranRace, veteranRace == "<10", 5)
veteranRace <- as.data.frame(sapply(veteranRace, as.numeric))
str(veteranRace)
veteranRace

vetRace <- data.frame(
  group = rep(names(veteranRace)[-c(1)], each = length(veteranRace$White)),
  suicide_rate = c(veteranRace$White, veteranRace$Black, veteranRace$` American Indian/Alaskan Native`,
                   veteranRace$`Asian, Native Hawaiian, or Pacific Islander`,
                   veteranRace$`Multiple Race`, veteranRace$`Unknown Race`)
)
whiteVet_data <- subset(vetRace, group == "White")
whiteVet_summary <- calculate_summary(whiteVet_data, "suicide_rate")
blackVet_data <- subset(vetRace, group == "Black")
blackVet_summary <- calculate_summary(blackVet_data, "suicide_rate")
amVet_data <- subset(vetRace, group == " American Indian/Alaskan Native")
amVet_summary <- calculate_summary(amVet_data, "suicide_rate")
asianVet_data <- subset(vetRace, group == "Asian, Native Hawaiian, or Pacific Islander")
asianVet_summary <- calculate_summary(asianVet_data, "suicide_rate")
multVet_data <- subset(vetRace, group == "Multiple Race")
multVet_summary <- calculate_summary(multVet_data, "suicide_rate")
unkVet_data <- subset(vetRace, group == "Unknown Race")
unkVet_summary <- calculate_summary(unkVet_data, "suicide_rate")

race_summary <- rbind(whiteVet_summary, blackVet_summary, amVet_summary, asianVet_summary,
                      multVet_summary, unkVet_summary)
rownames(race_summary) <- c("White", "Black", "American Indian/Alaskan Native",
                            "Asian, Native Hawaiian, or Pacific Islander", "Multiple Race",
                            "Unknown Race")
# Number 2
race_summary
# ***********

