
library(readxl)
library(dplyr)
library(ggplot2)
library(lmtest)
library(data.table)

weekly_visits <- read_excel("Web Analytics Case Student Spreadsheet.xls", sheet = "Weekly Visits", range = "A5:H71")
financials <- read_excel("Web Analytics Case Student Spreadsheet.xls", sheet = "Financials", range = "A5:E71")
lbs_sold <- read_excel("Web Analytics Case Student Spreadsheet.xls", sheet = "Lbs. Sold", range = "A5:B295")
daily_visits <- read_excel("Web Analytics Case Student Spreadsheet.xls", sheet = "Daily Visits", range = "A5:B467")
final <- read_excel("Web Analytics Case Student Spreadsheet.xls", sheet = "Final", range = "A1:L67")
View(weekly_visits)
View(financials)
View(lbs_sold)
View(daily_visits)
View(final)

# Display the summary of full data
summary(weekly_visits)
summary(financials)
summary(lbs_sold)
summary(daily_visits)
summary(final)

# Bar Chart for different variables
# Visits
visits_chart <- ggplot(weekly_visits, aes(x = final$`Week (2008-2009)`, y = final$Visits)) + 
  geom_bar(stat="identity", width = 0.6, fill = "Black") + 
  labs(title = "Number of Visits Per Week (2008-2009)", x = "Week (2008-2009)", y = "Visits") +
  theme(axis.text.x = element_text (size = 8, angle = 90, vjust = 0.8))
visits_chart

# Unique Visits
uv_chart <- ggplot(weekly_visits, aes(`Week (2008-2009)`, `Unique Visits`)) +
  geom_bar(stat="identity", width = 0.5, fill="Red") +
  labs(title="Unique Visit Per Week (2008-2009)") +
  theme(axis.text.x = element_text(angle=90, vjust=0.8))
uv_chart

# Revenue
rev_chart <- ggplot(financials, aes(`Week (2008-2009)`, `Revenue`))+ 
  geom_bar(stat="identity", width = 0.5, fill="Orange") +
  labs(title="Revenue Per Week (2008-2009)") +
  theme(axis.text.x = element_text(angle=90, vjust=0.8))
rev_chart

# Profit
profit_chart <- ggplot(financials, aes(`Week (2008-2009)`, `Profit`)) + 
  geom_bar(stat="identity", width = 0.5, fill="Green") +
  labs(title="Profit Per Week (2008-2009)") +
  theme(axis.text.x = element_text(angle=90, vjust=0.8))
profit_chart

#Lbs. Sold
lbs_chart <- ggplot(financials, aes(`Week (2008-2009)`, `Lbs. Sold`)) + 
  geom_bar(stat="identity", width = 0.5, fill="Blue") +
  labs(title="Lbs. Sold Per Week (2008-2009)") +
  theme(axis.text.x = element_text(angle=90, vjust=0.8))
lbs_chart

# Basic histogram for different variables
# Visits
visits_histogram <- hist(weekly_visits$Visits,
                         main="Histogram for Visits per Week", 
                         xlab="Visits per Week")

# Unique visits
uv_histogram <- hist(weekly_visits$`Unique Visits`,
                     main="Histogram for Unique Visits per Week", 
                     xlab="Unique Visits per Week")

# Revenue
rev_histogram <- hist(financials$Revenue,
                      main="Histogram for Revenue per Week", 
                      xlab="Revenue per Week")

# Profit
profit_histgram <- hist(financials$Profit,
                        main="Histogram for Profit per Week", 
                        xlab="Profit per Week")

#Lbs. Sold
lbs_histogram <- hist(lbs_sold$`Lbs. Sold`,
                      main="Histogram for Lbs. Sold per Week", 
                      xlab="Lbs. Sold per Week")

# Create Line Chart - Bounce rate per Week
bounce_rate <-ggplot(data = weekly_visits, 
                     aes(weekly_visits$`Week (2008-2009)`, weekly_visits$`Bounce Rate`, group=1)) + 
                     labs(title = "Bounce Rate per Week (2008-2009)", x = "Week (2008-2009)", y = "Bounce Rate") + 
                     geom_line() + geom_point()
bounce_rate

# Pie Chart of Browsers with Percentages
slices <- c(76,19,5) 
lbls <- c("Internet Explorer","Firefox","Others")
pct <- round(slices/sum(slices)*100)
lbls <- paste(lbls, pct) # add percents to labels 
lbls <- paste(lbls,"%",sep="") # ad % to labels 
pie(slices,labels = lbls, col=rainbow(length(lbls)),
    main="Browsers used to access website")

# Pie Chart of Regions
slices <- c(22.616,17.509,6.776,5.214,3.228,2.721,2.589,1.968,1.538,1.427) 
lbls <- c("South America","Northern America","Central America","Western Europe","Eastern Asia","Northern Europe","Southern Asia","Southern Europe","Eastern Europe")
pie(slices,labels = lbls, col=rainbow(length(lbls)),
    main="Regions with access to the website")

# Create boxplots to detect outliers
par(mfrow = c(2, 3))  # divide graph area
boxplot(weekly_visits$Visits, main="Visits") # box plot for "Visits"
boxplot(weekly_visits$`Unique Visits`, main="Unique Visits") # box plot for "Unique Visits"
boxplot(financials$Revenue, main="Revenue") # box plot for "Revenue"
boxplot(financials$Profit, main="Profit") # box plot for "Profit"
boxplot(financials$`Lbs. Sold`, main="Lbs. Sold") # box plot for "Lbs. Sold"

# Create a scatterplot to display the correlation between Lbs. Sold and Revenue
ggplot(final, aes(`Lbs. Sold`, Revenue)) +
  geom_point(shape = 16, size = 3.5, show.legend = FALSE) +
  theme_minimal() + geom_smooth(method="lm") + 
  labs(title = "Revenue vs. Lbs. Sold")
# Calculate correlation between Lbs. Sold and Revenue
cor(final$`Lbs. Sold`, final$Revenue) 

# Create a scatterplot to display the correlation between Visits and Revenue
ggplot(final, aes(Visits, Revenue)) +
  geom_point(shape = 16, size = 3.5, show.legend = FALSE) +
  theme_minimal() + labs(title = "Revenue vs. Visits")
# Calculate correlation between Visits and Revenue
cor(final$Visits, final$Revenue)

# Revenue per period
`5_weeks_revenue` <- c(sum(final$Revenue[1:5]),
                      sum(final$Revenue[6:10]),
                      sum(final$Revenue[11:14]),
                      sum(final$Revenue[15:21]),
                      sum(final$Revenue[22:26]),
                      sum(final$Revenue[27:31]),
                      sum(final$Revenue[32:35]),
                      sum(final$Revenue[36:41]),
                      sum(final$Revenue[42:46]),
                      sum(final$Revenue[47:52]),
                      sum(final$Revenue[53:57]),
                      sum(final$Revenue[58:62]),
                      sum(final$Revenue[63:66]))

rev_table <- data.table(`5_weeks` = seq(1,13,1), Revenue = `5_weeks_revenue`, Group = c(rep("Initial",3),
                                                                                        rep("Pre-promo",4),
                                                                                        rep("Promo",3),
                                                                                        rep("Post-promo",3)))

ggplot(data = rev_table, aes(x=rev_table$`5_weeks`, y=rev_table$Revenue/1000, fill = Group))  + 
  geom_bar(stat="identity", colour="Black") + 
  scale_x_discrete("5 Weeks Period") + 
  scale_y_continuous("Revenue(in thousand)") + 
  scale_fill_discrete(breaks=c("Initial","Pre-promo","Promo", "Post-promo")) +
  ggtitle(label = "Revenue") +
  theme(text = element_text(size=30),
        legend.title = element_text(size=30),
        legend.text = element_text(size=30),
        legend.position = "bottom")

# Visits
`5_weeks_visits` <- c(sum(final$Visits[1:5]),
                      sum(final$Visits[6:10]),
                      sum(final$Visits[11:14]),
                      sum(final$Visits[15:21]),
                      sum(final$Visits[22:26]),
                      sum(final$Visits[27:31]),
                      sum(final$Visits[32:35]),
                      sum(final$Visits[36:41]),
                      sum(final$Visits[42:46]),
                      sum(final$Visits[47:52]),
                      sum(final$Visits[53:57]),
                      sum(final$Visits[58:62]),
                      sum(final$Visits[63:66]))

visits_table <- data.table(`5_weeks` = seq(1,13,1), Visits = `5_weeks_visits`, Group = c(rep("Initial",3),
                                                                                      rep("Pre-promo",4),
                                                                                      rep("Promo",3),
                                                                                      rep("Post-promo",3)))

ggplot(data = visits_table, aes(x=visits_table$`5_weeks`, y=visits_table$Visits, fill = Group))  + 
  geom_bar(stat="identity", colour="Black") + 
  scale_x_discrete("5 Weeks Period") + 
  scale_y_continuous("Visits") + 
  scale_fill_discrete(breaks=c("Initial","Pre-promo","Promo", "Post-promo")) +
  ggtitle(label = "Visits") +
  theme(text = element_text(size=30),
        legend.title = element_text(size=30),
        legend.text = element_text(size=30),
        legend.position = "bottom")

# Lbs. Sold
`5_weeks_visits` <- c(sum(final$`Lbs. Sold`[1:5]),
                      sum(final$`Lbs. Sold`[6:10]),
                      sum(final$`Lbs. Sold`[11:14]),
                      sum(final$`Lbs. Sold`[15:21]),
                      sum(final$`Lbs. Sold`[22:26]),
                      sum(final$`Lbs. Sold`[27:31]),
                      sum(final$`Lbs. Sold`[32:35]),
                      sum(final$`Lbs. Sold`[36:41]),
                      sum(final$`Lbs. Sold`[42:46]),
                      sum(final$`Lbs. Sold`[47:52]),
                      sum(final$`Lbs. Sold`[53:57]),
                      sum(final$`Lbs. Sold`[58:62]),
                      sum(final$`Lbs. Sold`[63:66]))

lbs_table <- data.table(`5_weeks` = seq(1,13,1), `Lbs. Sold` = `5_weeks_visits`, Group = c(rep("Initial",3),
                                                                                           rep("Pre-promo",4),
                                                                                           rep("Promo",3),
                                                                                           rep("Post-promo",3)))

ggplot(data = lbs_table, aes(x=lbs_table$`5_weeks`, y=lbs_table$`Lbs. Sold`, fill = Group))  + 
  geom_bar(stat="identity", colour="Black") + 
  scale_x_discrete("5 Weeks Period") + 
  scale_y_continuous("Lbs. Sold") + 
  scale_fill_discrete(breaks=c("Initial","Pre-promo","Promo", "Post-promo")) +
  ggtitle(label = "Lbs. Sold") +
  theme(text = element_text(size=30),
        legend.title = element_text(size=30),
        legend.text = element_text(size=30),
        legend.position = "bottom")

# Build a linear regression model
linearMod <- lm(Revenue ~ Visits + `Lbs. Sold`, data = final)
summary(linearMod)

# Basic Linear Model Representation for Profit
regression_table <- data.table(`5_weeks` = seq(1,13,1), Revenue = `5_weeks_revenue`[1:13])

ggplot(regression_table, aes(x = `5_weeks`, y = Revenue/1000)) + 
  stat_smooth(method = "glm", formula = y ~ x,alpha = 0.2, size = 2, fill = "#E69F00", colour = "red") +
  geom_point(size = 3) +
  scale_x_continuous("5 Weeks Period", breaks=c(1:13), labels=c(1:13),limits=c(1,13)) + 
  scale_y_continuous("Revenue (in thousands)") +
  ggtitle("Revenue Linear Model") +
  theme(text = element_text(size = 22),
        legend.title = element_text(size=30))

