library(dplyr)
library(mcr)
library(VCA)

# Simulated Comparison -----------------------------------------------

# Coefficient of variation of 5%
cv <- 0.05

# Random set of values from 10 to 200.
x_ideal <- runif(40, 10, 200)

# Jitter function to add random noise following a normal distribution.
jitter <- function(x, cv) {
  (x + rnorm(1, 0, x * cv))
}

# Create x and y data and save to a data frame.
x <- vapply(x_ideal, jitter, numeric(1), cv)
y <- vapply(x_ideal, jitter, numeric(1), cv)

df <- data.frame(x, y)

write.csv(df, "data.csv")


# Run regressions
reg_paba <- mcreg(x, y, method.reg = "PaBa", method.ci = "bootstrap")
summary(reg_paba)

reg_wdem <- mcreg(x, y, method.reg = "WDeming", method.ci = "jackknife")
summary(reg_wdem)

reg_wdem <- mcreg(x, y,
  method.reg = "WDeming", method.ci = "jackknife",
  error.ratio = 0.8
)
summary(reg_wdem)

# Simulated precision data -----------------------------------------------

# Repeatability of 3%
cv_rep <- 0.03

# Between run imprecision of 1%
cv_run <- 0.01

# Between day imprecision of 2%
cv_day <- 0.02

target <- 100

values <- c()
days <- c()
runs <- c()

for (day in 1:20) {
  for (run in 1:2) {
    for (rep in 1:2) {
      days <- c(days, paste("Day", day))
      runs <- c(runs, paste("Run", run))
      values <- c(
        values,
        target +
          rnorm(1, 0, target * cv_rep) +
          rnorm(1, 0, target * cv_run) +
          rnorm(1, 0, target * cv_day)
      )
    }
  }
}

df <- data.frame(Day = days, Run = runs, Level1 = values)

write.csv(df, "data.csv")
df

# Load previous data -----------------------------------------------------

library(readxl)
library(VCA)

setwd("D:/workspace/METools/method-eval-tools/etc")
df <- as.data.frame(read_excel("mvw-test-data.xlsx", sheet = "Raw"))

vc <- anovaVCA(Level1 ~ Day / Run, Data = df)
vc

vc_ep <- anovaVCA(EP05_A1 ~ Day / Run, Data = df)
vc_ep

# remove rows 8 and 40
rows = c(seq(1,6,1), seq(9,39,1), seq(41,80,1))
df2 <- df[rows,]

vc2 <- anovaVCA(Level1 ~ Day / Run, Data = as.data.frame(df2))
vc2
vc2$aov.tab

# Regression -------------------------------------------------------------
library(mcr)

x = c(10377.5, 4056, 2654, 4747, 1459.5, 5880, 3871, 2461, 1802, 1607.5,
      4329, 7911.5, 1798.5, 6504, 9506.5, 4781, 2122, 17859, 7064, 963.5,
      408.5, 5623, 4923.5, 2745, 1738.5, 7057, 3001.5, 895.5, 12155.5, 1290.5,
      943, 11312.5, 2700, 3561, 1711, 1006, 1180.5, 9007, 4438, 3606)

y = c(10741.5, 3848.5, 2655.5, 5190.5, 1554.5, 6025, 3861.5, 2533.5, 1858, 1636, 
      4341.5, 8185, 1753.5, 6793, 9056, 4724.5, 2133.5, 16576.5, 7106.5, 986.5, 
      481, 5603, 5046, 2799.5, 1721.5, 7547.5, 3009.5, 986.5, 11712, 1416, 
      998.5, 11698.5, 2829, 3776.5, 1899, 1042, 1616, 9380.5, 4332.5, 3308.5)

# Passing Bablok Regression

reg = mcreg(x,y, method.reg = "PaBa", method.ci = "analytical")
summary(reg)

reg = mcreg(x,y, method.reg = "WDeming", method.ci = "jackknife")
summary(reg)
