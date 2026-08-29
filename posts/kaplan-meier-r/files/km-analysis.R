library(survival)

dat <- read.csv("km-data.csv")
fit <- survfit(Surv(time, status) ~ arm, data = dat)
print(summary(fit))
plot(fit, col = c("#277a43", "#52616b"), lwd = 2, xlab = "Time", ylab = "Event-free probability")
