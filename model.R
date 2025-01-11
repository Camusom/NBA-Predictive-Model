library(readxl)
library(car)
library(glmnet)
library(ggplot2)
balldata <- "C:\\Users\\Massimo Camuso\\Desktop\\NBA-NCAAB Model\\Masterfile.xlsx"
sheet_names <- excel_sheets(balldata)

sheets_data <- lapply(sheet_names, function(sheet) {
  read_excel(balldata, sheet = sheet)
})
names(sheets_data) <- sheet_names

list2env(sheets_data, envir = .GlobalEnv)

hometeam_abbr <- "LAL"
awayteam_abbr <- "OKC"

combined_data <- data.frame(
Hometeam_2pt = hometeam_data$'2ptpercent',
Awayteam_2pt = awayteam_data$'2ptpercent',

Hometeam_3pt = hometeam_data$'3ptpercent',
Awayteam_3pt = awayteam_data$'3ptpercent',

Hometeam_ft = hometeam_data$'ftpercent',
Awayteam_ft = awayteam_data$'ftpercent',

Hometeam_or = hometeam_data$'orbpercent',
Awayteam_or = awayteam_data$'orbpercent',

Hometeam_dr = hometeam_data$'drbpercent',
Awayteam_dr = awayteam_data$'drbpercent',

Hometeam_topercent = hometeam_data$'tov',
Awayteam_topercent = awayteam_data$'tov',

Hometeam_op2pt = hometeam_data$'op2ptpercent',
Awayteam_op2pt = awayteam_data$'op2ptpercent',

Hometeam_op3pt = hometeam_data$'op3ptpercent',
Awayteam_op3pt = awayteam_data$'op3ptpercent',

Hometeam_opft = hometeam_data$'opftpercent',
Awayteam_opft = awayteam_data$'opftpercent',

Hometeam_opor = hometeam_data$'oporpercent',
Awayteam_opor = awayteam_data$'oporpercent',

Hometeam_opdr = hometeam_data$'opdrpercent',
Awayteam_opdr = awayteam_data$'opdrpercent',


Hometeam_spread = hometeam_data[[paste(hometeam_abbr, "Line")]],
Awayteam_spread = awayteam_data[[paste(awayteam_abbr, "Line")]],

Hometeam_mov = hometeam_data$home_spread_movement,
Awayteam_mov = awayteam_data$away_spread_movement,

Awayteam_opstl = awayteam_data$'opstealpercent',

Hometeam_totalrebound = hometeam_data$totalrebound,
Awayteam_totalrebound = awayteam_data$totalrebound,

Hometeam_totalreboundpercent = hometeam_data$totalreboundpercent,
Awayteam_totalreboundpercent = awayteam_data$totalreboundpercent
)

combined_data <- na.omit(combined_data)

model_matrix <- model.matrix(model_spread_home, data = combined_data)  # Replace `y` with your response variable
qr(model_matrix)$rank
ncol(model_matrix)  # Compare the rank to the number of columns (predictors)

testmodel1 <- lm(combined_data$Hometeam_spread ~ combined_data$Hometeam_2pt)
#response variable vs predictor
ggplot(combined_data, aes(x = Hometeam_2pt, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)        # Linear regression
#residuals
ggplot(combined_data, aes(x = fitted(testmodel1), y = residuals(testmodel1))) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue") +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red")
#LINEAR include



testmodel3 <- lm(combined_data$Hometeam_spread ~ combined_data$Hometeam_3pt)
#response variable vs predictor
ggplot(combined_data, aes(x = Hometeam_3pt, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)        # Linear regression
#residuals
ggplot(combined_data, aes(x = fitted(testmodel3), y = residuals(testmodel3))) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue") +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red")
cor(combined_data$Hometeam_3pt, combined_data$Hometeam_spread, method = "pearson")
cor(combined_data$Hometeam_3pt, combined_data$Hometeam_spread, method = "spearman")
#weak linear - include




testmodel5 <- lm(combined_data$Hometeam_spread ~ combined_data$Hometeam_ft)
#response variable vs predictor
ggplot(combined_data, aes(x = Hometeam_ft, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)        # Linear regression
#residuals
ggplot(combined_data, aes(x = fitted(testmodel5), y = residuals(testmodel5))) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue") +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red")
cor(combined_data$Hometeam_ft, combined_data$Hometeam_spread, method = "pearson")
cor(combined_data$Hometeam_ft, combined_data$Hometeam_spread, method = "spearman")
#include





testmodel7 <- lm(combined_data$Hometeam_spread ~ combined_data$Hometeam_or)
#response variable vs predictor
ggplot(combined_data, aes(x = Hometeam_or, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)        # Linear regression
#residuals
ggplot(combined_data, aes(x = fitted(testmodel7), y = residuals(testmodel7))) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue") +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red")
cor(combined_data$Hometeam_or, combined_data$Hometeam_spread, method = "pearson")
cor(combined_data$Hometeam_or, combined_data$Hometeam_spread, method = "spearman")
#interactions


testmodel8 <- lm(combined_data$Hometeam_spread ~ combined_data$Awayteam_or)
#response variable vs predictor
ggplot(combined_data, aes(x = Awayteam_or, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)        # Linear regression
#residuals
ggplot(combined_data, aes(x = fitted(testmodel8), y = residuals(testmodel8))) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue") +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red")
cor(combined_data$Awayteam_or, combined_data$Hometeam_spread, method = "pearson")
cor(combined_data$Awayteam_or, combined_data$Hometeam_spread, method = "spearman")
#interactions

testmodel9 <- lm(combined_data$Hometeam_spread ~ combined_data$Hometeam_dr)
#response variable vs predictor
ggplot(combined_data, aes(x = Hometeam_dr, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)        # Linear regression
#residuals
ggplot(combined_data, aes(x = fitted(testmodel9), y = residuals(testmodel9))) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue") +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red")
cor(combined_data$Hometeam_dr, combined_data$Hometeam_spread, method = "pearson")
cor(combined_data$Hometeam_dr, combined_data$Hometeam_spread, method = "spearman")
#interactions

testmodel10 <- lm(combined_data$Hometeam_spread ~ combined_data$Awayteam_dr)
#response variable vs predictor
ggplot(combined_data, aes(x = Awayteam_dr, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)        # Linear regression
#residuals
ggplot(combined_data, aes(x = fitted(testmodel10), y = residuals(testmodel10))) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue") +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red")
cor(combined_data$Awayteam_dr, combined_data$Hometeam_spread, method = "pearson")
cor(combined_data$Awayteam_dr, combined_data$Hometeam_spread, method = "spearman")
#interactions

testmodel11 <- lm(combined_data$Hometeam_spread ~ combined_data$Hometeam_topg)
#response variable vs predictor
ggplot(combined_data, aes(x = Hometeam_topercent, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)        # Linear regression
#residuals
ggplot(combined_data, aes(x = fitted(testmodel11), y = residuals(testmodel11))) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue") +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red")
cor(combined_data$Hometeam_topg, combined_data$Hometeam_spread, method = "pearson")
cor(combined_data$Hometeam_topg, combined_data$Hometeam_spread, method = "spearman")
#check TOrate





testmodel13 <- lm(combined_data$Hometeam_spread ~ combined_data$Hometeam_op2pt)
#response variable vs predictor
ggplot(combined_data, aes(x = Awayteam_totalrebound, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)        # Linear regression
#residuals
ggplot(combined_data, aes(x = fitted(testmodel13), y = residuals(testmodel13))) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue") +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red")
cor(combined_data$Hometeam_op2pt, combined_data$Hometeam_spread, method = "pearson")
cor(combined_data$Hometeam_op2pt, combined_data$Hometeam_spread, method = "spearman")
#include




testmodel15 <- lm(combined_data$Hometeam_spread ~ combined_data$Hometeam_op3pt)
#response variable vs predictor
ggplot(combined_data, aes(x = Awayteam_3pt, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)        # Linear regression
#residuals
ggplot(combined_data, aes(x = fitted(testmodel15), y = residuals(testmodel15))) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue") +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red")
cor(combined_data$Hometeam_op3pt, combined_data$Hometeam_spread, method = "pearson")
cor(combined_data$Hometeam_op3pt, combined_data$Hometeam_spread, method = "spearman")
# f test




testmodel17 <- lm(combined_data$Hometeam_spread ~ combined_data$Hometeam_opft)
#response variable vs predictor
ggplot(combined_data, aes(x = Hometeam_opor, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)        # Linear regression
#residuals
ggplot(combined_data, aes(x = fitted(testmodel17), y = residuals(testmodel17))) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue") +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red")
cor(combined_data$Hometeam_opft, combined_data$Hometeam_spread, method = "pearson")
cor(combined_data$Hometeam_opft, combined_data$Hometeam_spread, method = "spearman")
#include

ggplot(combined_data, aes(x = Hometeam_opft*Awayteam_opft, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)




#include
ggplot(combined_data, aes(x = (Hometeam_opor-Awayteam_opor), y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)



ggplot(combined_data, aes(x = Awayteam_rebound_advantage, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)
ggplot(combined_data, aes(x = fitted(testmodelto), y = residuals(testmodelto))) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue") +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red")

Hometeam_rebound_advantage <- (combined_data$Hometeam_or + combined_data$Hometeam_dr) - (combined_data$Awayteam_or + combined_data$Awayteam_dr)
Awayteam_rebound_advantage <- (combined_data$Awayteam_or + combined_data$Awayteam_dr) - (combined_data$Hometeam_or + combined_data$Hometeam_dr)









#TRANSFORMATION TESTS
#REBOUND TRANSFORMATIONS
rebounddiff <- combined_data$Hometeam_dr - combined_data$Awayteam_dr
rebounddiffmod <- lm(combined_data$Hometeam_spread ~ rebounddiff)
ggplot(combined_data, aes(x = rebounddiff, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)
ggplot(combined_data, aes(x = fitted(rebounddiffmod), y = residuals(rebounddiffmod))) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue") +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red")
cor(combined_data$Awayteam_optopg, combined_data$Hometeam_spread, method = "pearson")
cor(combined_data$Awayteam_optopg, combined_data$Hometeam_spread, method = "spearman")





AwayteamDRxAwayteamopOR <- combined_data$Awayteam_dr*combined_data$Awayteam_opor
reboundintmod <- lm(combined_data$Hometeam_spread ~ AwayteamDRxAwayteamopOR)
ggplot(combined_data, aes(x = AwayteamDRxAwayteamopOR, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)


#NET REBOUNDING
netboards <- (combined_data$Hometeam_or-combined_data$Awayteam_dr)+(combined_data$Hometeam_dr-combined_data$Awayteam_or)
ggplot(combined_data, aes(x = Hometeam_topercent, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)



#SHOOTING INTERACTIONS






shootingmodel <- lm(combined_data$Hometeam_spread ~ combined_data$Awayteam_3pt)
ggplot(combined_data, aes(x = Awayteam_opstl, y = Hometeam_spread)) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue", se = FALSE) +  # Smoothed curve
  geom_smooth(method = "lm", color = "red", se = FALSE)
ggplot(combined_data, aes(x = fitted(shootingmodel), y = residuals(shootingmodel))) +
  geom_point() +
  geom_smooth(method = "loess", color = "blue") +
  geom_hline(yintercept = 0, linetype = "dashed", color = "red")
cor(combined_data$Awayteam_2pt*combined_data$Awayteam_opstl, combined_data$Hometeam_spread, method = "pearson")
cor(combined_data$Awayteam_2pt*combined_data$Awayteam_opstl, combined_data$Hometeam_spread, method = "spearman")



#MODELLLL
# Hometeam and Awayteam rebound percentage


Drebounddiff <- combined_data$Hometeam_dr - combined_data$Awayteam_dr
Awayteam2ptint <- combined_data$Awayteam_2pt*combined_data$Awayteam_opstl
Awayteam3ptint <- combined_data$Awayteam_3pt*combined_data$Awayteam_opstl
model_homespread <- lm(Hometeam_spread ~ Hometeam_2pt + Hometeam_3pt 
                       + Hometeam_ft + Hometeam_topercent 
                       + Hometeam_opft + Drebounddiff + Awayteam2ptint
                       + Awayteam3ptint, data = combined_data)
summary(model_homespread)
AIC(model_homespread)

model_test <- lm(Hometeam_spread ~ Drebounddiff, data = combined_data)
anova(model_test, model_homespread)
summary(model_test)
AIC(model_test)

vif(model_homespread)


# Load the glmnet package


# Prepare the data
# Assuming your dataset is called 'combined_data' and the response variable is 'Hometeam_spread'
x <- model.matrix(Hometeam_spread ~ Hometeam_2pt + Hometeam_3pt 
                  + Hometeam_ft + Hometeam_topercent 
                  + Hometeam_opft + Drebounddiff + Awayteam2ptint
                  + Awayteam3ptint, data = combined_data)[, -1]  # Remove intercept
y <- combined_data$Hometeam_spread

# Fit Ridge Regression with cross-validation
set.seed(123)  # For reproducibility
ridge_model <- cv.glmnet(x, y, alpha = 0, standardize = TRUE)

# Get the best lambda (regularization parameter)
best_lambda <- ridge_model$lambda.min

# Display the best lambda value
cat("Best Lambda:", best_lambda, "\n")

# Fit Ridge Regression using the best lambda
final_ridge_model <- glmnet(x, y, alpha = 0, lambda = best_lambda, standardize = TRUE)

# View coefficients of the final model
ridge_coefficients <- coef(final_ridge_model)
print(ridge_coefficients)

# Predict on the training data (or new data)
predictions <- predict(final_ridge_model, s = best_lambda, newx = x)

# Assess performance (optional)
mse_ridge <- mean((y - predictions)^2)  # Mean Squared Error
cat("Mean Squared Error:", mse, "\n")



#LASSO MODEL
set.seed(123)  # For reproducibility
lasso_model <- cv.glmnet(x, y, alpha = 1, standardize = TRUE)

# Get the best lambda (regularization parameter)
best_lambda_lasso <- lasso_model$lambda.min

# Display the best lambda value
cat("Best Lambda:", best_lambda_lasso, "\n")

# Fit Ridge Regression using the best lambda
final_lasso_model <- glmnet(x, y, alpha = 1, lambda = best_lambda_lasso, standardize = TRUE)

# View coefficients of the final model
lasso_coefficients <- coef(final_lasso_model)
print(lasso_coefficients)

# Predict on the training data (or new data)
predictions_lasso <- predict(final_lasso_model, s = best_lambda, newx = x)

# Assess performance (optional)
mse_lasso <- mean((y - predictions_lasso)^2)  # Mean Squared Error
cat("Mean Squared Error:", mse_lasso, "\n")



lasso_coefficients
ridge_coefficients