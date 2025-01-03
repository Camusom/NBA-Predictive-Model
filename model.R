library(readxl)
library(car)
balldata <- "C:\\Users\\Massimo Camuso\\Desktop\\NBA-NCAAB Model\\Masterwok.xlsx"
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

Hometeam_topg = hometeam_data$'tov',
Awayteam_topg = awayteam_data$'tov',

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

Hometeam_optopg = hometeam_data$'optov',
Awayteam_optopg = awayteam_data$'optov',

Hometeam_sos = hometeam_data$'home_team_sos',
Awayteam_sos = awayteam_data$'away_team_sos',

Hometeam_spread = hometeam_data[[paste(hometeam_abbr, "Line")]],
Awayteam_spread = awayteam_data[[paste(awayteam_abbr, "Line")]],

Hometeam_mov = hometeam_data$home_spread_movement,
Awayteam_mov = awayteam_data$away_spread_movement
)

combined_data <- na.omit(combined_data)
model_spread_home <- lm(Hometeam_spread ~ Hometeam_2pt + Awayteam_2pt + Hometeam_3pt + Awayteam_3pt
                        + Hometeam_ft + Awayteam_ft + Hometeam_or + Awayteam_or + Hometeam_dr
                        + Awayteam_dr + Hometeam_topg + Awayteam_topg + Hometeam_op2pt
                        + Awayteam_op2pt + Hometeam_op3pt + Awayteam_op3pt + Hometeam_opft
                        + Awayteam_opft + Hometeam_optopg + Awayteam_optopg + Hometeam_mov + Awayteam_spread, data=combined_data)
summary(model_spread_home)

model_matrix <- model.matrix(Hometeam_spread ~ ., data = combined_data)  # Replace `y` with your response variable
qr(model_matrix)$rank
ncol(model_matrix)  # Compare the rank to the number of columns (predictors)

newdata <- combined_data[1, ]
predict(model_spread_home, newdata=newdata)
