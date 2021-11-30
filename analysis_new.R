library(readxl)
library(RQuantLib)
library(plotly)

# This function re-formats the data for processing:
# - Rename the columns
# - Scale the risk-free rate
# - Scale data points that are over 1000
fix_data <- function(data) {
  no_columns <- length(data)
  names(data)[1] <- 'daystomaturity'
  names(data)[no_columns - 2] <- 'S'
  names(data)[no_columns - 1] <- 'rate'
  names(data)[no_columns] <- 'date'
  data$rate <- data$rate / 100
  for (t in 1:(length(data) - 1)) {
    for (g in 1:dim(data)[1]) {
      datapoint <- as.numeric(data[g,t])
      if (is.na(datapoint)) {
        data[g, t] <- NA
      } else {
        data[g, t] <- if (datapoint <= 1000) datapoint else datapoint / 1000
      }
    }
  }
  return (data)
}

# This function removes strikes w/o complete time series
# and selects the last 80 rows in order to retain comparability
strip_data <- function(data) {
  #data[, names(data)[!is.na(data[1,])]]
  tail(data[, c("daystomaturity", "340", "360", "380", "400", "420", "440", "460", "480", "500", "520", "S", "rate", "date")], 80)
}

# This function calculates the implied volatility
# and greeks for each data point
calculate_greeks <- function(data) {
  volas <- matrix(NA, length(data$daystomaturity), length(data) - 4)
  deltas <- volas
  gammas <- volas
  vegas <- volas
  for (i in 2:(length(data) - 3)) {
    strike <- as.numeric(names(data)[i])
    for (j in 1:length(data$daystomaturity)) {
      value <- as.numeric(data[j, i])
      underlying <- as.numeric(data[j, "S"])
      dividendYield <- 0
      riskFreeRate <- as.numeric(data[j, "rate"])
      maturity <- as.numeric(data[j, "daystomaturity"]) / 252
      vola <- NA
      tryCatch({
        vola <- EuropeanOptionImpliedVolatility("call", value,
                                                  underlying, strike, dividendYield, riskFreeRate, maturity, 0.01)[1]
        if (vola < 0.5) { # Sanity check: volatility cannot exceed 50 per cent
          volas[j, i - 1] <- vola
          result <- EuropeanOption("call", underlying, strike, dividendYield, riskFreeRate, maturity, vola)
          deltas[j, i - 1] <- result$delta
          gammas[j, i - 1] <- result$gamma
          vegas[j, i - 1] <- result$vega
        }
      }, error = function(e){} )
    }
  }
  ret <- list("volas" = volas, "deltas" = deltas, "gammas" = gammas, "vegas" = vegas)
}

# This function calculates the performance for the following strategy:
# - Long 1 call
# - Short delta units of underlying
single_delta_hedge <- function(data, greeks, strike_no, hedge_freq) {
  total_days <- length(data$daystomaturity)
  securities_value <- vector("double", total_days) # Value of the portfolio's securities on day N
  position_cash <- vector("double", total_days) # Cash position on day N
  tracking_error <- vector("double", total_days - 1) # Tracking error E on day N+1
  
  # Set the initial position
  position_call <- 1
  position_underlying <- -(greeks$deltas[1, strike_no])
  position_underlying <- if(is.na(position_underlying)) 0 else position_underlying
  days_after_rehedge <- if(is.na(position_underlying)) Inf else 1
  securities_value[1] <- position_call * data[1, strike_no + 1] + position_underlying * data$S[1]
  position_cash[1] <- -(securities_value[[1]]) # Purchase the initial portfolio
  
  for (day in 2:(total_days - 1)) {
    # The previously decided positions determine the change in portfolio's value. Cash stays the same
    securities_value[day] <- position_call * data[day, strike_no + 1] + position_underlying * data$S[day]
    position_cash[day] <- position_cash[day - 1]
    
    tracking_error[day - 1] <- position_call * (data[day, strike_no + 1] - data[day - 1, strike_no + 1]) + position_underlying * (data$S[day] - data$S[day - 1])
    
    # The positions might change according to the strategy. The trading affects cash
    if (!is.na(greeks$deltas[day, strike_no]) && days_after_rehedge >= hedge_freq) {
      position_cash[day] <- position_cash[day - 1] + (position_underlying + greeks$deltas[day, strike_no]) * data$S[day]
      position_underlying <- -(greeks$deltas[day, strike_no])
      days_after_rehedge <- 1
    } else {
      days_after_rehedge <- days_after_rehedge + 1
    }
  }
  
  # The positions are sold
  position_cash[total_days] <- position_cash[total_days - 1] + position_call * data[total_days, strike_no + 1] + position_underlying * data$S[total_days]
  tracking_error[total_days - 1] <- position_call * (data[total_days, strike_no + 1] - data[total_days - 1, strike_no + 1]) + position_underlying * (data$S[total_days] - data$S[total_days - 1])
  
  portfolio_value <- unlist(securities_value) + unlist(position_cash) # On each day, the portfolio's total value is sum of securities + cash
  mean_error_squared <- mean(unlist(tracking_error) ^ 2)
  ret <- list("portfolio_value" = portfolio_value, "mean_error_squared" = mean_error_squared)
  return (ret)
}

# First reporting item:
# - Rehedging frequency 1 day
# - Sheet 1
# - Delta hedging
sheet_no <- 1
data <- read_excel("isx2010C.xls", sheet=sheet_no)
data <- fix_data(data)
data <- strip_data(data)
greeks <- calculate_greeks(data)
print("Strike; Mean error squared; Total change")
total_days <- length(data$daystomaturity)
portfolio_values <- matrix(nrow = total_days, ncol = length(data) - 4)
for (i in 1:(length(data) - 4)) {
  ret <- single_delta_hedge(data, greeks, i, 1)
  print(paste(names(data)[i + 1], ret$mean_error_squared, ret$portfolio_value[total_days], sep = "; "))
  portfolio_values[1:total_days, i] <- ret$portfolio_value
}
plot_ly(z = ~portfolio_values, type = "surface") %>% layout(title="Rehedging frequency 1 day")

# Second reporting item:
# - Rehedging frequency 1 day
# - Average over all sheets
# - Delta hedging
portfolio_values <- matrix(nrow = total_days, ncol = length(data) - 4, data = 0)
for (sheet_no in 1:12) {
  data <- read_excel("isx2010C.xls", sheet=sheet_no)
  data <- fix_data(data)
  data <- strip_data(data)
  greeks <- calculate_greeks(data)
  total_days <- length(data$daystomaturity)
  for (i in 1:(length(data) - 4)) {
    ret <- single_delta_hedge(data, greeks, i, 1)
    portfolio_values[1:total_days, i] <- portfolio_values[1:total_days, i] + ret$portfolio_value
  }
}
portfolio_values <- portfolio_values / 12
plot_ly(z = ~portfolio_values, type = "surface") %>% layout(title="Rehedging frequency 1 day")
