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
        if (vola < 0.5) { # Sanity check: volatility cannot exceed 0.5
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
  mean_error_squared <- mean(unlist(tracking_error) ^ 2) * hedge_freq
  ret <- list("portfolio_value" = portfolio_value, "mean_error_squared" = mean_error_squared)
  return (ret)
}

# This function solves the following system of equations:
# position_call_2 * call_2_delta + position_underlying * 1 = -call_1_delta
# position_call_2 * call_2_gamma + position_underlying * 0 = -call_1_gamma
solve_delta_gamma_hedge <- function(call_1_delta, call_1_gamma, call_2_delta, call_2_gamma) {
  right_hand_side <- matrix(c(-call_1_delta, -call_1_gamma), nrow = 2)
  left_hand_side <- matrix(c(call_2_delta, call_2_gamma, 1, 0), nrow = 2)
  solve(left_hand_side, right_hand_side)
}

# This function calculates the performance for the following strategy:
# - Long 1 call
# - A position in underlying and an another option s.t. the portfolio is delta-gamma neutral
single_delta_gamma_hedge <- function(data, greeks, hedge_freq, strike_no, hedge_strike_no) {
  total_days <- length(data$daystomaturity)
  securities_value <- vector("double", total_days) # Value of the portfolio's securities on day N
  position_cash <- vector("double", total_days) # Cash position on day N
  tracking_error <- vector("double", total_days - 1) # Tracking error E on day N+1
  
  # Set the initial position. Fall back to delta hedge if the system is singular
  position_call_1 <- 1
  position_call_2 <- 0
  position_underlying <- -(greeks$deltas[1, strike_no])
  position_underlying <- if(is.na(position_underlying)) 0 else position_underlying
  tryCatch({
    res <- solve_delta_gamma_hedge(greeks$deltas[1, strike_no],
                                   greeks$gammas[1, strike_no],
                                   greeks$deltas[1, hedge_strike_no],
                                   greeks$gammas[1, hedge_strike_no])
    if (is.na(res[1]) || is.na(res[2])) {
      stop()
    }
    position_call_2 <- res[1]
    position_underlying <- res[2]
  }, error = function(e){})
  
  days_after_rehedge <- if(is.na(position_underlying)) Inf else 1
  securities_value[1] <- position_call_1 * data[1, strike_no + 1] + position_call_2 * data[1, hedge_strike_no + 1] + position_underlying * data$S[1]
  position_cash[1] <- -(securities_value[[1]]) # Purchase the initial portfolio
  
  for (day in 2:(total_days - 1)) {
    # The previously decided positions determine the change in portfolio's value. Cash stays the same
    securities_value[day] <- position_call_1 * data[day, strike_no + 1] + position_call_2 * data[day, hedge_strike_no + 1] + position_underlying * data$S[day]
    position_cash[day] <- position_cash[day - 1]
    
    tracking_error[day - 1] <- position_call_1 * (data[day, strike_no + 1] - data[day - 1, strike_no + 1]) + position_call_2 * (data[day, hedge_strike_no + 1] - data[day - 1, hedge_strike_no + 1]) + position_underlying * (data$S[day] - data$S[day - 1])
    
    # The positions might change according to the strategy. The trading affects cash
    if (!is.na(greeks$deltas[day, strike_no]) && !is.na(greeks$gammas[day, strike_no]) && days_after_rehedge >= hedge_freq) {
      tryCatch({
        res <- solve_delta_gamma_hedge(greeks$deltas[day, strike_no],
                                       greeks$gammas[day, strike_no],
                                       greeks$deltas[day, hedge_strike_no],
                                       greeks$gammas[day, hedge_strike_no])
        if (is.na(res[1]) || is.na(res[2])) {
          days_after_rehedge <- days_after_rehedge + 1
          stop()
        }
        position_cash[day] <- position_cash[[day - 1]] + (position_underlying - res[2]) * data$S[day] + (position_call_2 - res[1]) * data[day, hedge_strike_no + 1]
        position_call_2 <- res[1]
        position_underlying <- res[2]
        days_after_rehedge <- 1
      }, error = function(e) {
        days_after_rehedge <- days_after_rehedge + 1
      })
    } else {
      days_after_rehedge <- days_after_rehedge + 1
    }
    #print(paste(securities_value[[day]] + position_cash[[day]], position_underlying, position_call_2, data$daystomaturity[day]))
  }
  
  # The positions are sold
  position_cash[total_days] <- position_cash[total_days - 1] + position_call_1 * data[total_days, strike_no + 1] + position_call_2 * data[total_days, hedge_strike_no + 1] + position_underlying * data$S[total_days]
  tracking_error[total_days - 1] <- (position_call_1 * (data[total_days, strike_no + 1] - data[total_days - 1, strike_no + 1])
                                     + position_call_2 * (data[total_days, hedge_strike_no + 1] - data[total_days - 1, hedge_strike_no + 1])
                                     + position_underlying * (data$S[total_days] - data$S[total_days - 1]))
  
  portfolio_value <- unlist(securities_value) + unlist(position_cash) # On each day, the portfolio's total value is sum of securities + cash
  mean_error_squared <- mean(unlist(tracking_error) ^ 2)* hedge_freq
  ret <- list("portfolio_value" = portfolio_value, "mean_error_squared" = mean_error_squared)
  return (ret)
}

# This function calculates the performance for the following strategy:
# - Long 1 call with strike strike_1_no
# - Short 1 call with strike strike_2_no
# - A position in underlying s.t. the portfolio is delta neutral
bullspread_delta_hedge <- function(data, greeks, hedge_freq, strike_1_no, strike_2_no) {
  total_days <- length(data$daystomaturity)
  securities_value <- vector("double", total_days) # Value of the portfolio's securities on day N
  position_cash <- vector("double", total_days) # Cash position on day N
  tracking_error <- vector("double", total_days - 1) # Tracking error E on day N+1
  
  # Set the initial position
  position_call_1 <- 1
  position_call_2 <- -1
  position_underlying <- -(greeks$deltas[1, strike_1_no] - greeks$deltas[1, strike_2_no])
  days_after_rehedge <- if(is.na(position_underlying)) Inf else 1
  position_underlying <- if(is.na(position_underlying)) 0 else position_underlying
  securities_value[1] <- position_call_1 * data[1, strike_1_no + 1] + position_call_2 * data[1, strike_2_no + 1] + position_underlying * data$S[1]
  position_cash[1] <- -(securities_value[[1]]) # Purchase the initial portfolio
  
  for (day in 2:(total_days - 1)) {
    # The previously decided positions determine the change in portfolio's value. Cash stays the same
    securities_value[day] <- position_call_1 * data[day, strike_1_no + 1] + position_call_2 * data[day, strike_2_no + 1] + position_underlying * data$S[day]
    position_cash[day] <- position_cash[day - 1]
    
    tracking_error[day - 1] <- (position_call_1 * (data[day, strike_1_no + 1] - data[day - 1, strike_1_no + 1])
                                + position_call_2 * (data[day, strike_2_no + 1] - data[day - 1, strike_2_no + 1])
                                + position_underlying * (data$S[day] - data$S[day - 1]))
    
    # The positions might change according to the strategy. The trading affects cash
    if (!is.na(greeks$deltas[day, strike_1_no]) && !is.na(greeks$deltas[day, strike_2_no]) && days_after_rehedge >= hedge_freq) {
      position_cash[day] <- position_cash[day - 1] + (position_underlying + greeks$deltas[day, strike_1_no] - greeks$deltas[day, strike_2_no]) * data$S[day]
      position_underlying <- -(greeks$deltas[day, strike_1_no] - greeks$deltas[day, strike_2_no])
      days_after_rehedge <- 1
    } else {
      days_after_rehedge <- days_after_rehedge + 1
    }
  }
  
  # The positions are sold
  position_cash[total_days] <- position_cash[total_days - 1] + position_call_1 * data[total_days, strike_1_no + 1] + position_call_2 * data[total_days, strike_2_no + 1] + position_underlying * data$S[total_days]
  tracking_error[total_days - 1] <- position_call_1 * (data[total_days, strike_1_no + 1] - data[total_days - 1, strike_1_no + 1]) + position_call_2 * (data[total_days, strike_2_no + 1] - data[total_days - 1, strike_2_no + 1]) + position_underlying * (data$S[total_days] - data$S[total_days - 1])
  
  portfolio_value <- unlist(securities_value) + unlist(position_cash) # On each day, the portfolio's total value is sum of securities + cash
  mean_error_squared <- mean(unlist(tracking_error) ^ 2)* hedge_freq
  ret <- list("portfolio_value" = portfolio_value, "mean_error_squared" = mean_error_squared)
  return (ret)
}

# This function calculates the performance for the following strategy:
# - Long 1 call with strike strike_1_no
# - Short 1 call with strike strike_2_no
# - A position in underlying and an another option s.t. the portfolio is delta-gamma neutral
bullspread_delta_gamma_hedge <- function(data, greeks, hedge_freq, strike_1_no, strike_2_no, hedge_strike_no) {
  total_days <- length(data$daystomaturity)
  securities_value <- vector("double", total_days) # Value of the portfolio's securities on day N
  position_cash <- vector("double", total_days) # Cash position on day N
  tracking_error <- vector("double", total_days - 1) # Tracking error E on day N+1
  
  # Set the initial position. Fall back to delta hedge if the system is singular
  position_call_1 <- 1
  position_call_2 <- -1
  position_hedge_call <- 0
  position_underlying <- -(greeks$deltas[1, strike_1_no] - greeks$deltas[1, strike_2_no])
  position_underlying <- if(is.na(position_underlying)) 0 else position_underlying
  tryCatch({
    res <- solve_delta_gamma_hedge(greeks$deltas[1, strike_1_no] - greeks$deltas[1, strike_2_no],
                                   greeks$gammas[1, strike_1_no] - greeks$deltas[1, strike_2_no],
                                   greeks$deltas[1, hedge_strike_no],
                                   greeks$gammas[1, hedge_strike_no])
    if (is.na(res[1]) || is.na(res[2])) {
      stop()
    }
    position_hedge_call <- res[1]
    position_underlying <- res[2]
  }, error = function(e){})
  
  days_after_rehedge <- if(is.na(position_underlying)) Inf else 1
  securities_value[1] <- position_call_1 * data[1, strike_1_no + 1] + position_call_2 * data[1, strike_2_no + 1] + position_hedge_call * data[1, hedge_strike_no + 1] + position_underlying * data$S[1]
  position_cash[1] <- -(securities_value[[1]]) # Purchase the initial portfolio
  
  for (day in 2:(total_days - 1)) {
    # The previously decided positions determine the change in portfolio's value. Cash stays the same
    securities_value[day] <- position_call_1 * data[day, strike_1_no + 1] + position_call_2 * data[day, strike_2_no + 1] + position_hedge_call * data[day, hedge_strike_no + 1] + position_underlying * data$S[day]
    position_cash[day] <- position_cash[day - 1]
    
    tracking_error[day - 1] <- position_call_1 * (data[day, strike_1_no + 1] - data[day - 1, strike_1_no + 1]) + position_call_2 * (data[day, strike_2_no + 1] - data[day - 1, strike_2_no + 1]) + position_hedge_call * (data[day, hedge_strike_no + 1] - data[day - 1, hedge_strike_no + 1]) + position_underlying * (data$S[day] - data$S[day - 1])
    
    # The positions might change according to the strategy. The trading affects cash
    if (!is.na(greeks$deltas[day, strike_1_no]) && !is.na(greeks$deltas[day, strike_2_no]) && !is.na(greeks$gammas[day, strike_1_no]) && !is.na(greeks$gammas[day, strike_2_no]) && days_after_rehedge >= hedge_freq) {
      tryCatch({
        res <- solve_delta_gamma_hedge(greeks$deltas[day, strike_1_no] - greeks$deltas[day, strike_2_no],
                                       greeks$gammas[day, strike_1_no] - greeks$deltas[day, strike_2_no],
                                       greeks$deltas[day, hedge_strike_no],
                                       greeks$gammas[day, hedge_strike_no])
        if (is.na(res[1]) || is.na(res[2])) {
          days_after_rehedge <- days_after_rehedge + 1
          stop()
        }
        position_cash[day] <- position_cash[[day - 1]] + (position_underlying - res[2]) * data$S[day] + (position_hedge_call - res[1]) * data[day, hedge_strike_no + 1]
        position_hedge_call <- res[1]
        position_underlying <- res[2]
        days_after_rehedge <- 1
      }, error = function(e) {
        days_after_rehedge <- days_after_rehedge + 1
      })
    } else {
      days_after_rehedge <- days_after_rehedge + 1
    }
    #print(paste(securities_value[[day]] + position_cash[[day]], position_underlying, position_hedge_call, data$daystomaturity[day]))
  }
  
  # The positions are sold
  position_cash[total_days] <- position_cash[total_days - 1] + position_call_1 * data[total_days, strike_1_no + 1] + position_call_2 * data[total_days, strike_2_no + 1] + position_hedge_call * data[total_days, hedge_strike_no + 1] + position_underlying * data$S[total_days]
  tracking_error[total_days - 1] <- (position_call_1 * (data[total_days, strike_1_no + 1] - data[total_days - 1, strike_1_no + 1])
                                     + position_call_2 * (data[total_days, strike_2_no + 1] - data[total_days - 1, strike_2_no + 1])
                                     + position_hedge_call * (data[total_days, hedge_strike_no + 1] - data[total_days - 1, hedge_strike_no + 1])
                                     + position_underlying * (data$S[total_days] - data$S[total_days - 1]))
  
  portfolio_value <- unlist(securities_value) + unlist(position_cash) # On each day, the portfolio's total value is sum of securities + cash
  mean_error_squared <- mean(unlist(tracking_error) ^ 2)* hedge_freq
  ret <- list("portfolio_value" = portfolio_value, "mean_error_squared" = mean_error_squared)
  return (ret)
}


# First reporting item:
# - Rehedging frequency 1 day
# - Sheet 1
# - Delta hedging
first_reporting_item <- function () {
  sheet_no <- 1
  data <- read_excel("isx2010C.xls", sheet=sheet_no)
  data <- fix_data(data)
  data <- strip_data(data)
  greeks <- calculate_greeks(data)
  print("Strike; Mean error squared; Total change")
  total_days <- length(data$daystomaturity)
  portfolio_values <- matrix(nrow = total_days, ncol = length(data) - 4)
  for (i in 1:(length(data) - 4)) {
    ret <- single_delta_hedge(data, greeks, i, 7)
    print(paste(names(data)[i + 1], ret$mean_error_squared, ret$portfolio_value[total_days], sep = "; "))
    portfolio_values[1:total_days, i] <- ret$portfolio_value
  }
  plot_ly(z = ~portfolio_values, type = "surface") %>% layout(title="Rehedging frequency 1 day")
}
  
# Second reporting item:
# - Rehedging frequency 1 day
# - Average over all sheets
# - Delta hedging
second_reporting_item <- function () {
  total_days <- 80
  portfolio_values <- matrix(nrow = total_days, ncol = 10, data = 0)
  portfolio_errors <- c(rep(0,10))
  for (sheet_no in 1:12) {
    data <- read_excel("isx2010C.xls", sheet=sheet_no)
    data <- fix_data(data)
    data <- strip_data(data)
    greeks <- calculate_greeks(data)
    total_days <- length(data$daystomaturity)
    for (i in 1:(length(data) - 4)) {
      ret <- single_delta_hedge(data, greeks, i, 7)
      portfolio_values[1:total_days, i] <- portfolio_values[1:total_days, i] + ret$portfolio_value
      portfolio_errors[i] <- portfolio_errors[i] + ret$mean_error_squared
    }
  }
  portfolio_values <- portfolio_values / 12
  portfolio_errors <- portfolio_errors / 12
  plot_ly(z = ~portfolio_values, type = "surface") %>% layout(title="Rehedging frequency 1 day")
}

# Third reporting item:
# - Rehedging frequency 1 day
# - Sheet 1
# - Delta-gamma hedging using the adjacent strike
third_reporting_item <- function () {
  sheet_no <- 1
  data <- read_excel("isx2010C.xls", sheet=sheet_no)
  data <- fix_data(data)
  data <- strip_data(data)
  greeks <- calculate_greeks(data)
  print("Strike; Mean error squared; Total change")
  total_days <- length(data$daystomaturity)
  portfolio_values <- matrix(nrow = total_days, ncol = length(data) - 4)
  for (i in 1:(length(data) - 5)) {
    ret <- single_delta_gamma_hedge(data, greeks, 1, i, i + 1)
    print(paste(names(data)[i + 1], ret$mean_error_squared, ret$portfolio_value[total_days], sep = "; "))
    portfolio_values[1:total_days, i] <- ret$portfolio_value
  }
  plot_ly(z = ~portfolio_values, type = "surface") %>% layout(title="Rehedging frequency 1 day")
}

# Fourth reporting item:
# - Rehedging frequency 7 days
# - Sheet average
# - Delta-gamma hedging using the adjacent strike
fourth_reporting_item <- function () {
  total_days <- 80
  portfolio_values <- matrix(nrow = total_days, ncol = 10, data = 0)
  for (sheet_no in 1:12) {
    data <- read_excel("isx2010C.xls", sheet=sheet_no)
    data <- fix_data(data)
    data <- strip_data(data)
    greeks <- calculate_greeks(data)
    total_days <- length(data$daystomaturity)
    for (i in 1:10) {
      ret <- single_delta_gamma_hedge(data, greeks, 7, i, i + 1)
      portfolio_values[1:total_days, i] <- portfolio_values[1:total_days, i] + ret$portfolio_value
      #print(sheet_no)
      #print(portfolio_values)
    }
  }
  portfolio_values <- portfolio_values / 12
  plot_ly(z = ~portfolio_values, type = "surface") %>% layout(title="Rehedging frequency 7 day")
}

# Fifth reporting item:
# - Rehedging frequency 1 day
# - Sheet 1
# - Delta-gamma hedging using all combinations, where the hedge strike > base strike
fifth_reporting_item <- function () {
     total_days <- 80
     portfolio_values1 <- matrix(nrow = total_days, ncol = 10, data = 0)
     portfolio_errors1 <- matrix(nrow = total_days, ncol = 10, data = 0)
     for (sheet_no in 1:12){
       portfolio_values <- matrix(nrow = total_days, ncol = 10, data = 0)
       portfolio_errors <- matrix(nrow = total_days, ncol = 10, data = 0)
       data <- read_excel("isx2010C.xls", sheet=sheet_no)
       data <- fix_data(data)
       data <- strip_data(data)
       greeks <- calculate_greeks(data)
       total_days <- length(data$daystomaturity)
       count <- 0
       for (i in 1:10) {
           for (j in 1:10) {
               if (j > i) {
                   count <- count + 1
                   ret <- single_delta_gamma_hedge(data, greeks, 1, i, j)
                   portfolio_values[1:total_days, i] <- portfolio_values[1:total_days, i] + ret$portfolio_value
                   portfolio_errors[1:total_days, i] <- portfolio_errors[1:total_days, i] + ret$mean_error_squared
                 }
             }
         }
       portfolio_values[1:total_days, i] <- portfolio_values[1:total_days, i] / count
       portfolio_errors[1:total_days, i] <- portfolio_errors[1:total_days, i] / count
       portfolio_values1 <- portfolio_values1 + portfolio_values
       portfolio_errors1 <- portfolio_errors1 + portfolio_errors
       }
     portfolio_values1 <- portfolio_values1 / 12
     portfolio_errors1 <- portfolio_errors1 / 12
     plot_ly(z = ~portfolio_values1, type = "surface") %>% layout(title="Rehedging frequency 7 day")
}

# Sixth reporting item:
# - Bull/bear spread portfolio (all possible call spread combinations)
# - Rehedging frequency 1 day
# - Sheet 1
# - Delta hedging
sixth_reporting_item <- function () {
  sheet_no <- 1
  data <- read_excel("isx2010C.xls", sheet=sheet_no)
  data <- fix_data(data)
  data <- strip_data(data)
  greeks <- calculate_greeks(data)
  print("Strike; Mean error squared; Total change")
  total_days <- length(data$daystomaturity)
  for (i in 1:(length(data) - 4)) {
    for (j in 1:(length(data) - 4)) {
      ret <- bullspread_delta_hedge(data, greeks, 1, i, j)
      print(paste(names(data)[i + 1], names(data)[j + 1], ret$mean_error_squared, ret$portfolio_value[total_days], sep = "; "))
    }
  }
}

# Seventh reporting item:
# - Bullspread portfolio, i:th strike long and the adjacent strike i+1 short
# - Rehedging frequency 1 day
# - Average over all sheets
# - Delta hedging
seventh_reporting_item <- function () {
  total_days <- 80
  portfolio_values <- matrix(nrow = total_days, ncol = 10, data = 0)
  portfolio_errors <- c(rep(0,10))
  for (sheet_no in 1:12) {
    data <- read_excel("isx2010C.xls", sheet=sheet_no)
    data <- fix_data(data)
    data <- strip_data(data)
    greeks <- calculate_greeks(data)
    total_days <- length(data$daystomaturity)
    for (i in 1:(length(data) - 5)) {
      ret <- bullspread_delta_hedge(data, greeks, 1, i, i + 1)
      portfolio_values[1:total_days, i] <- portfolio_values[1:total_days, i] + ret$portfolio_value
      portfolio_errors[i] <- portfolio_errors[i] + ret$mean_error_squared
    }
  }
  portfolio_values <- portfolio_values / 12
  portfolio_errors <- portfolio_errors / 12
  plot_ly(z = ~portfolio_values, type = "surface") %>% layout(title="Rehedging frequency 1 day")
}

# Eigth reporting item:
# - Bullspread portfolio, i:th strike long and the adjacent strike i+1 short
# - Rehedging frequency 1-7 days
# - Average over all sheets
# - Delta hedging
eigth_reporting_item <- function () {
  freq <- 7 #Change this row
  total_days <- 80
  portfolio_values <- matrix(nrow = total_days, ncol = 10, data = 0)
  portfolio_errors <- c(rep(0,10))
  for (sheet_no in 1:12) {
    data <- read_excel("isx2010C.xls", sheet=sheet_no)
    data <- fix_data(data)
    data <- strip_data(data)
    greeks <- calculate_greeks(data)
    total_days <- length(data$daystomaturity)
    for (i in 1:(length(data) - 5)) {
      ret <- bullspread_delta_hedge(data, greeks, freq, i, i + 1)
      portfolio_values[1:total_days, i] <- portfolio_values[1:total_days, i] + ret$portfolio_value
      portfolio_errors[i] <- portfolio_errors[i] + ret$mean_error_squared
    }
  }
  portfolio_values <- portfolio_values / 12
  portfolio_errors <- portfolio_errors / 12
  print(paste("Frequency: ", freq, ", mean error = ", mean(portfolio_errors), ", value = ", mean(portfolio_values)))
}
