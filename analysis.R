library(readxl)
library(RQuantLib)
sheets_results_mean <- c()
sheets_results_std <- c()

for (sheet_no in 1:12) {
  tryCatch({
  data <- read_excel("isx2010C.xls",sheet=sheet_no)
  names(data)[1] <- 'daystomaturity'
  names(data)[length(names(data)) - 2] <- 'S'
  names(data)[length(names(data)) - 1] <- 'rate'
  data$rate<-data$rate/100
  for (t in 1:(length(data)-1)){
    for (g in 1:dim(data)[1]){
      data[g,t]=if (is.na(as.numeric(data[g,t]))){data[g,t]}else if(as.numeric(data[g,t])<=1000){data[g,t]}else{data[g,t]/1000}
    }
  }
  
  strikescolumns= 2 : (length(names(data)) - 3)
  volas <- matrix(0,length(data$daystomaturity)-1,length(strikescolumns))
  deltas <- matrix(0,length(data$daystomaturity)-1,length(strikescolumns))
  vegas <- matrix(0,length(data$daystomaturity)-1,length(strikescolumns))
  gammas <- matrix(0,length(data$daystomaturity)-1,length(strikescolumns))
  
  option_value_changes <- matrix(0,length(data$daystomaturity)-2,length(strikescolumns))
  delta_hedge_value_changes <- matrix(0,length(data$daystomaturity)-2,length(strikescolumns))
  
  for (i in 1:length(strikescolumns)){
    
    previous_value <- 0
    previous_delta <- 0
    previous_underlying <- 0
    
    for (k in 1:(length(data$daystomaturity)-1)){
      value <- as.numeric(data[k,strikescolumns[i]])
      if (is.na(value)) {
        break
      }
      if (value > data$S[k]) {
        value <- previous_value # If the value makes no sense, assume that it did not change
        if (previous_value == 0) {
          break # If the first value makes no sense, exclude this strike from the dataset
        }
      }
      underlying <- data$S[k]
      strike <- as.numeric(colnames(data)[strikescolumns[i]])
      riskFreeRate <- data$rate[k]
      maturity <- data$daystomaturity[k]/252
      
      guess_a <- 0
      guess_b <- 2 # Maximum guess may need to be adjusted according to the dataset
      
      while (abs(guess_a - guess_b) > 0.0000001)  {
        result <- EuropeanOption("call",underlying,strike,0,riskFreeRate,maturity, (guess_a+guess_b) / 2)$value - value
        
        if (result < 0) {
          guess_a = (guess_a+guess_b) / 2
        } else {
          guess_b = (guess_a+guess_b) / 2
        }
      }
      
      volas[k,i] = guess_a
      results = EuropeanOption("call",underlying,strike,0,riskFreeRate,maturity, guess_a)
      deltas[k,i] = results$delta
      vegas[k,i] = results$vega
      gammas[k,i] = results$gamma
      
      
      if (k>1) {
        option_value_changes[k-1,i] = value - previous_value
        delta_hedge_value_changes[k-1,i] = previous_delta * (underlying - previous_underlying)
      }
      
      # Update the state
      previous_value <- value
      previous_underlying <- underlying
      if (!is.na(results$delta)) {
        previous_delta <- results$delta
      }
    }
  }
  
  daily_errors_squared <- (option_value_changes - delta_hedge_value_changes) ^2
  print(paste("Calculated sheet no ",sheet_no))
  strike_performances <- colSums(daily_errors_squared) / (length(data$daystomaturity)-2)
  sheets_results_mean <- append(sheets_results_mean, mean(strike_performances[strike_performances != 0]))
  sheets_results_std <- append(sheets_results_std, sd(strike_performances[strike_performances != 0]))

  kokodatat <- vector()
  for (m in 2:(length(data)-3)){
    if(length(which(!is.na(data[,m])))==dim(data)[1]){kokodatat<-append(kokodatat,m)}
  }

  # Do this 10 times for each sheet
  for (s in 1:10){
    indeksit<-sample(kokodatat,3,replace=FALSE)
    eka<-sort(indeksit)[1]
    toka<-sort(indeksit)[2]
    kolmas<-sort(indeksit)[3]
    Portfoliodelta = deltas[1,eka-1]
    Portfoliogamma = gammas[1,eka-1]
    Portfoliovega = vegas[1,eka-1]
    optio_2_positio=0
    optio_3_positio=0
    underlying_positio=0
    counter=0

    dailyerrors=c(rep(0,(length(data$daystomaturity)-2)))
    for (n in 2:(length(data$daystomaturity)-1)){
      A=matrix(
        c(deltas[n-1,toka-1],deltas[n-1,kolmas-1],1,gammas[n-1,toka-1],gammas[n-1,kolmas-1],0,vegas[n-1,toka-1],vegas[n-1,kolmas-1],0),
        nrow=3,
        ncol=3,
        byrow=TRUE
      )

      B=c(-Portfoliodelta,-Portfoliogamma,-Portfoliovega)

      if (any(is.na(A))||any(is.na(B))){
        optio_2_positio=optio_2_positio
        optio_3_positio=optio_3_positio
        underlying_positio=underlying_positio
      } else {
      optio_2_positio= solve(A,B)[1]
      optio_3_positio= solve(A,B)[2]
      underlying_positio= solve(A,B)[3]
      }
      Portfoliodelta = deltas[n,eka-1]
      Portfoliogamma = gammas[n,eka-1]
      Portfoliovega = vegas[n,eka-1]

      Portfoliovaluechange = as.numeric(data[n,eka])-as.numeric(data[n-1,eka])
      Replicatingvaluechange = (optio_2_positio*(as.numeric(data[n,toka])-as.numeric(data[n-1,toka]))
                              +optio_3_positio*(as.numeric(data[n,kolmas])-as.numeric(data[n-1,kolmas]))
                              +underlying_positio*(as.numeric(data[n,(length(data)-2)])-as.numeric(data[n-1,(length(data)-2)])))
      dailyerrors[n-1]=if(optio_2_positio==0&&optio_3_positio==0&&underlying_positio==0){0}else{Portfoliovaluechange+Replicatingvaluechange}
      counter= if(optio_2_positio==0&&optio_3_positio==0&&underlying_positio==0){counter+1}else{counter}
    }
    dailyerrors
    performance=sum(dailyerrors^2)/(length(data$daystomaturity)-2-counter)
    print(performance)
    print(indeksit)
  }

},error=function(e){})
}