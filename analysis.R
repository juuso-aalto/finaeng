library(readxl)
library(RQuantLib)
sheets_results_mean <- c()
sheets_results_std <- c()
sheets_absdiff_mean <- c()
sheets_absdiff_std <- c()
sheets_results2_mean <- c()
sheets_results2_std <- c()

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
    
    cashdesk <- matrix(0,length(data$daystomaturity)-1,length(strikescolumns))
    option_value_changes <- matrix(0,length(data$daystomaturity)-2,length(strikescolumns))
    delta_hedge_value_changes <- matrix(0,length(data$daystomaturity)-2,length(strikescolumns))
    finalcash <- c()
    no_hedge_cash <- c()
    maxloss <- c()
    
    for (i in 1:length(strikescolumns)){
      
      previous_value <- 0
      previous_delta <- 0
      previous_underlying <- 0
      
      newdelta <- 0 #for cashdesk calculation
      pastdelta <- 0 #for cashdesk calculation
      maxloss <- if(!is.na(as.numeric(data[1,i+1]))){append(maxloss,-as.numeric(data[1,i+1]))}else{maxloss}
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
        guess_b <- 10 # Maximum guess may need to be adjusted according to the dataset
        
        while (abs(guess_a - guess_b) > 0.0000001)  {
          result <- EuropeanOption("call",underlying,strike,0,riskFreeRate,maturity, (guess_a+guess_b) / 2)$value - value
          
          if (result < 0) {
            guess_a = (guess_a+guess_b) / 2
          } else {
            guess_b = (guess_a+guess_b) / 2
          }
        }
        
        volas[k,i] = if (guess_a>0.5){0}else{guess_a}
        results = EuropeanOption("call",underlying,strike,0,riskFreeRate,maturity, (if (guess_a>0.5){0}else{guess_a}))
        deltas[k,i] = results$delta
        vegas[k,i] = results$vega
        gammas[k,i] = results$gamma
        if (sheet_no==1){
          volas1 <- volas
          gammas1 <- gammas
          vegas1 <- vegas
          deltas1 <- deltas
          data1 <- data
        }
        if (sheet_no==2){
          volas2 <- volas
          gammas2 <- gammas
          vegas2 <- vegas
          deltas2 <- deltas
          data2 <- data
        }
        if (sheet_no==3){
          volas3 <- volas
          gammas3 <- gammas
          vegas3 <- vegas
          deltas3 <- deltas
          data3 <- data
        }
        if (sheet_no==4){
          volas4 <- volas
          gammas4 <- gammas
          vegas4 <- vegas
          deltas4 <- deltas
          data4 <- data
        }
        if (sheet_no==5){
          volas5 <- volas
          gammas5 <- gammas
          vegas5 <- vegas
          deltas5 <- deltas
          data5 <- data
        }
        if (sheet_no==6){
          volas6 <- volas
          gammas6 <- gammas
          vegas6 <- vegas
          deltas6 <- deltas
          data6 <- data
        }
        if (sheet_no==7){
          volas7 <- volas
          gammas7 <- gammas
          vegas7 <- vegas
          deltas7 <- deltas
          data7 <- data
        }
        if (sheet_no==8){
          volas8 <- volas
          gammas8 <- gammas
          vegas8 <- vegas
          deltas8 <- deltas
          data8 <- data
        }
        if (sheet_no==9){
          volas9 <- volas
          gammas9 <- gammas
          vegas9 <- vegas
          deltas9 <- deltas
          data9 <- data
        }
        if (sheet_no==10){
          volas10 <- volas
          gammas10 <- gammas
          vegas10 <- vegas
          deltas10 <- deltas
          data10 <- data
        }
        if (sheet_no==11){
          volas11 <- volas
          gammas11 <- gammas
          vegas11 <- vegas
          deltas11 <- deltas
          data11 <- data
        }
        if (sheet_no==12){
          volas12 <- volas
          gammas12 <- gammas
          vegas12 <- vegas
          deltas12 <- deltas
          data12 <- data
        }
        
        cashdesk[1,i] <- -as.numeric(data[1,i+1])+deltas[1,i]*data$S[1] #initializing cashdesk at first day:
        #buying long call and selling short delta units of underlying 
        
        if (k>1) {
          newdelta <- if (!is.na(deltas[k,i])) {deltas[k,i]} else {newdelta} #if newdelta is not available, do not update
          pastdelta <- if (!is.na(deltas[k-1,i])) {deltas[k-1,i]} else {pastdelta} #if pastdelta is not available, do not update
          option_value_changes[k-1,i] = (value - previous_value)
          delta_hedge_value_changes[k-1,i] = -previous_delta * (underlying - previous_underlying)
          cashdesk[k,i] <- if (newdelta==0 || pastdelta ==0){cashdesk[k,i]}else{cashdesk[k-1,i]+(newdelta-pastdelta)*data$S[k]} #if new and pastdelta differ from zero, sell short delta change amount of current underlying.
        }
        if (k==(length(data$daystomaturity)-1)) {
          cashdesk[k,i] <- cashdesk[k,i]+max(0,underlying-strike)-newdelta*underlying
          finalcash <- append(finalcash,cashdesk[k,i])
          no_hedge_cash <- append(no_hedge_cash,max(0,underlying-strike)-as.numeric(data[1,i+1]))
        }
        # Update the state
        
        previous_value <- value
        previous_underlying <- underlying
        if (!is.na(results$delta)) {
          previous_delta <- results$delta
        }
      }
    }

    daily_errors_squared <- (option_value_changes + delta_hedge_value_changes) ^2
    print(paste("Calculated sheet no ",sheet_no))
    strike_performances <- colSums(daily_errors_squared) / (length(data$daystomaturity)-2)
    sheets_results_mean <- append(sheets_results_mean, mean(strike_performances[strike_performances != 0]))
    sheets_results_std <- append(sheets_results_std, sd(strike_performances[strike_performances != 0]))
    sheets_absdiff_mean <- append(sheets_absdiff_mean,mean(abs(no_hedge_cash-finalcash)))
    sheets_absdiff_std <- append(sheets_absdiff_std,sd(abs(no_hedge_cash-finalcash)))
  },error=function(e){})
}    
data1 <- data1[-86,]

    #delta-vega-gamma hedge:
    for (s in 1:1){
      tryCatch({
      
      kokodatat <- vector()
      dnam <- paste("deltas",s,sep="")
      gnam <- paste("gammas",s,sep="")
      vnam <- paste("vegas",s,sep="")
      datnam <- paste("data",s,sep="")
      
      sheetresults2 <- c()
      sheetstds2 <- c()
      sheetcashes <- c()
      for (m in 2:(length(get(datnam))-3)){
        if(length(which(!is.na(get(datnam)[,m])))==length(which(!is.na(get(datnam)[,2])))[1]){kokodatat<-append(kokodatat,m)}
      }
      
      
      for (z in 1:length(kokodatat)) {
        
      eka<-kokodatat[z]
      kokodatat2 <- kokodatat[-z]
      for (h in 1:(length(combn(kokodatat2,2))/2)){
        
        
        
      toka<-combn(kokodatat2,2)[,h][1]
      kolmas<-combn(kokodatat2,2)[,h][2]
      Portfoliodelta = get(dnam)[1,eka-1]
      Portfoliogamma = get(gnam)[1,eka-1]
      Portfoliovega = get(vnam)[1,eka-1]
      optio_2_positio=0
      optio_3_positio=0
      underlying_positio=0
      optio_2_prev=0
      optio_3_prev=0
      underlying_prev=0
      counter=0
      
      
      dailyerrors=c(rep(0,length(which(!is.na(get(datnam)[,2])))-1))
      cashdesk1=c(rep(0,length(which(!is.na(get(datnam)[,2])))))
      for (n in 2:(length(which(!is.na(get(datnam)[,2])))+1)){
        A=matrix(
          c(get(dnam)[n-1,toka-1],get(dnam)[n-1,kolmas-1],1,get(gnam)[n-1,toka-1],get(gnam)[n-1,kolmas-1],0,get(vnam)[n-1,toka-1],get(vnam)[n-1,kolmas-1],0),
          nrow=3,
          ncol=3,
          byrow=TRUE
        )
        
        B=c(-Portfoliodelta,-Portfoliogamma,-Portfoliovega)
        optio_2_prev=optio_2_positio
        optio_3_prev=optio_3_positio
        underlying_prev=underlying_positio
        
        if (any(is.na(A))||any(is.na(B))){
          optio_2_positio=optio_2_positio
          optio_3_positio=optio_3_positio
          underlying_positio=underlying_positio
        } else {
          optio_2_positio= solve(A,B)[1]
          optio_3_positio= solve(A,B)[2]
          underlying_positio= solve(A,B)[3]
        }
        if (n==2){
        cashdesk1[1] = -as.numeric(get(datnam)[1,z])-optio_2_positio*as.numeric(get(datnam)[1,toka])-optio_3_positio*as.numeric(get(datnam)[1,kolmas])-underlying_positio*as.numeric(get(datnam)$S[1])
        }
        if (n>2){
        cashdesk1[n-1] = cashdesk1[n-2]-(optio_2_positio-optio_2_prev)*as.numeric(get(datnam)[n-1,toka])-(optio_3_positio-optio_3_prev)*as.numeric(get(datnam)[n-1,kolmas])-(underlying_positio-underlying_prev)*as.numeric(get(datnam)$S[n-1])
        }
        
        if (n==(length(which(!is.na(get(datnam)[,2])))+1)){
        cashdesk1[n-1] = cashdesk1[n-1]+optio_2_positio*as.numeric(get(datnam)[n-1,toka])+optio_3_positio*as.numeric(get(datnam)[n-1,kolmas])+underlying_positio*as.numeric(get(datnam)$S[n-1])
        }
        
        if (n<=(length(which(!is.na(get(datnam)[,2]))))){
        Portfoliodelta = get(dnam)[n,eka-1]
        Portfoliogamma = get(gnam)[n,eka-1]
        Portfoliovega = get(vnam)[n,eka-1]
        
        Portfoliovaluechange = as.numeric(get(datnam)[n,eka])-as.numeric(get(datnam)[n-1,eka])
        Replicatingvaluechange = (optio_2_positio*(as.numeric(get(datnam)[n,toka])-as.numeric(get(datnam)[n-1,toka]))
                                  +optio_3_positio*(as.numeric(get(datnam)[n,kolmas])-as.numeric(get(datnam)[n-1,kolmas]))
                                  +underlying_positio*(as.numeric(get(datnam)[n,(length(get(datnam))-2)])-as.numeric(get(datnam)[n-1,(length(get(datnam))-2)])))
        dailyerrors[n-1]=if(optio_2_positio==0&&optio_3_positio==0&&underlying_positio==0){0}else{Portfoliovaluechange+Replicatingvaluechange}
        counter= if(optio_2_positio==0&&optio_3_positio==0&&underlying_positio==0){counter+1}else{counter}
        }
      }
      dailyerrors
      performance=sum(dailyerrors^2)/(length(which(!is.na(get(datnam)[,2])))-1-counter)
      sheetresults2 <- append(sheetresults2,performance)
      sheetcashes <- append(sheetcashes,cashdesk1[n-1])
      }
      }
    sheets_results2_mean <- append(sheets_results2_mean,mean(sheetresults2))
    sheets_results2_std <- append(sheets_results2_std,sd(sheetresults2))
      },error=function(e){})
    }
print(sheets_results_mean)
print(sheets_results_std)
print(sheets_absdiff_mean)
print(sheets_absdiff_std)
print(sheets_results2_mean)
print(sheets_results2_std)

