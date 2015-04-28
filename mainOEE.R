# ==========================================================================
# This is main program to simuating OOE (Overall Operation Effectiveness).
#                                               Created by Dr. Jaeho H. BAE
#                                                       Assistant Professor
#                Dept. of Logistics & Distribution Mgmt. at Hyechon College
#                                                                May., 2014.
#                                                      chillangri@gmail.com
# ==========================================================================

mainOEE <- function() {
  # 1. 분석에 사용할 파일명 입력 및 파일 데이터 로드
  dataFile <- getFileName()                # Determining file name and its deposit directory
    if(is.null(dataFile)==TRUE) on.exit()  # If working directory is not correct, this function will be terminated.
  rawData <- loadData(dataFile)            # Reading & saving data into memory
  
  # 2. 기본 분석 - ANOVA 및 Boxplot
  analysis.aov <- aov(R_act ~ P_Group, data=rawData)
  summary(analysis.aov)
  print(analysis.aov)
  boxplot(R_act ~ P_Group, data=rawData)
  
  gMethod <- selectMethod()                # Select a data-grouping method

  if (gMethod == 0) {                      # No grouping
    ctMethod <- selectCT()                   # Define a method of calculating theoretical Cycle time or Throughput rate
    result     <- as.list(ctMethod)
    result[[1]] <- calOEE(dataFile, rawData, ctMethod)
    sourceDate  <- rawData
    names(result)[1] <- "Total"
  } else if (gMethod == 1) {               # Group considering Processing materials and methods
    nameData <- names(table(rawData$"P_Group"))
    sourceData <- as.list(nameData)
    result     <- as.list(nameData)
    tempCtMethod <- selectCT(length(nameData))  # Define a method of calculating theoretical Cycle time or Throughput rate
    for(index in 1:length(nameData)) {
      sourceData[[index]] <- subset(x=rawData, subset=(rawData$"P_Group"==nameData[index]), select=c(P_Group, P_Code, Resource, W_Order, In_Qty, Good_Qty, NG_Qty, W_Time, L_T))
      ctMethod    <- list(method=tempCtMethod$method[index], value=tempCtMethod$value[index])
      result[[index]] <- calOEE(dataFile, sourceData[[index]], ctMethod)
    }
    names(result)  <- nameData

    result <- calTotal(length(nameData), result)
  }

  # 계산결과의 파일 저장
  dateLog <- format(Sys.time(), "%Y%m%d-%H%M%S")
  # saveResult(dateLog, sourceData)
  saveResult(dateLog, result)
  
  return(result)
}

saveResult <- function(saveDate, result) {

  fileName1 <- paste(getwd(), "/OEE-", saveDate, ".xlsx", sep="")
  nameData  <- names(result)
  lWB <- loadWorkbook(fileName1, create = TRUE)
  for(index in 1:length(nameData)) {
    createSheet(lWB, nameData[index])
    writeWorksheet(lWB, result[[index]], nameData[index], header=TRUE)
  }
  saveWorkbook(lWB)
}

calTotal <- function(numGroups, resultEachGroup) { 
  # ==========================================================================
  # Subfunction: calTotal("Number of groups", "result for each group")
  # ==========================================================================
  tResult <- resultEachGroup[[1]][,1]
  pResult <- resultEachGroup[[1]][,2]
  for (i in 2:numGroups) {
    tResult <- cbind(tResult, resultEachGroup[[i]][,1])
    pResult <- cbind(pResult, resultEachGroup[[i]][,2])
  }
  tX <- tResult[3,]/sum(tResult[3,])
  pX <- pResult[3,]/sum(pResult[3,])
  
  tICT  <- sum(tResult[1,]*tX)      # 1 * 1
  tACT  <- sum(tResult[2,]*tX)      # 2 * 1
  tTL   <- sum(tResult[3,])         # 3 * 1
  tTO   <- sum(tResult[4,])         # 4 * 1
  tLT   <- sum(tResult[5,])         # 5 * 1
  tTN   <- sum(tResult[6,])         # 6 * 1
  tLS   <- sum(tResult[7,])         # 7 * 1
  tTV   <- sum(tResult[8,])         # 8 * 1
  tLD   <- sum(tResult[9,])         # 9 * 1
  tA    <- tTO / tTL * 100          #10 * 1
  tP    <- tTN / tTO * 100          #11 * 1
  tQ    <- tTV / tTN * 100          #12 * 1
  tOEE  <- tTV / tTL * 100          #13 * 1

  pITH  <- sum(pResult[1,]*tX)      # 1 * 2
  pATH  <- sum(pResult[2,]*tX)      # 2 * 2  
  pPth  <- sum(pResult[3,])         # 3 * 2  
  pPO   <- sum(pResult[4,])         # 4 * 2
  pPthO <- sum(pResult[5,])         # 5 * 2
  pPact <- sum(pResult[6,])         # 6 * 2
  pPOact<- sum(pResult[7,])         # 7 * 2
  pPg   <- sum(pResult[8,])         # 8 * 2
  pPng  <- sum(pResult[9,])         # 9 * 2
#   pA    <- pPO / pPth               #10 * 2
#   pP    <- pPact / pPO              #11 * 2
#   pQ    <- pPg / pPact              #12 * 2
#   pOEE  <- pPg / pPth               #13 * 2
  pA    <- sum(pResult[10,]*tX)     #10 * 2
  pP    <- sum(pResult[11,]*tX)     #11 * 2
  pQ    <- sum(pResult[12,]*tX)     #12 * 2
  pOEE  <- pA * pP * pQ /10000      #13 * 2
  
  tResult <- zapsmall(c(tICT, tACT, tTL, tTO, tLT, tTN, tLS, tTV, tLD, tA, tP, tQ, tOEE), digits = 8)
  pResult <- zapsmall(c(pITH, pATH, pPth, pPO, pPthO, pPact, pPOact, pPg, pPng, pA, pP, pQ, pOEE), digits=9)
  newResult <- cbind(as.data.frame(tResult), as.data.frame(pResult))
  newResult <- as.data.frame(newResult)
  rownames(newResult)  <- c("Idl. CT | TH", "Act. CT | TH", "T(L) | P(th)", "T(O) | P(O)", "L(T) | P(th)-P(O)", "T(N) | P(act)", "L(S) | P(O)-P(act)", "T(V) | P(g)", "L(D) | P(ng)", "A(eff, %)", "P(eff, %)", "Q(eff, %)", "OEE(%)")
  colnames(newResult)  <- c("Time Based", "Qty. Based")

  resultEachGroup[[numGroups + 1]] <- newResult
  names(resultEachGroup)[numGroups + 1] <- "Total"

  return(resultEachGroup)
}

calOEE <- function(dataFile, rawData, thMethod) { 
  # ==========================================================================
  # Subfunction: calOEE("File Name", "Source Dataset", "Grouping Method)
  # ==========================================================================
  numRows <- length(rawData[,1])
  R_act   <- c(1:numRows)
  C_act   <- c(1:numRows)
  
  R_act   <- as.data.frame(R_act)
  C_act   <- as.data.frame(C_act)
  
  colnames(R_act)   <- "R_act"
  colnames(C_act)   <- "C_act"
  
  R_act[,1]   <- (rawData$Good_Qty + rawData$NG_Qty) / rawData$W_Time
  C_act[,1]   <- 1/R_act[,1]
  
  modData     <- cbind(rawData, R_act, C_act)
  
  wsSD        <- sd(modData$R_act)
  wsMedian    <- median(modData$R_act)
  wsMean      <- mean(modData$R_act)
  
  if (thMethod$method == "M") {
    C_th      <- thMethod$value
    R_th      <- 1/C_th
  } else {
    R_th        <- wsMedian + thMethod$value * wsSD
    C_th        <- 1/R_th
  }
  
  P_G         <- sum(modData$Good_Qty)  #total Good Product Quanltity
  P_NG        <- sum(modData$NG_Qty)    #total No good product Quantity
  P_act       <- P_G + P_NG             #total produced Quantity
  
  T_O         <- sum(modData$W_Time)   #Operating time
  L_T         <- sum(modData$L_T)      #Downtime Loss
  T_L         <- T_O + L_T             #Theoretical time to be utilized
  T_N         <- P_act * C_th          #Net operating time
  T_V         <- P_G * C_th            #Valuable operating time
  P_O         <- T_O * R_th             #Operating production quantity
  
  P_th        <- T_L * R_th            #Theoritical Produt Quantity
  R_act       <- P_act / T_O           #Actual Throughput
  C_act       <- 1/R_act               #Actual Cycle time
  
  A1         <- T_O / T_L *100            #Availablity Rate
  P1         <- T_N / T_O *100            #Performance Efficiency
  Q1         <- T_V / T_N *100            #Quality Rate
  OEE1       <- A1 * P1 * Q1 / 10000
  A2         <- P_O / P_th *100
  P2         <- P_act / P_O *100
  Q2         <- P_G / P_act *100
  OEE2       <- A2 * P2 * Q2 /10000
  
  result1     <- zapsmall(c(C_th, C_act, T_L, T_O, L_T, T_N, T_O-T_N, T_V, T_N-T_V, A1, P1, Q1, OEE1), digits = 8)
  result2     <- zapsmall(c(R_th, R_act, P_th, P_O, P_th - P_O, P_act, P_O-P_act, P_G, P_NG, A2, P2, Q2, OEE2), digits=9)
  result      <- cbind(result1, result2)
  result      <- as.data.frame(result)
  rownames(result)  <- c("Idl. CT | TH", "Act. CT | TH", "T(L) | P(th)", "T(O) | P(O)", "L(T) | P(th)-P(O)", "T(N) | P(act)", "L(S) | P(O)-P(act)", "T(V) | P(g)", "L(D) | P(ng)", "A(eff, %)", "P(eff, %)", "Q(eff, %)", "OEE(%)")
  colnames(result)  <- c("Time Based", "Qty. Based")
  #result    <- list(retName = c("A", "P", "Q", "OEE"), retValue = c(A, P, Q, OEE))
  
return(result)  
}

getFileName <- function() {
  # ==========================================================================
  # Subfunction: getFileName()
  # ==========================================================================
  currentWD <- getwd()
  cat("Your DATA file must be at ", currentWD, "\n", sep="")
  ANS <- readline("Is it right? (Y/y or N/n): ")
  
  if ((substr(ANS, 1, 1)=="n") || (substr(ANS, 1, 1)=="N")) {
    return()
  } else {
    fileName <- readline("What is the file Name? ")
    fileName <- paste(currentWD, "/", fileName, sep="")
    return(fileName)
  }
}

loadData <- function(dataFile) {
  # ==========================================================================
  # Subfunction: loadData("File Name")
  # ==========================================================================
  options(java.parameters = "-Xmx4096m")
  library(XLConnect)
  
  dataWB   <- loadWorkbook(dataFile)
  dataWS   <- getSheets(dataWB)
  
  cntWS    <- 0
  for(idx in dataWS) cntWS = cntWS+1
  
  if(cntWS==1){
    sourceData <- readWorksheet(dataWB, sheet=1)
    
  } else {
    cat("All Sheets on this file are followings:\n")
    for (i in 1:cntWS) cat("\t", i, "\t", dataWS[i], "\n")
    numSelect <- readline("Which Worksheet was made for Analysis? (number) ")
    numSelect <- as.integer(numSelect)
    sourceData <- readWorksheet(dataWB, sheet=dataWS[numSelect])
  }
  xlcFreeMemory()
  return(sourceData)
}

selectMethod <- function() {
  # ==========================================================================
  # Subfunction: selectMethod()
  # ==========================================================================
  cat("==================================================================================================\n")
  cat("Select a method grouping production data \n")
  cat("\t0:\tNo grouping(calculating OEE without considering processing materials and methods\n")
  cat("\t1:\tGrouping by product group (calculating OEE considering processing materials and methods\n")
  ANS <- readline("Input your data grouping method in numbers!! ")
  
  if ((substr(ANS, 1, 1)=="0") || (substr(ANS, 1, 1)=="1")) {
    return(ANS)
  } else {
    cat("Selected grouping method is not correct!!")
    return()
  }
}

selectCT <- function(numGroup = 1) {
  # ==========================================================================
  # Subfunction: selectCT()
  # ==========================================================================
  cat("Select a method determining Therotical Cycle Time or Throughput Rate \n")
  cat("\t0:\tTheoretical cycle time will be determined by user input!\n")
  cat("\t1~5:\tTheoretical cycle time = median + n times of standard derivation\n")
  ANS <- readline("How to determine theoretical cycle time? ")
  ANS <- as.numeric(ANS)

  if (ANS==0) {
    if (numGroup ==1) {
      ANS <- readline("Input theoretical cycle time ")
      returnValue <- list(method="M", value=as.numeric(ANS))
      return(returnValue)
    } else {
      for(i in 1:numGroup) {
        if (i == 1) {
          ANS <- readline("Input 1st Cycle Time\n")
          ANS <- as.numeric(ANS)
        } else {
          cat("Input ", i, "th Cycle Time")
          temp <- readline()
          temp <- as.numeric(temp)
          ANS  <- c(ANS, temp)
        }
      }
      returnValue <- list(method=rep("M", numGroup), value=ANS)
      return(returnValue)
    }
  } else if((ANS>0) || (ANS<=5)){
    returnValue <- list(method=rep("A", numGroup), value=rep(ANS, numGroup))
    return(returnValue)
  } else {
    cat("Your method is not proper!!")
    return()
  }
}