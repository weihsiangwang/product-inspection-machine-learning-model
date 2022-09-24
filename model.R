library(magrittr); library(stringr); library(scales); library(openxlsx); 
library(plyr); library(dplyr); library(tidyr); library(tibble); library(lubridate);

library(tidyverse); library(doParallel); library(foreach); library(smbinning); 
library(DMwR); library(h2o); library(glmnet); library(C50); library(rpart); library(caret)
library(future.apply)
library(PRROC)

setwd('D:\\Git Projects\\product-inspection-machine-learning-model')

# file name
output_name <- paste0('.\\result\\Model Analysis - final report.xlsx')
 
# 125 eigenfactor by data type
nvar_list <- as.vector(read.csv('.\\data\\nvar_list.csv')$NVAR.Factors)
cvar_list <- as.vector(read.csv('.\\data\\cvar_list.csv')$CVAR.Factors) %>%
  .[! . %in% c("Eigenfactor_104", "Eigenfactor_119", "Eigenfactor_120")]

# load preprocessing raw data
pre_data <- read.csv('.\\data\\raw_data.csv')
pre_data$Eigenfactor_126 <- pre_data$Eigenfactor_126 %>% as.character() %>% as.Date()



# training and test sets: splitting data by date (Eigenfactor_92)
## training data
training_data <- pre_data %>%
  filter(Eigenfactor_126 <= '2018-12-31')

## testing data 
testing_data <- pre_data %>%
  filter(Eigenfactor_126 >= '2019-01-01')


############################################# 
################## Binning ################## 
############################################# 
s_result <- function(train, test){
  # binning: training data
  cat('------ BINNING START ------', '\n')
  
  ### number factor first ###
  
  # determining the Eigenfactors which with more than 5 factors
  index <- lapply(apply(train[,nvar_list],2,unique),length)
  L5 <- colnames(train[,nvar_list])[which(index >=2 & index<5)]
  nvar_G5_list <- nvar_list[-which(nvar_list%in%L5)]
  
  n_sig_list <- c() # significant nvar factor
  cut_ref_output <- c() # output: cut point reference
  bins_iv <- c()

  # using Inspection_Result to bin each item in "nvar_G5_list"
  cat('1. binning the numeric factors', '\n')
  for(kk in 1:length(nvar_G5_list)) {
    cat(nvar_G5_list[kk], '\t')
    bins <- eval(parse(text=paste0("smbinning(train,'Inspection_Result','",nvar_G5_list[kk],"')")))
    bins_iv$var[kk] <- nvar_G5_list[kk]
    bins_iv$iv[kk] <- ifelse(is.list(bins), bins$iv, NA)
    
    # cut point
    if(is.list(bins)==TRUE){ 
      # renaming for each cut point 
      bin_names <- bins$ivtable$Cutpoint[1:(nrow(bins$ivtable)-2)]
      bin_names <- gsub('<= ','LE',bin_names) # eg. LE100
      bin_names <- gsub('> ','GT',bin_names) # eg. GT1000
      b <- c(-Inf,bins$cuts,Inf)
      # show answer
      cat('cut: ', b, '\n')
      
      cut_ref <- as.data.frame(cbind(factor_name=nvar_G5_list[kk],
                                   cut=paste(b,collapse = ','),
                                   name=paste(bin_names,collapse=',')))
      cut_ref_output <- rbind(cut_ref_output, cut_ref)
      
      bin_breaks <- cut_ref %>%
        dplyr::filter(factor_name == nvar_G5_list[kk]) %>%
        dplyr::select(cut) %>%
        unlist() %>%
        as.character() %>%
        strsplit(.,split=',') %>%
        unlist() %>%
        as.numeric()
      
      bin_names <- cut_ref %>%
        dplyr::filter(factor_name == nvar_G5_list[kk]) %>%
        dplyr::select(name) %>%
        unlist() %>%
        as.character() %>%
        strsplit(.,split=',') %>%
        unlist()    
      
      # insert cut point to training set
      eval(parse(text=paste0('train$',nvar_G5_list[kk], ' <- cut(train$',nvar_G5_list[kk],',breaks = bin_breaks, labels =  bin_names)')))
      n_sig_list <- c(n_sig_list,nvar_G5_list[kk])
      
    }else{
      # return "No significant splits"
      cat(nvar_G5_list[kk], '\t')
      cat(bins, '\n')
      }
    
  }
  bins_iv <- bins_iv %>% as.data.frame()
  # bins_iv1 <- bins_iv %>% 
  #   filter(!is.na(iv))
  
  cut_ref_output <- cbind(cut_ref_output, update_time=Sys.time())
  
  openxlsx::write.xlsx(cut_ref_output, paste0('.\\result\\Model Analysis - cut point reference.xlsx'))
  
  ### character factor ###
  cat('2. binning the character factors', '\n')
  cat('-- doing Fisher exact test', '\n')
  ct_result <- c()
  for(i in 1 : length(cvar_list)) {
    cat(cvar_list[i], '\n')
    # change Inspection_Result and each Eigenfactor in cvar_list to factor 
    t_tb <- data.frame(Inspection_Result = as.factor(ifelse(train$Inspection_Result==1, 'fail', 'pass')),
                          t_var=as.factor(eval(parse(text=paste0('train$', cvar_list[i])))))
    
    # Fisher's exact test if 2 factors at least
    if(length(levels(t_tb$t_var))>=2){
      ct <- table(t_tb$t_var, t_tb$Inspection_Result)
      p_value <- fisher.test(ct,simulate.p.value = TRUE)$p.value
      ct_result <- rbind(ct_result, cbind(cvar_list[i], p_value))
    }
    paste(i)
  }
  
  # significant character factor list
  c_sig_list <- ct_result[which(ct_result[,2] != 'NaN'), ]
  c_sig_list <- c_sig_list[which(as.numeric(as.character(c_sig_list[, 2])) <= 0.05), 1]
  
  # caculate the length of significant factors
  lvleng<-c()
  for(i in 1:length(c_sig_list)) {
    lvleng[i] <- nlevels(eval(parse(text=paste0('train$', c_sig_list[i]))))
  }
  
  # binng the factor with length more than 6
  gp_list <- as.vector(c_sig_list[which((lvleng>=6))])
  
  ## binng according the failure inspection rate
  cat('-- binng significant factors according the failure inspection rate', '\n')
  cut_p <- prop.table(table(train$Inspection_Result))[2]
  if(length(gp_list)>=1){
    for(i in 1:length(gp_list)) {
      cat(gp_list[i], '\n')
      
      ct <- table(eval(parse(text = paste0('train$', gp_list[i]))), train$Inspection_Result)
      p_ct <- as.data.frame(round(prop.table(ct, 1), 10))
      cut_level <- p_ct[which(p_ct$Var2 == 1 & p_ct$Freq >= cut_p), c('Var1')]
      
      str_t <- paste0('train <- train %>% mutate(',gp_list[i],'gp=ifelse(',gp_list[i],' %in% as.vector(cut_level),1,0))')
      eval(parse(text=str_t))
    }
    ## binning character factor list 
    cat1 <- paste0(gp_list, 'gp')
    ## non-binning character factor list 
    cat2 <- c_sig_list[(!c_sig_list %in% gp_list)]
    colname_trans <- c(cat1, cat2, n_sig_list)
    all_sig_list <- c('Inspection_Result',cat1, cat2, n_sig_list, L5)
  }else{
    ## non-binning character factor list 
    cat2 <- c_sig_list[(!c_sig_list %in% gp_list)]
    colname_trans <- c(cat2, n_sig_list)
    all_sig_list <- c('Inspection_Result', cat2, n_sig_list, L5)
  }
  
  train <- data.frame(train)
  # as.factor for each factor
  for( i in 1:length(colname_trans)) {
    eval(parse(text=paste0("train$",colname_trans[i]," <- as.factor(train$",colname_trans[i],")")))
  }
  
  # training data frame with all significant factors
  train1 <- train[, all_sig_list]
  
  #### TRAINING MODEL ####
  cat('------ TRAINING MODEL ------', '\n')
  # Do Parallel
  cl <- makeCluster(8)
  registerDoParallel(cl)
  train1 <- train1[complete.cases(train1), ]
  x <- model.matrix(Inspection_Result~ ., train1)
  x <- x[, sort(colnames(x))]
  y <- train1$Inspection_Result
  model <- cv.glmnet(x, y, 
                     family = "binomial", 
                     type.logistic = "modified.Newton", 
                     parallel = T, 
                     alpha = 1, 
                     lambda = exp(seq(log(0.001), log(5), length.out=500))
                     )
  
  colname <- colnames(train1)
  final_name <- c()
  for(i in 1:length(colname)){
    text <- paste0('train1$',colname[i])
    value <- eval(parse(text=text))
    final_name <- rbind(final_name, 
                        cbind(colname[i], 
                              paste0(colname[i], as.character(unique(value)))
                              )
                        )
  }
  
  coeff <- as.matrix(coef(model, s=min(model$lambda)))
  select_var <- row.names(coeff)[which(abs(coeff) >= 0.001)]
  select_var <- unique(final_name[which(final_name[, 2] %in% select_var), 1])
  select_var <- c(select_var, 'Inspection_Result')
  
  # Stop Parallel
  stopCluster(cl)
  
  
  
  ##### Risk Factor Analysis Results 
  cat('------ Risk Factor Analysis Results ------', '\n')
  select_var_r1 <- gsub(c('gp'), '', select_var)
  vars_all <- as.data.frame(cbind('Inspection_Result', c(nvar_list, cvar_list)))
  colnames(vars_all) <- c('Y', 'riskFactor')
  vars_all$sigleVar <- ifelse(vars_all$riskFactor %in% c(c_sig_list, n_sig_list), 'Y', 'N')
  vars_all$multiVar <- ifelse(vars_all$riskFactor %in% select_var, 'Y', 'N')
  
  openxlsx::write.xlsx(vars_all, paste0('.\\result\\Model Analysis - risk factors.xlsx'))
  
  train2 <- train1 %>%
    dplyr::select(select_var) %>%
    cbind(Document_No=train$Document_No, Product_No=train$Product_No)
  
  #### binning testing data ####
  var_unify <- function(test){
    out_data <- test
    # character factor 
    if(length(gp_list) >= 1){
      for(i in 1:length(gp_list)) {
        tb <- as.data.frame(table(eval(parse(text=paste0('train$',gp_list[i]))),eval(parse(text=paste0('train$',gp_list[i],'gp')))))
        cut_level <- tb[which(tb$Var2==1 & tb$Freq>=1), c('Var1')]
        
        str_t<-paste0('out_data<-out_data %>% mutate(',gp_list[i],'gp=ifelse(',gp_list[i],' %in% as.vector(cut_level),1,0))')
        eval(parse(text=str_t))
      }
    }
    
    return(out_data)
  }
  
  cut_var <- as.vector(cut_ref_output$factor_name)
  for(i in 1:length(cut_var)) {
    bin_breaks <- cut_ref_output %>%
      filter(factor_name==cut_var[i]) %>%
      dplyr::select(cut) %>%
      unlist() %>%
      as.character() %>%
      strsplit(.,split=',') %>%
      unlist() %>%
      as.numeric()
    
    bin_names <- cut_ref_output %>%
      filter(factor_name==cut_var[i]) %>%
      dplyr::select(name) %>%
      unlist() %>%
      as.character() %>%
      strsplit(.,split=',') %>%
      unlist()    
    
    # insert cut point to testing sets
    eval(parse(text=paste0('test$',cut_var[i], '<- cut(test$',cut_var[i],',breaks = bin_breaks, labels =  bin_names)')))
    
    # print(cut_var[i])
    # print(bin_breaks)
    # print(bin_names)
    
  }
  
  test1 <- var_unify(test)
  
  
  test2 <- test1 %>%
    var_unify() %>%
    dplyr::select(select_var) %>%
    cbind(Document_No=test$Document_No, Product_No=test$Product_No)
  
  
  TrainingData_binning <<- train2
  TestingData_binning <<- test2
  
}



############################################# 
################### SMOTE ################### 
############################################# 

goSMOTE <- function(x, y){
  x <- x %>% lapply(as.factor) %>% as.data.frame()
  y <- y %>% lapply(as.factor) %>% as.data.frame()
  
  
  xGO <- select(x, -c("Document_No", "Product_No"))
  # yGO <- select(y, -c("Document_No", "Product_No"))
  yGO <- y
  
  # 
  transcol <- setdiff(colnames(yGO), c("Document_No", "Product_No"))
  for (i in transcol) {
    eval(parse(text=paste0("yGO$", i, " <- factor(yGO$", i, 
                           ", levels = levels(xGO$", i, "))")
    ))
  }
  
  ### SMOTE
  set.seed(myseed)
  TrainingData_SMOTE = SMOTE(Inspection_Result~., data = xGO, perc.over = 2000, perc.under = 250) #3:7
  
  print( 
    paste0(
      'original ratio�G', nrow(filter(xGO, Inspection_Result == 0)) / nrow(filter(xGO, Inspection_Result == 1))
    )
  ) 
  
  print(
    paste0(
      'SMOTE ratio�G', nrow(filter(TrainingData_SMOTE, Inspection_Result == 0)) / nrow(filter(TrainingData_SMOTE, Inspection_Result == 1))
    )
  ) 
  
  
  TrainingData_SMOTE <<- TrainingData_SMOTE
  TestingData_binningGO <<- yGO
}



############################################# 
################### Model ################### 
############################################# 

sevenModel <- function(checo, rosberg){
  
  vettel <- select(rosberg, -c("Document_No", "Product_No"))
  
  ### Modeling
  h2oTrain <- checo
  names(h2oTrain) = seq(1, ncol(h2oTrain)) 
  colnames(h2oTrain)[ncol(h2oTrain)] = "class"
  
  h2oTest <- vettel
  names(h2oTest) = seq(1, ncol(h2oTest))  
  colnames(h2oTest)[ncol(h2oTest)] = "class"
  
  ncol(h2oTrain)
  y <- 'class' ; x <- setdiff(names(h2oTrain), y)
  
  train.h2o <- as.h2o(h2oTrain)
  test.h2o <- as.h2o(h2oTest)
  
  
  # Elastic Net
  cat('Elastic Net Modeling \n')
  EN.Train = model.matrix(Inspection_Result~ ., checo)
  
  set.seed(myseed)
  EN.final = cv.glmnet(
    x = EN.Train, 
    y = checo$Inspection_Result,
    family="binomial",
    type.logistic="modified.Newton",
    nfolds=3,
    parallel=)
  
  # GBM
  cat('GBM Modeling \n')
  h2o.gbm.final <- h2o.gbm(
    x = x,
    y = y,
    training_frame = train.h2o,
    nfolds = 3,
    fold_assignment = "Modulo",
    keep_cross_validation_predictions = TRUE,
    seed = myseed
  )
  
  # C5.0
  cat('C5.0 Modeling \n')
  set.seed(myseed)
  C50.final = C5.0(Inspection_Result ~ ., data = checo)
  
  # CART
  cat('CART Modeling \n')
  set.seed(myseed)
  CART.final = rpart(Inspection_Result ~ ., data = checo)
  
  # Naive Bayes 
  cat('Naive Bayes Modeling \n')
  h2o.NB.final = h2o.naiveBayes(
    x = x, 
    y = y, 
    training_frame = train.h2o,
    nfolds = 3,
    fold_assignment = "Modulo",
    keep_cross_validation_predictions = TRUE,
    seed = myseed
  )
  
  # Random Forest
  cat('Random Forest Modeling \n')
  h2o.RF.final = h2o.randomForest(
    x = x, 
    y = y, 
    training_frame = train.h2o, 
    nfolds = 3,
    fold_assignment = "Modulo",
    keep_cross_validation_predictions = TRUE,
    seed = myseed
  )
  
  # Logistic Regression
  cat('Logistic Regression Modeling \n')
  h2o.glm.final = h2o.glm(
    x = x, 
    y = y, 
    training_frame = train.h2o,
    nfolds = 3,
    fold_assignment = "Modulo",
    keep_cross_validation_predictions = TRUE,
    family = "binomial",
    lambda_search = TRUE,
    seed = myseed
  )
  
  
  
  ### Predict
  value.predict <- NULL
  # Elastic Net
  # cat('Elastic Net Predicting \n')
  # EN.Test = model.matrix(Inspection_Result~ ., vettel)
  # 
  # EN.predict = predict(EN.final, EN.Test, s = "lambda.min", type = "response") %>% as.data.frame() %>% .$`1`
  # 
  # EN.temp <- cbind(as.character(rosberg$Document_No), as.character(rosberg$Product_No), "EN", EN.predict) %>% data.frame() 
  # names(EN.temp) = c("Document_No", "Product_No", "Model", "Predicted_Probability")
  # 
  # value.predict <- value.predict %>% rbind(EN.temp)
  
  # EN.predict.R = ifelse(EN.predict >= 0.5, 1, 0) %>% as.factor() 
  # EN.report = cbind(as.character(EN.predict.R), as.character(vettel$Inspection_Result)) %>% data.frame() 
  # names(EN.report) = c("Predicted","Actual")
  # confusionMatrix(table(EN.report), positive = "1")$byClass[c('Balanced Accuracy', 'Sensitivity', 'Pos Pred Value', 'F1')]
  
  # GBM
  cat('GBM Predicting \n')
  GBM.predict = predict(h2o.gbm.final, test.h2o) %>% as.data.frame() %>% .$p1
  
  GBM.temp <- cbind(as.character(rosberg$Document_No), as.character(rosberg$Product_No), "GBM", GBM.predict) %>% data.frame() 
  names(GBM.temp) = c("Document_No", "Product_No", "Model", "Predicted_Probability")
  
  value.predict <- value.predict %>% rbind(GBM.temp)
  
  # GBM.predict.R = ifelse(GBM.predict >= 0.5, 1, 0) %>% as.factor() 
  # GBM.report = cbind(as.character(GBM.predict.R), as.character(h2oTest$class)) %>% data.frame() 
  # names(GBM.report) = c("Predicted","Actual")
  # confusionMatrix(table(GBM.report), positive = "1")$byClass[c('Balanced Accuracy', 'Sensitivity', 'Pos Pred Value', 'F1')]
  
  # C5.0
  cat('C5.0 Predicting \n')
  C50.predict = predict(C50.final, vettel, type = "prob")[, 2] %>%
    as.numeric()
  
  C50.temp <- cbind(as.character(rosberg$Document_No), as.character(rosberg$Product_No), "C50", C50.predict) %>% data.frame() 
  names(C50.temp) = c("Document_No", "Product_No", "Model", "Predicted_Probability")
  
  value.predict <- value.predict %>% rbind(C50.temp)
  
  # C50.predict.R = ifelse(C50.predict >= 0.5, 1, 0) %>% as.factor() 
  # C50.report = cbind(as.character(C50.predict.R), as.character(vettel$Inspection_Result)) %>% data.frame() 
  # names(C50.report) = c("Predicted","Actual")
  # confusionMatrix(table(C50.report), positive = "1")$byClass[c('Balanced Accuracy', 'Sensitivity', 'Pos Pred Value', 'F1')]
  
  # CART
  cat('CART Predicting \n')
  CART.predict = predict(CART.final, vettel, type = "prob")[, 2] %>%
    as.numeric()
  
  CART.temp <- cbind(as.character(rosberg$Document_No), as.character(rosberg$Product_No), "CART", CART.predict) %>% data.frame() 
  names(CART.temp) = c("Document_No", "Product_No", "Model", "Predicted_Probability")
  
  value.predict <- value.predict %>% rbind(CART.temp)
  
  # CART.predict.R = ifelse(CART.predict >= 0.5, 1, 0) %>% as.factor() 
  # CART.report = cbind(as.character(CART.predict.R), as.character(vettel$Inspection_Result)) %>% data.frame() 
  # names(CART.report) = c("Predicted","Actual")
  # confusionMatrix(table(CART.report), positive = "1")$byClass[c('Balanced Accuracy', 'Sensitivity', 'Pos Pred Value', 'F1')]
  
  # NaiveBayes 
  cat('Naive Bayes Predicting \n')
  NB.predict = predict(h2o.NB.final, test.h2o) %>% as.data.frame() %>% .$p1
  
  NB.temp <- cbind(as.character(rosberg$Document_No), as.character(rosberg$Product_No), "NB", NB.predict) %>% data.frame() 
  names(NB.temp) = c("Document_No", "Product_No", "Model", "Predicted_Probability")
  
  value.predict <- value.predict %>% rbind(NB.temp)
  
  # NB.predict.R = ifelse(NB.predict >= 0.5, 1, 0) %>% as.factor() 
  # NB.report = cbind(as.character(NB.predict.R), as.character(h2oTest$class)) %>% data.frame() 
  # names(NB.report) = c("Predicted","Actual")
  # confusionMatrix(table(NB.report), positive = "1")$byClass[c('Balanced Accuracy', 'Sensitivity', 'Pos Pred Value', 'F1')]
  
  # Random Forest
  cat('Random Forest Predicting \n')
  RF.predict = predict(h2o.RF.final, test.h2o) %>% as.data.frame() %>% .$p1
  
  RF.temp <- cbind(as.character(rosberg$Document_No), as.character(rosberg$Product_No), "RF", RF.predict) %>% data.frame() 
  names(RF.temp) = c("Document_No", "Product_No", "Model", "Predicted_Probability")
  
  value.predict <- value.predict %>% rbind(RF.temp)
  
  # RF.predict.R = ifelse(RF.predict >= 0.5, 1, 0) %>% as.factor() 
  # RF.report = cbind(as.character(RF.predict.R), as.character(h2oTest$class)) %>% data.frame() 
  # names(RF.report) = c("Predicted","Actual")
  # confusionMatrix(table(RF.report), positive = "1")$byClass[c('Balanced Accuracy', 'Sensitivity', 'Pos Pred Value', 'F1')]
  
  # Logistic Regression
  cat('Logistic Regression Predicting \n')
  GLM.predict = predict(h2o.glm.final, test.h2o) %>% as.data.frame() %>% .$p1
  
  GLM.temp <- cbind(as.character(rosberg$Document_No), as.character(rosberg$Product_No), "GLM", GLM.predict) %>% data.frame() 
  names(GLM.temp) = c("Document_No", "Product_No", "Model", "Predicted_Probability")
  
  value.predict <- value.predict %>% rbind(GLM.temp)
  
  # GLM.predict.R = ifelse(GLM.predict >= 0.5, 1, 0) %>% as.factor() 
  # GLM.report = cbind(as.character(GLM.predict.R), as.character(h2oTest$class)) %>% data.frame() 
  # names(GLM.report) = c("Predicted","Actual")
  # confusionMatrix(table(GLM.report), positive = "1")$byClass[c('Balanced Accuracy', 'Sensitivity', 'Pos Pred Value', 'F1')]
  
  ### Voting
  # cat('Voting \n')
  # vote.report = data.frame(ElasticNet = EN.predict.R %>% as.character() %>% as.integer(),
  #                          GBM = GBM.predict.R %>% as.character() %>% as.integer(),
  #                          C50 = C50.predict.R %>% as.character() %>% as.integer(),
  #                          CART = CART.predict.R %>% as.character() %>% as.integer(),
  #                          NaiveBayes = NB.predict.R %>% as.character() %>% as.integer(),
  #                          RandomForest = RF.predict.R %>% as.character() %>% as.integer(),
  #                          LogisticRegression = GLM.predict.R %>% as.character() %>% as.integer()) %>%
  #   mutate(votes = rowSums(.),
  #            voting_method = ifelse(rowSums(.)>=4, 1, 0) %>% as.factor(),
  #            actual = vettel$Inspection_Result)
  # 
  # CM <<- confusionMatrix(table(vote.report[, c('voting_method', 'actual')]), positive = "1")
  
  value.predict <<- value.predict
  
}


############################################# 
#################### PRF #################### 
############################################# 
### precision, recall, FPR
PRF <- function (x, y, B) {
  Table = table(x, y)
  if (length(Table) == 4) {
    precision = round(Table[1, 1]/sum(Table[1, ]), 6)
    recall = round(Table[1, 1]/sum(Table[, 1]), 6)
    FPR = round(Table[1, 2]/sum(Table[, 2]), 6)
    Fmeasure = ((1 + B^2) * precision * recall)/(B * 2 * 
                                                   precision + recall)
  }
  else {
    precision = NA
    recall = NA
    Fmeasure = NA
    FPR = NA
  }
  Result = c(precision, recall, FPR, Fmeasure)
  return(Result)
}


############################################# 
#################### GO ##################### 
############################################# 


### Parallel ###
plan(multicore, workers = 6L)


h2o.shutdown(prompt = TRUE)
h2o.no_progress()
h2o.init(nthreads = -1,
         max_mem_size = "32G")

QQQ1 <- Sys.time()

Result <- NULL
PRFscore <- NULL

# Binning
cat(paste0('Binning', '\n'))

Sys.time()
oldTime <- Sys.time()

s_result(training_data, testing_data)

newTime <- Sys.time()
runTime <- round(difftime(newTime, oldTime, units = 'secs'), 4)
cat(paste0('----------------time spend', runTime, 's'),'\n')


# modeling 10 times 
temp <- NULL
for (j in 1:10) {
  cat(paste0('---------------------------------',  j, ' time ---------------------------------\n'))
  
  myseed = 2021*j + 3*j^2 + 23*j^3
  
  # SMOTE
  cat(paste0('SMOTE', '\n'))
  
  Sys.time()
  oldTime <- Sys.time()
  
  goSMOTE(TrainingData_binning, TestingData_binning)
  
  newTime <- Sys.time()
  runTime <- round(difftime(newTime, oldTime, units = 'secs'), 4)
  cat(paste0('----------------time spend', runTime, 's'),'\n')
  
  
  # Modeling and Predicting
  sevenModel(TrainingData_SMOTE, TestingData_binningGO)
  
  temp <- temp %>% rbind(value.predict)
  rm(value.predict)
  
}


temp$Predicted_Probability <- temp$Predicted_Probability %>% as.character() %>% as.numeric()

temp.mean <- aggregate(Predicted_Probability ~ ., data = temp, mean)
temp.mean$predict = ifelse(temp.mean$Predicted_Probability >= 0.5, 1, 0) 


temp.all <- temp.mean %>%
  select("Document_No", "Product_No") %>%
  unique()


models = as.character(unique(temp.mean$Model))
for (k in models) {
  temp.model <- temp.mean %>%
    filter(Model == k)
  
  temp.model$predict = ifelse(temp.model$Predicted_Probability >= 0.5, 1, 0) 
  
  temp.all <- merge(temp.all, temp.model[, c("Document_No", "Product_No", "predict")])
  
  colnames(temp.all)[which(colnames(temp.all) == 'predict')] <- k
  
  rm(temp.model)
}


temp.all <- temp.all %>%
  mutate(votes = rowSums(temp.all[models])) %>%
  mutate(voting_method = ifelse(votes >= 4, 1, 0) %>% as.factor()) %>%
  merge(select(TestingData_binningGO, c("Document_No", "Product_No", "Inspection_Result")), all.x = TRUE) %>%
  unique()
  

# factor to 0, 1
temp.all <- temp.all %>% lapply(as.factor) %>% as.data.frame()
for (k in c(models, 'voting_method')) {
  eval(parse(text=paste0("temp.all$", k, " <- factor(temp.all$", k, 
                         ", levels = c(0, 1))")
  ))
}


for (k in c(models, 'voting_method')) {
  CM <- confusionMatrix(table(temp.all[, c(k, 'Inspection_Result')]), positive = "1")
  
  temp.result <- data.frame(
    'model' = k,
    'total_inspected' = CM$table %>% sum(),
    'TP' = CM$table['1', '1'],
    'FP' = CM$table['1', '0'],
    'FN' = CM$table['0', '1'],
    'TN' = CM$table['0', '0'],
    'Balanced.Accuracy' = CM$byClass['Balanced Accuracy'] %>% as.vector(),
    'Sensitivity' = CM$byClass['Sensitivity'] %>% as.vector(),
    'Specificity' = CM$byClass['Specificity'] %>% as.vector(),
    'Pos.Pred.Value' = CM$byClass['Pos Pred Value'] %>% as.vector(),
    'Neg.Pred.Value' = CM$byClass['Neg Pred Value'] %>% as.vector(),
    'F1' = CM$byClass['F1'] %>% as.vector()
  )
  
  Result <- Result %>% rbind(temp.result)
}

rm(temp, temp.result)




### AUC
temp.mean <- temp.mean %>%
  merge(select(TestingData_binningGO, c("Document_No", "Product_No", "Inspection_Result")), all.x = TRUE) %>%
  unique()

# inspection result to N�BY
temp.mean$Inspection_Result<- ifelse(temp.mean$Inspection_Result == 1, 'N', 'Y') %>% as.factor()

# filter the category 
PRFprod <- Result[Result[, c('TP', 'FN')] %>% rowSums() > 0, ] %>%
  filter(model != 'voting_method')

# GO
m=1
for (m in 1:nrow(PRFprod)) {
  modelF = PRFprod$model[m] %>% as.character()
  
  ModelData <- temp.mean %>% 
    filter(Model == modelF) 
  
  print(paste0('AUC   ', which(modelF == PRFprod$model), '/', nrow(PRFprod), '   model: ', modelF))
  
  cut <- t(seq(0, 1, 0.0001)) %>% as.list()
  p <- 0
  ypred <- NULL
  
  x <- ModelData$Predicted_Probability ; y <- ModelData$Inspection_Result ; B <- 1
  
  while (p < length(cut)) {
    ypred <- ypred %>% append(list(x))
    p <- p + 1
  }
  
  # detemining the threshold NY according cutting points
  AA <- future_mapply(function(x, y) { ifelse(x > y, "N", "Y") }, ypred, cut, SIMPLIFY = F)
  BB <- future_lapply(AA, function(x) { PRF(x, y, B) } ) %>% as.data.frame() %>% t() %>% as.data.frame()
  
  row.names(BB) <- NULL
  colnames(BB) <- c('Precision', 'Recall', 'FPR', 'Fbeta')
  BB$model <- modelF
  BB$cut <- cut %>% unlist()
  
  ### drawing AUC
  roc <- roc.curve(scores.class0 = ModelData[which(ModelData$Inspection_Result == 'N'), 'Predicted_Probability'], 
                   scores.class1 = ModelData[which(ModelData$Inspection_Result == 'Y'), 'Predicted_Probability'])
  
  pr <- pr.curve(scores.class0 = ModelData[which(ModelData$Inspection_Result == 'N'), 'Predicted_Probability'], 
                 scores.class1 = ModelData[which(ModelData$Inspection_Result == 'Y'), 'Predicted_Probability'])
  
  
  # BB$AUC <- auc_roc
  BB$AUC <- roc$auc
  BB$AUP <- pr$auc.davis.goadrich
  
  ### bind
  PRFscore <- PRFscore %>% rbind(BB)
  
  
  rm(modelF, ModelData, AA, BB, roc, pr)
}


rm(TrainingData_binning, TrainingData_SMOTE, TestingData_binning, TestingData_binningGO)


# write.csv(Result, 'Result.csv', row.names = FALSE)

QQQ1
Sys.time()


# rm(list=setdiff(ls(), c("SqlData", "InputData", "InputData1", "RawData", "PreData", "PreData1", "SqlLines", "nvar_list", "cvar_list")))



wb <- createWorkbook()

# Result
addWorksheet(wb, 'evaluation_indicators')
writeData(wb, 'evaluation_indicators', Result)

# AUC
addWorksheet(wb, 'AUC')
writeData(wb, 'AUC', PRFscore)



# save
saveWorkbook(wb, file = output_name, overwrite = TRUE)
