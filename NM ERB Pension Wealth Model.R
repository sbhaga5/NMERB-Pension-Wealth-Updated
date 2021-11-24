rm(list = ls())
library("readxl")
library(tidyverse)
library(zoo)
library(reshape2)
setwd(getwd())
FileName <- 'Model Inputs.xlsx'

#Update this every year
BaseYear <- 2021

#Assigning Variables
model_inputs <- read_excel(FileName, sheet = 'Main')
MeritIncreases <- read_excel(FileName, sheet = 'Merit Increases')
MeritIncreases <- MeritIncreases[,2]
for(i in 1:nrow(model_inputs)){
  if(!is.na(model_inputs[i,2])){
    assign(as.character(model_inputs[i,2]),as.double(model_inputs[i,3]))
  }
}

#These rates dont change so they're outside the function
#Mortality Rates
#2036 is the Age for the MP-2020 rates
MaleMortality <- read_excel(FileName, sheet = 'MP-2020_Male') %>% select(Age,'2036')
FemaleMortality <- read_excel(FileName, sheet = 'MP-2020_Female') %>% select(Age,'2036')
SurvivalRates <- read_excel(FileName, sheet = 'Mortality Rates')

#Function for determining retirement eligibility (including normal retirement, unreduced early retirement, and reduced early retirement)
IsRetirementEligible <- function(Age, YOS){
  Check = ifelse((Age >= NormalRetAgeI & YOS >= Vesting) |
                   YOS >= 30 |
                   #(Age >= NormalRetAgeII & YOS >= NormalYOSII) | 
                   (Age + YOS) >= ReduceRetAge, TRUE, FALSE)
  return(Check)
}

#Separation Rates
SeparationRates <- expand_grid(25:80,0:55)
colnames(SeparationRates) <- c('Age','YOS')
SeparationRates$SepProb <- 0

SeparationRatesBefore <- read_excel(FileName, sheet = 'Separation Rates Before')
MaleSeparationReduced <- read_excel(FileName, sheet = 'Male Separation Reduced')
FemaleSeparationReduced <- read_excel(FileName, sheet = 'Female Separation Reduced')
#The final rates are needed for YOS > 30 and Age > 70
FinalRatesReduced <- read_excel(FileName, sheet = 'Final Separation Rates Reduced')

SeparationRates <- left_join(SeparationRates,SeparationRatesBefore, by = 'YOS')
SeparationRates <- left_join(SeparationRates,MaleSeparationReduced,by = c('YOS','Age'))
SeparationRates <- left_join(SeparationRates,FemaleSeparationReduced,by = c('YOS','Age'))
SeparationRates <- left_join(SeparationRates,FinalRatesReduced,by = 'Age')
#Remove NAs
SeparationRates[is.na(SeparationRates)] <- 0

GetVestedBalanceData <- function(Hiring_Age = HiringAge, 
                          Starting_Salary = StartingSalary,
                          COLA_Plan = COLA,
                          DC_ReturnScenario = ReturnScenario,
                          DB_EE_Contrib_Rate = DBEEContrib_Rate,
                          DB_ER_Contrib_Rate = DBERContrib_Rate,
                          DC_EE_Contrib_Rate = DCEEContrib_Rate,
                          DC_ER_Contrib_Rate = DCERContrib_Rate,
                          Grad_Mult = GradMult,
                          Ben_Mult1 = BenMult1,
                          Ben_Mult2 = BenMult2,
                          Ben_Mult3 = BenMult3,
                          Ben_Mult4 = BenMult4){
  
  #Expand grid for ages 25-120 and years 2009 to 2019
  SurvivalMale <- expand_grid(25:120,2009:2129)
  colnames(SurvivalMale) <- c('Age','Years')
  SurvivalMale$Value <- 0
  #Join these tables to make the calculations easier
  SurvivalMale <- left_join(SurvivalMale,SurvivalRates,by = 'Age')
  SurvivalMale <- left_join(SurvivalMale,MaleMortality,by = 'Age') %>% group_by(Age) %>% 
    #MPValue2 is the cumulative product of the MP-2020 value for year 2036. We use it later so make life easy and calculate now
    mutate(MPValue2 = cumprod(1-lag(`2036`,2,default = 0)), YOS = Age - Hiring_Age,
           Value = ifelse(Age == 120, 1,
                          ifelse(!IsRetirementEligible(Age,YOS) & Years == 2010, PubG_2010_employee_male,
                                 ifelse(IsRetirementEligible(Age,YOS) & Years == 2010, PubG_2010_healthy_retiree_male,
                                        ifelse(!IsRetirementEligible(Age,YOS) & Years > 2010, PubG_2010_employee_male * MPValue2,
                                               ifelse(IsRetirementEligible(Age,YOS) & Years > 2010, PubG_2010_healthy_retiree_male * MPValue2, 0))))))
  #filter out the necessary variables
  SurvivalMale <- SurvivalMale %>% select('Age','Years','Value') %>% ungroup()
  
  #Expand grid for ages 25-120 and years 2009 to 2019
  SurvivalFemale <- expand_grid(25:120,2009:2129)
  colnames(SurvivalFemale) <- c('Age','Years')
  SurvivalFemale$Value <- 0
  #Join these tables to make the calculations easier
  SurvivalFemale <- left_join(SurvivalFemale,SurvivalRates,by = 'Age')
  SurvivalFemale <- left_join(SurvivalFemale,FemaleMortality,by = 'Age') %>% group_by(Age) %>% 
    #MPValue2 is the cumulative product of the MP-2020 value for year 2036. We use it later so make life easy and calculate now
    mutate(MPValue2 = cumprod(1-lag(`2036`,2,default = 0)), YOS = Age - Hiring_Age,
           Value = ifelse(Age == 120, 1,
                          ifelse(!IsRetirementEligible(Age,YOS) & Years == 2010, PubG_2010_employee_female,
                                 ifelse(IsRetirementEligible(Age,YOS) & Years == 2010, PubG_2010_healthy_retiree_female,
                                        ifelse(!IsRetirementEligible(Age,YOS) & Years > 2010, PubG_2010_employee_female * MPValue2,
                                               ifelse(IsRetirementEligible(Age,YOS) & Years > 2010, PubG_2010_healthy_retiree_female * MPValue2, 0))))))
  #filter out the necessary variables
  SurvivalFemale <- SurvivalFemale %>% select('Age','Years','Value') %>% ungroup()
  
  #Calculate Separation Rates
  SeparationRates <- SeparationRates %>% filter(Age - YOS == Hiring_Age) %>%
    mutate(male_sep_red_ben = ifelse(YOS > 30 | Age > 70, final_male_sep_red_ben, male_sep_red_ben),
           female_sep_red_ben = ifelse(YOS > 30 | Age > 70, final_female_sep_red_ben, female_sep_red_ben),
           SepProb = (pmax(male_sep_before_age, male_sep_red_ben) 
                      + pmax(female_sep_before_age, female_sep_red_ben))/2)
  #Filter out unecessary fields
  SeparationRates <- SeparationRates %>% select('Age','YOS','SepProb')
  
  #Change the sequence for the Age and YOS depending on the hiring age
  Age <- seq(Hiring_Age,120)
  YOS <- seq(0,(95-(Hiring_Age-25)))
  #Merit Increases need to have the same length as Age when the hiring age changes
  MeritIncreases <- MeritIncreases[1:length(Age),]
  TotalSalaryGrowth <- as.matrix(MeritIncreases) + salary_growth
  
  #Salary increases and other
  SalaryData <- tibble(Age,YOS) %>%
    mutate(Salary = Starting_Salary*cumprod(1+lag(TotalSalaryGrowth,default = 0)),
           IRSSalaryCap = pmin(Salary,IRSCompLimit),
           FinalAvgSalary = ifelse(YOS >= Vesting, rollmean(lag(Salary), k = FinAvgSalaryYears, fill = 0, align = "right"), 0),
           DB_EE_Contrib = DB_EE_Contrib_Rate*Salary, DB_ER_Contrib = DB_ER_Contrib_Rate*Salary,
           DC_EE_Contrib = DC_EE_Contrib_Rate*Salary, DC_ER_Contrib = DC_ER_Contrib_Rate*Salary,
           DB_ERVested = pmin(pmax(0+EnhER5*(YOS>=5)+EnhER610*(YOS-5)),1))
  
  #Because Credit Interest != inflation, you cant use NPV formulae for DB Balance
  for(i in 1:nrow(SalaryData)){
    if(SalaryData$YOS[i] == 0){
      SalaryData$DBEEBalance[i] <- 0
      SalaryData$DBERBalance[i] <- 0
      SalaryData$CumulativeWage[i] <- 0
      
      SalaryData$DCEEBalance[i] <- 0
      SalaryData$DCERBalance[i] <- 0
    } else {
      SalaryData$DBEEBalance[i] <- SalaryData$DBEEBalance[i-1]*(1+Interest) + SalaryData$DB_EE_Contrib[i-1]
      SalaryData$DBERBalance[i] <- SalaryData$DBERBalance[i-1]*(1+Interest) + SalaryData$DB_ER_Contrib[i-1]
      SalaryData$CumulativeWage[i] <- SalaryData$CumulativeWage[i-1]*(1+ARR) + SalaryData$Salary[i-1]
      
      SalaryData$DCEEBalance[i] <- SalaryData$DCEEBalance[i-1]*(1+DC_ReturnScenario) + SalaryData$DC_EE_Contrib[i-1]
      SalaryData$DCERBalance[i] <- SalaryData$DCERBalance[i-1]*(1+DC_ReturnScenario) + SalaryData$DC_ER_Contrib[i-1]
    }
  }
  
  #Adjusted Mortality Rates
  #The adjusted values follow the same difference between Year and Hiring Age. So if you start in 2020 and are 25, then 26 at 2021, etc.
  #The difference is always 1995. After this remove all unneccessary values
  #BaseYear here is 2020 but can be changed if need be
  AdjustedMale <- SurvivalMale %>% filter(Age - (Years - BaseYear) == Hiring_Age, Years >= BaseYear)
  AdjustedFemale <- SurvivalFemale %>% filter(Age - (Years - BaseYear) == Hiring_Age, Years >= BaseYear)
  
  #Rename columns for merging and taking average
  colnames(AdjustedMale) <- c('Age','Years','AdjMale')
  colnames(AdjustedFemale) <- c('Age','Years','AdjFemale')
  AdjustedValues <- left_join(AdjustedMale,AdjustedFemale) %>% mutate(AdjValue = (AdjMale + AdjFemale)/2)
  
  #Survival Probability and Annuity Factor
  AnnFactorData <- AdjustedValues %>% select(Age,AdjValue) %>%
    mutate(Prob = cumprod(1 - lag(AdjValue, default = 0)),
           DiscProb = Prob / (1+ARR)^(Age - Hiring_Age),
           surv_DR_COLA_Plan = DiscProb * (1+COLA_Plan)^(Age-Hiring_Age),
           AnnuityFactor = rev(cumsum(rev(surv_DR_COLA_Plan))) / surv_DR_COLA_Plan)
  
  #Retention Rates
  AFNormalRetAgeII <- AnnFactorData$AnnuityFactor[AnnFactorData$Age == NormalRetAgeII]
  SurvProbNormalRetAgeII <- AnnFactorData$Prob[AnnFactorData$Age  == NormalRetAgeII]
  RetentionRates_Inputs <- read_excel(FileName, sheet = 'Retention Rates')
  RetentionRates <- expand_grid(Hiring_Age:120,5:60)
  colnames(RetentionRates) <- c('Age','YOS')
  RetentionRates <- left_join(RetentionRates,RetentionRates_Inputs, by = 'Age') 
  RetentionRates <-  left_join(RetentionRates,AnnFactorData %>% select(Age,Prob,AnnuityFactor), by = 'Age') %>%
    mutate(AF = AnnuityFactor, SurvProb = Prob,
           #Replacement rate depending on the different age conditions
           RepRate = ifelse((Age >= NormalRetAgeI & YOS >= Vesting) |      
                              (Age >= NormalRetAgeII & YOS >= NormalYOSII) | 
                              ((Age + YOS) >= ReduceRetAge & Age >= 65), 1,
                            ifelse(Age < NormalRetAgeII & YOS >= NormalYOSII,
                                   AFNormalRetAgeII / (1+ARR)^(NormalRetAgeII - Age)*SurvProbNormalRetAgeII / SurvProb / AF,
                                   ifelse((Age + YOS) >= ReduceRetAge, Factor, 0))))
  #Rename this column so we can join it to the benefit table later
  colnames(RetentionRates)[1] <- 'Retirement Age'
  RetentionRates <- RetentionRates %>% select(`Retirement Age`, YOS, RepRate)
  
  #Benefits, Annuity Factor and Present Value for ages 45-120
  BenefitsTable <- expand_grid(Hiring_Age:120,45:120)
  colnames(BenefitsTable) <- c('Age','Retirement Age')
  BenefitsTable <- left_join(BenefitsTable,SalaryData, by = "Age") %>% 
                   left_join(RetentionRates, by = c("Retirement Age", "YOS")) %>% 
                   left_join(AnnFactorData %>%  select(Age, Prob, AnnuityFactor), by = c("Retirement Age" = "Age")) %>%
                   #Prob and AF at retirement
                   rename(Prob_Ret = Prob, AF_Ret = AnnuityFactor) %>% 
                   #Rejoin the table to get the regular AF and Prob
                   left_join(AnnFactorData %>% select(Age, Prob), by = c("Age"))
  
  BenefitsTable <- BenefitsTable %>% 
    mutate(GradedMult = Ben_Mult1*pmin(YOS,10) + Ben_Mult2*pmax(pmin(YOS,20)-10,0) + Ben_Mult3*pmax(pmin(YOS,30)-20,0) + Ben_Mult4*pmax(YOS-30,0),
           RepRateMult = case_when(Grad_Mult == 0 ~ RepRate*Ben_Mult2*YOS, TRUE ~ RepRate*GradedMult),
           AnnFactorAdj = (AF_Ret / ((1+ARR)^(`Retirement Age` - Age)))*(Prob_Ret / Prob),
           PensionBenefit = RepRateMult*FinalAvgSalary,
           PresentValue = ifelse(Age > `Retirement Age`,0,PensionBenefit*AnnFactorAdj))
  
  #The max benefit is done outside the table because it will be merged with Salary data
  OptimumBenefit <- BenefitsTable %>% group_by(Age) %>% summarise(MaxBenefit = max(PresentValue))
  SalaryData <- left_join(SalaryData,OptimumBenefit) 
  SalaryData <- left_join(SalaryData,SeparationRates,by = c('Age','YOS')) %>%
    mutate(PenWealth = ifelse(YOS < Vesting,(DBERBalance*DB_ERVested)+DBEEBalance,pmax((DBERBalance*DB_ERVested)+DBEEBalance,MaxBenefit)),
           PVPenWealth = PenWealth/(1+ARR)^(Age-Hiring_Age),
           DCEEBalance_Infl = DCEEBalance/(1+assum_infl)^(Age-HiringAge),
           DCERBalance_Infl = DCERBalance/(1+assum_infl)^(Age-HiringAge),
           VestedDCBalance = DCEEBalance_Infl + DCERBalance_Infl,
           HybridBalance = VestedDCBalance + PVPenWealth,
           #Fiter out the NAs because when you change the Hiring age, there are NAs in separation probability,
           #or in Pension wealth
           PVCumWage = CumulativeWage/(1+ARR)^(Age-Hiring_Age)) %>% filter(!is.na(PVPenWealth),!is.na(SepProb))
  
  return(SalaryData)
}

SalaryHeadcountData <- read_excel(FileName, sheet = 'Salary and Headcount')
GetNormalCostFinal <- function(SalaryHeadcountData){
  #This part requires a for loop since GetNormalCost cant be vectorized.
  for(i in 1:nrow(SalaryHeadcountData)){
    VestedBalanceData <- GetVestedBalanceData(Hiring_Age = SalaryHeadcountData$Hiring_Age[i], 
                                              Starting_Salary = SalaryHeadcountData$Starting_Salary[i])
    #Calc and return Normal Cost
    SalaryHeadcountData$NormalCost[i] <- sum(VestedBalanceData$SepProb*VestedBalanceData$PVPenWealth) / 
                                         sum(VestedBalanceData$SepProb*VestedBalanceData$PVCumWage)
    
  }
  
  #Calc the weighted average Normal Cost
  NormalCostFinal <- sum(SalaryHeadcountData$Average_Salary*SalaryHeadcountData$Headcount_Total*SalaryHeadcountData$NormalCost) /
                     sum(SalaryHeadcountData$Average_Salary*SalaryHeadcountData$Headcount_Total)
  
  return(NormalCostFinal)
}
GetNormalCostFinal(SalaryHeadcountData)

GetPlanBalance <- function(PlanType){
  #Can add more parameters later
  VestedBalanceData <- GetVestedBalanceData(Hiring_Age = 25)
  
  if(PlanType == 'DB'){
    return(VestedBalanceData$PVPenWealth)
  } else if(PlanType == 'DC'){
    return(VestedBalanceData$VestedDCBalance)
  } else if(PlanType == 'Hybrid'){
    return(VestedBalanceData$HybridBalance)
  }
}
GetPlanBalance('Hybrid')
