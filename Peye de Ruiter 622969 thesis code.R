setwd("C:/Users/Skikk/OneDrive - Erasmus University Rotterdam/Desktop/1 MASTER/1 Thesis/Code")
library(readxl)
library(dplyr) 
library(tidyverse)
library(reshape)
library(Hmisc)
library(ggplot2)
library(stargazer)
library(sandwich)
library (multiwayvcov)
library(tidyr)
library(lubridate)
library(bizdays)
library(tidymodels)
library(glmnet)  
library(leaps)
library(car)
library(skimr)
library(ggridges)
library(vip)
library(pdp)
library(lfe)
library(ggbeeswarm)
library(corrplot)
library(iml)

#-------------------------------------------------------------------------
#Loading datasets 
#-------------------------------------------------------------------------
OEFEN_finaledata22 <- read_excel("inbound_Details DCA v3.xlsx")
#FTE_tasks <- read_excel("FTE_tasks.xlsx")
OEFEN_finaledata <- read_excel("testdata stage peye (archief)-17-05-2025.xlsx")
#unloading_errors <- read_excel("Tickets 20-05-2025.xlsx", sheet = "Losplaatsen DCA")
Administration_errors <- read_excel("Tickets 20-05-2025.xlsx", sheet = "Administratie")
Received_more_items <- read_csv("C:/Users/Skikk/OneDrive - Erasmus University Rotterdam/Desktop/1 MASTER/1 Thesis/Code/te veel ontvangen.csv")
Received_less_items <- read_csv("C:/Users/Skikk/OneDrive - Erasmus University Rotterdam/Desktop/1 MASTER/1 Thesis/Code/deellevering.csv")
NB3_and_NB5_orders <- read_csv("C:/Users/Skikk/OneDrive - Erasmus University Rotterdam/Desktop/1 MASTER/1 Thesis/Code/NB3 orders.csv")


colnames(Received_more_items) <- gsub(" ", "_", colnames(Received_more_items))  
colnames(Received_less_items) <- gsub(" ", "_", colnames(Received_less_items))  
colnames(NB3_and_NB5_orders) <- gsub(" ", "_", colnames(NB3_and_NB5_orders))  

#-------------------------------------------------------------------------
#Cleaning and joining datasets
#-------------------------------------------------------------------------


#MAIN dataset cleaning and variable creation 
OEFEN_finaledata$ONTVANGEN_VLOER[is.na(OEFEN_finaledata$ONTVANGEN_VLOER)] <- 0
OEFEN_finaledata$VERWACHTE_COLLI[is.na(OEFEN_finaledata$VERWACHTE_COLLI)] <- 0
OEFEN_finaledata$ONTVANGEN_MIX[is.na(OEFEN_finaledata$ONTVANGEN_MIX)] <- 0
OEFEN_finaledata$ONTVANGEN_SIP[is.na(OEFEN_finaledata$ONTVANGEN_SIP)] <- 0

OEFEN_finaledata <- OEFEN_finaledata %>%
  mutate(
    ONTVANGEN_MIX = as.numeric(ONTVANGEN_MIX),
    ONTVANGEN_VLOER = as.numeric(ONTVANGEN_VLOER),
    ONTVANGEN_SIP = as.numeric(ONTVANGEN_SIP),
    ONTVANGEN_PALLETS = as.numeric(ONTVANGEN_PALLETS))

EDI_USERS<-c(10280, 24927,6023,14241,41103,30437,7765,16295,8425,778,19349,12136,17301,18564,13595,32607,12583,12815,7286,25619,7179,8961
             ,34884,67777,17236,240,12674,3863,17384,33829,42382,9043,14514,13367,18242,3129,6627,5306,4366,58685,5892,17293,70151,9597,48520,25346
             ,68783,2204,12047,11718,364,4267,11756,17335,29959,18135,35162,23051,35360,24398,83238,34066,10566,50328,36145,10645,6387,8698,18093
             ,28712,10454,18010,5546,28555,10158,11700,51466,497,56796,34025,34678,68809,41947,39024,35873,55376,14407,12096,55160,11411,10010
             ,12190,13706,1776,35568,46425,39354,37705,63412,13049,58883,25403,13904,36038,45070,25809,1164,3491,27037,10209,11597,60368,1156
             ,36087,17814,17707,38133,12278,14811,35410,30825,25049,38240,38919,37853,40022,53744,28175,20941,1867,6932,10769,14074,49155,43091
             ,15099,37887,34611,43935,52365,10298,16642,16717,38273,16691,37937,8466,37945,46672,60541,8847,15537,5009,16782,16709,17111,61275
             ,17581,36616,13417,24661,19868,23077,36137,3921,9282,16865,4523,43463,26963,3269,23580,13156,4317,12682,13102,11139,85381,37093
             ,4739,19867,10171,33811,12919,23952,58974,1107,39313,48215,17095,16139,7146,52795,30387,26526,2394)


#making the dummmy variable
OEFEN_finaledata<-OEFEN_finaledata%>%
  mutate(EDI_USER = ifelse(LEVNR %in% EDI_USERS, 1, 0))

#removing DCE and Cabel doors
sort(unique(OEFEN_finaledata$GEPLANDE_DEUR))

OEFEN_finaledata <- OEFEN_finaledata %>%
  filter(!GEPLANDE_DEUR %in% c("DCE-008","DCE-009","DCE-034","DCE-038","DCE-039","DCE-040","DCE-044","DCE-045","KBL-D35","KBL-D36"))

#MIX SIP control station
sort(unique(OEFEN_finaledata$MENU))
OEFEN_finaledata <- OEFEN_finaledata %>%
  filter(!MENU %in% c("TU - Receiving DCE", "TU - Receiving KBL"))

OEFEN_finaledata<-OEFEN_finaledata%>%
  mutate(SIP_ITEM = ifelse(MENU == "TU - Receiving SIP", 1, 0))
#------------------------------------------------------------------
#Calculating the timedifferences between arrival and other timestamps
#------------------------------------------------------------------

OEFEN_finaledata <- OEFEN_finaledata %>%
  mutate(
    aankomst_truck = as.POSIXct(CHECKIN_DTTM, format="%d-%m-%y %H:%M:%S"),
    vertrek_truck = as.POSIXct(CHECKOUT_DTTM, format="%d-%m-%y %H:%M:%S"),
    inruim = as.POSIXct(INRUIMMOMENT, format="%d-%m-%y %H:%M:%S"),
    aanmeld_ilpn = as.POSIXct(ONTVANGSTMOMENT, format="%d-%m-%y %H:%M:%S"),
    difference_arrival_inruim = as.numeric(difftime(inruim, aankomst_truck, units="hours"))) 

#Accounting for weekends and nights

create.calendar("workdays", weekdays = c("saturday", "sunday"),
                start.date = "2000-01-01", end.date = "2100-12-31")

#removing NAs
OEFEN_finaledata <- OEFEN_finaledata %>%
  filter(!is.na(aankomst_truck), !is.na(inruim)) %>%
  mutate(
    start_date = as_date(aankomst_truck),
    end_date = as_date(inruim),
    workdays_between = bizdays(start_date, end_date, "workdays"),
    datums_between = map2(start_date, end_date, ~ seq(.x, .y, by = "1 day")),
    contains_weekend = map_lgl(datums_between, ~ any(wday(.x) %in% c(1, 7))),
    dock_to_stock_time = difference_arrival_inruim - 8 * pmax(workdays_between, 0) -
      if_else(contains_weekend, 48, 0)
  )





#adding errors
#first getting the ASN numbers
OEFEN_finaledata$asn_number <- sapply(strsplit(as.character(OEFEN_finaledata$TC_ASN_ID), "_"), function(x) x[2])

#names(unloading_errors)[9] <- "Type_issue"
#sort(unique(unloading_errors$Type_issue))

names(Administration_errors)[8] <- "Type_Melding"
sort(unique(Administration_errors$Type_Melding))

#making the erroers wide
errors_wide <- Administration_errors %>%
  mutate(Fout_aanwezig = 1) %>% 
  pivot_wider(
    names_from = Type_Melding,
    values_from = Fout_aanwezig,
    values_fill = list(Fout_aanwezig = 0)  
  )
colnames(errors_wide) <- gsub(" ", "_", colnames(errors_wide)) 

names(errors_wide)[18] <- "Dubbele_EDI_PDF_berichten"


errors_wide_clean1<- errors_wide[,-(1:4)]
errors_wide_clean2<- errors_wide_clean1[,-(2:8)]
#removing obscure errors that happened only a few times and renaming the columns

amount_admin_errors <- sapply(errors_wide_clean2, function(col) sum(col == 1, na.rm = TRUE))
print(amount_admin_errors)

errors_wide_clean2$`Fysieke_pakbon_ontbreekt `<- NULL
errors_wide_clean2$Rechtstreekse_levering<- NULL
errors_wide_clean2$Dubbele_EDI_PDF_berichten<- NULL
errors_wide_clean2$`Purchase_order_voor_DC_Strijen `<- NULL


colnames(errors_wide_clean2)[2] <- "missing_information"
colnames(errors_wide_clean2)[3] <- "missing_ASN_on_packingslip"
colnames(errors_wide_clean2)[4] <- "rejected_in_process_director"

OEFEN_finaledata <- OEFEN_finaledata %>%
  left_join(errors_wide_clean2,by = c("APPOINTMENT_ID" = "Appointmentnummer"))


OEFEN_finaledata[, 36:38][is.na(OEFEN_finaledata[, 36:38])] <- 0



END_DATASET<- OEFEN_finaledata


CLEANING_FOR_MODEL <- END_DATASET %>%
  arrange(STATION, aanmeld_ilpn) %>%
  group_by(STATION) %>%
  mutate(
    vorige_inruim = lag(inruim),
    volgende_inruim = lead(inruim),
    vorige_aanmeld = lag(aanmeld_ilpn),
    volgende_aanmeld = lead(aanmeld_ilpn),
    vorige_asn = lag(asn_number),
    
    #Primary method
    checking_time_raw = as.numeric(difftime(inruim, vorige_inruim, units = "mins")),
    
    #Account for negative time
    checking_time_inruim_vs_prev_aanmeld = as.numeric(difftime(inruim, vorige_aanmeld, units = "mins")),
    break_time = as.numeric(difftime(volgende_aanmeld, aanmeld_ilpn, units = "mins")),
    
    time_for_recheck =as.numeric(difftime(volgende_inruim, inruim, units = "mins")),
    recheck_time  =as.numeric(difftime( inruim, aanmeld_ilpn, units = "mins")),
    
    recheck_case = ifelse(time_for_recheck < 0, 1, 0),
    checking_time = case_when(
      is.na(checking_time_raw) ~ NA_real_,
      checking_time_raw < 0    ~ checking_time_inruim_vs_prev_aanmeld,
      recheck_case == 1     ~ recheck_time,
      TRUE                     ~ checking_time_raw
    ),
    
    switched_asn = as.numeric(asn_number != vorige_asn)
  ) %>%
  ungroup()

CLEANING_FOR_MODEL$checking_time <- round(CLEANING_FOR_MODEL$checking_time, 2)
CLEANING_FOR_MODEL$checking_time_raw <- round(CLEANING_FOR_MODEL$checking_time_raw, 2)
CLEANING_FOR_MODEL$checking_time_inruim_vs_prev_aanmeld <- round(CLEANING_FOR_MODEL$checking_time_inruim_vs_prev_aanmeld, 2)



#adding the extra ASN errors

NB3_orders <- NB3_and_NB5_orders %>%
  filter(Soort == "NB3")

NB5_orders <- NB3_and_NB5_orders %>%
  filter(Soort == "NB5")

NB3_ASN <- NB3_orders$ASN_Nummer_leverancier
NB5_ASN <- NB5_orders$ASN_Nummer_leverancier
Received_less_list<-Received_less_items$ASN_Nummer_leverancier
Received_more_list<-Received_more_items$ASN_Nummer_leverancier

CLEANING_FOR_MODEL <- CLEANING_FOR_MODEL %>%
  mutate(
    NB3_error = as.integer(asn_number %in% NB3_ASN),
    NB5_error = as.integer(asn_number %in% NB5_ASN),
    Received_less_error = as.integer(asn_number %in% Received_less_list),
    Received_more_error = as.integer(asn_number %in% Received_more_list)
  )

#Adding time of day and weekday
CLEANING_FOR_MODEL <- CLEANING_FOR_MODEL %>%
  mutate(
    hour = hour(aanmeld_ilpn), 
    time_of_day = case_when(
      hour >= 5 & hour < 12 ~ "morning",
      hour >= 12 & hour < 17 ~ "noon",
      hour >= 17 | hour < 5 ~ "evening"
    ),
    day_of_week = wday(aanmeld_ilpn, label = TRUE, abbr = FALSE, locale = "C")
  )
CLEANING_FOR_MODEL <- CLEANING_FOR_MODEL %>%
  mutate(
    day_of_week = factor(day_of_week, ordered = FALSE)
  )
CLEANING_FOR_MODEL <- CLEANING_FOR_MODEL %>%
  mutate(
    time_for_break = ifelse(break_time > 45, 1,0)
  )


#CLEANING THE DATA
cleaned_dataset <- CLEANING_FOR_MODEL[,-(1:14)]

columns_to_remove <- c(
  "ILPN",
  "ONTVANGSTMOMENT",
  "INRUIMMOMENT",
  "MENU",
  "AANTAL_ILPNS_APPOINTMENT",
  "aankomst_truck",
  "vertrek_truck",
  "inruim",
  "difference_arrival_inruim",
  "start_date",
  "end_date",
  "workdays_between",
  "datums_between",
  "contains_weekend",
  "Week",
  "ILPNs_that_day",
  "aanmeld_ilpn",
  "vorige_inruim",
  "difference_each_ilpn",
  "Opleiding",
  "asn_number"
)
cleaned_dataset <- cleaned_dataset[, !(names(cleaned_dataset) %in% columns_to_remove)]


#--------------------------------
#Cleaned data
#---------------------------------

#-----------------------
dataset_for_efficiency<- cleaned_dataset
dataset_for_efficiency$dock_to_stock_time <- NULL


#here only relevant work tasks
dataset_for_efficiency$Opleiding <- NULL
dataset_for_efficiency$Leiding <- NULL
dataset_for_efficiency$Intern_transport <- NULL
dataset_for_efficiency$Administratie <- NULL
dataset_for_efficiency$AMAA <- NULL
dataset_for_efficiency$SIP <- NULL
dataset_for_efficiency$Total_FTE <- NULL
dataset_for_efficiency$Controle_werkstation <- NULL
dataset_for_efficiency$Key_user<- NULL
dataset_for_efficiency$Pakketpost <- NULL
dataset_for_efficiency$Perron <- NULL
dataset_for_efficiency$Weekday <- NULL
dataset_for_efficiency$asn_number <- NULL
dataset_for_efficiency$vorige_asn <- NULL
dataset_for_efficiency$rejected_in_process_director <- NULL
dataset_for_efficiency$missing_information <- NULL
dataset_for_efficiency$pallet_too_high <- NULL
dataset_for_efficiency$wrong_carrier <- NULL


dataset_for_efficiency$volgende_aanmeld <- NULL
dataset_for_efficiency$vorige_aanmeld <- NULL
dataset_for_efficiency$volgende_inruim <- NULL

#dataset_for_efficiency$checking_time <- NULL
dataset_for_efficiency$checking_time_inruim <- NULL
dataset_for_efficiency$checking_time_raw <- NULL
dataset_for_efficiency$checking_time_aanmeld_diff <- NULL
dataset_for_efficiency$checking_time_inruim_vs_prev_aanmeld <- NULL
dataset_for_efficiency$checking_time_aanmeld <- NULL
dataset_for_efficiency$hour<- NULL
dataset_for_efficiency$time_for_recheck<- NULL
dataset_for_efficiency$recheck_time<- NULL
dataset_for_efficiency$break_time<- NULL

dataset_for_efficiency$STATION <- NULL

dataset_for_efficiency$LEVNAAM<- as.factor(dataset_for_efficiency$LEVNAAM)
dataset_for_efficiency$day_of_week<- as.factor(dataset_for_efficiency$day_of_week)
dataset_for_efficiency$time_of_day<- as.factor(dataset_for_efficiency$time_of_day)



dataset_for_efficiency <- dataset_for_efficiency %>%
   filter(checking_time <500)%>%
  filter(checking_time >0)

ggplot(dataset_for_efficiency %>% filter(is.finite(checking_time)), 
       aes(x = checking_time)) +
  geom_histogram(binwidth = 15, fill = "steelblue", color = "red") +
  labs(title = "Distribution of Checking Time",
       x = "Checking Time in Minutes",
       y = "Amount of observations") +
  theme_minimal()


#filtering on checkingtime and removing the breaks
dataset_for_efficiency <- dataset_for_efficiency %>%
    filter(checking_time <180)%>%
  filter(checking_time >0) %>%
   filter(time_for_break == 0)



dataset_for_efficiency$time_for_break <- NULL


dataset_for_efficiency <- dataset_for_efficiency %>% drop_na()



#Creating the top suppliers list
top_suppliers_df <- dataset_for_efficiency %>%
  count(LEVNAAM, sort = TRUE)
top_suppliers_df <- top_suppliers_df %>%
  mutate(rank = row_number())

ggplot(top_suppliers_df, aes(x = rank, y = n)) +
  geom_bar(stat = "identity", fill = "steelblue") +
  geom_vline(xintercept = 100.5, linetype = "dashed", color = "red") +
  scale_x_continuous(breaks = top_suppliers_df$rank, labels = top_suppliers_df$LEVNAAM) +
  labs(title = "Number of Observations per Supplier",
       x = "Supplier",
       y = "Number of Observations") +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 90, hjust = 1))


ggplot(top_suppliers_df, aes(x = rank, y = n)) +
  geom_bar(stat = "identity", fill = "steelblue") +
  geom_hline(yintercept = 422,  color = "red") +
  scale_x_continuous(breaks = NULL) +
  labs(title = "Number of Observations per Supplier",
       x = "Supplier",
       y = "Number of Observations") +
  theme_minimal() +
  theme(axis.text.x = element_blank(),
        axis.ticks.x = element_blank())

top_suppliers <- dataset_for_efficiency %>%
  count(LEVNAAM, sort = TRUE) %>%
  slice_head(n = 100) %>%
  pull(LEVNAAM)

supplier_mapping <- setNames(paste0("Supplier ", 1:100), top_suppliers)

top_suppliers
#

dataset_for_efficiency <- dataset_for_efficiency %>%
 mutate(
  supplier = if_else(LEVNAAM %in% top_suppliers, LEVNAAM, "Other"),
 supplier = as.factor(supplier))
  


dataset_for_efficiency$LEVNAAM <- NULL
dataset_for_efficiency$checking_time <- round(dataset_for_efficiency$checking_time, 2)

#Descriptive statistics

skim(dataset_for_efficiency)

numeric_data <- dataset_for_efficiency[sapply(dataset_for_efficiency, is.numeric)]
cor_matrix <- cor(numeric_data, use = "complete.obs")
corrplot(cor_matrix, method = "color", type = "upper", tl.cex = 0.8)






#linear regression models

model_checking_efficiency<- lm(checking_time ~  EDI_USER + missing_ASN_on_packingslip +
                                 switched_asn + NB3_error + NB5_error + 
                                 Received_less_error + Received_more_error  + recheck_case + SIP_ITEM + time_of_day + day_of_week , data = dataset_for_efficiency)
summary(model_checking_efficiency)
vif(model_checking_efficiency)


fixed_suppliers<-felm(checking_time ~  EDI_USER + missing_ASN_on_packingslip +
                        switched_asn + NB3_error + NB5_error + 
                        Received_less_error + Received_more_error  + recheck_case + SIP_ITEM + time_of_day + day_of_week | supplier, data = dataset_for_efficiency)

summary(fixed_suppliers)
residuals <- residuals(fixed_suppliers)
rmse_felm <- sqrt(mean(residuals^2))
rmse_felm


#Holdout fit
set.seed(95)  
dataset_for_efficiency$day_of_week <- factor(dataset_for_efficiency$day_of_week, 
                                             levels = unique(dataset_for_efficiency$day_of_week))

dataset_for_efficiency$time_of_day <- factor(dataset_for_efficiency$time_of_day, 
                                             levels = unique(dataset_for_efficiency$time_of_day))
n <- nrow(dataset_for_efficiency)
train_indices <- sample(seq_len(n), size = 0.7 * n)
train_data <- dataset_for_efficiency[train_indices, ]
test_data  <- dataset_for_efficiency[-train_indices, ]

model_train <- felm(checking_time ~ EDI_USER + missing_ASN_on_packingslip +
                      switched_asn + NB3_error + NB5_error + 
                      Received_less_error + Received_more_error + 
                      recheck_case + SIP_ITEM + time_of_day + day_of_week |
                      supplier,
                    data = train_data)

coefs <- coef(model_train)
form <- as.formula(checking_time ~ EDI_USER + missing_ASN_on_packingslip +
                     switched_asn + NB3_error + NB5_error + 
                     Received_less_error + Received_more_error + 
                     recheck_case + SIP_ITEM + time_of_day + day_of_week)

X_test <- model.matrix(form, data = test_data)
X_test <- X_test[, colnames(X_test) != "(Intercept)"]
coefs <- coef(model_train)
coefs <- coefs[colnames(X_test)]

predictions <- X_test %*% coefs
colnames(X_test)
names(coefs)
predictions <- X_test %*% coefs
actuals <- test_data$checking_time

holdout_rmse <- sqrt(mean((actuals - predictions)^2))

holdout_rsq <- 1 - sum((actuals - predictions)^2) / sum((actuals - mean(actuals))^2)

holdout_rmse
holdout_rsq






stargazer(fixed_suppliers, type = "latex",
          title = "OLS Regression Results Checking Efficiency",
          label = "tab:ols_results checking efficiency",
          style = "default",  
          digits = 3)

supplier_effects <- getfe(fixed_suppliers)


#Plotting the supplier fixed effects
supplier_effects <- supplier_effects %>%
  filter(fe == "supplier") %>%
  arrange(obs) %>%
  mutate(supplier_label = if_else(row_number() <= 100,
                                  paste0("Supplier ", row_number()),
                                  "Other"))


getfe(fixed_suppliers) |>
  filter(fe == "supplier") %>%
  arrange(obs) %>%
  mutate(supplier_label = if_else(row_number() <= 100,
                                  paste0("Supplier ", row_number()),
                                  "Other"))%>%
  ggplot(aes(x = reorder(supplier_label, effect), y = effect)) +
  geom_point() +
  coord_flip() +
  labs(title = "Supplier Fixed Effects on Checking Time",
       y = "Fixed Effect (minutes)",
       x = "Supplier") +
  theme_bw()

fixed_effects_table <- supplier_effects[ ,-(3:5)]


supplier_summary <- dataset_for_efficiency %>%
  group_by(supplier) %>%
  summarise(
    n_obs = n(),    
    avg_SIP_ITEM = mean(SIP_ITEM),
    mean_checking_time = mean(checking_time, na.rm = TRUE),
    recheck_rate = mean(recheck_case, na.rm = TRUE),
    nb3_rate = mean(NB3_error, na.rm = TRUE),
    nb5_rate = mean(NB5_error, na.rm = TRUE),
    received_less_rate = mean(Received_less_error, na.rm = TRUE),
    received_more_rate = mean(Received_more_error, na.rm = TRUE),
    .groups = "drop"
  )

supplier_analysis_df <- left_join(supplier_effects,supplier_summary,by = c("idx" = "supplier"))

top_5_suppliers <- supplier_analysis_df %>%
  arrange(desc(effect)) %>%
  slice_head(n = 5)
bottom_5_suppliers <- supplier_analysis_df %>%
  arrange(effect) %>%        
  slice_head(n = 5) %>%      
  arrange(desc(effect))  

top_bottom_suppliers <- bind_rows(
  top_5_suppliers %>% mutate(type = "Highest Effect"),
  bottom_5_suppliers %>% mutate(type = "Lowest Effect")
)

top_bottom_suppliers <- top_bottom_suppliers[,-c(2,5,3,4,15)]

#Making the suppliers anonymously
supplier_map <- getfe(fixed_suppliers) |>
  filter(fe == "supplier") |>
  arrange(obs) |>
  mutate(supplier_label = if_else(row_number() <= 100,
                                  paste0("Supplier ", row_number()),
                                  "Other")) |>
  select(idx, supplier_label)
dataset_anon <- dataset_for_efficiency %>%
  left_join(supplier_map, by = c("supplier" = "idx"))

dataset_anon$supplier<- NULL
#Rechecks

recheck_model <- glm(recheck_case ~ EDI_USER + SIP_ITEM + 
                       NB3_error + NB5_error + missing_ASN_on_packingslip +
                       Received_less_error + Received_more_error + 
                       time_of_day + day_of_week + supplier_label,
                     family = binomial,
                     data = dataset_anon)

summary(recheck_model)
sort(round(exp(coef(recheck_model)),4), decreasing = TRUE)

stargazer(recheck_model, type = "latex",
          title = " Logistic Regression Estimating the Probability of a Recheck Case",
          label = " Logistic Regression Estimating the Probability of a Recheck Case",
          style = "default",  
          digits = 3)

mean(dataset_anon$EDI_USER)
#--------------------------------------------------------------------------------------------------
#LASSO REGRESSION:
#Splitting the data into training and testing set
set.seed(95)
efficiency_split <- initial_split(dataset_anon, prop = .7)
efficiency_train <- training(efficiency_split)
efficiency_test <- testing(efficiency_split)
cv_folds <- vfold_cv(efficiency_train, v = 10)


#setting up the recipe
linear_reg_recipe_efficiency <-
  recipe(checking_time ~ . , data = dataset_anon) |>
  step_dummy(all_nominal_predictors(), one_hot = TRUE) |>  
  #step_interact(~ all_predictors():all_predictors()) |>
  step_zv(all_predictors()) |>
  step_normalize(all_predictors())

#baking the recipe
efficiency_train_baked <-
  linear_reg_recipe_efficiency |>
  prep(efficiency_train) |>
  bake(efficiency_train)
efficiency_train_baked



efficiency_train_baked |>
  summarise(across(everything(), list(m = mean, s = sd))) |>
  mutate(across(everything(), ~ round(.x, 2)))



lasso_efficiency_linear_reg <-
  linear_reg(penalty = tune(), mixture = 1) |>
  set_engine("glmnet", intercept = FALSE)


lasso_efficiency_linear_reg |> translate()

lasso_efficiency_wf <-
  workflow() |>
  add_recipe(linear_reg_recipe_efficiency) |>
  add_model(lasso_efficiency_linear_reg)

#setting up the grid
grid_lasso_efficiency <-
  grid_regular(penalty(c(-1, 1), trans = log10_trans()),
               levels = 50
  )
#tuning the penalty
lasso_efficiency_tune <-
  lasso_efficiency_wf |>
  tune_grid(
    resamples = cv_folds,
    grid = grid_lasso_efficiency,
    metrics = metric_set(rmse)
  )

lasso_efficiency_tune_metrics <-
  lasso_efficiency_tune |>
  collect_metrics()

lasso_efficiency_tune$.metrics[[1]]
lasso_efficiency_tune_metrics |>
  filter(.metric == "rmse") |>
  mutate(
    ymin = mean - std_err,
    ymax = mean + std_err
  ) |>
  ggplot(aes(x = penalty, y = mean)) +
  geom_pointrange(aes(ymin = ymin, ymax = ymax), alpha = 0.5) +
  scale_x_log10() +
  labs(y = "RMSE", x = expression(lambda)) +
  theme_bw()


lasso_efficiency_tune |>
  autoplot() +
  theme_bw()

lasso_efficiency_tune |>
  show_best(metric = "rmse")

#taking the best model based on 1SE rule
lasso_efficiency_1se_model <-
  lasso_efficiency_tune |>
  select_by_one_std_err(metric = "rmse", desc(penalty))
lasso_efficiency_1se_model

lasso_efficiency_tune |>
  show_best(n = 10, metric = "rmse")

lasso_efficiency_wf_tuned <-
  lasso_efficiency_wf |>
  finalize_workflow(lasso_efficiency_1se_model)
lasso_efficiency_wf_tuned

#applying the mode
lasso_efficiency_last_fit <-
  lasso_efficiency_wf_tuned |>
  last_fit(efficiency_split, metrics = metric_set(rmse, rsq))

lasso_efficiency_test_metrics <-
  lasso_efficiency_last_fit |>
  collect_metrics()
lasso_efficiency_test_metrics

lasso_efficiency_test_metrics <-
  lasso_efficiency_test_metrics |>
  select(.metric, .estimate) |>
  mutate(model = "lasso_efficiency")


lasso_efficiency_last_fit |>
  extract_fit_parsnip() |>
  tidy() |>
  filter(estimate != 0) |>
  arrange(desc(abs(estimate))) |>
  print(n = 100)


#plotting the penalty graph
chosen_penalty <- lasso_efficiency_1se_model %>%
  pull(penalty)


lasso_efficiency_tune_metrics |>
  #collect_metrics() |>
  filter(.metric == "rmse") |>
  ggplot(aes(
    x = penalty, y = mean,
    ymin = mean - std_err, ymax = mean + std_err,
    colour = .config == "Preprocessor1_Model27"
  )) +
  geom_pointrange(alpha = 0.5, show.legend = FALSE) +
  scale_x_log10() +
  labs(y = "RMSE", x = expression(lambda)) +
  scale_colour_manual(values = c("black", "red")) +
  theme_bw()


lasso_metrics <- lasso_efficiency_test_metrics |>
  select(.metric, .estimate) |>
  pivot_wider(names_from = .metric, values_from = .estimate) |>
  mutate(model = "LASSO") |>
  select(model, rmse = rmse, rsq = rsq)
lasso_metrics








#---------------------------------------------
#Random forest
#---------------------------------------------
#splitting the data
random_forest <- dataset_for_efficiency
random_forest$LEVNAAM <- NULL
random_forest$STATION<- NULL

set.seed(95)
rf_checking_split <- initial_split(
  data = random_forest, prop = 0.7,
  strata = checking_time
)

rf_checking_train <- training(rf_checking_split)
rf_checking_test <- testing(rf_checking_split)

set.seed(95)
cv_folds <- rf_checking_train |>
  vfold_cv(v = 10, repeats = 5, strata = "checking_time")

rf_checking_train |>
  select_if(is.numeric) |>
  cor() |>
  corrplot::corrplot()

rf_checking_train |>
  group_by(checking_time)


#Recipe 
rf_recipe <- recipe(checking_time ~ ., data = random_forest)


#Model to were the parameters have to be tuned
rf_model_tune <- rand_forest(
  mtry = tune(),
  min_n = tune(),
  trees =500  
) %>%
  set_mode("regression") %>%
  set_engine("ranger", importance = "permutation")


rf_grid <- grid_regular(
  mtry(range = c(1, 11)),
  min_n(range = c(10, 20)),
  levels = 3  
)


#Creating the workflow 
rf_tune_wf <-
  workflow() |>
  add_recipe(rf_recipe) |>
  add_model(rf_model_tune)

#Regression metrics
reg_metrics <- metric_set(rmse)

rf_tune_grid <- grid_regular(mtry(range = c(1, 11)), levels = 11)
num_cores <- parallel::detectCores()
num_cores

doParallel::registerDoParallel(cores = num_cores - 1L)

#Grid tuning
rf_tune_res <- tune_grid(
  rf_tune_wf,
  resamples = cv_folds,
  grid = rf_grid,
  metrics = reg_metrics,
  control = control_grid(save_pred = TRUE)
)

#tuning for both grid optimized
rf_tune_res |>
  collect_metrics() |>
  filter(.metric == "rmse") |>
  ggplot(aes(x = mtry, y = min_n, fill = mean)) +
  geom_tile() +
  scale_fill_viridis_c() +
  labs(
    title = "RMSE Heatmap by Model Parameters",
    x = "Number of Variables at Split (mtry)",
    y = "Minimum Samples per Node (min_n)",
    fill = "RMSE"
  ) +
  theme_bw()


best_rf <- rf_tune_res %>%
  select_by_one_std_err(metric = "rmse", desc = FALSE)
best_rf

rf_final_wf <- finalize_workflow(rf_tune_wf, best_rf)
rf_final_wf




set.seed(95)
rf_final_fit <- last_fit(
  rf_final_wf,
  split = rf_checking_split,
  metrics = metric_set(rmse, rsq, mae)
)

rf_final_fit |>
  collect_metrics()








#importance scores
rf_final_fit |>
  extract_fit_parsnip() |>
  vip(geom = "point", num_features = 12) +
  theme_bw()

rf_final_fit |>
  extract_fit_parsnip() |>
  vi()

rf_final_fit |>
  extract_fit_parsnip() |>
  vip(num_features = 12)

#SHAP values
rf_model <- rf_final_fit |> extract_fit_parsnip() |> pluck("fit")

rf_recipe_prepped <- rf_final_fit %>%
  extract_recipe() %>%
  prep(training = rf_checking_train)

rf_baked_data <- bake(rf_recipe_prepped, new_data = rf_checking_train)

#Separate X and y for SHAP
X <- rf_baked_data |> select(-checking_time)
y <- rf_baked_data$checking_time

rf_model <- rf_final_fit |> extract_fit_parsnip() |> pluck("fit")

predictor <- Predictor$new(model = rf_model, data = X, y = y)

#Get SHAP values for one observation
shap <- Shapley$new(predictor, x.interest = X[1, ])
plot(shap)

# Loop over a sample of observations
shap_values_all <- lapply(1:50, function(i) {
  shap <- Shapley$new(predictor, x.interest = X[i, ])
  shap$results |> mutate(id = i)
})

shap_df <- bind_rows(shap_values_all)



ggplot(shap_df, aes(x = feature, y = phi)) +
  geom_quasirandom(aes(color = phi > 0), alpha = 0.6) +
  coord_flip() +
  scale_color_manual(values = c("TRUE" = "#009E73", "FALSE" = "#D55E00")) +
  labs(title = "SHAP Value Distribution by Feature",
       x = "Feature", y = "SHAP Value") +
  theme_bw()


rf_final_fit |> collect_metrics()

shap_df |> 
  group_by(feature) |>
  summarise(mean_abs_shap = mean(phi)) |>
  arrange(desc(mean_abs_shap))

rf_checking_train %>%
  filter(recheck_case == 1) %>%
  arrange(desc(checking_time)) %>%
  head()



rf_metrics <- rf_final_fit |>
  collect_metrics() |>
  select(.metric, .estimate) |>
  pivot_wider(names_from = .metric, values_from = .estimate) |>
  mutate(model = "Random Forest") |>
  select(model, rmse = rmse, rsq = rsq)
rf_metrics

#---------------------------------------------
#Best practice analysis
#---------------------------------------------
dataset_for_best<-dataset_anon%>%
  filter(recheck_case == 0, SIP_ITEM == 0)

best_practice_df<-dataset_for_best%>%
  filter(recheck_case == 0, SIP_ITEM == 0)


best_practice_df <- best_practice_df %>%
  left_join(supplier_effects, by = c("supplier_label" = "supplier_label"))


best_practice_data <- best_practice_df
best_practice_data$EDI_USER <- 1
best_practice_data$missing_ASN_on_packingslip <- 0
best_practice_data$switched_asn <- 0
best_practice_data$NB3_error <- 0
best_practice_data$NB5_error <- 0
best_practice_data$Received_less_error <- 0
best_practice_data$Received_more_error <- 0
 

colnames(best_practice_data)[13] <- "supplier"
best_practice_data$obs<-NULL
best_practice_data$comp<-NULL
best_practice_data$fe<-NULL
best_practice_data$idx<-NULL

cols_to_fix <- c("EDI_USER", "missing_ASN_on_packingslip", "switched_asn",
                 "NB3_error", "NB5_error", "Received_less_error", "Received_more_error",
                 "recheck_case", "SIP_ITEM")


best_practice_data$day_of_week <- factor(best_practice_data$day_of_week, levels = levels(dataset_for_best$day_of_week))
best_practice_data$time_of_day <- factor(best_practice_data$time_of_day, levels = levels(dataset_for_best$time_of_day))
best_practice_data[cols_to_fix] <- lapply(best_practice_data[cols_to_fix], function(x) as.numeric(as.character(x)))

#setting up the formula
model_formula <- as.formula("~ -1 +EDI_USER + missing_ASN_on_packingslip +
                             switched_asn + NB3_error + NB5_error + 
                             Received_less_error + Received_more_error  + 
                             recheck_case + SIP_ITEM + time_of_day + day_of_week")

#Taking the best coefficents
X_best <- model.matrix(model_formula, data = best_practice_data)
coefs_best <- coef(fixed_suppliers)

setdiff(colnames(X_best), names(coefs_best))
X_best <- X_best[, colnames(X_best) %in% names(coefs_best)]

colnames(X_best)


best_practice_data$predicted_linear <- as.numeric(X_best %*% coefs_best)

best_practice_data$supplier_fe <- ifelse(is.na(best_practice_data$effect), 0, best_practice_data$effect)
best_practice_data$predicted_best <- best_practice_data$predicted_linear + best_practice_data$supplier_fe


dataset_for_best[cols_to_fix] <- lapply(dataset_for_best[cols_to_fix], function(x) as.numeric(as.character(x)))

#Matching factors
dataset_for_best$day_of_week <- factor(dataset_for_best$day_of_week, levels = levels(best_practice_data$day_of_week))
dataset_for_best$time_of_day <- factor(dataset_for_best$time_of_day, levels = levels(best_practice_data$time_of_day))

#Building model matrix
X_actual <- model.matrix(model_formula, data = dataset_for_best)

#Matching columns to coefficients
X_actual <- X_actual[, colnames(X_actual) %in% names(coefs_best)]

dataset_for_best$predicted_linear <- as.numeric(X_actual %*% coefs)

dataset_for_best <- dataset_for_best %>%
  left_join(supplier_effects, by = c("supplier_label" = "supplier_label"))

dataset_for_best$supplier_fe <- ifelse(is.na(dataset_for_best$effect), 0, dataset_for_best$effect)

dataset_for_best$predicted_actual <- dataset_for_best$predicted_linear + dataset_for_best$supplier_fe
dataset_for_best$gain <- dataset_for_best$predicted_actual - best_practice_data$predicted_best
mean(dataset_for_best$gain)  
sum(dataset_for_best$gain)   
