install.packages("openxlsx")
library(openxlsx)
library(dplyr)
library(tidyr)
library("readxl")

#setwd("C:/NextEra MARL") 
setwd("C:/Users/kelsey.marshall/OneDrive - ERM/Documents/NextEra MARL")

SurveyCorridor <- read_excel("All_02_SSURGO_Soils_20250916_v1.xls", sheet = "ECC_ALL_SSURGO", trim_ws = TRUE, .name_repair = "universal")
SurveyCorridor <- SurveyCorridor %>%
  group_by(State.name, CountyName) %>%
  mutate(
    TotalAcrebyCounty = sum(Acres)) %>%
  ungroup()
colnames(SurveyCorridor)

Impact <- read_excel("All_02_Imp_SSURGO_Soils_20250916_v1.xls", sheet = "IMP_ALL_SSURGO", trim_ws = TRUE, .name_repair = "universal")
Impact <- Impact %>%
  group_by(State.name, CountyName) %>%
  mutate(
    TotalAcrebyCounty = sum(Acres)) %>%
  ungroup()
colnames(Impact)

FarmClassOrder <- c("All areas are prime farmland", "Farmland of statewide importance", "Farmland of local importance", "Not prime farmland")
CorrosionOrder <- c("High", "Moderate", "Low", "")
HydricOrder <- c("Yes", "No", "")
DrainOrder <- c("Somewhat excessively drained", "Well drained", "Moderately well drained", "Somewhat poorly drained", "Poorly drained", "")

SurveyFarmlandClass  <- SurveyCorridor %>%
  group_by(State.name, CountyName, Farmland.Class) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, CountyName, (match(Farmland.Class, FarmClassOrder)))
#SurveyFarmlandClass$Percent <- paste0(SurveyFarmlandClass$Percent, "%")

SurveySteelCorrosion  <- SurveyCorridor %>%
  group_by(State.name, CountyName, Steel.Corrosion) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, CountyName, (match(Steel.Corrosion, CorrosionOrder)))
#SurveySteelCorrosion$Percent <- paste0(SurveySteelCorrosion$Percent, "%")

SurveyConcreteCorrosion  <- SurveyCorridor %>%
  group_by(State.name, CountyName, Concrete.Corrosion) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, CountyName, (match(Concrete.Corrosion, CorrosionOrder)))
#SurveyConcreteCorrosion$Percent <- paste0(SurveyConcreteCorrosion$Percent, "%")

SurveyHydric  <- SurveyCorridor %>%
  group_by(State.name, CountyName, Hydric.Rating) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, CountyName, (match(Hydric.Rating, HydricOrder)))
#SurveyHydric$Percent <- paste0(SurveyHydric$Percent, "%")

SurveyDrainClass  <- SurveyCorridor %>%
  group_by(State.name, CountyName, Drainage.Class...Dominant.Condition) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, CountyName, (match(Drainage.Class...Dominant.Condition, DrainOrder)))
#SurveyDrainClass$Percent <- paste0(SurveyDrainClass$Percent, "%")

ImpactFarmlandClass  <- Impact %>%
  group_by(State.name, CountyName, Farmland.Class) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, CountyName, (match(Farmland.Class, FarmClassOrder)))
#ImpactFarmlandClass$Percent <- paste0(ImpactFarmlandClass$Percent, "%")

ImpactSteelCorrosion  <- Impact %>%
  group_by(State.name, CountyName, Steel.Corrosion) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, CountyName, (match(Steel.Corrosion, CorrosionOrder)))
#ImpactSteelCorrosion$Percent <- paste0(ImpactSteelCorrosion$Percent, "%")

ImpactConcreteCorrosion  <- Impact %>%
  group_by(State.name, CountyName, Concrete.Corrosion) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, CountyName, (match(Concrete.Corrosion, CorrosionOrder)))
#ImpactConcreteCorrosion$Percent <- paste0(ImpactConcreteCorrosion$Percent, "%")

ImpactHydric  <- Impact %>%
  group_by(State.name, CountyName, Hydric.Rating) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, CountyName, (match(Hydric.Rating, HydricOrder)))
#ImpactHydric$Percent <- paste0(ImpactHydric$Percent, "%")

ImpactDrainClass  <- Impact %>%
  group_by(State.name, CountyName, Drainage.Class...Dominant.Condition) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, CountyName, (match(Drainage.Class...Dominant.Condition, DrainOrder)))
#ImpactDrainClass$Percent <- paste0(ImpactDrainClass$Percent, "%")

Tables <- list("ImpactFarmlandClass" = ImpactFarmlandClass, "ImpactSteelCorrosions" = ImpactSteelCorrosion, 
               "ImpactConcreteCorrosion" = ImpactConcreteCorrosion, "ImpactHydric" = ImpactHydric, 
               "ImpactDrainClass" = ImpactDrainClass, "SurveyFarmlandClass" = SurveyFarmlandClass, 
               "SurveySteelCorrosions" = SurveySteelCorrosion, "SurveyConcreteCorrosion" = SurveyConcreteCorrosion, 
               "SurveyHydric" = SurveyHydric, "SurveyDrainClass" = SurveyDrainClass)
write.xlsx(Tables, file = "SoilTables.xlsx")


