install.packages("openxlsx")
library(openxlsx)
library(dplyr)
library(tidyr)
library("readxl")

setwd("C:/NextEra MARL")

SurveyCorridor <- read_excel("All_02_SSURGO_Soils_20250916_v1.xls", sheet = "ECC_ALL_SSURGO", trim_ws = TRUE, .name_repair = "universal")
SurveyCorridor <- SurveyCorridor %>%
  group_by(State.name) %>%
  mutate(
    TotalAcrebyState = sum(Acres)) %>%
  ungroup()
colnames(SurveyCorridor)

Impact <- read_excel("All_02_Imp_SSURGO_Soils_20250916_v1.xls", sheet = "IMP_ALL_SSURGO", trim_ws = TRUE, .name_repair = "universal")
Impact <- Impact %>%
  group_by(State.name) %>%
  mutate(
    TotalAcrebyState = sum(Acres)) %>%
  ungroup()
colnames(Impact)

SurveyFarmlandClass  <- SurveyCorridor %>%
  group_by(State.name, Farmland.Class) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyState))*100, 1))
SurveyFarmlandClass$Percent <- paste0(SurveyFarmlandClass$Percent, "%")

SurveySteelCorrosion  <- SurveyCorridor %>%
  group_by(State.name, Steel.Corrosion) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyState))*100, 1))
SurveySteelCorrosion$Percent <- paste0(SurveySteelCorrosion$Percent, "%")

SurveyConcreteCorrosion  <- SurveyCorridor %>%
  group_by(State.name, Concrete.Corrosion) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyState))*100, 1))
SurveyConcreteCorrosion$Percent <- paste0(SurveyConcreteCorrosion$Percent, "%")

SurveyHydric  <- SurveyCorridor %>%
  group_by(State.name, Hydric.Rating) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyState))*100, 1))
SurveyHydric$Percent <- paste0(SurveyHydric$Percent, "%")

SurveyDrainClass  <- SurveyCorridor %>%
  group_by(State.name, Drainage.Class...Dominant.Condition) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyState))*100, 1))
SurveyDrainClass$Percent <- paste0(SurveyDrainClass$Percent, "%")

ImpactFarmlandClass  <- Impact %>%
  group_by(State.name, Farmland.Class) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyState))*100, 1))
ImpactFarmlandClass$Percent <- paste0(ImpactFarmlandClass$Percent, "%")

ImpactSteelCorrosion  <- Impact %>%
  group_by(State.name, Steel.Corrosion) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyState))*100, 1))
ImpactSteelCorrosion$Percent <- paste0(ImpactSteelCorrosion$Percent, "%")

ImpactConcreteCorrosion  <- Impact %>%
  group_by(State.name, Concrete.Corrosion) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyState))*100, 1))
ImpactConcreteCorrosion$Percent <- paste0(ImpactConcreteCorrosion$Percent, "%")

ImpactHydric  <- Impact %>%
  group_by(State.name, Hydric.Rating) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyState))*100, 1))
ImpactHydric$Percent <- paste0(ImpactHydric$Percent, "%")

ImpactDrainClass  <- Impact %>%
  group_by(State.name, Drainage.Class...Dominant.Condition) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyState))*100, 1))
ImpactDrainClass$Percent <- paste0(ImpactDrainClass$Percent, "%")

Tables <- list("ImpactFarmlandClass" = ImpactFarmlandClass, "ImpactSteelCorrosions" = ImpactSteelCorrosion, 
               "ImpactConcreteCorrosion" = ImpactConcreteCorrosion, "ImpactHydric" = ImpactHydric, 
               "ImpactDrainClass" = ImpactDrainClass, "SurveyFarmlandClass" = SurveyFarmlandClass, 
               "SurveySteelCorrosions" = SurveySteelCorrosion, "SurveyConcreteCorrosion" = SurveyConcreteCorrosion, 
               "SurveyHydric" = SurveyHydric, "SurveyDrainClass" = SurveyDrainClass)
write.xlsx(Tables, file = "SoilTables.xlsx")

