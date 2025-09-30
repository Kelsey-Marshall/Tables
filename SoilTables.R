install.packages("openxlsx")
library(openxlsx)
library(dplyr)
library(tidyr)
library("readxl")

#setwd("C:/NextEra MARL") 
setwd("C:/Users/kelsey.marshall/OneDrive - ERM/Documents/NextEra MARL")

SurveyCorridor <- read_excel("All_02_SSURGO_Soils_20250916_v1.xls", sheet = "ECC_ALL_SSURGO", trim_ws = TRUE, .name_repair = "universal")
SurveyCorridor <- SurveyCorridor %>%
  group_by(State.name, RouteName, CountyName) %>%
  mutate(
    TotalAcrebyCounty = sum(Acres)) %>%
  ungroup()
colnames(SurveyCorridor)

Impact <- read_excel("All_02_Imp_SSURGO_Soils_20250916_v1.xls", sheet = "IMP_ALL_SSURGO", trim_ws = TRUE, .name_repair = "universal")
Impact <- Impact %>%
  group_by(State.name, Name, CountyName) %>%
  mutate(
    TotalAcrebyCounty = sum(Acres)) %>%
  ungroup()
colnames(Impact)

FarmClassOrder <- c("All areas are prime farmland", "Farmland of statewide importance", "Farmland of local importance", "Not prime farmland")
CorrosionOrder <- c("High", "Moderate", "Low", "")
HydricOrder <- c("Yes", "No", "")
DrainOrder <- c("Somewhat excessively drained", "Well drained", "Moderately well drained", "Somewhat poorly drained", "Poorly drained", "Very poorly drained", "")
WVCountyAcrePercSurvey <- c("State.name", "RouteName", "Farmland.Class", "Acres_Hampshire County", "Percent_Hampshire County", "Acres_Mineral County", "Percent_Mineral County", 
                      "Acres_Monongalia County", "Percent_Monongalia County", "Acres_Preston County", "Percent_Preston County")
PACountyAcrePercSurvey <- c("State.name", "RouteName", "Farmland.Class", "Acres_Fayette County", "Percent_Fayette County", "Acres_Greene County", 
                            "Percent_Greene County")

#####farm#####
SurveyFarmlandClassWV  <- SurveyCorridor %>%
  group_by(State.name, RouteName, CountyName, Farmland.Class) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, RouteName, CountyName, (match(Farmland.Class, FarmClassOrder))) %>%
  filter(State.name == "West Virginia") %>%
  group_by(RouteName, Farmland.Class) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(RouteName, (match(Farmland.Class, FarmClassOrder))) %>%
  select("State.name", "RouteName", "Farmland.Class", "Acres_Hampshire County", "Percent_Hampshire County", "Acres_Mineral County", "Percent_Mineral County", 
         "Acres_Monongalia County", "Percent_Monongalia County", "Acres_Preston County", "Percent_Preston County")

ImpactFarmlandClassWV<- Impact %>%
  group_by(State.name, Name, CountyName, Farmland.Class) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, Name, CountyName, (match(Farmland.Class, FarmClassOrder))) %>%
  filter(State.name == "West Virginia") %>%
  group_by(Name, Farmland.Class) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(Name, (match(Farmland.Class, FarmClassOrder))) %>%
  select("State.name", "Name", "Farmland.Class", "Acres_Hampshire County", "Percent_Hampshire County", "Acres_Mineral County", "Percent_Mineral County", 
         "Acres_Monongalia County", "Percent_Monongalia County", "Acres_Preston County", "Percent_Preston County")

SurveyFarmlandClassPA  <- SurveyCorridor %>%
  group_by(State.name, RouteName, CountyName, Farmland.Class) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, RouteName, CountyName, (match(Farmland.Class, FarmClassOrder))) %>%
  filter(State.name == "Pennsylvania") %>%
  group_by(RouteName, Farmland.Class) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(RouteName, (match(Farmland.Class, FarmClassOrder))) %>%
  select("State.name", "RouteName", "Farmland.Class", "Acres_Fayette County", "Percent_Fayette County", "Acres_Greene County", "Percent_Greene County")
#SurveyFarmlandClass$Percent <- paste0(SurveyFarmlandClass$Percent, "%")

ImpactFarmlandClassPA<- Impact %>%
  group_by(State.name, Name, CountyName, Farmland.Class) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, Name, CountyName, (match(Farmland.Class, FarmClassOrder))) %>%
  filter(State.name == "Pennsylvania") %>%
  group_by(Name, Farmland.Class) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(Name, (match(Farmland.Class, FarmClassOrder))) %>%
  select("State.name", "Name", "Farmland.Class", "Acres_Fayette County", "Percent_Fayette County", "Acres_Greene County", "Percent_Greene County")
#ImpactFarmlandClass$Percent <- paste0(ImpactFarmlandClass$Percent, "%")

#####steel#####
SurveySteelCorrosionWV  <- SurveyCorridor %>%
  group_by(State.name, RouteName, CountyName, Steel.Corrosion) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, RouteName, CountyName, (match(Steel.Corrosion, CorrosionOrder))) %>%
  filter(State.name == "West Virginia") %>%
  group_by(RouteName, Steel.Corrosion) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(RouteName, (match(Steel.Corrosion, CorrosionOrder))) %>%
  select("State.name", "RouteName", "Steel.Corrosion", "Acres_Hampshire County", "Percent_Hampshire County", "Acres_Mineral County", "Percent_Mineral County", 
         "Acres_Monongalia County", "Percent_Monongalia County", "Acres_Preston County", "Percent_Preston County")

ImpactSteelCorrosionWV  <- Impact %>%
  group_by(State.name, Name, CountyName, Steel.Corrosion) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, CountyName, (match(Steel.Corrosion, CorrosionOrder))) %>%
  filter(State.name == "West Virginia") %>%
  group_by(Name, Steel.Corrosion) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(Name, (match(Steel.Corrosion, CorrosionOrder))) %>%
  select("State.name", "Name", "Steel.Corrosion", "Acres_Hampshire County", "Percent_Hampshire County", "Acres_Mineral County", "Percent_Mineral County", 
         "Acres_Monongalia County", "Percent_Monongalia County", "Acres_Preston County", "Percent_Preston County")

SurveySteelCorrosionPA  <- SurveyCorridor %>%
  group_by(State.name, RouteName, CountyName, Steel.Corrosion) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, RouteName, CountyName, (match(Steel.Corrosion, CorrosionOrder))) %>%
  filter(State.name == "Pennsylvania") %>%
  group_by(RouteName, Steel.Corrosion) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(RouteName, (match(Steel.Corrosion, CorrosionOrder)))  %>%
  select("State.name", "RouteName", "Steel.Corrosion", "Acres_Fayette County", "Percent_Fayette County", "Acres_Greene County", "Percent_Greene County")


ImpactSteelCorrosionPA  <- Impact %>%
  group_by(State.name, Name, CountyName, Steel.Corrosion) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, CountyName, (match(Steel.Corrosion, CorrosionOrder))) %>%
  filter(State.name == "Pennsylvania") %>%
  group_by(Name, Steel.Corrosion) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(Name, (match(Steel.Corrosion, CorrosionOrder))) %>%
  select("State.name", "Name", "Steel.Corrosion", "Acres_Fayette County", "Percent_Fayette County", "Acres_Greene County", "Percent_Greene County")

#####concrete#####
SurveyConcreteCorrosionWV  <- SurveyCorridor %>%
  group_by(State.name, RouteName, CountyName, Concrete.Corrosion) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, RouteName, CountyName, (match(Concrete.Corrosion, CorrosionOrder))) %>%
  filter(State.name == "West Virginia") %>%
  group_by(RouteName, Concrete.Corrosion) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(RouteName, (match(Concrete.Corrosion, CorrosionOrder))) %>%
  select("State.name", "RouteName", "Concrete.Corrosion", "Acres_Hampshire County", "Percent_Hampshire County", "Acres_Mineral County", "Percent_Mineral County", 
         "Acres_Monongalia County", "Percent_Monongalia County", "Acres_Preston County", "Percent_Preston County")

ImpactConcreteCorrosionWV  <- Impact %>%
  group_by(State.name, Name, CountyName, Concrete.Corrosion) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, CountyName, (match(Concrete.Corrosion, CorrosionOrder))) %>%
  filter(State.name == "West Virginia") %>%
  group_by(Name, Concrete.Corrosion) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(Name, (match(Concrete.Corrosion, CorrosionOrder))) %>%
  select("State.name", "Name", "Concrete.Corrosion", "Acres_Hampshire County", "Percent_Hampshire County", "Acres_Mineral County", "Percent_Mineral County", 
         "Acres_Monongalia County", "Percent_Monongalia County", "Acres_Preston County", "Percent_Preston County")

SurveyConcreteCorrosionPA  <- SurveyCorridor %>%
  group_by(State.name, RouteName, CountyName, Concrete.Corrosion) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, RouteName, CountyName, (match(Concrete.Corrosion, CorrosionOrder))) %>%
  filter(State.name == "Pennsylvania") %>%
  group_by(RouteName, Concrete.Corrosion) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(RouteName, (match(Concrete.Corrosion, CorrosionOrder)))  %>%
  select("State.name", "RouteName", "Concrete.Corrosion", "Acres_Fayette County", "Percent_Fayette County", "Acres_Greene County", "Percent_Greene County")


ImpactConcreteCorrosionPA  <- Impact %>%
  group_by(State.name, Name, CountyName, Concrete.Corrosion) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, CountyName, (match(Concrete.Corrosion, CorrosionOrder))) %>%
  filter(State.name == "Pennsylvania") %>%
  group_by(Name, Concrete.Corrosion) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(Name, (match(Concrete.Corrosion, CorrosionOrder))) %>%
  select("State.name", "Name", "Concrete.Corrosion", "Acres_Fayette County", "Percent_Fayette County", "Acres_Greene County", "Percent_Greene County")


#####hydric#####
SurveyHydricWV  <- SurveyCorridor %>%
  group_by(State.name, RouteName, CountyName, Hydric.Rating) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, RouteName, CountyName, (match(Hydric.Rating, HydricOrder))) %>%
  filter(State.name == "WestVirginia") %>%
  group_by(RouteName, Hydric.Rating) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(RouteName, (match(Hydric.Rating, HydricOrder))) %>%
  select("State.name", "RouteName", "Hydric.Rating", "Acres_Hampshire County", "Percent_Hampshire County", "Acres_Mineral County", "Percent_Mineral County", 
         "Acres_Monongalia County", "Percent_Monongalia County", "Acres_Preston County", "Percent_Preston County")
#SurveyHydric$Percent <- paste0(SurveyHydric$Percent, "%")

ImpactHydricWV  <- Impact %>%
  group_by(State.name, Name, CountyName, Hydric.Rating) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, CountyName, (match(Hydric.Rating, HydricOrder))) %>%
  filter(State.name == "WestVirginia") %>%
  group_by(Name, Hydric.Rating) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(Name, (match(Hydric.Rating, HydricOrder))) %>%
  select("State.name", "Name", "Hydric.Rating", "Acres_Hampshire County", "Percent_Hampshire County", "Acres_Mineral County", "Percent_Mineral County", 
         "Acres_Monongalia County", "Percent_Monongalia County", "Acres_Preston County", "Percent_Preston County")

SurveyHydricPA  <- SurveyCorridor %>%
  group_by(State.name, RouteName, CountyName, Hydric.Rating) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, RouteName, CountyName, (match(Hydric.Rating, HydricOrder))) %>%
  filter(State.name == "Pennsylvania") %>%
  group_by(RouteName, Hydric.Rating) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(RouteName, (match(Hydric.Rating, HydricOrder))) %>%
  select("State.name", "RouteName", "Hydric.Rating", "Acres_Fayette County", "Percent_Fayette County", "Acres_Greene County", "Percent_Greene County")
#SurveyHydric$Percent <- paste0(SurveyHydric$Percent, "%")

ImpactHydricPA  <- Impact %>%
  group_by(State.name, Name, CountyName, Hydric.Rating) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, CountyName, (match(Hydric.Rating, HydricOrder))) %>%
  filter(State.name == "Pennsylvania") %>%
  group_by(Name, Hydric.Rating) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(Name, (match(Hydric.Rating, HydricOrder))) %>%
  select("State.name", "Name", "Hydric.Rating", "Acres_Fayette County", "Percent_Fayette County", "Acres_Greene County", "Percent_Greene County")

#####drain#####
SurveyDrainClassWV  <- SurveyCorridor %>%
  group_by(State.name, RouteName, CountyName, Drainage.Class...Dominant.Condition) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, RouteName, CountyName, (match(Drainage.Class...Dominant.Condition, DrainOrder))) %>%
  filter(State.name == "West Virginia") %>%
  group_by(RouteName, Drainage.Class...Dominant.Condition) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(RouteName, (match(Drainage.Class...Dominant.Condition, DrainOrder))) %>%
  select("State.name", "RouteName", "Drainage.Class...Dominant.Condition", "Acres_Hampshire County", "Percent_Hampshire County", "Acres_Mineral County", "Percent_Mineral County", 
         "Acres_Monongalia County", "Percent_Monongalia County", "Acres_Preston County", "Percent_Preston County")
#SurveyDrainClass$Percent <- paste0(SurveyDrainClass$Percent, "%")

ImpactDrainClassWV  <- Impact %>%
  group_by(State.name, Name, CountyName, Drainage.Class...Dominant.Condition) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, Name, CountyName, (match(Drainage.Class...Dominant.Condition, DrainOrder))) %>%
  filter(State.name == "West Virginia") %>%
  group_by(Name, Drainage.Class...Dominant.Condition) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(Name, (match(Drainage.Class...Dominant.Condition, DrainOrder))) %>%
  select("State.name", "Name", "Drainage.Class...Dominant.Condition", "Acres_Hampshire County", "Percent_Hampshire County", "Acres_Mineral County", "Percent_Mineral County", 
         "Acres_Monongalia County", "Percent_Monongalia County", "Acres_Preston County", "Percent_Preston County")

SurveyDrainClassPA  <- SurveyCorridor %>%
  group_by(State.name, RouteName, CountyName, Drainage.Class...Dominant.Condition) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, RouteName, CountyName, (match(Drainage.Class...Dominant.Condition, DrainOrder))) %>%
  filter(State.name == "Pennsylvania") %>%
  group_by(RouteName, Drainage.Class...Dominant.Condition) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0) %>%
  arrange(RouteName, (match(Drainage.Class...Dominant.Condition, DrainOrder))) %>%
  select("State.name", "RouteName", "Drainage.Class...Dominant.Condition", "Acres_Fayette County", "Percent_Fayette County", "Acres_Greene County", "Percent_Greene County")

ImpactDrainClassPA  <- Impact %>%
  group_by(State.name, Name, CountyName, Drainage.Class...Dominant.Condition) %>% 
  summarise(Acres = round(sum(Acres), 1), Percent = round(sum(Acres)/(mean(TotalAcrebyCounty))*100, 1)) %>%
  arrange(State.name, Name, CountyName, (match(Drainage.Class...Dominant.Condition, DrainOrder))) %>%
  filter(State.name == "Pennsylvania") %>%
  group_by(Name, Drainage.Class...Dominant.Condition) %>%
  pivot_wider(, names_from = CountyName, values_from = c(Acres, Percent), values_fill = 0)%>%
  arrange(Name, (match(Drainage.Class...Dominant.Condition, DrainOrder))) %>%
  select("State.name", "Name", "Drainage.Class...Dominant.Condition", "Acres_Fayette County", "Percent_Fayette County", "Acres_Greene County", "Percent_Greene County")


#####write tables#####

Tables <- list("SurveyFarmlandClassWV" = SurveyFarmlandClassWV, "ImpactFarmlandClassWV" = ImpactFarmlandClassWV,
               "SurveySteelCorrosionWV" = SurveySteelCorrosionWV, "ImpactSteelCorrosionWV" = ImpactSteelCorrosionWV,
               "SurveyConcreteCorrosionWV" = SurveyConcreteCorrosionWV, "ImpactConcreteCorrosionWV" = ImpactConcreteCorrosionWV,
               "SurveyHydricWV" = SurveyHydricWV, "ImpactHydricWV" = ImpactHydricWV, 
               "SurveyDrainClassWV" = SurveyDrainClassWV, "ImpactDrainClassWV" = ImpactDrainClassWV,
               "SurveyFarmlandClassPA" = SurveyFarmlandClassPA, "ImpactFarmlandClassPA" = ImpactFarmlandClassPA,
               "SurveySteelCorrosionPA" = SurveySteelCorrosionPA, "ImpactSteelCorrosionPA" = ImpactSteelCorrosionPA,
               "SurveyConcreteCorrosionPA" = SurveyConcreteCorrosionPA, "ImpactConcreteCorrosionPA" = ImpactConcreteCorrosionPA,
               "SurveyHydricPA" = SurveyHydricPA, "ImpactHydricPA" = ImpactHydricPA, 
               "SurveyDrainClassPA" = SurveyDrainClassPA, "ImpactDrainClassPA" = ImpactDrainClassPA
               )
write.xlsx(Tables, file = "SoilTables.xlsx")


#####State totals#####

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
