install.packages("pacman")
pacman::p_load(janitor,officer,gtsummary,tidyverse,gt,ggplot2,ggpubr,ggthemes,viridis,sjPlot)
library(readxl)
Short_Listed <- read_excel("Short-Listed.xlsx",sheet = 1)
view(Short_Listed)
A1<-Short_Listed %>% 
dplyr::select(Gender,`Level of Education`,`Professional Training`,`Types of farms`,`Name of Chicken family`,`No of Workers`,`Sanitation system`) %>%  
mutate(across(where(is.character),as.factor)) %>% 
tbl_summary(by=Gender,
            missing = "no") %>% 
  add_p() %>% 
  add_overall() %>% 
  bold_labels() %>% 
as_gt() %>% 
  gt::gtsave( filename = "C:/Users/Sc/Documents/Poultry/ARS2.docx")


ggplot(A1, aes(x=`Total surface area of poultry square feet`, y=`No of chicken`)) +
  geom_point() +
  labs(title = " Total Surface Area vs No of Chicken",
       x = "Total Surface Area of Poultry (square feet)",
       y = "Number of Chickens") +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5))

A1 %>% 
  plot_frq(`Sanitation system`)

A1 %>% 
  plot_frq(`Why antibiotics are used?`)

A1 %>% 
  plot_frq(`Mention the name of the antibiotics...22`)

A3<-Short_Listed %>% 
  dplyr::select(`Timing of antibiotic application`,`Vet. doctor's suggestion`,
                `Feed dealer's suggestion`,
                `Drug seller's suggestion`,
                `Pharmaceutical company representative's suggestion`,
                `Self decision`) %>%  
  mutate(across(where(is.character),as.factor)) %>% 
  tbl_summary(by=`Timing of antibiotic application`,
              missing = "no") %>% 
  add_p() %>% 
  add_overall() %>% 
  bold_labels() %>% 
  as_gt() %>% 
  gt::gtsave( filename = "C:/Users/Sc/Documents/Poultry/Suggestion.docx")
  

ggplot(Short_Listed2, aes(x=`Average cost for antibiotics per batch-cycle ( BDT)`,
                          y=`No of chicken`)) +
  geom_point(size=2, color="blue") +
  labs(title = "No of Chicken vs Average cost per batch",
       x = "Average cost for antibiotics per batch-cycle (BDT)",
       y = "Number of Chickens") +
  theme_minimal() +
  theme(plot.title = element_text(hjust = 0.5))

Short_Listed2 <- read_excel("Short-Listed2.xlsx",sheet = 1) 

A4<-Short_Listed2 %>%
  dplyr::select(`Level of Education`,`Do you know about the withdrawn period of antibiotics:`,
                `Do you know about the antimicrobial resistance (AMR):`) %>% 
mutate(across(where(is.character),as.factor)) %>% 
  mv_reg <- glm(`Level of Education` ~ `Do you know about the withdrawn period of antibiotics:` +`Do you know about the antimicrobial resistance (AMR):`
                ,family = "binomial", data = A4)

summary(mv_reg)
## show results table of final regression 
mv_tab <- tbl_regression(mv_reg, exponentiate = TRUE)

mv_tab

mv_tab_base<- data.frame(mv_tab_base) %>% 
  print(mv_tab_base, target = "C:/Users/Sc/Documents/Poultry/Regression.docx")
  
## choose a model using forward selection based on AIC
## you can also do "backward" or "both" by adjusting the direction
final_mv_reg <- mv_reg %>%
  step(direction = "forward", trace = FALSE)
options(scipen=999)

mv_tab_base <- final_mv_reg %>% 
  broom::tidy(exponentiate = TRUE, conf.int = TRUE) %>%  ## get a tidy dataframe of estimates 
  mutate(across(where(is.numeric), round, digits = 2)) 
## New


# Create a data frame with the data provided
data <- data.frame(
  term = c("(Intercept)", 
           "Do you know about the withdrawn period of antibiotics:", 
           "Do you know about the antimicrobial resistance (AMR):"),
  estimate = c(3.2e0, 7.65e7, 2.5e0),
  std.error = c(0.51, 3063, 1.18),
  statistic = c(2.27, 0.01, 0.78),
  p.value = c(0.02, 1, 0.44),
  conf.low = c(1.25, 0, 0.33),
  conf.high = c(9.78, NA, 52.2)
)
# Load the necessary package
library(knitr)

# Display the table in a nicely formatted way
kable(data, format = "pipe")


library(officer)
doc <- read_docx()
doc <- body_add_table(doc, value = data, style = "table_template")
print(doc, target = "C:/Users/Sc/Documents/Poultry/Regression.docx")

