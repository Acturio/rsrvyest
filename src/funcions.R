library(foreign)
library(readxl)
library(survey)
library(dplyr)
library(tidyr)
library(xlsx)
library(caret)

base = "BASE_CONACYT_260118.sav" #base.sav
lista = "Lista de Preguntas.xlsx" #archivo lista de variables

leer_datos <- function(base, lista){
  #se asume misma organización de carpetas 
  archivo <- paste0("data/", base)
  archivo2 <- paste0("lista/", lista)
  
  # Lectura de datos de spss
  dataset <- read.spss(archivo, to.data.frame = TRUE) 
  
  #hojas
  General <- "Total" # Nombre de estimación global (Puede ser nacional, cdmx, etc. Depende de la representatividad del estudio)
  
  Lista_vars <- read_xlsx(archivo2, sheet = "Lista Preguntas") %>%  
    pull(Pregunta) 
  Lista_Preg <- read_xlsx(archivo2, sheet = "Lista Preguntas") %>%  
    pull(Nombre) 
  DB_Mult <- read_xlsx(archivo2, sheet = "Múltiple") %>% 
    as.data.frame()
  Lista_Cont <- read_xlsx(archivo2, sheet = "Continuas") %>% 
    pull(VARIABLE)
  Dominios <- read_xlsx(archivo2, sheet = "Dominios") %>%  
    pull(Dominios)
  
  Multiples <- names(DB_Mult)
  Ponderador <- pull(dataset, Pondi1)
  #save = ""
  
  return(list(Lista_vars, Lista_Preg, DB_Mult, Lista_Cont, Dominios, Multiples, Ponderador))
  #return: lista de listas de preguntas, acceder []
}

#ver etiqueta vars continuas SPSS*
listaD <- leer_datos("BASE_CONACYT_260118.sav","Lista de Preguntas.xlsx")


library(srvyr)

crea_disenio <- function(data, id, cestrato, cpesos){
  
  disenio <- data %>% 
    as_survey_design(
      ids = {{id}},
      weights = {{cpesos}},
      strata = {{cestrato}}
    )
  
  return(disenio)
}

mdesign <- crea_disenio(dataset, CV_ESC, ESTRATO, Pondi1)

library(srvyr)
frecuencias_continuas <- function(disenio, pregunta, na.rm = TRUE,
                                  vartype = c("se", "ci", "cv", "var"), 
                                  level = 0.95, proportion = FALSE, prop_method = "likelihood",
                                  DEFF = TRUE) {

  estadisticas <- {{disenio}} %>% 
    #srvyr::group_by(!!sym(pregunta)) %>% 
    srvyr::summarise(
      prop = survey_mean(
        as.numeric(!!sym(pregunta)),
        na.rm = na.rm,
        vartype = vartype,
        level = level,
        proportion = proportion,
        prop_method = prop_method,
        deff = DEFF
      ),
      #total = survey_total(
      #  na.rm = na.rm,
      #  vartype = vartype,
      #  deff = DEFF
      #)
    ) %>% 
      mutate(prop_low = ifelse(prop_low < 0, 0, prop_low),
             prop_upp = ifelse(prop_upp > 1, 1, prop_upp)) %>% 
      as.data.frame()
  
  return(estadisticas)
  
}


#las buenas

cuantiles = c(0,0.25, 0.5, 0.75,1)
pregunta <- listaD[[4]][1]

estadisticas_continuas <- function(disenio, pregunta, na.rm = TRUE,
                                  vartype = c("se", "ci", "cv", "var"), 
                                  level = 0.95, proportion = FALSE, prop_method = "likelihood",
                                  DEFF = TRUE, cuantiles) {
  
  estadisticas <- disenio %>% 
    #srvyr::group_by(!!sym(pregunta)) %>% 
    srvyr::summarise(
      prop = survey_mean(
        as.numeric(!!sym(pregunta)),
        na.rm = na.rm,
        vartype = vartype,
        level = level,
        proportion = proportion,
        prop_method = prop_method,
        deff = DEFF
      ),
      cuantiles = survey_quantile(
        as.numeric(!!sym(pregunta)),
        quantiles = cuantiles,
        na.rm = na.rm
      )
    ) %>% 
    mutate(prop_low = ifelse(prop_low < 0, 0, prop_low),
           prop_upp = ifelse(prop_upp > 1, 1, prop_upp)) %>% 
    as.data.frame() %>% 
    select(prop, prop_low, prop_upp, cuantiles_q00, cuantiles_q25, cuantiles_q50,
           cuantiles_q75, cuantiles_q100, prop_se, prop_var, prop_cv, prop_deff)
  
  return(estadisticas)
  
}


library(reshape)
df <- estadisticas

acomoda_frecuencias <- function(df){
  df_t <- df %>% 
    melt()
  names(df_t) <- c("stat", "valor")
  
  lvars <- c("media", "lim_inf", "lim_sup", "mín", "Q25", "mediana", "Q75", "máx", 
             "sd", "var", "cv", "deff")
  
  lvars <- c("Media", "Lim_inf", "Lim_sup", "Mín", "Q25", "Mediana", "Q75", "Máx", 
             "Sd", "Var", "C.V.", "Deff")
  
  nvo_df <- data.frame(lvars, round(df_t[,2],3))
  names(nvo_df) <- c("Métrica", "Valor")
  
  
  
  return(nvo_df)
  
}


#total
#total
library(tibble)

total <- function(disenio, pregunta, na.rm = TRUE,
                  vartype = c("se", "ci", "cv", "var"), 
                  level = 0.95, proportion = FALSE, prop_method = "likelihood",
                  DEFF = TRUE, cuantiles) {
  
  total <- disenio %>% 
    #srvyr::group_by(!!sym(pregunta)) %>% 
    srvyr::summarise(
      prop = survey_mean(
        as.numeric(!!sym(pregunta)),
        na.rm = na.rm,
        vartype = vartype,
        level = level,
        deff = DEFF
      )
    )
  
  total %>% 
    as.data.frame()
  
  total <- add_column(total, dominio = "Total", .before = "prop")
  total <- add_column(total, var = "Total", .before = "dominio")
  
  return(total)
  
}


#for sobre lista de preguntas
#for sobre dominios
library(stringr)
Dominios <- listaD[[5]]
dominio = Dominios[1]

tabla_cruzada <- function(disenio, pregunta, dominio, na.rm = TRUE,
                          vartype = c("se", "ci", "cv", "var"), 
                          level = 0.95, DEFF = TRUE){
  
  cruce <- disenio %>% 
    srvyr::group_by(!!sym(dominio)) %>% 
    srvyr::summarise(
      prop = survey_mean(
        as.numeric(!!sym(pregunta)),
        na.rm = na.rm,
        vartype = vartype,
        level = level,
        deff = DEFF
      )
    )
  
  #cruce <- cruce %>% 
  #  as.data.frame()
  
  cruce %<>% 
    mutate(!!sym(dominio) := str_trim(!!sym(dominio), side = "both")) %>% 
    dplyr::rename(dominio = !!sym(dominio))
  
  cruce <- add_column(cruce, var = dominio, .before = "dominio")
  
  return(cruce)
}

tabla <- total
#prueba
for (dom in Dominios){
  
  print(dom)
  cruce <- tabla_cruzada(mdesign, pregunta, dom, na.rm = TRUE,
                         vartype = c("se", "ci", "cv", "var"), 
                         level = 0.95, DEFF = TRUE)
  
  
  
  tabla <- bind_rows(tabla, cruce)
  
}



#calcula totales
#formato tabla 
df = tabla

formato_tabla <- function(df){
  
  df1 <- df %>% 
    select(var, dominio, prop, prop_low, prop_upp)
  df2 <- df %>% 
    select(var, dominio, prop_se, prop_cv, prop_var, prop_deff)
  
  names(df1) <- c("Dominio", "Categoría", "Media", "Lim. inf.", "Lim. sup.")
  names(df2) <- c("Dominio", "Categoría", "Err. Est", "Coef. Var.", "Var.", "DEFF")
  
  ldf <- list(df1, df2)
  
  return(ldf)
  
}



#####################################################################3
#formato excel frecuencias

#1. estadisticas_continuas
#2. acomoda_frecuencias

#3. resultado = nvo_df

library(openxlsx)

wb <- createWorkbook()
addWorksheet(wb, "writeData auto-formatting")

hs1 <- createStyle(halign = "CENTER", textDecoration = "Bold",
                   border = "TopBottomLeftRight", fontColour = "black",
                   borderStyle = "medium", borderColour = "black")
writeData(wb, 1, nvo_df, startRow = 2, startCol = 2, headerStyle = hs1,
          borders = "columns", borderStyle = "medium", colNames = TRUE,
          borderColour = "black")
#writeData(wb, 1, df)
s <- createStyle(numFmt = "0")
addStyle(wb, 1, style = s, rows = 3:14, cols = 3, stack = T)


openXL(wb)
#saveWorkbook(wb, "results/frecuencias.xlsx", overwrite = TRUE)


#funcion_frecuencias

tabla_frec <- function(df){
  
}

#####################################################################3
#formato excel TC

library(openxlsx)

wb <- createWorkbook()
addWorksheet(wb, "writeData auto-formatting")

hs1 <- createStyle(halign = "CENTER", textDecoration = "Bold",
                   border = "TopBottomLeftRight", fontColour = "black",
                   borderStyle = "medium", borderColour = "black")

colini = 2
rowini = 2
#ncols df
kol <- ncol(df)
#rnows df
ren <- nrow(df)

writeData(wb, 1, df, startRow = rowini, startCol = colini, headerStyle = hs1,
          borders = "columns", borderStyle = "medium", colNames = TRUE,
          borderColour = "black")

#Formato de número y centrado
s <- createStyle(numFmt = "0.0", halign = "center", valign = "center")


rf <- rowini+ren
ri <- rowini+1
#cols afectadas con numero y centrados
addStyle(wb, 1, style = s, rows = ri:rf, cols = (colini+2):(colini+kol-1), stack = T, gridExpand = T)

#combino celdas
#mergeCells(wb, 1, cols = 2, rows = 4:5)
#mergeCells(wb, 1, cols = 2, rows = 6:12)

c=0
finicio = 2

centerStyle <- createStyle(valign = "center")
insideBorders <- createStyle(
  border = "bottom",
  borderStyle = "thin"
)

for (dom in Dominios){
  
  print(dom)
  
  for (k in finicio:ren){
    
    if (df[k,1]==dom){
      c = c + 1
    } else {
      c = c
    }
  
    i = c - 1
  }
  
  mergeCells(wb, 1, cols = 2, rows = (finicio + 2): (finicio + 2 + i))
  addStyle(wb, 1, centerStyle, rows = (finicio + 2): (finicio + 2 + i), cols = 2, stack = T, gridExpand = T)
  addStyle(wb, 1, insideBorders, rows = finicio + 2 + i, cols = 2:(kol+1), stack = T, gridExpand = T)
  finicio = finicio + c
  c = 0
}

openXL(wb)


####################################################################################
library(openxlsx)

wb <- createWorkbook()
addWorksheet(wb, "writeData auto-formatting")
#renglon y columna de inicio
colini = 2
rowini = 2

tabla_excel <- function(df, colini, rowini){
  
  hs1 <- createStyle(halign = "CENTER", textDecoration = "Bold",
                     border = "TopBottomLeftRight", fontColour = "black",
                     borderStyle = "medium", borderColour = "black")
  #Formato de número y centrado
  s <- createStyle(numFmt = "0.0", halign = "center", valign = "center")
  #centrado
  
  centerStyle <- createStyle(valign = "center")
  insideBorders <- createStyle(
    border = "bottom",
    borderStyle = "thin"
  )

  #ncols df
  kol <- ncol(df)
  #rnows df
  ren <- rowini+nrow(df)
  
  writeData(wb, 1, df, startRow = rowini, startCol = colini, headerStyle = hs1,
            borders = "columns", borderStyle = "medium", colNames = TRUE,
            borderColour = "black")
  #manipulo renglones
  rf <- rowini+ren
  ri <- rowini+1
  #cols afectadas con numero y centrados
  addStyle(wb, 1, style = s, rows = ri:rf, cols = (colini+2):(colini+kol-1), stack = T, gridExpand = T)
  #formato interior
  c=0
  finicio = rowini
  
  
  for (dom in Dominios){
    
    print(dom)
    
    for (k in finicio:ren){
      
      if (df[k,1]==dom){
        c = c + 1
      } else {
        c = c
      }
      
      i = c - 1
    }
    
    mergeCells(wb, 1, cols = 2, rows = (finicio + 2): (finicio + 2 + i))
    addStyle(wb, 1, centerStyle, rows = (finicio + 2): (finicio + 2 + i), cols = 2, stack = T, gridExpand = T)
    addStyle(wb, 1, insideBorders, rows = finicio + 2 + i, cols = 2:(kol+1), stack = T, gridExpand = T)
    finicio = finicio + c
    c = 0
  }
  
return(openXL(wb))  
  
}

#ejemplo funcion tabla_excel
tabla_excel(df, 2, 2)

######################3
Lista_Cont <- listaD[[4]]
pregunta <- listaD[[4]][1]
p <- listaD[[4]][2]
formato = 1
#for sobre todas las preguntas
wb <- createWorkbook()
addWorksheet(wb, "writeData auto-formatting")
#renglon y columna de inicio
colini = 2
rowini = 2

for (p in Lista) {
  
  if (p %in% Lista_Cont) {
    
    print(p)
    t <- total(disenio, p, na.rm = TRUE,
              vartype = c("se", "ci", "cv", "var"), 
              level = 0.95, proportion = FALSE, prop_method = "likelihood",
              DEFF = TRUE, cuantiles) 
    print(t)
    #tc
    tabla <- t
    for (dom in Dominios){
      cruce <- tabla_cruzada(mdesign, p, dom, na.rm = TRUE,
                             vartype = c("se", "ci", "cv", "var"), 
                             level = 0.95, DEFF = TRUE)
      tabla <- bind_rows(tabla, cruce)
      
    }
    print(tabla)
    #formato df
    ltabla <- formato_tabla(tabla)
    if (formato == 1) {
      tablaf <- as.data.frame(ltabla[1])
    } else {
      tablaf <- as.data.frame(ltabla[2])
    }
    
    #escribo el excel
      tabla_excel(tablaf, colini, rowini)
    
    #recalculo renglones
    rowini <- rowini + nrow(tablaf) +5
    
  }
  
}


