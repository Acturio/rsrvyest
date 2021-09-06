library(tidyverse)
library(survey)
library(srvyr)
library(magrittr)
library(tidyr)
library(foreign)
library(Hmisc)
library(openxlsx)

##### Disenio ######
disenio_multiples <- function(id, estrato, pesos, datos, pps = "brewer",
                              varianza = "HT", reps = TRUE, metodo = "subbootstrap", B=50, semilla=1234){
  
  disenio <- datos %>% 
    as_survey_design(
      ids = {{id}}, 
      weights = {{pesos}}, 
      strata = {{estrato}},
      pps = pps,
      variance = varianza
    )
  
  if (reps){
    disenio %<>% as_survey_rep(
      type = metodo, 
      replicates = B
    )
  }
  return(disenio)
}

disenio <- disenio_multiples(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1,
                             reps=FALSE, datos = dataset)

##### Función múltiples

############################################
############# Frecuencias simples ##########
############################################

frecuencias_simples_multiples <- function(pregunta, DB_Mult, disenio, na.rm = TRUE, datos,
                                          estadisticas = c("se","ci","cv", "var"), 
                                          significancia = 0.95, 
                                          proporcion = FALSE,
                                          metodo_prop = "likelihood",
                                          DEFF = TRUE){
  
  ## Onehot encoding
  ps <- DB_Mult %>% 
    dplyr::filter(!is.na(!!sym(pregunta))) %>% 
    dplyr::pull(!!sym(pregunta))
  
  df <- datos %>% 
    select(ps)
  
  categorias <- df %>% 
    pull() %>% 
    levels()
  
  numero_categorias <- length(categorias)
  
  df <- df %>% mutate(ID = row.names(df))
  
  one_hot <- caret::dummyVars(" ~ .", data=df)
  one_hot <- data.frame(predict(one_hot, newdata = df))
  one_hot[is.na(one_hot)] <- 0
  
  menciones_juntas <- matrix(NA, nrow(df), ncol=numero_categorias) %>% 
    as_tibble()
  names(menciones_juntas) <- categorias
  
  dum <- NULL
  for(j in 1:numero_categorias){
    dum <- one_hot[,j]
    for (i in 1:(length(ps)-1)) {
      dum <- dum + one_hot[,j+i*numero_categorias]  
    }
    menciones_juntas[,j] <- dum
  } 
  menciones_juntas[menciones_juntas > 1] <- 1
  
  menciones_vector <- menciones_juntas %>% names() %>% as_vector()
  
  ## Agregamos variables onehot a diseño
  
  for (i in menciones_vector){
    variable <- menciones_juntas %>% 
      pull(!!sym(i))
    disenio %<>% srvyr::mutate(!!sym(i) := variable)
  }
  
  frecuencias_simples = data_frame()
  
  ### Cálculo de frecuencias simples de todas las categorías de una pregunta
  
  for (categ in categorias) {
    nacional <- disenio %>% 
      srvyr::summarize(
        prop = survey_mean(!!sym(categ),
                           na.rm = na.rm, 
                           vartype = estadisticas,
                           level = significancia,
                           proportion = proporcion,
                           prop_method = metodo_prop,
                           deff = DEFF),
        total = survey_total(!!sym(categ),
          na.rm = TRUE)) %>% 
      mutate(prop_low = ifelse(prop_low < 0, 0, prop_low),
             prop_upp = ifelse(prop_upp > 1, 1, prop_upp),
             Respuesta := categ) 
    
    frecuencias_simples <- bind_rows(frecuencias_simples, nacional)
  }
  return(frecuencias_simples)
}


a = frecuencias_simples_multiples(pregunta = 'P1', DB_Mult=DB_Mult, disenio=disenio, datos = dataset)


#### Formato frecuencias simples en R (preguntas multiples)

formatear_tabla_multiple <- function(tabla){
  
  tabla1 <-  tabla %>%
    map_at(c('prop', 'prop_low', 'prop_upp', 'prop_se'), ~.x*100) %>%
    as_tibble() %>% 
    mutate(
      "Total" = round(total, 0),
      "Media" = prop,
      "Lim. inf." = prop_low,
      "Lim. sup." = prop_upp,
      "Err. Est." = prop_se,
      "Var." = prop_var,
      "Coef. Var." = prop_cv,
      "DEFF" = prop_deff
    ) %>% 
    select(Respuesta,
           "Total",
           "Media", "Lim. inf.", "Lim. sup.") 
  
  tabla2 <- tabla %>%
    mutate(
      "Total" = round(total, 0),
      "Media" = prop,
      "Lim. inf." = prop_low,
      "Lim. sup." = prop_upp,
      "Err. Est." = prop_se * 100,
      "Var." = prop_var * 10000,
      "Coef. Var." = prop_cv,
      "DEFF" = prop_deff
    ) %>% 
    select(Respuesta,"Total", "Err. Est." , "Coef. Var.","Var.", "DEFF") 
  
  return(list(tabla1, tabla2))
}

freq <- formatear_tabla_multiple(tabla = a)
freq

# Escribir en excel

escribir_tabla <- function(tabla, wb, hoja, renglon, columna,
                           bordes = 'surrounding', color_borde = 'black',
                           estilo_borde = 'medium', filtro = FALSE,
                           na = TRUE, reemplazo_na = '-',
                           nombres_columnas = FALSE, estilo_encabezado = NULL){
  
  
  writeData(wb = wb, sheet = hoja, x = tabla, startRow = renglon, 
            startCol = columna, colNames = nombres_columnas, borders = bordes,
            borderColour = color_borde, borderStyle = estilo_borde,
            withFilter = filtro, keepNA = na, na.string = reemplazo_na,
            col.names = nombres_columnas, headerStyle = estilo_encabezado)
  
}


####### Función escribir frecuencias simples en excel

formato_frecuencias_simples <- function(tabla, wb, hojas = c(1,2) ,renglon, columna,
                                        estilo_encabezado){
  
  # Primera Página renglón 2 columna 1
  
  escribir_tabla(tabla = tabla[[1]], wb = wb, hoja = hojas[1], renglon = renglon,
                 columna = columna, nombres_columnas = TRUE, estilo_borde = 'thin',
                 estilo_encabezado = estilo_encabezado)
  
  
  
  escribir_tabla(tabla = tabla[[2]], wb = wb, hoja = hojas[2], renglon = renglon,
                 columna = columna, estilo_borde = 'thin', nombres_columnas = TRUE,
                 estilo_encabezado = estilo_encabezado)
  
  
}

###### Función con todos los pasos:
frecuencias_simples <- function(pregunta, num_pregunta, datos, lista_preguntas,
                                DB_Mult, diseño, wb, renglon, columna = 1, hojas = c(1,2),
                                estilo_encabezado = headerStyle){
  
  # Título Pregunta
  
  np <- lista_preguntas[[num_pregunta]]
  
  escribir_tabla(tabla = np, wb = wb, hoja = hojas[1], renglon = renglon, columna = columna)
  
  escribir_tabla(tabla = np, wb = wb, hoja = hojas[2], renglon = renglon, columna = columna)
  
  
  # Frecuencias Simples
  
  estimaciones <- frecuencias_simples_multiples(pregunta, DB_Mult, diseño, datos = datos)
  
  # Formatear tabla 
  frecuencias_simples_formato <- formatear_tabla_multiple(tabla = estimaciones)
  
  
  formato_frecuencias_simples(tabla = frecuencias_simples_formato, wb = wb, renglon = (renglon + 1), 
                              columna = columna, estilo_encabezado = estilo_encabezado)
  
  
}


############################################
############# Tablas cruzadas ##############
############################################

### Función tablas cruzadas

tablas_cruzadas_multiples <- function(pregunta, frecuencias_simples, dominio, DB_Mult, disenio, na.rm = TRUE, datos,
                                          estadisticas = c("se","ci","cv", "var"), 
                                          significancia = 0.95, 
                                          proporcion = FALSE,
                                          metodo_prop = "likelihood",
                                          DEFF = TRUE){
  ## Onehot 
  ps <- DB_Mult %>% 
    dplyr::filter(!is.na(!!sym(pregunta))) %>% 
    dplyr::pull(!!sym(pregunta))
  
  df <- datos %>% 
    select(ps)
  
  categorias <- df %>% 
    pull() %>% 
    levels()
  
  numero_categorias <- length(categorias)
  
  df <- df %>% mutate(ID = row.names(df))
  
  one_hot <- caret::dummyVars(" ~ .", data=df)
  one_hot <- data.frame(predict(one_hot, newdata = df))
  one_hot[is.na(one_hot)] <- 0
  
  menciones_juntas <- matrix(NA, nrow(df), ncol=numero_categorias) %>% 
    as_tibble()
  names(menciones_juntas) <- categorias
  
  dum <- NULL
  for(j in 1:numero_categorias){
    dum <- one_hot[,j]
    for (i in 1:(length(ps)-1)) {
      dum <- dum + one_hot[,j+i*numero_categorias]  
    }
    menciones_juntas[,j] <- dum
  } 
  menciones_juntas[menciones_juntas > 1] <- 1
  
  menciones_vector <- menciones_juntas %>% names() %>% as_vector()
  
  for (i in menciones_vector){
    variable <- menciones_juntas %>% 
      pull(!!sym(i))
    disenio %<>% srvyr::mutate(!!sym(i) := variable)
  }
  
  tablas_cruzadas = data_frame()
  
    for (categ in categorias){
    Dominios_tabla <- disenio %>% 
      srvyr::group_by(!!sym(dominio), .drop = TRUE) %>% 
      srvyr::summarize(
        prop = survey_mean(!!sym(categ),
                           na.rm = na.rm, 
                           vartype = estadisticas,
                           level = significancia,
                           proportion = proporcion,
                           prop_method = metodo_prop,
                           deff = DEFF),
        total = survey_total(!!sym(categ),
          na.rm = TRUE
        )) %>% 
      mutate(prop_low = ifelse(prop_low < 0, 0, prop_low),
             prop_upp = ifelse(prop_upp > 1, 1, prop_upp),
             Dominio = dominio) %>% 
      dplyr::rename(Categorias := !!sym(dominio))
    
    tablas_cruzadas <- rbind(tablas_cruzadas, Dominios_tabla)
    }
  
  tablas_cruzadas %<>% select(Dominio, Categorias,total,
                            total_se, prop,prop_se, 
                            prop_low, prop_upp,
                            prop_cv, prop_var, prop_deff)
  return(tablas_cruzadas)
}

a = tablas_cruzadas_multiples(pregunta = 'P1',dominio = "Edad",frecuencias_simples = a, DB_Mult=DB_Mult, 
                              disenio=disenio, datos = dataset,
                                  na.rm = TRUE, estadisticas = c("se","ci","cv", "var"), 
                                  significancia = 0.95, proporcion = FALSE,
                                  metodo_prop = "likelihood", DEFF = TRUE)

################ Formatear tabla cruzada en R

formatear_tabla_cruzada <- function(tabla, datos, dominio, DB_Mult, pregunta){
  
  tabla1 <-  tabla %>%
    map_at(c('prop', 'prop_low', 'prop_upp', 'prop_se'), ~.x*100) %>%
    as_tibble() %>% 
    mutate(
      "Total" = round(total, 0),
      "Media" = prop,
      "Lim. inf." = prop_low,
      "Lim. sup." = prop_upp,
      "Err. Est." = prop_se,
      "Var." = prop_var,
      "Coef. Var." = prop_cv,
      "DEFF" = prop_deff
    ) %>% 
    select(Dominio,
           Categorias,
           "Total",
           "Media", "Lim. inf.", "Lim. sup.",-"Total") 
  
  tabla2 <- tabla %>%
    mutate(
      "Total" = round(total, 0),
      "Media" = prop,
      "Lim. inf." = prop_low,
      "Lim. sup." = prop_upp,
      "Err. Est." = prop_se * 100,
      "Var." = prop_var * 10000,
      "Coef. Var." = prop_cv,
      "DEFF" = prop_deff
    ) %>% 
    select(Dominio, Categorias, "Err. Est.", "Coef. Var.", "Var.", "DEFF", -"Total")
  
  suma_totales = tabla %>% dplyr::group_by(Categorias) %>% dplyr::summarise(Total = sum(total))
  
  len <- datos %>% select(!!sym(dominio)) %>% distinct() %>% nrow()
  
  ps <- DB_Mult %>% 
    dplyr::filter(!is.na(!!sym(pregunta))) %>% 
    dplyr::pull(!!sym(pregunta))
  
  df <- datos %>% 
    select(ps)
  
  categorias <- df %>% 
    pull() %>% 
    levels()
  
  nombres_tabla1 <- c('Dominio', 'Categorías', 'Total', rep(c('Media', 'Lim. inf.', 'Lim. sup.'),
                                                      length(categorias)))
  
  nombres_tabla2 <- c('Dominio', 'Categorías', 'Total', rep(c("Err. Est.", "Coef. Var.", "Var.", "DEFF"),
                                                      length(categorias)))
  
  
   tabla_final_1 <- tabla1[c(1:len),] 
   
   for (i in seq((len+1),nrow(tabla1), by = len)){
     tabla_c1 = tabla1[i:(i+len-1),c(-1,-2)]
     tabla_final_1 = cbind(tabla_final_1,tabla_c1)
   }
   
   suma_totales %<>% select(Total)
   
   
   tabla_final_1=cbind(tabla_final_1,suma_totales)
   
   tabla_final_1 %<>% relocate(Total,.after = Categorias)
   
   tabla_final_2 <- tabla2[c(1:len),] 
   
   for (i in seq((len+1),nrow(tabla2), by = len)){
     tabla_c2 = tabla2[i:(i+len-1),c(-1,-2)]
     tabla_final_2 = cbind(tabla_final_2,tabla_c2)
   }
  
  tabla_final_2 =cbind(tabla_final_2,suma_totales) 
  
  tabla_final_2 %<>% relocate(Total,.after = Categorias)
   
  names(tabla_final_1)=nombres_tabla1
  names(tabla_final_2)=nombres_tabla2
  
  return(list(tabla_final_1,tabla_final_2))
  }
  
b <-formatear_tabla_cruzada(a, datos = dataset, dominio = "Edad", DB_Mult = DB_Mult, pregunta = "P1")


### Funciones para crear renglón nacional para tablas cruzddas

desagregar_frecuencias_simples_tabla_estimaciones<- function(frecuencias_simples){
  
  categorias<-frecuencias_simples[[1]] %>% dplyr::pull(Respuesta)
  
  nombres_tabla_desagregada <- c('Dominio', 'Categorías', 'Total', rep(c('Media', 'Lim. inf.', 'Lim. sup.'),
                                                      length(categorias)))
  
  suma_totales = frecuencias_simples[[1]] %>% dplyr::summarise(Total = sum(Total))
  
  tabla_desagregada = frecuencias_simples[[1]][1,c(-1,-2)] %>% 
    dplyr::mutate(Dominio = "General", Categorias = "Total")
  
  for (i in seq(2,nrow(frecuencias_simples[[1]]), by = 1)){
    tabla_c1 = frecuencias_simples[[1]][i,c(-1,-2)]
    tabla_desagregada = cbind(tabla_desagregada,tabla_c1)
  }
  
  tabla_desagregada = cbind(suma_totales,tabla_desagregada)
  
  tabla_desagregada %<>% relocate(Categorias,.before = Total) %>% 
    relocate(Dominio,.before = Categorias)
  
  names(tabla_desagregada) = nombres_tabla_desagregada
  return(tabla_desagregada)
}

c = desagregar_frecuencias_simples_tabla_estimaciones(freq)  


desagregar_frecuencias_simples_tabla_errores<- function(frecuencias_simples){
  
  categorias<-frecuencias_simples[[2]] %>% dplyr::pull(Respuesta)
  
  nombres_tabla_desagregada <- c('Dominio', 'Categorías', 'Total', rep(c("Err. Est.", "Coef. Var.","Var.", "DEFF"),
                                                      length(categorias)))
  
  suma_totales = frecuencias_simples[[2]] %>% dplyr::summarise(Total = sum(Total))
  
  tabla_desagregada = frecuencias_simples[[2]][1,c(-1,-2)] %>% 
    dplyr::mutate(Dominio = "General", Categorias = "Total")
  
  for (i in seq(2,nrow(frecuencias_simples[[2]]), by = 1)){
    tabla_c1 = frecuencias_simples[[2]][i,c(-1,-2)]
    tabla_desagregada = cbind(tabla_desagregada,tabla_c1)
  }
  
  tabla_desagregada = cbind(suma_totales,tabla_desagregada)
  
  
  tabla_desagregada %<>% relocate(Categorias,.before = Total) %>% 
    relocate(Dominio,.before = Categorias)
  
  names(tabla_desagregada) = nombres_tabla_desagregada
  
  return(tabla_desagregada)
}

d = desagregar_frecuencias_simples_tabla_errores(freq) 

### Función final tabla cruzada por pregunta:

tabla_cruzada_por_pregunta <- function(disenio, pregunta, datos, dominios = Dominios, DB_Mult){
  
  # Frecuencias simples
  
  frecuencias_simples = frecuencias_simples_multiples(pregunta = pregunta, DB_Mult=DB_Mult, 
                                            disenio=disenio, datos = datos)
  
  frecuencias_simples_formato <- formatear_tabla_multiple(tabla = frecuencias_simples)
  
  # Renglón nacional
  
  # Estimaciones:
  estimaciones_generales = desagregar_frecuencias_simples_tabla_estimaciones(frecuencias_simples_formato)  
  
  # Errores
  errores_generales = desagregar_frecuencias_simples_tabla_errores(frecuencias_simples_formato)

  # Iterar sobre los dominios
  
  for (d in dominios) {
    tabla_cruzada = tablas_cruzadas_multiples(pregunta = pregunta, dominio = d, frecuencias_simples = frecuencias_simples,
                                  DB_Mult=DB_Mult,disenio=disenio, datos = datos,
                                  na.rm = TRUE, estadisticas = c("se","ci","cv", "var"), 
                                  significancia = 0.95, proporcion = FALSE,
                                  metodo_prop = "likelihood", DEFF = TRUE)
    
    tabla_cruzada_formato <-formatear_tabla_cruzada(tabla_cruzada, datos = datos, dominio = d, DB_Mult = DB_Mult, pregunta = pregunta)
    
    estimaciones_generales <- rbind(estimaciones_generales, tabla_cruzada_formato[[1]])
    errores_generales <- rbind(errores_generales, tabla_cruzada_formato[[2]])
  }
  return(list(estimaciones_generales,errores_generales))
}


tabla_cruzada_final <- tabla_cruzada_por_pregunta(disenio = disenio, pregunta = "P1", 
                                           datos = dataset, dominios = Dominios,
                                           DB_Mult = DB_Mult)



# Escribir tablas cruzadas en Excel

categorias_multiples_formato <- function(pregunta, datos, DB_Mult, metricas = TRUE){
  
  categs <- NULL
  
  ps <- DB_Mult %>% 
    dplyr::filter(!is.na(!!sym(pregunta))) %>% 
    dplyr::pull(!!sym(pregunta))
  
  df <- datos %>% 
    select(ps)
  
  categorias <- df %>% 
    pull() %>% 
    levels()%>% 
    as_tibble() %>% 
    mutate(
      a = NA,
      b = NA
    )
  
  for (i in 1:nrow(categorias)){
    categs <- c(categs, categorias[i,]) %>% unlist() %>% t() %>% as_tibble()
  } 
  
  
  if (metricas){
    
    categs <- NULL
    
    categorias %<>% 
      mutate(c = NA)
    
    for (i in 1:nrow(categorias)){
      categs <- c(categs, categorias[i,]) %>% unlist() %>% t() %>% as_tibble()
    } 
    
  }
  
  return(categs)
  
}



formato_categorias_multiples <- function(pregunta, datos, DB_Mult, wb,
                                         renglon = 2, columna = 4, hojas = c(3,4), estilo_cuerpo){
  
  
  simples <- categorias_multiples_formato(pregunta = pregunta, DB_Mult = DB_Mult,
                                          datos = datos, metricas = FALSE)
  
  escribir_tabla(tabla = simples, wb = wb, hoja = hojas[1], renglon = renglon,
                 columna = columna, bordes = 'surrounding', estilo_borde = 'thin',
                 na = FALSE)
  
  
  for (k in seq(columna, ncol(simples), by=3)){
    mergeCells(wb = wb, sheet = hojas[1], cols = k:(k+2), rows = renglon)
    addStyle(wb = wb, sheet = hojas[1], style = estilo_cuerpo, rows = renglon,
             cols = k:(k+2), stack = TRUE)
  }
  
  
  metricas <- categorias_multiples_formato(pregunta = pregunta, DB_Mult = DB_Mult,
                                           datos = datos, metricas = TRUE)
  
  escribir_tabla(tabla = metricas, wb = wb, hoja = hojas[2], renglon = renglon,
                 columna = columna, bordes = 'surrounding', estilo_borde = 'thin',
                 na = FALSE)
  
  for (k in seq(columna, ncol(metricas), by=4)) {
    mergeCells(wb = wb, sheet = hojas[2], cols = k:(k+3), rows = renglon)
    addStyle(wb = wb, sheet = hojas[2], style = estilo_cuerpo, rows = renglon,
             cols = k:(k+3), stack = TRUE)
  }
  
}


formato_tablas_cruzadas <- function(tabla, wb, renglon, columna = 1, hojas = c(3,4),
                                    estilo_encabezado = headerStyle,
                                    estilo_tabla = horizontalStyle){
  
  
  # renglón 3 columna 1
  escribir_tabla(tabla = tabla[[1]], wb = wb, hoja = hojas[1], renglon = renglon,
                 columna = columna, bordes = 'surrounding', estilo_borde = 'thin',
                 na = TRUE, reemplazo_na = '-', nombres_columnas = TRUE,
                 estilo_encabezado = estilo_encabezado)
  
  
  
  escribir_tabla(tabla = tabla[[2]], wb = wb, hoja = hojas[2], renglon = renglon,
                 columna = columna, bordes = 'surrounding', na = TRUE,
                 reemplazo_na = '-', nombres_columnas = TRUE, estilo_borde = 'thin',
                 estilo_encabezado = estilo_encabezado)
  
  
}

estilo_dominios <- function(tabla, wb, columna = 1, hojas = c(3,4), dominios,
                            renglon, estilo_dominios = horizontalStyle,
                            estilo_merge_dominios = bodyStyle){
  
  # Renglón donde se empieza a escribir la tabla cruzada + 1 para linea horizontal nacional 
  
  r <- renglon + 1
  
  dominios_general <- c('General', dominios)
  
  for (d in dominios_general) {
    tope <- tabla[[1]] %>% select(Dominio) %>% filter(Dominio == d) %>% nrow()
    
    for (i in hojas){
      mergeCells(wb = wb, sheet = i, cols = 1, rows = r:(r+tope-1))
      addStyle(wb = wb, sheet = i, cols = 1, rows = r:(r+tope-1), style = estilo_merge_dominios,
               stack = TRUE)
    }
    
    addStyle(wb = wb, sheet = hojas[1], cols = 1:ncol(tabla[[1]]), rows = (r+tope-1),
             style = estilo_dominios, stack = TRUE)
    
    addStyle(wb = wb, sheet = hojas[2], cols = 1:ncol(tabla[[2]]), rows = (r+tope-1),
             style = estilo_dominios, stack = TRUE)
    
    r <- r + tope
  }
}


estilo_columnas <- function(tabla, wb, hojas = c(3,4), estilo = verticalStyle,
                            renglon){
  
  secuencia_1 <- c(-1,0,1, seq(4, ncol(tabla[[1]]), by=3))
  
  for (k in secuencia_1){
    addStyle(wb = wb, sheet = hojas[1], cols = (k+2),
             rows = (renglon + 2):(renglon + 2 +nrow(tabla[[1]])),
             style = estilo, stack = TRUE)
  }
  secuencia_2 <- c(-2, -1, 0, seq(4, ncol(tabla[[2]]), by=4))
  
  for (k in secuencia_2 ) {
    addStyle(wb = wb, sheet = hojas[2], cols = (k+3), 
             rows = (renglon + 2):(renglon + 2 + nrow(tabla[[1]])),
             style = verticalStyle, stack = TRUE)
  }
  
}

## Estilos

{
  headerStyle <- createStyle(
    fontSize = 11, fontColour = "black", halign = "center",
    border = "TopBottom", borderColour = "black",
    borderStyle = c('thin', 'double'), textDecoration = 'bold')
  
  bodyStyle <- createStyle(halign = 'center', border = "TopBottomLeftRight",
                           borderColour = "black", borderStyle = 'thin',
                           valign = 'top', wrapText = TRUE)
  
  verticalStyle <- createStyle(border = "Right",
                               borderColour = "black", borderStyle = 'thin',
                               valign = 'top')
  
  totalStyle <-  createStyle(numFmt = "###,###,###")
  
  horizontalStyle <- createStyle(border = "bottom",
                                 borderColour = "black", borderStyle = 'thin')
  
  totalStyle <- createStyle(numFmt = "###,###,###")
  
}


wb <- openxlsx::createWorkbook()
options("openxlsx.numFmt" = "0.0")

{
  openxlsx::addWorksheet(wb, sheetName = 'General 1')
  openxlsx::addWorksheet(wb, sheetName = 'General 2')
  openxlsx::addWorksheet(wb, sheetName = 'Dominios 1')
  openxlsx::addWorksheet(wb, sheetName = 'Dominios 2')
}


######################################################################################################

### Función con todos los pasos tablas cruzadas

tablas_cruzadas <- function(pregunta, num_pregunta, datos, hojas = c(3,4),
                            dominios, DB_Mult, lista_preguntas, 
                            disenio, wb = wb, renglon = 1, columna = 1,
                            estilo_cuerpo = estilo_cuerpo, 
                            estilo_columnas = estilo_columnas){
  
  
  # Nombre de pregunta
  nombre_pregunta <- lista_preguntas[[num_pregunta]]
  
  escribir_tabla(tabla = nombre_pregunta, wb = wb, hoja = hojas[1], renglon = renglon,
                 columna = columna, bordes = 'none')
  
  escribir_tabla(tabla = nombre_pregunta, wb = wb, hoja = hojas[2], renglon = renglon,
                 columna = columna, bordes = 'none')
  
  # Tabla cruzada
  tabla_cruzada_final <- tabla_cruzada_por_pregunta(disenio = disenio, pregunta = pregunta, 
                                                    datos = datos, dominios = dominios,
                                                    DB_Mult = DB_Mult)
  
  # Formato categorias en excel
  formato_categorias_multiples(pregunta = pregunta, datos = datos,
                               DB_Mult = DB_Mult, wb = wb, renglon = (renglon+1), 
                               columna = (columna+3), hojas = hojas, 
                               estilo_cuerpo = estilo_cuerpo)
  
  # Escribir tabla cruzada en excel
  formato_tablas_cruzadas(tabla = tabla_cruzada_final, wb=wb, renglon = (renglon+2))
  
  # Formato de dominios en excel
  estilo_dominios(tabla = tabla_cruzada_final, wb=wb, renglon = (renglon+2),
                  dominios = dominios)
  
  # Formato de columnas en excel
  estilo_columnas(tabla = tabla_cruzada_final, hojas = hojas, wb = wb, 
                  estilo = estilo_columnas, renglon = renglon)
  
}

preg <- 'P1'
num_pregunta <- 1


frecuencias_simples(pregunta = preg, num_pregunta = num_pregunta, datos = dataset,
                    DB_Mult = DB_Mult, lista_preguntas = Lista_Preg, diseño = disenio,
                    wb = wb, renglon = 1)


tablas_cruzadas(pregunta = preg, num_pregunta= num_pregunta, datos = dataset, hojas = c(3,4),
                            dominios= Dominios, DB_Mult = DB_Mult, lista_preguntas = Lista_Preg, 
                            disenio = disenio, wb = wb, renglon = 1, columna = 1,
                            estilo_cuerpo = bodyStyle, 
                            estilo_columnas = verticalStyle)


##### Función preguntas múltiples

preguntas_multiples <- function(pregunta, numero_pregunta, datos, lista_preguntas,
                                  dominios, disenio, wb, renglon_tc, renglon_fs, DB_Mult,
                                  columna = 1, hojas_simples = c(1,2), hojas_cruzadas = c(3,4),
                                  estilo_cuerpo = bodyStyle, 
                                  estilo_columnas = verticalStyle,
                                  frecuencias_simples = TRUE, tablas_cruzadas = TRUE){
  
  if (frecuencias_simples){
    
    frecuencias_simples(pregunta = pregunta, num_pregunta = num_pregunta, datos = datos,
                                DB_Mult = DB_Mult, lista_preguntas = Lista_Preg, diseño = disenio,
                                wb = wb, renglon = renglon_fs)
  }
  
  if (tablas_cruzadas){
    
    tablas_cruzadas(pregunta = pregunta, num_pregunta= num_pregunta, datos = datos,
                          hojas = hojas_cruzadas, dominios= dominios, DB_Mult = DB_Mult, 
                          lista_preguntas = lista_preguntas, disenio = disenio, wb = wb, 
                          renglon = renglon_tc, columna = 1, estilo_cuerpo = estilo_cuerpo, 
                          estilo_columnas = estilo_columnas)
  }
  }



preguntas_multiples(pregunta = preg, numero_pregunta = 1, datos = dataset, lista_preguntas = Lista_Preg,
                                dominios = Dominios, disenio = disenio, wb = wb, renglon_tc =1, renglon_fs =1, DB_Mult,
                                columna = 1, hojas_simples = c(1,2), hojas_cruzadas = c(3,4),
                                estilo_cuerpo = bodyStyle, 
                                estilo_columnas = verticalStyle,
                                frecuencias_simples = TRUE, tablas_cruzadas = TRUE)




openxlsx::saveWorkbook(wb, "Prueba.xlsx", overwrite = TRUE)

