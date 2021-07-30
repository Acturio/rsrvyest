{
library(tidyverse)
library(srvyr)
library(survey)
library(magrittr)
library(stringr)
library(tidyr)
library(purrr)
library(ggplot2)
library(openxlsx)
library(xlsx)
}

##################### Función diseño muestreo

disenio_categorico <- function(id, estrato, pesos, datos, pps = "brewer",
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

disenio <- disenio_categorico(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1,
                              reps=FALSE, datos = dataset)

disenio

### Agregar pruebas unitarias librería argument check 

################### Función survey mean y estadísticas

estadisticas_categoricas <- function(diseño, pregunta, na.rm = TRUE, 
 estadisticas = c("se","ci","cv", "var"), significancia = 0.95, proporcion = FALSE, 
 metodo_prop = "likelihood", DEFF = TRUE){
  
  estadisticas <- {{diseño}} %>% 
    srvyr::group_by(!!sym(pregunta)) %>% 
    srvyr::summarize(
      prop = survey_mean(
      na.rm = na.rm, 
      vartype = estadisticas,
      level = significancia,
      proportion = proporcion,
      prop_method = metodo_prop,
      deff = DEFF
      ),
    total = survey_total(
      na.rm = na.rm
     )
  ) %>% 
    mutate(prop_low = ifelse(prop_low < 0, 0, prop_low),
           prop_upp = ifelse(prop_upp > 1, 1, prop_upp),
           !!sym(pregunta) := str_trim(!!sym(pregunta), side = 'both')) %>% 
    dplyr::rename('Respuesta' := !!sym(pregunta))
    

  return(estadisticas)
}

est <- estadisticas_categoricas(diseño = disenio, pregunta = 'P2')
est

################## Función formatear tabla
# Necesitamos pensar en el nombre de esta función

formatear_tabla_categorica <- function(pregunta, datos, tabla){
  
  # row.names(tabla) <- datos %>% 
  #   pull(!!sym(pregunta)) %>% 
  #   levels()
  
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
     select(Respuesta, "Err. Est." , "Var.", "Coef. Var.", "DEFF") 

  return(list(tabla1, tabla2))
}

freq <- formatear_tabla_categorica(pregunta = 'P2', datos = dataset, tabla = est)
freq

################## Función inicializar ínidces

actualizar_indices <- function(k, k1, k2, k3, k4, np, tabla1, tabla2, tabla3, tabla4){
  k = k+1
  # Actualización de índice de renglón para tablas de hoja 1,2 y 3. Se dejan 4 espacios entre tabla y tabla
  k1 = k1 + 1 + nrow(tabla1) + 4
  k2 = k2 + 1 + nrow(tabla2) + 4
  k3 = k3 + 1 + nrow(tabla3) + 4
  k4 = k4 + 1 + nrow(tabla4) + 4
  # Se actualiza número de pregunta
  np = np + 1
  
  return(c(k = k, k1 = k1, k2 = k2, k3 = k3, k4 = k4, np = np))
}

### Función estadisticas por dominios

tablas_cruzadas_categoricas <- function(diseño, pregunta, dominio,
                                        na.rm = TRUE, vartype = c("ci","se","var","cv"),
                                        significancia = 0.95, proporcion = FALSE, 
                                        metodo_prop = "likelihood", DEFF = TRUE){
  
  estadisticas <- diseño %>% 
    srvyr::group_by(!!sym(dominio), !!sym(pregunta), .drop = TRUE) %>% 
    srvyr::summarize(
      prop = survey_mean(
        na.rm = na.rm, 
        vartype = vartype,
        level = significancia,
        proportion = proporcion,
        prop_method = metodo_prop,
        deff = na.rm),
      total = survey_total(
        na.rm = na.rm
      )
    )%>% 
    mutate(prop_low = ifelse(prop_low < 0, 0, prop_low),
           prop_upp = ifelse(prop_upp > 1, 1, prop_upp),
           #Dominios = dominio,
           !!sym(dominio) := str_trim(!!sym(dominio), side = 'both')) 
  
 # Transformar estadísticas a wide

  estadisticas_wide <- estadisticas %>%
    pivot_wider(
      names_from = !!sym(pregunta),
      values_from = c(
        total,
        total_se,
        prop, 
        prop_low,
        prop_upp,
        prop_se,
        prop_cv,
        prop_var,
        prop_deff),
      names_glue = sprintf("{%s}_{.value}", {{pregunta}})) %>% 
    dplyr::rename(Categorias := !!sym(dominio))

  doms <- c(dominio, rep(NA, nrow(estadisticas_wide) - 1))
  
  estadisticas_wide <- bind_cols('Dominio' = doms, estadisticas_wide)

   return(estadisticas_wide)

}

tc <- tablas_cruzadas_categoricas(diseño = disenio, pregunta = 'P2', dominio = 'Sexo')
tc

### Función formatear tablas cruzadas
formatear_tabla_cruzada <- function(pregunta, datos, tabla){
  
  # Auxiliares
  
  categorias <- datos %>% 
    pull(!!sym(pregunta)) %>% 
    levels() 
  
  categorias <- str_trim(categorias, side = 'both')
  
  nombres1 <- c('Dominio', 'Categorías', 'Total', rep(c('Media', 'Lim. inf.', 'Lim. sup.'),
                                        length(categorias)))
  
  nombres2 <- c('Dominio', 'Categorías', 'Total', rep(c("Err. Est.", "Coef. Var.", "Var.", "DEFF"),
                                        length(categorias)))
  
  # Multiplicar por 100 y 10,000
  
   tabla %<>%
    map_if(str_detect(names(tabla), "_prop($|_low$|_upp$|_se$)"), ~.x*100) %>%
    map_if(str_detect(names(tabla), '_prop_var'), ~.x*10000) %>% 
    as_tibble()
  
  tabla %<>%
    rowwise() %>%
    mutate(Total = sum(c_across(ends_with('_total')), na.rm=TRUE)) %>%
     select(Dominio,
            Categorias,
            Total,
            ends_with(c( '_prop', '_prop_low', '_prop_upp',
                         '_prop_se','_prop_var', '_prop_cv','prop_deff')))
    
   
   ### Primera tabla 
   
   tabla1 <- tabla %>% 
     select( Dominio,
             Categorias,
             Total,
             ends_with(c( '_prop', '_prop_low', '_prop_upp')))
   
   for (i in 1:length(categorias)) {
     t1 <- tidyselect::vars_select(names(tabla1), matches(categorias)) %>% as.vector()
   }
   
   tabla_cruzada1 <- tabla1 %>% 
     select(
       Dominio,
       Categorias,
       Total,
       all_of(t1)) %>% 
     dplyr::ungroup()
   
   ## Segunda tabla
   
   tabla2 <- tabla %>% 
     select(
       Dominio,
       Categorias,
       Total,
       ends_with(c('_prop_se', '_prop_cv', '_prop_var','prop_deff')))
   
   for (i in 1:length(categorias)) {
     t2 <- tidyselect::vars_select(names(tabla2), matches(categorias)) %>% as.vector()
   }
  
  tabla_cruzada2 <- tabla2 %>% 
     select(
       Dominio,
       Categorias,
       Total,
       all_of(t2)) %>% 
    dplyr::ungroup()
  
  names(tabla_cruzada1) <- nombres1
  
  names(tabla_cruzada2) <- nombres2

  # 
  return(list(tabla_cruzada1, tabla_cruzada2))
   
  
}

f_tc <- formatear_tabla_cruzada(pregunta = 'P2', datos = dataset, tabla = tc)

f_tc

total_nacional <- function(diseño, pregunta, na.rm = TRUE,
                           estadisticas = c("se","ci","cv", "var"), significancia = 0.95, proporcion = FALSE, 
                           metodo_prop = "likelihood", DEFF = TRUE){
  
  estadisticas <- {{diseño}} %>% 
    srvyr::group_by(!!sym(pregunta)) %>% 
    srvyr::summarize(
      prop = survey_mean(
        na.rm = na.rm, 
        vartype = estadisticas,
        level = significancia,
        proportion = proporcion,
        prop_method = metodo_prop,
        deff = DEFF
      ),
      total = survey_total(
        na.rm = na.rm
      )
    ) %>% 
    mutate(prop_low = ifelse(prop_low < 0, 0, prop_low),
           prop_upp = ifelse(prop_upp > 1, 1, prop_upp),
           !!sym(pregunta) := str_trim(!!sym(pregunta), side = 'both')) 
  
  estadisticas_wide <- estadisticas %>%
    mutate(Dominio = 'General',
           Categorias = 'Total') %>% 
    pivot_wider(
      names_from = !!sym(pregunta),
      values_from = c(
        total,
        total_se,
        prop, 
        prop_low,
        prop_upp,
        prop_se,
        prop_cv,
        prop_var,
        prop_deff),
      names_glue = sprintf("{%s}_{.value}", {{pregunta}}))
   
  return(estadisticas_wide)
}

nacional <- total_nacional(diseño = disenio, pregunta = 'P2')

f_n <- formatear_tabla_cruzada(pregunta = 'P2', datos = dataset, tabla = nacional)

f_n

# Gráficas

## Numero pregunta, añadir total?, valores, categorías.

## Tabla creada en la función estadisticas_categoricas 

graficas <- function(tabla, x, y, pregunta, titulo, subtitulo, leyenda, 
                     y_lab, x_lab){
   grafica <- ggplot(tabla, aes_string(x = x, y = y)) +
  #   # geom_hline(yintercept = pct_gasto_estimado$mean_pct_gasto, 
  #   #            col = "blue", linetype = 'dotted') +
     geom_point(col = "red") +
     geom_errorbar(aes(ymin = !!sym(y) - qnorm(0.975) * prop_se,
                       ymax = !!sym(y) + qnorm(0.975) * prop_se),
                   width = 0.1) + 
       labs(title = titulo,
            subtitle = subtitulo,
            caption = paste(leyenda, pregunta))+
       ylab(y_lab) +
       xlab(x_lab)+
    theme_minimal()
  #theme(axis.text.x = element_text(angle = 45, hjust = 1))
  
  return(grafica)
  
}


graf <- graficas(tabla = est, x = 'P2', y = 'prop', pregunta = Lista_Preg[2],
                titulo = 'Aquí título', subtitulo = 'Aquí subtítulo',
                leyenda = 'Gráfica relacionada a la pregunta:',
                y_lab = 'Aquí y_lab',
                x_lab = 'Aquí x_lab')


graf


# FORS

est <- estadisticas_categoricas(diseño = disenio, pregunta = 'P2')
  
freq <- formatear_tabla_categorica(pregunta = 'P2', datos = dataset, tabla = est)
  
nacional <- total_nacional(diseño = disenio, pregunta = 'P2')

prueba <- formatear_tabla_cruzada(pregunta = 'P2', datos = dataset, tabla = nacional)

doms <- c('Sexo', 'Edad', 'Año_electivo', 'Ocupación_P', 'Ocupación_M')

  for (d in Dominios) {
    tc <- tablas_cruzadas_categoricas(diseño = disenio, pregunta = 'P2' , 
                                      dominio = 'Año_electivo')
    
    f_tc <- formatear_tabla_cruzada(pregunta = 'P2', datos = dataset, 
                                     tabla = tc)
    
    prueba[[1]] <- rbind(prueba[[1]], f_tc[[1]])
    prueba[[2]] <- rbind(prueba[[2]], f_tc[[2]])
    
  }

# Solo es para tabas cruzadas

categorias_pregunta_formato <- function(pregunta, datos, metricas = TRUE){
  
  categs <- NULL
  
  categorias <- datos %>%
    select(!!sym(pregunta)) %>%
    pull()%>%
    levels() %>% 
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

categorias_pregunta_formato(pregunta = 'P2', datos = dataset, metricas = TRUE)

## Estilos

{headerStyle <- createStyle(
  fontSize = 11, fontColour = "black", halign = "center",
  border = "TopBottom", borderColour = "black",
  borderStyle = c('thin', 'double'), textDecoration = 'bold'
)
  
  bodyStyle <- createStyle(halign = "center",
                           border = "TopBottomLeftRight",
                           borderColour = "black", borderStyle = 'thin',
                           valign = 'top')
  
  verticalStyle <- createStyle(halign = "center",
                               border = "Right",
                               borderColour = "black", borderStyle = 'thin',
                               valign = 'top')
  
  bodyStyleCategs <- createStyle(fontSize = 11, fontColour = "black", halign = "center",
                                 textDecoration = 'bold', border = "TopBottom",
                                 borderColour = "black",
                                 borderStyle = c('thin', 'thin'))}


## Escribir en excel

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


wb <- openxlsx::createWorkbook()
options("openxlsx.numFmt" = "0.0")

{
  openxlsx::addWorksheet(wb, sheetName = 'General 1')
  openxlsx::addWorksheet(wb, sheetName = 'General 2')
  openxlsx::addWorksheet(wb, sheetName = 'Dominios 1')
  openxlsx::addWorksheet(wb, sheetName = 'Dominios 2')
}


################## PRIMERA PROPUESTA
# Escribimos primero el nombre de la pregunta en el primer renglón primera columna

np <- Lista_Preg[2]

for (i in 1:4) {
  escribir_tabla(tabla = np, wb = wb, hoja = i, renglon = 1, columna = 1, bordes = 'none')
}

# Escribimos las tablas de fecuencias simples en las hojas 1 y 2 columna 1 renglón 2

escribir_tabla(tabla = freq[[1]], wb = wb, hoja = 1, renglon = 2, columna = 1, 
               nombres_columnas = TRUE, estilo_borde = 'thin',
               estilo_encabezado = headerStyle)

mergeCells(wb = wb, sheet = 1, cols = 1:5, rows = 1)

escribir_tabla(tabla = freq[[2]], wb = wb, hoja = 2, renglon = 2, columna = 1, 
              estilo_borde = 'thin', nombres_columnas = TRUE, estilo_encabezado = headerStyle)

mergeCells(wb = wb, sheet = 2, cols = 1:5, rows = 1)

# Escribimos las categorías en las de la pregunta en el segundo renglón cuarta columna (tablas cruzadas)

simples <- categorias_pregunta_formato(pregunta = 'P2', datos = dataset, metricas = FALSE)

escribir_tabla(tabla = simples, wb = wb, hoja = 3, renglon = 2, columna = 4, 
               bordes = 'surrounding', estilo_borde = 'thin',  na = FALSE)

addStyle(wb = wb, sheet = 3, style = bodyStyleCategs, rows = 2, cols = 4:15)

a <- f_tc[[1]] %>% nrow() -1 + 2 + 2

mergeCells(wb = wb, sheet = 3, cols = 4:6, rows = 2 )
addStyle(wb = wb, sheet = 3, style = bodyStyle, rows = 2, cols = 4:6, stack = TRUE)

mergeCells(wb = wb, sheet = 3, cols = 7:9, rows = 2 )
addStyle(wb = wb, sheet = 3, style = bodyStyle, rows = 2, cols = 7:9, stack = TRUE)

mergeCells(wb = wb, sheet = 3, cols = 10:12, rows = 2 )
addStyle(wb = wb, sheet = 3, style = bodyStyle, rows = 2, cols = 10:12, stack = TRUE)

mergeCells(wb = wb, sheet = 3, cols = 13:15, rows = 2 )
addStyle(wb = wb, sheet = 3, style = bodyStyle, rows = 2, cols = 13:15, stack = TRUE)


for (k in seq(4, 15, by=3)) {
  print(c(k, k+2))
  
   mergeCells(wb = wb, sheet = 3, cols = k:(k+2), rows = 2 )
   addStyle(wb = wb, sheet = 3, style = bodyStyle, rows = 2, cols = k:(k+2), stack = TRUE)
#   
#   
 }

metricas <- categorias_pregunta_formato(pregunta = 'P2', datos = dataset, metricas = TRUE)

escribir_tabla(tabla = metricas, wb = wb, hoja = 4, renglon = 2, columna = 4, 
               bordes = 'surrounding', estilo_borde = 'thin', na = FALSE)

addStyle(wb = wb, sheet = 4, style = bodyStyleCategs, rows = 2, cols = 4:16)


# Escribimos tablas cruzadas en hoja 3 y 4 

escribir_tabla(tabla = f_tc[[1]], wb = wb, hoja = 3, renglon = 3, columna = 1,
               bordes = 'surrounding', estilo_borde = 'thin', na = TRUE, 
               reemplazo_na = '-', nombres_columnas = TRUE,
               estilo_encabezado = headerStyle)


escribir_tabla(tabla = f_tc[[2]], wb = wb, hoja = 4, renglon = 3, columna = 1,
               bordes = 'surrounding', na = TRUE, reemplazo_na = '-', 
               nombres_columnas = TRUE, estilo_borde = 'thin',
               estilo_encabezado = headerStyle)


# Mergeamos celdas repetidas de dominios y Pregunta (renglón 1)

f_tc[[1]] %>% ncol() -1

mergeCells(wb = wb, sheet = 3, cols = 1, rows = 4:8)
mergeCells(wb = wb, sheet = 3, cols = 1:15, rows = 1)
addStyle(wb = wb, sheet = 3, style = bodyStyle,  cols = 1, rows = 4:8)

mergeCells(wb = wb, sheet = 4, cols = 1, rows = 4:8)
mergeCells(wb = wb, sheet = 4, cols = 1:15, rows = 1)
addStyle(wb = wb, sheet = 4, style = bodyStyle,  cols = 1, rows = 4:8)


# Líneas verticales


# Categorias 
 addStyle(wb = wb, sheet = 3, cols = 2, rows = 3:8, style = verticalStyle, stack = TRUE)
 
 # Total
 addStyle(wb = wb, sheet = 3, cols = 3, rows = 3:8, style = verticalStyle, stack = TRUE)
 
# Tablas cruzadas
 
addStyle(wb = wb, sheet = 3, cols = 6, rows = 3:8, style = verticalStyle, stack = TRUE)
addStyle(wb = wb, sheet = 3, cols = 9, rows = 3:8, style = verticalStyle, stack = TRUE)
addStyle(wb = wb, sheet = 3, cols = 12, rows = 3:8, style = verticalStyle, stack = TRUE)
addStyle(wb = wb, sheet = 3, cols = 15, rows = 3:8, style = verticalStyle, stack = TRUE)



openxlsx::openXL(wb)



openxlsx::saveWorkbook(wb, "Prueba.xlsx", overwrite = TRUE)





