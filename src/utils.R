
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

### Funciones categorias respuestas Solo es para tabas cruzadas

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

formatear_tabla_categorica <- function(tabla){
  
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
     select(Respuesta, "Err. Est." , "Coef. Var.", "Var.","DEFF") 

  return(list(tabla1, tabla2))
}

freq <- formatear_tabla_categorica(tabla = est)
freq


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
      total = round(survey_total(
        na.rm = na.rm
      ), 0)
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

  doms <- rep(dominio, nrow(estadisticas_wide))
  
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
      total = round(survey_total(
        na.rm = na.rm
      ),0)
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

###### Funcion completa tabla cruzada pega dominios con total nacional 

tabla_cruzada_total <- function(diseño, pregunta, datos, dominios = Dominios){
  
  # Tabla nacional
  nacional <- total_nacional(diseño = diseño, pregunta = pregunta)
  
  f_n <- formatear_tabla_cruzada(pregunta = pregunta, datos = datos,
                                 tabla = nacional)
  
  # Dominios de dispersión
  
  for (d in dominios) {
    tc <- tablas_cruzadas_categoricas(diseño = diseño, pregunta = pregunta, 
                                      dominio = d)
    
    f_tc <- formatear_tabla_cruzada(pregunta = pregunta, datos = datos, 
                                    tabla = tc)
    
    f_n[[1]] <- rbind(f_n[[1]], f_tc[[1]])
    f_n[[2]] <- rbind(f_n[[2]], f_tc[[2]])
    
    
  }
  
  return(f_n)
}


#######  Formato Categorías

        formato_categorias <- function(tabla, pregunta, datos, wb, renglon, columna = 4,
                                       hojas = c(3,4), estilo_cuerpo){
          
          # Renglón 2 columna 4
          
          simples <- categorias_pregunta_formato(pregunta = pregunta, datos = datos,
                                                 metricas = FALSE)
          
          escribir_tabla(tabla = simples, wb = wb, hoja = hojas[1], renglon = renglon,
                         columna = columna, bordes = 'surrounding', estilo_borde = 'thin',
                         na = FALSE)
          
          for (k in seq(columna, ncol(tabla[[1]]), by=3)){
            
            mergeCells(wb = wb, sheet = hojas[1], cols = k:(k+2), rows = renglon)
            addStyle(wb = wb, sheet = hojas[1], style = estilo_cuerpo, rows = renglon,
                     cols = k:(k+2), stack = TRUE)
          }
          
          
          metricas <- categorias_pregunta_formato(pregunta = pregunta, datos = datos,
                                                  metricas = TRUE)
          
          escribir_tabla(tabla = metricas, wb = wb, hoja = hojas[2], renglon = renglon,
                         columna = columna, bordes = 'surrounding', estilo_borde = 'thin',
                         na = FALSE)
          
          for (k in seq(columna, ncol(tabla[[2]]), by=4)) {
            mergeCells(wb = wb, sheet = hojas[2], cols = k:(k+3), rows = renglon)
            addStyle(wb = wb, sheet = hojas[2], style = estilo_cuerpo, rows = renglon,
                     cols = k:(k+3), stack = TRUE)
          }
          
        }

# Formato Tablas Cruzadas 

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


###############################

# Merge Dominios y Líneas Horizontales

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

######### Líneas Verticales

estilo_columnas <- function(tabla, wb, hojas = c(3,4), estilo = verticalStyle){
  
  secuencia_1 <- c(-1,0,1, seq(4, ncol(tabla[[1]]), by=3))
  
  for (k in secuencia_1){
    addStyle(wb = wb, sheet = hojas[1], cols = (k+2), rows = 3:(3+nrow(tabla[[1]])),
             style = estilo, stack = TRUE)
  }
  secuencia_2 <- c(-2, -1, 0, seq(4, ncol(tabla[[2]]), by=4))
  
  for (k in secuencia_2 ) {
    addStyle(wb = wb, sheet = hojas[2], cols = (k+3), rows = 3:(3+nrow(tabla[[1]])),
             style = verticalStyle, stack = TRUE)
  }
  
}



##### Gráficas

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


graf <- graficas(tabla = est, x = 'Respuesta', y = 'prop', pregunta = Lista_Preg[2],
                titulo = 'Aquí título', subtitulo = 'Aquí subtítulo',
                leyenda = 'Gráfica relacionada a la pregunta:',
                y_lab = 'Aquí y_lab',
                x_lab = 'Aquí x_lab')


graf



## Estilos

{
  headerStyle <- createStyle(
  fontSize = 11, fontColour = "black", halign = "center",
  border = "TopBottom", borderColour = "black",
  borderStyle = c('thin', 'double'), textDecoration = 'bold')
  
  bodyStyle <- createStyle(halign = 'center', border = "TopBottomLeftRight",
                           borderColour = "black", borderStyle = 'thin',
                           valign = 'top')
  
  verticalStyle <- createStyle(border = "Right",
                               borderColour = "black", borderStyle = 'thin',
                               valign = 'top')
  
  horizontalStyle <- createStyle(border = "bottom",
                                 borderColour = "black", borderStyle = 'thin')
  
  totalStyle <- createStyle(numFmt = "###,###,###")

  }


##############################################################################
##############################################################################
################################ FRECUENCIAS SIMPLES #########################
##############################################################################
##############################################################################

# Renglón 1, columna 1 siempre

frecuencias_simples <- function(pregunta, num_pregunta, datos, lista_preguntas,
                                diseño, wb, renglon, columna = 1, hojas = c(1,2),
                                estilo_encabezado = headerStyle){
  
  # Título Pregunta
  
  np <- lista_preguntas[[num_pregunta]]
  
  escribir_tabla(tabla = np, wb = wb, hoja = hojas[1], renglon = renglon, columna = columna)
  

  escribir_tabla(tabla = np, wb = wb, hoja = hojas[2], renglon = renglon, columna = columna)
  
  
  est <- estadisticas_categoricas(diseño = diseño, pregunta = pregunta)
  
  freq <- formatear_tabla_categorica(tabla = est)
  
  
  formato_frecuencias_simples(tabla = freq, wb = wb, renglon = (renglon + 1), 
                              columna = columna, estilo_encabezado = estilo_encabezado)
  
}

frecuencias_simples(pregunta = 'P2', num_pregunta = 2,
                    lista_preguntas = Lista_Preg, hojas = c(1,2), datos = dataset, 
                    diseño = disenio, wb = wb, renglon = 1)

################################################################################
################################################################################
################################## TABLAS CRUZADAS #############################
################################################################################
################################################################################

#renglón 1 columna 1 siempre

tabla_cruzada <- function(pregunta, num_pregunta, lista_preguntas, dominios, 
                          datos, diseño, wb, renglon, columna = 1, hojas = c(3,4),
                          estilo_encabezado = headerStyle, 
                          estilo_categorias = bodyStyle,
                          estilo_tabla = horizontalStyle
                          ){
  
  # Título pregunta
  
  np <- lista_preguntas[[num_pregunta]]
  
  escribir_tabla(tabla = np, wb = wb, hoja = hojas[1], renglon = renglon, columna = columna)
  
  escribir_tabla(tabla = np, wb = wb, hoja = hojas[2], renglon = renglon, columna = columna)
  
  # Tabla Cruzada 
  
  f <- tabla_cruzada_total(diseño = diseño, pregunta = pregunta, datos = datos,
                           dominios = dominios)

  # Formato Categorías empieza en el renglón 2
  
  formato_categorias(tabla = f, pregunta = pregunta, datos = datos, wb = wb,
  renglon = (renglon + 1), columna = (columna + 3), hojas = hojas,
  estilo_cuerpo = estilo_categorias)

  
  # Escribir tabla Cruzada empieza en el renglón 3

  formato_tablas_cruzadas(tabla = f, wb = wb, renglon = (renglon + 2),
                          columna = columna, hojas = hojas,
                          estilo_encabezado = estilo_encabezado,
                          estilo_tabla = estilo_tabla)
  
  # Estilo dominios empieza en el renglón 3
  
  estilo_dominios(tabla = f, wb = wb, columna = columna, dominios = Dominios,
                  hojas = hojas, renglon = (renglon + 2))
  
  
  # Estilo columnas 
  
  estilo_columnas(tabla = f, hojas = hojas ,wb = wb)

  
}


tabla_cruzada(pregunta = preg, lista_preguntas = Lista_Preg, num_pregunta = 2,
              dominios = Dominios, datos = dataset, hojas = c(3,4),
              diseño = disenio, wb = wb, renglon = 1, columna = 1)



################################################################################
################################################################################
########################## PREGUNTAS CATEGÓRICAS ###############################
################################################################################
################################################################################

preguntas_categoricas <- function(pregunta, numero_pregunta, datos, lista_preguntas,
                                  dominios, diseño, wb, renglon, columna,
                                  hojas_simples = c(1,2), hojas_cruzadas = c(3,4), 
                                  frecuencias_simples = TRUE, tablas_cruzadas = TRUE){
  
  if (frecuencias_simples){
    
    frecuencias_simples(pregunta = pregunta, num_pregunta = numero_pregunta,
                        lista_preguntas = lista_preguntas, datos = datos,
                        diseño = diseño, wb = wb, hojas = hojas_simples,
                        renglon = renglon)
    
  }
  
  if (tablas_cruzadas){
    
    tabla_cruzada(pregunta = pregunta, lista_preguntas = lista_preguntas,
                  num_pregunta = numero_pregunta, dominios = dominios,
                  datos = datos, diseño = diseño, wb = wb, hojas = hojas_cruzadas,
                  renglon = renglon, columna = columna)
    
  }
  
  
  
}


###### Prueba 

wb <- openxlsx::createWorkbook()
options("openxlsx.numFmt" = "0.0")

{
  openxlsx::addWorksheet(wb, sheetName = 'General 1')
  openxlsx::addWorksheet(wb, sheetName = 'General 2')
  openxlsx::addWorksheet(wb, sheetName = 'Dominios 1')
  openxlsx::addWorksheet(wb, sheetName = 'Dominios 2')
}


## Dinámico: pregunta, numero de pregunta, renglon

preguntas_categoricas(pregunta = 'P2', numero_pregunta = 2, datos = dataset,
                      lista_preguntas = Lista_Preg, dominios = Dominios,
                      diseño = disenio, wb = wb, renglon = 1, columna = 1,
                      frecuencias_simples = TRUE, tablas_cruzadas = TRUE)

# Estilo columna total 

addStyle(wb = wb, sheet = 1, style = totalStyle, rows = 3:100000, cols = 2 ,
         gridExpand = TRUE, stack = TRUE)

addStyle(wb = wb, sheet = 3, style = totalStyle, rows = 3:100000, cols = 3,
         gridExpand = TRUE, stack = TRUE)

openXL(wb)

openxlsx::saveWorkbook(wb, "Prueba.xlsx", overwrite = TRUE)

