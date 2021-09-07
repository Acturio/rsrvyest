
{
library(tidyverse)
library(srvyr)
library(survey)
library(magrittr)
library(reshape2)
library(foreign)
library(readxl)
library(stringr)
library(dplyr)
library(caret)
library(tidyr)
library(purrr)
library(Hmisc)
library(ggplot2)
library(openxlsx)
library(xlsx)
library(reshape)
library(tibble)
}

################################################################################
################################################################################
########################### PREGUNTAS CATEGÓRICAS ##############################
################################################################################
################################################################################

# FUNCIÓN DISEÑO DE MUESTREO

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


# FORMATO CATEGORIAS RESPUESTAS (TABLAS CRUZADAS) 

categorias_pregunta_formato <- function(pregunta, diseño, datos, metricas = TRUE){
  

  categorias <- datos %>%
    select(!!sym(pregunta)) %>%
    pull()%>%
    levels() %>% 
    str_trim(side = 'both')
  
  est <- estadisticas_categoricas(diseño = diseño, datos = dataset,
                                  pregunta = pregunta)
  
  
  categorias_tabla <- est %>% 
    select(Respuesta) %>% 
    as_vector()
  
  categorias_final <- intersect(categorias, categorias_tabla)
  
  categs <- NULL
  
  vector <- categorias_final %>% 
    as_tibble() %>% 
    mutate(
      a = NA,
      b = NA
    )
  
  for (i in 1:nrow(vector)){
    categs <- c(categs, vector[i,]) %>% unlist() %>% t() %>% as_tibble()
  } 
  
  
  if (metricas){
    
    categs <- NULL
    
    vector %<>% 
      mutate(c = NA)
    
    for (i in 1:nrow(vector)){
      categs <- c(categs, vector[i,]) %>% unlist() %>% t() %>% as_tibble()
    } 
    
  }
  
  return(categs)
  
}

# FUNCIÓN ESCRIBIR TABLA EN EXCEL

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

# FUNCIÓN FRECUENCIAS SIMPLES

estadisticas_categoricas <- function(diseño, datos, pregunta, na.rm = TRUE, 
                                     estadisticas = c("se","ci","cv", "var"), 
                                     significancia = 0.95, proporcion = FALSE, 
                                     metodo_prop = "likelihood", DEFF = TRUE){
  
  
  categorias <- datos %>% 
    pull(!!sym(pregunta)) %>% 
    levels() %>% 
    str_trim(side = 'both')
  
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
    dplyr::rename('Respuesta' := !!sym(pregunta)) #%>% 
    #dplyr::filter(Respuesta == categorias)
    

  return(estadisticas)
}


# FUNCIÓN FORMATEAR TABLA DE FRECUENCIAS SIMPLES

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
     select(Respuesta, 'Total', "Err. Est." , "Coef. Var.", "Var.","DEFF")
   
  return(list(tabla1, tabla2))
}


# FUNCIÓN ESCRIBIR FRECUENCIAS SIMPLES EN EXCEL (UTILIZA FUNCIÓN escribir_tabla)

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


# FUNCIÓN INICIALIZAR ÍNDICES 

actualizar_indices <- function(k1, k2, tabla1, tabla2, np){

  k1 = k1 + 1 + nrow(tabla1) + 4
  k2 = k2 + 1 + nrow(tabla2) + 4
  # Se actualiza número de pregunta
  np = np + 1
  
  return(c(k1, k2, np))
}

# FUNCIÓN TABLAS CRUZADAS

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
        deff = DEFF),
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
    mutate(Dominio = dominio) %>% 
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

   return(estadisticas_wide)

}

# FUNCIÓN FORMATEAR TABLA CRUZADA 

formatear_tabla_cruzada <- function(pregunta, datos, tabla){
  
  #Auxiliares
  
  categorias <- datos %>% 
    pull(!!sym(pregunta)) %>% 
    levels() %>% 
    str_trim(side = 'both')
  
  nombres1 <- c('Dominio', 'Categorías', 'Total', rep(c('Media', 'Lim. inf.', 'Lim. sup.'),
                                       length(categorias)))
  
  nombres2 <- c('Dominio', 'Categorías', 'Total', rep(c("Err. Est.", "Coef. Var.", "Var.", "DEFF"),
                                       length(categorias)))
  
  # Orden similar a las categorias
  
  tabla %<>% select(Dominio,Categorias, contains(categorias)) 
  
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
    
  tabla %<>% select(Dominio, Categorias, Total, contains(categorias))
   
   #Primera tabla 
   
   tabla1 <- tabla %>% 
     select( Dominio,
             Categorias,
             Total,
             ends_with(c( '_prop', '_prop_low', '_prop_upp'))) %>% 
     select(Dominio, Categorias, Total, contains(categorias)) %>% 
     dplyr::ungroup()
   
   # for (i in 1:length(categorias)) {
   #   t1 <- tidyselect::vars_select(names(tabla1), matches(categorias)) %>% as.vector()
   # }
   # 
   # tabla_cruzada1 <- tabla1 %>% 
   #   select(
   #     Dominio,
   #     Categorias,
   #     Total,
   #     all_of(t1)) %>% 
   #   dplyr::ungroup()
   
   
   #Segunda tabla
   
   tabla2 <- tabla %>% 
     select(
       Dominio,
       Categorias,
       Total,
       ends_with(c('_prop_se', '_prop_cv', '_prop_var','prop_deff'))) %>% 
     select(Dominio, Categorias, Total, contains(categorias)) %>% 
     dplyr::ungroup()
   
  #  for (i in 1:length(categorias)) {
  #    t2 <- tidyselect::vars_select(names(tabla2), matches(categorias)) %>% as.vector()
  #  }
  # 
  # tabla_cruzada2 <- tabla2 %>% 
  #    select(
  #      Dominio,
  #      Categorias,
  #      Total,
  #      all_of(t2)) %>% 
  #   dplyr::ungroup()
  
  names(tabla1) <- nombres1
  
  names(tabla2) <- nombres2

  return(list(tabla1, tabla2))
   
  
}

# FUNCIÓN TOTAL NACIONAL 

# total_nacional <- function(diseño, datos, pregunta, na.rm = TRUE,
#                            estadisticas = c("se","ci","cv", "var"), 
#                            significancia = 0.95, proporcion = FALSE, 
#                            metodo_prop = "likelihood", DEFF = TRUE){
#   
# 
#   est <- estadisticas_categoricas(diseño = diseño, pregunta = pregunta, datos = datos,
#                                            na.rm = na.rm, estadisticas = estadisticas,
#                                            significancia = significancia,
#                                            proporcion = proporcion, 
#                                            metodo_prop = metodo_prop, DEFF = DEFF)
#   
#   largo <- est %>% select(Respuesta) %>% nrow()
#   
#   
#   nombres_tabla1 <- c('Dominio', 'Categorías', 'Total',
#                       rep(c('Media', 'Lim. inf.', 'Lim. sup.'), largo))
#   
# 
#   nombres_tabla2 <- c('Dominio', 'Categorías', 'Total',
#                       rep(c("Err. Est.", "Coef. Var.","Var.", "DEFF"), largo))
#   
#   f_est <- formatear_tabla_categorica(tabla = est)
#   
#   # Suma total
#   
#   suma_total <- f_est[[1]] %>% select(Total) %>% sum()
#   
#   # Tabla 1
#   # Pasar a columnas tabla 1
#   
#   tabla_desagregada1 <- f_est[[1]][1, c(-1,-2)] %>% 
#     mutate(Dominio = 'General',
#            Categorias = 'Total') 
#   
#   for (i in seq(2,nrow(f_est[[1]]), by = 1)){
#     tabla_c1 = f_est[[1]][i,c(-1,-2)]
#     tabla_desagregada1 <- cbind(tabla_desagregada1,tabla_c1)
#   
#   }
#   
#   # Se pega suma total 
#   
#   tabla_desagregada1 <- cbind(Total = suma_total, tabla_desagregada1)
#   
#   tabla_desagregada1 %<>% 
#       relocate(Categorias,.before = Total) %>% 
#       relocate(Dominio,.before = Categorias)
#   
#  names(tabla_desagregada1) <- nombres_tabla1
#    
#  
#  # Tabla 2 
#  # Pasar a columnas tabla 2
#  
#  tabla_desagregada2 <- f_est[[2]][1, c(-1,-2)] %>% 
#    mutate(Dominio = 'General',
#           Categorias = 'Total') 
#  
#  for (i in seq(2,nrow(f_est[[2]]), by = 1)){
#    tabla_c2 = f_est[[2]][i,c(-1,-2)]
#    tabla_desagregada2 <- cbind(tabla_desagregada2,tabla_c2)
#    
#  }
#  
#  # Se pega suma total 
#  
#  tabla_desagregada2 <- cbind(Total = suma_total, tabla_desagregada2)
#  
#  tabla_desagregada2 %<>% 
#    relocate(Categorias,.before = Total) %>% 
#    relocate(Dominio,.before = Categorias)
#  
#  names(tabla_desagregada2) <- nombres_tabla2
#  
#   return(list(tabla_desagregada1, tabla_desagregada2))
# }


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

# FUNCIÓN TABLA CRUZADA TOTAL NACIONAL Y POR DOMINIOS 

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


# FUNCIÓN FORMATO CATEGORÍAS PARA TABLAS CRUZADAS 
# UTILIZA LAS FUNCIONES escribir_tabla Y categorias_pregunta_formato

formato_categorias <- function(tabla, pregunta, datos, diseño, wb, renglon, columna = 4,
                                       hojas = c(3,4), estilo_cuerpo){
          
          # Renglón 2 columna 4
  
          simples <- categorias_pregunta_formato(pregunta = pregunta, datos = datos,
                                                 diseño = diseño, metricas = FALSE)

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
                                                  diseño = diseño, metricas = TRUE)
                                

          escribir_tabla(tabla = metricas, wb = wb, hoja = hojas[2], renglon = renglon,
                         columna = columna, bordes = 'surrounding', estilo_borde = 'thin',
                         na = FALSE)
          
          for (k in seq(columna, ncol(tabla[[2]]), by=4)) {
            mergeCells(wb = wb, sheet = hojas[2], cols = k:(k+3), rows = renglon)
            addStyle(wb = wb, sheet = hojas[2], style = estilo_cuerpo, rows = renglon,
                     cols = k:(k+3), stack = TRUE)
          }
          
        }

# FUNCIÓN FORMATO TABLAS CRUZADAS, ESCRIBE LAS TABLAS CRUZADAS EN LAS HOJAS INDICADAS
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

# FUNCIÓN ESTILO DOMINIOS, DA FORMATO A LA COLUMNA DE DOMINIOS

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

# FUNCIÓN ESTILO COLUMNAS, DIVIDE LAS MÉTRICAS DE LAS TABLAS CRUZADAS POR CATEGORÍAS RESPUESTAS

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


# FUNCIÓN GRÁFICAS 

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

# FUNCIÓN FRECUENCIAS SIMPLES, DESDE CREAR LA TABLA, FORMATEARLA, HASTA INSERTARLA EN EXCEL

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
  
  escribir_tabla(tabla = np, wb = wb, hoja = hojas[1], renglon = renglon,
                 columna = columna,  bordes = 'none')

  escribir_tabla(tabla = np, wb = wb, hoja = hojas[1], renglon = renglon, columna = columna)
  

  escribir_tabla(tabla = np, wb = wb, hoja = hojas[2], renglon = renglon, columna = columna)

  escribir_tabla(tabla = np, wb = wb, hoja = hojas[2], renglon = renglon,
                 columna = columna, bordes = 'none')
  
  est <- estadisticas_categoricas(diseño = diseño, pregunta = pregunta, datos = datos)
  
  freq <- formatear_tabla_categorica(tabla = est)
  
  
  formato_frecuencias_simples(tabla = freq, wb = wb, renglon = (renglon + 1), 
                              columna = columna, estilo_encabezado = estilo_encabezado)
  
  return(freq)
  
}


# FUNCIÓN TABLAS CRUZADAS, DESDE CREAR LA TABLA CRUZADA, FORMATEARLA, HASTA INSERTARLA EN EXCEL

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
                          estilo_tabla = horizontalStyle,
                          estilo_columnas = verticalStyle){
  
  # Título pregunta
  
  np <- lista_preguntas[[num_pregunta]]
  
  escribir_tabla(tabla = np, wb = wb, hoja = hojas[1], renglon = renglon,
                 columna = columna, bordes = 'none')
  
  escribir_tabla(tabla = np, wb = wb, hoja = hojas[2], renglon = renglon,
                 columna = columna, bordes = 'none')
  
  # Tabla Cruzada 
  
  f <- tabla_cruzada_total(diseño = diseño, pregunta = pregunta, datos = datos,
                           dominios = dominios)
  

  # Formato Categorías empieza en el renglón 2
  
  formato_categorias(tabla = f, pregunta = pregunta, datos = datos, dis = diseño,
                     wb = wb, renglon = (renglon + 1), columna = (columna + 3),
                     hojas = hojas, estilo_cuerpo = estilo_categorias)

  
  # Escribir tabla Cruzada empieza en el renglón 3

  formato_tablas_cruzadas(tabla = f, wb = wb, renglon = (renglon + 2),
                          columna = columna, hojas = hojas,
                          estilo_encabezado = estilo_encabezado,
                          estilo_tabla = estilo_tabla)
  
  # Estilo dominios empieza en el renglón 3
  
  estilo_dominios(tabla = f, wb = wb, columna = columna, dominios = Dominios,
                  hojas = hojas, renglon = (renglon + 2))
  
  
  # Estilo columnas 
  
  estilo_columnas(tabla = f, hojas = hojas, wb = wb, estilo = estilo_columnas, 
                  renglon = renglon)

  return(f)
  
}


# FUNCIÓN PREGUNTAS CATEGÓRICAS, SE INDICA SI SE DESEAN ÚNICAMENTE LAS FRECUENCIAS SIMPLES
# O TABLAS CRUZADAS O AMBAS (FUNCIÓN FINAL)

preguntas_categoricas <- function(pregunta, numero_pregunta, datos, lista_preguntas,
                                  dominios, diseño, wb, renglon_tc, renglon_fs,
                                  columna = 1, hojas_simples = c(1,2), hojas_cruzadas = c(3,4), 
                                  frecuencias_simples = TRUE, tablas_cruzadas = TRUE){
  
  if (frecuencias_simples){
    
    freq <- frecuencias_simples(pregunta = pregunta, num_pregunta = numero_pregunta,
                        lista_preguntas = lista_preguntas, datos = datos,
                        diseño = diseño, wb = wb, hojas = hojas_simples,
                        renglon = renglon_fs, columna = columna)
  }
  
  if (tablas_cruzadas){
    
    tc <- tabla_cruzada(pregunta = pregunta, num_pregunta = numero_pregunta,
                        lista_preguntas = lista_preguntas, dominios = dominios, 
                         datos = datos, diseño = diseño, wb = wb,
                        renglon = renglon_tc , columna = columna, hojas = hojas_cruzadas)
  }
  
  return(list(freq, tc))
}

################################################################################
################################################################################
########################### PREGUNTAS NUMÉRICAS ################################
################################################################################
################################################################################

# FUNCIÓN LEER DATOS
# argumentos base (string): nombre del archivo con extención ("BASE_CONACYT_260118.sav")
# Lista de preguntas (string): nombre del archivo con extensión ("Lista de Preguntas.xlsx")             

leer_datos <- function(base, lista){
  #se asume misma organización de carpetas 
  archivo <- paste0("data/", base)
  archivo2 <- paste0("aux/", lista) # antes "lista/"
  
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


# FUNCIÓN CREAR DISEÑO 
# función para crear diseño muestral, los argumentos con data, id, estrato, y pesos.
# archivo <- paste0("data/", base)
# dataset <- read.spss(archivo, to.data.frame = TRUE) 

crea_disenio <- function(data, id, cestrato, cpesos){
  
  disenio <- data %>% 
    as_survey_design(
      ids = {{id}},
      weights = {{cpesos}},
      strata = {{cestrato}}
    )
  
  return(disenio)
}

# FUNCIÓN ESTADÍSTICAS CONTINUAS 
# función para crear df de frecuencias para variables continuas

estadisticas_continuas <- function(disenio, pregunta, na.rm = TRUE,
                                   vartype = c("se", "ci", "cv", "var"), 
                                   level = 0.95, proportion = FALSE, prop_method = "likelihood",
                                   DEFF = TRUE, cuantiles = c(0,0.25, 0.5, 0.75,1)) {
  
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

# FUNCIÓN ACOMODA FRECUENCIAS
# función para acomodar df de frecuencias para variables continuas

acomoda_frecuencias <- function(df){
  df_t <- df %>% 
    reshape2::melt()
  names(df_t) <- c("stat", "valor")
  
  lvars <- c("media", "lim_inf", "lim_sup", "mín", "Q25", "mediana", "Q75", "máx", 
             "sd", "var", "cv", "deff")
  
  lvars <- c("Media", "Lim_inf", "Lim_sup", "Mín", "Q25", "Mediana", "Q75", "Máx", 
             "Sd", "Var", "C.V.", "Deff")
  
  nvo_df <- data.frame(lvars, round(df_t[,2],3))
  names(nvo_df) <- c("Métrica", "Valor")
  
  return(nvo_df)
  
}

# FUNCIÓN TOTAL
# función para calcular estadisticas del total poblacional, la cual es complemento de la 
# función tabla_cruzada para generar la tabla cruzada final

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

# FUNCIÓN TABLA CRUZADA
# función para cruces con dominios

tabla_cruzada_cont <- function(disenio, pregunta, dominio, na.rm = TRUE,
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

# FORMATO TABLA
# si se elige tabla con media o limites o si se eligen las otras variables.

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

# FUNCIÓN FRECUENCIAS EN EXCEL 

tabla_frec_excel <- function(df, colini, rowini){
  hs1 <- createStyle(halign = "CENTER", textDecoration = "Bold",
                     border = "TopBottomLeftRight", fontColour = "black",
                     borderStyle = "medium", borderColour = "black")
  s <- createStyle(numFmt = "0", valign = "center")
  
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
  addStyle(wb, 1, style = s, rows = ri:rf, cols = 3, stack = T, gridExpand = T)
  #formato interior
  c=0
  finicio = rowini
  
  return(openXL(wb))  
}

# TABLA EXCEL TABLAS CRUZADAS CONTINUAS

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

################################################################################
################################################################################
########################### PREGUNTAS MÚLTIPLES ################################
################################################################################
################################################################################


# FUNCIÓN DISEÑO
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



# FRECUENCIAS SIMPLES PARA PREGUNTAS MÚLTIPLES

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
  
  frecuencias_simples = tibble()
  
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

# FORMATEAR TABLAS FRECUENCIAS MÚLTIPLES

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

# FUNCIÓN FRECUENCIAS SIMPLES COMPLETA

frecuencias_simples_mult <- function(pregunta, num_pregunta, datos, lista_preguntas,
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

# TABLAS CRUZADAS MÚLTIPLES

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

# FUNCIÓN FORMATEAR TABLA CRUZADA MÚLTIPLES

formatear_tabla_cruzada_multiples <- function(tabla, datos, dominio, DB_Mult, pregunta){
  
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

# FUNCIONES TOTAL NACIONAL

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

# FUNCIÓN FINAL TABLA CRUZADA MÚLTIPLES 

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
    
    tabla_cruzada_formato <-formatear_tabla_cruzada_multiples(tabla_cruzada, datos = datos, dominio = d, DB_Mult = DB_Mult, pregunta = pregunta)
    
    estimaciones_generales <- rbind(estimaciones_generales, tabla_cruzada_formato[[1]])
    errores_generales <- rbind(errores_generales, tabla_cruzada_formato[[2]])
  }
  
  return(list(estimaciones_generales,errores_generales))
}

# FUNCIÓN CATEGORÍAS MÚLTIPLES

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

# FORMATO CATEGORÍAS MÚLTIPLES

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

# FUNCIÓN CON TODOS LOS PASOS TABLAS CRUZADAS MÚLTIPLES

tablas_cruzadas_mult <- function(pregunta, num_pregunta, datos, hojas = c(3,4),
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

# PREGUNTAS MÚLTIPLES

preguntas_multiples <- function(pregunta, numero_pregunta, datos, lista_preguntas,
                                dominios, disenio, wb, renglon_tc, renglon_fs, DB_Mult,
                                columna = 1, hojas_simples = c(1,2),
                                hojas_cruzadas = c(3,4), estilo_cuerpo = bodyStyle, 
                                estilo_columnas = verticalStyle,
                                frecuencias_simples = TRUE, tablas_cruzadas = TRUE){
  
  if (frecuencias_simples){
    
    frecuencias_simples_mult(pregunta = pregunta, num_pregunta = numero_pregunta, datos = datos,
                        DB_Mult = DB_Mult, lista_preguntas = Lista_Preg, diseño = disenio,
                        wb = wb, renglon = renglon_fs)
  }
  
  if (tablas_cruzadas){
    
    tablas_cruzadas_mult(pregunta = pregunta, num_pregunta= numero_pregunta, datos = datos,
                    hojas = hojas_cruzadas, dominios= dominios, DB_Mult = DB_Mult, 
                    lista_preguntas = lista_preguntas, disenio = disenio, wb = wb, 
                    renglon = renglon_tc, columna = 1, estilo_cuerpo = estilo_cuerpo, 
                    estilo_columnas = estilo_columnas)
  }
}