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
  library(ggplot2)
  library(openxlsx)
  library(xlsx)
  library(reshape)
  library(tibble)
}

# FUNCIÓN DISEÑO DE MUESTREO 

disenio <- function(id, estrato, pesos, datos, pps = "brewer",
                      varianza = "HT", reps = TRUE, metodo = "subbootstrap",
                      B=50, semilla=1234){
    
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
  
#################################################################################
#################################################################################
############################# FRECUENCIAS SIMPLES ##############################
#################################################################################
#################################################################################

# FUNCIÓN FRECUENCIAS SIMPLES SEGÚN TIPO DE PREGUNTA

#' Crea tabla frecuencias simples
#' @description Se crean las frecuencias simples según tipo de pregunta (categórica, múltiple o continua)
#' @usage frecuencias_simples(
#' diseño,
#' datos,
#' pregunta,
#' DB_Mult,
#' na.rm = TRUE,
#' estadisticas =  c("se","ci","cv", "var"), 
#' cuantiles = c(0,0.25, 0.5, 0.75,1),
#' significancia = 0.95, 
#' proporcion = FALSE, 
#' metodo_prop = "likelihood", DEFF = TRUE,
#' tipo_pregunta = "categorica"
#' )
#' @param diseño Diseño muestral que se ocupará según el tipo de pregunta
#' @param datos Conjunto de datos en formato .sav
#' @param pregunta Pregunta de la cual se quieren obtener las frecuencias simples, por ejemplo, 'P_1'
#' @param DB_Mult Data frame con las preguntas múltiples
#' @param na.rm Valor lógico que indica si se deben de omitir valores faltantes
#' @param estadisticas Métricas de variabilidad: error estándar ("se"), intervalo de confianza ("ci"), varianza ("var") o coeficiente de variación ("cv")
#' @param cuantiles Vector de cuantiles a calcular
#' @param significancia Nivel de confianza: 0.95 por default
#' @param proporcion Valor lógico que indica si se desen usar métodos para calcular la proporción que puede tener intervalos de confianza más precisos cerca de 0 y 1
#' @param metodo_prop  Si proporcion = TRUE; tipo de método de proporción que se desea usar: "logit", "likelihood", "asin", "beta", "mean"
#' @param DEFF Valor lógico que indica si se desea calcular el efecto de diseño
#' @param tipo_pregunta Tipo de pregunta: 'categorica', 'multiple', 'continua'
#' @return Tabla tipo tibble con las estadísticas especificadas en el parámetro estadisticas por respuestas pertenecientes a la pregunta especificada en el parámetro pregunta
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{survey_mean}}
#' @example \dontrun{
#' frecuencias_simples(diseño = disenio_cat, datos = dataset, pregunta = 'P1',
#'  DB_Mult = DB_Mult, tipo_pregunta = 'categorica')
#' } 

frecuencias_simples <-  function(diseño, datos, pregunta, DB_Mult, na.rm = TRUE, 
                                 estadisticas = c("se","ci","cv", "var"), 
                                 cuantiles = c(0,0.25, 0.5, 0.75,1),
                                 significancia = 0.95, proporcion = FALSE, 
                                 metodo_prop = "likelihood", DEFF = TRUE,
                                 tipo_pregunta = "categorica"){
  
    if(tipo_pregunta == 'categorica'){
      
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
        dplyr::rename('Respuesta' := !!sym(pregunta)) 
    }
    
    if (tipo_pregunta == 'continua'){
      
      estadisticas <- {{diseño}} %>% 
        srvyr::summarise(
          prop = survey_mean(
            as.numeric(!!sym(pregunta)),
            na.rm = na.rm, 
            vartype = estadisticas,
            level = significancia,
            proportion = proporcion,
            prop_method = metodo_prop,
            deff = DEFF
          ),
          cuantiles = survey_quantile(
            as.numeric(!!sym(pregunta)),
            quantiles = cuantiles,
            na.rm = na.rm
          ),
          total = survey_total(
            na.rm = na.rm)
        ) %>% 
        select(total, prop, prop_low, prop_upp, cuantiles_q00, cuantiles_q25,
               cuantiles_q50, cuantiles_q75, cuantiles_q100, prop_se, prop_var,
               prop_cv, prop_deff)
      
    }
    
    if (tipo_pregunta == 'multiple'){
      
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
        diseño %<>% srvyr::mutate(!!sym(i) := variable)
      }
      
      frecuencias_simples = data_frame()
      
      ### Cálculo de frecuencias simples de todas las categorías de una pregunta
      
      for (categ in categorias) {
        nacional <- diseño %>% 
          srvyr::summarize(
            prop = survey_mean(!!sym(categ),
                               na.rm = na.rm,
                               vartype = c("se", "ci", "cv", "var"),
                               level = significancia,
                               proportion = proporcion,
                               prop_method = metodo_prop,
                               deff  = DEFF),
            total = survey_total(!!sym(categ),
                                 na.rm = na.rm)) %>% 
          mutate(prop_low = ifelse(prop_low < 0, 0, prop_low),
                 prop_upp = ifelse(prop_upp > 1, 1, prop_upp),
                 Respuesta = categ) 
        
        frecuencias_simples <- bind_rows(frecuencias_simples, nacional)
        
        estadisticas <- frecuencias_simples
        
      }
      
    }
  
    estadisticas %<>% mutate(
      prop_cv = ifelse(is.nan(prop_cv), NA, prop_cv),
      prop_deff = ifelse(is.nan(prop_deff), NA, prop_deff)
    ) 
    return(estadisticas)
  }
  

# FUNCIÓN FORMATEAR TABLA FRECUENCIAS SIMPLES SEGÚN TIPO DE PREGUNTA 

#'  Formatea tabla frecuencias simples 
#'  
#' @description Se formatea la tabla de frecuencias simples según tipo de pregunta 
#' @usage formatear_frecuencias_simples(
#' tabla,
#' tipo_pregunta = "categorica"
#' )
#' @param tabla Tabla de frecuencias simples creada con la función frecuencias_simples()
#' @param tipo_pregunta Tipo de pregunta: 'categorica', 'multiple', 'continua'
#' @return Lista de dos tibbles: en la primera tibble se encuentra el total estimado y las métricas media, límite inferior y superior; 
#' en la segunda tibble se encuentra el total y las métricas error estándar, varianza, coeficiente de variación y el efecto de diseño (si es que DEFF = TRUE en la función frecuencias_simples)
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{survey_mean}}
#' @example \dontrun{
#' freq_simples <- frecuencias_simples(diseño = disenio_cat, datos = dataset, pregunta = 'P1',
#'  DB_Mult = DB_Mult, tipo_pregunta = 'categorica') 
#'  
#' frecuencias_simples(tabla = freq_simples, tipo_pregunta = 'categorica')
#' } 


formatear_frecuencias_simples <- function(tabla,
                                          tipo_pregunta = 'categorica'){
  
  if (tipo_pregunta == 'categorica' || tipo_pregunta == 'multiple'){
    
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
  
  }
  
  if (tipo_pregunta == 'continua'){
    
    tabla1 <- tabla %>% 
      select(total, prop, prop_low, prop_upp, cuantiles_q00, cuantiles_q25,
             cuantiles_q50, cuantiles_q75, cuantiles_q100) %>% 
      dplyr::rename('Total' = total,
             'Media' = prop,
             'Lim. inf' = prop_low,
             'Lim. sup' = prop_upp,
             'Mín' = cuantiles_q00,
             'Q25' = cuantiles_q25,
             'Mediana' = cuantiles_q50,
             'Q75' = cuantiles_q75,
             'Máx' = cuantiles_q100) %>% 
      reshape2::melt()
    
    names(tabla1) <- c("Métrica", "Valor")
    
    tabla2 <- tabla %>% 
      select(total, prop_se, prop_var, prop_cv, prop_deff) %>% 
      dplyr::rename( "Err. Est." = prop_se,
                     "Total" = total,
                     "Var" = prop_var,
                     "Coef. Var." = prop_cv,
                     "DEFF" = prop_deff) %>% 
      reshape2::melt()

    names(tabla2) <- c("Métrica", "Valor")
    

  }
  return(list(tabla1, tabla2))
}

# FUNCIÓN ESCRIBIR FRECUENCIAS SIMPLES EN EXCEL (UTILIZA FUNCIÓN escribir_tabla)

#' Escribe frecuencias simples en dos hojas de Excel  
#'
#' @description Escribe la tabla de frecuencias simples formateadas en dos hojas de Excel
#' @usage formato_frecuencias_simples(
#' tabla,
#' wb,
#' hojas,
#' renglon,
#' columna,
#' estilo_encabezado,
#' estilo_horizontal,
#' estilo_total,
#' tipo_pregunta
#' )
#' @param tabla lista de las tibbles creadas por la función formatear_frecuencias_simples
#' @param wb Workbook de Excel que contiene al menos dos hojas 
#' @param hojas vector de número de hojas en el cual se desea insertar las tablas
#' @param renglon vetor tamaño 2 especificando el número de renglon en el cual se desea empezar a escribir la tabla 1 y tabla 2 respectivamente
#' @param columna columna en la cual se desea empezar a escribir las tablas 
#' @param estilo_encabezado estilo el cual se desea usar para los nombres de las columnas
#' @param estilo_horizontal estilo último renglón horizontal
#' @param estilo_total estilo el cual se desea usar para la columna total
#' @param tipo_pregunta tipo de pregunta_ 'categorica', 'multiple', 'continua'
#' @details Ambas tablas se escribirán en la misma columna pero en diferentes hojas (especificadas en el parámetro hojas).
#' @details El estilo_total se recomienda crear un estilo con la función createStyle de openxlsx con el formato que se desea, por ejemplo "###,###,###.0"
#' @details El estilo_horizontal hace referencia al tipo de lineado horizontal se desea en el úntimo renglón de la tabla
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{writeData}} \code{\link{createStyle}} 
#' @examples \dontrun{
#' # Creación del workbook
#' wb <- createWorkbook()
#' addWorksheet(wb, "Frecuencias simples")
#' addWorksheet(wb, "Frecuencias simples (dispersión)")
#' 
#' # Estilos 
#' headerStyle <- createStyle(fontSize = 11, fontColour = "black", halign = "center",
#' border = "TopBottom", borderColour = "black",
#' borderStyle = c('thin', 'double'), textDecoration = 'bold')
#' 
#' totalStyle <-  createStyle(numFmt = "###,###,###.0")
#' 
#' horizontalStyle <- createStyle(border = "bottom", borderColour = "black", borderStyle = 'thin', valign = 'center')
#' 
#' freq_simples <- frecuencias_simples(diseño = disenio_cat, datos = dataset, pregunta = 'P1',
#'  DB_Mult = DB_Mult, tipo_pregunta = 'categorica') 
#'  
#' freq_simples_format <- frecuencias_simples(tabla = freq_simples, tipo_pregunta = 'categorica')
#' 
#' formato_frecuencias_simples(tabla = freq_simples_format , wb = wb, hojas = c(1,2), 
#' renglon = c(1,1), columna = 1 , estilo_encabezado = headerStyle , 
#' estilo_horizontal = horizontalStyle , estilo_total = totalStyle, tipo_pregunta = 'categorica')
#' }
#' 

formato_frecuencias_simples <- function(tabla, wb, hojas = c(1,2), renglon = c(1,1),
                                        columna, estilo_encabezado, 
                                        estilo_horizontal, estilo_total,
                                        tipo_pregunta = 'categorica'){
  
  
  # Primera Página renglón 2 columna 1
  
  writeData(wb = wb, sheet = hojas[1], x = tabla[[1]], startRow = renglon[1], 
            startCol = columna, colNames = TRUE, borders = 'none', keepNA = TRUE,
            na.string = '-', headerStyle = estilo_encabezado)
  
  # Ultimo renglón
  addStyle(wb = wb, sheet = hojas[1], style = estilo_horizontal,
           rows = (renglon[1] + nrow(tabla[[1]])), 
           cols = columna:(ncol(tabla[[1]])), gridExpand = TRUE, stack = TRUE)
  
  
  writeData(wb = wb, sheet = hojas[2], x = tabla[[2]], startRow = renglon[2], 
            startCol = columna, colNames = TRUE, borders = 'none', keepNA = TRUE,
            na.string = '-', headerStyle = estilo_encabezado)
  
  # Ultimo renglón
  addStyle(wb = wb, sheet = hojas[2], style = estilo_horizontal,
           rows = (renglon[2] + nrow(tabla[[2]])), cols = columna:(ncol(tabla[[2]])),
           gridExpand = TRUE, stack = TRUE)
  
  # Estilo total
  
  if (tipo_pregunta == 'categorica' || tipo_pregunta == 'multiple'){
    
    # Columna total
    addStyle(wb = wb, sheet = hojas[1], style = estilo_total,
             rows = renglon[1]:(renglon[1] + nrow(tabla[[1]])), cols = 2,
             gridExpand = TRUE, stack = TRUE)
    
    # Columna total
    addStyle(wb = wb, sheet = hojas[2], style = estilo_total,
             rows = renglon[2]:(renglon[2] + 1 + nrow(tabla[[2]])), cols = 2,
             gridExpand = TRUE, stack = TRUE)
    
  }
  
  if (tipo_pregunta == 'continua'){
    
    # Columna total
    addStyle(wb = wb, sheet = hojas[1], style = estilo_total,
             rows = (renglon[1] + 1), cols = 2,
             gridExpand = TRUE, stack = TRUE)
    
    # Columna total
    addStyle(wb = wb, sheet = hojas[2], style = estilo_total,
             rows = (renglon[2] + 1), cols = 2,
             gridExpand = TRUE, stack = TRUE)
    
  }
  
}

# FUNCIÓN FRECUENCIAS SIMPLES EN EXCEL 

#' Función frecuencias simples en Excel
#'
#' @description Escribe los títulos 'Frecuencias simples' y 'Frecuencias simples (dispersión)', logo indicado, título de la pregunta, tabla de frecuencias simples formateada y pie de tabla en las hojas y renglones mencionados por el usuario
#' @usage frecuencias_simples_excel <- function(
#' pregunta,
#' num_pregunta,
#' datos,
#' DB_Mult,
#' lista_preguntas,
#' diseño,
#' wb,
#' renglon = c(1,1),
#' columna = 1,
#' hojas = c(1,2),
#' tipo_pregunta = 'categorica', 
#' fuente,
#' logo_path,
#' organismo_participacion,
#' estilo_encabezado = headerStyle,
#' estilo_horizontal = horizontalStyle,
#' estilo_total = totalStyle
#'  )
#' @param pregunta Nombre de la pregunta sobre la cual se desea obtener frecuencias simples e incluirlas en un workbook de Excel
#' @param num_pregunta Número de pregunta
#' @param datos Conjunto de datos en formato .sav
#' @param DB_Mult Data frame con las preguntas múltiples
#' @param lista_preguntas Data frame que contiene los títulos de las pregunta
#' @param diseño Diseño muestral que se ocupará según el tipo de pregunta
#' @param wb Workbook de Excel que contiene al menos dos hojas 
#' @param renglon Vector tamaño 2 especificando el número de renglon en el cual se desea empezar a escribir la tabla 1 y tabla 2 respectivamente
#' @param columna Columna en la cual se desea empezar a escribir las tablas 
#' @param hojas Vector de número de hojas en el cual se desea insertar las tablas
#' @param tipo_pregunta Tipo de pregunta_ 'categorica', 'multiple', 'continua'
#' @param fuente Nombre del proyecto
#' @param logo_path Path del logo de la UNAM 
#' @param organismo_participacion Organismos que participaron en el proyecto, por ejemplo, 'Ciudadanía Mexicana'
#' @param estilo_encabezado estilo el cual se desea usar para los nombres de las columnas
#' @param estilo_horizontal estilo último renglón horizontal
#' @param estilo_total estilo el cual se desea usar para la columna total
#' 
#' @details Esta función envuelve todas las funciones creadas para obtener las frecuencias simples, por lo que esta función es la única que se deberá llamar para crear las frecuencias simples de las preguntas deseadas e insertarlas en ciertas hojas de Excel
#' @details Es necesario crear al menos dos hojas de excel con la función addWorksheet de la paquetería openxlsx
#' @details El estilo_total se recomienda crear un estilo con la función createStyle de openxlsx con el formato que se desea, por ejemplo "###,###,###.0"
#' @details El estilo_horizontal hace referencia al tipo de lineado horizontal se desea en el úntimo renglón de la tabla
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{writeData}} \code{\link{createStyle}}  \code{\link(setRowHeights)} \code{\link{insertImage}} \code{\link{mergeCells}}
#' @examples \dontrun{
#' # Creación del workbook
#' wb <- createWorkbook()
#' addWorksheet(wb, "Frecuencias simples")
#' addWorksheet(wb, "Frecuencias simples (dispersión)")
#' 
#' # Estilos 
#' headerStyle <- createStyle(fontSize = 11, fontColour = "black", halign = "center",
#' border = "TopBottom", borderColour = "black",
#' borderStyle = c('thin', 'double'), textDecoration = 'bold')
#' totalStyle <-  createStyle(numFmt = "###,###,###.0")
#' horizontalStyle <- createStyle(border = "bottom",
#' borderColour = "black", borderStyle = 'thin', valign = 'center')
#' 
#' # Carga de datos
#'  dataset <- read.spss("data/BASE_CONACYT_260118.sav", to.data.frame = TRUE)
#'  Lista_Preg <- read_xlsx("aux/Lista de Preguntas.xlsx", 
#'  sheet = "Lista Preguntas")$Nombre %>% as.vector()
#'   DB_Mult <- read_xlsx("aux/Lista de Preguntas.xlsx", 
#'   sheet = "Múltiple") %>% as.data.frame()
#'
#' #Diseño
#'  disenio_mult <- disenio(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1, reps=FALSE, datos = dataset)
#'  
#' frecuencias_simples_excel(pregunta = 'P1',
#' num_pregunta = 1,
#' datos = dataset,
#' DB_Mult = DB_Mult,
#' lista_preguntas = Lista_Preg,
#' diseño = disenio_mult,
#' wb = wb,
#' renglon = c(1,1),
#' columna = 1, 
#' hojas = c(1,2),
#' tipo_pregunta = 'multiple',
#' fuente =  'Conacyt 2018',
#' organismo_participacion = 'Ciudadanía mexicana',
#' estilo_encabezado = headerStyle,
#' estilo_horizontal = horizontalStyle,
#' estilo_total = totalStyle
#' )
#' }
frecuencias_simples_excel <- function(pregunta, num_pregunta, datos, DB_Mult,
                                    lista_preguntas, diseño, wb, renglon = c(1,1),
                                    columna = 1, hojas = c(1,2),
                                    tipo_pregunta = 'categorica', fuente,
                                    logo_path = '~/Desktop/UNAM/DIAO/rsrvyest/img/logo_unam.png',
                                    organismo_participacion,
                                    estilo_encabezado = headerStyle,
                                    estilo_horizontal = horizontalStyle,
                                    estilo_total = totalStyle){
  
  
  titleStyle <- createStyle(fontSize = 24, fontColour = '#011f4b',
                            textDecoration = 'underline')
  
  subtitleStyle <- createStyle(fontSize = 14, fontColour = '#011f4b',
                            textDecoration = 'underline')
  
  table_Style <- createStyle(valign = 'center', halign = 'center')
  
  # Título general renglón 1 columna 2
  
  writeData(wb = wb, sheet = hojas[1], x = 'Frecuencias estadísticas',
            startCol = 2, startRow = 1)
  writeData(wb = wb, sheet = hojas[1], x = nombre_proyecto,
            startCol = 2, startRow = 2)
  addStyle(wb = wb, sheet = hojas[1], style = titleStyle, rows = 1, cols = 2,
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb = wb, sheet = hojas[1], style = subtitleStyle, rows = 2, cols = 2,
           gridExpand = TRUE, stack = TRUE)
  
  
  writeData(wb = wb, sheet = hojas[2], x = 'Frecuencias estadísticas (dispersión)',
            startCol = 2, startRow = 1)
  writeData(wb = wb, sheet = hojas[2], x = nombre_proyecto,
            startCol = 2, startRow = 2)
  addStyle(wb = wb, sheet = hojas[2], style = titleStyle, rows = 1, cols = 2,
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb = wb, sheet = hojas[2], style = subtitleStyle, rows = 2, cols = 2,
           gridExpand = TRUE, stack = TRUE)
  
  # Pegar logo UNAM renglón 1, columna 1
  

  logo <- logo_path
  
  insertImage(wb = wb, sheet = hojas[1],
              file = logo, startRow = 1, startCol = 1, 
              width = 2.08, height =2.2, units = 'cm')
  
  insertImage(wb = wb, sheet = hojas[2],
              file = logo, startRow = 1, startCol = 1, 
              width = 2.08, height = 2.2, units = 'cm')
  

  # Título Pregunta
  
  np <- lista_preguntas[[num_pregunta]]
  
  writeData(wb = wb, sheet = hojas[1], x = np, startRow = (renglon[1]), 
            startCol = columna, colNames = TRUE, borders = 'none')

  setRowHeights(wb = wb, sheet = hojas[1], rows = (renglon[1]), heights = 35)
  
  
  writeData(wb = wb, sheet = hojas[2], x = np, startRow = (renglon[2]), 
            startCol = columna, colNames = TRUE, borders = 'none')

  setRowHeights(wb = wb, sheet = hojas[2], rows = (renglon[2] + 1), heights = 35)
  
  
  # Tabla
  
  tabla_titulo <- paste0('Tabla ', num_pregunta, '*')
  
  writeData(wb = wb, sheet = hojas[1], x = tabla_titulo, startRow = (renglon[1] +1),
            startCol = columna, colNames = FALSE, borders = 'none')
  
  setRowHeights(wb = wb, sheet = hojas[1], rows = (renglon[1] + 1), heights = 30)
  
  

  writeData(wb = wb, sheet = hojas[2], x = tabla_titulo, startRow = (renglon[2] +1),
            startCol = columna, colNames = FALSE, borders = 'none')
  
  setRowHeights(wb = wb, sheet = hojas[2], rows = (renglon[1] + 1), heights = 30)
  
  
  est <- frecuencias_simples(diseño = diseño, datos = datos, pregunta = pregunta,
                      DB_Mult = DB_Mult, tipo_pregunta = tipo_pregunta)
  
  freq <-  formatear_frecuencias_simples(est, tipo_pregunta = tipo_pregunta)
  
  formato_frecuencias_simples(tabla = freq, wb = wb, hojas = hojas,
                              renglon = c((renglon[1] + 1 + 1), (renglon[2] + 1 + 1)),
                              columna = columna, estilo_encabezado = estilo_encabezado,
                              estilo_horizontal = estilo_horizontal,
                              estilo_total = estilo_total,
                              tipo_pregunta = tipo_pregunta)
  
  
  # Merge
  
  addStyle(wb = wb, sheet = hojas[1], style = table_Style,
           rows = (renglon[1] + 1), cols = 1:ncol(freq[[1]]),
           gridExpand = TRUE, stack = TRUE)
  
  mergeCells(wb = wb, sheet = hojas[1], cols = 1:ncol(freq[[1]]),
             rows = (renglon[1] + 1))

  
  addStyle(wb = wb, sheet = hojas[2], style = table_Style,
           rows = (renglon[2] + 1), cols = 1:ncol(freq[[2]]),
           gridExpand = TRUE, stack = TRUE)
  
  mergeCells(wb = wb, sheet = hojas[2], cols = 1:ncol(freq[[2]]),
             rows = (renglon[2] + 1))

  
  # Fuente
  
  fuente_preg <- paste0("Fuente: ", fuente)
  organismo <- paste0("Organismo de participacion ", organismo_participacion)
  tabla_preg <- paste0("* Tabla correspondiente a la pregunta ", num_pregunta,
                       " de ", fuente)
  
  fontStyle <- createStyle(fontSize = 9, fontColour = '#5b5b5b')
  
  
  writeData(wb = wb, sheet = hojas[1], x = fuente_preg,
            startRow = (renglon[1] + 2 + 1 + nrow(freq[[1]])), 
            startCol = columna, borders = 'none')
  
  addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
           rows = (renglon[1] + 2 + 1 + nrow(freq[[1]])), cols = columna,
           gridExpand = TRUE, stack = TRUE)
  
  setRowHeights(wb = wb, sheet = hojas[1],
                rows = (renglon[1] + 2 + 1 + nrow(freq[[1]])), heights = 10)

  
  
  writeData(wb = wb, sheet = hojas[1], x = organismo,
            startRow = (renglon[1] + 3+ 1 + nrow(freq[[1]])), 
            startCol = columna, borders = 'none')
  
  addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
           rows = (renglon[1] + 3 + 1 + nrow(freq[[1]])), cols = columna,
           gridExpand = TRUE, stack = TRUE)
  
  setRowHeights(wb = wb, sheet = hojas[1],
                rows = (renglon[1] + 3 + 1 + nrow(freq[[1]])), heights = 10)
  
  
  writeData(wb = wb, sheet = hojas[1], x = tabla_preg,
            startRow = (renglon[1] + 4 + 1+ nrow(freq[[1]])), 
            startCol = columna, borders = 'none')
  
  addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
           rows = (renglon[1] + 4 + 1 + nrow(freq[[1]])), cols = columna,
           gridExpand = TRUE, stack = TRUE)
  
  setRowHeights(wb = wb, sheet = hojas[1],
                rows = (renglon[1] + 4 + 1 + nrow(freq[[1]])), heights = 10)
  
  
  
  
  writeData(wb = wb, sheet = hojas[2], x = fuente_preg,
            startRow = (renglon[2] + 2+1 + nrow(freq[[2]])), 
            startCol = columna, borders = 'none')
  
  addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
           rows = (renglon[2] + 2 + 1 + nrow(freq[[2]])), cols = columna,
           gridExpand = TRUE, stack = TRUE)
  
  setRowHeights(wb = wb, sheet = hojas[2],
                rows = (renglon[2] + 2 + 1 + nrow(freq[[2]])), heights = 10)
  
  
  writeData(wb = wb, sheet = hojas[2], x = organismo,
            startRow = (renglon[2] + 3 +1+ nrow(freq[[2]])), 
            startCol = columna, borders = 'none')
  
  addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
           rows = (renglon[2] + 3 + 1 + nrow(freq[[2]])), cols = columna,
           gridExpand = TRUE, stack = TRUE)
  
  setRowHeights(wb = wb, sheet = hojas[2],
                rows = (renglon[2] + 3 + 1 + nrow(freq[[2]])), heights = 10)
  
  
  writeData(wb = wb, sheet = hojas[2], x = tabla_preg,
            startRow = (renglon[2] + 4+1 + nrow(freq[[2]])), 
            startCol = columna, borders = 'none')
  
  addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
           rows = (renglon[2] + 4 + 1 + nrow(freq[[2]])), cols = columna,
           gridExpand = TRUE, stack = TRUE)
  
  setRowHeights(wb = wb, sheet = hojas[2],
                rows = (renglon[2] + 4 + 1 + nrow(freq[[2]])), heights = 10)
  
  return(freq)
  
}

################################################################################
################################################################################
############################## TABLAS CRUZADAS #################################
################################################################################
################################################################################


# FUNCIÓN TABLAS CRUZADAS SEGÚN TIPO PREGUNTA

#' Función tablas cruzadas según tipo de pregunta y dominio
#' @description Se crea la tabla cruzada según dominio tipo de pregunta (categórica, múltiple o continua) 
#' @usage tablas_cruzadas(
#' diseño,
#' pregunta,
#' dominio,
#' datos,
#' DB_Mult,
#' na.rm,
#' vartype = c("ci","se","var","cv"),
#' cuantiles = c(0,0.25, 0.5, 0.75,1),
#' significancia = 0.95
#' proporcion = FALSE,
#' metodo_prop = 'likelihood',
#' DEFF = TRUE,
#' tipo_pregunta = 'categorica'
#' )
#' @param diseño Diseño muestral que se ocupará según el tipo de pregunta
#' @param pregunta Pregunta de la cual se quieren obtener las frecuencias simples, por ejemplo, 'P_1'
#' @param dominio Nombre del dominio del cual se desea obtener la tabla cruzada
#' @param datos Conjunto de datos en formato .sav
#' @param DB_Mult Data frame con las preguntas múltiples
#' @param na.rm Valor lógico que indica si se deben de omitir valores faltantes
#' @param vartype Métricas de variabilidad: error estándar ("se"), intervalo de confianza ("ci"), varianza ("var") o coeficiente de variación ("cv")
#' @param cuantiles Vector de cuantiles a calcular
#' @param significancia Nivel de confianza: 0.95 por default
#' @param proporcion Valor lógico que indica si se desen usar métodos para calcular la proporción que puede tener intervalos de confianza más precisos cerca de 0 y 1
#' @param metodo_prop  Si proporcion = TRUE; tipo de método de proporción que se desea usar: "logit", "likelihood", "asin", "beta", "mean"
#' @param DEFF Valor lógico que indica si se desea calcular el efecto de diseño
#' @param tipo_pregunta Tipo de pregunta: 'categorica', 'multiple', 'continua'
#' @return Tabla tipo tibble con las estadísticas especificadas en el parámetro estadisticas por respuestas pertenecientes a la pregunta y al dominio especificados
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{survey_mean}} \code{\link{srvyr::group_by}}
#' @example \dontrun{
#'  dataset <- read.spss("data/BASE_CONACYT_260118.sav", to.data.frame = TRUE) 
#'  Lista_Preg <- read_xlsx("aux/Lista de Preguntas.xlsx",
#'                        sheet = "Lista Preguntas")$Nombre %>% as.vector()
#'  DB_Mult <- read_xlsx("aux/Lista de Preguntas.xlsx",  sheet = "Múltiple") %>% as.data.frame()
#'  Lista_Cont <- read_xlsx("aux/Lista de Preguntas.xlsx", 
#'   sheet = "Continuas")$VARIABLE %>% as.vector()
#'  Dominios <- read_xlsx("aux/Lista de Preguntas.xlsx", sheet = "Dominios")$Dominios %>% as.vector()
#' 
#'  disenio_mult <- disenio(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1, reps=FALSE, datos = dataset)
#'  
#'  tablas_cruzadas(diseño = disenio_mult, pregunta = 'P1', dominio = 'Sexo', datos = dataset,
#'  DB_Mult = DB_Mult, tipo_pregunta = 'multiple')
#'  }

tablas_cruzadas <- function(diseño, pregunta, dominio, datos, DB_Mult,
                            na.rm = TRUE, vartype = c("ci","se","var","cv"),
                            cuantiles = c(0,0.25, 0.5, 0.75,1),
                            significancia = 0.95, proporcion = FALSE, 
                            metodo_prop = "likelihood", DEFF = TRUE,
                            tipo_pregunta = 'categorica'){
  
  if (tipo_pregunta == 'categorica'){
    
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
    
    tabla_cruzada <- estadisticas %>% 
      mutate(
      prop_cv = ifelse(is.nan(prop_cv), NA, prop_cv),
      prop_deff = ifelse(is.nan(prop_deff), NA, prop_deff)
      )
      
    tabla_cruzada %<>% 
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
      dplyr::rename(Categorias := !!sym(dominio)) %>% 
      select(Dominio, Categorias, everything())
    
  }
  
  if (tipo_pregunta == 'continua'){
    
    cruce <- diseño %>% 
      srvyr::group_by(!!sym(dominio)) %>% 
      srvyr::summarise(
        prop = survey_mean(
          as.numeric(!!sym(pregunta)),
          na.rm = na.rm, 
          vartype = vartype,
          level = significancia,
          proportion = proporcion,
          prop_method = metodo_prop,
          deff = DEFF
        ),
        total = round(survey_total(
          na.rm = na.rm
        ), 0),
        cuantiles = survey_quantile(
          as.numeric(!!sym(pregunta)),
          quantiles = cuantiles,
          na.rm = na.rm
        )
      )
    
    cruce %<>% 
      mutate(!!sym(dominio) := str_trim(!!sym(dominio), side = "both")) %>% 
      dplyr::rename('Categorias' = !!sym(dominio))
    
    tabla_cruzada <- add_column(cruce, 'Dominio' = dominio, .before = "Categorias")
    
    tabla_cruzada %<>% mutate(
      prop_cv = ifelse(is.nan(prop_cv), NA, prop_cv),
      prop_deff = ifelse(is.nan(prop_deff), NA, prop_deff)
    )

  }
  
  if (tipo_pregunta == 'multiple'){
    
    ## Onehot 
    ps <- DB_Mult %>% 
      dplyr::filter(!is.na(!!sym(pregunta))) %>% 
      dplyr::pull(!!sym(pregunta))
    
    df <- datos %>% 
      select(all_of(ps))
    
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
      diseño %<>% srvyr::mutate(!!sym(i) := variable)
    }
    
    tabla_cruzada = tibble()
    
    for (categ in categorias){
      Dominios_tabla <- {{diseño}} %>% 
        srvyr::group_by(!!sym(dominio), .drop = TRUE) %>% 
        srvyr::summarize(
          prop = survey_mean(!!sym(categ),
                             na.rm = na.rm, 
                             vartype = vartype,
                             level = significancia,
                             proportion = proporcion,
                             prop_method = metodo_prop,
                             deff = DEFF),
          total = round(survey_total(!!sym(categ),
                               na.rm = TRUE
          ),0)) %>% 
        mutate(prop_low = ifelse(prop_low < 0, 0, prop_low),
               prop_upp = ifelse(prop_upp > 1, 1, prop_upp),
               Dominio = dominio,
               Categorias := str_trim(!!sym(dominio), side = 'both')) 
      
      tabla_cruzada <- rbind(tabla_cruzada, Dominios_tabla)
    }
    
    tabla_cruzada %<>% select(Dominio, Categorias, total,
                                total_se, prop,prop_se, 
                                prop_low, prop_upp,
                                prop_cv, prop_var, prop_deff)
    
    tabla_cruzada %<>% mutate(
      prop_cv = ifelse(is.nan(prop_cv), NA, prop_cv),
      prop_deff = ifelse(is.nan(prop_deff), NA, prop_deff)
    )
  }
  
  
  return(tabla_cruzada)
}

# FUNCIÓN FORMATEAR TABLA CRUZADA

#'  Formatea tabla cruzada 
#'  
#' @description Se formatea la tabla cruzada según tipo de pregunta y dominio
#' @usage formatear_tabla_cruzada(
#' pregunta,
#' datos,
#' dominio,
#' tabla, 
#' DB_Mult,
#' tipo_pregunta
#' )
#' @param pregunta Nombre de la pregunta sobre la cual se desea obtener la tabla cruzada
#' @param datos Conjunto de datos en formato .sav
#' @param dominio Dominio sobre el cual se desea desagregar la tabla cruzada de cierta pregunta
#' @param tabla Tabla cruzada creada con la función tablas_cruzadas()
#' @param tipo_pregunta Tipo de pregunta: 'categorica', 'multiple', 'continua'
#' @return Lista de dos tibbles: en la primera tibble se encuentra el total estimado y las métricas media, límite inferior y superior; 
#' en la segunda tibble se encuentra el total y las métricas error estándar, varianza, coeficiente de variación y el efecto de diseño (si es que DEFF = TRUE)
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{survey_mean}}
#' @example \dontrun{
#'  dataset <- read.spss("data/BASE_CONACYT_260118.sav", to.data.frame = TRUE) 
#'  Lista_Preg <- read_xlsx("aux/Lista de Preguntas.xlsx",
#'                        sheet = "Lista Preguntas")$Nombre %>% as.vector()
#'  DB_Mult <- read_xlsx("aux/Lista de Preguntas.xlsx",  sheet = "Múltiple") %>% as.data.frame()
#'  Lista_Cont <- read_xlsx("aux/Lista de Preguntas.xlsx", 
#'   sheet = "Continuas")$VARIABLE %>% as.vector()
#'  Dominios <- read_xlsx("aux/Lista de Preguntas.xlsx", sheet = "Dominios")$Dominios %>% as.vector()
#' 
#'  disenio_mult <- disenio(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1, reps=FALSE, datos = dataset)
#'  
#'  tc <-  tablas_cruzadas(diseño = disenio_mult, pregunta = 'P1', dominio = 'Sexo', datos = dataset,
#'  DB_Mult = DB_Mult, tipo_pregunta = 'multiple')
#'  
#'  formatear_tabla_cruzada(pregunta = 'P1', datos = dataset, dominio = 'Sexo', tabla = tc, DB_Mult = DB_Mult, tipo_pregunta = 'multiple')
#' } 


formatear_tabla_cruzada <- function(pregunta, datos, dominio, tabla, DB_Mult,
                                    tipo_pregunta = 'categorica'){
  
  
  if (tipo_pregunta == 'categorica'){
    
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
  
  tabla_final_1 <- tabla %>% 
    select( Dominio,
            Categorias,
            Total,
            ends_with(c( '_prop', '_prop_low', '_prop_upp'))) %>% 
    select(Dominio, Categorias, Total, contains(categorias)) %>% 
    dplyr::ungroup()
  
  #Segunda tabla
  
  tabla_final_2 <- tabla %>% 
    select(
      Dominio,
      Categorias,
      Total,
      ends_with(c('_prop_se', '_prop_cv', '_prop_var','prop_deff'))) %>% 
    select(Dominio, Categorias, Total, contains(categorias)) %>% 
    dplyr::ungroup()
  
  names(tabla_final_1) <- nombres1
  
  names(tabla_final_2) <- nombres2
  
  }
  
  if(tipo_pregunta == 'continua'){
    
    tabla_final_1 <- tabla %>% 
      select(Dominio, Categorias, total, prop, prop_low, prop_upp, cuantiles_q00,
             cuantiles_q25, cuantiles_q50, cuantiles_q75, cuantiles_q100) %>% 
      dplyr::rename('Total' = total,
                    'Categorías' = Categorias,
                    'Media' = prop,
                    'Lim. inf' = prop_low,
                    'Lim. sup' = prop_upp,
                    'Mín' = cuantiles_q00,
                    'Q25' = cuantiles_q25,
                    'Mediana' = cuantiles_q50,
                    'Q75' = cuantiles_q75,
                    'Máx' = cuantiles_q100) 
    
    tabla_final_2 <- tabla %>% 
      select(Dominio, Categorias, total, prop_se,  prop_cv, prop_var, prop_deff) %>% 
      dplyr::rename( "Err. Est." = prop_se,
                     "Total" = total,
                     'Categorías' = Categorias,
                     "Var" = prop_var,
                     "Coef. Var." = prop_cv,
                     "DEFF" = prop_deff) 
  }
  
  if(tipo_pregunta == 'multiple'){
    
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
    
    nombres_tabla1 <- c('Dominio', 'Categorías', 'Total',
                        rep(c('Media', 'Lim. inf.', 'Lim. sup.'),
                            length(categorias)))
    
    nombres_tabla2 <- c('Dominio', 'Categorías', 'Total',
                        rep(c("Err. Est.", "Coef. Var.", "Var.", "DEFF"),
                            length(categorias)))
    
    
    tabla_final_1 <- tabla1[c(1:len),] 
    
    for (i in seq((len+1),nrow(tabla1), by = len)){
      tabla_c1 = tabla1[i:(i+len-1),c(-1,-2)]
      tabla_final_1 = cbind(tabla_final_1,tabla_c1)
    }
    
    suma_totales %<>% select(Total)
    
    
    tabla_final_1 = cbind(tabla_final_1,suma_totales)
    
    tabla_final_1 %<>% relocate(Total,.after = Categorias)
    
    tabla_final_2 <- tabla2[c(1:len),] 
    
    for (i in seq((len+1),nrow(tabla2), by = len)){
      tabla_c2 = tabla2[i:(i+len-1),c(-1,-2)]
      tabla_final_2 = cbind(tabla_final_2,tabla_c2)
    }
    
    tabla_final_2 =cbind(tabla_final_2,suma_totales) 
    
    tabla_final_2 %<>% relocate(Total,.after = Categorias)
    
    names(tabla_final_1) <- nombres_tabla1
    names(tabla_final_2) <- nombres_tabla2
    
  }
  
  
  return(list(tabla_final_1, tabla_final_2))
  
}

# FUNCIÓN TOTAL NACIONAL 

#' Función total nacional
#' @description  Esta función transforma la tabla de frecuencias simples de cierta pregunta especificada por el usuario en un formato similar al de tablas cruzadas (1 rengón)
#' @usage total_general(
#' diseño,
#' pregunta,
#' DB_Mult,
#' datos,
#' dominio = 'General',
#' tipo_pregunta,
#' na.rm = TRUE, 
#' vartype = c("se","ci","cv", "var"),
#' cuantiles =  c(0,0.25, 0.5, 0.75,1),
#' significancia = 0.95,
#' proporcion = FALSE, 
#' metodo_prop = "likelihood", 
#' DEFF = TRUE
#' )
#' @param diseño Diseño muestral que se ocupará según el tipo de pregunta
#' @param pregunta Nombre de la pregunta sobre la cual se desea obtener la tabla general
#' @param DB_Mult  Data frame con las preguntas múltiples
#' @param datos Conjunto de datos en formato .sav
#' @param dominio Nombre al cual se desea nombrar al total estimado, por ejemplo 'General', 'Total Nacional', etc.
#' @param tipo_pregunta Tipo de pregunta: 'categorica', 'multiple', 'continua'
#' @param na.rm Valor lógico que indica si se deben de omitir valores faltantes
#' @param vartype Métricas de variabilidad: error estándar ("se"), intervalo de confianza ("ci"), varianza ("var") o coeficiente de variación ("cv")
#' @param cuantiles Vector de cuantiles a calcular
#' @param significancia Nivel de confianza: 0.95 por default
#' @param proporcion Valor lógico que indica si se desen usar métodos para calcular la proporción que puede tener intervalos de confianza más precisos cerca de 0 y 1
#' @param metodo_prop  Si proporcion = TRUE; tipo de método de proporción que se desea usar: "logit", "likelihood", "asin", "beta", "mean"
#' @param DEFF Valor lógico que indica si se desea calcular el efecto de diseño
#' @return  Lista de dos tibbles: en la primera tibble se encuentra el total estimado y las métricas media, límite inferior y superior; 
#' en la segunda tibble se encuentra el total y las métricas error estándar, varianza, coeficiente de variación y el efecto de diseño (si es que DEFF = TRUE)
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{survey_mean}}
#' @example \dontrun{
#' # Lectura de datos
#'  dataset <- read.spss("data/BASE_CONACYT_260118.sav", to.data.frame = TRUE) 
#'  Lista_Preg <- read_xlsx("aux/Lista de Preguntas.xlsx",
#'                        sheet = "Lista Preguntas")$Nombre %>% as.vector()
#'  DB_Mult <- read_xlsx("aux/Lista de Preguntas.xlsx",  sheet = "Múltiple") %>% as.data.frame()
#'  Lista_Cont <- read_xlsx("aux/Lista de Preguntas.xlsx", 
#'   sheet = "Continuas")$VARIABLE %>% as.vector()
#'  Dominios <- read_xlsx("aux/Lista de Preguntas.xlsx", sheet = "Dominios")$Dominios %>% as.vector()
#' 
#' # Diseño
#'  disenio_mult <- disenio(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1, reps=FALSE, datos = dataset)
#'  
#'  total_general (diseño = disenio_mult,  pregunta = 'P1', dominio = 'General', datos = dataset,
#'  DB_Mult = DB_Mult, tipo_pregunta = 'multiple')
#' }

total_general <- function(diseño, pregunta, DB_Mult, datos, dominio = 'General',
                           tipo_pregunta = 'categorica', na.rm = TRUE, 
                           vartype = c("se","ci","cv", "var"),
                           cuantiles =  c(0,0.25, 0.5, 0.75,1),
                           significancia = 0.95, proporcion = FALSE, 
                           metodo_prop = "likelihood", DEFF = TRUE){
  
  if (tipo_pregunta == 'categorica'){
    
  estadisticas <- {{diseño}} %>% 
    srvyr::group_by(!!sym(pregunta)) %>% 
    srvyr::summarize(
      prop = survey_mean(
        na.rm = na.rm, 
        vartype = vartype,
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
  
  total <- estadisticas %>%
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
  
  total_formateado <- formatear_tabla_cruzada(pregunta = pregunta, datos = datos,
                                              tabla = total, dominio = dominio,
                                              DB_Mult = DB_Mult,
                                              tipo_pregunta = tipo_pregunta)
  
  }
  
  if (tipo_pregunta == 'continua'){
    
    total <- {{diseño}} %>% 
      srvyr::summarise(
        prop = survey_mean(
          as.numeric(!!sym(pregunta)),
          na.rm = na.rm, 
          vartype = vartype,
          level = significancia,
          proportion = proporcion,
          prop_method = metodo_prop,
          deff = DEFF
        ),
        cuantiles = survey_quantile(
          as.numeric(!!sym(pregunta)),
          quantiles = cuantiles,
          na.rm = na.rm
        ),
        total = survey_total(
          na.rm = na.rm)
      ) %>% 
      select(total, prop, prop_low, prop_upp, cuantiles_q00, cuantiles_q25,
             cuantiles_q50, cuantiles_q75, cuantiles_q100, prop_se, prop_var,
             prop_cv, prop_deff) %>% 
      mutate(Dominio = 'General',
             Categorias = 'Total') %>% 
      select(Dominio, Categorias, everything())
    
    
    total_formateado <- formatear_tabla_cruzada(pregunta = pregunta, datos = datos,
                                                tabla = total, dominio = dominio,
                                                DB_Mult = DB_Mult,
                                                tipo_pregunta = tipo_pregunta)
    
  }
  
  if(tipo_pregunta == 'multiple'){
    
    
   freqs <- frecuencias_simples(diseño = diseño, datos = datos,
                                pregunta = pregunta, DB_Mult = DB_Mult,
                                tipo_pregunta = "multiple")
   
  freqs <- formatear_frecuencias_simples(tabla = freqs,
                                         tipo_pregunta = 'multiple')

  categorias <- freqs[[1]] %>% dplyr::pull(Respuesta)
  
  # Primera tabla
  
  nombres_tabla_desagregada <- c('Dominio', 'Categorías', 'Total',
                                 rep(c('Media', 'Lim. inf.', 'Lim. sup.'),
                                     length(categorias)))
  
  suma_totales = freqs[[1]] %>% dplyr::summarise(Total = sum(Total))
  
  
  tabla_desagregada = freqs[[1]][1,c(-1,-2)] %>% 
    dplyr::mutate(Dominio = "General", Categorias = "Total")
  
  for (i in seq(2,nrow(freqs[[1]]), by = 1)){
    tabla_c1 = freqs[[1]][i,c(-1,-2)]
    tabla_desagregada = cbind(tabla_desagregada,tabla_c1)
  }
  
  tabla_desagregada = cbind(suma_totales,tabla_desagregada)
  
  tabla_desagregada %<>% relocate(Categorias,.before = Total) %>% 
    relocate(Dominio,.before = Categorias)
  
  names(tabla_desagregada) = nombres_tabla_desagregada

  # Segunda tabla
  
  nombres_tabla_desagregada_b <- c('Dominio', 'Categorías', 'Total',
                                   rep(c("Err. Est.", "Coef. Var.","Var.", "DEFF"),
                                       length(categorias)))
  
  tabla_desagregada_b = freqs[[2]][1,c(-1,-2)] %>% 
    dplyr::mutate(Dominio = "General", Categorias = "Total")
  
  for (i in seq(2,nrow(freqs[[2]]), by = 1)){
    tabla_c1 = freqs[[2]][i,c(-1,-2)]
    tabla_desagregada_b = cbind(tabla_desagregada_b,tabla_c1)
  }
  
  tabla_desagregada_b = cbind(suma_totales, tabla_desagregada_b)
  
  
  tabla_desagregada_b %<>% relocate(Categorias,.before = Total) %>% 
    relocate(Dominio,.before = Categorias)
  
  names(tabla_desagregada_b) = nombres_tabla_desagregada_b
  
  total_formateado <- list(tabla_desagregada, tabla_desagregada_b)

    }

  return(total_formateado)
}


# FUNCIÓN TABLA CRUZADA TOTAL NACIONAL Y POR DOMINIOS 

#' Función tabla cruzada total y por dominios
#' @description Esta función une la tabla cruzada general y las tablas cruzadas por dominios generadas por las funciones total_general() y tablas_cruzadas()
#' @usage tabla_cruzada_total(
#' diseño,
#' pregunta,
#' datos,
#' dominios,
#' tipo_pregunta
#' )
#' @param diseño Diseño muestral que se ocupará según el tipo de pregunta
#' @param pregunta Nombre de la pregunta sobre la cual se desea obtener la tabla cruzada total,
#' @param datos Conjunto de datos en formato .sav
#' @param dominios Vector el cual contiene los nombres de los dominios sobre los cuales se desean obtener las tablas cruzadas
#' @param tipo_pregunta Tipo de pregunta: 'categorica', 'multiple', 'continua'
#' @return Lista de dos tibbles: en la primera tibble se encuentra el total estimado y las métricas media, límite inferior y superior; 
#' en la segunda tibble se encuentra el total y las métricas error estándar, varianza, coeficiente de variación y el efecto de diseño (si es que DEFF = TRUE en la función tablas_cruzadas)
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{survey_mean}}
#' @example \dontrun{
#'  # Lectura de datos
#'  dataset <- read.spss("data/BASE_CONACYT_260118.sav", to.data.frame = TRUE) 
#'  Lista_Preg <- read_xlsx("aux/Lista de Preguntas.xlsx",
#'                        sheet = "Lista Preguntas")$Nombre %>% as.vector()
#'  DB_Mult <- read_xlsx("aux/Lista de Preguntas.xlsx",  sheet = "Múltiple") %>% as.data.frame()
#'  Lista_Cont <- read_xlsx("aux/Lista de Preguntas.xlsx", 
#'   sheet = "Continuas")$VARIABLE %>% as.vector()
#'  Dominios <- read_xlsx("aux/Lista de Preguntas.xlsx", sheet = "Dominios")$Dominios %>% as.vector()
#' 
#' # Diseño
#'  disenio_mult <- disenio(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1, reps=FALSE, datos = dataset)
#'  
#'  tabla_cruzada_total (diseño = disenio_mult,  pregunta = 'P1', dominios = Dominios, datos = dataset,
#'  DB_Mult = DB_Mult, tipo_pregunta = 'multiple')
#' }

tabla_cruzada_total <- function(diseño, pregunta, datos, DB_Mult,
                                dominios = Dominios, 
                                tipo_pregunta = 'categorica'){
  
  # Tabla nacional
  nacional <- total_general(diseño = diseño, datos = datos, DB_Mult = DB_Mult,
                             pregunta = pregunta, tipo_pregunta = tipo_pregunta)
  
  # Dominios de dispersión
  
  for (d in dominios) {
    
    f_n <- tablas_cruzadas(diseño = diseño, pregunta = pregunta, datos = datos,
                          DB_Mult = DB_Mult, dominio = d,
                          tipo_pregunta = tipo_pregunta)
    
    f_tc <- formatear_tabla_cruzada(pregunta = pregunta, datos = datos, 
                                    dominio = d, tabla = f_n, DB_Mult = DB_Mult,
                                    tipo_pregunta = tipo_pregunta)
    
    nacional[[1]] <- rbind(nacional[[1]], f_tc[[1]])
    nacional[[2]] <- rbind(nacional[[2]], f_tc[[2]])
    
  }
  
  return(nacional)
}

# FORMATO CATEGORIAS RESPUESTAS (TABLAS CRUZADAS) 

#' Formato categorías respuestas para preguntas múltiples y categóricas
#' @description Crea un vector para las respuestas de las preguntas múltiples y categóricas.
#' @usage categorias_pregunta_formato(
#' pregunta,
#' diseño,
#' datos,
#' DB_Mult,
#' tipo_pregunta,
#' metricas = TRUE
#' )
#' @param pregunta Nombre de la pregunta sobre la cual se desea obtener las categorías 
#' @param diseño Diseño muestral que se ocupará según el tipo de pregunta
#' @param datos Conjunto de datos en formato .sav
#' @param DB_Mult Data frame con las preguntas múltiples
#' @param tipo_pregunta Tipo de pregunta: 'categorica', 'multiple', 'continua'
#' @param metricas Valor lógica que indica si se desean las categorías de las respuestas para las métricas media,ímite inferior, límite superior o para las métricas error estándar, varianza, coeficiente de variación y el efecto de diseño
#' @return Vector con las categorías de las respuestas de la pregunta mecionada por el usuario, si métricas = TRUE se regresa un vector con las categorías y entre cada categoría se encuentran dos NA's, 
#' @return si metricas = FALSE se regresa un vector con las categorías de las respuestas y entre cada categoría se encuentran 3 NA's
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @example \dontrun{
#'  # Carga de datos
#'  dataset <- read.spss("data/BASE_CONACYT_260118.sav", to.data.frame = TRUE)
#'  Lista_Preg <- read_xlsx("aux/Lista de Preguntas.xlsx", 
#'  sheet = "Lista Preguntas")$Nombre %>% as.vector()
#'   DB_Mult <- read_xlsx("aux/Lista de Preguntas.xlsx", 
#'   sheet = "Múltiple") %>% as.data.frame()
#'
#' #Diseño
#'  disenio_mult <- disenio(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1, reps=FALSE, datos = dataset)
#' 
#' categorias_pregunta_formato(pregunta = 'P1', diseño = disenio_mult, datos = dataset, DB_Mult = DB_Mult, tipo_pregunta = 'multiple', metricas = TRUE)
#' }

categorias_pregunta_formato <- function(pregunta, diseño, datos, DB_Mult,
                                        tipo_pregunta = 'categorica',
                                        metricas = TRUE){
  
  if(tipo_pregunta == 'categorica'){
    
  categorias <- datos %>%
    select(!!sym(pregunta)) %>%
    pull()%>%
    levels() %>% 
    str_trim(side = 'both')
  
  est <- frecuencias_simples(diseño = diseño, datos = datos, pregunta = pregunta,
                             DB_Mult = DB_Mult, tipo_pregunta = 'categorica')
  
  
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
  }
  
  if(tipo_pregunta == 'multiple'){
    
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
  }
  
  return(categs)
  
}


# FUNCIÓN FORMATO CATEGORÍAS PARA PREGUNTAS CATEGÓRICAS Y MÚLTIPLES

#' Función formato categorías para preguntas categóricas y múltiples 
#' @description El vector obtenido con la función categorias_pregunta_formato() se escribe en un workbook de Excel
#' @usage formato_categorias(
#' tabla,
#' pregunta,
#' diseño,
#' datos,
#' DB_Mult,
#' wb,
#' renglon,
#' columna = 4,
#' hojas,
#' estilo_cuerpo,
#' tipo_pregunta
#' )
#' @param tabla Tabla cruzada creada por la función total_general()
#' @param pregunta  Nombre de la pregunta sobre la cual se creo la tabla
#' @param diseño Diseño muestral que se ocupará según el tipo de pregunta
#' @param datos Conjunto de datos en formato .sav
#' @param DB_Mult Data frame con las preguntas múltiples
#' @param wb Workbook de Excel que contiene al menos dos hojas 
#' @param renglon Vector tamaño 2 especificando el número de renglon en el cual se desea empezar a escribir el vector de categorías
#' @param columna Columna en la cual se desea empezar a escribir las tablas. SIEMPRE EN LA CUARTA COLUMNA
#' @param hojas Vector de número de hojas en el cual se desea insertar los vectores
#' @param estilo_cuerpo Estilo que indica cómo se desea formatear el vector en las hojas de excel
#' @param tipo_pregunta Tipo de pregunta_ 'categorica', 'multiple', 'continua'
#' @details Una vez escrito el vector colapsa celdas (3 o 4)
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{mergeCells}} \code{\link{addStyle}} \code{\link{createStyle}} 
#' @examples \dontrun{
#' # Creación del workbook
#' wb <- createWorkbook()
#' addWorksheet(wb, "Tablas cruzadas")
#' addWorksheet(wb, "Tablas cruzadas (dispersión)")
#' 
#' # Estilos
#'   bodyStyle <- createStyle(halign = 'center', border = "TopBottomLeftRight",
#'   borderColour = "black", borderStyle = 'thin',
#'   valign = 'center', wrapText = TRUE)
#'   
#' # Carga de datos   
#'  dataset <- read.spss("data/BASE_CONACYT_260118.sav", to.data.frame = TRUE) 
#'  Lista_Preg <- read_xlsx("aux/Lista de Preguntas.xlsx",
#'                        sheet = "Lista Preguntas")$Nombre %>% as.vector()
#'  DB_Mult <- read_xlsx("aux/Lista de Preguntas.xlsx",  sheet = "Múltiple") %>% as.data.frame()
#'  Lista_Cont <- read_xlsx("aux/Lista de Preguntas.xlsx", 
#'   sheet = "Continuas")$VARIABLE %>% as.vector()
#'  Dominios <- read_xlsx("aux/Lista de Preguntas.xlsx", sheet = "Dominios")$Dominios %>% as.vector()
#' 
#'  disenio_mult <- disenio(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1, reps=FALSE, datos = dataset)
#'  
#'  total <- total_general (diseño = disenio_mult,  pregunta = 'P1', dominio = 'General', datos = dataset,
#'  DB_Mult = DB_Mult, tipo_pregunta = 'multiple')
#'  
#'  formato_categorias(tabla = total, pregunta = 'P1', diseño = disenio_mult, datos = dataset,
#'  DB_Mult = DB_Mult, wb = wb, renglon = c(1,1), columna = 4, hojas = c(1,2), estilo_cuerpo = bodyStyle, tipo_pregunta = 'multiple')
#'  
#'  openxlsx::openXL(wb) 
#'  }
#'
#'
formato_categorias <- function(tabla, pregunta, diseño, datos, DB_Mult, wb,
                               renglon = c(1,1), columna = 4, hojas = c(3,4),
                               estilo_cuerpo, tipo_pregunta = 'categorica'){
  
  # Renglón 2 columna 4
  
  simples <- categorias_pregunta_formato(pregunta = pregunta,
                                         diseño = diseño, datos = datos,
                                         metricas = FALSE, DB_Mult = DB_Mult, 
                                         tipo_pregunta = tipo_pregunta)
  
  writeData(wb = wb, sheet =  hojas[1], x = simples, startRow = renglon[1], 
            startCol = columna, borders = 'surrounding',
            borderStyle = 'thin',  keepNA = FALSE, colNames = FALSE)
  
  
  for (k in seq(columna, ncol(tabla[[1]]), by=3)){
    
    mergeCells(wb = wb, sheet = hojas[1], cols = k:(k+2), rows = renglon[1])
    addStyle(wb = wb, sheet = hojas[1], style = estilo_cuerpo, rows = renglon[1],
             cols = k:(k+2), stack = TRUE, gridExpand = TRUE)
  }
  
  
  metricas <- categorias_pregunta_formato(pregunta = pregunta,
                                          diseño = diseño, datos = datos,
                                          metricas = TRUE, DB_Mult = DB_Mult, 
                                          tipo_pregunta = tipo_pregunta)
  
  writeData(wb = wb, sheet =  hojas[2], x = metricas, startRow = renglon[2], 
            startCol = columna, borders = 'surrounding',
            borderStyle = 'thin',  keepNA = FALSE, colNames = FALSE)
  
  for (k in seq(columna, ncol(tabla[[2]]), by=4)) {
    mergeCells(wb = wb, sheet = hojas[2], cols = k:(k+3), rows = renglon[2])
    addStyle(wb = wb, sheet = hojas[2], style = estilo_cuerpo, rows = renglon[2],
             cols = k:(k+3), stack = TRUE, gridExpand = TRUE)
  }
  
}

# FUNCIÓN FORMATO TABLAS CRUZADAS EXCEL

#' Escribe y da formato a la tablas cruzadas en el workbook
#' @description Función para escribir y dar formato a las tablas cruzadas en el workbook
#' @usage formato_tablas_cruzadas(
#' tabla,
#' wb,
#' renglon = c(1,1),
#' columna = 1,
#' hojas = c(3,4),
#' estilo_encabezado = headerStyle,
#' estilo_total = totalStyle
#' )
#' @param tabla Lista con las tablas cruzadas formateadas
#' @param wb Workbook en el que escribirán las tablas
#' @param renglon Renglón donde iniciará la tabla, un renglón por tipo de tabla cruzada
#' @param columna Columna en la que se iniciará la tabla
#' @param hojas Hojas del workbook en donde se ecribirán las tablas
#' @param estilo_encabezado Estilo para el ancabezado de la tabla
#' @param estilo_total Estilo para el resto de la tabla
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{survey_mean}}
#' @example \dontrun{
#' # Estilos 
#'  headerStyle <- createStyle( fontSize = 11, fontColour = "black", halign = "center",
#'  border = "TopBottom", borderColour = "black",
#'  borderStyle = c('thin', 'double'), textDecoration = 'bold')
#'  
#'  totalStyle <-  createStyle(numFmt = "###,###,###.0")
#' formato_tablas_cruzadas(tabla = tabla_cruzada, wb = wb, renglon = c(1,1),
#' columna = 1, hojas = c(3,4), estilo_encabezado = headerStyle, 
#' estilo_total = totalStyle)
#' }

formato_tablas_cruzadas <- function(tabla, wb, renglon = c(1,1), columna = 1,
                                    hojas = c(3,4),
                                    estilo_encabezado = headerStyle,
                                    estilo_total = totalStyle){
  
  
  # renglón 3 columna 1
  
  writeData(wb = wb, sheet =  hojas[1], x = tabla[[1]], startRow = renglon[1], 
            startCol = columna, borders = 'surrounding',
            borderStyle = 'thin',  keepNA = TRUE, na.string = '-', col.names = TRUE,
            headerStyle = estilo_encabezado)
  
  
  writeData(wb = wb, sheet =  hojas[2], x = tabla[[2]], startRow = renglon[2], 
            startCol = columna, borders = 'surrounding',
            borderStyle = 'thin',  keepNA = TRUE, na.string = '-', col.names = TRUE,
            headerStyle = estilo_encabezado)
}

# FUNCIÓN ESTILO DOMINIOS, DA FORMATO A LA COLUMNA DE DOMINIOS

#' Da formato a la columna de dominios de las tablas cruzadas en el workbook
#' @description Función para unir las celdas donde se encuentran los nombres de los dominios es las hojas solicitadas.
#' @usage estilo_dominios(
#' tabla,
#' wb, 
#' columna = 1,
#' hojas = c(3,4),
#' dominios,
#' renglon = c(1,1),
#' estilo_dominios = horizontalStyle,
#' estilo_merge_dominios = bodyStyle,
#' tipo_pregunta
#' )
#' @param tabla Lista con las tablas cruzadas formateadas
#' @param wb Workbook en el que aplicará el formato
#' @param columna Columna en la que se aplicará el formato
#' @param hojas Hojas del workbook a las que se aplicará el formato, una por tabla
#' @param dominios Vector con los dominios
#' @param renglon Renglón donde se iniciará el formato, uno por tipo de tabla, uno por tabla
#' @param estilo_dominios Estilo previamente definido con el borde en la parte inferior de la tabla
#' @param estilo_merge_dominios Estilo previamente definido con bordes en toda la tabla
#' @param tipo_pregunta Tipo de pregunta ("categorica", "multiple")
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{survey_mean}}
#' @example \dontrun{
#' estilo_dominios(tabla = tabla_cruzada, wb = wb, columna = 1, hojas = c(3,4), 
#' dominios = dominios, renglon = c(1,1), estilo_dominios = horizontalStyle,
#' estilo_merge_dominios = bodyStyle, tipo_pregunta = 'categorica')
#' }
#' 
#' 
estilo_dominios <- function(tabla, wb, columna = 1, hojas = c(3,4), dominios,
                            renglon = c(1,1), estilo_dominios = horizontalStyle,
                            estilo_merge_dominios = bodyStyle, tipo_pregunta){
  
  # Renglón donde se empieza a escribir la tabla cruzada + 1 para linea horizontal nacional 
  
  r <- renglon[1] + 1
  
  dominios_general <- c('General', dominios)
  
  for (d in dominios_general) {
    tope <- tabla[[1]] %>% select(Dominio) %>% filter(Dominio == d) %>% nrow()
    
    for (i in hojas){
      mergeCells(wb = wb, sheet = i, cols = 1, rows = r:(r+tope-1))
      addStyle(wb = wb, sheet = i, cols = 1, rows = r:(r+tope-1),
               style = estilo_merge_dominios, stack = TRUE, gridExpand = TRUE)
    }
    
    addStyle(wb = wb, sheet = hojas[1], cols = 1:ncol(tabla[[1]]), rows = (r+tope-1),
             style = estilo_dominios, stack = TRUE, gridExpand = TRUE)
    
    addStyle(wb = wb, sheet = hojas[2], cols = 1:ncol(tabla[[2]]), rows = (r+tope-1),
             style = estilo_dominios, stack = TRUE, gridExpand = TRUE)
    
    r <- r + tope
  }
  
  # Merge primeras tres columnas
  
  if (tipo_pregunta == 'categorica' || tipo_pregunta == 'multiple'){
    
    mergeCells(wb = wb, sheet = hojas[1], cols = 1:3, rows = (renglon[1] - 1))
    mergeCells(wb = wb, sheet = hojas[2], cols = 1:3, rows = (renglon[2] - 1))
  }
  
}

# FUNCIÓN ESTILO COLUMNAS, DIVIDE LAS MÉTRICAS DE LAS TABLAS CRUZADAS POR 
# CATEGORÍAS RESPUESTAS SOLO PREGUNTAS CATEGÓRICAS Y MÚLTIPLES

estilo_columnas <- function(tabla, wb, hojas = c(3,4), estilo = verticalStyle,
                            renglon = c(1,1), tipo_pregunta){
  
  if (tipo_pregunta == 'categorica' || tipo_pregunta == 'multiple'){
    
  secuencia_1_a<- c(-1, 0, 1)
  
  secuencia_1_b <- seq(4, ncol(tabla[[1]]), by=3)
  
  for (k in secuencia_1_a){
    addStyle(wb = wb, sheet = hojas[1], cols = (k+2),
             rows = (renglon[1] + 2):(renglon[1] + 2 + nrow(tabla[[1]])),
             style = estilo, stack = TRUE, gridExpand = TRUE)
  }
  
  for (k in secuencia_1_b){
    addStyle(wb = wb, sheet = hojas[1], cols = (k+2),
             rows = (renglon[1] + 1):(renglon[1] + 2 + nrow(tabla[[1]])),
             style = estilo, stack = TRUE, gridExpand = TRUE)
  }
  secuencia_2_a <- c(-2, -1, 0)
  
  secuencia_2_b <- seq(4, ncol(tabla[[2]]), by=4)
  
  for (k in secuencia_2_a ) {
    addStyle(wb = wb, sheet = hojas[2], cols = (k+3),
             rows = (renglon[2] + 2):(renglon[2] + 2 + nrow(tabla[[2]])),
             style = verticalStyle, stack = TRUE, gridExpand = TRUE)
  }
  
  for (k in secuencia_2_b) {
    addStyle(wb = wb, sheet = hojas[2], cols = (k+3),
             rows = (renglon[2] + 1):(renglon[2] + 2 + nrow(tabla[[2]])),
             style = verticalStyle, stack = TRUE, gridExpand = TRUE)
  }
  }
  
  if (tipo_pregunta == 'continua'){
    
    secuencia_1_a<- c(-1, 0, 1,(ncol(tabla[[1]]) - 2))
    
    for (k in secuencia_1_a){
      addStyle(wb = wb, sheet = hojas[1], cols = (k+2),
               rows = (renglon[1] + 1):(renglon[1] + 1 + nrow(tabla[[1]])),
               style = estilo, stack = TRUE, gridExpand = TRUE)
    }

    secuencia_2_a <- c(-2, -1, 0, (ncol(tabla[[2]]) - 3))
    

    for (k in secuencia_2_a) {
      addStyle(wb = wb, sheet = hojas[2], cols = (k+3),
               rows = (renglon[2] + 1):(renglon[2] + 1 + nrow(tabla[[2]])),
               style = verticalStyle, stack = TRUE, gridExpand = TRUE)
    }
    
  }
  
}

# FUNCIÓN TABLAS CRUZADAS EN EXCEL

#' Función tablas cruzadas en Excel
#' @description Escribe los títulos 'Tablas cruzadas' y 'Tablas cruzadas (dispersión)', logo indicado, título de la pregunta, tabla crzada total y pie de tabla en las hojas y renglones mencionados por el usuario
#' @usage tablas_cruzadas_excel(
#' pregunta,
#' num_pregunta,
#' dominios,
#' datos,
#' DB_Mult, 
#' lista_preguntas,
#' diseño,
#' wb,
#' renglon,
#' columna,
#' hojas ,
#' tipo_pregunta,
#' fuente,
#' organismo_participacion,
#' logo_path,
#' estilo_encabezado = headerStyle,
#' estilo_columnas = verticalStyle,
#' estilo_categorias = bodyStyle,
#' estilo_horizontal = horizontalStyle,
#' estilo_total = totalStyle
#' )
#' @param pregunta  Nombre de la pregunta sobre la cual se desea obtener la tabla cruzada e incluirla en un workbook de Excel
#' @param num_pregunta Número de pregunta
#' @param dominios Vector el cual contiene los nombres de los dominios sobre los cuales se desean obtener sus respectivas tablas cruzadas
#' @param DB_Mult Data frame con las preguntas múltiples
#' @param lista_preguntas Data frame que contiene los títulos de las pregunta
#' @param diseño Diseño muestral que se ocupará según el tipo de pregunta
#' @param wb Workbook de Excel que contiene al menos dos hojas 
#' @param renglon Vector tamaño 2 especificando el número de renglon en el cual se desea empezar a escribir la tabla 1 y tabla 2 respectivamente
#' @param columna Columna en la cual se desea empezar a escribir las tablas 
#' @param hojas Vector de número de hojas en el cual se desea insertar las tablas
#' @param tipo_pregunta Tipo de pregunta_ 'categorica', 'multiple', 'continua'
#' @param fuente Nombre del proyecto
#' @param organismo_participacion Organismos que participaron en el proyecto, por ejemplo, 'Ciudadanía Mexicana'
#' @param logo_path Path del logo de la UNAM 
#' @param estilo_encabezado estilo el cual se desea usar para los nombres de las columnas
#' @param estilo_horizontal estilo último renglón horizontal
#' @param estilo_total estilo el cual se desea usar para la columna total
#' 
#' @details Esta función envuelve todas las funciones creadas para obtener las tablas cruzadas, por lo que esta función es la única que se deberá llamar para crear las tablas cruzadas de las preguntas deseadas e insertarlas en ciertas hojas de Excel
#' @details Es necesario crear al menos dos hojas de excel con la función addWorksheet de la paquetería openxlsx
#' @details El estilo_total se recomienda crear un estilo con la función createStyle de openxlsx con el formato que se desea, por ejemplo "###,###,###.0"
#' @details El estilo_horizontal hace referencia al tipo de lineado horizontal se desea en el úntimo renglón de la tabla
#' @details El estilo_encabezado hace referencia al tipo de bordes, alineación, color, etc. que se desea obtener en el nombre de las columnas de la tabla cruzada final
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{writeData}} \code{\link{createStyle}}  \code{\link(setRowHeights)} \code{\link{insertImage}} \code{\link{mergeCells}}
#' @examples \dontrun{
#' # Creación del workbook
#' wb <- createWorkbook()
#' addWorksheet(wb, "Tablas cruzadas")
#' addWorksheet(wb, "Tablas cruzadas (dispersión)")
#' 
#' # Estilos 
#' headerStyle <- createStyle(fontSize = 11, fontColour = "black", halign = "center",
#' border = "TopBottom", borderColour = "black",
#' borderStyle = c('thin', 'double'), textDecoration = 'bold')
#' totalStyle <-  createStyle(numFmt = "###,###,###.0")
#' horizontalStyle <- createStyle(border = "bottom",
#' borderColour = "black", borderStyle = 'thin', valign = 'center')
#' 
#' # Carga de datos
#'  dataset <- read.spss("data/BASE_CONACYT_260118.sav", to.data.frame = TRUE)
#'  Lista_Preg <- read_xlsx("aux/Lista de Preguntas.xlsx", 
#'  sheet = "Lista Preguntas")$Nombre %>% as.vector()
#'   DB_Mult <- read_xlsx("aux/Lista de Preguntas.xlsx", 
#'   sheet = "Múltiple") %>% as.data.frame()
#'
#' #Diseño
#'  disenio_mult <- disenio(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1, reps=FALSE, datos = dataset)
#'  
#' tablas_cruzadas_excel(
#' pregunta = 'P1',
#' num_pregunta = 1,
#' dominios = Dominios,
#' datos = dataset,
#' DB_Mult = DB_Mult,
#' lista_preguntas = Lista_Preg,
#' diseño = disenio_mult,
#' wb = wb,
#' renglon = c(1,1),
#' columna = 1, 
#' hojas = c(1,2),
#' tipo_pregunta = 'multiple',
#' fuente =  'Conacyt 2018',
#' organismo_participacion = 'Ciudadanía mexicana',
#' logo_path = '~/Desktop/UNAM/DIAO/rsrvyest/img/logo_unam.png'
#' estilo_encabezado = headerStyle,
#' estilo_horizontal = horizontalStyle,
#' estilo_total = totalStyle
#' )
#' openxlsx::openXL(wb)
#' }
#' 
tablas_cruzadas_excel <- function(pregunta, num_pregunta, dominios, datos,
                                  DB_Mult, lista_preguntas, diseño, wb,
                                  renglon = c(1,1), columna = 1, hojas = c(3,4),
                                  tipo_pregunta = 'categorica', fuente,
                                  organismo_participacion, logo_path = '~/Desktop/UNAM/DIAO/rsrvyest/img/logo_unam.png',
                                  estilo_encabezado = headerStyle,
                                  estilo_columnas = verticalStyle,
                                  estilo_categorias = bodyStyle,
                                  estilo_horizontal = horizontalStyle,
                                  estilo_total = totalStyle){
  
  # Título pregunta
  
  titleStyle <- createStyle(fontSize = 24, fontColour = '#011f4b',
                            textDecoration = 'underline')
  
  subtitleStyle <- createStyle(fontSize = 14, fontColour = '#011f4b',
                               textDecoration = 'underline')
  
  table_Style <- createStyle(valign = 'center', halign = 'center')
  
  # Título general renglón 1 columna 2
  
  writeData(wb = wb, sheet = hojas[1], x = 'Tablas cruzadas',
            startCol = 2, startRow = 1)
  writeData(wb = wb, sheet = hojas[1], x = organismo_participacion,
            startCol = 2, startRow = 2)
  addStyle(wb = wb, sheet = hojas[1], style = titleStyle, rows = 1, cols = 2,
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb = wb, sheet = hojas[1], style = subtitleStyle, rows = 2, cols = 2,
           gridExpand = TRUE, stack = TRUE)
  
  
  writeData(wb = wb, sheet = hojas[2], x = 'Tablas cruzadas (dispersión)',
            startCol = 2, startRow = 1)
  writeData(wb = wb, sheet = hojas[2], x = nombre_proyecto,
            startCol = 2, startRow = 2)
  addStyle(wb = wb, sheet = hojas[2], style = titleStyle, rows = 1, cols = 2,
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb = wb, sheet = hojas[2], style = subtitleStyle, rows = 2, cols = 2,
           gridExpand = TRUE, stack = TRUE)
  
  
  # Pegar logo UNAM renglón 1, columna 1
  
  logo <- logo_path
  
  insertImage(wb = wb, sheet = hojas[1],
              file = logo, startRow = 1, startCol = 1, 
              width = 2.08, height = 2.2, units = 'cm')
  
  insertImage(wb = wb, sheet = hojas[2],
              file = logo, startRow = 1, startCol = 1, 
              width = 2.08, height = 2.2, units = 'cm')
  
  # Título pregunta
  
  np <- lista_preguntas[[num_pregunta]]
  
  writeData(wb = wb, sheet =  hojas[1], x = np, startCol = columna,
            startRow = renglon[1], borders = 'none')

  writeData(wb = wb, sheet =  hojas[2], x = np, startCol = columna,
            startRow = renglon[2], borders = 'none')
  
  setRowHeights(wb = wb, sheet = hojas[1], rows = (renglon[1]), heights = 35)
  
  setRowHeights(wb = wb, sheet = hojas[2], rows = (renglon[2] + 1), heights = 35)
  
  
  # Tabla título 
  
  tabla_titulo <- paste0('Tabla ', num_pregunta, '*')
  
  writeData(wb = wb, sheet = hojas[1], x = tabla_titulo, startRow = (renglon[1] +1),
            startCol = columna, colNames = FALSE, borders = 'none')
  
  setRowHeights(wb = wb, sheet = hojas[1], rows = (renglon[1] + 1), heights = 30)
  
  
  writeData(wb = wb, sheet = hojas[2], x = tabla_titulo, startRow = (renglon[2] +1),
            startCol = columna, colNames = FALSE, borders = 'none')
  
  setRowHeights(wb = wb, sheet = hojas[2], rows = (renglon[1] + 1), heights = 30)
  
  
  # Tabla Cruzada 
  
  f <- tabla_cruzada_total(diseño = diseño, pregunta = pregunta, datos = datos,
                           DB_Mult = DB_Mult, dominios = dominios, 
                           tipo_pregunta = tipo_pregunta)
  
  # Merge
  
  addStyle(wb = wb, sheet = hojas[1], style = table_Style,
           rows = (renglon[1] + 1), cols = 1:ncol(f[[1]]),
           gridExpand = TRUE, stack = TRUE)
  
  mergeCells(wb = wb, sheet = hojas[1], cols = 1:ncol(f[[1]]),
             rows = (renglon[1] + 1))
  
  
  
  addStyle(wb = wb, sheet = hojas[2], style = table_Style,
           rows = (renglon[2] + 1), cols = 1:nrow(f[[2]]),
           gridExpand = TRUE, stack = TRUE)
  
  mergeCells(wb = wb, sheet = hojas[2], cols = 1:ncol(f[[2]]),
             rows = (renglon[2] + 1))

  
  
  if(tipo_pregunta == 'categorica' || tipo_pregunta == 'multiple'){
    
  # Formato Categorías empieza en el renglón 2
  
  formato_categorias(tabla = f, pregunta = pregunta, datos = datos, 
                     DB_Mult = DB_Mult, diseño = diseño,
                     wb = wb, renglon = c((renglon[1] + 1+1),(renglon[2] + 1+1)),
                     columna = (columna + 3), hojas = hojas,
                     estilo_cuerpo = estilo_categorias,
                     tipo_pregunta = tipo_pregunta)
  
    setRowHeights(wb = wb, sheet = hojas[1], rows = (renglon[1] + 1+1), heights = 45)
    
    setRowHeights(wb = wb, sheet = hojas[2], rows = (renglon[2] + 1+1), heights = 45)
    
    # Escribir tabla Cruzada empieza en el renglón 3
    
    formato_tablas_cruzadas(tabla = f, wb = wb,
                            renglon = c((renglon[1] + 2+1), (renglon[2] + 2+1)),
                            columna = columna, hojas = hojas,
                            estilo_encabezado = estilo_encabezado,
                            estilo_total  = estilo_total)
    
    # Estilo dominios empieza en el renglón 3
    
    estilo_dominios(tabla = f, wb = wb, columna = columna, dominios = dominios,
                    hojas = hojas, renglon = c((renglon[1] + 2+1), (renglon[2] + 2+1)),
                    estilo_dominios = estilo_horizontal, tipo_pregunta = tipo_pregunta)
    # Estilo columnas 
    
    estilo_columnas(tabla = f, hojas = hojas, wb = wb, estilo = estilo_columnas, 
                    renglon = c(renglon[1]+1, renglon[2]+1), tipo_pregunta = tipo_pregunta)
    
    addStyle(wb = wb, sheet = hojas[2], style = estilo_total,
             rows = (renglon[2]+1):(renglon[2]+1 + nrow(f[[2]])), cols = 3,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = hojas[1], style = estilo_total,
             rows = (renglon[1]+1):(renglon[1]+1 + nrow(f[[1]])), cols = 3,
             gridExpand = TRUE, stack = TRUE)
    
    # Fuente
    
    fontStyle <- createStyle(fontSize = 9, fontColour = '#5b5b5b')
    
    fuente_preg <- paste0("Fuente: ", fuente)
    organismo <- paste0("Organismo de participacion ", organismo_participacion)
    tabla_preg <- paste0("Tabla correspondiente a la pregunta ", num_pregunta,
                         " de ", fuente)

    writeData(wb = wb, sheet = hojas[1], x = fuente_preg,
              startRow = (renglon[1] + 3+1 + nrow(f[[1]])), startCol = columna,
              borders = 'none')
    
    addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
             rows = (renglon[1] + 3 + 1 + nrow(f[[1]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)
    
    setRowHeights(wb = wb, sheet = hojas[1],
                  rows = (renglon[1] + 3 + 1 + nrow(f[[1]])),
                  heights = 10)
    
    
    writeData(wb = wb, sheet = hojas[1], x = organismo,
              startRow = (renglon[1] + 4+1 + nrow(f[[1]])), startCol = columna,
              borders = 'none')
    
    addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
             rows = (renglon[1] + 4 + 1 + nrow(f[[1]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)
    
    setRowHeights(wb = wb, sheet = hojas[1],
                  rows = (renglon[1] + 4 + 1 + nrow(f[[1]])),
                  heights = 10)
    
    
    writeData(wb = wb, sheet = hojas[1], x = tabla_preg,
              startRow = (renglon[1] + 5 +1+ nrow(f[[1]])), startCol = columna,
              borders = 'none')
    
    addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
             rows = (renglon[1] + 5 + 1 + nrow(f[[1]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)
    
    setRowHeights(wb = wb, sheet = hojas[1],
                  rows = (renglon[1] + 5 + 1 + nrow(f[[1]])),
                  heights = 10)
    
    
    writeData(wb = wb, sheet = hojas[2], x = fuente_preg,
              startRow = (renglon[2] + 3+1 + nrow(f[[2]])), startCol = columna,
              borders = 'none')
    
    addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
             rows = (renglon[2] + 3 + 1 + nrow(f[[2]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)
    
    setRowHeights(wb = wb, sheet = hojas[2],
                  rows = (renglon[2] + 3 + 1 + nrow(f[[2]])),
                  heights = 10)
    
    
    writeData(wb = wb, sheet = hojas[2], x = organismo,
              startRow = (renglon[2] + 4+1 + nrow(f[[2]])), startCol = columna,
              borders = 'none')
    
    addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
             rows = (renglon[2] + 4 + 1 + nrow(f[[2]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)
    
    setRowHeights(wb = wb, sheet = hojas[2],
                  rows = (renglon[2] + 4 + 1 + nrow(f[[2]])),
                  heights = 10)
    
    writeData(wb = wb, sheet = hojas[2], x = tabla_preg,
              startRow = (renglon[2] + 5+1 + nrow(f[[2]])), startCol = columna,
              borders = 'none')
    addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
             rows = (renglon[2] + 5 + 1 + nrow(f[[2]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)
    
    setRowHeights(wb = wb, sheet = hojas[2],
                  rows = (renglon[2] + 5 + 1 + nrow(f[[2]])),
                  heights = 10)
    
  }
  
  if(tipo_pregunta == 'continua'){
    
    # Escribir tabla Cruzada empieza en el renglón 3
    
    formato_tablas_cruzadas(tabla = f, wb = wb, 
                            renglon = c((renglon[1] + 1+1), (renglon[2] + 1+1)),
                            columna = columna, hojas = hojas,
                            estilo_encabezado = estilo_encabezado,
                            estilo_total  = estilo_total)
    
    # Estilo dominios empieza en el renglón 3
    
    estilo_dominios(tabla = f, wb = wb, columna = columna, dominios = dominios,
                    hojas = hojas, renglon = c((renglon[1] +1+ 1), (renglon[2]+1 + 1)),
                    estilo_dominios = estilo_horizontal, 
                    tipo_pregunta = tipo_pregunta)
    
    # Estilo columnas 
    
    estilo_columnas(tabla = f, hojas = hojas, wb = wb, estilo = estilo_columnas, 
                    renglon = c(renglon[1]+1, renglon[2]+1), tipo_pregunta = tipo_pregunta)
    
    
    # Fuente
    fontStyle <- createStyle(fontSize = 9, fontColour = '#5b5b5b')
    
    fuente_preg <- paste0("Fuente: ", fuente)
    organismo <- paste0("Organismo de participacion ", organismo_participacion)
    tabla_preg <- paste0("Tabla correspondiente a la pregunta ", num_pregunta,
                         " de ", fuente)
    
    writeData(wb = wb, sheet = hojas[1], x = fuente_preg,
              startRow = (renglon[1] + 2+1 + nrow(f[[1]])), startCol = columna,
              borders = 'none')
    
    addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
             rows = (renglon[1] + 2 + 1 + nrow(f[[1]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)
    
    setRowHeights(wb = wb, sheet = hojas[1],
                  rows = (renglon[1] + 2 + 1 + nrow(f[[1]])),
                  heights = 10)
    
    
    writeData(wb = wb, sheet = hojas[1], x = organismo,
              startRow = (renglon[1] + 3+1 + nrow(f[[1]])), startCol = columna,
              borders = 'none')
    
    addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
             rows = (renglon[1] + 3 + 1 + nrow(f[[1]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)
    
    setRowHeights(wb = wb, sheet = hojas[1],
                  rows = (renglon[1] + 3 + 1 + nrow(f[[1]])),
                  heights = 10)
    
    
    writeData(wb = wb, sheet = hojas[1], x = tabla_preg,
              startRow = (renglon[1] + 4+1 + nrow(f[[1]])), startCol = columna,
              borders = 'none')
    
    addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
             rows = (renglon[1] + 4 + 1 + nrow(f[[1]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)
    
    setRowHeights(wb = wb, sheet = hojas[1],
                  rows = (renglon[1] + 4 + 1 + nrow(f[[1]])),
                  heights = 10)
    
    
    
    writeData(wb = wb, sheet = hojas[2], x = fuente_preg,
              startRow = (renglon[2] + 2+1 + nrow(f[[2]])), startCol = columna,
              borders = 'none')
    
    addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
             rows = (renglon[2] + 2 + 1 + nrow(f[[2]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)
    
    setRowHeights(wb = wb, sheet = hojas[2],
                  rows = (renglon[2] + 2 + 1 + nrow(f[[2]])),
                  heights = 10)
    
    writeData(wb = wb, sheet = hojas[2], x = organismo,
              startRow = (renglon[2] + 3+1 + nrow(f[[2]])), startCol = columna,
              borders = 'none')
    
    addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
             rows = (renglon[2] + 3 + 1 + nrow(f[[2]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)
    
    setRowHeights(wb = wb, sheet = hojas[2],
                  rows = (renglon[2] + 3 + 1 + nrow(f[[2]])),
                  heights = 10)
    
    
    writeData(wb = wb, sheet = hojas[2], x = tabla_preg,
              startRow = (renglon[2] + 4+1 + nrow(f[[2]])), startCol = columna,
              borders = 'none')
    
    addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
             rows = (renglon[2] + 4 + 1 + nrow(f[[2]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)
    
    setRowHeights(wb = wb, sheet = hojas[2],
                  rows = (renglon[2] + 4 + 1 + nrow(f[[2]])),
                  heights = 10)
    
    # Estilo total
    
    addStyle(wb = wb, sheet = hojas[2], style = estilo_total,
             rows = (renglon[2] +1):(renglon[2]+1 + nrow(f[[2]])), cols = 3,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = hojas[1], style = estilo_total,
             rows = (renglon[1]+1):(renglon[1]+1 + nrow(f[[1]])), cols = 3,
             gridExpand = TRUE, stack = TRUE)
    
  }
  
  return(f)
  
}

#' Función preguntas 
#' @description Función envolvente de las funciones frecuencias_simples_excel() y tablas_cruzadas_excel
#' @usage preguntas(
#' pregunta,
#' num_pregunta,
#' datos,
#' DB_Mult,
#' dominios,
#' lista_preguntas,
#' diseño,
#' wb,
#' renglon_fs,
#' renglon_tc,
#' columna = 1,
#' hojas_fs,
#' hojas_tc,
#' fuente,
#' organismo_participacion,
#' logo_path,
#' tipo_pregunta,
#' estilo_encabezado = headerStyle,
#' estilo_categorias = bodyStyle,
#' estilo_horizontal = horizontalStyle,
#' estilo_total = totalStyle,
#' frecuencias_simples = TRUE,
#' tablas_cruzadas = TRUE
#' )
#' @param pregunta  Nombre de la pregunta sobre la cual se desea obtener las frecuencias simples y/o tabla cruzada
#' @param num_pregunta Número de pregunta
#' @param datos Conjunto de datos en formato .sav
#' @param DB_Mult Data frame con las preguntas múltiples
#' @param dominios Vector de dominios sobre los cuales se desea obtener sus respectivas tablas cruzadas
#' @param lista_preguntas Data frame que contiene los títulos de las pregunta
#' @param diseño Diseño muestral que se ocupará según el tipo de pregunta
#' @param wb Workbook de Excel que contiene al menos dos hojas 
#' @param renglon_fs Vector tamaño 2 especificando el número de renglon en el cual se desea empezar a escribir las tablas de frecuencias simples formateadas
#' @param renglon_tc Vector tamaño 2 especificando el número de tenglón el cual se desea empezar a escribir la tabla cruzada formateada
#' @param columna Columna en la cual se desea empezar a escribir las tablas 
#' @param hojas_fs Vector de número de hojas en el cual se desea insertar las tablas de frecuencias simples
#' @param hojas_tc Vector de número de hojas en el cual se desea insertar la tabla cruzada
#' @param fuente Nombre del proyecto
#' @param organismo_participacion Organismos que participaron en el proyecto, por ejemplo, 'Ciudadanía Mexicana'
#' @param logo_path Path del logo de la UNAM 
#' @param tipo_pregunta Tipo de pregunta_ 'categorica', 'multiple', 'continua'
#' @param estilo_encabezado Estilo el cual se desea usar para los nombres de las columnas
#' @param estilo_categorias Estilo el cual se desea usar para formatear las categorías de las tablas cruzadas para preguntas categóricas y múltiples
#' @param estilo_horizontal Estilo último renglones horizontales (último renglón para frecuencias simples)
#' @param estilo_total Estilo el cual se desea usar para la columna total
#' @param frecuencias_simples Valor lógico que indica si se desean realizar las frecuencias simples de la pregunta indicada
#' @param tablas_cruzadas Valor lógico que indica si se desean realizar las tablas cruzadas de la pregunta indicada
#' @details El estilo_total se recomienda crear un estilo con la función createStyle de openxlsx con el formato que se desea, por ejemplo "###,###,###.0"
#' @details El estilo_horizontal hace referencia al tipo de lineado horizontal se desea en el úntimo renglón de la tabla
#' @details El estilo_categoris hace referencia al tipo de bordes, fuente y alineación que se desea aplicar en las celdas de Excel donde se encuentra el vector de categorías
#' @details El estilo_encabezado hace referencia al tipo de formato que se desea conseguir para el nombre de las columnas de las tablas
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{openxlsx}}
#' @example \dontrun{ 
#' # Creación del workbook
#'   organismo <- 'Ciudadanía mexicana'
#'   nombre_proyecto <- 'Conacyt 2018'

#' openxlsx::addWorksheet(wb, sheetName = 'Frecuencias simples')
#' showGridLines(wb, sheet = 'Frecuencias simples', showGridLines = FALSE)

#' openxlsx::addWorksheet(wb, sheetName = 'Tablas cruzadas')
#' showGridLines(wb, sheet='Tablas cruzadas', showGridLines = FALSE)

#' openxlsx::addWorksheet(wb, sheetName = 'Frecuencias (dispersión)')
#' showGridLines(wb, sheet = 'Frecuencias (dispersión)', showGridLines = FALSE)

#' openxlsx::addWorksheet(wb, sheetName = 'Tablas cruzadas (dispersión)')
#' showGridLines(wb, sheet = 'Tablas cruzadas (dispersión)', showGridLines = FALSE)

#' 
#' # Estilos 
#'   headerStyle <- createStyle(fontSize = 11, fontColour = "black", halign = "center", border = "TopBottom", borderColour = "black", borderStyle = c('thin', 'double'), textDecoration = 'bold')
#'   bodyStyle <- createStyle(halign = 'center', border = "TopBottomLeftRight", borderColour = "black", borderStyle = 'thin', valign = 'center', wrapText = TRUE)
  
#'   verticalStyle <- createStyle(border = "Right", borderColour = "black", borderStyle = 'thin', valign = 'center')
  
#'  totalStyle <-  createStyle(numFmt = "###,###,###.0")
  
#'  horizontalStyle <- createStyle(border = "bottom", borderColour = "black", borderStyle = 'thin', valign = 'center')
#' 
#' # Carga de datos
#'  dataset <- read.spss("data/BASE_CONACYT_260118.sav", to.data.frame = TRUE)
#'  Lista_Preg <- read_xlsx("aux/Lista de Preguntas.xlsx", 
#'  sheet = "Lista Preguntas")$Nombre %>% as.vector()
#'   DB_Mult <- read_xlsx("aux/Lista de Preguntas.xlsx", 
#'   sheet = "Múltiple") %>% as.data.frame()
#'
#' #Diseño
#'  disenio_mult <- disenio(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1, reps=FALSE, datos = dataset)
#'  
#'   preguntas(pregunta = 'P1', num_pregunta = 1, datos=dataset,
#'   DB_Mult = DB_Mult, dominios = Dominios, 
#'   lista_preguntas=Lista_Preg,
#'   diseño = disenio_mult, wb = wb, renglon_fs = c(1,1),
#'   renglon_tc = c(1, 1), columna = 1, hojas_fs = c(1,3),
#'   hojas_tc = c(2,4), fuente = nombre_proyecto, 
#'   tipo_pregunta = 'multiple',
#'   organismo_participacion = organismo,
#'   estilo_encabezado = headerStyle,
#'   estilo_categorias = bodyStyle,
#'   estilo_horizontal = horizontalStyle,
#'   estilo_total = totalStyle,
#'   frecuencias_simples = TRUE, tablas_cruzadas = TRUE)
#'   
#'   openxlsx::openXL(wb)
#' }
#' 
preguntas <- function(pregunta, num_pregunta, datos, DB_Mult, dominios,
                      lista_preguntas, diseño, wb, renglon_fs, renglon_tc,
                      columna = 1, hojas_fs = c(1,2),
                      hojas_tc = c(3,4),
                      fuente = 'fuente',
                      organismo_participacion = 'organismo', logo_path = '~/Desktop/UNAM/DIAO/rsrvyest/img/logo_unam.png',
                      tipo_pregunta,
                      estilo_encabezado = headerStyle,
                      estilo_categorias = bodyStyle,
                      estilo_horizontal = horizontalStyle,
                      estilo_total = totalStyle,
                      frecuencias_simples = TRUE, tablas_cruzadas = TRUE){
  
  
  if(frecuencias_simples){
    
    fs_rows <- frecuencias_simples_excel(pregunta = pregunta,
                                         num_pregunta = num_pregunta,
                                         datos = datos, DB_Mult = DB_Mult,
                                         lista_preguntas = lista_preguntas,
                                         diseño = diseño, wb = wb,
                                         renglon = renglon_fs,
                                         columna = columna, hojas = hojas_fs,
                                         tipo_pregunta = tipo_pregunta, fuente = fuente,
                                         organismo_participacion = organismo_participacion,
                                         logo_path = logo_path,
                                         estilo_encabezado = estilo_encabezado,
                                         estilo_horizontal = estilo_horizontal,
                                         estilo_total = estilo_total)
    
    
  } else{
    fs_rows <- tibble()
  }
  
  if(tablas_cruzadas){
    
    tc_rows <- tablas_cruzadas_excel(pregunta = pregunta,
                                     num_pregunta = num_pregunta,
                                     dominios = dominios, datos = datos,
                                     DB_Mult = DB_Mult,
                                     lista_preguntas = lista_preguntas,
                                     diseño = diseño, wb = wb,
                                     renglon = renglon_tc, columna = columna,
                                     hojas = hojas_tc, 
                                     tipo_pregunta = tipo_pregunta, 
                                     fuente = fuente,
                                     organismo_participacion = organismo_participacion,
                                     logo_path = logo_path,
                                     estilo_encabezado = estilo_encabezado,
                                     estilo_categorias = estilo_categorias,
                                     estilo_horizontal = estilo_horizontal,
                                     estilo_total = estilo_total)
    
  }else{
    tc_rows <- tibble()
  }
  
  return(list(fs_rows, tc_rows))
}

