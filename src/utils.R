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

formatear_frecuencias_simples <- function(tabla, tipo_pregunta = 'categorica'){
  
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

frecuencias_simples_excel <- function(pregunta, num_pregunta, datos, DB_Mult,
                                    lista_preguntas, diseño, wb, renglon = c(1,1),
                                    columna = 1, hojas = c(1,2),
                                    tipo_pregunta = 'categorica', fuente,
                                    organismo_participacion,
                                    estilo_encabezado = headerStyle,
                                    estilo_horizontal = horizontalStyle,
                                    estilo_total = totalStyle){
  
  # Título Pregunta
  
  np <- lista_preguntas[[num_pregunta]]
  
  writeData(wb = wb, sheet = hojas[1], x = np, startRow = renglon[1], 
            startCol = columna, colNames = TRUE, borders = 'none')

  writeData(wb = wb, sheet = hojas[2], x = np, startRow = renglon[2], 
            startCol = columna, colNames = TRUE, borders = 'none')

  # Tabla
  est <- frecuencias_simples(diseño = diseño, datos = datos, pregunta = pregunta,
                      DB_Mult = DB_Mult, tipo_pregunta = tipo_pregunta)
  
  freq <-  formatear_frecuencias_simples(est, tipo_pregunta = tipo_pregunta)
  
  formato_frecuencias_simples(tabla = freq, wb = wb, hojas = hojas,
                              renglon = c((renglon[1] + 1), (renglon[2] + 1)),
                              columna = columna, estilo_encabezado = estilo_encabezado,
                              estilo_horizontal = estilo_horizontal,
                              estilo_total = estilo_total,
                              tipo_pregunta = tipo_pregunta)
  
  # Fuente
  
  fuente_preg <- paste0("Fuente: ", fuente)
  organismo <- paste0("Organismo de participacion ", organismo_participacion)
  tabla_preg <- paste0("Tabla correspondiente a la pregunta ", num_pregunta,
                       " de ", fuente)
  
  writeData(wb = wb, sheet = hojas[1], x = fuente_preg,
            startRow = (renglon[1] + 2 + nrow(freq[[1]])), 
            startCol = columna, borders = 'none')
  

  writeData(wb = wb, sheet = hojas[1], x = organismo,
            startRow = (renglon[1] + 3 + nrow(freq[[1]])), 
            startCol = columna, borders = 'none')
  
  writeData(wb = wb, sheet = hojas[1], x = tabla_preg,
            startRow = (renglon[1] + 4 + nrow(freq[[1]])), 
            startCol = columna, borders = 'none')
  
  
  writeData(wb = wb, sheet = hojas[2], x = fuente_preg,
            startRow = (renglon[2] + 2 + nrow(freq[[2]])), 
            startCol = columna, borders = 'none')
  
  
  writeData(wb = wb, sheet = hojas[2], x = organismo,
            startRow = (renglon[2] + 3 + nrow(freq[[2]])), 
            startCol = columna, borders = 'none')
  
  writeData(wb = wb, sheet = hojas[2], x = tabla_preg,
            startRow = (renglon[2] + 4 + nrow(freq[[2]])), 
            startCol = columna, borders = 'none')
  
  return(freq)
  
}

################################################################################
################################################################################
############################## TABLAS CRUZADAS #################################
################################################################################
################################################################################


# FUNCIÓN TABLAS CRUZADAS SEGÚN TIPO PREGUNTA

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
          total = survey_total(!!sym(categ),
                               na.rm = TRUE
          )) %>% 
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

tablas_cruzadas_excel <- function(pregunta, num_pregunta, dominios, datos,
                                  DB_Mult, lista_preguntas, diseño, wb,
                                  renglon = c(1,1), columna = 1, hojas = c(3,4),
                                  tipo_pregunta = 'categorica', fuente,
                                  organismo_participacion,
                                  estilo_encabezado = headerStyle,
                                  estilo_columnas = verticalStyle,
                                  estilo_categorias = bodyStyle,
                                  estilo_horizontal = horizontalStyle,
                                  estilo_total = totalStyle){
  
  # Título pregunta
  
  np <- lista_preguntas[[num_pregunta]]
  
  writeData(wb = wb, sheet =  hojas[1], x = np, startCol = columna,
            startRow = renglon[1], borders = 'none')
  
  writeData(wb = wb, sheet =  hojas[2], x = np, startCol = columna,
            startRow = renglon[2], borders = 'none')
  
  # Tabla Cruzada 
  
  f <- tabla_cruzada_total(diseño = diseño, pregunta = pregunta, datos = datos,
                           DB_Mult = DB_Mult, dominios = dominios, 
                           tipo_pregunta = tipo_pregunta)
  
  if(tipo_pregunta == 'categorica' || tipo_pregunta == 'multiple'){
    
  # Formato Categorías empieza en el renglón 2
  
  formato_categorias(tabla = f, pregunta = pregunta, datos = datos, 
                     DB_Mult = DB_Mult, diseño = diseño,
                     wb = wb, renglon = c((renglon[1] + 1),(renglon[2] + 1)),
                     columna = (columna + 3), hojas = hojas,
                     estilo_cuerpo = estilo_categorias,
                     tipo_pregunta = tipo_pregunta)
  
    setRowHeights(wb = wb, sheet = hojas[1], rows = (renglon[1] + 1), heights = 45)
    
    setRowHeights(wb = wb, sheet = hojas[2], rows = (renglon[2] + 1), heights = 45)
    
    # Escribir tabla Cruzada empieza en el renglón 3
    
    formato_tablas_cruzadas(tabla = f, wb = wb,
                            renglon = c((renglon[1] + 2), (renglon[2] + 2)),
                            columna = columna, hojas = hojas,
                            estilo_encabezado = estilo_encabezado,
                            estilo_total  = estilo_total)
    
    # Estilo dominios empieza en el renglón 3
    
    estilo_dominios(tabla = f, wb = wb, columna = columna, dominios = dominios,
                    hojas = hojas, renglon = c((renglon[1] + 2), (renglon[2] + 2)),
                    estilo_dominios = estilo_horizontal, tipo_pregunta = tipo_pregunta)
    # Estilo columnas 
    
    estilo_columnas(tabla = f, hojas = hojas, wb = wb, estilo = estilo_columnas, 
                    renglon = c(renglon[1], renglon[2]), tipo_pregunta = tipo_pregunta)
    
    addStyle(wb = wb, sheet = hojas[2], style = estilo_total,
             rows = renglon[2]:(renglon[2] + nrow(f[[2]])), cols = 3,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = hojas[1], style = estilo_total,
             rows = renglon[1]:(renglon[1] + nrow(f[[1]])), cols = 3,
             gridExpand = TRUE, stack = TRUE)
    
    # Fuente
    
    fuente_preg <- paste0("Fuente: ", fuente)
    organismo <- paste0("Organismo de participacion ", organismo_participacion)
    tabla_preg <- paste0("Tabla correspondiente a la pregunta ", num_pregunta,
                         " de ", fuente)

    writeData(wb = wb, sheet = hojas[1], x = fuente_preg,
              startRow = (renglon[1] + 3 + nrow(f[[1]])), startCol = columna,
              borders = 'none')
    
    writeData(wb = wb, sheet = hojas[1], x = organismo,
              startRow = (renglon[1] + 4 + nrow(f[[1]])), startCol = columna,
              borders = 'none')
    
    writeData(wb = wb, sheet = hojas[1], x = tabla_preg,
              startRow = (renglon[1] + 5 + nrow(f[[1]])), startCol = columna,
              borders = 'none')
    
    writeData(wb = wb, sheet = hojas[2], x = fuente_preg,
              startRow = (renglon[2] + 3 + nrow(f[[2]])), startCol = columna,
              borders = 'none')
    
    writeData(wb = wb, sheet = hojas[2], x = organismo,
              startRow = (renglon[2] + 4 + nrow(f[[2]])), startCol = columna,
              borders = 'none')
    
    writeData(wb = wb, sheet = hojas[2], x = tabla_preg,
              startRow = (renglon[2] + 5 + nrow(f[[2]])), startCol = columna,
              borders = 'none')
    
  }
  
  if(tipo_pregunta == 'continua'){
    
    # Escribir tabla Cruzada empieza en el renglón 3
    
    formato_tablas_cruzadas(tabla = f, wb = wb, 
                            renglon = c((renglon[1] + 1), (renglon[2] + 1)),
                            columna = columna, hojas = hojas,
                            estilo_encabezado = estilo_encabezado,
                            estilo_total  = estilo_total)
    
    # Estilo dominios empieza en el renglón 3
    
    estilo_dominios(tabla = f, wb = wb, columna = columna, dominios = dominios,
                    hojas = hojas, renglon = c((renglon[1] + 1), (renglon[2] + 1)),
                    estilo_dominios = estilo_horizontal, 
                    tipo_pregunta = tipo_pregunta)
    
    # Estilo columnas 
    
    estilo_columnas(tabla = f, hojas = hojas, wb = wb, estilo = estilo_columnas, 
                    renglon = c(renglon[1], renglon[2]), tipo_pregunta = tipo_pregunta)
    
    
    # Fuente
    
    fuente_preg <- paste0("Fuente: ", fuente)
    organismo <- paste0("Organismo de participacion ", organismo_participacion)
    tabla_preg <- paste0("Tabla correspondiente a la pregunta ", num_pregunta,
                         " de ", fuente)
    
    writeData(wb = wb, sheet = hojas[1], x = fuente_preg,
              startRow = (renglon[1] + 2 + nrow(f[[1]])), startCol = columna,
              borders = 'none')
    
    writeData(wb = wb, sheet = hojas[1], x = organismo,
              startRow = (renglon[1] + 3 + nrow(f[[1]])), startCol = columna,
              borders = 'none')
    
    writeData(wb = wb, sheet = hojas[1], x = tabla_preg,
              startRow = (renglon[1] + 4 + nrow(f[[1]])), startCol = columna,
              borders = 'none')
    
    writeData(wb = wb, sheet = hojas[2], x = fuente_preg,
              startRow = (renglon[2] + 2 + nrow(f[[2]])), startCol = columna,
              borders = 'none')
    
    writeData(wb = wb, sheet = hojas[2], x = organismo,
              startRow = (renglon[2] + 3 + nrow(f[[2]])), startCol = columna,
              borders = 'none')
    
    writeData(wb = wb, sheet = hojas[2], x = tabla_preg,
              startRow = (renglon[2] + 4 + nrow(f[[2]])), startCol = columna,
              borders = 'none')
    
    addStyle(wb = wb, sheet = hojas[2], style = estilo_total,
             rows = renglon[2]:(renglon[2] + nrow(f[[2]])), cols = 3,
             gridExpand = TRUE, stack = TRUE)
    
    addStyle(wb = wb, sheet = hojas[1], style = estilo_total,
             rows = renglon[1]:(renglon[1] + nrow(f[[1]])), cols = 3,
             gridExpand = TRUE, stack = TRUE)
    
  }
  
  return(f)
  
}



preguntas <- function(pregunta, num_pregunta, datos, DB_Mult, dominios,
                      lista_preguntas, diseño, wb, renglon_fs, renglon_tc,
                      columna = 1, hojas_fs = c(1,2),
                      hojas_tc = c(3,4),
                      fuente = 'fuente',
                      organismo_participacion = 'organismo',
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
                                     estilo_encabezado = estilo_encabezado,
                                     estilo_categorias = estilo_categorias,
                                     estilo_horizontal = estilo_horizontal,
                                     estilo_total = estilo_total)
    
  }else{
    tc_rows <- tibble()
  }
  
  return(list(fs_rows, tc_rows))
}

