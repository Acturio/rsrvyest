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
#' @examples \dontrun{
#' freq_simples <- frecuencias_simples(diseño = disenio_cat, datos = dataset, pregunta = 'P1',
#'  DB_Mult = DB_Mult, tipo_pregunta = 'categorica')
#'
#' frecuencias_simples(tabla = freq_simples, tipo_pregunta = 'categorica')
#' }
#' @import reshape2
#' @import dplyr
#' @export
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

    if (nrow(tabla) != 0){

      tabla1 %<>% add_row("Respuesta" = 'TOTAL',
                          "Total" = sum(tabla1$Total),
                          "Media" = sum(tabla1$Media),
                          "Lim. inf." = NA,
                          "Lim. sup." = NA
                          )
    }
    else{
      tabla1
    }

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

    if (nrow(tabla) != 0){

      tabla2 %<>% add_row("Respuesta" = 'TOTAL',
                          "Total" = sum(tabla2$Total),
                          "Err. Est." = NA ,
                          "Coef. Var." = NA,
                          "Var." = NA,
                          "DEFF" = NA
                          )
    }
    else{
      tabla2
    }


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
