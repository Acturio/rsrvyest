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
#' @examples \dontrun{
#' dataset <- read.spss("data/BASE_CONACYT_260118.sav", to.data.frame = TRUE)
#' Lista_Preg <- read_xlsx("aux/Lista de Preguntas.xlsx",
#'   sheet = "Lista Preguntas"
#' )$Nombre %>% as.vector()
#' DB_Mult <- read_xlsx("aux/Lista de Preguntas.xlsx", sheet = "Múltiple") %>% as.data.frame()
#' Lista_Cont <- read_xlsx("aux/Lista de Preguntas.xlsx",
#'   sheet = "Continuas"
#' )$VARIABLE %>% as.vector()
#' Dominios <- read_xlsx("aux/Lista de Preguntas.xlsx", sheet = "Dominios")$Dominios %>% as.vector()
#'
#' disenio_mult <- disenio(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1, reps = FALSE, datos = dataset)
#'
#' tc <- tablas_cruzadas(
#'   diseño = disenio_mult, pregunta = "P1", dominio = "Sexo", datos = dataset,
#'   DB_Mult = DB_Mult, tipo_pregunta = "multiple"
#' )
#'
#' formatear_tabla_cruzada(pregunta = "P1", datos = dataset, dominio = "Sexo", tabla = tc, DB_Mult = DB_Mult, tipo_pregunta = "multiple")
#' }
#' @import dplyr
#' @import purrr
#' @import stringr
#' @export
formatear_tabla_cruzada <- function(pregunta, datos, dominio, tabla, DB_Mult,
                                    tipo_pregunta = "categorica") {

  if (tipo_pregunta == "categorica") {

    categorias <- datos %>%
      pull(!!sym(pregunta)) %>%
      levels() %>%
      str_trim(side = "both") #%>%
      #str_c("_")

    nombres1 <- c("Dominio", "Categorías", "Total", rep(
      c("Media", "Lim. inf.", "Lim. sup."),
      length(categorias)
    ))

    nombres2 <- c("Dominio", "Categorías", "Total", rep(
      c("Err. Est.", "Coef. Var.", "Var.", "DEFF"),
      length(categorias)
    ))

    # Orden similar a las categorias
    tabla %<>% select(Dominio, Categorias, starts_with(categorias))

    # Multiplicar por 100 y 10,000

    tabla %<>%
      map_if(str_detect(names(tabla), "_prop($|_low$|_upp$|_se$)"), ~ .x * 100) %>%
      map_if(str_detect(names(tabla), "_prop_var"), ~ .x * 10000) %>%
      as_tibble()

    tabla %<>%
      rowwise() %>%
      mutate(Total = sum(c_across(ends_with("_total")), na.rm = TRUE)) %>%
      ungroup() %>%
      relocate(Total, .after = Categorias)

    # Primera tabla
    tabla_final_1 <- tabla %>%
      select(
        Dominio,
        Categorias,
        Total,
        ends_with(c("_prop", "_prop_low", "_prop_upp"))
      ) %>%
      select(Dominio, Categorias, Total, starts_with(categorias))

    # Segunda tabla
    tabla_final_2 <- tabla %>%
      select(
        Dominio,
        Categorias,
        Total,
        ends_with(c("_prop_se", "_prop_cv", "_prop_var", "prop_deff"))
      ) %>%
      select(Dominio, Categorias, Total, starts_with(categorias))

    names(tabla_final_1) <- nombres1
    names(tabla_final_2) <- nombres2
  }

  if (tipo_pregunta == "continua") {
    tabla_final_1 <- tabla %>%
      select(
        Dominio, Categorias, total, prop, prop_low, prop_upp, cuantiles_q00,
        cuantiles_q25, cuantiles_q50, cuantiles_q75, cuantiles_q100
      ) %>%
      dplyr::rename(
        "Total" = total,
        "Categorías" = Categorias,
        "Media" = prop,
        "Lim. inf" = prop_low,
        "Lim. sup" = prop_upp,
        "Mín" = cuantiles_q00,
        "Q25" = cuantiles_q25,
        "Mediana" = cuantiles_q50,
        "Q75" = cuantiles_q75,
        "Máx" = cuantiles_q100
      )

    tabla_final_2 <- tabla %>%
      select(Dominio, Categorias, total, prop_se, prop_cv, prop_var, prop_deff) %>%
      dplyr::rename(
        "Err. Est." = prop_se,
        "Total" = total,
        "Categorías" = Categorias,
        "Var" = prop_var,
        "Coef. Var." = prop_cv,
        "DEFF" = prop_deff
      )
  }

  if (tipo_pregunta == "multiple") {
    tabla1 <- tabla %>%
      map_at(c("prop", "prop_low", "prop_upp", "prop_se"), ~ .x * 100) %>%
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
      select(
        Dominio,
        Categorias,
        "Total",
        "Media", "Lim. inf.", "Lim. sup.", -"Total"
      )

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

    suma_totales <- tabla %>%
      dplyr::group_by(Categorias) %>%
      dplyr::summarise(Total = sum(total))

    len <- datos %>%
      select(!!sym(dominio)) %>%
      distinct() %>%
      nrow()

    ps <- DB_Mult %>%
      dplyr::filter(!is.na(!!sym(pregunta))) %>%
      dplyr::pull(!!sym(pregunta))

    df <- select(datos, ps)

    categorias <- df %>%
      pull() %>%
      levels() %>%
      str_trim(side = 'both')

    nombres_tabla1 <- c(
      "Dominio", "Categorías", "Total",
      rep(
        c("Media", "Lim. inf.", "Lim. sup."),
        length(categorias)
      )
    )

    nombres_tabla2 <- c(
      "Dominio", "Categorías", "Total",
      rep(
        c("Err. Est.", "Coef. Var.", "Var.", "DEFF"),
        length(categorias)
      )
    )


    tabla_final_1 <- tabla1[c(1:len), ]

    for (i in seq((len + 1), nrow(tabla1), by = len)) {
      tabla_c1 <- tabla1[i:(i + len - 1), c(-1, -2)]
      tabla_final_1 <- cbind(tabla_final_1, tabla_c1)
    }

    suma_totales %<>% select(Total)

    tabla_final_1 <- cbind(tabla_final_1, suma_totales) %>%
      relocate(Total, .after = Categorias)

    tabla_final_2 <- tabla2[c(1:len), ]

    for (i in seq((len + 1), nrow(tabla2), by = len)) {
      tabla_c2 <- tabla2[i:(i + len - 1), c(-1, -2)]
      tabla_final_2 <- cbind(tabla_final_2, tabla_c2)
    }

    tabla_final_2 <- cbind(tabla_final_2, suma_totales) %>%
      relocate(Total, .after = Categorias)

    names(tabla_final_1) <- nombres_tabla1
    names(tabla_final_2) <- nombres_tabla2
  }


  return(list(tabla_final_1, tabla_final_2))
}
