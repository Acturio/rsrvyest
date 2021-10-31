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
#' @examples \dontrun{
#' frecuencias_simples(diseño = disenio_cat, datos = dataset, pregunta = 'P1',
#'  DB_Mult = DB_Mult, tipo_pregunta = 'categorica')
#' }
#' @import dplyr
#' @import srvyr
#' @rawNamespace import(caret, except = lift)
#' @export
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
      filter(!is.na(!!sym(pregunta))) %>%
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
      filter(!is.na(!!sym(pregunta))) %>%
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
        filter(!is.na(!!sym(categ))) %>%
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

  estadisticas

  return(estadisticas)
}
