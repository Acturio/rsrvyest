#' Función tablas cruzadas según tipo de pregunta y dominio
#' @description Se crea la tabla cruzada según dominio tipo de pregunta (categórica, múltiple o continua)
#' @usage tablas_cruzadas(diseño,
#' pregunta,
#' dominio,
#' datos,
#' DB_Mult,
#' na.rm,
#' vartype = c("ci","se","var","cv"),
#' cuantiles = c(0,0.25, 0.5, 0.75,1),
#' significancia = 0.95,
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
#' @examples \dontrun{
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
#'  @import srvyr
#'  @import dplyr
#'  @import tidyr
#'  @importFrom  caret dummyVars
#'  @export
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
