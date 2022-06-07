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
#' @examples \dontrun{
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
#' @import srvyr
#' @import dplyr
#'
#' @export

total_general <- function(diseño, pregunta, DB_Mult, datos, dominio = 'General',
                          tipo_pregunta = 'categorica', na.rm = TRUE,
                          vartype = c("se","ci","cv", "var"),
                          cuantiles =  c(0,0.25, 0.5, 0.75,1),
                          significancia = 0.95, proporcion = FALSE,
                          metodo_prop = "likelihood", DEFF = TRUE){

if (tipo_pregunta == 'categorica'){

  estadisticas <- {{diseño}} %>%
    srvyr::filter(!is.na(!!sym(pregunta))) %>%
    srvyr::group_by(!!sym(pregunta)) %>%
    srvyr::summarize(
      prop = srvyr::survey_mean(
        na.rm = na.rm,
        vartype = vartype,
        level = significancia,
        proportion = proporcion,
        prop_method = metodo_prop,
        deff = DEFF
      ),
      total = round(srvyr::survey_total(
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
      srvyr::filter(!is.na(!!sym(pregunta))) %>%
      srvyr::summarise(
        prop = srvyr::survey_mean(
          as.numeric(!!sym(pregunta)),
          na.rm = na.rm,
          vartype = vartype,
          level = significancia,
          proportion = proporcion,
          prop_method = metodo_prop,
          deff = DEFF
        ),
        cuantiles = srvyr::survey_quantile(
          as.numeric(!!sym(pregunta)),
          quantiles = cuantiles,
          na.rm = na.rm
        ),
        total = srvyr::survey_total(
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
                                 tipo_pregunta = tipo_pregunta)

    freqs <- formatear_frecuencias_simples(tabla = freqs,
                                           tipo_pregunta = tipo_pregunta)

    freqs[[1]] %<>% filter(Respuesta != 'TOTAL')%>%
      select(- `Casos`, -`Media acumulada`)

    freqs[[2]] %<>% filter(Respuesta != 'TOTAL')%>%
      select(- `Casos`)


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
