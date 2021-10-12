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
#' @examples \dontrun{
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
#' @import dplyr
#' @importFrom magrittr '%<>%'
#' @export

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
