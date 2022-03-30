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
#' @examples \dontrun{
#' # Lectura de datos
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
#' # Diseño
#' disenio_mult <- disenio(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1, reps = FALSE, datos = dataset)
#'
#' tabla_cruzada_total(
#'   diseño = disenio_mult, pregunta = "P1", dominios = Dominios, datos = dataset,
#'   DB_Mult = DB_Mult, tipo_pregunta = "multiple"
#' )
#' }
#' @export
tabla_cruzada_total <- function(diseño, pregunta, datos, DB_Mult,
                                dominios = Dominios,
                                tipo_pregunta = "categorica") {

  # Tabla nacional
  gral <- total_general(
    diseño = diseño, datos = datos, DB_Mult = DB_Mult,
    pregunta = pregunta, tipo_pregunta = tipo_pregunta
  )

  # Dominios de dispersión

  nacional[[1]] <- tibble()
  nacional[[2]] <- tibble()

  for (d in dominios) {
    f_n <- tablas_cruzadas(
      diseño = diseño, pregunta = pregunta, datos = datos,
      DB_Mult = DB_Mult, dominio = d,
      tipo_pregunta = tipo_pregunta
    )

    f_tc <- formatear_tabla_cruzada(
      pregunta = pregunta, datos = datos,
      dominio = d, tabla = f_n, DB_Mult = DB_Mult,
      tipo_pregunta = tipo_pregunta
    )


    nacional[[1]] <- rbind(nacional[[1]], f_tc[[1]])
    nacional[[2]] <- rbind(nacional[[2]], f_tc[[2]])
  }

  gral[[1]] <- rbind(gral[[1]], nacional[[1]])
  gral[[2]] <- rbind(gral[[2]], nacional[[2]])

  return(gral)
}
