# FUNCIÓN ESTILO COLUMNAS, DIVIDE LAS MÉTRICAS DE LAS TABLAS CRUZADAS POR
# CATEGORÍAS RESPUESTAS SOLO PREGUNTAS CATEGÓRICAS Y MÚLTIPLES

#'  Divide las métricas de las tablas cruzadas
#'
#' @description Divide las métricas de las tablas cruzadas por categoría de respuestas en el caso de preguntas catégoricas y múltiples.
#' @usage estilo_columnas(
#' tabla,
#' wb,
#' hojas =  c(3,4),
#' estilo = verticalStyle,
#' renglon = c(1,1),
#' tipo_pregunta = "categorica"
#' )
#' @param tabla Tabla cruzada creada con la función tablas_cruzadas()
#' @param wb wb Workbook en el que aplicará el estilo de columnas
#' @param hojas Indicador de las hojas del workbook en las que se aplicará el estilo de columnas
#' @param estilo Un objeto de estilo devuelto por la función createStyle()
#' @param renglon Renglón donde se iniciará el formato
#' @param tipo_pregunta Tipo de pregunta: 'categorica', 'multiple', 'continua'
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{createStyle}} \code{\link{addStyle}}
#' @example \dontrun{
#' tabla_cruzada <- tablas_cruzadas(
#'   diseño = disenio_mult,
#'   pregunta = 'P1',
#'   dominio = 'Sexo',
#'   datos = dataset,
#'   DB_Mult = DB_Mult,
#'   tipo_pregunta = 'multiple'
#'   )
#'
#' estilo_columnas(
#'   tabla = tabla_cruzada,
#'   wb = wb,
#'   hojas = c(3,4),
#'   estilo = verticalStyle,
#'   renglon = c(1,1),
#'   tipo_pregunta = 'categorica'
#'   )
#' }
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
