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
#' @examples \dontrun{
#' # Estilos
#' headerStyle <- createStyle(
#'   fontSize = 11, fontColour = "black", halign = "center",
#'   border = "TopBottom", borderColour = "black",
#'   borderStyle = c("thin", "double"), textDecoration = "bold"
#' )
#'
#' totalStyle <- createStyle(numFmt = "###,###,###.0")
#' formato_tablas_cruzadas(
#'   tabla = tabla_cruzada, wb = wb, renglon = c(1, 1),
#'   columna = 1, hojas = c(3, 4), estilo_encabezado = headerStyle,
#'   estilo_total = totalStyle
#' )
#' }
#' @import openxlsx
#' @export
formato_tablas_cruzadas <- function(
  tabla,
  wb,
  renglon = c(1, 1),
  columna = 1,
  hojas = c(3, 4),
  estilo_encabezado = headerStyle,
  estilo_total = totalStyle) {


  # renglón 3 columna 1

  writeData(
    wb = wb, sheet = hojas[1], x = tabla[[1]], startRow = renglon[1],
    startCol = columna, borders = "surrounding",
    borderStyle = "thin", keepNA = TRUE, na.string = "-", colNames = TRUE,
    headerStyle = estilo_encabezado
  )


  writeData(
    wb = wb, sheet = hojas[2], x = tabla[[2]], startRow = renglon[2],
    startCol = columna, borders = "surrounding",
    borderStyle = "thin", keepNA = TRUE, na.string = "-", colNames = TRUE,
    headerStyle = estilo_encabezado
  )
}
