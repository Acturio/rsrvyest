# FUNCIÓN ESTILO DOMINIOS, DA FORMATO A LA COLUMNA DE DOMINIOS

#' Da formato a la columna de dominios de las tablas cruzadas en el workbook
#' @description Función para unir las celdas donde se encuentran los nombres de los dominios es las hojas solicitadas.
#' @usage estilo_dominios(
#' tabla,
#' wb,
#' columna = 1,
#' hojas = c(3,4),
#' dominios,
#' renglon = c(1,1),
#' estilo_dominios = horizontalStyle,
#' estilo_merge_dominios = bodyStyle,
#' tipo_pregunta
#' )
#' @param tabla Lista con las tablas cruzadas formateadas
#' @param wb Workbook en el que aplicará el formato
#' @param columna Columna en la que se aplicará el formato
#' @param hojas Hojas del workbook a las que se aplicará el formato, una por tabla
#' @param dominios Vector con los dominios
#' @param renglon Renglón donde se iniciará el formato, uno por tipo de tabla, uno por tabla
#' @param estilo_dominios Estilo previamente definido con el borde en la parte inferior de la tabla
#' @param estilo_merge_dominios Estilo previamente definido con bordes en toda la tabla
#' @param tipo_pregunta Tipo de pregunta ("categorica", "multiple")
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{survey_mean}}
#' @examples \dontrun{
#' estilo_dominios(tabla = tabla_cruzada, wb = wb, columna = 1, hojas = c(3,4),
#' dominios = dominios, renglon = c(1,1), estilo_dominios = horizontalStyle,
#' estilo_merge_dominios = bodyStyle, tipo_pregunta = 'categorica')
#' }
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
