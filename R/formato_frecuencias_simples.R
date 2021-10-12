# FUNCIÓN ESCRIBIR FRECUENCIAS SIMPLES EN EXCEL (UTILIZA FUNCIÓN escribir_tabla)

#' Escribe frecuencias simples en dos hojas de Excel
#'
#' @description Escribe la tabla de frecuencias simples formateadas en dos hojas de Excel
#' @usage formato_frecuencias_simples(
#' tabla,
#' wb,
#' hojas,
#' renglon,
#' columna,
#' estilo_encabezado,
#' estilo_horizontal,
#' estilo_total,
#' tipo_pregunta
#' )
#' @param tabla lista de las tibbles creadas por la función formatear_frecuencias_simples
#' @param wb Workbook de Excel que contiene al menos dos hojas
#' @param hojas vector de número de hojas en el cual se desea insertar las tablas
#' @param renglon vetor tamaño 2 especificando el número de renglon en el cual se desea empezar a escribir la tabla 1 y tabla 2 respectivamente
#' @param columna columna en la cual se desea empezar a escribir las tablas
#' @param estilo_encabezado estilo el cual se desea usar para los nombres de las columnas
#' @param estilo_horizontal estilo último renglón horizontal
#' @param estilo_total estilo el cual se desea usar para la columna total
#' @param tipo_pregunta tipo de pregunta_ 'categorica', 'multiple', 'continua'
#' @details Ambas tablas se escribirán en la misma columna pero en diferentes hojas (especificadas en el parámetro hojas).
#' @details El estilo_total se recomienda crear un estilo con la función createStyle de openxlsx con el formato que se desea, por ejemplo "###,###,###.0"
#' @details El estilo_horizontal hace referencia al tipo de lineado horizontal se desea en el úntimo renglón de la tabla
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{writeData}} \code{\link{createStyle}}
#' @examples \dontrun{
#' # Creación del workbook
#' wb <- createWorkbook()
#' addWorksheet(wb, "Frecuencias simples")
#' addWorksheet(wb, "Frecuencias simples (dispersión)")
#'
#' # Estilos
#' headerStyle <- createStyle(fontSize = 11, fontColour = "black", halign = "center",
#' border = "TopBottom", borderColour = "black",
#' borderStyle = c('thin', 'double'), textDecoration = 'bold')
#'
#' totalStyle <-  createStyle(numFmt = "###,###,###.0")
#'
#' horizontalStyle <- createStyle(border = "bottom", borderColour = "black", borderStyle = 'thin', valign = 'center')
#'
#' freq_simples <- frecuencias_simples(diseño = disenio_cat, datos = dataset, pregunta = 'P1',
#'  DB_Mult = DB_Mult, tipo_pregunta = 'categorica')
#'
#' freq_simples_format <- frecuencias_simples(tabla = freq_simples, tipo_pregunta = 'categorica')
#'
#' formato_frecuencias_simples(tabla = freq_simples_format , wb = wb, hojas = c(1,2),
#' renglon = c(1,1), columna = 1 , estilo_encabezado = headerStyle ,
#' estilo_horizontal = horizontalStyle , estilo_total = totalStyle, tipo_pregunta = 'categorica')
#' }
#' @import openxlsx
#' @export
formato_frecuencias_simples <- function(tabla, wb, hojas = c(1,2), renglon = c(1,1),
                                        columna, estilo_encabezado,
                                        estilo_horizontal, estilo_total,
                                        tipo_pregunta = 'categorica'){


  # Primera Página renglón 2 columna 1

  writeData(wb = wb, sheet = hojas[1], x = tabla[[1]], startRow = renglon[1],
            startCol = columna, colNames = TRUE, borders = 'none', keepNA = TRUE,
            na.string = '-', headerStyle = estilo_encabezado)

  # Ultimo renglón
  addStyle(wb = wb, sheet = hojas[1], style = estilo_horizontal,
           rows = (renglon[1] + nrow(tabla[[1]])),
           cols = columna:(ncol(tabla[[1]])), gridExpand = TRUE, stack = TRUE)


  writeData(wb = wb, sheet = hojas[2], x = tabla[[2]], startRow = renglon[2],
            startCol = columna, colNames = TRUE, borders = 'none', keepNA = TRUE,
            na.string = '-', headerStyle = estilo_encabezado)

  # Ultimo renglón
  addStyle(wb = wb, sheet = hojas[2], style = estilo_horizontal,
           rows = (renglon[2] + nrow(tabla[[2]])), cols = columna:(ncol(tabla[[2]])),
           gridExpand = TRUE, stack = TRUE)

  # Estilo total

  if (tipo_pregunta == 'categorica' || tipo_pregunta == 'multiple'){

    # Columna total
    addStyle(wb = wb, sheet = hojas[1], style = estilo_total,
             rows = renglon[1]:(renglon[1] + nrow(tabla[[1]])), cols = 2,
             gridExpand = TRUE, stack = TRUE)

    # Columna total
    addStyle(wb = wb, sheet = hojas[2], style = estilo_total,
             rows = renglon[2]:(renglon[2] + 1 + nrow(tabla[[2]])), cols = 2,
             gridExpand = TRUE, stack = TRUE)

  }

  if (tipo_pregunta == 'continua'){

    # Columna total
    addStyle(wb = wb, sheet = hojas[1], style = estilo_total,
             rows = (renglon[1] + 1), cols = 2,
             gridExpand = TRUE, stack = TRUE)

    # Columna total
    addStyle(wb = wb, sheet = hojas[2], style = estilo_total,
             rows = (renglon[2] + 1), cols = 2,
             gridExpand = TRUE, stack = TRUE)

  }

}
