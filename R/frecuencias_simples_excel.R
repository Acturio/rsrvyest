# FUNCIÓN FRECUENCIAS SIMPLES EN EXCEL

#' Función frecuencias simples en Excel
#'
#' @description Escribe los títulos 'Frecuencias simples' y 'Frecuencias simples (dispersión)', logo indicado, título de la pregunta, tabla de frecuencias simples formateada y pie de tabla en las hojas y renglones mencionados por el usuario
#' @usage frecuencias_simples_excel(
#' pregunta,
#' num_pregunta,
#' datos,
#' DB_Mult,
#' lista_preguntas,
#' diseño,
#' wb,
#' renglon = c(1,1),
#' columna = 1,
#' hojas = c(1,2),
#' tipo_pregunta = 'categorica',
#' fuente,
#' logo_path,
#' organismo_participacion,
#' estilo_encabezado = headerStyle,
#' estilo_horizontal = horizontalStyle,
#' estilo_total = totalStyle
#'  )
#' @param pregunta Nombre de la pregunta sobre la cual se desea obtener frecuencias simples e incluirlas en un workbook de Excel
#' @param num_pregunta Número de pregunta
#' @param datos Conjunto de datos en formato .sav
#' @param DB_Mult Data frame con las preguntas múltiples
#' @param lista_preguntas Data frame que contiene los títulos de las pregunta
#' @param diseño Diseño muestral que se ocupará según el tipo de pregunta
#' @param wb Workbook de Excel que contiene al menos dos hojas
#' @param renglon Vector tamaño 2 especificando el número de renglon en el cual se desea empezar a escribir la tabla 1 y tabla 2 respectivamente
#' @param columna Columna en la cual se desea empezar a escribir las tablas
#' @param hojas Vector de número de hojas en el cual se desea insertar las tablas
#' @param tipo_pregunta Tipo de pregunta_ 'categorica', 'multiple', 'continua'
#' @param fuente Nombre del proyecto
#' @param logo_path Path del logo de la UNAM
#' @param organismo_participacion Organismos que participaron en el proyecto, por ejemplo, 'Ciudadanía Mexicana'
#' @param estilo_encabezado estilo el cual se desea usar para los nombres de las columnas
#' @param estilo_horizontal estilo último renglón horizontal
#' @param estilo_total estilo el cual se desea usar para la columna total
#'
#' @details Esta función envuelve todas las funciones creadas para obtener las frecuencias simples, por lo que esta función es la única que se deberá llamar para crear las frecuencias simples de las preguntas deseadas e insertarlas en ciertas hojas de Excel
#' @details Es necesario crear al menos dos hojas de excel con la función addWorksheet de la paquetería openxlsx
#' @details El estilo_total se recomienda crear un estilo con la función createStyle de openxlsx con el formato que se desea, por ejemplo "###,###,###.0"
#' @details El estilo_horizontal hace referencia al tipo de lineado horizontal se desea en el úntimo renglón de la tabla
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{writeData}} \code{\link{createStyle}}  \code{\link{setRowHeights}} \code{\link{insertImage}} \code{\link{mergeCells}}
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
#' totalStyle <-  createStyle(numFmt = "###,###,###.0")
#' horizontalStyle <- createStyle(border = "bottom",
#' borderColour = "black", borderStyle = 'thin', valign = 'center')
#'
#' # Carga de datos
#'  dataset <- read.spss("data/BASE_CONACYT_260118.sav", to.data.frame = TRUE)
#'  Lista_Preg <- read_xlsx("aux/Lista de Preguntas.xlsx",
#'  sheet = "Lista Preguntas")$Nombre %>% as.vector()
#'   DB_Mult <- read_xlsx("aux/Lista de Preguntas.xlsx",
#'   sheet = "Múltiple") %>% as.data.frame()
#'
#' #Diseño
#'  disenio_mult <- disenio(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1, reps=FALSE, datos = dataset)
#'
#' frecuencias_simples_excel(pregunta = 'P1',
#' num_pregunta = 1,
#' datos = dataset,
#' DB_Mult = DB_Mult,
#' lista_preguntas = Lista_Preg,
#' diseño = disenio_mult,
#' wb = wb,
#' renglon = c(1,1),
#' columna = 1,
#' hojas = c(1,2),
#' tipo_pregunta = 'multiple',
#' fuente =  'Conacyt 2018',
#' organismo_participacion = 'Ciudadanía mexicana',
#' estilo_encabezado = headerStyle,
#' estilo_horizontal = horizontalStyle,
#' estilo_total = totalStyle
#' )
#' }
#' @export
frecuencias_simples_excel <- function(pregunta, num_pregunta, datos, DB_Mult,
                                      lista_preguntas, diseño, wb, renglon = c(1,1),
                                      columna = 1, hojas = c(1,2),
                                      tipo_pregunta = 'categorica', fuente,
                                      organismo_participacion, logo_path = 'imagen.png',
                                      estilo_encabezado = headerStyle,
                                      estilo_horizontal = horizontalStyle,
                                      estilo_total = totalStyle){


  titleStyle <- createStyle(fontSize = 24, fontColour = '#011f4b',
                            textDecoration = 'underline')

  subtitleStyle <- createStyle(fontSize = 14, fontColour = '#011f4b',
                               textDecoration = 'underline')

  table_Style <- createStyle(valign = 'center', halign = 'center')

  # Título general renglón 1 columna 2

  writeData(wb = wb, sheet = hojas[1], x = 'Frecuencias estadísticas',
            startCol = 2, startRow = 1)
  writeData(wb = wb, sheet = hojas[1], x = nombre_proyecto,
            startCol = 2, startRow = 2)
  addStyle(wb = wb, sheet = hojas[1], style = titleStyle, rows = 1, cols = 2,
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb = wb, sheet = hojas[1], style = subtitleStyle, rows = 2, cols = 2,
           gridExpand = TRUE, stack = TRUE)


  writeData(wb = wb, sheet = hojas[2], x = 'Frecuencias estadísticas (dispersión)',
            startCol = 2, startRow = 1)
  writeData(wb = wb, sheet = hojas[2], x = nombre_proyecto,
            startCol = 2, startRow = 2)
  addStyle(wb = wb, sheet = hojas[2], style = titleStyle, rows = 1, cols = 2,
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb = wb, sheet = hojas[2], style = subtitleStyle, rows = 2, cols = 2,
           gridExpand = TRUE, stack = TRUE)

  # Pegar logo UNAM renglón 1, columna 1


  #logo <- logo_path

  insertImage(wb = wb, sheet = hojas[1],
              file = logo_path, startRow = 1, startCol = 1,
              width = 2.08, height =2.2, units = 'cm')

  insertImage(wb = wb, sheet = hojas[2],
              file = logo_path, startRow = 1, startCol = 1,
              width = 2.08, height = 2.2, units = 'cm')

  # Título Pregunta

  np <- lista_preguntas[[num_pregunta]]

  writeData(wb = wb, sheet = hojas[1], x = np, startRow = (renglon[1]),
            startCol = columna, colNames = TRUE, borders = 'none')

  setRowHeights(wb = wb, sheet = hojas[1], rows = (renglon[1]), heights = 35)


  writeData(wb = wb, sheet = hojas[2], x = np, startRow = (renglon[2]),
            startCol = columna, colNames = TRUE, borders = 'none')

  setRowHeights(wb = wb, sheet = hojas[2], rows = (renglon[2] + 1), heights = 35)


  # Tabla

  tabla_titulo <- paste0('Tabla ', num_pregunta, '*')

  writeData(wb = wb, sheet = hojas[1], x = tabla_titulo, startRow = (renglon[1] +1),
            startCol = columna, colNames = FALSE, borders = 'none')

  setRowHeights(wb = wb, sheet = hojas[1], rows = (renglon[1] + 1), heights = 30)



  writeData(wb = wb, sheet = hojas[2], x = tabla_titulo, startRow = (renglon[2] +1),
            startCol = columna, colNames = FALSE, borders = 'none')

  setRowHeights(wb = wb, sheet = hojas[2], rows = (renglon[1] + 1), heights = 30)


  est <- frecuencias_simples(diseño = diseño, datos = datos, pregunta = pregunta,
                             DB_Mult = DB_Mult, tipo_pregunta = tipo_pregunta)

  freq <-  formatear_frecuencias_simples(est, tipo_pregunta = tipo_pregunta)

  formato_frecuencias_simples(tabla = freq, wb = wb, hojas = hojas,
                              renglon = c((renglon[1] + 1 + 1), (renglon[2] + 1 + 1)),
                              columna = columna, estilo_encabezado = estilo_encabezado,
                              estilo_horizontal = estilo_horizontal,
                              estilo_total = estilo_total,
                              tipo_pregunta = tipo_pregunta)


  # Merge

  addStyle(wb = wb, sheet = hojas[1], style = table_Style,
           rows = (renglon[1] + 1), cols = 1:ncol(freq[[1]]),
           gridExpand = TRUE, stack = TRUE)

  mergeCells(wb = wb, sheet = hojas[1], cols = 1:ncol(freq[[1]]),
             rows = (renglon[1] + 1))


  addStyle(wb = wb, sheet = hojas[2], style = table_Style,
           rows = (renglon[2] + 1), cols = 1:ncol(freq[[2]]),
           gridExpand = TRUE, stack = TRUE)

  mergeCells(wb = wb, sheet = hojas[2], cols = 1:ncol(freq[[2]]),
             rows = (renglon[2] + 1))


  # Fuente

  fuente_preg <- paste0("Fuente: ", fuente)
  organismo <- paste0("Organismo de participacion ", organismo_participacion)
  tabla_preg <- paste0("* Tabla correspondiente a la pregunta ", num_pregunta,
                       " de ", fuente)

  fontStyle <- createStyle(fontSize = 9, fontColour = '#5b5b5b')


  writeData(wb = wb, sheet = hojas[1], x = fuente_preg,
            startRow = (renglon[1] + 2 + 1 + nrow(freq[[1]])),
            startCol = columna, borders = 'none')

  addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
           rows = (renglon[1] + 2 + 1 + nrow(freq[[1]])), cols = columna,
           gridExpand = TRUE, stack = TRUE)

  setRowHeights(wb = wb, sheet = hojas[1],
                rows = (renglon[1] + 2 + 1 + nrow(freq[[1]])), heights = 10)



  writeData(wb = wb, sheet = hojas[1], x = organismo,
            startRow = (renglon[1] + 3+ 1 + nrow(freq[[1]])),
            startCol = columna, borders = 'none')

  addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
           rows = (renglon[1] + 3 + 1 + nrow(freq[[1]])), cols = columna,
           gridExpand = TRUE, stack = TRUE)

  setRowHeights(wb = wb, sheet = hojas[1],
                rows = (renglon[1] + 3 + 1 + nrow(freq[[1]])), heights = 10)


  writeData(wb = wb, sheet = hojas[1], x = tabla_preg,
            startRow = (renglon[1] + 4 + 1+ nrow(freq[[1]])),
            startCol = columna, borders = 'none')

  addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
           rows = (renglon[1] + 4 + 1 + nrow(freq[[1]])), cols = columna,
           gridExpand = TRUE, stack = TRUE)

  setRowHeights(wb = wb, sheet = hojas[1],
                rows = (renglon[1] + 4 + 1 + nrow(freq[[1]])), heights = 10)




  writeData(wb = wb, sheet = hojas[2], x = fuente_preg,
            startRow = (renglon[2] + 2+1 + nrow(freq[[2]])),
            startCol = columna, borders = 'none')

  addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
           rows = (renglon[2] + 2 + 1 + nrow(freq[[2]])), cols = columna,
           gridExpand = TRUE, stack = TRUE)

  setRowHeights(wb = wb, sheet = hojas[2],
                rows = (renglon[2] + 2 + 1 + nrow(freq[[2]])), heights = 10)


  writeData(wb = wb, sheet = hojas[2], x = organismo,
            startRow = (renglon[2] + 3 +1+ nrow(freq[[2]])),
            startCol = columna, borders = 'none')

  addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
           rows = (renglon[2] + 3 + 1 + nrow(freq[[2]])), cols = columna,
           gridExpand = TRUE, stack = TRUE)

  setRowHeights(wb = wb, sheet = hojas[2],
                rows = (renglon[2] + 3 + 1 + nrow(freq[[2]])), heights = 10)


  writeData(wb = wb, sheet = hojas[2], x = tabla_preg,
            startRow = (renglon[2] + 4+1 + nrow(freq[[2]])),
            startCol = columna, borders = 'none')

  addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
           rows = (renglon[2] + 4 + 1 + nrow(freq[[2]])), cols = columna,
           gridExpand = TRUE, stack = TRUE)

  setRowHeights(wb = wb, sheet = hojas[2],
                rows = (renglon[2] + 4 + 1 + nrow(freq[[2]])), heights = 10)

  return(freq)

}
