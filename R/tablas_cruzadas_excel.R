# FUNCIÓN TABLAS CRUZADAS EN EXCEL

#' Función tablas cruzadas en Excel
#' @description Escribe los títulos 'Tablas cruzadas' y 'Tablas cruzadas (dispersión)', logo indicado, título de la pregunta, tabla crzada total y pie de tabla en las hojas y renglones mencionados por el usuario
#' @usage tablas_cruzadas_excel(
#' pregunta,
#' num_pregunta,
#' dominios,
#' datos,
#' DB_Mult,
#' lista_preguntas,
#' diseño,
#' wb,
#' renglon,
#' columna,
#' hojas ,
#' tipo_pregunta,
#' fuente,
#' organismo_participacion,
#' logo_path,
#' estilo_encabezado = headerStyle,
#' estilo_columnas = verticalStyle,
#' estilo_categorias = bodyStyle,
#' estilo_horizontal = horizontalStyle,
#' estilo_total = totalStyle
#' )
#' @param pregunta  Nombre de la pregunta sobre la cual se desea obtener la tabla cruzada e incluirla en un workbook de Excel
#' @param num_pregunta Número de pregunta
#' @param dominios Vector el cual contiene los nombres de los dominios sobre los cuales se desean obtener sus respectivas tablas cruzadas
#' @param DB_Mult Data frame con las preguntas múltiples
#' @param lista_preguntas Data frame que contiene los títulos de las pregunta
#' @param diseño Diseño muestral que se ocupará según el tipo de pregunta
#' @param wb Workbook de Excel que contiene al menos dos hojas
#' @param renglon Vector tamaño 2 especificando el número de renglon en el cual se desea empezar a escribir la tabla 1 y tabla 2 respectivamente
#' @param columna Columna en la cual se desea empezar a escribir las tablas
#' @param hojas Vector de número de hojas en el cual se desea insertar las tablas
#' @param tipo_pregunta Tipo de pregunta_ 'categorica', 'multiple', 'continua'
#' @param fuente Nombre del proyecto
#' @param organismo_participacion Organismos que participaron en el proyecto, por ejemplo, 'Ciudadanía Mexicana'
#' @param logo_path Path del logo de la UNAM
#' @param estilo_encabezado estilo el cual se desea usar para los nombres de las columnas
#' @param estilo_horizontal estilo último renglón horizontal
#' @param estilo_total estilo el cual se desea usar para la columna total
#'
#' @details Esta función envuelve todas las funciones creadas para obtener las tablas cruzadas, por lo que esta función es la única que se deberá llamar para crear las tablas cruzadas de las preguntas deseadas e insertarlas en ciertas hojas de Excel
#' @details Es necesario crear al menos dos hojas de excel con la función addWorksheet de la paquetería openxlsx
#' @details El estilo_total se recomienda crear un estilo con la función createStyle de openxlsx con el formato que se desea, por ejemplo "###,###,###.0"
#' @details El estilo_horizontal hace referencia al tipo de lineado horizontal se desea en el úntimo renglón de la tabla
#' @details El estilo_encabezado hace referencia al tipo de bordes, alineación, color, etc. que se desea obtener en el nombre de las columnas de la tabla cruzada final
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{writeData}} \code{\link{createStyle}}  \code{\link{setRowHeights}} \code{\link{insertImage}} \code{\link{mergeCells}}
#' @examples \dontrun{
#' # Creación del workbook
#' wb <- createWorkbook()
#' addWorksheet(wb, "Tablas cruzadas")
#' addWorksheet(wb, "Tablas cruzadas (dispersión)")
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
#' tablas_cruzadas_excel(
#' pregunta = 'P1',
#' num_pregunta = 1,
#' dominios = Dominios,
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
#' logo_path = '~/Desktop/UNAM/DIAO/rsrvyest/img/logo_unam.png'
#' estilo_encabezado = headerStyle,
#' estilo_horizontal = horizontalStyle,
#' estilo_total = totalStyle
#' )
#' openxlsx::openXL(wb)
#' }
#' @export
tablas_cruzadas_excel <- function(pregunta, num_pregunta, dominios, datos,
                                  DB_Mult, lista_preguntas, diseño, wb,
                                  renglon = c(1,1), columna = 1, hojas = c(3,4),
                                  tipo_pregunta = 'categorica', fuente,
                                  organismo_participacion, logo_path = NULL,
                                  estilo_encabezado = headerStyle,
                                  estilo_columnas = verticalStyle,
                                  estilo_categorias = bodyStyle,
                                  estilo_horizontal = horizontalStyle,
                                  estilo_total = totalStyle){

  # Título pregunta

  titleStyle <- createStyle(fontSize = 24, fontColour = '#011f4b',
                            textDecoration = 'underline')

  subtitleStyle <- createStyle(fontSize = 14, fontColour = '#011f4b',
                               textDecoration = 'underline')

  table_Style <- createStyle(valign = 'center', halign = 'center')

  # Título general renglón 1 columna 2

  writeData(wb = wb, sheet = hojas[1], x = 'Tablas cruzadas',
            startCol = 2, startRow = 1)
  writeData(wb = wb, sheet = hojas[1], x = organismo_participacion,
            startCol = 2, startRow = 2)
  addStyle(wb = wb, sheet = hojas[1], style = titleStyle, rows = 1, cols = 2,
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb = wb, sheet = hojas[1], style = subtitleStyle, rows = 2, cols = 2,
           gridExpand = TRUE, stack = TRUE)


  writeData(wb = wb, sheet = hojas[2], x = 'Tablas cruzadas (dispersión)',
            startCol = 2, startRow = 1)
  writeData(wb = wb, sheet = hojas[2], x = nombre_proyecto,
            startCol = 2, startRow = 2)
  addStyle(wb = wb, sheet = hojas[2], style = titleStyle, rows = 1, cols = 2,
           gridExpand = TRUE, stack = TRUE)
  addStyle(wb = wb, sheet = hojas[2], style = subtitleStyle, rows = 2, cols = 2,
           gridExpand = TRUE, stack = TRUE)


  # Pegar logo UNAM renglón 1, columna 1

  insertImage(wb = wb, sheet = hojas[1],
              file = logo_path, startRow = 1, startCol = 1,
              width = 2.08, height = 2.2, units = 'cm')

  insertImage(wb = wb, sheet = hojas[2],
              file = logo_path, startRow = 1, startCol = 1,
              width = 2.08, height = 2.2, units = 'cm')

  # Título pregunta

  np <- lista_preguntas[[num_pregunta]]

  writeData(wb = wb, sheet =  hojas[1], x = np, startCol = columna,
            startRow = renglon[1], borders = 'none')

  writeData(wb = wb, sheet =  hojas[2], x = np, startCol = columna,
            startRow = renglon[2], borders = 'none')

  setRowHeights(wb = wb, sheet = hojas[1], rows = (renglon[1]), heights = 35)

  setRowHeights(wb = wb, sheet = hojas[2], rows = (renglon[2] + 1), heights = 35)


  # Tabla título

  tabla_titulo <- paste0('Tabla ', num_pregunta, '*')

  writeData(wb = wb, sheet = hojas[1], x = tabla_titulo, startRow = (renglon[1] +1),
            startCol = columna, colNames = FALSE, borders = 'none')

  setRowHeights(wb = wb, sheet = hojas[1], rows = (renglon[1] + 1), heights = 30)


  writeData(wb = wb, sheet = hojas[2], x = tabla_titulo, startRow = (renglon[2] +1),
            startCol = columna, colNames = FALSE, borders = 'none')

  setRowHeights(wb = wb, sheet = hojas[2], rows = (renglon[1] + 1), heights = 30)


  # Tabla Cruzada

  f <- tabla_cruzada_total(diseño = diseño, pregunta = pregunta, datos = datos,
                           DB_Mult = DB_Mult, dominios = dominios,
                           tipo_pregunta = tipo_pregunta)

  # Merge

  addStyle(wb = wb, sheet = hojas[1], style = table_Style,
           rows = (renglon[1] + 1), cols = 1:ncol(f[[1]]),
           gridExpand = TRUE, stack = TRUE)

  mergeCells(wb = wb, sheet = hojas[1], cols = 1:ncol(f[[1]]),
             rows = (renglon[1] + 1))



  addStyle(wb = wb, sheet = hojas[2], style = table_Style,
           rows = (renglon[2] + 1), cols = 1:nrow(f[[2]]),
           gridExpand = TRUE, stack = TRUE)

  mergeCells(wb = wb, sheet = hojas[2], cols = 1:ncol(f[[2]]),
             rows = (renglon[2] + 1))



  if(tipo_pregunta == 'categorica' || tipo_pregunta == 'multiple'){

    # Formato Categorías empieza en el renglón 2

    formato_categorias(tabla = f, pregunta = pregunta, datos = datos,
                       DB_Mult = DB_Mult, diseño = diseño,
                       wb = wb, renglon = c((renglon[1] + 1+1),(renglon[2] + 1+1)),
                       columna = (columna + 3), hojas = hojas,
                       estilo_cuerpo = estilo_categorias,
                       tipo_pregunta = tipo_pregunta)

    setRowHeights(wb = wb, sheet = hojas[1], rows = (renglon[1] + 1+1), heights = 45)

    setRowHeights(wb = wb, sheet = hojas[2], rows = (renglon[2] + 1+1), heights = 45)

    # Escribir tabla Cruzada empieza en el renglón 3

    formato_tablas_cruzadas(tabla = f, wb = wb,
                            renglon = c((renglon[1] + 2+1), (renglon[2] + 2+1)),
                            columna = columna, hojas = hojas,
                            estilo_encabezado = estilo_encabezado,
                            estilo_total  = estilo_total)

    # Estilo dominios empieza en el renglón 3

    estilo_dominios(tabla = f, wb = wb, columna = columna, dominios = dominios,
                    hojas = hojas, renglon = c((renglon[1] + 2+1), (renglon[2] + 2+1)),
                    estilo_dominios = estilo_horizontal, tipo_pregunta = tipo_pregunta,
                    estilo_merge_dominios = estilo_categorias)
    # Estilo columnas

    estilo_columnas(tabla = f, hojas = hojas, wb = wb, estilo = estilo_columnas,
                    renglon = c(renglon[1]+1, renglon[2]+1), tipo_pregunta = tipo_pregunta)

    addStyle(wb = wb, sheet = hojas[2], style = estilo_total,
             rows = (renglon[2]+1):(renglon[2]+1 + nrow(f[[2]])), cols = 3,
             gridExpand = TRUE, stack = TRUE)

    addStyle(wb = wb, sheet = hojas[1], style = estilo_total,
             rows = (renglon[1]+1):(renglon[1]+1 + nrow(f[[1]])), cols = 3,
             gridExpand = TRUE, stack = TRUE)

    # Fuente

    fontStyle <- createStyle(fontSize = 9, fontColour = '#5b5b5b')

    fuente_preg <- paste0("Fuente: ", fuente)
    organismo <- paste0("Organismo de participacion ", organismo_participacion)
    tabla_preg <- paste0("Tabla correspondiente a la pregunta ", num_pregunta,
                         " de ", fuente)

    writeData(wb = wb, sheet = hojas[1], x = fuente_preg,
              startRow = (renglon[1] + 3+1 + nrow(f[[1]])), startCol = columna,
              borders = 'none')

    addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
             rows = (renglon[1] + 3 + 1 + nrow(f[[1]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)

    setRowHeights(wb = wb, sheet = hojas[1],
                  rows = (renglon[1] + 3 + 1 + nrow(f[[1]])),
                  heights = 10)


    writeData(wb = wb, sheet = hojas[1], x = organismo,
              startRow = (renglon[1] + 4+1 + nrow(f[[1]])), startCol = columna,
              borders = 'none')

    addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
             rows = (renglon[1] + 4 + 1 + nrow(f[[1]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)

    setRowHeights(wb = wb, sheet = hojas[1],
                  rows = (renglon[1] + 4 + 1 + nrow(f[[1]])),
                  heights = 10)


    writeData(wb = wb, sheet = hojas[1], x = tabla_preg,
              startRow = (renglon[1] + 5 +1+ nrow(f[[1]])), startCol = columna,
              borders = 'none')

    addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
             rows = (renglon[1] + 5 + 1 + nrow(f[[1]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)

    setRowHeights(wb = wb, sheet = hojas[1],
                  rows = (renglon[1] + 5 + 1 + nrow(f[[1]])),
                  heights = 10)


    writeData(wb = wb, sheet = hojas[2], x = fuente_preg,
              startRow = (renglon[2] + 3+1 + nrow(f[[2]])), startCol = columna,
              borders = 'none')

    addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
             rows = (renglon[2] + 3 + 1 + nrow(f[[2]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)

    setRowHeights(wb = wb, sheet = hojas[2],
                  rows = (renglon[2] + 3 + 1 + nrow(f[[2]])),
                  heights = 10)


    writeData(wb = wb, sheet = hojas[2], x = organismo,
              startRow = (renglon[2] + 4+1 + nrow(f[[2]])), startCol = columna,
              borders = 'none')

    addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
             rows = (renglon[2] + 4 + 1 + nrow(f[[2]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)

    setRowHeights(wb = wb, sheet = hojas[2],
                  rows = (renglon[2] + 4 + 1 + nrow(f[[2]])),
                  heights = 10)

    writeData(wb = wb, sheet = hojas[2], x = tabla_preg,
              startRow = (renglon[2] + 5+1 + nrow(f[[2]])), startCol = columna,
              borders = 'none')
    addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
             rows = (renglon[2] + 5 + 1 + nrow(f[[2]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)

    setRowHeights(wb = wb, sheet = hojas[2],
                  rows = (renglon[2] + 5 + 1 + nrow(f[[2]])),
                  heights = 10)

  }

  if(tipo_pregunta == 'continua'){

    # Escribir tabla Cruzada empieza en el renglón 3

    formato_tablas_cruzadas(tabla = f, wb = wb,
                            renglon = c((renglon[1] + 1+1), (renglon[2] + 1+1)),
                            columna = columna, hojas = hojas,
                            estilo_encabezado = estilo_encabezado,
                            estilo_total  = estilo_total)

    # Estilo dominios empieza en el renglón 3

    estilo_dominios(tabla = f, wb = wb, columna = columna, dominios = dominios,
                    hojas = hojas, renglon = c((renglon[1] +1+ 1), (renglon[2]+1 + 1)),
                    estilo_dominios = estilo_horizontal,
                    estilo_merge_dominios = estilo_categorias,
                    tipo_pregunta = tipo_pregunta)

    # Estilo columnas

    estilo_columnas(tabla = f, hojas = hojas, wb = wb, estilo = estilo_columnas,
                    renglon = c(renglon[1]+1, renglon[2]+1), tipo_pregunta = tipo_pregunta)


    # Fuente
    fontStyle <- createStyle(fontSize = 9, fontColour = '#5b5b5b')

    fuente_preg <- paste0("Fuente: ", fuente)
    organismo <- paste0("Organismo de participacion ", organismo_participacion)
    tabla_preg <- paste0("Tabla correspondiente a la pregunta ", num_pregunta,
                         " de ", fuente)

    writeData(wb = wb, sheet = hojas[1], x = fuente_preg,
              startRow = (renglon[1] + 2+1 + nrow(f[[1]])), startCol = columna,
              borders = 'none')

    addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
             rows = (renglon[1] + 2 + 1 + nrow(f[[1]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)

    setRowHeights(wb = wb, sheet = hojas[1],
                  rows = (renglon[1] + 2 + 1 + nrow(f[[1]])),
                  heights = 10)


    writeData(wb = wb, sheet = hojas[1], x = organismo,
              startRow = (renglon[1] + 3+1 + nrow(f[[1]])), startCol = columna,
              borders = 'none')

    addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
             rows = (renglon[1] + 3 + 1 + nrow(f[[1]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)

    setRowHeights(wb = wb, sheet = hojas[1],
                  rows = (renglon[1] + 3 + 1 + nrow(f[[1]])),
                  heights = 10)


    writeData(wb = wb, sheet = hojas[1], x = tabla_preg,
              startRow = (renglon[1] + 4+1 + nrow(f[[1]])), startCol = columna,
              borders = 'none')

    addStyle(wb = wb, sheet = hojas[1], style = fontStyle,
             rows = (renglon[1] + 4 + 1 + nrow(f[[1]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)

    setRowHeights(wb = wb, sheet = hojas[1],
                  rows = (renglon[1] + 4 + 1 + nrow(f[[1]])),
                  heights = 10)



    writeData(wb = wb, sheet = hojas[2], x = fuente_preg,
              startRow = (renglon[2] + 2+1 + nrow(f[[2]])), startCol = columna,
              borders = 'none')

    addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
             rows = (renglon[2] + 2 + 1 + nrow(f[[2]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)

    setRowHeights(wb = wb, sheet = hojas[2],
                  rows = (renglon[2] + 2 + 1 + nrow(f[[2]])),
                  heights = 10)

    writeData(wb = wb, sheet = hojas[2], x = organismo,
              startRow = (renglon[2] + 3+1 + nrow(f[[2]])), startCol = columna,
              borders = 'none')

    addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
             rows = (renglon[2] + 3 + 1 + nrow(f[[2]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)

    setRowHeights(wb = wb, sheet = hojas[2],
                  rows = (renglon[2] + 3 + 1 + nrow(f[[2]])),
                  heights = 10)


    writeData(wb = wb, sheet = hojas[2], x = tabla_preg,
              startRow = (renglon[2] + 4+1 + nrow(f[[2]])), startCol = columna,
              borders = 'none')

    addStyle(wb = wb, sheet = hojas[2], style = fontStyle,
             rows = (renglon[2] + 4 + 1 + nrow(f[[2]])), cols = columna,
             gridExpand = TRUE, stack = TRUE)

    setRowHeights(wb = wb, sheet = hojas[2],
                  rows = (renglon[2] + 4 + 1 + nrow(f[[2]])),
                  heights = 10)

    # Estilo total

    addStyle(wb = wb, sheet = hojas[2], style = estilo_total,
             rows = (renglon[2] +1):(renglon[2]+1 + nrow(f[[2]])), cols = 3,
             gridExpand = TRUE, stack = TRUE)

    addStyle(wb = wb, sheet = hojas[1], style = estilo_total,
             rows = (renglon[1]+1):(renglon[1]+1 + nrow(f[[1]])), cols = 3,
             gridExpand = TRUE, stack = TRUE)

  }

  return(f)

}
