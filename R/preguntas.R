#' Función preguntas
#' @description Función envolvente de las funciones frecuencias_simples_excel() y tablas_cruzadas_excel
#' @usage preguntas(
#' pregunta,
#' num_pregunta,
#' datos,
#' DB_Mult,
#' dominios,
#' lista_preguntas,
#' diseño,
#' wb,
#' renglon_fs,
#' renglon_tc,
#' columna = 1,
#' hojas_fs,
#' hojas_tc,
#' fuente,
#' organismo_participacion,
#' logo_path,
#' tipo_pregunta,
#' estilo_encabezado = headerStyle,
#' estilo_categorias = bodyStyle,
#' estilo_horizontal = horizontalStyle,
#' estilo_total = totalStyle,
#' frecuencias_simples = TRUE,
#' tablas_cruzadas = TRUE
#' )
#' @param pregunta  Nombre de la pregunta sobre la cual se desea obtener las frecuencias simples y/o tabla cruzada
#' @param num_pregunta Número de pregunta
#' @param datos Conjunto de datos en formato .sav
#' @param DB_Mult Data frame con las preguntas múltiples
#' @param dominios Vector de dominios sobre los cuales se desea obtener sus respectivas tablas cruzadas
#' @param lista_preguntas Data frame que contiene los títulos de las pregunta
#' @param diseño Diseño muestral que se ocupará según el tipo de pregunta
#' @param wb Workbook de Excel que contiene al menos dos hojas
#' @param renglon_fs Vector tamaño 2 especificando el número de renglon en el cual se desea empezar a escribir las tablas de frecuencias simples formateadas
#' @param renglon_tc Vector tamaño 2 especificando el número de tenglón el cual se desea empezar a escribir la tabla cruzada formateada
#' @param columna Columna en la cual se desea empezar a escribir las tablas
#' @param hojas_fs Vector de número de hojas en el cual se desea insertar las tablas de frecuencias simples
#' @param hojas_tc Vector de número de hojas en el cual se desea insertar la tabla cruzada
#' @param fuente Nombre del proyecto
#' @param organismo_participacion Organismos que participaron en el proyecto, por ejemplo, 'Ciudadanía Mexicana'
#' @param logo_path Path del logo de la UNAM
#' @param tipo_pregunta Tipo de pregunta_ 'categorica', 'multiple', 'continua'
#' @param estilo_encabezado Estilo el cual se desea usar para los nombres de las columnas
#' @param estilo_categorias Estilo el cual se desea usar para formatear las categorías de las tablas cruzadas para preguntas categóricas y múltiples
#' @param estilo_horizontal Estilo último renglones horizontales (último renglón para frecuencias simples)
#' @param estilo_total Estilo el cual se desea usar para la columna total
#' @param frecuencias_simples Valor lógico que indica si se desean realizar las frecuencias simples de la pregunta indicada
#' @param tablas_cruzadas Valor lógico que indica si se desean realizar las tablas cruzadas de la pregunta indicada
#' @details El estilo_total se recomienda crear un estilo con la función createStyle de openxlsx con el formato que se desea, por ejemplo "###,###,###.0"
#' @details El estilo_horizontal hace referencia al tipo de lineado horizontal se desea en el úntimo renglón de la tabla
#' @details El estilo_categoris hace referencia al tipo de bordes, fuente y alineación que se desea aplicar en las celdas de Excel donde se encuentra el vector de categorías
#' @details El estilo_encabezado hace referencia al tipo de formato que se desea conseguir para el nombre de las columnas de las tablas
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{openxlsx}}
#' @examples \dontrun{
#' # Creación del workbook
#'   organismo <- 'Ciudadanía mexicana'
#'   nombre_proyecto <- 'Conacyt 2018'

#' openxlsx::addWorksheet(wb, sheetName = 'Frecuencias simples')
#' showGridLines(wb, sheet = 'Frecuencias simples', showGridLines = FALSE)

#' openxlsx::addWorksheet(wb, sheetName = 'Tablas cruzadas')
#' showGridLines(wb, sheet='Tablas cruzadas', showGridLines = FALSE)

#' openxlsx::addWorksheet(wb, sheetName = 'Frecuencias (dispersión)')
#' showGridLines(wb, sheet = 'Frecuencias (dispersión)', showGridLines = FALSE)

#' openxlsx::addWorksheet(wb, sheetName = 'Tablas cruzadas (dispersión)')
#' showGridLines(wb, sheet = 'Tablas cruzadas (dispersión)', showGridLines = FALSE)

#'
#' # Estilos
#'   headerStyle <- createStyle(fontSize = 11, fontColour = "black", halign = "center", border = "TopBottom", borderColour = "black", borderStyle = c('thin', 'double'), textDecoration = 'bold')
#'   bodyStyle <- createStyle(halign = 'center', border = "TopBottomLeftRight", borderColour = "black", borderStyle = 'thin', valign = 'center', wrapText = TRUE)

#'   verticalStyle <- createStyle(border = "Right", borderColour = "black", borderStyle = 'thin', valign = 'center')

#'  totalStyle <-  createStyle(numFmt = "###,###,###.0")

#'  horizontalStyle <- createStyle(border = "bottom", borderColour = "black", borderStyle = 'thin', valign = 'center')
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
#'   preguntas(pregunta = 'P1', num_pregunta = 1, datos=dataset,
#'   DB_Mult = DB_Mult, dominios = Dominios,
#'   lista_preguntas=Lista_Preg,
#'   diseño = disenio_mult, wb = wb, renglon_fs = c(1,1),
#'   renglon_tc = c(1, 1), columna = 1, hojas_fs = c(1,3),
#'   hojas_tc = c(2,4), fuente = nombre_proyecto,
#'   tipo_pregunta = 'multiple',
#'   organismo_participacion = organismo,
#'   estilo_encabezado = headerStyle,
#'   estilo_categorias = bodyStyle,
#'   estilo_horizontal = horizontalStyle,
#'   estilo_total = totalStyle,
#'   frecuencias_simples = TRUE, tablas_cruzadas = TRUE)
#'
#'   openxlsx::openXL(wb)
#' }
#' @export
#'
preguntas <- function(pregunta, num_pregunta, datos, DB_Mult, dominios,
                      lista_preguntas, diseño, wb, renglon_fs, renglon_tc,
                      columna = 1, hojas_fs = c(1,2),
                      hojas_tc = c(3,4),
                      fuente = 'fuente',
                      organismo_participacion = 'organismo',
                      logo_path = 'imagen.png',
                      tipo_pregunta,
                      estilo_encabezado = headerStyle,
                      estilo_categorias = bodyStyle,
                      estilo_horizontal = horizontalStyle,
                      estilo_total = totalStyle,
                      frecuencias_simples = TRUE, tablas_cruzadas = TRUE){


  if(frecuencias_simples){

    fs_rows <- frecuencias_simples_excel(pregunta = pregunta,
                                         num_pregunta = num_pregunta,
                                         datos = datos, DB_Mult = DB_Mult,
                                         lista_preguntas = lista_preguntas,
                                         diseño = diseño, wb = wb,
                                         renglon = renglon_fs,
                                         columna = columna, hojas = hojas_fs,
                                         tipo_pregunta = tipo_pregunta, fuente = fuente,
                                         organismo_participacion = organismo_participacion,
                                         logo_path = logo_path,
                                         estilo_encabezado = estilo_encabezado,
                                         estilo_horizontal = estilo_horizontal,
                                         estilo_total = estilo_total)

  } else{
    fs_rows <- tibble()
  }

  if(tablas_cruzadas){

    tc_rows <- tablas_cruzadas_excel(pregunta = pregunta,
                                     num_pregunta = num_pregunta,
                                     dominios = dominios, datos = datos,
                                     DB_Mult = DB_Mult,
                                     lista_preguntas = lista_preguntas,
                                     diseño = diseño, wb = wb,
                                     renglon = renglon_tc, columna = columna,
                                     hojas = hojas_tc,
                                     tipo_pregunta = tipo_pregunta,
                                     fuente = fuente,
                                     organismo_participacion = organismo_participacion,
                                     logo_path = logo_path,
                                     estilo_encabezado = estilo_encabezado,
                                     estilo_categorias = estilo_categorias,
                                     estilo_horizontal = estilo_horizontal,
                                     estilo_total = estilo_total)

  }else{
    tc_rows <- tibble()
  }

  return(list(fs_rows, tc_rows))
}
