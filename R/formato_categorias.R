# FUNCIÓN FORMATO CATEGORÍAS PARA PREGUNTAS CATEGÓRICAS Y MÚLTIPLES

#' Función formato categorías para preguntas categóricas y múltiples
#' @description El vector obtenido con la función categorias_pregunta_formato() se escribe en un workbook de Excel
#' @usage formato_categorias(
#' tabla,
#' pregunta,
#' diseño,
#' datos,
#' DB_Mult,
#' wb,
#' renglon,
#' columna = 4,
#' hojas,
#' estilo_cuerpo,
#' tipo_pregunta
#' )
#' @param tabla Tabla cruzada creada por la función total_general()
#' @param pregunta  Nombre de la pregunta sobre la cual se creo la tabla
#' @param diseño Diseño muestral que se ocupará según el tipo de pregunta
#' @param datos Conjunto de datos en formato .sav
#' @param DB_Mult Data frame con las preguntas múltiples
#' @param wb Workbook de Excel que contiene al menos dos hojas
#' @param renglon Vector tamaño 2 especificando el número de renglon en el cual se desea empezar a escribir el vector de categorías
#' @param columna Columna en la cual se desea empezar a escribir las tablas. SIEMPRE EN LA CUARTA COLUMNA
#' @param hojas Vector de número de hojas en el cual se desea insertar los vectores
#' @param estilo_cuerpo Estilo que indica cómo se desea formatear el vector en las hojas de excel
#' @param tipo_pregunta Tipo de pregunta_ 'categorica', 'multiple', 'continua'
#' @details Una vez escrito el vector colapsa celdas (3 o 4)
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{mergeCells}} \code{\link{addStyle}} \code{\link{createStyle}}
#' @examples \dontrun{
#' # Creación del workbook
#' wb <- createWorkbook()
#' addWorksheet(wb, "Tablas cruzadas")
#' addWorksheet(wb, "Tablas cruzadas (dispersión)")
#'
#' # Estilos
#'   bodyStyle <- createStyle(halign = 'center', border = "TopBottomLeftRight",
#'   borderColour = "black", borderStyle = 'thin',
#'   valign = 'center', wrapText = TRUE)
#'
#' # Carga de datos
#'  dataset <- read.spss("data/BASE_CONACYT_260118.sav", to.data.frame = TRUE)
#'  Lista_Preg <- read_xlsx("aux/Lista de Preguntas.xlsx",
#'                        sheet = "Lista Preguntas")$Nombre %>% as.vector()
#'  DB_Mult <- read_xlsx("aux/Lista de Preguntas.xlsx",  sheet = "Múltiple") %>% as.data.frame()
#'  Lista_Cont <- read_xlsx("aux/Lista de Preguntas.xlsx",
#'   sheet = "Continuas")$VARIABLE %>% as.vector()
#'  Dominios <- read_xlsx("aux/Lista de Preguntas.xlsx", sheet = "Dominios")$Dominios %>% as.vector()
#'
#'  disenio_mult <- disenio(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1, reps=FALSE, datos = dataset)
#'
#'  total <- total_general (diseño = disenio_mult,  pregunta = 'P1', dominio = 'General', datos = dataset,
#'  DB_Mult = DB_Mult, tipo_pregunta = 'multiple')
#'
#'  formato_categorias(tabla = total, pregunta = 'P1', diseño = disenio_mult, datos = dataset,
#'  DB_Mult = DB_Mult, wb = wb, renglon = c(1,1), columna = 4, hojas = c(1,2), estilo_cuerpo = bodyStyle, tipo_pregunta = 'multiple')
#'
#'  openxlsx::openXL(wb)
#'  }
#'  @import openxlsx
#' @export
formato_categorias <- function(tabla, pregunta, diseño, datos, DB_Mult, wb,
                               renglon = c(1,1), columna = 4, hojas = c(3,4),
                               estilo_cuerpo, tipo_pregunta = 'categorica'){

  # Renglón 2 columna 4

  simples <- categorias_pregunta_formato(pregunta = pregunta,
                                         diseño = diseño, datos = datos,
                                         metricas = FALSE, DB_Mult = DB_Mult,
                                         tipo_pregunta = tipo_pregunta)

  writeData(wb = wb, sheet =  hojas[1], x = simples, startRow = renglon[1],
            startCol = columna, borders = 'surrounding',
            borderStyle = 'thin',  keepNA = FALSE, colNames = FALSE)


  for (k in seq(columna, ncol(tabla[[1]]), by=3)){

    mergeCells(wb = wb, sheet = hojas[1], cols = k:(k+2), rows = renglon[1])
    addStyle(wb = wb, sheet = hojas[1], style = estilo_cuerpo, rows = renglon[1],
             cols = k:(k+2), stack = TRUE, gridExpand = TRUE)
  }


  metricas <- categorias_pregunta_formato(pregunta = pregunta,
                                          diseño = diseño, datos = datos,
                                          metricas = TRUE, DB_Mult = DB_Mult,
                                          tipo_pregunta = tipo_pregunta)

  writeData(wb = wb, sheet =  hojas[2], x = metricas, startRow = renglon[2],
            startCol = columna, borders = 'surrounding',
            borderStyle = 'thin',  keepNA = FALSE, colNames = FALSE)

  for (k in seq(columna, ncol(tabla[[2]]), by=4)) {
    mergeCells(wb = wb, sheet = hojas[2], cols = k:(k+3), rows = renglon[2])
    addStyle(wb = wb, sheet = hojas[2], style = estilo_cuerpo, rows = renglon[2],
             cols = k:(k+3), stack = TRUE, gridExpand = TRUE)
  }

}
