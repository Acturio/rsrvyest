% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/preguntas.R
\name{preguntas}
\alias{preguntas}
\title{Función preguntas}
\usage{
preguntas(
pregunta,
num_pregunta,
datos,
DB_Mult,
dominios,
lista_preguntas,
diseño,
wb,
renglon_fs,
renglon_tc,
columna = 1,
hojas_fs,
hojas_tc,
fuente,
pie_tabla,
organismo_participacion,
logo,
tipo_pregunta,
estilo_encabezado = headerStyle,
estilo_categorias = bodyStyle,
estilo_horizontal = horizontalStyle,
estilo_total = totalStyle,
frecuencias_simples = TRUE,
tablas_cruzadas = TRUE
)
}
\arguments{
\item{pregunta}{Nombre de la pregunta sobre la cual se desea obtener las frecuencias simples y/o tabla cruzada}

\item{num_pregunta}{Número de pregunta}

\item{datos}{Conjunto de datos en formato .sav}

\item{DB_Mult}{Data frame con las preguntas múltiples}

\item{dominios}{Vector de dominios sobre los cuales se desea obtener sus respectivas tablas cruzadas}

\item{lista_preguntas}{Data frame que contiene los títulos de las pregunta}

\item{diseño}{Diseño muestral que se ocupará según el tipo de pregunta}

\item{wb}{Workbook de Excel que contiene al menos dos hojas}

\item{renglon_fs}{Vector tamaño 2 especificando el número de renglon en el cual se desea empezar a escribir las tablas de frecuencias simples formateadas}

\item{renglon_tc}{Vector tamaño 2 especificando el número de tenglón el cual se desea empezar a escribir la tabla cruzada formateada}

\item{columna}{Columna en la cual se desea empezar a escribir las tablas}

\item{hojas_fs}{Vector de número de hojas en el cual se desea insertar las tablas de frecuencias simples}

\item{hojas_tc}{Vector de número de hojas en el cual se desea insertar la tabla cruzada}

\item{fuente}{Nombre del proyecto}

\item{pie_tabla}{Pie de tabla}

\item{organismo_participacion}{Organismos que participaron en el proyecto, por ejemplo, 'Ciudadanía Mexicana'}

\item{logo}{Path del logo de la UNAM}

\item{tipo_pregunta}{Tipo de pregunta_ 'categorica', 'multiple', 'continua'}

\item{estilo_encabezado}{Estilo el cual se desea usar para los nombres de las columnas}

\item{estilo_categorias}{Estilo el cual se desea usar para formatear las categorías de las tablas cruzadas para preguntas categóricas y múltiples}

\item{estilo_horizontal}{Estilo último renglones horizontales (último renglón para frecuencias simples)}

\item{estilo_total}{Estilo el cual se desea usar para la columna total}

\item{frecuencias_simples}{Valor lógico que indica si se desean realizar las frecuencias simples de la pregunta indicada}

\item{tablas_cruzadas}{Valor lógico que indica si se desean realizar las tablas cruzadas de la pregunta indicada}
}
\description{
Función envolvente de las funciones frecuencias_simples_excel() y tablas_cruzadas_excel
}
\details{
El estilo_total se recomienda crear un estilo con la función createStyle de openxlsx con el formato que se desea, por ejemplo "###,###,###.0"

El estilo_horizontal hace referencia al tipo de lineado horizontal se desea en el úntimo renglón de la tabla

El estilo_categoris hace referencia al tipo de bordes, fuente y alineación que se desea aplicar en las celdas de Excel donde se encuentra el vector de categorías

El estilo_encabezado hace referencia al tipo de formato que se desea conseguir para el nombre de las columnas de las tablas
}
\examples{
\dontrun{
# Creación del workbook
organismo <- "Ciudadanía mexicana"
nombre_proyecto <- "Conacyt 2018"
pie <- "Conacyt 2018"
openxlsx::addWorksheet(wb, sheetName = "Frecuencias simples")
showGridLines(wb, sheet = "Frecuencias simples", showGridLines = FALSE)
openxlsx::addWorksheet(wb, sheetName = "Tablas cruzadas")
showGridLines(wb, sheet = "Tablas cruzadas", showGridLines = FALSE)
openxlsx::addWorksheet(wb, sheetName = "Frecuencias (dispersión)")
showGridLines(wb, sheet = "Frecuencias (dispersión)", showGridLines = FALSE)
openxlsx::addWorksheet(wb, sheetName = "Tablas cruzadas (dispersión)")
showGridLines(wb, sheet = "Tablas cruzadas (dispersión)", showGridLines = FALSE)

# Estilos
headerStyle <- createStyle(fontSize = 11, fontColour = "black", halign = "center", border = "TopBottom", borderColour = "black", borderStyle = c("thin", "double"), textDecoration = "bold")
bodyStyle <- createStyle(halign = "center", border = "TopBottomLeftRight", borderColour = "black", borderStyle = "thin", valign = "center", wrapText = TRUE)
verticalStyle <- createStyle(border = "Right", borderColour = "black", borderStyle = "thin", valign = "center")
totalStyle <- createStyle(numFmt = "###,###,###.0")
horizontalStyle <- createStyle(border = "bottom", borderColour = "black", borderStyle = "thin", valign = "center")

# Carga de datos
dataset <- read.spss("data/BASE_CONACYT_260118.sav", to.data.frame = TRUE)
Lista_Preg <- read_xlsx("aux/Lista de Preguntas.xlsx",
  sheet = "Lista Preguntas"
)$Nombre \%>\% as.vector()
DB_Mult <- read_xlsx("aux/Lista de Preguntas.xlsx",
  sheet = "Múltiple"
) \%>\% as.data.frame()

# Diseño
disenio_mult <- disenio(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1, reps = FALSE, datos = dataset)

preguntas(
  pregunta = "P1", num_pregunta = 1, datos = dataset,
  DB_Mult = DB_Mult, dominios = Dominios,
  lista_preguntas = Lista_Preg,
  diseño = disenio_mult, wb = wb, renglon_fs = c(1, 1),
  renglon_tc = c(1, 1), columna = 1, hojas_fs = c(1, 3),
  hojas_tc = c(2, 4), fuente = nombre_proyecto, pie_tabla = pie,
  tipo_pregunta = "multiple",
  organismo_participacion = organismo,
  estilo_encabezado = headerStyle,
  estilo_categorias = bodyStyle,
  estilo_horizontal = horizontalStyle,
  estilo_total = totalStyle,
  frecuencias_simples = TRUE, tablas_cruzadas = TRUE
)

openxlsx::openXL(wb)
}
}
\seealso{
\code{\link{openxlsx}}
}
\author{
Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
}
