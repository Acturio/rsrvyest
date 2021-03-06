% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/frecuencias_simples.R
\name{frecuencias_simples}
\alias{frecuencias_simples}
\title{Crea tabla frecuencias simples}
\usage{
frecuencias_simples(
diseño,
datos,
pregunta,
DB_Mult,
na.rm = TRUE,
estadisticas =  c("se","ci","cv", "var"),
cuantiles = c(0,0.25, 0.5, 0.75,1),
significancia = 0.95,
proporcion = FALSE,
metodo_prop = "likelihood", DEFF = TRUE,
tipo_pregunta = "categorica"
)
}
\arguments{
\item{diseño}{Diseño muestral que se ocupará según el tipo de pregunta}

\item{datos}{Conjunto de datos en formato .sav}

\item{pregunta}{Pregunta de la cual se quieren obtener las frecuencias simples, por ejemplo, 'P_1'}

\item{DB_Mult}{Data frame con las preguntas múltiples}

\item{na.rm}{Valor lógico que indica si se deben de omitir valores faltantes}

\item{estadisticas}{Métricas de variabilidad: error estándar ("se"), intervalo de confianza ("ci"), varianza ("var") o coeficiente de variación ("cv")}

\item{cuantiles}{Vector de cuantiles a calcular}

\item{significancia}{Nivel de confianza: 0.95 por default}

\item{proporcion}{Valor lógico que indica si se desen usar métodos para calcular la proporción que puede tener intervalos de confianza más precisos cerca de 0 y 1}

\item{metodo_prop}{Si proporcion = TRUE; tipo de método de proporción que se desea usar: "logit", "likelihood", "asin", "beta", "mean"}

\item{DEFF}{Valor lógico que indica si se desea calcular el efecto de diseño}

\item{tipo_pregunta}{Tipo de pregunta: 'categorica', 'multiple', 'continua'}
}
\value{
Tabla tipo tibble con las estadísticas especificadas en el parámetro estadisticas por respuestas pertenecientes a la pregunta especificada en el parámetro pregunta
}
\description{
Se crean las frecuencias simples según tipo de pregunta (categórica, múltiple o continua)
}
\examples{
\dontrun{
frecuencias_simples(diseño = disenio_cat, datos = dataset, pregunta = 'P1',
 DB_Mult = DB_Mult, tipo_pregunta = 'categorica')
}
}
\seealso{
\code{\link{survey_mean}}
}
\author{
Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
}
