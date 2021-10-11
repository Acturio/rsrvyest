# FUNCIÓN DISEÑO DE MUESTREO

#' Crea un objeto con un diseño de muestreo
#' @description Se crea un objeto con el diseño de muestreo especificado.
#' @usage disenio(
#' id,
#' estrato,
#' pesos,
#' datos,
#' pps = "brewer",
#' varianza = "HT",
#' reps = TRUE,
#' metodo = "subbootstrap",
#' B = 500,
#' semilla = 1234
#' )
#' @param id Variable que especifica el indicador para cada uno de los estratos
#' @param estrato Variable que especifica los estratos.
#' @param pesos Variables que especifica las ponderaciones.
#' @param datos Conjunto de datos en formato .sav
#' @param pps Valor que indica el tipo de aproximación que se desea utilizar para el muestreo.
#' @param varianza Valor que indica el estimador que se desea utilizar ("HT" para el estimador Horvitz-Thompson, "YG" para el estimador Yates-Grundy)
#' @param reps Valor lógico que indica si se desea realizar un muestreo con pesos replicados.
#' @param metodo Si reps = TRUE; tipo de replicación de pesos que se desea usar.
#' @param B Si reps = TRUE; número de replicaciones que desean realizar.
#' @param semilla Número utilizado para inicializar la generación números aleatorios.
#' @return Objeto del tipo tbl_svy
#' @author Bringas Arturo, Rosales Cinthia, Salgado Iván, Torres Ana
#' @seealso \code{\link{as_survey_design, as_survey_rep}}
#' @examples \dontrun{
#' disenio(id = id_estrato, estrato = estrato, pesos = ponderador, datos = dataset,
#'   pps = "brewer", varianza = "HT", reps = TRUE, metodo = "subbootstrap",
#'   B = 50, semilla = 1234)
#' }
#' @export
disenio <- function(id, estrato, pesos, datos, pps = "brewer",
                    varianza = "HT", reps = TRUE, metodo = "subbootstrap",
                    B=50, semilla=1234){

  disenio <- datos %>%
    as_survey_design(
      ids = {{id}},
      weights = {{pesos}},
      strata = {{estrato}},
      pps = pps,
      variance = varianza
    )

  if (reps){
    disenio %<>% as_survey_rep(
      type = metodo,
      replicates = B
    )
  }
  return(disenio)
}

