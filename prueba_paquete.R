library(devtools)
devtools::install_github('Acturio/rsrvyest', force = TRUE)

{
  library(rsrvyest)
  library(foreign) #lectura de .sav
  library(dplyr)
  library(readxl) #lectura xlsx
  library(openxlsx)
  library(magrittr)
  library(tidyr)
  library(stringr)
  library(caret) #6.0-90
  library(tibble)
  library(purrr)
}

# Parámetros, datos, diseño

{
  organismo <- 'DENUE'
  nombre_proyecto <- 'BIENESTAR DE LOS HOGARES DE LA CDMX 2021'
  logo <- '~/Desktop/UNAM/DIAO/logo_unam.png'
  output_name <- 'EVALUA_frecuencias_simples.xlsx'
  fuente_tabla <- 'EVALUA 2021'
  FS <- TRUE
  TC <- FALSE

  dataset <- read.spss("data/BDD_UNAM_BIENESTAR_FINAL_22112021_L.sav", to.data.frame = TRUE) # Lectura de datos de spss

  Lista <- read_xlsx("aux/Lista de preguntas sa.xlsx",
                     sheet = "Lista de Preguntas")$Pregunta %>% as.vector()
  Lista_Preg <- read_xlsx("aux/Lista de preguntas sa.xlsx",
                          sheet = "Lista de Preguntas")$Nombre %>% as.vector()
  DB_Mult <- read_xlsx("aux/Lista de preguntas sa.xlsx",
                       sheet = "Múltiple") %>% as.data.frame()
  Lista_Cont <- read_xlsx("aux/Lista de preguntas sa.xlsx",
                          sheet = "Continuas")$Variable %>% as.vector()

  Dominios <- read_xlsx("aux/Lista de preguntas sa.xlsx",
                         sheet = "Dominios")$Dominio %>% as.vector()

  Multiples <- names(DB_Mult)

   dataset %<>% mutate(
     id_evalua = row_number(),
     Pondi1 = 1)

  # valid_column_names <- make.names(names = names(dataset), unique=TRUE, allow_ = TRUE)
  # names(dataset) <- valid_column_names

  # Diseño
  disenio_muestral <- disenio(
    datos = dataset,
    id = c(id_evalua),
    estrato = EDO,
    pesos = Pondi1,
    reps = T,
    B = 50
  )
}

# Inicializar workbook
{
  wb <- openxlsx::createWorkbook()
  options("openxlsx.numFmt" = "0.0")

  # Empezamos en el renglón 5
  k1 <- 5
  k2 <- 5
  k3 <- 5
  k4 <- 5
  np <- 1
    #215

  openxlsx::addWorksheet(wb, sheetName = 'Frecuencias simples')
  showGridLines(wb, sheet = 'Frecuencias simples', showGridLines = FALSE)

  openxlsx::addWorksheet(wb, sheetName = 'Tablas cruzadas')
  showGridLines(wb, sheet='Tablas cruzadas', showGridLines = FALSE)

  openxlsx::addWorksheet(wb, sheetName = 'Frecuencias (dispersión)')
  showGridLines(wb, sheet = 'Frecuencias (dispersión)', showGridLines = FALSE)

  openxlsx::addWorksheet(wb, sheetName = 'Tablas cruzadas (dispersión)')
  showGridLines(wb, sheet = 'Tablas cruzadas (dispersión)', showGridLines = FALSE)
}

# Estilos

{
  headerStyle <- createStyle(
    fontSize = 11, fontColour = "black", halign = "center",
    border = "TopBottom", borderColour = "black",
    borderStyle = c('thin', 'double'), textDecoration = 'bold')

  bodyStyle <- createStyle(halign = 'center', border = "TopBottomLeftRight",
                           borderColour = "black", borderStyle = 'thin',
                           valign = 'center', wrapText = TRUE)

  verticalStyle <- createStyle(border = "Right",
                               borderColour = "black", borderStyle = 'thin',
                               valign = 'center')

  totalStyle <-  createStyle(numFmt = "###,###,###.0")

  horizontalStyle <- createStyle(border = "bottom",
                                 borderColour = "black", borderStyle = 'thin',
                                 valign = 'center')
}


# For

ini <- Sys.time()

#[215:226]

for (p in Lista) {

  print(paste0("Estimando resultado de pregunta: ",p))

  if (p %in% Multiples){

    print('múltiple')

    multiples <- preguntas(pregunta = p, num_pregunta = np, datos=dataset,
                           DB_Mult = DB_Mult, dominios = Dominios,
                           lista_preguntas=Lista_Preg,
                           diseño = disenio_muestral, wb = wb, renglon_fs = c(k1, k2),
                           renglon_tc = c(k3, k4), columna = 1, hojas_fs = c(1,3),
                           hojas_tc = c(2,4), fuente = nombre_proyecto,
                           tipo_pregunta = 'multiple',
                           logo = logo, pie_tabla = fuente_tabla,
                           organismo_participacion = organismo,
                           estilo_encabezado = headerStyle,
                           estilo_categorias = bodyStyle,
                           estilo_horizontal = horizontalStyle,
                           estilo_total = totalStyle,
                           frecuencias_simples = FS, tablas_cruzadas = TC)


    # Frecuencias simples
    k1=k1 + 1 + nrow(multiples[[1]][[1]]) + 7
    k2=k2 + 1 + nrow(multiples[[1]][[2]]) + 7

    # Tablas cruzadas
    # k3=k3 + 1 + nrow(multiples[[2]][[1]]) + 8
    # k4=k4 + 1 + nrow(multiples[[2]][[2]]) + 8

    np=np + 1

  }
  else if (p %in% Lista_Cont){

    print('continua')

    continuas <-  preguntas(pregunta = p, num_pregunta = np, datos=dataset,
                            DB_Mult = DB_Mult, dominios = Dominios,
                            lista_preguntas=Lista_Preg, diseño = disenio_muestral,
                            wb = wb, renglon_fs = c(k1, k2), renglon_tc = c(k3, k4),
                            columna = 1, hojas_fs = c(1,3),
                            hojas_tc = c(2,4), fuente = nombre_proyecto,
                            tipo_pregunta = 'continua',
                            logo = logo, pie_tabla = fuente_tabla,
                            organismo_participacion = organismo,
                            estilo_encabezado = headerStyle,
                            estilo_categorias = bodyStyle,
                            estilo_horizontal = horizontalStyle,
                            estilo_total = totalStyle,
                            frecuencias_simples = FS, tablas_cruzadas = TC)

    # Frecuencias simples
    k1=k1 + 1 + nrow(continuas[[1]][[1]]) + 7
    k2=k2 + 1 + nrow(continuas[[1]][[2]]) + 7

    # Tablas cruzadas
    # k3=k3 + 1 + nrow(continuas[[2]][[1]]) + 7
    # k4=k4 + 1 + nrow(continuas[[2]][[2]]) + 7

    np=np + 1

  }
  else({

    print('categórica')

    categoricas <- preguntas(pregunta = p, num_pregunta = np, datos=dataset,
                             DB_Mult = DB_Mult, dominios = Dominios,
                             lista_preguntas=Lista_Preg, diseño = disenio_muestral,
                             wb = wb, renglon_fs = c(k1, k2), renglon_tc = c(k3, k4),
                             columna = 1, hojas_fs = c(1,3), hojas_tc = c(2,4),
                             fuente = nombre_proyecto, pie_tabla = fuente_tabla,
                             tipo_pregunta = 'categorica',
                             logo = logo,
                             organismo_participacion = organismo,
                             estilo_encabezado = headerStyle,
                             estilo_categorias = bodyStyle,
                             estilo_horizontal = horizontalStyle,
                             estilo_total = totalStyle,
                             frecuencias_simples = FS,
                             tablas_cruzadas = TC)

    # Frecuencias simples
    k1=k1 + 1 + nrow(categoricas[[1]][[1]]) + 7
    k2=k2 + 1 + nrow(categoricas[[1]][[2]]) + 7

    # Tablas cruzadas
    # k3=k3 + 1 + nrow(categoricas[[2]][[1]]) + 8
    # k4=k4 + 1 + nrow(categoricas[[2]][[2]]) + 8

    np=np + 1
  })

}

fin <- Sys.time()

fin - ini


ini <- Sys.time()

openxlsx::saveWorkbook(wb, output_name, overwrite = TRUE)

fin <- Sys.time()

fin - ini
