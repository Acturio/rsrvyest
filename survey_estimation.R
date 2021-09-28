# PRUEBA PREGUNTAS NUMÉRICAS Y CATEGÓRICAS
source("~/Desktop/UNAM/DIAO/rsrvyest/src/utils.R")
 
 # Lectura de datos

 {
   # Lectura de datos de spss
   dataset <- read.spss("data/BASE_CONACYT_260118.sav", to.data.frame = TRUE) 

   # Nombre de estimación global (Puede ser nacional, cdmx, etc. Depende de la representatividad del estudio)
   General <- "Nacional" 
   
   Lista <- read_xlsx("aux/Lista de Preguntas.xlsx", 
                      sheet = "Lista Preguntas")$Pregunta %>% as.vector()
   Lista_Preg <- read_xlsx("aux/Lista de Preguntas.xlsx", 
                           sheet = "Lista Preguntas")$Nombre %>% as.vector()
   DB_Mult <- read_xlsx("aux/Lista de Preguntas.xlsx", 
                        sheet = "Múltiple") %>% as.data.frame()
   Lista_Cont <- read_xlsx("aux/Lista de Preguntas.xlsx", 
                           sheet = "Continuas")$VARIABLE %>% as.vector()
   Dominios <- read_xlsx("aux/Lista de Preguntas.xlsx", 
                         sheet = "Dominios")$Dominios %>% as.vector()
   
   Multiples <- names(DB_Mult)
   Ponderador <- dataset$Pondi1
   save = ""
 }
 
# Diseños
 
{
  disenio_cat <- disenio(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1,
                               reps=FALSE, datos = dataset)
  disenio_cont <- disenio( id = CV_ESC, estrato = ESTRATO, pesos = Pondi1,
                      reps = FALSE, datos = dataset)
  
  disenio_mult <- disenio(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1,
                               reps=FALSE, datos = dataset)
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
  np <- 133

  organismo <- 'Ciudadanía mexicana'
  nombre_proyecto <- 'Conacyt 2018'
  
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
    fontSize = 11, 
    fontColour = "black", 
    halign = "center",
    border = "TopBottom", 
    borderColour = "black",
    borderStyle = c('thin', 'double'), 
    textDecoration = 'bold'
    )
  
  bodyStyle <- createStyle(
    halign = 'center', 
    border = "TopBottomLeftRight",
    borderColour = "black", 
    borderStyle = 'thin',
    valign = 'center', 
    wrapText = TRUE
    )
  
  verticalStyle <- createStyle(
    border = "Right",
    borderColour = "black", 
    borderStyle = 'thin',
    valign = 'center'
    )
  
  totalStyle <-  createStyle(numFmt = "###,###,###.0")
  
  horizontalStyle <- createStyle(
    border = "bottom",
    borderColour = "black", 
    borderStyle = 'thin',
    valign = 'center'
    )
}


# For

for (p in Lista[133:151]) {

  print(paste0("Estimando resultado de pregunta: ",p))
  dataset$Pondi1 <- Ponderador
  
  if (p %in% Multiples){

    print('múltiple')
    
    multiples <- preguntas(
      pregunta = p, 
      num_pregunta = np, 
      datos=dataset,
      DB_Mult = DB_Mult, 
      dominios = Dominios,
      lista_preguntas = Lista_Preg,
      diseño = disenio_mult, 
      wb = wb, 
      renglon_fs = c(k1, k2),
      renglon_tc = c(k3, k4), 
      columna = 1, 
      hojas_fs = c(1,3),
      hojas_tc = c(2,4), 
      fuente = nombre_proyecto,
      tipo_pregunta = 'multiple',
      organismo_participacion = organismo,
      estilo_encabezado = headerStyle,
      estilo_categorias = bodyStyle,
      estilo_horizontal = horizontalStyle,
      estilo_total = totalStyle,
      frecuencias_simples = TRUE, 
      tablas_cruzadas = TRUE
      )
    
    # Frecuencias simples
    k1=k1 + 1 + nrow(multiples[[1]][[1]]) + 7
    k2=k2 + 1 + nrow(multiples[[1]][[2]]) + 7
    
    # Tablas cruzadas
    k3=k3 + 1 + nrow(multiples[[2]][[1]]) + 8
    k4=k4 + 1 + nrow(multiples[[2]][[2]]) + 8
    
    np=np + 1
    
  }
  else if (p %in% Lista_Cont){

    print('continua')
    
   continuas <-  preguntas(
     pregunta = p, 
     num_pregunta = np, 
     datos=dataset, 
     DB_Mult = DB_Mult, 
     dominios = Dominios, 
     lista_preguntas=Lista_Preg, 
     diseño = disenio_cont,
     wb = wb, 
     renglon_fs = c(k1, k2), 
     renglon_tc = c(k3, k4),
     columna = 1, 
     hojas_fs = c(1,3),
     hojas_tc = c(2,4), 
     fuente = nombre_proyecto,
     tipo_pregunta = 'continua',
     organismo_participacion = organismo,
     estilo_encabezado = headerStyle,
     estilo_categorias = bodyStyle,
     estilo_horizontal = horizontalStyle,
     estilo_total = totalStyle,
     frecuencias_simples = TRUE, 
     tablas_cruzadas = TRUE
     )
   
   # Frecuencias simples
   k1=k1 + 1 + nrow(continuas[[1]][[1]]) + 7
   k2=k2 + 1 + nrow(continuas[[1]][[2]]) + 7
   
   # Tablas cruzadas
   k3=k3 + 1 + nrow(continuas[[2]][[1]]) + 7
   k4=k4 + 1 + nrow(continuas[[2]][[2]]) + 7
   
   np=np + 1

  }
  else({
    
    print('categórica')
  
    categoricas <- preguntas(
      pregunta = p, 
      num_pregunta = np, 
      datos = dataset,
      DB_Mult = DB_Mult, 
      dominios = Dominios,
      lista_preguntas = Lista_Preg, 
      diseño = disenio_cat, 
      wb = wb, 
      renglon_fs = c(k1, k2), 
      renglon_tc = c(k3, k4),
      columna = 1, 
      hojas_fs = c(1, 3), 
      hojas_tc = c(2, 4),
      fuente = nombre_proyecto,
      tipo_pregunta = 'categorica',
      organismo_participacion = organismo,
      estilo_encabezado = headerStyle,
      estilo_categorias = bodyStyle,
      estilo_horizontal = horizontalStyle,
      estilo_total = totalStyle,
      frecuencias_simples = TRUE, 
      tablas_cruzadas = TRUE)
    
    # Frecuencias simples
    k1=k1 + 1 + nrow(categoricas[[1]][[1]]) + 7
    k2=k2 + 1 + nrow(categoricas[[1]][[2]]) + 7
    
    # Tablas cruzadas
    k3=k3 + 1 + nrow(categoricas[[2]][[1]]) + 8
    k4=k4 + 1 + nrow(categoricas[[2]][[2]]) + 8
    
    np=np + 1
  })
  
}


openxlsx::openXL(wb)

openxlsx::saveWorkbook(wb, "Prueba.xlsx", overwrite = TRUE)

