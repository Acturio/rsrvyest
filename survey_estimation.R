# PRUEBA PREGUNTAS NUMÉRICAS Y CATEGÓRICAS
source("~/Desktop/UNAM/DIAO/rsrvyest/src/utils.R")
 
 # Lectura de datos
{ 
  base = "BASE_CONACYT_260118.sav" #base.sav
  lista = "Lista de Preguntas.xlsx" #archivo lista de variables
  listaD <- leer_datos("BASE_CONACYT_260118.sav","Lista de Preguntas.xlsx")
 }

 {
   dataset <- read.spss("data/BASE_CONACYT_260118.sav", to.data.frame = TRUE) # Lectura de datos de spss
   General <- "Nacional" # Nombre de estimación global (Puede ser nacional, cdmx, etc. Depende de la representatividad del estudio)
   
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
  disenio_cat <- disenio_categorico(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1,
                               reps=FALSE, datos = dataset)
  mdesign <- crea_disenio(dataset, CV_ESC, ESTRATO, Pondi1)
  
  disenio_mult <- disenio_multiples(id = c(CV_ESC, ID_DIAO), estrato = ESTRATO, pesos = Pondi1,
                               reps=FALSE, datos = dataset)
}

# Inicializar workbook
{
  wb <- openxlsx::createWorkbook()
  options("openxlsx.numFmt" = "0.0")
  
  k1 <- 1
  k2 <- 1
  np <- 133
  
  openxlsx::addWorksheet(wb, sheetName = 'General 1')
  openxlsx::addWorksheet(wb, sheetName = 'General 2')
  openxlsx::addWorksheet(wb, sheetName = 'Dominios 1')
  openxlsx::addWorksheet(wb, sheetName = 'Dominios 2')
  
}

# Estilos

{
  headerStyle <- createStyle(
    fontSize = 11, fontColour = "black", halign = "center",
    border = "TopBottom", borderColour = "black",
    borderStyle = c('thin', 'double'), textDecoration = 'bold')
  
  bodyStyle <- createStyle(halign = 'center', border = "TopBottom",
                           borderColour = "black", borderStyle = 'thin',
                           valign = 'top', wrapText = TRUE)
  
  verticalStyle <- createStyle(border = "Right",
                               borderColour = "black", borderStyle = 'thin',
                               valign = 'top')
  
  horizontalStyle <- createStyle(border = "bottom",
                                 borderColour = "black", borderStyle = 'thin')
  
  totalStyle <- createStyle(numFmt = "###,###,###")
}
 
# For

for (p in Lista[133:151]) {

  print(paste0("Estimando resultado de pregunta: ",p))
  dataset$Pondi1 <- Ponderador
  
  if (p %in% Multiples){
    
    print('múltiple')
    

    multiples <- preguntas_multiples(pregunta = p, numero_pregunta = np, datos = dataset,
                        lista_preguntas = Lista_Preg, dominios = Dominios,
                        disenio = disenio_mult, wb = wb, renglon_fs = k1, 
                        renglon_tc = k2, DB_Mult = DB_Mult, columna = 1,
                        hojas_simples = c(1,2), hojas_cruzadas = c(3,4),
                        estilo_cuerpo = bodyStyle, estilo_columnas = verticalStyle,
                        frecuencias_simples = TRUE, tablas_cruzadas = TRUE)
    
    k1=k1 + 1 + nrow(multiples[[1]][[1]]) + 5
    k2=k2 + 1 + nrow(multiples[[2]][[2]]) + 5
    np=np + 1
    
  }
  else if (p %in% Lista_Cont){
    
    print('continua')
    
    formato = 1
    
    # frecuencias simples
    
    tf <- estadisticas_continuas(mdesign, p)
    
    tf <- acomoda_frecuencias(tf)
    
    writeData(wb, 1, Lista_Preg[np], startRow = k1-1, startCol = 1)
    tabla_frec_excel(tf, 1, k1)
    
    # tablas cruzadas
    t <- ftotal(mdesign, p, na.rm = TRUE,
                vartype = c("se", "ci", "cv", "var"), 
                level = 0.95, proportion = FALSE, prop_method = "likelihood",
                DEFF = TRUE, cuantiles) 
    print(t)
    
    
    tabla <- t
    for (dom in Dominios){
      cruce <- tabla_cruzada(mdesign, p, dom, na.rm = TRUE,
                             vartype = c("se", "ci", "cv", "var"), 
                             level = 0.95, DEFF = TRUE)
      tabla <- bind_rows(tabla, cruce)
      
    }

    ltabla <- formato_tabla(tabla)
    if (formato == 1) {
      tablaf <- as.data.frame(ltabla[1])
    } else {
      tablaf <- as.data.frame(ltabla[2])
    }
    print(tablaf)
    
    #escribo el excel
    tabla_excel(tablaf, 1, k2)
    
    #escribo la pregunta  en rowini-1
    writeData(wb, 3, Lista_Preg[np], startRow = k2-1, startCol = 1)
    
    #recalculo renglones
    
    k1=k1+nrow(tf) + 5
    k2=k2 + nrow(tablaf) + 5
    np=np + 1
    
  } 
  else({
    
    print('categórica')
    ## Dinámico: pregunta, numero de pregunta, renglon
    
    categoricas <- preguntas_categoricas(pregunta = p, numero_pregunta = np, 
                                         datos = dataset, lista_preguntas = Lista_Preg,
                                         dominios = Dominios, diseño = disenio_cat,
                                         wb = wb, renglon_fs = k1,
                                         renglon_tc = k2, columna = 1,
                                         hojas_simples = c(1,2), hojas_cruzadas = c(3,4), 
                                         frecuencias_simples = TRUE,
                                         tablas_cruzadas = TRUE)
    
    k1=k1 + 1 + nrow(categoricas[[1]][[1]]) + 5
    k2=k2 + 1 + nrow(categoricas[[2]][[2]]) + 5
    np=np + 1
    
  })
  
}

# Estilo columna total 

{

  addStyle(wb = wb, sheet = 1, style = totalStyle, rows = 3:100000, cols = 2 ,
           gridExpand = TRUE, stack = TRUE)
  
  addStyle(wb = wb, sheet = 2, style = totalStyle, rows = 3:100000, cols = 2 ,
           gridExpand = TRUE, stack = TRUE)

  addStyle(wb = wb, sheet = 3, style = totalStyle, rows = 3:100000, cols = 3,
           gridExpand = TRUE, stack = TRUE)

  addStyle(wb = wb, sheet = 4, style = totalStyle, rows = 3:100000, cols = 3,
           gridExpand = TRUE, stack = TRUE)
}

openxlsx::openXL(wb)

openxlsx::saveWorkbook(wb, "Prueba.xlsx", overwrite = TRUE)

