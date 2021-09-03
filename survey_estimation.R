## PRUEBA PREGUNTAS NUMÉRICAS Y CATEGÓRICAS
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
}

# Inicializar workoob
{
  wb <- openxlsx::createWorkbook()
  options("openxlsx.numFmt" = "0.0")
  
  k1 <- 1
  k2 <- 1
  np <- 1
  
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
  
  bodyStyle <- createStyle(halign = 'center', border = "TopBottomLeftRight",
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

for (p in Lista) {
  
  print(paste0("Estimando resultado de pregunta: ",p))
  dataset$Pondi1 <- Ponderador
  
  if (p %in% Multiples){
    print('Iván')
    
    np=np + 1
  }
  else if (p %in% Lista_Cont){
    print('Cin')

    np=np + 1
  }
  else({
    
    ## Dinámico: pregunta, numero de pregunta, renglon
    
    categoricas <- preguntas_categoricas(pregunta = p, numero_pregunta = np, 
                                         datos = dataset, lista_preguntas = Lista_Preg,
                                         dominios = Dominios, diseño = disenio_cat,
                                         wb = wb, renglon_fs = k1,
                                         renglon_tc = k2, columna = 1,
                                         hojas_simples = c(1,2), hojas_cruzadas = c(3,4), 
                                         frecuencias_simples = TRUE,
                                         tablas_cruzadas = TRUE)
    
    k1=k1 + 1 + nrow(categoricas[[1]][[1]]) + 4
    k2=k2 + 1 + nrow(categoricas[[2]][[2]]) + 4
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

openXL(wb)

openxlsx::saveWorkbook(wb, "Prueba.xlsx", overwrite = TRUE)


total_nacional <- function(diseño, pregunta, na.rm = TRUE,
                           estadisticas = c("se","ci","cv", "var"), significancia = 0.95, proporcion = FALSE, 
                           metodo_prop = "likelihood", DEFF = TRUE){
  
  estadisticas <- {{diseño}} %>% 
    srvyr::group_by(!!sym(pregunta)) %>% 
    srvyr::summarize(
      prop = survey_mean(
        na.rm = na.rm, 
        vartype = estadisticas,
        level = significancia,
        proportion = proporcion,
        prop_method = metodo_prop,
        deff = DEFF
      ),
      total = round(survey_total(
        na.rm = na.rm
      ),0)
    ) %>% 
    mutate(prop_low = ifelse(prop_low < 0, 0, prop_low),
           prop_upp = ifelse(prop_upp > 1, 1, prop_upp),
           !!sym(pregunta) := str_trim(!!sym(pregunta), side = 'both')) 
  
  estadisticas_wide <- estadisticas %>%
    mutate(Dominio = 'General',
           Categorias = 'Total') %>% 
    pivot_wider(
      names_from = !!sym(pregunta),
      values_from = c(
        total,
        total_se,
        prop, 
        prop_low,
        prop_upp,
        prop_se,
        prop_cv,
        prop_var,
        prop_deff),
      names_glue = sprintf("{%s}_{.value}", {{pregunta}}))
  
  return(estadisticas_wide)
}

nacional <- total_nacional(diseño = disenio_cat, pregunta = 'P3_1a_1')

nac <- formatear_tabla_cruzada(pregunta = 'P3_1a_1', datos = dataset, tabla = nacional)

rbind(nacional)
