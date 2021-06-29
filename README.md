# Rsrvyest - R Survey Estimation

Repositorio desarrollado por el equipo de estadística y ciencia de datos del Departamento de Investigación Aplicada y Opinión (DIAO) de la UNAM.

El desarrollo de esta librería tiene por objetivo crear el reporte de las frecuencias simples, múltiples y estadísticos descriptivos ponderados para los resultados de las encuentas diseñadas por el departamento. Adicionalmente, se crean las funciones para realizar los reportes de las tablas cruzadas ponderadas, así como los estadísticos que reportan la calidad y precisión de las estimaciones de cada una de las preguntas.

## Instalación

> install.packages("devtools")
> library(devtools)
> install_github("Acturio/rsrvyest")

# Requisitos

Para poder hacer uso de la librería *rsrvyest*, es importante tener instalada la librería y contar con la siguiente estructura de carpetas:

  | data
    |- datos.sav
    |- lista_preguntas.xlsx
  | results
  | src
  
 # Uso
 
 Una vez que se cuenta con la carpeta *data* junto con los datos en formato .sav (spss) y la lista de pregunta que describe las preguntas para las cuales serán calculadas sus estadísticos univariados y bivariados, el análisis se realiza a través del siguiente flujo de funciones, las cuales pueden estar almacenadas en un archivo .R dentro de la carpeta *src*.
