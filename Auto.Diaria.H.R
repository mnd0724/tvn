#----------------------------------------------------------#
#Automatización Promedio Rating Hogares
#----------------------------------------------------------#
instalar_y_cargar <- function(paquetes) {
  paquetes_faltantes <- paquetes[!(paquetes %in% installed.packages()[, "Package"])]
  
  if (length(paquetes_faltantes) > 0) {
    cat("Instalando paquetes faltantes:", paquetes_faltantes, "\n")
    install.packages(paquetes_faltantes)
  }
  
  lapply(paquetes, require, character.only = TRUE)
}
paquetes_necesarios <- c("readxl", "openxlsx")
instalar_y_cargar(paquetes_necesarios)
library(readxl)
library(openxlsx)
#----------------------------------------------------------#
#Ingresar excel
#----------------------------------------------------------#
cat("Selecciona el archivo 24H\n")
archivo_excel <- file.choose()

cat("Selecciona el archivo PEAK\n")
archivo_excel2 <- file.choose()

cat("Selecciona el archivo Plantilla Rating\n")
archivo_destino <- file.choose()
#----------------------------------------------------------#
#1. Hojas del Excel ----
#----------------------------------------------------------#
if (file.exists(archivo_excel)) {
  hojas <- excel_sheets(archivo_excel)
  cat("Hojas disponibles en el archivo:\n")
  print(hojas)
  
  for (i in seq_along(hojas)) {
    assign(paste0("data", i), as.data.frame(read_excel(archivo_excel, sheet = hojas[i])))
  }
  
} else {
  cat("El archivo no se encuentra en la ruta especificada. Por favor verifica la ruta.\n")
}
#----------------------------------------------------------#
#2. Modificar Data ----
#----------------------------------------------------------# 
colnames(data2)<-as.character(data2[2,])
data2 <- data2[-2, ] 
colnames(data4)<-as.character(data4[1,])
data4<-data4[-c(1,2),]
colnames(data3)<-as.character(data3[1,])
data3<-data3[-c(1,2),]
#----------------------------------------------------------#
##2.1. Formato de Hora ----
#----------------------------------------------------------#
data3$`Hora Inicio` <- as.POSIXct(data3$`Hora Inicio`, format = "%H:%M:%S")
data3$Duración1 <- gsub("^0+", "", data3$Duración)  # Eliminar ceros iniciales
data3$Duración1 <- sapply(strsplit(data3$Duración1, ":"), function(x) {
  minutos <- as.numeric(x[1])
  segundos <- as.numeric(x[2])
  return(minutos * 60 + segundos)
})

data3$`Hora Final` <- data3$`Hora Inicio` + as.difftime(data3$Duración1, units = "secs")
data3$`Hora Rango` <- paste(format(data3$`Hora Inicio`, "%H:%M"), "-", format(data3$`Hora Final`, "%H:%M"))
fecha <- as.Date(data4$Fecha[1], format = "%d/%m/%Y")
dia_semana<- tools::toTitleCase(weekdays(fecha))
#----------------------------------------------------------#
# 4. Modificación del Título según Día de la Semana
#----------------------------------------------------------#
# Convertir la fecha a formato Date y obtener el día de la semana
fecha <- as.Date(data4$Fecha[1], format = "%d/%m/%Y")
dia_semana <- tools::toTitleCase(weekdays(fecha))

# Modificar el título "24 HORAS CENTRAL" según el día de la semana
if (dia_semana == "Sábado") {
  data3$Título[data3$Título == "24 HORAS CENTRAL SABADO"] <- "24 HORAS CENTRAL"
  data4$Título[data4$Título == "24 HORAS CENTRAL SABADO"] <- "24 HORAS CENTRAL"
} else if (dia_semana == "Domingo") {
  data3$Título[data3$Título == "24 HORAS CENTRAL DOMINGO"] <- "24 HORAS CENTRAL"
  data4$Título[data4$Título == "24 HORAS CENTRAL DOMINGO"] <- "24 HORAS CENTRAL"
}

data4 <- data4[data4$Título %in% c("24 HORAS CENTRAL", "TELETRECE") &
                 data4$Canal %in% c("Television Nacional CH7", "Canal 13"), ]
data3 <- data3[data3$Título %in% c("24 HORAS CENTRAL", "TELETRECE") &
                 data3$Canal %in% c("Television Nacional CH7", "Canal 13"), ]
#----------------------------------------------------------#
#4. Agregar valores al Excel ----
#----------------------------------------------------------# 
##4.1. Semana ----
#----------------------------------------------------------# 
###4.1.1. Título ----
#----------------------------------------------------------# 
wb <- loadWorkbook(archivo_destino)
titulo_actual <- readWorkbook(wb, sheet = 1, colNames = FALSE, rows = 3, cols = 3)
nuevo_titulo <- paste0("PROMEDIO RATING HOGARES ", fecha, " CANAL 24 HORAS")
writeData(wb, sheet = 1, nuevo_titulo, startCol = 3, startRow = 3)
saveWorkbook(wb, archivo_destino, overwrite = TRUE)
#----------------------------------------------------------# 
###4.1.2. Promedio Rating ----
#----------------------------------------------------------# 
writeData(wb, sheet = 1, x = dia_semana, startCol = 3, startRow = 6)
writeData(wb, sheet = 1, x = data2[2, "24 Horas SD+HD"], startCol = 4, startRow = 6)
writeData(wb, sheet = 1, x = data2[2, "CNN Chile"], startCol = 5, startRow = 6)
writeData(wb, sheet = 1, x = data2[2, "T13 en vivo"], startCol = 6, startRow = 6)
writeData(wb, sheet = 1, x = data2[2, "Meganoticias Ahora"], startCol = 7, startRow = 6)
writeData(wb, sheet = 1, x = data2[2, "TNT Sports Premium"], startCol = 8, startRow = 6)
writeData(wb, sheet = 1, x = data2[2, "ESPN"], startCol = 9, startRow = 6)
#----------------------------------------------------------# 
##4.2. Diario ----
#----------------------------------------------------------# 
writeData(wb, sheet = 1, x = fecha, startCol = 3, startRow = 8)
writeData(wb, sheet = 1, x = data3[data3$Título == "24 HORAS CENTRAL", "Hora Rango"], startCol = 4, startRow = 9)
writeData(wb, sheet = 1, x = data4[data4$Canal == "Television Nacional CH7", "24 Horas SD+HD"], startCol = 5, startRow = 9)
writeData(wb, sheet = 1, x = data4[data4$Canal == "Television Nacional CH7", "Television Nacional CH7"], startCol =6, startRow = 9)
writeData(wb, sheet = 1, x = data4[data4$Canal == "Television Nacional CH7", "TVN + 24H"], startCol =7, startRow = 9)
writeData(wb, sheet = 1, x = data3[data3$Canal == "Television Nacional CH7", "24 Horas SD+HD"], startCol = 8, startRow = 9)
writeData(wb, sheet = 1, x = data3[data3$Canal == "Television Nacional CH7", "Television Nacional CH7"], startCol =9, startRow = 9)
writeData(wb, sheet = 1, x = data3[data3$Canal == "Television Nacional CH7", "TVN + 24H"], startCol =10, startRow = 9)  

writeData(wb, sheet = 1, x = fecha, startCol = 3, startRow = 11)
writeData(wb, sheet = 1, x = data3[data3$Título == "TELETRECE", "Hora Rango"], startCol = 4, startRow = 12)
writeData(wb, sheet = 1, x = data4[data4$Canal == "Canal 13", "T13 en vivo"], startCol = 5, startRow = 12)
writeData(wb, sheet = 1, x = data4[data4$Canal == "Canal 13", "Canal 13"], startCol =6, startRow = 12)
writeData(wb, sheet = 1, x = data4[data4$Canal == "Canal 13", "T13 + C13"], startCol =7, startRow = 12)
writeData(wb, sheet = 1, x = data3[data3$Canal == "Canal 13", "T13 en vivo"], startCol = 8, startRow = 12)
writeData(wb, sheet = 1, x = data3[data3$Canal == "Canal 13", "Canal 13"], startCol =9, startRow = 12)
writeData(wb, sheet = 1, x = data3[data3$Canal == "Canal 13", "T13 + C13"], startCol =10, startRow = 12)

writeData(wb, sheet = 1, x = fecha, startCol = 3, startRow = 14)

# Guardar el archivo actualizado
saveWorkbook(wb, archivo_destino, overwrite = TRUE)
#----------------------------------------------------------#
###4.2.1. Datos Peak ----
#----------------------------------------------------------#
if (file.exists(archivo_excel2)) {
  cat("El archivo PEAK 24H.xlsx existe en la ruta especificada.\n")
  
  # Obtener la lista de hojas del archivo Excel
  hojas <- excel_sheets(archivo_excel2)
  cat("Hojas disponibles en el archivo PEAK 24H.xlsx:\n")
  print(hojas)
  
  # Leer la primera hoja del archivo Excel y guardarla como un data frame
  peak_data <- as.data.frame(read_excel(archivo_excel2, sheet = hojas[1]))
  cat("Data frame creado: peak_data\n")
  
  # Mostrar las primeras filas del data frame creado
  cat("\nPrimeras filas de peak_data:\n")
  print(head(peak_data))
} else {
  cat("El archivo PEAK 24H.xlsx no se encuentra en la ruta especificada. Por favor verifica la ruta.\n")
}

colnames(peak_data)<-as.character(peak_data[2,])
peak_data <- peak_data[-c(1,2), ] 

# Crear una función para ajustar las horas de Franjas
adjust_hours <- function(time_string) {
  # Separar la hora y los minutos
  time_parts <- strsplit(time_string, ":")[[1]]
  hour <- as.numeric(time_parts[1])
  minutes <- time_parts[2]
  
  # Ajustar la hora si es mayor o igual a 24
  if (hour >= 24) {
    hour <- hour - 24
  }
  
  # Formatear la nueva hora y mantener los minutos
  adjusted_time <- sprintf("%02d:%s", hour, minutes)
  return(adjusted_time)
}

# Aplicar la función a la variable Franjas en peak_data
peak_data$Franjas <- sapply(peak_data$Franjas, adjust_hours)

peak_data$Franjas<- as.POSIXct(peak_data$Franjas, format = "%H:%M")

#----------------------------------------------------------#
####4.2.1.1. 24H ----
#----------------------------------------------------------#

# Función para encontrar el intervalo continuo de tiempo con el puntaje máximo
encontrar_peak_continuo1 <- function(data, variable, tiempo) {
  max_valor <- max(data[[variable]], na.rm = TRUE)
  max_franjas <- sort(data[data[[variable]] == max_valor, tiempo])
  
  # Encontrar los intervalos continuos
  diferencias <- c(0, diff(max_franjas))
  cortes <- cumsum(diferencias != 60)
  
  intervalos <- split(max_franjas, cortes)
  intervalos_continuos <- lapply(intervalos, function(x) {
    if (length(x) > 1) {
      paste(format(min(x), "%H:%M"), "-", format(max(x), "%H:%M"))
    } else {
      paste(format(x, "%H:%M"), "-", format(x, "%H:%M"))
    }
  })
  
  return(unlist(intervalos_continuos))
}

# Ejecutar la función para encontrar el peak continuo de "24 Horas SD+HD"
intervalos_peak <- encontrar_peak_continuo1(peak_data, "24 Horas SD+HD", "Franjas")
cat("\nIntervalos de tiempo con el peak continuo:\n")
print(intervalos_peak)

# Función para encontrar los intervalos continuos de tiempo con el puntaje máximo
encontrar_peak_continuo <- function(intervalos_peak) {
  # Convertir los intervalos a formato de horas
  tiempos <- lapply(intervalos_peak, function(x) {
    horas <- strsplit(x, " - ")[[1]]
    as.POSIXct(horas, format = "%H:%M")
  })
  
  # Si solo hay un intervalo, devolverlo tal cual
  if (length(tiempos) == 1) {
    return(paste("[", format(tiempos[[1]][1], "%H:%M"), "-", format(tiempos[[1]][2], "%H:%M"), ">"))
  }
  
  # Inicializar variables para agrupar intervalos
  inicio <- tiempos[[1]][1]
  fin <- tiempos[[1]][2]
  intervalos_largos <- c()
  
  # Agrupar intervalos continuos
  for (i in 2:length(tiempos)) {
    if (difftime(tiempos[[i]][1], fin, units = "mins") == 1) {
      fin <- tiempos[[i]][2]
    } else {
      intervalos_largos <- c(intervalos_largos, paste("[", format(inicio, "%H:%M"), "-", format(fin, "%H:%M"), ">"))
      inicio <- tiempos[[i]][1]
      fin <- tiempos[[i]][2]
    }
  }
  
  # Añadir el último intervalo
  intervalos_largos <- c(intervalos_largos, paste("[", format(inicio, "%H:%M"), "-", format(fin, "%H:%M"), ">"))
  
  return(intervalos_largos)
}

intervalos_largos <- encontrar_peak_continuo(intervalos_peak)
print(intervalos_largos)
#----------------------------------------------------------#
####4.2.1.2. 24H Central----
#----------------------------------------------------------#
# Filtrar las filas de data3 donde el título es "24 HORAS CENTRAL"
data3_central <- data3[data3$Título == "24 HORAS CENTRAL", ]
peak_data_filtrado <- peak_data[peak_data$Franjas >= data3_central$`Hora Inicio` & peak_data$Franjas < data3_central$`Hora Final`, ]

# Ejecutar la función para encontrar el peak continuo de "24 Horas SD+HD"
intervalos_peak_c <- encontrar_peak_continuo1(peak_data_filtrado, "24 Horas SD+HD", "Franjas")
cat("\nIntervalos de tiempo con el peak continuo:\n")
print(intervalos_peak_c)
intervalos_largos_c <- encontrar_peak_continuo(intervalos_peak_c) 

# Función para limpiar el formato de los intervalos y calcular la duración
calcular_duracion <- function(intervalo) {
  intervalo <- gsub("\\[|\\>|\\s+", "", intervalo)  # Eliminar corchetes, flechas y espacios adicionales
  tiempos <- strsplit(intervalo, "-")[[1]]
  inicio <- as.POSIXct(tiempos[1], format = "%H:%M")
  fin <- as.POSIXct(tiempos[2], format = "%H:%M")
  return(as.numeric(difftime(fin, inicio, units = "mins")))
}

# Seleccionar primero el intervalo final para intervalos_largos_c (24 HORAS CENTRAL DOMINGO)
if (length(intervalos_largos_c) == 1) {
  intervalo_final_c <- intervalos_largos_c[1]
} else {
  duraciones_c <- sapply(intervalos_largos_c, calcular_duracion)
  intervalo_final_c <- intervalos_largos_c[which.max(duraciones_c)]
}

# Seleccionar luego el intervalo final para intervalos_largos (día completo), excluyendo el de 24 HORAS CENTRAL DOMINGO
if (length(intervalos_largos) == 1) {
  intervalo_final <- intervalos_largos[1]
} else {
  candidatos <- setdiff(intervalos_largos, intervalo_final_c)
  candidatos <- candidatos[!is.na(candidatos)]  # Eliminar posibles NAs
  duraciones <- sapply(candidatos, calcular_duracion)
  intervalo_final <- candidatos[which.max(duraciones)]
}

# Imprimir los intervalos finales seleccionados
cat("Intervalo final para 24 HORAS CENTRAL:", intervalo_final_c, "\n")
cat("Intervalo final para el día completo:", intervalo_final, "\n")

writeData(wb, sheet = 1, x = intervalo_final_c, startCol = 4, startRow = 15)
writeData(wb, sheet = 1, x = intervalo_final, startCol = 4, startRow = 16)
writeData(wb, sheet = 1, x = max(peak_data_filtrado$`24 Horas SD+HD`), startCol = 5, startRow = 15)
writeData(wb, sheet = 1, x = max(peak_data$`24 Horas SD+HD`), startCol = 5, startRow = 16)
saveWorkbook(wb, archivo_destino, overwrite = TRUE)



