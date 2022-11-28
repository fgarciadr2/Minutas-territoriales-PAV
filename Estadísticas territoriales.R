tictoc::tic()
rm(list = ls())

#####Elegir carpeta donde se guardarán los archivos#####

wdoutput <- "C:/Users/felip/OneDrive - Microsoft/PAV/Minutas Territoriales PAV"

#####Seleccionar años####

anoa <- 2022
anob <- 2022
mes <- month(Sys.Date()) - 1
#####Importar paquetes########

library(beepr)
library(xlsx)
library(data.table)
library(readxl)
library(scales)
library(ggrepel)
library(janitor)
library(lubridate)
library(stringr)
library(tidyverse)


####Funciones####

#Función para leer directorio github

read_git <- function(x){
  y = paste("https://raw.githubusercontent.com/fgarciadr2/Minutas-territoriales-PAV/fgarciadr2-Bases-de-datos",
        x,
        sep = "/") %>% 
    fread
  return(y)
}

#Función para eliminar caractéres del español y espacios en los nombres de columnas
limpiar_cat <- function(x){
  colnames(x) <- tolower(colnames(x))
  colnames(x) <- gsub(" ", ".", colnames(x))
  colnames(x) <- gsub("\\.de\\.", ".", colnames(x))
  colnames(x) <- gsub("°", "", colnames(x))
  colnames(x) <- chartr("áéíóúñ", "aeioun", colnames(x))
  return(x)
}

#Función para homogeneizar nombres de comunas que aparecen escritas de formas distintas en distintas bases de datos


limpiar_comunas <- function(x){
  x <- x %>% 
    mutate(comuna.delito = if_else(comuna.delito == "Alto Bío Bío", "Alto Biobío", comuna.delito),
           comuna.delito = if_else(comuna.delito == "Concon", "Concón", comuna.delito),
           comuna.delito = if_else(comuna.delito == "Curaco", "Curaco de Vélez", comuna.delito),
           comuna.delito = if_else(comuna.delito == "Llayllay", "Llay Llay", comuna.delito),
           comuna.delito = if_else(comuna.delito == "Los Alamos", "Los Álamos", comuna.delito),
           comuna.delito = if_else(comuna.delito == "Los Angeles", "Los Ángeles", comuna.delito),
           comuna.delito = if_else(comuna.delito == "Mafil", "Máfil", comuna.delito),
           comuna.delito = if_else(comuna.delito == "Marchigue", "Marchigüe", comuna.delito),
           comuna.delito = if_else(comuna.delito == "Puerto Natales", "Natales", comuna.delito),
           comuna.delito = if_else(comuna.delito == "O' Higgins", "O'Higgins", comuna.delito),
           comuna.delito = if_else(comuna.delito == "OHiggins", "O'Higgins", comuna.delito),
           comuna.delito = if_else(comuna.delito == "Quinta", "Quinta de Tilcoco", comuna.delito),
           comuna.delito = if_else(comuna.delito == "Requiona", "Requiona", comuna.delito),
           comuna.delito = if_else(comuna.delito == "San Vicente", "San Vicente de Tagua Tagua", comuna.delito),
           comuna.delito = if_else(comuna.delito == "Til Til", "Tiltil", comuna.delito),
           comuna.delito = if_else(comuna.delito == "Tirua", "Tirúa", comuna.delito)
    )
  return(x)
}

####Importar datos####


ssra <- read_git("SSR.csv") %>% 
  limpiar_cat() %>% 
  limpiar_comunas()
ccpa <- read_git("CCP.csv") %>% 
  limpiar_cat() %>% 
  limpiar_comunas()
llamadasa <- read_git("Llamadas.csv") %>% 
  limpiar_cat() %>% 
  limpiar_comunas()
ccp_casosa <- read_git("CCP_Casos.csv") %>% 
  limpiar_cat() %>% 
  limpiar_comunas()
ipa <- read_git("IP.csv") %>% 
  limpiar_cat()
rec_reg <- fread("https://raw.githubusercontent.com/fgarciadr2/Minutas-territoriales-PAV/fgarciadr2-Bases-de-datos/Regiones%20y%20comunas.csv")
sica <- read_git("SIC.csv") %>% 
  limpiar_cat() %>% 
  rename("comuna.delito" = "comuna.ocurrencia") %>% 
  limpiar_comunas()


####Joins####

#Utilizamos la base rec_reg que contiene la región, comuna, provincia y distrito electoral correspondiente a cada territorio

ssr <- left_join(ssra, rec_reg %>% 
                   rename("comuna.delito" = "comuna",
                          "region.delito" = "region",
                          "provincia.delito" = "provincia",
                          "distrito.delito" = "distrito"),
                 by = "comuna.delito")

sic <- left_join(sica, rec_reg %>% 
                   rename("comuna.delito" = "comuna",
                          "region.delito" = "region",
                          "provincia.delito" = "provincia",
                          "distrito.delito" = "distrito"),
                 by = "comuna.delito") %>% 
  mutate(ano.finalizacion = year(fecha.finalizacion),
         ano.recepcion = year(fecha.recepcion))

ccp <- left_join(ccpa, rec_reg %>% 
                   rename("comuna.delito" = "comuna",
                          "region.delito" = "region",
                          "provincia.delito" = "provincia",
                          "distrito.delito" = "distrito"),
                 by = "comuna.delito")

ip <- ipa %>% 
  filter(categoria == "Directa c/sesion",
         asistencia == "Si") %>% 
  rename(n.expediente = `n.expediente/s`) %>% 
  mutate(n.expediente = gsub(".*,", "", n.expediente),
         ano.actividad = year(`fecha/hora.inicio`),
         mes.actividad = month(`fecha/hora.inicio`)
  ) %>% 
  left_join(ssr %>% select(n.expediente, comuna.delito),
            by = "n.expediente") %>% 
  left_join(rec_reg %>% 
              rename("comuna.delito" = "comuna",
                     "region.delito" = "region",
                     "provincia.delito" = "provincia",
                     "distrito.delito" = "distrito"),
            by = "comuna.delito")

llamadas <- llamadasa %>% 
  mutate(fecha.llamada = dmy(fecha.llamada),
         ano.llamada = year(fecha.llamada),
         mes.llamada = month(fecha.llamada)) %>% 
  select(-comuna.delito) %>%
  rename(n = id.caso) %>% 
  left_join(sic %>% select(n, comuna.delito)) %>% 
  left_join(rec_reg %>% 
              rename("comuna.delito" = "comuna",
                     "region.delito" = "region",
                     "provincia.delito" = "provincia",
                     "distrito.delito" = "distrito"),
            by = "comuna.delito")


ccp_casos <- ccp_casosa %>% 
  select(-region.delito) %>% 
  left_join(rec_reg %>% 
              rename("comuna.delito" = "comuna",
                     "region.delito" = "region",
                     "provincia.delito" = "provincia",
                     "distrito.delito" = "distrito"),
            by = "comuna.delito")

#####Tablas ingresos####

if(wdoutput == "" | exists("wdoutput") == F){
  print("Se requiere elegir un nombre de la carpeta donde irán los archivos")
  stop()
} else{
  setwd(wdoutput) 
}

setwd(wdoutput)

dir.create("Comunas")
dir.create("Regiones")
dir.create("Provincias")
dir.create("Distritos")


####Tablas comunales####


setwd("Comunas")

for(i in 1 : nrow(rec_reg)){
  com <- rec_reg$comuna[i]
  
  ssr_ing <- ssr %>% 
    filter(comuna.delito == com,
           ano.ingreso >= anoa,
           ano.ingreso <= anob,
    )  %>% 
    tally(name = "SSR")
  
  sic_ing <- sic %>% 
    filter(comuna.delito == com,
           ano.finalizacion >= anoa,
           ano.finalizacion <= anob,
    )  %>% 
    tally(name = "SIC")
  
  ccp_ing <- ccp %>% 
    filter(comuna.delito == com,
           ano.ingreso >= anoa,
           ano.ingreso <= anob,
    )  %>% 
    tally(name = "CCP")
  
  tab_ing <- cbind(ssr_ing, sic_ing, ccp_ing) %>% 
    adorn_totals("col") %>% 
    t() %>% 
    as.data.frame() %>% 
    rownames_to_column()
  
  colnames(tab_ing) <- c("Servicio","Ingresos")
  
  ssr_at <- ssr %>% 
    filter(comuna.delito == com,
           (ano.ingreso <= anoa & ano.ingreso >= anob) |
             (ano.ingreso <= anoa & tipo.finalizacion == "NULL") 
    )  %>% 
    tally(name = "SSR")
  
  sic_at <- sic %>% 
    mutate(ano.recepcion = year(fecha.recepcion)) %>% 
    filter(between(ano.finalizacion, anoa, anob) |
             (ano.recepcion <= anoa & ano.finalizacion >= anob) |
             (ano.finalizacion <= anoa & tipo.finalizacion == "NULL")|
             (ano.finalizacion <= anoa & tipo.finalizacion == "")) %>% 
    filter(comuna.delito == com) %>% 
    tally(name = "SIC")
  
  ccp_at <- ccp %>% 
    filter(comuna.delito == com,
           (ano.ingreso <= anoa & ano.ingreso >= anob) |
             (ano.ingreso <= anoa & tipo.finalizacion == "NULL") 
    )  %>% 
    tally(name = "CCP")
  
  tab_at <- cbind(ssr_at, sic_at, ccp_at) %>% 
    adorn_totals("col") %>% 
    t() %>% 
    as.data.frame() %>% 
    rownames_to_column()
  
  colnames(tab_at) <- c("Servicio","Personas atendidas")
  
  tab_com <- full_join(tab_ing, tab_at, by = "Servicio")
  
  ssr_pres <- ip %>% 
    filter(ano.actividad >= anoa
           & !(ano.actividad == anob & mes.actividad > mes)
           & comuna.delito == com) %>% 
    tally(name = "SSR")
  
  sic_pres <- llamadas %>% 
    filter(ano.llamada >= anoa & ano.llamada <= anob
           & comuna.delito == com
    ) %>% 
    tally(name = "SIC")
  
  ccp_pres <- ccp_casos %>% 
    filter(ano.reporte >= anoa,
           ano.reporte <= anob,
           comuna.delito == com) %>% 
    tally(name = "CCP")
  
  tab_pres <- cbind(ssr_pres, sic_pres, ccp_pres) %>% 
    adorn_totals("col") %>% 
    t() %>% 
    as.data.frame() %>% 
    rownames_to_column()
  
  colnames(tab_pres) <- c("Servicio","Atenciones")
  
  tab_com <- full_join(tab_com, tab_pres, by = "Servicio")

  write.xlsx(tab_com, paste(com, "xlsx", sep = "."))  
  
}


####Tablas regionales####

setwd("..")
setwd("Regiones")

regiones <- unique(rec_reg$region)


for(i in 1 : length(unique(rec_reg$region))){
  reg <- unique(rec_reg$region)[i]
  
  ssr_ing <- ssr %>% 
    filter(region.delito == reg,
           ano.ingreso >= anoa,
           ano.ingreso <= anob,
    )  %>% 
    group_by(comuna.delito) %>% 
    tally(name = "SSR")
  
  sic_ing <- sic %>% 
    filter(region.delito == reg,
           ano.finalizacion >= anoa,
           ano.finalizacion <= anob,
    )  %>% 
    group_by(comuna.delito) %>% 
    tally(name = "SIC")
  
  ccp_ing <- ccp %>% 
    filter(region.delito == reg,
           ano.ingreso >= anoa,
           ano.ingreso <= anob,
    )  %>% 
    group_by(comuna.delito) %>% 
    tally(name = "CCP")
  
  tab_ing <- full_join(ssr_ing, sic_ing, by = "comuna.delito") %>% 
    full_join(ccp_ing, by = "comuna.delito") %>% 
    mutate_if(is.numeric, replace_na, 0) %>% 
    adorn_totals(c("col")) %>% 
    arrange(-Total) %>% 
    adorn_totals("row") %>% 
    rename("Comuna del delito" = comuna.delito)
  
  write.xlsx(tab_ing, paste(reg, "xlsx", sep = "."), sheetName = "Ingresos")
  
  ssr_at <- ssr %>% 
    filter(region.delito == reg,
           (ano.ingreso <= anoa & ano.ingreso >= anob) |
             (ano.ingreso <= anoa & tipo.finalizacion == "NULL") 
    )  %>% 
    group_by(comuna.delito) %>% 
    tally(name = "SSR")
  
  sic_at <- sic %>% 
    mutate(ano.recepcion = year(fecha.recepcion)) %>% 
    filter(between(ano.finalizacion, anoa, anob) |
             (ano.recepcion <= anoa & ano.finalizacion >= anob) |
             (ano.finalizacion <= anoa & tipo.finalizacion == "NULL")|
             (ano.finalizacion <= anoa & tipo.finalizacion == "")) %>% 
    filter(region.delito == reg) %>% 
    group_by(comuna.delito) %>% 
    tally(name = "SIC")
  
  ccp_at <- ccp %>% 
    filter(region.delito == reg,
           (ano.ingreso <= anoa & ano.ingreso >= anob) |
             (ano.ingreso <= anoa & tipo.finalizacion == "NULL") 
    )  %>% 
    group_by(comuna.delito) %>% 
    tally(name = "CCP")
  
  tab_at <- full_join(ssr_at, sic_at, by = "comuna.delito") %>% 
    full_join(ccp_at, by = "comuna.delito") %>% 
    mutate_if(is.numeric, replace_na, 0) %>% 
    adorn_totals(c("col")) %>% 
    arrange(-Total) %>% 
    adorn_totals("row") %>% 
    rename("Comuna del delito" = comuna.delito)
  
  write.xlsx(tab_at, paste(reg, "xlsx", sep = "."), sheetName = "Personas atendidas", append = T)
  
  ssr_pres <- ip %>% 
    filter(ano.actividad >= anoa
           & !(ano.actividad == anob & mes.actividad > mes)
           & region.delito == reg) %>%
    group_by(comuna.delito) %>% 
    tally(name = "SSR")
  
  sic_pres <- llamadas %>% 
    filter(ano.llamada >= anoa & ano.llamada <= anob
           & region.delito == reg
    ) %>% 
    group_by(comuna.delito) %>% 
    tally(name = "SIC")
  
  ccp_pres <- ccp_casos %>% 
    filter(ano.reporte >= anoa,
           ano.reporte <= anob,
           region.delito == reg) %>%
    group_by(comuna.delito) %>% 
    tally(name = "CCP")
  
  tab_pres <- full_join(ssr_pres, sic_pres, by = "comuna.delito") %>% 
    full_join(ccp_pres, by = "comuna.delito") %>% 
    mutate_if(is.numeric, replace_na, 0) %>% 
    adorn_totals(c("col")) %>% 
    arrange(-Total) %>% 
    adorn_totals("row") %>% 
    rename("Comuna del delito" = comuna.delito)
  
  write.xlsx(tab_pres, paste(reg, "xlsx", sep = "."), sheetName = "Atenciones", append = T)

  pav_tot <- tab_ing %>% select("Comuna del delito", Ingresos = Total) %>% 
    full_join(tab_at %>% select("Comuna del delito", "Personas atendidas" = Total), by = "Comuna del delito") %>% 
    full_join(tab_pres %>% select("Comuna del delito", "Atenciones" = Total), by = "Comuna del delito") %>% 
    filter(`Comuna del delito` != "Total") %>% 
    mutate_if(is.numeric, replace_na, 0) %>% 
    arrange(-Atenciones) %>% 
    untabyl() %>%
    adorn_totals("row")
  
  write.xlsx(pav_tot, paste(reg, "xlsx", sep = "."), sheetName = "Totales", append = T)
}


####Tablas provinciales####

setwd("..")
setwd("Provincias")


for(i in 1 : length(unique(rec_reg$provincia))){
  prov <- unique(rec_reg$provincia)[i]
  
  ssr_ing <- ssr %>% 
    filter(provincia.delito == prov,
           ano.ingreso >= anoa,
           ano.ingreso <= anob,
    )  %>% 
    group_by(comuna.delito) %>% 
    tally(name = "SSR")
  
  sic_ing <- sic %>% 
    filter(provincia.delito == prov,
           ano.finalizacion >= anoa,
           ano.finalizacion <= anob,
    )  %>% 
    group_by(comuna.delito) %>% 
    tally(name = "SIC")
  
  ccp_ing <- ccp %>% 
    filter(provincia.delito == prov,
           ano.ingreso >= anoa,
           ano.ingreso <= anob,
    )  %>% 
    group_by(comuna.delito) %>% 
    tally(name = "CCP")
  
  tab_ing <- full_join(ssr_ing, sic_ing, by = "comuna.delito") %>% 
    full_join(ccp_ing, by = "comuna.delito") %>% 
    mutate_if(is.numeric, replace_na, 0) %>% 
    adorn_totals(c("col")) %>% 
    arrange(-Total) %>% 
    adorn_totals("row") %>% 
    rename("Comuna del delito" = comuna.delito)
  
  write.xlsx(tab_ing, paste(prov, "xlsx", sep = "."), sheetName = "Ingresos")
  
  ssr_at <- ssr %>% 
    filter(provincia.delito == prov,
           (ano.ingreso <= anoa & ano.ingreso >= anob) |
             (ano.ingreso <= anoa & tipo.finalizacion == "NULL") 
    )  %>% 
    group_by(comuna.delito) %>% 
    tally(name = "SSR")
  
  sic_at <- sic %>% 
    mutate(ano.recepcion = year(fecha.recepcion)) %>% 
    filter(between(ano.finalizacion, anoa, anob) |
             (ano.recepcion <= anoa & ano.finalizacion >= anob) |
             (ano.finalizacion <= anoa & tipo.finalizacion == "NULL")|
             (ano.finalizacion <= anoa & tipo.finalizacion == "")) %>% 
    filter(provincia.delito == prov) %>% 
    group_by(comuna.delito) %>% 
    tally(name = "SIC")
  
  ccp_at <- ccp %>% 
    filter(provincia.delito == prov,
           (ano.ingreso <= anoa & ano.ingreso >= anob) |
             (ano.ingreso <= anoa & tipo.finalizacion == "NULL") 
    )  %>% 
    group_by(comuna.delito) %>% 
    tally(name = "CCP")
  
  tab_at <- full_join(ssr_at, sic_at, by = "comuna.delito") %>% 
    full_join(ccp_at, by = "comuna.delito") %>% 
    mutate_if(is.numeric, replace_na, 0) %>% 
    adorn_totals(c("col")) %>% 
    arrange(-Total) %>% 
    adorn_totals("row") %>% 
    rename("Comuna del delito" = comuna.delito)
  
  write.xlsx(tab_at, paste(prov, "xlsx", sep = "."), sheetName = "Personas atendidas", append = T)
  
  ssr_pres <- ip %>% 
    filter(ano.actividad >= anoa
           & !(ano.actividad == anob & mes.actividad > mes)
           & provincia.delito == prov) %>%
    group_by(comuna.delito) %>% 
    tally(name = "SSR")
  
  sic_pres <- llamadas %>% 
    filter(ano.llamada >= anoa & ano.llamada <= anob
           & provincia.delito == prov
    ) %>% 
    group_by(comuna.delito) %>% 
    tally(name = "SIC")
  
  ccp_pres <- ccp_casos %>% 
    filter(ano.reporte >= anoa,
           ano.reporte <= anob,
           provincia.delito == prov) %>%
    group_by(comuna.delito) %>% 
    tally(name = "CCP")
  
  tab_pres <- full_join(ssr_pres, sic_pres, by = "comuna.delito") %>% 
    full_join(ccp_pres, by = "comuna.delito") %>% 
    mutate_if(is.numeric, replace_na, 0) %>% 
    adorn_totals(c("col")) %>% 
    arrange(-Total) %>% 
    adorn_totals("row") %>% 
    rename("Comuna del delito" = comuna.delito)
  
  write.xlsx(tab_pres, paste(prov, "xlsx", sep = "."), sheetName = "Atenciones", append = T)
  
  pav_tot <- tab_ing %>% select("Comuna del delito", Ingresos = Total) %>% 
    full_join(tab_at %>% select("Comuna del delito", "Personas atendidas" = Total), by = "Comuna del delito") %>% 
    full_join(tab_pres %>% select("Comuna del delito", "Atenciones" = Total), by = "Comuna del delito") %>% 
    filter(`Comuna del delito` != "Total") %>% 
    mutate_if(is.numeric, replace_na, 0) %>% 
    arrange(-Atenciones) %>% 
    untabyl() %>%
    adorn_totals("row") 
  
  write.xlsx(pav_tot, paste(prov, "xlsx", sep = "."), sheetName = "Totales", append = T)
}


####Tablas distritales####

setwd("..")
setwd("Distritos")


for(i in 1 : length(unique(rec_reg$distrito))){
  dis <- distritos[i]
  
  ssr_ing <- ssr %>% 
    filter(distrito.delito == dis,
           ano.ingreso >= anoa,
           ano.ingreso <= anob,
    )  %>% 
    group_by(comuna.delito) %>% 
    tally(name = "SSR")
  
  sic_ing <- sic %>% 
    filter(distrito.delito == dis,
           ano.finalizacion >= anoa,
           ano.finalizacion <= anob,
    )  %>% 
    group_by(comuna.delito) %>% 
    tally(name = "SIC")
  
  ccp_ing <- ccp %>% 
    filter(distrito.delito == dis,
           ano.ingreso >= anoa,
           ano.ingreso <= anob,
    )  %>% 
    group_by(comuna.delito) %>% 
    tally(name = "CCP")
  
  tab_ing <- full_join(ssr_ing, sic_ing, by = "comuna.delito") %>% 
    full_join(ccp_ing, by = "comuna.delito") %>% 
    mutate_if(is.numeric, replace_na, 0) %>% 
    adorn_totals(c("col")) %>% 
    arrange(-Total) %>% 
    adorn_totals("row") %>% 
    rename("Comuna del delito" = comuna.delito)
  
  write.xlsx(tab_ing, paste(dis, "xlsx", sep = "."), sheetName = "Ingresos")
  
  ssr_at <- ssr %>% 
    filter(distrito.delito == dis,
           (ano.ingreso <= anoa & ano.ingreso >= anob) |
             (ano.ingreso <= anoa & tipo.finalizacion == "NULL") 
    )  %>% 
    group_by(comuna.delito) %>% 
    tally(name = "SSR")
  
  sic_at <- sic %>% 
    mutate(ano.recepcion = year(fecha.recepcion)) %>% 
    filter(between(ano.finalizacion, anoa, anob) |
             (ano.recepcion <= anoa & ano.finalizacion >= anob) |
             (ano.finalizacion <= anoa & tipo.finalizacion == "NULL")|
             (ano.finalizacion <= anoa & tipo.finalizacion == "")) %>% 
    filter(distrito.delito == dis) %>% 
    group_by(comuna.delito) %>% 
    tally(name = "SIC")
  
  ccp_at <- ccp %>% 
    filter(distrito.delito == dis,
           (ano.ingreso <= anoa & ano.ingreso >= anob) |
             (ano.ingreso <= anoa & tipo.finalizacion == "NULL") 
    )  %>% 
    group_by(comuna.delito) %>% 
    tally(name = "CCP")
  
  tab_at <- full_join(ssr_at, sic_at, by = "comuna.delito") %>% 
    full_join(ccp_at, by = "comuna.delito") %>% 
    mutate_if(is.numeric, replace_na, 0) %>% 
    adorn_totals(c("col")) %>% 
    arrange(-Total) %>% 
    adorn_totals("row") %>% 
    rename("Comuna del delito" = comuna.delito)
  
  write.xlsx(tab_at, paste(dis, "xlsx", sep = "."), sheetName = "Personas atendidas", append = T)
  
  ssr_pres <- ip %>% 
    filter(ano.actividad >= anoa
           & !(ano.actividad == anob & mes.actividad > mes)
           & distrito.delito == dis) %>%
    group_by(comuna.delito) %>% 
    tally(name = "SSR")
  
  sic_pres <- llamadas %>% 
    filter(ano.llamada >= anoa & ano.llamada <= anob
           & distrito.delito == dis
    ) %>% 
    group_by(comuna.delito) %>% 
    tally(name = "SIC")
  
  ccp_pres <- ccp_casos %>% 
    filter(ano.reporte >= anoa,
           ano.reporte <= anob,
           distrito.delito == dis) %>%
    group_by(comuna.delito) %>% 
    tally(name = "CCP")
  
  tab_pres <- full_join(ssr_pres, sic_pres, by = "comuna.delito") %>% 
    full_join(ccp_pres, by = "comuna.delito") %>% 
    mutate_if(is.numeric, replace_na, 0) %>% 
    adorn_totals(c("col")) %>% 
    arrange(-Total) %>% 
    adorn_totals("row") %>% 
    rename("Comuna del delito" = comuna.delito)
  
  write.xlsx(tab_pres, paste(dis, "xlsx", sep = "."), sheetName = "Atenciones", append = T)
  
  pav_tot <- tab_ing %>% select(`Comuna del delito`, Ingresos = Total) %>% 
    full_join(tab_at %>% select(`Comuna del delito`, "Personas atendidas" = Total), by = "Comuna del delito") %>% 
    full_join(tab_pres %>% select(`Comuna del delito`, "Atenciones" = Total), by = "Comuna del delito") %>% 
    filter(`Comuna del delito` != "Total") %>% 
    mutate_if(is.numeric, replace_na, 0) %>% 
    arrange(-Atenciones) %>% 
    untabyl() %>%
    adorn_totals("row") 

  write.xlsx(pav_tot, paste(dis, "xlsx", sep = "."), sheetName = "Totales", append = T)
}


tictoc::toc()
beepr::beep()