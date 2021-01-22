'Script Post-Trading'

#install.packages(c("readxl", "dplyr", "memisc", "gdata"))

library(readxl)
library(dplyr)
library(memisc)
library(gdata)
library(writexl)

Bloombergfile.input.path <- '/Users/luis.magosso/Documents/Boletas/Bloomberg/Boleta.xlsx'
Sqia.output.path <- '/Users/luis.magosso/Documents/Boletas/Systems/Sqia.txt'
Mitra.output.path <- '/Users/luis.magosso/Documents/Boletas/Systems/Mitra.xlsx'
HistoricalOp.output.path <- '/Users/luis.magosso/Documents/Boletas/Operations/'


Boleta <- read_excel(Bloombergfile.input.path)
colnames(Boleta) = gsub(" ", "", colnames(Boleta))
Boleta <- aggregate (ExecLastFill~ ExecAsofDate + OrderAccount + Side + Ticker + ExecLastFillPx + Broker, Boleta, sum)
Boleta <- Boleta[c("ExecAsofDate", "Broker", "OrderAccount", "Side", "Ticker", "ExecLastFill", "ExecLastFillPx")]

BoletaMitra <- Boleta %>%
  group_by(ExecAsofDate, OrderAccount, Side, Ticker, Broker) %>%
  summarise(ExecLastFillPx = weighted.mean(ExecLastFillPx, ExecLastFill))

QuantidadeSomada <- Boleta %>%
  group_by(ExecAsofDate, OrderAccount, Side, Ticker, Broker) %>%
  summarise(ExecLastFill = sum(ExecLastFill))

BoletaMitra["ExecLastFill"] <- QuantidadeSomada$ExecLastFill

remove(QuantidadeSomada)

BoletaMitra$ExecAsofDate <- format(BoletaMitra$ExecAsofDate, "%d-%m-%Y")

BoletaMitra$ExecLastFillPx <- format(round(BoletaMitra$ExecLastFillPx,10),nsmall = 2)

BoletaMitra$ExecLastFillPx <- as.numeric(BoletaMitra$ExecLastFillPx)

BoletaMitra <- BoletaMitra[c("ExecAsofDate", "OrderAccount", "Side", "Ticker", "ExecLastFill", "ExecLastFillPx", "Broker")]

Boleta$ExecAsofDate <- format(Boleta$ExecAsofDate,"%Y%m%d")

Boleta$ExecLastFill <- as.character(Boleta$ExecLastFill)

Boleta$ExecLastFillPx <- Boleta$ExecLastFillPx * 100

Boleta$ExecLastFillPx <- as.character(Boleta$ExecLastFillPx)

Boleta$Side <- ifelse(Boleta$Side == 'Compra', 'CV', 'VV')

Conta = cases( 
  "600" = Boleta$OrderAccount == "FIA IBOVESPA",
  "600" = Boleta$OrderAccount == "FIA Ibovespa",
  "615" = Boleta$OrderAccount == "FIA ATIVO",
  "615" = Boleta$OrderAccount == "FIA Ativo",
  "620" = Boleta$OrderAccount == "FIA PETROS SAL",
  "140" = Boleta$OrderAccount == "PP2 PART",
  "287222" = Boleta$OrderAccount == "PSP-R GIRO",
  "287208" = Boleta$OrderAccount == "PSP-NR GIRO",
  "287215" = Boleta$OrderAccount == "PP2 GIRO",
  "278306" = Boleta$OrderAccount == "LAN GIRO",
  "278307" = Boleta$OrderAccount == "NIT GIRO",
  "278308" = Boleta$OrderAccount == "PGA GIRO",
  "278309" = Boleta$OrderAccount == "ULT GIRO",
  "278314" = Boleta$OrderAccount == "TAP GIRO",
  "278315" = Boleta$OrderAccount == "SAN GIRO",
  "163" = Boleta$OrderAccount == "PSP-R PART",
  "162" = Boleta$OrderAccount == "PSP-NR PART",
  "122" = Boleta$OrderAccount == "LAN PART", 
  "123" = Boleta$OrderAccount == "NIT PART",
  "59" = Boleta$OrderAccount == "PGA PART",
  "121" = Boleta$OrderAccount == "ULT PART"
  
  )

Boleta$OrderAccount <- Conta

Corretora = cases(
  "127" = Boleta$Broker == "TPBR",
  "3" = Boleta$Broker == "XPBI",
  "45" = Boleta$Broker == "FB",
  "45" = Boleta$Broker == "AES",
  "77" = Boleta$Broker == "CBRZ",
  "40" = Boleta$Broker == "MS",
  "16" = Boleta$Broker == "JP",
  "114" = Boleta$Broker == "ITAA",
  "85" = Boleta$Broker == "BTGI",
  "72" = Boleta$Broker == "BDSG",
  "72" = Boleta$Broker == "BDMA",
  "13" = Boleta$Broker == "ML",
  "59" = Boleta$Broker == "SAFR",
  "8" = Boleta$Broker == "LINK"
)

Boleta$Broker <- Corretora

Boleta["Parametro"] <- "N"

write.fwf(Boleta, file= Sqia.output.path, width=rep(20,ncol(Boleta)), sep="", colnames=FALSE, justify = "left")



'Boleta MITRA - Pre?o M?dio'

BoletaMitraFills <- BoletaMitra

BoletaMitraFills["Grupo de Fundos"] <- " "
BoletaMitraFills["PL do Fundo"] <- " "
BoletaMitraFills["Valor"] <- " "
BoletaMitraFills["Al?ada"] <- "DISCRICIONARIA"
BoletaMitraFills["Observa??es"] <- " "
BoletaMitraFills["Impacta Margem?"] <- "N?o"
BoletaMitraFills["Pr?-Boleta Origem"] <- " "

BoletaMitraFills <- BoletaMitraFills[c("Grupo de Fundos", "OrderAccount", "Side", "Ticker", "ExecLastFill", "PL do Fundo", "ExecLastFillPx", "Valor", "Broker", "Al?ada", "Observa??es","Impacta Margem?", "Pr?-Boleta Origem")]

ContaMitra = cases( 
  "[FRONT] FP IBOVESPA FUNDO DE INVESTIMENTO EM ACOES" = BoletaMitraFills$OrderAccount == "FIA IBOVESPA",
  "[FRONT] FP IBOVESPA FUNDO DE INVESTIMENTO EM ACOES" = BoletaMitraFills$OrderAccount == "FIA Ibovespa",
  "[FRONT] FIA PETROS ATIVO" = BoletaMitraFills$OrderAccount == "FIA ATIVO",
  "[FRONT] FIA PETROS ATIVO" = BoletaMitraFills$OrderAccount == "FIA Ativo",
  "[FRONT] FIA PETROS SELECAO ALTA LIQUIDEZ" = BoletaMitraFills$OrderAccount == "FIA PETROS SAL",
  "[FRONT] RV.PART.PP2" = BoletaMitraFills$OrderAccount == "PP2 PART",
  "[FRONT] RV.GIRO.SPR" = BoletaMitraFills$OrderAccount == "PSP-R GIRO",
  "[FRONT] RV.GIRO.SPN" = BoletaMitraFills$OrderAccount == "PSP-NR GIRO",
  "[FRONT] RV.GIRO.PP2" = BoletaMitraFills$OrderAccount == "PP2 GIRO",
  "[FRONT] RV.GIRO.LAN" = BoletaMitraFills$OrderAccount == "LAN GIRO",
  "[FRONT] RV.GIRO.NIT" = BoletaMitraFills$OrderAccount == "NIT GIRO",
  "[FRONT] RV.GIRO.PGA" = BoletaMitraFills$OrderAccount == "PGA GIRO",
  "[FRONT] RV.GIRO.ULT" = BoletaMitraFills$OrderAccount == "ULT GIRO",
  "[FRONT] RV.GIRO.TAP" = BoletaMitraFills$OrderAccount == "TAP GIRO",
  "[FRONT] RV.GIRO.SAN" = BoletaMitraFills$OrderAccount == "SAN GIRO",
  "[FRONT] RV.PART.SPR" = BoletaMitraFills$OrderAccount == "PSP-R PART",
  "[FRONT] RV.PART.SPN" = BoletaMitraFills$OrderAccount == "PSP-NR PART",
  "[FRONT] RV.PART.LAN" = BoletaMitraFills$OrderAccount == "LAN PART",
  "[FRONT] RV.PART.NIT" = BoletaMitraFills$OrderAccount == "NIT PART",
  "[FRONT] RV.PART.PGA" = BoletaMitraFills$OrderAccount == "PGA PART",
  "[FRONT] RV.PART.ULT" = BoletaMitraFills$OrderAccount == "ULT PART"
  
)

BoletaMitraFills$OrderAccount <- ContaMitra

CorretoraMitra = cases( 
  
  "TULL PRE" = BoletaMitraFills$Broker == "TPBR",
  "XP_INV" = BoletaMitraFills$Broker == "XPBI",
  "C SUISSE" = BoletaMitraFills$Broker == "FB",
  "C SUISSE" = BoletaMitraFills$Broker == "AES",
  "CITIBANK" = BoletaMitraFills$Broker == "CBRZ",
  "MORGAN S" = BoletaMitraFills$Broker == "MS",
  "JP MORGAN" = BoletaMitraFills$Broker == "JP",
  "ITAU" = BoletaMitraFills$Broker == "ITAA",
  "PACTUAL" = BoletaMitraFills$Broker == "BTGI",
  "BRADESCO" = BoletaMitraFills$Broker == "BDSG",
  "BRADESCO" = BoletaMitraFills$Broker == "BDMA",
  "MERRIL" = BoletaMitraFills$Broker == "ML",
  "SAFRA" = BoletaMitraFills$Broker == "SAFR",
  "UBS" = BoletaMitraFills$Broker == "LINK"
  
)

BoletaMitraFills$Broker <- CorretoraMitra

file_name <- paste0(HistoricalOp.output.path, format(Sys.time(), "%Y-%m-%d"), ".xlsx")

writexl::write_xlsx(BoletaMitra, file_name)

writexl::write_xlsx(BoletaMitraFills, Mitra.output.path)