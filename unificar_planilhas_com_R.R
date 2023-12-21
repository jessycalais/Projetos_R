# UNIFICAR PLANILHAS (COM COLUNAS IGUAIS) VIA R

# Registrar hora de início
startTime <- Sys.time()

# Carregar pacotes
library(tidyverse)
library(readxl)
library(writexl)

# Unificar planilhas  
arquivos = list.files('ss', pattern='.xlsx', full.names=TRUE)
dados_unificados <- data.frame()

for (arquivo in arquivos) {
  df <- readxl::read_xlsx(arquivo)
  arquivo_de_origem <- c(rep(str_sub(arquivo, start=4, end=-6), each=dim(df)[1]))
  df['Arquivo de Origem'] <- arquivo_de_origem
  dados_unificados <- rbind(dados_unificados, df)
}

# Exportar arquivo - movendo "última coluna" para a 1a posição
DATA <- format(Sys.Date(), '%d-%m-%Y')
NOME_ARQUIVO <- paste0('planilhas_unificadas_com_R_', DATA, '.xlsx')

dados_unificados %>% 
  dplyr::relocate('Arquivo de Origem') %>% 
  writexl::write_xlsx(NOME_ARQUIVO)

# Total de arquivos unificados
print(paste0('Há ', length(arquivos), ' arquivos no formato ".xlsx" na pasta'))

# Verificar tempo gasto
endTime <- Sys.time()

print(endTime - startTime)
