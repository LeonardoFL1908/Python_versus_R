##### FEITO POR LEONARDO FARIA LIMA
##### 24\08/2024
##### PRECIFICADOR L2L

library(readxl) ### para ler arquivo em excel 
library(dplyr)
library(writexl) ### para salvar o arquvio em excel 

Sys.setlocale("LC_ALL", "pt_BR.UTF-8")

#### FAZER O INPUT DO NOME DO ARQUVIO
arquivo <- 'calculo_preco_OBMONOGENICA'

data_hora <- Sys.time()
data_hora <- format(data_hora, format="%Y-%m-%d_%Hh-%Mm")
data_hora <-  as.character(data_hora)

print(data_hora)
class(data_hora)


##############################################################################
### CARREGA O ARQUIVO 

caminho_arquivo <- sprintf("I:/Ger_Suporte a Cliente/Inteligencia de Mercado/005_Analytics/002_Bases/021_L2L/03_bases_e_template_lis/%s.xlsx", arquivo)

df <- read_excel(caminho_arquivo)

###############################################################################

#View(df)

### IMPOSTO 
tributo = 6.65/100

df <- df %>%  mutate(custo_capt_l2l = custo_total*(10/100))

###A FUNCAO MUTATE (PACOTE DPLYR)  É USADA PARA CRIAR NOVAS COLUNAS OU MODIFICAR COLUNAS EXISTENTES EM UM DATA FRAME. 
### %>% (PIPE) - ATALHO DE TECLADO CTRL + SHIFT + M.


#### IMPORTANTE: O CUSTO DE CAPTACAO L2L NAO PODE EXCEDER R$ 50,00.

df <- df %>% 
  mutate(custo_capt_l2l = ifelse(custo_capt_l2l>50, 50, custo_capt_l2l))

### CALCULA O CUSTO L2L
df <- df %>%  
  mutate(custo_l2l = custo_total+custo_capt_l2l)

### CALCULA LABOR3, PRECO MINIMO E MAXIMAO 

df <- df %>%
  mutate(
    labor3 = custo_l2l / ((1 - (mg_labor3 / 100)) * (1 - tributo)),
    minimo = custo_l2l / ((1 - (mg_minimo / 100)) * (1 - tributo)),
    maximo = labor3*5
)
         
### CALCULA MARGEM LABOR3 E PRECO MINIMO 

df <- df %>%
  mutate(
    `mg_min%` = (((minimo * (1 - tributo)) - custo_l2l) / (minimo * (1 - tributo))) * 100,
    `mg_labor3_custo%` = (((labor3 * (1 - tributo)) - custo_l2l) / (labor3 * (1 - tributo))) * 100,
    `mg_labor3_min%` = (((labor3 * (1 - tributo)) - minimo) / (labor3 * (1 - tributo))) * 100
  )

#### CALCULO MARGEM LABOR3 E MN (REBECA)
df <- df %>%
  mutate(
    `mg_min_r%` = ((1-tributo)-(custo_l2l/minimo))*100,
    `mg_labor3_r%` = ((1-tributo)-(custo_l2l/labor3))*100
  )

### MANTEM O DF ORIGINAL 

df_original <- df 

### EXCLUIR COLUNAS - select() PACOTE (dplyr)
df <- df %>% select(-c(mg_labor3, mg_minimo))


### MOSTRA O NOME DAS COLUNAS 
print('Nomes Colunas: ')
names(df)

### ARRENDONDA AS CASAS DECIMAIS 

df <- df %>%
  mutate(
    custo_total = round(custo_total, 2),
    custo_capt_l2l = round(custo_capt_l2l, 2),
    custo_l2l = round(custo_l2l, 2),
    minimo = round(minimo, 2),
    maximo = round(maximo, 2),
    labor3 = round(labor3, 2),
    maximo = round(maximo, 2),
    `mg_labor3_custo%` = round(`mg_labor3_custo%`, 2),
    `mg_labor3_min%` = round(`mg_labor3_min%`, 2),
    `mg_min_r%` = round(`mg_min_r%`, 2),
    `mg_labor3_r%` = round(`mg_labor3_r%`, 2)
  )

### REORGANIZANDO AS COLUNAS ### LIBRARY (dplyr)
df <- df %>% select(
  exame, 
  gprod,
  custo_total,
  custo_capt_l2l,
  custo_l2l,
  minimo,
  labor3,
  maximo,
  `mg_min%`,
  `mg_labor3_custo%`,
  `mg_labor3_min%`,
  `mg_min_r%`,
  `mg_labor3_r%`
  )  
### GERA O ARQUIVO DO ENVIO DE PRECO 

df_envio <- df %>% select(
  exame,	
  custo_total,	
  custo_l2l,	
  minimo,	
  maximo,	
  labor3,	
  gprod
)

#### CRIA UMA VARIAVEL COM O NOME DO EXAME (COL.1 | LINHA.1)
nome_arquivo <- df[1, "exame"]
# Substituindo '||' por '_'
nome_arquivo <- gsub("\\|\\|", "_", nome_arquivo)
print(nome_arquivo)

#### CAMINHOS PARA SALVAR OS ARQUVIVOS 
### A FUNCAO <<SPRINTF>> E USADA PARA FORMATAR STRINGS. SEMELHANTE A FSTRING DO PYTHON. 
### O %s É UM MARCADOR DE POSICAO PARA UMA STRING.

salva_preco <- sprintf(
  "i:/Ger_Suporte a Cliente/Inteligencia de Mercado/005_Analytics/002_Bases/021_L2L/05_resultado_calculo_preco/preco_%s_%s.xlsx", 
  nome_arquivo, data_hora
)
print(salva_preco)

write_xlsx(df, salva_preco)

salva_envio <- sprintf(
  "I:/Ger_Suporte a Cliente/Inteligencia de Mercado/005_Analytics/002_Bases/021_L2L/03_bases_e_template_lis/envio_precos_%s_%s.xlsx",
  nome_arquivo, data_hora
)

print(salva_envio)

write_xlsx(df, salva_envio)
