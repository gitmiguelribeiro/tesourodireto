# Sobre os títulos públicos -----------------------------------------------

#Página que explica o cálculo de rentabilidade:
browseURL("https://www.tesourodireto.com.br/titulos/tipos-de-tesouro.htm")

#Página fonte dos históricos:
browseURL("https://www.tesourodireto.com.br/titulos/historico-de-precos-e-taxas.htm")

# Bibliotecas -------------------------------------------------------------

pacotes <- 
  c("tidyverse","readxl","RSelenium","kableExtra","janitor")

if(sum(as.numeric(!pacotes %in% installed.packages())) != 0){
  instalador <- pacotes[!pacotes %in% installed.packages()]
  for(i in 1:length(instalador)) {
    install.packages(instalador, dependencies = T)
    break()}
  sapply(pacotes, require, character = T) 
} else {
  sapply(pacotes, require, character = T) 
}

rm(pacotes)

# Funções -----------------------------------------------------------------

ver <- 
  function(x){
    x %>% 
      kableExtra::kable() %>%
      kableExtra::kable_styling(bootstrap_options = c("striped", "hover", "condensed", "responsive"), 
                                full_width = F, 
                                fixed_thead = T,
                                font_size = 12)
  }


ler_excel_td <- 
  function(x){
    
    arquivo <- data.frame()
    
    num_plan <- 
      readxl::excel_sheets(
        path = paste0(
          pasta,
          x)) %>% 
      length(.) %>% 
      as.numeric(.)
    
    for (i in 1:num_plan){
      nome_plan <- 
        readxl::excel_sheets(path = paste0(pasta,x))[i]
      
      tipo <- 
        stringr::str_sub(
          nome_plan,
          1,
          stringr::str_locate(nome_plan,"[0-9]")[1,1]-2
        )
      
      venc <- 
        stringr::str_sub(
          nome_plan,
          stringr::str_locate(nome_plan,"[0-9]")[1,1],
          stringr::str_length(nome_plan)
        ) %>% 
        lubridate::dmy(.)
      
      planilha <- 
        readxl::read_excel(
          path = paste0(pasta,x),
          sheet = i,
          skip=1,
          n_max=260
        ) %>% 
        janitor::clean_names(.) %>% 
        dplyr::select(1:6) %>% 
        purrr::set_names("dt_dia","taxa_compra","taxa_venda",
                         "pu_compra","pu_venda","pu_base") %>%
        dplyr::mutate(
          titulo = 
            tipo %>% 
            stringr::str_to_upper(.),
          dt_vencimento = 
            venc,
          dt_dia =
            dt_dia %>% 
            lubridate::ymd(.),
          dias_ate_vencimento =
            lubridate::interval(dt_dia,dt_vencimento) %/% lubridate::days(1),
          anos_ate_vencimento =
            round(dias_ate_vencimento/365,digits = 1)) %>% 
        dplyr::filter(!(is.na(taxa_compra)&is.na(taxa_venda)&is.na(pu_compra)
                        &is.na(pu_venda)&is.na(pu_base))) %>% 
        dplyr::mutate(
          across(starts_with("taxa_")|starts_with("pu_"),
                 ~as.numeric(.)),
          titulo =
            titulo %>% 
            dplyr::recode(
              "LTN" = "PREFIXADO",
              "NTNF" = "PREFIXADO JS",
              "NTN-F" = "PREFIXADO JS",
              "NTNB PRINCIPAL" = "IPCA+",
              "NTN-B PRINCIPAL" = "IPCA+",
              "NTN-B PRINC" = "IPCA+",
              "NTNB PRINC" = "IPCA+",
              "NTNB" = "IPCA+ JS",
              "NTN-B" = "IPCA+ JS",
              "LFT" = "SELIC",
              "NTNC" = "IGPM+ JS",
              "NTN-C" = "IGPM+ JS"
            ),
          nome = 
            paste0("TESOURO ",titulo," ",dt_vencimento %>% lubridate::year(.) %>% as.character(.))) %>% 
        dplyr::relocate(nome,titulo,.before = dt_dia)

      arquivo = 
        
        dplyr::bind_rows(arquivo,planilha,.id=NULL)
    }
    
    return(
      arquivo
    )
  }

# Lista os anos a serem buscados --------------------------------------------------------

lista_anos <- #O site do tesouro fornece os dados somente a partir de 2003 
  seq(
    2003,
    lubridate::year(lubridate::today()),
    1) %>% 
  as.vector(.)

# Define o servidor do RSelenium e salva os arquivos em Download ----------

driver <- 
  RSelenium::rsDriver(
    browser = "chrome",
    chromever = "97.0.4692.36"
  )

naveg <- 
  driver$client

servidor <- 
  driver$server

for (i in lista_anos){
  for (j in 1:6){
    url <- 
      paste0(
        "https://apiapex.tesouro.gov.br/aria/v1/sistd/custom/historico?ano=",
        i,
        "&idSigla=",
        j)
    naveg$navigate(
      url)
  }
}

naveg$close()

servidor$stop()

rm(driver,i,j,lista_anos,naveg,servidor,url)
# Uma observação: ainda não consegui fazer com que o RSelenium salve os arquivos
# na pasta desejada. Então manualmente, temos que copiar e colar os arquivos lá.

# Lê, trata e une os arquivos  na pasta ----------------------------------------------------------

pasta <- 
  # Pasta onde colei os arquivos
  paste0(getwd(),"/estatisticas/")

lista_arquivos <- 
  list.files(
    path=pasta,
    pattern="*.xls"
  ) 

base_tesouro_direto <- 
  do.call(
    rbind,
    lapply(lista_arquivos,ler_excel_td)
  ) %>% 
  tibble::as_tibble() %>% 
  dplyr::mutate(titulo = 
                  titulo %>% 
                  forcats::as_factor(.),
                nome = 
                  nome %>% 
                  forcats::as_factor(.))

# Salva os arquivos -------------------------------------------------------

write.table(base_tesouro_direto,
            file = paste0(getwd(),"/base_tesouro_direto.csv"),
            append=F,
            eol = "\r\n",
            sep = ",")

save(base_tesouro_direto,
     file = paste0(getwd(),"/base_tesouro_direto.RData"))

# Limpa a memoria ---------------------------------------------------------
rm(list=ls())
gc()

