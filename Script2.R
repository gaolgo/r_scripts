#Para manipular os dados
library(dplyr)
library("stringr")
#Para lidar com o PNAD
library(PNADcIBGE)
library("pivottabler")
#Para exportar para o Excel
library(openxlsx)

generate_pnad_sheet_1<-function(year, quarter){
  
  #Importando o  PNAD ( por algum motivo não consigo puxar apenas as variávies selecionadas)###
  var<-c("UF", "RM_RIDE","V1028", "V2007","V2009","V2010","V3003A","V3009A",
         "V4010","V4012","V4013","V403311","V403312","VD2003","VD3004","VD3005",
         "VD4002","VD4010","VD4012","VD4016","VD4020")
  
  pnad<- get_pnadc(year, quarter = quarter, design=FALSE)
  
  pnad<-pnad%>%select(UF, RM_RIDE, V1028,V2007,V2009,V2010,V3003A,V3009A,
                      V4010,V4012,V4013,V403311,V403312,VD2003,VD3004,VD3005,
                      VD4002,VD4010,VD4012,VD4016,VD4020)
  #Nome das colunas
  colnames(pnad)<-c("UF","RM_RIDE","Peso","Sexo", "Idade", "Cor ou Raça", "Curso que frequenta", 
                    "Curso mais elevado que frequentou anteriormente?", "Cargo", 
                    "Vinculo", "Atividade", "Faixa rendimento", "Rendimento (valor)", "Pessoas no domicilio",
                    "Nível de instrução mais elevado alcançado (pessoas de 5 anos ou mais de idade) padronizado para o Ensino fundamental -  SISTEMA DE 9 ANOS",
                    "Anos de estudo", "Condição na ocupação", "Atividade Principal", "Contribuição Previdência", 
                    "Rendimento habitualmente", "Rendimento")
  
  ##Convertendo o CNAE para character
  pnad$Atividade <- as.character(as.numeric(pnad$Atividade))
  pnad$Cargo <- as.character(as.numeric(pnad$Cargo))
  pnad$Rendimento<- as.numeric(as.character(pnad$Rendimento))
  
  #Adicionando Região
  pnad<-pnad%>% mutate(Regiao = case_when(UF == "Amazonas" ~ "Norte",UF == "Acre" ~ "Norte", UF =="Amapá"~"Norte", UF== "Rondônia"~"Norte", UF=="Pará"~"Norte",UF=="Roraima"~"Norte", UF=="Tocantins"~"Norte",
                                          UF == "Pernambuco" ~ "Nordeste",UF=="Alagoas"~"Nordeste", UF=="Bahia"~"Nordeste", UF=="Ceará"~"Nordeste", UF=="Maranhão"~"Nordeste",UF=="Piauí"~"Nordeste", UF=="Paraíba"~"Nordeste", UF=="Rio Grande do Norte"~"Nordeste", UF=="Sergipe"~"Nordeste",
                                          UF == "Paraná" ~ "Sul",  UF == "Rio Grande do Sul" ~ "Sul", UF == "Santa Catarina" ~ "Sul",
                                          UF == 'Goiás' ~"Centro-Oeste",  UF == 'Mato Grosso' ~"Centro-Oeste",  UF == 'Mato Grosso do Sul' ~"Centro-Oeste",  UF == 'Distrito Federal' ~"Centro-Oeste", 
                                          UF == "São Paulo" ~ "Sudeste",UF == "Rio de Janeiro" ~ "Sudeste",UF == "Espírito Santo" ~ "Sudeste",UF == "Minas Gerais" ~ "Sudeste"))
  #Criando faixas-etárias
  pnad<-pnad%>% mutate(Idade = case_when(Idade> 14 & Idade <= 17 ~ "14 a 17 anos", 
                                         Idade> 17 & Idade <= 24 ~ "18 a 24 anos", 
                                         Idade> 25 & Idade <= 39 ~ "25 a 39 anos", 
                                         Idade> 40 & Idade <= 59 ~ "40 a 59 anos", 
                                         Idade >=60 ~ "60 anos ou mais"
  ))
  
  #Renomeando Atividades
  pnad<-pnad%>% mutate(Atividade = case_when(
    Atividade == "62000"~"Atividades dos Serviços de TI",
    Atividade == "63000"~"Atividades de Prestação de Serviços de TI",
    Atividade != "62000" ~"Demais Setores",
    Atividade != "63000" ~ "Demais Setores", 
  ))
  
  #Selecionando Vínculos(?) 
  pnad<-pnad%>% mutate(Vinculo = case_when(
    Vinculo=='Empregado do setor privado'~"Empregado do setor privado",
    Vinculo== 'Empregado do setor público (inclusive empresas de economia mista)'~"Empregado do setor público",
    Vinculo== 'Empregador' ~"Empregador", 
    Vinculo=='Conta própria'~"Conta própria",
    Vinculo=='Trabalhador familiar não remunerado'~" Trabalhador familiar não remunerado",
    Vinculo!='Empregado do setor privado'~"Demais",
    Vinculo!= 'Empregado do setor público (inclusive empresas de economia mista)'~"Demais",
    Vinculo!= 'Empregador' ~"Demais", 
    Vinculo!='Conta própria'~"Demais",
    Vinculo!='Trabalhador familiar não remunerado'~"Demais"
  )
  )
  #Renomeando Cargos
  pnad<-pnad%>% mutate(Cargo = case_when(
    Cargo==2511 ~"Desenvolvedores e analistas",
    Cargo==2512~ "Desenvolvedores e analistas",
    Cargo==2513~ "Desenvolvedores e analistas",
    Cargo==2514~ "Desenvolvedores e analistas",
    Cargo==2519~ 'Especialistas em base de dados e em redes',
    Cargo==2521~ 'Especialistas em base de dados e em redes',
    Cargo==2522~ 'Especialistas em base de dados e em redes',
    Cargo==2523~ 'Especialistas em base de dados e em redes',
    Cargo!=2511 ~"Demais",
    Cargo!=2512~ "Demais",
    Cargo!=2513~ "Demais",
    Cargo!=2514~ "Demais",
    Cargo!=2519~ "Demais",
    Cargo!=2521~ "Demais",
    Cargo!=2523~ "Demais",
  ))
  
  # Dando nome para as duas Atividades relevantes para a base 
  pnad_TI <-pnad %>% 
    filter(str_detect(Atividade, "Atividades dos Serviços de TI|Atividades de Prestação de Serviços de TI"))
  
  ##Criando um binding  para pode apresentar o total no Excel
  
  pnad_TI2<-pnad_TI%>% mutate(Atividade = case_when(
    Atividade=="Atividades dos Servi?os de TI"~"Total",
    Atividade=="Atividades de Presta??o de Servi?os de TI"~"Total",
  )
  )
  pnad_Ativ<-rbind(pnad_TI,pnad_TI2)
  
  pnad_2<-pnad%>% mutate(Regiao = case_when(
    Regiao!="Código" ~"Brasil",
  )
  )
  pnad_Reg<-rbind(pnad, pnad_2)
  #Criando Workbook
  wb <- createWorkbook(creator = Sys.getenv("Gabriel"))
  #V?nculo_Setor
  vs <- PivotTable$new()
  vs$addData(pnad_Ativ)
  vs$addColumnDataGroups("Atividade",
                         dataSortOrder="custom", 
                         customSortOrder=c("Atividades dos Serviços de TI","Atividades de Prestação de Serviços de TI","Demais Setores")
  )
  
  vs$addColumnDataGroups("Regiao",
                         dataSortOrder="custom", 
                         customSortOrder=c("Norte","Nordeste","Sudeste","Sul","Centro-Oeste") 
  )
  
  vs$addRowDataGroups("Vinculo",
                      dataSortOrder="custom", 
                      customSortOrder=c("Empregado do setor privado","Empregado do setor público","Empregador","Conta própria","Trabalhador familiar não remunerado", "Demais")
  )
  vs$defineCalculation(calculationName="Total", summariseExpression="sum(Peso, na.rm=TRUE)")
  vs$evaluatePivot()
  
  addWorksheet(wb, "Vinculo_Setor")
  vs$writeToExcelWorksheet(wb=wb, wsName="Vinculo_Setor", 
                           topRowNumber=1, leftMostColumnNumber=1, applyStyles=FALSE, exportOptions=list(skipNA=TRUE, skipNaN=TRUE))
  saveWorkbook(wb, file=sprintf("PNAD_%s_%s.xlsx",year,quartr), overwrite = TRUE)
  
  #Renda
  
  #Vinc_Rend_Setor
  
  vrs <- PivotTable$new()
  vrs$addData(pnad_Ativ)
  vrs$addColumnDataGroups("Atividade",
                          dataSortOrder="custom", 
                          customSortOrder=c("Atividades dos Serviços de TI","Atividades de Prestação de Serviços de TI","Demais Setores")
  )
  
  vrs$addColumnDataGroups("Regiao",
                          dataSortOrder="custom", 
                          customSortOrder=c("Norte","Nordeste","Sudeste","Sul","Centro-Oeste") 
  )
  
  vrs$addRowDataGroups("Vinculo",
                       dataSortOrder="custom", 
                       customSortOrder=c("Empregado do setor privado","Empregado do setor público","Empregador","Conta própria","Trabalhador familiar não remunerado", "Demais")
  )
  
  vrs$defineCalculation(calculationName="Total", summariseExpression="weighted.mean(Rendimento,Peso, na.rm=TRUE)")
  
  vrs$evaluatePivot()
  addWorksheet(wb, "Vinc_Rend_Setor")
  vrs$writeToExcelWorksheet(wb=wb, wsName="Vinc_Rend_Setor", 
                            topRowNumber=1, leftMostColumnNumber=1, applyStyles=FALSE, exportOptions=list(skipNA=TRUE, skipNaN=TRUE))
  saveWorkbook(wb, file=sprintf("PNAD_%s_%s.xlsx"), overwrite = TRUE)
  
  
  #Ocup_Atividade
  
  oa<- PivotTable$new()
  oa$addData(pnad_Reg)
  
  oa$addColumnDataGroups("Regiao",
                         dataSortOrder="custom", 
                         customSortOrder=c("Norte","Nordeste","Sudeste","Sul","Centro-Oeste") 
  )
  
  oa$addColumnDataGroups("Cargo", 
                         dataSortOrder="custom", 
                         customSortOrder=c("Desenvolvedores e Analistas","Especialistas em bases de dados e redes", "Demais"), 
  )
  
  oa$addRowDataGroups("Atividade",
                      dataSortOrder="custom", 
                      customSortOrder=c("Atividades dos Serviços de TI","Atividades de Prestação de Serviços de TI","Demais Setores")
  )
  oa$defineCalculation(calculationName="Total", summariseExpression="sum(Peso, na.rm=TRUE)")
  
  oa$evaluatePivot()
  addWorksheet(wb, "Ocup_Atividade")
  oa$writeToExcelWorksheet(wb=wb, wsName="Ocup_Atividade", 
                           topRowNumber=1, leftMostColumnNumber=1, applyStyles=FALSE, exportOptions=list(skipNA=TRUE, skipNaN=TRUE))
  saveWorkbook(wb, file=sprintf("PNAD_%s_%s.xlsx"), overwrite = TRUE)
  
  #Renda_ativi
  
  ra <- PivotTable$new()
  ra$addData(pnad_Reg)
  
  ra$addColumnDataGroups("Regiao",
                         dataSortOrder="custom", 
                         customSortOrder=c("Norte","Nordeste","Sudeste","Sul","Centro-Oeste", 
                                           addTotal=TRUE) 
  )
  
  ra$addColumnDataGroups("Cargo", 
                         dataSortOrder="custom", 
                         customSortOrder=c("Desenvolvedores e Analistas","Especialistas em bases de dados e redes", "Demais",
                                           addTotal=TRUE),
  )
  
  ra$addRowDataGroups("Atividade",
                      dataSortOrder="custom", 
                      customSortOrder=c("Atividades dos Serviços de TI","Atividades de Prestação de Serviços de TI","Demais Setores", 
                                        addTotal=TRUE)
  )
  ra$defineCalculation(calculationName="Total", summariseExpression="weighted.mean(Rendimento,Peso, na.rm=TRUE)")
  
  ra$evaluatePivot()
  addWorksheet(wb, "Renda_ativi")
  ra$writeToExcelWorksheet(wb=wb, wsName="Renda_ativi", 
                           topRowNumber=1, leftMostColumnNumber=1, applyStyles=FALSE, exportOptions=list(skipNA=TRUE, skipNaN=TRUE))
  saveWorkbook(wb, file=sprintf("PNAD_%s_%s.xlsx"), overwrite = TRUE)
  print("oi")
}

generate_pnad_sheet_1(2022,1)

generate_pnad_sheet_1(2021,4)
generate_pnad_sheet_1(2021,3)
generate_pnad_sheet_1(2021,2)
generate_pnad_sheet_1(2021,1)

generate_pnad_sheet_1(2020,4)
generate_pnad_sheet_1(2020,3)
generate_pnad_sheet_1(2020,2)
generate_pnad_sheet_1(2020,1)

generate_pnad_sheet_1(2019,4)
generate_pnad_sheet_1(2019,3)
generate_pnad_sheet_1(2019,2)
generate_pnad_sheet_1(2019,1)

generate_pnad_sheet_1(2018,4)
generate_pnad_sheet_1(2018,3)
generate_pnad_sheet_1(2018,2)
generate_pnad_sheet_1(2018,1)
