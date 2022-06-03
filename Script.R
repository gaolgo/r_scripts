#Para manipular os dados
library(dplyr)
library("stringr")
#Para lidar com o PNAD
library(PNADcIBGE)
library("pivottabler")
#Para exportar para o Excel
library(openxlsx)

generate_pnad_sheet<-function(year, quarter){
  
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
  # Dando nome para as duas ocupações relevantes para a base
  
  ##Converter o nome do dataframe para ocupar menos espaço
  
  #Dicionário por força bruta (Muito Bruto)
  pnad<-pnad%>% mutate(Cargo = case_when(
    Cargo == 1111~ "Legisladores",
    Cargo == 1112~ "Dirigentes superiores da administração pública",
    Cargo == 1113~ "Chefes de pequenas populações",
    Cargo == 1114~ "Dirigentes de organizações que apresentam um interesse especial",
    Cargo == 1120~ "Diretores gerais e gerentes gerais",
    Cargo == 1211~ "Dirigentes financeiros",
    Cargo == 1212~ "Dirigentes de recursos humanos",
    Cargo == 1213~ "Dirigentes de políticas e planejamento",
    Cargo == 1219~ "Dirigentes de administração e de serviços não classificados anteriormente",
    Cargo == 1221~ "Dirigentes de vendas e comercialização",
    Cargo == 1222~ "Dirigentes de publicidade e relações públicas",
    Cargo == 1223~ "Dirigentes de pesquisa e desenvolvimento",
    Cargo == 1311~ "Dirigentes de produção agropecuária e silvicultura",
    Cargo == 1312~ "Dirigentes de produção da aquicultura e pesca",
    Cargo == 1321~ "Dirigentes de indústria de transformação",
    Cargo == 1322~ "Dirigentes de explorações de mineração",
    Cargo == 1323~ "Dirigentes de empresas de construção",
    Cargo == 1324~ "Dirigentes de empresas de abastecimento, distribuição e afins",
    Cargo == 1330~ "Dirigentes de serviços de tecnologia da informação e comunicações",
    Cargo == 1341~ "Dirigentes de serviços de cuidados infantis",
    Cargo == 1342~ "Dirigentes de serviços de saúde",
    Cargo == 1343~ "Dirigentes de serviços de cuidado a pessoas idosas",
    Cargo == 1344~ "Dirigentes de serviços de bem-estar social",
    Cargo == 1345~ "Dirigentes de serviços de educação",
    Cargo == 1346~ "Gerentes de sucursais de bancos, de serviços financeiros e de seguros",
    Cargo == 1349~ "Dirigentes e gerentes de serviços profissionais não classificados anteriormente",
    Cargo == 1411~ "Gerentes de hotéis",
    Cargo == 1412~ "Gerentes de restaurantes",
    Cargo == 1420~ "Gerentes de comércios atacadistas e varejistas",
    Cargo == 1431~ "Gerentes de centros esportivos, de diversão e culturais",
    Cargo == 1439~ "Gerentes de serviços não classificados anteriormente",
    Cargo == 2111~ "Físicos e astrônomos",
    Cargo == 2112~ "Meteorologistas",
    Cargo == 2113~ "Químicos",
    Cargo == 2114~ "Geólogos e geofísicos",
    Cargo == 2120~ "Matemáticos, atuários e estatísticos",
    Cargo == 2131~ "Biólogos, botânicos, zoólogos e afins",
    Cargo == 2132~ "Agrônomos e afins",
    Cargo == 2133~ "Profissionais da proteção do meio ambiente",
    Cargo == 2141~ "Engenheiros industriais e de produção",
    Cargo == 2142~ "Engenheiros civis",
    Cargo == 2143~ "Engenheiros de meio ambiente",
    Cargo == 2144~ "Engenheiros mecânicos",
    Cargo == 2145~ "Engenheiros químicos",
    Cargo == 2146~ "Engenheiros de minas, metalúrgicos e afins",
    Cargo == 2149~ "Engenheiros não classificados anteriormente",
    Cargo == 2151~ "Engenheiros eletricistas",
    Cargo == 2152~ "Engenheiros eletrônicos",
    Cargo == 2153~ "Engenheiros em telecomunicações",
    Cargo == 2161~ "Arquitetos de edificações",
    Cargo == 2162~ "Arquitetos paisagistas",
    Cargo == 2163~ "Desenhistas de produtos e vestuário",
    Cargo == 2164~ "Urbanistas e engenheiros de trânsito",
    Cargo == 2165~ "Cartógrafos e agrimensores",
    Cargo == 2166~ "Desenhistas gráficos e de multimídia",
    Cargo == 2211~ "Médicos gerais",
    Cargo == 2212~ "Médicos especialistas",
    Cargo == 2221~ "Profissionais de enfermagem",
    Cargo == 2222~ "Profissionais de partos",
    Cargo == 2230~ "Profissionais da medicina tradicional e alternativa",
    Cargo == 2240~ "Paramédicos",
    Cargo == 2250~ "Veterinários",
    Cargo == 2261~ "Dentistas",
    Cargo == 2262~ "Farmacêuticos",
    Cargo == 2263~ "Profissionais da saúde e da higiene laboral e ambiental",
    Cargo == 2264~ "Fisioterapeutas",
    Cargo == 2265~ "Dietistas e nutricionistas",
    Cargo == 2266~ "Fonoaudiólogos e logopedistas",
    Cargo == 2267~ "Optometristas",
    Cargo == 2269~ "Profissionais da saúde não classificados anteriormente",
    Cargo == 2310~ "Professores de universidades e do ensino superior",
    Cargo == 2320~ "Professores de formação profissional",
    Cargo == 2330~ "Professores do ensino médio",
    Cargo == 2341~ "Professores do ensino fundamental",
    Cargo == 2342~ "Professores do ensino pré-escolar",
    Cargo == 2351~ "Especialistas em métodos pedagógicos",
    Cargo == 2352~ "Educadores para necessidades especiais",
    Cargo == 2353~ "Outros professores de idiomas",
    Cargo == 2354~ "Outros professores de música",
    Cargo == 2355~ "Outros professores de artes",
    Cargo == 2356~ "Instrutores em tecnologias da informação",
    Cargo == 2359~ "Profissionais de ensino não classificados anteriormente",
    Cargo == 2411~ "Contadores",
    Cargo == 2412~ "Assessores financeiros e em investimentos",
    Cargo == 2413~ "Analistas financeiros",
    Cargo == 2421~ "Analistas de gestão e administração",
    Cargo == 2422~ "Especialistas em políticas de administração",
    Cargo == 2423~ "Especialistas em políticas e serviços de pessoal e afins",
    Cargo == 2424~ "Especialistas em formação de pessoal",
    Cargo == 2431~ "Profissionais da publicidade e da comercialização",
    Cargo == 2432~ "Profissionais de relações públicas",
    Cargo == 2433~ "Profissionais de vendas técnicas e médicas (exclusive tic)",
    Cargo == 2434~ "Profissionais de vendas de tecnologia da informação e comunicações",
    Cargo == 2511~ "Analistas de sistemas",
    Cargo == 2512~ "Desenvolvedores de programas e aplicativos (software)",
    Cargo == 2513~ "Desenvolvedores de páginas de internet (web) e multimídia",
    Cargo == 2514~ "Programadores de aplicações",
    Cargo == 2519~ "Desenvolvedores e analistas de programas e aplicativos (software) e multimídia não classificados anteriormente",
    Cargo == 2521~ "Desenhistas e administradores de bases de dados",
    Cargo == 2522~ "Administradores de sistemas",
    Cargo == 2523~ "Profissionais em rede de computadores",
    Cargo == 2529~ "Especialistas em base de dados e em redes de computadores não classificados anteriormente",
    Cargo == 2611~ "Advogados e juristas",
    Cargo == 2612~ "Juízes",
    Cargo == 2619~ "Profissionais em direito não classificados anteriormente",
    Cargo == 2621~ "Arquivologistas e curadores de museus",
    Cargo == 2622~ "Bibliotecários, documentaristas e afins",
    Cargo == 2631~ "Economistas",
    Cargo == 2632~ "Sociólogos, antropólogos e afins",
    Cargo == 2633~ "Filósofos, historiadores e especialistas em ciência política",
    Cargo == 2634~ "Psicólogos",
    Cargo == 2635~ "Assistentes sociais",
    Cargo == 2636~ "Ministros de cultos religiosos, missionários e afins",
    Cargo == 2641~ "Escritores",
    Cargo == 2642~ "Jornalistas",
    Cargo == 2643~ "Tradutores, intérpretes e linguistas",
    Cargo == 2651~ "Artistas plásticos",
    Cargo == 2652~ "Músicos, cantores e compositores",
    Cargo == 2653~ "Bailarinos e coreógrafos",
    Cargo == 2654~ "Diretores de cinema, de teatro e afins",
    Cargo == 2655~ "Atores",
    Cargo == 2656~ "Locutores de rádio, televisão e outros meios de comunicação",
    Cargo == 2659~ "Artistas criativos e interpretativos não classificados anteriormente",
    Cargo == 3111~ "Técnicos em ciências físicas e químicas",
    Cargo == 3112~ "Técnicos em engenharia civil",
    Cargo == 3113~ "Eletrotécnicos",
    Cargo == 3114~ "Técnicos em eletrônica",
    Cargo == 3115~ "Técnicos em engenharia mecânica",
    Cargo == 3116~ "Técnicos em química industrial",
    Cargo == 3117~ "Técnicos em engenharia de minas e metalurgia",
    Cargo == 3118~ "Desenhistas e projetistas técnicos",
    Cargo == 3119~ "Técnicos em ciências físicas e da engenharia não classificados anteriormente",
    Cargo == 3121~ "Supervisores da mineração",
    Cargo == 3122~ "Supervisores de indústrias de transformação",
    Cargo == 3123~ "Supervisores da construção",
    Cargo == 3131~ "Operadores de instalações de produção de energia",
    Cargo == 3132~ "Operadores de incineradores, instalações de tratamento de água e afins",
    Cargo == 3133~ "Controladores de instalações de processamento de produtos químicos",
    Cargo == 3134~ "Operadores de instalações de refino de petróleo e gás natural",
    Cargo == 3135~ "Controladores de processos de produção de metais",
    Cargo == 3139~ "Técnicos em controle de processos não classificados anteriormente",
    Cargo == 3141~ "Técnicos e profissionais de nível médio em ciências biológicas (exclusive da medicina)",
    Cargo == 3142~ "Técnicos agropecuários",
    Cargo == 3143~ "Técnicos florestais",
    Cargo == 3151~ "Oficiais maquinistas em navegação",
    Cargo == 3152~ "Capitães, oficiais de coberta e práticos",
    Cargo == 3153~ "Pilotos de aviação e afins",
    Cargo == 3154~ "Controladores de tráfego aéreo",
    Cargo == 3155~ "Técnicos em segurança aeronáutica",
    Cargo == 3211~ "Técnicos em aparelhos de diagnóstico e tratamento médico",
    Cargo == 3212~ "Técnicos de laboratórios médicos",
    Cargo == 3213~ "Técnicos e assistentes farmacêuticos",
    Cargo == 3214~ "Técnicos de próteses médicas e dentárias",
    Cargo == 3221~ "Profissionais de nível médio de enfermagem",
    Cargo == 3222~ "Profissionais de nível médio de partos",
    Cargo == 3230~ "Profissionais de nível médio de medicina tradicional e alternativa",
    Cargo == 3240~ "Técnicos e assistentes veterinários",
    Cargo == 3251~ "Dentistas auxiliares e ajudantes de odontologia",
    Cargo == 3252~ "Técnicos em documentação sanitária",
    Cargo == 3253~ "Trabalhadores comunitários da saúde",
    Cargo == 3254~ "Técnicos em optometria e ópticos",
    Cargo == 3255~ "Técnicos e assistentes fisioterapeutas",
    Cargo == 3256~ "assistentes de medicina",
    Cargo == 3257~ "Inspetores de saúde laboral, ambiental e afins",
    Cargo == 3258~ "Ajudantes de ambulâncias",
    Cargo == 3259~ "Profissionais de nível médio da saúde não classificados anteriormente",
    Cargo == 3311~ "Agentes e corretores de bolsa, câmbio e outros serviços financeiros",
    Cargo == 3312~ "Agentes de empréstimos e financiamento",
    Cargo == 3313~ "Contabilistas e guarda livros",
    Cargo == 3314~ "Profissionais de nível médio de serviços estatísticos, matemáticos e afins",
    Cargo == 3315~ "Avaliadores",
    Cargo == 3321~ "Agentes de seguros",
    Cargo == 3322~ "Representantes comerciais",
    Cargo == 3323~ "Agentes de compras",
    Cargo == 3324~ "Corretores de comercialização",
    Cargo == 3331~ "Despachantes aduaneiros",
    Cargo == 3332~ "Organizadores de conferências e eventos",
    Cargo == 3333~ "Agentes de emprego e agenciadores de mão de obra",
    Cargo == 3334~ "Agentes imobiliários",
    Cargo == 3339~ "Agentes de serviços comerciais não classificados anteriormente",
    Cargo == 3341~ "Supervisores de secretaria",
    Cargo == 3342~ "Secretários jurídicos",
    Cargo == 3343~ "Secretários executivos e administrativos",
    Cargo == 3344~ "Secretários de medicina",
    Cargo == 3351~ "Agentes aduaneiros e inspetores de fronteiras",
    Cargo == 3352~ "Agentes da administração tributária",
    Cargo == 3353~ "Agentes de serviços de seguridade social",
    Cargo == 3354~ "Agentes de serviços de expedição de licenças e permissões",
    Cargo == 3355~ "Inspetores de polícia e detetives",
    Cargo == 3359~ "Agentes da administração pública para aplicação da lei e afins não classificados anteriormente",
    Cargo == 3411~ "Profissionais de nível médio do direito e serviços legais e afins",
    Cargo == 3412~ "Trabalhadores e assistentes sociais de nível médio",
    Cargo == 3413~ "Auxiliares leigos de religião",
    Cargo == 3421~ "Atletas e esportistas",
    Cargo == 3422~ "Treinadores, instrutores e árbitros de atividades esportivas",
    Cargo == 3423~ "Instrutores de educação física e atividades recreativas",
    Cargo == 3431~ "Fotógrafos",
    Cargo == 3432~ "Desenhistas e decoradores de interiores",
    Cargo == 3433~ "Técnicos em galerias de arte, museus e bibliotecas",
    Cargo == 3434~ "Chefes de cozinha",
    Cargo == 3435~ "Outros profissionais de nível médio em atividades culturais e artísticas",
    Cargo == 3511~ "Técnicos em operações de tecnologia da informação e das comunicações",
    Cargo == 3512~ "Técnicos em assistência ao usuário de tecnologia da informação e das comunicações",
    Cargo == 3513~ "Técnicos de redes e sistemas de computadores",
    Cargo == 3514~ "Técnicos da web",
    Cargo == 3521~ "Técnicos de radiodifusão e gravação audiovisual",
    Cargo == 3522~ "Técnicos de engenharia de telecomunicações",
    Cargo == 4110~ "Escriturários gerais",
    Cargo == 4120~ "Secretários (geral)",
    Cargo == 4131~ "Operadores de máquinas de processamento de texto e mecanógrafos",
    Cargo == 4132~ "Operadores de entrada de dados",
    Cargo == 4211~ "Caixas de banco e afins",
    Cargo == 4212~ "Coletores de apostas e de jogos",
    Cargo == 4213~ "Trabalhadores em escritórios de empréstimos e penhor",
    Cargo == 4214~ "Cobradores e afins",
    Cargo == 4221~ "Trabalhadores de agências de viagem",
    Cargo == 4222~ "Trabalhadores de centrais de atendimento",
    Cargo == 4223~ "Telefonistas",
    Cargo == 4224~ "Recepcionistas de hotéis",
    Cargo == 4225~ "Trabalhadores dos serviços de informações",
    Cargo == 4226~ "Recepcionistas em geral",
    Cargo == 4227~ "Entrevistadores de pesquisas de mercado",
    Cargo == 4229~ "Trabalhadores de serviços de informação ao cliente não classificados anteriormente",
    Cargo == 4311~ "Trabalhadores de contabilidade e cálculo de custos",
    Cargo == 4312~ "Trabalhadores de serviços estatísticos, financeiros e de seguros",
    Cargo == 4313~ "Trabalhadores encarregados de folha de pagamento",
    Cargo == 4321~ "Trabalhadores de controle de abastecimento e estoques",
    Cargo == 4322~ "Trabalhadores de serviços de apoio à produção",
    Cargo == 4323~ "Trabalhadores de serviços de transporte",
    Cargo == 4411~ "Trabalhadores de bibliotecas",
    Cargo == 4412~ "Trabalhadores de serviços de correios",
    Cargo == 4413~ "Codificadores de dados, revisores de provas de impressão e afins",
    Cargo == 4414~ "Outros escreventes",
    Cargo == 4415~ "Trabalhadores de arquivos",
    Cargo == 4416~ "Trabalhadores do serviço de pessoal",
    Cargo == 4419~ "Trabalhadores de apoio administrativo não classificados anteriormente",
    Cargo == 5111~ "Auxiliares de serviço de bordo",
    Cargo == 5112~ "Fiscais e cobradores de transportes públicos",
    Cargo == 5113~ "Guias de turismo",
    Cargo == 5120~ "Cozinheiros",
    Cargo == 5131~ "Garçons",
    Cargo == 5132~ "atendentes de bar",
    Cargo == 5141~ "Cabeleireiros",
    Cargo == 5142~ "Especialistas em tratamento de beleza e afins",
    Cargo == 5151~ "Supervisores de manutenção e limpeza de edifícios em escritórios, hotéis e estabelecimentos",
    Cargo == 5152~ "Governantas e mordomos domésticos",
    Cargo == 5153~ "Porteiros e zeladores",
    Cargo == 5161~ "Astrólogos, adivinhos e afins",
    Cargo == 5162~ "Acompanhantes e criados particulares",
    Cargo == 5163~ "Trabalhadores de funerárias e embalsamadores",
    Cargo == 5164~ "Cuidadores de animais",
    Cargo == 5165~ "Instrutores de autoescola",
    Cargo == 5168~ "Trabalhadores do sexo",
    Cargo == 5169~ "Trabalhadores de serviços pessoais não classificados anteriormente",
    Cargo == 5211~ "Vendedores de quiosques e postos de mercados",
    Cargo == 5212~ "Vendedores ambulantes de serviços de alimentação",
    Cargo == 5221~ "Comerciantes de lojas",
    Cargo == 5222~ "Supervisores de lojas",
    Cargo == 5223~ "Balconistas e vendedores de lojas",
    Cargo == 5230~ "Caixas e expedidores de bilhetes",
    Cargo == 5241~ "Modelos de moda, arte e publicidade",
    Cargo == 5242~ "Demonstradores de lojas",
    Cargo == 5243~ "Vendedores a domicilio",
    Cargo == 5244~ "Vendedores por telefone",
    Cargo == 5245~ "Frentistas de posto de gasolina",
    Cargo == 5246~ "Balconistas dos serviços de alimentação",
    Cargo == 5249~ "Vendedores não classificados anteriormente",
    Cargo == 5311~ "Cuidadores de crianças",
    Cargo == 5312~ "Ajudantes de professores",
    Cargo == 5321~ "Trabalhadores de cuidados pessoais em instituições",
    Cargo == 5322~ "Trabalhadores de cuidados pessoais a domicílios",
    Cargo == 5329~ "Trabalhadores de cuidados pessoais nos serviços de saúde não classificados anteriormente",
    Cargo == 5411~ "Bombeiros",
    Cargo == 5412~ "Policiais",
    Cargo == 5413~ "Guardiões de presídios",
    Cargo == 5414~ "Guardas de segurança",
    Cargo == 5419~ "Trabalhadores dos serviços de proteção e segurança não classificados anteriormente",
    Cargo == 6111~ "Agricultores e trabalhadores qualificados em atividades da agricultura (exclusive hortas, viveiros e jardins)",
    Cargo == 6112~ "Agricultores e trabalhadores qualificados no cultivo de hortas, viveiros e jardins",
    Cargo == 6114~ "Agricultores e trabalhadores qualificados de cultivos mistos",
    Cargo == 6121~ "Criadores de gado e trabalhadores qualificados da criação de gado",
    Cargo == 6122~ "Avicultores e trabalhadores qualificados da avicultura",
    Cargo == 6123~ "Apicultores, sericicultores e trabalhadores qualificados da apicultura e sericicultura",
    Cargo == 6129~ "Outros criadores e trabalhadores qualificados da pecuária não classificados anteriormente",
    Cargo == 6130~ "Produtores e trabalhadores qualificados de exploração agropecuária mista",
    Cargo == 6210~ "Trabalhadores florestais qualificados e afins",
    Cargo == 6221~ "Trabalhadores da aquicultura",
    Cargo == 6224~ "Caçadores",
    Cargo == 6225~ "Pescadores",
    Cargo == 7111~ "Construtores de casas",
    Cargo == 7112~ "Pedreiros",
    Cargo == 7113~ "Canteiros, cortadores e gravadores de pedras",
    Cargo == 7114~ "Trabalhadores em cimento e concreto armado",
    Cargo == 7115~ "Carpinteiros",
    Cargo == 7119~ "Outros trabalhadores qualificados e operários da construção não classificados anteriormente",
    Cargo == 7121~ "Telhadores",
    Cargo == 7122~ "Aplicadores de revestimentos cerâmicos, pastilhas, pedras e madeiras",
    Cargo == 7123~ "Gesseiros",
    Cargo == 7124~ "Instaladores de material isolante térmico e acústico",
    Cargo == 7125~ "Vidraceiros",
    Cargo == 7126~ "Bombeiros e encanadores",
    Cargo == 7127~ "Mecânicos-instaladores de sistemas de refrigeração e climatização",
    Cargo == 7131~ "Pintores e empapeladores",
    Cargo == 7132~ "Lustradores",
    Cargo == 7133~ "Limpadores de fachadas",
    Cargo == 7211~ "Moldadores de metal e macheiros",
    Cargo == 7212~ "Soldadores e oxicortadores",
    Cargo == 7213~ "Chapistas e caldeireiros",
    Cargo == 7214~ "Montadores de estruturas metálicas",
    Cargo == 7215~ "Aparelhadores e emendadores de cabos",
    Cargo == 7221~ "Ferreiros e forjadores",
    Cargo == 7222~ "Ferramenteiros e afins",
    Cargo == 7223~ "Reguladores e operadores de máquinas-ferramentas",
    Cargo == 7224~ "Polidores de metais e afiadores de ferramentas",
    Cargo == 7231~ "Mecânicos e reparadores de veículos a motor",
    Cargo == 7232~ "Mecânicos e reparadores de motores de avião",
    Cargo == 7233~ "Mecânicos e reparadores de máquinas agrícolas e industriais",
    Cargo == 7234~ "Reparadores de bicicletas e afins",
    Cargo == 7311~ "Mecânicos e reparadores de instrumentos de precisão",
    Cargo == 7312~ "Confeccionadores e afinadores de instrumentos musicais",
    Cargo == 7313~ "Joalheiros e lapidadores de gemas, artesãos de metais preciosos e semipreciosos",
    Cargo == 7314~ "Ceramistas e afins (preparação e fabricação)",
    Cargo == 7315~ "Cortadores, polidores, jateadores e gravadores de vidros e afins",
    Cargo == 7316~ "Redatores de cartazes, pintores decorativos e gravadores",
    Cargo == 7317~ "Artesãos de pedra, madeira, vime e materiais semelhantes",
    Cargo == 7318~ "Artesãos de tecidos, couros e materiais semelhantes",
    Cargo == 7319~ "Artesãos não classificados anteriormente",
    Cargo == 7321~ "Trabalhadores da pré-impressão gráfica",
    Cargo == 7322~ "Impressores",
    Cargo == 7323~ "Encadernadores e afins",
    Cargo == 7411~ "Eletricistas de obras e afins",
    Cargo == 7412~ "Mecânicos e ajustadores eletricistas",
    Cargo == 7413~ "Instaladores e reparadores de linhas elétricas",
    Cargo == 7421~ "Mecânicos e reparadores em eletrônica",
    Cargo == 7422~ "Instaladores e reparadores em tecnologias da informação e comunicações",
    Cargo == 7511~ "Magarefes e afins",
    Cargo == 7512~ "Padeiros, confeiteiros e afins",
    Cargo == 7513~ "Trabalhadores da pasteurização do leite e fabricação de laticínios e afins",
    Cargo == 7514~ "Trabalhadores da conservação de frutas, legumes e similares",
    Cargo == 7515~ "Trabalhadores da degustação e classificação de alimentos e bebidas",
    Cargo == 7516~ "Trabalhadores qualificados da preparação do fumo e seus produtos",
    Cargo == 7521~ "Trabalhadores de tratamento e preparação da madeira",
    Cargo == 7522~ "Marceneiros e afins",
    Cargo == 7523~ "Operadores de máquinas de lavrar madeira",
    Cargo == 7531~ "Alfaiates, modistas, chapeleiros e peleteiros",
    Cargo == 7532~ "Trabalhadores qualificados da preparação da confecção de roupas",
    Cargo == 7533~ "Costureiros, bordadeiros e afins",
    Cargo == 7534~ "Tapeceiros, colchoeiros e afins",
    Cargo == 7535~ "Trabalhadores qualificados do tratamento de couros e peles",
    Cargo == 7536~ "Sapateiros e afins",
    Cargo == 7541~ "Trabalhadores subaquáticos",
    Cargo == 7542~ "Dinamitadores e detonadores",
    Cargo == 7543~ "Classificadores e provadores de produtos (exceto de bebidas e alimentos)",
    Cargo == 7544~ "Fumigadores e outros controladores de pragas e ervas daninhas",
    Cargo == 7549~ "Outros trabalhadores qualificados e operários da indústria e do artesanato não classificados anteriormente",
    Cargo == 8111~ "Mineiros e operadores de máquinas e de instalações em minas e pedreiras",
    Cargo == 8112~ "Operadores de instalações de processamento de minerais e rochas",
    Cargo == 8113~ "Perfuradores e sondadores de poços e afins",
    Cargo == 8114~ "Operadores de máquinas para fabricar cimento, pedras e outros produtos minerais",
    Cargo == 8121~ "Operadores de instalações de processamento de metais",
    Cargo == 8122~ "Operadores de máquinas polidoras, galvanizadoras e recobridoras de metais",
    Cargo == 8131~ "Operadores de instalações e máquinas de produtos químicos",
    Cargo == 8132~ "Operadores de máquinas para fabricar produtos fotográficos",
    Cargo == 8141~ "Operadores de máquinas para fabricar produtos de borracha",
    Cargo == 8142~ "Operadores de máquinas para fabricar produtos de material plástico",
    Cargo == 8143~ "Operadores de máquinas para fabricar produtos de papel",
    Cargo == 8151~ "Operadores de máquinas de preparação de fibras, fiação e bobinamento de fios",
    Cargo == 8152~ "Operadores de teares e outras máquinas de tecelagem",
    Cargo == 8153~ "Operadores de máquinas de costura",
    Cargo == 8154~ "Operadores de máquinas de branqueamento, tingimento e limpeza de tecidos",
    Cargo == 8155~ "Operadores de máquinas de processamento de couros e peles",
    Cargo == 8156~ "Operadores de máquinas para fabricação de calçados e afins",
    Cargo == 8157~ "Operadores de máquinas de lavar, tingir e passar roupas",
    Cargo == 8159~ "Operadores de máquinas para fabricar produtos têxteis e artigos de couro e pele não classificados anteriormente",
    Cargo == 8160~ "Operadores de máquinas para elaborar alimentos e produtos afins",
    Cargo == 8171~ "Operadores de instalações para a preparação de pasta de papel e papel",
    Cargo == 8172~ "Operadores de instalações para processamento de madeira",
    Cargo == 8181~ "Operadores de instalações de vidraria e cerâmica",
    Cargo == 8182~ "Operadores de máquinas de vapor e caldeiras",
    Cargo == 8183~ "Operadores de máquinas de embalagem, engarrafamento e etiquetagem",
    Cargo == 8189~ "Operadores de máquinas e de instalações fixas não classificados anteriormente",
    Cargo == 8211~ "Mecânicos montadores de maquinaria mecânica",
    Cargo == 8212~ "Montadores de equipamentos elétricos e eletrônicos",
    Cargo == 8219~ "Montadores não classificados anteriormente",
    Cargo == 8311~ "Maquinistas de locomotivas",
    Cargo == 8312~ "Guarda-freios e agentes de manobras",
    Cargo == 8321~ "Condutores de motocicletas",
    Cargo == 8322~ "Condutores de automóveis, taxis e caminhonetes",
    Cargo == 8331~ "Condutores de ônibus e bondes",
    Cargo == 8332~ "Condutores de caminhões pesados",
    Cargo == 8341~ "Operadores de máquinas agrícolas e florestais móveis",
    Cargo == 8342~ "Operadores de máquinas de movimentação de terras e afins",
    Cargo == 8343~ "Operadores de guindastes, gruas, aparatos de elevação e afins",
    Cargo == 8344~ "Operadores de empilhadeiras",
    Cargo == 8350~ "Marinheiros de coberta e afins",
    Cargo == 9111~ "Trabalhadores dos serviços domésticos em geral",
    Cargo == 9112~ "Trabalhadores de limpeza de interior de edifícios, escritórios, hotéis e outros estabelecimentos",
    Cargo == 9121~ "Lavadeiros de roupas e passadeiros manuais",
    Cargo == 9122~ "Lavadores de veículos",
    Cargo == 9123~ "Limpadores de janelas",
    Cargo == 9129~ "Outros trabalhadores de limpeza",
    Cargo == 9211~ "Trabalhadores elementares da agricultura",
    Cargo == 9212~ "Trabalhadores elementares da pecuária",
    Cargo == 9213~ "Trabalhadores elementares da agropecuária",
    Cargo == 9214~ "Trabalhadores elementares da jardinagem e horticultura",
    Cargo == 9215~ "Trabalhadores florestais elementares",
    Cargo == 9216~ "Trabalhadores elementares da pesca e aquicultura",
    Cargo == 9311~ "Trabalhadores elementares de minas e pedreiras",
    Cargo == 9312~ "Trabalhadores elementares de obras públicas e da manutenção de estradas, represas e similares",
    Cargo == 9313~ "Trabalhadores elementares da construção de edifícios",
    Cargo == 9321~ "empacotadores manuais",
    Cargo == 9329~ "Trabalhadores elementares da indústria de transformação não classificados anteriormente",
    Cargo == 9331~ "Condutores de veículos acionados a pedal ou a braços",
    Cargo == 9332~ "Condutores de veículos e máquinas de tração animal",
    Cargo == 9333~ "Carregadores",
    Cargo == 9334~ "Repositores de prateleiras",
    Cargo == 9411~ "Preparadores de comidas rápidas",
    Cargo == 9412~ "Ajudantes de cozinha",
    Cargo == 9510~ "Trabalhadores ambulantes dos serviços e afins",
    Cargo == 9520~ "Vendedores ambulantes (exclusive de serviços de alimentação)",
    Cargo == 9611~ "Coletores de lixo e material reciclável",
    Cargo == 9612~ "Classificadores de resíduos",
    Cargo == 9613~ "Varredores e afins",
    Cargo == 9621~ "Mensageiros, carregadores de bagagens e entregadores de encomendas",
    Cargo == 9622~ "Pessoas que realizam várias tarefas",
    Cargo == 9623~ "Coletores de dinheiro em máquinas automáticas de venda e leitores de medidores",
    Cargo == 9624~ "Carregadores de água e coletores de lenha",
    Cargo == 9629~ "Outras ocupações elementares não classificadas anteriormente",
    Cargo == 0110~ "Oficiais das forças armadas",
    Cargo == 0210~ "Graduados e praças das forças armadas",
    Cargo == 0411~ "Oficiais de polícia militar",
    Cargo == 0412~ "Graduados e praças da polícia militar",
    Cargo == 0511~ "Oficiais de bombeiro militar",
    Cargo == 0512~ "Graduados e praças do corpo de bombeiros",))
  
  #filtrando só as de TI ( CNAE 6200 E 6300)
  pnad_TI <-pnad %>% 
    filter(str_detect(Atividade, "Atividades dos Serviços de TI|Atividades de Prestação de Serviços de TI"))
  #Criando um binding para poder apresentar o total no Excel?
  pnad_TI2<-pnad_TI%>% mutate(Atividade = case_when(
    Atividade=="Atividades dos Serviços de TI"~"Setor TI",
    Atividade=="Atividades de Prestação de Serviços de TI"~"Setor TI",
  )
  )
  pnad_Ativ<-rbind(pnad_TI,pnad_TI2)
  
  ##Criando as tabelas din?micas
  wb <- createWorkbook(creator = Sys.getenv("Gabriel"))
  #Cargo_Funcao
  
  cf<- PivotTable$new()
  cf$addData(pnad_Ativ)
  cf$addColumnDataGroups("Atividade",
                         dataSortOrder="custom", 
                         customSortOrder=c("Atividades dos Serviços de TI","Atividades de Prestação de Serviços de TI","Demais Setores")
  )
  
  cf$addColumnDataGroups("Regiao",
                         dataSortOrder="custom", 
                         customSortOrder=c("Norte","Nordeste","Sudeste","Sul","Centro-Oeste") 
  )
  
  cf$addRowDataGroups("Cargo")
  
  
  cf$defineCalculation(calculationName="Total", summariseExpression="sum(Peso, na.rm=TRUE)")
  
  cf$evaluatePivot()
  addWorksheet(wb, "cargo_func")
  cf$writeToExcelWorksheet(wb=wb, wsName="cargo_func", 
                           topRowNumber=1, leftMostColumnNumber=1, applyStyles=FALSE, exportOptions=list(skipNA=TRUE, skipNaN=TRUE))
  saveWorkbook(wb, file=sprintf("PNAD_%s_%s_2.xlsx",year,quarter), overwrite = TRUE)
  
  #Renda_Cargo
  rc<- PivotTable$new()
  rc$addData(pnad_Ativ)
  rc$addColumnDataGroups("Atividade",
                         dataSortOrder="custom", 
                         customSortOrder=c("Atividades dos Serviços de TI","Atividades de Prestação de Serviços de TI","Demais Setores")
  )
  
  rc$addColumnDataGroups("Regiao",
                         dataSortOrder="custom", 
                         customSortOrder=c("Norte","Nordeste","Sudeste","Sul","Centro-Oeste") 
  )
  
  rc$addRowDataGroups("Cargo")
  
  
  rc$defineCalculation(calculationName="Total", summariseExpression="weighted.mean(Rendimento,Peso, na.rm=TRUE)")
  
  rc$evaluatePivot()
  addWorksheet(wb, "rend_cargo")
  rc$writeToExcelWorksheet(wb=wb, wsName="rend_cargo", 
                           topRowNumber=1, leftMostColumnNumber=1, applyStyles=FALSE, exportOptions=list(skipNA=TRUE, skipNaN=TRUE))
  saveWorkbook(wb, file=sprintf("PNAD_%s_%s_2.xlsx",year,quarter), overwrite = TRUE)
  sprintf("concluído ano %s trimestre %s", year, quarter)
  rm(list=ls(all=TRUE))
  gc()
}
generate_pnad_sheet_2(2022,1)

generate_pnad_sheet_2(2021,4)
generate_pnad_sheet_2(2021,3)
generate_pnad_sheet_2(2021,2)
generate_pnad_sheet_2(2021,1)

generate_pnad_sheet_2(2020,4)
generate_pnad_sheet_2(2020,3)
generate_pnad_sheet_2(2020,2)
generate_pnad_sheet_2(2020,1)

generate_pnad_sheet_2(2019,4)
generate_pnad_sheet_2(2019,3)
generate_pnad_sheet_2(2019,2)
generate_pnad_sheet_2(2019,1)

generate_pnad_sheet_2(2018,4)
generate_pnad_sheet_2(2018,3)
generate_pnad_sheet_2(2018,2)
generate_pnad_sheet_2(2018,1)
