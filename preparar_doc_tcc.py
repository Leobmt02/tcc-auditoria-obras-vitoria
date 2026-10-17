import os
import shutil
import fitz
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image as RLImage, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table as RLTable, TableStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT

# 1. Copiar imagens das 4 abas do dashboard enviadas pelo usuário
os.makedirs("img_painel", exist_ok=True)
user_imgs = [
    (r"C:\Users\Leonardo Matias\.gemini\antigravity\brain\459a3da5-49bc-4411-8d38-ccfbdce56f9d\.user_uploaded\media_1792203833441_c7f2a944.png", "img_painel/aba1_obras.png"),
    (r"C:\Users\Leonardo Matias\.gemini\antigravity\brain\459a3da5-49bc-4411-8d38-ccfbdce56f9d\.user_uploaded\media_1792203841625_c3fea504.png", "img_painel/aba2_empresas.png"),
    (r"C:\Users\Leonardo Matias\.gemini\antigravity\brain\459a3da5-49bc-4411-8d38-ccfbdce56f9d\.user_uploaded\media_1792203851229_fd9cdae0.png", "img_painel/aba3_painel_fisico.png"),
    (r"C:\Users\Leonardo Matias\.gemini\antigravity\brain\459a3da5-49bc-4411-8d38-ccfbdce56f9d\.user_uploaded\media_1792203861126_46de6e9f.png", "img_painel/aba4_todos_contratos.png"),
]
for src, dst in user_imgs:
    shutil.copyfile(src, dst)

# Extrair a imagem do e-mail da página 12 do PDF original para manter o Anexo original também
pdf_orig = [f for f in os.listdir(".") if "(2).pdf" in f][0]
doc_pdf = fitz.open(pdf_orig)
for img_info in doc_pdf.get_page_images(11):
    xref = img_info[0]
    pix = fitz.Pixmap(doc_pdf, xref)
    pix.save("img_painel/anexo_email.png")
    break
doc_pdf.close()

# ==============================================================================
# PARTE A: ATUALIZAR O RELATÓRIO TÉCNICO (Analise_Contratos_PMV_2025_2026.pdf)
# ==============================================================================
# Vamos chamar o gerar_pdf.py mas adicionando a seção com as 4 telas do painel!

# ==============================================================================
# PARTE B: GERAR O DOCUMENTO OFICIAL DO TCC (DOCX + PDF)
# ==============================================================================

SECOES = [
    ("1 INTRODUÇÃO", [
        ("1.1 O PROBLEMA", [
            "No Brasil, e especificamente no Estado do Espírito Santo e em sua capital, Vitória, a contratação de obras e serviços públicos segue os parâmetros instituídos pela Lei nº 14.133/2021 (Nova Lei de Licitações e Contratos Administrativos). Entre os critérios de julgamento utilizados nos processos licitatórios, destaca-se o menor preço ou maior desconto, cujo objetivo é assegurar maior economicidade e eficiência na aplicação dos recursos da Administração Pública.",
            "Entretanto, durante a execução contratual, diversos instrumentos sofrem alterações financeiras e de vigência por meio de termos aditivos e apostilamentos, resultando em divergências expressivas entre o valor inicialmente contratado, o valor atualizado e o montante efetivamente empenhado, liquidado e pago. Nesse contexto, torna-se relevante analisar historicamente e quantitativamente essas alterações para compreender os padrões de comportamento contratual e verificar a aderência aos limites legais previstos no Art. 125 da Lei nº 14.133/2021 (25% para obras, serviços e compras em geral, e 50% para reforma de edifício ou equipamento).",
            "Tanto a Prefeitura Municipal de Vitória (PMV), por meio de seu Portal da Transparência, quanto o Tribunal de Contas do Estado do Espírito Santo (TCE-ES), por meio do sistema CidadES Contratações, e o Portal Nacional de Contratações Públicas (PNCP) disponibilizam dados abertos sobre licitações, contratos, termos aditivos, medições e execução orçamentária. Apesar da disponibilidade pública dessas informações, os dados encontram-se fragmentados em diferentes plataformas, com estruturas heterogêneas, chaves distintas e inconsistências de formatação, dificultando análises integradas e a identificação de padrões relevantes para apoio à tomada de decisão. Dessa forma, a engenharia, o cruzamento e a análise estruturada desses dados contribuem para uma visão estratégica, transparente e auditável sobre a execução dos contratos e obras públicas."
        ]),
        ("1.2 FORMULAÇÃO DO PROBLEMA", [
            "De que forma a integração e a análise estruturada dos dados abertos da Prefeitura Municipal de Vitória (PMV), do Tribunal de Contas do Estado do Espírito Santo (TCE-ES) e do Portal Nacional de Contratações Públicas (PNCP) podem ser sistematizadas para identificar padrões de termos aditivos, recorrência de fornecedores, desvios em relação aos limites legais e prever o risco de sobrecusto em contratações e obras públicas?"
        ]),
        ("1.3 HIPÓTESE", [
            "A integração automatizada e a análise estruturada das bases de dados abertos da PMV, do TCE-ES e do PNCP permitem identificar padrões objetivos relacionados à ocorrência de termos aditivos, à concentração de contratos em determinados fornecedores, à natureza real do objeto contratado e às diferenças entre valores inicialmente contratados, atualizados e efetivamente pagos, possibilitando a formulação de indicadores quantitativos de risco e painéis analíticos aplicados à auditoria de contratos públicos."
        ]),
        ("1.4 OBJETIVOS", [
            "1.4.1 Geral",
            "Construir um sistema analítico de suporte à decisão e auditoria quantitativa, integrando a coleta, o tratamento e o cruzamento de dados abertos da Prefeitura Municipal de Vitória (PMV), do TCE-ES (CidadES) e do PNCP, visando identificar padrões de termos aditivos, recorrência de empresas e risco financeiro de sobrecusto em obras e contratos públicos.",
            "1.4.2 Específicos",
            "• Extrair, higienizar e consolidar os dados públicos de contratos administrativos (2025–2026) e do Painel de Obras (2021–2026) do Portal da Transparência da Prefeitura Municipal de Vitória, incluindo unidades da Administração Direta e Indireta (IPAMV e CDTIV);",
            "• Integrar programaticamente (via API REST e bases de Dados Abertos) as informações de licitações do Portal Nacional de Contratações Públicas (PNCP) e os registros de instrumentos contratuais, termos aditivos, apostilamentos e medições do TCE-ES (CidadES Contratações);",
            "• Mapear e classificar os contratos segundo a natureza real do objeto (Subcategoria do Objeto) e avaliar a conformidade das alterações financeiras frente aos limites legais do Art. 125 da Lei nº 14.133/2021 (25% para obras/serviços e 50% para reformas);",
            "• Mensurar a recorrência de empresas contratadas, a proporção de contratos aditados por fornecedor e as discrepâncias entre valor estimado, homologado, contratado, aditado, empenhado, liquidado e pago;",
            "• Apresentar visualmente os indicadores auditados por meio de relatórios quantitativos automatizados e de um painel interativo web (Dashboard de Business Intelligence) publicado em ambiente aberto para apoio à gestão e ao controle social."
        ]),
        ("1.5 JUSTIFICATIVA", [
            "A gestão eficiente dos recursos públicos representa um dos principais desafios da administração pública contemporânea. Nesse contexto, torna-se relevante o desenvolvimento de mecanismos computacionais que auxiliem na auditoria e no acompanhamento da execução contratual, especialmente em obras de engenharia e serviços que apresentam alterações financeiras ao longo do tempo. Para a sociedade, a importância deste trabalho está relacionada à ampliação da transparência ativa e ao controle social, reduzindo assimetrias de informação associadas a atrasos, paralisações, aditivos sucessivos e sobrecustos em obras e serviços públicos.",
            "O projeto realiza a integração inédita de quatro fontes oficiais — Contratos PMV, Painel de Obras PMV, Licitações PNCP e CidadES/TCE-ES — superando limitações individuais de cada portal (como a ausência de detalhamento de aditivos no Painel de Obras ou a ausência de valores empenhados/pagos na base bruta de Contratos). Dessa forma, a pesquisa entrega uma ferramenta prática de apoio à tomada de decisão e análise de risco em processos licitatórios.",
            "Para a área de Engenharia da Computação, a pesquisa demonstra a aplicação prática da Engenharia de Dados e Ciência de Dados no contexto da governança pública, integrando coleta via APIs governamentais, expressões regulares para recuperação de arquivos corrompidos, algoritmos de cruzamento determinístico (data linkage), estatística descritiva e visualização interativa em arquitetura web."
        ]),
        ("1.6 DELIMITAÇÃO DO ESCOPO: RECORTE NOS PROJETOS DE OBRAS DA PREFEITURA DE VITÓRIA", [
            "Embora a problemática de sobrecusto em licitações abranja a Administração Pública em todo o Estado do Espírito Santo, para garantir profundidade analítica, padronização metodológica e rastreabilidade ponta a ponta dos contratos, o escopo empírico deste trabalho foi delimitado especificamente aos projetos e contratos de obras públicas da Prefeitura Municipal de Vitória (PMV).",
            "A escolha de restringir o foco principal às obras de engenharia do Município de Vitória fundamenta-se em três critérios técnicos e metodológicos:",
            "• Relevância Orçamentária e Complexidade Executiva: Mesmo representando cerca de 9,5% da quantidade total de contratos firmados pelo município em 2025 e 2026 (101 contratos de obras dentro de um universo de 1.065 contratos), os projetos de obras concentram R$ 545,99 milhões em valor inicial contratado — mais da metade (54,1%) de todo o volume financeiro contratado pela Prefeitura no período (R$ 1,009 bilhão, excluído outlier de digitação), além de R$ 2,18 bilhões acompanhados no Painel de Obras (2021–2026);",
            "• Enquadramento Legal Específico na Lei nº 14.133/2021: Projetos de obras e reformas possuem dinâmica própria de alterações quantitativas e qualitativas de escopo (falhas de projeto básico, imprevistos geotécnicos e reequilíbrios), sujeitando-se diretamente aos tetos legais de aditamento do Art. 125 da Nova Lei de Licitações (25% para obras em geral e 50% para reforma de edifício ou equipamento), ao passo que contratos de prestação de serviços contínuos ou fornecimento de materiais possuem regras distintas de prorrogação plurianual;",
            "• Disponibilidade de Cruzamento em Quatro Bases Oficiais: O recorte na Prefeitura de Vitória viabilizou o cruzamento determinístico entre a base de Contratos Administrativos da PMV, o Painel Físico/Financeiro de Obras da PMV (que detalha valores empenhados, liquidados, pagos, bairros e estágios de execução), o Portal Nacional de Contratações Públicas (PNCP) e o sistema CidadES do TCE-ES. A base geral dos 1.065 contratos municipais foi mantida no sistema exclusivamente como linha de base comparativa (baseline), enquanto a auditoria de limites legais, subcategorias de engenharia, recorrência de construtoras e execução físico-financeira concentra-se estritamente nos projetos de obras da capital."
        ])
    ]),
    ("2 REFERENCIAL TEÓRICO", [
        ("2.1 LICITAÇÕES PÚBLICAS, TERMOS ADITIVOS E A LEI Nº 14.133/2021", [
            "A licitação pública é o procedimento administrativo utilizado pela Administração Pública para contratar obras, serviços e adquirir bens, buscando garantir igualdade de condições entre os concorrentes e selecionar a proposta mais vantajosa para o interesse público. No Brasil, a Nova Lei de Licitações e Contratos Administrativos, Lei nº 14.133/2021, modernizou o marco regulatório, substituindo a antiga Lei nº 8.666/1993 e instituindo o Portal Nacional de Contratações Públicas (PNCP) como repositório centralizado obrigatório.",
            "Segundo Di Pietro (2023), a licitação possui como princípios fundamentais a legalidade, impessoalidade, moralidade, publicidade e eficiência. Durante a execução contratual, o Art. 124 da Lei nº 14.133/2021 prevê que os contratos podem ser alterados unilateralmente pela Administração (quando houver modificação do projeto ou necessidade de acréscimo ou supressão quantitativa do objeto) ou por acordo entre as partes (para restabelecer o equilíbrio econômico-financeiro ou prorrogar vigências).",
            "De forma estrita, o Art. 125 da Lei nº 14.133/2021 estabelece os limites legais para alterações unilaterais quantitativas e qualitativas: nas obras, nos serviços e nas compras, o contratado é obrigado a aceitar, nas mesmas condições contratuais, acréscimos ou supressões de até 25% (vinte e cinco por cento) do valor inicial atualizado do contrato; e, no caso particular de reforma de edifício ou de equipamento, o limite para os acréscimos é de até 50% (cinquenta por cento). Embora permitidos pela legislação, o uso recorrente de aditivos próximos ao teto legal ou acima dele exige auditoria rigorosa, pois pode sinalizar falhas no projeto básico/executivo, subestimação de quantitativos na licitação ou risco de desvirtuamento da proposta originalmente vencedora."
        ]),
        ("2.2 DADOS ABERTOS GOVERNAMENTAIS E TRANSPARÊNCIA PÚBLICA", [
            "Os dados abertos governamentais consistem em informações disponibilizadas pelos órgãos públicos em formato acessível, legível por máquina e não proprietário, permitindo reutilização e auditoria cívica pela sociedade. Conforme a Controladoria-Geral da União (CGU, 2023) e a Lei de Acesso à Informação (Lei nº 12.527/2011), a política de dados abertos fortalece a transparência, o controle social e a tomada de decisão baseada em evidências.",
            "No âmbito municipal e estadual capixaba, os dados de contratações e obras distribuem-se em três esferas complementares: (i) o Portal da Transparência da Prefeitura Municipal de Vitória, que publica as bases de Contratos Administrativos e o Painel de Obras Físico/Financeiro; (ii) o Portal de Dados Abertos do Estado do Espírito Santo e o Painel de Controle do TCE-ES, alimentados pelo sistema CidadES Contratações (Resolução TCE-ES nº 379/2024); e (iii) a API de Consulta do Portal Nacional de Contratações Públicas (PNCP). A dispersão dessas informações entre múltiplos sistemas justifica o emprego de técnicas de Engenharia de Dados para unificação e cruzamento."
        ]),
        ("2.3 ENGENHARIA DE DADOS E PROCESSOS ETL", [
            "A engenharia de dados é a área responsável pela coleta, transformação, higienização, armazenamento e disponibilização de dados para análise. Uma das principais arquiteturas utilizadas é o processo ETL (Extract, Transform and Load), que consiste em extrair dados de múltiplas fontes heterogêneas, transformá-los em um modelo padronizado e íntegro, e carregá-los em estruturas analíticas (Kimball; Ross, 2013).",
            "Em bases públicas reais, o processo de transformação exige tratamento avançado de codificação de caracteres (como conversão de CP1252/Latin-1 para UTF-8), reconstrução de registros corrompidos por quebras de linha não escapadas em campos textuais livres, normalização de chaves compostas (número/ano do contrato, número/ano do processo administrativo e CNPJ do fornecedor) e detecção de valores discrepantes (outliers) decorrentes de erros materiais de digitação na origem."
        ]),
        ("2.4 MODELAGEM RELACIONAL E CRUZAMENTO DE DADOS PÚBLICOS", [
            "Segundo Elmasri e Navathe (2019), a modelagem relacional permite organizar grandes volumes de dados de forma consistente, estabelecendo chaves primárias e estrangeiras que viabilizam junções (joins) entre entidades distintas. No contexto da auditoria governamental, a ausência de um identificador único universal entre os sistemas municipal (PMV), estadual (TCE-ES) e federal (PNCP) exige a construção de chaves determinísticas compostas:",
            "• Chave de Processo Administrativo (Processo/AnoProcesso): conecta o contrato assinado na Prefeitura ao procedimento licitatório original no PNCP;",
            "• Chave de Instrumento e Fornecedor (NumeroContrato/AnoContrato + CNPJ): conecta o contrato municipal ao registro de instrumentos, termos aditivos e medições no CidadES/TCE-ES, evitando colisões de numeração entre diferentes Unidades Gestoras (como Fundos Municipais, IPAMV e CDTIV);",
            "• Chave de Obra (NumeroContrato/AnoContrato extraído por expressão regular): vincula o instrumento contratual ao Painel de Obras da PMV para recuperação dos montantes empenhados, liquidados, pagos, bairros atendidos e estágio físico."
        ]),
        ("2.5 BUSINESS INTELLIGENCE E ANÁLISE QUANTITATIVA DE RISCO", [
            "Business Intelligence (BI) refere-se ao conjunto de técnicas, arquiteturas e ferramentas utilizadas para transformar dados brutos em informações estratégicas e visuais para apoio à tomada de decisão. Segundo Provost e Fawcett (2016), a combinação de estatística descritiva robusta (média, mediana, desvio padrão, extremos) com visualização interativa e indicadores de recorrência permite detectar anomalias, concentração de mercado e riscos operacionais de forma imediata e auditável."
        ])
    ]),
    ("3 METODOLOGIA, ORIGEM DOS DADOS E LÓGICA DAS ANÁLISES", [
        ("3.1 NATUREZA DA PESQUISA E DELIMITAÇÃO AOS PROJETOS DE OBRAS DA PMV", [
            "A presente pesquisa caracteriza-se como aplicada, pois desenvolve uma solução tecnológica concreta — composta por pipeline de dados em Python, base relacional consolidada, relatório quantitativo automatizado e dashboard web interativo — voltada à auditoria de contratos e análise de sobrecustos na gestão pública. Quanto à abordagem, a pesquisa é estritamente quantitativa e objetiva, fundamentada exclusivamente em métricas financeiras, variações percentuais, testes de consistência matemática e parâmetros legais (Lei nº 14.133/2021), abstendo-se de juízos de valor subjetivos. Quanto aos objetivos, classifica-se como exploratória e descritiva.",
            "Conforme estabelecido na delimitação de escopo (Seção 1.6), o estudo empírico foi direcionado e limitado aos projetos e contratos de obras públicas da Prefeitura Municipal de Vitória (PMV). Para contextualizar o peso das obras dentro da administração municipal, coletou-se inicialmente a totalidade dos contratos administrativos de 2025 e 2026 da PMV (1.065 contratos), aplicando-se em seguida o filtro metodológico sobre os contratos de obras de engenharia (101 contratos classificados como TipoContrato = Obra em 2025–2026) e sobre o histórico físico-financeiro do Painel de Obras municipal (179 obras de 2021 a 2026)."
        ]),
        ("3.2 ORIGEM DOS DADOS E FONTES UTILIZADAS", [
            "Para superar a fragmentação informacional identificada na formulação do problema, o estudo coletou e integrou quatro bases de dados públicas oficiais referentes ao Município de Vitória/ES (CNPJ do órgão principal: 27.142.058/0001-26):"
        ]),
        ("TABELA_FONTES", []),
        ("3.3 PIPELINE DE COLETA, HIGIENIZAÇÃO E INTEGRIDADE (ETL)", [
            "O processamento dos dados foi implementado em linguagem Python (utilizando as bibliotecas Pandas, NumPy, Requests, Regular Expressions, Matplotlib e ReportLab) seguindo cinco etapas rigorosas de tratamento:",
            "1. Recuperação de Registros Quebrados nos CSVs da PMV (limpar_contratos_pmv.py): Os arquivos brutos exportados pelo Portal da Transparência da PMV (codificação CP1252) apresentavam quebras de linha literais (CRLF) no interior do campo textual Objeto, o que fazia com que leitores convencionais partissem um único contrato em duas linhas (gerando 472 fragmentos órfãos contendo apenas o valor ',Não' da última coluna). Foi desenvolvido um parser baseado em expressão regular (\\r?\\n(?=\\d+,20\\d\\d,)) que identifica o início exato de cada registro contratual, recompõe os 20 campos originais sem perda de dados e converte os valores monetários para ponto flutuante de dupla precisão.",
            "2. Consolidação Multi-Entidades (consolidar_tudo.py): Além dos arquivos centrais da Prefeitura (583 contratos de 2025 e 418 de 2026, totalizando 1.001 contratos), foram incorporados os arquivos específicos das entidades da Administração Indireta com dados válidos: IPAMV (14 contratos em 2025 e 15 em 2026) e CDTIV (21 contratos em 2025 e 14 em 2026), totalizando 1.065 contratos únicos analisados.",
            "3. Extração Automatizada da API do PNCP (baixar_pncp_licitacoes.py): Por meio de requisições HTTP GET ao endpoint /api/consulta/v1/contratacoes/publicacao, iterando sobre janelas mensais (janeiro/2025 a outubro/2026) e sobre todas as 13 modalidades de contratação previstas na Lei nº 14.133/2021, foram extraídas 1.287 licitações do Município de Vitória (518 inexigibilidades, 425 pregões eletrônicos, 264 dispensas, 45 concorrências eletrônicas, 21 credenciamentos, 9 leilões, 4 pregões presenciais e 1 concorrência presencial).",
            "4. Processamento em Streaming das Bases Estaduais do TCE-ES (baixar_tcees.py): A base estadual de alterações e medições do CidadES possui 1,40 GB (990.412 registros). Foi implementado um filtro por lotes (chunks de 100.000 linhas) selecionando estritamente a esfera administrativa 'Vitória', extraindo 1.681 instrumentos contratuais de 2025–2026 (dos quais 695 são do tipo Contrato), 228 registros de obras e 44.531 registros históricos de termos aditivos, apostilamentos e medições.",
            "5. Cruzamento Determinístico entre as Quatro Fontes: (a) O vínculo PMV × PNCP foi realizado pela chave Processo/AnoProcesso, alcançando correspondência para 792 contratos (62 obras); (b) O vínculo PMV × TCE-ES foi realizado pela chave composta NumeroContrato/AnoContrato + CNPJ do Fornecedor, alcançando 100% de exatidão no valor inicial para os 615 contratos localizados; (c) O vínculo Contratos PMV × Painel de Obras PMV foi realizado extraindo os pares número/ano da coluna Contratos da base de obras, vinculando 40 das 101 obras de 2025–2026 aos seus respectivos valores empenhados, liquidados, pagos, bairros e estágios de execução."
        ]),
        ("3.4 LÓGICA DAS ANÁLISES E FÓRMULAS MATEMÁTICAS", [
            "Todas as análises foram conduzidas sob ótica estritamente quantitativa, aplicando-se as seguintes definições e regras de negócio sobre cada contrato i:",
            "1. Variação Absoluta (Δ Valor) e Percentual (% Δ): Calculadas a partir do Valor Inicial assinado (ValorContratado), do montante líquido de alterações (Aditivo) e do Valor Atualizado (ValorTotal):",
            "   • Δ Valor_i = ValorTotal_i - ValorContratado_i  (em R$)",
            "   • % Δ_i = (Aditivo_i / ValorContratado_i) × 100  (para ValorContratado_i > 0)",
            "2. Verificação de Consistência Interna (Discrepância de Soma): Para 100% dos contratos, verificou-se a identidade algébrica |ValorContratado_i + Aditivo_i - ValorTotal_i| <= R$ 0,01, constatando-se 0 (zero) divergências aritméticas na base higienizada.",
            "3. Tratamento de Outlier por Erro Material na Fonte: Foi identificado que o Contrato nº 300/2025 (Secretaria de Educação — Ensinatech Ltda) possui ValorContratado = R$ 2.139.480,00 e Aditivo = R$ 2.139.480.000,00 (+100.000,00%), evidenciando erro de digitação de três ordens de grandeza (multiplicação por 1.000) no portal municipal. Para evitar a contaminação das médias e somas globais, todas as estatísticas gerais são reportadas em dupla visão: 'Todos (n = 1.065)' e 'Todos sem outlier (n = 1.064)'.",
            "4. Enquadramento nos Limites Legais (Art. 125 da Lei nº 14.133/2021): Para os 101 contratos classificados como TipoContrato = Obra, definiu-se o limite legal aplicável (L_i) de forma automatizada:",
            "   • L_i = 50,0%, quando o campo ReformaDeEdificioOuEquipamento do TCE-ES é verdadeiro ou quando a descrição do Objeto contém o termo 'reforma';",
            "   • L_i = 25,0%, para todas as demais obras de construção, urbanização, contenção, pavimentação e engenharia.",
            "   • O contrato é sinalizado como 'Acima do Limite Legal' quando % Δ_i > L_i, e na 'Faixa de Atenção (Próximo ao Limite)' quando 20,0% <= % Δ_i <= L_i.",
            "5. Mapeamento Semântico por Subcategoria do Objeto (Subcategoria_Objeto): A análise textual dos 101 contratos cadastrados como 'Obra' pela Prefeitura revelou que a classificação original agrupa naturezas distintas. Criou-se um classificador baseado em regras lexicais sobre o campo Objeto, segregando os contratos de obras em 10 subcategorias (Construção/Reforma de Equipamentos Públicos; Urbanização/Praças/Vias; Projetos e Consultoria; Instalações/Tecnologia/Elevadores/Iluminação; Manutenção Predial; Reforma Elétrica/Climatização; Contenção de Encostas; Manutenção de Vias/Drenagem; Melhorias Habitacionais; e Outros Serviços Cadastrados como Obra). Isso permitiu comprovar que contratos aparentemente acima do teto de 25% de obras civis (como o Contrato 223/2025, com +150,51%, e o Contrato 293/2025, com +85,31%) tratam-se, na realidade, de serviços contínuos de manutenção de equipamentos e elevadores cadastrados sob o rótulo de 'Obra'.",
            "6. Desconto Licitatório (PNCP) e Execução Financeira (Empenhado/Liquidado/Pago): Para os contratos vinculados ao PNCP, calculou-se o desconto obtido na fase externa da licitação: Desconto_% = ((ValorEstimado - ValorHomologado) / ValorEstimado) × 100. Para as obras vinculadas ao Painel de Obras, comparou-se o ValorTotal (Inicial + Aditivo) com o ValorEmpenhado, ValorLiquidado e ValorPago, identificando casos em que o empenho ou o pagamento já supera o valor contratual atualizado (por efeito de reajustamentos ou apostilamentos financeiros).",
            "7. Índice de Recorrência e Concentração por Empresa: Agruparam-se os contratos por CPFCNPJ e Fornecedor, contabilizando a quantidade de contratos assinados por empresa (Qtd_Contratos), a quantidade e proporção de contratos que sofreram aditivo (Taxa_Aditamento_Empresa) e o percentual global de acréscimo financeiro por construtora."
        ]),
        ("3.5 SÍNTESE DOS RESULTADOS QUANTITATIVOS OBTIDOS", [
            "A aplicação da metodologia descrita sobre a base consolidada de 2025 e 2026 (1.065 contratos) e sobre o Painel de Obras (179 obras de 2021 a 2026) produziu os seguintes resultados auditados:"
        ]),
        ("TABELA_RESUMO_GRUPOS", []),
        ("TEXTO_RESULTADOS_OBRAS", [
            "No recorte específico dos 101 contratos classificados como Obra (2025–2026), que somam R$ 545.988.525,59 em valor inicial e R$ 14.161.570,88 em aditivos líquidos (+2,59% no agregado), verificou-se que 21 contratos tiveram acréscimo positivo (média de +39,37% e mediana de +24,94% entre os aditados) e 1 contrato apresentou supressão (Contrato 39/2026, de -R$ 1.470.524,90, exatamente compensado por acréscimo de idêntico valor no Contrato 47/2026 da mesma empresa e programa habitacional). Ao todo, 5 contratos ultrapassaram o limite legal aplicável (25% ou 50%) e 8 contratos situaram-se na faixa de 20% até o limite legal (com destaque para aditivos cravados em 24,94% e 25,00%).",
            "Quanto à recorrência empresarial nas obras de 2025–2026, identificou-se que 57 empresas distintas firmaram contratos, das quais 21 são recorrentes (2 ou mais contratos) e concentram 65 dos 101 contratos. As três líderes em quantidade são: Ilumiterra Construções e Montagens Ltda (10 contratos; R$ 18,54 milhões iniciais e R$ 5,38 milhões em aditivos, +29,03%), M.T.F. Construções e Montagens Ltda (7 contratos; R$ 25,74 milhões iniciais) e Engesan Engenharia e Saneamento Ltda (6 contratos; R$ 41,82 milhões iniciais). Destaca-se ainda a empresa Renova Construções e Reformas Ltda, que obteve aditivos em 100% de seus 3 contratos de 2025 (+48,35%, +44,24% e +12,50%), todos enquadrados no limite de 50% por tratarem de reformas."
        ]),
        ("TABELA_SUBCATEGORIAS", []),
        ("3.6 ARQUITETURA E DEMONSTRAÇÃO DO PAINEL INTERATIVO (DASHBOARD WEB)", [
            "Para atender ao objetivo específico de disponibilizar uma ferramenta visual de suporte à decisão para gestores públicos e auditores, foi desenvolvido um Dashboard Web Interativo em arquitetura cliente autocontida (HTML5, CSS3 moderno em tema claro de alta legibilidade, JavaScript ES6 e biblioteca gráfica Chart.js), com os 1.065 contratos e as 179 obras embutidos em estrutura JSON otimizada.",
            "O painel foi publicado em repositório público no GitHub Pages (disponível em: https://leobmt02.github.io/tcc-auditoria-obras-vitoria/) e encontra-se estruturado em quatro visões analíticas (abas), ilustradas nas Figuras 1 a 4 a seguir:"
        ]),
        ("IMAGENS_PAINEL", [])
    ])
]

print("Estrutura de texto pronta!")
