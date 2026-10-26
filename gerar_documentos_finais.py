import os
import pandas as pd
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    Image as RLImage,
    KeepTogether,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table as RLTable,
    TableStyle,
)
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT

# ==============================================================================
# DADOS PARA AS TABELAS DO DOCUMENTO OFICIAL DO TCC
# ==============================================================================

TABELA_FONTES_DATA = [
    ["Fonte / Sistema", "Órgão Responsável", "Escopo Temporal", "Volume de Registros", "Principais Campos Utilizados"],
    [
        "Portal da Transparência — Contratos (Admin. Direta + IPAMV + CDTIV)",
        "Prefeitura Municipal de Vitória (PMV)",
        "2025 e 2026",
        "1.065 contratos únicos (sendo 101 classificados como Obra)",
        "NumeroContrato, AnoContrato, Processo, AnoProcesso, TipoContrato, UnidadeGestora, Fornecedor, CPFCNPJ, ValorContratado, Aditivo, ValorTotal, Objeto"
    ],
    [
        "Portal da Transparência — Painel de Obras Físico/Financeiro",
        "Prefeitura Municipal de Vitória (PMV)",
        "2021 a 2026",
        "179 obras (40 cruzadas diretamente com Contratos 2025–2026)",
        "Contratos vinculados, Valor Inicial, Valor Atual Contratado, ValorEmpenhado, ValorLiquidado, ValorPago, Estagio, Bairros, PrazoInicialDias"
    ],
    [
        "API REST de Consulta de Contratações (/v1/contratacoes/publicacao)",
        "Portal Nacional de Contratações Públicas (PNCP)",
        "2025 e 2026",
        "1.287 licitações publicadas (792 vinculadas a contratos PMV)",
        "numeroControlePNCP, modalidadeNome, processo, valorTotalEstimado, valorTotalHomologado, amparoLegal"
    ],
    [
        "Dados Abertos CidadES Contratações (Instrumentos, Alterações e Obras)",
        "Tribunal de Contas do Estado do ES (TCE-ES)",
        "2025 e 2026 (histórico estadual)",
        "1.681 instrumentos em Vitória (695 Contratos) e 44.531 alterações/medições",
        "IdentificadorContrato, ValorInicial, ReformaDeEdificioOuEquipamento, TipoAlteracao, ValorAlteracao, ValorMedicao"
    ]
]

TABELA_GRUPOS_DATA = [
    ["Grupo Analisado", "N (Contratos)", "Com Aditivo (>0)", "Supressão (<0)", "Valor Inicial Total (R$)", "Soma Aditivos Líquida (R$)", "Valor Total Atualizado (R$)", "% Δ Agregado", "Média % Δ (c/ adit.)", "Mediana % Δ (c/ adit.)"],
    ["Todos os Contratos (2025–2026)", "1.065", "141", "5", "1.011.430.558,80", "2.246.945.642,73", "3.258.376.201,53", "+222,16%", "+744,41%", "+24,58%"],
    ["Todos sem Outlier (s/ 300/2025)*", "1.064", "140", "5", "1.009.291.078,80", "107.465.642,73", "1.116.756.721,53", "+10,65%", "+35,44%", "+24,49%"],
    ["Contratos de Obras (Tipo = Obra)", "101", "21", "1", "545.988.525,59", "14.161.570,88", "560.150.096,47", "+2,59%", "+39,37%", "+24,94%"],
    ["Demais Contratos (s/ 300/2025)", "963", "119", "4", "463.302.553,21", "93.304.071,85", "556.606.625,06", "+20,14%", "+34,75%", "+24,40%"]
]

TABELA_SUBCAT_DATA = [
    ["Subcategoria Real do Objeto (101 Contratos de Obras)", "Qtd.", "C/ Adit.", "> Limite", "20%–Lim.", "Valor Inicial (R$)", "Soma Aditivos (R$)", "Valor Atualizado (R$)", "% Δ Grupo"],
    ["Obra Civil — Construção / Reforma de Equip. Públicos", "36", "5", "1", "2", "183.758.597,22", "1.466.520,56", "185.225.117,78", "+0,80%"],
    ["Obra Civil — Urbanização, Praças, Calçadas e Vias", "22", "3", "0", "3", "218.234.434,53", "3.266.203,95", "221.500.638,48", "+1,50%"],
    ["Serviço Técnico — Projetos, Consultoria e Apoio", "8", "1", "0", "1", "11.559.171,71", "1.484.331,14", "13.043.502,85", "+12,84%"],
    ["Instalações / Tecnologia / Elevadores / Iluminação", "6", "4", "4", "0", "5.318.925,75", "4.828.144,61", "10.147.070,36", "+90,77%"],
    ["Serviço de Engenharia — Manutenção Predial", "6", "1", "0", "0", "25.490.289,14", "324.442,63", "25.814.731,77", "+1,27%"],
    ["Obra Civil — Reforma Elétrica / Climatização", "6", "1", "0", "1", "5.226.404,96", "513.562,67", "5.739.967,63", "+9,83%"],
    ["Obra Civil — Contenção de Encostas / Geotecnia", "5", "1", "0", "0", "57.095.893,85", "468.932,46", "57.564.826,31", "+0,82%"],
    ["Serviço de Engenharia — Manutenção de Vias / Drenagem", "5", "3", "0", "0", "28.910.427,25", "1.787.319,31", "30.697.746,56", "+6,18%"],
    ["Obra Civil — Melhorias Habitacionais (Casa Feliz)", "5", "1", "0", "1", "9.644.996,87", "0,00", "9.644.996,87", "0,00%"],
    ["Outros Serviços Cadastrados como Obra", "2", "1", "0", "0", "749.384,31", "22.113,55", "771.497,86", "+2,95%"]
]

TABELA_ACIMA_LIMITE_DATA = [
    ["Contrato", "Fornecedor", "Subcategoria do Objeto", "Valor Inicial (R$)", "Aditivo (R$)", "Valor Total (R$)", "% Δ", "Limite Legal"],
    ["223/2025", "Karisten Instalação e Manut. de Equipamentos Ltda", "Instalações / Tecnologia / Equipamentos", "122.905,56", "184.989,80", "307.895,36", "+150,51%", "25,0%"],
    ["325/2025", "Ilumiterra Construções e Montagens Ltda", "Instalações / Iluminação Pública", "2.625.276,80", "2.625.276,80", "5.250.553,60", "+100,00%", "25,0%"],
    ["165/2025", "Ilumiterra Construções e Montagens Ltda", "Reforma / Iluminação Campo Futebol", "1.805.280,29", "1.805.280,29", "3.610.560,58", "+100,00%", "50,0% (Reforma)"],
    ["293/2025", "Inove Elevadores Ltda", "Instalações / Elevadores", "266.000,00", "226.934,52", "492.934,52", "+85,31%", "25,0%"],
    ["2/2025", "Matix Construções e Reformas Eireli", "Construção / Reforma Equip. Públicos", "284.337,15", "140.258,28", "424.595,43", "+49,33%", "25,0%"]
]

TABELA_EMPRESAS_OBRAS_DATA = [
    ["Pos.", "Fornecedor (Construtora / Empresa)", "CNPJ", "Qtd. Contratos", "Com Aditivo", "> Limite Legal", "Valor Inicial Total (R$)", "Soma Aditivos (R$)", "Valor Atualizado (R$)", "% Δ Global"],
    ["1º", "Ilumiterra Construções e Montagens Ltda", "08.698.240/0001-30", "10", "3 (30%)", "2", "18.544.611,01", "5.383.861,69", "23.928.472,70", "+29,03%"],
    ["2º", "M.T.F Construções e Montagens Ltda", "40.232.028/0001-63", "7", "2 (29%)", "0", "25.736.939,81", "463.702,21", "26.200.642,02", "+1,80%"],
    ["3º", "Engesan Engenharia e Saneamento Ltda", "27.392.463/0001-74", "6", "1 (17%)", "0", "41.818.422,00", "468.932,46", "42.287.354,46", "+1,12%"],
    ["4º", "C. R. Engenharia Ltda", "29.972.371/0001-95", "4", "0 (0%)", "0", "19.992.950,60", "0,00", "19.992.950,60", "0,00%"],
    ["5º", "Construtora Intra Brasil Ltda", "46.845.106/0001-32", "3", "0 (0%)", "0", "18.880.000,00", "0,00", "18.880.000,00", "0,00%"],
    ["6º", "Sope - Sociedade Projetista e Eng. Ltda", "30.768.261/0001-17", "3", "2 (67%)", "0", "8.622.393,51", "932.226,41", "9.554.619,92", "+10,81%"],
    ["7º", "Sólida Engenharia e Construções Ltda", "11.559.104/0001-33", "3", "1 (33%)", "0", "6.427.674,97", "513.562,67", "6.941.237,64", "+7,99%"],
    ["8º", "Renova Construções e Reformas Ltda", "12.927.589/0001-03", "3", "3 (100%)", "0", "2.626.809,75", "862.559,97", "3.489.369,72", "+32,84%"],
    ["9º", "G Dourado Construtora Ltda", "20.130.654/0001-13", "3", "0 (0%)", "0", "1.454.996,87", "0,00", "1.454.996,87", "0,00%"],
    ["10º", "A.R.C. Comércio Construções e Serv. Ltda", "02.838.689/0001-59", "2", "1 (50%)", "0", "70.312.692,89", "324.442,63", "70.637.135,52", "+0,46%"]
]

CRONOGRAMA_DATA = [
    ["Etapa / Atividade do Projeto", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"],
    ["Definição do tema, problema e objetivos", "X", "", "", "", "", "", "", "", "", "", ""],
    ["Levantamento bibliográfico e Lei 14.133/2021", "X", "X", "", "", "", "", "", "", "", "", ""],
    ["Coleta e ETL das bases PMV, PNCP e TCE-ES", "", "X", "X", "X", "", "", "", "", "", "", ""],
    ["Higienização de CSVs e Cruzamento Relacional", "", "", "X", "X", "X", "", "", "", "", "", ""],
    ["Cálculo de métricas, limites legais e subcategorias", "", "", "", "", "X", "X", "X", "", "", "", ""],
    ["Desenvolvimento e publicação do Dashboard Web", "", "", "", "", "", "", "X", "X", "X", "", ""],
    ["Modelagem preditiva e validação dos resultados", "", "", "", "", "", "", "", "", "X", "X", ""],
    ["Redação final da monografia e defesa do TCC", "", "", "", "", "", "", "", "", "X", "X", "X"]
]

PAINEL_FIGURAS = [
    (
        "a) Visão 1 — Aba '1. Contratos de Obras (101)': Auditoria de Limites Legais e Subcategorias",
        [
            "Finalidade da Visão: Constitui a tela principal do sistema, dedicada exclusivamente ao recorte dos 101 contratos classificados como Obras assinados em 2025 e 2026 pela Prefeitura Municipal de Vitória.",
            "Indicadores Superiores (Cartões KPI): O primeiro cartão exibe o total de contratos filtrados (101 contratos distribuídos em 57 empresas, sendo 21 contratos aditados, ou 20,8%); o segundo apresenta o Valor Inicial Contratado somado (R$ 545.988.525,59); o terceiro destaca o Total em Aditivos líquidos (R$ 14.161.570,88, equivalente a +2,59% no agregado); o quarto mostra o Valor Atualizado Total (R$ 560.150.096,47) confrontado diretamente com o Valor Empenhado recuperado do Painel de Obras (R$ 99,73 milhões para as obras vinculadas); e o quinto cartão alerta para os 5 contratos que ultrapassaram o limite legal do Art. 125 da Lei nº 14.133/2021, trazendo também o montante efetivamente pago mapeado (R$ 69,57 milhões).",
            "Gráficos e Tabela de Auditoria: O gráfico horizontal à esquerda ('Top Contratos por % de Aditivo vs. Limite Legal') ordena os contratos com maior variação percentual, aplicando codificação visual automática: barras em vermelho indicam contratos que excedem o limite legal aplicável (25% para obras em geral ou 50% para reformas), enquanto barras em azul indicam aditivos dentro do teto legal (como aqueles estacionados em 24,94% e 25,00%). O gráfico à direita ('Mapeamento por Subcategoria do Objeto') compara lado a lado o volume inicial contratado (azul) e os aditivos acumulados (laranja) em cada uma das 10 subcategorias reais da engenharia. Na parte inferior, a tabela detalhada permite inspecionar contrato a contrato o valor estimado e homologado no PNCP, o desconto licitatório, os termos no TCE-ES e os valores empenhado e pago."
        ],
        "img_painel/aba1_obras.png",
        "Figura 1 — Aba 1 do Dashboard Interativo: Auditoria dos 101 Contratos de Obras (2025–2026), exibindo cartões de resumo financeiro, comparativo de contratos aditados frente ao limite legal (25%/50%) e distribuição por Subcategoria Real do Objeto."
    ),
    (
        "b) Visão 2 — Aba '2. Ranking de Empresas & Recorrência': Concentração de Fornecedores",
        [
            "Finalidade da Visão: Permite auditar a distribuição dos contratos entre as construtoras e fornecedores, respondendo quantas empresas concentram múltiplos contratos e qual é o comportamento histórico de termos aditivos de cada empresa.",
            "Controles Específicos e Indicadores: Dispõe de um seletor de escopo ('Somente Contratos de Obras — 101' ou 'Todos os Contratos da PMV — 1.064') e de um filtro de recorrência ('Empresas Recorrentes com 2 ou mais, 3 ou mais, 5 ou mais contratos, ou empresas que receberam aditivo'). Os cinco cartões KPI revelam que, nas obras de 2025–2026, atuam 57 empresas distintas, das quais 21 são recorrentes (concentrando 65 dos 101 contratos), destacando automaticamente a empresa líder em quantidade (Ilumiterra Construções, com 10 contratos).",
            "Gráficos Comparativos e Drill-Down ('Ver Contratos'): O gráfico à esquerda ('Empresas com Maior Quantidade de Contratos') exibe para cada empresa duas barras paralelas: o total de contratos assinados (azul) e a quantidade de contratos que sofreram aditivo (laranja), permitindo visualizar imediatamente casos como o da empresa Ilumiterra (10 contratos, 3 aditados) ou Renova Construções (3 contratos, 3 aditados — 100% de taxa de aditamento). O gráfico à direita ('Empresas com Maior Volume Financeiro e Aditivos') ordena os maiores contratados em milhões de reais. Na tabela inferior de classificação geral das empresas, o botão interativo 'Ver Contratos' aplica automaticamente o filtro daquela construtora e redireciona o usuário para a lista completa de seus contratos individuais."
        ],
        "img_painel/aba2_empresas.png",
        "Figura 2 — Aba 2 do Dashboard Interativo: Ranking de Empresas e Recorrência, evidenciando as 57 construtoras contratadas em obras, as 21 empresas recorrentes (2 ou mais contratos) e a proporção de contratos com e sem aditivo por fornecedor."
    ),
    (
        "c) Visão 3 — Aba '3. Painel Físico/Financeiro (179)': Execução Orçamentária (2021–2026)",
        [
            "Finalidade da Visão: Integra os dados históricos do Painel de Obras da Prefeitura Municipal de Vitória (179 intervenções acompanhadas entre 2021 e 2026), suprindo a ausência de dados de execução orçamentária e avanço físico da planilha básica de contratos.",
            "Indicadores de Execução Orçamentária: Os cinco cartões superiores consolidam o ciclo financeiro completo das 179 obras: Valor Atual Contratado de R$ 2.183.310.390,85; Valor Total Empenhado de R$ 1.703.411.375,30 (78,0% do contratado); Valor Total Liquidado (medido e atestado pela fiscalização) de R$ 1.526.084.258,81; e Valor Efetivamente Pago de R$ 1.525.153.301,58 (69,9% do contratado), além de sinalizar as 10 obras paralisadas ou rescindidas unilateralmente.",
            "Gráficos de Estágio Físico e Filtros por Bairro: O gráfico de rosca à esquerda ('Obras por Estágio de Execução') demonstra a proporção física da carteira municipal (62 obras em execução, 50 concluídas e recebidas provisoriamente, 39 recebidas definitivamente, 11 reiniciadas, 8 paralisadas, 6 concluídas e não recebidas e 2 com rescisão unilateral). O gráfico de colunas agrupadas à direita confronta, dentro de cada estágio físico, o montante Contratado (azul), Empenhado (laranja) e Pago (verde). A barra de filtros específica desta aba permite pesquisar obras por bairro atendido (ex.: São Pedro, Jardim Camburi, Maruípe), estágio ou secretaria."
        ],
        "img_painel/aba3_painel_fisico.png",
        "Figura 3 — Aba 3 do Dashboard Interativo: Painel Físico e Financeiro das 179 Obras da PMV (2021–2026), demonstrando os montantes Contratado (R$ 2,18 bi), Empenhado (R$ 1,70 bi), Liquidado e Pago (R$ 1,53 bi), além da distribuição por estágio de execução."
    ),
    (
        "d) Visão 4 — Aba '4. Todos os Contratos PMV (1.065)': Linha de Base Municipal e Tratamento de Outlier",
        [
            "Finalidade da Visão: Apresenta o panorama completo dos 1.065 contratos administrativos firmados pela Prefeitura de Vitória, IPAMV e CDTIV em 2025 e 2026 (incluindo Prestação de Serviços, Fornecimento de Materiais, Locação e Obras), funcionando como linha de base comparativa para o estudo.",
            "Demonstração Prática do Controle de Outlier: Nesta aba, o seletor 'Outlier (300/2025)' vem configurado por padrão na opção recomendada ('Ocultar Outlier'), exibindo os 1.064 contratos consistentes que somam R$ 1.009.291.078,80 iniciais e R$ 107.465.642,73 em aditivos (+10,65% global, com 76 contratos acima de 25%). Caso o auditor alterne o filtro para 'Incluir Outlier', o painel recalcula instantaneamente os cartões para R$ 2,24 bilhões em aditivos (+222,16%), demonstrando visualmente o impacto da distorção causada pelo erro material de digitação no Contrato nº 300/2025.",
            "Comparativo entre Serviços Contínuos e Obras: Os gráficos desta aba evidenciam que os maiores percentuais de acréscimo na administração geral ocorrem em contratos de terceirização de mão de obra, limpeza, transporte e tecnologia da informação (que sofrem prorrogações plurianuais sucessivas), justificando tecnicamente o recorte metodológico do TCC nos contratos de obras de engenharia."
        ],
        "img_painel/aba4_todos_contratos.png",
        "Figura 4 — Aba 4 do Dashboard Interativo: Visão Geral de Todos os 1.065 Contratos da PMV (2025–2026), com filtro ativo de ocultação do outlier (Contrato 300/2025), distribuição por subcategoria geral e ranking de aditivos."
    )
]

REFERENCIAS_LIST = [
    "BRASIL. Lei nº 12.527, de 18 de novembro de 2011. Regula o acesso a informações previsto no inciso XXXIII do art. 5º, no inciso II do § 3º do art. 37 e no § 2º do art. 216 da Constituição Federal. Diário Oficial da União, Brasília, DF, 2011.",
    "BRASIL. Lei nº 14.133, de 1º de abril de 2021. Lei de Licitações e Contratos Administrativos. Diário Oficial da União, Brasília, DF, 2021.",
    "CONTROLADORIA-GERAL DA UNIÃO (CGU). Portal Brasileiro de Dados Abertos: Guia de Transparência Ativa. Brasília: CGU, 2023.",
    "DI PIETRO, Maria Sylvia Zanella. Direito Administrativo. 36. ed. Rio de Janeiro: Forense, 2023.",
    "ELMASRI, Ramez; NAVATHE, Shamkant B. Sistemas de Banco de Dados. 7. ed. São Paulo: Pearson Education do Brasil, 2019.",
    "ESPÍRITO SANTO (Estado). Tribunal de Contas do Estado do Espírito Santo (TCE-ES). Portal CidadES Contratações e Dados Abertos. Vitória: TCE-ES, 2026. Disponível em: https://dados.es.gov.br/dataset/portal-cidades-contratacoes. Acesso em: out. 2026.",
    "KIMBALL, Ralph; ROSS, Margy. The Data Warehouse Toolkit: Guia Completo para Modelagem Dimensional. 3. ed. Rio de Janeiro: Elsevier, 2013.",
    "PORTAL NACIONAL DE CONTRATAÇÕES PÚBLICAS (PNCP). Manual de Integração e API de Consulta do PNCP (Versão 1.0). Brasília: Ministério da Gestão e da Inovação em Serviços Públicos, 2025. Disponível em: https://pncp.gov.br/api/consulta/swagger-ui/index.html. Acesso em: out. 2026.",
    "PREFEITURA MUNICIPAL DE VITÓRIA (PMV). Portal da Transparência: Contratos Administrativos e Painel de Obras Públicas. Vitória: PMV, 2026. Disponível em: https://transparencia.vitoria.es.gov.br/. Acesso em: out. 2026.",
    "PROVOST, Foster; FAWCETT, Tom. Data Science para Negócios: O que você precisa saber sobre mineração de dados e pensamento analítico de dados. Rio de Janeiro: Alta Books, 2016."
]

# ==============================================================================
# 1. GERADOR DO DOCUMENTO WORD (.DOCX) — PADRÃO ACADÊMICO FAESA / ABNT
# ==============================================================================

def set_cell_bg(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def add_formatted_table_docx(doc, data, caption_text):
    p_cap = doc.add_paragraph()
    p_cap.paragraph_format.space_before = Pt(8)
    p_cap.paragraph_format.space_after = Pt(4)
    run_cap = p_cap.add_run(caption_text)
    run_cap.bold = True
    run_cap.font.size = Pt(10)
    run_cap.font.name = "Arial"

    table = doc.add_table(rows=len(data), cols=len(data[0]))
    table.style = 'Table Grid'
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for r_idx, row in enumerate(data):
        for c_idx, val in enumerate(row):
            cell = table.cell(r_idx, c_idx)
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.05
            if r_idx == 0:
                set_cell_bg(cell, "1E293B")
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(str(val))
                run.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(8.5)
                run.font.name = "Arial"
            else:
                if r_idx % 2 == 1:
                    set_cell_bg(cell, "F8FAFC")
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT if c_idx in (0, 1, 2) and len(str(val)) > 14 else WD_ALIGN_PARAGRAPH.RIGHT if any(ch.isdigit() for ch in str(val)) else WD_ALIGN_PARAGRAPH.CENTER
                run = p.add_run(str(val))
                run.font.size = Pt(8.5)
                run.font.name = "Arial"
                if r_idx == 1 and "Todos os Contratos" not in str(row[0]):
                    pass
    doc.add_paragraph().paragraph_format.space_after = Pt(4)

def add_bookmark_docx(paragraph, bookmark_name, bm_id):
    start = OxmlElement('w:bookmarkStart')
    start.set(qn('w:id'), str(bm_id))
    start.set(qn('w:name'), bookmark_name)
    paragraph._p.insert(0, start)
    end = OxmlElement('w:bookmarkEnd')
    end.set(qn('w:id'), str(bm_id))
    paragraph._p.append(end)

def add_toc_hyperlink_docx(paragraph, anchor_name, text, is_bold=False):
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('w:anchor'), anchor_name)
    hyperlink.set(qn('w:history'), '1')

    new_run = OxmlElement('w:r')
    rPr = OxmlElement('w:rPr')

    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:ascii'), 'Arial')
    rFonts.set(qn('w:hAnsi'), 'Arial')
    rFonts.set(qn('w:cs'), 'Arial')
    rPr.append(rFonts)

    if is_bold:
        b = OxmlElement('w:b')
        rPr.append(b)

    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), '20')  # 10pt
    rPr.append(sz)

    color = OxmlElement('w:color')
    color.set(qn('w:val'), '000000')
    rPr.append(color)

    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'none')
    rPr.append(u)

    new_run.append(rPr)
    t = OxmlElement('w:t')
    t.set(qn('xml:space'), 'preserve')
    t.text = text
    new_run.append(t)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

def make_anchor_id(title_str):
    clean = "".join(ch if ch.isalnum() else "_" for ch in title_str.strip())
    while "__" in clean:
        clean = clean.replace("__", "_")
    return "sec_" + clean[:40].strip("_")

def gerar_docx_oficial(output_docx):
    from preparar_doc_tcc import SECOES
    doc = Document()

    # Margens ABNT
    for section in doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(2.0)

    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Arial'
    style_normal.font.size = Pt(11)
    style_normal.font.color.rgb = RGBColor(0, 0, 0)

    for h_name, h_sz in [('Heading 1', 12), ('Heading 2', 11)]:
        if h_name in doc.styles:
            st_h = doc.styles[h_name]
            st_h.font.name = 'Arial'
            st_h.font.size = Pt(h_sz)
            st_h.font.bold = True
            st_h.font.color.rgb = RGBColor(0, 0, 0)

    # --- CAPA ---
    for txt, sz, bld, sp_after in [
        ("FAESA CENTRO UNIVERSITÁRIO", 12, True, 4),
        ("UNIDADE DE EDUCAÇÃO TECNOLÓGICA E COMPUTAÇÃO", 11, True, 36),
        ("CAMILA RAMOS BASSETE", 12, True, 2),
        ("LEONARDO BANDEIRA MATIAS", 12, True, 2),
        ("MATHEUS GOMES EVANGELISTA", 12, True, 90),
        ("PREDIÇÃO DE SOBRECUSTO EM LICITAÇÕES DE PROJETOS GOVERNAMENTAIS DO ESTADO DO ESPÍRITO SANTO", 14, True, 175),
        ("VITÓRIA", 12, True, 2),
        ("2026", 12, True, 0),
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(sp_after)
        r = p.add_run(txt)
        r.bold = bld
        r.font.size = Pt(sz)
        r.font.name = "Arial"
        r.font.color.rgb = RGBColor(0, 0, 0)

    doc.add_page_break()

    # --- FOLHA DE ROSTO ---
    for txt, sz, bld, sp_after in [
        ("CAMILA RAMOS BASSETE", 12, True, 2),
        ("LEONARDO BANDEIRA MATIAS", 12, True, 2),
        ("MATHEUS GOMES EVANGELISTA", 12, True, 70),
        ("PREDIÇÃO DE SOBRECUSTO EM LICITAÇÕES DE PROJETOS GOVERNAMENTAIS DO ESTADO DO ESPÍRITO SANTO", 14, True, 50),
    ]:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(sp_after)
        r = p.add_run(txt)
        r.bold = bld
        r.font.size = Pt(sz)
        r.font.name = "Arial"
        r.font.color.rgb = RGBColor(0, 0, 0)

    p_nat = doc.add_paragraph()
    p_nat.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_nat.paragraph_format.left_indent = Cm(7.5)
    p_nat.paragraph_format.space_after = Pt(12)
    r_nat = p_nat.add_run(
        "Trabalho de Conclusão de Curso apresentado ao Curso de Engenharia da Computação "
        "do Centro Universitário FAESA, como requisito parcial para obtenção do grau de "
        "Bacharel em Engenharia da Computação."
    )
    r_nat.font.size = Pt(10)
    r_nat.font.name = "Arial"

    p_ori = doc.add_paragraph()
    p_ori.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_ori.paragraph_format.left_indent = Cm(7.5)
    p_ori.paragraph_format.space_after = Pt(105)
    r_ori = p_ori.add_run("Prof. Orientador: Wesley Pereira da Silva.")
    r_ori.bold = True
    r_ori.font.size = Pt(10)
    r_ori.font.name = "Arial"

    for txt_l in ["VITÓRIA", "2026"]:
        p_loc = doc.add_paragraph()
        p_loc.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_loc.paragraph_format.space_after = Pt(2)
        r_loc = p_loc.add_run(txt_l)
        r_loc.bold = True
        r_loc.font.size = Pt(12)
        r_loc.font.name = "Arial"

    doc.add_page_break()

    # --- SUMÁRIO CLICÁVEL ---
    p_sum = doc.add_paragraph()
    p_sum.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sum.paragraph_format.space_after = Pt(18)
    r_sum = p_sum.add_run("SUMÁRIO")
    r_sum.bold = True
    r_sum.font.size = Pt(12)
    r_sum.font.name = "Arial"

    itens_sumario = [
        ("1 INTRODUÇÃO", "1 INTRODUÇÃO", "4"),
        ("  1.1 O PROBLEMA", "1.1 O PROBLEMA", "4"),
        ("  1.2 FORMULAÇÃO DO PROBLEMA", "1.2 FORMULAÇÃO DO PROBLEMA", "4"),
        ("  1.3 HIPÓTESE", "1.3 HIPÓTESE", "4"),
        ("  1.4 OBJETIVOS", "1.4 OBJETIVOS", "4"),
        ("  1.5 JUSTIFICATIVA", "1.5 JUSTIFICATIVA", "5"),
        ("  1.6 DELIMITAÇÃO DO ESCOPO: RECORTE NOS PROJETOS DE OBRAS DA PMV", "1.6 DELIMITAÇÃO DO ESCOPO: RECORTE NOS PROJETOS DE OBRAS DA PREFEITURA DE VITÓRIA", "5"),
        ("2 REFERENCIAL TEÓRICO", "2 REFERENCIAL TEÓRICO", "6"),
        ("  2.1 LICITAÇÕES PÚBLICAS, TERMOS ADITIVOS E A LEI Nº 14.133/2021", "2.1 LICITAÇÕES PÚBLICAS, TERMOS ADITIVOS E A LEI Nº 14.133/2021", "6"),
        ("  2.2 DADOS ABERTOS GOVERNAMENTAIS E TRANSPARÊNCIA PÚBLICA", "2.2 DADOS ABERTOS GOVERNAMENTAIS E TRANSPARÊNCIA PÚBLICA", "6"),
        ("  2.3 ENGENHARIA DE DADOS E PROCESSOS ETL", "2.3 ENGENHARIA DE DADOS E PROCESSOS ETL", "6"),
        ("  2.4 MODELAGEM RELACIONAL E CRUZAMENTO DE DADOS PÚBLICOS", "2.4 MODELAGEM RELACIONAL E CRUZAMENTO DE DADOS PÚBLICOS", "7"),
        ("  2.5 BUSINESS INTELLIGENCE E ANÁLISE QUANTITATIVA DE RISCO", "2.5 BUSINESS INTELLIGENCE E ANÁLISE QUANTITATIVA DE RISCO", "7"),
        ("3 METODOLOGIA, ORIGEM DOS DADOS E LÓGICA DAS ANÁLISES", "3 METODOLOGIA, ORIGEM DOS DADOS E LÓGICA DAS ANÁLISES", "7"),
        ("  3.1 NATUREZA DA PESQUISA E DELIMITAÇÃO AOS PROJETOS DE OBRAS DA PMV", "3.1 NATUREZA DA PESQUISA E DELIMITAÇÃO AOS PROJETOS DE OBRAS DA PMV", "7"),
        ("  3.2 ORIGEM DOS DADOS E FONTES UTILIZADAS", "3.2 ORIGEM DOS DADOS E FONTES UTILIZADAS", "8"),
        ("  3.3 PIPELINE DE COLETA, HIGIENIZAÇÃO E INTEGRIDADE (ETL)", "3.3 PIPELINE DE COLETA, HIGIENIZAÇÃO E INTEGRIDADE (ETL)", "8"),
        ("  3.4 LÓGICA DAS ANÁLISES E REGRAS DE CÁLCULO", "3.4 LÓGICA DAS ANÁLISES E REGRAS DE CÁLCULO", "9"),
        ("  3.5 SÍNTESE DOS RESULTADOS QUANTITATIVOS OBTIDOS", "3.5 SÍNTESE DOS RESULTADOS QUANTITATIVOS OBTIDOS", "10"),
        ("  3.6 DESENVOLVIMENTO, DOCUMENTAÇÃO E GUIA DE USO DO PAINEL INTERATIVO", "3.6 DESENVOLVIMENTO, DOCUMENTAÇÃO E GUIA DE USO DO PAINEL INTERATIVO (DASHBOARD WEB)", "11"),
        ("4 CRONOGRAMA DE EXECUÇÃO (2026)", "4 CRONOGRAMA DE EXECUÇÃO (2026)", "16"),
        ("REFERÊNCIAS", "REFERÊNCIAS", "17"),
        ("ANEXO A — REGISTRO DE SOLICITAÇÃO TÉCNICA DE DADOS ABERTOS AO TCE-ES", "ANEXO A — REGISTRO DE SOLICITAÇÃO TÉCNICA DE DADOS ABERTOS AO TCE-ES", "18"),
    ]
    for item_label, target_heading, pg_num in itens_sumario:
        p = doc.add_paragraph()
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        anchor = make_anchor_id(target_heading)
        dots = "." * max(4, 78 - len(item_label) - len(pg_num))
        line_str = f"{item_label} {dots} {pg_num}"
        add_toc_hyperlink_docx(p, anchor, line_str, is_bold=(not item_label.startswith("  ")))

    doc.add_page_break()

    bm_counter = 100

    # --- SEÇÕES 1, 2 E 3 ---
    for sec_title, subsecs in SECOES:
        p_sec = doc.add_paragraph(style='Heading 1')
        p_sec.paragraph_format.space_before = Pt(14)
        p_sec.paragraph_format.space_after = Pt(8)
        r_sec = p_sec.add_run(sec_title)
        r_sec.bold = True
        r_sec.font.size = Pt(12)
        r_sec.font.name = "Arial"
        r_sec.font.color.rgb = RGBColor(0, 0, 0)
        add_bookmark_docx(p_sec, make_anchor_id(sec_title), bm_counter)
        bm_counter += 1

        for sub_title, pars in subsecs:
            if sub_title == "TABELA_FONTES":
                add_formatted_table_docx(
                    doc,
                    TABELA_FONTES_DATA,
                    "Tabela 1 — Síntese das Fontes de Dados Abertos Integradas no Projeto (Vitória/ES)"
                )
                continue
            elif sub_title == "TABELA_RESUMO_GRUPOS":
                add_formatted_table_docx(
                    doc,
                    TABELA_GRUPOS_DATA,
                    "Tabela 2 — Resumo Quantitativo por Grupo de Contratos da PMV (2025–2026)"
                )
                p_nota = doc.add_paragraph()
                p_nota.paragraph_format.space_after = Pt(6)
                r_nota = p_nota.add_run(
                    "*Nota Técnica: O Contrato nº 300/2025 (Ensinatech Tecnologia Ltda) registra Valor Inicial de R$ 2.139.480,00 "
                    "e Aditivo de R$ 2.139.480.000,00 (+100.000,00%), tratando-se de erro material de preenchimento na fonte municipal. "
                    "Por rigor metodológico, os indicadores globais são reportados com e sem esse registro."
                )
                r_nota.font.size = Pt(9)
                r_nota.italic = True
                continue
            elif sub_title == "TEXTO_RESULTADOS_OBRAS":
                for par_txt in pars:
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    p.paragraph_format.first_line_indent = Cm(1.25)
                    p.paragraph_format.line_spacing = 1.35
                    p.paragraph_format.space_after = Pt(6)
                    r = p.add_run(par_txt)
                    r.font.name = "Arial"
                    r.font.size = Pt(11)
                continue
            elif sub_title == "TABELA_SUBCATEGORIAS":
                add_formatted_table_docx(
                    doc,
                    TABELA_SUBCAT_DATA,
                    "Tabela 3 — Classificação dos 101 Contratos de Obras por Subcategoria Real do Objeto (2025–2026)"
                )
                add_formatted_table_docx(
                    doc,
                    TABELA_ACIMA_LIMITE_DATA,
                    "Tabela 4 — Contratos Classificados como Obra com Percentual de Aditivo Acima do Limite Legal (Art. 125 da Lei nº 14.133/2021)"
                )
                add_formatted_table_docx(
                    doc,
                    TABELA_EMPRESAS_OBRAS_DATA,
                    "Tabela 5 — Ranking das 10 Principais Empresas por Recorrência de Contratos de Obras na PMV (2025–2026)"
                )
                continue
            elif sub_title == "IMAGENS_PAINEL":
                for vis_title, vis_pars, img_path, fig_cap in PAINEL_FIGURAS:
                    p_vt = doc.add_paragraph()
                    p_vt.paragraph_format.space_before = Pt(10)
                    p_vt.paragraph_format.space_after = Pt(4)
                    r_vt = p_vt.add_run(vis_title)
                    r_vt.bold = True
                    r_vt.font.size = Pt(11)
                    r_vt.font.name = "Arial"

                    for vp in vis_pars:
                        p_vp = doc.add_paragraph()
                        p_vp.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                        p_vp.paragraph_format.first_line_indent = Cm(1.25)
                        p_vp.paragraph_format.line_spacing = 1.35
                        p_vp.paragraph_format.space_after = Pt(5)
                        r_vp = p_vp.add_run(vp)
                        r_vp.font.name = "Arial"
                        r_vp.font.size = Pt(10.5)

                    if os.path.exists(img_path):
                        p_img = doc.add_paragraph()
                        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_img.paragraph_format.space_before = Pt(6)
                        p_img.paragraph_format.space_after = Pt(2)
                        run_img = p_img.add_run()
                        run_img.add_picture(img_path, width=Cm(16.0))

                        p_fc = doc.add_paragraph()
                        p_fc.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        p_fc.paragraph_format.space_after = Pt(12)
                        r_fc = p_fc.add_run(fig_cap)
                        r_fc.font.size = Pt(9.5)
                        r_fc.italic = True
                        r_fc.font.name = "Arial"
                continue

            # Subtítulo normal com Bookmark para o Sumário Clicável
            p_sub = doc.add_paragraph(style='Heading 2')
            p_sub.paragraph_format.space_before = Pt(10)
            p_sub.paragraph_format.space_after = Pt(6)
            r_sub = p_sub.add_run(sub_title)
            r_sub.bold = True
            r_sub.font.size = Pt(11)
            r_sub.font.name = "Arial"
            r_sub.font.color.rgb = RGBColor(0, 0, 0)
            add_bookmark_docx(p_sub, make_anchor_id(sub_title), bm_counter)
            bm_counter += 1

            for par_txt in pars:
                p = doc.add_paragraph()
                is_sub3 = par_txt.startswith("1.4.") or par_txt.startswith("3.6.")
                is_bullet_or_sub = par_txt.startswith("•") or par_txt.startswith("   •") or is_sub3 or (len(par_txt) > 2 and par_txt[0].isdigit() and par_txt[1] == ".")
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                if not is_bullet_or_sub:
                    p.paragraph_format.first_line_indent = Cm(1.25)
                p.paragraph_format.line_spacing = 1.35
                p.paragraph_format.space_after = Pt(6)
                r = p.add_run(par_txt)
                r.font.name = "Arial"
                r.font.size = Pt(11)
                if is_sub3:
                    r.bold = True

    # --- SEÇÃO 4: CRONOGRAMA ---
    sec4_title = "4 CRONOGRAMA DE EXECUÇÃO (2026)"
    p_sec4 = doc.add_paragraph(style='Heading 1')
    p_sec4.paragraph_format.space_before = Pt(14)
    p_sec4.paragraph_format.space_after = Pt(8)
    r_sec4 = p_sec4.add_run(sec4_title)
    r_sec4.bold = True
    r_sec4.font.size = Pt(12)
    r_sec4.font.name = "Arial"
    r_sec4.font.color.rgb = RGBColor(0, 0, 0)
    add_bookmark_docx(p_sec4, make_anchor_id(sec4_title), bm_counter)
    bm_counter += 1

    p_c_intro = doc.add_paragraph()
    p_c_intro.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p_c_intro.paragraph_format.first_line_indent = Cm(1.25)
    p_c_intro.paragraph_format.space_after = Pt(6)
    p_c_intro.add_run(
        "A Tabela 6 apresenta o cronograma físico de execução das etapas do Trabalho de Conclusão de Curso "
        "ao longo do ano letivo de 2026, contemplando desde a definição do problema e coleta das bases governamentais "
        "até a publicação do painel analítico e redação final da monografia."
    )
    add_formatted_table_docx(doc, CRONOGRAMA_DATA, "Tabela 6 — Cronograma de Execução das Etapas do Projeto (2026)")

    # --- REFERÊNCIAS ---
    doc.add_page_break()
    ref_title = "REFERÊNCIAS"
    p_ref_t = doc.add_paragraph(style='Heading 1')
    p_ref_t.paragraph_format.space_before = Pt(10)
    p_ref_t.paragraph_format.space_after = Pt(12)
    r_ref_t = p_ref_t.add_run(ref_title)
    r_ref_t.bold = True
    r_ref_t.font.size = Pt(12)
    r_ref_t.font.name = "Arial"
    r_ref_t.font.color.rgb = RGBColor(0, 0, 0)
    add_bookmark_docx(p_ref_t, make_anchor_id(ref_title), bm_counter)
    bm_counter += 1

    for ref in REFERENCIAS_LIST:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(ref)
        r.font.name = "Arial"
        r.font.size = Pt(10.5)

    # --- ANEXO A ---
    doc.add_page_break()
    anx_title = "ANEXO A — REGISTRO DE SOLICITAÇÃO TÉCNICA DE DADOS ABERTOS AO TCE-ES"
    p_anx = doc.add_paragraph(style='Heading 1')
    p_anx.paragraph_format.space_before = Pt(10)
    p_anx.paragraph_format.space_after = Pt(8)
    r_anx = p_anx.add_run(anx_title)
    r_anx.bold = True
    r_anx.font.size = Pt(12)
    r_anx.font.name = "Arial"
    r_anx.font.color.rgb = RGBColor(0, 0, 0)
    add_bookmark_docx(p_anx, make_anchor_id(anx_title), bm_counter)

    if os.path.exists("img_painel/anexo_email.png"):
        p_aimg = doc.add_paragraph()
        p_aimg.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_aimg.add_run().add_picture("img_painel/anexo_email.png", width=Cm(15.5))

    doc.save(output_docx)
    print(f"DOCX salvo com sucesso (com Sumário Clicável): {output_docx}")


# ==============================================================================
# 2. GERADOR DO DOCUMENTO OFICIAL DO TCC EM PDF (.PDF) COM SUMÁRIO CLICÁVEL
# ==============================================================================

def header_footer_tcc(canvas, doc):
    canvas.saveState()
    if doc.page > 3:
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.black)
        canvas.drawString(2.2 * cm, A4[1] - 1.3 * cm, "FAESA — Engenharia da Computação | Predição de Sobrecusto em Licitações e Auditoria de Contratos (Vitória/ES)")
        canvas.drawRightString(A4[0] - 2.0 * cm, A4[1] - 1.3 * cm, f"{doc.page}")
        canvas.setStrokeColor(colors.HexColor("#CBD5E1"))
        canvas.setLineWidth(0.5)
        canvas.line(2.2 * cm, A4[1] - 1.45 * cm, A4[0] - 2.0 * cm, A4[1] - 1.45 * cm)
    canvas.restoreState()

def build_rl_table(data, col_widths, styles):
    cell_header = ParagraphStyle("TH", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=7.2, leading=9, textColor=colors.white, alignment=TA_CENTER)
    cell_body_l = ParagraphStyle("TBL", parent=styles["Normal"], fontName="Helvetica", fontSize=7.2, leading=9.2, textColor=colors.black, alignment=TA_LEFT)
    cell_body_r = ParagraphStyle("TBR", parent=styles["Normal"], fontName="Helvetica", fontSize=7.2, leading=9.2, textColor=colors.black, alignment=TA_RIGHT)
    cell_body_c = ParagraphStyle("TBC", parent=styles["Normal"], fontName="Helvetica", fontSize=7.2, leading=9.2, textColor=colors.black, alignment=TA_CENTER)

    formatted_rows = []
    for r_idx, row in enumerate(data):
        f_row = []
        for c_idx, val in enumerate(row):
            s_val = str(val)
            if r_idx == 0:
                f_row.append(Paragraph(s_val, cell_header))
            else:
                if c_idx in (0, 1, 2) and len(s_val) > 12:
                    f_row.append(Paragraph(s_val, cell_body_l))
                elif any(ch.isdigit() for ch in s_val) and s_val != "X":
                    f_row.append(Paragraph(s_val, cell_body_r))
                else:
                    f_row.append(Paragraph(s_val, cell_body_c))
        formatted_rows.append(f_row)

    t = RLTable(formatted_rows, colWidths=col_widths, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E293B")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))
    return t

def gerar_pdf_oficial_tcc(output_pdf):
    from preparar_doc_tcc import SECOES
    doc = SimpleDocTemplate(
        output_pdf,
        pagesize=A4,
        leftMargin=2.2 * cm,
        rightMargin=2.0 * cm,
        topMargin=2.0 * cm,
        bottomMargin=2.0 * cm,
        title="PREDIÇÃO DE SOBRECUSTO EM LICITAÇÕES DE PROJETOS GOVERNAMENTAIS DO ESTADO DO ESPÍRITO SANTO",
        author="Camila Ramos Bassete, Leonardo Bandeira Matias, Matheus Gomes Evangelista"
    )
    styles = getSampleStyleSheet()
    st_center_b = ParagraphStyle("CB", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=12, leading=16, alignment=TA_CENTER, textColor=colors.black)
    st_title_cover = ParagraphStyle("TC", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=14, leading=19, alignment=TA_CENTER, textColor=colors.black)
    st_nat = ParagraphStyle("NAT", parent=styles["Normal"], fontName="Helvetica", fontSize=9.5, leading=13.5, leftIndent=6.5*cm, alignment=TA_JUSTIFY, textColor=colors.black)
    st_h1 = ParagraphStyle("H1", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=12, leading=16, spaceBefore=12, spaceAfter=6, textColor=colors.black)
    st_h2 = ParagraphStyle("H2", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10.5, leading=14, spaceBefore=9, spaceAfter=4, textColor=colors.black)
    st_body = ParagraphStyle("BD", parent=styles["Normal"], fontName="Helvetica", fontSize=10, leading=14.2, firstLineIndent=1.0*cm, spaceAfter=5, alignment=TA_JUSTIFY, textColor=colors.black)
    st_bullet = ParagraphStyle("BL", parent=styles["Normal"], fontName="Helvetica", fontSize=9.8, leading=13.8, leftIndent=0.5*cm, spaceAfter=4, alignment=TA_JUSTIFY, textColor=colors.black)
    st_cap = ParagraphStyle("CP", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9, leading=12, spaceBefore=6, spaceAfter=3, textColor=colors.black)
    st_fig_cap = ParagraphStyle("FC", parent=styles["Normal"], fontName="Helvetica-Oblique", fontSize=8.5, leading=11.5, spaceBefore=3, spaceAfter=10, alignment=TA_CENTER, textColor=colors.black)
    st_note = ParagraphStyle("NT", parent=styles["Normal"], fontName="Helvetica-Oblique", fontSize=8.2, leading=11, spaceBefore=3, spaceAfter=6, textColor=colors.black)

    story = []

    # CAPA
    story.append(Spacer(1, 1.0 * cm))
    story.append(Paragraph("FAESA CENTRO UNIVERSITÁRIO", st_center_b))
    story.append(Paragraph("UNIDADE DE EDUCAÇÃO TECNOLÓGICA E COMPUTAÇÃO", st_center_b))
    story.append(Spacer(1, 3.0 * cm))
    story.append(Paragraph("CAMILA RAMOS BASSETE<br/>LEONARDO BANDEIRA MATIAS<br/>MATHEUS GOMES EVANGELISTA", st_center_b))
    story.append(Spacer(1, 4.2 * cm))
    story.append(Paragraph("PREDIÇÃO DE SOBRECUSTO EM LICITAÇÕES DE PROJETOS GOVERNAMENTAIS DO ESTADO DO ESPÍRITO SANTO", st_title_cover))
    story.append(Spacer(1, 8.5 * cm))
    story.append(Paragraph("VITÓRIA<br/>2026", st_center_b))
    story.append(PageBreak())

    # FOLHA DE ROSTO
    story.append(Spacer(1, 1.5 * cm))
    story.append(Paragraph("CAMILA RAMOS BASSETE<br/>LEONARDO BANDEIRA MATIAS<br/>MATHEUS GOMES EVANGELISTA", st_center_b))
    story.append(Spacer(1, 3.5 * cm))
    story.append(Paragraph("PREDIÇÃO DE SOBRECUSTO EM LICITAÇÕES DE PROJETOS GOVERNAMENTAIS DO ESTADO DO ESPÍRITO SANTO", st_title_cover))
    story.append(Spacer(1, 2.5 * cm))
    story.append(Paragraph(
        "Trabalho de Conclusão de Curso apresentado ao Curso de Engenharia da Computação do Centro Universitário FAESA, "
        "como requisito parcial para obtenção do grau de Bacharel em Engenharia da Computação.<br/><br/>"
        "<b>Prof. Orientador:</b> Wesley Pereira da Silva.",
        st_nat
    ))
    story.append(Spacer(1, 7.5 * cm))
    story.append(Paragraph("VITÓRIA<br/>2026", st_center_b))
    story.append(PageBreak())

    # SUMÁRIO CLICÁVEL (links internos <a href="#anchor">)
    story.append(Paragraph("SUMÁRIO", st_center_b))
    story.append(Spacer(1, 0.6 * cm))
    sum_rows = [
        ("1 INTRODUÇÃO", "1 INTRODUÇÃO", "4"),
        ("   1.1 O PROBLEMA", "1.1 O PROBLEMA", "4"),
        ("   1.2 FORMULAÇÃO DO PROBLEMA", "1.2 FORMULAÇÃO DO PROBLEMA", "4"),
        ("   1.3 HIPÓTESE", "1.3 HIPÓTESE", "4"),
        ("   1.4 OBJETIVOS (GERAL E ESPECÍFICOS)", "1.4 OBJETIVOS", "4"),
        ("   1.5 JUSTIFICATIVA", "1.5 JUSTIFICATIVA", "5"),
        ("   1.6 DELIMITAÇÃO DO ESCOPO: RECORTE NOS PROJETOS DE OBRAS DA PMV", "1.6 DELIMITAÇÃO DO ESCOPO: RECORTE NOS PROJETOS DE OBRAS DA PREFEITURA DE VITÓRIA", "5"),
        ("2 REFERENCIAL TEÓRICO", "2 REFERENCIAL TEÓRICO", "6"),
        ("   2.1 LICITAÇÕES PÚBLICAS, TERMOS ADITIVOS E A LEI Nº 14.133/2021", "2.1 LICITAÇÕES PÚBLICAS, TERMOS ADITIVOS E A LEI Nº 14.133/2021", "6"),
        ("   2.2 DADOS ABERTOS GOVERNAMENTAIS E TRANSPARÊNCIA PÚBLICA", "2.2 DADOS ABERTOS GOVERNAMENTAIS E TRANSPARÊNCIA PÚBLICA", "6"),
        ("   2.3 ENGENHARIA DE DADOS E PROCESSOS ETL", "2.3 ENGENHARIA DE DADOS E PROCESSOS ETL", "6"),
        ("   2.4 MODELAGEM RELACIONAL E CRUZAMENTO DE DADOS PÚBLICOS", "2.4 MODELAGEM RELACIONAL E CRUZAMENTO DE DADOS PÚBLICOS", "7"),
        ("   2.5 BUSINESS INTELLIGENCE E ANÁLISE QUANTITATIVA DE RISCO", "2.5 BUSINESS INTELLIGENCE E ANÁLISE QUANTITATIVA DE RISCO", "7"),
        ("3 METODOLOGIA, ORIGEM DOS DADOS E LÓGICA DAS ANÁLISES", "3 METODOLOGIA, ORIGEM DOS DADOS E LÓGICA DAS ANÁLISES", "7"),
        ("   3.1 NATUREZA DA PESQUISA E DELIMITAÇÃO AOS PROJETOS DE OBRAS DA PMV", "3.1 NATUREZA DA PESQUISA E DELIMITAÇÃO AOS PROJETOS DE OBRAS DA PMV", "7"),
        ("   3.2 ORIGEM DOS DADOS E FONTES UTILIZADAS (TABELA 1)", "3.2 ORIGEM DOS DADOS E FONTES UTILIZADAS", "8"),
        ("   3.3 PIPELINE DE COLETA, HIGIENIZAÇÃO E INTEGRIDADE (ETL)", "3.3 PIPELINE DE COLETA, HIGIENIZAÇÃO E INTEGRIDADE (ETL)", "8"),
        ("   3.4 LÓGICA DAS ANÁLISES E REGRAS DE CÁLCULO", "3.4 LÓGICA DAS ANÁLISES E REGRAS DE CÁLCULO", "9"),
        ("   3.5 SÍNTESE DOS RESULTADOS QUANTITATIVOS OBTIDOS (TABELAS 2 A 5)", "3.5 SÍNTESE DOS RESULTADOS QUANTITATIVOS OBTIDOS", "10"),
        ("   3.6 DESENVOLVIMENTO, DOCUMENTAÇÃO E GUIA DE USO DO PAINEL INTERATIVO", "3.6 DESENVOLVIMENTO, DOCUMENTAÇÃO E GUIA DE USO DO PAINEL INTERATIVO (DASHBOARD WEB)", "11"),
        ("4 CRONOGRAMA DE EXECUÇÃO (2026)", "4 CRONOGRAMA DE EXECUÇÃO (2026)", "16"),
        ("REFERÊNCIAS", "REFERÊNCIAS", "17"),
        ("ANEXO A — REGISTRO DE SOLICITAÇÃO TÉCNICA DE DADOS ABERTOS AO TCE-ES", "ANEXO A — REGISTRO DE SOLICITAÇÃO TÉCNICA DE DADOS ABERTOS AO TCE-ES", "18"),
    ]
    st_sum_l = ParagraphStyle("SL", parent=styles["Normal"], fontName="Helvetica", fontSize=9.5, leading=15, textColor=colors.black)
    st_sum_b = ParagraphStyle("SB", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9.5, leading=15, textColor=colors.black)
    st_sum_r = ParagraphStyle("SR", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9.5, leading=15, alignment=TA_RIGHT, textColor=colors.black)
    t_sum_data = []
    for txt_s, target_h, pg_s in sum_rows:
        st_use = st_sum_l if txt_s.startswith("   ") else st_sum_b
        anc = make_anchor_id(target_h)
        disp_txt = txt_s.replace("   ", "&nbsp;&nbsp;&nbsp;&nbsp;")
        t_sum_data.append([
            Paragraph(f'<a href="#{anc}">{disp_txt}</a>', st_use),
            Paragraph(f'<a href="#{anc}">{pg_s}</a>', st_sum_r)
        ])
    t_sum = RLTable(t_sum_data, colWidths=[(A4[0]-4.2*cm)*0.88, (A4[0]-4.2*cm)*0.12])
    t_sum.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (-1, -1), 0.3, colors.HexColor("#E2E8F0")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(t_sum)
    story.append(PageBreak())

    # CONTEÚDO DAS SEÇÕES (com âncoras <a name="..."/> em cada título)
    avail_w = A4[0] - 4.2 * cm
    for sec_title, subsecs in SECOES:
        anc_sec = make_anchor_id(sec_title)
        story.append(Paragraph(f'<a name="{anc_sec}"/>{sec_title}', st_h1))
        for sub_title, pars in subsecs:
            if sub_title == "TABELA_FONTES":
                story.append(Paragraph("Tabela 1 — Síntese das Fontes de Dados Abertos Integradas no Projeto (Vitória/ES)", st_cap))
                story.append(build_rl_table(TABELA_FONTES_DATA, [avail_w*0.23, avail_w*0.18, avail_w*0.11, avail_w*0.20, avail_w*0.28], styles))
                story.append(Spacer(1, 0.25*cm))
                continue
            elif sub_title == "TABELA_RESUMO_GRUPOS":
                story.append(Paragraph("Tabela 2 — Resumo Quantitativo por Grupo de Contratos da PMV (2025–2026)", st_cap))
                story.append(build_rl_table(TABELA_GRUPOS_DATA, [avail_w*0.20, avail_w*0.06, avail_w*0.07, avail_w*0.07, avail_w*0.13, avail_w*0.13, avail_w*0.13, avail_w*0.07, avail_w*0.07, avail_w*0.07], styles))
                story.append(Paragraph(
                    "*Nota Técnica: O Contrato nº 300/2025 (Ensinatech Tecnologia Ltda) registra Valor Inicial de R$ 2.139.480,00 e Aditivo de R$ 2.139.480.000,00 (+100.000,00%), "
                    "tratando-se de erro material de preenchimento na fonte municipal. Por rigor metodológico, os indicadores globais são reportados com e sem esse registro.",
                    st_note
                ))
                continue
            elif sub_title == "TEXTO_RESULTADOS_OBRAS":
                for par_txt in pars:
                    story.append(Paragraph(par_txt, st_body))
                continue
            elif sub_title == "TABELA_SUBCATEGORIAS":
                story.append(Paragraph("Tabela 3 — Classificação dos 101 Contratos de Obras por Subcategoria Real do Objeto (2025–2026)", st_cap))
                story.append(build_rl_table(TABELA_SUBCAT_DATA, [avail_w*0.31, avail_w*0.05, avail_w*0.06, avail_w*0.06, avail_w*0.06, avail_w*0.13, avail_w*0.12, avail_w*0.13, avail_w*0.08], styles))
                story.append(Spacer(1, 0.2*cm))
                story.append(Paragraph("Tabela 4 — Contratos Classificados como Obra com Percentual de Aditivo Acima do Limite Legal (Art. 125 da Lei nº 14.133/2021)", st_cap))
                story.append(build_rl_table(TABELA_ACIMA_LIMITE_DATA, [avail_w*0.08, avail_w*0.24, avail_w*0.20, avail_w*0.11, avail_w*0.11, avail_w*0.11, avail_w*0.07, avail_w*0.08], styles))
                story.append(Spacer(1, 0.2*cm))
                story.append(Paragraph("Tabela 5 — Ranking das 10 Principais Empresas por Recorrência de Contratos de Obras na PMV (2025–2026)", st_cap))
                story.append(build_rl_table(TABELA_EMPRESAS_OBRAS_DATA, [avail_w*0.05, avail_w*0.24, avail_w*0.12, avail_w*0.07, avail_w*0.08, avail_w*0.06, avail_w*0.11, avail_w*0.10, avail_w*0.11, avail_w*0.06], styles))
                story.append(Spacer(1, 0.25*cm))
                continue
            elif sub_title == "IMAGENS_PAINEL":
                for vis_title, vis_pars, img_path, fig_cap in PAINEL_FIGURAS:
                    story.append(Paragraph(f"<b>{vis_title}</b>", st_h2))
                    for vp in vis_pars:
                        story.append(Paragraph(vp, st_body))
                    if os.path.exists(img_path):
                        img_obj = RLImage(img_path, width=16.2*cm, height=8.2*cm, kind='proportional')
                        story.append(KeepTogether([
                            Spacer(1, 0.15*cm),
                            img_obj,
                            Paragraph(fig_cap, st_fig_cap)
                        ]))
                continue

            anc_sub = make_anchor_id(sub_title)
            story.append(Paragraph(f'<a name="{anc_sub}"/>{sub_title}', st_h2))
            for par_txt in pars:
                is_sub3 = par_txt.startswith("1.4.") or par_txt.startswith("3.6.")
                is_bullet_or_sub = par_txt.startswith("•") or par_txt.startswith("   •") or is_sub3 or (len(par_txt) > 2 and par_txt[0].isdigit() and par_txt[1] == ".")
                if is_sub3:
                    story.append(Paragraph(f"<b>{par_txt}</b>", st_bullet))
                elif is_bullet_or_sub:
                    story.append(Paragraph(par_txt, st_bullet))
                else:
                    story.append(Paragraph(par_txt, st_body))

    # SEÇÃO 4: CRONOGRAMA
    sec4_t = "4 CRONOGRAMA DE EXECUÇÃO (2026)"
    story.append(Paragraph(f'<a name="{make_anchor_id(sec4_t)}"/>{sec4_t}', st_h1))
    story.append(Paragraph(
        "A Tabela 6 apresenta o cronograma físico de execução das etapas do Trabalho de Conclusão de Curso ao longo do ano letivo de 2026, "
        "contemplando desde a definição do problema e coleta das bases governamentais até a publicação do painel analítico e redação final da monografia.",
        st_body
    ))
    story.append(Paragraph("Tabela 6 — Cronograma de Execução das Etapas do Projeto (2026)", st_cap))
    story.append(build_rl_table(CRONOGRAMA_DATA, [avail_w*0.34] + [avail_w*0.06]*11, styles))

    # REFERÊNCIAS
    story.append(PageBreak())
    ref_t = "REFERÊNCIAS"
    story.append(Paragraph(f'<a name="{make_anchor_id(ref_t)}"/>{ref_t}', st_h1))
    for ref in REFERENCIAS_LIST:
        story.append(Paragraph(ref, st_bullet))

    # ANEXO A
    if os.path.exists("img_painel/anexo_email.png"):
        story.append(PageBreak())
        anx_t = "ANEXO A — REGISTRO DE SOLICITAÇÃO TÉCNICA DE DADOS ABERTOS AO TCE-ES"
        story.append(Paragraph(f'<a name="{make_anchor_id(anx_t)}"/>{anx_t}', st_h1))
        story.append(Spacer(1, 0.2*cm))
        story.append(RLImage("img_painel/anexo_email.png", width=15.5*cm, height=19.0*cm, kind='proportional'))

    doc.build(story, onFirstPage=header_footer_tcc, onLaterPages=header_footer_tcc)

    # Pós-processamento com PyMuPDF: expandir área clicável para toda a linha do Sumário, tornar todas as URLs clicáveis e criar Marcadores (Outline) do PDF
    import pymupdf
    import re
    import shutil
    pdf_doc = pymupdf.open(output_pdf)
    pg_sum = pdf_doc[2]
    links = pg_sum.get_links()
    toc_list = []
    for lk in links:
        r = lk["from"]
        if r.x0 < 200:  # Link do título à esquerda -> estende até a margem direita (x1 = 535)
            lk["from"] = pymupdf.Rect(62, r.y0 - 2, 535, r.y1 + 2)
            pg_sum.update_link(lk)
    for txt_s, _, pg_s in sum_rows:
        lvl = 2 if txt_s.startswith("   ") else 1
        toc_list.append([lvl, txt_s.strip(), int(pg_s)])
    pdf_doc.set_toc(toc_list)

    # Tornar todas as URLs (https://...) clicáveis em qualquer página do PDF
    for page in pdf_doc:
        text_page = page.get_text("text")
        for m_url in re.finditer(r"https?://[^\s\)\>\"']+", text_page):
            raw_url = m_url.group(0).rstrip(".,;)")
            rects = page.search_for(raw_url)
            for rc in rects:
                page.insert_link({"kind": pymupdf.LINK_URI, "from": rc, "uri": raw_url})

    pdf_doc.saveIncr()
    pdf_doc.close()

    shutil.copyfile(output_pdf, "TCC_Oficial_Predicao_Sobrecusto_PMV.pdf")
    print(f"PDF Oficial do TCC salvo com sucesso (com Sumário e URLs Clicáveis): {output_pdf}")


if __name__ == "__main__":
    import shutil
    gerar_docx_oficial("PREDIÇAO_DE_SOBRECUSTO_EM_LICITAÇOES_TCC_ATUALIZADO.docx")
    shutil.copyfile("PREDIÇAO_DE_SOBRECUSTO_EM_LICITAÇOES_TCC_ATUALIZADO.docx", "TCC_Oficial_Predicao_Sobrecusto_PMV.docx")
    gerar_pdf_oficial_tcc("PREDIÇAO_DE_SOBRECUSTO_EM_LICITAÇOES_TCC_ATUALIZADO.pdf")



