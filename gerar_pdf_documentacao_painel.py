import os
import pymupdf
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

OUTPUT_PDF = "Documentacao_e_Guia_de_Apresentacao_Painel_TCC.pdf"


def rodape_guia(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#334155"))
    canvas.drawString(
        1.8 * cm,
        1.1 * cm,
        "Documentação Técnica e Guia de Apresentação do Painel Interativo — TCC Engenharia da Computação (FAESA, 2026)"
    )
    canvas.drawRightString(A4[0] - 1.8 * cm, 1.1 * cm, f"Página {doc.page}")
    canvas.setStrokeColor(colors.HexColor("#CBD5E1"))
    canvas.setLineWidth(0.5)
    canvas.line(1.8 * cm, 1.45 * cm, A4[0] - 1.8 * cm, 1.45 * cm)
    canvas.restoreState()


def make_box(title, bullets, bg_hex="#F8FAFC", border_hex="#CBD5E1", title_color="#0F172A"):
    styles = getSampleStyleSheet()
    st_t = ParagraphStyle("BT", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=9.8, leading=13, textColor=colors.HexColor(title_color), spaceAfter=4)
    st_b = ParagraphStyle("BB", parent=styles["Normal"], fontName="Helvetica", fontSize=9.2, leading=13, textColor=colors.black, spaceAfter=3, alignment=TA_JUSTIFY)

    content = [Paragraph(title, st_t)]
    for b in bullets:
        content.append(Paragraph(b, st_b))

    avail_w = A4[0] - 3.6 * cm
    t = RLTable([[content]], colWidths=[avail_w])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(bg_hex)),
        ("BOX", (0, 0), (-1, -1), 0.8, colors.HexColor(border_hex)),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
    ]))
    return t


def gerar_guia_pdf():
    doc = SimpleDocTemplate(
        OUTPUT_PDF,
        pagesize=A4,
        leftMargin=1.8 * cm,
        rightMargin=1.8 * cm,
        topMargin=1.6 * cm,
        bottomMargin=1.8 * cm,
        title="Documentação Técnica e Guia de Apresentação do Painel Interativo — TCC FAESA",
        author="Camila Ramos Bassete, Leonardo Bandeira Matias, Matheus Gomes Evangelista",
    )

    styles = getSampleStyleSheet()
    st_title = ParagraphStyle("T", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=15, leading=20, textColor=colors.black, spaceAfter=4)
    st_sub = ParagraphStyle("S", parent=styles["Normal"], fontName="Helvetica", fontSize=9.8, leading=14, textColor=colors.HexColor("#334155"), spaceAfter=10)
    st_h1 = ParagraphStyle("H1", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=12, leading=16, textColor=colors.black, spaceBefore=12, spaceAfter=6)
    st_h2 = ParagraphStyle("H2", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=10.5, leading=14, textColor=colors.black, spaceBefore=9, spaceAfter=4)
    st_body = ParagraphStyle("BD", parent=styles["Normal"], fontName="Helvetica", fontSize=9.6, leading=13.8, textColor=colors.black, spaceAfter=5, alignment=TA_JUSTIFY)
    st_bullet = ParagraphStyle("BL", parent=styles["Normal"], fontName="Helvetica", fontSize=9.4, leading=13.5, leftIndent=0.4 * cm, textColor=colors.black, spaceAfter=4, alignment=TA_JUSTIFY)
    st_cap = ParagraphStyle("FC", parent=styles["Normal"], fontName="Helvetica-Oblique", fontSize=8.5, leading=11.5, alignment=TA_CENTER, textColor=colors.HexColor("#334155"), spaceBefore=3, spaceAfter=10)

    story = []
    avail_w = A4[0] - 3.6 * cm

    # CABEÇALHO DO DOCUMENTO
    story.append(Paragraph("DOCUMENTAÇÃO TÉCNICA, MANUAL DE USO E ROTEIRO DE APRESENTAÇÃO DO PAINEL INTERATIVO (DASHBOARD WEB)", st_title))
    story.append(Paragraph(
        "<b>Projeto de TCC (FAESA — Engenharia da Computação):</b> Predição de Sobrecusto em Licitações e Auditoria Quantitativa de Obras da Prefeitura Municipal de Vitória/ES<br/>"
        "<b>Autores:</b> Camila Ramos Bassete, Leonardo Bandeira Matias e Matheus Gomes Evangelista | <b>Orientador:</b> Prof. Wesley Pereira da Silva<br/>"
        "<b>Link Público do Painel Online (GitHub Pages):</b> <a href='https://leobmt02.github.io/tcc-auditoria-obras-vitoria/'><u>https://leobmt02.github.io/tcc-auditoria-obras-vitoria/</u></a>",
        st_sub
    ))

    # SEÇÃO 1: COMO FOI DESENVOLVIDO
    story.append(Paragraph("1. COMO O PAINEL FOI DESENVOLVIDO (ARQUITETURA E ENGENHARIA DE DADOS)", st_h1))
    story.append(Paragraph(
        "O painel interativo foi projetado para permitir que qualquer avaliador, gestor público ou cidadão audite os contratos e obras da Prefeitura de Vitória "
        "diretamente pelo navegador (computador ou celular), com resposta instantânea e sem depender de instalação de programas. "
        "O seu desenvolvimento foi dividido em duas camadas integradas:",
        st_body
    ))

    story.append(Paragraph("1.1 Camada Back-End de Engenharia de Dados (Python)", st_h2))
    for item in [
        "• <b>Coleta e Integração de 4 Fontes Oficiais:</b> Os scripts em Python (<i>consolidar_tudo.py</i> e <i>gerar_dashboard.py</i>) unificam quatro bases abertas: "
        "(1) <b>Contratos Administrativos da PMV 2025–2026</b> (Administração Direta + IPAMV + CDTIV, totalizando 1.065 contratos únicos, sendo 101 de Obras); "
        "(2) <b>Painel Físico/Financeiro de Obras da PMV 2021–2026</b> (179 obras com valores Empenhado, Liquidado, Pago, Bairros e Estágio Físico); "
        "(3) <b>API Oficial do PNCP</b> (1.287 licitações extraídas com valores estimado, homologado e modalidade); e "
        "(4) <b>Dados Abertos do TCE-ES / CidadES</b> (termos aditivos, apostilamentos, medições e indicador de reforma).",
        "• <b>Higienização e Regras de Negócio:</b> O pipeline reconstrói 472 registros que vinham quebrados nos CSVs originais da Prefeitura, calcula a variação financeira (em R$ e %), "
        "aplica a regra do Art. 125 da Lei nº 14.133/2021 (teto de <b>25% para obras em geral</b> e <b>50% para reformas</b>), classifica os objetos das obras em <b>10 Subcategorias Reais</b> "
        "e calcula a recorrência de contratos por CNPJ.",
        "• <b>Serialização JSON Embutida:</b> Ao final do processamento, o Python converte todas as tabelas tratadas em objetos JSON compactados e os injeta diretamente dentro do arquivo HTML final."
    ]:
        story.append(Paragraph(item, st_bullet))

    story.append(Paragraph("1.2 Camada Front-End de Visualização Interativa (HTML5, CSS3 e JavaScript ES6)", st_h2))
    for item in [
        "• <b>Interface Visual em Tema Claro (HTML5 + CSS3):</b> Desenvolvida com paleta clara de alto contraste (fundo branco/gelo, tipografia <i>Inter</i> e cartões com codificação semântica de cores: azul para valores iniciais, laranja para aditivos, verde para pagamentos e vermelho para alertas de limite legal), ideal para projeção em sala de aula e leitura rápida.",
        "• <b>Motor Analítico em Memória (JavaScript ES6 + Chart.js):</b> Quando o usuário abre o link, os 1.065 contratos e 179 obras já estão em memória no navegador. Qualquer filtro acionado recalcula todos os cartões, gráficos (renderizados com a biblioteca <i>Chart.js</i>) e tabelas em menos de 10 milissegundos.",
        "• <b>Publicação Contínua no GitHub Pages:</b> Por ser uma aplicação web autocontida (<i>Single-File Web Application</i>), foi publicada gratuitamente no GitHub Pages do projeto, garantindo disponibilidade de 100% e acesso público imediato."
    ]:
        story.append(Paragraph(item, st_bullet))

    # SEÇÃO 2: MANUAL DE FUNCIONALIDADES E FILTROS
    story.append(Paragraph("2. MANUAL DE FUNCIONALIDADES E CONTROLES INTERATIVOS (COMO USAR)", st_h1))
    story.append(Paragraph(
        "Todos os controles do painel são interativos e cumulativos — ou seja, você pode combinar múltiplos filtros ao mesmo tempo para responder perguntas específicas durante a apresentação:",
        st_body
    ))

    t_filtros_data = [
        ["Controle / Funcionalidade", "Onde Fica", "O que Faz e Como Usar na Prática"],
        ["Botões de Download (PDF, Word e CSV)", "Topo direito do cabeçalho", "Permitem baixar com 1 clique o Documento Oficial do TCC (PDF e Word), o Relatório Técnico (PDF) e a base consolidada das 101 obras (CSV)."],
        ["Seletor de Abas (Visões 1 a 4)", "Cabeçalho superior", "Alterna instantaneamente entre as 4 visões analíticas do sistema sem recarregar a página."],
        ["Filtro: Ano do Contrato", "Barra de filtros", "Isola os contratos assinados apenas em 2025, apenas em 2026 ou exibe os dois anos juntos."],
        ["Filtro: Subcategoria do Objeto", "Barra de filtros", "Segmenta as obras pelas 10 categorias reais (ex.: Contenção de Encostas, Reforma Elétrica, Urbanização, Instalações/Elevadores)."],
        ["Filtro: Empresa / Fornecedor (com Qtd.)", "Barra de filtros", "Lista todas as empresas ordenadas da que possui mais contratos para a que possui menos — ex.: ILUMITERRA (10). Filtra toda a tela para 1 empresa."],
        ["Filtro: Secretaria / Unidade", "Barra de filtros", "Filtra os contratos por órgão municipal (Secretaria de Obras, Educação, Saúde, Central de Serviços, IPAMV ou CDTIV)."],
        ["Filtro: Situação de Aditivo", "Barra de filtros", "Oferece 4 recortes prontos de auditoria: 'Somente com Aditivo (> R$ 0)', 'Aditivo > 25%', 'Acima do Limite Legal (25%/50%)' e 'Com Supressão (< R$ 0)'."],
        ["Filtro: Outlier (300/2025)", "Barra de filtros (Aba 4)", "Alterna entre 'Ocultar Outlier (Recomendado)' e 'Incluir Outlier (R$ 2,13 bi)', mostrando ao vivo o impacto do erro de digitação da Prefeitura."],
        ["Busca Livre em Tempo Real", "Barra de filtros", "Caixa de texto onde você digita qualquer termo (ex.: 'encosta', 'São Pedro', 'elevador', '223/2025') e a tela filtra a cada letra digitada."],
        ["Ordenação das Tabelas + Drill-Down", "Cabeçalho das tabelas", "Clicar no nome de qualquer coluna ordena os valores (maior/menor). Na Aba 2, o botão 'Ver Contratos' abre a lista de contratos da empresa clicada."]
    ]

    st_th = ParagraphStyle("TH", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8, leading=10, textColor=colors.white, alignment=TA_CENTER)
    st_td_b = ParagraphStyle("TDB", parent=styles["Normal"], fontName="Helvetica-Bold", fontSize=8, leading=10.5, textColor=colors.black)
    st_td = ParagraphStyle("TD", parent=styles["Normal"], fontName="Helvetica", fontSize=8, leading=10.5, textColor=colors.black)

    rows_f = []
    for i, r in enumerate(t_filtros_data):
        if i == 0:
            rows_f.append([Paragraph(c, st_th) for c in r])
        else:
            rows_f.append([Paragraph(r[0], st_td_b), Paragraph(r[1], st_td), Paragraph(r[2], st_td)])

    tf = RLTable(rows_f, colWidths=[avail_w * 0.25, avail_w * 0.17, avail_w * 0.58], repeatRows=1)
    tf.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1E293B")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#CBD5E1")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(tf)

    # SEÇÃO 3: EXPLICAÇÃO DE CADA ABA COM IMAGENS E DICAS DE APRESENTAÇÃO
    story.append(PageBreak())
    story.append(Paragraph("3. COMO EXPLICAR CADA VISÃO DO PAINEL NA APRESENTAÇÃO (ABAS 1 A 4)", st_h1))

    # --- ABA 1 ---
    story.append(Paragraph("3.1 Visão 1 — Aba '1. Contratos de Obras (101)': Auditoria de Limites Legais e Subcategorias", st_h2))
    if os.path.exists("img_painel/aba1_obras.png"):
        story.append(RLImage("img_painel/aba1_obras.png", width=17.0 * cm, height=8.4 * cm, kind="proportional"))
        story.append(Paragraph("Figura 1 — Tela da Aba 1: Auditoria dos 101 Contratos de Obras de 2025 e 2026.", st_cap))

    for p in [
        "• <b>O que esta visão representa:</b> É a tela principal do TCC, dedicada ao recorte dos <b>101 contratos classificados como Obras (2025–2026)</b>.",
        "• <b>Leitura dos 5 Cartões KPI no topo:</b> (1) <b>101 contratos</b> assinados por 57 empresas, sendo <b>21 com aditivo positivo (20,8%)</b>; "
        "(2) <b>R$ 545,99 milhões</b> de Valor Inicial Contratado (54,1% de todo o orçamento contratado pelo município no período); "
        "(3) <b>R$ 14,16 milhões</b> em aditivos líquidos (+2,59% no agregado geral, mas com média de +39,37% entre as obras aditadas); "
        "(4) <b>R$ 560,15 milhões</b> de Valor Atualizado Total, exibindo também os <b>R$ 99,73 milhões já empenhados</b> nas obras cruzadas com o Painel de Obras; e "
        "(5) Alerta vermelho para os <b>5 contratos acima do limite legal (25% ou 50%)</b> e <b>R$ 69,57 milhões efetivamente pagos</b>.",
        "• <b>Como explicar o Gráfico da Esquerda (Top Contratos por % de Aditivo vs. Limite Legal):</b> Ordena os contratos com maiores percentuais de aditivo. "
        "Explique que as <b>barras vermelhas</b> representam contratos que ultrapassaram o limite do Art. 125 da Lei 14.133/2021 (ex.: <i>Karisten 223/2025</i> com +150,51%, "
        "<i>Ilumiterra 325/2025</i> e <i>165/2025</i> com +100,00%, <i>Inove Elevadores 293/2025</i> com +85,31% e <i>Matix 2/2025</i> com +49,33%), enquanto as <b>barras azuis</b> "
        "mostram contratos que pararam exatamente no teto legal de 25% (25,00% e 24,94%), evidenciando que o limite legal funciona como uma âncora.",
        "• <b>Como explicar o Gráfico da Direita (Mapeamento por Subcategoria do Objeto):</b> Compara o Valor Inicial (azul) e os Aditivos (laranja) nas 10 categorias reais. "
        "Mostre que os contratos que estouraram o limite de 25% pertencem majoritariamente à subcategoria de <b>Instalações, Tecnologia, Elevadores e Iluminação (+90,77% de aditivo)</b>, "
        "revelando que a Prefeitura cadastrou serviços contínuos de engenharia sob o rótulo de 'Obra'."
    ]:
        story.append(Paragraph(p, st_bullet))

    story.append(Spacer(1, 4))
    story.append(make_box(
        "💡 DICA DO QUE FALAR PARA O PROFESSOR / BANCA NESTA ABA 1:",
        [
            "\"Professor, nesta primeira aba isolamos os 101 contratos de obras de 2025 e 2026, que somam R$ 546 milhões. Ao aplicar os limites do Art. 125 da Nova Lei de Licitações (25% para obras e 50% para reformas), identificamos 5 contratos acima do teto (em vermelho no gráfico à esquerda) e outros 8 estacionados entre 20% e o limite (vários em 24,94% e 25,00%). Porém, graças à nossa categorização semântica no gráfico à direita, descobrimos que 4 dos 5 contratos acima do limite são, na verdade, serviços de manutenção de iluminação pública e elevadores que foram cadastrados como 'Obra' no portal municipal.\""
        ],
        bg_hex="#EFF6FF", border_hex="#93C5FD", title_color="#1E40AF"
    ))

    # --- ABA 2 ---
    story.append(PageBreak())
    story.append(Paragraph("3.2 Visão 2 — Aba '2. Ranking de Empresas & Recorrência': Concentração de Fornecedores", st_h2))
    if os.path.exists("img_painel/aba2_empresas.png"):
        story.append(RLImage("img_painel/aba2_empresas.png", width=17.0 * cm, height=8.4 * cm, kind="proportional"))
        story.append(Paragraph("Figura 2 — Tela da Aba 2: Ranking de Empresas por Quantidade de Contratos, Volume Financeiro e Taxa de Aditamento.", st_cap))

    for p in [
        "• <b>O que esta visão representa:</b> Audita o grau de concentração das licitações de obras entre as construtoras e verifica se empresas com múltiplos contratos têm maior propensão a aditivos.",
        "• <b>Leitura dos 5 Cartões KPI no topo:</b> Mostram que as 101 obras estão distribuídas entre <b>57 empresas</b>, mas <b>21 empresas são recorrentes (possuem 2 ou mais contratos)</b> e concentram sozinhas <b>65 dos 101 contratos (64,4%)</b>. O cartão central destaca a líder em quantidade: <b>Ilumiterra Construções e Montagens Ltda (10 contratos)</b>.",
        "• <b>Como explicar o Gráfico da Esquerda (Empresas com Maior Quantidade de Contratos):</b> Para cada empresa, exibe a <b>barra azul (Total de Contratos assinados)</b> ao lado da <b>barra laranja (Contratos que sofreram Aditivo)</b>. Permite comparar perfis distintos: a <i>Ilumiterra</i> possui 10 contratos (3 aditados); a <i>M.T.F. Construções</i> possui 7 contratos; a <i>Engesan</i> possui 6 contratos; já a <i>Renova Construções</i> possui 3 contratos e <b>todos os 3 sofreram aditivos (100% de recorrência de aditivo)</b>.",
        "• <b>Como explicar o Gráfico da Direita e a Tabela Inferior:</b> O gráfico da direita ordena os maiores volumes financeiros contratados (liderados pelos consórcios de macroinfraestrutura: <i>Nova Beira Mar</i>, <i>Tancredão</i> e <i>Transbordo Vix</i>). Na tabela inferior, ao clicar no botão <b>'Ver Contratos (N)'</b> de qualquer empresa, o sistema filtra automaticamente todos os contratos daquela construtora."
    ]:
        story.append(Paragraph(p, st_bullet))

    story.append(Spacer(1, 4))
    story.append(make_box(
        "💡 DICA DO QUE FALAR PARA O PROFESSOR / BANCA NESTA ABA 2:",
        [
            "\"Na Aba 2 respondemos à questão da recorrência de fornecedores: descobrimos que 21 construtoras detêm 65 dos 101 contratos de obras do município. No gráfico da esquerda, comparamos o total de contratos de cada empresa (azul) com quantos tiveram aditivo (laranja): a Ilumiterra lidera com 10 contratos e R$ 5,38 milhões em aditivos, enquanto empresas como a Renova Construções tiveram aditivos em 100% dos seus contratos (3 de 3). Além disso, clicando em 'Ver Contratos' na tabela, conseguimos auditar instantaneamente todos os contratos de qualquer empresa.\""
        ],
        bg_hex="#EFF6FF", border_hex="#93C5FD", title_color="#1E40AF"
    ))

    # --- ABA 3 ---
    story.append(PageBreak())
    story.append(Paragraph("3.3 Visão 3 — Aba '3. Painel Físico/Financeiro (179)': Execução Orçamentária e Estágio das Obras", st_h2))
    if os.path.exists("img_painel/aba3_painel_fisico.png"):
        story.append(RLImage("img_painel/aba3_painel_fisico.png", width=17.0 * cm, height=8.4 * cm, kind="proportional"))
        story.append(Paragraph("Figura 3 — Tela da Aba 3: Acompanhamento Físico e Financeiro das 179 Obras do Painel Municipal (2021–2026).", st_cap))

    for p in [
        "• <b>O que esta visão representa:</b> Integra a base histórica do <b>Painel de Obras da PMV (179 obras de 2021 a 2026)</b>, trazendo os dados de execução orçamentária (Empenhado, Liquidado e Pago), bairros e estágio físico que não existem na planilha bruta de contratos.",
        "• <b>Leitura dos 5 Cartões KPI no topo:</b> (1) <b>179 obras</b> acompanhadas, destacando que <b>10 estão paralisadas (8) ou rescindidas unilateralmente (2)</b>; "
        "(2) <b>R$ 2,18 bilhões</b> de Valor Atual Contratado; (3) <b>R$ 1,70 bilhão Empenhado (78,0% do contratado)</b>; "
        "(4) <b>R$ 1,526 bilhão Liquidado</b> (serviço já medido e atestado pela fiscalização); e (5) <b>R$ 1,525 bilhão Efetivamente Pago (69,9% do contratado)</b>.",
        "• <b>Como explicar os 2 Gráficos desta Aba:</b> O gráfico de rosca à esquerda mostra a distribuição física das 179 obras (62 em execução, 50 concluídas provisoriamente, 39 concluídas definitivamente, 11 reiniciadas e 8 paralisadas). O gráfico de colunas à direita compara lado a lado o valor <b>Contratado (azul)</b>, <b>Empenhado (laranja)</b> e <b>Pago (verde)</b> dentro de cada estágio de execução, permitindo auditar inclusive quanto já foi pago em obras hoje paralisadas."
    ]:
        story.append(Paragraph(p, st_bullet))

    story.append(Spacer(1, 4))
    story.append(make_box(
        "💡 DICA DO QUE FALAR PARA O PROFESSOR / BANCA NESTA ABA 3:",
        [
            "\"Como a base de contratos da Prefeitura não informa quanto já foi efetivamente gasto em cada obra, integramos na Aba 3 as 179 obras do Painel de Obras (2021–2026), que somam R$ 2,18 bilhões contratados. Aqui conseguimos acompanhar todo o ciclo da despesa pública: R$ 1,70 bilhão empenhado (78%), R$ 1,53 bilhão liquidado após medição técnica e R$ 1,525 bilhão pago (69,9%), além de identificar 8 obras paralisadas e 2 rescindidas unilateralmente.\""
        ],
        bg_hex="#EFF6FF", border_hex="#93C5FD", title_color="#1E40AF"
    ))

    # --- ABA 4 ---
    story.append(PageBreak())
    story.append(Paragraph("3.4 Visão 4 — Aba '4. Todos os Contratos PMV (1.065)': Linha de Base e Demonstração do Outlier", st_h2))
    if os.path.exists("img_painel/aba4_todos_contratos.png"):
        story.append(RLImage("img_painel/aba4_todos_contratos.png", width=17.0 * cm, height=8.4 * cm, kind="proportional"))
        story.append(Paragraph("Figura 4 — Tela da Aba 4: Panorama Geral dos 1.065 Contratos da PMV (2025–2026) com Filtro de Outlier.", st_cap))

    for p in [
        "• <b>O que esta visão representa:</b> Exibe a base completa de <b>1.065 contratos</b> da Prefeitura (incluindo IPAMV e CDTIV) de 2025 e 2026, permitindo comparar o comportamento das Obras frente aos contratos de Prestação de Serviços e Fornecimento de Materiais.",
        "• <b>Demonstração Prática do Tratamento de Outlier (Contrato 300/2025):</b> Com o seletor 'Outlier (300/2025)' na opção padrão (<i>Ocultar Outlier</i>), os 1.064 contratos somam <b>R$ 1,009 bilhão inicial</b> e <b>R$ 107,47 milhões em aditivos (+10,65% global)</b>. Ao alternar para <i>Incluir Outlier</i>, o total de aditivos salta para <b>R$ 2,24 bilhões (+222,16%)</b> por conta do erro de digitação de 1.000x no Contrato 300/2025 (Ensinatech Ltda, R$ 2,139 milhões lançados como R$ 2,139 bilhões de aditivo).",
        "• <b>Comparação Serviços vs. Obras:</b> Os gráficos mostram que nos contratos gerais (especialmente terceirização de mão de obra, limpeza e TI) há 76 contratos com aditivo acima de 25% devido a prorrogações de vigência, justificando por que o nosso estudo isolou os projetos de obras para avaliar sobrecusto de escopo."
    ]:
        story.append(Paragraph(p, st_bullet))

    story.append(Spacer(1, 4))
    story.append(make_box(
        "💡 ROTEIRO PRÁTICO DE 5 PASSOS PARA DEMONSTRAR O PAINEL AO VIVO:",
        [
            "<b>Passo 1 (Aba 1 — Visão Geral das Obras):</b> Abra o painel na Aba 1 e mostre os R$ 545,99 milhões em 101 contratos de obras e os 5 contratos em vermelho que ultrapassaram o limite legal (25%/50%).",
            "<b>Passo 2 (Aba 1 — Filtro de Subcategoria):</b> Clique no filtro 'Situação de Aditivo' -> 'Acima do Limite Legal (25% / 50%)' para a tabela listar exatamente os 5 contratos e mostre como o cruzamento com PNCP, TCE-ES e Painel de Obras aparece na mesma linha.",
            "<b>Passo 3 (Aba 2 — Empresas Recorrentes):</b> Clique na Aba 2, mostre que 21 empresas concentram 65 contratos (lideradas pela Ilumiterra com 10 contratos) e clique no botão 'Ver Contratos (10)' da Ilumiterra para mostrar o drill-down automático.",
            "<b>Passo 4 (Aba 3 — Empenhado e Pago):</b> Clique na Aba 3 e mostre o funil financeiro das 179 obras (R$ 2,18 bi Contratado -> R$ 1,70 bi Empenhado -> R$ 1,53 bi Pago) e filtre as 8 obras 'Paralisadas' no seletor de Estágio.",
            "<b>Passo 5 (Aba 4 — Efeito do Outlier):</b> Clique na Aba 4 e alterne o filtro 'Outlier (300/2025)' de 'Ocultar' para 'Incluir', mostrando ao vivo o salto de R$ 107 milhões (+10,65%) para R$ 2,24 bilhões (+222,16%), encerrando com a importância da higienização de dados."
        ],
        bg_hex="#F0FDF4", border_hex="#86EFAC", title_color="#166534"
    ))

    doc.build(story, onFirstPage=rodape_guia, onLaterPages=rodape_guia)

    # Adicionar índice lateral (Bookmarks) no PDF do Guia
    pdf_doc = pymupdf.open(OUTPUT_PDF)
    pdf_doc.set_toc([
        [1, "1. Como o Painel Foi Desenvolvido (Arquitetura)", 1],
        [1, "2. Manual de Funcionalidades e Filtros (Como Usar)", 1],
        [1, "3. Como Explicar Cada Visão na Apresentação (Abas 1 a 4)", 2],
        [2, "3.1 Aba 1: Contratos de Obras (101)", 2],
        [2, "3.2 Aba 2: Ranking de Empresas e Recorrência", 3],
        [2, "3.3 Aba 3: Painel Físico/Financeiro (179)", 4],
        [2, "3.4 Aba 4: Todos os Contratos (1.065) e Roteiro ao Vivo", 5],
    ])
    pdf_doc.saveIncr()
    pdf_doc.close()
    print(f"PDF de Documentação e Guia gerado com sucesso: {OUTPUT_PDF}")


if __name__ == "__main__":
    gerar_guia_pdf()
