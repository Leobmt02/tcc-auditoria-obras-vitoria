import json
import pandas as pd
import numpy as np

# Carregar dados consolidados
df_all = pd.read_csv(
    "analise_contratos_metricas.csv",
    sep=";",
    decimal=",",
    dtype={"NumeroContrato": str, "AnoContrato": str, "Processo": str, "AnoProcesso": str}
)
df_painel = pd.read_csv("painel_obras_pmv_2021_2026.csv", sep=";", decimal=",")


def clean_records(df, cols):
    sub = df[cols].copy()
    for c in sub.columns:
        if pd.api.types.is_bool_dtype(sub[c]):
            sub[c] = sub[c].fillna(False).astype(bool)
        elif pd.api.types.is_numeric_dtype(sub[c]):
            sub[c] = sub[c].replace([np.inf, -np.inf], np.nan).fillna(0).round(2)
        else:
            sub[c] = sub[c].fillna("").astype(str)
    return sub.to_dict(orient="records")


cols_contratos = [
    "NumeroContrato", "AnoContrato", "TipoContrato", "Subcategoria_Objeto",
    "Processo", "AnoProcesso", "ValorContratado", "Aditivo", "ValorTotal",
    "pct_aditivo", "limite", "excede", "reforma",
    "ValorEmpenhado_PMV", "ValorLiquidado_PMV", "ValorPago_PMV",
    "Estagio_Obra_PMV", "Bairros_Obra_PMV", "Nome_Obra_PMV",
    "CPFCNPJ", "Fornecedor", "Secretaria", "Fiscal",
    "DataInicio", "DataFinal", "Situacao", "Objeto",
    "modalidadeNome", "valorTotalEstimado", "valorTotalHomologado",
    "n_aditivos", "n_prorrogacao", "n_acresc_supress", "n_reequilibrio", "n_apostilamentos", "total_pago"
]

data_contratos = clean_records(df_all, cols_contratos)
data_painel = clean_records(
    df_painel,
    ["Id", "Nome", "Contratos", "Secretaria", "Tipo", "ValorContratado", "Estagio", "Vigencia", "Bairros", "Situacao", "ValorPago", "ValorEmpenhado", "ValorLiquidado"]
)

payload = {
    "contratos": data_contratos,
    "painel_obras": data_painel
}

html_template = r"""<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Auditoria de Contratos e Obras — Prefeitura de Vitória/ES</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap" rel="stylesheet">
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>
<style>
  :root {
    --bg: #f8fafc;
    --bg-card: #ffffff;
    --bg-hover: #f1f5f9;
    --border: #e2e8f0;
    --border-strong: #cbd5e1;
    --text: #0f172a;
    --text-secondary: #334155;
    --text-muted: #64748b;
    --primary: #0284c7;
    --primary-light: #e0f2fe;
    --indigo: #4f46e5;
    --success: #16a34a;
    --success-bg: #dcfce7;
    --warning: #d97706;
    --warning-bg: #fef3c7;
    --danger: #dc2626;
    --danger-bg: #fee2e2;
    --shadow-sm: 0 1px 2px 0 rgba(15, 23, 42, 0.04);
    --shadow: 0 4px 12px -2px rgba(15, 23, 42, 0.06), 0 2px 4px -2px rgba(15, 23, 42, 0.04);
  }
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    background: var(--bg);
    color: var(--text);
    line-height: 1.5;
    padding-bottom: 50px;
    -webkit-font-smoothing: antialiased;
  }
  header {
    background: #ffffff;
    border-bottom: 1px solid var(--border);
    padding: 18px 32px;
    display: flex;
    flex-wrap: wrap;
    justify-content: space-between;
    align-items: center;
    gap: 16px;
    position: sticky;
    top: 0;
    z-index: 50;
    box-shadow: var(--shadow-sm);
  }
  .brand h1 {
    font-size: 1.3rem;
    font-weight: 800;
    letter-spacing: -0.025em;
    color: var(--text);
    display: flex;
    align-items: center;
    gap: 10px;
  }
  .brand p {
    font-size: 0.82rem;
    color: var(--text-muted);
    margin-top: 2px;
    font-weight: 500;
  }
  .badge-pill {
    background: var(--primary-light);
    color: var(--primary);
    border: 1px solid #bae6fd;
    font-size: 0.72rem;
    padding: 3px 10px;
    border-radius: 999px;
    font-weight: 700;
  }
  .tabs {
    display: flex;
    flex-wrap: wrap;
    gap: 6px;
    background: #f1f5f9;
    padding: 5px;
    border-radius: 10px;
    border: 1px solid var(--border);
  }
  .tab-btn {
    background: transparent;
    color: var(--text-secondary);
    border: none;
    padding: 8px 15px;
    border-radius: 7px;
    font-size: 0.82rem;
    font-weight: 600;
    cursor: pointer;
    transition: all 0.18s;
  }
  .tab-btn:hover {
    color: var(--text);
    background: rgba(255,255,255,0.6);
  }
  .tab-btn.active {
    background: #ffffff;
    color: var(--primary);
    box-shadow: 0 1px 3px rgba(15, 23, 42, 0.1);
  }
  .container {
    max-width: 1500px;
    margin: 0 auto;
    padding: 24px 32px;
  }
  .filter-bar {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 16px 20px;
    margin-bottom: 22px;
    display: flex;
    flex-wrap: wrap;
    gap: 14px;
    align-items: flex-end;
    box-shadow: var(--shadow-sm);
  }
  .filter-group {
    display: flex;
    flex-direction: column;
    gap: 5px;
    flex: 1;
    min-width: 165px;
  }
  .filter-group label {
    font-size: 0.71rem;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    color: var(--text-muted);
    font-weight: 700;
  }
  .filter-group select, .filter-group input {
    background: #ffffff;
    border: 1px solid var(--border-strong);
    color: var(--text);
    padding: 8px 11px;
    border-radius: 8px;
    font-size: 0.83rem;
    font-family: inherit;
    outline: none;
    transition: border-color 0.15s, box-shadow 0.15s;
  }
  .filter-group select:focus, .filter-group input:focus {
    border-color: var(--primary);
    box-shadow: 0 0 0 3px rgba(2, 132, 199, 0.12);
  }
  .btn-reset {
    background: #f1f5f9;
    color: var(--text-secondary);
    border: 1px solid var(--border-strong);
    padding: 8px 15px;
    border-radius: 8px;
    cursor: pointer;
    font-size: 0.82rem;
    font-weight: 600;
    height: 36px;
    transition: all 0.15s;
  }
  .btn-reset:hover {
    background: #e2e8f0;
    color: var(--text);
  }
  .active-company-banner {
    background: var(--primary-light);
    border: 1px solid #7dd3fc;
    color: #0369a1;
    padding: 10px 16px;
    border-radius: 10px;
    margin-bottom: 18px;
    display: flex;
    justify-content: space-between;
    align-items: center;
    font-size: 0.85rem;
    font-weight: 600;
  }
  .active-company-banner button {
    background: #0284c7;
    color: #fff;
    border: none;
    padding: 5px 12px;
    border-radius: 6px;
    cursor: pointer;
    font-size: 0.78rem;
    font-weight: 600;
  }
  .kpi-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(225px, 1fr));
    gap: 16px;
    margin-bottom: 24px;
  }
  .kpi-card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 18px 20px;
    position: relative;
    overflow: hidden;
    box-shadow: var(--shadow);
  }
  .kpi-card::before {
    content: "";
    position: absolute;
    top: 0; left: 0; width: 5px; height: 100%;
    background: var(--primary);
  }
  .kpi-card.warn::before { background: var(--warning); }
  .kpi-card.danger::before { background: var(--danger); }
  .kpi-card.success::before { background: var(--success); }
  .kpi-card.indigo::before { background: var(--indigo); }
  .kpi-title {
    font-size: 0.74rem;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    color: var(--text-muted);
    font-weight: 700;
  }
  .kpi-value {
    font-size: 1.5rem;
    font-weight: 800;
    margin: 6px 0 4px;
    color: var(--text);
    letter-spacing: -0.02em;
  }
  .kpi-sub {
    font-size: 0.78rem;
    color: var(--text-muted);
    font-weight: 500;
  }
  .charts-grid {
    display: grid;
    grid-template-columns: repeat(12, 1fr);
    gap: 20px;
    margin-bottom: 24px;
  }
  .chart-card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 20px;
    grid-column: span 6;
    display: flex;
    flex-direction: column;
    box-shadow: var(--shadow);
  }
  .chart-card.span-12 { grid-column: span 12; }
  .chart-card.span-5 { grid-column: span 5; }
  .chart-card.span-7 { grid-column: span 7; }
  @media (max-width: 1100px) {
    .chart-card, .chart-card.span-5, .chart-card.span-7 { grid-column: span 12; }
  }
  .chart-header {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    margin-bottom: 14px;
    gap: 8px;
  }
  .chart-title {
    font-size: 0.96rem;
    font-weight: 700;
    color: var(--text);
  }
  .chart-subtitle {
    font-size: 0.77rem;
    color: var(--text-muted);
    margin-top: 2px;
  }
  .chart-wrap {
    position: relative;
    height: 315px;
    width: 100%;
  }
  .chart-wrap.tall { height: 370px; }
  .table-card {
    background: var(--bg-card);
    border: 1px solid var(--border);
    border-radius: 14px;
    padding: 20px;
    overflow: hidden;
    box-shadow: var(--shadow);
    margin-bottom: 24px;
  }
  .table-scroll {
    overflow-x: auto;
    max-height: 540px;
    overflow-y: auto;
    margin-top: 12px;
    border: 1px solid var(--border);
    border-radius: 10px;
  }
  table {
    width: 100%;
    border-collapse: collapse;
    font-size: 0.81rem;
  }
  thead th {
    position: sticky;
    top: 0;
    background: #f1f5f9;
    color: var(--text-secondary);
    text-align: left;
    padding: 11px 13px;
    font-weight: 700;
    border-bottom: 2px solid var(--border-strong);
    white-space: nowrap;
    z-index: 2;
  }
  tbody tr {
    border-bottom: 1px solid var(--border);
    transition: background 0.12s;
    cursor: pointer;
  }
  tbody tr:nth-child(even) { background: #fafbfc; }
  tbody tr:hover { background: #e0f2fe !important; }
  td {
    padding: 10px 13px;
    vertical-align: top;
    color: var(--text-secondary);
  }
  td strong { color: var(--text); }
  td.num { text-align: right; font-variant-numeric: tabular-nums; white-space: nowrap; }
  .tag {
    display: inline-block;
    padding: 3px 9px;
    border-radius: 999px;
    font-size: 0.71rem;
    font-weight: 700;
    white-space: nowrap;
  }
  .tag-danger { background: var(--danger-bg); color: var(--danger); border: 1px solid #fecaca; }
  .tag-warn { background: var(--warning-bg); color: var(--warning); border: 1px solid #fde68a; }
  .tag-ok { background: var(--success-bg); color: var(--success); border: 1px solid #bbf7d0; }
  .tag-info { background: var(--primary-light); color: var(--primary); border: 1px solid #bae6fd; }
  .obj-cell {
    max-width: 350px;
    display: -webkit-box;
    -webkit-line-clamp: 2;
    -webkit-box-orient: vertical;
    overflow: hidden;
    font-size: 0.76rem;
    color: var(--text-muted);
  }
  /* Modal de detalhes */
  .modal-backdrop {
    display: none;
    position: fixed;
    inset: 0;
    background: rgba(15, 23, 42, 0.55);
    backdrop-filter: blur(2px);
    z-index: 100;
    align-items: center;
    justify-content: center;
    padding: 20px;
  }
  .modal-backdrop.open { display: flex; }
  .modal {
    background: #ffffff;
    border: 1px solid var(--border-strong);
    border-radius: 16px;
    max-width: 820px;
    width: 100%;
    max-height: 88vh;
    overflow-y: auto;
    padding: 26px;
    box-shadow: 0 25px 50px -12px rgba(15, 23, 42, 0.25);
  }
  .modal-header {
    display: flex;
    justify-content: space-between;
    align-items: flex-start;
    border-bottom: 1px solid var(--border);
    padding-bottom: 14px;
    margin-bottom: 18px;
  }
  .modal-close {
    background: #f1f5f9;
    border: 1px solid var(--border-strong);
    color: var(--text-secondary);
    width: 34px;
    height: 34px;
    border-radius: 8px;
    cursor: pointer;
    font-size: 1rem;
    font-weight: 700;
  }
  .modal-close:hover { background: #e2e8f0; color: var(--text); }
  .detail-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(215px, 1fr));
    gap: 12px;
    margin-bottom: 18px;
  }
  .detail-item {
    background: #f8fafc;
    border: 1px solid var(--border);
    padding: 11px 13px;
    border-radius: 10px;
  }
  .detail-item span { display: block; font-size: 0.69rem; color: var(--text-muted); text-transform: uppercase; font-weight: 700; margin-bottom: 2px; }
  .detail-item strong { font-size: 0.92rem; color: var(--text); }
  .obj-box {
    background: #f8fafc;
    border: 1px solid var(--border);
    padding: 14px;
    border-radius: 10px;
    font-size: 0.85rem;
    color: var(--text-secondary);
    margin-top: 6px;
    line-height: 1.6;
  }
  .hidden { display: none !important; }
</style>
</head>
<body>

<header>
  <div class="brand">
    <h1>
      <span>Auditoria Quantitativa — Prefeitura de Vitória/ES</span>
      <span class="badge-pill">TCC • 2025–2026</span>
      <a href="PREDIÇAO_DE_SOBRECUSTO_EM_LICITAÇOES_TCC_ATUALIZADO.pdf" target="_blank" download style="text-decoration:none;background:#4f46e5;color:#fff;font-size:0.73rem;padding:4px 10px;border-radius:6px;font-weight:600;">📘 Documento Oficial TCC (PDF)</a>
      <a href="PREDIÇAO_DE_SOBRECUSTO_EM_LICITAÇOES_TCC_ATUALIZADO.docx" download style="text-decoration:none;background:#eef2ff;color:#4338ca;border:1px solid #c7d2fe;font-size:0.73rem;padding:4px 10px;border-radius:6px;font-weight:600;">📝 TCC Editável (Word)</a>
      <a href="Analise_Contratos_PMV_2025_2026.pdf" target="_blank" download style="text-decoration:none;background:#0284c7;color:#fff;font-size:0.73rem;padding:4px 10px;border-radius:6px;font-weight:600;">📄 Relatório Técnico PDF</a>
      <a href="contratos_obras_2025_2026.csv" download style="text-decoration:none;background:#f1f5f9;color:#0f172a;border:1px solid #cbd5e1;font-size:0.73rem;padding:4px 10px;border-radius:6px;font-weight:600;">📊 CSV de Obras</a>
    </h1>
    <p>Análise de Contratos, Termos Aditivos (Lei 14.133/2021), Empresas Recorrentes, Licitações (PNCP), TCE-ES e Painel de Obras</p>
  </div>
  <div class="tabs">
    <button class="tab-btn active" onclick="setTab('obras')">1. Contratos de Obras (101)</button>
    <button class="tab-btn" onclick="setTab('empresas')">2. Ranking de Empresas &amp; Recorrência</button>
    <button class="tab-btn" onclick="setTab('painel')">3. Painel Físico/Financeiro (179)</button>
    <button class="tab-btn" onclick="setTab('geral')">4. Todos os Contratos PMV (1.065)</button>
  </div>
</header>

<div class="container">
  <!-- Barra de Filtros Contratos -->
  <div class="filter-bar" id="filterBarContratos">
    <div class="filter-group">
      <label>Ano do Contrato</label>
      <select id="fAno" onchange="render()">
        <option value="">Todos (2025 e 2026)</option>
        <option value="2025">2025</option>
        <option value="2026">2026</option>
      </select>
    </div>
    <div class="filter-group" id="fgTipo">
      <label>Tipo de Contrato</label>
      <select id="fTipo" onchange="render()">
        <option value="">Todos os Tipos</option>
      </select>
    </div>
    <div class="filter-group">
      <label>Subcategoria do Objeto</label>
      <select id="fSubcat" onchange="render()">
        <option value="">Todas as Subcategorias</option>
      </select>
    </div>
    <div class="filter-group">
      <label>Empresa / Fornecedor (com Qtd.)</label>
      <select id="fEmp" onchange="render()">
        <option value="">Todas as Empresas</option>
      </select>
    </div>
    <div class="filter-group">
      <label>Secretaria / Unidade</label>
      <select id="fSec" onchange="render()">
        <option value="">Todas as Secretarias</option>
      </select>
    </div>
    <div class="filter-group">
      <label>Situação de Aditivo</label>
      <select id="fAditivo" onchange="render()">
        <option value="">Todos os Contratos</option>
        <option value="com">Somente com Aditivo (&gt; R$ 0)</option>
        <option value="acima25">Aditivo &gt; 25%</option>
        <option value="excede">Acima do Limite Legal (25% / 50%)</option>
        <option value="supressao">Com Supressão (&lt; R$ 0)</option>
      </select>
    </div>
    <div class="filter-group" id="fgOutlier">
      <label>Outlier (300/2025)</label>
      <select id="fOutlier" onchange="render()">
        <option value="sem">Ocultar Outlier (Recomendado)</option>
        <option value="com">Incluir Outlier (R$ 2,13 bi)</option>
      </select>
    </div>
    <div class="filter-group" style="min-width:190px;">
      <label>Busca Livre (Objeto, Empresa, Nº)</label>
      <input type="text" id="fBusca" placeholder="Ex.: Ilumiterra, encosta, 223..." oninput="render()">
    </div>
    <button class="btn-reset" onclick="resetFilters()">Limpar Filtros</button>
  </div>

  <!-- Barra de Filtros Empresas -->
  <div class="filter-bar hidden" id="filterBarEmpresas">
    <div class="filter-group">
      <label>Escopo de Contratos</label>
      <select id="feEscopo" onchange="renderEmpresas()">
        <option value="obras">Somente Contratos de Obras (101 contratos • 2025–2026)</option>
        <option value="painel">Histórico Painel de Obras (179 obras • 2021–2026)</option>
        <option value="geral">Todos os Contratos da Prefeitura Sem Outlier (1.064 contratos)</option>
      </select>
    </div>
    <div class="filter-group">
      <label>Filtro de Recorrência</label>
      <select id="feMin" onchange="renderEmpresas()">
        <option value="1">Todas as Empresas (1 ou mais contratos)</option>
        <option value="2">Recorrentes (2 ou mais contratos)</option>
        <option value="3">Alta Recorrência (3 ou mais contratos)</option>
        <option value="aditivo">Somente Empresas que receberam Aditivo</option>
      </select>
    </div>
    <div class="filter-group">
      <label>Buscar Empresa ou CNPJ</label>
      <input type="text" id="feBusca" placeholder="Ex.: Ilumiterra, M.T.F., Renova..." oninput="renderEmpresas()">
    </div>
    <button class="btn-reset" onclick="resetEmpresasFilters()">Limpar Filtros</button>
  </div>

  <!-- Barra de Filtros do Painel de Obras -->
  <div class="filter-bar hidden" id="filterBarPainel">
    <div class="filter-group">
      <label>Estágio da Obra</label>
      <select id="fpEstagio" onchange="renderPainel()">
        <option value="">Todos os Estágios</option>
      </select>
    </div>
    <div class="filter-group">
      <label>Secretaria</label>
      <select id="fpSec" onchange="renderPainel()">
        <option value="">Todas as Secretarias</option>
      </select>
    </div>
    <div class="filter-group">
      <label>Busca por Obra, Bairro ou Empresa</label>
      <input type="text" id="fpBusca" placeholder="Ex.: São Pedro, CMEI, Monjardim..." oninput="renderPainel()">
    </div>
    <button class="btn-reset" onclick="resetPainelFilters()">Limpar Filtros</button>
  </div>

  <!-- KPIs -->
  <div class="kpi-grid" id="kpiGrid"></div>

  <!-- Gráficos Aba 1 e 4 (Contratos) -->
  <div class="charts-grid" id="chartsContratos">
    <div class="chart-card span-6">
      <div class="chart-header">
        <div>
          <div class="chart-title">Top Contratos por % de Aditivo vs. Limite Legal</div>
          <div class="chart-subtitle">Vermelho = ultrapassa o limite legal (25% obras / 50% reformas)</div>
        </div>
      </div>
      <div class="chart-wrap tall"><canvas id="chartTopAditivos"></canvas></div>
    </div>

    <div class="chart-card span-6">
      <div class="chart-header">
        <div>
          <div class="chart-title">Mapeamento por Subcategoria do Objeto</div>
          <div class="chart-subtitle">Volume Inicial Contratado vs. Aditivos Acumulados (R$ milhões)</div>
        </div>
      </div>
      <div class="chart-wrap tall"><canvas id="chartSubcat"></canvas></div>
    </div>

    <div class="chart-card span-6">
      <div class="chart-header">
        <div>
          <div class="chart-title">Top Empresas por Quantidade de Contratos (no filtro atual)</div>
          <div class="chart-subtitle">Quantidade total de contratos vs. quantos tiveram termo aditivo</div>
        </div>
      </div>
      <div class="chart-wrap"><canvas id="chartEmpQtd"></canvas></div>
    </div>

    <div class="chart-card span-6">
      <div class="chart-header">
        <div>
          <div class="chart-title">Modalidade de Licitação (Cruzamento PNCP)</div>
          <div class="chart-subtitle">Comparativo entre Valor Inicial e Aditivos por Modalidade</div>
        </div>
      </div>
      <div class="chart-wrap"><canvas id="chartModalidade"></canvas></div>
    </div>
  </div>

  <!-- Gráficos Aba 2 (Ranking de Empresas) -->
  <div class="charts-grid hidden" id="chartsEmpresas">
    <div class="chart-card span-6">
      <div class="chart-header">
        <div>
          <div class="chart-title">Empresas com Maior Quantidade de Contratos</div>
          <div class="chart-subtitle">Azul = Total de Contratos assinados | Laranja = Contratos que sofreram Aditivo</div>
        </div>
      </div>
      <div class="chart-wrap tall"><canvas id="chartRankingQtd"></canvas></div>
    </div>
    <div class="chart-card span-6">
      <div class="chart-header">
        <div>
          <div class="chart-title">Empresas com Maior Volume Financeiro e Aditivos (R$ milhões)</div>
          <div class="chart-subtitle">Comparativo entre Valor Inicial Somado e Total de Aditivos Recebidos</div>
        </div>
      </div>
      <div class="chart-wrap tall"><canvas id="chartRankingValor"></canvas></div>
    </div>
  </div>

  <!-- Gráficos Aba 3 (Painel de Obras) -->
  <div class="charts-grid hidden" id="chartsPainel">
    <div class="chart-card span-5">
      <div class="chart-header">
        <div>
          <div class="chart-title">Obras por Estágio de Execução (2021–2026)</div>
          <div class="chart-subtitle">Quantidade de obras em execução, concluídas ou paralisadas</div>
        </div>
      </div>
      <div class="chart-wrap"><canvas id="chartEstagio"></canvas></div>
    </div>
    <div class="chart-card span-7">
      <div class="chart-header">
        <div>
          <div class="chart-title">Execução Financeira por Estágio (Contratado vs. Empenhado vs. Pago)</div>
          <div class="chart-subtitle">Valores em R$ milhões no Painel de Obras da PMV</div>
        </div>
      </div>
      <div class="chart-wrap"><canvas id="chartExecFinanceira"></canvas></div>
    </div>
  </div>

  <!-- Tabela Interativa -->
  <div class="table-card">
    <div class="chart-header">
      <div>
        <div class="chart-title" id="tableTitle">Detalhamento dos Contratos</div>
        <div class="chart-subtitle" id="tableSubtitle">Clique em qualquer linha para abrir a ficha completa do contrato e o texto integral do Objeto</div>
      </div>
      <span class="badge-pill" id="tableCount">0 registros</span>
    </div>
    <div class="table-scroll">
      <table id="mainTable">
        <thead id="tableHead"></thead>
        <tbody id="tableBody"></tbody>
      </table>
    </div>
  </div>
</div>

<!-- Modal Detalhe -->
<div class="modal-backdrop" id="modalBackdrop" onclick="if(event.target===this) closeModal()">
  <div class="modal">
    <div class="modal-header">
      <div>
        <h2 id="mTitle" style="font-size:1.15rem;color:var(--text);">Detalhes do Contrato</h2>
        <p id="mSubtitle" style="font-size:0.82rem;color:var(--text-muted);margin-top:2px;"></p>
      </div>
      <button class="modal-close" onclick="closeModal()">✕</button>
    </div>
    <div class="detail-grid" id="mGrid"></div>
    <h4 style="font-size:0.76rem;text-transform:uppercase;color:var(--text-muted);margin-top:8px;font-weight:700;">Objeto Integral do Contrato</h4>
    <div class="obj-box" id="mObjeto"></div>
    <div id="mOtherContracts" style="margin-top:16px;"></div>
  </div>
</div>

<script>
const DATA = __PAYLOAD_JSON__;
let currentTab = 'obras';
let charts = {};

Chart.defaults.color = '#475569';
Chart.defaults.font.family = "'Inter', -apple-system, BlinkMacSystemFont, sans-serif";

const fmtBRL = (v) => {
  if (v === null || v === undefined || isNaN(v)) return '-';
  return Number(v).toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
};
const fmtNum = (v, d=2) => {
  if (v === null || v === undefined || isNaN(v)) return '-';
  return Number(v).toLocaleString('pt-BR', { minimumFractionDigits: d, maximumFractionDigits: d });
};

// Mapa rápido de quantos contratos cada empresa tem no escopo atual
function getCompanyCounts(baseRows) {
  const counts = {};
  baseRows.forEach(d => {
    const k = d.Fornecedor || 'Não informado';
    counts[k] = (counts[k] || 0) + 1;
  });
  return counts;
}

function populateFilters() {
  const base = currentTab === 'obras'
    ? DATA.contratos.filter(d => d.TipoContrato === 'Obra')
    : DATA.contratos;

  const setOptions = (id, values, defaultLabel) => {
    const sel = document.getElementById(id);
    const cur = sel.value;
    sel.innerHTML = `<option value="">${defaultLabel}</option>`;
    [...new Set(values.filter(Boolean))].sort().forEach(v => {
      const opt = document.createElement('option');
      opt.value = v;
      opt.textContent = v;
      sel.appendChild(opt);
    });
    if ([...sel.options].some(o => o.value === cur)) sel.value = cur;
  };

  setOptions('fTipo', DATA.contratos.map(d => d.TipoContrato), 'Todos os Tipos');
  setOptions('fSubcat', base.map(d => d.Subcategoria_Objeto), 'Todas as Subcategorias');
  setOptions('fSec', base.map(d => d.Secretaria), 'Todas as Secretarias');

  // Popular dropdown de empresas ordenado por quantidade de contratos
  const counts = getCompanyCounts(base);
  const empSel = document.getElementById('fEmp');
  const curEmp = empSel.value;
  empSel.innerHTML = `<option value="">Todas as Empresas (${Object.keys(counts).length})</option>`;
  Object.keys(counts)
    .sort((a, b) => counts[b] - counts[a] || a.localeCompare(b))
    .forEach(emp => {
      const opt = document.createElement('option');
      opt.value = emp;
      opt.textContent = `${emp.slice(0, 38)} (${counts[emp]} contr.)`;
      empSel.appendChild(opt);
    });
  if ([...empSel.options].some(o => o.value === curEmp)) empSel.value = curEmp;

  document.getElementById('fgTipo').style.display = currentTab === 'obras' ? 'none' : 'flex';
  document.getElementById('fgOutlier').style.display = currentTab === 'obras' ? 'none' : 'flex';

  setOptions('fpEstagio', DATA.painel_obras.map(d => d.Estagio), 'Todos os Estágios');
  setOptions('fpSec', DATA.painel_obras.map(d => d.Secretaria), 'Todas as Secretarias');
}

function setTab(tab) {
  currentTab = tab;
  const order = ['obras', 'empresas', 'painel', 'geral'];
  document.querySelectorAll('.tab-btn').forEach((b, i) => {
    b.classList.toggle('active', order[i] === tab);
  });

  document.getElementById('filterBarContratos').classList.toggle('hidden', tab !== 'obras' && tab !== 'geral');
  document.getElementById('chartsContratos').classList.toggle('hidden', tab !== 'obras' && tab !== 'geral');
  document.getElementById('filterBarEmpresas').classList.toggle('hidden', tab !== 'empresas');
  document.getElementById('chartsEmpresas').classList.toggle('hidden', tab !== 'empresas');
  document.getElementById('filterBarPainel').classList.toggle('hidden', tab !== 'painel');
  document.getElementById('chartsPainel').classList.toggle('hidden', tab !== 'painel');

  populateFilters();
  if (tab === 'painel') renderPainel();
  else if (tab === 'empresas') renderEmpresas();
  else render();
}

function resetFilters() {
  ['fAno','fTipo','fSubcat','fEmp','fSec','fAditivo','fBusca'].forEach(id => document.getElementById(id).value = '');
  document.getElementById('fOutlier').value = 'sem';
  render();
}
function resetEmpresasFilters() {
  document.getElementById('feEscopo').value = 'obras';
  document.getElementById('feMin').value = '1';
  document.getElementById('feBusca').value = '';
  renderEmpresas();
}
function resetPainelFilters() {
  ['fpEstagio','fpSec','fpBusca'].forEach(id => document.getElementById(id).value = '');
  renderPainel();
}

function filterByCompany(fornecedor) {
  const escopo = document.getElementById('feEscopo').value;
  if (escopo === 'painel') {
    setTab('painel');
    document.getElementById('fpBusca').value = fornecedor;
    renderPainel();
  } else {
    setTab(escopo === 'obras' ? 'obras' : 'geral');
    populateFilters();
    document.getElementById('fEmp').value = fornecedor;
    render();
  }
}

function getFilteredContratos() {
  const ano = document.getElementById('fAno').value;
  const tipo = document.getElementById('fTipo').value;
  const sub = document.getElementById('fSubcat').value;
  const emp = document.getElementById('fEmp').value;
  const sec = document.getElementById('fSec').value;
  const adit = document.getElementById('fAditivo').value;
  const out = document.getElementById('fOutlier').value;
  const q = document.getElementById('fBusca').value.trim().toLowerCase();

  return DATA.contratos.filter(d => {
    if (currentTab === 'obras' && d.TipoContrato !== 'Obra') return false;
    if (currentTab !== 'obras' && out === 'sem' && d.pct_aditivo > 10000) return false;
    if (ano && String(d.AnoContrato) !== ano) return false;
    if (currentTab !== 'obras' && tipo && d.TipoContrato !== tipo) return false;
    if (sub && d.Subcategoria_Objeto !== sub) return false;
    if (emp && d.Fornecedor !== emp) return false;
    if (sec && d.Secretaria !== sec) return false;
    if (adit === 'com' && !(d.Aditivo > 0)) return false;
    if (adit === 'acima25' && !(d.pct_aditivo > 25)) return false;
    if (adit === 'excede' && !d.excede) return false;
    if (adit === 'supressao' && !(d.Aditivo < 0)) return false;
    if (q) {
      const hay = `${d.NumeroContrato}/${d.AnoContrato} ${d.Fornecedor} ${d.CPFCNPJ} ${d.Objeto} ${d.Subcategoria_Objeto} ${d.Bairros_Obra_PMV}`.toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });
}

function render() {
  const rows = getFilteredContratos();
  const baseScope = currentTab === 'obras' ? DATA.contratos.filter(d => d.TipoContrato === 'Obra') : DATA.contratos;
  const companyCounts = getCompanyCounts(baseScope);

  const sumIni = rows.reduce((a, b) => a + (b.ValorContratado || 0), 0);
  const sumAd = rows.reduce((a, b) => a + (b.Aditivo || 0), 0);
  const sumTot = rows.reduce((a, b) => a + (b.ValorTotal || 0), 0);
  const varPct = sumIni > 0 ? (sumAd / sumIni * 100) : 0;
  const nComAd = rows.filter(d => d.Aditivo > 0).length;
  const nExcede = rows.filter(d => d.excede || (d.TipoContrato !== 'Obra' && d.pct_aditivo > 25)).length;
  const uniqueEmp = new Set(rows.map(d => d.Fornecedor)).size;
  const sumEmp = rows.reduce((a, b) => a + (b.ValorEmpenhado_PMV || 0), 0);
  const sumPago = rows.reduce((a, b) => a + (b.ValorPago_PMV || b.total_pago || 0), 0);

  document.getElementById('kpiGrid').innerHTML = `
    <div class="kpi-card">
      <div class="kpi-title">Contratos / Empresas</div>
      <div class="kpi-value">${rows.length} <span style="font-size:0.95rem;font-weight:600;color:var(--text-muted)">(${uniqueEmp} empr.)</span></div>
      <div class="kpi-sub">${nComAd} contratos com aditivo (${rows.length ? fmtNum(nComAd/rows.length*100,1) : 0}%)</div>
    </div>
    <div class="kpi-card indigo">
      <div class="kpi-title">Valor Inicial Contratado</div>
      <div class="kpi-value">${fmtBRL(sumIni)}</div>
      <div class="kpi-sub">Soma original assinada</div>
    </div>
    <div class="kpi-card warn">
      <div class="kpi-title">Total em Aditivos (Δ R$)</div>
      <div class="kpi-value">${fmtBRL(sumAd)}</div>
      <div class="kpi-sub">Variação global: <strong>+${fmtNum(varPct, 2)}%</strong></div>
    </div>
    <div class="kpi-card success">
      <div class="kpi-title">Valor Atualizado Total</div>
      <div class="kpi-value">${fmtBRL(sumTot)}</div>
      <div class="kpi-sub">${sumEmp > 0 ? 'Empenhado (Painel): ' + fmtBRL(sumEmp) : 'Inicial + Aditivos'}</div>
    </div>
    <div class="kpi-card danger">
      <div class="kpi-title">${currentTab === 'obras' ? 'Acima do Limite Legal (25%/50%)' : 'Aditivo > 25%'}</div>
      <div class="kpi-value">${nExcede} contratos</div>
      <div class="kpi-sub">${sumPago > 0 ? 'Pago mapeado: ' + fmtBRL(sumPago) : 'Atenção de auditoria'}</div>
    </div>
  `;

  renderChartsContratos(rows);
  renderTableContratos(rows, companyCounts);
}

function destroyChart(key) {
  if (charts[key]) { charts[key].destroy(); delete charts[key]; }
}

function renderChartsContratos(rows) {
  // 1. Top contratos por % aditivo
  destroyChart('top');
  const topRows = [...rows].filter(d => d.Aditivo > 0).sort((a, b) => b.pct_aditivo - a.pct_aditivo).slice(0, 13);
  charts.top = new Chart(document.getElementById('chartTopAditivos'), {
    type: 'bar',
    data: {
      labels: topRows.map(d => `${d.NumeroContrato}/${d.AnoContrato} - ${d.Fornecedor.slice(0, 22)}`),
      datasets: [{
        label: '% de Aditivo',
        data: topRows.map(d => d.pct_aditivo),
        backgroundColor: topRows.map(d => (d.excede || d.pct_aditivo > 25) ? '#dc2626' : '#0284c7'),
        borderRadius: 5
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: (ctx) => {
              const r = topRows[ctx.dataIndex];
              return `${fmtNum(r.pct_aditivo)}% | Aditivo: ${fmtBRL(r.Aditivo)} | Limite: ${r.limite || 25}%`;
            },
            afterLabel: (ctx) => `Subcategoria: ${topRows[ctx.dataIndex].Subcategoria_Objeto}`
          }
        }
      },
      scales: {
        x: { ticks: { callback: v => v + '%' }, grid: { color: '#e2e8f0' } },
        y: { ticks: { color: '#1e293b', font: { size: 11, weight: '500' } }, grid: { display: false } }
      }
    }
  });

  // 2. Subcategorias do Objeto
  destroyChart('subcat');
  const bySub = {};
  rows.forEach(d => {
    const k = d.Subcategoria_Objeto || 'Não classificado';
    if (!bySub[k]) bySub[k] = { ini: 0, ad: 0, n: 0 };
    bySub[k].ini += d.ValorContratado || 0;
    bySub[k].ad += d.Aditivo || 0;
    bySub[k].n += 1;
  });
  const subKeys = Object.keys(bySub).sort((a, b) => bySub[b].ini - bySub[a].ini).slice(0, 10);
  charts.subcat = new Chart(document.getElementById('chartSubcat'), {
    type: 'bar',
    data: {
      labels: subKeys.map(k => `${k.slice(0, 34)} (${bySub[k].n})`),
      datasets: [
        { label: 'Valor Inicial (R$ mi)', data: subKeys.map(k => bySub[k].ini / 1e6), backgroundColor: '#0284c7', borderRadius: 4 },
        { label: 'Aditivos (R$ mi)', data: subKeys.map(k => bySub[k].ad / 1e6), backgroundColor: '#f59e0b', borderRadius: 4 }
      ]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        tooltip: { callbacks: { label: (ctx) => `${ctx.dataset.label}: R$ ${fmtNum(ctx.raw, 2)} milhões` } }
      },
      scales: {
        x: { ticks: { callback: v => 'R$ ' + v + 'M' }, grid: { color: '#e2e8f0' } },
        y: { ticks: { color: '#1e293b', font: { size: 11, weight: '500' } }, grid: { display: false } }
      }
    }
  });

  // 3. Top Empresas por Quantidade de Contratos
  destroyChart('empQtd');
  const byEmp = {};
  rows.forEach(d => {
    const k = d.Fornecedor || 'Não informado';
    if (!byEmp[k]) byEmp[k] = { n: 0, adCount: 0, ini: 0, adVal: 0 };
    byEmp[k].n += 1;
    if (d.Aditivo > 0) byEmp[k].adCount += 1;
    byEmp[k].ini += d.ValorContratado || 0;
    byEmp[k].adVal += d.Aditivo || 0;
  });
  const empKeys = Object.keys(byEmp).sort((a, b) => byEmp[b].n - byEmp[a].n || byEmp[b].ini - byEmp[a].ini).slice(0, 10);
  charts.empQtd = new Chart(document.getElementById('chartEmpQtd'), {
    type: 'bar',
    data: {
      labels: empKeys.map(k => k.slice(0, 28)),
      datasets: [
        { label: 'Total de Contratos', data: empKeys.map(k => byEmp[k].n), backgroundColor: '#4f46e5', borderRadius: 4 },
        { label: 'Contratos com Aditivo', data: empKeys.map(k => byEmp[k].adCount), backgroundColor: '#f59e0b', borderRadius: 4 }
      ]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        tooltip: {
          callbacks: {
            afterBody: (items) => {
              const k = empKeys[items[0].dataIndex];
              return `Inicial somado: ${fmtBRL(byEmp[k].ini)}\nAditivos somados: ${fmtBRL(byEmp[k].adVal)}`;
            }
          }
        }
      },
      scales: {
        x: { ticks: { stepSize: 1 }, grid: { color: '#e2e8f0' } },
        y: { ticks: { color: '#1e293b', font: { size: 11, weight: '500' } }, grid: { display: false } }
      }
    }
  });

  // 4. Por Modalidade PNCP
  destroyChart('mod');
  const byMod = {};
  rows.forEach(d => {
    const k = d.modalidadeNome || 'Sem vínculo PNCP';
    if (!byMod[k]) byMod[k] = { ini: 0, ad: 0, n: 0 };
    byMod[k].ini += d.ValorContratado || 0;
    byMod[k].ad += d.Aditivo || 0;
    byMod[k].n += 1;
  });
  const modKeys = Object.keys(byMod).sort((a, b) => byMod[b].ini - byMod[a].ini);
  charts.mod = new Chart(document.getElementById('chartModalidade'), {
    type: 'bar',
    data: {
      labels: modKeys.map(k => `${k} (${byMod[k].n})`),
      datasets: [
        { label: 'Valor Inicial (R$ mi)', data: modKeys.map(k => byMod[k].ini / 1e6), backgroundColor: '#0284c7', borderRadius: 4 },
        { label: 'Aditivos (R$ mi)', data: modKeys.map(k => byMod[k].ad / 1e6), backgroundColor: '#16a34a', borderRadius: 4 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { ticks: { color: '#1e293b', font: { size: 10, weight: '500' } }, grid: { display: false } },
        y: { ticks: { callback: v => 'R$ ' + v + 'M' }, grid: { color: '#e2e8f0' } }
      }
    }
  });
}

function renderTableContratos(rows, companyCounts) {
  document.getElementById('tableTitle').textContent = currentTab === 'obras'
    ? 'Contratos de Obras (2025–2026) com Qtd. por Empresa, Objeto, Limite Legal e Execução'
    : 'Todos os Contratos da Prefeitura (2025–2026)';
  document.getElementById('tableSubtitle').textContent = 'Clique em qualquer linha para abrir a ficha completa do contrato e ver todos os contratos da mesma empresa';
  document.getElementById('tableCount').textContent = `${rows.length} contratos`;

  document.getElementById('tableHead').innerHTML = `
    <tr>
      <th>Contrato</th>
      <th>Subcategoria / Tipo</th>
      <th>Fornecedor (Qtd. Contratos)</th>
      <th>Objeto do Contrato</th>
      <th style="text-align:right">Valor Inicial</th>
      <th style="text-align:right">Aditivo</th>
      <th style="text-align:right">% Adit.</th>
      <th>Limite</th>
      <th style="text-align:right">Empenhado</th>
      <th style="text-align:right">Pago</th>
    </tr>
  `;

  const sorted = [...rows].sort((a, b) => (b.pct_aditivo || 0) - (a.pct_aditivo || 0));
  document.getElementById('tableBody').innerHTML = sorted.slice(0, 350).map(d => {
    let tagClass = 'tag-ok';
    let statusTxt = 'Normal';
    if (d.TipoContrato === 'Obra') {
      if (d.excede) { tagClass = 'tag-danger'; statusTxt = `Acima (${d.limite}%)`; }
      else if (d.pct_aditivo >= 20) { tagClass = 'tag-warn'; statusTxt = `Próx. (${d.limite}%)`; }
      else { statusTxt = `Até ${d.limite || 25}%`; }
    } else {
      if (d.pct_aditivo > 25) { tagClass = 'tag-warn'; statusTxt = '> 25%'; }
    }
    const qtdEmp = companyCounts[d.Fornecedor] || 1;
    const empBadge = qtdEmp > 1
      ? `<span class="tag tag-info" style="margin-left:4px">${qtdEmp} contr.</span>`
      : `<span style="font-size:0.7rem;color:var(--text-muted);margin-left:4px">(1)</span>`;
    const emp = d.ValorEmpenhado_PMV > 0 ? fmtBRL(d.ValorEmpenhado_PMV) : '-';
    const pag = d.ValorPago_PMV > 0 ? fmtBRL(d.ValorPago_PMV) : (d.total_pago > 0 ? fmtBRL(d.total_pago) : '-');
    return `
      <tr onclick='openModalContrato(${JSON.stringify(d.NumeroContrato)}, ${JSON.stringify(d.AnoContrato)}, ${JSON.stringify(d.CPFCNPJ)})'>
        <td><strong>${d.NumeroContrato}/${d.AnoContrato}</strong><br><span style="font-size:0.71rem;color:var(--text-muted)">${d.Secretaria.slice(0,25)}</span></td>
        <td><span class="tag tag-info">${d.Subcategoria_Objeto.slice(0, 26)}</span></td>
        <td><strong>${d.Fornecedor.slice(0, 28)}</strong> ${empBadge}</td>
        <td><div class="obj-cell" title="${d.Objeto.replace(/"/g, '&quot;')}">${d.Objeto}</div></td>
        <td class="num">${fmtBRL(d.ValorContratado)}</td>
        <td class="num" style="font-weight:600;color:${d.Aditivo>0?'var(--warning)':(d.Aditivo<0?'var(--danger)':'var(--text-secondary)')}">${fmtBRL(d.Aditivo)}</td>
        <td class="num"><strong>${fmtNum(d.pct_aditivo)}%</strong></td>
        <td><span class="tag ${tagClass}">${statusTxt}</span></td>
        <td class="num">${emp}</td>
        <td class="num">${pag}</td>
      </tr>
    `;
  }).join('');
}

// ABA 2: RANKING DE EMPRESAS E RECORRÊNCIA
function renderEmpresas() {
  const escopo = document.getElementById('feEscopo').value;
  const minFilter = document.getElementById('feMin').value;
  const q = document.getElementById('feBusca').value.trim().toLowerCase();

  let grouped = [];
  if (escopo === 'painel') {
    const map = {};
    DATA.painel_obras.forEach(d => {
      const m = String(d.Contratos).match(/\d{1,4}\/20\d\d\s*-\s*([^,\n]+)/);
      const emp = m ? m[1].trim() : (d.Contratos || 'Não informado');
      if (!map[emp]) map[emp] = { fornecedor: emp, cnpj: '-', qtd: 0, comAditivo: 0, ini: 0, ad: 0, tot: 0, empVal: 0, pagVal: 0 };
      map[emp].qtd += 1;
      map[emp].ini += d.ValorContratado || 0;
      map[emp].tot += d.ValorContratado || 0;
      map[emp].empVal += d.ValorEmpenhado || 0;
      map[emp].pagVal += d.ValorPago || 0;
    });
    grouped = Object.values(map);
  } else {
    const src = escopo === 'obras'
      ? DATA.contratos.filter(d => d.TipoContrato === 'Obra')
      : DATA.contratos.filter(d => d.pct_aditivo <= 10000);
    const map = {};
    src.forEach(d => {
      const emp = d.Fornecedor || 'Não informado';
      if (!map[emp]) map[emp] = { fornecedor: emp, cnpj: d.CPFCNPJ, qtd: 0, comAditivo: 0, excedeCount: 0, ini: 0, ad: 0, tot: 0, empVal: 0, pagVal: 0 };
      map[emp].qtd += 1;
      if (d.Aditivo > 0) map[emp].comAditivo += 1;
      if (d.excede || d.pct_aditivo > 25) map[emp].excedeCount += 1;
      map[emp].ini += d.ValorContratado || 0;
      map[emp].ad += d.Aditivo || 0;
      map[emp].tot += d.ValorTotal || 0;
      map[emp].empVal += d.ValorEmpenhado_PMV || 0;
      map[emp].pagVal += (d.ValorPago_PMV || d.total_pago || 0);
    });
    grouped = Object.values(map);
  }

  grouped = grouped.filter(g => {
    if (minFilter === '2' && g.qtd < 2) return false;
    if (minFilter === '3' && g.qtd < 3) return false;
    if (minFilter === 'aditivo' && g.comAditivo < 1 && g.ad <= 0) return false;
    if (q && !(`${g.fornecedor} ${g.cnpj}`.toLowerCase().includes(q))) return false;
    return true;
  });

  grouped.sort((a, b) => b.qtd - a.qtd || b.ini - a.ini);

  const totalEmp = grouped.length;
  const recorrentes = grouped.filter(g => g.qtd >= 2).length;
  const totalContratos = grouped.reduce((a, b) => a + b.qtd, 0);
  const sumIni = grouped.reduce((a, b) => a + b.ini, 0);
  const sumAd = grouped.reduce((a, b) => a + b.ad, 0);

  document.getElementById('kpiGrid').innerHTML = `
    <div class="kpi-card">
      <div class="kpi-title">Empresas Mapeadas</div>
      <div class="kpi-value">${totalEmp}</div>
      <div class="kpi-sub">Total de ${totalContratos} contratos/obras</div>
    </div>
    <div class="kpi-card indigo">
      <div class="kpi-title">Empresas Recorrentes (≥ 2 contratos)</div>
      <div class="kpi-value">${recorrentes} empresas</div>
      <div class="kpi-sub">Concentram ${grouped.filter(g=>g.qtd>=2).reduce((a,b)=>a+b.qtd,0)} contratos</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-title">Líder em Quantidade</div>
      <div class="kpi-value" style="font-size:1.15rem">${grouped[0] ? grouped[0].fornecedor.slice(0,24) : '-'}</div>
      <div class="kpi-sub">${grouped[0] ? grouped[0].qtd + ' contratos (' + fmtBRL(grouped[0].ini) + ')' : '-'}</div>
    </div>
    <div class="kpi-card warn">
      <div class="kpi-title">Volume Inicial Somado</div>
      <div class="kpi-value">${fmtBRL(sumIni)}</div>
      <div class="kpi-sub">Soma das empresas filtradas</div>
    </div>
    <div class="kpi-card danger">
      <div class="kpi-title">Aditivos nas Empresas Filtradas</div>
      <div class="kpi-value">${fmtBRL(sumAd)}</div>
      <div class="kpi-sub">${sumIni > 0 ? '+' + fmtNum(sumAd/sumIni*100, 2) + '% sobre o inicial' : '-'}</div>
    </div>
  `;

  destroyChart('rankQtd');
  destroyChart('rankVal');

  const topQtd = grouped.slice(0, 12);
  charts.rankQtd = new Chart(document.getElementById('chartRankingQtd'), {
    type: 'bar',
    data: {
      labels: topQtd.map(g => g.fornecedor.slice(0, 28)),
      datasets: [
        { label: 'Qtd. Total de Contratos', data: topQtd.map(g => g.qtd), backgroundColor: '#0284c7', borderRadius: 4 },
        { label: 'Qtd. com Aditivo', data: topQtd.map(g => g.comAditivo), backgroundColor: '#f59e0b', borderRadius: 4 }
      ]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { ticks: { stepSize: 1 }, grid: { color: '#e2e8f0' } },
        y: { ticks: { color: '#1e293b', font: { size: 11, weight: '600' } }, grid: { display: false } }
      }
    }
  });

  const topVal = [...grouped].sort((a, b) => (b.ad || b.ini) - (a.ad || a.ini)).slice(0, 12);
  charts.rankVal = new Chart(document.getElementById('chartRankingValor'), {
    type: 'bar',
    data: {
      labels: topVal.map(g => `${g.fornecedor.slice(0, 25)} (${g.qtd})`),
      datasets: [
        { label: 'Valor Inicial (R$ mi)', data: topVal.map(g => g.ini / 1e6), backgroundColor: '#4f46e5', borderRadius: 4 },
        { label: 'Aditivos Recebidos (R$ mi)', data: topVal.map(g => g.ad / 1e6), backgroundColor: '#dc2626', borderRadius: 4 }
      ]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { ticks: { callback: v => 'R$ ' + v + 'M' }, grid: { color: '#e2e8f0' } },
        y: { ticks: { color: '#1e293b', font: { size: 11, weight: '600' } }, grid: { display: false } }
      }
    }
  });

  document.getElementById('tableTitle').textContent = 'Ranking de Empresas por Quantidade de Contratos e Volume de Aditivos';
  document.getElementById('tableSubtitle').textContent = 'Clique em qualquer empresa para filtrar imediatamente todos os contratos dela na aba de Contratos';
  document.getElementById('tableCount').textContent = `${grouped.length} empresas`;

  document.getElementById('tableHead').innerHTML = `
    <tr>
      <th>#</th>
      <th>Empresa / Fornecedor</th>
      <th>CNPJ</th>
      <th style="text-align:right">Qtd. Contratos</th>
      <th style="text-align:right">Com Aditivo</th>
      <th style="text-align:right">Valor Inicial Somado</th>
      <th style="text-align:right">Total de Aditivos</th>
      <th style="text-align:right">% Aditivo Global</th>
      <th style="text-align:right">Valor Atualizado</th>
      <th style="text-align:right">Pago Mapeado</th>
    </tr>
  `;

  document.getElementById('tableBody').innerHTML = grouped.map((g, i) => {
    const pct = g.ini > 0 ? (g.ad / g.ini * 100) : 0;
    const tagClass = pct > 25 ? 'tag-danger' : (pct > 0 ? 'tag-warn' : 'tag-ok');
    return `
      <tr onclick='filterByCompany(${JSON.stringify(g.fornecedor)})'>
        <td><strong>${i + 1}º</strong></td>
        <td><strong>${g.fornecedor}</strong></td>
        <td>${g.cnpj}</td>
        <td class="num"><span class="tag tag-info">${g.qtd} contrato${g.qtd>1?'s':''}</span></td>
        <td class="num">${g.comAditivo} (${fmtNum(g.comAditivo/g.qtd*100, 0)}%)</td>
        <td class="num">${fmtBRL(g.ini)}</td>
        <td class="num" style="font-weight:700;color:${g.ad>0?'var(--warning)':'var(--text-secondary)'}">${fmtBRL(g.ad)}</td>
        <td class="num"><span class="tag ${tagClass}">${fmtNum(pct)}%</span></td>
        <td class="num"><strong>${fmtBRL(g.tot)}</strong></td>
        <td class="num">${g.pagVal > 0 ? fmtBRL(g.pagVal) : '-'}</td>
      </tr>
    `;
  }).join('');
}

// ABA 3: PAINEL DE OBRAS
function renderPainel() {
  const est = document.getElementById('fpEstagio').value;
  const sec = document.getElementById('fpSec').value;
  const q = document.getElementById('fpBusca').value.trim().toLowerCase();

  const rows = DATA.painel_obras.filter(d => {
    if (est && d.Estagio !== est) return false;
    if (sec && d.Secretaria !== sec) return false;
    if (q) {
      const hay = `${d.Nome} ${d.Contratos} ${d.Bairros} ${d.Secretaria}`.toLowerCase();
      if (!hay.includes(q)) return false;
    }
    return true;
  });

  const sumCon = rows.reduce((a, b) => a + (b.ValorContratado || 0), 0);
  const sumEmp = rows.reduce((a, b) => a + (b.ValorEmpenhado || 0), 0);
  const sumLiq = rows.reduce((a, b) => a + (b.ValorLiquidado || 0), 0);
  const sumPag = rows.reduce((a, b) => a + (b.ValorPago || 0), 0);
  const nParal = rows.filter(d => d.Estagio === 'Paralisado' || d.Estagio === 'Rescisão Unilateral').length;

  document.getElementById('kpiGrid').innerHTML = `
    <div class="kpi-card">
      <div class="kpi-title">Obras no Painel (2021–2026)</div>
      <div class="kpi-value">${rows.length}</div>
      <div class="kpi-sub">${nParal} paralisadas ou rescindidas</div>
    </div>
    <div class="kpi-card indigo">
      <div class="kpi-title">Valor Atual Contratado</div>
      <div class="kpi-value">${fmtBRL(sumCon)}</div>
      <div class="kpi-sub">Soma com aditivos no Painel</div>
    </div>
    <div class="kpi-card warn">
      <div class="kpi-title">Valor Total Empenhado</div>
      <div class="kpi-value">${fmtBRL(sumEmp)}</div>
      <div class="kpi-sub">${sumCon > 0 ? fmtNum(sumEmp/sumCon*100, 1) : 0}% do contratado</div>
    </div>
    <div class="kpi-card">
      <div class="kpi-title">Valor Total Liquidado</div>
      <div class="kpi-value">${fmtBRL(sumLiq)}</div>
      <div class="kpi-sub">Medido e atestado</div>
    </div>
    <div class="kpi-card success">
      <div class="kpi-title">Valor Efetivamente Pago</div>
      <div class="kpi-value">${fmtBRL(sumPag)}</div>
      <div class="kpi-sub">${sumCon > 0 ? fmtNum(sumPag/sumCon*100, 1) : 0}% pago sobre o contratado</div>
    </div>
  `;

  destroyChart('est');
  destroyChart('exec');
  const byEst = {};
  rows.forEach(d => {
    const k = d.Estagio || 'Não informado';
    if (!byEst[k]) byEst[k] = { n: 0, con: 0, emp: 0, pag: 0 };
    byEst[k].n += 1;
    byEst[k].con += d.ValorContratado || 0;
    byEst[k].emp += d.ValorEmpenhado || 0;
    byEst[k].pag += d.ValorPago || 0;
  });
  const estKeys = Object.keys(byEst).sort((a, b) => byEst[b].n - byEst[a].n);

  charts.est = new Chart(document.getElementById('chartEstagio'), {
    type: 'doughnut',
    data: {
      labels: estKeys.map(k => `${k} (${byEst[k].n})`),
      datasets: [{
        data: estKeys.map(k => byEst[k].n),
        backgroundColor: ['#0284c7', '#16a34a', '#4f46e5', '#f59e0b', '#dc2626', '#9333ea', '#db2777']
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { position: 'right', labels: { color: '#1e293b', font: { size: 11, weight: '500' } } } }
    }
  });

  charts.exec = new Chart(document.getElementById('chartExecFinanceira'), {
    type: 'bar',
    data: {
      labels: estKeys,
      datasets: [
        { label: 'Contratado (R$ mi)', data: estKeys.map(k => byEst[k].con / 1e6), backgroundColor: '#0284c7', borderRadius: 4 },
        { label: 'Empenhado (R$ mi)', data: estKeys.map(k => byEst[k].emp / 1e6), backgroundColor: '#f59e0b', borderRadius: 4 },
        { label: 'Pago (R$ mi)', data: estKeys.map(k => byEst[k].pag / 1e6), backgroundColor: '#16a34a', borderRadius: 4 }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      scales: {
        x: { ticks: { color: '#1e293b', font: { size: 10, weight: '500' } }, grid: { display: false } },
        y: { ticks: { callback: v => 'R$ ' + v + 'M' }, grid: { color: '#e2e8f0' } }
      }
    }
  });

  document.getElementById('tableTitle').textContent = 'Obras do Painel Oficial da Prefeitura (2021–2026)';
  document.getElementById('tableSubtitle').textContent = 'Relação completa das 179 obras com estágio físico, bairros e execução orçamentária';
  document.getElementById('tableCount').textContent = `${rows.length} obras`;
  document.getElementById('tableHead').innerHTML = `
    <tr>
      <th>ID / Contrato</th>
      <th>Nome da Obra</th>
      <th>Bairros</th>
      <th>Estágio</th>
      <th style="text-align:right">Valor Contratado</th>
      <th style="text-align:right">Valor Empenhado</th>
      <th style="text-align:right">Valor Liquidado</th>
      <th style="text-align:right">Valor Pago</th>
    </tr>
  `;
  const sorted = [...rows].sort((a, b) => (b.ValorContratado || 0) - (a.ValorContratado || 0));
  document.getElementById('tableBody').innerHTML = sorted.map(d => {
    const isBad = d.Estagio === 'Paralisado' || d.Estagio === 'Rescisão Unilateral';
    return `
      <tr>
        <td><strong>#${d.Id}</strong><br><span style="font-size:0.72rem;color:var(--text-muted)">${d.Contratos.slice(0, 36)}</span></td>
        <td><strong>${d.Nome}</strong><br><span style="font-size:0.72rem;color:var(--text-muted)">${d.Secretaria}</span></td>
        <td style="font-size:0.76rem">${d.Bairros}</td>
        <td><span class="tag ${isBad ? 'tag-danger' : 'tag-info'}">${d.Estagio}</span></td>
        <td class="num">${fmtBRL(d.ValorContratado)}</td>
        <td class="num">${fmtBRL(d.ValorEmpenhado)}</td>
        <td class="num">${fmtBRL(d.ValorLiquidado)}</td>
        <td class="num"><strong>${fmtBRL(d.ValorPago)}</strong></td>
      </tr>
    `;
  }).join('');
}

function openModalContrato(num, ano, cnpj) {
  const d = DATA.contratos.find(r => String(r.NumeroContrato)===String(num) && String(r.AnoContrato)===String(ano) && String(r.CPFCNPJ)===String(cnpj));
  if (!d) return;
  const outrosDaEmpresa = DATA.contratos.filter(r => r.Fornecedor === d.Fornecedor);

  document.getElementById('mTitle').textContent = `Contrato ${d.NumeroContrato}/${d.AnoContrato} — ${d.Fornecedor}`;
  document.getElementById('mSubtitle').textContent = `${d.Secretaria} • CNPJ: ${d.CPFCNPJ} • Processo: ${d.Processo}/${d.AnoProcesso} • Esta empresa possui ${outrosDaEmpresa.length} contrato(s) na base`;
  document.getElementById('mGrid').innerHTML = `
    <div class="detail-item"><span>Tipo / Subcategoria</span><strong>${d.TipoContrato} — ${d.Subcategoria_Objeto}</strong></div>
    <div class="detail-item"><span>Valor Inicial Contratado</span><strong>${fmtBRL(d.ValorContratado)}</strong></div>
    <div class="detail-item"><span>Aditivo Acumulado</span><strong>${fmtBRL(d.Aditivo)} (${fmtNum(d.pct_aditivo)}%)</strong></div>
    <div class="detail-item"><span>Valor Atualizado Total</span><strong>${fmtBRL(d.ValorTotal)}</strong></div>
    <div class="detail-item"><span>Limite Legal Aplicável</span><strong>${d.limite ? d.limite + '%' : 'N/A'} ${d.excede ? '(ACIMA DO LIMITE)' : ''}</strong></div>
    <div class="detail-item"><span>Modalidade PNCP</span><strong>${d.modalidadeNome || 'Não vinculado'}</strong></div>
    <div class="detail-item"><span>Empenhado / Pago (Painel PMV)</span><strong>${d.ValorEmpenhado_PMV > 0 ? fmtBRL(d.ValorEmpenhado_PMV) + ' / ' + fmtBRL(d.ValorPago_PMV) : 'Não consta no Painel de Obras'}</strong></div>
    <div class="detail-item"><span>Estágio / Bairros (Painel PMV)</span><strong>${d.Estagio_Obra_PMV || '-'} ${d.Bairros_Obra_PMV ? '• ' + d.Bairros_Obra_PMV : ''}</strong></div>
    <div class="detail-item"><span>Termos Registrados no TCE-ES</span><strong>${d.n_aditivos} aditivos | ${d.n_prorrogacao} prorrog. | ${d.n_apostilamentos} apost.</strong></div>
    <div class="detail-item"><span>Vigência</span><strong>${d.DataInicio || '-'} a ${d.DataFinal || '-'}</strong></div>
  `;
  document.getElementById('mObjeto').textContent = d.Objeto;

  if (outrosDaEmpresa.length > 1) {
    document.getElementById('mOtherContracts').innerHTML = `
      <h4 style="font-size:0.76rem;text-transform:uppercase;color:var(--text-muted);margin-bottom:8px;font-weight:700;">
        Todos os ${outrosDaEmpresa.length} contratos de ${d.Fornecedor} na base (2025–2026)
      </h4>
      <div style="border:1px solid var(--border);border-radius:8px;overflow:hidden;">
        <table style="font-size:0.76rem;">
          <thead>
            <tr>
              <th>Contrato</th>
              <th>Tipo</th>
              <th>Subcategoria</th>
              <th style="text-align:right">Valor Inicial</th>
              <th style="text-align:right">Aditivo</th>
              <th style="text-align:right">% Adit.</th>
            </tr>
          </thead>
          <tbody>
            ${outrosDaEmpresa.map(o => `
              <tr style="${o.NumeroContrato===d.NumeroContrato && o.AnoContrato===d.AnoContrato ? 'background:#e0f2fe;font-weight:600;' : ''}">
                <td>${o.NumeroContrato}/${o.AnoContrato}</td>
                <td>${o.TipoContrato}</td>
                <td>${o.Subcategoria_Objeto}</td>
                <td class="num">${fmtBRL(o.ValorContratado)}</td>
                <td class="num">${fmtBRL(o.Aditivo)}</td>
                <td class="num">${fmtNum(o.pct_aditivo)}%</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    `;
  } else {
    document.getElementById('mOtherContracts').innerHTML = '';
  }

  document.getElementById('modalBackdrop').classList.add('open');
}

function closeModal() {
  document.getElementById('modalBackdrop').classList.remove('open');
}

populateFilters();
render();
</script>
</body>
</html>
"""

html_final = html_template.replace("__PAYLOAD_JSON__", json.dumps(payload, ensure_ascii=False))
with open("dashboard_obras_pmv.html", "w", encoding="utf-8") as f:
    f.write(html_final)

print("Dashboard claro atualizado com sucesso! Tamanho:", len(html_final), "bytes")
