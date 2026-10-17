import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Image, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

m = pd.read_csv("analise_contratos_metricas.csv", sep=";", decimal=",", dtype={"NumeroContrato": str, "AnoContrato": str})
OBRA = m.TipoContrato == "Obra"
out = m.pct_aditivo > 10000
base = m[~out]
ob = m[OBRA]


def br(v, d=2):
    if pd.isna(v):
        return "-"
    return f"{v:,.{d}f}".replace(",", "X").replace(".", ",").replace("X", ".")


def stats(s):
    s = s.dropna()
    return [len(s), s.mean(), s.median(), s.std(), s.min(), s.max()]


# ---------- gráficos ----------
plt.rcParams.update({"font.size": 8})
o = ob[ob.pct_aditivo >= 20].sort_values("pct_aditivo")
fig, ax = plt.subplots(figsize=(7.2, 3.6))
cores = ["#c0392b" if e else "#2e86c1" for e in o.excede]
ax.barh([f"{n}/{a} {f[:22]}" for n, a, f in zip(o.NumeroContrato, o.AnoContrato, o.Fornecedor)], o.pct_aditivo, color=cores)
ax.axvline(25, color="k", ls="--", lw=0.8)
ax.axvline(50, color="gray", ls=":", lw=0.8)
ax.set_xlabel("% de aditivo sobre o valor contratado  (linhas: 25% e 50%)")
ax.set_title("Obras com aditivo >= 20% (vermelho = acima do limite aplicado)")
plt.tight_layout()
plt.savefig("g1.png", dpi=170)
plt.close()

mo = m[OBRA & m.achou_pncp].groupby("modalidadeNome").agg(n=("ValorContratado", "size"), ini=("ValorContratado", "sum"), ad=("Aditivo", "sum"))
mo["pct"] = mo.ad / mo.ini * 100
fig, ax = plt.subplots(figsize=(7.2, 2.8))
ax.bar(mo.index, mo.pct, color="#2e86c1")
for i, (p, n) in enumerate(zip(mo.pct, mo.n)):
    ax.text(i, p, f"{br(p)}%\n(n={n})", ha="center", va="bottom", fontsize=7)
ax.set_ylabel("Aditivo / valor inicial (%)")
ax.set_title("Obras: aditivo sobre o valor inicial, por modalidade (PNCP)")
ax.set_ylim(0, mo.pct.max() * 1.3)
plt.tight_layout()
plt.savefig("g2.png", dpi=170)
plt.close()

ss = base.groupby("Secretaria").agg(ini=("ValorContratado", "sum"), ad=("Aditivo", "sum")).sort_values("ad", ascending=False).head(8).iloc[::-1]
fig, ax = plt.subplots(figsize=(7.2, 3.2))
ax.barh([s[:38] for s in ss.index], ss.ad / 1e6, color="#2e86c1")
ax.set_xlabel("Aditivos acumulados (R$ milhões, sem o outlier)")
ax.set_title("Secretarias com maior valor de aditivos (2025-2026)")
plt.tight_layout()
plt.savefig("g3.png", dpi=170)
plt.close()

# ---------- PDF ----------
ss_ = getSampleStyleSheet()
H1 = ParagraphStyle("H1", parent=ss_["Heading1"], fontSize=15, spaceAfter=6, textColor=colors.HexColor("#1b4f72"))
H2 = ParagraphStyle("H2", parent=ss_["Heading2"], fontSize=11.5, spaceBefore=10, spaceAfter=4, textColor=colors.HexColor("#1b4f72"))
P = ParagraphStyle("P", parent=ss_["BodyText"], fontSize=8.8, leading=11.5)
S = ParagraphStyle("S", parent=P, fontSize=7.4, leading=9.0)
SB = ParagraphStyle("SB", parent=S, fontName="Helvetica-Bold", textColor=colors.white)


def tbl(data, widths=None, header=True):
    rows = []
    for r_i, r in enumerate(data):
        rows.append([Paragraph(str(c), SB if (header and r_i == 0) else S) for c in r])
    t = Table(rows, colWidths=widths, repeatRows=1 if header else 0)
    st = [("GRID", (0, 0), (-1, -1), 0.3, colors.HexColor("#b0b7bf")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
          ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]
    if header:
        st.append(("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1b4f72")))
        for i in range(1, len(rows)):
            if i % 2 == 0:
                st.append(("BACKGROUND", (0, i), (-1, i), colors.HexColor("#eef3f7")))
    t.setStyle(TableStyle(st))
    return t


def rodape(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 7)
    canvas.setFillColor(colors.grey)
    canvas.drawString(1.8 * cm, 1.0 * cm, "Analise quantitativa de contratos e obras - Prefeitura Municipal de Vitoria/ES - 2025 e 2026")
    canvas.drawRightString(A4[0] - 1.8 * cm, 1.0 * cm, f"pagina {doc.page}")
    canvas.restoreState()


E = []
E.append(Paragraph("Análise quantitativa de contratos e obras - Prefeitura Municipal de Vitória/ES", H1))
E.append(Paragraph("Contratos de 2025 e 2026: valores iniciais, aditivos, execução financeira (empenhado, liquidado e pago), limites legais, licitação (PNCP) e termos (TCE-ES). "
                   "Documento com valores numéricos e cálculos, sem juízo de valor.", P))

E.append(Paragraph("1. Dados e método", H2))
E.append(tbl([
    ["Fonte", "Conteúdo usado"],
    ["Portal da Transparência de Vitória (Contratos 2025/2026 + IPAMV e CDTIV)", f"{len(m)} contratos consolidados, sendo {OBRA.sum()} com TipoContrato = Obra. Valor contratado, aditivo acumulado, valor total, fornecedor, secretaria, datas e objeto."],
    ["Painel de Obras da PMV (Obras-2025/2026.csv)", f"179 obras cadastradas (2021-2026), sendo {int(ob['ValorEmpenhado_PMV'].notna().sum())} casadas com contratos de 2025-2026. Traz Valor Empenhado, Valor Liquidado, Valor Pago, Estágio e Bairros."],
    ["PNCP (API de consulta)", f"1.287 licitações de Vitória (2025 a 16/10/2026). Cruzamento pelo número do processo: {int(m.achou_pncp.sum())} contratos com licitação ({int((m.achou_pncp & OBRA).sum())} obras)."],
    ["TCE-ES (dados abertos, CidadES Contratações)", f"Instrumentos, termos de alteração e medições/pagamentos. Cruzamento por número/ano e CNPJ: {int(m.casou_tce.sum())} contratos ({int((m.casou_tce & OBRA).sum())} obras)."],
], [5.2 * cm, 11.6 * cm]))
E.append(Spacer(1, 4))
E.append(Paragraph("<b>Cálculos:</b> Variação (R$) = valor total - valor contratado; Variação % = variação (R$) / valor contratado x 100. "
                   "Limite legal aplicado às obras: 25% do valor inicial, ou 50% quando o objeto é reforma (Lei 14.133, art. 125). "
                   "<b>Outlier:</b> contrato 300/2025 (aditivo de R$ 2.139.480.000,00 sobre R$ 2.139.480,00 contratados, 100.000%), apresentado à parte nas tabelas &quot;sem outlier&quot;.", P))

E.append(Paragraph("2. Totais e variação (R$)", H2))
grp = [("Todos", m), ("Todos sem outlier", base), ("Obras", ob), ("Serviço", m[m.TipoContrato.str.startswith("Presta", na=False)]),
       ("Material", m[m.TipoContrato.str.startswith("Fornec", na=False)]), ("Outros", m[m.TipoContrato == "Outros"])]
d = [["Grupo", "n", "Valor inicial", "Aditivos", "Valor atual", "Var. %", "Com aditivo", "Supressões"]]
for k, g in grp:
    ini, ad = g.ValorContratado.sum(), g.Aditivo.sum()
    d.append([k, len(g), br(ini), br(ad), br(g.ValorTotal.sum()), br(ad / ini * 100), int((g.Aditivo > 0).sum()), int((g.Aditivo < 0).sum())])
E.append(tbl(d, [3 * cm, 1.1 * cm, 2.9 * cm, 2.9 * cm, 2.9 * cm, 1.3 * cm, 1.4 * cm, 1.4 * cm]))

E.append(Paragraph("3. Estatística descritiva", H2))
d = [["Métrica", "Contratado (sem outlier)", "Aditivo (sem outlier)", "% aditivo (sem outlier)", "Contratado (obras)", "Aditivo (obras)", "% aditivo (obras)"]]
a, b = base, ob
nomes = ["n", "Média", "Mediana", "Desvio padrão", "Mínimo", "Máximo"]
S1 = [stats(a.ValorContratado), stats(a.Aditivo), stats(a.pct_aditivo), stats(b.ValorContratado), stats(b.Aditivo), stats(b.pct_aditivo)]
for i, n in enumerate(nomes):
    d.append([n] + [br(s[i], 0 if i == 0 else 2) for s in S1])
E.append(tbl(d, [2.2 * cm, 2.7 * cm, 2.4 * cm, 2.3 * cm, 2.7 * cm, 2.3 * cm, 2.2 * cm]))
pa, pb = base[base.Aditivo > 0].pct_aditivo, ob[ob.Aditivo > 0].pct_aditivo
E.append(Spacer(1, 3))
E.append(Paragraph(f"Considerando só contratos com aditivo positivo: todos sem outlier, n={len(pa)}, média {br(pa.mean())}%, mediana {br(pa.median())}%; "
                   f"obras, n={len(pb)}, média {br(pb.mean())}%, mediana {br(pb.median())}%.", P))

E.append(PageBreak())
E.append(Paragraph("4. Obras, limite legal e execução financeira (Empenhado / Pago)", H2))
ac = int(ob.excede.sum())
E.append(Paragraph(f"Das {len(ob)} obras: {int((ob.Aditivo > 0).sum())} têm aditivo positivo, {int((ob.Aditivo < 0).sum())} tem supressão, "
                   f"<b>{ac} estão acima do limite aplicado</b> e {int(((ob.pct_aditivo >= 20) & (~ob.excede)).sum())} ficam entre 20% e o limite.", P))
E.append(Image("g1.png", width=16.0 * cm, height=7.6 * cm))
o2 = ob[ob.pct_aditivo >= 20].sort_values("pct_aditivo", ascending=False)
d = [["Contrato", "Fornecedor / Subcategoria", "Inicial", "Aditivo", "%", "Lim.", "Acima?", "Empenhado", "Pago"]]
for _, r in o2.iterrows():
    E_emp = br(r.ValorEmpenhado_PMV) if pd.notna(r.ValorEmpenhado_PMV) else "-"
    E_pag = br(r.ValorPago_PMV) if pd.notna(r.ValorPago_PMV) else (br(r.total_pago) if pd.notna(r.total_pago) else "-")
    d.append([
        f"{r.NumeroContrato}/{r.AnoContrato}",
        f"{r.Fornecedor[:28]} ({str(r.Subcategoria_Objeto)[:24]})",
        br(r.ValorContratado), br(r.Aditivo), br(r.pct_aditivo),
        f"{br(r.limite, 0)}%", "Sim" if r.excede else "Não",
        E_emp, E_pag
    ])
E.append(tbl(d, [1.5 * cm, 4.7 * cm, 2.0 * cm, 1.9 * cm, 1.1 * cm, 1.0 * cm, 1.1 * cm, 1.9 * cm, 1.8 * cm]))

E.append(Spacer(1, 4))
E.append(Paragraph("<b>Mapeamento por Subcategoria do Objeto (101 contratos classificados como Obra):</b>", P))
sub_df = ob.groupby("Subcategoria_Objeto").agg(n=("ValorContratado", "size"), ini=("ValorContratado", "sum"), ad=("Aditivo", "sum"), ca=("Aditivo", lambda s: (s > 0).sum()), emp=("ValorEmpenhado_PMV", "sum"), pag=("ValorPago_PMV", "sum")).sort_values("n", ascending=False)
d_sub = [["Subcategoria do Objeto", "n", "Valor Inicial", "Aditivos", "Var. %", "Com adit.", "Empenhado (Painel)", "Pago (Painel)"]]
for k, r in sub_df.iterrows():
    d_sub.append([k, int(r.n), br(r.ini), br(r.ad), br(r.ad / r.ini * 100 if r.ini > 0 else 0), int(r.ca), br(r.emp) if r.emp > 0 else "-", br(r.pag) if r.pag > 0 else "-"])
E.append(tbl(d_sub, [5.4 * cm, 0.9 * cm, 2.3 * cm, 2.0 * cm, 1.2 * cm, 1.2 * cm, 2.0 * cm, 2.0 * cm]))

E.append(PageBreak())
E.append(Paragraph("5. Discrepâncias matemáticas e cruzamento com o Painel de Obras", H2))
t = m[m.casou_tce]
ob_cas = ob[ob.ValorEmpenhado_PMV.notna()]
E.append(tbl([
    ["Verificação", "Resultado"],
    ["Valor contratado + aditivo diferente de valor total (tolerância R$ 0,01)", f"{int((m.div_soma.abs() > 0.01).sum())} contratos"],
    ["Valor contratado menor ou igual a zero", f"{int((m.ValorContratado <= 0).sum())} contratos"],
    ["Aditivo maior que o valor contratado (> 100%)", f"{int((m.pct_aditivo > 100).sum())} contratos"],
    ["Obras (2025-2026) localizadas no Painel de Obras da PMV", f"{len(ob_cas)} de {len(ob)} obras (Empenhado: R$ {br(ob_cas.ValorEmpenhado_PMV.sum())} | Pago: R$ {br(ob_cas.ValorPago_PMV.sum())})"],
    ["Obras em que o Valor Empenhado supera o Valor Total (Inicial + Aditivo)", f"{int((ob_cas.ValorEmpenhado_PMV > ob_cas.ValorTotal + 0.01).sum())} obras"],
    ["Valor inicial PMV igual ao TCE-ES (contratos casados)", f"{int(t.valor_confere.sum())} de {len(t)}"],
    ["Aditivo > 0 na PMV e nenhum termo no TCE-ES (contratos casados)", f"{int(((t.Aditivo > 0) & (t.n_termos == 0)).sum())} contratos"],
], [10.8 * cm, 6 * cm]))

E.append(Paragraph("6. Licitação (PNCP) e contrato", H2))
c = m[m.valorTotalHomologado.notna() & (m.valorTotalHomologado > 0)]
cb = c[c.TipoContrato == "Obra"]
d = [["Grupo", "n", "Média %", "Mediana %", "Desvio %", "Mín %", "Máx %", "Contr. > homolog.", "Atual > homolog."]]
for k, g in [("Todos", c), ("Obras", cb)]:
    s = stats(g.desc_est_hom_pct)
    d.append([k, int(s[0])] + [br(x) for x in s[1:]] + [int((g.contr_vs_hom > 0.01).sum()), int((g.ValorTotal > g.valorTotalHomologado + 0.01).sum())])
E.append(Paragraph("Desconto da licitação = (valor estimado - valor homologado) / valor estimado.", P))
E.append(tbl(d, [1.8 * cm, 1.1 * cm, 1.8 * cm, 1.9 * cm, 1.8 * cm, 1.7 * cm, 1.7 * cm, 2.5 * cm, 2.5 * cm]))
E.append(Spacer(1, 4))
E.append(Image("g2.png", width=15 * cm, height=5.5 * cm))
mm = base[base.achou_pncp].groupby("modalidadeNome").agg(n=("ValorContratado", "size"), ini=("ValorContratado", "sum"), ad=("Aditivo", "sum"), ca=("Aditivo", lambda s: (s > 0).sum()))
d = [["Modalidade (todos, sem outlier)", "n", "Valor inicial", "Aditivo", "Aditivo/inicial %", "Com aditivo"]]
for k, r in mm.iterrows():
    d.append([k, int(r.n), br(r.ini), br(r.ad), br(r.ad / r.ini * 100), int(r.ca)])
E.append(tbl(d, [4.6 * cm, 1.3 * cm, 3 * cm, 2.7 * cm, 2.8 * cm, 2.4 * cm]))

E.append(PageBreak())
E.append(Paragraph("7. Termos e pagamentos (TCE-ES, apenas contratos casados)", H2))
E.append(tbl([
    ["Indicador", "Valor"],
    ["Contratos com termo de alteração", int((t.n_termos > 0).sum())],
    ["Termos aditivos / apostilamentos", f"{int(t.n_aditivos.sum())} / {int(t.n_apostilamentos.sum())}"],
    ["Prorrogações de prazo", int(t.n_prorrogacao.sum())],
    ["Acréscimos ou supressões quantitativas", int(t.n_acresc_supress.sum())],
    ["Reequilíbrios econômico-financeiros", int(t.n_reequilibrio.sum())],
    ["Contratos com pagamento registrado no TCE-ES / total pago (R$)", f"{int((t.n_pagamentos > 0).sum())} / {br(t.total_pago.sum())}"],
], [11.8 * cm, 5 * cm]))

E.append(Paragraph("8. Aditivos por secretaria / entidade", H2))
E.append(Image("g3.png", width=15 * cm, height=6.4 * cm))
ss = base.groupby("Secretaria").agg(n=("ValorContratado", "size"), ini=("ValorContratado", "sum"), ad=("Aditivo", "sum"), ca=("Aditivo", lambda s: (s > 0).sum())).sort_values("ad", ascending=False).head(12)
d = [["Secretaria / Entidade (sem outlier)", "n", "Valor inicial", "Aditivo", "Aditivo/inicial %", "Com aditivo"]]
for k, r in ss.iterrows():
    d.append([k[:55], int(r.n), br(r.ini), br(r.ad), br(r.ad / r.ini * 100), int(r.ca)])
E.append(tbl(d, [6 * cm, 1.1 * cm, 3 * cm, 2.6 * cm, 2.4 * cm, 1.7 * cm]))

E.append(Paragraph("9. Limitações", H2))
for tx in [
    "O Painel de Obras da PMV traz Valor Empenhado, Liquidado e Pago para 40 das 101 obras de 2025-2026 (e mais 139 obras de 2021-2024 salvas em painel_obras_pmv_2021_2026.csv).",
    "Obras = campo TipoContrato 'Obra' da própria base. Há itens de tecnologia, manutenção e eventos nessa classificação (mapeados na Subcategoria do Objeto).",
    "Reforma identificada pelo campo do TCE-ES ou pela palavra 'reforma' no objeto; pode haver erros.",
    "O limite de 25% foi aplicado apenas às obras. Serviços contínuos têm regras próprias.",
]:
    E.append(Paragraph("• " + tx, P))

E.append(PageBreak())
E.append(Paragraph("10. Recorrência de empresas nos contratos de obras (2025–2026)", H2))
E.append(Paragraph("Nos 101 contratos classificados como Obra, atuam 57 empresas distintas, sendo 21 empresas recorrentes (2 ou mais contratos), que concentram 65 contratos:", P))
emp_ob = ob.groupby(["CPFCNPJ", "Fornecedor"], as_index=False).agg(
    qtd=("ValorContratado", "size"),
    c_adit=("Aditivo", lambda s: int((s > 0).sum())),
    ac_lim=("excede", lambda s: int((s == True).sum())),
    ini=("ValorContratado", "sum"),
    ad=("Aditivo", "sum"),
    tot=("ValorTotal", "sum"),
).sort_values(["qtd", "tot"], ascending=[False, False]).head(15)
d_emp = [["Pos.", "Fornecedor", "CNPJ", "Contratos", "C/ adit.", "> Lim.", "Valor inicial", "Aditivo", "Valor atual", "% Δ"]]
for idx, r in enumerate(emp_ob.itertuples(), 1):
    pct_e = (r.ad / r.ini * 100) if r.ini > 0 else 0
    d_emp.append([f"{idx}º", str(r.Fornecedor)[:36], str(r.CPFCNPJ), int(r.qtd), f"{r.c_adit} ({int(round(r.c_adit/r.qtd*100))}%)", int(r.ac_lim), br(r.ini), br(r.ad), br(r.tot), br(pct_e) + "%"])
E.append(tbl(d_emp, [0.8 * cm, 4.5 * cm, 2.5 * cm, 1.3 * cm, 1.4 * cm, 1.0 * cm, 2.2 * cm, 1.9 * cm, 2.2 * cm, 1.2 * cm]))

E.append(Spacer(1, 6))
E.append(Paragraph("11. Telas do Painel Interativo Web (Dashboard Publicado)", H2))
E.append(Paragraph("Link público do painel interativo: https://leobmt02.github.io/tcc-auditoria-obras-vitoria/", P))
for img_p, cap_tx in [
    ("img_painel/aba1_obras.png", "Figura A — Aba 1: Contratos de Obras (101) — Subcategorias, limites legais (25%/50%) e filtros."),
    ("img_painel/aba2_empresas.png", "Figura B — Aba 2: Ranking de Empresas e Recorrência — Quantidade de contratos e proporção de aditivos."),
    ("img_painel/aba3_painel_fisico.png", "Figura C — Aba 3: Painel Físico/Financeiro (179 obras, 2021–2026) — Contratado, Empenhado, Liquidado e Pago."),
    ("img_painel/aba4_todos_contratos.png", "Figura D — Aba 4: Todos os Contratos PMV (1.065) — Visão com e sem outlier 300/2025."),
]:
    E.append(Spacer(1, 4))
    E.append(Image(img_p, width=16.5 * cm, height=8.2 * cm))
    E.append(Paragraph(cap_tx, P))

doc = SimpleDocTemplate("Analise_Contratos_PMV_2025_2026.pdf", pagesize=A4, leftMargin=1.8 * cm, rightMargin=1.8 * cm, topMargin=1.6 * cm, bottomMargin=1.6 * cm,
                        title="Análise quantitativa de contratos e obras - PMV 2025-2026")
doc.build(E, onFirstPage=rodape, onLaterPages=rodape)
print("PDF ATUALIZADO OK")

