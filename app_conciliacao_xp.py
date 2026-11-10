"""Conciliação mensal de patrimônio XP | M Wealth.

Execute: streamlit run app_conciliacao_xp.py
Dependências: streamlit, pandas, openpyxl, xlsxwriter

Método: ganho = PL final - PL inicial - fluxos externos líquidos.
Modified Dietz: ganho / (PL inicial + soma(fluxo externo * peso temporal)).
O resultado depende de extrato integral, classificação correta, e PLs de bases compatíveis.
"""
from __future__ import annotations

from datetime import date, datetime, timedelta
from io import BytesIO
import calendar
import re
import unicodedata

import pandas as pd
import streamlit as st

st.set_page_config(page_title="Conciliação XP | M Wealth", layout="wide", page_icon="📊")

st.markdown("""<style>
.block-container {max-width: 1480px; padding-top: 1.7rem}
:root {--mw:#DCC9B1}
h1,h2,h3 {letter-spacing:-.025em}
.mw-kpi {background:#192532; border:1px solid #344355; border-radius:14px;
         padding:16px 17px; min-height:112px; box-sizing:border-box; margin-bottom:8px}
.mw-kpi .kpi-label {color:#cbd4df;font-size:.83rem;line-height:1.4;margin-bottom:12px}
.mw-kpi .kpi-value {color:#fff;font-size:clamp(1.18rem,1.9vw,1.75rem);
                   font-weight:750;line-height:1.2;overflow-wrap:anywhere;font-variant-numeric:tabular-nums}
.mw-kpi .kpi-note {color:#b2bfca;font-size:.76rem;margin-top:9px}
.mw-kpi.accent {border-color:#857760;background:#252923}
.mw-kpi.accent .kpi-value {color:#f5e8d5}
</style>""", unsafe_allow_html=True)


def kpi(label, value, note="", accent=False):
    from html import escape
    cl = " accent" if accent else ""
    note_html = f'<div class="kpi-note">{escape(note)}</div>' if note else ""
    st.markdown(f'<div class="mw-kpi{cl}"><div class="kpi-label">{escape(label)}</div>'
                f'<div class="kpi-value">{escape(value)}</div>{note_html}</div>',
                unsafe_allow_html=True)



def normalize(value: object) -> str:
    txt = unicodedata.normalize("NFKD", str(value or ""))
    return "".join(c for c in txt if not unicodedata.combining(c)).upper().strip()


def money(value: object) -> float:
    if pd.isna(value):
        return float("nan")
    if isinstance(value, (int, float)):
        return float(value)
    txt = str(value).replace("R$", "").replace(" ", "").replace("\u00a0", "")
    if "," in txt:
        txt = txt.replace(".", "").replace(",", ".")
    return float(txt)


def parse_date(value: object):
    if pd.isna(value) or value == "":
        return pd.NaT
    if isinstance(value, (pd.Timestamp, date, datetime)):
        return pd.Timestamp(value).normalize()
    if isinstance(value, (int, float)) and 30000 < value < 70000:
        return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(value))
    txt = str(value).strip()
    if re.fullmatch(r"\d{5}(?:\.0)?", txt):
        return pd.Timestamp("1899-12-30") + pd.Timedelta(days=float(txt))
    return pd.to_datetime(txt, dayfirst=True, errors="coerce")


def account(value: object) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip()
    return re.sub(r"\.0$", "", s)


def read_excel(upload, **kwargs):
    # A leitura parte sempre do início, inclusive após alterações de parâmetros no Streamlit.
    return pd.read_excel(BytesIO(upload.getvalue()), **kwargs)


def pl_file(upload):
    frame = read_excel(upload)
    names = {normalize(k): k for k in frame.columns}
    try:
        c_acc = names["CONTA"]
        c_dt = names["DATA"] if "DATA" in names else names["DATA-BASE"]
        c_pl = names["PL"]
    except KeyError as exc:
        raise ValueError("Controle de PL deve conter as colunas Conta, Data e PL.") from exc
    f = frame[[c_acc, c_dt, c_pl]].copy()
    f.columns = ["Conta", "Data", "PL"]
    f["Conta"] = f["Conta"].map(account)
    f["Data"] = f["Data"].map(parse_date)
    f["PL"] = f["PL"].map(money)
    f = f.dropna(subset=["Data", "PL"])
    f = f[f["Conta"] != ""]
    if f.duplicated(["Conta", "Data"]).any():
        dup = f.loc[f.duplicated(["Conta", "Data"], keep=False), ["Conta", "Data"]]
        raise ValueError("Controle de PL tem registros repetidos para conta/data: " + dup.to_string(index=False))
    return f


def position_file(upload, base_date):
    data = read_excel(upload, header=None)
    raw_account = str(data.iloc[0, 5]) if data.shape[1] > 5 else ""
    m = re.search(r"Conta\s*:\s*(\d+)", raw_account, flags=re.I)
    if not m:
        raise ValueError(f"Não encontrei o número da conta na posição {upload.name} (célula F1).")
    value = money(data.iloc[3, 0])
    return {"Conta": m.group(1), "Data": pd.Timestamp(base_date), "PL": value}


def movement_files(uploads):
    out = []
    required = ["CLIENTE", "DATA DE REFERENCIA", "DATA DE LIQUIDACAO", "DESCRICAO", "VALOR"]
    for upload in uploads:
        f = read_excel(upload)
        cols = {normalize(col): col for col in f.columns}
        missing = [name for name in required if name not in cols]
        if missing:
            raise ValueError(f"Extrato {upload.name} sem colunas: {', '.join(missing)}")
        part = pd.DataFrame({
            "Conta": f[cols["CLIENTE"]].map(account),
            "Data referência": f[cols["DATA DE REFERENCIA"]].map(parse_date),
            "Data liquidação": f[cols["DATA DE LIQUIDACAO"]].map(parse_date),
            "Descrição": f[cols["DESCRICAO"]].fillna("").astype(str),
            "Valor": f[cols["VALOR"]].map(money),
            "Arquivo": upload.name,
        })
        if "SALDO" in cols:
            part["Saldo caixa"] = f[cols["SALDO"]].map(money)
        else:
            part["Saldo caixa"] = float("nan")
        out.append(part)
    if not out:
        return pd.DataFrame()
    frame = pd.concat(out, ignore_index=True).dropna(subset=["Data referência", "Valor"])
    # Não deduplicar por valor/descrição: operações legítimas podem ser idênticas.
    return frame.reset_index(drop=True)


CATEGORIES = ["Interna / investimento", "Aporte externo", "Retirada externa", "Revisar"]


def classify(desc: str) -> str:
    d = normalize(desc)
    # Uma TED de aplicação de fundos é movimentação interna, não saque bancário.
    if any(k in d for k in ["TED APLICACAO FUND", "APLICACAO FUND", "RESGATE FUND",
                           "RECOMPRA COMPROMISS", "COMPRA COMPROMISS", "IRRF", "IOF",
                           "TAXA GEST", "DEVOLUCAO TX", "COTIZACAO", "DIVIDENDO",
                           "JUROS", "RENDIMENTO", "AMORTIZACAO", "VENCIMENTO",
                           "BOLETA", "LIQUIDACAO", "CUPOM", "PROVENTO"]):
        return "Interna / investimento"
    if "RETIRADA EM C/C" in d or "RETIRADA EM CC" in d:
        return "Retirada externa"
    if "TED" in d and any(k in d for k in ["RECEB", "CREDITO EM C/C", "DEPOSITO"]):
        return "Aporte externo"
    # Desconhecidos devem ser revisados, não presumidos como entradas ou saídas.
    if any(k in d for k in ["TED", "PIX", "TRANSFER", "DOC", "DEPOSITO", "RETIRADA", "SAQUE", "CREDITO EXTERNO"]):
        return "Revisar"
    return "Interna / investimento"


def fmt_brl(n):
    if pd.isna(n): return "—"
    return "R$ " + f"{n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def result_book(summary, details, pl, checks, notes, comparison=None):
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
        sheets = [("Resultado por conta", summary), ("Memória movimentações", details),
                  ("Controle PL", pl), ("Conferências", checks), ("Metodologia", notes)]
        if comparison is not None:
            sheets.insert(1, ("Comparativo referência", comparison))
        for name, df in sheets:
            df.to_excel(writer, sheet_name=name, index=False)
            ws = writer.sheets[name]
            wb = writer.book
            header = wb.add_format({"bold": True, "bg_color": "#173D57", "font_color": "white", "border": 0})
            for col_no, col_name in enumerate(df.columns):
                ws.write(0, col_no, col_name, header)
                width = min(65, max(15, int(max(len(str(col_name)), *(len(str(x)) for x in df[col_name].head(60).fillna("")))) + 3)) if len(df) else 22
                ws.set_column(col_no, col_no, width)
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(1, len(df)), max(0, len(df.columns)-1))
    return bio.getvalue()


st.title("Conciliação de PL e Ganho Financeiro | XP")
st.caption("M Wealth • Apuração mensal independente do relatório de performance da XP")
st.info("**Rotina mensal:** atualize cinco PLs no Controle de Fechamento e importe o extrato completo. A posição consolidada detalhada só é necessária na implantação, se o PL inicial ainda não estiver no controle.")

col_a, col_b = st.columns([1, 1])
with col_a:
    month_ref = st.date_input("Mês de referência", value=date(2026, 10, 1), format="DD/MM/YYYY")
    year, month = month_ref.year, month_ref.month
    end_dt = pd.Timestamp(year, month, calendar.monthrange(year, month)[1])
    start_dt = pd.Timestamp(year, month, 1) - pd.Timedelta(days=1)
    st.caption(f"Abertura: {start_dt:%d/%m/%Y} · Fechamento: {end_dt:%d/%m/%Y} (último dia calendário; PL do último dia útil aceito)")
with col_b:
    flow_date_type = st.radio("Data para os fluxos externos", ["Data de referência", "Data de liquidação"], horizontal=True, help="Use a data que produz o lançamento efetivo no PL. Confira especialmente TEDs na virada do mês.")

upload_pl = st.file_uploader("1. Controle de PL Fechamento (.xlsx)", type="xlsx", key="pl")
upload_mov = st.file_uploader("2. Extrato XP (.xlsx) — um ou mais arquivos", type="xlsx", accept_multiple_files=True, key="extrato")
with st.expander("Implantação: posição inicial (opcional, somente se faltar no controle)"):
    st.write("Envie as posições consolidadas com data-base de abertura. A data no cabeçalho da XP pode ser apenas a data de emissão; utilizaremos a data-base informada aqui.")
    upload_pos = st.file_uploader("Posições consolidadas das contas", type="xlsx", accept_multiple_files=True, key="posicoes")
    initial_date = st.date_input("Data-base das posições anexadas", value=start_dt.date(), format="DD/MM/YYYY")

if not upload_pl or not upload_mov:
    st.stop()

try:
    pl = pl_file(upload_pl)
    if upload_pos:
        seeds = pd.DataFrame([position_file(u, initial_date) for u in upload_pos])
        # O arquivo Controle de PL tem prioridade quando a mesma conta/data existe.
        pl = pd.concat([seeds, pl], ignore_index=True).drop_duplicates(["Conta", "Data"], keep="last")
    movements = movement_files(upload_mov)
except Exception as exc:
    st.error(str(exc))
    st.stop()

# O fechamento deve ser o último PL registrado no mês, mas não pode ser de um dia
# anterior ao último dia útil sem sinalização.
# Se o último dia calendário do mês anterior não foi útil, usar o fechamento do
# último dia útil (p.ex., 30/10 para a abertura de novembro/2026).
last_open_bday = (start_dt if start_dt.weekday() < 5 else start_dt - pd.offsets.BDay(1)).normalize()
opening = pl.loc[(pl["Data"] >= last_open_bday) & (pl["Data"] <= start_dt)].copy()
opening = opening.sort_values("Data").drop_duplicates("Conta", keep="last")
opening = opening[["Conta", "PL"]].rename(columns={"PL": "PL inicial"})
# Uma data-base de último dia útil anterior a uma data de fechamento não útil é aceita.
last_bday = (end_dt if end_dt.weekday() < 5 else end_dt - pd.offsets.BDay(1)).normalize()
close_rows = pl.loc[(pl["Data"] >= last_bday) & (pl["Data"] <= end_dt)].copy()
close_rows = close_rows.sort_values("Data").drop_duplicates("Conta", keep="last")
closing = close_rows[["Conta", "PL", "Data"]].rename(columns={"PL": "PL final", "Data": "Data PL final"})
accounts = sorted(set(opening["Conta"]) | set(closing["Conta"]) | set(movements["Conta"]))
base = pd.DataFrame({"Conta": accounts}).merge(opening, on="Conta", how="left").merge(closing, on="Conta", how="left")
base["Status"] = base.apply(lambda r: "OK" if pd.notna(r["PL inicial"]) and pd.notna(r["PL final"]) else "PL inicial/final ausente", axis=1)

m = movements.copy()
# A seleção da interface usa "de", enquanto as colunas do extrato usam
# nomes internos padronizados. Não acessar o DataFrame pela legenda da tela.
DATE_COLUMNS = {
    "Data de referência": "Data referência",
    "Data de liquidação": "Data liquidação",
}
chosen_date_column = DATE_COLUMNS.get(flow_date_type)
if chosen_date_column not in m.columns:
    st.error(f"O extrato não possui a coluna de data esperada: {chosen_date_column}. Colunas lidas: {', '.join(map(str, m.columns))}")
    st.stop()
m["Data apuração"] = m[chosen_date_column]
m = m[(m["Data apuração"] >= start_dt + pd.Timedelta(days=1)) & (m["Data apuração"] <= end_dt)].copy()
m = m.sort_values(["Conta", "Data apuração", "Descrição", "Valor"]).reset_index(drop=True)
m.insert(0, "ID", range(1, len(m) + 1))
m["Classificação"] = m["Descrição"].map(classify)

st.subheader("1. Conferência das movimentações")
st.caption("Revise apenas as movimentações com 'Revisar' e qualquer TED/transferência que não represente um aporte ou retirada efetiva do patrimônio. Alterações valem para esta apuração e aparecem na memória exportada.")
counts = m["Classificação"].value_counts()
if counts.get("Revisar", 0):
    st.warning(f"Existem {counts['Revisar']} lançamentos sem classificação conclusiva. Resolva antes de considerar o resultado final.")
with st.expander(f"Editar classificação de {len(m)} movimentações", expanded=bool(counts.get("Revisar", 0))):
    m_edit = st.data_editor(m[["ID", "Conta", "Data apuração", "Descrição", "Valor", "Classificação"]],
                            hide_index=True, use_container_width=True, height=420,
                            disabled=["ID", "Conta", "Data apuração", "Descrição", "Valor"],
                            column_config={"Classificação": st.column_config.SelectboxColumn("Classificação", options=CATEGORIES, required=True)},
                            key=f"ed_{year}_{month}_{flow_date_type}")
    m["Classificação"] = m_edit["Classificação"].values

if not len(m):
    st.error("O extrato não contém movimentações no mês selecionado (conforme a data escolhida). Confira as datas e a integridade do arquivo.")
    st.stop()

if m["Classificação"].eq("Revisar").any():
    st.error("Há movimentos em 'Revisar'. A apuração será exibida como preliminar.")

m["Fluxo externo"] = 0.0
is_in = m["Classificação"] == "Aporte externo"
is_out = m["Classificação"] == "Retirada externa"
m.loc[is_in | is_out, "Fluxo externo"] = m.loc[is_in | is_out, "Valor"]
wrong_sign = m.loc[(is_in & (m["Valor"] < 0)) | (is_out & (m["Valor"] > 0))]
if len(wrong_sign):
    st.error(f"{len(wrong_sign)} movimentação(ões) com sinal divergente da classificação. Revise a memória.")

m["Peso Dietz"] = (end_dt - m["Data apuração"]).dt.days / (end_dt - start_dt).days
m["Fluxo ponderado"] = m["Fluxo externo"] * m["Peso Dietz"]
flow = m.groupby("Conta", as_index=False).agg(
    **{"Fluxo líquido": ("Fluxo externo", "sum"), "Fluxo ponderado": ("Fluxo ponderado", "sum")})
summary = base.merge(flow, on="Conta", how="left")
summary[["Fluxo líquido", "Fluxo ponderado"]] = summary[["Fluxo líquido", "Fluxo ponderado"]].fillna(0)
summary["Ganho financeiro"] = summary["PL final"] - summary["PL inicial"] - summary["Fluxo líquido"]
summary["Capital ponderado"] = summary["PL inicial"] + summary["Fluxo ponderado"]
summary["Modified Dietz (%)"] = (summary["Ganho financeiro"] / summary["Capital ponderado"]).where(summary["Capital ponderado"] > 0) * 100
summary.loc[summary["Capital ponderado"] <= 0, "Status"] = "Base Dietz inválida"

# Lançamentos que influenciam o ganho patrimonial, mas NÃO são fluxos externos.
# Valores sinalizados no extrato: IRRF e taxas são negativos, cashback é positivo.
desc_upper = m["Descrição"].map(normalize)
m["Natureza econômica"] = "Demais movimentos / operações internas"
m.loc[desc_upper.str.contains("IRRF|IMPOSTO DE RENDA", regex=True), "Natureza econômica"] = "IRRF"
m.loc[desc_upper.str.contains("IOF", regex=True), "Natureza econômica"] = "IOF"
m.loc[desc_upper.str.contains("TAXA GEST|TAXA DE GEST", regex=True), "Natureza econômica"] = "Taxa de gestão"
m.loc[desc_upper.str.contains("DEVOLUCAO TX|CASHBACK|DEVOLUCAO DE TAXA", regex=True), "Natureza econômica"] = "Cashback / devolução de distribuição"
m.loc[m["Classificação"].isin(["Aporte externo", "Retirada externa"]), "Natureza econômica"] = "Fluxo externo"

recon = m.pivot_table(index="Conta", columns="Natureza econômica", values="Valor", aggfunc="sum", fill_value=0)
for cat in ["IRRF", "IOF", "Taxa de gestão", "Cashback / devolução de distribuição"]:
    if cat not in recon.columns:
        recon[cat] = 0.0
recon = recon[["IRRF", "IOF", "Taxa de gestão", "Cashback / devolução de distribuição"]].reset_index()
recon.columns = ["Conta", "IRRF lançado", "IOF lançado", "Taxa gestão lançada", "Cashback lançado"]
summary = summary.merge(recon, on="Conta", how="left")
for cat in recon.columns[1:]:
    summary[cat] = summary[cat].fillna(0.0)
# O PL final já inclui os efeitos dos impostos/taxas/cashbacks contabilizados.
# Para reconstruir o ganho ANTES de IR e taxas, adicionamos de volta os desembolsos.
# Cashback NÃO é somado outra vez: já está incluído na variação patrimonial.
summary["IR + IOF (custo)"] = -(summary["IRRF lançado"] + summary["IOF lançado"])
summary["Taxa gestão (custo)"] = -summary["Taxa gestão lançada"]
summary["Ganho bruto (antes IR e gestão)"] = (summary["Ganho financeiro"]
                                             + summary["IR + IOF (custo)"]
                                             + summary["Taxa gestão (custo)"])
summary["Retorno bruto Dietz (%)"] = (summary["Ganho bruto (antes IR e gestão)"] / summary["Capital ponderado"]).where(summary["Capital ponderado"] > 0) * 100

st.subheader("2. Resultado mensal")
st.caption("**Ganho líquido** = variação de PL ajustada pelos fluxos externos. **Ganho bruto ajustado** = ganho líquido + IRRF/IOF + taxa de gestão identificados no extrato. O cashback já está no PL, portanto não é somado novamente.")
complete = summary["PL inicial"].notna() & summary["PL final"].notna()
complete_data = summary.loc[complete]
if (~complete).any():
    st.warning("Faltam posições de abertura/fechamento para: " + ", ".join(summary.loc[~complete, "Conta"]) + ". O consolidado abaixo inclui apenas as contas completas.")
if len(complete_data):
    g = complete_data["Ganho financeiro"].sum()
    g_bruto = complete_data["Ganho bruto (antes IR e gestão)"].sum()
    start_sum = complete_data["PL inicial"].sum()
    end_sum = complete_data["PL final"].sum()
    flows = complete_data["Fluxo líquido"].sum()
    weighted = complete_data["Fluxo ponderado"].sum()
    rate = 100 * g/(start_sum + weighted) if (start_sum + weighted) > 0 else float("nan")
    rate_bruto = 100 * g_bruto/(start_sum + weighted) if (start_sum + weighted) > 0 else float("nan")
    cols = st.columns(4)
    with cols[0]: kpi("PL inicial",fmt_brl(start_sum))
    with cols[1]: kpi("PL final",fmt_brl(end_sum))
    with cols[2]: kpi("Fluxo externo líquido",fmt_brl(flows))
    with cols[3]: kpi("Ganho financeiro líquido",fmt_brl(g),accent=True)
    cols2=st.columns(4)
    with cols2[0]: kpi("Ganho bruto ajustado", fmt_brl(g_bruto), "Antes de IR e taxa de gestão")
    with cols2[1]: kpi("IRRF + IOF identificados",fmt_brl(complete_data["IR + IOF (custo)"].sum()))
    with cols2[2]: kpi("Taxas de gestão identificadas",fmt_brl(complete_data["Taxa gestão (custo)"].sum()))
    with cols2[3]: kpi("Cashback contabilizado",fmt_brl(complete_data["Cashback lançado"].sum()),"Já incluído no ganho líquido")
    st.caption(f"Rentabilidade Modified Dietz: **líquida {rate:.4f}%** · **bruta ajustada {rate_bruto:.4f}%** (aproximações; o cálculo bruto pressupõe custos lançados no período).")

view_cols = ["Conta", "PL inicial", "PL final", "Fluxo líquido", "Ganho financeiro", "IR + IOF (custo)", "Taxa gestão (custo)", "Cashback lançado", "Ganho bruto (antes IR e gestão)", "Modified Dietz (%)", "Retorno bruto Dietz (%)", "Status"]
view = summary[view_cols].copy()
fmts={c:fmt_brl for c in ["PL inicial", "PL final", "Fluxo líquido", "Ganho financeiro", "IR + IOF (custo)", "Taxa gestão (custo)", "Cashback lançado", "Ganho bruto (antes IR e gestão)"]}
fmts.update({c:(lambda v: f"{v:.4f}%" if pd.notna(v) else "—") for c in ["Modified Dietz (%)", "Retorno bruto Dietz (%)"]})
st.dataframe(view.style.format(fmts), use_container_width=True, hide_index=True)

st.subheader("3. Comparação com apuração XP / controle manual")
st.caption("Referência opcional para conferência. A comparação não altera o resultado calculado. Digite o ganho anotado para cada conta e escolha o critério de comparação.")
bench_defaults = {"9445242":339.27,"5166121":110215.15,"4163084":46839.88,"11370136":32549.55,"7983962":103587.85}
bench = pd.DataFrame({"Conta":summary["Conta"],"Ganho de referência (R$)":[bench_defaults.get(acc, float("nan")) if year==2026 and month==10 else float("nan") for acc in summary["Conta"]]})
bench_edit = st.data_editor(bench, hide_index=True, use_container_width=True,
    disabled=["Conta"],column_config={"Ganho de referência (R$)":st.column_config.NumberColumn("Ganho de referência (R$)",format="%.2f")},
    key=f"bench_{year}_{month}")
criterion = st.radio("Comparar referência com",["Ganho bruto (antes IR e gestão)","Ganho financeiro"],horizontal=True,
    format_func=lambda v:"Bruto (antes de IR e gestão)" if v.startswith("Ganho bruto") else "Líquido (após IR e gestão)")
compare = summary[["Conta","Ganho financeiro","Ganho bruto (antes IR e gestão)","IR + IOF (custo)","Taxa gestão (custo)","Cashback lançado"]].merge(bench_edit,on="Conta",how="left")
compare["Diferença (app - referência)"]=compare[criterion]-compare["Ganho de referência (R$)"]
st.dataframe(compare.style.format({c:fmt_brl for c in compare.columns if c!="Conta"}),hide_index=True,use_container_width=True)
if compare["Diferença (app - referência)"].abs().gt(1).any():
    discrepant=compare.loc[compare["Diferença (app - referência)"].abs()>1,["Conta","Diferença (app - referência)"]]
    st.warning("Existem diferenças superiores a R$ 1,00: " + "; ".join(f"conta {r['Conta']}: {fmt_brl(r['Diferença (app - referência)'])}" for _,r in discrepant.iterrows()) + ". Confira a natureza dos lançamentos, eventuais IRs não refletidos no extrato e os critérios de apuração XP.")

st.subheader("4. Validações e exportação")
checks = pd.DataFrame([
    {"Verificação": "Contas sem PL completo", "Resultado": ", ".join(summary.loc[~complete,"Conta"]) or "Nenhuma"},
    {"Verificação": "Movimentações pendentes de classificação", "Resultado": str(int(m["Classificação"].eq("Revisar").sum()))},
    {"Verificação": "Movimentações com sinal divergente", "Resultado": str(len(wrong_sign))},
    {"Verificação": "Extrato: 1ª data de referência", "Resultado": str(movements["Data referência"].min().date())},
    {"Verificação": "Extrato: última data de referência", "Resultado": str(movements["Data referência"].max().date())},
    {"Verificação": "Extrato: última data de liquidação", "Resultado": str(movements["Data liquidação"].max().date())},
    {"Verificação": "Data usada para fluxos", "Resultado": flow_date_type},
    {"Verificação": "Ganho consolidado = PL final - PL inicial - fluxos", "Resultado": fmt_brl(complete_data["Ganho financeiro"].sum())},
    {"Verificação": "Ganho bruto = líquido + impostos + taxa gestão", "Resultado": fmt_brl(complete_data["Ganho bruto (antes IR e gestão)"].sum())},
    {"Verificação": "Cashback (já incluso no ganho)", "Resultado": fmt_brl(complete_data["Cashback lançado"].sum())},
])

last_mov_day = m["Data apuração"].max()
if last_mov_day < last_bday:
    st.warning("A última movimentação disponível antecede o fechamento. Isso não prova que há extrato incompleto, mas é importante confirmar a cobertura até o fim do mês.")
if last_bday != end_dt:
    st.caption(f"Fechamento financeiro aceito em {last_bday:%d/%m/%Y}; os pesos Modified Dietz usam o encerramento calendário de {end_dt:%d/%m/%Y}. Se necessário, ajuste o critério no código para combinar com o padrão de sua instituição.")

st.dataframe(checks, use_container_width=True, hide_index=True)
notes = pd.DataFrame({"Metodologia / alertas": [
    "Ganho = PL final - PL inicial - soma dos fluxos externos assinados (positivo=aporte; negativo=retirada).",
    "Modified Dietz = ganho / [PL inicial + soma(fluxo externo * (dias até fim do mês / dias do período))].",
    "Modified Dietz é uma aproximação de retorno, não TWR exato. O valor do ganho não depende do Modified Dietz.",
    "Ganho líquido reflete IRRF, IOF, taxas de gestão e cashback quando integrados aos patrimônios de abertura e fechamento.",
    "Ganho bruto ajustado = ganho líquido + IRRF + IOF + taxa de gestão debitados no mês e identificados no extrato.",
    "Cashback/devolução de distribuição é receita da conta e já está no PL final; não somar novamente ao ganho.",
    "Se quiser rentabilidade antes também do efeito do cashback, subtraia a devolução do ganho bruto, apenas como análise adicional.",
    "Transferências entre contas do próprio grupo devem ser classificadas como fluxo externo nas contas individuais se saírem/entrarem no PL, mesmo que no consolidado se anulem.",
    "Aplicações, resgates e recompras de investimentos dentro da mesma conta NÃO são fluxos externos.",
    "Os PLs inicial/final devem estar no mesmo critério (bruto/líquido) e representar os respectivos fechamentos.",
    "A completude do extrato, a classificação dos fluxos e a data efetiva do crédito/débito devem ser validadas.",
    "O relatório é gerencial e não substitui o informe fiscal, contábil ou o relatório oficial da XP.",
]})

export_view = view.copy()
if len(complete_data):
    consolidated_row = {
        "Conta": "CONSOLIDADO (contas completas)",
        "PL inicial": start_sum,
        "PL final": end_sum,
        "Fluxo líquido": flows,
        "Ganho financeiro": g,
        "IR + IOF (custo)": complete_data["IR + IOF (custo)"].sum(),
        "Taxa gestão (custo)": complete_data["Taxa gestão (custo)"].sum(),
        "Cashback lançado": complete_data["Cashback lançado"].sum(),
        "Ganho bruto (antes IR e gestão)": g_bruto,
        "Modified Dietz (%)": rate,
        "Retorno bruto Dietz (%)": rate_bruto,
        "Status": "OK" if complete.all() and not m["Classificação"].eq("Revisar").any() and not len(wrong_sign) else "PRELIMINAR",
    }
    export_view = pd.concat([export_view, pd.DataFrame([consolidated_row])], ignore_index=True)
out = result_book(export_view, m, pl.sort_values(["Data", "Conta"]), checks, notes, compare)
st.download_button("⬇️ Baixar apuração consolidada e memória de cálculo (.xlsx)",
                   data=out, file_name=f"conciliacao_xp_{year}_{month:02d}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary", use_container_width=True)

with st.expander("Próximos fechamentos: rotina mínima"):
    st.markdown("""1. No **Controle de PL**, acrescente uma linha por conta com data-base e PL do fechamento (cinco linhas por mês). Não precisa de PL diário.\n2. Exporte o **extrato completo do mês**, incluindo todos os fluxos externos.\n3. Selecione o mês, importe os arquivos, revise as transferências e baixe a memória.\n4. Guarde o mesmo Controle de PL histórico: o fechamento anterior vira a abertura do mês seguinte.""")
