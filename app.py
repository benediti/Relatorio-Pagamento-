import io
import re
import pandas as pd
import streamlit as st

# (Opcional) PDF simples
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas


st.set_page_config(page_title="Relatório de Pagamento", layout="wide")
st.title("Relatório de Pagamento (Funcionários + Totais Líquidos)")

st.markdown(
    """
    **Entrada:**
    - `funcionarios.csv` com: matricula, nome, cargo, localizacao, Data de Admissão, Data de Desligamento
    - `Totais liquidos.xls` com: matricula, Valor Liquido, DataPagto, Referencia

    **Saída:**
    - Base final com as colunas do relatório + total de Valor Líquido
    """
)

# -----------------------
# Helpers
# -----------------------
def normalize_col(s: str) -> str:
    # normaliza para facilitar match de colunas (sem acentos, minúsculo, só letras/números)
    s = str(s).strip().lower()
    s = re.sub(r"\s+", " ", s)
    # remove acentos de forma simples via encode/decode
    s = s.encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^a-z0-9 ]", "", s)
    s = s.replace(" ", "_")
    return s

def pick_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    norm_map = {normalize_col(c): c for c in df.columns}
    for cand in candidates:
        key = normalize_col(cand)
        if key in norm_map:
            return norm_map[key]
    return None

def to_str_matricula(x):
    if pd.isna(x):
        return ""
    # evita 1234.0
    s = str(x).strip()
    s = s.replace(".0", "") if re.fullmatch(r"\d+\.0", s) else s
    # remove espaços
    s = re.sub(r"\s+", "", s)
    return s

def parse_money_br(x):
    """Aceita 1.234,56 / 1234,56 / 1234.56 / numérico -> float"""
    if pd.isna(x):
        return 0.0
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if s == "":
        return 0.0
    # remove separador de milhar e padroniza decimal
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except:
        return 0.0

def parse_date(x):
    # pandas lida com datas variadas
    if pd.isna(x) or str(x).strip() == "":
        return pd.NaT
    return pd.to_datetime(x, errors="coerce", dayfirst=True)

def format_mes_ano_pt(x):
    meses = [
        "Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
        "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"
    ]
    dt = parse_date(x)
    if pd.isna(dt):
        s = str(x).strip()
        return s
    return f"{meses[dt.month - 1]}/{dt.year}"

def df_to_excel_bytes(df: pd.DataFrame, sheet_name="Relatorio"):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()

def df_to_csv_bytes(df: pd.DataFrame):
    return df.to_csv(index=False, sep=";", encoding="utf-8-sig").encode("utf-8-sig")

def df_to_pdf_bytes(df: pd.DataFrame, title="Relatório de Pagamento"):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=landscape(A4))
    width, height = landscape(A4)

    c.setFont("Helvetica-Bold", 14)
    c.drawString(30, height - 30, title)

    c.setFont("Helvetica", 9)
    y = height - 55

    # cabeçalho
    cols = list(df.columns)
    col_text = " | ".join(cols)
    c.drawString(30, y, col_text[:200])  # limita para caber
    y -= 14
    c.line(30, y, width - 30, y)
    y -= 14

    # linhas (simples, sem tabela complexa)
    for _, row in df.iterrows():
        line = " | ".join([str(row[col]) for col in cols])
        c.drawString(30, y, line[:200])
        y -= 12
        if y < 30:
            c.showPage()
            c.setFont("Helvetica", 9)
            y = height - 30

    c.save()
    return buf.getvalue()

# -----------------------
# Upload
# -----------------------
col1, col2 = st.columns(2)
with col1:
    up_func = st.file_uploader("Upload funcionarios.csv", type=["csv"])
with col2:
    up_tot = st.file_uploader("Upload Totais liquidos.xls/.xlsx", type=["xls", "xlsx"])

st.divider()

if up_func and up_tot:
    # Lê funcionarios.csv
    try:
        func = pd.read_csv(up_func, sep=None, engine="python", dtype=str)
    except Exception:
        up_func.seek(0)
        func = pd.read_csv(up_func, sep=";", dtype=str)

    # Lê totais (.xls ou .xlsx), detectando a linha de cabeçalho
    try:
        up_tot.seek(0)
        tot_raw = pd.read_excel(up_tot, header=None)

        header_row = None
        for i in range(len(tot_raw)):
            if "matricula" in str(tot_raw.iloc[i].values).lower():
                header_row = i
                break

        up_tot.seek(0)
        if header_row is not None:
            tot = pd.read_excel(up_tot, skiprows=header_row)
        else:
            tot = pd.read_excel(up_tot)
    except Exception as e:
        st.error(f"Não consegui ler o Excel de totais: {e}")
        st.stop()

    # Detecta colunas automaticamente
    col_matric_func = pick_col(func, ["matricula", "matrícula", "codigo", "código"])
    col_nome = pick_col(func, ["nome"])
    col_cargo = pick_col(func, ["cargo"])
    col_local = pick_col(func, ["localizacao", "localização", "local"])
    col_adm = pick_col(func, ["data_de_admissao", "data de admissao", "data de admissão"])
    col_desl = pick_col(func, ["data_de_desligamento", "data de desligamento"])

    col_matric_tot = pick_col(tot, ["matricula", "matrícula", "codigo", "código"])
    col_val = pick_col(tot, ["valor_liquido", "valor liquido", "valor líquido", "valor"])
    col_pag = pick_col(tot, ["datapagto", "data pagto", "data_pagto", "data pagamento", "datapgto"])
    col_ref = pick_col(tot, ["referencia", "referência", "ref"])

    missing = []
    if not col_matric_func: missing.append("matricula (funcionarios.csv)")
    if not col_nome: missing.append("nome (funcionarios.csv)")
    if not col_cargo: missing.append("cargo (funcionarios.csv)")
    if not col_local: missing.append("localizacao (funcionarios.csv)")
    if not col_adm: missing.append("Data de Admissão (funcionarios.csv)")
    if not col_desl: missing.append("Data de Desligamento (funcionarios.csv)")

    if not col_matric_tot: missing.append("matricula (Totais)")
    if not col_val: missing.append("Valor Liquido (Totais)")
    if not col_pag: missing.append("DataPagto (Totais)")
    if missing:
        st.error("Faltam colunas obrigatórias ou não consegui identificar automaticamente:\n- " + "\n- ".join(missing))
        st.stop()

    referencia_manual = None
    if not col_ref:
        st.warning("Não encontrei a coluna Referencia em Totais. Informe uma data para usar como referência.")
        referencia_manual = st.date_input("Data de referência", format="DD/MM/YYYY")

    # Padroniza matrículas
    func["_matricula"] = func[col_matric_func].map(to_str_matricula)
    tot["_matricula"] = tot[col_matric_tot].map(to_str_matricula)

    # Converte valores e datas
    tot["_valor_liquido"] = tot[col_val].map(parse_money_br)
    tot["_datapagto"] = tot[col_pag].map(parse_date)
    # referência: usa coluna do arquivo ou data informada manualmente
    if col_ref:
        tot["_referencia"] = tot[col_ref].map(format_mes_ano_pt)
    else:
        tot["_referencia"] = format_mes_ano_pt(referencia_manual)

    func["_admissao"] = func[col_adm].map(parse_date)
    func["_deslig"] = func[col_desl].map(parse_date)

    # Merge
    base = tot.merge(
        func,
        on="_matricula",
        how="left",
        suffixes=("", "_func")
    )

    # Monta saída final com os nomes exatos que você pediu
    saida = pd.DataFrame({
        "matricula": base["_matricula"],
        "nome": base[col_nome].fillna(""),
        "cargo": base[col_cargo].fillna(""),
        "localizacao": base[col_local].fillna(""),
        "Data de Admissão": base["_admissao"].dt.strftime("%d/%m/%Y").fillna(""),
        "Data de Desligamento": base["_deslig"].dt.strftime("%d/%m/%Y").fillna(""),
        "Valor Liquido": base["_valor_liquido"].round(2),
        "DataPagto": base["_datapagto"].dt.strftime("%d/%m/%Y").fillna(""),
        "Referencia": base["_referencia"],
    })

    # Ordenação opcional
    st.subheader("Prévia do relatório")
    sort_cols = st.multiselect("Ordenar por", ["localizacao", "nome", "matricula", "DataPagto", "Referencia"], default=["localizacao", "nome"])
    if sort_cols:
        saida = saida.sort_values(by=sort_cols, kind="stable")

    st.dataframe(saida, use_container_width=True, height=520)

    total = saida["Valor Liquido"].sum()
    st.metric("Total (Valor Liquido)", f"R$ {total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    # Pendências (matrícula sem cadastro)
    pend = saida[saida["nome"].astype(str).str.strip() == ""].copy()
    if len(pend) > 0:
        st.warning(f"Há {len(pend)} registros em Totais que não encontrei em funcionarios.csv (matrícula sem cadastro).")
        st.dataframe(pend[["matricula", "Valor Liquido", "DataPagto", "Referencia"]], use_container_width=True)

    st.divider()
    st.subheader("Exportar")

    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button(
            "Baixar CSV (;)",
            data=df_to_csv_bytes(saida),
            file_name="relatorio_pagamento.csv",
            mime="text/csv"
        )
    with c2:
        st.download_button(
            "Baixar Excel (.xlsx)",
            data=df_to_excel_bytes(saida),
            file_name="relatorio_pagamento.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with c3:
        st.download_button(
            "Baixar PDF (simples)",
            data=df_to_pdf_bytes(saida),
            file_name="relatorio_pagamento.pdf",
            mime="application/pdf"
        )
else:
    st.info("Faça upload dos dois arquivos para gerar o relatório.")