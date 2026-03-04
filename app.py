import io
import re
import html
import pandas as pd
import streamlit as st

# (Opcional) PDF simples
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer


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

st.info("1) Leia as diretrizes na barra lateral  •  2) Faça upload dos dois arquivos  •  3) Gere e exporte o relatório")

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

def format_money_br(x):
    if pd.isna(x):
        return "R$ 0,00"
    try:
        n = float(x)
        return f"R$ {n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "R$ 0,00"

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

def template_funcionarios_csv_bytes():
    modelo = pd.DataFrame([
        {
            "matricula": "1234",
            "nome": "NOME SOBRENOME",
            "cargo": "AUXILIAR DE SERVIÇOS GERAIS",
            "localizacao": "UNIDADE X - 123,45",
            "Data de Admissão": "01/02/2024",
            "Data de Desligamento": "",
        }
    ])
    return df_to_csv_bytes(modelo)

def template_totais_excel_bytes():
    modelo = pd.DataFrame([
        {
            "matricula": "1234",
            "Valor Liquido": "1882,59",
            "DataPagto": "06/03/2026",
            "Referencia": "Março/2026",
        }
    ])
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        modelo.to_excel(writer, index=False, sheet_name="Totais")
    return output.getvalue()

st.sidebar.subheader("Diretrizes de importação")
st.sidebar.markdown(
    """
    **1) funcionarios.csv (separador ;)**
        - Extrair no Ahgora em: Relatórios > Funcionários
    - Colunas obrigatórias e nesta ordem:
      matricula, nome, cargo, localizacao, Data de Admissão, Data de Desligamento
    - matricula: somente números, sem espaços e sem texto
    - Datas: DD/MM/AAAA (ex.: 06/03/2026)

    **2) Totais liquidos.xls/.xlsx (Phoenix)**
    - Colunas obrigatórias: matricula, Valor Liquido, DataPagto
    - Coluna opcional: Referencia
    - Valor Liquido: 1882,59 ou 1882.59
    - Remover a última linha de total do Phoenix

    **Boas práticas**
    - Não mesclar células no Excel
    - Evitar linhas em branco antes do cabeçalho
    """
)

st.sidebar.download_button(
    "Baixar modelo funcionarios.csv",
    data=template_funcionarios_csv_bytes(),
    file_name="modelo_funcionarios.csv",
    mime="text/csv",
)
st.sidebar.download_button(
    "Baixar modelo Totais.xlsx",
    data=template_totais_excel_bytes(),
    file_name="modelo_totais.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
st.sidebar.divider()

def df_to_pdf_bytes(df: pd.DataFrame, title="Relatório de Pagamento"):
    def money_br(v):
        try:
            n = float(v)
            return f"R$ {n:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return str(v)

    rel = df.copy()
    rel["Valor Liquido"] = rel["Valor Liquido"].map(money_br)

    total = pd.to_numeric(df["Valor Liquido"], errors="coerce").fillna(0).sum()
    total_fmt = money_br(total)
    periodo = ""
    if "Referencia" in df.columns and not df["Referencia"].dropna().empty:
        periodo = str(df["Referencia"].dropna().iloc[0])

    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=landscape(A4),
        leftMargin=20,
        rightMargin=20,
        topMargin=20,
        bottomMargin=20,
    )

    styles = getSampleStyleSheet()
    cell_style = styles["Normal"].clone("CellStyle")
    cell_style.fontName = "Helvetica"
    cell_style.fontSize = 7
    cell_style.leading = 8

    header_style = styles["Normal"].clone("HeaderStyle")
    header_style.fontName = "Helvetica-Bold"
    header_style.fontSize = 8
    header_style.leading = 9

    story = []
    story.append(Paragraph(f"<b>{title}</b>", styles["Title"]))
    if periodo:
        story.append(Paragraph(f"Período de referência: <b>{periodo}</b>", styles["Normal"]))
    story.append(Paragraph(f"Total líquido: <b>{total_fmt}</b>", styles["Normal"]))
    story.append(Spacer(1, 10))

    headers = list(rel.columns)
    header_row = [Paragraph(f"<b>{html.escape(str(h))}</b>", header_style) for h in headers]
    body_rows = []
    for row in rel.fillna("").astype(str).values.tolist():
        body_rows.append([Paragraph(html.escape(str(value)), cell_style) for value in row])

    total_row = [""] * len(headers)
    total_row[5] = Paragraph("<b>TOTAL DA FOLHA</b>", cell_style)
    total_row[6] = Paragraph(f"<b>{html.escape(total_fmt)}</b>", cell_style)

    data = [header_row] + body_rows + [total_row]

    col_widths = [45, 165, 115, 138, 66, 70, 78, 56, 55]
    table = Table(data, colWidths=col_widths, repeatRows=1)

    total_row_idx = len(data) - 1

    table_style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("ALIGN", (0, 1), (0, -1), "CENTER"),
        ("ALIGN", (6, 1), (6, -1), "RIGHT"),
        ("ALIGN", (7, 1), (8, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#D0D7DE")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F9FC")]),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("BACKGROUND", (0, total_row_idx), (-1, total_row_idx), colors.HexColor("#EAF2FB")),
        ("SPAN", (0, total_row_idx), (5, total_row_idx)),
        ("ALIGN", (0, total_row_idx), (5, total_row_idx), "RIGHT"),
        ("ALIGN", (6, total_row_idx), (6, total_row_idx), "RIGHT"),
        ("LINEABOVE", (0, total_row_idx), (-1, total_row_idx), 1, colors.HexColor("#1f4e79")),
    ]
    table.setStyle(TableStyle(table_style))

    story.append(table)

    doc.build(story)
    pdf_bytes = pdf_buffer.getvalue()
    pdf_buffer.seek(0)
    return pdf_bytes

# -----------------------
# Upload
# -----------------------
st.sidebar.subheader("Importar arquivos")
up_func = st.sidebar.file_uploader("Upload funcionarios.csv", type=["csv"])
up_tot = st.sidebar.file_uploader("Upload Totais liquidos.xls/.xlsx", type=["xls", "xlsx"])

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

    # Proteção: remove linha de total do Phoenix (quando vier no final ou no meio)
    tot = tot[tot[col_matric_tot].notna()].copy()
    mask_total_phoenix = tot[col_matric_tot].astype(str).str.strip().str.lower().str.contains("total", na=False)
    if mask_total_phoenix.any():
        tot = tot[~mask_total_phoenix].copy()

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

    saida_preview = saida.copy()
    saida_preview["Valor Liquido"] = saida_preview["Valor Liquido"].map(format_money_br)
    st.dataframe(saida_preview, use_container_width=True, height=520)

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
            "Baixar PDF (executivo)",
            data=df_to_pdf_bytes(saida),
            file_name="relatorio_pagamento.pdf",
            mime="application/pdf"
        )
else:
    st.info("Faça upload dos dois arquivos para gerar o relatório.")