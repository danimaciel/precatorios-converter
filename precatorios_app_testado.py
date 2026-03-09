
import re
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Conversor de Precatórios", layout="wide")
st.title("Conversor de Planilha de Precatórios")
st.write("Envie a planilha .xlsx baixada do sistema e baixe o arquivo final limpo.")


COLUNAS_FINAIS = [
    "ORDEM DE PAGAMENTO",
    "MOMENTO DE APRESENTAÇÃO DO PRECATÓRIO",
    "PROCESSO",
    "PRECATÓRIO",
    "RP",
    "VENCIMENTO",
    "EXEQUENTE",
    "CPF",
    "VALOR DEVIDO / SALDO A PAGAR POR EXEQUENTE",
    "TIPO DE PREFERÊNCIA",
]


def limpar(valor):
    if valor is None or pd.isna(valor):
        return ""
    texto = str(valor).replace("\n", " ").replace("\r", " ").strip()
    texto = re.sub(r"\s+", " ", texto)
    return texto


def separar_exequente_cpf(texto):
    texto = limpar(texto)
    if not texto:
        return "", ""

    # Primeiro tenta achar CPF no fim da célula
    m = re.search(r"^(.*?)\s*-\s*(\d{3}\.\d{3}\.\d{3}-\d{2})", texto)
    if m:
        nome = limpar(m.group(1))
        cpf = m.group(2)
        return nome, cpf

    # Fallback: separa no primeiro hífen encontrado
    partes = re.split(r"\s*-\s*", texto, maxsplit=1)
    if len(partes) == 2:
        return limpar(partes[0]), limpar(partes[1])

    return texto, ""


def carregar_linhas(arquivo):
    wb = load_workbook(arquivo, data_only=True)
    ws = wb.active
    return list(ws.iter_rows(values_only=True))


def converter_planilha(arquivo):
    linhas = carregar_linhas(arquivo)

    registros = []
    bloco_atual = ""

    for linha in linhas:
        vals = [limpar(v) for v in linha]
        primeira = vals[0] if len(vals) > 0 else ""
        texto_linha = " | ".join([v for v in vals if v])
        texto_upper = texto_linha.upper()

        if not texto_linha:
            continue

        # identifica o bloco atual
        if "LISTA CONSOLIDADA - OFÍCIOS PRECATÓRIOS" in texto_upper:
            m = re.search(r"\((.*?)\)", texto_linha)
            bloco_atual = limpar(m.group(1)) if m else ""
            continue

        # descarta linhas que não são dados
        if (
            primeira == "ORDEM DE PAGAMENTO"
            or texto_upper.startswith("MUNICÍPIO DE")
            or "PODER JUDICIÁRIO" in texto_upper
            or "TRIBUNAL REGIONAL" in texto_upper
            or "SECRETARIA DE PRECATÓRIOS" in texto_upper
            or texto_upper.startswith("PREFERÊNCIAS")
            or texto_upper.startswith("ORDEM CRONOLÓGICA")
        ):
            continue

        # só entra se a primeira célula for a ordem de pagamento
        if not primeira.isdigit():
            continue

        nome, cpf = separar_exequente_cpf(vals[9] if len(vals) > 9 else "")

        tipo = "Idoso" if bloco_atual.upper().startswith("PREFER") else bloco_atual

        registros.append({
            "ORDEM DE PAGAMENTO": vals[0] if len(vals) > 0 else "",
            "MOMENTO DE APRESENTAÇÃO DO PRECATÓRIO": vals[1] if len(vals) > 1 else "",
            "PROCESSO": vals[2] if len(vals) > 2 else "",
            "PRECATÓRIO": vals[3] if len(vals) > 3 else "",
            "RP": vals[4] if len(vals) > 4 else "",
            "VENCIMENTO": vals[6] if len(vals) > 6 else "",
            "EXEQUENTE": nome,
            "CPF": cpf,
            "VALOR DEVIDO / SALDO A PAGAR POR EXEQUENTE": vals[14] if len(vals) > 14 else "",
            "TIPO DE PREFERÊNCIA": tipo,
        })

    df_final = pd.DataFrame(registros, columns=COLUNAS_FINAIS)

    if df_final.empty:
        raise ValueError("Nenhum registro foi extraído. Verifique se o layout da planilha é o mesmo do arquivo de exemplo.")

    # ordenação por ordem de pagamento
    df_final["_ord"] = pd.to_numeric(df_final["ORDEM DE PAGAMENTO"], errors="coerce")
    df_final = df_final.sort_values("_ord").drop(columns="_ord")
    df_final = df_final.drop_duplicates()

    return df_final


arquivo = st.file_uploader("Upload da planilha .xlsx", type=["xlsx"])

if arquivo is not None:
    try:
        df_final = converter_planilha(arquivo)

        st.success(f"Planilha processada com sucesso. Registros extraídos: {len(df_final)}")
        st.dataframe(df_final, use_container_width=True)

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df_final.to_excel(writer, index=False, sheet_name="Final")

        st.download_button(
            label="Baixar planilha final",
            data=output.getvalue(),
            file_name="2_Final.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except Exception as e:
        st.error(f"Erro ao processar a planilha: {e}")
