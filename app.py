"""
app.py — Preenchimento automático de orçamentos de licitação (Streamlit)

Hierarquia de busca : banco interno  →  SINAPI
Tipo de busca       : correspondência EXATA após normalização (sem fuzzy matching)
Execução            : streamlit run app.py
"""
from __future__ import annotations

import io
import os
import re
import unicodedata
from dataclasses import dataclass
from numbers import Number

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------------
# Constantes
# ----------------------------------------------------------------------------
TOLERANCIA_VALOR = 0.005  # preços que diferem menos que meio centavo = mesmo preço

FILL_AMARELO = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
FILL_LARANJA = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")

ST_OK = "Preenchido"
ST_NAO = "Não encontrado"
ST_AMB = "Ambíguo (valores conflitantes)"
ST_JA = "Já preenchido (mantido)"
ST_MESCLADA = "Célula de valor mesclada"
FONTE_INTERNO = "Banco interno"
FONTE_SINAPI = "SINAPI"
COLUNAS_RELATORIO = ["Linha Excel", "Descrição", "Status", "Fonte", "Valor unitário", "Observação"]
COR_STATUS = {ST_NAO: "#FFFF00", ST_AMB: "#FFC000", ST_MESCLADA: "#FFC000", ST_JA: "#D9D9D9"}


@dataclass
class ConfigPlanilha:
    raw: pd.DataFrame        # aba inteira, sem cabeçalho (índice 0 = linha 1 do Excel)
    aba: str
    linha_cab: int           # nº da linha do cabeçalho no Excel (base 1)
    col_desc: int            # índice da coluna de descrição (base 0)
    col_valor: int           # índice da coluna de valor (base 0)
    col_qtd: int | None = None


# ----------------------------------------------------------------------------
# Normalização e conversões
# ----------------------------------------------------------------------------
def normalizar(texto) -> str:
    """Remove acentos, converte para MAIÚSCULAS e limpa espaços.

    NFD separa a letra do acento ("é" -> "e" + "´") e removemos só as marcas
    combinantes (categoria Mn). Usamos NFD (e não NFKD) de propósito: NFKD também
    converteria "m²" em "m2" e "º" em "o", unificando especificações que podem
    ser diferentes. Espaços/tabs/quebras de linha/NBSP viram um único espaço.
    """
    if texto is None or pd.isna(texto):
        return ""
    s = unicodedata.normalize("NFD", str(texto))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    return re.sub(r"\s+", " ", s.upper()).strip()


def para_numero(valor) -> float | None:
    """Converte para float (aceita 'R$ 1.234,56'). Vazio, NaN, '-', zero e negativos -> None."""
    if valor is None or isinstance(valor, bool):
        return None
    if isinstance(valor, Number):
        numero = float(valor)
    else:
        txt = re.sub(r"[^\d,.\-]", "", str(valor))
        if "," in txt:  # formato brasileiro: '.' = milhar, ',' = decimal
            txt = txt.replace(".", "").replace(",", ".")
        try:
            numero = float(txt)
        except ValueError:
            return None
    return numero if numero == numero and numero > 0 else None  # numero == numero descarta NaN


def _mesma_celula(valor_openpyxl, valor_pandas) -> bool:
    """Trava de segurança: confere se a linha lida pelo pandas é a mesma que o openpyxl vai escrever."""
    if isinstance(valor_openpyxl, str) and valor_openpyxl.startswith("="):
        return True  # fórmula: o openpyxl não expõe o resultado calculado
    if normalizar(valor_openpyxl) == normalizar(valor_pandas):
        return True
    return isinstance(valor_openpyxl, Number) and isinstance(valor_pandas, Number) and valor_openpyxl == valor_pandas


def _fmt_conflito(valores: list[float]) -> str:
    return "Valores conflitantes: " + "; ".join(f"{v:.2f}".replace(".", ",") for v in valores)


# ----------------------------------------------------------------------------
# Leitura (pandas) e índices de busca
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def listar_abas(conteudo: bytes) -> list[str]:
    return pd.ExcelFile(io.BytesIO(conteudo)).sheet_names


@st.cache_data(show_spinner="Lendo planilha...")
def ler_aba(conteudo: bytes, aba: str) -> pd.DataFrame:
    # header=None: o índice do DataFrame bate com a linha do Excel (índice + 1)
    return pd.read_excel(io.BytesIO(conteudo), sheet_name=aba, header=None, dtype=object)


def montar_indice(cfg: ConfigPlanilha) -> tuple[dict[str, float], dict[str, list[float]]]:
    """Retorna (únicos, ambíguos): {descrição normalizada: valor} e {descrição normalizada: [valores]}.

    Se a mesma descrição aparece com preços diferentes, NÃO escolhemos um: vai para 'ambíguos'.
    """
    dados = cfg.raw.iloc[cfg.linha_cab:, [cfg.col_desc, cfg.col_valor]].copy()
    dados.columns = ["desc", "valor"]
    dados["chave"] = dados["desc"].map(normalizar)
    dados["valor"] = dados["valor"].map(para_numero)
    dados = dados[(dados["chave"] != "") & dados["valor"].notna()]

    unicos: dict[str, float] = {}
    ambiguos: dict[str, list[float]] = {}
    for chave, grupo in dados.groupby("chave")["valor"]:
        valores = grupo.tolist()
        if max(valores) - min(valores) <= TOLERANCIA_VALOR:
            unicos[chave] = valores[0]
        else:
            ambiguos[chave] = sorted(set(valores))
    return unicos, ambiguos


# ----------------------------------------------------------------------------
# Preenchimento (escrita com openpyxl para preservar a formatação original)
# ----------------------------------------------------------------------------
def preencher(conteudo_lic: bytes, cfg: ConfigPlanilha, interno, sinapi, sobrescrever: bool):
    ok_int, amb_int = interno
    ok_sin, amb_sin = sinapi

    wb = load_workbook(io.BytesIO(conteudo_lic))
    ws = wb[cfg.aba]
    registros = []

    for i in range(cfg.linha_cab, len(cfg.raw)):  # i base 0 -> linha Excel = i + 1
        descricao = cfg.raw.iat[i, cfg.col_desc]
        chave = normalizar(descricao)
        if not chave:
            continue
        if cfg.col_qtd is not None and not normalizar(cfg.raw.iat[i, cfg.col_qtd]):
            continue  # sem quantidade = título de grupo, não é item

        linha = i + 1
        c_desc = ws.cell(row=linha, column=cfg.col_desc + 1)
        c_valor = ws.cell(row=linha, column=cfg.col_valor + 1)
        if not _mesma_celula(c_desc.value, descricao):
            raise ValueError(
                f"Desalinhamento na linha {linha}: o pandas leu '{descricao}', mas o Excel tem "
                f"'{c_desc.value}'. Nada foi gerado para evitar preencher a linha errada."
            )

        # Hierarquia: banco interno -> SINAPI (correspondência exata da chave normalizada)
        valor, fonte, status, obs = None, "", ST_NAO, ""
        if chave in ok_int:
            valor, fonte, status = ok_int[chave], FONTE_INTERNO, ST_OK
        elif chave in amb_int:  # existe no interno com preços conflitantes: não cai para o SINAPI
            fonte, status, obs = FONTE_INTERNO, ST_AMB, _fmt_conflito(amb_int[chave])
        elif chave in ok_sin:
            valor, fonte, status = ok_sin[chave], FONTE_SINAPI, ST_OK
        elif chave in amb_sin:
            fonte, status, obs = FONTE_SINAPI, ST_AMB, _fmt_conflito(amb_sin[chave])

        if isinstance(c_valor, MergedCell):
            valor, status = None, ST_MESCLADA
        elif status == ST_OK and c_valor.value not in (None, "") and not sobrescrever:
            valor, status, obs = None, ST_JA, f"Valor existente: {c_valor.value}"

        if status == ST_OK:
            c_valor.value = valor
            if c_valor.number_format == "General":
                c_valor.number_format = "#,##0.00"
        elif status in (ST_NAO, ST_AMB, ST_MESCLADA):  # exige preenchimento manual -> destaca
            fill = FILL_AMARELO if status == ST_NAO else FILL_LARANJA
            c_desc.fill = fill
            if not isinstance(c_valor, MergedCell):
                c_valor.fill = fill

        registros.append({
            "Linha Excel": linha, "Descrição": str(descricao), "Status": status,
            "Fonte": fonte, "Valor unitário": valor, "Observação": obs,
        })

    saida = io.BytesIO()
    wb.save(saida)
    return saida.getvalue(), pd.DataFrame(registros, columns=COLUNAS_RELATORIO)


def df_para_xlsx(df: pd.DataFrame) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as escritor:
        df.to_excel(escritor, index=False, sheet_name="Relatório")
    return buf.getvalue()


# ----------------------------------------------------------------------------
# Interface
# ----------------------------------------------------------------------------
def detectar_cabecalho(raw: pd.DataFrame) -> int:
    """Primeira linha (até a 40ª) que contém 'DESCRI...'; senão, linha 1."""
    for i in range(min(40, len(raw))):
        if any("DESCRI" in normalizar(v) for v in raw.iloc[i]):
            return i + 1
    return 1


def opcoes_colunas(raw: pd.DataFrame, linha_cab: int) -> dict[int, str]:
    cab = raw.iloc[linha_cab - 1]
    opcoes = {}
    for j in range(raw.shape[1]):
        titulo = "" if pd.isna(cab.iloc[j]) else str(cab.iloc[j])[:60]
        opcoes[j] = f"{get_column_letter(j + 1)} · {titulo}".rstrip(" ·")
    return opcoes


def sugerir(opcoes: dict[int, str], palavras: tuple[str, ...]) -> int | None:
    for j, rotulo in opcoes.items():
        if any(p in normalizar(rotulo) for p in palavras):
            return j
    return None


def configurar(titulo, arquivo, prefixo, rotulo_valor, usar_qtd=False) -> ConfigPlanilha | None:
    conteudo = arquivo.getvalue()
    with st.expander(titulo, expanded=True):
        c1, c2 = st.columns(2)
        aba = c1.selectbox("Aba", listar_abas(conteudo), key=f"{prefixo}_aba")
        raw = ler_aba(conteudo, aba)
        if raw.empty:
            st.error("A aba selecionada está vazia.")
            return None

        linha_cab = int(c2.number_input(
            "Linha do cabeçalho (nº da linha no Excel)", min_value=1, max_value=len(raw),
            value=detectar_cabecalho(raw), step=1, key=f"{prefixo}_cab_{aba}"))
        opcoes = opcoes_colunas(raw, linha_cab)
        chave_w = f"{prefixo}_{aba}_{linha_cab}"  # muda ao trocar aba/cabeçalho -> recalcula sugestões

        cols = st.columns(3 if usar_qtd else 2)
        col_desc = cols[0].selectbox(
            "Coluna de Descrição/Especificação", list(opcoes), format_func=opcoes.get,
            index=sugerir(opcoes, ("DESCRI", "ESPECIFIC")) or 0, key=f"desc_{chave_w}")
        col_valor = cols[1].selectbox(
            rotulo_valor, list(opcoes), format_func=opcoes.get,
            index=sugerir(opcoes, ("VALOR UNIT", "PRECO UNIT", "CUSTO UNIT", "UNITARIO")) or 0,
            key=f"valor_{chave_w}")
        col_qtd = None
        if usar_qtd:
            s = sugerir(opcoes, ("QUANT", "QTD"))
            col_qtd = cols[2].selectbox(
                "Coluna de Quantidade (opcional)", [None] + list(opcoes),
                format_func=lambda j: "(não usar)" if j is None else opcoes[j],
                index=0 if s is None else s + 1, key=f"qtd_{chave_w}",
                help="Linhas sem quantidade são tratadas como títulos de grupo e ignoradas.")

        prev = raw.head(15).astype(str).replace({"nan": "", "None": "", "NaT": ""})
        prev.columns = [get_column_letter(j + 1) for j in range(prev.shape[1])]
        prev.index = range(1, len(prev) + 1)
        st.caption("Pré-visualização (letras e números = os do Excel):")
        st.dataframe(prev)

    return ConfigPlanilha(raw, aba, linha_cab, col_desc, col_valor, col_qtd)


def exibir_resultado(res: dict) -> None:
    rel: pd.DataFrame = res["relatorio"]
    ok = rel["Status"] == ST_OK

    st.subheader("Resultado")
    m = st.columns(5)
    m[0].metric("Itens analisados", len(rel))
    m[1].metric("Do banco interno", int((ok & (rel["Fonte"] == FONTE_INTERNO)).sum()))
    m[2].metric("Do SINAPI", int((ok & (rel["Fonte"] == FONTE_SINAPI)).sum()))
    m[3].metric("Não encontrados (amarelo)", int((rel["Status"] == ST_NAO).sum()))
    m[4].metric("Ambíguos / outros (laranja)", int(rel["Status"].isin([ST_AMB, ST_MESCLADA]).sum()))

    d1, d2 = st.columns(2)
    d1.download_button(
        "⬇️ Baixar planilha preenchida (.xlsx)", data=res["xlsx"], file_name=res["nome"],
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    d2.download_button(
        "⬇️ Baixar relatório de auditoria (.xlsx)", data=df_para_xlsx(rel),
        file_name=res["nome"].replace("_preenchida", "_relatorio"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    def pintar(linha):
        cor = COR_STATUS.get(linha["Status"])
        return [f"background-color: {cor}; color: black" if cor else ""] * len(linha)

    st.dataframe(rel.style.apply(pintar, axis=1).format({"Valor unitário": "{:,.2f}"}, na_rep=""))


def main() -> None:
    st.set_page_config(page_title="Orçamento de Licitação", page_icon="📊", layout="wide")
    st.title("📊 Preenchimento automático de orçamento de licitação")
    st.caption("Correspondência exata (após normalização) · 1º banco interno, 2º SINAPI · sem fuzzy matching.")

    u1, u2, u3 = st.columns(3)
    f_lic = u1.file_uploader("1 · Planilha da licitação", type=["xlsx"])
    f_int = u2.file_uploader("2 · Banco de dados interno", type=["xlsx", "xlsm", "xls"])
    f_sin = u3.file_uploader("3 · Tabela SINAPI", type=["xlsx", "xlsm", "xls"])
    if not (f_lic and f_int and f_sin):
        st.info("Envie os três arquivos para continuar.")
        return

    st.subheader("Configuração das colunas")
    cfg_lic = configurar("Planilha da licitação", f_lic, "lic", "Coluna de valor unitário (destino)", usar_qtd=True)
    cfg_int = configurar("Banco de dados interno", f_int, "int", "Coluna de valor unitário")
    cfg_sin = configurar("SINAPI", f_sin, "sin", "Coluna de preço (escolha a sua UF / regime)")
    if any(c is None for c in (cfg_lic, cfg_int, cfg_sin)):
        return

    sobrescrever = st.checkbox("Sobrescrever valores que já existem na coluna de destino", value=False)

    if st.button("⚙️ Processar", type="primary"):
        st.session_state.pop("resultado", None)
        with st.spinner("Processando..."):
            try:
                xlsx, rel = preencher(
                    f_lic.getvalue(), cfg_lic, montar_indice(cfg_int), montar_indice(cfg_sin), sobrescrever)
            except ValueError as erro:
                st.error(str(erro))
                return
        base = os.path.splitext(f_lic.name)[0]
        st.session_state["resultado"] = {"xlsx": xlsx, "relatorio": rel, "nome": f"{base}_preenchida.xlsx"}

    # o resultado fica no session_state para sobreviver ao rerun causado pelo clique no download
    if "resultado" in st.session_state:
        exibir_resultado(st.session_state["resultado"])


if __name__ == "__main__":
    main()
