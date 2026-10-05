"""
app.py — interface do técnico orçamentista (Streamlit).

Três telas:
  Preencher planilha -> envia a planilha da licitação e recebe ela preenchida + relatório
  Cadastrar preços   -> alimenta o banco com planilhas de licitações passadas
  Consultar banco    -> busca preços, exporta e desfaz importações

Regras de negócio: core.py. Banco de dados: storage.py.
Logo: salve o arquivo em assets/logo.png (também aceita .svg, .jpg, .webp).
"""
from __future__ import annotations

import base64
import html
import io
import os
from dataclasses import dataclass
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

import core
import storage

# ----------------------------------------------------------------------------
# Identidade visual
# ----------------------------------------------------------------------------
AZUL_MARINHO = "#0E2A47"
AZUL = "#1565A8"
NEVOA = "#EEF3F8"
TINTA = "#16263A"
CINZA_TEXTO = "#5A6B7D"

# As cores dos status são as MESMAS que o core.py pinta no Excel.
COR_BARRA = {core.ST_OK: AZUL, **core.COR_STATUS}
COR_CELULA_STATUS = {core.ST_OK: "#D6E6F5", **core.COR_STATUS}
COR_SITUACAO = {storage.SIT_NOVO: "#D8EFE0", storage.SIT_EXISTE: "#E3E3E3", storage.SIT_REPETIDO: "#FFE7A8"}
ORDEM_STATUS = [core.ST_OK, core.ST_NAO, core.ST_AMB, core.ST_MESCLADA, core.ST_JA]
LEGENDA_STATUS = {
    core.ST_OK: "preenchidos",
    core.ST_NAO: "sem preço encontrado",
    core.ST_AMB: "com preços conflitantes",
    core.ST_MESCLADA: "com célula de valor mesclada",
    core.ST_JA: "já tinham valor (mantidos)",
}

CSS = f"""
<style>
.block-container {{ padding-top: 1.5rem; max-width: 1180px; }}

.atl-topo {{
  display: flex; align-items: center; gap: 1.25rem; flex-wrap: wrap;
  padding: 0 0 1rem; margin-bottom: 1.5rem; border-bottom: 3px solid {AZUL_MARINHO};
}}
.atl-topo img {{ height: 52px; width: auto; display: block; }}
.atl-marca {{ font-size: 1.35rem; font-weight: 700; color: {AZUL_MARINHO}; }}
.atl-titulo {{ border-left: 1px solid #C9D3DE; padding-left: 1.25rem; }}
.atl-titulo .nome {{ font-size: 1.2rem; font-weight: 600; color: {TINTA}; line-height: 1.3; }}
.atl-titulo .sub {{ font-size: .9rem; color: {CINZA_TEXTO}; }}
.atl-banco {{
  margin-left: auto; font-size: .85rem; padding: .35rem .8rem; border-radius: 999px;
  background: {NEVOA}; color: {TINTA}; white-space: nowrap;
}}
.atl-banco.alerta {{ background: #FFF4D6; color: #7A5300; }}

.atl-resumo {{ font-size: 1.05rem; color: {TINTA}; margin: .25rem 0 0; }}
.atl-resumo strong {{ font-size: 2rem; color: {AZUL_MARINHO}; font-weight: 700; margin-right: .35rem; }}
.atl-barra {{
  display: flex; height: 20px; border-radius: 3px; overflow: hidden;
  background: #E6EBF0; margin: .6rem 0 .7rem; border: 1px solid #D5DDE6;
}}
.atl-barra span {{ display: block; height: 100%; }}
.atl-legenda {{ display: flex; flex-wrap: wrap; gap: .35rem 1.4rem; font-size: .9rem; color: {TINTA}; margin-bottom: 1rem; }}
.atl-legenda i {{
  display: inline-block; width: .8rem; height: .8rem; border-radius: 2px; margin-right: .4rem;
  vertical-align: -1px; border: 1px solid rgba(0, 0, 0, .18);
}}

@media (max-width: 640px) {{
  .atl-titulo {{ border-left: 0; padding-left: 0; }}
  .atl-banco {{ margin-left: 0; }}
}}
</style>
"""


# ----------------------------------------------------------------------------
# Formatação
# ----------------------------------------------------------------------------
def fmt_int(n: int) -> str:
    return f"{n:,}".replace(",", ".")


def brl(valor) -> str:
    if valor is None or pd.isna(valor):
        return ""
    return "R$ " + f"{float(valor):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def letra(col: int | None) -> str:
    return "" if col is None else get_column_letter(col + 1)


def arquivo_logo() -> Path | None:
    pasta = Path(__file__).parent / "assets"
    for ext in ("png", "svg", "jpg", "jpeg", "webp"):
        caminho = pasta / f"logo.{ext}"
        if caminho.exists():
            return caminho
    return None


def logo_html() -> str:
    caminho = arquivo_logo()
    if caminho is None:
        return '<span class="atl-marca">Atlântico Engenharia</span>'
    mime = {"svg": "image/svg+xml", "jpg": "image/jpeg", "jpeg": "image/jpeg"}.get(caminho.suffix[1:], f"image/{caminho.suffix[1:]}")
    b64 = base64.b64encode(caminho.read_bytes()).decode()
    return f'<img src="data:{mime};base64,{b64}" alt="Atlântico Engenharia">'


def avisar(chave: str, mensagem: str) -> None:
    """Guarda uma mensagem de sucesso para mostrar depois do st.rerun()."""
    st.session_state[f"aviso_{chave}"] = mensagem


def mostrar_aviso(chave: str) -> None:
    mensagem = st.session_state.pop(f"aviso_{chave}", None)
    if mensagem:
        st.success(mensagem)


# ----------------------------------------------------------------------------
# Leitura de planilhas
# ----------------------------------------------------------------------------
@st.cache_data(show_spinner=False, max_entries=16)
def listar_abas(conteudo: bytes) -> list[str]:
    wb = load_workbook(io.BytesIO(conteudo), read_only=True)
    abas = wb.sheetnames
    wb.close()
    return abas


@st.cache_data(show_spinner="Lendo a planilha...", max_entries=16)
def ler_aba(conteudo: bytes, aba: str) -> pd.DataFrame:
    """Aba inteira, sem cabeçalho. Índice 0 = linha 1 do Excel.

    Lemos com o openpyxl (e não com pd.read_excel) para que a grade seja exatamente
    a mesma que o core.preencher vai escrever, linha por linha.
    """
    wb = load_workbook(io.BytesIO(conteudo), data_only=True)
    linhas = list(wb[aba].iter_rows(values_only=True))
    wb.close()
    return pd.DataFrame(linhas, dtype=object)


# ----------------------------------------------------------------------------
# Detecção automática de cabeçalho e colunas
# ----------------------------------------------------------------------------
@dataclass(frozen=True)
class Campo:
    nome: str
    rotulo: str
    chaves: tuple[tuple[str, ...], ...]   # tentativas em ordem de preferência
    evitar: tuple[str, ...] = ()
    obrigatorio: bool = True
    ajuda: str | None = None


DESC = Campo("desc", "Descrição do item", (("DESCRI", "ESPECIFICA", "DISCRIMINA"),))
VALOR = Campo("valor", "Valor unitário", (("UNIT",), ("VALOR", "PRECO", "CUSTO")), evitar=("TOTAL", "QUANT"))
CAMPOS_LICITACAO = [
    DESC,
    Campo("valor", "Valor unitário (onde o preço será escrito)", VALOR.chaves, VALOR.evitar),
    Campo("qtd", "Quantidade", (("QUANT", "QTD"),), obrigatorio=False,
          ajuda="Linhas sem quantidade são tratadas como títulos de grupo e não são preenchidas."),
]
CAMPOS_SINAPI = [DESC, Campo("valor", "Preço", (("PRECO", "CUSTO", "VALOR"),), evitar=("TOTAL",))]
CAMPOS_CADASTRO = [
    DESC,
    VALOR,
    Campo("codigo", "Código", (("CODIGO", "COD."),), obrigatorio=False),
    Campo("fornecedor", "Fornecedor", (("FORNECEDOR", "EMPRESA", "LICITANTE"),), obrigatorio=False),
    Campo("vendedor", "Vendedor", (("VENDEDOR", "REPRESENTANTE", "CONTATO"),), obrigatorio=False),
    Campo("data", "Data da cotação", (("DATA",),), obrigatorio=False),
]
GRUPOS_CABECALHO = (DESC.chaves[0], ("UNIT", "VALOR", "PRECO", "CUSTO"), ("QUANT", "QTD"))


def _linha_norm(raw: pd.DataFrame, i: int) -> list[str]:
    return [core.normalizar(v) for v in raw.iloc[i].tolist()]


def detectar_cabecalho(raw: pd.DataFrame, limite: int = 60) -> int:
    """Linha (base 1) que mais parece cabeçalho: precisa ter 'descrição' e, de preferência, valor e quantidade."""
    melhor, melhor_pontos = 1, 0
    for i in range(min(len(raw), limite)):
        celulas = _linha_norm(raw, i)
        grupos = [any(any(k in c for k in grupo) for c in celulas) for grupo in GRUPOS_CABECALHO]
        pontos = sum(grupos) if grupos[0] else 0
        if pontos > melhor_pontos:
            melhor, melhor_pontos = i + 1, pontos
    return melhor


def detectar_coluna(raw: pd.DataFrame, linha_cab: int, campo: Campo) -> int | None:
    if not 1 <= linha_cab <= len(raw):
        return None
    celulas = _linha_norm(raw, linha_cab - 1)
    for tentativa in campo.chaves:
        for j, c in enumerate(celulas):
            if any(k in c for k in tentativa) and not any(e in c for e in campo.evitar):
                return j
    return None


@dataclass
class Selecao:
    raw: pd.DataFrame
    aba: str
    linha_cab: int
    cols: dict[str, int | None]
    chave: str  # prefixo único para widgets que dependem desta seleção


def seletor_planilha(prefixo: str, arquivo, campos: list[Campo], previa: bool = True) -> Selecao | None:
    """Aba, linha do cabeçalho e colunas, já sugeridos automaticamente. O usuário só confere."""
    conteudo = arquivo.getvalue()
    try:
        abas = listar_abas(conteudo)
    except Exception:
        st.error("Não foi possível abrir este arquivo. Abra no Excel, salve como .xlsx e envie de novo.")
        return None

    base = f"{prefixo}_{arquivo.name}_{arquivo.size}"
    c_aba, c_cab = st.columns([3, 1])
    aba = c_aba.selectbox("Aba", abas, key=f"{base}_aba")
    raw = ler_aba(conteudo, aba)
    if raw.empty:
        st.warning("Esta aba está vazia. Escolha outra aba.")
        return None
    linha_cab = int(c_cab.number_input(
        "Linha do cabeçalho", min_value=1, max_value=len(raw), value=detectar_cabecalho(raw),
        key=f"{base}_{aba}_cab", help="Linha onde estão os títulos das colunas (Descrição, Valor...).",
    ))

    chave = f"{base}_{aba}_{linha_cab}"
    titulos = raw.iloc[linha_cab - 1].tolist()
    nomes = {j: f"Coluna {letra(j)}: {core.texto(titulos[j])[:45] or '(sem título)'}" for j in range(raw.shape[1])}
    opcoes = list(range(raw.shape[1]))

    cols: dict[str, int | None] = {}
    for inicio in range(0, len(campos), 3):
        grade = st.columns(3)
        for pos, campo in enumerate(campos[inicio:inicio + 3]):
            sugerida = detectar_coluna(raw, linha_cab, campo)
            if campo.obrigatorio:
                ops, idx = opcoes, (sugerida if sugerida is not None else 0)
            else:
                ops, idx = [None] + opcoes, (0 if sugerida is None else sugerida + 1)
            cols[campo.nome] = grade[pos].selectbox(
                campo.rotulo, ops, index=idx, key=f"{chave}_{campo.nome}", help=campo.ajuda,
                format_func=lambda j: "Não tem" if j is None else nomes[j],
            )

    usadas = [c for c in campos if cols[c.nome] is not None]
    repetidas = {cols[c.nome] for c in usadas if sum(cols[o.nome] == cols[c.nome] for o in usadas) > 1}
    if repetidas:
        st.error(f"A coluna {', '.join(letra(j) for j in sorted(repetidas))} foi escolhida para mais de um campo. Cada campo precisa de uma coluna diferente.")
        return None

    if previa:
        fatia = raw.iloc[linha_cab:linha_cab + 6]
        tabela = pd.DataFrame({c.rotulo: fatia[cols[c.nome]].map(core.texto) for c in usadas})
        tabela.index = [f"Linha {i + 1}" for i in fatia.index]
        with st.expander("Conferir as primeiras linhas lidas"):
            st.dataframe(tabela)

    return Selecao(raw=raw, aba=aba, linha_cab=linha_cab, cols=cols, chave=chave)


# ----------------------------------------------------------------------------
# Banco de dados
# ----------------------------------------------------------------------------
def _url_banco() -> str | None:
    try:
        url = st.secrets.get("DATABASE_URL")
    except Exception:  # sem arquivo de secrets (rodando localmente)
        url = None
    return url or os.environ.get("DATABASE_URL")


@st.cache_resource(show_spinner="Conectando ao banco de preços...")
def obter_engine():
    engine = storage.criar_engine(_url_banco())
    storage.inicializar(engine)
    return engine


@st.cache_data(ttl=300, show_spinner=False)
def carregar_banco(_engine) -> pd.DataFrame:
    return storage.ler_precos(_engine)


# ----------------------------------------------------------------------------
# Componentes
# ----------------------------------------------------------------------------
def topo(engine, banco: pd.DataFrame | None) -> None:
    if engine is None or banco is None:
        estado, classe = "Banco de preços indisponível", "atl-banco alerta"
    elif not storage.eh_remoto(engine):
        estado, classe = f"Banco temporário (teste): {fmt_int(len(banco))} preços", "atl-banco alerta"
    else:
        estado, classe = f"Banco de preços: {fmt_int(len(banco))} registros", "atl-banco"
    st.markdown(
        f"""<div class="atl-topo">{logo_html()}
        <div class="atl-titulo"><div class="nome">Orçamento de licitações</div>
        <div class="sub">Preenchimento de planilhas com o banco de preços da empresa e o SINAPI</div></div>
        <div class="{classe}">{html.escape(estado)}</div></div>""",
        unsafe_allow_html=True,
    )


def barra_cobertura(contagem: pd.Series, total: int) -> None:
    ok = int(contagem.get(core.ST_OK, 0))
    segmentos, legenda = [], []
    for status in ORDEM_STATUS:
        n = int(contagem.get(status, 0))
        if not n:
            continue
        segmentos.append(f'<span style="width:{n / total * 100:.3f}%;background:{COR_BARRA[status]}" title="{n} {LEGENDA_STATUS[status]}"></span>')
        legenda.append(f'<span><i style="background:{COR_BARRA[status]}"></i>{fmt_int(n)} {LEGENDA_STATUS[status]}</span>')
    st.markdown(
        f"""<p class="atl-resumo"><strong>{fmt_int(ok)}</strong>de {fmt_int(total)} itens preenchidos</p>
        <div class="atl-barra" role="img" aria-label="{ok} de {total} itens preenchidos">{''.join(segmentos)}</div>
        <div class="atl-legenda">{''.join(legenda)}</div>""",
        unsafe_allow_html=True,
    )


def _pintar(cores: dict[str, str]):
    def estilo(valor):
        cor = cores.get(valor)
        return f"background-color: {cor}; color: #1A1A1A" if cor else ""
    return estilo


def tabela_relatorio(df: pd.DataFrame) -> None:
    vis = df.copy()
    vis["Valor unitário"] = vis["Valor unitário"].map(brl)
    st.dataframe(
        vis.style.map(_pintar(COR_CELULA_STATUS), subset=["Status"]),
        hide_index=True,
        column_config={
            "Linha Excel": st.column_config.NumberColumn("Linha", format="%d", width="small"),
            "Descrição": st.column_config.TextColumn(width="large"),
        },
    )


# ----------------------------------------------------------------------------
# Tela 1: preencher planilha
# ----------------------------------------------------------------------------
def tela_preencher(banco: pd.DataFrame | None) -> None:
    st.subheader("1. Planilha da licitação")
    arquivo = st.file_uploader(
        "Planilha que será preenchida", type=["xlsx", "xlsm"], key="up_licitacao",
        help="Arquivos .xls antigos: abra no Excel e salve como .xlsx.",
    )
    if arquivo is None:
        st.info(
            "Envie a planilha da licitação para começar. O app procura o preço de cada item no banco "
            "da empresa e no SINAPI e devolve a planilha preenchida, com os itens sem preço destacados em cor."
        )
        return
    selecao = seletor_planilha("lic", arquivo, CAMPOS_LICITACAO)
    if selecao is None:
        return

    st.subheader("2. Fontes de preço")
    tem_banco = banco is not None and not banco.empty
    estrategia = core.ESTRATEGIAS[0]
    col_banco, col_sinapi = st.columns(2, gap="large")
    with col_banco:
        st.markdown("**Banco de preços da empresa** (consultado primeiro)")
        if banco is None:
            st.caption("Indisponível no momento. O preenchimento vai usar só o SINAPI.")
        elif banco.empty:
            st.caption("Ainda não há preços cadastrados. Use a aba Cadastrar preços para alimentar o banco.")
        else:
            st.caption(f"{fmt_int(len(banco))} preços cadastrados.")
            estrategia = st.radio(
                "Quando o item tiver mais de um preço no banco, usar", core.ESTRATEGIAS, horizontal=True, key="estrategia",
                help="Mais recente: a cotação de data mais nova. Menor preço: o mais barato. Mediana: o valor do meio.",
            )
    with col_sinapi:
        st.markdown("**Tabela SINAPI** (opcional, consultada depois do banco)")
        arquivo_sinapi = st.file_uploader("Planilha do SINAPI", type=["xlsx", "xlsm"], key="up_sinapi",
                                          label_visibility="collapsed")

    cfg_sinapi = None
    if arquivo_sinapi is not None:
        with st.expander("Colunas do SINAPI (detectadas automaticamente)"):
            sel_sinapi = seletor_planilha("sin", arquivo_sinapi, CAMPOS_SINAPI, previa=False)
        if sel_sinapi is not None:
            cfg_sinapi = core.ConfigPlanilha(
                raw=sel_sinapi.raw, aba=sel_sinapi.aba, linha_cab=sel_sinapi.linha_cab,
                col_desc=sel_sinapi.cols["desc"], col_valor=sel_sinapi.cols["valor"],
            )
            st.caption(f"SINAPI: descrição na coluna {letra(cfg_sinapi.col_desc)}, preço na coluna "
                       f"{letra(cfg_sinapi.col_valor)}, cabeçalho na linha {cfg_sinapi.linha_cab}.")

    st.subheader("3. Preencher")
    sobrescrever = st.toggle(
        "Substituir valores que já estão na planilha", value=False,
        help="Desligado: itens que já têm valor ficam como estão e aparecem no relatório como mantidos.",
    )
    sem_fonte = not tem_banco and cfg_sinapi is None
    if sem_fonte:
        st.warning("Não há fonte de preço disponível. Envie a tabela SINAPI ou cadastre preços no banco.")

    assinatura = f"{arquivo.name}-{arquivo.size}"
    if st.button("Preencher planilha", type="primary", disabled=sem_fonte):
        cfg = core.ConfigPlanilha(
            raw=selecao.raw, aba=selecao.aba, linha_cab=selecao.linha_cab,
            col_desc=selecao.cols["desc"], col_valor=selecao.cols["valor"], col_qtd=selecao.cols["qtd"],
        )
        with st.spinner("Procurando os preços..."):
            interno = core.montar_indice_banco(banco, estrategia) if tem_banco else core.Indice()
            sinapi = core.montar_indice(cfg_sinapi) if cfg_sinapi is not None else core.Indice()
            try:
                saida, relatorio = core.preencher(arquivo.getvalue(), cfg, interno, sinapi, sobrescrever)
            except ValueError as erro:
                st.error(f"{erro} Confira a aba e a linha do cabeçalho e tente de novo.")
                return
        st.session_state["resultado"] = {
            "assinatura": assinatura, "planilha": saida, "relatorio": relatorio,
            "nome": f"{Path(arquivo.name).stem}_preenchida.xlsx", "xlsm": arquivo.name.lower().endswith(".xlsm"),
        }

    resultado = st.session_state.get("resultado")
    if resultado and resultado["assinatura"] == assinatura:
        mostrar_resultado(resultado)


def mostrar_resultado(resultado: dict) -> None:
    relatorio: pd.DataFrame = resultado["relatorio"]
    st.subheader("4. Resultado")
    if relatorio.empty:
        st.warning(
            "Nenhum item foi encontrado abaixo do cabeçalho. Confira a linha do cabeçalho, a coluna de "
            "descrição e, se escolheu, a coluna de quantidade."
        )
        return

    barra_cobertura(relatorio["Status"].value_counts(), len(relatorio))

    c1, c2, _ = st.columns([1.3, 1, 1.2])
    c1.download_button(
        "Baixar planilha preenchida", resultado["planilha"], file_name=resultado["nome"], type="primary",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    c2.download_button(
        "Baixar relatório", core.df_para_xlsx(relatorio), file_name=resultado["nome"].replace("_preenchida", "_relatorio"),
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    if resultado["xlsm"]:
        st.caption("A planilha original tinha macros. A versão preenchida sai em .xlsx, sem as macros.")

    pendentes = relatorio[relatorio["Status"].isin([core.ST_NAO, core.ST_AMB, core.ST_MESCLADA])]
    opcoes = [f"Itens para revisar ({fmt_int(len(pendentes))})", f"Todos os itens ({fmt_int(len(relatorio))})"]
    escolha = st.radio("Mostrar", opcoes, horizontal=True, key="filtro_resultado", label_visibility="collapsed")
    if escolha == opcoes[0]:
        if pendentes.empty:
            st.success("Todos os itens receberam preço. Nada para revisar.")
            return
        st.caption("Estes itens ficaram destacados com a mesma cor na planilha baixada e precisam de preço manual.")
        tabela_relatorio(pendentes)
    else:
        tabela_relatorio(relatorio)


# ----------------------------------------------------------------------------
# Tela 2: cadastrar preços
# ----------------------------------------------------------------------------
def tela_cadastrar(engine) -> None:
    mostrar_aviso("cadastro")
    if engine is None:
        st.error("O banco de preços não está acessível agora, então não é possível cadastrar. "
                 "Confira a DATABASE_URL em Manage app > Settings > Secrets.")
        return
    if not storage.eh_remoto(engine):
        st.warning("Modo de teste: este banco é temporário e é apagado quando o app reinicia. "
                   "Para guardar os preços de vez, configure a DATABASE_URL nos Secrets do Streamlit.")

    st.subheader("1. Planilha de uma licitação passada")
    arquivo = st.file_uploader(
        "Planilha com preços já cotados ou contratados", type=["xlsx", "xlsm"], key="up_cadastro",
        help="Cada linha com descrição e valor unitário vira um registro no banco.",
    )
    if arquivo is None:
        st.info("Envie uma planilha de licitação passada para alimentar o banco de preços. "
                "Antes de gravar, você confere o que é novo e o que já está cadastrado.")
        return
    selecao = seletor_planilha("cad", arquivo, CAMPOS_CADASTRO)
    if selecao is None:
        return
    cols = selecao.cols

    st.subheader("2. Dados da licitação")
    c1, c2 = st.columns(2)
    origem = c1.text_input(
        "Identificação da licitação", value=Path(arquivo.name).stem, key=f"{selecao.chave}_origem",
        help="Ex.: Pregão 12/2025, TRE-DF. Fica gravado junto com cada preço.",
    )
    data_fixa = None
    if cols["data"] is None:
        data_fixa = c2.date_input("Data da cotação (vale para todos os itens)", value=None, format="DD/MM/YYYY",
                                  key=f"{selecao.chave}_fixo_data")
    else:
        c2.caption(f"Data da cotação lida da coluna {letra(cols['data'])}.")
    c3, c4 = st.columns(2)
    fornecedor_fixo = vendedor_fixo = None
    if cols["fornecedor"] is None:
        fornecedor_fixo = c3.text_input("Fornecedor (vale para todos os itens)", key=f"{selecao.chave}_forn")
    else:
        c3.caption(f"Fornecedor lido da coluna {letra(cols['fornecedor'])}.")
    if cols["vendedor"] is None:
        vendedor_fixo = c4.text_input("Vendedor (vale para todos os itens)", key=f"{selecao.chave}_vend")
    else:
        c4.caption(f"Vendedor lido da coluna {letra(cols['vendedor'])}.")

    mapa = core.MapaIngestao(
        linha_cab=selecao.linha_cab, col_desc=cols["desc"], col_valor=cols["valor"], col_codigo=cols["codigo"],
        col_fornecedor=cols["fornecedor"], fornecedor_fixo=(fornecedor_fixo or "").strip() or None,
        col_vendedor=cols["vendedor"], vendedor_fixo=(vendedor_fixo or "").strip() or None,
        col_data=cols["data"], data_fixa=data_fixa,
    )
    validos, ignorados = core.extrair_registros(selecao.raw, mapa)

    st.subheader("3. Conferir e gravar")
    if not validos:
        st.warning("Nenhuma linha com descrição e valor unitário válido. Confira a linha do cabeçalho e as colunas.")
        if ignorados:
            st.dataframe(pd.DataFrame(ignorados), hide_index=True)
        return
    try:
        situacoes = storage.classificar(engine, validos)
    except Exception as erro:
        st.error(f"Não foi possível consultar o banco: {erro}")
        return

    tabela = pd.DataFrame(validos)
    tabela.insert(0, "Situação", situacoes)
    novos = int((tabela["Situação"] == storage.SIT_NOVO).sum())
    existentes = int((tabela["Situação"] == storage.SIT_EXISTE).sum())
    repetidos = int((tabela["Situação"] == storage.SIT_REPETIDO).sum())
    texto_resumo = (f"**{fmt_int(len(tabela))} preços encontrados:** {fmt_int(novos)} novos, "
                    f"{fmt_int(existentes)} já estão no banco e {fmt_int(repetidos)} se repetem na própria planilha.")
    if ignorados:
        texto_resumo += f" {fmt_int(len(ignorados))} linhas foram ignoradas."
    st.markdown(texto_resumo)

    vis = pd.DataFrame({
        "Situação": tabela["Situação"],
        "Linha": tabela["linha_excel"],
        "Descrição": tabela["descricao"],
        "Valor unitário": tabela["valor_unitario"].map(brl),
        "Fornecedor": tabela["fornecedor"],
        "Data": pd.to_datetime(tabela["data_cotacao"], errors="coerce").dt.strftime("%d/%m/%Y").fillna(""),
    })
    st.dataframe(vis.style.map(_pintar(COR_SITUACAO), subset=["Situação"]), hide_index=True,
                 column_config={"Descrição": st.column_config.TextColumn(width="large")})
    if ignorados:
        with st.expander(f"Linhas ignoradas ({fmt_int(len(ignorados))})"):
            st.dataframe(pd.DataFrame(ignorados), hide_index=True)

    if st.button(f"Gravar {fmt_int(novos)} preços novos no banco", type="primary", disabled=novos == 0):
        if not origem.strip():
            st.error("Informe a identificação da licitação antes de gravar.")
            return
        with st.spinner("Gravando..."):
            resultado = storage.inserir(engine, validos, origem.strip())
        carregar_banco.clear()
        avisar("cadastro", f"{fmt_int(resultado['inseridos'])} preços gravados de \"{origem.strip()}\". "
                           "Se precisar desfazer, remova esta importação na aba Consultar banco.")
        st.rerun()
    if novos == 0:
        st.caption("Todos os preços desta planilha já estão no banco.")


# ----------------------------------------------------------------------------
# Tela 3: consultar banco
# ----------------------------------------------------------------------------
def tela_consultar(engine, banco: pd.DataFrame | None) -> None:
    mostrar_aviso("consulta")
    if engine is None or banco is None:
        st.error("O banco de preços não está acessível agora. Confira a DATABASE_URL em Manage app > Settings > Secrets.")
        return
    if banco.empty:
        st.info("O banco ainda está vazio. Cadastre a primeira planilha na aba Cadastrar preços.")
        return

    busca = st.text_input("Buscar item", placeholder="Ex.: cabo flexível 2,5 mm²",
                          help="Mostra os itens que contêm todas as palavras digitadas, sem diferenciar acentos.")
    dados = banco
    termos = core.normalizar(busca).split()
    if termos:
        descricoes = dados["descricao_norm"].fillna("")
        filtro = pd.Series(True, index=dados.index)
        for termo in termos:
            filtro &= descricoes.str.contains(termo, regex=False)
        dados = dados[filtro]

    datas = pd.to_datetime(dados["data_cotacao"], errors="coerce")
    exportar = pd.DataFrame({
        "Descrição": dados["descricao"],
        "Valor unitário": dados["valor_unitario"],
        "Fornecedor": dados["fornecedor"],
        "Vendedor": dados["vendedor"],
        "Data da cotação": datas.dt.date,
        "Licitação": dados["origem"],
        "Código": dados["codigo"],
    }).assign(_ordem=dados["descricao_norm"], _data=datas)
    exportar = exportar.sort_values(["_ordem", "_data"], ascending=[True, False]).drop(columns=["_ordem", "_data"])

    st.caption(f"{fmt_int(len(exportar))} de {fmt_int(len(banco))} registros")
    vis = exportar.assign(**{
        "Valor unitário": exportar["Valor unitário"].map(brl),
        "Data da cotação": pd.to_datetime(exportar["Data da cotação"]).dt.strftime("%d/%m/%Y").fillna(""),
    })
    st.dataframe(vis, hide_index=True, column_config={"Descrição": st.column_config.TextColumn(width="large")})
    st.download_button("Exportar resultado (.xlsx)", core.df_para_xlsx(exportar, "Banco de preços"),
                       file_name="banco_de_precos.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.subheader("Importações")
    st.caption("Cada planilha cadastrada é uma importação. Se algo foi gravado errado, remova a importação "
               "inteira e cadastre a planilha de novo.")
    lotes = storage.listar_lotes(engine)
    lotes["importado_em"] = pd.to_datetime(lotes["importado_em"], errors="coerce")
    st.dataframe(
        pd.DataFrame({
            "Licitação": lotes["origem"],
            "Importada em": lotes["importado_em"].dt.strftime("%d/%m/%Y %H:%M"),
            "Preços": lotes["registros"],
        }),
        hide_index=True,
    )
    info = {r.lote: r for r in lotes.itertuples()}
    lote = st.selectbox(
        "Importação para remover", list(info), index=None, placeholder="Escolha uma importação",
        format_func=lambda l: f"{info[l].origem} ({fmt_int(info[l].registros)} preços, "
                              f"{info[l].importado_em:%d/%m/%Y %H:%M})",
    )
    if lote is not None:
        confirmar = st.checkbox(f"Confirmo que quero apagar os {fmt_int(info[lote].registros)} preços desta importação")
        if st.button("Remover importação", disabled=not confirmar):
            removidos = storage.remover_lote(engine, lote)
            carregar_banco.clear()
            avisar("consulta", f"Importação removida: {fmt_int(removidos)} preços apagados do banco.")
            st.rerun()


# ----------------------------------------------------------------------------
# Página
# ----------------------------------------------------------------------------
def main() -> None:
    logo = arquivo_logo()
    st.set_page_config(
        page_title="Orçamento de licitações | Atlântico Engenharia",
        page_icon=str(logo) if logo and logo.suffix != ".svg" else "📋",
        layout="wide",
    )
    st.markdown(CSS, unsafe_allow_html=True)

    engine, banco, erro = None, None, None
    try:
        engine = obter_engine()
        banco = carregar_banco(engine)
    except Exception as e:  # o app continua funcionando só com o SINAPI
        erro = e

    topo(engine, banco)
    if erro is not None:
        st.error(f"Não foi possível conectar ao banco de preços ({type(erro).__name__}). "
                 "O preenchimento com SINAPI continua funcionando. Confira a DATABASE_URL em Settings > Secrets.")

    aba_preencher, aba_cadastrar, aba_consultar = st.tabs(["Preencher planilha", "Cadastrar preços", "Consultar banco"])
    with aba_preencher:
        tela_preencher(banco)
    with aba_cadastrar:
        tela_cadastrar(engine)
    with aba_consultar:
        tela_consultar(engine, banco)


if __name__ == "__main__":
    main()
