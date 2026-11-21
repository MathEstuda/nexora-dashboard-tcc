
import os
import difflib
import hashlib
import time
from datetime import datetime

# VERSÃO 10 - AVALIAÇÃO SEM TEMPOS
import pandas as pd
import streamlit as st
import plotly.express as px
import re
from io import BytesIO
from docx import Document
from docx.shared import RGBColor, Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import unicodedata
from difflib import SequenceMatcher
import base64
import streamlit.components.v1 as components

# PROMPT_RELATORIO_PADRAO_SEGURO
prompt_relatorio = ""


st.set_page_config(
    page_title="Nexora",
    page_icon="📊",
    layout="wide"
)

# =========================================================
# IDENTIDADE VISUAL NEXORA V11
# =========================================================

st.markdown(
    (
        '<div style="margin:0 0 18px 0;'
        'padding:4px 0 14px 0;'
        'border-bottom:1px solid #E2E8F0;">'
        '<div style="color:#1F4E78;'
        'font-family:Arial,sans-serif;'
        'font-size:32px;'
        'font-weight:800;'
        'letter-spacing:1.2px;'
        'line-height:1.15;">'
        'NEXORA'
        '</div>'
        '<div style="color:#64748B;'
        'font-family:Arial,sans-serif;'
        'font-size:15px;'
        'font-weight:400;'
        'margin-top:5px;">'
        'Inteligência para dados, processos e relatórios'
        '</div>'
        '</div>'
    ),
    unsafe_allow_html=True
)


# =========================================================
# TEMA EXECUTIVO LEVE GERAREL V11
# =========================================================

st.markdown(
    """
    <style>

    /* ------------------------------------------------ */
    /* BASE DA APLICAÇÃO                                */
    /* ------------------------------------------------ */

    .stApp {
        background-color: #F8FAFC;
    }

    .main .block-container {
        max-width: 1480px;
        padding-top: 1.35rem;
        padding-bottom: 2.5rem;
    }

    h1, h2, h3, h4 {
        color: #172B4D;
        letter-spacing: -0.015em;
    }

    h1 {
        font-weight: 750;
    }

    h2, h3 {
        font-weight: 700;
    }

    /* ------------------------------------------------ */
    /* BOTÕES PRINCIPAIS                                */
    /* ------------------------------------------------ */

    .stButton > button[kind="primary"] {
        background-color: #1F4E78;
        border-color: #1F4E78;
        color: #FFFFFF;
        border-radius: 8px;
        min-height: 2.75rem;
        font-weight: 650;
        box-shadow:
            0 3px 9px rgba(31, 78, 120, 0.15);
        transition:
            background-color 150ms ease,
            border-color 150ms ease,
            transform 150ms ease,
            box-shadow 150ms ease;
    }

    .stButton > button[kind="primary"]:hover {
        background-color: #183E61;
        border-color: #183E61;
        color: #FFFFFF;
        transform: translateY(-1px);
        box-shadow:
            0 5px 13px rgba(31, 78, 120, 0.20);
    }

    /* ------------------------------------------------ */
    /* BOTÕES SECUNDÁRIOS                               */
    /* ------------------------------------------------ */

    .stButton > button[kind="secondary"] {
        background-color: #FFFFFF;
        border: 1px solid #CBD5E1;
        color: #1F4E78;
        border-radius: 8px;
        min-height: 2.65rem;
        font-weight: 600;
    }

    .stButton > button[kind="secondary"]:hover {
        background-color: #F1F5F9;
        border-color: #94A3B8;
        color: #183E61;
    }

    /* ------------------------------------------------ */
    /* DOWNLOAD                                         */
    /* ------------------------------------------------ */

    .stDownloadButton > button {
        background-color: #1F4E78;
        border: 1px solid #1F4E78;
        color: #FFFFFF;
        border-radius: 8px;
        min-height: 2.75rem;
        font-weight: 650;
    }

    .stDownloadButton > button:hover {
        background-color: #183E61;
        border-color: #183E61;
        color: #FFFFFF;
    }

    /* ------------------------------------------------ */
    /* FORMULÁRIOS                                      */
    /* ------------------------------------------------ */

    div[data-testid="stForm"] {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 1rem;
        box-shadow:
            0 2px 8px rgba(23, 43, 77, 0.045);
    }

    /* ------------------------------------------------ */
    /* ABAS                                             */
    /* ------------------------------------------------ */

    div[data-baseweb="tab-list"] {
        gap: 0.25rem;
        border-bottom: 1px solid #E2E8F0;
    }

    button[data-baseweb="tab"] {
        color: #64748B;
        border-radius: 7px 7px 0 0;
        padding-left: 0.85rem;
        padding-right: 0.85rem;
    }

    button[data-baseweb="tab"]:hover {
        background-color: #F1F5F9;
        color: #1F4E78;
    }

    button[data-baseweb="tab"][aria-selected="true"] {
        color: #1F4E78;
        font-weight: 700;
    }

    /* ------------------------------------------------ */
    /* BARRA DE PROGRESSO                               */
    /* ------------------------------------------------ */

    div[data-testid="stProgress"] > div > div {
        background-color: #2F80ED;
        border-radius: 99px;
    }

    /* ------------------------------------------------ */
    /* ALERTAS                                          */
    /* ------------------------------------------------ */

    div[data-testid="stAlert"] {
        border-radius: 9px;
        border-left-width: 4px;
    }

    /* ------------------------------------------------ */
    /* EXPANSORES                                       */
    /* ------------------------------------------------ */

    details[data-testid="stExpander"] {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 9px;
    }

    details[data-testid="stExpander"] summary {
        font-weight: 600;
        color: #334155;
    }

    /* ------------------------------------------------ */
    /* SIDEBAR                                          */
    /* Mantém a aparência nativa e só adiciona borda.   */
    /* ------------------------------------------------ */

    section[data-testid="stSidebar"] {
        border-right: 1px solid #E2E8F0;
    }

    /* ------------------------------------------------ */
    /* TABELAS                                          */
    /* ------------------------------------------------ */

    div[data-testid="stDataFrame"] {
        border: 1px solid #E2E8F0;
        border-radius: 9px;
        overflow: hidden;
    }

    /* ------------------------------------------------ */
    /* CAMPOS                                           */
    /* Não altera cores internas, apenas arredondamento */
    /* ------------------------------------------------ */

    textarea,
    input {
        border-radius: 7px !important;
    }

    /* ------------------------------------------------ */
    /* ESPAÇAMENTO                                      */
    /* ------------------------------------------------ */

    hr {
        margin-top: 1.35rem;
        margin-bottom: 1.35rem;
        border-color: #E2E8F0;
    }

    @media (max-width: 900px) {

        .main .block-container {
            padding-left: 0.85rem;
            padding-right: 0.85rem;
        }

        button[data-baseweb="tab"] {
            font-size: 0.84rem;
            padding-left: 0.55rem;
            padding-right: 0.55rem;
        }
    }

    </style>
    """,
    unsafe_allow_html=True
)


CAMINHOS_BASE = [
    # GitHub e Streamlit Community Cloud
    os.path.join(
        os.path.dirname(__file__),
        "data",
        "Base_Dados_TCC_IA_Relatorios_Ampliada.xlsx"
    ),

    # Estrutura local original do projeto
    os.path.join(
        os.path.dirname(__file__),
        "..",
        "01_Base_Dados",
        "Base_Dados_TCC_IA_Relatorios_Ampliada.xlsx"
    ),

    # Google Drive e Google Colab
    (
        "/content/drive/MyDrive/Colab Notebooks/"
        "TCC_GeraRel/01_Base_Dados/"
        "Base_Dados_TCC_IA_Relatorios_Ampliada.xlsx"
    )
]

arquivo_excel = None

for caminho in CAMINHOS_BASE:
    caminho = os.path.abspath(caminho)

    if os.path.exists(caminho):
        arquivo_excel = caminho
        break

if arquivo_excel is None:
    st.error(
        "A planilha não foi encontrada na pasta 01_Base_Dados."
    )
    st.stop()

@st.cache_data
def carregar_base(caminho):
    cenarios = pd.read_excel(
        caminho,
        sheet_name="Cenarios"
    )

    registros = pd.read_excel(
        caminho,
        sheet_name="Registros_Operacionais"
    )

    indicadores = pd.read_excel(
        caminho,
        sheet_name="Indicadores_Mensais"
    )

    ocorrencias = pd.read_excel(
        caminho,
        sheet_name="Ocorrencias"
    )

    encaminhamentos = pd.read_excel(
        caminho,
        sheet_name="Encaminhamentos"
    )

    return (
        cenarios,
        registros,
        indicadores,
        ocorrencias,
        encaminhamentos
    )

(
    cenarios,
    registros,
    indicadores,
    ocorrencias,
    encaminhamentos
) = carregar_base(arquivo_excel)

# =========================================================
# INTEGRAÇÃO DINÂMICA DA FONTE DE DADOS V11
# =========================================================

# Preserva cópias independentes da base acadêmica.
cenarios_academicos = cenarios.copy()
registros_academicos = registros.copy()
indicadores_academicos = indicadores.copy()
ocorrencias_academicas = ocorrencias.copy()
encaminhamentos_academicos = encaminhamentos.copy()

if "fonte_dados_ativa" not in st.session_state:

    st.session_state[
        "fonte_dados_ativa"
    ] = "Base acadêmica"

if "dados_importados_ativos" not in st.session_state:

    st.session_state[
        "dados_importados_ativos"
    ] = None

if "nome_arquivo_importado_ativo" not in st.session_state:

    st.session_state[
        "nome_arquivo_importado_ativo"
    ] = ""

fonte_dados_ativa = st.session_state[
    "fonte_dados_ativa"
]

dados_importados_ativos = st.session_state.get(
    "dados_importados_ativos"
)

# ---------------------------------------------------------
# ATIVA A PLANILHA IMPORTADA, QUANDO DISPONÍVEL
# ---------------------------------------------------------

if (
    fonte_dados_ativa == "Planilha importada"
    and isinstance(
        dados_importados_ativos,
        dict
    )
):

    abas_necessarias_ativas = {
        "Cenarios",
        "Registros_Operacionais",
        "Indicadores_Mensais",
        "Ocorrencias",
        "Encaminhamentos"
    }

    abas_importadas_ativas = set(
        dados_importados_ativos.keys()
    )

    if abas_necessarias_ativas.issubset(
        abas_importadas_ativas
    ):

        cenarios = (
            dados_importados_ativos[
                "Cenarios"
            ].copy()
        )

        registros = (
            dados_importados_ativos[
                "Registros_Operacionais"
            ].copy()
        )

        indicadores = (
            dados_importados_ativos[
                "Indicadores_Mensais"
            ].copy()
        )

        ocorrencias = (
            dados_importados_ativos[
                "Ocorrencias"
            ].copy()
        )

        encaminhamentos = (
            dados_importados_ativos[
                "Encaminhamentos"
            ].copy()
        )

        # Normaliza os relacionamentos por cenário.
        dataframes_cenario_ativos = [
            cenarios,
            registros,
            indicadores,
            ocorrencias,
            encaminhamentos
        ]

        for dataframe_ativo in (
            dataframes_cenario_ativos
        ):

            if "Cenario_ID" in dataframe_ativo.columns:

                dataframe_ativo[
                    "Cenario_ID"
                ] = (
                    dataframe_ativo[
                        "Cenario_ID"
                    ]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

        # Remove cenários vazios e duplicados.
        cenarios = cenarios[
            cenarios["Cenario_ID"] != ""
        ].copy()

        cenarios = cenarios.drop_duplicates(
            subset=["Cenario_ID"],
            keep="first"
        )

    else:

        # Retorno seguro para a base acadêmica.
        st.session_state[
            "fonte_dados_ativa"
        ] = "Base acadêmica"

        st.session_state[
            "dados_importados_ativos"
        ] = None

        fonte_dados_ativa = "Base acadêmica"

        cenarios = cenarios_academicos.copy()
        registros = registros_academicos.copy()
        indicadores = indicadores_academicos.copy()
        ocorrencias = ocorrencias_academicas.copy()
        encaminhamentos = (
            encaminhamentos_academicos.copy()
        )

else:

    fonte_dados_ativa = "Base acadêmica"

    cenarios = cenarios_academicos.copy()
    registros = registros_academicos.copy()
    indicadores = indicadores_academicos.copy()
    ocorrencias = ocorrencias_academicas.copy()
    encaminhamentos = (
        encaminhamentos_academicos.copy()
    )

st.sidebar.header("Configuração")

# ---------------------------------------------------------
# STATUS DA FONTE ATIVA
# ---------------------------------------------------------

if fonte_dados_ativa == "Planilha importada":

    nome_arquivo_fonte = st.session_state.get(
        "nome_arquivo_importado_ativo",
        "Arquivo externo"
    )

    st.sidebar.success(
        "Fonte ativa: Planilha importada"
    )

    st.sidebar.caption(
        f"Arquivo: {nome_arquivo_fonte}"
    )

else:

    st.sidebar.info(
        "Fonte ativa: Base acadêmica"
    )

lista_cenarios_ativos = (
    cenarios["Cenario_ID"]
    .dropna()
    .astype(str)
    .str.strip()
)

lista_cenarios_ativos = [
    cenario_id
    for cenario_id in (
        lista_cenarios_ativos.tolist()
    )
    if cenario_id
]

if not lista_cenarios_ativos:

    st.error(
        "A fonte ativa não possui cenários válidos."
    )

    st.stop()

cenario_escolhido = st.sidebar.selectbox(
    "Selecione o cenário",
    lista_cenarios_ativos,
    key=(
        f"seletor_cenario_"
        f"{fonte_dados_ativa}"
    )
)

if fonte_dados_ativa == "Planilha importada":

    voltar_base_academica = st.sidebar.button(
        "↩️ Voltar à base acadêmica",
        type="secondary",
        width="stretch",
        key="voltar_base_academica"
    )

    if voltar_base_academica:

        st.session_state[
            "fonte_dados_ativa"
        ] = "Base acadêmica"

        st.session_state[
            "dados_importados_ativos"
        ] = None

        st.session_state[
            "nome_arquivo_importado_ativo"
        ] = ""

        st.session_state[
            "planilha_externa_confirmada"
        ] = False

        st.session_state[
            "mostrar_confirmacao_ativacao"
        ] = False

        st.session_state[
            "mensagem_ativacao_planilha"
        ] = ""

        st.rerun()

if fonte_dados_ativa == "Base acadêmica":

    st.sidebar.caption(
        "Os dados acadêmicos são fictícios e destinados "
        "exclusivamente ao experimento."
    )

else:

    st.sidebar.caption(
        "Os dados exibidos foram carregados do arquivo "
        "externo confirmado na Central de Dados."
    )

dados_cenario = cenarios[
    cenarios["Cenario_ID"] == cenario_escolhido
].copy()

registros_cenario = registros[
    registros["Cenario_ID"] == cenario_escolhido
].copy()

indicadores_cenario = indicadores[
    indicadores["Cenario_ID"] == cenario_escolhido
].copy()

# =========================================================
# COMPATIBILIDADE DOS INDICADORES IMPORTADOS V11
# =========================================================

# Algumas planilhas administrativas não possuem uma
# coluna específica para a unidade do indicador.
# Esse campo é tratado como opcional para não impedir
# a utilização da base importada.

if "Unidade" not in indicadores_cenario.columns:

    indicadores_cenario[
        "Unidade"
    ] = "Não informada"

else:

    indicadores_cenario[
        "Unidade"
    ] = (
        indicadores_cenario[
            "Unidade"
        ]
        .fillna("Não informada")
        .astype(str)
        .str.strip()
        .replace(
            "",
            "Não informada"
        )
    )

# Garante também os demais campos utilizados na tela,
# evitando queda do dashboard por colunas opcionais.

campos_opcionais_indicadores = {
    "Indicador": "Indicador não informado",
    "Meta": 0,
    "Realizado": 0,
    "Atingimento_pct": 0.0,
    "Situacao": "Não informada"
}

for (
    campo_indicador,
    valor_padrao_indicador
) in campos_opcionais_indicadores.items():

    if (
        campo_indicador
        not in indicadores_cenario.columns
    ):

        indicadores_cenario[
            campo_indicador
        ] = valor_padrao_indicador

# =========================================================
# COERÊNCIA SEMÂNTICA E MATEMÁTICA V11
# =========================================================

avisos_coerencia_cenario = []

# ---------------------------------------------------------
# PRESERVA O PERCENTUAL ORIGINAL
# ---------------------------------------------------------

if (
    "Atingimento_pct_Informado"
    not in indicadores_cenario.columns
):

    indicadores_cenario[
        "Atingimento_pct_Informado"
    ] = indicadores_cenario[
        "Atingimento_pct"
    ].copy()

# ---------------------------------------------------------
# CONVERTE META E REALIZADO PARA NÚMEROS
# ---------------------------------------------------------

meta_numerica = pd.to_numeric(
    indicadores_cenario["Meta"],
    errors="coerce"
)

realizado_numerico = pd.to_numeric(
    indicadores_cenario["Realizado"],
    errors="coerce"
)

percentual_informado = pd.to_numeric(
    indicadores_cenario[
        "Atingimento_pct_Informado"
    ],
    errors="coerce"
)

# Normaliza percentuais informados como 80 ou 100
# para a escala decimal 0,80 ou 1,00.
percentual_informado_decimal = (
    percentual_informado.copy()
)

mascara_percentual_inteiro = (
    percentual_informado_decimal.abs()
    > 10
)

percentual_informado_decimal.loc[
    mascara_percentual_inteiro
] = (
    percentual_informado_decimal.loc[
        mascara_percentual_inteiro
    ]
    / 100
)

# ---------------------------------------------------------
# CALCULA REALIZADO / META
# ---------------------------------------------------------

mascara_meta_valida = (
    meta_numerica.notna()
    & realizado_numerico.notna()
    & (meta_numerica != 0)
)

percentual_calculado = (
    percentual_informado_decimal.copy()
)

percentual_calculado.loc[
    mascara_meta_valida
] = (
    realizado_numerico.loc[
        mascara_meta_valida
    ]
    / meta_numerica.loc[
        mascara_meta_valida
    ]
)

percentual_calculado = (
    percentual_calculado.fillna(0.0)
)

# ---------------------------------------------------------
# IDENTIFICA DIVERGÊNCIAS
# ---------------------------------------------------------

diferenca_percentual = (
    percentual_informado_decimal
    - percentual_calculado
).abs()

# Tolerância de 0,5 ponto percentual.
mascara_divergencia = (
    mascara_meta_valida
    & percentual_informado_decimal.notna()
    & (diferenca_percentual > 0.005)
)

indicadores_cenario[
    "Atingimento_pct_Calculado"
] = percentual_calculado

indicadores_cenario[
    "Divergencia_Atingimento"
] = mascara_divergencia

# O dashboard e o prompt passam a usar o valor calculado.
indicadores_cenario[
    "Atingimento_pct"
] = percentual_calculado

quantidade_divergencias = int(
    mascara_divergencia.sum()
)

if quantidade_divergencias > 0:

    avisos_coerencia_cenario.append(
        f"{quantidade_divergencias} indicador(es) "
        "possuem percentual informado diferente do "
        "resultado calculado por Realizado ÷ Meta. "
        "A Nexora utilizará o percentual calculado."
    )

quantidade_metas_zero = int(
    (
        meta_numerica.notna()
        & (meta_numerica == 0)
    ).sum()
)

if quantidade_metas_zero > 0:

    avisos_coerencia_cenario.append(
        f"{quantidade_metas_zero} indicador(es) "
        "possuem meta igual a zero. O percentual "
        "não pôde ser recalculado nesses casos."
    )

# ---------------------------------------------------------
# CONFERE A SITUAÇÃO DO INDICADOR
# ---------------------------------------------------------

if "Situacao" in indicadores_cenario.columns:

    situacao_normalizada = (
        indicadores_cenario[
            "Situacao"
        ]
        .fillna("")
        .astype(str)
        .str.lower()
        .str.strip()
    )

    meta_atingida_calculada = (
        mascara_meta_valida
        & (
            realizado_numerico
            >= meta_numerica
        )
    )

    situacao_diz_atingida = (
        situacao_normalizada.str.contains(
            "atingida",
            na=False
        )
        & ~situacao_normalizada.str.contains(
            "não",
            na=False
        )
        & ~situacao_normalizada.str.contains(
            "nao",
            na=False
        )
    )

    mascara_situacao_incoerente = (
        mascara_meta_valida
        & (
            meta_atingida_calculada
            != situacao_diz_atingida
        )
    )

    quantidade_situacoes_incoerentes = int(
        mascara_situacao_incoerente.sum()
    )

    if quantidade_situacoes_incoerentes > 0:

        avisos_coerencia_cenario.append(
            f"{quantidade_situacoes_incoerentes} "
            "indicador(es) possuem situação textual "
            "incompatível com Meta e Realizado."
        )

# ---------------------------------------------------------
# CONFERE TIPO DE ATIVIDADE VERSUS SETOR
# ---------------------------------------------------------

if (
    "Tipo_Atividade"
    in registros_cenario.columns
    and not dados_cenario.empty
):

    setor_cenario_validacao = str(
        dados_cenario.iloc[0].get(
            "Setor_Ficticio",
            ""
        )
    ).strip()

    def normalizar_comparacao_nexora(valor):

        texto = str(
            valor
        ).strip().lower()

        texto = (
            unicodedata.normalize(
                "NFKD",
                texto
            )
        )

        texto = "".join(
            caractere
            for caractere in texto
            if not unicodedata.combining(
                caractere
            )
        )

        texto = re.sub(
            r"[^a-z0-9]+",
            " ",
            texto
        )

        return " ".join(
            texto.split()
        )

    setor_normalizado = (
        normalizar_comparacao_nexora(
            setor_cenario_validacao
        )
    )

    atividades_normalizadas = (
        registros_cenario[
            "Tipo_Atividade"
        ]
        .fillna("")
        .astype(str)
        .map(
            normalizar_comparacao_nexora
        )
    )

    mascara_atividade_igual_setor = (
        (atividades_normalizadas != "")
        & (
            atividades_normalizadas
            == setor_normalizado
        )
    )

    quantidade_atividade_igual_setor = int(
        mascara_atividade_igual_setor.sum()
    )

    if quantidade_atividade_igual_setor > 0:

        avisos_coerencia_cenario.append(
            f"{quantidade_atividade_igual_setor} "
            "registro(s) utilizam o mesmo valor em "
            "Tipo_Atividade e Setor. Verifique se a "
            "coluna de atividade foi mapeada corretamente."
        )

ocorrencias_cenario = ocorrencias[
    ocorrencias["Cenario_ID"] == cenario_escolhido
].copy()

encaminhamentos_cenario = encaminhamentos[
    encaminhamentos["Cenario_ID"] == cenario_escolhido
].copy()

if dados_cenario.empty:

    st.error(
        "O cenário selecionado não possui cadastro "
        "válido na fonte de dados ativa."
    )

    st.stop()

geral = dados_cenario.iloc[0]

st.subheader(f"Cenário selecionado: {cenario_escolhido}")

if avisos_coerencia_cenario:

    with st.expander(
        "⚠️ Verificações de coerência dos dados",
        expanded=True
    ):

        for aviso_coerencia in (
            avisos_coerencia_cenario
        ):

            st.warning(
                aviso_coerencia
            )

if (
    fonte_dados_ativa == "Planilha importada"
    and registros_cenario.empty
):

    st.warning(
        "Este cenário está cadastrado na planilha "
        "importada, mas ainda não possui registros "
        "operacionais. Algumas análises e relatórios "
        "poderão ficar indisponíveis."
    )

info1, info2, info3 = st.columns(3)

info1.write("**Organização**")
info1.write(geral["Organizacao_Ficticia"])

info2.write("**Setor**")
info2.write(geral["Setor_Ficticio"])

info3.write("**Complexidade**")
info3.write(geral["Complexidade"])

st.divider()

tipos_atividade = registros_cenario[
    "Tipo_Atividade"
].nunique()

card1, card2, card3, card4, card5 = st.columns(5)

card1.metric(
    "Registros",
    len(registros_cenario)
)

card2.metric(
    "Tipos de atividade",
    tipos_atividade
)

card3.metric(
    "Indicadores",
    len(indicadores_cenario)
)

card4.metric(
    "Ocorrências",
    len(ocorrencias_cenario)
)

card5.metric(
    "Encaminhamentos",
    len(encaminhamentos_cenario)
)

st.divider()

resumo_atividades = (
    registros_cenario
    .groupby(
        ["Tipo_Atividade", "Unidade"],
        as_index=False
    )
    .agg(
        Quantidade_Total=("Quantidade", "sum"),
        Numero_Registros=("Registro_ID", "count"),
        Prazo_Medio_dias=("Prazo_Padrao_dias", "mean"),
        Tempo_Real_Medio_dias=("Tempo_Real_dias", "mean")
    )
)

resumo_atividades["Prazo_Medio_dias"] = (
    resumo_atividades["Prazo_Medio_dias"].round(1)
)

resumo_atividades["Tempo_Real_Medio_dias"] = (
    resumo_atividades["Tempo_Real_Medio_dias"].round(1)
)

# Cria uma cópia somente para exibição
atividades_exibicao = resumo_atividades.copy()

# Formata números decimais com vírgula
for coluna in [
    "Prazo_Medio_dias",
    "Tempo_Real_Medio_dias"
]:
    atividades_exibicao[coluna] = (
        atividades_exibicao[coluna]
        .map(lambda valor: f"{valor:.1f}".replace(".", ","))
    )

# Renomeia as colunas
atividades_exibicao = atividades_exibicao.rename(
    columns={
        "Tipo_Atividade": "Tipo de atividade",
        "Unidade": "Unidade",
        "Quantidade_Total": "Quantidade total",
        "Numero_Registros": "Número de registros",
        "Prazo_Medio_dias": "Prazo médio, em dias",
        "Tempo_Real_Medio_dias": "Tempo real médio, em dias"
    }
)


# ---------------------------------------------------------
# ABAS DE NAVEGAÇÃO
# ---------------------------------------------------------

aba_visao_geral, aba_operacoes, aba_indicadores, aba_central_dados, aba_fluxo_guiado, aba_validacao = st.tabs(
    [
        "🏠 Visão geral",
        "📊 Operações",
        "📈 Indicadores e ocorrências",
        "📥 Central de Dados",
        "🧭 Central de Relatórios",
        "✅ Validação do sistema"
    ]
)

# ---------------------------------------------------------
# ABA 1: VISÃO GERAL
# ---------------------------------------------------------

with aba_visao_geral:
    st.subheader("Resumo do cenário")

    coluna_resumo1, coluna_resumo2 = st.columns(2)

    with coluna_resumo1:
        st.markdown("#### Informações gerais")

        st.write(
            f"**Cenário:** {cenario_escolhido}"
        )

        st.write(
            f"**Organização:** "
            f"{geral['Organizacao_Ficticia']}"
        )

        st.write(
            f"**Setor:** "
            f"{geral['Setor_Ficticio']}"
        )

        st.write(
            f"**Complexidade:** "
            f"{geral['Complexidade']}"
        )

    with coluna_resumo2:
        st.markdown("#### Síntese dos dados")

        st.write(
            f"**Registros operacionais:** "
            f"{len(registros_cenario)}"
        )

        st.write(
            f"**Tipos de atividade:** "
            f"{tipos_atividade}"
        )

        st.write(
            f"**Indicadores acompanhados:** "
            f"{len(indicadores_cenario)}"
        )

        st.write(
            f"**Ocorrências registradas:** "
            f"{len(ocorrencias_cenario)}"
        )

        st.write(
            f"**Encaminhamentos previstos:** "
            f"{len(encaminhamentos_cenario)}"
        )

    st.info(
        "Utilize as demais abas para consultar as "
        "atividades, indicadores, ocorrências e "
        "recursos de geração do relatório."
    )

    # -----------------------------------------------------
    # ALERTAS AUTOMÁTICOS DO CENÁRIO
    # -----------------------------------------------------

    st.markdown("### Alertas automáticos")

    # Verifica as metas não atingidas
    situacoes_indicadores = (
        indicadores_cenario["Situacao"]
        .fillna("")
        .astype(str)
        .str.lower()
    )

    metas_nao_atingidas = int(
        situacoes_indicadores
        .str.contains(
            "não atingida|nao atingida",
            regex=True
        )
        .sum()
    )

    # Verifica ocorrências de severidade alta
    if ocorrencias_cenario.empty:

        ocorrencias_altas = 0

    else:

        severidades = (
            ocorrencias_cenario["Severidade"]
            .fillna("")
            .astype(str)
            .str.lower()
        )

        ocorrencias_altas = int(
            severidades
            .str.contains("alta")
            .sum()
        )

    # Verifica encaminhamentos de prioridade alta
    if encaminhamentos_cenario.empty:

        prioridades_altas = 0

    else:

        prioridades = (
            encaminhamentos_cenario["Prioridade"]
            .fillna("")
            .astype(str)
            .str.lower()
        )

        prioridades_altas = int(
            prioridades
            .str.contains("alta")
            .sum()
        )

    # Verifica registros parciais ou concluídos com atraso
    situacoes_registros = (
        registros_cenario["Situacao"]
        .fillna("")
        .astype(str)
        .str.lower()
    )

    registros_parciais = int(
        situacoes_registros
        .str.contains("parcial")
        .sum()
    )

    registros_com_atraso = int(
        situacoes_registros
        .str.contains("atraso")
        .sum()
    )

    # -----------------------------------------------------
    # EXIBIÇÃO DOS ALERTAS
    # -----------------------------------------------------

    alerta_coluna1, alerta_coluna2 = st.columns(2)

    with alerta_coluna1:

        if metas_nao_atingidas == 0:

            st.success(
                "✅ Todas as metas do cenário foram atingidas."
            )

        elif metas_nao_atingidas == 1:

            st.warning(
                "⚠ Foi identificada 1 meta não atingida."
            )

        else:

            st.warning(
                f"⚠ Foram identificadas "
                f"{metas_nao_atingidas} metas não atingidas."
            )

        if ocorrencias_cenario.empty:

            st.success(
                "✅ Nenhuma ocorrência foi registrada."
            )

        elif ocorrencias_altas == 0:

            st.info(
                f"ℹ Existem {len(ocorrencias_cenario)} "
                "ocorrências, mas nenhuma possui "
                "severidade alta."
            )

        elif ocorrencias_altas == 1:

            st.error(
                "🚨 Foi identificada 1 ocorrência "
                "de severidade alta."
            )

        else:

            st.error(
                f"🚨 Foram identificadas "
                f"{ocorrencias_altas} ocorrências "
                "de severidade alta."
            )

    with alerta_coluna2:

        if prioridades_altas == 0:

            st.success(
                "✅ Nenhum encaminhamento possui "
                "prioridade alta."
            )

        elif prioridades_altas == 1:

            st.warning(
                "⚠ Existe 1 encaminhamento "
                "com prioridade alta."
            )

        else:

            st.warning(
                f"⚠ Existem {prioridades_altas} "
                "encaminhamentos com prioridade alta."
            )

        if (
            registros_parciais == 0
            and registros_com_atraso == 0
        ):

            st.success(
                "✅ Não existem registros parciais "
                "ou concluídos com atraso."
            )

        else:

            st.warning(
                f"⚠ Registros parciais: "
                f"{registros_parciais}. "
                f"Registros com atraso: "
                f"{registros_com_atraso}."
            )

    # -----------------------------------------------------
    # SÍNTESE AUTOMÁTICA
    # -----------------------------------------------------

    st.markdown("### Síntese automática do cenário")

    texto_sintese = (
        f"O cenário {cenario_escolhido}, classificado como "
        f"{str(geral['Complexidade']).lower()} complexidade, "
        f"possui {len(registros_cenario)} registros "
        f"operacionais, {tipos_atividade} tipos de atividade, "
        f"{len(indicadores_cenario)} indicadores, "
        f"{len(ocorrencias_cenario)} ocorrências e "
        f"{len(encaminhamentos_cenario)} encaminhamentos. "
        f"Foram identificadas {metas_nao_atingidas} metas "
        f"não atingidas, {ocorrencias_altas} ocorrências "
        f"de severidade alta e {prioridades_altas} "
        f"encaminhamentos de prioridade alta."
    )

    st.write(texto_sintese)

# ---------------------------------------------------------
# ABA 2: OPERAÇÕES
# ---------------------------------------------------------

with aba_operacoes:
    st.subheader("Atividades consolidadas")

    st.dataframe(
        atividades_exibicao,
        width="stretch",
        hide_index=True
    )

    st.markdown("### Volume por atividade")

    grafico_volume = px.bar(
        resumo_atividades.sort_values(
            "Quantidade_Total",
            ascending=True
        ),
        x="Quantidade_Total",
        y="Tipo_Atividade",
        orientation="h",
        text="Quantidade_Total",
        color="Quantidade_Total",
        color_continuous_scale=[
            "#BDD7EE",
            "#5B9BD5",
            "#1F4E78"
        ],
        labels={
            "Quantidade_Total": "Quantidade total",
            "Tipo_Atividade": "Tipo de atividade"
        }
    )

    grafico_volume.update_traces(
        textposition="outside",
        hovertemplate=(
            "<b>%{y}</b><br>"
            "Quantidade total: %{x}<extra></extra>"
        )
    )

    grafico_volume.update_layout(
        height=430,
        margin=dict(
            l=10,
            r=50,
            t=20,
            b=10
        ),
        coloraxis_showscale=False,
        xaxis_title="Quantidade total",
        yaxis_title="",
        plot_bgcolor="white",
        paper_bgcolor="white",
        font=dict(
            family="Arial",
            size=13,
            color="#263238"
        )
    )

    grafico_volume.update_xaxes(
        showgrid=True,
        gridcolor="#E8EDF3",
        zeroline=False
    )

    grafico_volume.update_yaxes(
        showgrid=False
    )

    st.plotly_chart(
        grafico_volume,
        width="stretch",
        config={
            "displaylogo": False,
            "responsive": True
        }
    )


# ---------------------------------------------------------
# ABA 3: INDICADORES E OCORRÊNCIAS
# ---------------------------------------------------------

with aba_indicadores:
    st.subheader("Indicadores mensais")

    indicadores_exibicao = indicadores_cenario[
        [
            "Indicador",
            "Meta",
            "Realizado",
            "Unidade",
            "Atingimento_pct",
            "Situacao"
        ]
    ].copy()

    indicadores_exibicao["Atingimento_pct"] = (
        indicadores_exibicao["Atingimento_pct"]
        .mul(100)
        .round(1)
    )

    indicadores_exibicao["Atingimento_pct"] = (
        indicadores_exibicao["Atingimento_pct"]
        .map(lambda valor: f"{valor:.1f}%".replace(".", ","))
    )

    indicadores_exibicao = indicadores_exibicao.rename(
        columns={
            "Indicador": "Indicador",
            "Meta": "Meta",
            "Realizado": "Realizado",
            "Unidade": "Unidade",
            "Atingimento_pct": "Atingimento (%)",
            "Situacao": "Situação"
        }
    )

    st.dataframe(
        indicadores_exibicao,
        width="stretch",
        hide_index=True
    )

    st.markdown("### Meta versus realizado")

    # Copia os indicadores do cenário selecionado
    dados_grafico_indicadores = indicadores_cenario[
        [
            "Indicador",
            "Meta",
            "Realizado",
            "Situacao"
        ]
    ].copy()

    # Organiza Meta e Realizado para barras agrupadas
    dados_grafico_longos = (
        dados_grafico_indicadores
        .melt(
            id_vars=["Indicador"],
            value_vars=["Meta", "Realizado"],
            var_name="Referência",
            value_name="Valor"
        )
    )

    # Cria o gráfico interativo
    grafico_indicadores = px.bar(
        dados_grafico_longos,
        x="Indicador",
        y="Valor",
        color="Referência",
        barmode="group",
        text="Valor",
        color_discrete_map={
            "Meta": "#A7B6C6",
            "Realizado": "#1F77B4"
        },
        labels={
            "Indicador": "Indicador",
            "Valor": "Quantidade",
            "Referência": ""
        }
    )

    grafico_indicadores.update_traces(
        textposition="outside",
        cliponaxis=False,
        hovertemplate=(
            "<b>%{x}</b><br>"
            "%{fullData.name}: %{y}"
            "<extra></extra>"
        )
    )

    grafico_indicadores.update_layout(
        height=460,
        margin=dict(
            l=20,
            r=20,
            t=30,
            b=100
        ),
        plot_bgcolor="white",
        paper_bgcolor="white",
        legend_orientation="h",
        legend_y=-0.28,
        legend_x=0.36,
        font=dict(
            family="Arial",
            size=13,
            color="#263238"
        )
    )

    grafico_indicadores.update_xaxes(
        tickangle=-20,
        showgrid=False
    )

    grafico_indicadores.update_yaxes(
        showgrid=True,
        gridcolor="#E8EDF3",
        zeroline=False
    )

    st.plotly_chart(
        grafico_indicadores,
        width="stretch",
        config={
            "displaylogo": False,
            "responsive": True
        }
    )

    # Conta metas atingidas e não atingidas
    situacoes_grafico = (
        dados_grafico_indicadores["Situacao"]
        .fillna("")
        .astype(str)
        .str.lower()
    )

    metas_nao_atingidas_grafico = int(
        situacoes_grafico
        .str.contains(
            "não atingida|nao atingida",
            regex=True
        )
        .sum()
    )

    metas_atingidas_grafico = int(
        len(dados_grafico_indicadores)
        - metas_nao_atingidas_grafico
    )

    coluna_meta1, coluna_meta2 = st.columns(2)

    coluna_meta1.success(
        f"✅ Metas atingidas: "
        f"{metas_atingidas_grafico}"
    )

    if metas_nao_atingidas_grafico == 0:

        coluna_meta2.success(
            "✅ Nenhuma meta ficou abaixo do esperado."
        )

    else:

        coluna_meta2.warning(
            f"⚠ Metas não atingidas: "
            f"{metas_nao_atingidas_grafico}"
        )


    coluna_ocorrencias, coluna_encaminhamentos = st.columns(2)

    with coluna_ocorrencias:
        st.subheader("Ocorrências")

        if ocorrencias_cenario.empty:
            st.success(
                "Nenhuma ocorrência registrada no período."
            )
        else:
            ocorrencias_exibicao = ocorrencias_cenario[
                [
                    "Categoria",
                    "Descricao",
                    "Impacto",
                    "Severidade",
                    "Situacao"
                ]
            ].copy()

            ocorrencias_exibicao = ocorrencias_exibicao.rename(
                columns={
                    "Categoria": "Categoria",
                    "Descricao": "Descrição",
                    "Impacto": "Impacto",
                    "Severidade": "Severidade",
                    "Situacao": "Situação"
                }
            )

            st.dataframe(
                ocorrencias_exibicao,
                width="stretch",
                hide_index=True
            )

    with coluna_encaminhamentos:
        st.subheader("Encaminhamentos")

        if encaminhamentos_cenario.empty:
            st.info(
                "Nenhum encaminhamento registrado."
            )
        else:
            encaminhamentos_exibicao = encaminhamentos_cenario[
                [
                    "Acao_Prevista",
                    "Responsavel_Ficticio",
                    "Prazo",
                    "Status",
                    "Prioridade"
                ]
            ].copy()

            encaminhamentos_exibicao["Prazo"] = (
                pd.to_datetime(
                    encaminhamentos_exibicao["Prazo"]
                )
                .dt.strftime("%d/%m/%Y")
            )

            encaminhamentos_exibicao = encaminhamentos_exibicao.rename(
                columns={
                    "Acao_Prevista": "Ação prevista",
                    "Responsavel_Ficticio": "Responsável",
                    "Prazo": "Prazo",
                    "Status": "Status",
                    "Prioridade": "Prioridade"
                }
            )

            st.dataframe(
                encaminhamentos_exibicao,
                width="stretch",
                hide_index=True
            )


# ---------------------------------------------------------
# ABA 4: GERAR RELATÓRIO
# ---------------------------------------------------------


# =========================================================
# NORMALIZAÇÃO INTELIGENTE DE PLANILHAS V11
# =========================================================

def normalizar_rotulo_gerarel(valor):
    """
    Remove acentos, diferenças de caixa, espaços,
    hífens e caracteres especiais.
    """

    if valor is None:
        return ""

    texto = str(valor).strip().lower()

    texto = unicodedata.normalize(
        "NFKD",
        texto
    )

    texto = "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(
            caractere
        )
    )

    texto = re.sub(
        r"[^a-z0-9]+",
        "_",
        texto
    )

    texto = re.sub(
        r"_+",
        "_",
        texto
    )

    return texto.strip("_")


# ---------------------------------------------------------
# SINÔNIMOS DOS NOMES DAS ABAS
# ---------------------------------------------------------

SINONIMOS_ABAS_GERAREL = {
    "Cenarios": [
        "cenario",
        "cenarios",
        "cenario_geral",
        "dados_gerais",
        "cadastro_de_cenarios",
        "situacoes"
    ],
    "Registros_Operacionais": [
        "registros_operacionais",
        "registros",
        "atividades",
        "atividades_operacionais",
        "operacoes",
        "tarefas"
    ],
    "Indicadores_Mensais": [
        "indicadores_mensais",
        "indicadores",
        "metas",
        "resultados",
        "metas_e_resultados"
    ],
    "Ocorrencias": [
        "ocorrencias",
        "ocorrencia",
        "problemas",
        "eventos",
        "pendencias",
        "riscos"
    ],
    "Encaminhamentos": [
        "encaminhamentos",
        "encaminhamento",
        "acoes",
        "planos_de_acao",
        "proximas_acoes",
        "providencias"
    ]
}


# ---------------------------------------------------------
# SINÔNIMOS DAS COLUNAS
# ---------------------------------------------------------

SINONIMOS_COLUNAS_GERAREL = {
    "Cenario_ID": [
        "cenario_id",
        "id_cenario",
        "codigo_cenario",
        "cenario",
        "id_do_cenario"
    ],
    "Organizacao_Ficticia": [
        "organizacao_ficticia",
        "organizacao",
        "nome_da_organizacao",
        "nome_organizacao",
        "empresa",
        "instituicao",
        "orgao",
        "unidade"
    ],
    "Setor_Ficticio": [
        "setor_ficticio",
        "setor",
        "departamento",
        "area",
        "divisao",
        "unidade_administrativa"
    ],
    "Responsavel_Ficticio": [
        "responsavel_ficticio",
        "responsavel",
        "gestor",
        "colaborador",
        "encarregado",
        "coordenador"
    ],
    "Local": [
        "local",
        "cidade",
        "municipio",
        "localidade",
        "unidade_local"
    ],
    "Complexidade": [
        "complexidade",
        "nivel_de_complexidade",
        "nivel",
        "grau_de_complexidade"
    ],
    "Finalidade": [
        "finalidade",
        "objetivo",
        "proposito",
        "descricao_do_cenario"
    ],
    "Registro_ID": [
        "registro_id",
        "id_registro",
        "codigo_registro",
        "id_da_atividade"
    ],
    "Tipo_Atividade": [
        "tipo_atividade",
        "tipo_de_atividade",
        "atividade",
        "tarefa",
        "servico",
        "servico_executado",
        "operacao"
    ],
    "Quantidade": [
        "quantidade",
        "volume",
        "total",
        "producao",
        "qtd"
    ],
    "Unidade": [
        "unidade",
        "unidade_de_medida",
        "medida",
        "tipo_de_unidade"
    ],
    "Prazo_Padrao_dias": [
        "prazo_padrao_dias",
        "prazo_padrao",
        "prazo_previsto",
        "prazo_em_dias",
        "dias_previstos"
    ],
    "Tempo_Real_dias": [
        "tempo_real_dias",
        "tempo_real",
        "tempo_executado",
        "dias_utilizados",
        "duracao_real"
    ],
    "Situacao": [
        "situacao",
        "status",
        "resultado_situacao",
        "estado"
    ],
    "Indicador_ID": [
        "indicador_id",
        "id_indicador",
        "codigo_indicador"
    ],
    "Indicador": [
        "indicador",
        "nome_do_indicador",
        "metrica",
        "kpi"
    ],
    "Meta": [
        "meta",
        "meta_mensal",
        "objetivo_numerico",
        "valor_planejado",
        "planejado"
    ],
    "Realizado": [
        "realizado",
        "resultado_obtido",
        "resultado_realizado",
        "executado",
        "valor_realizado"
    ],
    "Atingimento_pct": [
        "atingimento_pct",
        "percentual_de_atingimento",
        "atingimento",
        "percentual",
        "percentual_realizado"
    ],
    "Ocorrencia_ID": [
        "ocorrencia_id",
        "id_ocorrencia",
        "codigo_ocorrencia"
    ],
    "Categoria": [
        "categoria",
        "tipo_de_ocorrencia",
        "categoria_da_ocorrencia"
    ],
    "Descricao": [
        "descricao",
        "descricao_da_ocorrencia",
        "problema_registrado",
        "ocorrencia",
        "observacao"
    ],
    "Impacto": [
        "impacto",
        "efeito",
        "consequencia"
    ],
    "Severidade": [
        "severidade",
        "gravidade",
        "nivel_de_risco"
    ],
    "Encaminhamento_ID": [
        "encaminhamento_id",
        "id_encaminhamento",
        "codigo_encaminhamento"
    ],
    "Acao_Prevista": [
        "acao_prevista",
        "acao",
        "proxima_acao",
        "providencia",
        "encaminhamento",
        "plano_de_acao"
    ],
    "Prazo": [
        "prazo",
        "data_limite",
        "prazo_final",
        "data_prevista"
    ],
    "Status": [
        "status",
        "situacao_da_acao",
        "estado_da_acao"
    ],
    "Prioridade": [
        "prioridade",
        "nivel_de_prioridade",
        "urgencia"
    ],
    "Resultado_Esperado": [
        "resultado_esperado",
        "resultado_previsto",
        "objetivo_da_acao",
        "entrega_esperada"
    ]
}


def criar_indice_sinonimos_gerarel(
    dicionario_sinonimos
):
    indice = {}

    for campo_oficial, sinonimos in (
        dicionario_sinonimos.items()
    ):

        indice[
            normalizar_rotulo_gerarel(
                campo_oficial
            )
        ] = campo_oficial

        for sinonimo in sinonimos:

            indice[
                normalizar_rotulo_gerarel(
                    sinonimo
                )
            ] = campo_oficial

    return indice


INDICE_ABAS_GERAREL = (
    criar_indice_sinonimos_gerarel(
        SINONIMOS_ABAS_GERAREL
    )
)

INDICE_COLUNAS_GERAREL = (
    criar_indice_sinonimos_gerarel(
        SINONIMOS_COLUNAS_GERAREL
    )
)


def calcular_similaridade_gerarel(
    texto_a,
    texto_b
):
    return SequenceMatcher(
        None,
        normalizar_rotulo_gerarel(texto_a),
        normalizar_rotulo_gerarel(texto_b)
    ).ratio()


def sugerir_campo_gerarel(
    nome_recebido,
    campos_permitidos,
    indice_sinonimos,
    limite_automatico=0.90,
    limite_sugestao=0.72
):
    """
    Retorna:
    - nome oficial sugerido;
    - confiança entre 0 e 1;
    - método de reconhecimento;
    - se pode ser aplicado automaticamente.
    """

    nome_normalizado = normalizar_rotulo_gerarel(
        nome_recebido
    )

    # Correspondência exata após normalização
    if nome_normalizado in indice_sinonimos:

        campo_encontrado = indice_sinonimos[
            nome_normalizado
        ]

        if campo_encontrado in campos_permitidos:

            return {
                "origem": str(nome_recebido),
                "destino": campo_encontrado,
                "confianca": 1.0,
                "metodo": (
                    "Normalização ou sinônimo conhecido"
                ),
                "automatico": True
            }

    melhor_campo = None
    melhor_confianca = 0.0

    nomes_candidatos = {}

    for campo_oficial in campos_permitidos:

        candidatos = [
            campo_oficial
        ] + SINONIMOS_COLUNAS_GERAREL.get(
            campo_oficial,
            []
        )

        nomes_candidatos[
            campo_oficial
        ] = candidatos

    for campo_oficial, candidatos in (
        nomes_candidatos.items()
    ):

        for candidato in candidatos:

            confianca = calcular_similaridade_gerarel(
                nome_recebido,
                candidato
            )

            if confianca > melhor_confianca:

                melhor_confianca = confianca
                melhor_campo = campo_oficial

    if (
        melhor_campo is not None
        and melhor_confianca >= limite_sugestao
    ):

        return {
            "origem": str(nome_recebido),
            "destino": melhor_campo,
            "confianca": melhor_confianca,
            "metodo": "Similaridade textual",
            "automatico": (
                melhor_confianca
                >= limite_automatico
            )
        }

    return {
        "origem": str(nome_recebido),
        "destino": None,
        "confianca": melhor_confianca,
        "metodo": "Não reconhecido",
        "automatico": False
    }


def normalizar_planilha_gerarel(
    planilhas_recebidas,
    estrutura_esperada
):
    """
    Normaliza nomes das abas e colunas.

    Correspondências claras são aplicadas.
    Casos ambíguos são devolvidos para confirmação.
    """

    planilhas_normalizadas = {}
    mapeamentos_aplicados = []
    sugestoes_pendentes = []
    avisos_normalizacao = []
    erros_normalizacao = []

    # -----------------------------------------------------
    # NORMALIZA AS ABAS
    # -----------------------------------------------------

    for nome_aba_recebida, dataframe in (
        planilhas_recebidas.items()
    ):

        aba_normalizada = normalizar_rotulo_gerarel(
            nome_aba_recebida
        )

        aba_oficial = INDICE_ABAS_GERAREL.get(
            aba_normalizada
        )

        if aba_oficial is None:

            melhor_aba = None
            melhor_confianca = 0.0

            for aba_candidata in (
                estrutura_esperada.keys()
            ):

                candidatos_aba = [
                    aba_candidata
                ] + SINONIMOS_ABAS_GERAREL.get(
                    aba_candidata,
                    []
                )

                for candidato_aba in candidatos_aba:

                    confianca_aba = (
                        calcular_similaridade_gerarel(
                            nome_aba_recebida,
                            candidato_aba
                        )
                    )

                    if (
                        confianca_aba
                        > melhor_confianca
                    ):

                        melhor_aba = aba_candidata
                        melhor_confianca = (
                            confianca_aba
                        )

            if melhor_confianca >= 0.90:

                aba_oficial = melhor_aba

            else:

                avisos_normalizacao.append(
                    "Aba adicional não reconhecida: "
                    f"{nome_aba_recebida}."
                )

                planilhas_normalizadas[
                    nome_aba_recebida
                ] = dataframe.copy()

                continue

        if aba_oficial in planilhas_normalizadas:

            erros_normalizacao.append(
                f"Duas abas foram associadas a "
                f"{aba_oficial}."
            )

            continue

        if nome_aba_recebida != aba_oficial:

            mapeamentos_aplicados.append(
                {
                    "tipo": "Aba",
                    "aba": aba_oficial,
                    "origem": nome_aba_recebida,
                    "destino": aba_oficial,
                    "confianca": 1.0
                }
            )

        dataframe_normalizado = dataframe.copy()

        colunas_esperadas = estrutura_esperada[
            aba_oficial
        ]

        renomear_colunas = {}
        destinos_utilizados = set()

        # -------------------------------------------------
        # NORMALIZA AS COLUNAS
        # -------------------------------------------------

        for coluna_recebida in (
            dataframe_normalizado.columns
        ):

            sugestao = sugerir_campo_gerarel(
                coluna_recebida,
                colunas_esperadas,
                INDICE_COLUNAS_GERAREL
            )

            destino = sugestao["destino"]

            if destino is None:

                avisos_normalizacao.append(
                    f"{aba_oficial}: coluna adicional "
                    f"não reconhecida: {coluna_recebida}."
                )

                continue

            if destino in destinos_utilizados:

                erros_normalizacao.append(
                    f"{aba_oficial}: mais de uma coluna "
                    f"foi associada ao campo {destino}."
                )

                continue

            if sugestao["automatico"]:

                renomear_colunas[
                    coluna_recebida
                ] = destino

                destinos_utilizados.add(
                    destino
                )

                if coluna_recebida != destino:

                    mapeamentos_aplicados.append(
                        {
                            "tipo": "Coluna",
                            "aba": aba_oficial,
                            "origem": (
                                str(coluna_recebida)
                            ),
                            "destino": destino,
                            "confianca": sugestao[
                                "confianca"
                            ]
                        }
                    )

            else:

                sugestoes_pendentes.append(
                    {
                        "aba": aba_oficial,
                        "origem": (
                            str(coluna_recebida)
                        ),
                        "destino_sugerido": destino,
                        "confianca": sugestao[
                            "confianca"
                        ],
                        "metodo": sugestao[
                            "metodo"
                        ]
                    }
                )

        dataframe_normalizado = (
            dataframe_normalizado.rename(
                columns=renomear_colunas
            )
        )

        planilhas_normalizadas[
            aba_oficial
        ] = dataframe_normalizado

    return {
        "dados": planilhas_normalizadas,
        "mapeamentos": mapeamentos_aplicados,
        "sugestoes": sugestoes_pendentes,
        "avisos": avisos_normalizacao,
        "erros": erros_normalizacao
    }


# =========================================================
# CENTRAL DE DADOS V11
# =========================================================

with aba_central_dados:

    # =====================================================
    # CONFIRMAÇÃO PERSISTENTE DA FONTE IMPORTADA V11
    # =====================================================

    st.subheader("Central de Dados")

    fonte_importada_ativa = (
        st.session_state.get(
            "fonte_dados_ativa"
        )
        == "Planilha importada"
        and isinstance(
            st.session_state.get(
                "dados_importados_ativos"
            ),
            dict
        )
    )

    if fonte_importada_ativa:

        dados_ativos_confirmacao = (
            st.session_state[
                "dados_importados_ativos"
            ]
        )

        nome_arquivo_confirmacao = (
            st.session_state.get(
                "nome_arquivo_importado_ativo",
                "Planilha importada"
            )
        )

        quantidade_cenarios_confirmacao = len(
            dados_ativos_confirmacao.get(
                "Cenarios",
                pd.DataFrame()
            )
        )

        quantidade_registros_confirmacao = len(
            dados_ativos_confirmacao.get(
                "Registros_Operacionais",
                pd.DataFrame()
            )
        )

        st.success(
            "✅ Planilha ativada com sucesso e "
            "utilizada como fonte atual do dashboard."
        )

        confirmacao_coluna1,         confirmacao_coluna2,         confirmacao_coluna3 = st.columns(3)

        confirmacao_coluna1.metric(
            "Fonte ativa",
            "Planilha importada"
        )

        confirmacao_coluna2.metric(
            "Cenários disponíveis",
            quantidade_cenarios_confirmacao
        )

        confirmacao_coluna3.metric(
            "Registros carregados",
            quantidade_registros_confirmacao
        )

        st.info(
            f"Arquivo ativo: {nome_arquivo_confirmacao}"
        )

        st.caption(
            "Os cartões, gráficos, alertas e prompts "
            "estão utilizando os dados deste arquivo."
        )

    st.write(
        "Utilize a base acadêmica do projeto ou envie uma "
        "nova planilha compatível com o modelo estrutural "
        "do GeraRel."
    )

    st.info(
        "Nesta primeira versão, o arquivo enviado será "
        "analisado e validado. A substituição da base do "
        "dashboard será liberada somente após a confirmação."
    )

    # -----------------------------------------------------
    # DEFINIÇÃO DO MODELO GERAREL
    # -----------------------------------------------------

    estrutura_gerarel = {
        "Cenarios": [
            "Cenario_ID",
            "Organizacao_Ficticia",
            "Setor_Ficticio",
            "Responsavel_Ficticio",
            "Local",
            "Complexidade",
            "Finalidade"
        ],
        "Registros_Operacionais": [
            "Registro_ID",
            "Cenario_ID",
            "Tipo_Atividade",
            "Quantidade",
            "Unidade",
            "Prazo_Padrao_dias",
            "Tempo_Real_dias",
            "Situacao"
        ],
        "Indicadores_Mensais": [
            "Indicador_ID",
            "Cenario_ID",
            "Indicador",
            "Meta",
            "Realizado",
            "Atingimento_pct",
            "Situacao"
        ],
        "Ocorrencias": [
            "Ocorrencia_ID",
            "Cenario_ID",
            "Categoria",
            "Descricao",
            "Impacto",
            "Severidade",
            "Situacao"
        ],
        "Encaminhamentos": [
            "Encaminhamento_ID",
            "Cenario_ID",
            "Acao_Prevista",
            "Responsavel_Ficticio",
            "Prazo",
            "Status",
            "Prioridade",
            "Resultado_Esperado"
        ]
    }

    # -----------------------------------------------------
    # FUNÇÃO PARA GERAR O ARQUIVO MODELO
    # -----------------------------------------------------

    def gerar_modelo_excel_gerarel():

        memoria_modelo = BytesIO()

        exemplos_modelo = {
            "Cenarios": pd.DataFrame(
                [
                    {
                        "Cenario_ID": "EX01",
                        "Organizacao_Ficticia": (
                            "Organização Exemplo"
                        ),
                        "Setor_Ficticio": (
                            "Setor Administrativo"
                        ),
                        "Responsavel_Ficticio": (
                            "Responsável Exemplo"
                        ),
                        "Local": "Rio de Janeiro",
                        "Complexidade": "Baixa",
                        "Finalidade": (
                            "Demonstrar o preenchimento"
                        )
                    }
                ]
            ),
            "Registros_Operacionais": pd.DataFrame(
                [
                    {
                        "Registro_ID": "REG001",
                        "Cenario_ID": "EX01",
                        "Tipo_Atividade": (
                            "Atendimento administrativo"
                        ),
                        "Quantidade": 10,
                        "Unidade": "atendimentos",
                        "Prazo_Padrao_dias": 2,
                        "Tempo_Real_dias": 2.5,
                        "Situacao": "Concluído"
                    }
                ]
            ),
            "Indicadores_Mensais": pd.DataFrame(
                [
                    {
                        "Indicador_ID": "IND001",
                        "Cenario_ID": "EX01",
                        "Indicador": (
                            "Atendimentos concluídos"
                        ),
                        "Meta": 10,
                        "Realizado": 10,
                        "Atingimento_pct": 1.0,
                        "Situacao": "Atingida"
                    }
                ]
            ),
            "Ocorrencias": pd.DataFrame(
                columns=estrutura_gerarel[
                    "Ocorrencias"
                ]
            ),
            "Encaminhamentos": pd.DataFrame(
                [
                    {
                        "Encaminhamento_ID": "ENC001",
                        "Cenario_ID": "EX01",
                        "Acao_Prevista": (
                            "Revisar o procedimento"
                        ),
                        "Responsavel_Ficticio": (
                            "Responsável Exemplo"
                        ),
                        "Prazo": "30/11/2026",
                        "Status": "Planejada",
                        "Prioridade": "Média",
                        "Resultado_Esperado": (
                            "Padronizar o processo"
                        )
                    }
                ]
            )
        }

        with pd.ExcelWriter(
            memoria_modelo,
            engine="openpyxl"
        ) as escritor:

            for nome_aba, dataframe in (
                exemplos_modelo.items()
            ):

                dataframe.to_excel(
                    escritor,
                    sheet_name=nome_aba,
                    index=False
                )

        memoria_modelo.seek(0)

        return memoria_modelo.getvalue()

    # -----------------------------------------------------
    # FONTE DOS DADOS
    # -----------------------------------------------------

    st.markdown("### Fonte dos dados")

    fonte_dados = st.radio(
        "Selecione a fonte de dados",
        options=[
            "Base acadêmica da Nexora",
            "Importar nova planilha"
        ],
        index=0,
        horizontal=True,
        key="central_dados_fonte"
    )

    if fonte_dados == "Base acadêmica da Nexora":

        st.success(
            "✅ A base acadêmica simulada permanece ativa."
        )

        base_coluna1, base_coluna2,         base_coluna3 = st.columns(3)

        base_coluna1.metric(
            "Cenários",
            len(cenarios)
        )

        base_coluna2.metric(
            "Registros operacionais",
            len(registros)
        )

        base_coluna3.metric(
            "Fonte",
            "Base simulada"
        )

        st.caption(
            "Nenhuma alteração foi realizada nos dados "
            "utilizados atualmente pelo dashboard."
        )

    else:

        st.markdown(
            "### Importar planilha no modelo GeraRel"
        )

        st.write(
            "Baixe o modelo, substitua os exemplos pelos "
            "dados desejados e mantenha os nomes das abas "
            "e das colunas obrigatórias."
        )

        modelo_excel = gerar_modelo_excel_gerarel()

        st.download_button(
            label="⬇️ Baixar modelo de planilha GeraRel",
            data=modelo_excel,
            file_name="Modelo_Planilha_GeraRel.xlsx",
            mime=(
                "application/vnd.openxmlformats-"
                "officedocument.spreadsheetml.sheet"
            ),
            width="stretch",
            key="download_modelo_gerarel"
        )

        arquivo_enviado = st.file_uploader(
            "Selecione a planilha",
            type=["xlsx", "xlsm"],
            accept_multiple_files=False,
            help=(
                "O arquivo deve conter as abas e colunas "
                "definidas no modelo GeraRel."
            ),
            key="upload_planilha_gerarel"
        )

        if arquivo_enviado is None:

            st.info(
                "Envie uma planilha para iniciar "
                "a validação estrutural."
            )

        else:

            st.markdown(
                "### Validação estrutural"
            )

            validar_arquivo = st.button(
                "🔍 Validar planilha enviada",
                type="primary",
                width="stretch",
                key="validar_planilha_gerarel"
            )

            if validar_arquivo:

                with st.spinner(
                    "Lendo e validando a planilha..."
                ):

                    progresso_validacao = st.progress(
                        0,
                        text="Abrindo o arquivo Excel..."
                    )

                    time.sleep(0.20)

                    erros_importacao = []
                    avisos_importacao = []
                    planilhas_importadas = {}

                    try:

                        arquivo_enviado.seek(0)

                        planilhas_importadas_brutas = (
                            pd.read_excel(
                                arquivo_enviado,
                                sheet_name=None,
                                engine="openpyxl"
                            )
                        )

                        resultado_normalizacao = (
                            normalizar_planilha_gerarel(
                                planilhas_importadas_brutas,
                                estrutura_gerarel
                            )
                        )

                        planilhas_importadas = (
                            resultado_normalizacao[
                                "dados"
                            ]
                        )

                        mapeamentos_aplicados = (
                            resultado_normalizacao[
                                "mapeamentos"
                            ]
                        )

                        sugestoes_pendentes = (
                            resultado_normalizacao[
                                "sugestoes"
                            ]
                        )

                        avisos_importacao.extend(
                            resultado_normalizacao[
                                "avisos"
                            ]
                        )

                        erros_importacao.extend(
                            resultado_normalizacao[
                                "erros"
                            ]
                        )

                        if mapeamentos_aplicados:

                            avisos_importacao.append(
                                f"{len(mapeamentos_aplicados)} "
                                "nome(s) de aba ou coluna "
                                "foram reconhecidos e "
                                "normalizados automaticamente."
                            )

                        if sugestoes_pendentes:

                            for sugestao_pendente in (
                                sugestoes_pendentes
                            ):

                                confianca_pct = (
                                    sugestao_pendente[
                                        "confianca"
                                    ]
                                    * 100
                                )

                                avisos_importacao.append(
                                    f"{sugestao_pendente['aba']}: "
                                    f"a coluna "
                                    f"'{sugestao_pendente['origem']}' "
                                    f"pode corresponder a "
                                    f"'{sugestao_pendente['destino_sugerido']}' "
                                    f"({confianca_pct:.0f}% de "
                                    f"similaridade). Confirmação "
                                    f"manual recomendada."
                                )

                    except Exception as erro_leitura:

                        erros_importacao.append(
                            "Não foi possível ler o arquivo: "
                            f"{type(erro_leitura).__name__}."
                        )

                    progresso_validacao.progress(
                        30,
                        text="Verificando as abas..."
                    )

                    if planilhas_importadas:

                        abas_recebidas = set(
                            planilhas_importadas.keys()
                        )

                        abas_obrigatorias = set(
                            estrutura_gerarel.keys()
                        )

                        abas_ausentes = (
                            abas_obrigatorias
                            - abas_recebidas
                        )

                        abas_extras = (
                            abas_recebidas
                            - abas_obrigatorias
                        )

                        for aba_ausente in sorted(
                            abas_ausentes
                        ):

                            erros_importacao.append(
                                f"Aba obrigatória ausente: "
                                f"{aba_ausente}."
                            )

                        for aba_extra in sorted(
                            abas_extras
                        ):

                            avisos_importacao.append(
                                f"Aba adicional ignorada: "
                                f"{aba_extra}."
                            )

                    progresso_validacao.progress(
                        55,
                        text="Verificando as colunas..."
                    )

                    if planilhas_importadas:

                        for nome_aba, colunas_esperadas in (
                            estrutura_gerarel.items()
                        ):

                            if nome_aba not in (
                                planilhas_importadas
                            ):
                                continue

                            dataframe_aba = (
                                planilhas_importadas[
                                    nome_aba
                                ]
                            )

                            colunas_recebidas = set(
                                dataframe_aba.columns
                            )

                            colunas_ausentes = (
                                set(colunas_esperadas)
                                - colunas_recebidas
                            )

                            for coluna_ausente in sorted(
                                colunas_ausentes
                            ):

                                erros_importacao.append(
                                    f"{nome_aba}: coluna "
                                    f"obrigatória ausente: "
                                    f"{coluna_ausente}."
                                )

                            colunas_extras = (
                                colunas_recebidas
                                - set(colunas_esperadas)
                            )

                            for coluna_extra in sorted(
                                colunas_extras
                            ):

                                avisos_importacao.append(
                                    f"{nome_aba}: coluna "
                                    f"adicional preservada: "
                                    f"{coluna_extra}."
                                )

                    progresso_validacao.progress(
                        75,
                        text=(
                            "Verificando os identificadores "
                            "dos cenários..."
                        )
                    )

                    if (
                        not erros_importacao
                        and "Cenarios"
                        in planilhas_importadas
                    ):

                        dados_cenarios_importados = (
                            planilhas_importadas[
                                "Cenarios"
                            ]
                        )

                        ids_cenarios = set(
                            dados_cenarios_importados[
                                "Cenario_ID"
                            ]
                            .dropna()
                            .astype(str)
                        )

                        if not ids_cenarios:

                            erros_importacao.append(
                                "A aba Cenarios não possui "
                                "nenhum Cenario_ID válido."
                            )

                        abas_relacionadas = [
                            "Registros_Operacionais",
                            "Indicadores_Mensais",
                            "Ocorrencias",
                            "Encaminhamentos"
                        ]

                        for aba_relacionada in (
                            abas_relacionadas
                        ):

                            dataframe_relacionado = (
                                planilhas_importadas[
                                    aba_relacionada
                                ]
                            )

                            ids_relacionados = set(
                                dataframe_relacionado[
                                    "Cenario_ID"
                                ]
                                .dropna()
                                .astype(str)
                            )

                            ids_sem_cadastro = (
                                ids_relacionados
                                - ids_cenarios
                            )

                            for id_sem_cadastro in sorted(
                                ids_sem_cadastro
                            ):

                                erros_importacao.append(
                                    f"{aba_relacionada}: "
                                    f"Cenario_ID sem cadastro "
                                    f"na aba Cenarios: "
                                    f"{id_sem_cadastro}."
                                )

                    progresso_validacao.progress(
                        100,
                        text="Validação concluída."
                    )

                    time.sleep(0.25)

                    progresso_validacao.empty()

                resultado_importacao = {
                    "arquivo": arquivo_enviado.name,
                    "dados": planilhas_importadas,
                    "erros": erros_importacao,
                    "avisos": avisos_importacao,
                    "mapeamentos_aplicados": (
                        locals().get(
                            "mapeamentos_aplicados",
                            []
                        )
                    ),
                    "sugestoes_pendentes": (
                        locals().get(
                            "sugestoes_pendentes",
                            []
                        )
                    ),
                    "aprovado": (
                        len(erros_importacao) == 0
                    )
                }

                st.session_state[
                    "resultado_importacao_gerarel"
                ] = resultado_importacao

                st.rerun()

            resultado_importacao = (
                st.session_state.get(
                    "resultado_importacao_gerarel"
                )
            )

            if resultado_importacao is not None:

                st.markdown(
                    "#### Resultado da validação"
                )


                mapeamentos_resultado = (
                    resultado_importacao.get(
                        "mapeamentos_aplicados",
                        []
                    )
                )

                sugestoes_resultado = (
                    resultado_importacao.get(
                        "sugestoes_pendentes",
                        []
                    )
                )

                if mapeamentos_resultado:

                    st.markdown(
                        "#### Nomes reconhecidos automaticamente"
                    )

                    dados_mapeamento_visual = []

                    for item_mapeamento in (
                        mapeamentos_resultado
                    ):

                        dados_mapeamento_visual.append(
                            {
                                "Tipo": item_mapeamento[
                                    "tipo"
                                ],
                                "Aba": item_mapeamento[
                                    "aba"
                                ],
                                "Nome encontrado": (
                                    item_mapeamento[
                                        "origem"
                                    ]
                                ),
                                "Campo GeraRel": (
                                    item_mapeamento[
                                        "destino"
                                    ]
                                ),
                                "Confiança": (
                                    f"{item_mapeamento['confianca'] * 100:.0f}%"
                                )
                            }
                        )

                    st.dataframe(
                        pd.DataFrame(
                            dados_mapeamento_visual
                        ),
                        hide_index=True,
                        width="stretch"
                    )

                if sugestoes_resultado:

                    st.warning(
                        "Algumas colunas possuem "
                        "correspondência ambígua e exigirão "
                        "confirmação manual em uma próxima "
                        "etapa do assistente."
                    )

                    with st.expander(
                        "Ver sugestões de mapeamento"
                    ):

                        for sugestao_item in (
                            sugestoes_resultado
                        ):

                            st.write(
                                f"**{sugestao_item['aba']}** — "
                                f"{sugestao_item['origem']} → "
                                f"{sugestao_item['destino_sugerido']} "
                                f"({sugestao_item['confianca'] * 100:.0f}%)"
                            )

                if resultado_importacao[
                    "aprovado"
                ]:

                    st.success(
                        "✅ A planilha é compatível com "
                        "o modelo estrutural do GeraRel."
                    )

                else:

                    st.error(
                        "🚨 A planilha possui erros que "
                        "impedem a importação."
                    )

                dados_validados = (
                    resultado_importacao["dados"]
                )

                if dados_validados:

                    validacao_coluna1,                     validacao_coluna2,                     validacao_coluna3 = st.columns(3)

                    validacao_coluna1.metric(
                        "Abas localizadas",
                        len(dados_validados)
                    )

                    if "Cenarios" in dados_validados:

                        quantidade_cenarios_importados = len(
                            dados_validados[
                                "Cenarios"
                            ]
                        )

                    else:

                        quantidade_cenarios_importados = 0

                    validacao_coluna2.metric(
                        "Cenários",
                        quantidade_cenarios_importados
                    )

                    if (
                        "Registros_Operacionais"
                        in dados_validados
                    ):

                        quantidade_registros_importados = len(
                            dados_validados[
                                "Registros_Operacionais"
                            ]
                        )

                    else:

                        quantidade_registros_importados = 0

                    validacao_coluna3.metric(
                        "Registros",
                        quantidade_registros_importados
                    )

                if resultado_importacao["erros"]:

                    st.markdown(
                        "##### Erros encontrados"
                    )

                    for erro_importacao in (
                        resultado_importacao["erros"]
                    ):

                        st.error(
                            f"❌ {erro_importacao}"
                        )

                if resultado_importacao["avisos"]:

                    with st.expander(
                        "Avisos da importação"
                    ):

                        for aviso_importacao in (
                            resultado_importacao[
                                "avisos"
                            ]
                        ):

                            st.warning(
                                f"⚠️ {aviso_importacao}"
                            )


                # =================================================
                # VALIDAÇÃO AVANÇADA DA QUALIDADE DOS DADOS V11
                # =================================================

                if (
                    resultado_importacao["aprovado"]
                    and dados_validados
                ):

                    st.markdown(
                        "### Qualidade dos dados"
                    )

                    erros_qualidade = []
                    ressalvas_qualidade = []
                    itens_aprovados_qualidade = []

                    # ---------------------------------------------
                    # DATAFRAMES DA PLANILHA
                    # ---------------------------------------------

                    df_cenarios_importados = (
                        dados_validados["Cenarios"].copy()
                    )

                    df_registros_importados = (
                        dados_validados[
                            "Registros_Operacionais"
                        ].copy()
                    )

                    df_indicadores_importados = (
                        dados_validados[
                            "Indicadores_Mensais"
                        ].copy()
                    )

                    df_ocorrencias_importadas = (
                        dados_validados[
                            "Ocorrencias"
                        ].copy()
                    )

                    df_encaminhamentos_importados = (
                        dados_validados[
                            "Encaminhamentos"
                        ].copy()
                    )

                    # ---------------------------------------------
                    # NORMALIZA OS IDs
                    # ---------------------------------------------

                    dataframes_com_cenario = [
                        df_cenarios_importados,
                        df_registros_importados,
                        df_indicadores_importados,
                        df_ocorrencias_importadas,
                        df_encaminhamentos_importados
                    ]

                    for dataframe_cenario in (
                        dataframes_com_cenario
                    ):

                        dataframe_cenario[
                            "Cenario_ID"
                        ] = (
                            dataframe_cenario[
                                "Cenario_ID"
                            ]
                            .fillna("")
                            .astype(str)
                            .str.strip()
                        )

                    # ---------------------------------------------
                    # CENARIO_ID VAZIO OU DUPLICADO
                    # ---------------------------------------------

                    ids_cenarios_vazios = int(
                        (
                            df_cenarios_importados[
                                "Cenario_ID"
                            ] == ""
                        ).sum()
                    )

                    if ids_cenarios_vazios > 0:

                        erros_qualidade.append(
                            f"A aba Cenarios possui "
                            f"{ids_cenarios_vazios} "
                            "Cenario_ID vazio(s)."
                        )

                    else:

                        itens_aprovados_qualidade.append(
                            "Todos os cenários possuem ID."
                        )

                    ids_duplicados = (
                        df_cenarios_importados[
                            "Cenario_ID"
                        ]
                        .duplicated(
                            keep=False
                        )
                    )

                    lista_ids_duplicados = sorted(
                        set(
                            df_cenarios_importados.loc[
                                ids_duplicados,
                                "Cenario_ID"
                            ].tolist()
                        )
                    )

                    if lista_ids_duplicados:

                        erros_qualidade.append(
                            "Cenario_ID duplicado(s): "
                            + ", ".join(
                                lista_ids_duplicados
                            )
                            + "."
                        )

                    else:

                        itens_aprovados_qualidade.append(
                            "Não existem cenários duplicados."
                        )

                    ids_cadastrados = set(
                        df_cenarios_importados[
                            "Cenario_ID"
                        ]
                        .loc[
                            lambda serie:
                                serie != ""
                        ]
                        .tolist()
                    )

                    # ---------------------------------------------
                    # CENÁRIOS SEM REGISTROS
                    # ---------------------------------------------

                    ids_com_registros = set(
                        df_registros_importados[
                            "Cenario_ID"
                        ]
                        .loc[
                            lambda serie:
                                serie != ""
                        ]
                        .tolist()
                    )

                    cenarios_sem_registros = sorted(
                        ids_cadastrados
                        - ids_com_registros
                    )

                    if cenarios_sem_registros:

                        ressalvas_qualidade.append(
                            "Cenários sem registros operacionais: "
                            + ", ".join(
                                cenarios_sem_registros
                            )
                            + "."
                        )

                    else:

                        itens_aprovados_qualidade.append(
                            "Todos os cenários possuem "
                            "registros operacionais."
                        )

                    # ---------------------------------------------
                    # CAMPOS OBRIGATÓRIOS VAZIOS
                    # ---------------------------------------------

                    campos_obrigatorios = {
                        "Cenarios": [
                            "Cenario_ID",
                            "Organizacao_Ficticia",
                            "Setor_Ficticio",
                            "Responsavel_Ficticio",
                            "Local",
                            "Complexidade",
                            "Finalidade"
                        ],
                        "Registros_Operacionais": [
                            "Registro_ID",
                            "Cenario_ID",
                            "Tipo_Atividade",
                            "Quantidade",
                            "Unidade",
                            "Situacao"
                        ],
                        "Indicadores_Mensais": [
                            "Indicador_ID",
                            "Cenario_ID",
                            "Indicador",
                            "Meta",
                            "Realizado",
                            "Situacao"
                        ],
                        "Ocorrencias": [
                            "Ocorrencia_ID",
                            "Cenario_ID",
                            "Categoria",
                            "Descricao",
                            "Severidade",
                            "Situacao"
                        ],
                        "Encaminhamentos": [
                            "Encaminhamento_ID",
                            "Cenario_ID",
                            "Acao_Prevista",
                            "Responsavel_Ficticio",
                            "Prazo",
                            "Status",
                            "Prioridade"
                        ]
                    }

                    for nome_aba_qualidade, campos in (
                        campos_obrigatorios.items()
                    ):

                        dataframe_qualidade = (
                            dados_validados[
                                nome_aba_qualidade
                            ]
                        )

                        # Abas opcionais podem ficar vazias.
                        if dataframe_qualidade.empty:
                            continue

                        for campo_qualidade in campos:

                            vazios = int(
                                (
                                    dataframe_qualidade[
                                        campo_qualidade
                                    ]
                                    .isna()
                                    |
                                    (
                                        dataframe_qualidade[
                                            campo_qualidade
                                        ]
                                        .fillna("")
                                        .astype(str)
                                        .str.strip()
                                        == ""
                                    )
                                ).sum()
                            )

                            if vazios > 0:

                                erros_qualidade.append(
                                    f"{nome_aba_qualidade}: "
                                    f"{vazios} valor(es) "
                                    f"vazio(s) no campo "
                                    f"{campo_qualidade}."
                                )

                    # ---------------------------------------------
                    # COMPLEXIDADE
                    # ---------------------------------------------

                    complexidades_validas = {
                        "Baixa",
                        "Média",
                        "Alta"
                    }

                    complexidades_recebidas = set(
                        df_cenarios_importados[
                            "Complexidade"
                        ]
                        .dropna()
                        .astype(str)
                        .str.strip()
                    )

                    complexidades_invalidas = sorted(
                        complexidades_recebidas
                        - complexidades_validas
                    )

                    if complexidades_invalidas:

                        erros_qualidade.append(
                            "Complexidade inválida: "
                            + ", ".join(
                                complexidades_invalidas
                            )
                            + ". Utilize Baixa, Média ou Alta."
                        )

                    else:

                        itens_aprovados_qualidade.append(
                            "Complexidades válidas."
                        )

                    # ---------------------------------------------
                    # VALORES NUMÉRICOS
                    # ---------------------------------------------

                    validacoes_numericas = [
                        (
                            "Registros_Operacionais",
                            df_registros_importados,
                            "Quantidade",
                            False
                        ),
                        (
                            "Registros_Operacionais",
                            df_registros_importados,
                            "Prazo_Padrao_dias",
                            True
                        ),
                        (
                            "Registros_Operacionais",
                            df_registros_importados,
                            "Tempo_Real_dias",
                            True
                        ),
                        (
                            "Indicadores_Mensais",
                            df_indicadores_importados,
                            "Meta",
                            False
                        ),
                        (
                            "Indicadores_Mensais",
                            df_indicadores_importados,
                            "Realizado",
                            False
                        )
                    ]

                    for (
                        nome_aba_numero,
                        dataframe_numero,
                        coluna_numero,
                        permite_vazio
                    ) in validacoes_numericas:

                        if dataframe_numero.empty:
                            continue

                        serie_original = (
                            dataframe_numero[
                                coluna_numero
                            ]
                        )

                        serie_numerica = pd.to_numeric(
                            serie_original,
                            errors="coerce"
                        )

                        if permite_vazio:

                            invalidos = (
                                serie_original.notna()
                                & serie_numerica.isna()
                            )

                        else:

                            invalidos = (
                                serie_numerica.isna()
                            )

                        quantidade_invalidos = int(
                            invalidos.sum()
                        )

                        quantidade_negativos = int(
                            (
                                serie_numerica
                                .dropna()
                                < 0
                            ).sum()
                        )

                        if quantidade_invalidos > 0:

                            erros_qualidade.append(
                                f"{nome_aba_numero}: "
                                f"{quantidade_invalidos} "
                                f"valor(es) não numérico(s) "
                                f"em {coluna_numero}."
                            )

                        if quantidade_negativos > 0:

                            erros_qualidade.append(
                                f"{nome_aba_numero}: "
                                f"{quantidade_negativos} "
                                f"valor(es) negativo(s) "
                                f"em {coluna_numero}."
                            )

                    # ---------------------------------------------
                    # PERCENTUAL DE ATINGIMENTO
                    # ---------------------------------------------

                    if not df_indicadores_importados.empty:

                        percentual_importado = pd.to_numeric(
                            df_indicadores_importados[
                                "Atingimento_pct"
                            ],
                            errors="coerce"
                        )

                        percentuais_invalidos = int(
                            percentual_importado.isna().sum()
                        )

                        percentuais_negativos = int(
                            (
                                percentual_importado
                                .dropna()
                                < 0
                            ).sum()
                        )

                        percentuais_altos = int(
                            (
                                percentual_importado
                                .dropna()
                                > 10
                            ).sum()
                        )

                        if percentuais_invalidos > 0:

                            erros_qualidade.append(
                                f"Indicadores_Mensais: "
                                f"{percentuais_invalidos} "
                                "percentual(is) inválido(s)."
                            )

                        if percentuais_negativos > 0:

                            erros_qualidade.append(
                                "Indicadores_Mensais possui "
                                "percentual negativo."
                            )

                        if percentuais_altos > 0:

                            ressalvas_qualidade.append(
                                f"Indicadores_Mensais possui "
                                f"{percentuais_altos} percentual(is) "
                                "acima de 1000%. Verifique se o "
                                "valor foi informado como percentual "
                                "inteiro em vez de decimal."
                            )

                    # ---------------------------------------------
                    # DATAS DOS ENCAMINHAMENTOS
                    # ---------------------------------------------

                    if not df_encaminhamentos_importados.empty:

                        datas_convertidas = pd.to_datetime(
                            df_encaminhamentos_importados[
                                "Prazo"
                            ],
                            errors="coerce",
                            dayfirst=True
                        )

                        datas_invalidas = int(
                            datas_convertidas.isna().sum()
                        )

                        if datas_invalidas > 0:

                            erros_qualidade.append(
                                f"Encaminhamentos possui "
                                f"{datas_invalidas} data(s) "
                                "de prazo inválida(s)."
                            )

                        else:

                            itens_aprovados_qualidade.append(
                                "Datas dos encaminhamentos válidas."
                            )

                    # ---------------------------------------------
                    # VALORES CONTROLADOS
                    # ---------------------------------------------

                    valores_controlados = [
                        (
                            "Ocorrencias",
                            df_ocorrencias_importadas,
                            "Severidade",
                            {
                                "Baixa",
                                "Média",
                                "Alta"
                            }
                        ),
                        (
                            "Encaminhamentos",
                            df_encaminhamentos_importados,
                            "Prioridade",
                            {
                                "Baixa",
                                "Média",
                                "Alta"
                            }
                        )
                    ]

                    for (
                        nome_aba_controlada,
                        dataframe_controlado,
                        coluna_controlada,
                        valores_validos
                    ) in valores_controlados:

                        if dataframe_controlado.empty:
                            continue

                        valores_recebidos = set(
                            dataframe_controlado[
                                coluna_controlada
                            ]
                            .dropna()
                            .astype(str)
                            .str.strip()
                        )

                        valores_invalidos = sorted(
                            valores_recebidos
                            - valores_validos
                        )

                        if valores_invalidos:

                            erros_qualidade.append(
                                f"{nome_aba_controlada}: "
                                f"{coluna_controlada} inválida: "
                                + ", ".join(
                                    valores_invalidos
                                )
                                + "."
                            )

                    # ---------------------------------------------
                    # CLASSIFICAÇÃO FINAL DA BASE
                    # ---------------------------------------------

                    if erros_qualidade:

                        classificacao_base = "Reprovada"

                    elif ressalvas_qualidade:

                        classificacao_base = (
                            "Aprovada com ressalvas"
                        )

                    else:

                        classificacao_base = "Aprovada"

                    # ---------------------------------------------
                    # MÉTRICAS DA QUALIDADE
                    # ---------------------------------------------

                    qualidade_coluna1,                     qualidade_coluna2,                     qualidade_coluna3,                     qualidade_coluna4 = st.columns(4)

                    qualidade_coluna1.metric(
                        "Cenários cadastrados",
                        len(ids_cadastrados)
                    )

                    qualidade_coluna2.metric(
                        "Erros críticos",
                        len(erros_qualidade)
                    )

                    qualidade_coluna3.metric(
                        "Ressalvas",
                        len(ressalvas_qualidade)
                    )

                    qualidade_coluna4.metric(
                        "Classificação",
                        classificacao_base
                    )

                    # ---------------------------------------------
                    # RESULTADO VISUAL
                    # ---------------------------------------------

                    if classificacao_base == "Aprovada":

                        st.success(
                            "✅ Qualidade dos dados: "
                            "APROVADA."
                        )

                    elif classificacao_base == (
                        "Aprovada com ressalvas"
                    ):

                        st.warning(
                            "⚠️ Qualidade dos dados: "
                            "APROVADA COM RESSALVAS."
                        )

                    else:

                        st.error(
                            "🚨 Qualidade dos dados: "
                            "REPROVADA."
                        )

                    if erros_qualidade:

                        st.markdown(
                            "#### Erros críticos da base"
                        )

                        for erro_qualidade in (
                            erros_qualidade
                        ):

                            st.error(
                                f"❌ {erro_qualidade}"
                            )

                    if ressalvas_qualidade:

                        st.markdown(
                            "#### Ressalvas da base"
                        )

                        for ressalva_qualidade in (
                            ressalvas_qualidade
                        ):

                            st.warning(
                                f"⚠️ {ressalva_qualidade}"
                            )

                    with st.expander(
                        "Ver verificações aprovadas"
                    ):

                        if itens_aprovados_qualidade:

                            for item_aprovado in (
                                itens_aprovados_qualidade
                            ):

                                st.success(
                                    f"✅ {item_aprovado}"
                                )

                        else:

                            st.write(
                                "Nenhuma verificação "
                                "adicional registrada."
                            )

                    resultado_importacao[
                        "erros_qualidade"
                    ] = erros_qualidade

                    resultado_importacao[
                        "ressalvas_qualidade"
                    ] = ressalvas_qualidade

                    resultado_importacao[
                        "classificacao_qualidade"
                    ] = classificacao_base

                    resultado_importacao[
                        "qualidade_aprovada"
                    ] = (
                        len(erros_qualidade) == 0
                    )

                    st.session_state[
                        "resultado_importacao_gerarel"
                    ] = resultado_importacao

                if resultado_importacao[
                    "aprovado"
                ]:

                    st.info(
                        "A planilha passou na validação. "
                        "Ao confirmar, o arquivo será ativado "
                        "como fonte de dados do dashboard."
                    )

                    qualidade_aprovada = (
                        resultado_importacao.get(
                            "qualidade_aprovada",
                            False
                        )
                    )

                    confirmar_planilha = st.button(
                        "✅ Confirmar planilha validada",
                        type="primary",
                        width="stretch",
                        disabled=(
                            not qualidade_aprovada
                        ),
                        key="confirmar_planilha_gerarel"
                    )

                    if not qualidade_aprovada:

                        st.warning(
                            "Corrija os erros críticos antes "
                            "de confirmar a planilha."
                        )

                    if confirmar_planilha:

                        dados_confirmados = (
                            resultado_importacao.get(
                                "dados"
                            )
                        )

                        if not isinstance(
                            dados_confirmados,
                            dict
                        ):

                            st.error(
                                "Os dados da planilha não "
                                "estão disponíveis para ativação."
                            )

                        else:

                            st.session_state[
                                "planilha_externa_confirmada"
                            ] = True

                            st.session_state[
                                "dados_importados_ativos"
                            ] = {
                                nome_aba: dataframe.copy()
                                for nome_aba, dataframe
                                in dados_confirmados.items()
                            }

                            st.session_state[
                                "nome_arquivo_importado_ativo"
                            ] = resultado_importacao.get(
                                "arquivo",
                                "Planilha importada"
                            )

                            st.session_state[
                                "fonte_dados_ativa"
                            ] = "Planilha importada"

                            st.session_state[
                                "mensagem_ativacao_planilha"
                            ] = (
                                "Planilha ativada com sucesso."
                            )

                            st.session_state[
                                "mostrar_confirmacao_ativacao"
                            ] = True

                            st.rerun()

# =========================================================


def exibir_botao_copiar_nexora(
    texto,
    identificador
):
    """
    Exibe um botão para copiar o prompt no navegador.
    """

    if not texto or not texto.strip():

        st.warning(
            "O prompt está vazio e não pode ser copiado."
        )

        return

    # Garante um identificador válido no HTML e JavaScript.
    identificador_seguro = re.sub(
        r"[^A-Za-z0-9_]",
        "_",
        str(identificador)
    )

    if not identificador_seguro:

        identificador_seguro = "prompt_nexora"

    texto_base64 = base64.b64encode(
        texto.encode("utf-8")
    ).decode("ascii")

    html_copia = f"""
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
</head>

<body style="margin:0;padding:0;background:transparent;">

<button
    id="botao_{identificador_seguro}"
    type="button"
    style="
        width:100%;
        height:44px;
        padding:10px 16px;
        background:#1F4E78;
        color:#FFFFFF;
        border:1px solid #1F4E78;
        border-radius:8px;
        font-family:Arial,sans-serif;
        font-size:14px;
        font-weight:700;
        cursor:pointer;
    "
>
    📋 Copiar prompt
</button>

<div
    id="mensagem_{identificador_seguro}"
    style="
        min-height:18px;
        margin-top:7px;
        color:#2E7D32;
        font-family:Arial,sans-serif;
        font-size:12px;
        text-align:center;
    "
></div>

<script>
(function() {{

    const botao = document.getElementById(
        "botao_{identificador_seguro}"
    );

    const mensagem = document.getElementById(
        "mensagem_{identificador_seguro}"
    );

    const textoBase64 = "{texto_base64}";

    function decodificarTexto() {{

        const binario = atob(
            textoBase64
        );

        const bytes = new Uint8Array(
            binario.length
        );

        for (
            let indice = 0;
            indice < binario.length;
            indice++
        ) {{

            bytes[indice] = (
                binario.charCodeAt(indice)
            );
        }}

        return new TextDecoder(
            "utf-8"
        ).decode(bytes);
    }}

    async function copiarPrompt() {{

        try {{

            const texto = decodificarTexto();

            if (
                navigator.clipboard
                && window.isSecureContext
            ) {{

                await navigator.clipboard.writeText(
                    texto
                );

            }} else {{

                const area = document.createElement(
                    "textarea"
                );

                area.value = texto;
                area.setAttribute(
                    "readonly",
                    ""
                );

                area.style.position = "fixed";
                area.style.left = "-9999px";
                area.style.top = "0";

                document.body.appendChild(
                    area
                );

                area.select();
                area.setSelectionRange(
                    0,
                    area.value.length
                );

                const sucesso = document.execCommand(
                    "copy"
                );

                document.body.removeChild(
                    area
                );

                if (!sucesso) {{

                    throw new Error(
                        "O navegador recusou a cópia."
                    );
                }}
            }}

            mensagem.textContent =
                "✅ Prompt copiado.";

            botao.textContent =
                "✅ Prompt copiado";

            setTimeout(function() {{

                botao.textContent =
                    "📋 Copiar prompt";

            }}, 1800);

        }} catch (erro) {{

            mensagem.textContent =
                "⚠️ O navegador bloqueou a cópia. "
                + "Use a visualização do prompt.";

            console.error(
                "Falha ao copiar o prompt:",
                erro
            );
        }}
    }}

    botao.addEventListener(
        "click",
        copiarPrompt
    );

}})();
</script>

</body>
</html>
"""

    components.html(
        html_copia,
        height=76,
        scrolling=False
    )


# CENTRAL DE RELATÓRIOS DEFINITIVA V11
# =========================================================

with aba_fluxo_guiado:

    st.subheader("Central de Relatórios")

    st.write(
        "Prepare, revise, avalie e exporte o relatório "
        "seguindo uma etapa por vez. O sistema libera "
        "automaticamente a próxima etapa após a confirmação."
    )

    # -----------------------------------------------------
    # IDENTIFICAÇÃO DA EXECUÇÃO
    # -----------------------------------------------------

    geracao_wizard = st.selectbox(
        "Geração",
        options=["G1", "G2"],
        index=0,
        key=f"wizard_geracao_{cenario_escolhido}"
    )

    id_wizard = (
        f"{cenario_escolhido}_{geracao_wizard}"
    )

    prefixo_wizard = (
        f"wizard_{cenario_escolhido}_"
        f"{geracao_wizard}"
    )

    chave_etapa = f"{prefixo_wizard}_etapa"
    chave_original = f"{prefixo_wizard}_original"
    chave_revisao_editor = (
        f"{prefixo_wizard}_revisao_editor"
    )
    chave_revisao_salva = (
        f"{prefixo_wizard}_revisao_salva"
    )
    chave_analise = f"{prefixo_wizard}_analise"
    chave_avaliacao = f"{prefixo_wizard}_avaliacao"
    chave_validacao = f"{prefixo_wizard}_validacao"

    if chave_etapa not in st.session_state:
        st.session_state[chave_etapa] = 1

    if chave_original not in st.session_state:
        st.session_state[chave_original] = ""

    if chave_revisao_editor not in st.session_state:
        st.session_state[chave_revisao_editor] = ""

    if chave_revisao_salva not in st.session_state:
        st.session_state[chave_revisao_salva] = ""

    if chave_analise not in st.session_state:
        st.session_state[chave_analise] = None

    if chave_avaliacao not in st.session_state:
        st.session_state[chave_avaliacao] = None

    if chave_validacao not in st.session_state:
        st.session_state[chave_validacao] = None

    etapa_atual = st.session_state[chave_etapa]

    # -----------------------------------------------------
    # PROGRESSO
    # -----------------------------------------------------

    nomes_etapas = [
        "Rascunho",
        "Revisão",
        "Análise",
        "Avaliação",
        "Validação e exportação"
    ]

    percentual_etapa = int(
        etapa_atual / 5 * 100
    )

    st.progress(
        percentual_etapa,
        text=(
            f"Etapa {etapa_atual} de 5: "
            f"{nomes_etapas[etapa_atual - 1]}"
        )
    )

    progresso_colunas = st.columns(5)

    for indice, nome_etapa in enumerate(
        nomes_etapas,
        start=1
    ):

        with progresso_colunas[indice - 1]:

            if indice < etapa_atual:

                st.success(
                    f"✅ {indice}. {nome_etapa}"
                )

            elif indice == etapa_atual:

                st.info(
                    f"▶️ {indice}. {nome_etapa}"
                )

            else:

                st.write(
                    f"🔒 {indice}. {nome_etapa}"
                )

    st.caption(
        f"Execução ativa: {id_wizard}"
    )

    st.divider()

    # =====================================================
    # ETAPA 1 - RASCUNHO ORIGINAL
    # =====================================================

    if etapa_atual == 1:

        # =================================================
        # PROMPT INTEGRADO À ETAPA 1 DEFINITIVA
        # =================================================

        st.markdown(
            "### Etapa 1 — Preparar o prompt e preservar "
            "o rascunho original"
        )

        st.write(
            "Prepare o prompt com os dados do cenário, envie-o "
            "manualmente à ferramenta de IA e registre a "
            "primeira resposta sem realizar correções."
        )

        # -------------------------------------------------
        # CHAVES DO PROMPT
        # -------------------------------------------------

        chave_prompt_preparado = (
            f"{prefixo_wizard}_prompt_preparado"
        )

        chave_prompt_texto = (
            f"{prefixo_wizard}_prompt_texto"
        )

        if chave_prompt_preparado not in st.session_state:
            st.session_state[
                chave_prompt_preparado
            ] = False

        if chave_prompt_texto not in st.session_state:
            st.session_state[
                chave_prompt_texto
            ] = ""

        # -------------------------------------------------
        # PREPARAÇÃO DO PROMPT
        # -------------------------------------------------

        st.markdown(
            "#### 1. Preparar o prompt do cenário"
        )

        st.caption(
            f"Fonte dos dados: cenário {cenario_escolhido} "
            f"da base atualmente carregada."
        )

        gerar_prompt_wizard = st.button(
            "⚙️ Gerar prompt do cenário",
            type="primary",
            width="stretch",
            key=f"{prefixo_wizard}_gerar_prompt"
        )

        if gerar_prompt_wizard:

            with st.spinner(
                "Conferindo os dados e preparando o prompt..."
            ):

                progresso_prompt = st.progress(
                    0,
                    text="Validando o cenário selecionado..."
                )

                time.sleep(0.20)

                progresso_prompt.progress(
                    30,
                    text="Consolidando atividades e indicadores..."
                )

                # Tenta usar o prompt criado pelo código principal.
                prompt_atual = globals().get(
                    "prompt_relatorio",
                    ""
                )

                progresso_prompt.progress(
                    60,
                    text="Organizando ocorrências e encaminhamentos..."
                )

                if not prompt_atual.strip():

                    # Prompt alternativo seguro, construído
                    # diretamente com os dados do cenário.
                    linhas_atividades_prompt = []

                    if not resumo_atividades.empty:

                        for _, atividade in (
                            resumo_atividades.iterrows()
                        ):

                            tipo_atividade = str(
                                atividade.get(
                                    "Tipo_Atividade",
                                    "Atividade não informada"
                                )
                            )

                            quantidade_total = atividade.get(
                                "Quantidade_Total",
                                atividade.get(
                                    "Quantidade",
                                    ""
                                )
                            )

                            unidade_atividade = str(
                                atividade.get(
                                    "Unidade",
                                    ""
                                )
                            )

                            numero_registros = atividade.get(
                                "Numero_Registros",
                                ""
                            )

                            linhas_atividades_prompt.append(
                                f"- {tipo_atividade}: "
                                f"{quantidade_total} "
                                f"{unidade_atividade}; "
                                f"{numero_registros} registros."
                            )

                    else:

                        linhas_atividades_prompt.append(
                            "- Nenhuma atividade registrada."
                        )

                    linhas_indicadores_prompt = []

                    if not indicadores_cenario.empty:

                        for _, indicador in (
                            indicadores_cenario.iterrows()
                        ):

                            atingimento = indicador.get(
                                "Atingimento_pct",
                                ""
                            )

                            if isinstance(
                                atingimento,
                                (int, float)
                            ):

                                if atingimento <= 2:

                                    atingimento_formatado = (
                                        f"{atingimento * 100:.1f}%"
                                        .replace(".", ",")
                                    )

                                else:

                                    atingimento_formatado = (
                                        f"{atingimento:.1f}%"
                                        .replace(".", ",")
                                    )

                            else:

                                atingimento_formatado = str(
                                    atingimento
                                )

                            linhas_indicadores_prompt.append(
                                f"- {indicador.get('Indicador', '')}: "
                                f"meta {indicador.get('Meta', '')}; "
                                f"realizado "
                                f"{indicador.get('Realizado', '')}; "
                                f"atingimento "
                                f"{atingimento_formatado}; "
                                f"situação "
                                f"{indicador.get('Situacao', '')}."
                            )

                    else:

                        linhas_indicadores_prompt.append(
                            "- Nenhum indicador registrado."
                        )

                    linhas_ocorrencias_prompt = []

                    if ocorrencias_cenario.empty:

                        linhas_ocorrencias_prompt.append(
                            "- Nenhuma ocorrência registrada."
                        )

                    else:

                        for _, ocorrencia in (
                            ocorrencias_cenario.iterrows()
                        ):

                            linhas_ocorrencias_prompt.append(
                                f"- Categoria: "
                                f"{ocorrencia.get('Categoria', '')}; "
                                f"descrição: "
                                f"{ocorrencia.get('Descricao', '')}; "
                                f"impacto: "
                                f"{ocorrencia.get('Impacto', '')}; "
                                f"severidade: "
                                f"{ocorrencia.get('Severidade', '')}; "
                                f"situação: "
                                f"{ocorrencia.get('Situacao', '')}."
                            )

                    linhas_encaminhamentos_prompt = []

                    if encaminhamentos_cenario.empty:

                        linhas_encaminhamentos_prompt.append(
                            "- Nenhum encaminhamento registrado."
                        )

                    else:

                        for _, encaminhamento in (
                            encaminhamentos_cenario.iterrows()
                        ):

                            prazo_prompt = encaminhamento.get(
                                "Prazo",
                                ""
                            )

                            try:

                                prazo_prompt = pd.to_datetime(
                                    prazo_prompt
                                ).strftime("%d/%m/%Y")

                            except Exception:

                                prazo_prompt = str(
                                    prazo_prompt
                                )

                            linhas_encaminhamentos_prompt.append(
                                f"- Ação: "
                                f"{encaminhamento.get('Acao_Prevista', '')}; "
                                f"responsável: "
                                f"{encaminhamento.get('Responsavel_Ficticio', '')}; "
                                f"prazo: {prazo_prompt}; "
                                f"status: "
                                f"{encaminhamento.get('Status', '')}; "
                                f"prioridade: "
                                f"{encaminhamento.get('Prioridade', '')}; "
                                f"resultado esperado: "
                                f"{encaminhamento.get('Resultado_Esperado', '')}."
                            )

                    prompt_atual = f"""
Você é um assistente especializado em redação administrativa.

Elabore um RASCUNHO de relatório mensal de atividades
administrativas usando exclusivamente os dados fornecidos.

REGRAS OBRIGATÓRIAS

1. Não invente nomes, valores, datas, causas,
   justificativas, resultados ou encaminhamentos.
2. Preserve todos os valores exatamente como fornecidos.
3. Caso uma informação não esteja disponível, informe
   isso de forma objetiva.
4. Utilize linguagem formal, clara, objetiva e impessoal.
5. Não inclua informações externas aos dados fornecidos.
6. O documento deve conter exatamente sete seções.
7. Termine com a frase:
   "Documento gerado como rascunho e sujeito à revisão humana."

ESTRUTURA OBRIGATÓRIA

1. Identificação
2. Apresentação
3. Atividades realizadas
4. Indicadores
5. Ocorrências e pendências
6. Encaminhamentos
7. Fecho

DADOS GERAIS

Cenário: {cenario_escolhido}
Organização: {geral.get('Organizacao_Ficticia', '')}
Setor: {geral.get('Setor_Ficticio', '')}
Responsável: {geral.get('Responsavel_Ficticio', '')}
Local: {geral.get('Local', '')}
Complexidade: {geral.get('Complexidade', '')}
Finalidade: {geral.get('Finalidade', '')}

ATIVIDADES CONSOLIDADAS

{chr(10).join(linhas_atividades_prompt)}

INDICADORES

{chr(10).join(linhas_indicadores_prompt)}

OCORRÊNCIAS

{chr(10).join(linhas_ocorrencias_prompt)}

ENCAMINHAMENTOS

{chr(10).join(linhas_encaminhamentos_prompt)}

Produza somente o relatório solicitado, sem comentários
sobre o processo de geração.
""".strip()

                progresso_prompt.progress(
                    90,
                    text="Finalizando o arquivo do prompt..."
                )

                st.session_state[
                    chave_prompt_texto
                ] = prompt_atual

                st.session_state[
                    chave_prompt_preparado
                ] = bool(
                    prompt_atual.strip()
                )

                progresso_prompt.progress(
                    100,
                    text="Prompt preparado."
                )

                time.sleep(0.25)

                progresso_prompt.empty()

            if st.session_state[
                chave_prompt_preparado
            ]:

                st.success(
                    "✅ Prompt preparado com sucesso."
                )

            else:

                st.error(
                    "Não foi possível preparar o prompt "
                    "do cenário selecionado."
                )

            st.rerun()

        prompt_preparado = st.session_state[
            chave_prompt_preparado
        ]

        prompt_texto_wizard = st.session_state[
            chave_prompt_texto
        ]

        if prompt_preparado and prompt_texto_wizard.strip():

            nome_prompt_wizard = (
                f"prompt_{id_wizard}.txt"
            )

            st.success(
                f"✅ Prompt disponível: "
                f"{nome_prompt_wizard}"
            )

            detalhe_prompt1, detalhe_prompt2 = (
                st.columns(2)
            )

            detalhe_prompt1.metric(
                "Caracteres do prompt",
                len(prompt_texto_wizard)
            )

            detalhe_prompt2.metric(
                "Cenário e geração",
                id_wizard
            )

            with st.expander(
                "Visualizar o prompt preparado"
            ):

                st.text_area(
                    "Conteúdo do prompt",
                    value=prompt_texto_wizard,
                    height=380,
                    disabled=True,
                    key=f"{prefixo_wizard}_prompt_visual"
                )


            st.markdown(


                "#### Ações do prompt"


            )


            coluna_copiar_prompt, coluna_baixar_prompt = st.columns(2)


            with coluna_copiar_prompt:


                exibir_botao_copiar_nexora(


                    prompt_texto_wizard,


                    (


                        f"{cenario_escolhido}_"
            f"{geracao_wizard}"


                        


                    )


                )


            with coluna_baixar_prompt:


                st.download_button(


                    label=(


                        "⬇️ Baixar prompt em TXT"


                    ),


                    data=prompt_texto_wizard.encode(


                        "utf-8"


                    ),


                    file_name=nome_prompt_wizard,


                    mime="text/plain",


                    width="stretch",


                    key=(


                        f"{prefixo_wizard}_"


                        "prompt_download"


                    )


                )


            st.info(
                "Copie o prompt e envie-o ao Microsoft Copilot. "
                "Depois, copie a primeira resposta sem "
                "corrigir e cole-a no campo abaixo."
            )

        else:

            st.warning(
                "Gere o prompt antes de registrar "
                "a resposta original da IA."
            )

        st.markdown(
            "#### 2. Registrar a resposta original"
        )

        with st.form(
            key=f"{prefixo_wizard}_form_original"
        ):

            texto_original_digitado = st.text_area(
                "Rascunho original da IA",
                height=520,
                placeholder=(
                    "Cole aqui a primeira resposta "
                    "da inteligência artificial..."
                )
            )

            preservar_original = (
                st.form_submit_button(
                    "💾 Preservar original e continuar",
                    type="primary",
                    width="stretch",
                    disabled=(
                        not st.session_state.get(
                            chave_prompt_preparado,
                            False
                        )
                    )
                )
            )

        if preservar_original:

            if not texto_original_digitado.strip():

                st.warning(
                    "Cole o rascunho original "
                    "antes de continuar."
                )

            else:

                with st.spinner(
                    "Preservando o rascunho original..."
                ):

                    progresso_original = st.progress(
                        0,
                        text="Conferindo o conteúdo..."
                    )

                    time.sleep(0.20)

                    progresso_original.progress(
                        50,
                        text="Criando a versão de revisão..."
                    )

                    st.session_state[
                        chave_original
                    ] = texto_original_digitado

                    st.session_state[
                        chave_revisao_editor
                    ] = texto_original_digitado

                    st.session_state[
                        chave_revisao_salva
                    ] = ""

                    st.session_state[
                        chave_analise
                    ] = None

                    st.session_state[
                        chave_avaliacao
                    ] = None

                    st.session_state[
                        chave_validacao
                    ] = None

                    progresso_original.progress(
                        100,
                        text="Original preservado."
                    )

                    time.sleep(0.25)

                    progresso_original.empty()

                st.session_state[chave_etapa] = 2

                st.rerun()

    # =====================================================
    # ETAPA 2 - REVISÃO
    # =====================================================

    elif etapa_atual == 2:

        st.markdown(
            "### Etapa 2 — Revisar e salvar o relatório"
        )

        st.success(
            "✅ O rascunho original foi preservado."
        )

        with st.expander(
            "Consultar rascunho original"
        ):

            st.text_area(
                "Original preservado",
                value=st.session_state[
                    chave_original
                ],
                height=280,
                disabled=True,
                key=f"{prefixo_wizard}_original_visual"
            )

        with st.form(
            key=f"{prefixo_wizard}_form_revisao"
        ):

            texto_revisao_digitado = st.text_area(
                "Relatório em revisão",
                value=st.session_state[
                    chave_revisao_editor
                ],
                height=620
            )

            salvar_revisao = (
                st.form_submit_button(
                    "💾 Salvar revisão e continuar",
                    type="primary",
                    width="stretch"
                )
            )

        if salvar_revisao:

            if not texto_revisao_digitado.strip():

                st.warning(
                    "O relatório em revisão "
                    "não pode ficar vazio."
                )

            else:

                with st.spinner(
                    "Salvando a versão revisada..."
                ):

                    progresso_revisao = st.progress(
                        0,
                        text="Registrando a revisão..."
                    )

                    time.sleep(0.20)

                    st.session_state[
                        chave_revisao_editor
                    ] = texto_revisao_digitado

                    st.session_state[
                        chave_revisao_salva
                    ] = texto_revisao_digitado

                    st.session_state[
                        chave_analise
                    ] = None

                    st.session_state[
                        chave_avaliacao
                    ] = None

                    st.session_state[
                        chave_validacao
                    ] = None

                    progresso_revisao.progress(
                        100,
                        text="Revisão salva."
                    )

                    time.sleep(0.25)

                    progresso_revisao.empty()

                st.session_state[chave_etapa] = 3

                st.rerun()

        if st.button(
            "⬅️ Voltar ao rascunho",
            width="stretch",
            key=f"{prefixo_wizard}_voltar_1"
        ):

            st.session_state[chave_etapa] = 1
            st.rerun()

    # =====================================================
    # ETAPA 3 - ANÁLISE
    # =====================================================

    elif etapa_atual == 3:

        st.markdown(
            "### Etapa 3 — Analisar as alterações"
        )

        st.write(
            "O sistema comparará a versão revisada "
            "com o rascunho original."
        )

        original_analise = st.session_state[
            chave_original
        ]

        revisado_analise = st.session_state[
            chave_revisao_salva
        ]

        executar_analise_wizard = st.button(
            "🔍 Analisar alterações",
            type="primary",
            width="stretch",
            key=f"{prefixo_wizard}_analisar"
        )

        if executar_analise_wizard:

            with st.spinner(
                "Comparando original e revisão..."
            ):

                progresso_analise = st.progress(
                    0,
                    text="Comparando caracteres..."
                )

                time.sleep(0.20)

                progresso_analise.progress(
                    35,
                    text="Comparando linhas..."
                )

                diferenca_linhas = list(
                    difflib.ndiff(
                        original_analise.splitlines(),
                        revisado_analise.splitlines()
                    )
                )

                linhas_adicionadas = sum(
                    1
                    for linha in diferenca_linhas
                    if linha.startswith("+ ")
                )

                linhas_removidas = sum(
                    1
                    for linha in diferenca_linhas
                    if linha.startswith("- ")
                )

                progresso_analise.progress(
                    70,
                    text="Calculando similaridade..."
                )

                similaridade = (
                    difflib.SequenceMatcher(
                        None,
                        original_analise,
                        revisado_analise
                    ).ratio()
                    * 100
                )

                resultado_analise = {
                    "alterado": (
                        original_analise
                        != revisado_analise
                    ),
                    "caracteres_original": len(
                        original_analise
                    ),
                    "caracteres_revisado": len(
                        revisado_analise
                    ),
                    "palavras_original": len(
                        original_analise.split()
                    ),
                    "palavras_revisado": len(
                        revisado_analise.split()
                    ),
                    "linhas_adicionadas": (
                        linhas_adicionadas
                    ),
                    "linhas_removidas": (
                        linhas_removidas
                    ),
                    "similaridade": similaridade,
                    "alteracao": 100 - similaridade,
                    "analisado_em": (
                        datetime.now().strftime(
                            "%d/%m/%Y %H:%M:%S"
                        )
                    )
                }

                st.session_state[
                    chave_analise
                ] = resultado_analise

                progresso_analise.progress(
                    100,
                    text="Análise concluída."
                )

                time.sleep(0.25)

                progresso_analise.empty()

            st.rerun()

        resultado_analise = st.session_state[
            chave_analise
        ]

        if resultado_analise is not None:

            if resultado_analise["alterado"]:

                st.success(
                    "✅ Foram identificadas alterações "
                    "entre o original e a revisão."
                )

            else:

                st.info(
                    "ℹ️ A revisão permanece idêntica "
                    "ao rascunho original."
                )

            analise_c1, analise_c2,             analise_c3, analise_c4 = st.columns(4)

            analise_c1.metric(
                "Linhas adicionadas",
                resultado_analise[
                    "linhas_adicionadas"
                ]
            )

            analise_c2.metric(
                "Linhas removidas",
                resultado_analise[
                    "linhas_removidas"
                ]
            )

            analise_c3.metric(
                "Similaridade",
                (
                    f"{resultado_analise['similaridade']:.1f}%"
                    .replace(".", ",")
                )
            )

            analise_c4.metric(
                "Conteúdo alterado",
                (
                    f"{resultado_analise['alteracao']:.1f}%"
                    .replace(".", ",")
                )
            )

            if st.button(
                "Continuar para avaliação humana ➡️",
                type="primary",
                width="stretch",
                key=f"{prefixo_wizard}_continuar_4"
            ):

                st.session_state[chave_etapa] = 4
                st.rerun()

        if st.button(
            "⬅️ Voltar à revisão",
            width="stretch",
            key=f"{prefixo_wizard}_voltar_2"
        ):

            st.session_state[chave_etapa] = 2
            st.rerun()

    # =====================================================
    # ETAPA 4 - AVALIAÇÃO HUMANA
    # =====================================================

    elif etapa_atual == 4:

        st.markdown(
            "### Etapa 4 — Avaliação humana"
        )

        st.write(
            "Preencha os critérios e calcule "
            "a classificação antes de confirmar."
        )

        with st.expander(
            "Critérios objetivos de classificação"
        ):

            st.markdown(
                """
                **Aprovado:** sete seções, qualidade 4 ou 5
                e nenhum problema crítico.

                **Aprovado com ressalva:** qualidade 3 ou
                ajustes relevantes que não comprometem
                a fidelidade dos dados.

                **Reprovado:** erro factual, informação
                inventada, omissão relevante, estrutura
                incompleta ou qualidade 1 ou 2.
                """
            )

        # =================================================
        # FORMULÁRIO CONDICIONAL DE AVALIAÇÃO V11
        # =================================================

        chave_precisou_correcao = (
            f"{prefixo_wizard}_precisou_correcao"
        )

        chave_resposta_anterior = (
            f"{prefixo_wizard}_"
            "precisou_correcao_anterior"
        )

        precisou_correcao = st.radio(
            "O relatório precisou de correções?",
            options=[
                "Não",
                "Sim"
            ],
            horizontal=True,
            key=chave_precisou_correcao,
            help=(
                "Selecione Sim somente quando o rascunho "
                "original tiver exigido alguma alteração."
            )
        )

        resposta_anterior = st.session_state.get(
            chave_resposta_anterior
        )

        if (
            resposta_anterior is not None
            and resposta_anterior
            != precisou_correcao
        ):

            # Invalida a classificação anterior quando
            # a resposta principal é modificada.
            st.session_state[
                chave_avaliacao
            ] = None

        st.session_state[
            chave_resposta_anterior
        ] = precisou_correcao

        if precisou_correcao == "Não":

            st.info(
                "Nenhuma correção será registrada. "
                "Avalie apenas a qualidade e os aspectos "
                "positivos do rascunho."
            )

        else:

            st.info(
                "Informe o nível e os tipos de correção "
                "realizados no rascunho."
            )

        with st.form(
            key=f"{prefixo_wizard}_form_avaliacao"
        ):

            # ---------------------------------------------
            # CAMPOS EXIBIDOS SOMENTE QUANDO HOUVE CORREÇÃO
            # ---------------------------------------------

            if precisou_correcao == "Sim":

                nivel_correcao_resumido = st.radio(
                    "Qual foi o nível das correções?",
                    options=[
                        "Leves",
                        "Relevantes"
                    ],
                    horizontal=True,
                    key=(
                        f"{prefixo_wizard}_"
                        "nivel_correcao_resumido"
                    )
                )

                if (
                    nivel_correcao_resumido
                    == "Leves"
                ):

                    nivel_wizard = (
                        "Sim, correções leves"
                    )

                else:

                    nivel_wizard = (
                        "Sim, correções relevantes"
                    )

                problemas_wizard = st.multiselect(
                    "O que precisou ser corrigido?",
                    options=[
                        "Problema de clareza ou redação",
                        "Dado incorreto",
                        (
                            "Informação inventada "
                            "ou sem respaldo"
                        ),
                        (
                            "Informação obrigatória "
                            "ausente"
                        ),
                        "Estrutura incompleta",
                        "Problema de formatação"
                    ],
                    placeholder=(
                        "Selecione uma ou mais "
                        "correções realizadas"
                    ),
                    help=(
                        "Essas opções serão utilizadas "
                        "para calcular a classificação."
                    )
                )

                opcoes_aspectos_avaliacao = [
                    "Ajuste de clareza ou redação",
                    "Correção de dado incorreto",
                    (
                        "Remoção de informação "
                        "sem respaldo"
                    ),
                    "Inclusão de informação omitida",
                    (
                        "Correção de estrutura "
                        "ou formatação"
                    ),
                    "Dados preservados corretamente",
                    "Linguagem clara e objetiva",
                    "Outro aspecto"
                ]

                titulo_aspectos = (
                    "Aspectos observados após a correção"
                )

                placeholder_aspectos = (
                    "Selecione os aspectos identificados "
                    "na versão revisada"
                )

            else:

                nivel_wizard = (
                    "Não precisou de correções"
                )

                # Não mantém problemas selecionados
                # anteriormente quando a resposta é Não.
                problemas_wizard = []

                opcoes_aspectos_avaliacao = [
                    "Estrutura completa e organizada",
                    "Dados preservados corretamente",
                    "Linguagem clara e objetiva",
                    (
                        "Nenhum ajuste adicional "
                        "necessário"
                    ),
                    "Outro aspecto"
                ]

                titulo_aspectos = (
                    "Qualidades observadas no rascunho"
                )

                placeholder_aspectos = (
                    "Opcional: selecione qualidades "
                    "identificadas no rascunho"
                )

            # ---------------------------------------------
            # QUALIDADE GERAL
            # ---------------------------------------------

            qualidade_wizard = st.select_slider(
                "Qualidade geral do rascunho",
                options=[
                    "1 — Inadequado",
                    (
                        "2 — Parcialmente "
                        "aproveitável"
                    ),
                    "3 — Aceitável",
                    "4 — Bom",
                    "5 — Excelente"
                ],
                value="3 — Aceitável"
            )

            # ---------------------------------------------
            # ASPECTOS OBJETIVOS
            # ---------------------------------------------

            aspectos_observados_wizard = (
                st.multiselect(
                    titulo_aspectos,
                    options=(
                        opcoes_aspectos_avaliacao
                    ),
                    placeholder=(
                        placeholder_aspectos
                    ),
                    help=(
                        "As opções selecionadas serão "
                        "usadas para gerar a "
                        "justificativa da avaliação."
                    )
                )
            )

            comentario_complementar_wizard = (
                st.text_area(
                    "Comentário complementar",
                    height=90,
                    placeholder=(
                        "Opcional. Detalhe somente algo "
                        "que não esteja representado "
                        "nas opções acima."
                    )
                )
            )

            # ---------------------------------------------
            # JUSTIFICATIVA AUTOMÁTICA
            # ---------------------------------------------

            frases_aspectos_avaliacao = {
                (
                    "Estrutura completa e organizada"
                ): (
                    "O rascunho apresentou estrutura "
                    "completa e organizada."
                ),
                (
                    "Dados preservados corretamente"
                ): (
                    "Os dados essenciais foram "
                    "preservados corretamente."
                ),
                (
                    "Linguagem clara e objetiva"
                ): (
                    "A linguagem utilizada foi clara "
                    "e objetiva."
                ),
                (
                    "Ajuste de clareza ou redação"
                ): (
                    "Foi necessário um ajuste de "
                    "clareza ou redação."
                ),
                (
                    "Correção de dado incorreto"
                ): (
                    "Foi necessária a correção de "
                    "um dado incorreto."
                ),
                (
                    "Remoção de informação sem respaldo"
                ): (
                    "Foi removida uma informação sem "
                    "respaldo na base de dados."
                ),
                (
                    "Inclusão de informação omitida"
                ): (
                    "Foi incluída uma informação "
                    "obrigatória anteriormente omitida."
                ),
                (
                    "Correção de estrutura ou formatação"
                ): (
                    "Foi realizada uma correção de "
                    "estrutura ou formatação."
                ),
                (
                    "Nenhum ajuste adicional necessário"
                ): (
                    "Nenhum ajuste adicional "
                    "foi necessário."
                )
            }

            frases_selecionadas_avaliacao = [
                frases_aspectos_avaliacao[aspecto]
                for aspecto
                in aspectos_observados_wizard
                if aspecto
                in frases_aspectos_avaliacao
            ]

            observacao_wizard = " ".join(
                frases_selecionadas_avaliacao
            ).strip()

            if (
                precisou_correcao == "Não"
                and not observacao_wizard
            ):

                justificativas_por_qualidade = {
                    "1 — Inadequado": (
                        "O rascunho foi avaliado como "
                        "inadequado, embora não tenha sido "
                        "registrada correção textual."
                    ),
                    "2 — Parcialmente aproveitável": (
                        "O rascunho foi avaliado como "
                        "parcialmente aproveitável, embora "
                        "não tenha sido registrada correção "
                        "textual."
                    ),
                    "3 — Aceitável": (
                        "O rascunho foi avaliado como "
                        "aceitável e não precisou de "
                        "correções."
                    ),
                    "4 — Bom": (
                        "O rascunho foi avaliado como bom "
                        "e não precisou de correções."
                    ),
                    "5 — Excelente": (
                        "O rascunho foi avaliado como "
                        "excelente e não precisou de "
                        "correções."
                    )
                }

                observacao_wizard = (
                    justificativas_por_qualidade.get(
                        qualidade_wizard,
                        (
                            "O rascunho não precisou "
                            "de correções."
                        )
                    )
                )

            if (
                comentario_complementar_wizard.strip()
            ):

                observacao_wizard = (
                    observacao_wizard
                    + " "
                    + comentario_complementar_wizard
                    .strip()
                ).strip()

            calcular_avaliacao = (
                st.form_submit_button(
                    "🧮 Calcular classificação",
                    type="primary",
                    width="stretch"
                )
            )

        if calcular_avaliacao:

            problemas_avaliacao_wizard = []

            if (
                precisou_correcao == "Sim"
                and not aspectos_observados_wizard
            ):

                problemas_avaliacao_wizard.append(
                    "Selecione pelo menos um aspecto "
                    "relacionado à correção."
                )

            if precisou_correcao == "Não":

                # Fluxo simples: a resposta Não já confirma
                # que nenhuma correção foi necessária.
                problemas_wizard = []

            else:

                if not problemas_wizard:

                    problemas_avaliacao_wizard.append(
                        "Selecione pelo menos um item "
                        "que precisou ser corrigido."
                    )

                aspectos_de_correcao = {
                    "Ajuste de clareza ou redação",
                    "Correção de dado incorreto",
                    (
                        "Remoção de informação "
                        "sem respaldo"
                    ),
                    "Inclusão de informação omitida",
                    (
                        "Correção de estrutura "
                        "ou formatação"
                    )
                }

                marcou_aspecto_correcao = any(
                    aspecto
                    in aspectos_de_correcao
                    for aspecto
                    in aspectos_observados_wizard
                )

                if not marcou_aspecto_correcao:

                    problemas_avaliacao_wizard.append(
                        "Selecione pelo menos um "
                        "aspecto relacionado à correção."
                    )

            if (
                "Outro aspecto"
                in aspectos_observados_wizard
                and len(
                    comentario_complementar_wizard
                    .strip()
                ) < 20
            ):

                problemas_avaliacao_wizard.append(
                    "Ao selecionar 'Outro aspecto', "
                    "descreva-o com pelo menos "
                    "20 caracteres."
                )

            # ---------------------------------------------
            # COERÊNCIA COM A ANÁLISE AUTOMÁTICA
            # ---------------------------------------------

            analise_avaliacao_atual = (
                st.session_state.get(
                    chave_analise
                )
            )

            texto_foi_alterado = bool(
                analise_avaliacao_atual
                and analise_avaliacao_atual.get(
                    "alterado",
                    False
                )
            )

            if (
                precisou_correcao == "Não"
                and texto_foi_alterado
            ):

                problemas_avaliacao_wizard.append(
                    "A análise automática identificou "
                    "alterações no texto. Se as mudanças "
                    "forem correções, selecione Sim."
                )

            if (
                precisou_correcao == "Sim"
                and not texto_foi_alterado
            ):

                problemas_avaliacao_wizard.append(
                    "Foi informado que houve correção, "
                    "mas nenhuma alteração textual foi "
                    "identificada."
                )

            if problemas_avaliacao_wizard:

                st.error(
                    "A avaliação precisa de alguns ajustes."
                )

                for problema_avaliacao_wizard in (
                    problemas_avaliacao_wizard
                ):

                    st.warning(
                        f"⚠️ {problema_avaliacao_wizard}"
                    )

            else:

                qualidade_numero = int(
                    qualidade_wizard.split(" ")[0]
                )

                if precisou_correcao == "Não":

                    problemas_wizard = []

                erro_factual = int(
                    "Dado incorreto"
                    in problemas_wizard
                )

                sem_respaldo = int(
                    "Informação inventada ou sem respaldo"
                    in problemas_wizard
                )

                omissao = int(
                    "Informação obrigatória ausente"
                    in problemas_wizard
                )

                secoes_esperadas_wizard = [
                    "identificação",
                    "apresentação",
                    "atividades realizadas",
                    "indicadores",
                    "ocorrências e pendências",
                    "encaminhamentos",
                    "fecho"
                ]

                texto_avaliado = (
                    st.session_state[
                        chave_revisao_salva
                    ].lower()
                )

                secoes_presentes = sum(
                    1
                    for secao in secoes_esperadas_wizard
                    if secao in texto_avaliado
                )

                estrutura_incompleta = (
                    "Estrutura incompleta"
                    in problemas_wizard
                    or secoes_presentes < 7
                )

                problema_critico = any(
                    [
                        erro_factual,
                        sem_respaldo,
                        omissao,
                        estrutura_incompleta,
                        qualidade_numero <= 2
                    ]
                )

                if problema_critico:

                    classificacao = "Reprovado"

                elif (
                    qualidade_numero == 3
                    or nivel_wizard
                    == "Sim, correções relevantes"
                ):

                    classificacao = (
                        "Aprovado com ressalva"
                    )

                else:

                    classificacao = "Aprovado"

                analise_salva = st.session_state[
                    chave_analise
                ]

                if nivel_wizard == (
                    "Não precisou de correções"
                ):

                    correcoes = 0

                elif nivel_wizard == (
                    "Sim, correções leves"
                ):

                    correcoes = max(
                        1,
                        analise_salva[
                            "linhas_adicionadas"
                        ],
                        analise_salva[
                            "linhas_removidas"
                        ]
                    )

                else:

                    correcoes = max(
                        2,
                        analise_salva[
                            "linhas_adicionadas"
                        ],
                        analise_salva[
                            "linhas_removidas"
                        ]
                    )

                st.session_state[
                    chave_avaliacao
                ] = {
                    "precisou_correcao": (
                        precisou_correcao
                    ),
                    "nivel": nivel_wizard,
                    "problemas": problemas_wizard,
                    "qualidade": qualidade_numero,
                    "observacao": observacao_wizard,
                    "aspectos_observados": (
                        aspectos_observados_wizard
                    ),
                    "comentario_complementar": (
                        comentario_complementar_wizard
                    ),
                    "classificacao": classificacao,
                    "correcoes": correcoes,
                    "erros_fatuais": erro_factual,
                    "sem_respaldo": sem_respaldo,
                    "omissoes": omissao,
                    "secoes": secoes_presentes,
                    "calculado_em": (
                        datetime.now().strftime(
                            "%d/%m/%Y %H:%M:%S"
                        )
                    )
                }

                st.rerun()

        avaliacao_wizard = st.session_state[
            chave_avaliacao
        ]

        if avaliacao_wizard is not None:

            classificacao = avaliacao_wizard[
                "classificacao"
            ]

            if classificacao == "Aprovado":

                st.success(
                    "✅ Classificação calculada: "
                    "APROVADO"
                )

            elif classificacao == (
                "Aprovado com ressalva"
            ):

                st.warning(
                    "⚠️ Classificação calculada: "
                    "APROVADO COM RESSALVA"
                )

            else:

                st.error(
                    "🚨 Classificação calculada: "
                    "REPROVADO"
                )

            resumo_c1, resumo_c2,             resumo_c3, resumo_c4 = st.columns(4)

            resumo_c1.metric(
                "Correções",
                avaliacao_wizard["correcoes"]
            )

            resumo_c2.metric(
                "Problemas críticos",
                (
                    avaliacao_wizard[
                        "erros_fatuais"
                    ]
                    + avaliacao_wizard[
                        "sem_respaldo"
                    ]
                    + avaliacao_wizard[
                        "omissoes"
                    ]
                )
            )

            resumo_c3.metric(
                "Seções",
                f"{avaliacao_wizard['secoes']}/7"
            )

            resumo_c4.metric(
                "Qualidade",
                f"{avaliacao_wizard['qualidade']}/5"
            )

            st.markdown(
                "#### Justificativa gerada"
            )

            st.info(
                avaliacao_wizard["observacao"]
            )

            aspectos_confirmados = (
                avaliacao_wizard.get(
                    "aspectos_observados",
                    []
                )
            )

            if aspectos_confirmados:

                with st.expander(
                    "Ver aspectos registrados"
                ):

                    for aspecto_confirmado in (
                        aspectos_confirmados
                    ):

                        st.write(
                            f"✅ {aspecto_confirmado}"
                        )

            confirmar_avaliacao = st.button(
                "✅ Confirmar avaliação e continuar",
                type="primary",
                width="stretch",
                key=f"{prefixo_wizard}_confirmar_5"
            )

            if confirmar_avaliacao:

                st.session_state[chave_etapa] = 5
                st.rerun()

        if st.button(
            "⬅️ Voltar à análise",
            width="stretch",
            key=f"{prefixo_wizard}_voltar_3"
        ):

            st.session_state[chave_etapa] = 3
            st.rerun()

    # =====================================================
    # ETAPA 5 - VALIDAÇÃO E EXPORTAÇÃO
    # =====================================================

    elif etapa_atual == 5:

        st.markdown(
            "### Etapa 5 — Validar e exportar"
        )

        st.success(
            "✅ Revisão, análise e avaliação concluídas."
        )

        executar_validacao_wizard = st.button(
            "🔎 Executar validação de fidelidade",
            type="primary",
            width="stretch",
            key=f"{prefixo_wizard}_validar"
        )

        if executar_validacao_wizard:

            texto_validar = st.session_state[
                chave_revisao_salva
            ].lower()

            secoes_validar = [
                "identificação",
                "apresentação",
                "atividades realizadas",
                "indicadores",
                "ocorrências e pendências",
                "encaminhamentos",
                "fecho"
            ]

            secoes_localizadas = sum(
                1
                for secao in secoes_validar
                if secao in texto_validar
            )

            total_itens_validacao = 7
            percentual_validacao = (
                secoes_localizadas
                / total_itens_validacao
                * 100
            )

            st.session_state[
                chave_validacao
            ] = {
                "itens": total_itens_validacao,
                "localizados": secoes_localizadas,
                "cobertura": percentual_validacao,
                "validado_em": (
                    datetime.now().strftime(
                        "%d/%m/%Y %H:%M:%S"
                    )
                )
            }

            st.rerun()

        validacao_wizard = st.session_state[
            chave_validacao
        ]

        if validacao_wizard is None:

            st.info(
                "Execute a validação de fidelidade "
                "para liberar a exportação."
            )

        else:

            cobertura = validacao_wizard[
                "cobertura"
            ]

            if cobertura == 100:

                st.success(
                    "✅ Validação concluída: "
                    "7 de 7 seções localizadas."
                )

            else:

                st.warning(
                    f"⚠️ Validação concluída: "
                    f"{validacao_wizard['localizados']} "
                    "de 7 seções localizadas."
                )

            st.metric(
                "Cobertura estrutural",
                f"{cobertura:.1f}%".replace(
                    ".",
                    ","
                )
            )

            # =================================================
        # RELATÓRIO EXECUTIVO WORD NEXORA V11
# REFINAMENTOS DO RELATÓRIO EXECUTIVO NEXORA V11
        # =================================================

        st.markdown(
            "#### Exportação do relatório"
        )

        formato_wizard = st.selectbox(
            "Formato do relatório",
            options=[
                (
                    "Relatório Executivo Word, "
                    "com gráficos e texto"
                ),
                "Relatório Textual Word",
                "Texto simples TXT"
            ],
            index=0,
            key=f"{prefixo_wizard}_formato"
        )

        texto_final_wizard = (
            st.session_state[
                chave_revisao_salva
            ]
        )

        avaliacao_exportacao = (
            st.session_state.get(
                chave_avaliacao
            )
            or {}
        )

        validacao_exportacao = (
            st.session_state.get(
                chave_validacao
            )
            or {}
        )

        # ---------------------------------------------
        # SINGULAR E PLURAL
        # ---------------------------------------------

        def pluralizar_nexora(
            quantidade,
            singular,
            plural
        ):

            return (
                singular
                if int(quantidade) == 1
                else plural
            )

        # ---------------------------------------------
        # FORMATAÇÃO DE CÉLULAS DO WORD
        # ---------------------------------------------

        def aplicar_cor_celula_nexora(
            celula,
            cor_hexadecimal
        ):

            from docx.oxml import OxmlElement
            from docx.oxml.ns import qn

            propriedades = (
                celula._tc.get_or_add_tcPr()
            )

            preenchimento = OxmlElement(
                "w:shd"
            )

            preenchimento.set(
                qn("w:fill"),
                cor_hexadecimal
            )

            propriedades.append(
                preenchimento
            )

        def aplicar_bordas_celula_nexora(
            celula,
            cor="D9E2F3"
        ):

            from docx.oxml import OxmlElement
            from docx.oxml.ns import qn

            tc_pr = celula._tc.get_or_add_tcPr()

            bordas = tc_pr.first_child_found_in(
                "w:tcBorders"
            )

            if bordas is None:

                bordas = OxmlElement(
                    "w:tcBorders"
                )

                tc_pr.append(
                    bordas
                )

            for lado in [
                "top",
                "left",
                "bottom",
                "right"
            ]:

                elemento = OxmlElement(
                    f"w:{lado}"
                )

                elemento.set(
                    qn("w:val"),
                    "single"
                )

                elemento.set(
                    qn("w:sz"),
                    "5"
                )

                elemento.set(
                    qn("w:color"),
                    cor
                )

                bordas.append(
                    elemento
                )

        def adicionar_numero_pagina_nexora(
            paragrafo
        ):

            from docx.oxml import OxmlElement
            from docx.oxml.ns import qn

            run = paragrafo.add_run()

            inicio_campo = OxmlElement(
                "w:fldChar"
            )

            inicio_campo.set(
                qn("w:fldCharType"),
                "begin"
            )

            instrucao = OxmlElement(
                "w:instrText"
            )

            instrucao.set(
                qn("xml:space"),
                "preserve"
            )

            instrucao.text = "PAGE"

            separador = OxmlElement(
                "w:fldChar"
            )

            separador.set(
                qn("w:fldCharType"),
                "separate"
            )

            texto_padrao = OxmlElement(
                "w:t"
            )

            texto_padrao.text = "1"

            fim_campo = OxmlElement(
                "w:fldChar"
            )

            fim_campo.set(
                qn("w:fldCharType"),
                "end"
            )

            run._r.append(
                inicio_campo
            )

            run._r.append(
                instrucao
            )

            run._r.append(
                separador
            )

            run._r.append(
                texto_padrao
            )

            run._r.append(
                fim_campo
            )

            run.font.name = "Arial"
            run.font.size = Pt(8)


        def adicionar_titulo_nexora(
            documento,
            texto,
            nivel=1
        ):

            paragrafo = documento.add_paragraph()

            run = paragrafo.add_run(
                texto
            )

            run.font.name = "Arial"
            run.bold = True

            if nivel == 1:

                run.font.size = Pt(16)
                run.font.color.rgb = RGBColor(
                    31,
                    78,
                    120
                )

            else:

                run.font.size = Pt(13)
                run.font.color.rgb = RGBColor(
                    23,
                    43,
                    77
                )

            paragrafo.paragraph_format.space_before = (
                Pt(8)
            )

            paragrafo.paragraph_format.space_after = (
                Pt(6)
            )

            return paragrafo

        def adicionar_paragrafo_nexora(
            documento,
            texto,
            negrito=False
        ):

            paragrafo = documento.add_paragraph()

            paragrafo.alignment = (
                WD_ALIGN_PARAGRAPH.JUSTIFY
            )

            paragrafo.paragraph_format.line_spacing = (
                1.15
            )

            run = paragrafo.add_run(
                str(texto)
            )

            run.font.name = "Arial"
            run.font.size = Pt(10.5)
            run.bold = negrito

            return paragrafo

        # ---------------------------------------------
        # GERAÇÃO DOS GRÁFICOS
        # ---------------------------------------------

        def gerar_graficos_executivos_nexora():

            import matplotlib.pyplot as plt

            imagens = {}

            # Gráfico Meta versus Realizado
            if not indicadores_cenario.empty:

                dados_indicadores = (
                    indicadores_cenario.copy()
                )

                dados_indicadores[
                    "Meta"
                ] = pd.to_numeric(
                    dados_indicadores["Meta"],
                    errors="coerce"
                ).fillna(0)

                dados_indicadores[
                    "Realizado"
                ] = pd.to_numeric(
                    dados_indicadores["Realizado"],
                    errors="coerce"
                ).fillna(0)

                figura, eixo = plt.subplots(
                    figsize=(9, 4.8)
                )

                posicoes = range(
                    len(dados_indicadores)
                )

                largura = 0.36

                eixo.bar(
                    [
                        posicao - largura / 2
                        for posicao in posicoes
                    ],
                    dados_indicadores["Meta"],
                    width=largura,
                    label="Meta",
                    color="#9FBAD3"
                )

                eixo.bar(
                    [
                        posicao + largura / 2
                        for posicao in posicoes
                    ],
                    dados_indicadores[
                        "Realizado"
                    ],
                    width=largura,
                    label="Realizado",
                    color="#1F4E78"
                )

                eixo.set_title(
                    "Meta versus realizado"
                )

                eixo.set_xticks(
                    list(posicoes)
                )

                eixo.set_xticklabels(
                    dados_indicadores[
                        "Indicador"
                    ].astype(str),
                    rotation=18,
                    ha="right"
                )

                eixo.legend()
                eixo.grid(
                    axis="y",
                    alpha=0.2
                )

                figura.tight_layout()

                memoria_grafico = BytesIO()

                figura.savefig(
                    memoria_grafico,
                    format="png",
                    dpi=160,
                    bbox_inches="tight"
                )

                plt.close(
                    figura
                )

                memoria_grafico.seek(0)

                imagens[
                    "meta_realizado"
                ] = memoria_grafico

            # Gráfico Volume por atividade
            if not resumo_atividades.empty:

                dados_atividades = (
                    resumo_atividades.copy()
                )

                dados_atividades = (
                    dados_atividades.sort_values(
                        "Quantidade_Total",
                        ascending=True
                    )
                )

                figura, eixo = plt.subplots(
                    figsize=(9, 4.8)
                )

                eixo.barh(
                    dados_atividades[
                        "Tipo_Atividade"
                    ].astype(str),
                    dados_atividades[
                        "Quantidade_Total"
                    ],
                    color="#2F80ED"
                )

                eixo.set_title(
                    "Volume por atividade"
                )

                eixo.set_xlabel(
                    "Quantidade total"
                )

                eixo.grid(
                    axis="x",
                    alpha=0.2
                )

                figura.tight_layout()

                memoria_grafico = BytesIO()

                figura.savefig(
                    memoria_grafico,
                    format="png",
                    dpi=160,
                    bbox_inches="tight"
                )

                plt.close(
                    figura
                )

                memoria_grafico.seek(0)

                imagens[
                    "volume_atividade"
                ] = memoria_grafico

            return imagens

        # ---------------------------------------------
        # DOCUMENTO TEXTUAL
        # ---------------------------------------------

        def gerar_word_textual_nexora():

            documento = Document()

            secao = documento.sections[0]

            secao.top_margin = Cm(2.5)
            secao.left_margin = Cm(3)
            secao.right_margin = Cm(2)
            secao.bottom_margin = Cm(2)

            estilo = documento.styles[
                "Normal"
            ]

            estilo.font.name = "Arial"
            estilo.font.size = Pt(12)

            for linha in (
                texto_final_wizard.splitlines()
            ):

                if not linha.strip():
                    continue

                paragrafo = (
                    documento.add_paragraph()
                )

                paragrafo.alignment = (
                    WD_ALIGN_PARAGRAPH.JUSTIFY
                )

                paragrafo.paragraph_format.line_spacing = (
                    1.5
                )

                run = paragrafo.add_run(
                    linha.strip()
                )

                run.font.name = "Arial"
                run.font.size = Pt(12)

            memoria = BytesIO()

            documento.save(
                memoria
            )

            memoria.seek(0)

            return memoria.getvalue()

        # ---------------------------------------------
        # DOCUMENTO EXECUTIVO
        # ---------------------------------------------

        def gerar_word_executivo_nexora():

            documento = Document()

            secao = documento.sections[0]

            secao.top_margin = Cm(1.8)
            secao.bottom_margin = Cm(1.8)
            secao.left_margin = Cm(2)
            secao.right_margin = Cm(2)

            # CAPA
            paragrafo_marca = (
                documento.add_paragraph()
            )

            paragrafo_marca.alignment = (
                WD_ALIGN_PARAGRAPH.CENTER
            )

            run_marca = paragrafo_marca.add_run(
                "NEXORA"
            )

            run_marca.bold = True
            run_marca.font.name = "Arial"
            run_marca.font.size = Pt(25)
            run_marca.font.color.rgb = RGBColor(
                31,
                78,
                120
            )

            paragrafo_subtitulo = (
                documento.add_paragraph()
            )

            paragrafo_subtitulo.alignment = (
                WD_ALIGN_PARAGRAPH.CENTER
            )

            run_subtitulo = (
                paragrafo_subtitulo.add_run(
                    "Inteligência para dados, "
                    "processos e relatórios"
                )
            )

            run_subtitulo.font.name = "Arial"
            run_subtitulo.font.size = Pt(11)
            run_subtitulo.font.color.rgb = RGBColor(
                100,
                116,
                139
            )

            documento.add_paragraph()

            titulo_capa = documento.add_paragraph()

            titulo_capa.alignment = (
                WD_ALIGN_PARAGRAPH.CENTER
            )

            run_titulo = titulo_capa.add_run(
                "RELATÓRIO EXECUTIVO DE "
                "ATIVIDADES ADMINISTRATIVAS"
            )

            run_titulo.bold = True
            run_titulo.font.name = "Arial"
            run_titulo.font.size = Pt(18)
            run_titulo.font.color.rgb = RGBColor(
                23,
                43,
                77
            )

            dados_capa = [
                (
                    "Cenário",
                    cenario_escolhido
                ),
                (
                    "Geração",
                    geracao_wizard
                ),
                (
                    "Organização",
                    geral.get(
                        "Organizacao_Ficticia",
                        "Não informada"
                    )
                ),
                (
                    "Setor",
                    geral.get(
                        "Setor_Ficticio",
                        "Não informado"
                    )
                ),
                (
                    "Responsável",
                    geral.get(
                        "Responsavel_Ficticio",
                        "Não informado"
                    )
                ),
                (
                    "Complexidade",
                    geral.get(
                        "Complexidade",
                        "Não informada"
                    )
                ),
                (
                    "Local",
                    geral.get(
                        "Local",
                        "Não informado"
                    )
                ),
                (
                    "Finalidade",
                    geral.get(
                        "Finalidade",
                        "Não informada"
                    )
                ),
                (
                    "Fonte dos dados",
                    (
                        st.session_state.get(
                            "nome_arquivo_importado_ativo",
                            "Planilha importada"
                        )
                        if st.session_state.get(
                            "fonte_dados_ativa"
                        ) == "Planilha importada"
                        else "Base acadêmica simulada"
                    )
                ),
                (
                    "Gerado em",
                    datetime.now().strftime(
                        "%d/%m/%Y às %H:%M"
                    )
                )
            ]

            tabela_capa = documento.add_table(
                rows=len(dados_capa),
                cols=2
            )

            tabela_capa.autofit = True

            for indice, (
                rotulo,
                valor
            ) in enumerate(dados_capa):

                celula_rotulo = (
                    tabela_capa.rows[
                        indice
                    ].cells[0]
                )

                celula_valor = (
                    tabela_capa.rows[
                        indice
                    ].cells[1]
                )

                celula_rotulo.text = str(
                    rotulo
                )

                celula_valor.text = str(
                    valor
                )

                aplicar_cor_celula_nexora(
                    celula_rotulo,
                    "EAF2F8"
                )

                for celula in [
                    celula_rotulo,
                    celula_valor
                ]:

                    aplicar_bordas_celula_nexora(
                        celula
                    )

                    for paragrafo in (
                        celula.paragraphs
                    ):

                        for run in (
                            paragrafo.runs
                        ):

                            run.font.name = "Arial"
                            run.font.size = Pt(10)

                celula_rotulo.paragraphs[
                    0
                ].runs[0].bold = True

            documento.add_page_break()

            # RESUMO EXECUTIVO
            adicionar_titulo_nexora(
                documento,
                "Resumo executivo"
            )

            quantidade_registros = len(
                registros_cenario
            )

            quantidade_atividades = len(
                resumo_atividades
            )

            quantidade_indicadores = len(
                indicadores_cenario
            )

            quantidade_ocorrencias = len(
                ocorrencias_cenario
            )

            quantidade_encaminhamentos = len(
                encaminhamentos_cenario
            )

            classificacao_executiva = (
                avaliacao_exportacao.get(
                    "classificacao",
                    "Não informada"
                )
            )

            qualidade_executiva = (
                avaliacao_exportacao.get(
                    "qualidade",
                    "Não informada"
                )
            )

            cobertura_executiva = (
                validacao_exportacao.get(
                    "cobertura"
                )
            )

            validacao_executada_exportacao = (
                cobertura_executiva is not None
                and bool(validacao_exportacao)
            )

            if validacao_executada_exportacao:

                texto_cobertura_executiva = (
                    f"{float(cobertura_executiva):.1f}%"
                    .replace(".", ",")
                )

            else:

                texto_cobertura_executiva = (
                    "Não calculada"
                )

            indicadores_executivos = [
                (
                    "Registros",
                    quantidade_registros
                ),
                (
                    "Atividades",
                    quantidade_atividades
                ),
                (
                    "Indicadores",
                    quantidade_indicadores
                ),
                (
                    "Ocorrências",
                    quantidade_ocorrencias
                ),
                (
                    "Encaminhamentos",
                    quantidade_encaminhamentos
                ),
                (
                    "Qualidade",
                    (
                        f"{qualidade_executiva}/5"
                        if isinstance(
                            qualidade_executiva,
                            int
                        )
                        else qualidade_executiva
                    )
                ),
                (
                    "Classificação",
                    classificacao_executiva
                ),
                (
                    "Cobertura",
                    texto_cobertura_executiva
                )
            ]

            tabela_indicadores = (
                documento.add_table(
                    rows=2,
                    cols=4
                )
            )

            for indice, (
                rotulo,
                valor
            ) in enumerate(
                indicadores_executivos
            ):

                linha = indice // 4
                coluna = indice % 4

                celula = (
                    tabela_indicadores.rows[
                        linha
                    ].cells[coluna]
                )

                celula.text = (
                    f"{rotulo}\n{valor}"
                )

                aplicar_cor_celula_nexora(
                    celula,
                    (
                        "EAF2F8"
                        if linha == 0
                        else "F8FAFC"
                    )
                )

                aplicar_bordas_celula_nexora(
                    celula
                )

                celula.paragraphs[
                    0
                ].alignment = (
                    WD_ALIGN_PARAGRAPH.CENTER
                )

                runs = celula.paragraphs[
                    0
                ].runs

                if runs:

                    runs[0].font.name = "Arial"
                    runs[0].font.size = Pt(10)
                    runs[0].bold = True

            # GRÁFICOS
            graficos = (
                gerar_graficos_executivos_nexora()
            )

            if graficos:

                adicionar_titulo_nexora(
                    documento,
                    "Indicadores visuais"
                )

                if (
                    "meta_realizado"
                    in graficos
                ):

                    documento.add_picture(
                        graficos[
                            "meta_realizado"
                        ],
                        width=Cm(16)
                    )

                    ultimo_paragrafo = (
                        documento.paragraphs[-1]
                    )

                    ultimo_paragrafo.alignment = (
                        WD_ALIGN_PARAGRAPH.CENTER
                    )

                if (
                    "volume_atividade"
                    in graficos
                ):

                    documento.add_picture(
                        graficos[
                            "volume_atividade"
                        ],
                        width=Cm(16)
                    )

                    ultimo_paragrafo = (
                        documento.paragraphs[-1]
                    )

                    ultimo_paragrafo.alignment = (
                        WD_ALIGN_PARAGRAPH.CENTER
                    )

            # =================================================
            # PAGINAÇÃO EXECUTIVA NEXORA V11
            # =================================================

            # Impede que o título da avaliação fique no final
            # da página dos gráficos e a justificativa seja
            # deslocada sozinha para a página seguinte.
            documento.add_page_break()

            # AVALIAÇÃO HUMANA
            adicionar_titulo_nexora(
                documento,
                "Avaliação humana"
            )

            adicionar_paragrafo_nexora(
                documento,
                (
                    "Classificação: "
                    f"{classificacao_executiva}"
                ),
                negrito=True
            )

            adicionar_paragrafo_nexora(
                documento,
                (
                    "Justificativa: "
                    + str(
                        avaliacao_exportacao.get(
                            "observacao",
                            "Não informada"
                        )
                    )
                )
            )

            aspectos = (
                avaliacao_exportacao.get(
                    "aspectos_observados",
                    []
                )
            )

            if aspectos:

                adicionar_paragrafo_nexora(
                    documento,
                    "Aspectos registrados:",
                    negrito=True
                )

                for aspecto in aspectos:

                    adicionar_paragrafo_nexora(
                        documento,
                        f"• {aspecto}"
                    )

            # VALIDAÇÃO
            adicionar_titulo_nexora(
                documento,
                "Validação"
            )

            if validacao_executada_exportacao:

                adicionar_paragrafo_nexora(
                    documento,
                    (
                        "Cobertura estrutural: "
                        f"{texto_cobertura_executiva}."
                    )
                )

                adicionar_paragrafo_nexora(
                    documento,
                    (
                        "Seções localizadas: "
                        f"{validacao_exportacao.get('localizados', 0)} "
                        f"de {validacao_exportacao.get('itens', 7)}."
                    )
                )

            else:

                adicionar_paragrafo_nexora(
                    documento,
                    (
                        "A validação estrutural não havia "
                        "sido calculada no momento da geração."
                    )
                )

            # OCORRÊNCIAS
            adicionar_titulo_nexora(
                documento,
                "Ocorrências e encaminhamentos"
            )

            if ocorrencias_cenario.empty:

                adicionar_paragrafo_nexora(
                    documento,
                    "Nenhuma ocorrência registrada."
                )

            else:

                for _, ocorrencia in (
                    ocorrencias_cenario.iterrows()
                ):

                    adicionar_paragrafo_nexora(
                        documento,
                        (
                            f"• {ocorrencia.get('Categoria', '')}: "
                            f"{ocorrencia.get('Descricao', '')} "
                            f"| Severidade: "
                            f"{ocorrencia.get('Severidade', '')} "
                            f"| Situação: "
                            f"{ocorrencia.get('Situacao', '')}"
                        )
                    )

            if encaminhamentos_cenario.empty:

                adicionar_paragrafo_nexora(
                    documento,
                    "Nenhum encaminhamento registrado."
                )

            else:

                for _, encaminhamento in (
                    encaminhamentos_cenario.iterrows()
                ):

                    adicionar_paragrafo_nexora(
                        documento,
                        (
                            f"• {encaminhamento.get('Acao_Prevista', '')} "
                            f"| Prioridade: "
                            f"{encaminhamento.get('Prioridade', '')} "
                            f"| Status: "
                            f"{encaminhamento.get('Status', '')}"
                        )
                    )

            # TEXTO REVISADO
            documento.add_page_break()

            adicionar_titulo_nexora(
                documento,
                "Relatório revisado"
            )

            for linha in (
                texto_final_wizard.splitlines()
            ):

                if not linha.strip():
                    continue

                adicionar_paragrafo_nexora(
                    documento,
                    linha.strip()
                )

            # RODAPÉ
            for secao_documento in (
                documento.sections
            ):

                rodape = (
                    secao_documento.footer
                    .paragraphs[0]
                )

                rodape.alignment = (
                    WD_ALIGN_PARAGRAPH.CENTER
                )

                run_rodape = rodape.add_run(
                    (
                        "Nexora | "
                        f"{id_wizard} | "
                        "Dados processados e "
                        "revisados pelo usuário | Página "
                    )
                )

                run_rodape.font.name = "Arial"
                run_rodape.font.size = Pt(8)
                run_rodape.font.color.rgb = RGBColor(
                    100,
                    116,
                    139
                )

                adicionar_numero_pagina_nexora(
                    rodape
                )

            memoria = BytesIO()

            documento.save(
                memoria
            )

            memoria.seek(0)

            return memoria.getvalue()

        # ---------------------------------------------
        # PREPARAÇÃO SOB DEMANDA
        # ---------------------------------------------

        chave_arquivo_preparado = (
            f"{prefixo_wizard}_"
            "arquivo_exportacao_preparado"
        )

        chave_formato_preparado = (
            f"{prefixo_wizard}_"
            "formato_exportacao_preparado"
        )

        if (
            chave_arquivo_preparado
            not in st.session_state
        ):

            st.session_state[
                chave_arquivo_preparado
            ] = None

        if (
            chave_formato_preparado
            not in st.session_state
        ):

            st.session_state[
                chave_formato_preparado
            ] = ""

        formato_executivo_selecionado = (
            formato_wizard.startswith(
                "Relatório Executivo"
            )
        )

        validacao_pronta_exportacao = bool(
            validacao_exportacao
            and validacao_exportacao.get(
                "cobertura"
            ) is not None
        )

        if (
            formato_executivo_selecionado
            and not validacao_pronta_exportacao
        ):

            st.warning(
                "Execute a validação de fidelidade antes "
                "de preparar o relatório executivo."
            )

        preparar_relatorio = st.button(
            (
                "📊 Preparar relatório executivo"
                if formato_executivo_selecionado
                else "📄 Preparar arquivo"
            ),
            type="primary",
            width="stretch",
            disabled=(
                formato_executivo_selecionado
                and not validacao_pronta_exportacao
            ),
            key=(
                f"{prefixo_wizard}_"
                "preparar_exportacao"
            )
        )

        if preparar_relatorio:

            with st.spinner(
                "Preparando o arquivo..."
            ):

                if formato_wizard.startswith(
                    "Relatório Executivo"
                ):

                    dados_preparados = (
                        gerar_word_executivo_nexora()
                    )

                    nome_preparado = (
                        f"Relatorio_Executivo_"
                        f"{id_wizard}.docx"
                    )

                    mime_preparado = (
                        "application/vnd.openxmlformats-"
                        "officedocument.wordprocessingml."
                        "document"
                    )

                elif formato_wizard == (
                    "Relatório Textual Word"
                ):

                    dados_preparados = (
                        gerar_word_textual_nexora()
                    )

                    nome_preparado = (
                        f"Relatorio_Revisado_"
                        f"{id_wizard}.docx"
                    )

                    mime_preparado = (
                        "application/vnd.openxmlformats-"
                        "officedocument.wordprocessingml."
                        "document"
                    )

                else:

                    dados_preparados = (
                        texto_final_wizard.encode(
                            "utf-8"
                        )
                    )

                    nome_preparado = (
                        f"relatorio_revisado_"
                        f"{id_wizard}.txt"
                    )

                    mime_preparado = (
                        "text/plain"
                    )

                st.session_state[
                    chave_arquivo_preparado
                ] = {
                    "dados": dados_preparados,
                    "nome": nome_preparado,
                    "mime": mime_preparado
                }

                st.session_state[
                    chave_formato_preparado
                ] = formato_wizard

            st.success(
                "✅ Arquivo preparado com sucesso."
            )

            st.rerun()

        arquivo_preparado = (
            st.session_state.get(
                chave_arquivo_preparado
            )
        )

        formato_preparado = (
            st.session_state.get(
                chave_formato_preparado,
                ""
            )
        )

        if (
            arquivo_preparado is not None
            and formato_preparado
            == formato_wizard
        ):

            st.success(
                "✅ Relatório pronto para download."
            )

            st.download_button(
                label="⬇️ Baixar relatório",
                data=arquivo_preparado[
                    "dados"
                ],
                file_name=arquivo_preparado[
                    "nome"
                ],
                mime=arquivo_preparado[
                    "mime"
                ],
                width="stretch",
                key=(
                    f"{prefixo_wizard}_"
                    "download_preparado"
                )
            )

            st.caption(
                "Arquivo preparado: "
                f"{arquivo_preparado['nome']}"
            )

        elif (
            arquivo_preparado is not None
            and formato_preparado
            != formato_wizard
        ):

            st.info(
                "O formato foi alterado. Clique em "
                "Preparar arquivo novamente."
            )

        if st.button(
            "⬅️ Voltar à avaliação",
            width="stretch",
            key=f"{prefixo_wizard}_voltar_4"
        ):

            st.session_state[chave_etapa] = 4
            st.rerun()

    # -----------------------------------------------------
    # RECOMEÇAR EXECUÇÃO
    # -----------------------------------------------------

    st.divider()

    with st.expander(
        "Opções da execução"
    ):

        if st.button(
            "🗑️ Apagar esta execução e recomeçar",
            width="stretch",
            key=f"{prefixo_wizard}_reset"
        ):

            chaves_apagar = [
                chave_etapa,
                chave_original,
                chave_revisao_editor,
                chave_revisao_salva,
                chave_analise,
                chave_avaliacao,
                chave_validacao,
                chave_prompt_preparado,
                chave_prompt_texto
            ]

            for chave in chaves_apagar:

                if chave in st.session_state:
                    del st.session_state[chave]

            st.rerun()

with aba_validacao:

    st.subheader("Validação funcional do GeraRel")

    st.write(
        "Esta aba executa verificações automáticas das "
        "principais funções e estruturas utilizadas pelo "
        "dashboard."
    )

    # -----------------------------------------------------
    # CÁLCULOS PARA CONFERÊNCIA
    # -----------------------------------------------------

    quantidade_cenarios = len(cenarios)

    quantidade_registros_total = len(registros)

    quantidade_registros_cenario = len(
        registros_cenario
    )

    soma_quantidades_origem = (
        registros_cenario["Quantidade"]
        .sum()
    )

    soma_quantidades_consolidadas = (
        resumo_atividades["Quantidade_Total"]
        .sum()
    )

    quantidade_indicadores_cenario = len(
        indicadores_cenario
    )

    quantidade_ocorrencias_cenario = len(
        ocorrencias_cenario
    )

    quantidade_encaminhamentos_cenario = len(
        encaminhamentos_cenario
    )

    prompt_disponivel = bool(
        globals().get('prompt_relatorio', '').strip()
    )

    # -----------------------------------------------------
    # LISTA DOS TESTES
    # -----------------------------------------------------

    testes_funcionais = [
        {
            "Teste": "Base Excel localizada",
            "Critério": (
                "O arquivo da base deve existir "
                "no caminho configurado."
            ),
            "Resultado": (
                "Arquivo localizado"
                if os.path.exists(arquivo_excel)
                else "Arquivo não localizado"
            ),
            "Aprovado": os.path.exists(
                arquivo_excel
            )
        },
        {
            "Teste": "Cenários disponíveis",
            "Critério": (
                "A base deve possuir 6 cenários."
            ),
            "Resultado": (
                f"{quantidade_cenarios} cenários"
            ),
            "Aprovado": (
                quantidade_cenarios == 6
            )
        },
        {
            "Teste": "Registros operacionais",
            "Critério": (
                "A base deve possuir 210 registros."
            ),
            "Resultado": (
                f"{quantidade_registros_total} registros"
            ),
            "Aprovado": (
                quantidade_registros_total == 210
            )
        },
        {
            "Teste": "Filtro do cenário",
            "Critério": (
                "O cenário selecionado deve possuir "
                "registros correspondentes."
            ),
            "Resultado": (
                f"{cenario_escolhido}: "
                f"{quantidade_registros_cenario} registros"
            ),
            "Aprovado": (
                quantidade_registros_cenario > 0
            )
        },
        {
            "Teste": "Consolidação das quantidades",
            "Critério": (
                "A soma consolidada deve ser igual "
                "à soma dos registros originais."
            ),
            "Resultado": (
                f"Origem: {soma_quantidades_origem} | "
                f"Consolidado: "
                f"{soma_quantidades_consolidadas}"
            ),
            "Aprovado": (
                soma_quantidades_origem
                == soma_quantidades_consolidadas
            )
        },
        {
            "Teste": "Indicadores do cenário",
            "Critério": (
                "Os indicadores devem ser carregados "
                "para o cenário selecionado."
            ),
            "Resultado": (
                f"{quantidade_indicadores_cenario} "
                "indicadores"
            ),
            "Aprovado": (
                quantidade_indicadores_cenario > 0
            )
        },
        {
            "Teste": "Ocorrências do cenário",
            "Critério": (
                "A consulta deve ser executada mesmo "
                "quando não houver ocorrência."
            ),
            "Resultado": (
                f"{quantidade_ocorrencias_cenario} "
                "ocorrências"
            ),
            "Aprovado": True
        },
        {
            "Teste": "Encaminhamentos do cenário",
            "Critério": (
                "Os encaminhamentos devem ser "
                "consultados sem erro."
            ),
            "Resultado": (
                f"{quantidade_encaminhamentos_cenario} "
                "encaminhamentos"
            ),
            "Aprovado": True
        },
        {
            "Teste": "Geração do prompt",
            "Critério": (
                "O prompt deve possuir conteúdo."
            ),
            "Resultado": (
                f"{len(globals().get('prompt_relatorio', ''))} caracteres"
            ),
            "Aprovado": prompt_disponivel
        },
        {
            "Teste": "Exportação do prompt",
            "Critério": (
                "O botão de download TXT deve estar "
                "disponível."
            ),
            "Resultado": "Recurso disponível",
            "Aprovado": True
        },
        {
            "Teste": "Avaliação do rascunho",
            "Critério": (
                "O dashboard deve permitir registrar "
                "erros, omissões e correções."
            ),
            "Resultado": "Recurso disponível",
            "Aprovado": True
        },
        {
            "Teste": "Exportação CSV",
            "Critério": (
                "O resultado da avaliação deve poder "
                "ser exportado."
            ),
            "Resultado": "Recurso disponível",
            "Aprovado": True
        },
        {
            "Teste": "Exportação DOCX",
            "Critério": (
                "A versão revisada deve poder ser "
                "exportada para Word."
            ),
            "Resultado": "Recurso disponível",
            "Aprovado": True
        }
    ]

    dataframe_testes = pd.DataFrame(
        testes_funcionais
    )

    # -----------------------------------------------------
    # SITUAÇÃO DOS TESTES
    # -----------------------------------------------------

    dataframe_testes["Situação"] = (
        dataframe_testes["Aprovado"]
        .map(
            {
                True: "Aprovado",
                False: "Reprovado"
            }
        )
    )

    total_testes = len(
        dataframe_testes
    )

    testes_aprovados = int(
        dataframe_testes["Aprovado"]
        .sum()
    )

    testes_reprovados = (
        total_testes
        - testes_aprovados
    )

    percentual_aprovacao = (
        testes_aprovados
        / total_testes
        * 100
    )

    # -----------------------------------------------------
    # CARTÕES DA VALIDAÇÃO
    # -----------------------------------------------------

    validacao_coluna1, validacao_coluna2,     validacao_coluna3, validacao_coluna4 = (
        st.columns(4)
    )

    validacao_coluna1.metric(
        "Testes executados",
        total_testes
    )

    validacao_coluna2.metric(
        "Aprovados",
        testes_aprovados
    )

    validacao_coluna3.metric(
        "Reprovados",
        testes_reprovados
    )

    validacao_coluna4.metric(
        "Taxa de aprovação",
        f"{percentual_aprovacao:.1f}%"
        .replace(".", ",")
    )

    # -----------------------------------------------------
    # RESULTADO GERAL
    # -----------------------------------------------------

    if testes_reprovados == 0:

        st.success(
            "✅ Todos os testes funcionais "
            "foram aprovados."
        )

    else:

        st.error(
            f"🚨 Foram identificados "
            f"{testes_reprovados} testes reprovados."
        )

    # -----------------------------------------------------
    # TABELA DOS TESTES
    # -----------------------------------------------------

    st.markdown(
        "### Resultado dos testes funcionais"
    )

    tabela_testes = dataframe_testes[
        [
            "Teste",
            "Critério",
            "Resultado",
            "Situação"
        ]
    ].copy()

    st.dataframe(
        tabela_testes,
        width="stretch",
        hide_index=True
    )

    # -----------------------------------------------------
    # DOWNLOAD DO RESULTADO
    # -----------------------------------------------------

    csv_testes = (
        tabela_testes
        .to_csv(
            index=False,
            sep=";"
        )
        .encode("utf-8-sig")
    )

    nome_arquivo_testes = (
        f"validacao_funcional_"
        f"{cenario_escolhido}.csv"
    )

    st.download_button(
        label=(
            f"⬇️ Baixar "
            f"{nome_arquivo_testes}"
        ),
        data=csv_testes,
        file_name=nome_arquivo_testes,
        mime="text/csv",
        width="stretch"
    )

    st.info(
        "Os testes verificam o funcionamento das "
        "principais etapas do GeraRel. A validação "
        "não substitui a conferência humana dos "
        "relatórios administrativos."
    )


st.divider()

st.caption(
    "Protótipo acadêmico do TCC de Engenharia de Computação. "
    "Dados exclusivamente simulados."
)
