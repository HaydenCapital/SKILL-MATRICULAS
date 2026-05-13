import streamlit as st
import pandas as pd
import sys
import os
import tempfile
import time
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "app"))

from modulo1_leitura import carregar_imoveis
from modulo2_cartorios import carregar_cartorios
from modulo3_match import cruzar
from modulo4_email import _obter_token_delegado, _graph_enviar, _assunto, _corpo_html, DELAY_ENTRE_ENVIOS

# ── Config ──────────────────────────────────────────────────────
st.set_page_config(
    page_title="Hayden Capital — Consulta de Matrículas",
    page_icon="🏛️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Identidade visual Hayden Capital ────────────────────────────
st.markdown("""
<style>
    /* Fundo geral */
    .stApp { background-color: #0d1117; }
    [data-testid="stAppViewContainer"] { background-color: #0d1117; }
    [data-testid="stHeader"] { background-color: #0d1117; }

    /* Sidebar */
    [data-testid="stSidebar"] { background-color: #111827; }

    /* Tipografia */
    html, body, [class*="css"] {
        font-family: 'Inter', 'Segoe UI', sans-serif;
        color: #e5e7eb;
    }

    /* Header Hayden */
    .hayden-header {
        background: linear-gradient(135deg, #111827 0%, #1a2540 100%);
        border-bottom: 2px solid #c9a84c;
        padding: 1.5rem 2rem;
        border-radius: 8px;
        margin-bottom: 1.5rem;
        display: flex;
        align-items: center;
        gap: 1.5rem;
    }
    .hayden-logo {
        font-size: 1.8rem;
        font-weight: 800;
        color: #c9a84c;
        letter-spacing: 1px;
    }
    .hayden-subtitle {
        color: #9ca3af;
        font-size: 0.9rem;
        margin-top: 2px;
    }

    /* Cards de métrica */
    [data-testid="metric-container"] {
        background: #111827;
        border: 1px solid #1f2937;
        border-radius: 8px;
        padding: 1rem;
        border-left: 3px solid #c9a84c;
    }
    [data-testid="metric-container"] label { color: #9ca3af !important; }
    [data-testid="metric-container"] [data-testid="stMetricValue"] {
        color: #f9fafb !important;
        font-size: 2rem !important;
    }

    /* Seção */
    .section-title {
        color: #f9fafb;
        font-size: 1.1rem;
        font-weight: 600;
        padding: 0.5rem 0;
        border-bottom: 1px solid #1f2937;
        margin-bottom: 1rem;
    }
    .step-badge {
        background: #c9a84c;
        color: #0d1117;
        font-weight: 700;
        padding: 2px 10px;
        border-radius: 12px;
        font-size: 0.8rem;
        margin-right: 8px;
    }

    /* Upload */
    [data-testid="stFileUploader"] {
        background: #111827;
        border: 2px dashed #374151;
        border-radius: 8px;
        padding: 1rem;
    }
    [data-testid="stFileUploader"]:hover { border-color: #c9a84c; }

    /* Botões */
    .stButton > button {
        border-radius: 6px;
        font-weight: 600;
        padding: 0.55rem 1.2rem;
        transition: all 0.2s;
    }
    .stButton > button[kind="primary"] {
        background: #c9a84c !important;
        color: #0d1117 !important;
        border: none !important;
    }
    .stButton > button[kind="primary"]:hover {
        background: #b8973d !important;
        transform: translateY(-1px);
    }
    .stButton > button[kind="secondary"] {
        background: transparent !important;
        color: #c9a84c !important;
        border: 1px solid #c9a84c !important;
    }
    .stButton > button[kind="secondary"]:hover {
        background: #c9a84c22 !important;
    }

    /* Dataframe */
    [data-testid="stDataFrame"] {
        border: 1px solid #1f2937;
        border-radius: 8px;
        overflow: hidden;
    }

    /* Expander */
    [data-testid="stExpander"] {
        background: #111827;
        border: 1px solid #1f2937;
        border-radius: 8px;
    }

    /* Divider */
    hr { border-color: #1f2937 !important; }

    /* Checkbox */
    [data-testid="stCheckbox"] label { color: #e5e7eb !important; }

    /* Info / success / warning / error */
    [data-testid="stAlert"] { border-radius: 6px; }

    /* Progress */
    [data-testid="stProgressBar"] > div > div { background-color: #c9a84c !important; }

    /* Caption */
    .stCaption { color: #6b7280 !important; }
</style>
""", unsafe_allow_html=True)


# ── Header ──────────────────────────────────────────────────────
st.markdown("""
<div class="hayden-header">
    <div>
        <div class="hayden-logo">HAYDEN CAPITAL</div>
        <div class="hayden-subtitle">Plataforma de Consulta de Matrículas · Registro de Imóveis</div>
    </div>
</div>
""", unsafe_allow_html=True)


# ── Estado da sessão ────────────────────────────────────────────
for key in ["df_match", "df_imoveis", "arquivo_nome", "log_enviado", "etapa"]:
    if key not in st.session_state:
        st.session_state[key] = None
if st.session_state.etapa is None:
    st.session_state.etapa = "upload"


# ── Nova busca ──────────────────────────────────────────────────
def resetar():
    st.session_state.df_match    = None
    st.session_state.df_imoveis  = None
    st.session_state.arquivo_nome = None
    st.session_state.log_enviado = None
    st.session_state.etapa       = "upload"


# ════════════════════════════════════════════════════════════════
# ETAPA 1 — Upload
# ════════════════════════════════════════════════════════════════
if st.session_state.etapa == "upload":

    st.markdown('<div class="section-title"><span class="step-badge">1</span>Carregar planilha de imóveis</div>', unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "Arraste ou selecione o arquivo Excel (.xls / .xlsx)",
        type=["xls", "xlsx"],
        help="A planilha deve conter a aba 'Imóveis Rurais Nacional' com as colunas padrão.",
        label_visibility="collapsed",
    )

    if uploaded:
        st.session_state.arquivo_nome = uploaded.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=os.path.splitext(uploaded.name)[1]) as tmp:
            tmp.write(uploaded.read())
            tmp_path = tmp.name

        with st.spinner("Lendo planilha e cruzando com cartórios CNJ..."):
            try:
                df_imoveis = carregar_imoveis(tmp_path)
                df_cnj     = carregar_cartorios()
                df_match   = cruzar(df_imoveis, df_cnj)
                st.session_state.df_match   = df_match
                st.session_state.df_imoveis = df_imoveis
                st.session_state.log_enviado = None
                st.session_state.etapa = "match"
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao processar arquivo: {e}")
            finally:
                os.unlink(tmp_path)


# ════════════════════════════════════════════════════════════════
# ETAPA 2 — Match
# ════════════════════════════════════════════════════════════════
elif st.session_state.etapa in ("match", "envio", "concluido"):

    df_match = st.session_state.df_match

    # Barra de ações topo
    col_arq, col_btn = st.columns([6, 1])
    with col_arq:
        st.markdown(f"📄 **{st.session_state.arquivo_nome}**  ·  {len(st.session_state.df_imoveis)} imóveis")
    with col_btn:
        if st.button("↩ Nova busca", type="secondary"):
            resetar()
            st.rerun()

    st.divider()

    # ── Métricas ────────────────────────────────────────────────
    st.markdown('<div class="section-title"><span class="step-badge">2</span>Resultado do cruzamento</div>', unsafe_allow_html=True)

    aptos = df_match[
        (df_match["match_metodo"] != "NAO_ENCONTRADO") &
        df_match["cartorio_email"].notna() &
        (df_match["cartorio_email"] != "")
    ]
    nao_enc   = len(df_match[df_match["match_metodo"] == "NAO_ENCONTRADO"])
    sem_email = len(df_match[
        (df_match["match_metodo"] != "NAO_ENCONTRADO") &
        (df_match["cartorio_email"].isna() | (df_match["cartorio_email"] == ""))
    ])

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Imóveis na base",        len(st.session_state.df_imoveis))
    c2.metric("Cartórios identificados", len(df_match[df_match["match_metodo"] != "NAO_ENCONTRADO"]))
    c3.metric("Não encontrados",         nao_enc)
    c4.metric("📧 Prontos para envio",   len(aptos))

    # ── Tabela ──────────────────────────────────────────────────
    cols_ok = [c for c in ["nirf_crf","denominacao","municipio","comarca","uf_sigla",
                            "cartorio_nome","cartorio_email","match_metodo"] if c in df_match.columns]
    df_display = df_match[cols_ok].rename(columns={
        "nirf_crf":"NIRF/CRF", "denominacao":"Denominação", "municipio":"Município",
        "comarca":"Comarca", "uf_sigla":"UF", "cartorio_nome":"Cartório",
        "cartorio_email":"E-mail", "match_metodo":"Método"
    })

    with st.expander("🔎 Ver detalhes do cruzamento", expanded=True):
        st.dataframe(df_display, hide_index=True, width="stretch")

    # ── Disparo ──────────────────────────────────────────────────
    st.divider()
    st.markdown('<div class="section-title"><span class="step-badge">3</span>Envio de e-mails</div>', unsafe_allow_html=True)

    if len(aptos) == 0:
        st.warning("Nenhum cartório com e-mail identificado para envio.")

    elif st.session_state.etapa == "concluido" and st.session_state.log_enviado is not None:
        df_log    = st.session_state.log_enviado
        enviados  = len(df_log[df_log["Status"] == "✅ Enviado"])
        erros     = len(df_log[df_log["Status"].str.startswith("❌")])
        st.success(f"Envio concluído — {enviados} enviados · {erros} erros")
        st.dataframe(df_log, hide_index=True, width="stretch")

    else:
        col_t, col_r = st.columns(2)

        with col_t:
            st.markdown("**🧪 Modo Teste**")
            st.caption("Todos os e-mails chegam no seu inbox para você conferir o template.")
            btn_teste = st.button(f"Enviar {len(aptos)} e-mails para meu e-mail", type="secondary")

        with col_r:
            st.markdown("**🚀 Modo Produção**")
            st.caption("E-mails enviados diretamente para os cartórios. Ação irreversível.")
            confirmar = st.checkbox("Confirmo o envio para os cartórios reais")
            btn_real  = st.button(f"Enviar {len(aptos)} e-mails para os cartórios", type="primary", disabled=not confirmar)

        modo_teste = None
        if btn_teste:           modo_teste = True
        elif btn_real and confirmar: modo_teste = False

        if modo_teste is not None:
            from dotenv import load_dotenv
            load_dotenv()
            EMAIL_TESTE = os.getenv("EMAIL_TESTE", "")

            progress = st.progress(0)
            status_txt = st.empty()
            logs = []

            try:
                token = _obter_token_delegado()

                for i, (_, row) in enumerate(aptos.iterrows(), 1):
                    rd   = row.to_dict()
                    dest = EMAIL_TESTE if modo_teste and EMAIL_TESTE else row["cartorio_email"]
                    try:
                        _graph_enviar(token, row["cartorio_email"],
                                      _assunto(rd, modo_teste), _corpo_html(rd), modo_teste)
                        status = "✅ Enviado"
                    except Exception as e:
                        status = f"❌ Erro: {e}"

                    logs.append({
                        "NIRF/CRF":    row.get("nirf_crf"),
                        "Denominação": row.get("denominacao"),
                        "Cartório":    row.get("cartorio_nome"),
                        "E-mail":      row.get("cartorio_email"),
                        "Método":      row.get("match_metodo"),
                        "Status":      status,
                    })

                    progress.progress(i / len(aptos))
                    status_txt.markdown(f"`[{i}/{len(aptos)}]` {status} → **{dest}**")

                    if i < len(aptos):
                        time.sleep(DELAY_ENTRE_ENVIOS)

                st.session_state.log_enviado = pd.DataFrame(logs)
                st.session_state.etapa = "concluido"
                st.rerun()

            except Exception as e:
                st.error(f"Erro no disparo: {e}")
