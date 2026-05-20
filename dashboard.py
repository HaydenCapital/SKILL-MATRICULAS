import streamlit as st
import pandas as pd
import sys
import os
import tempfile
import time
import unicodedata

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "app"))

from modulo1_leitura import carregar_imoveis
from modulo2_cartorios import carregar_cartorios
from modulo3_match import cruzar
from modulo4_email import obter_token_streamlit, _graph_enviar, _assunto, _corpo_html, DELAY_ENTRE_ENVIOS

OVERRIDE_PATH = os.path.join(os.path.dirname(__file__), "data", "overrides", "municipio_ri_override.csv")


def _norm(s: str) -> str:
    s = str(s).strip().upper()
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


# ── Config ──────────────────────────────────────────────────────
st.set_page_config(
    page_title="Hayden Capital — Consulta de Matrículas",
    page_icon="H",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ── Identidade visual Hayden Capital ────────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #0d1117; }
    [data-testid="stAppViewContainer"] { background-color: #0d1117; }
    [data-testid="stHeader"] { background-color: #0d1117; }
    [data-testid="stSidebar"] { background-color: #111827; }

    html, body, [class*="css"] {
        font-family: 'Inter', 'Segoe UI', sans-serif;
        color: #e5e7eb;
    }

    .hayden-header {
        background: linear-gradient(135deg, #111827 0%, #1a2540 100%);
        border-bottom: 2px solid #c9a84c;
        padding: 1.5rem 2rem;
        border-radius: 8px;
        margin-bottom: 1.5rem;
    }

    /* Barra superior sticky */
    .topbar-fixa {
        position: -webkit-sticky;
        position: sticky;
        top: 0;
        z-index: 999;
        background: linear-gradient(135deg, #111827 0%, #1a2540 100%);
        border-bottom: 2px solid #c9a84c;
        padding: 0.7rem 2rem;
        display: flex;
        align-items: center;
        justify-content: space-between;
        margin-bottom: 1rem;
    }
    .topbar-logo {
        font-size: 1.1rem;
        font-weight: 800;
        color: #c9a84c;
        letter-spacing: 1px;
    }
    .topbar-info {
        font-size: 0.82rem;
        color: #9ca3af;
    }
    .topbar-spacer {
        height: 4px;
    }
    /* Remove padding extra do bloco principal para sticky funcionar */
    section.main > div { padding-top: 0 !important; }
    .block-container { padding-top: 1rem !important; }
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

    [data-testid="stFileUploader"] {
        background: #111827;
        border: 2px dashed #374151;
        border-radius: 8px;
        padding: 1rem;
    }
    [data-testid="stFileUploader"]:hover { border-color: #c9a84c; }

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

    [data-testid="stDataFrame"] {
        border: 1px solid #1f2937;
        border-radius: 8px;
        overflow: hidden;
    }
    /* Remove botão de expandir e handle de redimensionar */
    [data-testid="stDataFrame"] [data-testid="StyledFullScreenButton"] {
        display: none !important;
    }
    [data-testid="stDataFrame"] .stDataFrameResizeHandle {
        display: none !important;
    }
    [data-testid="stDataFrame"] iframe {
        resize: none !important;
        pointer-events: auto;
    }

    [data-testid="stExpander"] {
        background: #111827;
        border: 1px solid #1f2937;
        border-radius: 8px;
    }

    hr { border-color: #1f2937 !important; }
    [data-testid="stCheckbox"] label { color: #e5e7eb !important; }
    [data-testid="stAlert"] { border-radius: 6px; }
    [data-testid="stProgressBar"] > div > div { background-color: #c9a84c !important; }
    .stCaption { color: #6b7280 !important; }

    /* Tabs */
    [data-testid="stTabs"] [data-baseweb="tab-list"] {
        background: transparent;
        border-bottom: 1px solid #1f2937;
    }
    [data-testid="stTabs"] [data-baseweb="tab"] {
        color: #9ca3af;
        font-weight: 500;
    }
    [data-testid="stTabs"] [aria-selected="true"] {
        color: #c9a84c !important;
        border-bottom: 2px solid #c9a84c !important;
    }
</style>
""", unsafe_allow_html=True)


# ── Header ──────────────────────────────────────────────────────
st.markdown(f"""
<div class="hayden-header">
    <div>
        <img src="data:image/png;base64,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" alt="Hayden Capital" height="160" style="display:block;margin-bottom:6px;filter:brightness(0) invert(1);">
        <div class="hayden-subtitle">Plataforma de Consulta de Matrículas · Registro de Imóveis</div>
    </div>
</div>
""", unsafe_allow_html=True)


# ── Estado da sessão ────────────────────────────────────────────
for key in ["df_match", "df_imoveis", "df_cnj", "arquivo_nome", "log_enviado", "etapa"]:
    if key not in st.session_state:
        st.session_state[key] = None
if st.session_state.etapa is None:
    st.session_state.etapa = "upload"


def resetar():
    for key in ["df_match", "df_imoveis", "df_cnj", "arquivo_nome", "log_enviado"]:
        st.session_state[key] = None
    st.session_state.etapa = "upload"


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
                st.session_state.df_imoveis  = df_imoveis
                st.session_state.df_cnj      = df_cnj
                st.session_state.df_match    = df_match
                st.session_state.log_enviado = None
                st.session_state.etapa       = "match"
                st.rerun()
            except Exception as e:
                st.error(f"Erro ao processar arquivo: {e}")
            finally:
                os.unlink(tmp_path)


# ════════════════════════════════════════════════════════════════
# ETAPA 2 — Resultados
# ════════════════════════════════════════════════════════════════
elif st.session_state.etapa in ("match", "concluido"):

    df_match = st.session_state.df_match

    # ── Barra de ações topo ──────────────────────────────────────
    col_arq, col_btn = st.columns([6, 1])
    with col_arq:
        st.markdown(f"Arquivo: **{st.session_state.arquivo_nome}** · {len(st.session_state.df_imoveis)} imóveis")
    with col_btn:
        if st.button("Nova busca", type="secondary"):
            resetar()
            st.rerun()

    st.divider()

    # ── Métricas ─────────────────────────────────────────────────
    aptos = df_match[
        (df_match["match_metodo"] != "NAO_ENCONTRADO") &
        df_match["cartorio_email"].notna() &
        (df_match["cartorio_email"] != "")
    ]
    nao_enc   = df_match[df_match["match_metodo"] == "NAO_ENCONTRADO"]
    sem_email = df_match[
        (df_match["match_metodo"] != "NAO_ENCONTRADO") &
        (df_match["cartorio_email"].isna() | (df_match["cartorio_email"] == ""))
    ]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Imóveis na base",        len(st.session_state.df_imoveis))
    c2.metric("Cartórios identificados", len(df_match[df_match["match_metodo"] != "NAO_ENCONTRADO"]))
    c3.metric("Não encontrados",         len(nao_enc))
    c4.metric("Prontos para envio",      len(aptos))

    st.divider()

    # ── Abas ─────────────────────────────────────────────────────
    tab_resultado, tab_envio, tab_overrides = st.tabs([
        "Resultado do Cruzamento",
        "Envio de E-mails",
        f"Municipios sem RI  ({len(nao_enc)})",
    ])

    # ────────────────────────────────────────────────────────────
    # ABA 1 — Resultado
    # ────────────────────────────────────────────────────────────
    with tab_resultado:
        cols_ok = [c for c in ["nirf_crf", "denominacao", "municipio", "comarca",
                                "uf_sigla", "cartorio_nome", "cartorio_email", "match_metodo"]
                   if c in df_match.columns]
        df_display = df_match[cols_ok].copy()

        # Coluna CRI Indicado — município sede do RI, apenas para overrides manuais
        df_display.insert(
            df_display.columns.get_loc("comarca") + 1,
            "cri_indicado",
            df_match.apply(
                lambda r: r["cartorio_municipio"] if r["match_metodo"] == "override_manual" else "—",
                axis=1,
            ),
        )

        df_display = df_display.rename(columns={
            "nirf_crf":      "NIRF/CRF",
            "denominacao":   "Denominacao",
            "municipio":     "Municipio",
            "comarca":       "Comarca",
            "cri_indicado":  "CRI Indicado",
            "uf_sigla":      "UF",
            "cartorio_nome": "Cartorio",
            "cartorio_email":"E-mail",
            "match_metodo":  "Metodo",
        })

        altura = min(38 + len(df_display) * 35, 600)
        st.dataframe(
            df_display,
            hide_index=True,
            use_container_width=True,
            height=altura,
            column_config={
                "NIRF/CRF":     st.column_config.TextColumn("NIRF/CRF",     width=130),
                "Denominacao":  st.column_config.TextColumn("Denominacao",   width=180),
                "Municipio":    st.column_config.TextColumn("Municipio",     width=150),
                "Comarca":      st.column_config.TextColumn("Comarca",       width=150),
                "CRI Indicado": st.column_config.TextColumn("CRI Indicado",  width=130),
                "UF":           st.column_config.TextColumn("UF",            width=50),
                "Cartorio":     st.column_config.TextColumn("Cartorio",      width=220),
                "E-mail":       st.column_config.TextColumn("E-mail",        width=180),
                "Metodo":       st.column_config.TextColumn("Metodo",        width=160),
            },
        )

    # ────────────────────────────────────────────────────────────
    # ABA 2 — Municípios sem RI
    # ────────────────────────────────────────────────────────────
    with tab_overrides:
        st.markdown('<div class="section-title">Municipios sem Cartorio de RI identificado</div>', unsafe_allow_html=True)

        if len(nao_enc) == 0:
            st.success("Todos os municípios foram identificados. Nenhuma ação necessária.")
        else:
            # Carregar overrides existentes
            if os.path.exists(OVERRIDE_PATH):
                df_ov = pd.read_csv(OVERRIDE_PATH)
            else:
                df_ov = pd.DataFrame(columns=["municipio_norm", "uf_sigla", "municipio_ri_norm", "observacao"])

            st.markdown(
                f"**{len(nao_enc)} linha(s)** sem RI identificado. "
                "Informe o município do cartório competente para cada um:"
            )
            st.caption(
                "O nome deve ser o municipio onde fica o Cartório de RI (ex: Umuarama). "
                "Após salvar, baixe o CSV e substitua o arquivo data/overrides/ no repositório."
            )

            st.divider()

            municipios_unicos = nao_enc.drop_duplicates(subset=["municipio_norm", "uf_sigla"])
            novos_overrides = []

            for _, row in municipios_unicos.iterrows():
                mun_norm = _norm(str(row.get("municipio_norm", "")))
                uf       = _norm(str(row.get("uf_sigla", "")))
                mun_nome = row.get("municipio", mun_norm)

                ja_existe = not df_ov[
                    (df_ov["municipio_norm"].apply(_norm) == mun_norm) &
                    (df_ov["uf_sigla"].apply(_norm) == uf)
                ].empty

                col_mun, col_input, col_alerta = st.columns([2, 3, 2])

                with col_mun:
                    st.markdown(f"**{mun_nome}**")
                    st.caption(f"UF: {uf}  |  Comarca: {row.get('comarca', '—')}")

                with col_input:
                    if ja_existe:
                        ri_atual = df_ov[
                            (df_ov["municipio_norm"].apply(_norm) == mun_norm) &
                            (df_ov["uf_sigla"].apply(_norm) == uf)
                        ]["municipio_ri_norm"].values[0]
                        st.text_input(
                            "Municipio do RI competente",
                            value=ri_atual,
                            key=f"ri_{mun_norm}_{uf}",
                            disabled=True,
                        )
                    else:
                        ri_digitado = st.text_input(
                            "Municipio do RI competente",
                            placeholder="Ex: Umuarama",
                            key=f"ri_{mun_norm}_{uf}",
                        )
                        if ri_digitado.strip():
                            novos_overrides.append({
                                "municipio_norm":   mun_norm,
                                "uf_sigla":         uf,
                                "municipio_ri_norm": _norm(ri_digitado),
                                "observacao":       "MANUAL_DASHBOARD",
                            })

                with col_alerta:
                    if ja_existe:
                        st.warning("Ja cadastrado")
                    else:
                        st.markdown("")

                st.markdown("---")

            # Botão salvar
            if novos_overrides:
                if st.button("Salvar e re-aplicar cruzamento", type="primary"):
                    df_novos       = pd.DataFrame(novos_overrides)
                    df_ov_updated  = pd.concat([df_ov, df_novos], ignore_index=True)
                    os.makedirs(os.path.dirname(OVERRIDE_PATH), exist_ok=True)
                    df_ov_updated.to_csv(OVERRIDE_PATH, index=False)

                    with st.spinner("Re-aplicando cruzamento com novos overrides..."):
                        df_novo = cruzar(st.session_state.df_imoveis, st.session_state.df_cnj)
                        st.session_state.df_match = df_novo

                    st.success(f"{len(novos_overrides)} override(s) adicionado(s). Cruzamento atualizado.")
                    st.rerun()

            # Download CSV atualizado
            st.divider()
            df_ov_dl = pd.read_csv(OVERRIDE_PATH) if os.path.exists(OVERRIDE_PATH) else df_ov
            csv_bytes = df_ov_dl.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="Baixar CSV de overrides atualizado",
                data=csv_bytes,
                file_name="municipio_ri_override.csv",
                mime="text/csv",
                type="secondary",
            )
            st.caption(
                "Substitua o arquivo data/overrides/municipio_ri_override.csv no repositório "
                "e faça commit para que as alterações persistam nos próximos acessos."
            )

        # ── Planilha de revisão ──────────────────────────────────
        st.divider()
        st.markdown('<div class="section-title">Planilha de Revisao de RI por Municipio</div>', unsafe_allow_html=True)

        REVISAO_PATH = os.path.join(os.path.dirname(__file__), "data", "overrides", "municipios_sem_ri_para_revisao.xlsx")

        if not os.path.exists(REVISAO_PATH):
            st.info("Arquivo de revisão não encontrado: data/overrides/municipios_sem_ri_para_revisao.xlsx")
        else:
            import io as _io
            df_rev = pd.read_excel(REVISAO_PATH, header=1)

            COL_SUGERIDO   = "RI Sugerido (distancia)"
            COL_CONFIRMADO = "RI Confirmado\n(preencher se errado)"
            COL_DIST       = "Distancia (km)"

            # Renomeia para display mais limpo
            df_rev = df_rev.rename(columns={COL_CONFIRMADO: "RI Confirmado (preencher se errado)"})
            COL_CONFIRMADO = "RI Confirmado (preencher se errado)"

            pendentes = df_rev[df_rev[COL_CONFIRMADO].isna() | (df_rev[COL_CONFIRMADO].astype(str).str.strip() == "") | (df_rev[COL_CONFIRMADO].astype(str) == "nan")]
            preenchidos = len(df_rev) - len(pendentes)

            st.markdown(
                f"Total: **{len(df_rev)}** municípios · "
                f"Revisados: **{preenchidos}** · "
                f"Pendentes: **{len(pendentes)}**"
            )
            st.progress(preenchidos / len(df_rev) if len(df_rev) > 0 else 0)
            st.caption(
                "Instrucao: se o RI Sugerido estiver correto, nao preencha nada. "
                "Se estiver errado, escreva o municipio correto na coluna 'RI Confirmado'."
            )

            st.divider()

            # Filtros
            col_f1, col_f2 = st.columns([2, 3])
            with col_f1:
                mostrar = st.radio(
                    "Exibir",
                    ["Todos", "Apenas pendentes"],
                    horizontal=True,
                )
            with col_f2:
                busca = st.text_input(
                    "Pesquisar município",
                    placeholder="Ex: Tapejara",
                )

            df_exibir = pendentes if mostrar == "Apenas pendentes" else df_rev

            if busca.strip():
                df_exibir = df_exibir[
                    df_exibir["Municipio"].str.contains(busca.strip(), case=False, na=False)
                ]
                if df_exibir.empty:
                    st.warning(f"Nenhum município encontrado com '{busca}'.")

            # Editor — apenas a coluna RI Confirmado é editável
            colunas_fixas = [c for c in df_exibir.columns if c != COL_CONFIRMADO]
            df_editado = st.data_editor(
                df_exibir.reset_index(drop=True),
                disabled=colunas_fixas,
                hide_index=True,
                use_container_width=True,
                key="editor_revisao",
            )

            # Salvar edições de volta ao df_rev completo
            if st.button("Salvar alteracoes e baixar planilha", type="primary"):
                df_rev.update(df_editado)
                buf = _io.BytesIO()
                df_rev.to_excel(buf, index=False)
                buf.seek(0)
                st.download_button(
                    label="Baixar municipios_sem_ri_para_revisao.xlsx atualizado",
                    data=buf,
                    file_name="municipios_sem_ri_para_revisao.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="secondary",
                )
                st.caption("Substitua o arquivo no repositório e faça commit para persistir as alterações.")

    # ────────────────────────────────────────────────────────────
    # ABA 3 — Envio de e-mails
    # ────────────────────────────────────────────────────────────
    with tab_envio:

        if st.session_state.etapa == "concluido" and st.session_state.log_enviado is not None:
            df_log   = st.session_state.log_enviado
            enviados = len(df_log[df_log["Status"] == "Enviado"])
            erros    = len(df_log[df_log["Status"].str.startswith("Erro")])
            st.success(f"Envio concluído — {enviados} enviado(s) · {erros} erro(s)")
            st.dataframe(df_log, hide_index=True, width="stretch")

        elif len(aptos) == 0:
            st.warning("Nenhum cartório com e-mail identificado para envio.")

        else:
            col_t, col_r = st.columns(2)

            with col_t:
                st.markdown("**Modo Teste**")
                st.caption("Todos os e-mails chegam no seu inbox para conferir o template antes do envio real.")
                btn_teste = st.button(f"Enviar {len(aptos)} e-mails para meu e-mail", type="secondary")

            with col_r:
                st.markdown("**Modo Producao**")
                st.caption("E-mails enviados diretamente para os cartórios. Ação irreversível.")
                confirmar = st.checkbox("Confirmo o envio para os cartórios reais")
                btn_real  = st.button(
                    f"Enviar {len(aptos)} e-mails para os cartórios",
                    type="primary",
                    disabled=not confirmar,
                )

            # Salva modo_teste na session para persistir durante o login Microsoft
            if btn_teste:
                st.session_state["modo_envio_pendente"] = True
            elif btn_real and confirmar:
                st.session_state["modo_envio_pendente"] = False

            modo_teste = st.session_state.get("modo_envio_pendente")

            if modo_teste is not None:
                from dotenv import load_dotenv
                load_dotenv()
                EMAIL_TESTE = os.getenv("EMAIL_TESTE", "")

                # Autenticação Microsoft — exibe UI de login se necessário
                token = obter_token_streamlit(st)
                if not token:
                    st.stop()

                progress   = st.progress(0)
                status_txt = st.empty()
                logs       = []

                try:
                    for i, (_, row) in enumerate(aptos.iterrows(), 1):
                        rd   = row.to_dict()
                        dest = EMAIL_TESTE if modo_teste and EMAIL_TESTE else row["cartorio_email"]
                        try:
                            _graph_enviar(
                                token,
                                row["cartorio_email"],
                                _assunto(rd, modo_teste),
                                _corpo_html(rd, modo_teste),
                                modo_teste,
                            )
                            status = "Enviado"
                        except Exception as e:
                            status = f"Erro: {e}"

                        logs.append({
                            "NIRF/CRF":    row.get("nirf_crf"),
                            "Denominacao": row.get("denominacao"),
                            "Cartorio":    row.get("cartorio_nome"),
                            "E-mail":      row.get("cartorio_email"),
                            "Metodo":      row.get("match_metodo"),
                            "Status":      status,
                        })

                        progress.progress(i / len(aptos))
                        status_txt.markdown(f"[{i}/{len(aptos)}] {status} — {dest}")

                        if i < len(aptos):
                            time.sleep(DELAY_ENTRE_ENVIOS)

                    st.session_state.log_enviado = pd.DataFrame(logs)
                    st.session_state.etapa       = "concluido"
                    st.session_state.pop("modo_envio_pendente", None)
                    st.rerun()

                except Exception as e:
                    st.session_state.pop("modo_envio_pendente", None)
                    st.error(f"Erro no disparo: {e}")
