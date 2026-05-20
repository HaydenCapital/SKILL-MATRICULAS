import time
import os
import json
import html
from dotenv import load_dotenv
import pandas as pd
import msal
import urllib.request
import urllib.parse

load_dotenv()

GRAPH_TENANT_ID = os.getenv("GRAPH_TENANT_ID", "")
GRAPH_CLIENT_ID = os.getenv("GRAPH_CLIENT_ID", "")

REMETENTE       = os.getenv("EMAIL_REMETENTE", "")
NOME_REM        = os.getenv("EMAIL_NOME_REMETENTE", "")
REMETENTE_TESTE = os.getenv("EMAIL_REMETENTE_TESTE", "")
NOME_REM_TESTE  = os.getenv("EMAIL_NOME_REMETENTE_TESTE", "")
EMAIL_TESTE     = os.getenv("EMAIL_TESTE", "")
MODO_ENVIO      = os.getenv("MODO_ENVIO", "delegado").lower()

DELAY_ENTRE_ENVIOS = 3
TOKEN_CACHE_PATH   = "data/cache/token_cache.json"
SCOPES             = ["Mail.Send"]


# ─────────────────────────────────────────────
# Template
# ─────────────────────────────────────────────

LOGO_H_B64 = (
    "iVBORw0KGgoAAAANSUhEUgAAADwAAAA8CAYAAAA6/NlyAAADgUlEQVR4nO2avWscRxTA33v7Id2X"
    "0JlwQWCcIq6cFAHXRhDIHyCkOQl3ViH/CTFpbDdOkf/gQjqpuQVVqoxNkHuTJiSFiiSduMIn9u"
    "72uLvRvPCWWyPHLnS7Kzgy82v2g50f82Z2Z+cLwOFwOBwOh8NxM2DedEopkpNer/eRo9VqsRyj"
    "KDIAwEvktg/M8TxvbW21qtXqN1rrujGmKfeYOUHEFQAIEfFdGIajJEn+PD4+/idLdx23UurOys"
    "rKvel0WmPmWwAwZeYJIlblGSLqE9HI87zfDg8Pe9d0v8eHxUjFvu8PEPEMEdeJyGPmp/V6/ask"
    "Sc6NMU+Y+e/ZbHahtX53Nd113JPJRAI6Y+Z1AKgj4o+1Wm1jPB7/AQDPEPEvz/MuJpPJYAF34W"
    "/4A7a3tx+sra39OhwOH0dR9AuUyM7OzqNGo/HzaDT6ttvtminq8wukxYODA7/f76eNhzHGAwCt"
    "lJKjF0XRrECjIg1XAACXRKTnbhB3s9mkTqej87oJ8sMSbBRFl1czKte9Xq9oC8riEJcx5n0e5T"
    "or4LxigiUHEUv99RBYBoFlEFgGgWUQWAaBZRBYBoFlEFiGX7KPpL87GAw8pVShHlLmQERa5oCH"
    "87711f51XlLH3t7eEJYwYE9rGcDA90qph1LTiCid/Nwws9SsMcbcvrxMY09HTEsRMBEZovTNe4"
    "2IL+eZK1TLiJg6mPk7IrrPzIUKsNSAjTE8D/htt9t9BSWyu7vbRMTSRk0+lEstmwAo4TvOHDUo"
    "Eb9MmVS2NFqbm5t4enpaKODM0W63S3mVrf0PE1gGgWUQWAaBZRBYBoFlEFgGgWVQGRJjTLoox8"
    "ylLM5dJXOW5fYLpEVZ2FJKpflJbyCaeV+aFl23/a+71WqlkwkyEpu7OVtMkzLO6/YhP9zpdGbz"
    "5VJZqJa+s1/SBABHUTSVk3a77YtbargMN+ZJpJSqBIGwobVeZ+a7RPS0UqncG4/H58z8AwD8Ho"
    "Zhn5l7R0dH8SLu/f39xnQ6/Xzu/hoAXlQqlQ3ZTYCIz5n5zPf9iziOz09OTpKbrmGU0tdaN8Iw"
    "vMvMNUSU4dtPcRyPgyCQ7Q4hM99BxFtJksgbEC+y5WEwGDRXV1e/NMbI6r/sLngSx/EkCIIKM1"
    "eJ6AtE/CwIApn6SQp+Ov9/MG+6bGvRp7YX3dS2pcxbwO1wOBwOh8PhgJviX+tUn7cB0Ai1AAAA"
    "AElFTkSuQmCC"
)


def _assinatura_simples(nome: str, cargo: str, email: str, tel: str) -> str:
    """Assinatura em texto simples — sem card visual (modo teste)."""
    linhas = [f"<strong>{nome}</strong>"]
    if cargo:
        linhas.append(cargo)
    linhas.append("Hayden Capital")
    if tel:
        linhas.append(tel)
    linhas.append(f'<a href="mailto:{email}" style="color:#1F4E79;">{email}</a>')
    return "<p style='margin:0;line-height:1.8;'>" + "<br>".join(linhas) + "</p>"


def _assinatura_card(nome: str, cargo: str, email: str, tel: str) -> str:
    """Assinatura corporativa card-style — painel branco + cinza (modo produção)."""
    return f"""<table style="border-collapse:collapse;width:480px;margin:8px 0 24px 0;font-family:Georgia,serif;font-size:12px;">
    <tr>
      <td style="width:200px;background:#ffffff;padding:22px 20px;border:1px solid #d8dde3;border-right:none;vertical-align:middle;">
        <img src="data:image/png;base64,{LOGO_H_B64}"
             alt="H" width="38" height="38" style="display:block;margin-bottom:14px;opacity:0.75;">
        <div style="font-size:13px;font-weight:400;color:#2c2c2c;letter-spacing:0.5px;">Hayden Capital</div>
        <div style="border-top:1px solid #b0b8c4;margin:10px 0;width:120px;"></div>
        <div style="font-size:13px;font-weight:600;color:#1a1a1a;margin-bottom:4px;">{nome}</div>
        <div style="font-size:11px;color:#666;letter-spacing:0.3px;">{cargo}</div>
      </td>
      <td style="background:#636363;padding:22px 18px;vertical-align:middle;border:1px solid #636363;">
        <table cellpadding="0" cellspacing="0" border="0">
          <tr><td style="padding-bottom:10px;color:#ffffff;font-size:12px;font-family:Arial,sans-serif;">&#128222;&nbsp; {tel}</td></tr>
          <tr><td style="padding-bottom:10px;font-size:12px;font-family:Arial,sans-serif;">
            <a href="mailto:{email}" style="color:#ffffff;text-decoration:none;">&#9993;&nbsp; {email}</a>
          </td></tr>
          <tr><td style="color:#ffffff;font-size:12px;font-family:Arial,sans-serif;line-height:1.6;">
            &#128205;&nbsp; Rua Urussui n&#186;125, 4&#186; andar<br>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Itaim Bibi, S&#227;o Paulo<br>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;CEP: 04.542-050
          </td></tr>
        </table>
      </td>
    </tr>
  </table>"""


def _assunto(row: dict, teste: bool) -> str:
    base = f"Solicitação de Número de Matrícula – Imóvel Rural | NIRF {row['nirf_crf']}"
    return f"[TESTE] {base}" if teste else base


def _corpo_html(row: dict, modo_teste: bool = False) -> str:
    nome_cartorio = html.escape(str(row.get('cartorio_nome') or '').strip())
    saudacao = (
        f"Prezado(a) Oficial de Registro de Imóveis – {nome_cartorio},"
        if nome_cartorio else
        "Prezado(a) Oficial de Registro de Imóveis,"
    )
    # Assinatura sempre usa dados de produção (Luiza é sempre a remetente)
    # Modo teste só muda o DESTINO do e-mail, não a assinatura
    email_sig = REMETENTE
    nome_sig  = NOME_REM
    nome_disp = nome_sig.split("|")[0].strip() if nome_sig else "Hayden Capital"
    cargo_sig = os.getenv("EMAIL_CARGO", "")
    tel_sig   = os.getenv("EMAIL_TELEFONE", "")

    denominacao = html.escape(str(row.get('denominacao') or '—'))
    municipio   = html.escape(str(row.get('municipio')   or '—'))
    uf_sigla    = html.escape(str(row.get('uf_sigla')    or '—'))
    comarca     = html.escape(str(row.get('comarca')     or '—'))
    nirf_crf    = html.escape(str(row.get('nirf_crf')    or '—'))
    titular     = html.escape(str(row.get('titular')     or '—'))

    return f"""<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"></head>
<body style="font-family:Arial,sans-serif;font-size:14px;color:#222;max-width:600px;margin:0 auto;line-height:1.6;">

  <p>{saudacao}</p>

  <p>
    A <strong>Hayden Capital</strong> é uma empresa de gestão de ativos com atuação em
    reestruturação e recuperação de crédito. No contexto de nossa análise patrimonial,
    identificamos um imóvel rural possivelmente registrado nessa Serventia cujos dados
    constam abaixo:
  </p>

  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin:14px 0;">
    <tr><td align="center">
  <table style="border-collapse:collapse;width:auto;min-width:420px;font-size:13px;">
    <thead>
      <tr style="background-color:#1F4E79;color:white;">
        <th style="padding:6px 14px;text-align:left;font-weight:600;">Campo</th>
        <th style="padding:6px 14px;text-align:left;font-weight:600;">Informação</th>
      </tr>
    </thead>
    <tbody>
      <tr style="background-color:#f2f7fc;">
        <td style="padding:5px 14px;border-bottom:1px solid #e0e0e0;color:#555;">Denominação</td>
        <td style="padding:5px 14px;border-bottom:1px solid #e0e0e0;">{denominacao}</td>
      </tr>
      <tr>
        <td style="padding:5px 14px;border-bottom:1px solid #e0e0e0;color:#555;">Município / UF</td>
        <td style="padding:5px 14px;border-bottom:1px solid #e0e0e0;">{municipio} – {uf_sigla}</td>
      </tr>
      <tr style="background-color:#f2f7fc;">
        <td style="padding:5px 14px;border-bottom:1px solid #e0e0e0;color:#555;">Comarca</td>
        <td style="padding:5px 14px;border-bottom:1px solid #e0e0e0;">{comarca}</td>
      </tr>
      <tr>
        <td style="padding:5px 14px;border-bottom:1px solid #e0e0e0;color:#555;">Código NIRF / CRF</td>
        <td style="padding:5px 14px;border-bottom:1px solid #e0e0e0;"><strong>{nirf_crf}</strong></td>
      </tr>
      <tr style="background-color:#f2f7fc;">
        <td style="padding:5px 14px;color:#555;">Titular</td>
        <td style="padding:5px 14px;">{titular}</td>
      </tr>
    </tbody>
  </table>
  </td></tr></table>

  <p>
    Solicitamos, por gentileza, a confirmação do <strong>número de matrícula</strong>
    desse imóvel registrado nessa Serventia, ou, caso não conste em seus registros,
    que nos informe para que possamos buscar a serventia competente.
  </p>

  <p>
    Desde já agradecemos a colaboração e nos colocamos à disposição para quaisquer esclarecimentos.
  </p>

  <p style="margin-top:20px;">Atenciosamente,</p>

  {_assinatura_card(nome_disp, cargo_sig, email_sig, tel_sig)}

  <hr style="border:none;border-top:1px solid #e5e5e5;margin-top:4px;">
  <p style="font-size:11px;color:#999;margin-top:8px;">
    Esta mensagem é de caráter informativo e destinada exclusivamente ao destinatário indicado.
    Caso tenha recebido por engano, pedimos que nos informe pelo e-mail acima.
  </p>
</body>
</html>"""


# ─────────────────────────────────────────────
# Autenticação Delegada (Device Code Flow)
# ─────────────────────────────────────────────

def _carregar_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_CACHE_PATH):
        with open(TOKEN_CACHE_PATH, "r", encoding="utf-8") as f:
            cache.deserialize(f.read())
    return cache


def _salvar_cache(cache: msal.SerializableTokenCache):
    os.makedirs(os.path.dirname(TOKEN_CACHE_PATH), exist_ok=True)
    if cache.has_state_changed:
        with open(TOKEN_CACHE_PATH, "w", encoding="utf-8") as f:
            f.write(cache.serialize())


def _obter_token_delegado() -> str:
    cache  = _carregar_cache()
    app    = msal.PublicClientApplication(
        GRAPH_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}",
        token_cache=cache,
    )

    # Tenta usar token em cache primeiro
    contas = app.get_accounts()
    if contas:
        resultado = app.acquire_token_silent(SCOPES, account=contas[0])
        if resultado and "access_token" in resultado:
            _salvar_cache(cache)
            return resultado["access_token"]

    # Device code flow — abre login no navegador
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Falha ao iniciar device flow: {flow}")

    print("\n" + "="*60)
    print("LOGIN NECESSÁRIO — faça isso uma vez:")
    print(f"\n  1. Acesse: {flow['verification_uri']}")
    print(f"  2. Digite o código: {flow['user_code']}")
    print("\nAguardando login...")
    print("="*60 + "\n")

    resultado = app.acquire_token_by_device_flow(flow)

    if "access_token" not in resultado:
        raise RuntimeError(f"Falha no login: {resultado.get('error_description', resultado)}")

    _salvar_cache(cache)
    print("Login realizado com sucesso! Token salvo em cache.\n")
    return resultado["access_token"]


# ─────────────────────────────────────────────
# Envio via Graph API
# ─────────────────────────────────────────────

def _graph_enviar(token: str, destinatario: str, assunto: str, corpo_html: str, modo_teste: bool):
    dest = EMAIL_TESTE if modo_teste and EMAIL_TESTE else destinatario
    payload = json.dumps({
        "message": {
            "subject": assunto,
            "body":    {"contentType": "HTML", "content": corpo_html},
            "toRecipients": [{"emailAddress": {"address": dest}}],
        },
        "saveToSentItems": "true"
    }).encode("utf-8")

    url = f"https://graph.microsoft.com/v1.0/me/sendMail"
    req = urllib.request.Request(url, data=payload, method="POST")
    req.add_header("Authorization", f"Bearer {token}")
    req.add_header("Content-Type",  "application/json")

    with urllib.request.urlopen(req) as resp:
        if resp.status not in (200, 202):
            raise RuntimeError(f"Graph retornou {resp.status}")


# ─────────────────────────────────────────────
# Ponto de entrada público
# ─────────────────────────────────────────────

def disparar(df_resultado: pd.DataFrame, modo_teste: bool = True) -> pd.DataFrame:
    df = df_resultado[
        df_resultado["cartorio_email"].notna() &
        (df_resultado["cartorio_email"] != "") &
        (df_resultado["match_metodo"] != "NAO_ENCONTRADO")
    ].copy()

    if df.empty:
        print("Nenhum e-mail para disparar.")
        return pd.DataFrame()

    print(f"\n{'='*60}")
    print(f"Modo teste  : {'SIM → ' + EMAIL_TESTE if modo_teste else 'NÃO (e-mails reais)'}")
    print(f"Total emails: {len(df)}")
    print(f"{'='*60}")

    token = _obter_token_delegado()

    logs = []
    for i, (_, row) in enumerate(df.iterrows(), 1):
        rd   = row.to_dict()
        dest = EMAIL_TESTE if modo_teste and EMAIL_TESTE else row["cartorio_email"]
        try:
            _graph_enviar(token, row["cartorio_email"],
                          _assunto(rd, modo_teste), _corpo_html(rd, modo_teste), modo_teste)
            status = "Enviado"
            print(f"[{i:>3}/{len(df)}] OK   → {dest} | NIRF {row.get('nirf_crf')} | {str(row.get('cartorio_nome',''))[:45]}")
        except Exception as e:
            status = f"Erro: {e}"
            print(f"[{i:>3}/{len(df)}] ERRO → {dest} | {e}")

        logs.append({
            "nirf_crf":       row.get("nirf_crf"),
            "denominacao":    row.get("denominacao"),
            "comarca":        row.get("comarca"),
            "uf_sigla":       row.get("uf_sigla"),
            "cartorio_nome":  row.get("cartorio_nome"),
            "cartorio_email": row.get("cartorio_email"),
            "match_metodo":   row.get("match_metodo"),
            "status_envio":   status,
        })

        if i < len(df):
            time.sleep(DELAY_ENTRE_ENVIOS)

    return pd.DataFrame(logs)


# ─────────────────────────────────────────────
# Autenticação via Streamlit (sem terminal)
# ─────────────────────────────────────────────

def obter_token_streamlit(st_ref) -> str:
    """
    Gerencia autenticação Microsoft dentro do Streamlit.
    Retorna o token se disponível; exibe UI de login e retorna "" se aguardando.
    st_ref = módulo streamlit (passado pelo dashboard para evitar import circular)
    """
    import concurrent.futures

    # 1. Token já na sessão (autenticado nesta sessão)
    if st_ref.session_state.get("msal_token"):
        return st_ref.session_state["msal_token"]

    cache = _carregar_cache()
    app   = msal.PublicClientApplication(
        GRAPH_CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}",
        token_cache=cache,
    )

    # 2. Token em cache de disco (login anterior ainda válido)
    contas = app.get_accounts()
    if contas:
        resultado = app.acquire_token_silent(SCOPES, account=contas[0])
        if resultado and "access_token" in resultado:
            _salvar_cache(cache)
            st_ref.session_state["msal_token"] = resultado["access_token"]
            return resultado["access_token"]

    # 3. Precisa de novo login — iniciar Device Code Flow
    if "msal_flow" not in st_ref.session_state:
        flow = app.initiate_device_flow(scopes=SCOPES)
        if "user_code" not in flow:
            st_ref.error(f"Erro ao iniciar autenticação Microsoft: {flow}")
            return ""
        st_ref.session_state["msal_flow"] = flow
        st_ref.session_state["msal_app"]  = app

    flow  = st_ref.session_state["msal_flow"]
    app_s = st_ref.session_state["msal_app"]

    # Exibe UI de login no dashboard
    st_ref.divider()
    st_ref.markdown("### 🔐 Login Microsoft necessário")
    st_ref.markdown(
        f"Para enviar e-mails, faça login com sua conta **@haydencapital.com.br**:"
    )
    col_info, col_btn = st_ref.columns([3, 1])
    with col_info:
        st_ref.markdown(
            f"1. Acesse **[{flow['verification_uri']}]({flow['verification_uri']})**\n"
            f"2. Digite o código abaixo e aprove o acesso"
        )
        st_ref.code(flow["user_code"], language=None)
        st_ref.caption("O código expira em ~15 min. Após confirmar no browser, clique no botão.")
    with col_btn:
        st_ref.markdown("<br><br><br>", unsafe_allow_html=True)
        if st_ref.button("✅ Já fiz o login", type="primary", key="btn_msal_confirm"):
            with st_ref.spinner("Verificando autenticação..."):
                try:
                    with concurrent.futures.ThreadPoolExecutor() as ex:
                        future = ex.submit(app_s.acquire_token_by_device_flow, flow)
                        resultado = future.result(timeout=15)
                    if "access_token" in resultado:
                        _salvar_cache(cache)
                        token = resultado["access_token"]
                        st_ref.session_state["msal_token"] = token
                        del st_ref.session_state["msal_flow"]
                        del st_ref.session_state["msal_app"]
                        return token
                    else:
                        st_ref.error("Login não confirmado. Verifique o browser e tente novamente.")
                        del st_ref.session_state["msal_flow"]
                        del st_ref.session_state["msal_app"]
                except concurrent.futures.TimeoutError:
                    st_ref.error("Tempo esgotado (15s). Complete o login no browser primeiro.")
                    del st_ref.session_state["msal_flow"]
                    del st_ref.session_state["msal_app"]
    return ""


def testar_conexao() -> bool:
    print("Obtendo token de acesso (delegado)...")
    try:
        token = _obter_token_delegado()
        print(f"Token OK — {len(token)} chars")
        return True
    except Exception as e:
        print(f"Falha: {e}")
        return False


if __name__ == "__main__":
    import sys
    sys.path.insert(0, os.path.dirname(__file__))
    from modulo1_leitura import carregar_imoveis
    from modulo2_cartorios import carregar_cartorios
    from modulo3_match import cruzar

    if not testar_conexao():
        print("\nVerifique o arquivo .env e as permissões no Azure.")
        sys.exit(1)

    df_imoveis   = carregar_imoveis("data/input/Pesquisa de Bens - Antonio Francischini.xls")
    df_cnj       = carregar_cartorios()
    df_resultado = cruzar(df_imoveis, df_cnj)

    df_log = disparar(df_resultado, modo_teste=True)

    if not df_log.empty:
        os.makedirs("data/output", exist_ok=True)
        log_path = "data/output/log_disparos.xlsx"
        df_log.to_excel(log_path, index=False)
        print(f"\nLog salvo: {log_path}")
        print(df_log["status_envio"].value_counts().to_string())
