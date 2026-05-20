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
    """Assinatura corporativa com imagem oficial da Luiza Nohra."""
    return (
        '<img src="data:image/jpeg;base64,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aAPzZooByAfaigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP12pKWigBNoo2ilooATAoxS0UAJtFBUGlooATaMUAYpaKAEKg9RSbRinUUAJgUYpaKAExRilooATaKMClooATApaKKACmTIssbI6hkYEMD0I70+kIzQB+XXxe8HN4B+JXiHQ9pWG3umaDI6xP8AMn6HH4Vx9fVf7dngn7LrGgeK4Y/kuYzp9yw/vL80ZJ9xuH4V8qUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQB+u1FFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQB5l+0d4I/4Tz4Ra9Yxpvu7eL7Zbeokj+bj6gEfjX5rBtwz69q/XORFkUqwDKeCD0Ir8v/AIx+DG8AfE7xDom0rDDctJBnvE/zofyOPwoA4yiiigAooooAKKKKACjpXtPw8/Zp1HxNp8Oo63eNo9rKoeK3RN07qehOeFH15rr9Y/ZN0uS1b+y9cvILgD5ftaLIhPvgAgUAfM9Fbfi/wdqngbWpNL1aDybhfmVlOUlXsynuKxKACiiigAooo6UAFFFFABRRR6e/SgAooo60AFFFLigBKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA/XaiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooASvjn9u3wV5GpeH/ABXDH8syNp9ywH8Q+aMn6jcPwr7Hrzb9ofwT/wAJ78I9f0+OPzLuGH7XbDv5kfzDH1AI/GgD80qKAdwB9aKACiiigAruPgr4eg8TfEzRrS5QSW6M1w6N0YIM4/PH5Vw9dL8N/FY8E+NtL1hwWggl2zBepjYYb9Dn8KAPYf2svGPjHwPqvgXU9E1STR/Dwvgl/dKpMQkLgKJgOqbd3H19q+iLeZLq2juIJEmglQSJNEcoykAhlI6g1l6npmjePPDUtneQwavomow/MrfNHKh7g9j3yOQa+eX/AGbfiR4O8U6Qvgj4hzwe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GPxoA/MX3oowRwQQe4PUGigAooooAKKKKAClzSUUAFFFFAB3ooooAP50dsfzoooAXNJjiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA/XaiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigApGGRS0h6UAfmb8f/BZ8BfFvxBpqJstZZvtlsOxjk+bA+hJH4V55X1/+3d4K32vh/xZBGcxO2n3LAfwtloyT9dw/GvkCgAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP12ooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAOG+NvgtfiB8LvEGi7Q08ts0kGe0qfMv6jH41+YZDAncCrZwwPY9xX65sM+9fmh+0J4IPgL4u+INORNlpNL9st/+ucnzYH0O4fhQB51RRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAfrtRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFfJv7d3goT6foHiqGPLQOdPuWA/gb5kJ+hDD8a+sq4j4z+DF8f/DLxBou0PPPbM8HqJU+ZMfiMfjQB+YFFKVZSQ42uOGB7HvSUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFAH67UUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABSEUtFAH5pftFeCv+EE+L+v2KR+Xa3Ev222AHBSXLYH0bcK82r7G/bt8Fefpeg+KoU+a2kNhcsB/A/KE/wDAgfzr45oAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKAP12ooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigDi/jD4LX4gfDXxBoZQNLcWzNBxnEq/Mn6jH41+XzKyEq42upwy+hBwR+dfroa/Nf9o/wV/wAIN8YNetI02Wl1L9utuONknJA+jbhQB5lRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQB+u1FFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFfH37e+l2yXHhDUlTF3Is9uz+qDawH4Fj+dFFAHyVRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQB/9k=" ' +
        'alt="Assinatura Luiza Nohra" width="500" ' +
        'style="display:block;max-width:100%;margin:8px 0 24px 0;">' 
    )


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
