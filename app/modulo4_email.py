import smtplib
import time
import os
import json
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import pandas as pd
import msal
import urllib.request
import urllib.parse

load_dotenv()

GRAPH_TENANT_ID     = os.getenv("GRAPH_TENANT_ID", "")
GRAPH_CLIENT_ID     = os.getenv("GRAPH_CLIENT_ID", "")
GRAPH_CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET", "")

REMETENTE  = os.getenv("EMAIL_REMETENTE", "")
NOME_REM   = os.getenv("EMAIL_NOME_REMETENTE", "Felipe | Hayden Capital")
EMAIL_TESTE = os.getenv("EMAIL_TESTE", "")
MODO_ENVIO  = os.getenv("MODO_ENVIO", "delegado").lower()

DELAY_ENTRE_ENVIOS = 3
TOKEN_CACHE_PATH   = "data/cache/token_cache.json"
SCOPES             = ["Mail.Send"]


# ─────────────────────────────────────────────
# Template
# ─────────────────────────────────────────────

def _assunto(row: dict, teste: bool) -> str:
    base = f"Solicitação de Número de Matrícula – Imóvel Rural | NIRF {row['nirf_crf']}"
    return f"[TESTE] {base}" if teste else base


def _corpo_html(row: dict) -> str:
    nome_cartorio = row.get('cartorio_nome', '').strip()
    saudacao = (
        f"Prezado(a) Oficial de Registro de Imóveis – {nome_cartorio},"
        if nome_cartorio else
        "Prezado(a) Oficial de Registro de Imóveis,"
    )

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

  <table style="border-collapse:collapse;width:auto;min-width:420px;margin:14px 0;font-size:13px;">
    <thead>
      <tr style="background-color:#1F4E79;color:white;">
        <th style="padding:6px 14px;text-align:left;font-weight:600;">Campo</th>
        <th style="padding:6px 14px;text-align:left;font-weight:600;">Informação</th>
      </tr>
    </thead>
    <tbody>
      <tr style="background-color:#f2f7fc;">
        <td style="padding:5px 14px;border-bottom:1px solid #e0e0e0;color:#555;">Denominação</td>
        <td style="padding:5px 14px;border-bottom:1px solid #e0e0e0;">{row.get('denominacao','—')}</td>
      </tr>
      <tr>
        <td style="padding:5px 14px;border-bottom:1px solid #e0e0e0;color:#555;">Município / UF</td>
        <td style="padding:5px 14px;border-bottom:1px solid #e0e0e0;">{row.get('municipio','—')} – {row.get('uf_sigla','—')}</td>
      </tr>
      <tr style="background-color:#f2f7fc;">
        <td style="padding:5px 14px;border-bottom:1px solid #e0e0e0;color:#555;">Comarca</td>
        <td style="padding:5px 14px;border-bottom:1px solid #e0e0e0;">{row.get('comarca','—')}</td>
      </tr>
      <tr>
        <td style="padding:5px 14px;border-bottom:1px solid #e0e0e0;color:#555;">Código NIRF / CRF</td>
        <td style="padding:5px 14px;border-bottom:1px solid #e0e0e0;"><strong>{row.get('nirf_crf','—')}</strong></td>
      </tr>
      <tr style="background-color:#f2f7fc;">
        <td style="padding:5px 14px;color:#555;">Titular</td>
        <td style="padding:5px 14px;">{row.get('titular','—')}</td>
      </tr>
    </tbody>
  </table>

  <div style="background:#f0f5fb;border-left:4px solid #1F4E79;padding:12px 16px;margin:18px 0;border-radius:0 6px 6px 0;">
    <p style="margin:0 0 6px 0;">
      Solicitamos, por gentileza, a confirmação do <strong>número de matrícula</strong>
      desse imóvel registrado nessa Serventia, ou, caso não conste em seus registros,
      que nos informe para que possamos buscar a serventia competente.
    </p>
    <p style="margin:0;">
      Desde já agradecemos a colaboração e nos colocamos à disposição para quaisquer esclarecimentos.
    </p>
  </div>

  <p style="margin-top:20px;">Atenciosamente,</p>
  <p style="margin:0;">
    <strong>Felipe Serra Silva</strong><br>
    Hayden Capital<br>
    <a href="mailto:{REMETENTE}" style="color:#1F4E79;">{REMETENTE}</a>
  </p>

  <hr style="border:none;border-top:1px solid #e5e5e5;margin-top:28px;">
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
        cache.deserialize(open(TOKEN_CACHE_PATH).read())
    return cache


def _salvar_cache(cache: msal.SerializableTokenCache):
    os.makedirs(os.path.dirname(TOKEN_CACHE_PATH), exist_ok=True)
    if cache.has_state_changed:
        open(TOKEN_CACHE_PATH, "w").write(cache.serialize())


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
                          _assunto(rd, modo_teste), _corpo_html(rd), modo_teste)
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
