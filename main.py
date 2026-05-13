"""
Uso:
    python main.py "data/input/NomeDoArquivo.xls"

Flags opcionais:
    --real        Dispara para os cartórios reais (padrão: modo teste)
    --so-match    Só faz o match, sem enviar e-mail (útil para revisar antes)
"""

import sys
import os
import argparse
import pandas as pd
from datetime import datetime

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "app"))

from modulo1_leitura import carregar_imoveis
from modulo2_cartorios import carregar_cartorios
from modulo3_match import cruzar
from modulo4_email import disparar


def main():
    parser = argparse.ArgumentParser(description="Automação de consulta de matrículas — Hayden Capital")
    parser.add_argument("excel", help="Caminho para o arquivo Excel de imóveis")
    parser.add_argument("--real",     action="store_true", help="Dispara para os cartórios reais (sem --real roda em modo teste)")
    parser.add_argument("--so-match", action="store_true", help="Só faz o match e salva resultado, sem enviar e-mail")
    args = parser.parse_args()

    if not os.path.exists(args.excel):
        print(f"Arquivo não encontrado: {args.excel}")
        sys.exit(1)

    modo_teste = not args.real
    nome_base  = os.path.splitext(os.path.basename(args.excel))[0]
    timestamp  = datetime.now().strftime("%Y%m%d_%H%M%S")
    os.makedirs("data/output", exist_ok=True)

    # ── Etapa 1: Leitura do Excel ──────────────────────────────
    print(f"\n[1/4] Lendo: {args.excel}")
    df_imoveis = carregar_imoveis(args.excel)
    print(f"      {len(df_imoveis)} imóveis carregados")

    # ── Etapa 2: Cartórios CNJ ─────────────────────────────────
    print("\n[2/4] Carregando cartórios CNJ...")
    df_cnj = carregar_cartorios()

    # ── Etapa 3: Match ─────────────────────────────────────────
    print("\n[3/4] Cruzando dados...")
    df_resultado = cruzar(df_imoveis, df_cnj)

    print(f"\n      Resultado do match:")
    print(df_resultado["match_metodo"].value_counts().to_string())

    # Salva resultado do match sempre
    match_path = f"data/output/{nome_base}_{timestamp}_match.xlsx"
    cols_match = ["nirf_crf", "denominacao", "municipio", "comarca", "uf_sigla",
                  "cartorio_nome", "cartorio_email", "cartorio_municipio", "match_metodo"]
    df_resultado[[c for c in cols_match if c in df_resultado.columns]].to_excel(match_path, index=False)
    print(f"\n      Match salvo: {match_path}")

    if args.so_match:
        print("\nModo --so-match: encerrado sem enviar e-mails.")
        return

    # ── Etapa 4: Disparo de e-mails ────────────────────────────
    sem_email = df_resultado[df_resultado["cartorio_email"].isna() | (df_resultado["cartorio_email"] == "")]
    if not sem_email.empty:
        print(f"\n⚠  {len(sem_email)} linha(s) sem e-mail de cartório — serão ignoradas:")
        print(sem_email[["nirf_crf", "cartorio_nome", "match_metodo"]].to_string(index=False))

    print(f"\n[4/4] Disparando e-mails ({'TESTE' if modo_teste else 'PRODUÇÃO'})...")
    df_log = disparar(df_resultado, modo_teste=modo_teste)

    if df_log.empty:
        print("Nenhum e-mail enviado.")
        return

    # Salva log
    log_path = f"data/output/{nome_base}_{timestamp}_log.xlsx"
    df_log.to_excel(log_path, index=False)

    print(f"\n{'='*60}")
    print(f"CONCLUÍDO")
    print(f"  Arquivo    : {args.excel}")
    print(f"  Match      : {match_path}")
    print(f"  Log envios : {log_path}")
    print(f"\nResumo de envios:")
    print(df_log["status_envio"].value_counts().to_string())
    print("="*60)


if __name__ == "__main__":
    main()
