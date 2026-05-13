"""
Lê o Excel de revisão preenchido pelo usuário e atualiza o override CSV.

Regra:
  - Se "RI Confirmado" está preenchido → usa o valor confirmado
  - Se "RI Confirmado" está vazio      → usa o RI Sugerido (aceito como correto)
  - Linhas com "SEM COORDENADA" sem confirmação → ignoradas (aviso)
"""

import pandas as pd
import unicodedata
import os

REVISAO_PATH  = 'data/overrides/municipios_sem_ri_para_revisao.xlsx'
OVERRIDE_PATH = 'data/overrides/municipio_ri_override.csv'

def norm(t):
    if not isinstance(t, str): return ''
    t = unicodedata.normalize('NFKD', t)
    t = ''.join(c for c in t if not unicodedata.combining(c))
    return t.strip().upper()

def importar():
    df = pd.read_excel(REVISAO_PATH, sheet_name='Para Revisar', header=1, engine='openpyxl')
    df.columns = ['municipio', 'uf', 'ri_sugerido', 'distancia_km', 'ri_confirmado', 'status']
    df = df.dropna(subset=['municipio', 'uf'])

    novas = []
    ignoradas = []

    for _, row in df.iterrows():
        confirmado = str(row.get('ri_confirmado', '')).strip()
        sugerido   = str(row.get('ri_sugerido',   '')).strip()
        municipio  = str(row['municipio']).strip()
        uf         = str(row['uf']).strip()

        ri_final = confirmado if confirmado and confirmado.lower() not in ('nan', '') else sugerido

        if not ri_final or ri_final == 'SEM COORDENADA':
            ignoradas.append(f"  {municipio}-{uf}: sem RI definido")
            continue

        novas.append({
            'municipio_norm':   norm(municipio),
            'uf_sigla':         norm(uf),
            'municipio_ri_norm': norm(ri_final),
            'observacao':       f'Importado da revisao - sugestao distancia ({ri_final})'
                                if not confirmado or confirmado.lower() == 'nan'
                                else f'Confirmado manualmente ({ri_final})'
        })

    df_novas = pd.DataFrame(novas).drop_duplicates(subset=['municipio_norm', 'uf_sigla'])

    # Carrega override existente e mescla sem duplicar
    if os.path.exists(OVERRIDE_PATH):
        df_atual = pd.read_csv(OVERRIDE_PATH)
        df_atual['municipio_norm'] = df_atual['municipio_norm'].apply(norm)
        df_atual['uf_sigla']       = df_atual['uf_sigla'].apply(norm)
        chaves_existentes = set(df_atual['municipio_norm'] + '|' + df_atual['uf_sigla'])
        df_novas = df_novas[~(df_novas['municipio_norm'] + '|' + df_novas['uf_sigla']).isin(chaves_existentes)]
        df_final = pd.concat([df_atual, df_novas], ignore_index=True)
    else:
        df_final = df_novas

    df_final.to_csv(OVERRIDE_PATH, index=False)

    print(f"Override atualizado: {OVERRIDE_PATH}")
    print(f"  Entradas novas adicionadas : {len(df_novas)}")
    print(f"  Total no arquivo agora     : {len(df_final)}")

    if ignoradas:
        print(f"\nIgnorados por falta de RI definido ({len(ignoradas)}):")
        for msg in ignoradas:
            print(msg)

if __name__ == '__main__':
    importar()
