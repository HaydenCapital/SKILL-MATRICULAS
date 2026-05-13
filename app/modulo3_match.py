import pandas as pd
import unicodedata
from rapidfuzz import fuzz, process

FUZZY_THRESHOLD = 90
OVERRIDE_PATH = "data/overrides/municipio_ri_override.csv"


def _normalizar(texto):
    if not isinstance(texto, str):
        return ""
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto.strip().upper()


def _carregar_overrides() -> pd.DataFrame:
    try:
        df = pd.read_csv(OVERRIDE_PATH)
        df["municipio_norm"] = df["municipio_norm"].apply(_normalizar)
        df["uf_sigla"] = df["uf_sigla"].apply(_normalizar)
        df["municipio_ri_norm"] = df["municipio_ri_norm"].apply(_normalizar)
        return df
    except FileNotFoundError:
        return pd.DataFrame(columns=["municipio_norm", "uf_sigla", "municipio_ri_norm"])


def _match_override(row: pd.Series, df_cnj: pd.DataFrame, df_overrides: pd.DataFrame) -> pd.DataFrame:
    """Tentativa 0: verifica se o município tem override manual de RI competente."""
    hit = df_overrides[
        (df_overrides["municipio_norm"] == row["municipio_norm"]) &
        (df_overrides["uf_sigla"] == row["uf_sigla"])
    ]
    if hit.empty:
        # tenta também pela comarca
        hit = df_overrides[
            (df_overrides["municipio_norm"] == row["comarca_norm"]) &
            (df_overrides["uf_sigla"] == row["uf_sigla"])
        ]
    if hit.empty:
        return pd.DataFrame()

    municipio_ri = hit.iloc[0]["municipio_ri_norm"]
    return df_cnj[
        (df_cnj["municipio_norm"] == municipio_ri) &
        (df_cnj["uf_norm"] == row["uf_sigla"])
    ]


def _match_exato(row: pd.Series, df_cnj: pd.DataFrame) -> pd.DataFrame:
    """Tentativa 1: municipio_norm do Excel == municipio_norm do CNJ + uf_sigla."""
    return df_cnj[
        (df_cnj["municipio_norm"] == row["municipio_norm"]) &
        (df_cnj["uf_norm"] == row["uf_sigla"])
    ]


def _match_comarca(row: pd.Series, df_cnj: pd.DataFrame) -> pd.DataFrame:
    """Tentativa 2: comarca_norm do Excel == municipio_norm do CNJ + uf_sigla."""
    return df_cnj[
        (df_cnj["municipio_norm"] == row["comarca_norm"]) &
        (df_cnj["uf_norm"] == row["uf_sigla"])
    ]


def _match_fuzzy(row: pd.Series, df_cnj: pd.DataFrame) -> pd.DataFrame:
    """Tentativa 3: rapidfuzz entre comarca_norm e municipio_norm do CNJ, mesma UF."""
    candidatos = df_cnj[df_cnj["uf_norm"] == row["uf_sigla"]].copy()
    if candidatos.empty:
        return pd.DataFrame()

    municipios = candidatos["municipio_norm"].tolist()
    resultado = process.extractOne(
        row["comarca_norm"],
        municipios,
        scorer=fuzz.token_sort_ratio,
        score_cutoff=FUZZY_THRESHOLD,
    )
    if resultado is None:
        return pd.DataFrame()

    melhor_municipio = resultado[0]
    return candidatos[candidatos["municipio_norm"] == melhor_municipio]


def cruzar(df_imoveis: pd.DataFrame, df_cnj: pd.DataFrame) -> pd.DataFrame:
    """
    Para cada imóvel do Excel, encontra os cartórios correspondentes no CNJ.
    Retorna DataFrame com colunas do imóvel + colunas do cartório + metadados do match.

    Cascata de tentativas:
      0. Override manual (municipio_ri_override.csv) — para cidades sem RI próprio
      1. municipio_norm exato + uf_sigla
      2. comarca_norm como município + uf_sigla
      3. fuzzy (rapidfuzz token_sort_ratio >= 90%) na comarca vs municipio CNJ
    """
    df_cnj = df_cnj.copy()
    df_cnj["uf_norm"] = df_cnj["uf_norm"].str.strip().str.upper()
    df_overrides = _carregar_overrides()

    resultados = []

    for _, imovel in df_imoveis.iterrows():
        # Tentativa 0 — override manual
        encontrados = _match_override(imovel, df_cnj, df_overrides)
        metodo = "override_manual"

        # Tentativa 1 — municipio exato
        if encontrados.empty:
            encontrados = _match_exato(imovel, df_cnj)
            metodo = "municipio_exato"

        # Tentativa 2 — comarca como município
        if encontrados.empty:
            encontrados = _match_comarca(imovel, df_cnj)
            metodo = "comarca_como_municipio"

        # Tentativa 3 — fuzzy na comarca
        if encontrados.empty:
            encontrados = _match_fuzzy(imovel, df_cnj)
            metodo = "fuzzy_comarca"

        if encontrados.empty:
            resultados.append({
                **imovel.to_dict(),
                "cartorio_nome": None,
                "cartorio_email": None,
                "cartorio_municipio": None,
                "cartorio_cns": None,
                "match_metodo": "NAO_ENCONTRADO",
            })
        else:
            for _, cartorio in encontrados.iterrows():
                resultados.append({
                    **imovel.to_dict(),
                    "cartorio_nome": cartorio.get("nome"),
                    "cartorio_email": cartorio.get("email"),
                    "cartorio_municipio": cartorio.get("municipio"),
                    "cartorio_cns": cartorio.get("cns"),
                    "match_metodo": metodo,
                })

    return pd.DataFrame(resultados)


if __name__ == "__main__":
    import sys
    import os
    sys.path.insert(0, os.path.dirname(__file__))

    from modulo1_leitura import carregar_imoveis
    from modulo2_cartorios import carregar_cartorios

    df_imoveis = carregar_imoveis("data/input/Pesquisa de Bens - Antonio Francischini.xls")
    df_cnj = carregar_cartorios()

    df_resultado = cruzar(df_imoveis, df_cnj)

    print(f"\n{'='*60}")
    print(f"Total de linhas resultado: {len(df_resultado)}")
    print(f"\nDistribuicao por metodo de match:")
    print(df_resultado["match_metodo"].value_counts().to_string())
    print(f"\nResultado completo:")
    cols = ["nirf_crf", "denominacao", "comarca", "uf_sigla",
            "cartorio_nome", "cartorio_email", "match_metodo"]
    print(df_resultado[[c for c in cols if c in df_resultado.columns]].to_string())
