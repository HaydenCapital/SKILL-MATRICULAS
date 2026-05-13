import pandas as pd
import unicodedata

UF_NOME_PARA_SIGLA = {
    "ACRE": "AC", "ALAGOAS": "AL", "AMAPA": "AP", "AMAZONAS": "AM",
    "BAHIA": "BA", "CEARA": "CE", "DISTRITO FEDERAL": "DF",
    "ESPIRITO SANTO": "ES", "GOIAS": "GO", "MARANHAO": "MA",
    "MATO GROSSO": "MT", "MATO GROSSO DO SUL": "MS", "MINAS GERAIS": "MG",
    "PARA": "PA", "PARAIBA": "PB", "PARANA": "PR", "PERNAMBUCO": "PE",
    "PIAUI": "PI", "RIO DE JANEIRO": "RJ", "RIO GRANDE DO NORTE": "RN",
    "RIO GRANDE DO SUL": "RS", "RONDONIA": "RO", "RORAIMA": "RR",
    "SANTA CATARINA": "SC", "SAO PAULO": "SP", "SERGIPE": "SE",
    "TOCANTINS": "TO",
}

# Mapeamento posicional para resistir a encoding corrompido
# Col: 1=Titular, 2=CPF/CNPJ, 3=Código, 4=Denominação, 5=Área, 6=Município, 7=Comarca, 8=UF
COL_MAP_POSICIONAL = {
    1: "titular",
    2: "cpf_cnpj",
    3: "nirf_crf",
    4: "denominacao",
    5: "area",
    6: "municipio",
    7: "comarca",
    8: "uf",
}

COL_MAP_NOME = {
    "Titular": "titular",
    "CPF/CNPJ": "cpf_cnpj",
    "Código": "nirf_crf",
    "Denominação": "denominacao",
    "Área": "area",
    "Município": "municipio",
    "Comarca Imóvel Registrado": "comarca",
    "UF": "uf",
}

def normalizar_texto(texto):
    if not isinstance(texto, str):
        return ""
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto.strip().upper()

def _renomear_colunas(df: pd.DataFrame) -> pd.DataFrame:
    """Tenta renomear por nome; se falhar para alguma coluna, usa posição."""
    colunas_presentes = set(df.columns)
    colunas_esperadas = set(COL_MAP_NOME.keys())

    if colunas_esperadas.issubset(colunas_presentes):
        return df.rename(columns=COL_MAP_NOME)

    # Fallback posicional
    rename_pos = {df.columns[i]: nome for i, nome in COL_MAP_POSICIONAL.items() if i < len(df.columns)}
    return df.rename(columns=rename_pos)

def _uf_para_sigla(uf_valor: str) -> str:
    """Converte nome completo ou sigla de UF para sigla de 2 letras."""
    v = normalizar_texto(str(uf_valor))
    if v in UF_NOME_PARA_SIGLA:
        return UF_NOME_PARA_SIGLA[v]
    if len(v) == 2:
        return v  # já é sigla
    return v

def carregar_imoveis(caminho_excel: str) -> pd.DataFrame:
    df = pd.read_excel(caminho_excel, sheet_name="Imóveis Rurais Nacional", engine="openpyxl")
    df = _renomear_colunas(df)
    df = df.dropna(subset=["nirf_crf"])

    df["municipio_norm"] = df["municipio"].apply(normalizar_texto)
    df["comarca_norm"]   = df["comarca"].apply(normalizar_texto)
    df["uf_sigla"]       = df["uf"].apply(_uf_para_sigla)

    colunas_saida = ["nirf_crf", "titular", "cpf_cnpj", "denominacao",
                     "municipio", "comarca", "uf", "uf_sigla",
                     "municipio_norm", "comarca_norm"]
    return df[[c for c in colunas_saida if c in df.columns]]

if __name__ == "__main__":
    df = carregar_imoveis("data/input/Pesquisa de Bens - Antonio Francischini.xls")
    print(f"{len(df)} imoveis carregados")
    print(df[["nirf_crf", "municipio", "comarca", "uf", "uf_sigla"]].to_string())