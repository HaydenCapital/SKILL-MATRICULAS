import requests
import pandas as pd
import unicodedata
import os
import time

BASE_URL = "https://justicaabertaapi.cnj.jus.br/v1/api"
CACHE_PATH = "data/cache/cartorios_cnj.csv"

UFS = ["AC","AL","AP","AM","BA","CE","DF","ES","GO","MA","MT","MS",
       "MG","PA","PB","PR","PE","PI","RJ","RN","RS","RO","RR","SC",
       "SP","SE","TO"]

HEADERS = {
    "Content-Type": "application/json",
    "Accept": "*/*",
    "Origin": "https://justicaaberta.cnj.jus.br",
    "Referer": "https://justicaaberta.cnj.jus.br/",
}

def normalizar_texto(texto):
    if not isinstance(texto, str):
        return ""
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto.strip().upper()

def buscar_cartorios_ri_por_cidade(cidade_id: int) -> list:
    """Busca cartórios de Registro de Imóveis (atribuição 4) por cidade via POST."""
    todos = []
    pagina = 1

    while True:
        params = {
            "assignments": "[4]",
            "page": pagina,
            "perPage": 50,
            "search": ""
        }
        body = {"cidade_id": cidade_id}

        try:
            r = requests.post(
                f"{BASE_URL}/serventias",
                params=params,
                json=body,
                headers=HEADERS,
                timeout=15
            )
            if r.status_code != 200:
                break

            data = r.json()
            itens = data.get("data", [])
            todos.extend(itens)

            meta = data.get("meta", {})
            total = meta.get("total", 0)
            if len(todos) >= total or not itens:
                break

            pagina += 1
            time.sleep(0.1)

        except Exception:
            break

    return todos

def baixar_cartorios_cnj() -> pd.DataFrame:
    print("Baixando cartórios de Registro de Imóveis do CNJ (API Justiça Aberta)...")
    todos = []

    for uf in UFS:
        print(f"  {uf}", end="", flush=True)
        try:
            r = requests.get(f"{BASE_URL}/cidades/listar/{uf}", timeout=15)
            r.raise_for_status()
            cidades = r.json()

            count_uf = 0
            for cidade in cidades:
                cartorios = buscar_cartorios_ri_por_cidade(cidade["id"])
                for c in cartorios:
                    todos.append({
                        "cns": c.get("cns", ""),
                        "nome": c.get("denominacao_fantasia", "") or c.get("denominacao_padrao", ""),
                        "municipio": cidade["nome"],
                        "uf": uf,
                        "email": c.get("email", ""),
                        "telefone": c.get("telefone", ""),
                        "endereco": c.get("endereco", ""),
                        "bairro": c.get("bairro", ""),
                        "cep": c.get("cep", ""),
                        "status": c.get("status", ""),
                        "natureza": c.get("natureza", ""),
                    })
                    count_uf += 1
                time.sleep(0.1)

            print(f" -> {count_uf} cartórios")

        except Exception as e:
            print(f" -> ERRO: {e}")

        time.sleep(0.2)

    df = pd.DataFrame(todos)
    print(f"\nTotal nacional: {len(df)} cartórios de Registro de Imóveis")
    return df

def carregar_cartorios(forcar_download=False) -> pd.DataFrame:
    if not forcar_download and os.path.exists(CACHE_PATH):
        print(f"Cache encontrado: {CACHE_PATH}")
        df = pd.read_csv(CACHE_PATH)
        print(f"  {len(df)} cartórios carregados")
    else:
        os.makedirs("data/cache", exist_ok=True)
        df = baixar_cartorios_cnj()
        df.to_csv(CACHE_PATH, index=False)
        print(f"Cache salvo: {CACHE_PATH}")

    df["municipio_norm"] = df["municipio"].apply(normalizar_texto)
    df["uf_norm"] = df["uf"].apply(normalizar_texto)
    return df

if __name__ == "__main__":
    df = carregar_cartorios(forcar_download=True)
    print("\nAmostra MT:")
    print(df[df["uf"] == "MT"][["nome", "municipio", "email"]].to_string())