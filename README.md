# Skill Matrículas — Hayden Capital

Automação para identificar o Cartório de Registro de Imóveis competente para imóveis rurais e solicitar o número de matrícula via e-mail.

---

## O que faz

1. Lê uma planilha Excel com imóveis rurais (NIRF/CRF, município, comarca, UF)
2. Cruza os dados com a base de cartórios do CNJ (Justiça Aberta API)
3. Identifica o RI competente para cada imóvel — com sistema de overrides para municípios sem RI próprio
4. Dispara e-mails personalizados para cada cartório solicitando o número de matrícula
5. Gera log de envios com status por imóvel

---

## Estrutura do projeto

```
SKILL MATRICULAS/
├── app/
│   ├── modulo1_leitura.py        # Leitura e normalização do Excel
│   ├── modulo2_cartorios.py      # Download e cache da base CNJ
│   ├── modulo3_match.py          # Cruzamento imóveis × cartórios
│   └── modulo4_email.py          # Disparo de e-mails via Graph API
├── data/
│   ├── cache/                    # Cache da API CNJ e token OAuth (gerado automaticamente)
│   ├── input/                    # Coloque aqui os Excels de imóveis
│   ├── output/                   # Logs e resultados gerados
│   └── overrides/
│       └── municipio_ri_override.csv   # Mapeamento manual município → RI
├── dashboard.py                  # Interface Streamlit
├── main.py                       # CLI (linha de comando)
├── gerar_revisao_ri.py           # Gera planilha de revisão de municípios sem RI
├── importar_confirmacoes_ri.py   # Importa confirmações para o override CSV
├── .env                          # Credenciais (NUNCA commitar)
├── .env.example                  # Modelo de configuração
└── requirements.txt
```

---

## Instalação

```bash
pip install -r requirements.txt
```

### Configurar o `.env`

Copie `.env.example` para `.env` e preencha:

```env
GRAPH_TENANT_ID=SEU_TENANT_ID
GRAPH_CLIENT_ID=SEU_CLIENT_ID
GRAPH_CLIENT_SECRET=SEU_CLIENT_SECRET

EMAIL_REMETENTE=seu@email.com.br
EMAIL_NOME_REMETENTE=Nome | Empresa
EMAIL_TESTE=seu@email.com.br      # E-mail que recebe no modo teste
```

---

## Como usar — Dashboard (recomendado)

```bash
python -m streamlit run dashboard.py
```

Acesse **http://localhost:8501** no navegador.

### Passo a passo

**1. Upload da planilha**
- Arraste ou selecione o arquivo Excel (`.xls` / `.xlsx`)
- A aba deve ser `Imóveis Rurais Nacional` com as colunas padrão (NIRF, Município, Comarca, UF, etc.)

**2. Resultado do cruzamento**
- O sistema exibe quantos imóveis foram identificados e quantos cartórios têm e-mail disponível
- A tabela mostra o cartório encontrado para cada imóvel e o método de match utilizado

**3. Envio de e-mails**

| Modo | Comportamento |
|------|--------------|
| **Teste** | Todos os e-mails chegam no seu inbox (`EMAIL_TESTE`) para conferir o template |
| **Produção** | E-mails enviados diretamente para os cartórios (requer confirmação via checkbox) |

- O progresso é exibido em tempo real
- Ao final, aparece o log com status de cada envio

---

## Como usar — CLI

```bash
# Modo teste (e-mails vão para EMAIL_TESTE)
python main.py "data/input/Arquivo.xls"

# Modo produção (e-mails reais para os cartórios)
python main.py "data/input/Arquivo.xls" --real

# Só faz o match, sem enviar e-mail
python main.py "data/input/Arquivo.xls" --so-match
```

Os resultados são salvos em `data/output/` com timestamp.

---

## Métodos de match (ordem de prioridade)

| Método | Descrição |
|--------|-----------|
| `OVERRIDE` | Mapeamento manual em `municipio_ri_override.csv` |
| `EXATO` | Município do imóvel == município do cartório (mesma UF) |
| `COMARCA` | Comarca do imóvel == município do cartório |
| `FUZZY` | Similaridade textual ≥ 90% (rapidfuzz) |
| `NAO_ENCONTRADO` | Nenhum cartório identificado |

Imóveis `NAO_ENCONTRADO` são ignorados no envio. Para adicioná-los, inclua uma entrada em `data/overrides/municipio_ri_override.csv`.

---

## Adicionar novos municípios sem RI

Alguns municípios não possuem Cartório de Registro de Imóveis próprio — a competência é de outro município. Para mapeá-los:

```bash
# Gera planilha com sugestões baseadas em distância geográfica
python gerar_revisao_ri.py --uf PR

# Após revisar/confirmar a planilha, importa para o override CSV
python importar_confirmacoes_ri.py
```

O arquivo `data/overrides/municipio_ri_override.csv` atualmente cobre **665 municípios** de MT, PR, AM e SP.

---

## Sobre o envio de e-mail

> **Importante:** O envio usa a **Microsoft Graph API com OAuth2 (Device Code Flow)**, não SMTP.

Isso é necessário porque o Microsoft 365 corporativo bloqueia SMTP AUTH por padrão. O fluxo funciona assim:

1. **Primeiro uso:** ao clicar em enviar, o sistema exibe um código e um link para login no navegador. Faça o login com sua conta corporativa uma única vez.
2. **Usos seguintes:** o token fica em cache em `data/cache/token_cache.json` e é reutilizado automaticamente (sem novo login).
3. **Expiração:** se o token expirar (normalmente após 1 hora de inatividade), o login será solicitado novamente.

O cache do token **não deve ser commitado** — já está no `.gitignore`.

### Configuração no Azure (já feita)

A aplicação requer um App Registration no Azure com:
- Tipo: **Public client** (permite Device Code Flow)
- Permissão delegada: `Mail.Send` (não requer consentimento do administrador)

---

## Dependências principais

| Biblioteca | Uso |
|---|---|
| `pandas` / `openpyxl` | Leitura de Excel |
| `requests` | Consulta à API CNJ |
| `msal` | Autenticação OAuth2 Microsoft |
| `rapidfuzz` | Match fuzzy de nomes de municípios |
| `streamlit` | Dashboard web |
| `python-dotenv` | Configuração via `.env` |
