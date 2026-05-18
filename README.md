# Skill Matrículas — Hayden Capital

Automação para identificar o Cartório de Registro de Imóveis competente para imóveis rurais e solicitar o número de matrícula via e-mail.

> **Repositório público.** O código não contém credenciais — todas as chaves de API e senhas são configuradas via variáveis de ambiente (`.env` local ou Secrets no Streamlit Cloud) e nunca são commitadas.

---

## Acesso ao dashboard

O sistema está disponível em:

**[hayden-matriculas.streamlit.app](https://hayden-matriculas.streamlit.app)**

Para solicitar acesso, entre em contato com o administrador do repositório.

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
│   ├── cache/
│   │   └── cartorios_cnj.csv     # Cache da base CNJ (evita re-download no deploy)
│   ├── input/                    # Coloque aqui os Excels de imóveis (não commitado)
│   ├── output/                   # Logs e resultados gerados (não commitado)
│   └── overrides/
│       ├── municipio_ri_override.csv          # Mapeamento município → RI (665 entradas)
│       └── municipios_sem_ri_para_revisao.xlsx # Planilha de revisão por estado
├── dashboard.py                  # Interface Streamlit
├── main.py                       # CLI (linha de comando)
├── gerar_revisao_ri.py           # Gera planilha de sugestões por distância geográfica
├── importar_confirmacoes_ri.py   # Importa confirmações para o override CSV
├── .env                          # Credenciais (NUNCA commitar — está no .gitignore)
├── .env.example                  # Modelo de configuração sem dados sensíveis
└── requirements.txt
```

---

## Configuração para rodar localmente

### 1. Instalar dependências

```bash
pip install -r requirements.txt
```

### 2. Configurar credenciais

Copie `.env.example` para `.env` e preencha com as credenciais do Azure:

```env
GRAPH_TENANT_ID=SEU_TENANT_ID
GRAPH_CLIENT_ID=SEU_CLIENT_ID
GRAPH_CLIENT_SECRET=SEU_CLIENT_SECRET

# Remetente produção
EMAIL_REMETENTE=luiza@haydencapital.com.br
EMAIL_NOME_REMETENTE=Luiza | Hayden Capital

# Remetente teste (e-mails de teste chegam aqui)
EMAIL_REMETENTE_TESTE=felipess@haydencapital.com.br
EMAIL_NOME_REMETENTE_TESTE=Felipe Serra Silva | Hayden Capital
EMAIL_TESTE=felipess@haydencapital.com.br
```

As credenciais do Azure (Tenant ID, Client ID, Client Secret) estão registradas no App Registration **"Hayden Matriculas Bot"** no portal Azure da empresa.

### 3. Rodar o dashboard

```bash
python -m streamlit run dashboard.py
```

Acesse **http://localhost:8501** no navegador.

---

## Como usar o dashboard

**Aba 1 — Resultado do Cruzamento**
- Faça upload do Excel de imóveis (aba `Imóveis Rurais Nacional`)
- O sistema cruza com a base do CNJ e exibe o cartório identificado para cada imóvel
- A coluna **CRI Indicado** mostra o município do RI para imóveis mapeados via override manual

**Aba 2 — Envio de E-mails**

| Modo | Comportamento |
|------|--------------|
| **Teste** | Todos os e-mails chegam no inbox de teste para conferir o template antes do envio real |
| **Producao** | E-mails enviados diretamente para os cartórios (requer confirmação via checkbox) |

- No primeiro uso da sessão, o sistema solicita login Microsoft via QR code (Device Code Flow)
- O progresso é exibido em tempo real com status por cartório

**Aba 3 — Municípios sem RI**
- Lista municípios do Excel que não tiveram cartório identificado automaticamente
- Permite informar manualmente o RI competente e re-aplicar o cruzamento
- Inclui a planilha de revisão com 663 municípios mapeados por distância geográfica — é possível pesquisar por nome e corrigir sugestões incorretas
- Botão para baixar o CSV atualizado e commitar no repositório

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
| `override_manual` | Mapeamento manual em `municipio_ri_override.csv` |
| `municipio_exato` | Município do imóvel == município do cartório (mesma UF) |
| `comarca_como_municipio` | Comarca do imóvel == município do cartório |
| `fuzzy_comarca` | Similaridade textual ≥ 90% (rapidfuzz) |
| `NAO_ENCONTRADO` | Nenhum cartório identificado — requer mapeamento manual |

---

## Adicionar novos municípios sem RI

Municípios sem Cartório de Registro de Imóveis próprio precisam ser mapeados para o RI competente da circunscrição. Há duas formas:

**Via dashboard (recomendado):** acesse a aba "Municípios sem RI", informe o município do RI competente e baixe o CSV atualizado para commitar.

**Via CLI:**
```bash
# Gera planilha com sugestões baseadas em distância geográfica
python gerar_revisao_ri.py --uf PR

# Após revisar/confirmar a planilha, importa para o override CSV
python importar_confirmacoes_ri.py
```

O arquivo `data/overrides/municipio_ri_override.csv` cobre **665 municípios** de MT, PR, AM e SP.

---

## Sobre o envio de e-mail

O envio usa a **Microsoft Graph API com OAuth2 (Device Code Flow)**, não SMTP — necessário porque o Microsoft 365 corporativo bloqueia SMTP AUTH por padrão.

**Como funciona:**
1. Ao clicar em enviar, o sistema exibe um código e um link (`microsoft.com/devicelogin`)
2. O usuário faz login com sua conta `@haydencapital.com.br` uma única vez por sessão
3. O token é armazenado em cache local (`data/cache/token_cache.json`) — nunca commitado
4. No Streamlit Cloud, o login é solicitado a cada nova sessão

**Configuração Azure (já realizada):**
- App Registration: **"Hayden Matriculas Bot"** — tenant Hayden Capital
- Permissão delegada: `Mail.Send`
- Tipo: Public client (Device Code Flow habilitado)
- Não requer consentimento do administrador

---

## Atualizar a base de cartórios CNJ

A base é commitada em `data/cache/cartorios_cnj.csv` para evitar re-download no deploy. Para atualizar:

```bash
python -c "from app.modulo2_cartorios import carregar_cartorios; carregar_cartorios(forcar_download=True)"
git add data/cache/cartorios_cnj.csv
git commit -m "feat: atualiza base de cartorios CNJ"
git push
```

Recomenda-se atualizar mensalmente ou quando identificar cartórios desatualizados.

---

## Dependências principais

| Biblioteca | Uso |
|---|---|
| `pandas` / `openpyxl` / `xlrd` | Leitura de Excel |
| `requests` | Consulta à API CNJ |
| `msal` | Autenticação OAuth2 Microsoft |
| `rapidfuzz` | Match fuzzy de nomes de municípios |
| `streamlit` | Dashboard web |
| `python-dotenv` | Configuração via `.env` |
| `unidecode` | Normalização de texto para comparação |
