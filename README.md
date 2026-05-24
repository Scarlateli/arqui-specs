# Arqui Specs

Assistente de chat que monta planilhas de especificações técnicas (memorial descritivo) para projetos de arquitetura de interiores. Você descreve os ambientes em linguagem natural, ele organiza tudo num `.xlsx` formatado e pronto pra entregar ao cliente.

## O que faz

Você manda mensagens curtas descrevendo os itens de cada ambiente — "cozinha: cuba Tramontina inox 50x40, torneira Docol gourmet preta" — e o assistente:

- Identifica o ambiente
- Busca as especificações reais do produto na internet (dimensões, voltagem, características)
- Preenche a planilha com marca, modelo, acabamento e observações técnicas (pontos hidráulicos, elétricos, alturas de instalação)
- Marca como `(definir)` o que ainda precisa de decisão
- Mantém a formatação profissional: cabeçalho do projeto, separadores por ambiente, bordas, larguras ajustadas

Você baixa o `.xlsx` a qualquer momento.

## Como rodar localmente

Precisa de Python 3.11+ e uma chave da API da Anthropic.

```bash
git clone https://github.com/Scarlateli/arqui-specs.git
cd arqui-specs
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Crie o arquivo `.streamlit/secrets.toml`:

```toml
ANTHROPIC_API_KEY = "sua-chave-aqui"
APP_PASSWORD = "sua-senha-aqui"
```

Rode:

```bash
streamlit run app.py
```

## Stack

Python, Streamlit para a interface, Anthropic SDK para o Claude, openpyxl para manipular o Excel. Deploy no Streamlit Community Cloud.

## Segurança

A chave da API e a senha do app ficam em `secrets.toml` (ignorado pelo Git) ou no painel do Streamlit Cloud quando em produção. O acesso ao app é protegido por senha. O limite de gasto da API está topado em USD 25/mês no console da Anthropic.

Se a chave vazar:

1. Revogar em https://console.anthropic.com/settings/keys
2. Gerar nova chave
3. Atualizar o secret no Streamlit Cloud
4. Verificar o histórico do Git: `git log -p | grep -i "sk-ant"`

## Estrutura

```
arqui-specs/
├── .gitignore
├── .streamlit/
│   └── secrets.toml        # não versionado
├── app.py                  # entrada do Streamlit
├── claude_client.py        # comunicação com a API
├── excel_handler.py        # leitura, escrita e formatação do .xlsx
├── prompts/
│   └── system_prompt.md    # instruções base para o Claude
├── requirements.txt
└── LICENSE
```

## Licença

MIT.
