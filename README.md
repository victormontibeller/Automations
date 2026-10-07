# Automations

Repositório de automações pessoais. Cada automação fica em uma pasta própria
na raiz do repositório, com código, dependências, configuração, testes e dados
de execução independentes.

## Automações disponíveis

| Automação | Finalidade | Documentação |
| --- | --- | --- |
| Cartões | Gerar e enviar pelo Gmail os resumos individuais dos cartões Black e Latam, com execução pelo Hermes. | [Instalação e uso](cartoes/README.md) |

## Estrutura

```text
Automations/
├── README.md
├── .gitignore
├── libs/
│   └── automation_core/  # installable, application-independent Google clients
└── cartoes/
    ├── README.md
    ├── pyproject.toml
    ├── cartoes.py    # compatibility CLI shim
    ├── src/
    │   ├── cartoes.py  # installed legacy main import
    │   └── resumos_cartoes/
    ├── config.example.json
    ├── .env.example  # nota de migração; não copiar para .env
    ├── hermes/
    ├── tests/
    ├── inputs/       # arquivos opcionais no modo local
    ├── outputs/      # prévias HTML e texto
    └── var/          # histórico e trava; criado durante o uso
```

Planilhas, prévias, histórico, credenciais e ambientes Python são locais e ficam
fora do Git. Os arquivos de exemplo de configuração fazem parte do repositório.

## Executar a automação dos cartões

A partir da raiz:

```bash
cd cartoes
python3 -m venv .venv
.venv/bin/python -m pip install ../libs/automation_core .
cp config.example.json config.json
# configure owner_name, recipients, personal_copy_email and payment_footer in config.json
.venv/bin/python cartoes.py --config config.json --card black --month 2026-09
.venv/bin/python cartoes.py --config config.json --card latam --month 2026-09
```

Por padrão, a planilha deve estar no Google Drive em
`Meu Drive/Cartão/2026/2026-09.xlsx`; o ano e o nome esperado acompanham o mês
solicitado. Para desenvolvimento local, configure `input_source: "local"` e
`input_dir`. Para configurar contatos, Gmail e agendamentos, siga o
[guia dos cartões](cartoes/README.md). A planilha pessoal não acompanha
um clone do repositório.

## Shared integrations

`libs/automation_core/` is the installable `automation-core` package. It provides
explicit-path Google authentication, caller-composed Gmail transport, and exact
Drive lookups/downloads. It does not depend on cards, Excel, Hermes, schedules,
recipient policy or the delivery ledger. See its [API and test guide](libs/automation_core/README.md).

The shared package is **local to this repository, not published on PyPI**. From
`cartoes/`, install both local packages together:

```bash
.venv/bin/python -m pip install ../libs/automation_core .
.venv/bin/python -m pip check
.venv/bin/python -m unittest discover -s ../libs/automation_core/tests -v
.venv/bin/python -m unittest discover -s tests -v
```

Keep a separate `.venv/` per automation. Reuse means installing the same shared
library into each application's own environment, not sharing a virtualenv or
changing `sys.path`. Regular installs must be reinstalled after source edits.
Development installations do not update production launchers or scheduled jobs.

## Adicionar outra automação

1. Crie `<nome>/`, com um nome curto e descritivo.
2. Adicione um `README.md` com finalidade, configuração, execução e testes.
3. Defina as dependências dentro dessa pasta; em Python, use seu próprio
   `pyproject.toml` e `.venv/`.
4. Mantenha configurações e credenciais próprias, com exemplos sem dados pessoais
   ou senhas. Use `inputs/`, `outputs/` e `var/` quando precisar dessas pastas.
5. Coloque os lançadores em `hermes/` quando a execução usar o Hermes, com o
   caminho de configuração e o interpretador dessa automação.
6. Acrescente a automação à tabela acima.

Execute instalações e testes dentro da pasta da automação correspondente. Cada
agendamento deve usar o ambiente e o histórico da sua automação.
