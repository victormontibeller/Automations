# Resumos individuais dos cartões

Esta automação fica em `automations/cartoes/`. Execute os comandos deste guia
dentro dessa pasta. Consulte o [catálogo do repositório](../../README.md) para
ver a organização das demais automações.

O script `cartoes.py` lê o Excel mensal e gera e-mails semelhantes às abas pessoais:
faixa azul, colunas **Data, Lançamento, Parcelas, Total e Rateio**, seguidas de
**TOTAL**, **PAGAMENTOS** e **TOTAL GERAL**. Cada amigo recebe apenas os seus
lançamentos do cartão correspondente. O valor da compra fica em **Total**; o valor
individual e o saldo ficam em **Rateio**. O total da fatura inteira não aparece
nas linhas de resumo individual.
Quando não houver pagamentos, o resumo exibe **PAGAMENTOS: R$ 0,00**.

| Cartão | Execução, horário de Brasília | Arquivo usado em outubro |
| --- | --- | --- |
| Black | Dia 5, às 09h | `2026-09.xlsx` |
| Latam | Dia 20, às 09h | `2026-09.xlsx` |

É necessário Python **3.11 ou superior**, em Linux ou macOS. Os valores são lidos
diretamente das colunas dos participantes nas abas `Black` e `Latam`; o script não
precisa abrir o Excel nem recalcular `FILTER`. A planilha original não é alterada.

### 1. Preparar o projeto

No servidor, coloque o repositório em `~/Automations`. Prepare um ambiente
Python exclusivo desta automação:

```bash
cd ~/Automations/automations/cartoes
python3 -m venv .venv
.venv/bin/python -m pip install .
cp config.example.json config.json
cp .env.example .env
chmod 600 .env
```

Edite `config.json`:

- `input_dir`: pasta em que você copiará os arquivos `AAAA-MM.xlsx`; o exemplo
  usa `inputs`, dentro da pasta desta automação.
- `output_dir`: pasta das prévias HTML e texto.
- `state_dir`: pasta do histórico SQLite e da trava de execução.
- `secrets_file`: caminho do arquivo de credenciais.
- `sender_name`: nome exibido como remetente.
- `owner_name`: seu nome, usado na apresentação “Sou o Hermes, assistente pessoal
  de…”. Preencha esse campo antes de enviar. Enquanto estiver `null`, as prévias
  mostram `[seu nome]` e o envio é bloqueado.
- `recipients`: associação entre `sheet` (aba pessoal), `participant` (cabeçalho
  da coluna na linha 2 do cartão) e `email`. Preencha os `null` com os endereços.
  Para excluir explicitamente alguém dos envios, acrescente `"enabled": false`.

Os caminhos relativos são resolvidos a partir da pasta de `config.json`, não da
pasta corrente do Hermes. O cadastro inicial corresponde às 12 abas do exemplo,
incluindo `Dora` associada a `Dora/Cae`. Colunas de pessoas sem aba pessoal não
produzem e-mails. Ao adicionar, remover ou renomear abas pessoais, atualize o
cadastro; abas desconhecidas ou ausentes interrompem a execução.

Edite `.env`, sem enviar os valores pela conversa:

```dotenv
GMAIL_USER=seu.endereco@gmail.com
GMAIL_APP_PASSWORD=senha_de_app_de_16_letras
```

`GMAIL_USER` deve ser a sua conta: ela envia os resumos e recebe, sozinha, os
avisos de erro da automação. O cadastro dos amigos nunca é usado para esses avisos.

Use uma **senha de app** da conta Google, que exige verificação em duas etapas e
disponibilidade dessa opção na conta. Não use a senha comum do Gmail.
[Instruções oficiais do Google](https://support.google.com/mail/answer/185833?hl=pt-BR).
O script aceita a senha de app com os espaços exibidos pelo Google. Conecta-se a
`smtp.gmail.com:465` com TLS e validação do certificado. Também aceita as duas
credenciais em variáveis de ambiente; elas têm precedência sobre o arquivo.

O `.env` é carregado explicitamente pelo script, pois um serviço do Hermes pode
não herdar as variáveis do seu terminal. `config.json`, `.env`, planilhas,
histórico e prévias estão excluídos do Git.

### 2. Gerar as prévias

Copie a planilha salva para `inputs/` (ou a pasta definida em `input_dir`). O nome deve ser exatamente
`AAAA-MM.xlsx`; não há seleção automática do arquivo mais recente.

```bash
.venv/bin/python cartoes.py --config config.json --card black --month 2026-09
.venv/bin/python cartoes.py --config config.json --card latam --month 2026-09
```

Sem `--send`, o script **somente gera prévias**. Você pode usar
`--config config.example.json` antes de preencher os destinatários. Abra:

- `outputs/2026-09/black/index.html`
- `outputs/2026-09/latam/index.html`

Cada índice aponta para os resumos HTML atuais. São geradas também versões `.txt`.
O HTML é o corpo completo do e-mail, incluindo a saudação do Hermes, o resumo e
a orientação para falar diretamente com você em caso de dúvidas. A apresentação
usa o mês no formato `MM/AAAA`; o nome do cartão aparece no título da tabela.
Uma nova prévia substitui os arquivos que o script gerou anteriormente no mesmo
mês e cartão. Nenhuma conexão ao Gmail é realizada e o histórico não é alterado.

Com o exemplo fornecido, são esperados **11 resumos Black e 3 Latam**. Pessoas com
pagamentos ou compras recebem resumo mesmo com saldo quitado. Quem não teve
movimento é ignorado. Créditos e estornos conservam seu sinal.

O script soma valores monetários em centavos com `Decimal`. As compras mantêm a
ordem da planilha, inclusive lançamentos repetidos e parcelas com data de compra
antiga. As seções Azul e Black da aba `Black` fazem parte do mesmo resumo, como no
exemplo. As linhas de pagamentos são somadas com o sinal registrado na planilha:
um pagamento negativo reduz o saldo; um positivo aumenta o saldo. As fórmulas
de subtotal não entram novamente nos cálculos.

Nos campos individuais dos lançamentos, valores devem ser numéricos e datas
devem ser datas do Excel. Fórmulas ou erros nos valores necessários interrompem
o lote; resultados em cache não são usados como substituto. Espera-se uma linha
`TOTAL`, pagamentos após ela e uma linha `TOTAL GERAL`, como no exemplo.

### 3. Testar o e-mail e enviar

Depois de configurar o Gmail, envie inicialmente **um resumo apenas para você**:

```bash
.venv/bin/python cartoes.py --config config.json --card latam --month 2026-09 --recipient Iago --send --test-to seu.endereco@gmail.com
```

`--test-to` troca o destinatário pelo endereço informado, acrescenta `[TESTE]` ao
assunto e não registra um envio de produção. Sem `--recipient`, todos os resumos
daquele cartão serão redirecionados a você. Confira o e-mail no Gmail do
computador e do celular, incluindo as cinco colunas e os valores de rateio.

Para enviar aos amigos:

```bash
.venv/bin/python cartoes.py --config config.json --card black --month 2026-09 --send
.venv/bin/python cartoes.py --config config.json --card latam --month 2026-09 --send
```

Um e-mail ausente ou inválido em qualquer resumo do lote impede os envios antes
da conexão SMTP. Cada mensagem tem um único destinatário, sem CC ou BCC. Não há
anexo com a planilha inteira nem conteúdo de outros amigos.

### 4. Agendar no Hermes

A configuração abaixo deve ser feita **no servidor**, após validar o teste de
e-mail. O Hermes precisa oferecer cron em modo `no_agent` e a opção
`--interpreter`. Confira `hermes cron create --help` na versão instalada.
[Documentação oficial de scripts sem IA](https://hermes-agent.nousresearch.com/docs/guides/cron-script-only).

Copie os dois lançadores para a pasta de scripts do perfil usado pelo gateway.
Não use links simbólicos para fora dessa pasta:

```bash
mkdir -p ~/.hermes/scripts
cp hermes/cartao_black.py ~/.hermes/scripts/cartao_black.py
cp hermes/cartao_latam.py ~/.hermes/scripts/cartao_latam.py
```

Os lançadores usam `~/Automations/automations/cartoes/config.json`. Se o projeto ficar em outro
local, edite o caminho padrão nos dois arquivos copiados. Também é possível
definir `CARTOES_CONFIG` no ambiente do gateway. Eles importam o módulo `cartoes`
instalado no ambiente Python próprio, sem depender do diretório corrente.

Configure o fuso do Hermes **antes de criar os agendamentos**. Essa configuração
afeta os outros horários do mesmo perfil; mantenha o fuso já existente se ele
for `America/Sao_Paulo`. Reinicie o gateway após uma mudança:

```bash
hermes config set HERMES_TIMEZONE America/Sao_Paulo
hermes gateway restart
```

[Referência de configuração de fuso do Hermes](https://hermes-agent.nousresearch.com/docs/user-guide/configuration#timezone).

Crie os agendamentos inicialmente pausados, com saída operacional local:

```bash
hermes cron create "0 9 5 * *" --no-agent --script cartao_black.py --interpreter "$HOME/Automations/automations/cartoes/.venv/bin/python" --deliver local --name "Resumo cartão Black" --paused
hermes cron create "0 9 20 * *" --no-agent --script cartao_latam.py --interpreter "$HOME/Automations/automations/cartoes/.venv/bin/python" --deliver local --name "Resumo cartão Latam" --paused
hermes cron list
hermes cron status
```

O e-mail é enviado pelo **Python**; `--deliver local` guarda somente o resultado
operacional no Hermes. Os logs não imprimem compras, valores individuais ou
credenciais. Para receber alertas no seu canal pessoal já configurado no Hermes,
você pode trocar `local` por esse canal; não use os contatos dos amigos como alvo
do log operacional.

Confira que a próxima execução será às **09h no fuso de Brasília**. Depois do
teste de Gmail e dessa conferência, substitua os identificadores pelos IDs
retornados pelo Hermes:

```bash
hermes cron resume ID_DO_BLACK
hermes cron resume ID_DO_LATAM
hermes cron doctor
```

Mantenha o gateway ativo e a pasta `var` em armazenamento persistente. Se o
Hermes estiver em contêiner, projeto, planilhas, ambiente Python e histórico
precisam estar acessíveis dentro do contêiner, com os caminhos correspondentes.

Os lançadores verificam o dia em `America/Sao_Paulo`: Black só executa no dia 5 e
Latam no dia 20. Fora do dia esperado, interrompem o envio. Para recuperar uma
execução atrasada ou com arquivo ausente, use o comando manual com `--month`;
não force um mês diferente por `--scheduled`. O teste manual pelo Hermes, com
`hermes cron run`, chama o lançador de **produção**; para testar sem enviar aos
amigos, use o comando Python com `--test-to` da etapa anterior.

### Histórico, repetição e falhas

Em execuções com `--send`, incluindo os lançadores do Hermes, uma falha interrompe
o lote e tenta enviar **um aviso somente ao proprietário**, no endereço
`GMAIL_USER`. Isso inclui arquivo ausente, Excel corrompido, dados ou destinatários
inválidos e erros operacionais. A abertura e a leitura das células são verificadas
antes dos envios aos amigos. Se o Excel falhar nessa etapa, nenhum resumo é enviado.

O aviso identifica cartão, referência, etapa e motivo, sem anexar a planilha,
resumos ou erros internos que possam expor valores e credenciais. Mesmo com
`--test-to`, o aviso de erro vai para `GMAIL_USER`. Prévias, consulta e manutenção
do histórico continuam sem envio de e-mail.

Se o erro ocorrer após o início dos envios, mensagens anteriores podem já ter
sido aceitas pelo Gmail; o aviso orienta a conferir o histórico e a pasta Enviados.
Se o Gmail, a conexão ou as credenciais também estiverem indisponíveis, o aviso
pode falhar: nesse caso o script registra a falha nos logs do Hermes, termina com
código 1 e não tenta enviar avisos recursivamente. Uma nova execução com erro pode
gerar um novo aviso; isso não altera a proteção contra resumos duplicados.

Se `config.json` não puder ser carregado, o aviso tenta usar o `.env` ao lado
desse arquivo ou as variáveis `GMAIL_USER` e `GMAIL_APP_PASSWORD` do ambiente.
Um `secrets_file` personalizado depende de uma configuração válida.

O histórico fica em `var/deliveries.sqlite3`. Cada envio é identificado por
**mês + cartão + aba pessoal**. Uma execução posterior pula mensagens aceitas
anteriormente pelo SMTP, inclusive se a planilha ou o endereço forem alterados.
O histórico também guarda endereço usado, identificador da mensagem e hash do
arquivo de entrada. Faça backup da pasta `var`: perder ou restaurar um histórico
antigo pode permitir mensagens repetidas.

```bash
.venv/bin/python cartoes.py --config config.json --history --month 2026-09
```

O comando mostra mês, cartão, aba, estado, **endereço usado no envio**, última
atualização em UTC e `Message-ID`. O endereço exibido vem do histórico, mesmo
quando o cadastro atual foi alterado.

Estados:

- `sent`: Gmail aceitou a mensagem; isso não confirma leitura ou entrega final.
- `failed`: SMTP recusou a mensagem explicitamente; o comando de envio pode ser
  repetido para tentar os destinatários ainda não enviados.
- `unknown`: houve desconexão, timeout ou interrupção após iniciar o envio. O
  lote fica bloqueado até você conferir no Gmail para evitar duplicação.

Ao conferir um envio `unknown`, procure na pasta **Enviados** por destinatário,
assunto e `Message-ID` do histórico. Se foi enviado, marque como concluído:

```bash
.venv/bin/python cartoes.py --config config.json --card latam --month 2026-09 --recipient Iago --resolve sent
```

Se confirmar que **não** foi enviado:

```bash
.venv/bin/python cartoes.py --config config.json --card latam --month 2026-09 --recipient Iago --resolve not-sent
.venv/bin/python cartoes.py --config config.json --card latam --month 2026-09 --recipient Iago --send
```

Há uma trava por pasta de histórico que impede dois processos de envio ou
manutenção ao mesmo tempo. Não há repetição automática de um resultado incerto:
SMTP e SQLite não oferecem uma transação conjunta de envio exatamente uma vez.

### Testes

```bash
.venv/bin/python -m unittest discover -s tests -v
```

A suíte usa planilhas sintéticas e SMTP simulado; nenhum teste envia e-mails reais.
A amostra fictícia versionada em `tests/fixtures/anonymous_reports.json` cobre
**24 blocos**, com totais esperados explícitos, incluindo pagamentos parciais,
quitação, créditos, parcelas antigas e participantes sem movimento. Ela permite
validar **11 mensagens Black e 3 Latam** em um clone sem arquivos pessoais.

Se o exemplo pessoal
`inputs/2026-09.xlsx` estiver na pasta, compara também os **24 blocos** com os dados
salvos das abas pessoais e os totais disponíveis. A planilha pessoal não é incluída no Git. Os testes
cobrem rateios, pagamentos, créditos, ausência de movimento, validações,
formatação, virada do ano, fuso, repetição, falhas parciais, falhas após aceite SMTP,
trava entre processos, avisos exclusivos ao proprietário e Excel corrompido na
abertura ou durante a leitura das células.

O [workflow de testes](../../.github/workflows/tests.yml) executa a suíte em pushes
e pull requests que alterem esta automação, usando Linux com Python 3.11 e 3.12
e macOS com Python 3.12. Os testes usam credenciais fictícias e dispensam secrets
do GitHub. Também é possível iniciar o workflow manualmente.

Depois de alterar o script no servidor, reinstale no ambiente usado pelo Hermes:

```bash
.venv/bin/python -m pip install --upgrade .
```

### Instalações anteriores à reorganização

Se esta automação já estava instalada na raiz do repositório ou em
`~/resumos-cartoes`, os caminhos do servidor também precisam ser ajustados:

- Transfira `config.json` e `.env` para a nova pasta da automação e atualize
  `input_dir` para `inputs` se mover as planilhas para essa pasta. Caminhos
  absolutos que continuam válidos podem ser mantidos.
- Preserve a pasta de histórico indicada por `state_dir`, incluindo
  `deliveries.sqlite3`, para manter a proteção contra envios repetidos. Se mover
  `var/`, faça isso com os processos de envio parados.
- Crie a `.venv` no novo local e instale o pacote conforme a etapa 1.
- Copie novamente os lançadores para `~/.hermes/scripts/`, confira
  `CARTOES_CONFIG` se estiver definido e ajuste o caminho do interpretador dos
  agendamentos existentes para a nova `.venv`. Mantenha apenas um agendamento
  ativo por cartão.

A reorganização local não altera uma instalação já existente no servidor.
