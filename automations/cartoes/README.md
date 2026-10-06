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
| Latam | Dia 20, às 09h | `2026-10.xlsx` |

É necessário Python **3.11 ou superior**, em Linux ou macOS. Os valores são lidos
diretamente das colunas dos participantes nas abas `Black` e `Latam`; o script não
precisa abrir o Excel nem recalcular `FILTER`. A planilha original não é alterada.

### 1. Preparar o projeto

No servidor, use o repositório em `~/Projetos/Automations`. Prepare um ambiente
Python exclusivo desta automação:

```bash
cd ~/Projetos/Automations/automations/cartoes
python3 -m venv .venv
.venv/bin/python -m pip install .
cp config.example.json config.json
```

Edite `config.json`:

- `input_source`: `drive` (padrão) busca em `Meu Drive/<drive_folder_name>/<ano>/<AAAA-MM>.xlsx`;
  `local` habilita a pasta `input_dir` apenas como alternativa local.
- `drive_folder_name`: nome exato da pasta na raiz de `Meu Drive`; o padrão é `Cartão`.
- `input_dir`: pasta local usada somente quando `input_source` for `local`; o exemplo
  define `input` (relativo à pasta do `config.json`).
- `output_dir`: pasta das prévias HTML e texto.
- `state_dir`: pasta do histórico SQLite e da trava de execução.
- `google_token_file` (opcional): caminho do token OAuth do Hermes. Por padrão,
  usa `$HERMES_HOME/google_token.json` ou `~/.hermes/google_token.json`.
- `sender_name`: nome exibido como remetente; o padrão neutro é `Hermes`. Defina
  o nome desejado apenas no `config.json` privado.
- `owner_name`: seu nome, usado na apresentação “Sou o Hermes, assistente pessoal
  do…”. Preencha esse campo antes de enviar. Enquanto estiver `null`, as prévias
  mostram `[seu nome]` e o envio é bloqueado.
- `payment_footer` (opcional): texto de uma linha anexado ao fim dos corpos HTML e
  texto. Cadastre dados de pagamento somente no `config.json` privado; no exemplo
  versionado, este campo fica `null`.
- `personal_copy_email`: endereço pessoal incluído em `Cc` em cada resumo; preencha
  antes de enviar. Os avisos também usam esse `Cc` quando o endereço está configurado.
- `recipients`: associação entre `sheet` (aba pessoal), `participant` (cabeçalho
  da coluna na linha 2 do cartão) e `email`. Preencha os `null` com os endereços.
  Para excluir explicitamente alguém dos envios, acrescente `"enabled": false`.

Os caminhos relativos são resolvidos a partir da pasta de `config.json`, não da
pasta corrente do Hermes. O exemplo contém 12 destinatários fictícios
(`Pessoa01`–`Pessoa12`). Substitua os campos `sheet` e `participant` pelos nomes
exatos da aba pessoal e do cabeçalho na planilha. Atualize o cadastro se adicionar,
remover ou renomear abas; abas desconhecidas ou ausentes interrompem a execução.

A autenticação da planilha reutiliza o mesmo token OAuth do Hermes, que precisa ter
acesso de leitura ao Drive, além dos escopos `gmail.send` e `gmail.readonly` para
os e-mails. O endereço remetente e o destino dos avisos são obtidos do perfil Gmail.

Com `input_source: "drive"`, a busca é exata: `Meu Drive/Cartão/<ano>/<AAAA-MM>.xlsx`,
onde o ano vem do mês solicitado. Se pasta ou arquivo não for encontrado, a execução
falha com aviso; não cai silenciosamente para `inputs/`. Para usar uma pasta local,
defina explicitamente `input_source: "local"` e `input_dir`.

O processo que executa a automação deve rodar com o mesmo usuário Linux que tem
acesso a `~/.hermes/google_token.json`. Se o token estiver em outro local, defina
`google_token_file` em `config.json`. A automação não lê nem altera um `.env` antigo;
`config.json`, `.env`, planilhas, histórico e prévias continuam excluídos do Git.

### 2. Gerar as prévias

Com `input_source: "drive"` (padrão), salve o arquivo no Google Drive em
`Meu Drive/Cartão/<ano>/<AAAA-MM>.xlsx`. Por exemplo, `2026-09.xlsx` fica em
`Meu Drive/Cartão/2026/`. O script seleciona o mês informado ou calculado pelo
agendamento; não usa o arquivo mais recente por nome ou data. A pasta `inputs/` só é
usada se você configurar explicitamente `input_source` como `local`.

```bash
.venv/bin/python cartoes.py --config config.json --card black --month 2026-09
.venv/bin/python cartoes.py --config config.json --card latam --month 2026-09
```

Sem `--send`, o script **somente gera prévias**. Para a sua planilha, copie
`config.example.json` para `config.json` e substitua os nomes fictícios em `sheet`
e `participant` pelos rótulos exatos da planilha; os e-mails podem ficar `null`
para gerar prévias. O arquivo de exemplo só corresponde a planilhas com esses
placeholders. Abra:

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
.venv/bin/python cartoes.py --config config.json --card latam --month 2026-09 --recipient NOME_DA_ABA --send --test-to seu.endereco@gmail.com
```

Troque `NOME_DA_ABA` pelo valor exato de `sheet` no `config.json` e o endereço
ilustrativo pelo destinatário de teste.

`--test-to` troca o destinatário em `To` pelo endereço informado, acrescenta `[TESTE]`
ao assunto e não registra um envio de produção. A cópia `personal_copy_email` continua
em `Cc` e fica visível ao destinatário do teste; se for o mesmo endereço de `To`, não
é repetida. Sem `--recipient`, todos os resumos daquele cartão vão ao endereço de teste.
Confira o e-mail no Gmail do computador e do celular, incluindo as cinco colunas e os valores de rateio.

Para enviar aos amigos:

```bash
.venv/bin/python cartoes.py --config config.json --card black --month 2026-09 --send
.venv/bin/python cartoes.py --config config.json --card latam --month 2026-09 --send
```

Um e-mail ausente ou inválido em qualquer resumo do lote impede os envios antes
da conexão à API do Gmail. Cada resumo vai para um único destinatário em `To` e
leva `personal_copy_email` em `Cc` (visível aos destinatários), sem Bcc. Não há anexo
com a planilha inteira nem conteúdo de outros amigos.

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

Os lançadores usam `~/Projetos/Automations/automations/cartoes/config.json`. Se o projeto ficar em outro
local, edite o caminho padrão nos dois arquivos copiados. Também é possível
definir `CARTOES_CONFIG` no ambiente do gateway. Eles importam o módulo `cartoes`
instalado no ambiente Python próprio, sem depender do diretório corrente.

Configure o fuso do Hermes **antes de criar os agendamentos**. O fuso usado pelo
cron no perfil padrão precisa ser `America/Sao_Paulo`; essa configuração afeta
os outros agendamentos do mesmo perfil:

```bash
hermes config set timezone America/Sao_Paulo
hermes gateway restart
```

[Referência de configuração de fuso do Hermes](https://hermes-agent.nousresearch.com/docs/user-guide/configuration#timezone).

Crie os agendamentos inicialmente pausados, com saída operacional local:

```bash
hermes cron create "0 9 5 * *" --no-agent --script cartao_black.py --interpreter "$HOME/Projetos/Automations/automations/cartoes/.venv/bin/python" --deliver local --name "Resumo cartão Black" --paused
hermes cron create "0 9 20 * *" --no-agent --script cartao_latam.py --interpreter "$HOME/Projetos/Automations/automations/cartoes/.venv/bin/python" --deliver local --name "Resumo cartão Latam" --paused
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
envia a competência do mês anterior; Latam só executa no dia 20 e envia a competência
do próprio mês. Fora do dia esperado, interrompem o envio. Para recuperar uma
execução atrasada ou com arquivo ausente, use o comando manual com `--month`;
não force um mês diferente por `--scheduled`. O teste manual pelo Hermes, com
`hermes cron run`, chama o lançador de **produção**; para testar sem enviar aos
amigos, use o comando Python com `--test-to` da etapa anterior.

### Histórico, repetição e falhas

Em execuções com `--send`, incluindo os lançadores do Hermes, uma falha interrompe
o lote e tenta enviar **um aviso ao proprietário**, para a conta Google
autenticada pelo token OAuth, com cópia para `personal_copy_email` quando esse endereço
está configurado. Isso inclui arquivo ausente, Excel corrompido, dados ou destinatários
inválidos e erros operacionais. A abertura e a leitura das células são verificadas
antes dos envios aos amigos. Se o Excel falhar nessa etapa, nenhum resumo é enviado.

O aviso identifica cartão, referência, etapa e motivo, sem anexar a planilha,
resumos ou erros internos que possam expor valores e credenciais. Mesmo com
`--test-to`, os resumos vão para o endereço informado em `To` e para `personal_copy_email`
em `Cc`; o aviso de erro vai para a conta Google autenticada e também ao `Cc` configurado.
Prévias, consulta e manutenção do histórico continuam sem envio de e-mail.

Se o erro ocorrer após o início dos envios, mensagens anteriores podem já ter
sido aceitas pelo Gmail; o aviso orienta a conferir o histórico e a pasta Enviados.
Se o Gmail, a conexão ou as credenciais também estiverem indisponíveis, o aviso
pode falhar: nesse caso o script registra a falha nos logs do Hermes, termina com
código 1 e não tenta enviar avisos recursivamente. Uma nova execução com erro pode
gerar um novo aviso; isso não altera a proteção contra resumos duplicados.

Se `config.json` não puder ser carregado, o aviso tenta usar o token padrão do
Hermes em `~/.hermes/google_token.json` (ou `$HERMES_HOME/google_token.json`).

O histórico fica em `var/deliveries.sqlite3`. Cada envio é identificado por
**mês + cartão + aba pessoal**. Uma execução posterior pula mensagens aceitas
anteriormente pela API do Gmail, inclusive se a planilha ou o endereço forem alterados.
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
- `failed`: a API do Gmail recusou a mensagem explicitamente; o comando de envio pode ser
  repetido para tentar os destinatários ainda não enviados.
- `unknown`: houve desconexão, timeout ou interrupção após iniciar o envio. O
  lote fica bloqueado até você conferir no Gmail para evitar duplicação.

Ao conferir um envio `unknown`, procure na pasta **Enviados** por destinatário,
assunto e `Message-ID` do histórico. Se foi enviado, marque como concluído:

```bash
.venv/bin/python cartoes.py --config config.json --card latam --month 2026-09 --recipient NOME_DA_ABA --resolve sent
```

Se confirmar que **não** foi enviado:

```bash
.venv/bin/python cartoes.py --config config.json --card latam --month 2026-09 --recipient NOME_DA_ABA --resolve not-sent
.venv/bin/python cartoes.py --config config.json --card latam --month 2026-09 --recipient NOME_DA_ABA --send
```

Há uma trava por pasta de histórico que impede dois processos de envio ou
manutenção ao mesmo tempo. Não há repetição automática de um resultado incerto:
a API do Gmail e o SQLite não oferecem uma transação conjunta de envio exatamente uma vez.

### Testes

```bash
.venv/bin/python -m unittest discover -s tests -v
```

A suíte usa planilhas sintéticas e a API do Gmail simulada; nenhum teste envia e-mails reais.
A amostra fictícia versionada em `tests/fixtures/anonymous_reports.json` cobre
**24 blocos**, com totais esperados explícitos, incluindo pagamentos parciais,
quitação, créditos, parcelas antigas e participantes sem movimento. Ela permite
validar **11 mensagens Black e 3 Latam** em um clone sem arquivos pessoais.

O teste com uma planilha real é opcional e somente local: se houver um arquivo
privado `inputs/2026-09.xlsx` no clone, a suíte compara os **24 blocos** com os dados
salvos. Esse teste usa esse caminho local independentemente do `input_source` da
configuração de produção. A planilha pessoal não é incluída no Git. Os testes
cobrem rateios, pagamentos, créditos, ausência de movimento, validações,
formatação, virada do ano, fuso, repetição, falhas parciais, falhas após aceite da API do Gmail,
trava entre processos, avisos de falha ao proprietário e Excel corrompido na
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

- Transfira `config.json` para a nova pasta da automação. O padrão atual busca em
  `Meu Drive/Cartão/<ano>/`; mantenha `input_source: "drive"`. Se ainda quiser
  arquivos locais, defina `input_source: "local"` e ajuste `input_dir`.
- Preserve a pasta de histórico indicada por `state_dir`, incluindo
  `deliveries.sqlite3`, para manter a proteção contra envios repetidos. Se mover
  `var/`, faça isso com os processos de envio parados.
- Crie a `.venv` no novo local e instale o pacote conforme a etapa 1.
- Copie novamente os lançadores para `~/.hermes/scripts/`, confira
  `CARTOES_CONFIG` se estiver definido e ajuste o caminho do interpretador dos
  agendamentos existentes para a nova `.venv`. Mantenha apenas um agendamento
  ativo por cartão.

A reorganização local não altera uma instalação já existente no servidor.
