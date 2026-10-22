# Retenção de Clientes (`century_churn_retention`)

> Pipeline de retenção para os clientes que o modelo preditivo de churn aponta como em risco: o gerente importa a lista em Excel e cada vendedor trabalha os seus clientes até recuperar ou confirmar a perda.

## Para que serve

O modelo de previsão de churn (feito em notebook, fora do Odoo) gera uma planilha com os clientes que podem parar de comprar, com a probabilidade de perda, a curva ABC e dados de receita. Este módulo leva essa planilha para o Odoo e transforma cada cliente em um cartão num pipeline de retenção, com responsável, prazo, histórico de conversa e resultado final.

O objetivo é que o contato com o cliente em risco aconteça e seja registrado, e que o gerente consiga medir quantos foram recuperados e por que os outros se perderam.

Cada cliente é ligado ao seu vendedor e ao líder de equipe pelo cadastro do cliente no Odoo. Cada pessoa vê apenas os clientes da sua carteira ou da sua equipe.

## Onde fica no Odoo

Aplicativo **Retenção** no menu principal.

- **Retenção › Pipeline de Retenção**: o quadro de cartões (kanban), com lista, formulário, gráfico e tabela dinâmica.
- **Retenção › Importar Lista de Churn**: só para o grupo Gerente.
- **Retenção › Configuração › Estágios**: só para o grupo Gerente.

## Como usar

### Gerente: importar a lista do mês

1. Entre em **Retenção › Importar Lista de Churn**.
2. Anexe o arquivo `.xlsx` exportado pelo notebook de churn. A primeira linha deve ser o cabeçalho. É preciso ter ao menos a coluna de CNPJ ou a de nome do cliente. As demais colunas: curva, probabilidade de churn, risco, receita total, número de pedidos, recência, variação de receita. Os cabeçalhos aceitam variações comuns (por exemplo, `cnpj`, `cpf`, `vat`, `probabilidade`, `receita`).
3. Confira o **Nome do Lote** (sugestão: "Churn AAAA-MM") e o **Prazo para Contato** (sugestão: dia 28 do mês).
4. Escolha o que fazer **se o cliente já estiver no pipeline**: atualizar a probabilidade ou ignorar.
5. Deixe marcado **Importar apenas Curva A e B**, se quiser pular a curva C.
6. A tela mostra uma prévia das 10 primeiras linhas, com o total de linhas e quantas acharam cliente no Odoo.
7. Clique em **Importar Clientes**. Você cai no pipeline filtrado pelo lote importado.

### Vendedor ou líder: trabalhar o pipeline

1. Abra **Retenção › Pipeline de Retenção**. A tela já vem filtrada para **Em Processo**, **Curva A + Risco Alto** e **Meus Clientes**. Remova os filtros para ver o restante.
2. Abra o cartão do cliente. Clientes de risco alto mostram um alerta vermelho pedindo contato urgente.
3. Clique em **Registrar Contato** quando falar com o cliente. Isso grava a data do contato e deixa uma nota no histórico.
4. Arraste o cartão entre os estágios conforme o trabalho avança e use a prioridade (estrelas) e o estado do cartão (Em Progresso, Pronto para Avançar, Bloqueado).
5. Registre o diagnóstico na aba **Observações** e use o histórico do cartão para conversar e agendar atividades.
6. No fim, clique em **Marcar como Recuperado** ou **Confirmar Churn**. Em caso de churn, preencha o **Motivo do Churn**: sell-out fraco, concorrência, preço, atendimento, mix inadequado, fechamento da loja ou outro.

## Regras de negócio

- **Nível de risco vem da probabilidade**: 70% ou mais é Alto, de 40% a menos de 70% é Médio, abaixo de 40% é Baixo. É recalculado sozinho.
- **Probabilidade pode vir em escala 0 a 1 ou 0 a 100**: na importação, valores até 1 são tratados como fração e multiplicados por 100 (0,85 vira 85%).
- **Curva C fica fora por padrão**, porque o churn de curva C é estruturalmente alto e o retorno do esforço de retenção é baixo.
- **O cliente é encontrado pelo CNPJ** (ignorando pontuação e zeros à esquerda). Se o CNPJ não achar, tenta o nome do cliente (igual, depois parecido). O cartão é criado no cadastro principal da empresa. Linhas sem cliente correspondente são puladas, e os primeiros erros aparecem ao fim da importação.
- **Vendedor, equipe e líder vêm do cadastro do cliente**: o vendedor é o responsável do cliente no Odoo; a equipe é a primeira equipe de vendas em que esse vendedor está; o líder é o líder dessa equipe. Se o cliente não tem vendedor, o cartão fica sem equipe e sem líder, e só o gerente o vê.
- **Todo cartão importado começa no primeiro estágio** (Identificado) e com resultado "Em Processo".
- **Cliente já em processo**: se já existe cartão em processo para o mesmo cliente, a importação atualiza (probabilidade, curva, receita e lote) ou ignora, conforme a escolha. Cliente recuperado ou churned anteriormente gera um cartão novo.
- **Recuperado e Churned movem o cartão para o estágio correspondente** (o marcado como recuperação ou churn) e deixam registro no histórico.
- **Arrastar o cartão para o estágio Recuperado ou Churned não altera o campo Resultado**; só os botões fazem isso (o código não faz essa sincronização).
- **Ordenação**: maior probabilidade primeiro, depois curva A antes de B e C.

Estágios que o módulo cria (editáveis pelo gerente):

| Estágio | Significado |
|---|---|
| Identificado | Cliente importado pelo modelo. Ainda sem contato. |
| Contato Realizado | Primeiro contato feito. Diagnóstico em andamento. |
| Em Tratamento | Ação de retenção em andamento (troca de showroom, proposta comercial etc.). |
| Monitoramento | Ação executada. Aguardando retorno de compra. |
| Recuperado | Cliente voltou a comprar. |
| Churned | Perda confirmada. Registrar o motivo. |

## Quem pode o quê

Os grupos aparecem em Configurações › Usuários, na categoria Vendas. Cada grupo inclui o anterior.

| Grupo | O que vê e faz |
|---|---|
| Retencao / Vendedor | Vê e edita apenas os cartões dos seus clientes (onde é o vendedor). Cria cartões. Não apaga. Só consulta os estágios. |
| Retencao / Lider de Vendas | Vê e edita os cartões da sua equipe (onde é líder). Pode apagar. Mais o que o vendedor faz. |
| Retencao / Gerente | Vê tudo. Importa listas e edita os estágios (menus Importar e Configuração). |

Quem não está em nenhum grupo não vê o aplicativo.

## Configurações e parâmetros

Sem parâmetros do sistema e sem rotinas agendadas. A única configuração é a lista de estágios, em Retenção › Configuração › Estágios (nome, sequência, cor, se é estágio de recuperação, se é de churn confirmado, se fica dobrado no quadro).

## Integrações

- **Modelo preditivo de churn** (notebook, fora do Odoo): produz o Excel que a importação lê. Não há ligação automática; o arquivo é carregado à mão.
- **Cadastro de clientes do Odoo**: liga o cartão ao cliente (por CNPJ ou nome), ao vendedor responsável e, por ele, à equipe e ao líder.
- **Focco e DW**: sem integração.

## Limitações conhecidas

- **Importação manual**: não há agendamento nem leitura automática do notebook.
- **A atualização de cartão existente é parcial**: só probabilidade, curva, receita e lote são atualizados. Pedidos, recência, variação de receita e prazo ficam como estavam, e o estágio não volta ao início.
- **O vendedor vem do cadastro do cliente**: se o responsável do cliente mudar, o cartão acompanha (o campo é calculado do cadastro). Cliente sem vendedor só aparece para o gerente.
- **Equipe**: se o vendedor participa de mais de uma equipe de vendas, vale a primeira cadastrada.
- **Filtros de abertura**: o pipeline abre com filtros pré-aplicados (Em Processo, Curva A + Risco Alto, Meus Clientes). Quem não é vendedor de nenhum cliente, como o gerente, verá a tela vazia até tirar o filtro Meus Clientes.
- **Resultado e estágio não são sincronizados** ao arrastar o cartão (veja Regras de negócio).
- **Não há testes automatizados** nem README anterior; o módulo ainda não tem TODO registrado.

## Perguntas frequentes

**Que arquivo eu importo?**
O Excel (.xlsx) exportado pelo notebook de previsão de churn, com a primeira linha de cabeçalho.

**A importação criou menos cartões que linhas da planilha. Por quê?**
Linhas de curva C (se a opção estiver marcada), clientes não encontrados no Odoo e clientes já em processo (se escolheu ignorar) são puladas.

**Por que não vejo nenhum cliente no pipeline?**
O pipeline abre com o filtro "Meus Clientes". Tire-o, e confira se você está em um dos grupos de Retenção.

**O que muda entre risco Alto, Médio e Baixo?**
Só a classificação e o destaque na tela: 70% ou mais é Alto, de 40% até 69% é Médio, abaixo de 40% é Baixo.

**Quem vê os clientes da minha equipe?**
O líder de vendas da equipe e o gerente. O vendedor vê apenas os seus.

**Como marco um cliente como recuperado?**
Abra o cartão e clique em Marcar como Recuperado. Arrastar para a coluna Recuperado não registra o resultado.

**Posso apagar um cartão?**
Só líderes de vendas e gerentes.

**O cliente aparece duas vezes. Por quê?**
Se o cartão anterior já foi marcado como recuperado ou churned, uma nova importação cria um cartão novo.

## Para desenvolvedores

Dependências: `crm`, `sale_management`, `mail`. Versão `18.0.1.0.0`. A importação exige a biblioteca Python `openpyxl` no servidor.

```
century_churn_retention/
├── models/retention_lead.py      century.retention.lead (cartão) e century.retention.stage (estágio)
├── wizard/import_churn_wizard.py century.import.churn.wizard: leitura do xlsx, prévia, importação
├── wizard/import_churn_views.xml formulário do assistente
├── security/security.xml         3 grupos (user, coordinator, manager) e 3 regras de registro
├── security/ir.model.access.csv  ACLs por grupo
├── data/crm_stage_data.xml       6 estágios (noupdate=1)
├── views/retention_views.xml     kanban, lista, formulário, busca, gráfico, pivot, ações
├── views/retention_menu.xml      menus
└── static/description/icon.png
```

- `century.retention.lead` herda `mail.thread` e `mail.activity.mixin`. Campos importantes: `partner_id` (domínio: empresa), `cnpj` (related do `vat`, stored), `representante_id` (related de `partner_id.user_id`, stored), `team_id` (computado a partir de `crm.team.member`, stored, primeira por `id`), `team_leader_id` (related de `team_id.user_id`, stored), `nivel_risco` (computado de `prob_churn`, stored).
- Regras de registro (`ir.rule`): vendedor `representante_id = user.id`; líder `team_id.user_id = user.id`; gerente sem filtro. Como regras de grupos diferentes se somam (OU), a hierarquia de grupos entrega o acesso acumulado.
- `_find_partner` no assistente busca por `vat` com `ilike`; se não achar, percorre todos os parceiros com `vat` comparando só dígitos. Isso pode ser lento em base grande.
- A prévia (`_onchange_file_preview`) lê o arquivo a cada alteração de arquivo ou de `only_ab` e consulta só as 10 primeiras linhas.
- Colunas do Excel: mapeadas por `EXPECTED_COLS` em `import_churn_wizard.py` (cabeçalhos em minúsculas e com `_` no lugar de espaço).
- O assistente roda como o usuário logado; só o grupo Gerente tem acesso ao modelo do assistente.
- Atenção: a ação do pipeline traz `search_default_meus` no contexto, e o assistente devolve a mesma ação com esse default após a importação.
- Instalação/upgrade: `-i`/`-u century_churn_retention`. Sem testes automatizados no repositório.
