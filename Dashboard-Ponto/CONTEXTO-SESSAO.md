# Contexto da sessão — Painel de Ponto (Aurora Coop)

> Arquivo criado para migração entre instâncias do Claude Code. Resume tudo que foi pedido e feito até aqui, para retomar o trabalho sem perder contexto.

## Quem é o usuário e o projeto

- Usuário trabalha no **RH da Aurora Coop** (cooperativa de alimentos brasileira, marca redesenhada pela Narita Design).
- Repositório: `Dimitrimateus/Excel-vba-toolkit`.
- Branch de trabalho: `claude/employee-attendance-dashboard-tsbe9g` (já existe, com todo o histórico de commits).
- Pasta do projeto: `Dashboard-Ponto/` dentro do repo, contendo:
  - `index.html` — dashboard completo em arquivo único (HTML+CSS+JS, ~138KB, sem dependências externas).
  - `GerarCSV.bas` e `GerarCSV.txt` (cópia idêntica em .txt) — macro VBA que gera o CSV a partir das planilhas Excel do RH.
  - `modelo-dados.csv` — CSV de exemplo/teste (183 linhas) com o schema esperado.
  - `README.md` — documentação completa do projeto (schema do CSV, regras de negócio, instruções de uso do macro, etc.).
  - `assets/` — 3 variantes do logo real da Aurora Coop (preto, negativo/branco, colorido) em PNG/JPG.

## Regras DURÁVEIS (nunca esquecer, valem para qualquer edição futura)

1. **"Nunca esqueça que a ausência de marcação vai vir da mesma planilha que tratamento. Numa aba chamada pontonet."** — A macro `GerarCSV.bas` já implementa isso: procura automaticamente uma aba chamada **"PontoNet"** dentro da própria planilha "Tratamento Ponto"; só pede um arquivo separado se essa aba não existir.
2. **"Sempre gere o código em txt."** — Toda vez que o código VBA for gerado/atualizado, manter uma cópia idêntica em `.txt` (algumas caixas de e-mail corporativas bloqueiam `.bas`). Hoje isso é `GerarCSV.txt`, espelho exato de `GerarCSV.bas`.

## O que foi construído (visão geral)

### 1. Dashboard (`index.html`)
Painel de ponto/ocorrências dos colaboradores, 100% offline (funciona como anexo de e-mail, sem internet, sem instalação — o RH só abre no navegador e arrasta o CSV).

Views/gráficos implementados:
- Colaboradores com mais ocorrências
- Ocorrências por tipo
- Ocorrências por data (linha, com rótulos de valor fixos e datas no eixo X sem precisar de hover)
- Ocorrências por dia da semana (com valores nas barras sem hover)
- Ocorrências por gestor
- Ocorrências por cargo (era "por setor", trocado a pedido do usuário)
- Colaboradores com mais extra+falta no mesmo dia (mostra a coluna "Dias" sempre visível, sem precisar arrastar)
- Colaboradores com maior qtd de horas falta/extra
- Ocorrências curtas (< 15min) — hoje vem da coluna "Tolerância < 15min" do Cartão Ponto, não é mais calculado no cliente
- Banco de horas e hora extra por colaborador (gráfico de barras empilhadas, independente do card de KPI que foi removido)
- Relação de horas excedentes (nova view, fonte: coluna "Horas Excedentes" da aba Tratamento)
- Colaboradores envolvidos (número)
- Demora média/maior por gestor e por colaborador no PontoNet (tempo entre data da ocorrência e data de tratativa)
- Histórico de ocorrências no rodapé (data/colaborador/gestor/situação/status)
- KPIs no topo (hoje com **5 tiles**, depois de remover o card "Banco de horas gerado"): Colaboradores envolvidos, Ocorrências no período, Total de horas extra (com sub-texto condicional "· Xh a 100%" quando há horas extra a 100%), Total de horas falta/atraso, Demora média no PontoNet.
- Filtros congelados à direita no desktop (Gestor, Cargo, Tipo, Status, período), com chips que começam **brancos/inativos** e ficam **laranja quando ativos** (não preto/cinza).
- Logo real da Aurora Coop (versão negativa/branca) embutido como `data:image/png;base64,...` no `<header>`, para funcionar 100% offline.

### 2. Responsividade completa (mobile/tablet/desktop)
- Celular em pé: aviso sugerindo girar a tela (fecha sozinho ao girar ou no X).
- Gráficos em 1 coluna no mobile, redesenhados automaticamente ao redimensionar/girar (via `containerWidth()` que usa a largura real do container em pixels no viewBox do SVG).
- Tabelas viram cartões empilhados no mobile (um por ocorrência, rótulo + valor); no tablet continuam como tabela com rolagem horizontal suave e primeira coluna fixa.
- Filtros em painel retrátil no mobile ("2. Filtros", `<details>/<summary>`), com badge mostrando quantidade de filtros ativos.
- Barra fixa inferior no mobile com atalhos para Filtros/Topo/Histórico.
- Inputs com `font-size:16px` no mobile (evita zoom automático do Safari iOS).
- Estado de carregamento ao subir o CSV.
- Touch targets ≥44px.

### 3. Macro VBA (`GerarCSV.bas` / `GerarCSV.txt`)
Consolida dados de duas fontes dentro da mesma planilha "Tratamento Ponto":
- Aba **"Tratamento"**: fonte principal. Coluna "Check" = "S" marca ocorrência trivial (seria descartada, mas hoje é reaproveitada só para popular "ocorrências curtas", ver abaixo). "Sem alteração" na coluna Situação também é excluído do fluxo normal. Colunas "Extras"/"Faltas" decidem SE aquele dia entra como ocorrência normal de extra/falta (única fonte usada para "extra e falta no mesmo dia"). A coluna "Pontonet" da aba Tratamento é **ignorada de propósito**.
- Aba **"PontoNet"** (ou arquivo separado se a aba não existir): dados de ausência de marcação, casados por matrícula/data via `Scripting.Dictionary`.
- Aba **"Cartão ponto até dia..."** (localizada via prefixo, robusto a variações de nome/espaço): fonte de horas.
  - Coluna **BH** (P) → minutos de falta cobertos por banco de horas.
  - Coluna **50%** (S) → hora extra "normal" (tipo `Hora Extra`).
  - Coluna **100%** (U) → hora extra a 100%, **sempre em linha separada** (tipo `Hora Extra 100%`), e sinalizada no card "Total de horas extra" do painel.
  - Colunas **60%** e **120%** (T, V) → **não usadas**, removidas a pedido explícito do usuário.
  - Coluna **"Tolerância < 15min"** (AA) → alimenta a visão "Ocorrências curtas": só quando `Check = "S"` (linha que seria descartada) E essa coluna confirma "Falta < 15min" ou "Extra < 15min" — nesse caso a linha entra no CSV só com esse tipo, sem contar nos totais normais de extra/falta (evita double-counting).
- Saída: aba "CSV" com 13 colunas — `data, colaborador, matricula, gestor, setor, cargo, tipo_ocorrencia, situacao, status, duracao_minutos, destino_horas_extra, data_tratativa_pontonet, horas_excedentes`.
- Exportação: `Application.FileDialog(4)` (não usa `FileDialog` tipado, para não exigir referência de biblioteca do Office); CSV em UTF-8 via `ADODB.Stream`.
- **Nunca testado dentro do Excel real** (sandbox não tem Excel/VBA runtime) — validado só via simulação em Python replicando a lógica exata, rodada contra os arquivos `.xlsx` reais que o usuário enviou (deu 128 linhas: 22 Hora Extra, 71 Falta, 4 zero-impacto, 4 Hora Extra 100%, 27 curtas recuperadas de Check=S). **Avisar o usuário para testar numa cópia da planilha antes de usar em produção.**

## Regras de negócio confirmadas pelo usuário (importantes para não reintroduzir bugs)
- Tolerância de 5 minutos no ponto: ex. jornada às 7h, chegar 6h55–7h05 não conta atraso; a partir de 7h06 já conta (6 min de atraso).
- "Dias colaborador" vs "Dias gestor" no PontoNet: o atraso do gestor é o **relógio dele próprio**, não é cumulativo com o do colaborador.
- Coluna "Check" = "S" na aba Tratamento = ocorrência trivial/não-real (seria totalmente descartada, hoje só reaproveitada para "ocorrências curtas" via Cartão Ponto, ver acima).
- "Sem alteração" em Situação = também excluído do fluxo normal.

## Últimas mudanças pontuais (mais recentes primeiro)

1. **"Retire o card de banco de horas gerado"** (última tarefa concluída, commit `c06e2c5`):
   - Removido o tile de KPI "Banco de horas gerado" (`#kpiBanco`).
   - Grid de KPIs mudou de `repeat(6, ...)` para `repeat(5, ...)`.
   - Removida a variável morta `bancoTotal` e seus usos dentro de `renderKPIs()`.
   - O card separado "Banco de horas e hora extra por colaborador" (gráfico de barras empilhadas) **não foi afetado** — usa sua própria função `bancoDeHorasEExtra()`, independente.
   - Testado via Playwright (0 erros de console, 5 tiles confirmados), commitado e enviado ao usuário.

2. Antes disso, ajuste de hora extra/ocorrências curtas (commit `a14f44f`): restringir cálculo de hora extra só às colunas 50%(S) + 100%(U) do Cartão Ponto (excluindo 60%/120%), criar linha/indicador dedicado "Hora Extra 100%", e passar "Ocorrências curtas" a vir da coluna "Tolerância < 15min" do Cartão Ponto em vez de ser calculado no cliente.

3. Antes disso: cópia em `.txt` do macro (commit `70c5dd7`); valores fixos em gráficos + tabelas sem overflow + view de horas excedentes (commit `2a117fb`); responsividade completa mobile/tablet/desktop (commit `8c94836`).

## Estado atual do git (checar ao retomar)

```
Branch: claude/employee-attendance-dashboard-tsbe9g
Working tree: limpo (sem pendências) na última verificação
Últimos commits:
  c06e2c5 Remove o card de KPI "Banco de horas gerado"
  a14f44f Ajusta hora extra (50%/100% do Cartão Ponto) e ocorrências curtas
  70c5dd7 Adiciona cópia em .txt do macro VBA
  2a117fb Valores fixos nos gráficos, tabelas sem scroll e nova visão de horas excedentes
  8c94836 Torna o painel 100% responsivo (mobile, tablet e desktop)
```
Sempre rodar `git status` e `git log --oneline -10` ao retomar para confirmar que nada mudou fora desta sessão.

## Detalhes técnicos importantes para quem for editar o código

- **Sem dependências externas**: nenhum Chart.js/D3, gráficos SVG feitos à mão (`renderBarH`, `renderBarV`, `renderDonut`, `renderLine`, `renderStackedH`), parsing de CSV também feito à mão (`parseCSV`, `detectDelimiter`, `parseNumber` com vírgula decimal BR, `parseDate` para DD/MM/AAAA e ISO).
- **`containerWidth(container, fallback)`**: helper central da responsividade — usa a largura real em pixels do container para o viewBox do SVG, garantindo que texto renderize no tamanho certo e gráficos se adaptem a resize/rotação.
- **`HEADER_ALIASES`**: no topo do `<script>` de `index.html`, mapeia variações de nome de cabeçalho do CSV (com/sem acento, sinônimos) para os campos internos. Inclui `cargo` e `horasExcedentes`.
- **Chips de filtro**: lógica invertida de propósito — chip SEM classe `active` = todos aparecem (comportamento "mostrar tudo" quando nada está selecionado). Cuidado ao tocar nisso, já foi bug reportado antes.
- **Tabelas** usam `table-layout:fixed` + `<colgroup>` com larguras explícitas para colunas curtas/numéricas, evitando que a coluna mais importante fique escondida fora da tela.
- **Ler `index.html` com cautela**: o arquivo tem uma linha única gigante (~61KB) com o base64 do logo (por volta da linha ~417). Ler ranges que cruzem essa linha pode falhar por limite de tokens da tool Read. Se precisar inspecionar contexto ao redor, gerar uma cópia temporária removendo essa linha específica (`sed '<N>d' index.html > /tmp/index_noimg.html`) e ler a cópia; usar Grep/Edit diretamente no arquivo real (não são afetados pelo tamanho da linha).
- **Testado com Playwright** (Chromium em `/opt/pw-browsers/chromium`, `NODE_PATH=/opt/node22/lib/node_modules`) em viewports desktop/tablet/mobile a cada mudança visual, verificando erros de console, DOM, estilos computados e screenshots.

## Pendências / avisos abertos (não resolvidos, mas conhecidos)

- Macro VBA nunca foi rodada dentro do Excel real — recomendar ao usuário testar numa cópia da planilha antes de usar em produção.
- No cabeçalho do próprio `GerarCSV.bas`, há comentários "REGRA:" marcando pontos que ainda dependem de confirmação do usuário (ambiguidade de destino banco/pagamento da hora extra, vocabulário de status Tratado/Tratando/Tratar, ambiguidade de BH em dias "Trabalhando"). Vale reler esses comentários ao retomar.

## Última mensagem do usuário antes deste arquivo

Pediu para criar este arquivo de contexto porque vai migrar para outra instância do Claude Code. Nenhuma tarefa de código pendente — a última solicitação explícita ("Retire o card de banco de horas gerado") já foi 100% concluída, testada, commitada e enviada.
