# Painel de Ponto (Aurora Coop)

Dashboard em um único arquivo HTML (`index.html`) para visualizar as ocorrências de ponto dos colaboradores. Não depende de internet nem de instalação: o RH carrega o CSV de dados, e o próprio arquivo HTML desenha os gráficos no navegador.

## Responsividade (mobile e tablet)

O painel funciona em celular e tablet, não só no desktop:

- **Celular em pé (retrato):** aparece um aviso sugerindo girar a tela (some sozinho ao girar, ou pode ser fechado no X). Gráficos em 1 coluna, tabelas viram cartões empilhados (um por ocorrência/colaborador, com rótulo + valor), filtros ficam num painel retrátil ("2. Filtros", toque para abrir) e uma barra fixa embaixo dá atalho para Filtros/Topo/Histórico.
- **Tablet:** os cartões se reorganizam sozinhos conforme o espaço (2 ou 3 colunas), e as tabelas continuam como tabela mesmo, com rolagem horizontal suave e a primeira coluna fixa.
- **Desktop:** segue igual a antes (2 colunas, filtros na lateral direita, sempre abertos).

Os gráficos são redesenhados automaticamente ao redimensionar a janela ou girar o aparelho, e os campos de filtro usam fonte de 16px no mobile para o Safari do iPhone não dar zoom sozinho ao tocar neles.

## Como funciona o envio

1. O RH mantém **um único `index.html`**, sempre igual para todo mundo.
2. Para cada **gestor**, gera-se um CSV contendo apenas a equipe dele e envia-se os dois arquivos juntos (HTML + CSV daquele gestor).
3. Para o **supervisor geral**, envia-se o mesmo `index.html` com um CSV único contendo todos os gestores e todas as equipes.
4. Sem o CSV, o HTML abre vazio (nenhum gráfico contém dados por padrão). Assim, cada pessoa só vê os dados do arquivo que você mandou para ela.
5. Quem recebe os arquivos abre o `index.html` no navegador (duplo clique) e arrasta o CSV para o quadro "Carregar dados", no painel à direita. Nada é enviado para a internet: tudo é processado localmente no navegador de quem abrir o arquivo.
6. O filtro de **Gestor**, na barra de filtros, existe pensando no arquivo do supervisor geral (que tem todos os gestores). No arquivo enviado a um gestor específico, o filtro aparece mas normalmente não precisa ser usado, já que só existe um gestor naquele CSV.

## Estrutura do CSV

Use `modelo-dados.csv` (nesta mesma pasta) como modelo: ele já está no formato certo, com dados fictícios de exemplo. Basta apagar as linhas de exemplo e preencher com os dados reais, mantendo o cabeçalho.

Colunas esperadas (a ordem das colunas não importa, o que importa é o nome do cabeçalho):

| Coluna | Obrigatória | Descrição | Exemplos de valor |
|---|---|---|---|
| `data` | Sim | Data da ocorrência. | `24/08/2026` ou `2026-08-24` |
| `colaborador` | Sim | Nome do colaborador. | `Ana Paula Ribeiro` |
| `matricula` | Não | Matrícula/ID do colaborador, útil se houver homônimos. | `5242` |
| `gestor` | Não* | Gestor responsável pelo colaborador. | `Marina Souza` |
| `setor` | Não* | Setor/área do colaborador. | `Produção` |
| `cargo` | Não* | Cargo/função do colaborador (usada no gráfico "Ocorrências por cargo"). | `Auxiliar de Logística` |
| `tipo_ocorrencia` | Não* | Tipo da ocorrência. | `Atraso`, `Falta`, `Saída Antecipada`, `Hora Extra`, `Hora Extra 100%`, `Esquecimento de Registro`, `Abono`, `Falta < 15min`, `Extra < 15min` |
| `situacao` | Não | Descrição/justificativa da ocorrência (aparece no histórico). | `Atraso na entrada` |
| `status` | Não | Situação do tratamento da ocorrência. | `Pendente`, `Aprovado`, `Reprovado`, `Regularizado` |
| `duracao_minutos` | Não | Duração em minutos, **com sinal**: negativo para falta/atraso, positivo para hora extra. | `-30` (30 min de atraso), `120` (2h de hora extra) |
| `destino_horas_extra` | Não | Só faz sentido quando `duracao_minutos` é positivo. | `Banco de Horas` ou `Pagamento` |
| `data_tratativa_pontonet` | Não | Data (e hora, se tiver) em que a ocorrência foi tratada/aprovada/regularizada no PontoNet. Usada para calcular a demora. | `25/08/2026 14:30` |
| `horas_excedentes` | Não | Marca a ocorrência como "hora excedente" (aparece na lista "Relação de horas excedentes"). Qualquer valor preenchido conta como marcado; vazio = não marcado. | `Sim` |
| `ocorrencias_extra_falta_mes_atual` | Não | Só usada em linhas com `tipo_ocorrencia = Auditoria Extra e Falta (3 Meses)` (uma por colaborador, geradas pela macro): quantos dias de extra+falta no mesmo dia esse colaborador teve no mês atual. | `4` |
| `ocorrencias_extra_falta_3_meses` | Não | Mesma linha de auditoria acima: soma de dias de extra+falta no mesmo dia nos últimos meses (até 3). A partir de 10, o painel destaca a linha em vermelho na visão "Extra e falta no mesmo dia" (regra da auditoria). | `12` |
| `pausas_corretas`, `pausas_menor_20min`, `pausas_maior_20min`, `trabalho_correto_140`, `trabalho_maior_140`, `trabalho_menor_140`, `pausas_marcacoes_impares` | Não | Só usadas em linhas com `tipo_ocorrencia = Resumo Pausas Térmicas` (uma por colaborador, geradas pela macro a partir da aba "Pausas térmicas"): os mesmos números da seção "Totais por Colaborador" desse relatório, repassados sem recálculo. Alimentam a visão "Pausas Térmicas — Resumo por colaborador". | `19`, `0`, `0`, `13`, `2`, `17`, `0` |

`* `Se a coluna não existir ou vier vazia numa linha, o painel usa "Sem gestor" / "Sem setor" / "Sem cargo" / "Não informado" no lugar, mas os gráficos correspondentes perdem o sentido. Vale a pena preencher.

Detalhes que o painel já resolve automaticamente:
- Aceita CSV separado por vírgula ou por ponto e vírgula (comum em exportações do Excel em português).
- Aceita datas em `DD/MM/AAAA` ou `AAAA-MM-DD`, com ou sem hora.
- Aceita números com vírgula decimal (`"60,5"`) além do formato com ponto.
- Os nomes das colunas no cabeçalho podem ter variações razoáveis (com/sem acento, maiúsculas, `tipo` em vez de `tipo_ocorrencia`, etc.). Veja a lista de sinônimos aceitos no início do `<script>` de `index.html`, na constante `HEADER_ALIASES`, caso precise adaptar a exportação do sistema de ponto da empresa.

## O que "demora no PontoNet" significa aqui

É o tempo entre a **data da ocorrência** (coluna `data`) e a **data em que ela foi tratada/aprovada/regularizada** (coluna `data_tratativa_pontonet`). Se essa segunda coluna ficar vazia numa linha, aquela ocorrência simplesmente não entra nas contas de demora (mas continua contando nos outros gráficos).

Se, na prática da empresa, "demora no PontoNet" significa outra coisa (por exemplo, o atraso entre o horário previsto de bater o ponto e o horário real do registro, e não o tempo de tratativa), é só preencher `data_tratativa_pontonet` com o dado que corresponda a essa definição. O cálculo em si (diferença de tempo) não muda.

## Sobre o logotipo e as cores

O painel usa a paleta oficial da Aurora Coop (vermelho, laranja, amarelo, preto, branco e dois tons de cinza), a partir das imagens da marca fornecidas diretamente pelo RH. Os valores hexadecimais usados são uma leitura aproximada dessas imagens (vermelho `#E31E24`, laranja `#F7931E`, amarelo `#FFC72C`, preto `#1A1A1A`, cinza escuro `#58595B`, cinza claro `#A7A9AC`). Se você tiver o guia de marca com os códigos exatos, basta ajustar os valores no topo do `<style>`, dentro de `:root { ... }` (variáveis `--c-orange`, `--c-gold`, `--c-terracotta`, etc.).

O cabeçalho usa o **logotipo oficial de verdade**: a versão negativa (branca, para fundo escuro) que você enviou está embutida direto no `<header>` de `index.html`, como `data:` URI (imagem `<img class="brand-mark">`), então o arquivo continua funcionando sem internet e sem depender de nenhum outro arquivo solto.

As três variantes que você mandou (preta, negativa/branca e colorida) também estão salvas em `assets/`, caso precise delas soltas para outra coisa (um e-mail, um documento, etc.). Se quiser trocar a versão usada no cabeçalho (por exemplo, para a colorida ou a preta, caso o fundo mude para claro no futuro):

1. Gere o `data:image/png;base64,...` (ou `image/jpeg;base64,...`) do arquivo em `assets/` que você quiser usar (qualquer conversor online, ou `base64 arquivo.png` no terminal).
2. Abra `index.html`, procure `<img class="brand-mark"` dentro do `<header>` e troque o valor do `src` pela nova string.

## Testando antes de enviar

Abra `index.html` em qualquer navegador (Chrome, Edge, Firefox) e carregue o `modelo-dados.csv` para ver o painel funcionando com dados de exemplo antes de trocá-lo pelos dados reais.

## Gerando o CSV a partir do PontoNet (macro VBA)

`GerarCSV.bas`, nesta mesma pasta, é uma primeira versão do macro que consolida a aba "Tratamento" (da sua planilha "Tratamento Ponto") com os dados de "Ausência de marcação" do PontoNet numa aba "CSV" pronta para exportar. A macro procura automaticamente uma aba chamada **"PontoNet"** dentro da própria planilha "Tratamento Ponto" (é lá que você deve colar o relatório de ausência de marcação); se essa aba não existir, ela pede para você selecionar um arquivo separado. Esse arquivo **não foi testado dentro do Excel** (não há Excel disponível no ambiente onde ele foi escrito), então revise e rode primeiro numa cópia da planilha. O cabeçalho do próprio arquivo `.bas` explica o passo a passo de instalação e as regras de negócio que ainda precisam da sua confirmação (procure por "REGRA:").

A macro também usa a aba de Cartão Ponto do mês atual (a mais recente entre as abas "Cartão ponto ...", ver seção abaixo) para pegar a quantidade de horas:
- Coluna **"BH"** → hora de falta coberta por banco de horas.
- Coluna **"50%"** → hora extra "normal" (tipo `Hora Extra`).
- Coluna **"100%"** → hora extra a 100%, sempre reportada em separado (tipo `Hora Extra 100%`) e indicada no card "Total de horas extra" do painel. As colunas "60%" e "120%" não são usadas.

As colunas "Extras"/"Faltas" da aba Tratamento continuam decidindo só SE aquele dia entra como ocorrência de extra/falta "normal" nas linhas do dia a dia do CSV; a coluna "Pontonet" da Tratamento é ignorada de propósito.

> **Correção importante:** até esta versão, a leitura da aba de Cartão Ponto nunca encontrava as colunas (o cabeçalho de verdade está na **linha 2** dessa aba, não na linha 1, e a coluna de matrícula lá se chama "Matrícula" com acento, diferente da Tratamento). Na prática isso significa que "Hora Extra 100%" nunca aparecia e os valores de hora extra/falta sempre vinham só da Tratamento, nunca do Cartão Ponto. Já está corrigido — se você já revisou números de uma exportação anterior, vale conferir de novo.

### Ocorrências curtas (< 15 minutos)

Antes, essa visão só pegava linhas que tinham `Check = "S"` na Tratamento **e** confirmação na coluna "Tolerância < 15min" do Cartão Ponto — um subconjunto pequeno. Agora a macro varre a coluna **"Tolerância < 15min"** da aba de Cartão Ponto do mês atual direto, sem depender da Tratamento: toda linha com essa coluna preenchida ("Falta < 15min", "Extra < 15min" ou "Extra e Falta < 15min") vira uma ou duas linhas no CSV (uma de falta, uma de extra, quando for "e"), usando as colunas "BH"/"50%"/"100%" da própria linha do Cartão Ponto como duração.

### Auditoria de "extra e falta no mesmo dia" (últimos meses)

A macro procura automaticamente **todas** as abas cujo nome comece com "Cartão ponto" (ex.: "Cartão ponto Julho", "Cartão ponto Agosto", "Cartão ponto Atual" — o nome depois de "Cartão ponto" pode ser qualquer coisa) e usa até as **3 mais recentes** (pela maior data encontrada na coluna "DT" de cada aba, não pelo nome — então não precisa renomear nada de mês a mês, só manter no máximo 3 abas desse tipo na planilha). A aba com a data mais recente é tratada como "mês atual".

Para cada colaborador, a macro conta quantos dias tiveram "Sim" na coluna **"Banco de horas e extra no mesmo dia"** de cada uma dessas abas (esse é o sinal usado tanto para "este mês" quanto para o total dos últimos meses — uma fonte só, consistente entre os períodos) e grava isso numa linha extra no CSV, com `tipo_ocorrencia = Auditoria Extra e Falta (3 Meses)` e as colunas `ocorrencias_extra_falta_mes_atual` / `ocorrencias_extra_falta_3_meses`. Essas linhas não aparecem nos gráficos/KPIs/histórico normais do painel — só alimentam a tabela "Extra e falta no mesmo dia", que mostra as duas colunas lado a lado e destaca em vermelho quem atingiu **10 ou mais ocorrências na soma dos últimos meses** (regra atual da auditoria interna).

Colaboradores que só aparecem numa aba de Cartão Ponto antiga (não estão mais na Tratamento do mês atual) ainda entram nessa auditoria, mas ficam sem gestor preenchido se não tiverem nenhuma linha na Tratamento do período — não há outro lugar na planilha com essa informação para eles.

### Resumo de Pausas Térmicas

Se a planilha tiver uma aba chamada **"Pausas térmicas"** já formatada pelo macro `FormatarPausasTermicas.bas` (ver `Pausas-Termicas/` no repositório), o `GerarAbaCSV` lê a seção "Totais por Colaborador" dessa aba (não recalcula nada, só repassa os números) e gera uma linha por colaborador no CSV, com `tipo_ocorrencia = Resumo Pausas Térmicas`. O painel mostra isso na seção "Pausas Térmicas — Resumo por colaborador": pausas corretas/curtas, trabalho correto/longo/curto entre pausas e marcações inválidas, com a linha em vermelho para quem tem marcação inválida. Se a aba não existir, essa seção simplesmente fica vazia — não trava o resto do CSV.

`GerarCSV.txt` é o mesmo código, salvo em `.txt` (algumas caixas de e-mail/corporativas bloqueiam anexos `.bas` por segurança). Pra usar: importe `GerarCSV.bas` normalmente pelo VBA (Alt+F11 > Arquivo > Importar Arquivo); se só tiver o `.txt`, renomeie a extensão pra `.bas` antes de importar, ou abra um módulo novo em branco no VBA e cole o conteúdo do `.txt` dentro.
