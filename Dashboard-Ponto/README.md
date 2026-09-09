# Painel de Ponto (Aurora Coop)

Dashboard em um único arquivo HTML (`index.html`) para visualizar as ocorrências de ponto dos colaboradores. Não depende de internet nem de instalação: o RH carrega o CSV de dados, e o próprio arquivo HTML desenha os gráficos no navegador.

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
| `tipo_ocorrencia` | Não* | Tipo da ocorrência. | `Atraso`, `Falta`, `Saída Antecipada`, `Hora Extra`, `Esquecimento de Registro`, `Abono` |
| `situacao` | Não | Descrição/justificativa da ocorrência (aparece no histórico). | `Atraso na entrada` |
| `status` | Não | Situação do tratamento da ocorrência. | `Pendente`, `Aprovado`, `Reprovado`, `Regularizado` |
| `duracao_minutos` | Não | Duração em minutos, **com sinal**: negativo para falta/atraso, positivo para hora extra. | `-30` (30 min de atraso), `120` (2h de hora extra) |
| `destino_horas_extra` | Não | Só faz sentido quando `duracao_minutos` é positivo. | `Banco de Horas` ou `Pagamento` |
| `data_tratativa_pontonet` | Não | Data (e hora, se tiver) em que a ocorrência foi tratada/aprovada/regularizada no PontoNet. Usada para calcular a demora. | `25/08/2026 14:30` |

`* `Se a coluna não existir ou vier vazia numa linha, o painel usa "Sem gestor" / "Sem setor" / "Não informado" no lugar, mas os gráficos correspondentes perdem o sentido. Vale a pena preencher.

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

No cabeçalho, o ícone (três triângulos ascendentes nas cores da marca) é uma recriação em SVG inspirada no símbolo real, e não uma cópia vetorial exata: como as imagens do logotipo foram enviadas coladas na conversa (sem um arquivo de imagem por trás), não foi possível extrair o vetor original com precisão de pixel. Se você anexar os arquivos do logotipo (o mesmo jeito que anexou as planilhas, como arquivo e não colado na mensagem), dá pra embutir a arte oficial em vez desta recriação. Para trocar depois:

1. Abra `index.html` num editor de texto e procure pelo bloco `<svg class="brand-mark">` no `<header>`. Substitua o conteúdo por um `<img>` apontando para o logotipo oficial (como um `data:` URI embutido, para manter o arquivo funcionando sem internet) ou pelo SVG oficial.
2. O texto "AURORA COOP" ao lado do ícone é HTML/CSS (classe `.brand-name`), não uma imagem. Se o logotipo oficial já incluir a palavra "AURORA COOP" desenhada, você pode remover esse texto para não duplicar.

## Testando antes de enviar

Abra `index.html` em qualquer navegador (Chrome, Edge, Firefox) e carregue o `modelo-dados.csv` para ver o painel funcionando com dados de exemplo antes de trocá-lo pelos dados reais.
