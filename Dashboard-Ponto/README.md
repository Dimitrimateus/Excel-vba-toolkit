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
| `cargo` | Não* | Cargo/função do colaborador (usada no gráfico "Ocorrências por cargo"). | `Auxiliar de Logística` |
| `tipo_ocorrencia` | Não* | Tipo da ocorrência. | `Atraso`, `Falta`, `Saída Antecipada`, `Hora Extra`, `Esquecimento de Registro`, `Abono` |
| `situacao` | Não | Descrição/justificativa da ocorrência (aparece no histórico). | `Atraso na entrada` |
| `status` | Não | Situação do tratamento da ocorrência. | `Pendente`, `Aprovado`, `Reprovado`, `Regularizado` |
| `duracao_minutos` | Não | Duração em minutos, **com sinal**: negativo para falta/atraso, positivo para hora extra. | `-30` (30 min de atraso), `120` (2h de hora extra) |
| `destino_horas_extra` | Não | Só faz sentido quando `duracao_minutos` é positivo. | `Banco de Horas` ou `Pagamento` |
| `data_tratativa_pontonet` | Não | Data (e hora, se tiver) em que a ocorrência foi tratada/aprovada/regularizada no PontoNet. Usada para calcular a demora. | `25/08/2026 14:30` |

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
